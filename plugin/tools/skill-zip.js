"use strict";
// 技能包（.zip）解压：只取文本文件，给前端「导入多文件技能」用。
// WPS 各端 WebView 版本参差（DecompressionStream 不一定有），所以解压放在代理侧，用 Node 自带 zlib。
// 只读内存里的 buffer、只返回文本，不落盘 —— 没有 zip-slip 风险；体积/数量上限防 zip 炸弹。
const zlib = require("zlib");

const DEFAULT_LIMITS = {
  maxZipBytes: 20 * 1024 * 1024,
  maxEntries: 2000,
  maxFiles: 100,
  maxFileBytes: 512 * 1024,
  maxTotalBytes: 4 * 1024 * 1024
};

const SIG_EOCD = 0x06054b50;
const SIG_CENTRAL = 0x02014b50;
const SIG_LOCAL = 0x04034b50;

function findEocd(buf) {
  // EOCD 定长 22 字节 + 最长 65535 字节注释，从尾部往前找签名
  const min = Math.max(0, buf.length - 22 - 0xffff);
  for (let i = buf.length - 22; i >= min; i -= 1) {
    if (buf.readUInt32LE(i) === SIG_EOCD) return i;
  }
  return -1;
}

function looksBinary(data) {
  const n = Math.min(data.length, 8192);
  for (let i = 0; i < n; i += 1) if (data[i] === 0) return true;
  return false;
}

// 读 zip 中央目录，列出全部条目（不解压）。docx / xlsx / pptx 本质也是 zip，文档解析（doc-extract.js）复用这里。
function listZipEntries(buf, limits) {
  const lim = Object.assign({}, DEFAULT_LIMITS, limits || {});
  if (!Buffer.isBuffer(buf) || buf.length < 22) throw new Error("不是有效的 zip 文件");
  if (buf.length > lim.maxZipBytes) throw new Error(`zip 超过 ${Math.round(lim.maxZipBytes / 1024 / 1024)}MB 上限`);
  const eocd = findEocd(buf);
  if (eocd < 0) throw new Error("不是有效的 zip 文件（找不到目录结尾）");
  const total = buf.readUInt16LE(eocd + 10);
  const cdOffset = buf.readUInt32LE(eocd + 16);
  if (total === 0xffff || cdOffset === 0xffffffff) throw new Error("暂不支持 zip64 格式的压缩包");
  if (total > lim.maxEntries) throw new Error(`zip 内条目过多（${total}）`);
  const entries = [];
  let p = cdOffset;
  for (let i = 0; i < total; i += 1) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== SIG_CENTRAL) throw new Error("zip 目录损坏");
    const nameLen = buf.readUInt16LE(p + 28);
    entries.push({
      name: buf.slice(p + 46, p + 46 + nameLen).toString("utf8").replace(/\\/g, "/"),
      flags: buf.readUInt16LE(p + 8),
      method: buf.readUInt16LE(p + 10),
      compSize: buf.readUInt32LE(p + 20),
      rawSize: buf.readUInt32LE(p + 24),
      localOffset: buf.readUInt32LE(p + 42)
    });
    p += 46 + nameLen + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32);
  }
  return entries;
}

// 解压单个条目；不支持 / 损坏 / 超限时抛错，message 即跳过原因
function readZipEntry(buf, entry, maxBytes) {
  if (entry.flags & 0x1) throw new Error("加密");
  if (entry.method !== 0 && entry.method !== 8) throw new Error("不支持的压缩方式");
  if (entry.rawSize > maxBytes) throw new Error("文件过大");
  const lo = entry.localOffset;
  if (lo + 30 > buf.length || buf.readUInt32LE(lo) !== SIG_LOCAL) throw new Error("条目损坏");
  const dataStart = lo + 30 + buf.readUInt16LE(lo + 26) + buf.readUInt16LE(lo + 28);
  const comp = buf.slice(dataStart, dataStart + entry.compSize);
  let data;
  try {
    // maxOutputLength：声明的 rawSize 不可信，解压时再卡一道
    data = entry.method === 0 ? comp : zlib.inflateRawSync(comp, { maxOutputLength: maxBytes });
  } catch (e) { throw new Error("解压失败"); }
  if (data.length > maxBytes) throw new Error("文件过大");
  return data;
}

/**
 * @param {Buffer} buf zip 文件内容
 * @returns {{ files: {path:string, content:string}[], skipped: {path:string, reason:string}[] }}
 */
function extractTextFiles(buf, limits) {
  const lim = Object.assign({}, DEFAULT_LIMITS, limits || {});
  const files = [];
  const skipped = [];
  let totalBytes = 0;
  for (const entry of listZipEntries(buf, lim)) {
    const name = entry.name;
    if (!name || name.endsWith("/")) continue; // 目录条目
    const skip = (reason) => skipped.push({ path: name, reason });
    if (entry.rawSize <= lim.maxFileBytes && files.length >= lim.maxFiles) { skip("文件数超过上限"); continue; }
    if (entry.rawSize <= lim.maxFileBytes && totalBytes + entry.rawSize > lim.maxTotalBytes) { skip("总大小超过上限"); continue; }
    let data;
    try { data = readZipEntry(buf, entry, lim.maxFileBytes); } catch (e) { skip(e.message); continue; }
    if (looksBinary(data)) { skip("非文本文件"); continue; }
    totalBytes += data.length;
    files.push({ path: name, content: data.toString("utf8").replace(/^\uFEFF/, "") });
  }
  return { files, skipped };
}

module.exports = { extractTextFiles, listZipEntries, readZipEntry, DEFAULT_LIMITS };
