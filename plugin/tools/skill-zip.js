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

/**
 * @param {Buffer} buf zip 文件内容
 * @returns {{ files: {path:string, content:string}[], skipped: {path:string, reason:string}[] }}
 */
function extractTextFiles(buf, limits) {
  const lim = Object.assign({}, DEFAULT_LIMITS, limits || {});
  if (!Buffer.isBuffer(buf) || buf.length < 22) throw new Error("不是有效的 zip 文件");
  if (buf.length > lim.maxZipBytes) throw new Error(`zip 超过 ${Math.round(lim.maxZipBytes / 1024 / 1024)}MB 上限`);
  const eocd = findEocd(buf);
  if (eocd < 0) throw new Error("不是有效的 zip 文件（找不到目录结尾）");
  const total = buf.readUInt16LE(eocd + 10);
  const cdOffset = buf.readUInt32LE(eocd + 16);
  if (total === 0xffff || cdOffset === 0xffffffff) throw new Error("暂不支持 zip64 格式的压缩包");
  if (total > lim.maxEntries) throw new Error(`zip 内条目过多（${total}）`);

  const files = [];
  const skipped = [];
  let totalBytes = 0;
  let p = cdOffset;
  for (let i = 0; i < total; i += 1) {
    if (p + 46 > buf.length || buf.readUInt32LE(p) !== SIG_CENTRAL) throw new Error("zip 目录损坏");
    const flags = buf.readUInt16LE(p + 8);
    const method = buf.readUInt16LE(p + 10);
    const compSize = buf.readUInt32LE(p + 20);
    const rawSize = buf.readUInt32LE(p + 24);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localOffset = buf.readUInt32LE(p + 42);
    const name = buf.slice(p + 46, p + 46 + nameLen).toString("utf8").replace(/\\/g, "/");
    p += 46 + nameLen + extraLen + commentLen;

    if (!name || name.endsWith("/")) continue; // 目录条目
    const skip = (reason) => skipped.push({ path: name, reason });
    if (flags & 0x1) { skip("加密"); continue; }
    if (method !== 0 && method !== 8) { skip("不支持的压缩方式"); continue; }
    if (rawSize > lim.maxFileBytes) { skip("文件过大"); continue; }
    if (files.length >= lim.maxFiles) { skip("文件数超过上限"); continue; }
    if (totalBytes + rawSize > lim.maxTotalBytes) { skip("总大小超过上限"); continue; }
    if (localOffset + 30 > buf.length || buf.readUInt32LE(localOffset) !== SIG_LOCAL) { skip("条目损坏"); continue; }
    const dataStart = localOffset + 30 + buf.readUInt16LE(localOffset + 26) + buf.readUInt16LE(localOffset + 28);
    const comp = buf.slice(dataStart, dataStart + compSize);
    let data;
    try {
      // maxOutputLength：声明的 rawSize 不可信，解压时再卡一道
      data = method === 0 ? comp : zlib.inflateRawSync(comp, { maxOutputLength: lim.maxFileBytes });
    } catch (e) { skip("解压失败"); continue; }
    if (data.length > lim.maxFileBytes) { skip("文件过大"); continue; }
    if (looksBinary(data)) { skip("非文本文件"); continue; }
    totalBytes += data.length;
    files.push({ path: name, content: data.toString("utf8").replace(/^﻿/, "") });
  }
  return { files, skipped };
}

module.exports = { extractTextFiles, DEFAULT_LIMITS };
