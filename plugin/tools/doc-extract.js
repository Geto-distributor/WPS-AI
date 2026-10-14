"use strict";
// 参考文档的磁盘读取：把 .docx / .xlsx / .pptx / .pdf / 纯文本解析成「带分节标题的纯文本」。
// 给合同评审这类场景当背景资料用 —— 只求文字和表格数据准确，不还原版式。
// docx / xlsx / pptx 本质是 zip + XML，复用 skill-zip 的解压；PDF 由调用方注入（代理里已有 pdfjs 实现）。
const fs = require("fs");
const path = require("path");
const { listZipEntries, readZipEntry } = require("./skill-zip.js");

const LIMITS = {
  maxFileBytes: 60 * 1024 * 1024,   // 源文件
  maxPartBytes: 40 * 1024 * 1024,   // zip 内单个 XML 部件（大表格的 sheet.xml 可能很大）
  maxChars: 3 * 1000 * 1000,        // 解析结果总字符上限，超出截断
  maxSheetRows: 20000,
  maxSheetCols: 200
};

const KIND_BY_EXT = {
  ".docx": "word", ".docm": "word",
  ".xlsx": "excel", ".xlsm": "excel",
  ".pptx": "ppt", ".pptm": "ppt",
  ".pdf": "pdf",
  ".txt": "text", ".md": "text", ".markdown": "text", ".csv": "text", ".tsv": "text", ".json": "text", ".xml": "text", ".html": "text", ".htm": "text"
};
// WPS 能打开、但这里解析不了的老格式 / 自有格式：提示用户在 WPS 里打开后走委托代读
const LEGACY_EXT = new Set([".doc", ".wps", ".wpt", ".dot", ".rtf", ".xls", ".et", ".ett", ".ppt", ".dps", ".dpt", ".pps"]);

function kindOf(filePath) {
  const ext = path.extname(String(filePath || "")).toLowerCase();
  if (KIND_BY_EXT[ext]) return KIND_BY_EXT[ext];
  if (LEGACY_EXT.has(ext)) return "legacy";
  return "unknown";
}

function decodeXml(s) {
  return String(s)
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_m, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_m, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&amp;/g, "&");
}

function zipReader(buf) {
  const entries = listZipEntries(buf, { maxZipBytes: LIMITS.maxFileBytes, maxEntries: 20000 });
  const byName = new Map(entries.map((e) => [e.name, e]));
  return {
    has: (name) => byName.has(name),
    names: () => entries.map((e) => e.name),
    text: (name) => {
      const e = byName.get(name);
      return e ? readZipEntry(buf, e, LIMITS.maxPartBytes).toString("utf8").replace(/^﻿/, "") : null;
    }
  };
}

// _rels：Id → Target（相对 baseDir 解析成 zip 内路径）
function readRels(zip, relsPath, baseDir) {
  const xml = zip.text(relsPath);
  const map = new Map();
  if (!xml) return map;
  const re = /<Relationship\b([^>]*)\/?>/g;
  let m;
  while ((m = re.exec(xml))) {
    const id = (m[1].match(/\bId="([^"]*)"/) || [])[1];
    const target = (m[1].match(/\bTarget="([^"]*)"/) || [])[1];
    if (!id || !target) continue;
    const resolved = target.startsWith("/") ? target.slice(1) : path.posix.normalize(path.posix.join(baseDir, target));
    map.set(id, resolved);
  }
  return map;
}

// ---------- Word ----------
// 顺序扫 document.xml 的标签：段落成行；表格逐行输出成「| 单元格 | 单元格 |」，单元格内多段用 / 连接。
function docxToText(zip) {
  const xml = zip.text("word/document.xml");
  if (xml == null) throw new Error("不是有效的 Word 文档（缺少 word/document.xml）");
  const out = [];
  const tables = []; // 栈：{ rows: [], row: null, cell: null }
  let para = null;
  let inText = false;
  const re = /<(\/?)(w:p|w:tbl|w:tr|w:tc|w:t|w:tab|w:br|w:cr|w:delText)\b[^>]*?(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(xml))) {
    if (m[4] != null) { if (inText && para != null) para.push(decodeXml(m[4])); continue; }
    const closing = m[1] === "/";
    const selfClosing = m[3] === "/";
    const tag = m[2];
    const tbl = tables[tables.length - 1];
    if (tag === "w:t") { inText = !closing && !selfClosing; continue; }
    if (tag === "w:delText") { inText = false; continue; } // 修订里被删掉的文字不算正文
    if (tag === "w:tab") { if (para) para.push("\t"); continue; }
    if (tag === "w:br" || tag === "w:cr") { if (para) para.push("\n"); continue; }
    if (tag === "w:p") {
      if (selfClosing) { if (!tbl) out.push(""); continue; }
      if (!closing) { para = []; continue; }
      const text = (para || []).join("").replace(/[ \t]+$/g, "");
      para = null;
      if (tbl && tbl.cell) tbl.cell.push(text);
      else out.push(text);
      continue;
    }
    if (tag === "w:tbl") {
      if (!closing) { tables.push({ row: null, cell: null }); out.push(""); }
      else { tables.pop(); out.push(""); }
      continue;
    }
    if (!tbl) continue;
    if (tag === "w:tr") {
      if (!closing) tbl.row = [];
      else { if (tbl.row) out.push("| " + tbl.row.join(" | ") + " |"); tbl.row = null; }
      continue;
    }
    if (tag === "w:tc") {
      if (!closing) tbl.cell = [];
      else {
        const cellText = (tbl.cell || []).map((t) => t.trim()).filter(Boolean).join(" / ").replace(/\|/g, "｜").replace(/\n/g, " ");
        if (tbl.row) tbl.row.push(cellText);
        tbl.cell = null;
      }
    }
  }
  const body = out.join("\n").replace(/\n{3,}/g, "\n\n").trim();
  return { sections: [{ title: "正文", text: body }] };
}

// ---------- Excel ----------
function colIndex(ref) {
  const letters = String(ref || "").replace(/[^A-Z]/gi, "").toUpperCase();
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function xlsxToText(zip) {
  const wbXml = zip.text("xl/workbook.xml");
  if (wbXml == null) throw new Error("不是有效的 Excel 文件（缺少 xl/workbook.xml）");
  const rels = readRels(zip, "xl/_rels/workbook.xml.rels", "xl");
  const shared = [];
  const ssXml = zip.text("xl/sharedStrings.xml");
  if (ssXml) {
    const siRe = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
    let si;
    while ((si = siRe.exec(ssXml))) {
      // 富文本里的注音（rPh）不是单元格内容
      const inner = si[1].replace(/<rPh\b[\s\S]*?<\/rPh>/g, "");
      let text = "";
      const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
      let t;
      while ((t = tRe.exec(inner))) text += decodeXml(t[1]);
      shared.push(text);
    }
  }
  const sections = [];
  const sheetRe = /<sheet\b([^>]*)\/?>/g;
  let sm;
  while ((sm = sheetRe.exec(wbXml))) {
    const name = decodeXml((sm[1].match(/\bname="([^"]*)"/) || [])[1] || "Sheet");
    const rid = (sm[1].match(/\br:id="([^"]*)"/) || [])[1];
    const state = (sm[1].match(/\bstate="([^"]*)"/) || [])[1];
    const target = rid && rels.get(rid);
    if (!target || !zip.has(target)) continue;
    const xml = zip.text(target);
    const rows = [];
    let maxCol = -1;
    let truncated = false;
    const rowRe = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
    let rm;
    while ((rm = rowRe.exec(xml))) {
      const rIdx = parseInt((rm[1].match(/\br="(\d+)"/) || [])[1] || String(rows.length + 1), 10) - 1;
      if (rIdx >= LIMITS.maxSheetRows) { truncated = true; break; }
      const cells = [];
      const cRe = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let cm;
      while ((cm = cRe.exec(rm[2]))) {
        const attrs = cm[1];
        const ref = (attrs.match(/\br="([A-Z]+)\d+"/i) || [])[1];
        const ci = ref ? colIndex(ref) : cells.length;
        if (ci >= LIMITS.maxSheetCols) continue;
        const type = (attrs.match(/\bt="([^"]*)"/) || [])[1] || "n";
        const inner = cm[2] || "";
        let value = "";
        if (type === "inlineStr") {
          const tRe = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
          let t;
          while ((t = tRe.exec(inner))) value += decodeXml(t[1]);
        } else {
          const v = (inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
          if (v != null) {
            if (type === "s") value = shared[parseInt(v, 10)] || "";
            else if (type === "b") value = v === "1" ? "TRUE" : "FALSE";
            else value = decodeXml(v);
          }
        }
        cells[ci] = value.replace(/[\t\r\n]+/g, " ");
        if (value !== "" && ci > maxCol) maxCol = ci;
      }
      rows[rIdx] = cells;
    }
    const lines = [];
    for (let r = 0; r < rows.length; r += 1) {
      const cells = rows[r] || [];
      const line = [];
      for (let c = 0; c <= maxCol; c += 1) line.push(cells[c] == null ? "" : cells[c]);
      if (line.some((v) => v !== "")) lines.push(`${r + 1}\t${line.join("\t")}`);
    }
    let title = `工作表「${name}」`;
    if (state === "hidden" || state === "veryHidden") title += "（隐藏）";
    const header = maxCol >= 0
      ? `行号\t${Array.from({ length: maxCol + 1 }, (_v, i) => columnLetter(i)).join("\t")}`
      : "";
    let text = lines.length ? `${header}\n${lines.join("\n")}` : "（空表）";
    if (truncated) text += `\n（行数超过 ${LIMITS.maxSheetRows}，后面的行未读取）`;
    sections.push({ title, text });
  }
  if (!sections.length) sections.push({ title: "工作簿", text: "（没有可读取的工作表）" });
  return { sections, note: "日期单元格显示为 Excel 序列号（如 45200），公式单元格显示上次保存时的计算结果。" };
}

function columnLetter(i) {
  let n = i + 1;
  let s = "";
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

// ---------- PPT ----------
function pptxToText(zip) {
  const presXml = zip.text("ppt/presentation.xml");
  if (presXml == null) throw new Error("不是有效的 PPT 文件（缺少 ppt/presentation.xml）");
  const rels = readRels(zip, "ppt/_rels/presentation.xml.rels", "ppt");
  const ids = [];
  const idRe = /<p:sldId\b([^>]*)\/?>/g;
  let m;
  while ((m = idRe.exec(presXml))) {
    const rid = (m[1].match(/\br:id="([^"]*)"/) || [])[1];
    if (rid && rels.get(rid)) ids.push(rels.get(rid));
  }
  const sections = ids.map((target, i) => {
    const xml = zip.text(target) || "";
    const paras = [];
    const pRe = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
    let pm;
    while ((pm = pRe.exec(xml))) {
      let text = "";
      const tRe = /<a:t\b[^>]*>([\s\S]*?)<\/a:t>|<a:br\b[^>]*\/>/g;
      let t;
      while ((t = tRe.exec(pm[1]))) text += t[1] != null ? decodeXml(t[1]) : "\n";
      if (text.trim()) paras.push(text.trim());
    }
    return { title: `第 ${i + 1} 页`, text: paras.join("\n") || "（本页没有文字）" };
  });
  if (!sections.length) sections.push({ title: "演示文稿", text: "（没有幻灯片）" });
  return { sections };
}

// 分节拼成一整段文本，并记下每节的起始字符位置，方便分段读取和按节跳转
function assemble(sections) {
  let text = "";
  const index = [];
  sections.forEach((s) => {
    if (text) text += "\n\n";
    index.push({ title: s.title, offset: text.length });
    text += `## ${s.title}\n${s.text}`;
  });
  let truncated = false;
  if (text.length > LIMITS.maxChars) { text = text.slice(0, LIMITS.maxChars); truncated = true; }
  return { text, sections: index.filter((s) => s.offset < text.length), truncated };
}

/**
 * 读取并解析一个文件。
 * @param {string} filePath
 * @param {{ extractPdfText?: (p: string) => Promise<{pages: {page:number,text:string}[], hasText?: boolean}> }} deps
 * @returns {Promise<{ kind, text, sections: {title, offset}[], truncated, note? }>}
 */
async function extractDocument(filePath, deps = {}) {
  const kind = kindOf(filePath);
  if (kind === "legacy") throw new Error("这是老格式（.doc / .xls / .ppt / WPS 格式），无法直接从磁盘解析，请在 WPS 里打开后再读取");
  if (kind === "unknown") throw new Error("不支持的文件类型：" + (path.extname(filePath) || "无扩展名"));
  const st = fs.statSync(filePath);
  if (!st.isFile()) throw new Error("不是文件：" + filePath);
  if (st.size > LIMITS.maxFileBytes) throw new Error(`文件太大（${(st.size / 1048576).toFixed(0)}MB）`);

  let parsed;
  if (kind === "pdf") {
    if (typeof deps.extractPdfText !== "function") throw new Error("PDF 解析不可用");
    const r = await deps.extractPdfText(filePath);
    parsed = { sections: (r.pages || []).map((p) => ({ title: `第 ${p.page} 页`, text: p.text || "（本页没有文字）" })) };
    if (r.hasText === false) parsed.note = "这份 PDF 几乎没有文字层（可能是扫描件），提取出来的内容不可靠";
  } else if (kind === "text") {
    parsed = { sections: [{ title: path.basename(filePath), text: fs.readFileSync(filePath, "utf8").replace(/^﻿/, "") }] };
  } else {
    const zip = zipReader(fs.readFileSync(filePath));
    parsed = kind === "word" ? docxToText(zip) : kind === "excel" ? xlsxToText(zip) : pptxToText(zip);
  }
  const assembled = assemble(parsed.sections);
  return Object.assign({ kind, mtimeMs: st.mtimeMs, size: st.size }, assembled, parsed.note ? { note: parsed.note } : {});
}

// 按 mtime 缓存解析结果：同一份文件分页读取时不用反复解析
function createExtractor(deps) {
  const cache = new Map();
  async function get(filePath) {
    const key = path.resolve(filePath).toLowerCase();
    const st = fs.statSync(filePath);
    const hit = cache.get(key);
    if (hit && hit.mtimeMs === st.mtimeMs && hit.size === st.size) return hit;
    const result = await extractDocument(filePath, deps);
    if (cache.size >= 16) cache.delete(cache.keys().next().value);
    cache.set(key, result);
    return result;
  }
  return { get };
}

// 从解析结果里切一页：offset / maxChars，或按节名（section）跳到某节开头
function sliceDocument(doc, { offset, maxChars, section } = {}) {
  const max = Math.min(Math.max(Number(maxChars) || 12000, 500), 60000);
  let start = Math.max(0, Number(offset) || 0);
  if (section) {
    const want = String(section).trim();
    const hit = doc.sections.find((s) => s.title === want) || doc.sections.find((s) => s.title.includes(want));
    if (hit) start = hit.offset;
  }
  const end = Math.min(doc.text.length, start + max);
  return {
    kind: doc.kind,
    text: doc.text.slice(start, end),
    offset: start,
    nextOffset: end < doc.text.length ? end : null,
    total: doc.text.length,
    sections: doc.sections,
    truncated: !!doc.truncated,
    note: doc.note
  };
}

module.exports = { extractDocument, createExtractor, sliceDocument, kindOf, LIMITS, _test: { docxToText, xlsxToText, pptxToText, zipReader, assemble } };
