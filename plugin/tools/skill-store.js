"use strict";
// 技能的磁盘存储：~/.lingxi-ai/skills/<技能目录>/SKILL.md (+ references/ 等文本文件)。
// 目录就是技能本身 —— 用户可以直接用编辑器改、拷文件夹进来；前端 WpsAiSkills 经 /skills/* 读写。
// 只处理文本文件；所有路径都校验后再落盘，绝不写出技能根目录。
const fs = require("fs");
const os = require("os");
const path = require("path");
const crypto = require("crypto");

const LIMITS = {
  maxSkills: 500,
  maxFilesPerSkill: 100,
  maxFileBytes: 512 * 1024,
  maxDepth: 6
};
const TRASH_DIR = ".trash";
const MAIN_FILE = "SKILL.md";
// 随插件发布的技能装到磁盘后，在技能目录里放这个标记文件：记下安装时每个文件的指纹，
// 升级时据此区分「用户没动过（可以更新）」和「用户改过（保留用户的）」。以 . 开头，不会被当成技能文件。
const BUNDLED_MARKER = ".lingxi-bundled.json";
const sha1 = (text) => crypto.createHash("sha1").update(text, "utf8").digest("hex");

function defaultRoot() {
  return process.env.LINGXI_SKILLS_DIR || path.join(os.homedir(), ".lingxi-ai", "skills");
}

// 技能目录名：单段、不以 . 开头、不含路径分隔符和 Windows 保留字符
function assertDirName(dir) {
  const s = String(dir == null ? "" : dir);
  // eslint-disable-next-line no-control-regex
  if (!s || s.length > 120 || s.startsWith(".") || /[\\/:*?"<>|\u0000-\u001f]/.test(s) || s !== s.trim() || /[. ]$/.test(s)) {
    throw new Error(`技能目录名无效：${s}`);
  }
  return s;
}

// 技能内的相对路径：正斜杠、无 .. / 空段 / 隐藏段
function assertRelPath(rel) {
  const parts = String(rel == null ? "" : rel).replace(/\\/g, "/").split("/");
  // eslint-disable-next-line no-control-regex
  if (!parts.length || parts.some((seg) => !seg || seg === "." || seg === ".." || seg.startsWith(".") || /[:*?"<>|\u0000-\u001f]/.test(seg))) {
    throw new Error(`文件路径无效：${rel}`);
  }
  if (parts.length > LIMITS.maxDepth) throw new Error("目录层级过深");
  return parts.join("/");
}

function resolveInside(base, rel) {
  const full = path.resolve(base, rel);
  if (full !== base && !full.startsWith(base + path.sep)) throw new Error("路径越界");
  return full;
}

function looksBinary(buf) {
  const n = Math.min(buf.length, 8192);
  for (let i = 0; i < n; i += 1) if (buf[i] === 0) return true;
  return false;
}

function createSkillStore(rootDir) {
  const root = path.resolve(rootDir || defaultRoot());
  const skillDir = (dir) => resolveInside(root, assertDirName(dir));

  function readSkillFiles(base) {
    const files = [];
    const walk = (abs, rel, depth) => {
      if (depth > LIMITS.maxDepth) return;
      let entries;
      try { entries = fs.readdirSync(abs, { withFileTypes: true }); } catch (e) { return; }
      entries.sort((a, b) => a.name.localeCompare(b.name));
      for (const ent of entries) {
        if (ent.name.startsWith(".") || files.length >= LIMITS.maxFilesPerSkill) continue;
        const childAbs = path.join(abs, ent.name);
        const childRel = rel ? `${rel}/${ent.name}` : ent.name;
        if (ent.isDirectory()) { walk(childAbs, childRel, depth + 1); continue; }
        if (!ent.isFile()) continue; // 不跟符号链接，避免读到技能目录之外
        try {
          if (fs.statSync(childAbs).size > LIMITS.maxFileBytes) continue;
          const buf = fs.readFileSync(childAbs);
          if (looksBinary(buf)) continue;
          files.push({ path: childRel, content: buf.toString("utf8").replace(/^﻿/, "") });
        } catch (e) { /* 读不了的文件跳过 */ }
      }
    };
    walk(base, "", 1);
    return files;
  }

  // 列出全部技能：根目录下每个含 SKILL.md 的子目录算一个。返回原始文本，解析交给前端。
  function list() {
    let entries;
    try { entries = fs.readdirSync(root, { withFileTypes: true }); } catch (e) { return { root, skills: [] }; }
    const skills = [];
    entries.sort((a, b) => a.name.localeCompare(b.name));
    for (const ent of entries) {
      if (!ent.isDirectory() || ent.name.startsWith(".") || skills.length >= LIMITS.maxSkills) continue;
      const base = path.join(root, ent.name);
      let mainName = null;
      try { mainName = fs.readdirSync(base).find((n) => /^skill\.(md|markdown)$/i.test(n)); } catch (e) {}
      if (!mainName) continue;
      let mtime = 0;
      try { mtime = fs.statSync(path.join(base, mainName)).mtimeMs; } catch (e) {}
      const item = { dir: ent.name, mtime, files: readSkillFiles(base) };
      if (fs.existsSync(path.join(base, BUNDLED_MARKER))) item.bundled = true;
      skills.push(item);
    }
    return { root, skills };
  }

  function writeFile(dir, rel, content) {
    const base = skillDir(dir);
    const full = resolveInside(base, assertRelPath(rel));
    const text = String(content == null ? "" : content);
    if (Buffer.byteLength(text, "utf8") > LIMITS.maxFileBytes) throw new Error("文件过大");
    fs.mkdirSync(path.dirname(full), { recursive: true });
    // 先写临时文件再改名：写到一半崩了不会留下半截 SKILL.md
    const tmp = `${full}.${process.pid}.tmp`;
    fs.writeFileSync(tmp, text, "utf8");
    fs.renameSync(tmp, full);
  }

  // 整包写入（导入 / 迁移）：files = [{ path, content }]，必须含 SKILL.md。目录已存在则报错，不覆盖。
  function put(dir, files) {
    const base = skillDir(dir);
    const list_ = Array.isArray(files) ? files : [];
    if (!list_.some((f) => f && f.path === MAIN_FILE)) throw new Error("缺少 SKILL.md");
    if (list_.length > LIMITS.maxFilesPerSkill) throw new Error("文件数超过上限");
    if (fs.existsSync(base)) throw new Error(`技能目录已存在：${dir}`);
    list_.forEach((f) => assertRelPath(f.path)); // 全部校验通过再动磁盘
    try {
      list_.forEach((f) => writeFile(dir, f.path, f.content));
    } catch (e) {
      try { fs.rmSync(base, { recursive: true, force: true }); } catch (err) {}
      throw e;
    }
  }

  function deleteFile(dir, rel) {
    const base = skillDir(dir);
    const safeRel = assertRelPath(rel);
    if (/^skill\.(md|markdown)$/i.test(safeRel)) throw new Error("不能单独删除 SKILL.md");
    const full = resolveInside(base, safeRel);
    if (!fs.existsSync(full)) return false;
    fs.unlinkSync(full);
    // 顺手清掉变空的父目录（不越过技能根）
    let parent = path.dirname(full);
    while (parent !== base && parent.startsWith(base + path.sep)) {
      try { if (fs.readdirSync(parent).length) break; fs.rmdirSync(parent); } catch (e) { break; }
      parent = path.dirname(parent);
    }
    return true;
  }

  function rename(from, to) {
    const src = skillDir(from);
    const dst = skillDir(to);
    if (!fs.existsSync(src)) throw new Error(`技能目录不存在：${from}`);
    if (fs.existsSync(dst)) throw new Error(`技能目录已存在：${to}`);
    fs.renameSync(src, dst);
  }

  // 删除技能 = 挪进 .trash/<目录>-<时间戳>，不直接抹掉；误删可以手动捞回来。
  function remove(dir) {
    const src = skillDir(dir);
    if (!fs.existsSync(src)) return false;
    const trash = path.join(root, TRASH_DIR);
    fs.mkdirSync(trash, { recursive: true });
    fs.renameSync(src, path.join(trash, `${dir}-${Date.now()}`));
    return true;
  }

  // 把随插件发布的技能（sourceDir：含 SKILL.md + 数据文件）装到 <root>/<dir>/，让它的参考文件能被按需读取。
  //   - 目录不存在 → 整包装上
  //   - 目录存在但没有标记文件 → 是用户自己的同名技能，不碰
  //   - 插件升级（源文件变了）→ 只更新用户没改过的文件；用户改过的保留
  // 返回 "installed" | "current" | "updated" | "skipped"
  function installBundled(sourceDir, dir) {
    const base = skillDir(dir);
    const source = readSkillFiles(path.resolve(sourceDir));
    if (!source.some((f) => /^skill\.(md|markdown)$/i.test(f.path))) throw new Error(`内置技能源目录缺少 SKILL.md：${sourceDir}`);
    const hashes = {};
    source.forEach((f) => { hashes[f.path] = sha1(f.content); });
    const markerPath = path.join(base, BUNDLED_MARKER);
    const writeMarker = () => fs.writeFileSync(markerPath, JSON.stringify({ files: hashes }, null, 2), "utf8");

    if (!fs.existsSync(base)) {
      source.forEach((f) => writeFile(dir, f.path, f.content));
      writeMarker();
      return "installed";
    }
    let marker = null;
    try { marker = JSON.parse(fs.readFileSync(markerPath, "utf8")); } catch (e) {}
    if (!marker || typeof marker.files !== "object") return "skipped";
    const old = marker.files;
    const same = Object.keys(old).length === source.length && source.every((f) => old[f.path] === hashes[f.path]);
    if (same) return "current";

    const diskHash = (rel) => {
      try { return sha1(fs.readFileSync(resolveInside(base, rel), "utf8").replace(/^\uFEFF/, "")); } catch (e) { return null; }
    };
    source.forEach((f) => {
      const onDisk = diskHash(f.path);
      // 磁盘上没有 / 和上次安装时一模一样 → 用户没动过，换成新版；否则是用户改过的，留着
      if (onDisk === null || onDisk === old[f.path]) writeFile(dir, f.path, f.content);
    });
    Object.keys(old).forEach((rel) => {
      if (hashes[rel] !== undefined) return;
      try { if (diskHash(rel) === old[rel]) deleteFile(dir, rel); } catch (e) {} // 新版删掉的文件：用户没改过才跟着删
    });
    writeMarker();
    return "updated";
  }

  return { root, list, writeFile, put, deleteFile, rename, remove, installBundled };
}

module.exports = { createSkillStore, defaultRoot, LIMITS, BUNDLED_MARKER };
