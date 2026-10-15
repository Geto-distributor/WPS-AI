#!/usr/bin/env node
"use strict";
// 安装插件时顺带安装技能：把 plugin/skills/<名字>/ 装进用户的技能目录（~/.lingxi-ai/skills/<名字>/）。
// 装好后和用户自己上传的技能没有区别（可改、可改名、可删）；技能还在就不动，不在了才装。
// 由各平台 post-install 脚本调用。任何失败都只打日志、退出码 0，不影响插件本身的安装。
//
// 用法：node install-bundled-skills.js [--skills <源技能目录>] [--root <用户技能目录>]
const fs = require("fs");
const path = require("path");
const { createSkillStore, defaultRoot } = require("./skill-store.js");

function arg(name, fallback) {
  const i = process.argv.indexOf(name);
  return i >= 0 && process.argv[i + 1] ? process.argv[i + 1] : fallback;
}

function main() {
  const sourceRoot = path.resolve(arg("--skills", path.join(__dirname, "..", "skills")));
  const root = path.resolve(arg("--root", defaultRoot()));
  let names = [];
  try {
    names = fs.readdirSync(sourceRoot, { withFileTypes: true }).filter((d) => d.isDirectory() && !d.name.startsWith(".")).map((d) => d.name);
  } catch (e) {
    console.log(`[skills] 没有随包技能（${sourceRoot}）`);
    return;
  }
  const store = createSkillStore(root);
  for (const name of names) {
    try {
      const result = store.preinstall(path.join(sourceRoot, name), name);
      console.log(`[skills] ${name}: ${result === "installed" ? "已安装" : "已存在，跳过"} → ${root}`);
    } catch (e) {
      console.log(`[skills] ${name}: 安装失败（${e && e.message ? e.message : e}）`);
    }
  }
}

try { main(); } catch (e) { console.log("[skills] 安装技能出错：" + (e && e.message ? e.message : e)); }
process.exit(0);
