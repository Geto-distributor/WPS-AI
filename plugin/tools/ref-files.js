"use strict";
// 参考文档的系统级操作：弹系统文件选择框（拿到真实路径），以及让 WPS 打开文件。
// WebView 里的 <input type=file> 只给文件内容不给路径，而「在 WPS 打开 / 委托代读」都要路径，所以选文件走这里。
const fs = require("fs");
const path = require("path");
const { spawn, spawnSync } = require("child_process");

const PICK_TIMEOUT_MS = 10 * 60 * 1000; // 用户在对话框里慢慢挑
const DOC_FILTER_EXTS = ["docx", "doc", "wps", "xlsx", "xls", "et", "csv", "pptx", "ppt", "dps", "pdf", "txt", "md"];

function run(cmd, args, timeout) {
  try { return spawnSync(cmd, args, { encoding: "utf8", timeout, windowsHide: false }); } catch (e) { return { error: e }; }
}

// 文件选择框要等用户慢慢挑，必须异步：用 spawnSync 会把整个代理的事件循环卡住（聊天流式输出也跟着停）
function runAsync(cmd, args, timeout) {
  return new Promise((resolve) => {
    let child;
    try { child = spawn(cmd, args, { windowsHide: false }); } catch (e) { resolve({ error: e }); return; }
    let stdout = "";
    let stderr = "";
    const timer = setTimeout(() => { try { child.kill(); } catch (e) {} }, timeout);
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (d) => { stdout += d; });
    child.stderr.on("data", (d) => { stderr += d; });
    child.on("error", (error) => { clearTimeout(timer); resolve({ error, stdout, stderr }); });
    child.on("close", (status) => { clearTimeout(timer); resolve({ status, stdout, stderr }); });
  });
}

function splitPaths(stdout) {
  return String(stdout || "").split(/\r?\n/).map((s) => s.trim()).filter(Boolean).filter((p) => {
    try { return fs.statSync(p).isFile(); } catch (e) { return false; }
  });
}

async function pickMac() {
  const script = [
    "try",
    '  set picked to choose file with prompt "选择参考文档（可多选）" with multiple selections allowed',
    '  set out to ""',
    "  repeat with f in picked",
    "    set out to out & POSIX path of f & linefeed",
    "  end repeat",
    "  out",
    "on error number -128",
    '  ""',
    "end try"
  ].join("\n");
  const r = await runAsync("/usr/bin/osascript", ["-e", script], PICK_TIMEOUT_MS);
  if (r.error || r.status !== 0) throw new Error("打开文件选择框失败：" + String(r.stderr || (r.error && r.error.message) || r.status).slice(0, 200));
  return splitPaths(r.stdout);
}

async function pickWindows() {
  const filter = `文档|${DOC_FILTER_EXTS.map((e) => "*." + e).join(";")}|所有文件|*.*`;
  const script = [
    "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8",
    "Add-Type -AssemblyName System.Windows.Forms",
    "$dialog = New-Object System.Windows.Forms.OpenFileDialog",
    "$dialog.Title = '选择参考文档（可多选）'",
    "$dialog.Multiselect = $true",
    `$dialog.Filter = '${filter}'`,
    "if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { [Console]::Write(($dialog.FileNames -join \"`n\")) }"
  ].join("; ");
  const r = await runAsync("powershell.exe", ["-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-Command", script], PICK_TIMEOUT_MS);
  if (r.error || r.status !== 0) throw new Error("打开文件选择框失败：" + String(r.stderr || (r.error && r.error.message) || r.status).slice(0, 200));
  return splitPaths(r.stdout);
}

async function pickLinux() {
  const r = await runAsync("zenity", ["--file-selection", "--multiple", "--separator=\n", "--title=选择参考文档"], PICK_TIMEOUT_MS);
  if (r.error && r.error.code === "ENOENT") throw new Error("没有找到 zenity，无法弹出文件选择框");
  if (r.status === 1) return []; // 用户取消
  if (r.error || r.status !== 0) throw new Error("打开文件选择框失败：" + String(r.stderr || r.status).slice(0, 200));
  return splitPaths(r.stdout);
}

// 弹系统文件选择框，返回选中的文件路径（取消时返回空数组）
async function pickFiles() {
  if (process.platform === "darwin") return pickMac();
  if (process.platform === "win32") return pickWindows();
  return pickLinux();
}

// Windows：找正在运行的 WPS 所在目录，按文件类型挑组件程序打开；找不到就交给系统默认程序
function windowsWpsLauncher(filePath) {
  const r = run("powershell.exe", ["-NoProfile", "-Command",
    "[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; (Get-Process -Name wps,et,wpp,wpspdf -ErrorAction SilentlyContinue | Where-Object { $_.Path } | Select-Object -First 1).Path"], 5000);
  const running = String(r.stdout || "").trim();
  if (!running) return null;
  const dir = path.dirname(running);
  const ext = path.extname(filePath).toLowerCase();
  const exe = /^\.(xlsx?|xlsm|et|ett|csv)$/.test(ext) ? "et.exe"
    : /^\.(pptx?|pptm|dps|dpt|pps)$/.test(ext) ? "wpp.exe"
    : ext === ".pdf" ? "wpspdf.exe"
    : "wps.exe";
  const candidate = path.join(dir, exe);
  if (fs.existsSync(candidate)) return candidate;
  return fs.existsSync(path.join(dir, "wps.exe")) ? path.join(dir, "wps.exe") : null;
}

// 让 WPS 打开文件。尽量不抢前台（macOS 用 open -g），但多数情况下 WPS 仍会切到新打开的文档 ——
// 切回原文档由插件侧负责（它知道原文档是哪个）。返回每个文件的结果。
function openInWps(paths) {
  return (Array.isArray(paths) ? paths : []).map((p) => {
    const filePath = String(p || "");
    try {
      if (!filePath || !fs.statSync(filePath).isFile()) return { path: filePath, ok: false, error: "文件不存在" };
    } catch (e) { return { path: filePath, ok: false, error: "文件不存在" }; }
    try {
      if (process.platform === "darwin") {
        let r = run("/usr/bin/open", ["-g", "-a", "wpsoffice", filePath], 10000);
        if (r.status !== 0) r = run("/usr/bin/open", ["-g", filePath], 10000);
        return r.status === 0 ? { path: filePath, ok: true, via: "open" } : { path: filePath, ok: false, error: String(r.stderr || r.status).slice(0, 200) };
      }
      if (process.platform === "win32") {
        const exe = windowsWpsLauncher(filePath);
        const child = exe
          ? spawn(exe, [filePath], { detached: true, stdio: "ignore", windowsHide: false })
          : spawn("cmd.exe", ["/c", "start", "", filePath], { detached: true, stdio: "ignore", windowsHide: true });
        child.unref();
        return { path: filePath, ok: true, via: exe ? path.basename(exe) : "default-app" };
      }
      const ext = path.extname(filePath).toLowerCase();
      const cmd = /^\.(xlsx?|et|csv)$/.test(ext) ? "et" : /^\.(pptx?|dps)$/.test(ext) ? "wpp" : ext === ".pdf" ? "xdg-open" : "wps";
      const child = spawn(cmd, [filePath], { detached: true, stdio: "ignore" });
      child.on("error", () => { try { spawn("xdg-open", [filePath], { detached: true, stdio: "ignore" }).unref(); } catch (e) {} });
      child.unref();
      return { path: filePath, ok: true, via: cmd };
    } catch (e) {
      return { path: filePath, ok: false, error: e && e.message ? e.message : String(e) };
    }
  });
}

module.exports = { pickFiles, openInWps, DOC_FILTER_EXTS };
