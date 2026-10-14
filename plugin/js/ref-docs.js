// WpsAiRefDocs：参考文档（合同评审时的背景资料、材料清单等）。
//
// 用法：在 A 文档的 AI 面板里添加 B / C / D 作为参考文档，AI 评审 A 时按需读取它们。
// 读取有两条路，自动选：
//   1. 委托代读：B 已在 WPS 里打开 → 交给 B 所在组件（文字 / 表格 / 演示）的那份插件用 WPS 接口读，
//      能拿到未保存的修改。各组件的插件互相看不到对方的文档（实测 wps.EtApplication() 在文字窗口里为空），
//      所以经本地代理中转：每份插件定时登记自己开着的文档，并长轮询等「代读」请求。
//   2. 磁盘读取：没打开 / 对应插件不在线 / 代读失败 → 代理直接解析磁盘文件（docx / xlsx / pptx / pdf / 文本）。
// 本模块在每个加载了 main.js 的页面里都跑（各组件的后台页 + 面板页），设置等弹窗页不跑登记和代读。
(function attachRefDocs(global) {
  "use strict";

  const STORE_KEY = "lingxi_ref_docs_v1"; // { [conversationId]: [{ path, name, kind, addedAt }] }
  const PRESENCE_INTERVAL_MS = 5000;
  const MAX_SHEET_ROWS = 20000;
  const MAX_SHEET_COLS = 200;
  const instanceId = "ri-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
  const listeners = new Set();

  function base() { return global.WpsAiRuntime?.proxyBase?.() || "http://127.0.0.1:3890"; }
  async function call(route, body, method) {
    const resp = await fetch(base() + "/refs/" + route, body === undefined && method !== "POST"
      ? { method: "GET", cache: "no-store" }
      : { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body || {}) });
    const json = await resp.json().catch(() => null);
    if (!resp.ok || !json?.ok) {
      const err = new Error(json?.error || `参考文档服务请求失败（${resp.status}）`);
      err.status = resp.status;
      throw err;
    }
    return json;
  }

  function normPath(p) { return String(p || "").replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase(); }
  function baseName(p) { return String(p || "").split(/[\\/]/).pop() || String(p || ""); }
  function kindOf(p) {
    const ext = (String(p || "").match(/\.[^.\\/]+$/) || [""])[0].toLowerCase();
    if (/^\.(docx|docm|doc|wps|wpt|dot|rtf)$/.test(ext)) return "word";
    if (/^\.(xlsx|xlsm|xls|et|ett|csv)$/.test(ext)) return "excel";
    if (/^\.(pptx|pptm|ppt|dps|dpt|pps)$/.test(ext)) return "ppt";
    if (ext === ".pdf") return "pdf";
    return "text";
  }

  // ===== 参考文档清单（按对话存） =====
  function conversationId() {
    try { return global.WpsAiConversations?.getCurrent?.()?.id || "default"; } catch (e) { return "default"; }
  }
  function readAll() {
    try { const raw = global.WpsAiStore?.getItem?.(STORE_KEY); const v = raw ? JSON.parse(raw) : {}; return v && typeof v === "object" ? v : {}; } catch (e) { return {}; }
  }
  function writeAll(all) {
    try { global.WpsAiStore?.setItem?.(STORE_KEY, JSON.stringify(all)); } catch (e) {}
    listeners.forEach((cb) => { try { cb(); } catch (e) {} });
  }
  function list() {
    const items = readAll()[conversationId()];
    return Array.isArray(items) ? items.slice() : [];
  }
  function add(paths) {
    const all = readAll();
    const id = conversationId();
    const cur = Array.isArray(all[id]) ? all[id] : [];
    const added = [];
    (paths || []).forEach((p) => {
      const path = String(p || "").trim();
      if (!path || cur.some((r) => normPath(r.path) === normPath(path))) return;
      const ref = { path, name: baseName(path), kind: kindOf(path), addedAt: Date.now() };
      cur.push(ref);
      added.push(ref);
    });
    all[id] = cur;
    writeAll(all);
    return added;
  }
  function remove(path) {
    const all = readAll();
    const id = conversationId();
    all[id] = (all[id] || []).filter((r) => normPath(r.path) !== normPath(path));
    if (!all[id].length) delete all[id];
    writeAll(all);
  }
  function onChange(cb) { listeners.add(cb); return () => listeners.delete(cb); }

  // ===== 当前组件里已打开的文档（WPS 接口） =====
  async function getApp() {
    try { return await global.WpsAiDocument?.getApplication?.(); } catch (e) { return null; }
  }
  async function currentHost() {
    try { return (await global.WpsAiDocument?.getHostInfo?.())?.host || "unknown"; } catch (e) { return "unknown"; }
  }
  function each(coll, fn) {
    const n = Number(coll?.Count) || 0;
    for (let i = 1; i <= n; i += 1) {
      let item = null;
      try { item = coll.Item(i); } catch (e) { item = null; }
      if (item) fn(item, i);
    }
  }
  function collectionFor(app, host) {
    try {
      if (host === "wps") return app.Documents;
      if (host === "et") return app.Workbooks;
      if (host === "wpp") return app.Presentations;
    } catch (e) {}
    return null;
  }
  function activeFor(app, host) {
    try {
      if (host === "wps") return app.ActiveDocument;
      if (host === "et") return app.ActiveWorkbook;
      if (host === "wpp") return app.ActivePresentation;
    } catch (e) {}
    return null;
  }
  function docPath(doc) {
    try { return String(doc.FullName || ""); } catch (e) { return ""; }
  }

  async function collectOpenDocs() {
    const host = await currentHost();
    const app = await getApp();
    if (!app || host === "unknown") return { host, docs: [] };
    const docs = [];
    if (host === "pdf") {
      // PDF 组件没有统一的文档集合；只登记当前这一份，路径取不到就不登记（代读 PDF 也只走磁盘）
      try {
        const p = await global.WpsAiHostPdf?.getActivePdfPath?.();
        if (p) docs.push({ name: baseName(p), path: p, saved: true, active: true });
      } catch (e) {}
      return { host, docs };
    }
    const activePath = normPath(docPath(activeFor(app, host)));
    each(collectionFor(app, host), (doc) => {
      const path = docPath(doc);
      let name = "";
      let saved = true;
      try { name = String(doc.Name || ""); } catch (e) {}
      try { saved = doc.Saved !== false; } catch (e) {}
      docs.push({ name: name || baseName(path), path, saved, active: !!path && normPath(path) === activePath });
    });
    return { host, docs };
  }

  // ===== 用 WPS 接口读一份已打开的文档（代读时由它所在组件的插件执行）=====
  function columnLetter(i) {
    let n = i + 1;
    let s = "";
    while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  // Word 的 Range.Text 里，表格单元格以 \r\x07 结尾、每行末尾还有一个 \r\x07，空单元格和行尾分不开，
  // 所以表格单独按 Rows / Columns 切：能对上行列数就输出「| 单元格 | 单元格 |」，对不上（合并单元格）就粗略拼接。
  function formatWordTable(tbl) {
    let raw = "";
    try { raw = String(tbl.Range.Text || ""); } catch (e) { return null; }
    let rows = 0;
    let cols = 0;
    try { rows = Number(tbl.Rows.Count) || 0; cols = Number(tbl.Columns.Count) || 0; } catch (e) {}
    const tokens = raw.split("\r\x07");
    if (tokens[tokens.length - 1] === "") tokens.pop();
    const clean = (t) => t.replace(/\r/g, " / ").replace(/\x07/g, "").replace(/\|/g, "｜").trim();
    let lines;
    if (rows && cols && tokens.length === rows * (cols + 1)) {
      lines = [];
      for (let r = 0; r < rows; r += 1) lines.push("| " + tokens.slice(r * (cols + 1), r * (cols + 1) + cols).map(clean).join(" | ") + " |");
    } else {
      lines = ["| " + tokens.map(clean).filter(Boolean).join(" | ") + " |"];
    }
    return { raw, formatted: "\n" + lines.join("\n") + "\n\n" }; // 表格前后各空一行，和磁盘解析的输出一致
  }
  function readWordDoc(doc) {
    let text = "";
    try { text = String(doc.Content.Text || ""); } catch (e) { text = ""; }
    let from = 0;
    each(doc.Tables, (tbl) => {
      const t = formatWordTable(tbl);
      if (!t || !t.raw) return;
      const at = text.indexOf(t.raw, from);
      if (at < 0) return;
      text = text.slice(0, at) + t.formatted + text.slice(at + t.raw.length);
      from = at + t.formatted.length;
    });
    text = text.replace(/\x07/g, "").replace(/\r/g, "\n").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
    return [{ title: "正文", text }];
  }
  function readWorkbook(wb) {
    const sections = [];
    each(wb.Worksheets, (ws) => {
      let name = "Sheet";
      let hidden = false;
      try { name = String(ws.Name); } catch (e) {}
      try { hidden = ws.Visible === 0 || ws.Visible === false || ws.Visible === 2; } catch (e) {}
      let values = null;
      let row0 = 1;
      let col0 = 1;
      try {
        const used = ws.UsedRange;
        row0 = Number(used.Row) || 1;
        col0 = Number(used.Column) || 1;
        values = used.Value2;
      } catch (e) {}
      let rows = values == null ? [] : Array.isArray(values) ? (Array.isArray(values[0]) ? values : [values]) : [[values]];
      const truncated = rows.length > MAX_SHEET_ROWS;
      if (truncated) rows = rows.slice(0, MAX_SHEET_ROWS);
      let maxCol = -1;
      const lines = [];
      rows.forEach((row, r) => {
        const cells = (Array.isArray(row) ? row : [row]).slice(0, MAX_SHEET_COLS).map((v) => (v == null ? "" : String(v).replace(/[\t\r\n]+/g, " ")));
        let last = -1;
        cells.forEach((v, c) => { if (v !== "") last = c; });
        if (last < 0) return;
        if (last > maxCol) maxCol = last;
        lines.push({ r: row0 + r, cells });
      });
      let title = `工作表「${name}」` + (hidden ? "（隐藏）" : "");
      let text = "（空表）";
      if (lines.length) {
        const header = "行号\t" + Array.from({ length: maxCol + 1 }, (_v, i) => columnLetter(col0 - 1 + i)).join("\t");
        text = header + "\n" + lines.map((l) => `${l.r}\t${l.cells.slice(0, maxCol + 1).concat(Array(Math.max(0, maxCol + 1 - l.cells.length)).fill("")).join("\t")}`).join("\n");
      }
      if (truncated) text += `\n（行数超过 ${MAX_SHEET_ROWS}，后面的行未读取）`;
      sections.push({ title, text });
    });
    return sections.length ? sections : [{ title: "工作簿", text: "（没有可读取的工作表）" }];
  }
  function readPresentation(pres) {
    const sections = [];
    each(pres.Slides, (slide, i) => {
      const paras = [];
      each(slide.Shapes, (shape) => {
        try {
          if (shape.HasTextFrame) {
            const t = String(shape.TextFrame.TextRange.Text || "").replace(/\r/g, "\n").trim();
            if (t) paras.push(t);
          } else if (shape.HasTable) {
            const tbl = shape.Table;
            const rows = Number(tbl.Rows.Count) || 0;
            const cols = Number(tbl.Columns.Count) || 0;
            for (let r = 1; r <= rows; r += 1) {
              const cells = [];
              for (let c = 1; c <= cols; c += 1) {
                try { cells.push(String(tbl.Cell(r, c).Shape.TextFrame.TextRange.Text || "").replace(/\s+/g, " ").trim()); } catch (e) { cells.push(""); }
              }
              paras.push("| " + cells.join(" | ") + " |");
            }
          }
        } catch (e) {}
      });
      sections.push({ title: `第 ${i} 页`, text: paras.join("\n") || "（本页没有文字）" });
    });
    return sections.length ? sections : [{ title: "演示文稿", text: "（没有幻灯片）" }];
  }
  function assemble(sections) {
    let text = "";
    const index = [];
    sections.forEach((s) => {
      if (text) text += "\n\n";
      index.push({ title: s.title, offset: text.length });
      text += `## ${s.title}\n${s.text}`;
    });
    return { text, sections: index };
  }

  async function readOpenDocument(filePath) {
    const host = await currentHost();
    const app = await getApp();
    const want = normPath(filePath);
    let target = null;
    each(collectionFor(app, host), (doc) => { if (!target && normPath(docPath(doc)) === want) target = doc; });
    if (!target) return { ok: false, notOpen: true, error: "这份文档已不在 WPS 里打开" };
    const sections = host === "wps" ? readWordDoc(target) : host === "et" ? readWorkbook(target) : readPresentation(target);
    let saved = true;
    try { saved = target.Saved !== false; } catch (e) {}
    return Object.assign({ ok: true, source: "wps", kind: host === "wps" ? "word" : host === "et" ? "excel" : "ppt", saved }, assemble(sections));
  }

  // ===== 读取（给 AI 工具用）：能委托就委托，否则磁盘 =====
  async function listOpenDocs() {
    try { return await call("open-docs"); } catch (e) { return { ok: false, docs: [], online: {}, error: e.message }; }
  }
  function findIn(docs, filePath) {
    const want = normPath(filePath);
    return (docs || []).find((d) => d.path && normPath(d.path) === want) || null;
  }

  async function read({ path, offset, maxChars, section } = {}) {
    if (!path) throw new Error("需要文档路径");
    const open = await listOpenDocs();
    const hit = findIn(open.docs, path);
    let delegateError = "";
    if (hit && hit.host !== "pdf" && open.online?.[hit.host]) {
      try {
        const r = await call("delegate", { host: hit.host, path, offset, maxChars, section });
        if (r && r.text != null) return Object.assign(r, { name: baseName(path), path });
        delegateError = r?.error || "";
      } catch (e) { delegateError = e.message; }
    }
    try {
      const r = await call("read", { path, offset, maxChars, section });
      const notes = [r.note, delegateError ? `委托 WPS 读取失败（${delegateError}），已改为读取磁盘上保存的版本` : "",
        hit && hit.saved === false ? "这份文档在 WPS 里有未保存的修改，磁盘上读到的是上次保存的版本" : ""].filter(Boolean);
      return Object.assign(r, { name: baseName(path), path, note: notes.join("；") || undefined });
    } catch (e) {
      if (!hit && /老格式/.test(e.message)) throw new Error(e.message + "（可以在参考文档列表里点它在 WPS 中打开，打开后会自动改用 WPS 读取）");
      throw e;
    }
  }

  // ===== 选文件 / 在 WPS 打开 =====
  async function pick() {
    const r = await call("pick", {}, "POST");
    return Array.isArray(r.paths) ? r.paths : [];
  }

  // 让 WPS 打开文件，然后尽量切回原来的文档（跨组件打开时 WPS 通常会跳到新文档）
  async function openInWps(paths, { returnFocus = true } = {}) {
    const app = await getApp();
    const host = await currentHost();
    const back = returnFocus ? activeFor(app, host) : null;
    const r = await call("open", { paths });
    if (back) {
      [1500, 3500].forEach((ms) => setTimeout(() => {
        try { back.Activate(); } catch (e) {}
        try { back.ActiveWindow && back.ActiveWindow.Activate(); } catch (e) {}
      }, ms));
    }
    return r.results || [];
  }

  // ===== 后台：登记 + 等代读 =====
  const isDialogPage = /[?&]mode=/.test(String(global.location?.search || ""));
  let presenceTimer = null;
  let polling = false;
  let backoffUntil = 0;

  async function reportPresence() {
    if (Date.now() < backoffUntil) return;
    try {
      const { host, docs } = await collectOpenDocs();
      if (host === "unknown") return;
      await call("presence", { instanceId, host, docs });
    } catch (e) {
      // 老版本代理没有 /refs/*：一分钟后再试，别刷屏
      backoffUntil = Date.now() + (e.status === 404 ? 60000 : 10000);
    }
  }

  async function pollLoop() {
    if (polling) return;
    polling = true;
    while (polling) {
      if (Date.now() < backoffUntil) { await new Promise((r) => setTimeout(r, 5000)); continue; }
      const host = await currentHost();
      if (host === "unknown" || host === "pdf") { await new Promise((r) => setTimeout(r, 10000)); continue; }
      let job = null;
      try {
        job = (await call(`poll?host=${encodeURIComponent(host)}&instance=${encodeURIComponent(instanceId)}`)).call;
      } catch (e) {
        backoffUntil = Date.now() + (e.status === 404 ? 60000 : 5000);
        continue;
      }
      if (!job) continue;
      let result;
      try { result = job.op === "read" ? await readOpenDocument(job.path) : { ok: false, error: "未知请求：" + job.op }; } catch (e) { result = { ok: false, error: e?.message || String(e) }; }
      try { await call("result", { callId: job.callId, result }); } catch (e) {}
    }
  }

  function start() {
    if (isDialogPage || presenceTimer) return;
    presenceTimer = setInterval(reportPresence, PRESENCE_INTERVAL_MS);
    setTimeout(reportPresence, 1500);
    setTimeout(pollLoop, 2000);
  }
  // WPS 在组件启动时就加载插件的后台页（功能区按钮也是这时注册的），所以没点开面板也会登记和等活
  if (typeof global.setTimeout === "function" && global.fetch) start();

  global.WpsAiRefDocs = {
    list, add, remove, onChange,
    pick, openInWps, listOpenDocs, read,
    kindOf, baseName, normPath,
    _internal: { collectOpenDocs, readOpenDocument, readWorkbook, readWordDoc, readPresentation, assemble, instanceId }
  };
})(window);
