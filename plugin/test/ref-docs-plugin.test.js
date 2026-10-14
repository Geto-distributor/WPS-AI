const test = require("node:test");
const assert = require("node:assert");
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");

const read = (f) => fs.readFileSync(path.join(__dirname, "..", f), "utf8");
// 模块跑在独立的 vm 上下文里，返回的数组 / 对象原型和这里不同，比较前先转成普通 JSON
const plain = (v) => JSON.parse(JSON.stringify(v));

// 假 WPS 应用对象：集合用 Count + Item(i)（1 起），和 JSAPI 一致
const coll = (items) => ({ Count: items.length, Item: (i) => items[i - 1] });

function load({ host = "wps", app = {}, routes = {}, conversationId = "c1" } = {}) {
  const mem = new Map();
  const store = { getItem: (k) => (mem.has(k) ? mem.get(k) : null), setItem: (k, v) => mem.set(k, String(v)), removeItem: (k) => mem.delete(k) };
  const calls = [];
  const fetchImpl = async (url, init) => {
    const u = new URL(url);
    const route = u.pathname.replace("/refs/", "");
    const body = init && init.body ? JSON.parse(init.body) : undefined;
    calls.push({ route, body });
    const handler = routes[route];
    const out = handler ? await handler(body, u) : { status: 404, json: { ok: false, error: "no route" } };
    const status = out.status || 200;
    return { ok: status === 200, status, json: async () => out.json || out };
  };
  const window = {
    WpsAiStore: store,
    WpsAiRuntime: { proxyBase: () => "http://127.0.0.1:3890" },
    WpsAiConversations: { getCurrent: () => ({ id: conversationId }) },
    WpsAiDocument: { getApplication: async () => app, getHostInfo: async () => ({ host }) },
    WpsAiToolRegistry: null,
    location: { search: "?mode=settings" } // 弹窗页：不启动登记 / 轮询循环，测试只测函数
  };
  const run = (file) => vm.runInNewContext(read(file), { window, fetch: fetchImpl, console, setTimeout, clearTimeout, setInterval, clearInterval, URL });
  run("js/ref-docs.js");
  // 工具注册表 + 参考文档工具
  const reg = new Map();
  window.WpsAiToolRegistry = { registerTool: (d) => reg.set(d.name, d) };
  run("js/tools/refs.js");
  return { R: window.WpsAiRefDocs, tools: reg, calls, store };
}

test("参考文档清单按对话保存，去重，移除", () => {
  const a = load({ conversationId: "c1" });
  const added = a.R.add(["C:\\资料\\清单.xlsx", "C:/资料/清单.xlsx", "D:\\合同\\附件.pdf", "  "]);
  assert.deepStrictEqual(plain(added.map((r) => `${r.name}:${r.kind}`)), ["清单.xlsx:excel", "附件.pdf:pdf"], "同一路径不同写法只算一份");
  assert.strictEqual(a.R.list().length, 2);
  a.R.remove("c:\\资料\\清单.xlsx");
  assert.deepStrictEqual(plain(a.R.list().map((r) => r.name)), ["附件.pdf"]);
  const raw = JSON.parse(a.store.getItem("lingxi_ref_docs_v1"));
  assert.deepStrictEqual(Object.keys(raw), ["c1"], "按对话 id 分开存");
});

test("登记：只报本组件开着的文档，标出当前文档和未保存的", async () => {
  const docs = [
    { Name: "合同.docx", FullName: "C:\\a\\合同.docx", Saved: true },
    { Name: "草稿.docx", FullName: "C:\\a\\草稿.docx", Saved: false }
  ];
  const { R } = load({ host: "wps", app: { Documents: coll(docs), ActiveDocument: docs[0] } });
  const r = await R._internal.collectOpenDocs();
  assert.deepStrictEqual(plain(r), { host: "wps", docs: [
    { name: "合同.docx", path: "C:\\a\\合同.docx", saved: true, active: true },
    { name: "草稿.docx", path: "C:\\a\\草稿.docx", saved: false, active: false }
  ] });
});

test("代读：表格按 UsedRange 输出带行号列号的文本；Word 表格标记转成制表符；按路径找文档", async () => {
  const sheet = (name, row, col, values, visible = -1) => ({ Name: name, Visible: visible, UsedRange: { Row: row, Column: col, Value2: values } });
  const wb = { FullName: "C:\\a\\清单.xlsx", Saved: false, Worksheets: coll([
    sheet("清单", 2, 2, [["品名", "数量"], [null, null], ["铝模板", 1200]]),
    sheet("备注", 1, 1, "只有一格", 0)
  ]) };
  const et = load({ host: "et", app: { Workbooks: coll([wb]) } });
  const r = await et.R._internal.readOpenDocument("c:/A/清单.xlsx");
  assert.strictEqual(r.ok, true);
  assert.strictEqual(r.saved, false);
  assert.strictEqual(r.text, "## 工作表「清单」\n行号\tB\tC\n2\t品名\t数量\n4\t铝模板\t1200\n\n## 工作表「备注」（隐藏）\n行号\tA\n1\t只有一格");
  assert.deepStrictEqual(plain(r.sections.map((s) => s.title)), ["工作表「清单」", "工作表「备注」（隐藏）"]);
  assert.strictEqual((await et.R._internal.readOpenDocument("C:\\a\\别的.xlsx")).notOpen, true);

  // 第二行第二格是空的：「铝模板」后面紧跟空单元格的结束符和行尾符，光看文本分不出来，要靠行列数
  const tableText = "品名\r\x07数量\r\x07\r\x07铝模板\r标准件\r\x07\r\x07\r\x07";
  const doc = {
    FullName: "C:\\a\\合同.docx",
    Content: { Text: "第一条\r" + tableText + "第二条\r" },
    Tables: coll([{ Range: { Text: tableText }, Rows: { Count: 2 }, Columns: { Count: 2 } }])
  };
  const wps = load({ host: "wps", app: { Documents: coll([doc]) } });
  const w = await wps.R._internal.readOpenDocument("C:\\a\\合同.docx");
  assert.strictEqual(w.text, "## 正文\n第一条\n\n| 品名 | 数量 |\n| 铝模板 / 标准件 |  |\n\n第二条");
});

test("读取路由：开着且对应组件在线 → 委托；委托失败 / 没开 → 读磁盘，并说明原因", async () => {
  const xlsx = "C:\\a\\清单.xlsx";
  let delegateFails = false;
  const routes = {
    "open-docs": () => ({ ok: true, docs: [{ host: "et", name: "清单.xlsx", path: xlsx, saved: false }], online: { et: true } }),
    delegate: (b) => (delegateFails ? { status: 400, json: { ok: false, error: "超时" } } : { ok: true, source: "wps", text: "实时内容", offset: 0, nextOffset: null, total: 4, sections: [] }),
    read: (b) => ({ ok: true, source: "disk", text: "磁盘内容:" + b.path, offset: 0, nextOffset: null, total: 9, sections: [] })
  };
  const { R, calls } = load({ routes });
  const live = await R.read({ path: xlsx, maxChars: 5000 });
  assert.strictEqual(live.source, "wps");
  assert.strictEqual(live.text, "实时内容");
  assert.deepStrictEqual(plain(calls.find((c) => c.route === "delegate").body), { host: "et", path: xlsx, maxChars: 5000 });

  delegateFails = true;
  const fallback = await R.read({ path: xlsx });
  assert.strictEqual(fallback.source, "disk");
  assert.match(fallback.note, /委托 WPS 读取失败（超时）/);
  assert.match(fallback.note, /未保存的修改/);

  const notOpen = await R.read({ path: "D:\\x\\附件.pdf" });
  assert.strictEqual(notOpen.source, "disk");
  assert.strictEqual(calls.filter((c) => c.route === "delegate").length, 2, "没打开的不去委托");
});

test("AI 工具：只能读已添加或已打开的文档；列表区分已添加 / 其它打开着的（排除当前文档）", async () => {
  const routes = {
    "open-docs": () => ({ ok: true, online: { wps: true, et: true }, docs: [
      { host: "wps", name: "合同.docx", path: "C:\\a\\合同.docx", active: true },
      { host: "et", name: "清单.xlsx", path: "C:\\a\\清单.xlsx", saved: false },
      { host: "pdf", name: "资料.pdf", path: "C:\\a\\资料.pdf" }
    ] }),
    delegate: () => ({ ok: true, source: "wps", text: "x", offset: 0, nextOffset: null, total: 1, sections: [] }),
    read: () => ({ ok: true, source: "disk", text: "y", offset: 0, nextOffset: null, total: 1, sections: [] })
  };
  const { R, tools } = load({ routes });
  R.add(["C:\\a\\清单.xlsx", "D:\\背景\\说明.docx"]);
  const listed = await tools.get("list_reference_docs").handler({});
  assert.deepStrictEqual(plain(listed.added.map((a) => `${a.name}|${a.readVia}`)), ["清单.xlsx|已在 WPS 打开（读实时内容）", "说明.docx|读取磁盘文件"]);
  assert.deepStrictEqual(plain(listed.otherOpenDocs.map((d) => d.name)), ["资料.pdf"], "当前文档和已添加的不再重复列出");

  const read = tools.get("read_reference_doc").handler;
  assert.strictEqual((await read({ path: "C:\\a\\资料.pdf" })).readVia, "读取磁盘文件", "PDF 一律读磁盘");
  assert.strictEqual((await read({ path: "c:/a/清单.xlsx" })).readVia, "已在 WPS 打开（读实时内容）");
  await assert.rejects(read({ path: "C:\\Windows\\win.ini" }), /只能读取已添加的参考文档或 WPS 里已打开的文档/);
  await assert.rejects(read({}), /需要 path/);
});
