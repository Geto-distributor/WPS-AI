const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const { makeZip } = require("./helpers/make-zip.js");
const { extractDocument, createExtractor, sliceDocument, kindOf } = require("../tools/doc-extract.js");
const { createRefRelay } = require("../tools/ref-relay.js");

function tmpFile(name, buf) {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "lingxi-refdoc-"));
  const p = path.join(dir, name);
  fs.writeFileSync(p, buf);
  return p;
}

const DOCX_XML = `<?xml version="1.0"?><w:document xmlns:w="w"><w:body>
<w:p><w:r><w:t>第一条 &amp; 定义</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">甲方：</w:t></w:r><w:r><w:tab/><w:t>志特</w:t></w:r></w:p>
<w:p/>
<w:tbl><w:tr><w:tc><w:p><w:r><w:t>品名</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>数量</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:p><w:r><w:t>铝模板</w:t></w:r></w:p><w:p><w:r><w:t>标准件</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>1200</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p><w:r><w:t>第二条</w:t></w:r><w:r><w:delText>已删掉的字</w:delText></w:r><w:r><w:br/><w:t>换行后</w:t></w:r></w:p>
</w:body></w:document>`;

test("docx：段落 / 制表符 / 换行 / 表格 / 修订删除的文字", async () => {
  const p = tmpFile("合同.docx", makeZip([{ name: "word/document.xml", data: DOCX_XML }]));
  const doc = await extractDocument(p);
  assert.strictEqual(doc.kind, "word");
  assert.strictEqual(doc.text, [
    "## 正文",
    "第一条 & 定义",
    "甲方：\t志特",
    "",
    "| 品名 | 数量 |",
    "| 铝模板 / 标准件 | 1200 |",
    "",
    "第二条",
    "换行后"
  ].join("\n"));
});

test("xlsx：共享字符串 / 内联字符串 / 布尔 / 空单元格占位 / 隐藏工作表 / 按 rels 找表", async () => {
  const p = tmpFile("材料清单.xlsx", makeZip([
    { name: "xl/workbook.xml", data: `<workbook><sheets><sheet name="清单" sheetId="1" r:id="rId1"/><sheet name="备注" sheetId="2" state="hidden" r:id="rId2"/><sheet name="空表" sheetId="3" r:id="rId3"/></sheets></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", data: `<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Target="worksheets/sheet9.xml"/><Relationship Id="rId3" Target="/xl/worksheets/sheet3.xml"/></Relationships>` },
    { name: "xl/sharedStrings.xml", data: `<sst><si><t>品名</t></si><si><r><t>铝</t></r><r><t>模板</t></r><rPh><t>lv</t></rPh></si></sst>` },
    { name: "xl/worksheets/sheet1.xml", data: `<worksheet><sheetData>
      <row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="inlineStr"><is><t>单价 &lt;元&gt;</t></is></c></row>
      <row r="3"><c r="A3" t="s"><v>1</v></c><c r="B3"><v>1200</v></c><c r="C3" t="b"><v>1</v></c></row>
    </sheetData></worksheet>` },
    { name: "xl/worksheets/sheet9.xml", data: `<worksheet><sheetData><row r="2"><c r="B2" t="str"><v>内部</v></c></row></sheetData></worksheet>` },
    { name: "xl/worksheets/sheet3.xml", data: `<worksheet><sheetData/></worksheet>` }
  ]));
  const doc = await extractDocument(p);
  assert.strictEqual(doc.kind, "excel");
  assert.strictEqual(doc.text, [
    "## 工作表「清单」",
    "行号\tA\tB\tC",
    "1\t品名\t\t单价 <元>",
    "3\t铝模板\t1200\tTRUE",
    "",
    "## 工作表「备注」（隐藏）",
    "行号\tA\tB",
    "2\t\t内部",
    "",
    "## 工作表「空表」",
    "（空表）"
  ].join("\n"));
  assert.deepStrictEqual(doc.sections.map((s) => s.title), ["工作表「清单」", "工作表「备注」（隐藏）", "工作表「空表」"]);
  assert.match(doc.note, /序列号/);
});

test("pptx：按 presentation.xml 的顺序读页，不按文件名顺序", async () => {
  const p = tmpFile("汇报.pptx", makeZip([
    { name: "ppt/presentation.xml", data: `<p:presentation><p:sldIdLst><p:sldId id="256" r:id="rId7"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst></p:presentation>` },
    { name: "ppt/_rels/presentation.xml.rels", data: `<Relationships><Relationship Id="rId3" Target="slides/slide1.xml"/><Relationship Id="rId7" Target="slides/slide2.xml"/></Relationships>` },
    { name: "ppt/slides/slide1.xml", data: `<p:sld><a:p><a:r><a:t>后一页</a:t></a:r></a:p></p:sld>` },
    { name: "ppt/slides/slide2.xml", data: `<p:sld><a:p><a:r><a:t>封面</a:t></a:r><a:br/><a:r><a:t>副标题</a:t></a:r></a:p><a:p></a:p></p:sld>` }
  ]));
  const doc = await extractDocument(p);
  assert.strictEqual(doc.text, "## 第 1 页\n封面\n副标题\n\n## 第 2 页\n后一页");
});

test("pdf 走注入的解析；扫描件给出提示；老格式和未知类型报清楚的错", async () => {
  const p = tmpFile("资料.pdf", Buffer.from("%PDF-1.4"));
  const doc = await extractDocument(p, { extractPdfText: async () => ({ pages: [{ page: 1, text: "甲" }, { page: 2, text: "" }], hasText: false }) });
  assert.strictEqual(doc.text, "## 第 1 页\n甲\n\n## 第 2 页\n（本页没有文字）");
  assert.match(doc.note, /扫描件/);
  await assert.rejects(extractDocument(tmpFile("旧.doc", Buffer.from("x"))), /老格式/);
  await assert.rejects(extractDocument(tmpFile("图.png", Buffer.from("x"))), /不支持的文件类型/);
  await assert.rejects(extractDocument(tmpFile("坏.docx", Buffer.from("not a zip at all, definitely not"))), /zip/);
  assert.deepStrictEqual(["a.DOCX", "b.xlsm", "c.pptx", "d.pdf", "e.csv", "f.wps", "g.zip"].map(kindOf), ["word", "excel", "ppt", "pdf", "text", "legacy", "unknown"]);
});

test("分段读取：offset / maxChars / 按节名跳转；解析结果按 mtime 缓存", async () => {
  const p = tmpFile("长文.txt", "a".repeat(1500) + "\n## 不是节\n" + "b".repeat(1500));
  let calls = 0;
  const ex = createExtractor({});
  const orig = fs.readFileSync;
  fs.readFileSync = function (...args) { if (args[0] === p) calls += 1; return orig.apply(this, args); };
  try {
    const doc = await ex.get(p);
    await ex.get(p);
    assert.strictEqual(calls, 1, "未改动的文件不重复解析");
    const first = sliceDocument(doc, { maxChars: 1000 });
    assert.strictEqual(first.offset, 0);
    assert.strictEqual(first.text.length, 1000);
    assert.strictEqual(first.nextOffset, 1000);
    const last = sliceDocument(doc, { offset: first.total - 10, maxChars: 1000 });
    assert.strictEqual(last.nextOffset, null);
    assert.strictEqual(sliceDocument(doc, { maxChars: 1 }).text.length, 500, "单次至少 500 字");
  } finally { fs.readFileSync = orig; }

  const x = tmpFile("两表.xlsx", makeZip([
    { name: "xl/workbook.xml", data: `<workbook><sheets><sheet name="甲" r:id="r1"/><sheet name="乙" r:id="r2"/></sheets></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", data: `<Relationships><Relationship Id="r1" Target="worksheets/s1.xml"/><Relationship Id="r2" Target="worksheets/s2.xml"/></Relationships>` },
    { name: "xl/worksheets/s1.xml", data: `<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>一</v></c></row></sheetData></worksheet>` },
    { name: "xl/worksheets/s2.xml", data: `<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>二</v></c></row></sheetData></worksheet>` }
  ]));
  const doc2 = await createExtractor({}).get(x);
  assert.ok(sliceDocument(doc2, { section: "乙" }).text.startsWith("## 工作表「乙」"));
});

test("ref-relay：登记汇总去重、过期清理；委托代读按组件分发、超时报错", async () => {
  let t = 1000;
  const relay = createRefRelay({ now: () => t, staleMs: 5000, callTimeoutMs: 200, pollTimeoutMs: 100 });
  relay.report({ instanceId: "w1", host: "wps", docs: [{ name: "合同.docx", path: "C:\\docs\\合同.docx", active: true }] });
  relay.report({ instanceId: "w2", host: "wps", docs: [{ name: "合同.docx", path: "c:/docs/合同.docx" }] }); // 同组件的另一页再报一次
  relay.report({ instanceId: "e1", host: "et", docs: [{ name: "清单.xlsx", path: "C:\\docs\\清单.xlsx", saved: false }] });
  assert.deepStrictEqual(relay.listOpenDocs().map((d) => `${d.host}:${d.name}:${d.saved}`), ["wps:合同.docx:true", "et:清单.xlsx:false"]);
  assert.strictEqual(relay.findOpenDoc("c:\\DOCS\\清单.xlsx").host, "et", "路径比对忽略大小写和分隔符");

  assert.strictEqual(relay.workerOnline("et"), false, "只登记过、没来取过活的不算能代读");
  const polled = relay.poll("et", "e1");
  assert.strictEqual(relay.workerOnline("et"), true);
  const answer = relay.request("et", { op: "read", path: "C:\\docs\\清单.xlsx" });
  const call = await polled;
  assert.strictEqual(call.op, "read");
  assert.strictEqual(relay.resolveCall(call.callId, { source: "wps", text: "实时内容" }), true);
  assert.deepStrictEqual(await answer, { source: "wps", text: "实时内容" });
  assert.strictEqual(relay.resolveCall(call.callId, {}), false, "同一个 callId 只认一次");

  assert.strictEqual(await relay.poll("wpp", "p1"), null, "没活时长轮询到点返回 null");
  await assert.rejects(relay.request("pdf", { op: "read" }), /委托读取超时/);

  relay.report({ instanceId: "e2", host: "et", docs: [], page: "background", url: "/et/index.html" });
  const inst = relay.listInstances().find((i) => i.instanceId === "e2");
  assert.deepStrictEqual([inst.host, inst.page, inst.url, inst.fresh], ["et", "background", "/et/index.html", true]);

  t += 6000;
  assert.deepStrictEqual(relay.listOpenDocs(), [], "超过 staleMs 没再登记的实例被清掉");
  assert.strictEqual(relay.workerOnline("et"), false);
});

test("随包的 pdfjs（tools/vendor）不靠 node_modules 就能解析中文 PDF；各平台打包都不会排除它", async () => {
  const vendor = path.join(__dirname, "..", "tools", "vendor", "pdfjs-dist");
  const pdfjs = require(path.join(vendor, "legacy", "build", "pdf.js"));
  const data = new Uint8Array(fs.readFileSync(path.join(__dirname, "fixtures", "ref-doc-cn.pdf")));
  const doc = await pdfjs.getDocument({ data, useSystemFonts: true, isEvalSupported: false, cMapUrl: path.join(vendor, "cmaps") + path.sep, cMapPacked: true }).promise;
  try {
    const text = (await (await doc.getPage(1)).getTextContent()).items.map((i) => i.str).join("");
    assert.match(text, /合同评审测试/);
    assert.match(text, /付款条件：预付款/);
  } finally { await doc.destroy(); }

  const root = path.join(__dirname, "..", "..");
  const variants = fs.readFileSync(path.join(root, "plugin", "tools", "build-variants.js"), "utf8");
  const skip = (variants.match(/const SKIP = new Set\(\[([\s\S]*?)\]\)/) || [])[1] || "";
  assert.doesNotMatch(skip, /"vendor"|"tools"/, "build-variants 不能跳过 tools/vendor");
  const iss = fs.readFileSync(path.join(root, "installer", "lingxi-ai.iss"), "utf8");
  assert.doesNotMatch((iss.match(/Excludes: "([^"]*)"/) || [])[1] || "", /vendor/, "Windows 安装包不能排除 vendor");
  for (const f of ["installer-mac/build-dmg.sh", "installer-linux/build.sh"]) {
    assert.doesNotMatch(fs.readFileSync(path.join(root, f), "utf8"), /--exclude='[^']*vendor/, `${f} 不能排除 vendor`);
  }
});
