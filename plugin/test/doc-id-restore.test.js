const test = require("node:test");
const assert = require("node:assert");
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");

// 假的 CustomDocumentProperties：Item(name) 不存在时抛错，和 WPS 一致
function makeDoc(initialId) {
  const props = new Map(initialId ? [["LingxiDocId", initialId]] : []);
  const coll = {
    Item(name) { if (!props.has(name)) throw new Error("not found"); return { Value: props.get(name) }; },
    Add(name, _link, _type, value) { props.set(name, value); }
  };
  return { FullName: "C:\\a\\合同.docx", CustomDocumentProperties: coll, _props: props };
}

function loadBackup(doc) {
  const window = { Application: { ActiveDocument: doc }, crypto: { randomUUID: () => "new-uuid" } };
  vm.runInNewContext(fs.readFileSync(path.join(__dirname, "..", "js", "backup.js"), "utf8"), { window, console, setTimeout, clearTimeout });
  return window.WpsAiBackup;
}

test("同一个文件被重新载入、内存里的文档 id 丢了：writeDocId 能把原来的 id 补回去", () => {
  const reloaded = makeDoc(null);
  const B = loadBackup(reloaded);
  assert.strictEqual(B.readDocId(), null);
  assert.strictEqual(B.writeDocId("old-uuid"), "old-uuid");
  assert.strictEqual(B.readDocId(), "old-uuid");
  assert.strictEqual(B.writeDocId("old-uuid"), "old-uuid", "已经是这个 id：不重复写");
  assert.strictEqual(B.writeDocId("another"), null, "文档已有别的 id：不覆盖");
  assert.strictEqual(B.readDocId(), "old-uuid");
  assert.strictEqual(B.writeDocId(""), null);
});

test("ensureDocId 行为不变：没有就新建，有就沿用", () => {
  const B = loadBackup(makeDoc(null));
  assert.strictEqual(B.ensureDocId(), "new-uuid");
  assert.strictEqual(B.ensureDocId(), "new-uuid");
  assert.strictEqual(loadBackup(makeDoc("kept")).ensureDocId(), "kept");
});
