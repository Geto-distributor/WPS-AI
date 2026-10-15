const test = require("node:test");
const assert = require("node:assert");
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");

const skillZip = require("../tools/skill-zip.js");

function loadWindow() {
  const mem = new Map();
  const store = {
    getItem: (k) => (mem.has(k) ? mem.get(k) : null),
    setItem: (k, v) => mem.set(k, String(v)),
    removeItem: (k) => mem.delete(k)
  };
  const window = { WpsAiStore: store };
  const run = (file) => vm.runInThisContext(
    "(function(window, localStorage, console, fetch){ " + fs.readFileSync(path.join(__dirname, "..", file), "utf8") + "\n})"
  )(window, store, console, async () => { throw new Error("no fetch"); });
  run("js/skills.js");
  run("js/tools/registry.js");
  run("js/tools/common.js");
  return window;
}

const { makeZip } = require("./helpers/make-zip.js");

const SKILL_MD = "---\nname: 合同审查\ndescription: 审合同时用\n---\n# 步骤\n先看 references/checklist.md\n";

test("skill-zip: 解出文本文件，跳过二进制，目录条目忽略", () => {
  const zip = makeZip([
    { name: "contract/", data: "", store: true },
    { name: "contract/SKILL.md", data: SKILL_MD },
    { name: "contract/references/checklist.md", data: "# 清单\n- 违约责任", store: true },
    { name: "contract/assets/logo.png", data: Buffer.from([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 1]) }
  ]);
  const out = skillZip.extractTextFiles(zip);
  assert.deepStrictEqual(out.files.map((f) => f.path), ["contract/SKILL.md", "contract/references/checklist.md"]);
  assert.strictEqual(out.files[0].content, SKILL_MD);
  assert.strictEqual(out.files[1].content, "# 清单\n- 违约责任");
  assert.deepStrictEqual(out.skipped, [{ path: "contract/assets/logo.png", reason: "非文本文件" }]);
});

test("skill-zip: 上限生效（单文件过大 / 解压炸弹），坏文件直接报错", () => {
  const big = "a".repeat(5000);
  const out = skillZip.extractTextFiles(makeZip([{ name: "SKILL.md", data: big }]), { maxFileBytes: 1000 });
  assert.strictEqual(out.files.length, 0);
  assert.strictEqual(out.skipped[0].reason, "文件过大");
  assert.throws(() => skillZip.extractTextFiles(Buffer.from("not a zip at all, definitely not")), /不是有效的 zip/);
});

test("buildSkillFromFiles: 取最浅的 SKILL.md 为根，references 进 files，scripts / 非文本 / 隐藏文件跳过", () => {
  const S = loadWindow().WpsAiSkills;
  const built = S.buildSkillFromFiles([
    { path: "contract/SKILL.md", content: SKILL_MD },
    { path: "contract/references/checklist.md", content: "# 清单" },
    { path: "contract\\references\\clauses.csv", content: "a,b" },
    { path: "contract/scripts/run.py", content: "print(1)" },
    { path: "contract/assets/logo.png", content: "xx" },
    { path: "contract/.DS_Store", content: "xx" },
    { path: "__MACOSX/contract/._SKILL.md", content: "xx" }
  ]);
  assert.strictEqual(built.name, "合同审查");
  assert.strictEqual(built.description, "审合同时用");
  assert.match(built.content, /^# 步骤/);
  assert.deepStrictEqual(built.files.map((f) => f.path), ["references/checklist.md", "references/clauses.csv"]);
  assert.deepStrictEqual(built.skipped.map((s) => s.path).sort(), ["assets/logo.png", "scripts/run.py"]);
});

test("buildSkillFromFiles: 无 frontmatter 时用文件夹名；没有 SKILL.md 报错", () => {
  const S = loadWindow().WpsAiSkills;
  const built = S.buildSkillFromFiles([{ path: "my-skill/SKILL.md", content: "照着做就行" }]);
  assert.strictEqual(built.name, "my-skill");
  assert.deepStrictEqual(built.files, []);
  assert.throws(() => S.buildSkillFromFiles([{ path: "x/readme.md", content: "hi" }]), /没找到 SKILL\.md/);
});

test("addUser 存下 files；updateUser 改正文不丢 files；readFile 拒绝越界路径", () => {
  const S = loadWindow().WpsAiSkills;
  const saved = S.addUser({ name: "A", content: "正文", files: [
    { path: "references/a.md", content: "AAA" },
    { path: "../secret.md", content: "x" },
    { path: "references/empty.md", content: "  " }
  ] });
  assert.deepStrictEqual(S.listFiles(S.get(saved.id)), ["references/a.md"]);
  S.updateUser(saved.id, { content: "新正文" });
  const again = S.get(saved.id);
  assert.strictEqual(again.content, "新正文");
  assert.strictEqual(S.readFile(again, "references/a.md"), "AAA");
  assert.strictEqual(S.readFile(again, "References/A.md"), "AAA");
  assert.strictEqual(S.readFile(again, "references/../references/a.md"), null);
  assert.strictEqual(S.readFile(again, "nope.md"), null);
  assert.strictEqual(S.addUser({ name: "B", content: "x" }).files, undefined, "单文件技能不带空 files 字段");
});

test("use_skill 返回 files 清单；read_skill_file 按需读取，读不到时报可读清单", async () => {
  const w = loadWindow();
  const S = w.WpsAiSkills;
  const R = w.WpsAiToolRegistry;
  const multi = S.addUser(S.buildSkillFromFiles([
    { path: "contract/SKILL.md", content: SKILL_MD },
    { path: "contract/references/checklist.md", content: "# 清单" }
  ]));
  const single = S.addUser({ name: "单文件", content: "正文" });
  S.setEnabled(multi.id, true);
  S.setEnabled(single.id, true);
  const call = (name, args) => R.getDefinition(name).handler(args);

  const used = await call("use_skill", { name: "合同审查" });
  assert.deepStrictEqual(used.files, ["references/checklist.md"]);
  assert.match(used.filesHint, /read_skill_file/);
  assert.ok(!used.content.includes("# 清单"), "参考文件不随 use_skill 一次性加载");

  const plain = await call("use_skill", { name: "单文件" });
  assert.strictEqual(plain.files, undefined);
  assert.strictEqual(plain.filesHint, undefined);

  const read = await call("read_skill_file", { name: "合同审查", path: "references/checklist.md" });
  assert.strictEqual(read.content, "# 清单");
  await assert.rejects(call("read_skill_file", { name: "合同审查", path: "references/nope.md" }), /可读文件：references\/checklist\.md/);
  await assert.rejects(call("read_skill_file", { name: "不存在", path: "a.md" }), /没有启用的技能/);
});

test("buildFileTree: 目录在前、文件在后，嵌套目录逐层展开", () => {
  const S = loadWindow().WpsAiSkills;
  const tree = S.buildFileTree(["references/b.md", "notes.txt", "references/sub/data.csv", "references/a.md", "../bad.md"]);
  const flat = (nodes, depth = 0) => nodes.flatMap((n) => [`${"  ".repeat(depth)}${n.type === "dir" ? n.name + "/" : n.name}`, ...(n.children ? flat(n.children, depth + 1) : [])]);
  assert.deepStrictEqual(flat(tree), ["references/", "  sub/", "    data.csv", "  a.md", "  b.md", "notes.txt"]);
  assert.strictEqual(tree[0].children[0].children[0].path, "references/sub/data.csv");
});

test("SKILL.md 编辑往返：buildSkillMarkdown ⇄ saveSkillMarkdown 改名 / 改简介 / 改正文，不丢 files", () => {
  const S = loadWindow().WpsAiSkills;
  const saved = S.addUser({ name: "旧名", description: "旧简介", content: "旧正文", files: [{ path: "references/a.md", content: "A" }] });
  assert.strictEqual(S.buildSkillMarkdown(saved), "---\nname: 旧名\ndescription: 旧简介\n---\n旧正文\n");
  S.saveSkillMarkdown(saved.id, "---\nname: 新名\ndescription: 新简介\n---\n# 新正文\n");
  const cur = S.get(saved.id);
  assert.deepStrictEqual([cur.name, cur.description, cur.content], ["新名", "新简介", "# 新正文"]);
  assert.deepStrictEqual(S.listFiles(cur), ["references/a.md"]);
  assert.throws(() => S.saveSkillMarkdown(saved.id, "---\nname: x\n---\n   "), /正文不能为空/);
  assert.strictEqual(S.get(saved.id).name, "新名", "保存失败不改动原技能");
});

test("writeFile / deleteFile：新建、覆盖、删除参考文件，非法路径和内置技能拒绝", () => {
  const S = loadWindow().WpsAiSkills;
  const saved = S.addUser({ name: "A", content: "正文" });
  S.writeFile(saved.id, "references\\new.md", "v1");
  S.writeFile(saved.id, "references/sub/data.csv", "a,b");
  S.writeFile(saved.id, "References/New.md", "v2"); // 同一个文件（忽略大小写）→ 覆盖，保留原路径
  let cur = S.get(saved.id);
  assert.deepStrictEqual(S.listFiles(cur), ["references/new.md", "references/sub/data.csv"]);
  assert.strictEqual(S.readFile(cur, "references/new.md"), "v2");

  assert.throws(() => S.writeFile(saved.id, "scripts/run.py", "print(1)"), /scripts/);
  assert.throws(() => S.writeFile(saved.id, "SKILL.md", "x"), /SKILL\.md/);
  assert.throws(() => S.writeFile(saved.id, "../x.md", "x"), /路径无效/);
  assert.throws(() => S.writeFile(saved.id, "logo.png", "x"), /只支持文本文件/);
  assert.throws(() => S.writeFile(saved.id, "references/empty.md", "  "), /不能为空/);
  assert.throws(() => S.writeFile("not-a-user-skill", "references/x.md", "x"), /只能编辑自定义技能/);

  assert.strictEqual(S.deleteFile(saved.id, "references/nope.md"), false);
  assert.strictEqual(S.deleteFile(saved.id, "references/new.md"), true);
  assert.strictEqual(S.deleteFile(saved.id, "references/sub/data.csv"), true);
  cur = S.get(saved.id);
  assert.strictEqual(cur.files, undefined, "删到一个不剩时去掉 files 字段");
  assert.strictEqual(cur.content, "正文");
});
