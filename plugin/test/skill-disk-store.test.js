const test = require("node:test");
const assert = require("node:assert");
const vm = require("node:vm");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const { createSkillStore } = require("../tools/skill-store.js");

function tmpRoot() {
  return fs.mkdtempSync(path.join(os.tmpdir(), "lingxi-skills-test-"));
}
const read = (root, ...p) => fs.readFileSync(path.join(root, ...p), "utf8");
const exists = (root, ...p) => fs.existsSync(path.join(root, ...p));

// 前端 skills.js + 一个把 /skills/* 直接转给磁盘模块的假 fetch（等价于代理的那几条路由）
function loadSkills(root, { legacy, failRoutes } = {}) {
  const store = createSkillStore(root);
  const mem = new Map();
  if (legacy) mem.set("lingxi_skills_user_v1", JSON.stringify({ entries: legacy.entries }));
  if (legacy?.enabled) mem.set("lingxi_skills_enabled_v1", JSON.stringify({ ids: legacy.enabled }));
  const kv = {
    getItem: (k) => (mem.has(k) ? mem.get(k) : null),
    setItem: (k, v) => mem.set(k, String(v)),
    removeItem: (k) => mem.delete(k)
  };
  const calls = [];
  const state = { down: false };
  const fetchImpl = async (url, init) => {
    if (state.down) throw new Error("ECONNREFUSED");
    const route = new URL(url).pathname.replace("/skills/", "");
    const body = init?.body ? JSON.parse(init.body) : undefined;
    calls.push(route);
    if (route === "list" && state.listGate) await state.listGate; // 让测试能把 list 卡在半路
    const respond = (status, json) => ({ ok: status === 200, status, json: async () => json });
    try {
      if (failRoutes?.includes(route)) throw new Error("磁盘已满");
      if (route === "list") return respond(200, Object.assign({ ok: true }, store.list()));
      if (route === "put") store.put(body.dir, body.files);
      else if (route === "write") store.writeFile(body.dir, body.path, body.content);
      else if (route === "delete-file") store.deleteFile(body.dir, body.path);
      else if (route === "rename") store.rename(body.from, body.to);
      else if (route === "delete") store.remove(body.dir);
      else return respond(404, { ok: false, error: "unknown" });
      return respond(200, { ok: true });
    } catch (e) { return respond(400, { ok: false, error: e.message }); }
  };
  const window = { WpsAiStore: kv };
  vm.runInThisContext(
    "(function(window, localStorage, console, fetch){ " + fs.readFileSync(path.join(__dirname, "..", "js", "skills.js"), "utf8") + "\n})"
  )(window, kv, { warn() {}, error() {}, log() {} }, fetchImpl);
  return { S: window.WpsAiSkills, kv, calls, state, store };
}

test("skill-store: put / list / write / delete-file / rename / remove 走一遍", () => {
  const root = tmpRoot();
  const store = createSkillStore(root);
  assert.deepStrictEqual(store.list().skills, [], "根目录不存在时返回空而不是报错");
  store.put("合同审查", [
    { path: "SKILL.md", content: "---\nname: 合同审查\n---\n正文" },
    { path: "references/a.md", content: "A" }
  ]);
  fs.writeFileSync(path.join(root, "合同审查", "logo.png"), Buffer.from([0x89, 0x50, 0, 0, 1]));
  fs.mkdirSync(path.join(root, "不是技能"));
  fs.writeFileSync(path.join(root, "不是技能", "readme.md"), "x");

  const listed = store.list().skills;
  assert.deepStrictEqual(listed.map((s) => s.dir), ["合同审查"], "没有 SKILL.md 的目录不算技能");
  assert.deepStrictEqual(listed[0].files.map((f) => f.path), ["references/a.md", "SKILL.md"], "二进制文件不读");

  store.writeFile("合同审查", "references/sub/b.csv", "1,2");
  assert.strictEqual(read(root, "合同审查", "references", "sub", "b.csv"), "1,2");
  assert.strictEqual(store.deleteFile("合同审查", "references/sub/b.csv"), true);
  assert.ok(!exists(root, "合同审查", "references", "sub"), "删空的子目录一并清掉");
  assert.ok(exists(root, "合同审查", "references", "a.md"));
  assert.throws(() => store.deleteFile("合同审查", "SKILL.md"), /不能单独删除/);

  store.rename("合同审查", "合同审查 v2");
  assert.ok(exists(root, "合同审查 v2", "logo.png"), "改名是真 rename，不受管的文件也跟着走");
  assert.throws(() => store.put("合同审查 v2", [{ path: "SKILL.md", content: "x" }]), /已存在/);

  assert.strictEqual(store.remove("合同审查 v2"), true);
  assert.ok(!exists(root, "合同审查 v2"));
  const trashed = fs.readdirSync(path.join(root, ".trash"));
  assert.strictEqual(trashed.length, 1, "删除是挪进 .trash，不是直接抹掉");
  assert.ok(exists(root, ".trash", trashed[0], "logo.png"));
  assert.deepStrictEqual(store.list().skills, [], ".trash 不会被当成技能列出来");
});

test("skill-store: 越界 / 非法路径一律拒绝，磁盘上不留痕迹", () => {
  const root = tmpRoot();
  const store = createSkillStore(root);
  store.put("ok", [{ path: "SKILL.md", content: "x" }]);
  for (const dir of ["..", "../evil", "a/b", ".hidden", "", "con:1", "x "]) {
    assert.throws(() => store.writeFile(dir, "a.md", "x"), /目录名无效/, `dir=${JSON.stringify(dir)}`);
  }
  for (const rel of ["../a.md", "a/../../b.md", "/etc/passwd", ".env", "a//b.md", "refs/.secret"]) {
    assert.throws(() => store.writeFile("ok", rel, "x"), /路径无效/, `path=${rel}`);
  }
  assert.throws(() => store.put("bad", [{ path: "SKILL.md", content: "x" }, { path: "../x.md", content: "y" }]), /路径无效/);
  assert.ok(!exists(root, "bad"), "校验不过时整包都不写");
  assert.throws(() => store.put("nomain", [{ path: "a.md", content: "x" }]), /缺少 SKILL\.md/);
  assert.deepStrictEqual(fs.readdirSync(root), ["ok"]);
  assert.deepStrictEqual(fs.readdirSync(path.dirname(root)).filter((n) => n === "evil"), []);
});

test("前端磁盘模式：新建 / 改文件 / 删文件 / 删技能都落到 <技能名>/ 目录", async () => {
  const root = tmpRoot();
  const { S } = loadSkills(root);
  assert.strictEqual(await S.init(), "disk");

  const saved = S.addUser({ name: "合同审查", description: "审合同时用", content: "# 步骤", files: [{ path: "references/a.md", content: "A" }] });
  assert.strictEqual(saved.id, "user:合同审查");
  await S.whenSaved();
  assert.strictEqual(read(root, "合同审查", "SKILL.md"), "---\nname: 合同审查\ndescription: 审合同时用\n---\n# 步骤\n");
  assert.strictEqual(read(root, "合同审查", "references", "a.md"), "A");

  S.writeFile(saved.id, "references/a.md", "A2");
  S.writeFile(saved.id, "references/sub/b.csv", "1,2");
  S.deleteFile(saved.id, "references/a.md");
  await S.whenSaved();
  assert.ok(!exists(root, "合同审查", "references", "a.md"));
  assert.strictEqual(read(root, "合同审查", "references", "sub", "b.csv"), "1,2");

  const dup = S.addUser({ name: "合同审查", content: "另一个同名技能" });
  assert.strictEqual(dup.id, "user:合同审查-2", "重名目录自动加序号");
  const odd = S.addUser({ name: "a/b:c?.", content: "x" });
  assert.strictEqual(odd.dir, "a-b-c-", "目录名里的非法字符被替换");
  await S.whenSaved();

  S.removeUser(dup.id);
  await S.whenSaved();
  assert.ok(!exists(root, "合同审查-2"));
  assert.strictEqual(fs.readdirSync(path.join(root, ".trash")).length, 1);
  assert.deepStrictEqual(fs.readdirSync(root).filter((n) => !n.startsWith(".")).sort(), ["a-b-c-", "合同审查"]);
});

test("前端磁盘模式：改名连目录一起改，启用状态跟着新 id；SKILL.md 原文里的其它 frontmatter 字段保留", async () => {
  const root = tmpRoot();
  const { S } = loadSkills(root);
  await S.init();
  const saved = S.addUser({ name: "旧名", content: "正文", files: [{ path: "references/a.md", content: "A" }] });
  S.setEnabled(saved.id, true);
  await S.whenSaved();
  fs.writeFileSync(path.join(root, "旧名", "logo.png"), Buffer.from([0, 1, 2]));

  const raw = "---\nname: 新名\ndescription: 新简介\nlicense: MIT\n---\n# 新正文\n";
  const updated = S.saveSkillMarkdown(saved.id, raw);
  assert.strictEqual(updated.id, "user:新名");
  await S.whenSaved();
  assert.ok(!exists(root, "旧名"));
  assert.strictEqual(read(root, "新名", "SKILL.md"), raw, "按用户写的原文落盘，不丢 license 这类字段");
  assert.ok(exists(root, "新名", "logo.png"), "目录里不受管的文件随改名保留");
  assert.strictEqual(read(root, "新名", "references", "a.md"), "A");
  assert.strictEqual(S.isEnabled("user:新名"), true);
  assert.strictEqual(S.isEnabled("user:旧名"), false);
  assert.strictEqual(S.buildSkillMarkdown(S.get("user:新名")), raw);
});

test("前端磁盘模式：外部直接改磁盘，refresh 后生效；没动过的技能不会被重写", async () => {
  const root = tmpRoot();
  fs.mkdirSync(path.join(root, "手工技能", "references"), { recursive: true });
  const handRaw = "---\nname: 手写的\ndescription: 外部创建\nmetadata:\n  short-description: x\n---\n\n\n正文\n";
  fs.writeFileSync(path.join(root, "手工技能", "SKILL.md"), handRaw);
  fs.writeFileSync(path.join(root, "手工技能", "references", "r.md"), "R1");
  const { S, calls } = loadSkills(root);
  await S.init();
  const sk = S.get("user:手工技能");
  assert.deepStrictEqual([sk.name, sk.description, sk.content, S.listFiles(sk)], ["手写的", "外部创建", "正文", ["references/r.md"]]);

  S.addUser({ name: "另一个", content: "x" });
  await S.whenSaved();
  assert.strictEqual(read(root, "手工技能", "SKILL.md"), handRaw, "别的技能的写入不会顺带规范化这份 SKILL.md");
  assert.deepStrictEqual(calls.filter((c) => c !== "list"), ["put"]);

  fs.writeFileSync(path.join(root, "手工技能", "references", "r.md"), "R2");
  fs.writeFileSync(path.join(root, "手工技能", "references", "new.md"), "N");
  await S.refresh();
  const again = S.get("user:手工技能");
  assert.strictEqual(S.readFile(again, "references/r.md"), "R2");
  assert.deepStrictEqual(S.listFiles(again), ["references/new.md", "references/r.md"]);
});

test("迁移：KV 里的老技能首次启动搬到磁盘，启用状态换新 id，只搬一次", async () => {
  const root = tmpRoot();
  const legacy = {
    entries: [
      { id: "sk-old1", name: "老技能", description: "d", content: "老正文", builtin: false, files: [{ path: "references/a.md", content: "A" }] },
      { id: "sk-old2", name: "没启用的", content: "x", builtin: false }
    ],
    enabled: ["builtin-ppt-designer", "sk-old1"]
  };
  const first = loadSkills(root, { legacy });
  assert.strictEqual(first.S.list().filter((s) => !s.builtin).length, 2, "init 之前仍从 KV 读");
  await first.S.init();
  assert.strictEqual(read(root, "老技能", "references", "a.md"), "A");
  assert.ok(exists(root, "没启用的", "SKILL.md"));
  assert.strictEqual(first.S.isEnabled("user:老技能"), true);
  assert.strictEqual(first.S.isEnabled("user:没启用的"), false);
  assert.strictEqual(first.S.isEnabled("builtin-ppt-designer"), true);
  assert.ok(first.kv.getItem("lingxi_skills_disk_migrated_v1"));
  assert.ok(first.kv.getItem("lingxi_skills_user_v1"), "KV 里的原数据留作备份");

  // 第二次启动（同一份 KV + 同一个磁盘）：不重复搬
  const second = loadSkills(root, { legacy });
  second.kv.setItem("lingxi_skills_disk_migrated_v1", "1");
  await second.S.init();
  assert.deepStrictEqual(fs.readdirSync(root).sort(), ["没启用的", "老技能"]);
  assert.strictEqual(second.S.list().filter((s) => !s.builtin).length, 2);
});

test("迁移写到一半失败：不打标记，下次重试且不产生重复目录", async () => {
  const root = tmpRoot();
  const legacy = { entries: [{ id: "sk-1", name: "技能一", content: "x", builtin: false }] };
  const first = loadSkills(root, { legacy, failRoutes: ["put"] });
  await first.S.init();
  assert.strictEqual(first.kv.getItem("lingxi_skills_disk_migrated_v1"), null);
  assert.deepStrictEqual(fs.readdirSync(root), []);

  const createSkillStoreAgain = loadSkills(root, { legacy });
  await createSkillStoreAgain.S.init();
  await loadSkills(root, { legacy }).S.init(); // 标记没带过来的第三次启动：靠同名同内容去重
  assert.deepStrictEqual(fs.readdirSync(root), ["技能一"]);
});

test("写盘失败：whenSaved 抛错，缓存按磁盘回滚", async () => {
  const root = tmpRoot();
  const { S } = loadSkills(root, { failRoutes: ["write"] });
  await S.init();
  const saved = S.addUser({ name: "A", content: "正文" });
  await S.whenSaved();
  S.writeFile(saved.id, "references/a.md", "A");
  assert.deepStrictEqual(S.listFiles(S.get(saved.id)), ["references/a.md"], "先乐观更新缓存");
  await assert.rejects(S.whenSaved(), /磁盘已满/);
  assert.deepStrictEqual(S.listFiles(S.get(saved.id)), [], "失败后以磁盘为准");
  await S.whenSaved(); // 错误只抛一次
});

test("代理连不上：没迁移过 → 继续用 KV；迁移过 → 不显示旧数据且拒绝写入", async () => {
  const fresh = loadSkills(tmpRoot());
  fresh.state.down = true;
  assert.strictEqual(await fresh.S.init(), "kv");
  const saved = fresh.S.addUser({ name: "离线也能存", content: "x" });
  assert.match(saved.id, /^sk-/);
  assert.ok(fresh.kv.getItem("lingxi_skills_user_v1").includes("离线也能存"));

  const migrated = loadSkills(tmpRoot(), { legacy: { entries: [{ id: "sk-1", name: "旧副本", content: "x", builtin: false }] } });
  migrated.kv.setItem("lingxi_skills_disk_migrated_v1", "1");
  migrated.state.down = true;
  assert.strictEqual(await migrated.S.init(), "offline");
  assert.deepStrictEqual(migrated.S.list().filter((s) => !s.builtin), []);
  assert.throws(() => migrated.S.addUser({ name: "x", content: "y" }), /本地服务未连接/);
});

test("重读磁盘在途时本地又改了：过期的读取结果不会把刚做的改动盖回去", async () => {
  const root = tmpRoot();
  const { S, state, calls } = loadSkills(root);
  await S.init();
  const saved = S.addUser({ name: "旧名", content: "正文" });
  await S.whenSaved();

  let release;
  state.listGate = new Promise((r) => { release = r; });
  calls.length = 0;
  const pending = S.refresh();
  await new Promise((r) => setImmediate(r));
  assert.deepStrictEqual(calls, ["list"], "list 请求已发出、还没返回");
  const updated = S.saveSkillMarkdown(saved.id, "---\nname: 新名\n---\n新正文"); // 这时候改名
  state.listGate = null;
  release();
  await pending;
  await S.whenSaved();
  assert.deepStrictEqual(S.list().filter((s) => !s.builtin).map((s) => s.id), ["user:新名"]);
  assert.strictEqual(updated.id, "user:新名");
  assert.deepStrictEqual(fs.readdirSync(root), ["新名"]);
  S.writeFile("user:新名", "references/a.md", "A"); // 后续操作还能接着做
  await S.whenSaved();
  assert.strictEqual(read(root, "新名", "references", "a.md"), "A");
});

const BUNDLED_SRC = path.join(__dirname, "..", "skills", "ui-ux-pro-max");

test("内置 UI/UX Pro Max 装到磁盘：SKILL.md + 7 个 CSV 落盘，前端把它们挂到内置技能上而不是变成自定义技能", async () => {
  const root = tmpRoot();
  const { S, store, calls } = loadSkills(root);
  assert.strictEqual(store.installBundled(BUNDLED_SRC, "ui-ux-pro-max"), "installed");
  assert.strictEqual(store.installBundled(BUNDLED_SRC, "ui-ux-pro-max"), "current", "再装一次什么都不做");
  const csvOnDisk = fs.readdirSync(path.join(root, "ui-ux-pro-max", "data")).sort();
  assert.deepStrictEqual(csvOnDisk, ["charts.csv", "colors.csv", "design.csv", "styles.csv", "typography.csv", "ui-reasoning.csv", "ux-guidelines.csv"]);
  assert.strictEqual(read(root, "ui-ux-pro-max", "data", "colors.csv"), fs.readFileSync(path.join(BUNDLED_SRC, "data", "colors.csv"), "utf8"));

  await S.init();
  assert.deepStrictEqual(S.list().filter((s) => !s.builtin), [], "不会多出一条自定义技能");
  const builtin = S.get("builtin-ui-ux-pro-max");
  assert.strictEqual(builtin.builtin, true);
  assert.strictEqual(builtin.name, "UI/UX Pro Max 设计智能", "名字 / 简介 / hostFilter 仍用内置定义");
  assert.deepStrictEqual(builtin.hostFilter, ["wpp"]);
  assert.deepStrictEqual(S.listFiles(builtin), csvOnDisk.map((n) => "data/" + n), "7 个 CSV 都能读，全部在单文件 / 总量上限内");
  assert.match(S.readFile(builtin, "data/colors.csv"), /^No,Product Type,Primary/);
  assert.ok((await S.loadContent(builtin)).length > 10000, "正文也来自磁盘，不用再 fetch contentPath");

  // 前端对这个目录只读：别的技能增删改不会碰它
  S.addUser({ name: "别的技能", content: "x" });
  S.removeUser("user:别的技能");
  await S.whenSaved();
  assert.deepStrictEqual(calls.filter((c) => c !== "list"), ["put", "delete"]);
  assert.strictEqual(fs.readdirSync(path.join(root, "ui-ux-pro-max", "data")).length, 7);
});

test("内置技能升级：没动过的文件换新版，用户改过的保留；用户自己的同名目录不碰", () => {
  const root = tmpRoot();
  const src = tmpRoot();
  const store = createSkillStore(root);
  const writeSrc = (rel, text) => { fs.mkdirSync(path.dirname(path.join(src, rel)), { recursive: true }); fs.writeFileSync(path.join(src, rel), text); };
  writeSrc("SKILL.md", "---\nname: demo\n---\nv1");
  writeSrc("data/a.csv", "a-v1");
  writeSrc("data/b.csv", "b-v1");
  writeSrc("data/gone.csv", "gone-v1");
  assert.strictEqual(store.installBundled(src, "demo"), "installed");

  fs.writeFileSync(path.join(root, "demo", "data", "b.csv"), "b-用户改过");
  fs.writeFileSync(path.join(root, "demo", "data", "mine.csv"), "用户自己加的");
  writeSrc("SKILL.md", "---\nname: demo\n---\nv2");
  writeSrc("data/a.csv", "a-v2");
  writeSrc("data/b.csv", "b-v2");
  writeSrc("data/new.csv", "new-v2");
  fs.unlinkSync(path.join(src, "data", "gone.csv"));

  assert.strictEqual(store.installBundled(src, "demo"), "updated");
  assert.match(read(root, "demo", "SKILL.md"), /v2$/);
  assert.strictEqual(read(root, "demo", "data", "a.csv"), "a-v2");
  assert.strictEqual(read(root, "demo", "data", "b.csv"), "b-用户改过", "用户改过的不覆盖");
  assert.strictEqual(read(root, "demo", "data", "new.csv"), "new-v2");
  assert.strictEqual(read(root, "demo", "data", "mine.csv"), "用户自己加的");
  assert.ok(!exists(root, "demo", "data", "gone.csv"), "新版去掉、用户也没改过的文件跟着删");
  assert.strictEqual(store.installBundled(src, "demo"), "current");
  assert.strictEqual(read(root, "demo", "data", "b.csv"), "b-用户改过", "之后的启动也不会再去覆盖它");

  // 整个目录被用户删了 → 重新装回来
  fs.rmSync(path.join(root, "demo"), { recursive: true });
  assert.strictEqual(store.installBundled(src, "demo"), "installed");

  // 用户自己建的同名技能（没有标记文件）→ 不碰，前端照常当自定义技能
  store.put("mine", [{ path: "SKILL.md", content: "用户的" }]);
  assert.strictEqual(store.installBundled(src, "mine"), "skipped");
  assert.strictEqual(read(root, "mine", "SKILL.md"), "用户的");
  assert.ok(!store.list().skills.find((s) => s.dir === "mine").bundled);
  assert.strictEqual(store.list().skills.find((s) => s.dir === "demo").bundled, true);
});
