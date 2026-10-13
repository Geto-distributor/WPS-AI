// 技能（Skills）模块：
// - 内置技能（hardcoded）：针对 WPS 各宿主调好的指令模板（PPT 设计师 / Word 润色 / Excel 数据助手 / PDF 阅读）
// - 用户导入技能：单个 .md / .txt，或「技能包」（文件夹 / .zip：SKILL.md + references/ 等附属文本文件）
// - 启用的技能在 chat 时拼到 system prompt 后面
// - 自定义技能存磁盘：~/.lingxi-ai/skills/<技能名>/SKILL.md (+ references/ …)，经本地代理 /skills/* 读写。
//   本模块保留一份内存缓存让 list()/get() 仍是同步的；改动先落缓存，再排队异步写盘（whenSaved() 可等结果）。
//   代理从没连上过（init 没成功）时退回老的 KV 存法，保证离线全新安装也能用。
//
// 数据形状：
//   skill = { id, name, description, content, builtin?: boolean, files?: [{ path, content }] }
//   files：技能包里 SKILL.md 之外的文本文件（如 references/xxx.md），AI 用 read_skill_file 按需读取，
//          不随 use_skill 一次性塞进上下文。只存文本、只读，不执行（scripts/ 导入时直接跳过）。
//   localStorage:
//     - lingxi_skills_user_v1: { entries: [skill, ...], savedAt }
//     - lingxi_skills_enabled_v1: { ids: [skillId, ...] }
(function attachSkills(global) {
  "use strict";

  const USER_KEY = "lingxi_skills_user_v1";
  const ENABLED_KEY = "lingxi_skills_enabled_v1";
  const CLOUD_KEY = "lingxi_skills_cloud_v1"; // 缓存从 OSS 拉到的云端技能元数据（离线也能列出）
  // 云端技能目录索引（OSS）：{ skills: [{ id, name, description, url|content|contentPath, hostFilter? }] }
  const DEFAULT_CLOUD_SKILLS_URL = "https://llteac-file.oss-cn-hangzhou.aliyuncs.com/wps-ai/skills/index.json";

  // 经本地 proxy 的 /forward/ 转发拉 OSS（避开 WebView 跨域）。返回原始响应体。
  function proxyForward(url) {
    const pre = global.WpsAiRuntime?.forwardPrefix?.() || "http://127.0.0.1:3890/forward/";
    return pre + encodeURIComponent(url);
  }

  // ===== 内置技能（开箱即用，针对 WPS 各宿主场景）=====
  const BUILTIN = [
    {
      id: "builtin-ppt-designer",
      name: "PPT 视觉设计师",
      description: "做 PPT 时遵循专业视觉设计原则：强对比、留白、层级、配色。AI 生成或修改 PPT 时启用。",
      builtin: true,
      content: [
        "你做 PPT 时严格遵循以下视觉设计原则：",
        "1. **强对比**：标题用大字号 + 加粗，正文 18-22pt，标题至少 40pt；标题颜色用 primaryColor，正文用 bodyColor（自动跟着 stylePreset 来）。",
        "2. **充足留白**：四边 margin ≥ 80px（1920×1080 画布下），元素之间 gap 至少 40px。绝对避免内容贴边或挤成一坨。",
        "3. **清晰层级**：一页最多 3 个视觉层级；用 size + weight + color 三种手段拉开；每页只有一个视觉重点。",
        "4. **配色克制**：每页用色板里的 2-3 个颜色，**绝对不超过 5 种**；accentColor 只用于关键数字/重点；surfaceColor 用于卡片底色。",
        "5. **字号映射**：1pt = 2px（1920×1080 画布）。封面巨标题 120-192px / H1 80-108px / H2 56-72px / 正文 36-44px / 卡片 32-40px / 脚注 22-28px。**绝对不要 < 20px**。",
        "6. **数据可视化优先**：有数字/趋势/占比/多维比较时用 ECharts，比纯文字描述强 10 倍。",
        "7. **避免反模式**：项目符号铺满整页、大段正文、4 色以上配色、3 种以上字体、装饰元素抢戏、贴边不留白。"
      ].join("\n")
    },
    {
      id: "builtin-word-polish",
      name: "Word 中文润色",
      description: "改 Word 文档时按中文写作规范润色：简练、主动语态、去口语化、专业措辞。",
      builtin: true,
      content: [
        "你处理 Word 文档时遵循中文写作规范：",
        "1. **简练**：删冗余词（「进行」「一些」「的话」「其实」「然后」「那么」等填充词）、合并短句。",
        "2. **主动语态优先**：把「被 X 所 Y」改成「X Y」；少用「被」。",
        "3. **去口语化**：「我觉得」→「显然」或直接陈述；「挺好的」→ 具体描述；不用「嗯」「啊」「呢」等语气词。",
        "4. **术语统一**：同一个概念在全文里用一致的术语，第一次出现给定义或英文原文。",
        "5. **段落结构**：每段一个核心论点，首句点题；段落之间用逻辑词承接（「因此」「然而」「相比之下」）。",
        "6. **数字规范**：正文 0-10 用汉字，11+ 用阿拉伯数字；百分比、年份、版本号一律阿拉伯数字。",
        "7. **标点严谨**：中文标点（，。：；？！「」）；不用中英文混标。",
        "8. **保留原意**：润色不是改写，保持作者原意和事实，只调表达方式。"
      ].join("\n")
    },
    {
      id: "builtin-excel-data",
      name: "Excel 数据助手",
      description: "处理 Excel 时按数据分析最佳实践：表头清晰、公式规范、命名区域、不破坏原数据。",
      builtin: true,
      content: [
        "你处理 Excel 表格时遵循以下规范：",
        "1. **保留原数据**：原始数据 sheet 不动；新增公式/汇总放新 sheet 或新区域。",
        "2. **公式规范**：用绝对引用锁定区域（$A$1:$A$100），命名区域代替 magic range（用「销售额」代替 D2:D100）。",
        "3. **避免硬编码**：阈值/系数放单独单元格做参数，公式引用它，方便用户调参。",
        "4. **错误处理**：用 IFERROR / IFNA 包住可能出错的公式，避免满屏 #N/A。",
        "5. **表头**：第一行始终是表头（加粗 + 浅底色）；下面不要插空行。",
        "6. **数据类型**：日期统一格式（YYYY-MM-DD），数字保留必要小数位，文本不要混数字。",
        "7. **汇总优先用 SUMIFS/COUNTIFS/AVERAGEIFS**：可读性 > VLOOKUP 嵌套；XLOOKUP 在新版 Excel 里替代 VLOOKUP。",
        "8. **透视表 / Power Query**：复杂聚合用透视表；数据清洗用 Power Query 步骤，记得告诉用户怎么 refresh。"
      ].join("\n")
    },
    {
      id: "builtin-pdf-reader",
      name: "PDF 阅读助手",
      description: "处理 PDF 文档时按知识工作者节奏：先抓结构、再概要、再深挖；引用具体页码。",
      builtin: true,
      content: [
        "你处理 PDF 文档时按这套节奏：",
        "1. **先看目录/章节结构**：用户问之前，先扫一遍 outline 或前几页，建立全文骨架。",
        "2. **概要回答**：用户问问题先用 1-2 句回答核心，再展开细节，避免直接铺长答案。",
        "3. **引用页码**：所有论断都标注来源页（如「第 5 页提到...」「图 3 显示...」），便于用户回查。",
        "4. **量化优先**：PDF 里出现数字、表格、图表时优先引用具体数字，比抽象描述更有信息量。",
        "5. **术语保留**：专业术语用原文（中文 PDF 保留中文，英文 PDF 保留英文），不强行翻译让用户产生歧义。",
        "6. **结构化输出**：列表 / 表格 / 时间线等结构化内容用 markdown 还原，不要全段落叙述。",
        "7. **不知道就说不知道**：PDF 没覆盖的问题直接讲「原文未提及」，不要瞎编。"
      ].join("\n")
    },
    {
      // UI/UX Pro Max 设计智能 —— 引入 nextlevelbuilder/ui-ux-pro-max-skill 的设计知识库。
      // content 大（44KB SKILL.md），用 contentPath 懒加载，第一次用到才 fetch + 缓存。
      // 这条 skill 用于「用 HTML 生成 PPT」自由设计场景：当用户在 prompt 里明确风格时
      // (cyberpunk / 极简 / 玻璃拟态 / 暗黑等), 用户的样式预设让位给这套设计自由度。
      id: "builtin-ui-ux-pro-max",
      name: "UI/UX Pro Max 设计智能",
      description: "50+ 风格 / 161 色板 / 57 字体配对 / 99 UX 准则 / 25 图表类型。AI 用 HTML 生成 PPT 时按需引用，让幻灯片不再千篇一律。来自 nextlevelbuilder/ui-ux-pro-max-skill（MIT）。",
      builtin: true,
      contentPath: "skills/ui-ux-pro-max/SKILL.md",
      // 代理会把整个技能目录（含 data/*.csv）装到 ~/.lingxi-ai/skills/ui-ux-pro-max/；
      // 装好后正文和参考文件都从磁盘读（见 loadFromDisk），contentPath 只是代理不可用时的兜底。
      diskDir: "ui-ux-pro-max",
      // 只在 PPT 宿主注入；其他宿主聊天不浪费 ~10K 个 token
      hostFilter: ["wpp"]
    }
  ];

  // 懒加载缓存：内容来源(url/contentPath) → 已 fetch 的 content
  const _contentCache = new Map();
  async function loadContent(skill) {
    if (skill.content) return skill.content;
    const src = skill.url || skill.contentPath;
    if (!src) return "";
    if (_contentCache.has(src)) return _contentCache.get(src);
    try {
      // 绝对 http(s)（云端技能）走 proxy 转发避跨域；相对路径（内置随包）直接取
      const fetchUrl = /^https?:\/\//i.test(src) ? proxyForward(src) : src;
      const resp = await fetch(fetchUrl, { cache: "no-cache" });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const text = await resp.text();
      _contentCache.set(src, text);
      return text;
    } catch (e) {
      console.warn(`[skills] 加载 ${src} 失败:`, e?.message || e);
      return "";
    }
  }

  // ===== 云端技能（从 OSS 目录加载）=====
  // 把一条云端技能原始条目规整成内部 skill 形态；无有效内容来源则返回 null。
  function normalizeCloudSkill(raw) {
    if (!raw) return null;
    const rid = String(raw.id || raw.name || "").trim();
    if (!rid) return null;
    const id = rid.indexOf("cloud-") === 0 ? rid : "cloud-" + rid;
    const out = {
      id,
      name: String(raw.name || rid).trim(),
      description: String(raw.description || "").trim(),
      builtin: false,
      source: "cloud"
    };
    if (typeof raw.content === "string" && raw.content.trim()) out.content = raw.content.trim();
    else if (raw.url) out.url = String(raw.url).trim();
    else if (raw.contentPath) out.contentPath = String(raw.contentPath).trim();
    else return null;
    if (Array.isArray(raw.hostFilter) && raw.hostFilter.length) out.hostFilter = raw.hostFilter.map(String);
    return out;
  }

  // 解析云端索引：接受 { skills: [...] } 或直接数组。纯函数。
  function parseCloudIndex(data) {
    const arr = Array.isArray(data) ? data : (Array.isArray(data?.skills) ? data.skills : []);
    const seen = new Set();
    const out = [];
    arr.map(normalizeCloudSkill).forEach((s) => {
      if (!s || seen.has(s.id)) return;
      seen.add(s.id);
      out.push(s);
    });
    return out;
  }

  function readCloudSkills() {
    try {
      const raw = global.WpsAiStore.getItem(CLOUD_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed?.entries) ? parsed.entries : [];
    } catch (e) { return []; }
  }

  function writeCloudSkills(entries) {
    try { global.WpsAiStore.setItem(CLOUD_KEY, JSON.stringify({ entries, savedAt: Date.now() })); } catch (e) {}
  }

  // 从 OSS 拉云端技能目录，刷新本地缓存。失败回退到缓存（离线可用）。best-effort。
  async function loadCloud(opts) {
    const url = (opts && opts.url)
      || global.WpsAiProviderRegistry?.loadSettings?.().cloudSkillsUrl
      || DEFAULT_CLOUD_SKILLS_URL;
    try {
      const resp = await fetch(proxyForward(url), { cache: "no-cache" });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const data = await resp.json();
      const skills = parseCloudIndex(data);
      writeCloudSkills(skills);
      // 刷新时清掉云端（绝对 URL）的内容缓存，让下次注入拿最新正文，
      // 避免 OSS 同 URL 更新内容后本会话仍用旧缓存。内置的相对 contentPath 缓存不动。
      for (const k of Array.from(_contentCache.keys())) {
        if (/^https?:/i.test(k)) _contentCache.delete(k);
      }
      return skills;
    } catch (e) {
      console.warn("[skills] 云端技能加载失败，用缓存兜底:", e?.message || e);
      return readCloudSkills();
    }
  }

  // ===== 自定义技能的存储后端 =====
  //   "kv"      老存法：整包 JSON 存 WpsAiStore（init 之前 / 代理从没连上过）
  //   "disk"    磁盘目录，_cache 是内存镜像
  //   "offline" 已迁到磁盘但这次连不上代理：不显示也不允许改（KV 里那份是迁移前的旧数据，不能再当真）
  const DISK_MIGRATED_KEY = "lingxi_skills_disk_migrated_v1";
  const MAIN_FILE = "SKILL.md";
  let _backend = "kv";
  let _cache = [];
  let _builtinDisk = new Map(); // 内置技能 id → { content, files }：随包技能装到磁盘后的正文 + 参考文件
  let _snapshot = new Map(); // 磁盘上已知状态：dir → Map(path → content)，用来算最小写盘操作
  let _chain = Promise.resolve();
  let _writeError = null;
  let _rev = 0; // 本地每改一次 +1；重读磁盘期间它变了，说明读回来的是改动前的旧状态，不能拿来覆盖缓存

  function proxyBase() {
    return global.WpsAiRuntime?.proxyBase?.() || "http://127.0.0.1:3890";
  }
  async function diskCall(route, body) {
    const resp = await fetch(proxyBase() + "/skills/" + route, body === undefined
      ? { method: "GET", cache: "no-store" }
      : { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    const json = await resp.json().catch(() => null);
    if (!resp.ok || !json?.ok) throw new Error(json?.error || `技能存储请求失败（${resp.status}）`);
    return json;
  }

  function readLegacyUserSkills() {
    try {
      const raw = global.WpsAiStore.getItem(USER_KEY);
      if (!raw) return [];
      const parsed = JSON.parse(raw);
      return Array.isArray(parsed?.entries) ? parsed.entries : [];
    } catch (e) {
      console.warn("[skills] 读取用户技能失败：", e?.message || e);
      return [];
    }
  }

  function readUserSkills() {
    return _backend === "kv" ? readLegacyUserSkills() : _cache;
  }

  function writeUserSkills(entries) {
    if (_backend === "offline") throw new Error("本地服务未连接，暂时无法保存技能，请稍后重试");
    if (_backend === "disk") {
      _cache = entries;
      _rev += 1;
      syncToDisk();
      return;
    }
    try {
      global.WpsAiStore.setItem(USER_KEY, JSON.stringify({ entries, savedAt: Date.now() }));
    } catch (e) {
      console.error("[skills] 写入用户技能失败（localStorage 可能已满）:", e?.message || e);
    }
  }

  // 技能名 → 目录名：去掉路径分隔符 / Windows 保留字符 / 首尾的点和空格；重名加 -2、-3…
  function dirNameFor(name, takenDirs) {
    // eslint-disable-next-line no-control-regex
    let base = String(name || "").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "-").replace(/^[.\s]+|[.\s]+$/g, "").slice(0, 80).replace(/[.\s]+$/g, "");
    if (!base) base = "skill";
    const taken = new Set(Array.from(takenDirs || []).map((d) => String(d).toLowerCase()));
    let dir = base;
    for (let n = 2; taken.has(dir.toLowerCase()); n += 1) dir = `${base}-${n}`;
    return dir;
  }
  const diskId = (dir) => "user:" + dir;

  // 一条技能在磁盘上应有的文件：SKILL.md（优先原文 raw，保住用户写的其它 frontmatter 字段）+ 参考文件
  function diskFilesOf(entry) {
    const files = new Map();
    files.set(entry.mainPath || MAIN_FILE, typeof entry.raw === "string" ? entry.raw : buildSkillMarkdown(entry));
    (Array.isArray(entry.files) ? entry.files : []).forEach((f) => files.set(f.path, f.content));
    return files;
  }

  // 缓存 vs 磁盘快照 → 最小操作序列，排进写队列。只动本模块管的文本文件：
  // 目录里的图片 / 脚本等不在快照里，改名用真 rename、删除是挪进 .trash，都不会丢。
  function syncToDisk() {
    const ops = [];
    _cache.forEach((e) => {
      if (e._from && _snapshot.has(e._from) && !_snapshot.has(e.dir)) {
        ops.push(["rename", { from: e._from, to: e.dir }]);
        _snapshot.set(e.dir, _snapshot.get(e._from));
        _snapshot.delete(e._from);
      }
      delete e._from;
    });
    const desired = new Map(_cache.map((e) => [e.dir, diskFilesOf(e)]));
    Array.from(_snapshot.keys()).forEach((dir) => {
      if (desired.has(dir)) return;
      ops.push(["delete", { dir }]);
      _snapshot.delete(dir);
    });
    desired.forEach((files, dir) => {
      const old = _snapshot.get(dir);
      if (!old) {
        ops.push(["put", { dir, files: Array.from(files, ([path, content]) => ({ path, content })) }]);
      } else {
        files.forEach((content, path) => { if (old.get(path) !== content) ops.push(["write", { dir, path, content }]); });
        old.forEach((_c, path) => { if (!files.has(path)) ops.push(["delete-file", { dir, path }]); });
      }
      _snapshot.set(dir, files);
    });
    if (!ops.length) return;
    _chain = _chain.then(async () => {
      try {
        for (const [route, body] of ops) await diskCall(route, body);
      } catch (e) {
        // 写盘失败：缓存已经和磁盘对不上了，重读一遍以磁盘为准，错误留给 whenSaved() 抛给界面
        _writeError = e;
        console.error("[skills] 写盘失败：", e?.message || e);
        try { await loadFromDisk(); } catch (err) {}
      }
    });
  }

  // 等排队中的写盘做完；期间有失败就抛出来（只抛一次）。界面保存后 await 它来决定报成功还是失败。
  async function whenSaved() {
    await _chain;
    if (_writeError) { const e = _writeError; _writeError = null; throw e; }
  }

  async function loadFromDisk() {
    const rev = _rev;
    const json = await diskCall("list");
    // 请求在途时本地又改了：这份结果已过期，丢掉（改动自己的写盘排在后面，缓存本来就是对的）
    if (rev !== _rev && _backend === "disk") return;
    const entries = [];
    const snapshot = new Map();
    const builtinDisk = new Map();
    (json.skills || []).forEach((raw) => {
      try {
        const main = (raw.files || []).find((f) => SKILL_MD_RE.test(f.path));
        if (!main) return;
        const built = buildSkillFromFiles(raw.files, raw.dir);
        // 随包技能的目录（代理装的，带标记）：不当成自定义技能，而是给对应内置技能补上正文和参考文件；
        // 不进 snapshot，所以本模块永远不会去写 / 删它。没标记的同名目录是用户自己的，照常当自定义技能。
        const builtin = raw.bundled ? BUILTIN.find((b) => b.diskDir === raw.dir) : null;
        if (builtin) {
          builtinDisk.set(builtin.id, { content: built.content, files: built.files });
          return;
        }
        const entry = {
          id: diskId(raw.dir), dir: raw.dir, mainPath: main.path, raw: main.content,
          name: built.name, description: built.description, content: built.content,
          builtin: false, ts: Number(raw.mtime) || 0
        };
        if (built.files.length) entry.files = built.files;
        entries.push(entry);
        snapshot.set(raw.dir, diskFilesOf(entry));
      } catch (e) { /* SKILL.md 为空等：这个目录不算技能 */ }
    });
    _cache = entries;
    _snapshot = snapshot;
    _builtinDisk = builtinDisk;
    _backend = "disk";
  }

  // 把 KV 里的老技能搬到磁盘目录，只做一次。启用状态跟着换成新 id。KV 里的原数据不删，留作备份。
  async function migrateLegacyToDisk() {
    if (global.WpsAiStore.getItem(DISK_MIGRATED_KEY)) return;
    const legacy = readLegacyUserSkills();
    if (legacy.length) {
      const enabled = readEnabledIds();
      const next = _cache.slice();
      legacy.forEach((old) => {
        if (!old?.name || typeof old.content !== "string" || !old.content.trim()) return;
        // 上次迁移写到一半失败过：已经在磁盘上的同名同内容技能不再重复搬
        const dup = next.find((e) => e.name === String(old.name).trim() && e.content === old.content.trim());
        if (dup) { if (enabled.delete(old.id)) enabled.add(dup.id); return; }
        const dir = dirNameFor(old.name, next.map((e) => e.dir));
        const entry = {
          id: diskId(dir), dir, name: String(old.name).trim(), description: String(old.description || "").trim(),
          content: old.content.trim(), builtin: false, ts: Number(old.ts) || Date.now()
        };
        const files = normalizeSkillFiles(old.files);
        if (files.length) entry.files = files;
        next.push(entry);
        if (enabled.delete(old.id)) enabled.add(entry.id);
      });
      writeUserSkills(next);
      await whenSaved(); // 写盘失败会抛：不打标记，下次启动再试
      writeEnabledIds(enabled);
    }
    global.WpsAiStore.setItem(DISK_MIGRATED_KEY, String(Date.now()));
  }

  // 启动时调（WpsAiStore 就绪后）：连得上代理就切到磁盘存储并做一次性迁移。返回当前后端名。
  let _initPromise = null;
  function init() {
    if (!_initPromise) _initPromise = doInit();
    return _initPromise;
  }
  async function doInit() {
    try {
      await loadFromDisk();
    } catch (e) {
      if (global.WpsAiStore.getItem(DISK_MIGRATED_KEY)) { _backend = "offline"; _cache = []; }
      return _backend;
    }
    try { await migrateLegacyToDisk(); } catch (e) { console.warn("[skills] 迁移到磁盘失败，下次启动重试：", e?.message || e); }
    return _backend;
  }

  // 重读磁盘（用户可能在外部改了文件 / 设置窗口里改了技能）。排在未完成的写盘之后，不会读到半截状态。
  async function refresh() {
    if (_initPromise) await _initPromise; // 启动时的首次加载 / 迁移还没完就先等它
    if (_backend === "kv") return _backend;
    _chain = _chain.then(() => loadFromDisk().catch(() => {}));
    return _chain.then(() => _backend);
  }

  // 内置技能默认全部启用 —— 用户装完插件开箱即用，少一步「为什么 AI 不按 PPT 设计师那套来？」
  // 用户不需要的可以在设置 → 技能里逐条关掉。BUILTIN 数组变化时这里也要跟。
  const DEFAULT_ENABLED = BUILTIN.map((s) => s.id);

  function readEnabledIds() {
    try {
      const raw = global.WpsAiStore.getItem(ENABLED_KEY);
      if (!raw) return new Set(DEFAULT_ENABLED);
      const parsed = JSON.parse(raw);
      return new Set(Array.isArray(parsed?.ids) ? parsed.ids : []);
    } catch (e) { return new Set(DEFAULT_ENABLED); }
  }

  function writeEnabledIds(set) {
    try {
      global.WpsAiStore.setItem(ENABLED_KEY, JSON.stringify({ ids: Array.from(set), savedAt: Date.now() }));
    } catch (e) {}
  }

  // 列出全部（内置 → 云端 → 用户自定义）
  function list() {
    const builtins = BUILTIN.map((b) => {
      const onDisk = _builtinDisk.get(b.id);
      if (!onDisk) return b;
      const merged = Object.assign({}, b, { content: onDisk.content });
      if (onDisk.files.length) merged.files = onDisk.files;
      return merged;
    });
    return builtins.concat(readCloudSkills()).concat(readUserSkills());
  }

  function get(id) {
    return list().find((s) => s.id === id) || null;
  }

  function isEnabled(id) {
    return readEnabledIds().has(id);
  }

  function setEnabled(id, on) {
    const set = readEnabledIds();
    if (on) set.add(id);
    else set.delete(id);
    writeEnabledIds(set);
  }

  function getEnabledSkills() {
    const set = readEnabledIds();
    return list().filter((s) => set.has(s.id));
  }

  // ===== 技能包附属文件（references/ 等）=====
  const FILE_LIMITS = { maxFiles: 50, maxFileChars: 200 * 1000, maxTotalChars: 1000 * 1000 };
  const TEXT_FILE_RE = /\.(md|markdown|txt|csv|tsv|json|ya?ml|xml|html?|css)$/i;
  const SKILL_MD_RE = /^skill\.(md|markdown)$/i;

  // 统一成正斜杠相对路径；含 .. / 绝对路径 / 空段的一律判无效（返回 ""）
  function normalizeFilePath(path) {
    const parts = String(path == null ? "" : path).replace(/\\/g, "/").split("/").filter((seg) => seg && seg !== ".");
    if (!parts.length || parts.some((seg) => seg === "..")) return "";
    return parts.join("/");
  }

  // 规整一组附属文件：去重、去空、套上限。超限的进 skipped，不抛错（导入尽量成功）。
  function normalizeSkillFiles(files, skipped) {
    const out = [];
    const seen = new Set();
    let total = 0;
    const skip = (path, reason) => { if (skipped) skipped.push({ path, reason }); };
    (Array.isArray(files) ? files : []).forEach((f) => {
      const path = normalizeFilePath(f?.path);
      const content = typeof f?.content === "string" ? f.content : "";
      if (!path || seen.has(path.toLowerCase())) return;
      if (!content.trim()) return skip(path, "空文件");
      if (content.length > FILE_LIMITS.maxFileChars) return skip(path, "文件过大");
      if (out.length >= FILE_LIMITS.maxFiles) return skip(path, "文件数超过上限");
      if (total + content.length > FILE_LIMITS.maxTotalChars) return skip(path, "总大小超过上限");
      seen.add(path.toLowerCase());
      total += content.length;
      out.push({ path, content });
    });
    return out;
  }

  // 把一个技能包（文件夹 / zip 展开后的 [{path, content}]）组装成一条技能：
  //   - 取层级最浅的 SKILL.md 当正文，它所在目录就是技能根目录（兼容「包了一层同名文件夹」的 zip）
  //   - 根目录下其它文本文件按相对路径收进 files；scripts/、隐藏文件、非文本文件跳过
  // 返回 { name, description, content, files, skipped }；找不到 SKILL.md 抛错。
  function buildSkillFromFiles(entries, fallbackName) {
    const all = (Array.isArray(entries) ? entries : [])
      .map((e) => ({ path: normalizeFilePath(e?.path), content: typeof e?.content === "string" ? e.content : "" }))
      .filter((e) => e.path && !/(^|\/)(__MACOSX|\.[^/]*)(\/|$)/.test(e.path));
    const depth = (path) => path.split("/").length;
    const main = all
      .filter((e) => SKILL_MD_RE.test(e.path.split("/").pop()))
      .sort((x, y) => depth(x.path) - depth(y.path))[0];
    if (!main) throw new Error("技能包里没找到 SKILL.md");
    const root = main.path.split("/").slice(0, -1).join("/");
    const prefix = root ? root + "/" : "";
    const parsed = parseMarkdownSkill(main.content, (root.split("/").pop() || fallbackName || "") + ".md");
    if (!parsed.content) throw new Error("SKILL.md 内容为空");
    const skipped = [];
    const candidates = [];
    all.forEach((e) => {
      if (e === main || !e.path.startsWith(prefix)) return;
      const rel = e.path.slice(prefix.length);
      if (/^scripts\//i.test(rel)) return skipped.push({ path: rel, reason: "脚本不支持执行" });
      if (!TEXT_FILE_RE.test(rel)) return skipped.push({ path: rel, reason: "非文本文件" });
      candidates.push({ path: rel, content: e.content });
    });
    const files = normalizeSkillFiles(candidates, skipped);
    return Object.assign(parsed, { files, skipped });
  }

  function listFiles(skill) {
    return (Array.isArray(skill?.files) ? skill.files : []).map((f) => f.path);
  }

  // 按路径读技能的附属文件；找不到返回 null。大小写不敏感兜底（模型偶尔会改大小写）。
  function readFile(skill, path) {
    const want = normalizeFilePath(path);
    const files = Array.isArray(skill?.files) ? skill.files : [];
    if (!want) return null;
    const hit = files.find((f) => f.path === want) || files.find((f) => f.path.toLowerCase() === want.toLowerCase());
    return hit ? hit.content : null;
  }

  // 把一组路径（如 ["references/a.md", "references/sub/b.csv"]）整理成目录树，给设置里的技能详情用。
  // 节点：{ type: "dir"|"file", name, path, children? }；每层目录在前、文件在后，各自按名字排序。
  function buildFileTree(paths) {
    const root = { type: "dir", name: "", path: "", children: [] };
    (Array.isArray(paths) ? paths : []).forEach((raw) => {
      const path = normalizeFilePath(raw);
      if (!path) return;
      const parts = path.split("/");
      let node = root;
      parts.forEach((name, i) => {
        const isFile = i === parts.length - 1;
        const nodePath = parts.slice(0, i + 1).join("/");
        let child = node.children.find((c) => c.name === name && c.type === (isFile ? "file" : "dir"));
        if (!child) {
          child = isFile ? { type: "file", name, path: nodePath } : { type: "dir", name, path: nodePath, children: [] };
          node.children.push(child);
        }
        node = child;
      });
    });
    const sort = (node) => {
      node.children.sort((x, y) => (x.type === y.type ? x.name.localeCompare(y.name) : (x.type === "dir" ? -1 : 1)));
      node.children.forEach((c) => { if (c.type === "dir") sort(c); });
    };
    sort(root);
    return root.children;
  }

  // 技能 → SKILL.md 全文（frontmatter + 正文），编辑 SKILL.md 时给用户看的就是这个；
  // 存回去时走 parseMarkdownSkill 拆回 name / description / content，两者互逆。
  // 磁盘技能带 raw（SKILL.md 原文）时直接给原文，保住 name/description 之外的 frontmatter 字段。
  function buildSkillMarkdown(skill) {
    if (typeof skill?.raw === "string" && skill.raw.trim()) return skill.raw;
    const oneLine = (v) => String(v == null ? "" : v).replace(/\s*\r?\n\s*/g, " ").trim();
    const head = ["---", `name: ${oneLine(skill?.name)}`];
    if (oneLine(skill?.description)) head.push(`description: ${oneLine(skill.description)}`);
    head.push("---", "");
    return head.join("\n") + String(skill?.content || "").trim() + "\n";
  }

  // 用编辑后的 SKILL.md 全文更新一条用户技能。正文为空抛错；id 不是用户技能返回 null。
  function saveSkillMarkdown(id, text) {
    const parsed = parseMarkdownSkill(text, "");
    if (!parsed.content) throw new Error("SKILL.md 正文不能为空");
    return updateUser(id, parsed, String(text));
  }

  // 新建或覆盖用户技能的一个参考文件。校验不过抛错（路径非法 / 非文本 / 超限等）。
  function writeFile(id, path, content) {
    const rel = normalizeFilePath(path);
    const text = typeof content === "string" ? content : "";
    if (!rel) throw new Error("文件路径无效");
    if (SKILL_MD_RE.test(rel)) throw new Error("SKILL.md 是技能正文，不能作为参考文件新建");
    if (/^scripts\//i.test(rel)) throw new Error("不支持 scripts/ 目录（脚本无法执行）");
    if (!TEXT_FILE_RE.test(rel)) throw new Error("只支持文本文件（.md / .txt / .csv / .json / .yaml / .xml / .html 等）");
    if (!text.trim()) throw new Error("文件内容不能为空");
    if (text.length > FILE_LIMITS.maxFileChars) throw new Error("文件过大");
    const entries = readUserSkills();
    const skill = entries.find((e) => e.id === id);
    if (!skill) throw new Error("只能编辑自定义技能");
    const files = Array.isArray(skill.files) ? skill.files : [];
    const idx = files.findIndex((f) => f.path.toLowerCase() === rel.toLowerCase());
    if (idx < 0 && files.length >= FILE_LIMITS.maxFiles) throw new Error(`参考文件最多 ${FILE_LIMITS.maxFiles} 个`);
    const others = files.reduce((sum, f, i) => (i === idx ? sum : sum + f.content.length), 0);
    if (others + text.length > FILE_LIMITS.maxTotalChars) throw new Error("参考文件总大小超过上限");
    if (idx >= 0) files[idx] = { path: files[idx].path, content: text };
    else files.push({ path: rel, content: text });
    skill.files = files;
    skill.ts = Date.now();
    writeUserSkills(entries);
    return skill;
  }

  // 删除用户技能的一个参考文件；删到一个不剩时去掉 files 字段。文件不存在返回 false。
  function deleteFile(id, path) {
    const rel = normalizeFilePath(path);
    const entries = readUserSkills();
    const skill = entries.find((e) => e.id === id);
    if (!skill || !rel || !Array.isArray(skill.files)) return false;
    const next = skill.files.filter((f) => f.path !== rel);
    if (next.length === skill.files.length) return false;
    if (next.length) skill.files = next; else delete skill.files;
    skill.ts = Date.now();
    writeUserSkills(entries);
    return true;
  }

  function genId() {
    return "sk-" + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  }

  // 导入用户技能。skill 可以来源于 .md/.txt 解析或直接 JS object。
  // skill: { name (必填), description?, content (必填), files? }
  function addUser(skill) {
    if (!skill?.name) throw new Error("name 必填");
    if (typeof skill.content !== "string" || !skill.content.trim()) {
      throw new Error("content 必填且非空");
    }
    const entries = readUserSkills();
    const disk = _backend !== "kv";
    const dir = disk ? dirNameFor(skill.name, entries.map((e) => e.dir)) : "";
    const saved = {
      id: disk ? diskId(dir) : genId(),
      name: String(skill.name).trim(),
      description: String(skill.description || "").trim(),
      content: String(skill.content).trim(),
      builtin: false,
      ts: Date.now()
    };
    const files = normalizeSkillFiles(skill.files);
    if (files.length) saved.files = files;
    if (disk) saved.dir = dir;
    writeUserSkills(entries.concat([saved]));
    return saved;
  }

  function removeUser(id) {
    const entries = readUserSkills();
    const filtered = entries.filter((e) => e.id !== id);
    writeUserSkills(filtered);
    // 同步清掉 enabled
    const set = readEnabledIds();
    if (set.delete(id)) writeEnabledIds(set);
  }

  // 更新一条用户技能（AI 用 save_skill 持续优化技能时走这里）。id 不在用户技能里 → 返回 null。
  // rawText：SKILL.md 原文（界面里直接编辑全文时传），不传则按 name/description/content 重新生成。
  // 磁盘技能改名会连目录一起改名，id 随之变化 —— 调用方要用返回值里的新 id。
  function updateUser(id, patch, rawText) {
    const entries = readUserSkills();
    const idx = entries.findIndex((e) => e.id === id);
    if (idx < 0) return null;
    const cur = entries[idx];
    if (patch && patch.name != null && String(patch.name).trim()) cur.name = String(patch.name).trim();
    if (patch && patch.description != null) cur.description = String(patch.description).trim();
    if (patch && patch.content != null && String(patch.content).trim()) cur.content = String(patch.content).trim();
    cur.ts = Date.now();
    if (cur.dir) {
      delete cur.raw;
      cur.raw = typeof rawText === "string" ? rawText : buildSkillMarkdown(cur);
      const others = entries.filter((e) => e !== cur).map((e) => e.dir);
      const wantDir = dirNameFor(cur.name, others);
      if (wantDir !== cur.dir) {
        const oldId = cur.id;
        cur._from = cur.dir;
        cur.dir = wantDir;
        cur.id = diskId(wantDir);
        const enabled = readEnabledIds();
        if (enabled.delete(oldId)) { enabled.add(cur.id); writeEnabledIds(enabled); }
      }
    }
    writeUserSkills(entries);
    return cur;
  }

  // markdown 解析：支持开头 frontmatter（--- name: ... description: ... ---）+ 剩余内容
  // 不带 frontmatter 时整篇当 content，name 用 # 标题或文件名
  function parseMarkdownSkill(text, filename) {
    const s = String(text || "");
    const fmMatch = s.match(/^---\s*\n([\s\S]+?)\n---\s*\n([\s\S]*)$/);
    let name = "";
    let description = "";
    let content = s;
    if (fmMatch) {
      const front = fmMatch[1];
      content = fmMatch[2];
      front.split(/\r?\n/).forEach((line) => {
        const m = line.match(/^([A-Za-z_-]+)\s*:\s*(.+)$/);
        if (!m) return;
        const k = m[1].toLowerCase();
        const v = m[2].trim().replace(/^["']|["']$/g, "");
        if (k === "name") name = v;
        else if (k === "description") description = v;
      });
    }
    if (!name) {
      // 找第一个 # 标题
      const h1 = content.match(/^#\s+(.+)$/m);
      if (h1) name = h1[1].trim();
    }
    if (!name && filename) {
      name = filename.replace(/\.(md|txt|markdown)$/i, "");
    }
    if (!name) name = "未命名技能";
    return { name, description, content: content.trim() };
  }

  // 拿"启用的技能"并把每条的真实 content 解析出来（contentPath 形式的会 fetch）。
  // 给 chat 路径用——它的 system prompt 是 async 构造的，所以这里 async 直接拿全文。
  // opts.host: 当前宿主名（wps/wpp/et/pdf）。skill.hostFilter 设了且不包含此 host → 跳过
  async function getEnabledSkillsWithContent(opts) {
    const host = opts?.host;
    const skills = getEnabledSkills().filter((s) => {
      if (!Array.isArray(s.hostFilter) || !s.hostFilter.length) return true;
      return !host || s.hostFilter.includes(host);
    });
    const out = [];
    for (const s of skills) {
      const content = await loadContent(s);
      if (!content) continue;
      out.push({ id: s.id, name: s.name, description: s.description, content });
    }
    return out;
  }

  global.WpsAiSkills = {
    list,
    get,
    isEnabled,
    setEnabled,
    getEnabledSkills,
    getEnabledSkillsWithContent,
    addUser,
    updateUser,
    removeUser,
    init,
    refresh,
    whenSaved,
    backend: () => _backend,
    parseMarkdownSkill,
    buildSkillFromFiles,
    listFiles,
    readFile,
    writeFile,
    deleteFile,
    buildFileTree,
    buildSkillMarkdown,
    saveSkillMarkdown,
    loadContent,
    loadCloud,
    parseCloudIndex,
    readCloudSkills,
    _builtinIds: BUILTIN.map((b) => b.id)
  };
})(window);
