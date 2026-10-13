/**
 * AI 改动记录：每次 AI 调用写入型工具，把"改了什么"记下来。
 *
 * 数据条目结构：
 *   {
 *     id: 'h-1683...',           // 唯一 id
 *     ts: 1683830400000,         // 时间戳 ms
 *     host: 'wps'|'et'|'wpp',
 *     toolName: 'wpp_apply_template',
 *     friendlyName: '套用 PPT 模板',  // 人话名
 *     target: { kind, label },         // { kind:'slide', label:'第 3 页' }
 *     params: {...},                   // 工具入参（裁剪后）
 *     before: {...} | null,            // 调用前快照
 *     after:  {...} | null,            // 调用后快照
 *     ok: true | false,
 *     resultSummary: '...',            // 结果一两句话
 *     error: null | '...'
 *   }
 *
 * 存储：localStorage key 'lingxi_history_v1'，限 MAX_ENTRIES 条，FIFO 淘汰。
 * 序列化时单条 > 64KB 的 before/after 会被截断，避免占满 storage。
 */
(function attachHistory(global) {
  "use strict";

  const STORAGE_KEY = "lingxi_history_v1";
  const MAX_ENTRIES = 200;
  const MAX_SNAPSHOT_BYTES = 64 * 1024;

  // 内存里的条目列表（按写入顺序，最新在末尾）
  let entries = [];
  // 订阅者，UI 用来重渲染
  const listeners = new Set();

  // ---- 持久化 ----

  function load() {
    try {
      const raw = global.WpsAiStore.getItem(STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) entries = parsed;
    } catch (e) {
      console.warn("[history] load failed", e);
      entries = [];
    }
  }

  // 排序（ts 升序）+ 限 MAX 条尾部保留 —— 合并后统一收口本地视图。
  function capEntries(arr) {
    const sorted = (Array.isArray(arr) ? arr.slice() : []).sort((a, b) => (a.ts || 0) - (b.ts || 0));
    return sorted.length > MAX_ENTRIES ? sorted.slice(-MAX_ENTRIES) : sorted;
  }

  // 修 Critical 1：4 个宿主（wps/et/wpp/pdf）共用同一 SQLite。把"读-合并-写"放服务端原子完成，
  // 前端交给 WpsAiStore.mergeList（sqlite → POST /kv/merge-list；降级 → 本地客户端合并 + setItem），
  // 避免同时打开的宿主互相冲掉对方的历史条目。fire-and-forget：服务端合并原子，安全。
  function persist() {
    let p;
    try { p = global.WpsAiStore.mergeList(STORAGE_KEY, entries, "id", "ts"); } catch (e) { return; }
    Promise.resolve(p).then((arr) => {
      if (Array.isArray(arr)) {
        const capped = capEntries(arr);
        entries = capped;
        if (capped.length !== arr.length) persistExact();
        notify();
      }
    }).catch((e) => {
      // localStorage 满了 → 砍一半重试
      if (e?.name === "QuotaExceededError" && entries.length > 20) {
        entries = entries.slice(-Math.floor(entries.length / 2));
        try { global.WpsAiStore.setItem(STORAGE_KEY, JSON.stringify(entries)); } catch (_) {}
      }
    });
  }

  // mergeList is upsert-only. Deletes and MAX_ENTRIES pruning need exact replacement.
  function persistExact() {
    try {
      entries = capEntries(entries);
      global.WpsAiStore.setItem(STORAGE_KEY, JSON.stringify(entries));
    } catch (e) {}
  }

  function truncateSnapshot(snap) {
    if (!snap) return snap;
    try {
      const s = JSON.stringify(snap);
      if (s.length <= MAX_SNAPSHOT_BYTES) return snap;
      return { _truncated: true, _originalBytes: s.length, _excerpt: s.slice(0, MAX_SNAPSHOT_BYTES) + "…" };
    } catch (e) {
      return { _error: "snapshot 不可序列化" };
    }
  }

  function notify() {
    listeners.forEach((fn) => { try { fn(); } catch (e) {} });
  }

  // ---- 公共 API ----

  function addEntry(entry) {
    const id = "h-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 6);
    // 文档身份：docId 优先，docPath 兜底。ensureDocId 会在必要时写一次 UUID 到 doc
    // 的 CustomDocumentProperties；已有 UUID 就直接返回。写不成功（PDF / 不支持宿主）
    // 就退回按路径记录。
    const backup = global.WpsAiBackup;
    const docPath = entry.docPath || backup?.getCurrentDocPath?.() || null;
    const docId = entry.docId || backup?.ensureDocId?.() || null;
    const full = Object.assign(
      { id, ts: Date.now(), turnId: currentTurn?.id || null, docPath, docId },
      entry,
      {
        before: truncateSnapshot(entry.before),
        after: truncateSnapshot(entry.after)
      }
    );
    entries.push(full);
    let pruned = false;
    if (entries.length > MAX_ENTRIES) {
      entries = entries.slice(-MAX_ENTRIES);
      pruned = true;
    }
    // 迁移：这条 entry 拿到了 docId，把之前只有 docPath 的老条目也回填 docId。
    // 这样切走 -> Save As / 重命名 -> 切回来后老历史仍能匹配上。
    if (docId && docPath) backfillDocIdByPath(docPath, docId);
    if (pruned) persistExact();
    else persist();
    notify();
    return full;
  }

  // 老条目只有 docPath，没有 docId。第一次给某个 path 分配到 docId 时，把这个映射
  // 回写到所有匹配的老条目上，避免"迁移前"的历史孤立。同名 path 一律吃同 id（用户
  // 视角就是"这个文件"）。
  function backfillDocIdByPath(docPath, docId) {
    let touched = 0;
    entries.forEach((e) => {
      if (!e.docId && pathsEqual(e.docPath, docPath)) {
        e.docId = docId;
        touched += 1;
      }
    });
    if (touched > 0) persist();
    // turns 表里 backup.docPath 也可能存在，顺手带上
    let turnsTouched = 0;
    Object.values(turns).forEach((t) => {
      if (t?.backup && !t.backup.docId && pathsEqual(t.backup.docPath, docPath)) {
        t.backup.docId = docId;
        turnsTouched += 1;
      }
    });
    if (turnsTouched > 0) persistTurns();
  }

  // ===== Turn 概念：一次 AI 对话视为一个 turn =====
  // app.js 在 runChatTurn 入口调 startTurn；execute 触发第一个修改型工具时 lazy 抓 backup。

  let currentTurn = null;           // 内存中的"当前 turn"
  let turns = {};                   // 已结束的 turn 索引，持久化

  function loadTurns() {
    try {
      const raw = global.WpsAiStore.getItem("lingxi_history_turns_v1");
      return raw ? JSON.parse(raw) : {};
    } catch (e) { return {}; }
  }

  // 修 Critical 1：turns 索引是 {turnId: turn} 对象，多宿主共用同一 SQLite。用服务端原子
  // mergeObject(assign)：patch（本宿主 turns）逐 key 覆盖磁盘上别的宿主写入的 turn，其余保留。
  // 拿回权威合并结果换掉本地 turns。fire-and-forget。
  function persistTurns() {
    let p;
    try { p = global.WpsAiStore.mergeObject("lingxi_history_turns_v1", turns, "assign"); } catch (e) { return; }
    Promise.resolve(p).then((obj) => { if (obj && typeof obj === "object") { turns = obj; notify(); } }).catch(() => {});
  }

  // mergeObject(assign) cannot express deleted object keys.
  function persistTurnsExact() {
    try {
      global.WpsAiStore.setItem("lingxi_history_turns_v1", JSON.stringify(turns || {}));
    } catch (e) {}
  }

  function newTurnId() {
    return "t-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 6);
  }

  // 用户按下「发送」时调用，把上一个 turn 收尾，开新的
  function startTurn(prompt) {
    if (currentTurn && currentTurn.id) {
      // 收掉上一 turn 的 UndoRecord(如果开过的话)。这样下次 Application.Undo
      // 一次性撤回上一 turn 整组改动。不收的话 UndoRecord 一直开着,新 turn 的
      // 改动会被合并进去 — 撤回时连新 turn 也跟着撤了。
      try { global.WpsAiBackup?.endUndoGroup?.(); } catch (e) {}
      turns[currentTurn.id] = currentTurn;
      persistTurns();
    }
    // 记一下开 turn 时候的文档身份，方便渲染时对齐 —— captureCurrentDoc 里也会补 docId。
    const openDocId = global.WpsAiBackup?.readDocId?.() || null;
    const openDocPath = global.WpsAiBackup?.getCurrentDocPath?.() || null;
    currentTurn = {
      id: newTurnId(),
      startedAt: Date.now(),
      prompt: prompt ? String(prompt).slice(0, 200) : "",
      docId: openDocId,      // may be null；ensureBackupForTurn 会补上
      docPath: openDocPath,
      backup: null      // { docPath, docId, backupPath, size, ts, undoGroup } 由 ensureBackupForTurn 填
    };
    notify();
    return currentTurn.id;
  }

  // 修改型工具调用前调一下；同一个 turn 只会 capture 一次
  async function ensureBackupForTurn() {
    if (!currentTurn) return null;
    if (currentTurn.backup) return currentTurn.backup;  // 已经抓过
    const backup = global.WpsAiBackup;
    if (!backup) return null;
    try {
      const res = await backup.captureCurrentDoc();
      if (res?.ok) {
        currentTurn.backup = {
          docPath: res.docPath,
          docId: res.docId || null,
          backupPath: res.backupPath,
          size: res.size,
          ts: res.timestamp || Date.now(),
          // 是否启动了 UndoRecord。回退时优先走 Application.Undo,失败再走文件层。
          undoGroup: !!res.undoGroup
        };
        // 顺手把 currentTurn 顶层的 docId / docPath 也补齐
        if (res.docId && !currentTurn.docId) currentTurn.docId = res.docId;
        if (res.docPath && !currentTurn.docPath) currentTurn.docPath = res.docPath;
        // 修 B36：备份路径已拿到，立即把 currentTurn 落盘。否则要等下一次 startTurn 才写入，
        // 期间用户关闭/刷新 TaskPane 或 WPS 崩溃，这个 turn 的回滚入口（backupPath）就丢了——
        // 恰恰是"AI 改坏了文档、重启回滚"最需要恢复的场景。
        turns[currentTurn.id] = currentTurn;
        persistTurns();
        notify();
        return currentTurn.backup;
      }
      // 没存盘的新文档之类失败原因，记下不再重试
      currentTurn.backup = { error: res?.error || "备份失败", ts: Date.now() };
      notify();
      return null;
    } catch (e) {
      currentTurn.backup = { error: e?.message || String(e), ts: Date.now() };
      return null;
    }
  }

  // 历史 turn 加上当前 turn 一起返回
  function listTurns() {
    const out = Object.assign({}, turns);
    if (currentTurn) out[currentTurn.id] = currentTurn;
    return out;
  }

  function getCurrentTurnId() { return currentTurn?.id || null; }

  function deleteTurn(turnId) {
    delete turns[turnId];
    persistTurnsExact();
    entries = entries.filter((e) => e.turnId !== turnId);
    persistExact();
    notify();
  }

  // 标记某 turn 已被恢复 —— 不删 entry，加 restoredAt 字段，UI 自己渲染"已恢复"徽章。
  // 之前用 deleteTurn 收尾恢复操作会把 entry 全删掉，重启 TaskPane 后看不到历史；现在保留可查。
  // currentTurn 被恢复时一并提前 flush 到 turns 表，保证 reload 后还能拿到这条记录。
  function markTurnRestored(turnId) {
    if (!turnId) return;
    const stamp = Date.now();
    if (turns[turnId]) {
      turns[turnId].restoredAt = stamp;
    } else if (currentTurn?.id === turnId) {
      currentTurn.restoredAt = stamp;
      turns[turnId] = Object.assign({}, currentTurn);
    } else {
      // 既不在 turns 表也不是 currentTurn —— 兜底建一条最小记录
      turns[turnId] = { id: turnId, restoredAt: stamp };
    }
    persistTurns();
    notify();
  }

  // 倒序返回（最新在前）。可选 filter：
  //   filter.docId   —— 优先：按文档 UUID 匹配（跨重命名 / Save As 稳）
  //   filter.docPath —— 退路：docId 空时按路径匹配
  // 同时传两者时按 "OR" 语义：任一命中即算这个文档的记录，帮助覆盖迁移期混合数据。
  function listEntries(filter) {
    let out = entries.slice();
    if (filter && (filter.docId || filter.docPath)) {
      out = out.filter((e) => entryMatchesDoc(e, filter));
    }
    return out.reverse();
  }

  // 判断一条 entry 是否属于某个文档。docId 一致优先，其次退到 docPath 相等。
  // filter.docId 存在但 entry.docId 空 → 用 pathsEqual 兜底（老数据迁移期）。
  function entryMatchesDoc(entry, filter) {
    if (filter.docId && entry.docId && entry.docId === filter.docId) return true;
    if (filter.docPath && pathsEqual(entry.docPath, filter.docPath)) return true;
    return false;
  }

  // 跨平台路径比较（大小写不敏感 + 反斜杠归一化）
  function pathsEqual(a, b) {
    if (!a || !b) return false;
    const norm = (s) => String(s).replace(/\\/g, "/").toLowerCase().replace(/\/+$/, "");
    return norm(a) === norm(b);
  }

  // 返回所有出现过的 docPath 集合（UI 切换 / 调试用）
  function listDocPaths() {
    const set = new Set();
    entries.forEach((e) => { if (e.docPath) set.add(e.docPath); });
    return Array.from(set);
  }

  // 返回 docId → 最新 docPath 的映射（UI 切换 / 调试）
  function listDocIds() {
    const out = new Map();
    entries.forEach((e) => {
      if (e.docId) {
        const prev = out.get(e.docId);
        if (!prev || (e.ts || 0) > (prev.ts || 0)) {
          out.set(e.docId, { docPath: e.docPath || prev?.docPath || "", ts: e.ts || 0 });
        }
      }
    });
    return out;
  }

  function clear() {
    entries = [];
    persistExact();
    notify();
  }

  function size() { return entries.length; }

  function subscribe(fn) {
    listeners.add(fn);
    return () => listeners.delete(fn);
  }

  // ---- 工具名 → 人话映射 ----

  const FRIENDLY_NAMES = {
    // 演示
    wpp_apply_template: "套用 PPT 模板（A 类）",
    wpp_apply_visual_template: "套用 PPT 视觉模板（B 类）",
    wpp_render_chart: "插入数据图表",
    wpp_add_slide: "添加幻灯片",
    wpp_delete_slide: "删除幻灯片",
    wpp_duplicate_slide: "复制幻灯片",
    wpp_move_slide: "移动幻灯片",
    wpp_replace_shape_text: "替换形状文字",
    wpp_set_title: "设置标题",
    wpp_set_notes: "设置演讲者备注",
    wpp_add_text_box: "添加文本框",
    wpp_add_shape: "添加形状",
    wpp_add_picture: "添加图片",
    wpp_set_slide_background: "设置幻灯片背景",
    wpp_set_slide_layout: "切换版式",
    wpp_select_slide: "选中幻灯片",
    wpp_set_slide_transition: "设置切换动画",
    wpp_apply_style_preset: "统一应用风格预设",
    // 表格
    et_write_range: "写入单元格区域",
    et_apply_style: "应用表格样式",
    et_autofit_columns: "自动调整列宽",
    et_autofit_rows: "自动调整行高",
    et_clear_range: "清空单元格区域",
    et_set_formula: "设置公式",
    et_insert_image: "插入图片",
    // 文字
    wps_replace_selection: "替换选区内容",
    wps_insert_text: "插入文字",
    wps_insert_image: "插入图片",
    wps_apply_paragraph_style: "应用段落样式",
    wps_insert_paragraph: "插入段落",
    wps_remove_paragraph: "删除段落",
    // 通用
    generate_image: "生成图片",
    use_skill: "调用技能",
    read_skill_file: "读取技能参考文件",
    save_skill: "保存 / 优化技能"
  };

  function getFriendlyName(toolName) {
    return FRIENDLY_NAMES[toolName] || toolName;
  }

  // 哪些工具会"改动文档"——只有这些才记录
  // 读取型 / 信息查询型不记录（防噪音）
  function isMutatingTool(toolName) {
    if (!toolName) return false;
    // 黑名单关键字：所有以 _get_ / _list_ / _read_ 开头的视为只读
    if (/(^|_)(get|list|read)_/.test(toolName)) return false;
    if (/^wpp_get_/.test(toolName)) return false;
    // 显式只读（不改文档：查素材 / 联网抓取 / 生图。之前被默认判为修改型，会在未保存的新文档上被
    // registry.execute 的"先存盘"拦截误挡，见 bug M5）
    const readonly = new Set([
      "wpp_list_slides", "wpp_read_slide", "wpp_get_notes",
      "wpp_get_presentation_info", "wpp_get_style_preset",
      "query_materials", "web_fetch", "web_image_search", "generate_image",
      "reveal_location", // 只定位/滚动/高亮，不改文档——不进 history、不触发"先存盘"拦截
      "use_skill", // 只加载技能指引文本给 AI，不改文档
      "save_skill", // 沉淀/优化技能到本地技能库，不改文档
      "wps_find_colored_text", // 只扫描颜色/高亮/底纹，不改文档（名字非 read_/get_/list_，需显式列为只读）
      "wps_export_pdf", "et_export_pdf", "wpp_export_pdf", // 导出另存 PDF，不改原文档内容——不必快照/记录
      // 另存/打印/文档属性：不改文档正文（属性改动内容快照也捕获不到），跳过快照避免刷无意义的改动记录
      "wps_save_as", "et_save_as", "wpp_save_as",
      "wps_print", "et_print", "wpp_print",
      "wps_doc_properties", "et_doc_properties", "wpp_doc_properties",
      "wps_set_view", "et_set_view", "wpp_set_view" // 只调缩放/定位，不改文档
    ]);
    if (readonly.has(toolName)) return false;
    // 其余默认为修改型
    return true;
  }

  // 模块脚本解析时（top-level）就会跑到这一段，早于 app.js boot 里 WpsAiStore.init()
  // 把 sqlite 数据灌进内存 Map —— 那时候读到的必然是空表。暴露 reloadFromStore()，
  // 让 app.js 在 init() 完成后再补一次读，把 entries/turns 换成真实数据。
  function reloadFromStore() {
    load();
    turns = loadTurns();
  }
  reloadFromStore();

  global.WpsAiHistory = {
    addEntry,
    listEntries,
    entryMatchesDoc,
    listDocPaths,
    listDocIds,
    pathsEqual,
    clear,
    size,
    subscribe,
    getFriendlyName,
    isMutatingTool,
    // turn 管理
    startTurn,
    ensureBackupForTurn,
    listTurns,
    getCurrentTurnId,
    deleteTurn,
    markTurnRestored,
    reloadFromStore,
    MAX_ENTRIES
  };
})(window);
