/**
 * AI 工作期间的文档硬锁定。
 *
 * 三个宿主各有"最强能用"的锁定方式：
 *
 * - Word（WPS 文字）
 *     Document.Protect(wdAllowOnlyReading=3, NoReset, password, false, false)
 *     副作用：COM 修改也被拦。所以每次工具调用前要 Unprotect → 改 → 再 Protect。
 *     由 tempUnlock(fn) 包装。
 *
 * - Excel（WPS 表格）
 *     ActiveSheet.Protect(password, ..., UserInterfaceOnly=true)
 *     UserInterfaceOnly=true 是关键：UI 锁住但 COM/VBA 仍可改。一次锁住，
 *     工具调用全部直接走，不需要解-改-再锁。
 *
 * - PowerPoint（WPS 演示）
 *     没有原生 UI 锁。降级到 Application.Interactive=false（部分版本支持）
 *     + 选区轮询警告（在 app.js）。
 *
 * 失败一律 try/catch 静默；解锁也要在 finally 路径里能跑（unlock 永远尝试
 * 还原所有曾经动过的状态）。
 */
(function attachDocLock(global) {
  "use strict";

  // wdProtectionType
  const WD_NO_PROTECTION = -1;
  const WD_ALLOW_ONLY_REVISIONS = 0; // 只允许修订：编辑强制记为原生修订，COM 仍可直接改
  const WD_ALLOW_ONLY_READING = 3;

  // 固定 token —— 之前用随机密码,unprotect 万一失败就永远卡在被锁状态,
  // 用户也不知道密码(随机只存在内存里,plugin reload 就丢)。
  // 固定 token 的代价是任何拿到源码的人都能解锁——但这本来就是 UI 锁不是密码学保护,
  // 收益是出错时还能自救:下次启动看见残留保护 → 用 token 一解就解开了。
  const LOCK_TOKEN = "lingxi-ai-doc-lock-v1";

  // 修 B33：Word/Excel 的 Unprotect(password) 在"目标保护本身没设密码"时，传任何密码都成功。
  // 因此不能用"Unprotect 成功"判定是不是我们加的锁——用户用无密码"限制编辑"的文档会被误判、
  // 误解、再永久移除。改用一个文档自定义属性作为"这是志特 AI 加的锁"的标记，只有带标记的
  // 保护才允许我们解除；用户自己的保护（无标记）一律不碰。
  const LOCK_MARKER = "LingxiAiLock";

  function hasOurMarker(doc) {
    try {
      const props = doc.CustomDocumentProperties;
      const p = props?.Item ? props.Item(LOCK_MARKER) : null;
      return !!p && String(p.Value) === "1";
    } catch (e) { return false; }
  }

  function setOurMarker(doc, on) {
    try {
      const props = doc.CustomDocumentProperties;
      if (!props) return;
      let existing = null;
      try { existing = props.Item(LOCK_MARKER); } catch (e) { existing = null; }
      if (on) {
        if (existing) existing.Value = "1";
        else if (typeof props.Add === "function") props.Add(LOCK_MARKER, false, 4 /* msoPropertyTypeString */, "1");
      } else if (existing && typeof existing.Delete === "function") {
        try { existing.Delete(); } catch (e) {}
      }
    } catch (e) {}
  }

  let state = null;

  function getApp() {
    return global.wps?.WpsApplication?.()
      || global.wps?.EtApplication?.()
      || global.wps?.WppApplication?.()
      || global.wps?.Application
      || global.Application
      || null;
  }

  function setInteractive(app, v) {
    try {
      if ("Interactive" in app) {
        const prev = app.Interactive;
        app.Interactive = v;
        return prev;
      }
    } catch (e) {}
    return null;
  }

  function lockWord(app, reviseMode) {
    const doc = app.ActiveDocument;
    if (!doc) return null;

    if (reviseMode) {
      // 修订模式：不加 Protect 硬锁（AI 改动本就以原生修订形式记录、由用户审阅；受保护反而挡住接受/拒绝修订）。
      // 但必须保证文档「可写」——用户开修订模式就是要让 AI 改。所以先尽量解掉任何残留保护：
      // 我们自己的锁 / 无密码"限制编辑"都用 token 解得开；用户设了密码的保护解不开（保持只读，AI 会据此报告）。
      // 这一步不做 marker 门控，专治"保护残留但 marker 丢了 → 一直只读、写入静默失败"的情况。
      try {
        if (doc.ProtectionType != null && doc.ProtectionType !== WD_NO_PROTECTION) {
          try { doc.Unprotect(LOCK_TOKEN); setOurMarker(doc, false); } catch (e) {}
        }
      } catch (e) {}
      try { doc.TrackRevisions = true; } catch (e) {}
      return { kind: "wps-revisions", doc };
    }

    // 非修订模式：检查 doc 是否已经被保护。如果是,先用我们的 token 试解 ——
    // 解开了说明是上次没清干净的残留锁(可继续);解不开是用户/别处加的锁,不动它。
    try {
      if (doc.ProtectionType != null && doc.ProtectionType !== WD_NO_PROTECTION) {
        // 修 B33：只有带我们标记的保护才是"残留 AI 锁"，才解开重建；否则是用户自己的保护，不接手。
        if (hasOurMarker(doc)) {
          try { doc.Unprotect(LOCK_TOKEN); } catch (e) {}
        } else {
          console.warn("[doc-lock] 文档已被外部保护（无 AI 标记），跳过 AI 锁定");
          return { kind: "wps-already-protected", doc };
        }
      }
    } catch (e) {}
    try {
      doc.Protect(WD_ALLOW_ONLY_READING, true, LOCK_TOKEN, false, false);
      setOurMarker(doc, true);
      return { kind: "wps-protect", doc };
    } catch (e) {
      console.error("[doc-lock] Word Protect 失败:", e?.message || e);
      return null;
    }
  }

  function unlockWord(lockInfo) {
    if (!lockInfo) return;
    if (lockInfo.kind === "wps-already-protected") return; // 不是我们加的，不动
    if (lockInfo.kind === "wps-revisions") return;         // 修订模式没加 Protect，无需 Unprotect
    try { lockInfo.doc.Unprotect(LOCK_TOKEN); setOurMarker(lockInfo.doc, false); }
    catch (e) { console.error("[doc-lock] Word Unprotect 失败:", e?.message || e); }
  }

  // 启动时主动尝试用 token 解一次：如果上次 AI session 没干净清掉保护,这里能自救
  function clearStaleWordLock() {
    try {
      const app = getApp();
      const doc = app?.ActiveDocument;
      if (!doc) return;
      if (doc.ProtectionType != null && doc.ProtectionType !== WD_NO_PROTECTION) {
        // 修 B33：只清带我们标记的残留锁，绝不碰用户自己的无密码保护。
        if (hasOurMarker(doc)) {
          try {
            doc.Unprotect(LOCK_TOKEN);
            setOurMarker(doc, false);
            console.log("[doc-lock] 清理了上次残留的 AI 锁定");
          } catch (e) { /* 忽略 */ }
        }
      }
    } catch (e) {}
  }

  // 用户手动触发的强制解锁：内存 state + 残留保护一起清。
  // 返回 { word, sheet, interactive, hadLock } —— word/sheet=true 表示真的把保护解掉了。
  // 始终用 LOCK_TOKEN 尝试 Unprotect，token 不匹配（用户自己加的密码保护）时不会越权解开。
  function forceUnlock() {
    const result = { word: false, sheet: false, interactive: false, hadLock: !!state };
    try { unlock(); } catch (e) {}
    const app = getApp();
    if (!app) return result;
    try {
      const doc = app.ActiveDocument;
      // 修 B33：只解带我们标记的 AI 锁，不移除用户自己的（无密码）保护。
      if (doc && doc.ProtectionType != null && doc.ProtectionType !== WD_NO_PROTECTION && hasOurMarker(doc)) {
        try { doc.Unprotect(LOCK_TOKEN); setOurMarker(doc, false); result.word = true; }
        catch (e) { /* 忽略 */ }
      }
    } catch (e) {}
    try {
      const sheet = app.ActiveSheet;
      const wb = (() => { try { return app.ActiveWorkbook; } catch (e) { return null; } })();
      if (sheet && sheet.ProtectContents && wb && hasOurMarker(wb)) {
        try { sheet.Unprotect(LOCK_TOKEN); setOurMarker(wb, false); result.sheet = true; }
        catch (e) { /* 忽略 */ }
      }
    } catch (e) {}
    try {
      if ("Interactive" in app && app.Interactive === false) {
        app.Interactive = true;
        result.interactive = true;
      }
    } catch (e) {}
    return result;
  }

  function lockExcel(app) {
    // 【不再对 Excel 硬锁工作表】
    // 原来用 sheet.Protect(UserInterfaceOnly=true) 锁 UI。但只要用户在 AI 工作期间去碰表格，
    // WPS 就会弹出原生的「工作表受保护」对话框——它是 WPS 应用级模态框，会把 ribbon、表格、
    // 以及右侧 taskpane（含「停止」按钮、整个聊天区）全部挡死，用户必须先关掉它才能操作 /
    // 打断 AI。防用户误编辑这点好处，远抵不过「整个 WPS 被模态框冻住、连停止都点不了」的代价。
    // 改为不加保护：AI 仍通过 COM 正常写入；用户理论上能同时编辑（概率低、影响有限），
    // 换来全程 UI 不被冻、随时可打断。仍由聊天面板的「AI 工作中」banner 做软提示。
    //
    // 若当前 sheet 是用户自己加的保护（无我们的标记），保持原样、不接手、不解锁。
    try {
      const sheet = app.ActiveSheet;
      const wb = (() => { try { return app.ActiveWorkbook; } catch (e) { return null; } })();
      if (sheet && sheet.ProtectContents && !(wb && hasOurMarker(wb))) {
        return { kind: "et-already-protected", sheet };
      }
    } catch (e) {}
    return { kind: "et-noop" };
  }

  function unlockExcel(lockInfo) {
    if (!lockInfo) return;
    if (lockInfo.kind === "et-noop") return;              // 没加保护，无需解锁
    if (lockInfo.kind === "et-already-protected") return; // 不是我们加的，不动
    try {
      lockInfo.sheet.Unprotect(LOCK_TOKEN);
      if (lockInfo.wb) setOurMarker(lockInfo.wb, false);
    } catch (e) { console.error("[doc-lock] Excel Unprotect 失败:", e?.message || e); }
  }

  // 进入锁定。host 必须传准 ("wps" / "et" / "wpp")；opts.reviseMode → Word 完全不硬锁
  function lock(host, opts) {
    if (state) return state;
    const app = getApp();
    if (!app) return null;
    const reviseMode = host === "wps" && !!(opts && opts.reviseMode);
    // 【不再对整个 app 设 Interactive=false】
    // 它会把整个 WPS app（含右侧 taskpane）一起冻住：AI 输出整轮都点不了「停止」、滚不动聊天区，
    // 连纯思考（没调用任何工具）时也照样冻，直到回合结束 unlock 才恢复——这是用户反馈"发消息后
    // 点不动、输出完才能点"的真正根因。COM 写入在 Interactive=true 下正常工作（本就不依赖它）；
    // 防用户误操作文档改由聊天面板 banner 软提示 + Word 的 Document.Protect 承担。
    const prevInteractive = null;
    let docLock = null;
    if (host === "wps") {
      docLock = lockWord(app, reviseMode);
    } else if (host === "et") {
      docLock = lockExcel(app);
    }
    // wpp 没有靠谱的 UI 锁，跳过 docLock
    state = { app, host, prevInteractive, docLock };
    return state;
  }

  function unlock() {
    if (!state) return;
    const { app, host, prevInteractive, docLock } = state;
    if (host === "wps") unlockWord(docLock);
    else if (host === "et") unlockExcel(docLock);
    if (prevInteractive != null) {
      try { app.Interactive = prevInteractive; } catch (e) {}
    }
    state = null;
  }

  // 给 Word 用：临时解锁执行 fn，结束后立即再锁回去。
  // Excel 不用调（UserInterfaceOnly 已经允许 COM 写）。
  // 没锁 / PPT / 用户已有保护一律直接调 fn，不影响主流程。
  async function tempUnlock(fn) {
    if (!state || state.host !== "wps" || !state.docLock || state.docLock.kind !== "wps-protect") {
      return await fn();
    }
    let unlocked = false;
    let prevInteractive = null;
    try {
      state.docLock.doc.Unprotect(LOCK_TOKEN);
      unlocked = true;
    } catch (e) {
      console.error("[doc-lock] tempUnlock 时 Unprotect 失败,工具写入可能被拒:", e?.message || e);
    }
    try {
      if ("Interactive" in state.app) {
        prevInteractive = state.app.Interactive;
        if (prevInteractive === false) state.app.Interactive = true;
      }
    } catch (e) {}
    try {
      return await fn();
    } finally {
      if (prevInteractive != null) {
        try { state.app.Interactive = prevInteractive; } catch (e) {}
      }
      // 只有刚才确实 unlock 成功了才需要再 Protect 回去
      if (unlocked) {
        try {
          state.docLock.doc.Protect(WD_ALLOW_ONLY_READING, true, LOCK_TOKEN, false, false);
        } catch (e) { console.error("[doc-lock] tempUnlock 后 re-Protect 失败:", e?.message || e); }
      }
    }
  }

  function isLocked() { return !!state; }

  function isDocumentLocked() {
    if (!state || !state.docLock) return false;
    return state.docLock.kind === "wps-protect" || state.docLock.kind === "et-protect-sheet";
  }

  // 审批弹窗要用户点面板，但 lock() 期间 app.Interactive=false 会把整个 WPS（含任务窗格）
  // 的交互一起禁掉，导致"确认框弹出来什么都点不了"。这里在已锁定时临时把交互开回来（true）
  // 或再关掉（false）。文档本身仍由 Protect 守着，不影响写入安全。未锁定时无操作。
  function setInteractiveLive(v) {
    if (!state || !state.app) return;
    try { if ("Interactive" in state.app) state.app.Interactive = !!v; } catch (e) {}
  }

  // 启动时 / Pane 重启时主动清掉残留的 AI 锁定（如果有）。AI 锁定的 token 是固定的,
  // Unprotect 用 token 试一下:成功说明是我们之前没清干净,清掉即可;失败是用户加的锁,不动。
  function cleanupStaleLocks() {
    clearStaleWordLock();
    // Excel sheet 残留锁的清理略复杂(要遍历所有 sheet),先不做,等用户报告再补
  }

  // 自动启动时跑一次 cleanup
  setTimeout(cleanupStaleLocks, 1500);

  global.WpsAiLock = { lock, unlock, tempUnlock, isLocked, isDocumentLocked, setInteractiveLive, cleanupStaleLocks, forceUnlock, LOCK_TOKEN };
})(window);
