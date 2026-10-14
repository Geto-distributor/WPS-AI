(function registerRefDocTools(global) {
  "use strict";

  // 参考文档工具：读取用户添加的参考资料（以及 WPS 里其它已打开的文档），用来评审 / 撰写当前文档。
  // 读取方式由 WpsAiRefDocs 自动选：已在 WPS 打开就委托对应组件的插件读（含未保存修改），否则读磁盘。
  const registry = global.WpsAiToolRegistry;
  if (!registry) return;
  const refs = () => {
    if (!global.WpsAiRefDocs) throw new Error("参考文档模块未加载。");
    return global.WpsAiRefDocs;
  };
  const STATUS_LABEL = { wps: "已在 WPS 打开（读实时内容）", disk: "读取磁盘文件" };

  registry.registerTool({
    name: "list_reference_docs",
    hosts: ["*"],
    description: "列出参考文档：用户为本次对话添加的参考资料（added），以及 WPS 里当前打开着、但没被添加的其它文档（otherOpenDocs）。评审 / 对照 / 引用资料前先调它看有哪些可读。",
    parameters: { type: "object", properties: {} },
    handler: async () => {
      const R = refs();
      const open = await R.listOpenDocs();
      const openByPath = new Map((open.docs || []).map((d) => [R.normPath(d.path), d]));
      const added = R.list().map((r) => {
        const o = openByPath.get(R.normPath(r.path));
        const viaWps = !!(o && o.host !== "pdf" && open.online?.[o.host]);
        return { name: r.name, path: r.path, kind: r.kind, readVia: STATUS_LABEL[viaWps ? "wps" : "disk"], unsavedChanges: o ? o.saved === false : undefined };
      });
      const addedSet = new Set(added.map((a) => R.normPath(a.path)));
      const otherOpenDocs = (open.docs || [])
        .filter((d) => d.path && !addedSet.has(R.normPath(d.path)) && !d.current)
        .map((d) => ({ name: d.name, path: d.path, component: d.host, unsavedChanges: d.saved === false }));
      return { added, otherOpenDocs, note: open.ok === false ? `暂时拿不到 WPS 已打开文档列表：${open.error}` : undefined };
    }
  });

  registry.registerTool({
    name: "read_reference_doc",
    hosts: ["*"],
    description: [
      "读取一份参考文档的内容（Word / Excel / PPT / PDF / 文本），按段返回。",
      "path 取自 list_reference_docs；内容长时用返回的 nextOffset 作为下次的 offset 继续读，或用 section 直接跳到某节（如「工作表「清单」」「第 3 页」）。",
      "返回 sections 是全部分节标题和起始位置，可据此只读需要的部分；Excel 每行开头是行号，列用制表符分隔。"
    ].join(""),
    parameters: {
      type: "object",
      required: ["path"],
      properties: {
        path: { type: "string", description: "文档完整路径（来自 list_reference_docs）" },
        offset: { type: "number", description: "从第几个字符开始读，默认 0" },
        maxChars: { type: "number", description: "本次最多读多少字符，默认 12000，最大 60000" },
        section: { type: "string", description: "直接跳到这一节的开头（节名取自 sections），优先于 offset" }
      }
    },
    handler: async ({ path, offset, maxChars, section } = {}) => {
      const R = refs();
      const p = String(path || "").trim();
      if (!p) throw new Error("需要 path（先调 list_reference_docs 看可读的文档）。");
      // 只允许读：用户添加的参考文档，或 WPS 里确实开着的文档 —— 不让模型凭空读磁盘上的任意文件
      const isAdded = R.list().some((r) => R.normPath(r.path) === R.normPath(p));
      if (!isAdded) {
        const open = await R.listOpenDocs();
        if (!(open.docs || []).some((d) => R.normPath(d.path) === R.normPath(p))) {
          throw new Error("只能读取已添加的参考文档或 WPS 里已打开的文档。请先调 list_reference_docs。");
        }
      }
      const r = await R.read({ path: p, offset, maxChars, section });
      return {
        name: r.name,
        path: r.path,
        readVia: STATUS_LABEL[r.source === "wps" ? "wps" : "disk"],
        kind: r.kind,
        offset: r.offset,
        nextOffset: r.nextOffset,
        total: r.total,
        sections: r.sections,
        note: r.note,
        text: r.text
      };
    }
  });
})(window);
