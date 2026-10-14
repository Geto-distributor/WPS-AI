"use strict";
// 参考文档的「在线登记 + 委托代读」中转。
//
// WPS 的文字 / 表格 / 演示 / PDF 各自加载一份插件，彼此看不到对方打开的文档（实测：文字窗口里
// wps.EtApplication() 等都返回空）。每份插件只能读自己这一类的文档，于是由本模块当中间人：
//   1. 登记：每份插件定时报「我是哪个组件、开着哪些文档」→ 汇总成一张已打开文档总表
//   2. 代读：文字窗口要读一个开在表格里的文件 → 把请求排给表格那份插件（它在长轮询等活）→ 结果原路带回
// 纯内存、无 IO，便于单测；HTTP 路由在 proxy-server.js。
const crypto = require("crypto");

function normPath(p) {
  return String(p || "").replace(/\\/g, "/").replace(/\/+$/, "").toLowerCase();
}

function createRefRelay(opts = {}) {
  const now = opts.now || (() => Date.now());
  const staleMs = opts.staleMs || 15000;      // 超过这么久没登记的插件实例视为已下线
  const callTimeoutMs = opts.callTimeoutMs || 20000;
  const pollTimeoutMs = opts.pollTimeoutMs || 20000;

  const instances = new Map(); // instanceId → { host, docs, at, polledAt }
  const queues = new Map();    // host → [call]
  const waiters = new Map();   // host → [{ instanceId, resolve, timer }]
  const pending = new Map();   // callId → { resolve, reject, timer, host }

  function fresh(inst) { return now() - Math.max(inst.at || 0, inst.polledAt || 0) < staleMs; }

  function report({ instanceId, host, docs }) {
    if (!instanceId || !host) throw new Error("instanceId / host 必填");
    const list = (Array.isArray(docs) ? docs : []).slice(0, 200).map((d) => ({
      name: String(d && d.name || ""),
      path: String(d && d.path || ""),
      saved: d && d.saved !== false,
      active: !!(d && d.active)
    })).filter((d) => d.name || d.path);
    const prev = instances.get(instanceId);
    instances.set(instanceId, { host: String(host), docs: list, at: now(), polledAt: prev ? prev.polledAt : 0 });
  }

  // 全部组件的已打开文档，按路径去重（同一组件的面板页和后台页可能各报一份）
  function listOpenDocs() {
    const out = [];
    const seen = new Set();
    instances.forEach((inst, id) => {
      if (!fresh(inst)) { instances.delete(id); return; }
      inst.docs.forEach((d) => {
        const key = d.path ? normPath(d.path) : `${inst.host}:${d.name}`;
        if (seen.has(key)) return;
        seen.add(key);
        out.push(Object.assign({ host: inst.host }, d));
      });
    });
    return out;
  }

  // 某个组件当前有没有正在等活的插件实例（能不能委托代读）
  function workerOnline(host) {
    let ok = false;
    instances.forEach((inst) => { if (inst.host === host && inst.polledAt && now() - inst.polledAt < staleMs) ok = true; });
    return ok || ((waiters.get(host) || []).length > 0);
  }

  function hand(host) {
    const q = queues.get(host) || [];
    const w = waiters.get(host) || [];
    while (q.length && w.length) {
      const waiter = w.shift();
      clearTimeout(waiter.timer);
      waiter.resolve(q.shift());
    }
  }

  // 委托：把请求排给 host 组件的插件，等它回结果；超时 reject
  function request(host, payload) {
    const callId = crypto.randomBytes(8).toString("hex");
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        pending.delete(callId);
        const q = queues.get(host) || [];
        const i = q.findIndex((c) => c.callId === callId);
        if (i >= 0) q.splice(i, 1);
        reject(new Error(`委托读取超时（${Math.round(callTimeoutMs / 1000)} 秒内${host}组件的插件没有响应）`));
      }, callTimeoutMs);
      pending.set(callId, { resolve, reject, timer, host });
      if (!queues.has(host)) queues.set(host, []);
      queues.get(host).push(Object.assign({ callId }, payload));
      hand(host);
    });
  }

  // 插件长轮询取活：有活立即给，没活最多等 pollTimeoutMs 后返回 null
  function poll(host, instanceId) {
    const inst = instances.get(instanceId);
    if (inst) inst.polledAt = now();
    else instances.set(instanceId, { host, docs: [], at: 0, polledAt: now() });
    const q = queues.get(host) || [];
    if (q.length) return Promise.resolve(q.shift());
    return new Promise((resolve) => {
      const waiter = { instanceId, resolve, timer: null };
      waiter.timer = setTimeout(() => {
        const w = waiters.get(host) || [];
        const i = w.indexOf(waiter);
        if (i >= 0) w.splice(i, 1);
        resolve(null);
      }, pollTimeoutMs);
      if (!waiters.has(host)) waiters.set(host, []);
      waiters.get(host).push(waiter);
    });
  }

  function resolveCall(callId, result) {
    const p = pending.get(callId);
    if (!p) return false;
    pending.delete(callId);
    clearTimeout(p.timer);
    p.resolve(result);
    return true;
  }

  function findOpenDoc(filePath) {
    const want = normPath(filePath);
    return listOpenDocs().find((d) => d.path && normPath(d.path) === want) || null;
  }

  return { report, listOpenDocs, workerOnline, request, poll, resolveCall, findOpenDoc };
}

module.exports = { createRefRelay, normPath };
