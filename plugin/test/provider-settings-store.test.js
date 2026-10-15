const test = require("node:test");
const assert = require("node:assert");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const SETTINGS_KEY = "wps_ai_provider_settings_v1";

function makeStorage(seed = {}) {
  const mem = new Map(Object.entries(seed));
  return {
    getItem: (key) => (mem.has(key) ? mem.get(key) : null),
    setItem: (key, value) => { mem.set(key, String(value)); },
    removeItem: (key) => { mem.delete(key); },
    _mem: mem
  };
}

function loadRegistry({ storeSeed = {}, localSeed = {} } = {}) {
  const localStorage = makeStorage(localSeed);
  const store = makeStorage(storeSeed);
  const sandbox = {
    window: { WpsAiStore: store },
    localStorage,
    console
  };
  vm.createContext(sandbox);
  const code = fs.readFileSync(path.join(__dirname, "..", "js", "providers", "registry.js"), "utf8");
  vm.runInContext(code, sandbox);
  return { registry: sandbox.window.WpsAiProviderRegistry, store, localStorage };
}

test("provider 设置从共享 WpsAiStore 读取，避免 Word/Excel localStorage 隔离", () => {
  const settings = {
    activeChatModel: "excel-shared::qwen-plus",
    chatProviders: [
      {
        id: "excel-shared",
        type: "openai",
        label: "Excel 可见的共享模型",
        enabled: true,
        baseUrl: "https://example.test/v1",
        apiKey: "sk-shared",
        defaultModel: "qwen-plus"
      }
    ]
  };
  const { registry } = loadRegistry({
    storeSeed: { [SETTINGS_KEY]: JSON.stringify(settings) }
  });

  const loaded = registry.loadSettings();

  assert.equal(loaded.activeChatModel, "excel-shared::qwen-plus");
  assert.equal(loaded.chatProviders[0].id, "excel-shared");
  assert.equal(loaded.chatProviders[0].apiKey, "sk-shared");
});

test("provider 设置保存到共享 WpsAiStore，不写宿主隔离的 localStorage", () => {
  const { registry, store, localStorage } = loadRegistry();
  const settings = registry.loadSettings();
  settings.chatProviders = [
    {
      id: "word-configured",
      type: "openai",
      label: "Word 配置的模型",
      enabled: true,
      baseUrl: "https://example.test/v1",
      apiKey: "sk-word",
      defaultModel: "deepseek-chat"
    }
  ];

  registry.saveSettings(settings);

  assert.match(store.getItem(SETTINGS_KEY), /word-configured/);
  assert.equal(localStorage.getItem(SETTINGS_KEY), null);
});

test("provider 设置读取时使用更新的 localStorage 备份并回灌共享 WpsAiStore", () => {
  const oldSettings = {
    __updatedAt: 100,
    activeChatModel: "old-provider::old-model",
    chatProviders: [
      {
        id: "old-provider",
        type: "openai",
        label: "旧供应商",
        enabled: true,
        baseUrl: "https://old.example/v1",
        apiKey: "sk-old",
        defaultModel: "old-model"
      }
    ]
  };
  const newSettings = {
    __updatedAt: 200,
    activeChatModel: "new-provider::new-model",
    chatProviders: [
      {
        id: "new-provider",
        type: "openai",
        label: "新供应商",
        enabled: false,
        baseUrl: "https://new.example/v1",
        apiKey: "sk-new",
        defaultModel: "new-model"
      }
    ]
  };
  const { registry, store } = loadRegistry({
    storeSeed: { [SETTINGS_KEY]: JSON.stringify(oldSettings) },
    localSeed: { [SETTINGS_KEY]: JSON.stringify(newSettings) }
  });

  const loaded = registry.loadSettings();

  assert.equal(loaded.activeChatModel, "new-provider::new-model");
  assert.equal(loaded.chatProviders[0].id, "new-provider");
  assert.equal(loaded.chatProviders[0].enabled, false);
  assert.match(store.getItem(SETTINGS_KEY), /new-provider/);
});

test("聊天供应商只允许一家启用：存量多家启用时保留 activeChatModel 指向的那家", () => {
  const mk = (id, enabled) => ({ id, type: "openai", label: id, enabled, baseUrl: "https://example.test/v1", apiKey: "sk", defaultModel: `${id}-model` });
  const load = (activeChatModel) => loadRegistry({
    storeSeed: { [SETTINGS_KEY]: JSON.stringify({ activeChatModel, chatProviders: [mk("a", true), mk("b", true), mk("c", false)] }) }
  }).registry.loadSettings();
  const enabledIds = (s) => s.chatProviders.filter((p) => p.enabled).map((p) => p.id);

  assert.deepEqual(enabledIds(load("b::b-model")), ["b"]);
  assert.deepEqual(enabledIds(load("")), ["a"], "没有指向时保留排在最前的那家");
  assert.deepEqual(enabledIds(load("c::c-model")), ["a"], "指向未启用的那家时不擅自启用它");
});
