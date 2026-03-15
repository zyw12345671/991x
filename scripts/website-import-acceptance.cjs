const fs = require("fs");
const path = require("path");
const vm = require("vm");

const appPath = path.join(__dirname, "..", "app.js");
const source = fs.readFileSync(appPath, "utf8");

function sliceBetween(startMarker, endMarker) {
  const start = source.indexOf(startMarker);
  if (start < 0) throw new Error(`Cannot find start marker: ${startMarker}`);
  const end = source.indexOf(endMarker, start);
  if (end < 0) throw new Error(`Cannot find end marker: ${endMarker}`);
  return source.slice(start, end);
}

const bootstrapConstants = `
const WEBSITE_COLLECTION_TYPE_DEFINITIONS = Object.freeze([
  { key: "common", label: "常用网站", icon: "🌐" },
  { key: "news", label: "新闻资讯", icon: "📰" },
  { key: "life", label: "生活购物", icon: "🛍️" },
  { key: "entertainment", label: "娱乐旅行", icon: "🎬" },
  { key: "education", label: "教育学习", icon: "🎓" },
  { key: "ai", label: "人工智能", icon: "🤖" },
  { key: "custom", label: "自定义", icon: "🏷️" },
]);
const WEBSITE_COLLECTION_TYPE_LABELS = Object.freeze(
  WEBSITE_COLLECTION_TYPE_DEFINITIONS.reduce((output, item) => {
    output[item.key] = item.label;
    return output;
  }, {})
);
const FAVORITES_COLLECTION_TYPE_LABEL_OVERRIDES = Object.freeze({
  news: "教程",
  life: "视频",
  entertainment: "参考",
  education: "AI",
});
const WEBSITE_COLLECTION_TYPE_DEFAULT_ICONS = Object.freeze(
  WEBSITE_COLLECTION_TYPE_DEFINITIONS.reduce((output, item) => {
    output[item.key] = item.icon;
    return output;
  }, {})
);
const WEBSITE_COLLECTION_ICON_OPTIONS = Object.freeze([
  "🌐","📰","🛍️","🧭","🎯","🎓","🤖","💼","📚","📊","📺","🎬","🎮","🎧","🧠","💡",
  "🧰","⚙️","🔍","🔐","🏠","🍽️","✈️","🚗","🚄","🏥","💰","📈","🧾","🛒","🏷️","📷",
  "🖼️","✍️","🧪","🔬","🛰️","🌍","🌟","⭐"
]);
const WEBSITE_COLLECTION_ICON_OPTION_SET = new Set(WEBSITE_COLLECTION_ICON_OPTIONS);
const LINK_IMPORT_MODE_SKIP = "skip";
const LINK_IMPORT_MODE_OVERWRITE = "overwrite";
const LINK_IMPORT_MODE_RENAME = "rename";
const PROTECTED_FEATURE_IMPORT_MAX_BYTES = 1024 * 1024 * 2;
`;

const extractedCode = [
  sliceBetween("function createLocalShortcutId", "function getEmptyLocalShortcuts"),
  sliceBetween("function formatLinkImportModeLabel", "function formatTasksImportModeLabel"),
  sliceBetween("function sanitizeProtectedFeatureText", "function sanitizeAiTextPreserveWhitespace"),
  sliceBetween("function normalizeProtectedFeatureUrl", "function createNetworkImageId"),
  sliceBetween("function normalizeProtectedFeatureItem", "function clearProtectedFeatureStatus"),
  sliceBetween("async function importProtectedFeatureFromJsonFile", "async function clearProtectedFeatureItems"),
].join("\n\n");

const statusLog = [];
const confirmLog = [];
const promptLog = [];
const promptQueue = [];
const backupStore = new Map();

const context = {
  console,
  Date,
  Math,
  URL,
  Set,
  Map,
  Number,
  String,
  Object,
  Array,
  JSON,
  Promise,
  window: {
    crypto: {
      randomUUID: (() => {
        let i = 0;
        return () => `uuid-${++i}`;
      })(),
    },
    confirm: (message) => {
      confirmLog.push(String(message || ""));
      return true;
    },
    prompt: (message, defaultValue) => {
      promptLog.push({ message: String(message || ""), defaultValue: String(defaultValue || "") });
      if (promptQueue.length === 0) return "2";
      return String(promptQueue.shift());
    },
    alert: () => {},
  },
  notesState: {
    authenticated: true,
    dataKey: { mock: true },
    account: { username: "qa_user" },
  },
  protectedFeatureConfigs: {
    website: { addLabel: "网站" },
    favorites: { addLabel: "收藏" },
    tasks: { addLabel: "任务" },
    vault: { addLabel: "密盒" },
  },
  protectedFeatureState: {
    website: { items: [], loaded: true },
    favorites: { items: [], loaded: true },
    tasks: { items: [], loaded: true },
    vault: { items: [], loaded: true },
  },
  ensureProtectedFeatureLoaded: async () => {},
  decryptTextWithAesGcm: async () => {
    throw new Error("Encrypted payload test is out of scope in this acceptance run");
  },
  saveProtectedFeatureCollectionToStorage: async () => true,
  renderProtectedFeatureList: () => {},
  syncTasksImportUndoButtonState: () => {},
  syncWebsiteImportUndoButtonState: () => {},
  setProtectedFeatureStatus: (feature, message, isError = false) => {
    statusLog.push({ feature: String(feature || ""), message: String(message || ""), isError: Boolean(isError) });
  },
  saveProtectedFeatureImportBackup: async (feature, items, metadata = {}) => {
    const normalized = context.normalizeProtectedFeatureCollection(feature, items);
    backupStore.set(feature, {
      feature,
      items: normalized,
      importMode: String(metadata.importMode || ""),
      previousCount: Number.isFinite(Number(metadata.previousCount)) ? Math.floor(Number(metadata.previousCount)) : normalized.length,
      createdAt: new Date().toISOString(),
    });
    return true;
  },
  loadProtectedFeatureImportBackup: async (feature) => {
    const payload = backupStore.get(feature);
    return payload ? { ...payload, items: [...payload.items] } : null;
  },
  clearProtectedFeatureImportBackup: (feature) => {
    backupStore.delete(feature);
  },
  formatImportBackupTimeLabel: (raw) => String(raw || ""),
};

vm.createContext(context);
vm.runInContext(`${bootstrapConstants}\n${extractedCode}`, context, { filename: "qa-extracted-app.js" });

function normalizeWebItems(items) {
  return context.normalizeProtectedFeatureCollection("website", items).map((item) => ({
    id: item.id,
    title: item.title,
    link: item.link,
    note: item.note,
  }));
}

function canonicalizeWebItems(items) {
  return [...items]
    .map((item) => ({
      id: String(item.id || ""),
      title: String(item.title || ""),
      link: String(item.link || ""),
      note: String(item.note || ""),
    }))
    .sort((a, b) => {
      const idCompare = a.id.localeCompare(b.id, "zh-CN");
      if (idCompare !== 0) return idCompare;
      const titleCompare = a.title.localeCompare(b.title, "zh-CN");
      if (titleCompare !== 0) return titleCompare;
      return a.link.localeCompare(b.link, "zh-CN");
    });
}

function makeImportFile(payloadObject) {
  const raw = JSON.stringify(payloadObject, null, 2);
  return {
    size: Buffer.byteLength(raw, "utf8"),
    async text() {
      return raw;
    },
  };
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function getLatestStatus(feature) {
  const matched = statusLog.filter((entry) => entry.feature === feature);
  return matched.length ? matched[matched.length - 1] : null;
}

function hasInvalidLinks(items) {
  return items.some((item) => !context.normalizeProtectedFeatureUrl(item.link));
}

function getTitleSet(items) {
  return new Set(items.map((item) => item.title));
}

async function runModeCase(modeChoice, expected) {
  statusLog.length = 0;
  confirmLog.length = 0;
  promptLog.length = 0;
  promptQueue.length = 0;
  backupStore.clear();

  const existing = [
    { id: "site_google_001", title: "Google", link: "https://google.com", note: "搜索", updatedAtMs: 1000 },
    { id: "site_bing_001", title: "Bing", link: "https://bing.com", note: "备用", updatedAtMs: 900 },
  ];
  const expectedRestore = normalizeWebItems(existing);
  context.protectedFeatureState.website.items = normalizeWebItems(existing);

  const importedMixed = [
    { title: "Google", link: "https://google.com", order: 1 },
    { title: "Bing Mirror", link: "https://bing.com", order: 2 },
    { title: "OpenAI", link: "openai.com", order: 3 },
    { title: "Google", link: "https://example.com", order: 4 },
    { title: "Invalid Script", link: "javascript:alert(1)", order: 5 },
    { title: "Invalid FTP", link: "ftp://files.example.com", order: 6 },
    { title: "No Link", order: 7 },
  ];

  promptQueue.push(modeChoice);
  await context.importProtectedFeatureFromJsonFile(
    "website",
    makeImportFile({
      source: "991x-protected-feature-website-collection",
      feature: "website",
      items: importedMixed,
    })
  );

  const afterImport = normalizeWebItems(context.protectedFeatureState.website.items);
  const latestStatus = getLatestStatus("website");
  const backup = backupStore.get("website");

  assert(afterImport.length === expected.count, `${expected.name}: 导入后数量应为 ${expected.count}，实际 ${afterImport.length}`);
  assert(!hasInvalidLinks(afterImport), `${expected.name}: 导入后不应包含无效网址`);
  assert(latestStatus && latestStatus.message.includes("已跳过 3 条无效网址"), `${expected.name}: 状态文案应包含无效网址跳过数量`);
  assert(backup && backup.previousCount === 2, `${expected.name}: 应创建可撤销备份且 previousCount=2`);

  if (expected.check === "skip") {
    const titles = getTitleSet(afterImport);
    assert(titles.has("Google") && titles.has("Bing") && titles.has("OpenAI"), `${expected.name}: 应保留原有并新增 OpenAI`);
  } else if (expected.check === "overwrite") {
    const byTitle = new Map(afterImport.map((item) => [item.title, item]));
    assert(byTitle.has("Bing Mirror"), `${expected.name}: 应覆盖为 Bing Mirror`);
    assert(Array.from(byTitle.values()).some((item) => item.link.includes("example.com")), `${expected.name}: 应将 Google 冲突覆盖为 example.com`);
  } else if (expected.check === "rename") {
    const titles = afterImport.map((item) => item.title);
    assert(titles.some((title) => title.startsWith("Google（导入")), `${expected.name}: 应生成重命名标题 Google（导入）`);
  }

  await context.undoLastProtectedFeatureImport("website");
  const afterUndo = normalizeWebItems(context.protectedFeatureState.website.items);
  const undoStatus = getLatestStatus("website");

  assert(
    JSON.stringify(canonicalizeWebItems(afterUndo)) === JSON.stringify(canonicalizeWebItems(expectedRestore)),
    `${expected.name}: 撤销后应恢复到导入前状态`
  );
  assert(!backupStore.has("website"), `${expected.name}: 撤销后应清理备份`);
  assert(undoStatus && undoStatus.message.includes("已撤销上次导入"), `${expected.name}: 撤销后应显示成功文案`);

  return {
    mode: expected.name,
    importCount: afterImport.length,
    importStatus: latestStatus?.message || "",
    undoStatus: undoStatus?.message || "",
    sampleTitles: afterImport.map((item) => item.title),
  };
}

async function runAllInvalidCase() {
  statusLog.length = 0;
  confirmLog.length = 0;
  promptLog.length = 0;
  promptQueue.length = 0;
  backupStore.clear();

  const existing = [
    { id: "site_google_001", title: "Google", link: "https://google.com", note: "搜索", updatedAtMs: 1000 },
    { id: "site_bing_001", title: "Bing", link: "https://bing.com", note: "备用", updatedAtMs: 900 },
  ];
  const before = normalizeWebItems(existing);
  context.protectedFeatureState.website.items = normalizeWebItems(existing);

  promptQueue.push("2");
  await context.importProtectedFeatureFromJsonFile(
    "website",
    makeImportFile({
      source: "991x-protected-feature-website-collection",
      feature: "website",
      items: [
        { title: "Bad 1", link: "javascript:alert(1)" },
        { title: "Bad 2", link: "ftp://bad.example" },
        { title: "Bad 3" },
      ],
    })
  );

  const after = normalizeWebItems(context.protectedFeatureState.website.items);
  const latestStatus = getLatestStatus("website");

  assert(
    JSON.stringify(canonicalizeWebItems(after)) === JSON.stringify(canonicalizeWebItems(before)),
    "全无效网址: 数据不应被改动"
  );
  assert(latestStatus && latestStatus.isError, "全无效网址: 应提示错误状态");
  assert(
    latestStatus && latestStatus.message.includes("已跳过 3 条无效网址") && latestStatus.message.includes("未发现可导入条目"),
    "全无效网址: 应提示跳过数量与无可导入条目"
  );
  assert(!backupStore.has("website"), "全无效网址: 不应创建备份");

  return {
    mode: "全无效网址拦截",
    status: latestStatus?.message || "",
    preservedCount: after.length,
  };
}

(async () => {
  const report = [];
  report.push(await runModeCase("1", { name: "跳过冲突", check: "skip", count: 3 }));
  report.push(await runModeCase("2", { name: "覆盖冲突", check: "overwrite", count: 3 }));
  report.push(await runModeCase("3", { name: "重命名冲突", check: "rename", count: 4 }));
  report.push(await runAllInvalidCase());

  console.log("验收结果：全部通过");
  console.log(JSON.stringify(report, null, 2));
})().catch((error) => {
  console.error("验收失败:", error && error.stack ? error.stack : error);
  process.exit(1);
});
