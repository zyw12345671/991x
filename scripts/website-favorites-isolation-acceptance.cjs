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
    confirm: () => true,
    prompt: () => (promptQueue.length ? String(promptQueue.shift()) : "2"),
    alert: () => {},
  },
  notesState: {
    authenticated: true,
    dataKey: { mock: true },
    account: { username: "qa_user" },
  },
  protectedFeatureConfigs: {
    website: { addLabel: "网站" },
    favorites: { addLabel: "内容" },
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
vm.runInContext(`${bootstrapConstants}\n${extractedCode}`, context, { filename: "qa-isolation-extracted-app.js" });

function assert(condition, message) {
  if (!condition) throw new Error(message);
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

function snapshot(feature) {
  return context
    .normalizeProtectedFeatureCollection(feature, context.protectedFeatureState[feature].items)
    .map((item) => ({ title: item.title, link: item.link }));
}

function latestStatus(feature) {
  const list = statusLog.filter((s) => s.feature === feature);
  return list.length ? list[list.length - 1] : null;
}

async function runIsolation() {
  statusLog.length = 0;
  promptQueue.length = 0;
  backupStore.clear();

  context.protectedFeatureState.website.items = [
    { id: "w1", title: "站点A", link: "https://a.com", updatedAtMs: 1000 },
  ];
  context.protectedFeatureState.favorites.items = [
    { id: "f1", title: "内容A", link: "https://c.com", updatedAtMs: 1000 },
  ];

  const websiteBefore = snapshot("website");
  const favoritesBefore = snapshot("favorites");

  promptQueue.push("2");
  await context.importProtectedFeatureFromJsonFile(
    "website",
    makeImportFile({
      source: "991x-protected-feature-website-collection",
      feature: "website",
      items: [{ title: "站点B", link: "https://b.com" }],
    })
  );
  const websiteAfterWebsiteImport = snapshot("website");
  const favoritesAfterWebsiteImport = snapshot("favorites");
  assert(websiteAfterWebsiteImport.length === 2, "网站区导入后数量应为2");
  assert(JSON.stringify(favoritesAfterWebsiteImport) === JSON.stringify(favoritesBefore), "网站区导入不应影响网络内容区");

  promptQueue.push("2");
  await context.importProtectedFeatureFromJsonFile(
    "favorites",
    makeImportFile({
      source: "991x-protected-feature-favorites-content",
      feature: "favorites",
      items: [{ title: "内容B", link: "https://d.com" }],
    })
  );
  const websiteAfterFavoritesImport = snapshot("website");
  const favoritesAfterFavoritesImport = snapshot("favorites");
  assert(JSON.stringify(websiteAfterFavoritesImport) === JSON.stringify(websiteAfterWebsiteImport), "网络内容区导入不应影响网站区");
  assert(favoritesAfterFavoritesImport.length === 2, "网络内容区导入后数量应为2");

  await context.importProtectedFeatureFromJsonFile(
    "favorites",
    makeImportFile({
      source: "991x-protected-feature-website-collection",
      feature: "website",
      items: [{ title: "跨区文件", link: "https://x.com" }],
    })
  );
  const mismatchStatus = latestStatus("favorites");
  assert(mismatchStatus && mismatchStatus.isError && mismatchStatus.message.includes("文件类型与当前功能不匹配"), "跨区导入应被阻止并报错");
  assert(snapshot("favorites").length === 2, "跨区导入失败后网络内容区数据应保持不变");

  await context.undoLastProtectedFeatureImport("website");
  const websiteAfterUndo = snapshot("website");
  const favoritesAfterWebsiteUndo = snapshot("favorites");
  assert(JSON.stringify(websiteAfterUndo) === JSON.stringify(websiteBefore), "撤销网站区导入应恢复网站区原数据");
  assert(JSON.stringify(favoritesAfterWebsiteUndo) === JSON.stringify(favoritesAfterFavoritesImport), "撤销网站区导入不应影响网络内容区");

  await context.undoLastProtectedFeatureImport("favorites");
  const favoritesAfterUndo = snapshot("favorites");
  assert(JSON.stringify(favoritesAfterUndo) === JSON.stringify(favoritesBefore), "撤销网络内容区导入应恢复网络内容区原数据");

  return {
    websiteBefore,
    favoritesBefore,
    websiteAfterWebsiteImport,
    favoritesAfterFavoritesImport,
    finalWebsite: snapshot("website"),
    finalFavorites: snapshot("favorites"),
    mismatchStatus: mismatchStatus.message,
  };
}

runIsolation()
  .then((report) => {
    console.log("隔离验收结果：全部通过");
    console.log(JSON.stringify(report, null, 2));
  })
  .catch((error) => {
    console.error("隔离验收失败:", error && error.stack ? error.stack : error);
    process.exit(1);
  });
