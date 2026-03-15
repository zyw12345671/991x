const http = require("node:http");
const fs = require("node:fs");
const fsp = require("node:fs/promises");
const os = require("node:os");
const path = require("node:path");
const url = require("node:url");
const { spawn } = require("node:child_process");
const crypto = require("node:crypto");
const { PDFParse } = require("pdf-parse");
const JSZip = require("jszip");
const XLSX = require("xlsx");
const CFB = require("cfb");
const WordExtractor = require("word-extractor");
let officeParser = null;
let pptToText = null;
let playwrightChromium = null;

try {
  officeParser = require("officeparser");
} catch {
  officeParser = null;
}

try {
  pptToText = require("ppt-to-text");
} catch {
  pptToText = null;
}

try {
  ({ chromium: playwrightChromium } = require("playwright-core"));
} catch {
  playwrightChromium = null;
}

try {
  const cptable = require("xlsx/dist/cpexcel.js");
  if (cptable && typeof XLSX.set_cptable === "function") {
    XLSX.set_cptable(cptable);
  }
} catch {
  // Optional codepage table for legacy XLS decoding.
}

const HOST = "127.0.0.1";
const PORT = Number(process.env.PORT || 9910);
const ROOT_DIR = __dirname;
const GAOZHI_DIR = path.join(ROOT_DIR, "gaozhi");
const FENXIANG_POSTS_DIR = path.join(ROOT_DIR, "fenxiang", "_posts");
const DATA_DIR = path.join(ROOT_DIR, "data");
const VISIT_DURATIONS_FILE = path.join(DATA_DIR, "visit-durations.json");
const LOCAL_SHORTCUTS_FILE = path.join(DATA_DIR, "local-shortcuts.json");
const LOCAL_SHORTCUTS_HISTORY_DIR = path.join(DATA_DIR, "local-shortcuts-history");
const LOCAL_SHORTCUTS_HISTORY_KEEP_LIMIT = 180;
const LOCAL_AGENT_MAX_SCAN_FILES = parseIntegerEnv("LOCAL_AGENT_MAX_SCAN_FILES", 3200, 100, 30000);
const LOCAL_AGENT_MAX_SCAN_DEPTH = parseIntegerEnv("LOCAL_AGENT_MAX_SCAN_DEPTH", 8, 1, 32);
const LOCAL_AGENT_PREVIEW_LIMIT = parseIntegerEnv("LOCAL_AGENT_PREVIEW_LIMIT", 120, 10, 2000);
const LOCAL_AGENT_MAX_RENAME_ITEMS = parseIntegerEnv("LOCAL_AGENT_MAX_RENAME_ITEMS", 1200, 20, 20000);
const LOCAL_AGENT_MAX_EDIT_BYTES = parseIntegerEnv("LOCAL_AGENT_MAX_EDIT_BYTES", 2 * 1024 * 1024, 8 * 1024, 20 * 1024 * 1024);
const LOCAL_AGENT_COMMAND_TIMEOUT_MS = parseTimeoutEnv("LOCAL_AGENT_COMMAND_TIMEOUT_MS", 12000, 1500);
const LOCAL_AGENT_OPEN_SEARCH_MAX_DEPTH = parseIntegerEnv("LOCAL_AGENT_OPEN_SEARCH_MAX_DEPTH", 5, 1, 12);
const LOCAL_AGENT_OPEN_SEARCH_MAX_DIRS = parseIntegerEnv("LOCAL_AGENT_OPEN_SEARCH_MAX_DIRS", 1200, 100, 20000);
const LOCAL_AGENT_OPEN_SEARCH_MAX_ENTRIES = parseIntegerEnv("LOCAL_AGENT_OPEN_SEARCH_MAX_ENTRIES", 50000, 5000, 200000);
const LOCAL_AGENT_OPEN_SEARCH_MAX_MATCHES = parseIntegerEnv("LOCAL_AGENT_OPEN_SEARCH_MAX_MATCHES", 40, 5, 200);
const LOCAL_AGENT_ACTION_SYSTEM_INFO = "system-info";
const LOCAL_AGENT_ACTION_ORGANIZE = "organize-files";
const LOCAL_AGENT_ACTION_RENAME = "rename-files";
const LOCAL_AGENT_ACTION_EDIT = "edit-text";
const LOCAL_AGENT_ACTION_OPEN = "open-path";
const LOCAL_AGENT_TEXT_FILE_EXTENSIONS = new Set([
  ".txt",
  ".md",
  ".markdown",
  ".json",
  ".xml",
  ".yaml",
  ".yml",
  ".ini",
  ".conf",
  ".csv",
  ".log",
  ".html",
  ".css",
  ".js",
  ".mjs",
  ".cjs",
  ".ts",
  ".tsx",
  ".jsx",
  ".py",
  ".java",
  ".c",
  ".cpp",
  ".cc",
  ".h",
  ".hpp",
  ".go",
  ".rs",
  ".php",
  ".sh",
  ".bat",
  ".ps1",
  ".sql",
]);
const LOCAL_AGENT_CATEGORY_RULES = Object.freeze([
  { name: "图片", folder: "图片", extensions: new Set([".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".svg", ".avif", ".ico"]) },
  { name: "视频", folder: "视频", extensions: new Set([".mp4", ".mov", ".mkv", ".avi", ".wmv", ".flv", ".webm", ".m4v"]) },
  { name: "音频", folder: "音频", extensions: new Set([".mp3", ".wav", ".flac", ".aac", ".ogg", ".m4a", ".opus"]) },
  { name: "文档", folder: "文档", extensions: new Set([".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".txt", ".md"]) },
  { name: "压缩包", folder: "压缩包", extensions: new Set([".zip", ".rar", ".7z", ".tar", ".gz", ".bz2"]) },
  { name: "代码", folder: "代码", extensions: new Set([".js", ".ts", ".tsx", ".jsx", ".py", ".java", ".go", ".rs", ".c", ".cpp", ".h", ".hpp"]) },
]);
const AI_IMAGE_CACHE_DIR = path.join(DATA_DIR, "ai-image-cache");
function parseTimeoutEnv(name, fallbackMs, minMs = 5000) {
  const raw = Number(process.env[name]);
  if (!Number.isFinite(raw)) return fallbackMs;
  return Math.max(minMs, Math.floor(raw));
}

function parseIntegerEnv(name, fallbackValue, minValue = 1, maxValue = Number.MAX_SAFE_INTEGER) {
  const raw = Number(process.env[name]);
  if (!Number.isFinite(raw)) return Math.max(minValue, Math.min(maxValue, Math.floor(fallbackValue)));
  return Math.max(minValue, Math.min(maxValue, Math.floor(raw)));
}

function parseSizeLimitEnv(
  name,
  fallbackValue,
  minValue = 1,
  maxValue = Number.MAX_SAFE_INTEGER,
  unlimitedSentinel = Number.POSITIVE_INFINITY
) {
  const fallback = Number(fallbackValue);
  const fallbackNormalized = Number.isFinite(fallback)
    ? Math.max(minValue, Math.min(maxValue, Math.floor(fallback)))
    : unlimitedSentinel;
  const rawText = String(process.env[name] || "").trim();
  if (!rawText) return fallbackNormalized;
  const raw = Number(rawText);
  if (!Number.isFinite(raw)) return fallbackNormalized;
  if (raw <= 0) return unlimitedSentinel;
  return Math.max(minValue, Math.min(maxValue, Math.floor(raw)));
}

const AI_REQUEST_TIMEOUT_MS = parseTimeoutEnv("AI_REQUEST_TIMEOUT_MS", 120000, 5000);
const AI_REASONING_CHAT_TIMEOUT_MS = parseTimeoutEnv("AI_REASONING_CHAT_TIMEOUT_MS", 10 * 60 * 1000, 30000);
const AI_IMAGE_REQUEST_TIMEOUT_MS = parseTimeoutEnv("AI_IMAGE_REQUEST_TIMEOUT_MS", 8 * 60 * 1000, 30000);
const AI_IMAGE_EDIT_TIMEOUT_MS = parseTimeoutEnv("AI_IMAGE_EDIT_TIMEOUT_MS", 12 * 60 * 1000, 30000);
const AI_OLLAMA_CHAT_TIMEOUT_MS = parseTimeoutEnv("AI_OLLAMA_CHAT_TIMEOUT_MS", 15 * 60 * 1000, 30000);
const AI_OLLAMA_KEEP_ALIVE = String(process.env.AI_OLLAMA_KEEP_ALIVE || "30m").trim() || "30m";
const AI_WEB_SEARCH_TIMEOUT_MS = 18000;
const AI_WEBPAGE_FETCH_TIMEOUT_MS = parseTimeoutEnv("AI_WEBPAGE_FETCH_TIMEOUT_MS", 18000, 3000);
const AI_WEBPAGE_FETCH_MAX_BYTES = parseIntegerEnv("AI_WEBPAGE_FETCH_MAX_BYTES", 2 * 1024 * 1024, 200 * 1024, 8 * 1024 * 1024);
const AI_WEBPAGE_EXTRACT_TEXT_LIMIT = parseIntegerEnv("AI_WEBPAGE_EXTRACT_TEXT_LIMIT", 24000, 4000, 120000);
const AI_WEB_SEARCH_ENGINE_TIMEOUT_MS = parseTimeoutEnv(
  "AI_WEB_SEARCH_ENGINE_TIMEOUT_MS",
  Math.min(AI_WEB_SEARCH_TIMEOUT_MS, 9000),
  2000
);
const AI_ON_THIS_DAY_SOURCE_TIMEOUT_MS = parseTimeoutEnv("AI_ON_THIS_DAY_SOURCE_TIMEOUT_MS", 12000, 3000);
const AI_WEATHER_REQUEST_TIMEOUT_MS = parseTimeoutEnv("AI_WEATHER_REQUEST_TIMEOUT_MS", 12000, 3000);
const AI_MAP_REQUEST_TIMEOUT_MS = parseTimeoutEnv("AI_MAP_REQUEST_TIMEOUT_MS", 7000, 2000);
const AI_WEB_SEARCH_CACHE_TTL_MS = parseIntegerEnv("AI_WEB_SEARCH_CACHE_TTL_MS", 3 * 60 * 1000, 10000, 60 * 60 * 1000);
const AI_WEB_SEARCH_CACHE_MAX_ITEMS = parseIntegerEnv("AI_WEB_SEARCH_CACHE_MAX_ITEMS", 360, 20, 5000);
const AI_RESEARCH_DEFAULT_MAX_PAGES = parseIntegerEnv("AI_RESEARCH_DEFAULT_MAX_PAGES", 3, 1, 8);
const AI_RESEARCH_DEFAULT_PAGE_MAX_CHARS = parseIntegerEnv("AI_RESEARCH_DEFAULT_PAGE_MAX_CHARS", 5200, 1000, 24000);
const AI_WEB_SEARCH_ENGINE_HEALTH_CACHE_MS = parseIntegerEnv(
  "AI_WEB_SEARCH_ENGINE_HEALTH_CACHE_MS",
  1000 * 60 * 3,
  10000,
  1000 * 60 * 60
);
const AI_WEB_SEARCH_ENGINE_PROBE_TIMEOUT_MS = parseTimeoutEnv(
  "AI_WEB_SEARCH_ENGINE_PROBE_TIMEOUT_MS",
  Math.min(AI_WEB_SEARCH_ENGINE_TIMEOUT_MS, 6000),
  2000
);
const AI_WEB_SEARCH_ENGINE_PROBE_QUERY = String(process.env.AI_WEB_SEARCH_ENGINE_PROBE_QUERY || "openai").trim() || "openai";
const AI_WEB_SEARCH_SEARX_SPACE_INSTANCES_URL = String(
  process.env.AI_WEB_SEARCH_SEARX_SPACE_INSTANCES_URL || "https://searx.space/data/instances.json"
).trim();
const AI_WEB_SEARCH_SEARX_SPACE_TIMEOUT_MS = parseTimeoutEnv("AI_WEB_SEARCH_SEARX_SPACE_TIMEOUT_MS", 12000, 3000);
const AI_WEB_SEARCH_SEARX_POOL_CACHE_MS = parseIntegerEnv(
  "AI_WEB_SEARCH_SEARX_POOL_CACHE_MS",
  1000 * 60 * 60 * 6,
  60000,
  1000 * 60 * 60 * 24 * 7
);
const AI_WEB_SEARCH_SEARX_MAX_INSTANCE_ATTEMPTS = parseIntegerEnv("AI_WEB_SEARCH_SEARX_MAX_INSTANCE_ATTEMPTS", 2, 1, 5);
const AI_WEB_SEARCH_SEARX_MAX_QUERY_ROUNDS = parseIntegerEnv("AI_WEB_SEARCH_SEARX_MAX_QUERY_ROUNDS", 3, 1, 6);
const AI_WEB_SEARCH_SEARX_INSTANCES_ENV = String(process.env.AI_WEB_SEARCH_SEARX_INSTANCES || "").trim();
const AI_WEB_SEARCH_SEARX_DEFAULT_INSTANCES = [
  "https://searx.tiekoetter.com",
  "https://search.ononoki.org",
  "https://search.sapti.me",
  "https://search.bus-hit.me",
  "https://searx.be",
];
let aiWebSearchSearxPoolCache = {
  expiresAtMs: 0,
  instances: [],
};
const aiWebSearchResultCache = new Map();
const aiWebSearchEngineHealthCache = new Map();
const AI_ATTACHMENT_PARSE_MAX_BYTES = parseSizeLimitEnv(
  "AI_ATTACHMENT_PARSE_MAX_BYTES",
  Number.POSITIVE_INFINITY,
  2 * 1024 * 1024
);
const AI_ATTACHMENT_PARSE_TEXT_LIMIT = parseIntegerEnv("AI_ATTACHMENT_PARSE_TEXT_LIMIT", 600000, 8000, 5000000);
const AI_ATTACHMENT_IMAGE_MAX_COUNT = parseIntegerEnv("AI_ATTACHMENT_IMAGE_MAX_COUNT", 80, 1, 600);
const AI_ATTACHMENT_IMAGE_MAX_BYTES = parseIntegerEnv(
  "AI_ATTACHMENT_IMAGE_MAX_BYTES",
  12 * 1024 * 1024,
  32 * 1024,
  80 * 1024 * 1024
);
const AI_ATTACHMENT_IMAGE_TOTAL_MAX_BYTES = parseIntegerEnv(
  "AI_ATTACHMENT_IMAGE_TOTAL_MAX_BYTES",
  180 * 1024 * 1024,
  128 * 1024,
  512 * 1024 * 1024
);
const AI_ATTACHMENT_OPENXML_ENTRY_MAX_COUNT = parseIntegerEnv("AI_ATTACHMENT_OPENXML_ENTRY_MAX_COUNT", 6000, 200, 50000);
const AI_ATTACHMENT_PARSE_REQUEST_MAX_BYTES = parseSizeLimitEnv(
  "AI_ATTACHMENT_PARSE_REQUEST_MAX_BYTES",
  Number.isFinite(AI_ATTACHMENT_PARSE_MAX_BYTES) ? Math.ceil(AI_ATTACHMENT_PARSE_MAX_BYTES * 1.7) : Number.POSITIVE_INFINITY,
  4 * 1024 * 1024
);
const AI_OFFICE_CONVERT_TIMEOUT_MS = parseTimeoutEnv("AI_OFFICE_CONVERT_TIMEOUT_MS", 45000, 8000);
const AI_OFFICE_FALLBACK_ENABLED = String(process.env.AI_OFFICE_FALLBACK_ENABLED || "0").trim() !== "0";
const AI_OFFICE_FALLBACK_CACHE_MS = parseIntegerEnv("AI_OFFICE_FALLBACK_CACHE_MS", 1000 * 60 * 10, 10000, 1000 * 60 * 60);
const AI_LIBREOFFICE_BIN_ENV = String(process.env.AI_LIBREOFFICE_BIN || "").trim();
let aiOfficeFallbackBinaryCache = {
  expiresAtMs: 0,
  binary: "",
  error: "",
};
const AI_ATTACHMENT_TRANSCRIBE_TIMEOUT_MS = 180000;
const AI_IMAGE_FETCH_TIMEOUT_MS = parseTimeoutEnv("AI_IMAGE_FETCH_TIMEOUT_MS", 30000, 5000);
const AI_IMAGE_FETCH_MAX_BYTES = 25 * 1024 * 1024;
const AI_IMAGE_CACHE_MAX_BYTES = 25 * 1024 * 1024;
const AI_IMAGE_CACHE_BODY_MAX_BYTES = 80 * 1024 * 1024;
const AI_IMAGE_CACHE_MAX_FILES = 600;
const AI_BROWSER_AGENT_NAV_TIMEOUT_MS = parseTimeoutEnv("AI_BROWSER_AGENT_NAV_TIMEOUT_MS", 30000, 5000);
const AI_BROWSER_AGENT_ACTION_TIMEOUT_MS = parseTimeoutEnv("AI_BROWSER_AGENT_ACTION_TIMEOUT_MS", 12000, 2000);
const AI_BROWSER_AGENT_DEFAULT_TEXT_MAX_CHARS = parseIntegerEnv("AI_BROWSER_AGENT_DEFAULT_TEXT_MAX_CHARS", 7000, 1200, 30000);
const AI_BROWSER_AGENT_DEFAULT_MAX_LINKS = parseIntegerEnv("AI_BROWSER_AGENT_DEFAULT_MAX_LINKS", 12, 4, 40);
const AI_BROWSER_AGENT_MAX_FORMS = parseIntegerEnv("AI_BROWSER_AGENT_MAX_FORMS", 6, 1, 20);
const AI_BROWSER_AGENT_MAX_FORM_FIELDS = parseIntegerEnv("AI_BROWSER_AGENT_MAX_FORM_FIELDS", 14, 2, 40);
const AI_BROWSER_AGENT_MAX_BUTTONS = parseIntegerEnv("AI_BROWSER_AGENT_MAX_BUTTONS", 18, 4, 60);
const AI_BROWSER_AGENT_MAX_AUTOFILL_FIELDS = parseIntegerEnv("AI_BROWSER_AGENT_MAX_AUTOFILL_FIELDS", 16, 1, 40);
const AI_BROWSER_AGENT_EXECUTABLE_PATH_ENV = String(
  process.env.AI_BROWSER_AGENT_EXECUTABLE_PATH || process.env.AI_BROWSER_AGENT_EXECUTABLE || ""
).trim();
const AI_AUDIO_EXTENSIONS = new Set(["mp3", "wav", "m4a", "flac", "aac", "ogg", "opus", "webm"]);
const AI_VIDEO_EXTENSIONS = new Set(["mp4", "mov", "mkv", "webm", "avi", "m4v"]);
const AI_DOCUMENT_EXTENSIONS = new Set(["pdf", "doc", "docx", "xls", "xlsx", "ppt", "pptx", "odt", "ods", "odp"]);
const AI_DOCUMENT_MIME_TO_EXTENSION = new Map([
  ["application/pdf", "pdf"],
  ["application/x-pdf", "pdf"],
  ["application/msword", "doc"],
  ["application/vnd.openxmlformats-officedocument.wordprocessingml.document", "docx"],
  ["application/vnd.ms-excel", "xls"],
  ["application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "xlsx"],
  ["application/vnd.ms-excel.sheet.macroenabled.12", "xlsx"],
  ["application/vnd.ms-powerpoint", "ppt"],
  ["application/mspowerpoint", "ppt"],
  ["application/x-mspowerpoint", "ppt"],
  ["application/vnd.openxmlformats-officedocument.presentationml.presentation", "pptx"],
  ["application/vnd.oasis.opendocument.text", "odt"],
  ["application/vnd.oasis.opendocument.spreadsheet", "ods"],
  ["application/vnd.oasis.opendocument.presentation", "odp"],
]);
const AI_OLLAMA_DEFAULT_BASE_URL = "http://127.0.0.1:11434";

const MIME_MAP = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".md": "text/markdown; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".bmp": "image/bmp",
  ".avif": "image/avif",
  ".svg": "image/svg+xml",
  ".webp": "image/webp",
  ".ico": "image/x-icon",
  ".woff2": "font/woff2",
  ".woff": "font/woff",
  ".ttf": "font/ttf",
  ".otf": "font/otf",
};

function writeJson(res, statusCode, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Content-Length": Buffer.byteLength(body),
  });
  res.end(body);
}

function normalizeWebPathname(pathname) {
  const decoded = decodeURIComponent(pathname || "/");
  if (decoded === "/") return "/index.html";
  return decoded;
}

function toSafeFilePath(pathname) {
  const normalized = normalizeWebPathname(pathname);
  const trimmed = normalized.replace(/^\/+/, "");
  const resolved = path.resolve(ROOT_DIR, trimmed);
  if (!resolved.startsWith(path.resolve(ROOT_DIR))) return null;
  return resolved;
}

function sanitizeVisitDailyTotals(value) {
  if (!value || typeof value !== "object") return {};
  const sanitized = {};
  Object.entries(value).forEach(([key, rawAmount]) => {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(key)) return;
    const amount = Number(rawAmount);
    if (!Number.isFinite(amount) || amount <= 0) return;
    sanitized[key] = Math.floor(amount);
  });
  return sanitized;
}

function createLocalShortcutId(prefix = "sc") {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function sanitizeLocalShortcutCategory(rawValue) {
  const normalized = String(rawValue || "").trim().toLowerCase();
  if (normalized === "applications") return "applications";
  if (normalized === "files") return "files";
  return "";
}

function normalizeLocalShortcutPath(rawPath) {
  if (typeof rawPath !== "string") return "";
  let value = rawPath.trim();
  if (!value) return "";
  const quoted =
    (value.startsWith('"') && value.endsWith('"')) ||
    (value.startsWith("'") && value.endsWith("'"));
  if (quoted) value = value.slice(1, -1).trim();
  if (!value || /[\0\r\n]/.test(value)) return "";
  return value;
}

function normalizeAbsoluteLocalPath(rawPath) {
  const pathText = normalizeLocalShortcutPath(rawPath);
  if (!pathText) return "";
  const windowsPath = pathText.replaceAll("/", "\\");
  if (!path.win32.isAbsolute(windowsPath)) return "";
  return path.win32.normalize(windowsPath);
}

function inferShortcutName(rawPath) {
  const fullPath = normalizeLocalShortcutPath(rawPath);
  if (!fullPath) return "未命名";
  const normalized = fullPath.replaceAll("\\", "/");
  const parts = normalized.split("/").filter(Boolean);
  return parts[parts.length - 1] || normalized;
}

function sanitizeLocalShortcutEntry(rawItem, category) {
  if (!rawItem || typeof rawItem !== "object") return null;
  const safePath = normalizeAbsoluteLocalPath(rawItem.path ?? rawItem.target ?? "");
  if (!safePath) return null;
  const idRaw = typeof rawItem.id === "string" ? rawItem.id.trim() : "";
  const safeId =
    /^[a-zA-Z0-9_-]{6,80}$/.test(idRaw) ? idRaw : createLocalShortcutId(category === "applications" ? "app" : "file");
  const nameRaw = typeof rawItem.name === "string" ? rawItem.name.trim() : "";
  const name = nameRaw || inferShortcutName(safePath);
  const createdAtMsRaw = Number(rawItem.createdAtMs);
  const createdAtMs = Number.isFinite(createdAtMsRaw) && createdAtMsRaw > 0 ? Math.floor(createdAtMsRaw) : Date.now();
  return {
    id: safeId,
    name,
    path: safePath,
    createdAtMs,
  };
}

function sanitizeLocalShortcuts(value) {
  const sanitized = { applications: [], files: [] };
  if (!value || typeof value !== "object") return sanitized;
  ["applications", "files"].forEach((category) => {
    const entries = Array.isArray(value[category]) ? value[category] : [];
    const map = new Map();
    entries.forEach((entry) => {
      const safeItem = sanitizeLocalShortcutEntry(entry, category);
      if (!safeItem) return;
      if (!map.has(safeItem.id)) map.set(safeItem.id, safeItem);
    });
    sanitized[category] = Array.from(map.values()).sort((a, b) => b.createdAtMs - a.createdAtMs);
  });
  return sanitized;
}

async function readRequestBody(req, limitBytes = 1024 * 1024) {
  return new Promise((resolve, reject) => {
    let raw = "";
    let bytes = 0;
    let overflow = false;
    req.setEncoding("utf8");
    req.on("data", (chunk) => {
      if (overflow) return;
      bytes += Buffer.byteLength(chunk);
      if (bytes > limitBytes) {
        overflow = true;
        return;
      }
      raw += chunk;
    });
    req.on("end", () => {
      if (overflow) {
        const err = new Error("Payload Too Large");
        err.code = "PAYLOAD_TOO_LARGE";
        reject(err);
        return;
      }
      resolve(raw);
    });
    req.on("error", reject);
  });
}

async function readRequestBuffer(req, limitBytes = 1024 * 1024) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let bytes = 0;
    let overflow = false;
    req.on("data", (chunk) => {
      if (overflow) return;
      const part = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk || "");
      bytes += part.length;
      if (bytes > limitBytes) {
        overflow = true;
        return;
      }
      chunks.push(part);
    });
    req.on("end", () => {
      if (overflow) {
        const err = new Error("Payload Too Large");
        err.code = "PAYLOAD_TOO_LARGE";
        reject(err);
        return;
      }
      resolve(Buffer.concat(chunks));
    });
    req.on("error", reject);
  });
}

function normalizeAiHttpUrl(rawValue, fallback = "") {
  const raw = String(rawValue || "").trim();
  if (!raw) return fallback;
  const normalized = /^https?:\/\//i.test(raw) ? raw : `https://${raw}`;
  try {
    const parsed = new URL(normalized);
    if (!/^https?:$/i.test(parsed.protocol)) return fallback;
    const pathname = parsed.pathname.replace(/\/+$/g, "");
    return `${parsed.origin}${pathname}`;
  } catch {
    return fallback;
  }
}

function normalizeAiHttpPath(rawValue, fallbackPath = "/") {
  const raw = String(rawValue || "").trim();
  if (!raw) return fallbackPath;
  if (/^https?:\/\//i.test(raw)) return raw;
  let safePath = raw.replace(/\s+/g, "");
  if (!safePath.startsWith("/")) safePath = `/${safePath}`;
  return safePath;
}

function buildAiEndpointUrl(baseUrl, pathOrUrl) {
  const safeBase = normalizeAiHttpUrl(baseUrl, "");
  const safePath = normalizeAiHttpPath(pathOrUrl, "/");
  if (!safeBase && /^https?:\/\//i.test(safePath)) return safePath;
  if (!safeBase) return "";
  if (/^https?:\/\//i.test(safePath)) return safePath;
  return `${safeBase}${safePath}`;
}

function splitAiBaseUrlAndKnownEndpoint(rawBaseUrl) {
  const normalized = normalizeAiHttpUrl(rawBaseUrl, "");
  if (!normalized) return { baseUrl: "", endpointPath: "", endpointType: "", adjusted: false };
  try {
    const parsed = new URL(normalized);
    const pathname = String(parsed.pathname || "").replace(/\/+$/g, "");
    const knownEndpoints = [
      { type: "image", path: "/images/edits" },
      { type: "image", path: "/images/generations" },
      { type: "chat", path: "/chat/completions" },
      { type: "video", path: "/videos/generations" },
      { type: "models", path: "/models" },
    ];
    const lowerPathname = pathname.toLowerCase();
    for (const item of knownEndpoints) {
      const lowerKnown = item.path.toLowerCase();
      if (!lowerPathname.endsWith(lowerKnown)) continue;
      const keepLength = pathname.length - lowerKnown.length;
      const basePath = keepLength > 0 ? pathname.slice(0, keepLength) : "";
      const baseUrl = `${parsed.origin}${basePath}` || parsed.origin;
      return {
        baseUrl,
        endpointPath: item.path,
        endpointType: item.type,
        adjusted: true,
      };
    }
  } catch {
    // ignore
  }
  return { baseUrl: normalized, endpointPath: "", endpointType: "", adjusted: false };
}

async function fetchAiJson(targetUrl, options = {}) {
  if (!targetUrl) {
    const error = new Error("目标地址无效");
    error.code = "INVALID_TARGET_URL";
    throw error;
  }
  const controller = new AbortController();
  const timeout = Number.isFinite(Number(options.timeoutMs)) ? Math.max(5000, Math.floor(Number(options.timeoutMs))) : AI_REQUEST_TIMEOUT_MS;
  let abortedByTimeout = false;
  const onExternalAbort = () => controller.abort();
  if (options.signal) {
    if (options.signal.aborted) {
      controller.abort();
    } else {
      options.signal.addEventListener("abort", onExternalAbort, { once: true });
    }
  }
  const timer = setTimeout(() => {
    abortedByTimeout = true;
    controller.abort();
  }, timeout);
  try {
    const response = await fetch(targetUrl, {
      method: options.method || "GET",
      headers: options.headers || undefined,
      body: options.body || undefined,
      signal: controller.signal,
      cache: "no-store",
    });
    const text = await response.text();
    let data = null;
    if (text) {
      try {
        data = JSON.parse(text);
      } catch {
        data = { rawText: text };
      }
    }
    return { ok: response.ok, status: response.status, data, text };
  } catch (error) {
    if (error?.name === "AbortError") {
      const wrapped = new Error(
        abortedByTimeout
          ? `请求超时（${Math.ceil(timeout / 1000)}秒），模型可能仍在加载，请稍后重试`
          : "请求已取消"
      );
      wrapped.code = abortedByTimeout ? "REQUEST_TIMEOUT" : "REQUEST_ABORTED";
      wrapped.status = abortedByTimeout ? 504 : 408;
      throw wrapped;
    }
    throw error;
  } finally {
    clearTimeout(timer);
    if (options.signal) options.signal.removeEventListener("abort", onExternalAbort);
  }
}

function sanitizeAiRemoteImageUrl(rawValue) {
  const raw = String(rawValue || "").trim();
  if (!raw || raw.length > 3000) return "";
  try {
    const parsed = new URL(raw);
    if (!/^https?:$/i.test(parsed.protocol)) return "";
    return parsed.toString();
  } catch {
    return "";
  }
}

async function fetchAiBinary(targetUrl, options = {}) {
  if (!targetUrl) {
    const error = new Error("目标地址无效");
    error.code = "INVALID_TARGET_URL";
    error.status = 400;
    throw error;
  }
  const controller = new AbortController();
  const timeout = Number.isFinite(Number(options.timeoutMs)) ? Math.max(5000, Math.floor(Number(options.timeoutMs))) : AI_REQUEST_TIMEOUT_MS;
  const maxBytes = Number.isFinite(Number(options.maxBytes)) ? Math.max(1024, Math.floor(Number(options.maxBytes))) : AI_IMAGE_FETCH_MAX_BYTES;
  let abortedByTimeout = false;
  const onExternalAbort = () => controller.abort();
  if (options.signal) {
    if (options.signal.aborted) {
      controller.abort();
    } else {
      options.signal.addEventListener("abort", onExternalAbort, { once: true });
    }
  }
  const timer = setTimeout(() => {
    abortedByTimeout = true;
    controller.abort();
  }, timeout);
  try {
    const response = await fetch(targetUrl, {
      method: options.method || "GET",
      headers: options.headers || undefined,
      body: options.body || undefined,
      signal: controller.signal,
      cache: "no-store",
      redirect: "follow",
    });
    const contentLength = Number(response.headers.get("content-length") || 0);
    if (Number.isFinite(contentLength) && contentLength > maxBytes) {
      const sizeError = new Error(`远程图片过大（>${Math.floor(maxBytes / 1024 / 1024)}MB）`);
      sizeError.code = "REMOTE_IMAGE_TOO_LARGE";
      sizeError.status = 413;
      throw sizeError;
    }
    const buffer = Buffer.from(await response.arrayBuffer());
    if (buffer.length > maxBytes) {
      const sizeError = new Error(`远程图片过大（>${Math.floor(maxBytes / 1024 / 1024)}MB）`);
      sizeError.code = "REMOTE_IMAGE_TOO_LARGE";
      sizeError.status = 413;
      throw sizeError;
    }
    return {
      ok: response.ok,
      status: response.status,
      contentType: String(response.headers.get("content-type") || "").trim(),
      buffer,
      text: buffer.length ? buffer.toString("utf8", 0, Math.min(buffer.length, 1200)) : "",
    };
  } catch (error) {
    if (error?.name === "AbortError") {
      const wrapped = new Error(
        abortedByTimeout
          ? `请求超时（${Math.ceil(timeout / 1000)}秒），请稍后重试`
          : "请求已取消"
      );
      wrapped.code = abortedByTimeout ? "REQUEST_TIMEOUT" : "REQUEST_ABORTED";
      wrapped.status = abortedByTimeout ? 504 : 408;
      throw wrapped;
    }
    throw error;
  } finally {
    clearTimeout(timer);
    if (options.signal) options.signal.removeEventListener("abort", onExternalAbort);
  }
}

async function fetchAiText(targetUrl, options = {}) {
  if (!targetUrl) {
    const error = new Error("目标地址无效");
    error.code = "INVALID_TARGET_URL";
    error.status = 400;
    throw error;
  }
  const controller = new AbortController();
  const timeout = Number.isFinite(Number(options.timeoutMs)) ? Math.max(3000, Math.floor(Number(options.timeoutMs))) : AI_REQUEST_TIMEOUT_MS;
  const maxBytes = Number.isFinite(Number(options.maxBytes)) ? Math.max(1024, Math.floor(Number(options.maxBytes))) : AI_WEBPAGE_FETCH_MAX_BYTES;
  let abortedByTimeout = false;
  const onExternalAbort = () => controller.abort();
  if (options.signal) {
    if (options.signal.aborted) {
      controller.abort();
    } else {
      options.signal.addEventListener("abort", onExternalAbort, { once: true });
    }
  }
  const timer = setTimeout(() => {
    abortedByTimeout = true;
    controller.abort();
  }, timeout);
  try {
    const response = await fetch(targetUrl, {
      method: options.method || "GET",
      headers: options.headers || undefined,
      body: options.body || undefined,
      signal: controller.signal,
      cache: "no-store",
      redirect: "follow",
    });
    const contentLength = Number(response.headers.get("content-length") || 0);
    if (Number.isFinite(contentLength) && contentLength > maxBytes) {
      const sizeError = new Error(`网页内容过大（>${Math.floor(maxBytes / 1024)}KB）`);
      sizeError.code = "WEBPAGE_FETCH_TOO_LARGE";
      sizeError.status = 413;
      throw sizeError;
    }
    const buffer = Buffer.from(await response.arrayBuffer());
    if (buffer.length > maxBytes) {
      const sizeError = new Error(`网页内容过大（>${Math.floor(maxBytes / 1024)}KB）`);
      sizeError.code = "WEBPAGE_FETCH_TOO_LARGE";
      sizeError.status = 413;
      throw sizeError;
    }
    const text = buffer.length ? buffer.toString("utf8") : "";
    return {
      ok: response.ok,
      status: response.status,
      contentType: String(response.headers.get("content-type") || "").trim(),
      text,
      sizeBytes: buffer.length,
    };
  } catch (error) {
    if (error?.name === "AbortError") {
      const wrapped = new Error(
        abortedByTimeout
          ? `请求超时（${Math.ceil(timeout / 1000)}秒），请稍后重试`
          : "请求已取消"
      );
      wrapped.code = abortedByTimeout ? "REQUEST_TIMEOUT" : "REQUEST_ABORTED";
      wrapped.status = abortedByTimeout ? 504 : 408;
      throw wrapped;
    }
    throw error;
  } finally {
    clearTimeout(timer);
    if (options.signal) options.signal.removeEventListener("abort", onExternalAbort);
  }
}

function isAiBlockedPrivateHostname(hostname) {
  const host = String(hostname || "").trim().toLowerCase();
  if (!host) return true;
  if (host === "localhost" || host === "::1" || host.endsWith(".local")) return true;
  if (/^127\./.test(host)) return true;
  if (/^10\./.test(host)) return true;
  if (/^192\.168\./.test(host)) return true;
  if (/^169\.254\./.test(host)) return true;
  const octets = host.match(/^(\d{1,3})(?:\.(\d{1,3})){3}$/);
  if (octets) {
    const segments = host.split(".").map((part) => Number(part));
    if (segments.some((part) => !Number.isInteger(part) || part < 0 || part > 255)) return true;
    if (segments[0] === 172 && segments[1] >= 16 && segments[1] <= 31) return true;
    return false;
  }
  if (/^(?:\[)?[a-f0-9:]+(?:\])?$/i.test(host)) {
    const safe = host.replace(/^\[|\]$/g, "");
    if (safe === "::1") return true;
    if (/^(fc|fd)/i.test(safe)) return true;
    if (/^fe80:/i.test(safe)) return true;
  }
  return false;
}

function sanitizeAiWebPageUrl(rawValue) {
  const raw = String(rawValue || "").trim();
  if (!raw || raw.length > 3000) return "";
  let normalized = raw;
  if (!/^[a-zA-Z][a-zA-Z\d+.-]*:/.test(normalized)) {
    normalized = `https://${normalized}`;
  }
  try {
    const parsed = new URL(normalized);
    if (!/^https?:$/i.test(parsed.protocol)) return "";
    if (isAiBlockedPrivateHostname(parsed.hostname)) return "";
    parsed.hash = "";
    return parsed.toString();
  } catch {
    return "";
  }
}

function sanitizeAiWebPageExtractedText(rawText, maxLength = AI_WEBPAGE_EXTRACT_TEXT_LIMIT) {
  const safeMax = Number.isFinite(Number(maxLength)) ? Math.max(1000, Math.floor(Number(maxLength))) : AI_WEBPAGE_EXTRACT_TEXT_LIMIT;
  return decodeHtmlEntities(String(rawText || ""))
    .replace(/\u0000/g, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\u00a0/g, " ")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim()
    .slice(0, safeMax);
}

function extractAiWebPageTitle(htmlText) {
  const matched = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(String(htmlText || ""));
  if (!matched) return "";
  return sanitizeAiWebPageExtractedText(matched[1], 240);
}

function extractAiWebPageTextFromHtml(htmlText, maxLength = AI_WEBPAGE_EXTRACT_TEXT_LIMIT) {
  const raw = String(htmlText || "");
  if (!raw) return "";
  const withoutNoise = raw
    .replace(/<!--[\s\S]*?-->/g, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, " ")
    .replace(/<svg[\s\S]*?<\/svg>/gi, " ")
    .replace(/<template[\s\S]*?<\/template>/gi, " ")
    .replace(/<iframe[\s\S]*?<\/iframe>/gi, " ")
    .replace(/<\/(p|div|article|section|header|footer|aside|li|ul|ol|h[1-6]|tr|table|blockquote|pre)>/gi, "\n")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<[^>]+>/g, " ");
  return sanitizeAiWebPageExtractedText(withoutNoise, maxLength);
}

async function executeAiWebPageExtractRequest(rawPayload) {
  const targetUrl = sanitizeAiWebPageUrl(rawPayload?.url);
  if (!targetUrl) {
    const error = new Error("网页链接无效，或属于受限地址");
    error.code = "WEBPAGE_URL_INVALID";
    error.status = 400;
    throw error;
  }
  const maxCharsRaw = Number(rawPayload?.maxChars);
  const maxChars = Number.isFinite(maxCharsRaw)
    ? Math.max(1000, Math.min(AI_WEBPAGE_EXTRACT_TEXT_LIMIT, Math.floor(maxCharsRaw)))
    : AI_WEBPAGE_EXTRACT_TEXT_LIMIT;
  const fetched = await fetchAiText(targetUrl, {
    timeoutMs: AI_WEBPAGE_FETCH_TIMEOUT_MS,
    maxBytes: AI_WEBPAGE_FETCH_MAX_BYTES,
  });
  if (!fetched.ok) {
    const error = new Error(`网页抓取失败（${fetched.status}）`);
    error.code = "WEBPAGE_FETCH_FAILED";
    error.status = fetched.status || 502;
    throw error;
  }
  const contentType = String(fetched.contentType || "").toLowerCase();
  const rawText = String(fetched.text || "");
  const isHtml = contentType.includes("text/html") || /<html[\s>]|<!doctype\s+html/i.test(rawText);
  const title = isHtml ? extractAiWebPageTitle(rawText) : "";
  const fullText = isHtml
    ? extractAiWebPageTextFromHtml(rawText, maxChars * 2)
    : sanitizeAiWebPageExtractedText(rawText, maxChars * 2);
  if (!fullText) {
    const error = new Error("网页未提取到可用正文");
    error.code = "WEBPAGE_TEXT_EMPTY";
    error.status = 422;
    throw error;
  }
  const truncated = fullText.length > maxChars;
  const text = fullText.slice(0, maxChars);
  return {
    url: targetUrl,
    title,
    text,
    characterCount: fullText.length,
    truncated,
    contentType: fetched.contentType,
  };
}

function clampAiInteger(rawValue, fallbackValue, minValue = 1, maxValue = Number.MAX_SAFE_INTEGER) {
  const fallback = Number.isFinite(Number(fallbackValue)) ? Math.floor(Number(fallbackValue)) : minValue;
  const safeFallback = Math.max(minValue, Math.min(maxValue, fallback));
  const value = Number(rawValue);
  if (!Number.isFinite(value)) return safeFallback;
  return Math.max(minValue, Math.min(maxValue, Math.floor(value)));
}

function extractAiFirstHttpUrlFromText(rawText) {
  const text = String(rawText || "");
  if (!text) return "";
  const urlRegex = /https?:\/\/[^\s"'`<>]+/gi;
  let matched = urlRegex.exec(text);
  while (matched) {
    const rawCandidate = String(matched[0] || "").trim().replace(/[),.;!?]+$/g, "");
    const candidate = sanitizeAiWebPageUrl(rawCandidate);
    if (candidate) return candidate;
    matched = urlRegex.exec(text);
  }
  return "";
}

function detectAiBrowserAgentPromptOpenIntent(rawText) {
  const text = sanitizeAiWebSearchText(rawText, 1600);
  if (!text) return false;
  return /(打开|访问|前往|open|visit|go\s*to)/i.test(text);
}

function resolveAiBrowserAgentSearchEngineFromPrompt(rawText) {
  const text = sanitizeAiWebSearchText(rawText, 1600).toLowerCase();
  if (!text) return AI_WEB_SEARCH_ENGINE_BING;
  if (/(bing|必应)/i.test(text)) return AI_WEB_SEARCH_ENGINE_BING;
  if (/(baidu|百度)/i.test(text)) return AI_WEB_SEARCH_ENGINE_BAIDU;
  if (/(google|谷歌)/i.test(text)) return AI_WEB_SEARCH_ENGINE_GOOGLE;
  return AI_WEB_SEARCH_ENGINE_BING;
}

function resolveAiBrowserAgentSearchSiteFromPrompt(rawText) {
  const text = sanitizeAiWebSearchText(rawText, 1600).toLowerCase();
  if (!text) return "";
  if (/(youtube|youtu\.be|油管)/i.test(text)) return "youtube";
  if (/(douyin|抖音)/i.test(text)) return "douyin";
  if (/(bilibili|b站|哔哩哔哩)/i.test(text)) return "bilibili";
  if (/(zhihu|知乎)/i.test(text)) return "zhihu";
  if (/(xiaohongshu|xhs|rednote|小红书)/i.test(text)) return "xiaohongshu";
  if (/(京东|jd\.com|\bjd\b)/i.test(text)) return "jd";
  if (/(淘宝|taobao)/i.test(text)) return "taobao";
  if (/(tmall|天猫)/i.test(text)) return "tmall";
  if (/(携程|ctrip)/i.test(text)) return "ctrip";
  if (/(12306|铁路12306)/i.test(text)) return "12306";
  if (/(学信网|chsi|xuexin)/i.test(text)) return "xuexin";
  if (/(同程|同程旅游|ly\.com)/i.test(text)) return "tongcheng";
  if (/(闲鱼|xianyu)/i.test(text)) return "xianyu";
  if (/(高德地图|高德|amap)/i.test(text)) return "amap";
  if (/(百度地图|baidu\s*map|map\.baidu)/i.test(text)) return "baidu-map";
  if (/(中国大学\s*mooc|icourse163|mooc)/i.test(text)) return "icourse163";
  return "";
}

function isAiBrowserAgentSearchIntent(rawText) {
  const text = sanitizeAiWebSearchText(rawText, 1600);
  if (!text) return false;
  return /(?:搜索|检索|查询|查找|搜|search(?:\s+for)?)/i.test(text);
}

function getAiBrowserAgentSiteAliasPattern(siteId) {
  const site = sanitizeAiWebSearchText(siteId, 40).toLowerCase();
  if (site === "youtube") return /(youtube|youtu\.be|油管)/gi;
  if (site === "douyin") return /(douyin|抖音)/gi;
  if (site === "bilibili") return /(bilibili|b站|哔哩哔哩)/gi;
  if (site === "zhihu") return /(zhihu|知乎)/gi;
  if (site === "xiaohongshu") return /(xiaohongshu|xhs|rednote|小红书)/gi;
  if (site === "jd") return /(京东|jd\.com|\bjd\b)/gi;
  if (site === "taobao") return /(淘宝|taobao)/gi;
  if (site === "tmall") return /(tmall|天猫)/gi;
  if (site === "ctrip") return /(携程|ctrip)/gi;
  if (site === "12306") return /(12306|铁路12306)/gi;
  if (site === "xuexin") return /(学信网|chsi|xuexin)/gi;
  if (site === "tongcheng") return /(同程|同程旅游|ly\.com)/gi;
  if (site === "xianyu") return /(闲鱼|xianyu)/gi;
  if (site === "amap") return /(高德地图|高德|amap)/gi;
  if (site === "baidu-map") return /(百度地图|baidu\s*map|map\.baidu)/gi;
  if (site === "icourse163") return /(中国大学\s*mooc|icourse163|mooc)/gi;
  return null;
}

function extractAiBrowserAgentImplicitKeywordFromPrompt(rawText, siteId = "") {
  const text = sanitizeAiWebSearchText(rawText, 1600);
  if (!text) return "";
  const quoted = extractAiWebSearchQuotedTerms(text);
  if (quoted.length) return sanitizeAiWebSearchQuery(quoted[0]);
  const site = sanitizeAiWebSearchText(siteId, 40).toLowerCase();
  const normalizeCandidate = (rawValue) => {
    let candidate = sanitizeAiWebSearchQuery(rawValue || "");
    candidate = candidate
      .replace(/[“”"'`]/g, " ")
      .replace(/^(?:到|去|关于|有关|一下|一下子|相关|的)\s*/i, "")
      .replace(/\b(?:please|plz|pls)\b/gi, " ")
      .replace(/(?:怎么买|怎么选|哪个好|哪款好|哪家好|推荐吗|求推荐|推荐下|推荐一下|多少钱|价格|价位|对比|测评|评测)$/i, "")
      .replace(/\s+/g, " ")
      .trim();
    return sanitizeAiWebSearchQuery(candidate).slice(0, 120);
  };
  if (site === "12306") {
    const routeFromToMatched = text.match(
      /(?:^|[^A-Za-z\u4e00-\u9fff])(?:从)?([A-Za-z\u4e00-\u9fff]{1,10})\s*(?:到|至)\s*([A-Za-z\u4e00-\u9fff]{1,10}?)(?=(?:怎么买|怎么去|怎么到|怎么走|火车票|高铁票|动车票|车票|机票|余票|票|$|[，。！？、,.!?;；:：\s]))/
    );
    if (routeFromToMatched && routeFromToMatched[1] && routeFromToMatched[2]) {
      const fromCity = sanitizeAiWebSearchQuery(routeFromToMatched[1]).replace(/^(?:查|搜|找|看|买|购|订|去|到)/, "");
      const toCity = sanitizeAiWebSearchQuery(routeFromToMatched[2]).replace(
        /(?:怎么买|怎么去|怎么到|怎么走|火车票|高铁票|动车票|车票|机票|余票|票)+$/g,
        ""
      );
      const routeKeyword = normalizeCandidate(`${fromCity}到${toCity}`);
      if (routeKeyword) return routeKeyword;
    }
    const routeMatched = text.match(
      /(?:^|[^A-Za-z\u4e00-\u9fff])([A-Za-z\u4e00-\u9fff]{1,10}\s*(?:到|至)\s*[A-Za-z\u4e00-\u9fff]{1,10}?)(?=(?:怎么买|怎么去|怎么到|怎么走|火车票|高铁票|动车票|车票|机票|余票|票|$|[，。！？、,.!?;；:：\s]))/
    );
    if (routeMatched && routeMatched[1]) {
      const routeKeyword = normalizeCandidate(
        String(routeMatched[1] || "")
          .replace(/^(?:查|搜|找|看|买|购|订)/, "")
          .replace(/(?:火车票|高铁票|动车票|车票|机票|余票|票)+$/g, "")
      );
      if (routeKeyword) return routeKeyword;
    }
  }
  if (site === "amap" || site === "baidu-map") {
    const nearbyMatched = text.match(/(?:附近|周边)\s*(.{1,60})$/i);
    if (nearbyMatched && nearbyMatched[1]) {
      const nearbyKeyword = normalizeCandidate(nearbyMatched[1]);
      if (nearbyKeyword) return nearbyKeyword;
    }
    const routeHowMatched =
      text.match(/(?:去|到)\s*(.{1,60})\s*(?:怎么走|怎么去|怎么到|路线|导航)/i) ||
      text.match(/(?:怎么走|怎么去|怎么到)\s*(?:去|到)?\s*(.{1,60})$/i);
    if (routeHowMatched && routeHowMatched[1]) {
      const routeKeyword = normalizeCandidate(routeHowMatched[1]);
      if (routeKeyword) return routeKeyword;
    }
    const navMatched =
      text.match(/(?:导航(?:到|去)?|前往|路线(?:到|去)?|开车去|坐车去)\s*(.{1,80})$/i) || text.match(/(?:去|到)\s*(.{1,80})$/i);
    if (navMatched && navMatched[1]) {
      const navKeyword = normalizeCandidate(navMatched[1]);
      if (navKeyword) return navKeyword;
    }
  }
  const actionMatched =
    text.match(
      /(?:买|购买|想买|要买|想入手|入手|求推荐|推荐|推荐下|推荐一下|安利|种草|找|查|看看|看|学|学习|搜|搜索|检索|查询|查找|订|预订|预定|订票|买票|购票|book|buy|find|search(?:\s+for)?)\s*(?:一下|下|一张|一趟|一个|一间|一件)?\s*(.{1,100})$/i
    ) || text.match(/(?:关于|有关)\s*(.{1,100})$/i);
  if (actionMatched && actionMatched[1]) {
    const actionKeyword = normalizeCandidate(
      String(actionMatched[1] || "")
        .replace(/^(?:我想|我想要|我要|想|求|想看|想学|想买)\s*/i, "")
        .replace(/(?:吧|吗|呢|呀|啊)\s*$/i, "")
    );
    if (actionKeyword) return actionKeyword;
  }
  let fallback = text
    .replace(/https?:\/\/[^\s"'`<>]+/gi, " ")
    .replace(/[“”"'`]/g, " ")
    .replace(/(?:请你|请|帮我|麻烦你|直接|立刻|马上|现在|一下|帮忙|我想要|我想|我要|我想入手|求推荐)\s*/gi, " ")
    .replace(/(?:打开|访问|前往|进入|open|visit|go\s*to)\s*/gi, " ");
  const siteAliasPattern = getAiBrowserAgentSiteAliasPattern(site);
  if (siteAliasPattern) fallback = fallback.replace(siteAliasPattern, " ");
  fallback = fallback
    .replace(
      /(?:搜索|检索|查询|查找|搜|search(?:\s+for)?|买|购买|想入手|入手|求推荐|推荐|安利|种草|找|看|学|学习|订|预订|预定|订票|买票|购票|导航(?:到|去)?|前往|路线(?:到|去)?|开车去|坐车去|怎么走|怎么去|怎么到|去|到|附近|周边)\s*/gi,
      " "
    )
    .replace(/[，。！？、,.!?;；:：()（）[\]【】]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return normalizeCandidate(fallback);
}

function extractAiBrowserAgentSearchKeywordFromPrompt(rawText) {
  const text = sanitizeAiWebSearchText(rawText, 1600);
  if (!text) return "";
  const quoted = extractAiWebSearchQuotedTerms(text);
  if (quoted.length) return sanitizeAiWebSearchQuery(quoted[0]);
  const matched =
    text.match(/(?:搜索|检索|查询|查找|搜|search(?:\s+for)?)\s*(?:[:：]|\s+)?(.{1,120})$/i) ||
    text.match(/(?:搜索|检索|查询|查找|搜|search(?:\s+for)?)\s*(?:["“”'‘’])([^"“”'‘’]{1,120})(?:["“”'‘’])/i);
  if (matched && matched[1]) return sanitizeAiWebSearchQuery(matched[1]);
  const normalized = normalizeAiWebSearchTaskQuery(text);
  return sanitizeAiWebSearchQuery(normalized);
}

function buildAiBrowserAgentEngineSearchUrl(engineId, keyword) {
  const safeKeyword = sanitizeAiWebSearchQuery(keyword);
  if (!safeKeyword) return "";
  const engine = sanitizeAiWebSearchEngineId(engineId) || AI_WEB_SEARCH_ENGINE_BING;
  if (engine === AI_WEB_SEARCH_ENGINE_BAIDU) {
    return `https://www.baidu.com/s?wd=${encodeURIComponent(safeKeyword)}`;
  }
  if (engine === AI_WEB_SEARCH_ENGINE_GOOGLE) {
    return `https://www.google.com/search?q=${encodeURIComponent(safeKeyword)}`;
  }
  return `https://www.bing.com/search?q=${encodeURIComponent(safeKeyword)}`;
}

function buildAiBrowserAgentSiteSearchUrl(siteId, keyword) {
  const safeKeyword = sanitizeAiWebSearchQuery(keyword);
  if (!safeKeyword) return "";
  const site = sanitizeAiWebSearchText(siteId, 40).toLowerCase();
  if (site === "youtube") return `https://www.youtube.com/results?search_query=${encodeURIComponent(safeKeyword)}`;
  if (site === "douyin") return `https://www.douyin.com/search/${encodeURIComponent(safeKeyword)}`;
  if (site === "bilibili") return `https://search.bilibili.com/all?keyword=${encodeURIComponent(safeKeyword)}`;
  if (site === "zhihu") return `https://www.zhihu.com/search?type=content&q=${encodeURIComponent(safeKeyword)}`;
  if (site === "xiaohongshu") return `https://www.xiaohongshu.com/search_result?keyword=${encodeURIComponent(safeKeyword)}`;
  if (site === "jd") return `https://search.jd.com/Search?keyword=${encodeURIComponent(safeKeyword)}`;
  if (site === "taobao") return `https://s.taobao.com/search?q=${encodeURIComponent(safeKeyword)}`;
  if (site === "tmall") return `https://list.tmall.com/search_product.htm?q=${encodeURIComponent(safeKeyword)}`;
  if (site === "ctrip") return `https://you.ctrip.com/searchsite/Sight?query=${encodeURIComponent(safeKeyword)}`;
  if (site === "12306") return `https://www.12306.cn/index/`;
  if (site === "xuexin") return `https://www.chsi.com.cn/`;
  if (site === "tongcheng") return `https://so.ly.com/scenery?q=${encodeURIComponent(safeKeyword)}`;
  if (site === "xianyu") return `https://www.goofish.com/search?q=${encodeURIComponent(safeKeyword)}`;
  if (site === "amap") return `https://ditu.amap.com/search?query=${encodeURIComponent(safeKeyword)}`;
  if (site === "baidu-map") return `https://map.baidu.com/search/${encodeURIComponent(safeKeyword)}`;
  if (site === "icourse163") return `https://www.icourse163.org/search.htm?search=${encodeURIComponent(safeKeyword)}`;
  return "";
}

function buildAiBrowserAgentSiteHomeUrl(siteId) {
  const site = sanitizeAiWebSearchText(siteId, 40).toLowerCase();
  if (site === "youtube") return "https://www.youtube.com/";
  if (site === "douyin") return "https://www.douyin.com/";
  if (site === "bilibili") return "https://www.bilibili.com/";
  if (site === "zhihu") return "https://www.zhihu.com/";
  if (site === "xiaohongshu") return "https://www.xiaohongshu.com/";
  if (site === "jd") return "https://www.jd.com/";
  if (site === "taobao") return "https://www.taobao.com/";
  if (site === "tmall") return "https://www.tmall.com/";
  if (site === "ctrip") return "https://www.ctrip.com/";
  if (site === "12306") return "https://www.12306.cn/index/";
  if (site === "xuexin") return "https://www.chsi.com.cn/";
  if (site === "tongcheng") return "https://www.ly.com/";
  if (site === "xianyu") return "https://www.goofish.com/";
  if (site === "amap") return "https://ditu.amap.com/";
  if (site === "baidu-map") return "https://map.baidu.com/";
  if (site === "icourse163") return "https://www.icourse163.org/";
  return "";
}

function resolveAiBrowserAgentPromptNavigationIntent(rawText) {
  const text = sanitizeAiWebSearchText(rawText, 1600);
  if (!text) {
    return {
      targetUrl: "",
      targetSource: "",
      openInDesktopSuggested: false,
      searchKeyword: "",
      searchEngine: "",
      searchSite: "",
    };
  }
  const openIntent = detectAiBrowserAgentPromptOpenIntent(text);
  const searchSite = resolveAiBrowserAgentSearchSiteFromPrompt(text);
  const explicitSearchIntent = isAiBrowserAgentSearchIntent(text);
  const implicitSearchKeyword = searchSite ? extractAiBrowserAgentImplicitKeywordFromPrompt(text, searchSite) : "";
  const searchIntent = explicitSearchIntent || Boolean(implicitSearchKeyword);
  if (!searchIntent) {
    const homeTargetUrl = openIntent && searchSite ? sanitizeAiWebPageUrl(buildAiBrowserAgentSiteHomeUrl(searchSite)) : "";
    return {
      targetUrl: homeTargetUrl,
      targetSource: homeTargetUrl ? "prompt-site-home-url" : "",
      openInDesktopSuggested: openIntent,
      searchKeyword: "",
      searchEngine: "",
      searchSite,
    };
  }
  const searchKeyword = explicitSearchIntent ? extractAiBrowserAgentSearchKeywordFromPrompt(text) || implicitSearchKeyword : implicitSearchKeyword;
  if (!searchKeyword) {
    const homeTargetUrl = openIntent && searchSite ? sanitizeAiWebPageUrl(buildAiBrowserAgentSiteHomeUrl(searchSite)) : "";
    return {
      targetUrl: homeTargetUrl,
      targetSource: homeTargetUrl ? "prompt-site-home-url" : "",
      openInDesktopSuggested: openIntent,
      searchKeyword: "",
      searchEngine: "",
      searchSite,
    };
  }
  const searchEngine = resolveAiBrowserAgentSearchEngineFromPrompt(text);
  const rawUrl = searchSite
    ? buildAiBrowserAgentSiteSearchUrl(searchSite, searchKeyword)
    : buildAiBrowserAgentEngineSearchUrl(searchEngine, searchKeyword);
  const targetUrl = sanitizeAiWebPageUrl(rawUrl);
  return {
    targetUrl,
    targetSource: targetUrl ? (searchSite ? "prompt-site-search-url" : "prompt-search-url") : "",
    openInDesktopSuggested: openIntent,
    searchKeyword,
    searchEngine,
    searchSite,
  };
}

function openAiBrowserAgentDesktopUrl(rawUrl) {
  const safeUrl = sanitizeAiWebPageUrl(rawUrl);
  if (!safeUrl) {
    const error = new Error("桌面打开失败：URL 无效");
    error.code = "BROWSER_AGENT_DESKTOP_URL_INVALID";
    throw error;
  }
  let command = "";
  let args = [];
  if (process.platform === "win32") {
    command = "cmd.exe";
    args = ["/c", "start", "", safeUrl];
  } else if (process.platform === "darwin") {
    command = "open";
    args = [safeUrl];
  } else {
    command = "xdg-open";
    args = [safeUrl];
  }
  const child = spawn(command, args, {
    detached: true,
    stdio: "ignore",
    windowsHide: true,
  });
  child.unref();
  return safeUrl;
}

function resolveAiBrowserAgentExecutablePath() {
  const envPath = String(AI_BROWSER_AGENT_EXECUTABLE_PATH_ENV || "").trim();
  if (envPath) {
    try {
      if (fs.existsSync(envPath)) return envPath;
    } catch {
      // ignore invalid configured path
    }
  }

  const candidates = [];
  if (process.platform === "win32") {
    const programFiles = String(process.env.ProgramFiles || "C:\\Program Files").trim();
    const programFilesX86 = String(process.env["ProgramFiles(x86)"] || "C:\\Program Files (x86)").trim();
    const localAppData = String(process.env.LOCALAPPDATA || "").trim();
    candidates.push(
      path.join(programFiles, "Google", "Chrome", "Application", "chrome.exe"),
      path.join(programFilesX86, "Google", "Chrome", "Application", "chrome.exe"),
      path.join(programFiles, "Microsoft", "Edge", "Application", "msedge.exe"),
      path.join(programFilesX86, "Microsoft", "Edge", "Application", "msedge.exe")
    );
    if (localAppData) {
      candidates.push(
        path.join(localAppData, "Google", "Chrome", "Application", "chrome.exe"),
        path.join(localAppData, "Microsoft", "Edge", "Application", "msedge.exe")
      );
    }
  } else if (process.platform === "darwin") {
    candidates.push(
      "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
      "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
      "/Applications/Chromium.app/Contents/MacOS/Chromium"
    );
  } else {
    candidates.push(
      "/usr/bin/google-chrome",
      "/usr/bin/google-chrome-stable",
      "/usr/bin/chromium",
      "/usr/bin/chromium-browser",
      "/snap/bin/chromium"
    );
  }

  for (const candidate of candidates) {
    if (!candidate) continue;
    try {
      if (fs.existsSync(candidate)) return candidate;
    } catch {
      // ignore
    }
  }
  return "";
}

function normalizeAiBrowserAgentAutofillEnabled(rawValue) {
  if (rawValue === true) return true;
  const value = String(rawValue || "").trim().toLowerCase();
  if (!value) return false;
  return ["1", "true", "yes", "on", "enable", "enabled"].includes(value);
}

function normalizeAiBrowserAgentAutofillEntries(rawValue) {
  const output = [];
  const seen = new Set();
  const pushEntry = (entry) => {
    if (!entry || typeof entry !== "object") return;
    if (output.length >= AI_BROWSER_AGENT_MAX_AUTOFILL_FIELDS) return;
    const selector = String(entry.selector || "").trim().slice(0, 220);
    const name = String(entry.name || "").trim().slice(0, 120);
    const label = String(entry.label || "").trim().slice(0, 120);
    const value = String(entry.value ?? "").replace(/\u0000/g, "").slice(0, 600);
    if (!value && value !== "") return;
    if (!selector && !name && !label) return;
    const dedupeKey = `${selector.toLowerCase()}|${name.toLowerCase()}|${label.toLowerCase()}`;
    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);
    output.push({ selector, name, label, value });
  };
  if (Array.isArray(rawValue)) {
    rawValue.forEach((item) => pushEntry(item));
    return output;
  }
  if (rawValue && typeof rawValue === "object") {
    Object.entries(rawValue).forEach(([name, value]) => {
      pushEntry({ name, value: String(value ?? "") });
    });
  }
  return output;
}

async function executeAiResearchRequest(rawPayload) {
  const query = normalizeAiWebSearchTaskQuery(rawPayload?.query || rawPayload?.prompt || rawPayload?.text || "");
  if (!query) {
    const error = new Error("缂哄皯鏈夋晥璋冪爺鍏抽敭璇?");
    error.code = "SEARCH_QUERY_REQUIRED";
    error.status = 400;
    throw error;
  }
  const maxResults = clampAiInteger(rawPayload?.maxResults, 5, 1, 8);
  const maxPages = clampAiInteger(rawPayload?.maxPages, Math.min(AI_RESEARCH_DEFAULT_MAX_PAGES, maxResults), 1, 8);
  const pageMaxChars = clampAiInteger(
    rawPayload?.pageMaxChars,
    AI_RESEARCH_DEFAULT_PAGE_MAX_CHARS,
    1000,
    AI_WEBPAGE_EXTRACT_TEXT_LIMIT
  );

  const searchResponse = await executeAiWebSearchRequest({
    query,
    maxResults,
    enginePreference: rawPayload?.enginePreference,
    strictEngine: rawPayload?.strictEngine,
    allowEngineFallback: rawPayload?.allowEngineFallback,
    bypassCache: rawPayload?.bypassCache,
    regionHint: rawPayload?.regionHint,
    clientTimeZone: rawPayload?.clientTimeZone,
    clientLocale: rawPayload?.clientLocale,
  });
  const baseResults = Array.isArray(searchResponse?.results) ? searchResponse.results : [];
  const fetchTargets = [];
  for (let index = 0; index < baseResults.length; index += 1) {
    const safeUrl = sanitizeAiWebPageUrl(baseResults[index]?.url);
    if (!safeUrl) continue;
    fetchTargets.push({ index, url: safeUrl });
    if (fetchTargets.length >= maxPages) break;
  }

  const extractedByIndex = new Map();
  const errorByIndex = new Map();
  await Promise.all(
    fetchTargets.map(async (target) => {
      try {
        const extracted = await executeAiWebPageExtractRequest({
          url: target.url,
          maxChars: pageMaxChars,
        });
        extractedByIndex.set(target.index, extracted);
      } catch (error) {
        errorByIndex.set(target.index, {
          message: sanitizeAiWebSearchText(error?.message || "", 220) || "缃戦〉鎶撳彇澶辫触",
          code: sanitizeAiWebSearchText(error?.code || "", 80),
          status: Number.isFinite(Number(error?.status)) ? Number(error.status) : 0,
        });
      }
    })
  );

  const enrichedResults = baseResults.map((entry, index) => {
    const base = entry && typeof entry === "object" ? { ...entry } : {};
    base.title = sanitizeAiWebSearchText(base.title || "", 260);
    base.url = sanitizeAiWebPageUrl(base.url) || sanitizeAiWebSearchText(base.url || "", 800);
    base.snippet = sanitizeAiWebSearchText(base.snippet || "", 900);
    if (extractedByIndex.has(index)) {
      const page = extractedByIndex.get(index);
      base.webpage = {
        url: sanitizeAiWebPageUrl(page?.url) || "",
        title: sanitizeAiWebSearchText(page?.title || "", 260),
        text: sanitizeAiWebPageExtractedText(page?.text || "", pageMaxChars),
        characterCount: Number.isFinite(Number(page?.characterCount)) ? Math.max(0, Math.floor(Number(page.characterCount))) : 0,
        truncated: page?.truncated === true,
        contentType: sanitizeAiWebSearchText(page?.contentType || "", 80),
      };
    } else if (errorByIndex.has(index)) {
      const pageError = errorByIndex.get(index) || {};
      base.webpageError = pageError.message || "缃戦〉鎶撳彇澶辫触";
      if (pageError.code) base.webpageErrorCode = pageError.code;
      if (pageError.status) base.webpageErrorStatus = pageError.status;
    }
    return base;
  });

  return {
    ...searchResponse,
    query,
    results: enrichedResults,
    maxPages,
    pageMaxChars,
    pageFetchCount: fetchTargets.length,
    pageContextCount: extractedByIndex.size,
    generatedAtMs: Date.now(),
  };
}

async function executeAiBrowserAgentRequest(rawPayload) {
  if (!playwrightChromium) {
    const error = new Error("鏈娴嬪埌 playwright-core锛屾棤娉曞惎鐢ㄦ祻瑙堝櫒 Agent");
    error.code = "BROWSER_AGENT_UNAVAILABLE";
    error.status = 503;
    throw error;
  }

  const promptText = sanitizeAiWebSearchText(rawPayload?.prompt || rawPayload?.query || rawPayload?.text || "", 1600);
  const maxLinks = clampAiInteger(rawPayload?.maxLinks, AI_BROWSER_AGENT_DEFAULT_MAX_LINKS, 1, AI_BROWSER_AGENT_DEFAULT_MAX_LINKS);
  const textMaxChars = clampAiInteger(
    rawPayload?.textMaxChars,
    AI_BROWSER_AGENT_DEFAULT_TEXT_MAX_CHARS,
    1000,
    AI_BROWSER_AGENT_DEFAULT_TEXT_MAX_CHARS
  );
  const maxForms = clampAiInteger(rawPayload?.maxForms, AI_BROWSER_AGENT_MAX_FORMS, 1, AI_BROWSER_AGENT_MAX_FORMS);
  const maxFormFields = clampAiInteger(rawPayload?.maxFormFields, AI_BROWSER_AGENT_MAX_FORM_FIELDS, 1, AI_BROWSER_AGENT_MAX_FORM_FIELDS);
  const maxButtons = clampAiInteger(rawPayload?.maxButtons, AI_BROWSER_AGENT_MAX_BUTTONS, 1, AI_BROWSER_AGENT_MAX_BUTTONS);

  const directUrl = sanitizeAiWebPageUrl(rawPayload?.url);
  const promptUrl = directUrl ? "" : extractAiFirstHttpUrlFromText(promptText);
  const promptNavigationIntent = directUrl || promptUrl ? null : resolveAiBrowserAgentPromptNavigationIntent(promptText);
  let targetUrl = directUrl || promptUrl || promptNavigationIntent?.targetUrl || "";
  let targetSource = directUrl
    ? "payload-url"
    : promptUrl
      ? "prompt-url"
      : promptNavigationIntent?.targetSource || (promptNavigationIntent?.targetUrl ? "prompt-search-url" : "");
  let searchFallback = null;

  if (!targetUrl) {
    const query = normalizeAiWebSearchTaskQuery(rawPayload?.query || promptText);
    if (query) {
      searchFallback = await executeAiWebSearchRequest({
        query,
        maxResults: clampAiInteger(rawPayload?.maxResults, 4, 1, 8),
        enginePreference: rawPayload?.enginePreference,
        strictEngine: rawPayload?.strictEngine,
        allowEngineFallback: rawPayload?.allowEngineFallback,
        bypassCache: rawPayload?.bypassCache,
        regionHint: rawPayload?.regionHint,
        clientTimeZone: rawPayload?.clientTimeZone,
        clientLocale: rawPayload?.clientLocale,
      });
      const fallbackResults = Array.isArray(searchFallback?.results) ? searchFallback.results : [];
      for (const item of fallbackResults) {
        const safeUrl = sanitizeAiWebPageUrl(item?.url);
        if (!safeUrl) continue;
        targetUrl = safeUrl;
        targetSource = "search-result";
        break;
      }
    }
  }

  if (!targetUrl) {
    const error = new Error("鏈彁渚涘彲鐢ㄧ綉椤甸摼鎺ワ紝璇疯緭鍏?URL 鎴栧彲妫€绱㈢殑鍏抽敭璇?");
    error.code = "BROWSER_AGENT_TARGET_REQUIRED";
    error.status = 400;
    throw error;
  }

  const openInDesktopRequested = normalizeAiWebSearchBoolean(
    rawPayload?.openInDesktop ?? rawPayload?.desktopOpen,
    promptNavigationIntent?.openInDesktopSuggested === true
  );
  const desktopOpen = {
    requested: openInDesktopRequested,
    attempted: false,
    opened: false,
    url: "",
    error: "",
  };
  if (openInDesktopRequested) {
    desktopOpen.attempted = true;
    try {
      desktopOpen.url = openAiBrowserAgentDesktopUrl(targetUrl);
      desktopOpen.opened = true;
    } catch (error) {
      desktopOpen.error = sanitizeAiWebSearchText(error?.message || "", 220) || "桌面浏览器打开失败";
    }
  }

  const autofillEntries = normalizeAiBrowserAgentAutofillEntries(rawPayload?.formData);
  const autofillEnabled = normalizeAiBrowserAgentAutofillEnabled(rawPayload?.autofill) && autofillEntries.length > 0;
  const executablePath = resolveAiBrowserAgentExecutablePath();
  const launchOptions = {
    headless: true,
    timeout: AI_BROWSER_AGENT_ACTION_TIMEOUT_MS,
  };
  if (executablePath) {
    launchOptions.executablePath = executablePath;
  }

  let browser = null;
  let context = null;
  const allowSoftNavigationFailure = Boolean(promptNavigationIntent?.targetUrl);
  let executionWarning = "";
  try {
    try {
      browser = await playwrightChromium.launch(launchOptions);
    } catch (error) {
      const wrapped = new Error(
        sanitizeAiWebSearchText(error?.message || "", 240) ||
          "娴忚鍣ㄥ惎鍔ㄥけ璐ワ紝璇峰畨瑁匔hrome/Edge 鎴栬缃?AI_BROWSER_AGENT_EXECUTABLE_PATH"
      );
      wrapped.code = "BROWSER_AGENT_LAUNCH_FAILED";
      wrapped.status = 503;
      throw wrapped;
    }

    context = await browser.newContext({
      ignoreHTTPSErrors: true,
    });
    const page = await context.newPage();
    page.setDefaultNavigationTimeout(AI_BROWSER_AGENT_NAV_TIMEOUT_MS);
    page.setDefaultTimeout(AI_BROWSER_AGENT_ACTION_TIMEOUT_MS);
    try {
      await page.goto(targetUrl, {
        waitUntil: "domcontentloaded",
        timeout: AI_BROWSER_AGENT_NAV_TIMEOUT_MS,
      });
    } catch (error) {
      const wrapped = new Error(sanitizeAiWebSearchText(error?.message || "", 220) || "缃戦〉璁块棶澶辫触");
      wrapped.code = "BROWSER_AGENT_NAV_FAILED";
      wrapped.status = 502;
      if (!allowSoftNavigationFailure) throw wrapped;
      executionWarning = sanitizeAiWebSearchText(
        `${wrapped.message}，页面抓取失败但已保留目标链接，可在桌面浏览器继续操作`,
        240
      );
    }

    let pageSnapshot = null;
    if (!executionWarning) {
      try {
        pageSnapshot = await page.evaluate(
          ({ textMaxCharsInner, maxLinksInner, maxFormsInner, maxFormFieldsInner, maxButtonsInner }) => {
            const normalizeText = (value, max = 6000) =>
              String(value || "")
                .replace(/\u0000/g, "")
                .replace(/\r\n/g, "\n")
                .replace(/\r/g, "\n")
                .replace(/[ \t]+\n/g, "\n")
                .replace(/[ \t]{2,}/g, " ")
                .replace(/\n{3,}/g, "\n\n")
                .trim()
                .slice(0, max);

            const readMainText = () => {
              const selectorCandidates = ["main", "article", "[role='main']", ".main", "body"];
              for (const selector of selectorCandidates) {
                const node = document.querySelector(selector);
                if (!node) continue;
                const content = normalizeText(node.textContent || "", textMaxCharsInner * 2);
                if (content.length >= 120) return content;
              }
              return normalizeText(document.body?.textContent || "", textMaxCharsInner * 2);
            };

            const links = [];
            const seenLink = new Set();
            document.querySelectorAll("a[href]").forEach((node) => {
              if (links.length >= maxLinksInner) return;
              const href = String(node.getAttribute("href") || "").trim();
              const absoluteUrl = String(node.href || "").trim();
              if (!href || !absoluteUrl) return;
              if (/^(javascript:|mailto:|tel:|#)/i.test(href)) return;
              const key = absoluteUrl.toLowerCase();
              if (seenLink.has(key)) return;
              seenLink.add(key);
              links.push({
                text: normalizeText(node.textContent || "", 180),
                url: absoluteUrl.slice(0, 1200),
              });
            });

            const forms = [];
            document.querySelectorAll("form").forEach((form, formIndex) => {
              if (forms.length >= maxFormsInner) return;
              const fields = [];
              const fieldNodes = form.querySelectorAll("input, textarea, select");
              fieldNodes.forEach((fieldNode) => {
                if (fields.length >= maxFormFieldsInner) return;
                const tag = String(fieldNode.tagName || "").toLowerCase();
                const type = String(
                  fieldNode.getAttribute("type") || (tag === "textarea" ? "textarea" : tag === "select" ? "select" : "text")
                )
                  .toLowerCase()
                  .slice(0, 40);
                fields.push({
                  name: normalizeText(fieldNode.getAttribute("name") || "", 100),
                  id: normalizeText(fieldNode.getAttribute("id") || "", 100),
                  label: normalizeText(fieldNode.getAttribute("aria-label") || fieldNode.getAttribute("placeholder") || "", 140),
                  tag,
                  type,
                  required: fieldNode.hasAttribute("required"),
                  disabled: fieldNode.disabled === true,
                });
              });
              forms.push({
                index: formIndex,
                action: normalizeText(form.getAttribute("action") || "", 600),
                method: normalizeText(form.getAttribute("method") || "get", 20).toLowerCase(),
                name: normalizeText(form.getAttribute("name") || "", 120),
                id: normalizeText(form.getAttribute("id") || "", 120),
                fields,
              });
            });

            const buttons = [];
            document.querySelectorAll("button, input[type='submit'], input[type='button']").forEach((buttonNode) => {
              if (buttons.length >= maxButtonsInner) return;
              const tag = String(buttonNode.tagName || "").toLowerCase();
              const type = String(buttonNode.getAttribute("type") || (tag === "button" ? "button" : "submit"))
                .toLowerCase()
                .slice(0, 30);
              buttons.push({
                text: normalizeText(buttonNode.textContent || buttonNode.getAttribute("value") || "", 160),
                type,
                id: normalizeText(buttonNode.getAttribute("id") || "", 120),
                name: normalizeText(buttonNode.getAttribute("name") || "", 120),
                disabled: buttonNode.disabled === true,
              });
            });

            return {
              title: normalizeText(document.title || "", 260),
              url: String(location.href || "").trim().slice(0, 1400),
              text: readMainText().slice(0, textMaxCharsInner),
              links,
              forms,
              buttons,
            };
          },
          {
            textMaxCharsInner: textMaxChars,
            maxLinksInner: maxLinks,
            maxFormsInner: maxForms,
            maxFormFieldsInner: maxFormFields,
            maxButtonsInner: maxButtons,
          }
        );
      } catch (error) {
        if (!allowSoftNavigationFailure) {
          const wrapped = new Error(sanitizeAiWebSearchText(error?.message || "", 220) || "页面内容提取失败");
          wrapped.code = "BROWSER_AGENT_NAV_FAILED";
          wrapped.status = 502;
          throw wrapped;
        }
        executionWarning = sanitizeAiWebSearchText(
          `${sanitizeAiWebSearchText(error?.message || "", 180) || "页面内容提取失败"}，已保留目标链接，可在桌面浏览器继续操作`,
          240
        );
      }
    }
    if (!pageSnapshot || typeof pageSnapshot !== "object") {
      pageSnapshot = {
        title: "",
        url: targetUrl,
        text: "",
        links: [],
        forms: [],
        buttons: [],
      };
    }

    let autofill = {
      enabled: autofillEnabled,
      attempted: autofillEntries.length,
      applied: 0,
      failed: autofillEntries.length,
      details: [],
    };

    if (autofillEnabled) {
      const autofillResult = await page.evaluate((entries) => {
        const normalizeText = (value, max = 160) =>
          String(value || "")
            .replace(/\s+/g, " ")
            .trim()
            .slice(0, max);
        const normalizeValue = (value) => String(value ?? "").replace(/\u0000/g, "").slice(0, 600);
        const findByName = (name) => {
          const target = normalizeText(name, 120).toLowerCase();
          if (!target) return null;
          const fields = Array.from(document.querySelectorAll("input, textarea, select"));
          for (const node of fields) {
            const nodeName = normalizeText(node.getAttribute("name") || "", 120).toLowerCase();
            if (nodeName && nodeName === target) return node;
          }
          return null;
        };
        const findByLabel = (label) => {
          const target = normalizeText(label, 120).toLowerCase();
          if (!target) return null;
          const labels = Array.from(document.querySelectorAll("label"));
          for (const labelNode of labels) {
            const text = normalizeText(labelNode.textContent || "", 160).toLowerCase();
            if (!text || !text.includes(target)) continue;
            const htmlFor = String(labelNode.getAttribute("for") || "").trim();
            if (htmlFor) {
              const byId = document.getElementById(htmlFor);
              if (byId && /^(input|textarea|select)$/i.test(byId.tagName || "")) return byId;
            }
            const nested = labelNode.querySelector("input, textarea, select");
            if (nested) return nested;
          }
          return null;
        };
        const details = [];
        let applied = 0;
        let failed = 0;
        entries.forEach((entry, index) => {
          const selector = normalizeText(entry?.selector || "", 220);
          const name = normalizeText(entry?.name || "", 120);
          const label = normalizeText(entry?.label || "", 120);
          const value = normalizeValue(entry?.value || "");
          let targetNode = null;
          if (selector) {
            try {
              targetNode = document.querySelector(selector);
            } catch {
              targetNode = null;
            }
          }
          if (!targetNode && name) targetNode = findByName(name);
          if (!targetNode && label) targetNode = findByLabel(label);
          if (!targetNode || !/^(input|textarea|select)$/i.test(targetNode.tagName || "")) {
            failed += 1;
            details.push({
              index,
              selector,
              name,
              label,
              status: "not-found",
            });
            return;
          }
          if (targetNode.disabled === true || targetNode.readOnly === true) {
            failed += 1;
            details.push({
              index,
              selector,
              name,
              label,
              status: "blocked",
            });
            return;
          }
          const type = String(targetNode.getAttribute("type") || "").toLowerCase();
          if (type === "checkbox" || type === "radio") {
            const checked = !["0", "false", "off", "no"].includes(String(value || "").trim().toLowerCase());
            targetNode.checked = checked;
          } else {
            targetNode.value = value;
          }
          targetNode.dispatchEvent(new Event("input", { bubbles: true }));
          targetNode.dispatchEvent(new Event("change", { bubbles: true }));
          applied += 1;
          details.push({
            index,
            selector,
            name,
            label,
            status: "filled",
          });
        });
        return {
          attempted: entries.length,
          applied,
          failed,
          details: details.slice(0, 40),
        };
      }, autofillEntries);
      autofill = {
        enabled: true,
        attempted: clampAiInteger(autofillResult?.attempted, autofillEntries.length, 0, AI_BROWSER_AGENT_MAX_AUTOFILL_FIELDS),
        applied: clampAiInteger(autofillResult?.applied, 0, 0, AI_BROWSER_AGENT_MAX_AUTOFILL_FIELDS),
        failed: clampAiInteger(autofillResult?.failed, 0, 0, AI_BROWSER_AGENT_MAX_AUTOFILL_FIELDS),
        details: Array.isArray(autofillResult?.details)
          ? autofillResult.details
              .map((item) => ({
                index: clampAiInteger(item?.index, 0, 0, AI_BROWSER_AGENT_MAX_AUTOFILL_FIELDS),
                selector: sanitizeAiWebSearchText(item?.selector || "", 220),
                name: sanitizeAiWebSearchText(item?.name || "", 120),
                label: sanitizeAiWebSearchText(item?.label || "", 120),
                status: sanitizeAiWebSearchText(item?.status || "", 40),
              }))
              .slice(0, AI_BROWSER_AGENT_MAX_AUTOFILL_FIELDS)
          : [],
      };
    }

    return {
      targetUrl,
      targetSource: targetSource || "direct",
      finalUrl: sanitizeAiWebPageUrl(pageSnapshot?.url) || targetUrl,
      title: sanitizeAiWebSearchText(pageSnapshot?.title || "", 260),
      warning: sanitizeAiWebSearchText(executionWarning || "", 240),
      text: sanitizeAiWebPageExtractedText(pageSnapshot?.text || "", textMaxChars),
      links: Array.isArray(pageSnapshot?.links)
        ? pageSnapshot.links
            .map((item) => ({
              text: sanitizeAiWebSearchText(item?.text || "", 220),
              url: sanitizeAiWebPageUrl(item?.url) || sanitizeAiWebSearchText(item?.url || "", 1200),
            }))
            .filter((item) => item.url)
            .slice(0, maxLinks)
        : [],
      forms: Array.isArray(pageSnapshot?.forms)
        ? pageSnapshot.forms
            .map((form) => ({
              index: clampAiInteger(form?.index, 0, 0, AI_BROWSER_AGENT_MAX_FORMS),
              action: sanitizeAiWebSearchText(form?.action || "", 600),
              method: sanitizeAiWebSearchText(form?.method || "", 20).toLowerCase(),
              name: sanitizeAiWebSearchText(form?.name || "", 120),
              id: sanitizeAiWebSearchText(form?.id || "", 120),
              fields: Array.isArray(form?.fields)
                ? form.fields
                    .map((field) => ({
                      name: sanitizeAiWebSearchText(field?.name || "", 120),
                      id: sanitizeAiWebSearchText(field?.id || "", 120),
                      label: sanitizeAiWebSearchText(field?.label || "", 150),
                      tag: sanitizeAiWebSearchText(field?.tag || "", 20),
                      type: sanitizeAiWebSearchText(field?.type || "", 40),
                      required: field?.required === true,
                      disabled: field?.disabled === true,
                    }))
                    .slice(0, maxFormFields)
                : [],
            }))
            .slice(0, maxForms)
        : [],
      buttons: Array.isArray(pageSnapshot?.buttons)
        ? pageSnapshot.buttons
            .map((button) => ({
              text: sanitizeAiWebSearchText(button?.text || "", 180),
              type: sanitizeAiWebSearchText(button?.type || "", 30),
              id: sanitizeAiWebSearchText(button?.id || "", 120),
              name: sanitizeAiWebSearchText(button?.name || "", 120),
              disabled: button?.disabled === true,
            }))
            .slice(0, maxButtons)
        : [],
      search:
        searchFallback && typeof searchFallback === "object"
          ? {
              query: sanitizeAiWebSearchQuery(searchFallback.query || ""),
              queryUsed: sanitizeAiWebSearchQuery(searchFallback.queryUsed || ""),
              source: sanitizeAiWebSearchText(searchFallback.source || "", 80),
              warning: sanitizeAiWebSearchText(searchFallback.warning || "", 200),
              results: Array.isArray(searchFallback.results)
                ? searchFallback.results
                    .map((entry) => ({
                      title: sanitizeAiWebSearchText(entry?.title || "", 220),
                      url: sanitizeAiWebPageUrl(entry?.url) || sanitizeAiWebSearchText(entry?.url || "", 800),
                      snippet: sanitizeAiWebSearchText(entry?.snippet || "", 500),
                    }))
                    .filter((entry) => entry.url || entry.title || entry.snippet)
                    .slice(0, 5)
                : [],
            }
          : null,
      desktopOpen,
      promptNavigation:
        promptNavigationIntent && typeof promptNavigationIntent === "object"
          ? {
              searchKeyword: sanitizeAiWebSearchQuery(promptNavigationIntent.searchKeyword || ""),
              searchEngine: sanitizeAiWebSearchEngineId(promptNavigationIntent.searchEngine || ""),
              searchSite: sanitizeAiWebSearchText(promptNavigationIntent.searchSite || "", 40).toLowerCase(),
            }
          : null,
      autofill,
      executablePath: executablePath || "",
      generatedAtMs: Date.now(),
    };
  } finally {
    if (context) {
      try {
        await context.close();
      } catch {
        // ignore context close error
      }
    }
    if (browser) {
      try {
        await browser.close();
      } catch {
        // ignore browser close error
      }
    }
  }
}

function sanitizeAiImageDataUrlPayload(rawValue) {
  const value = String(rawValue || "").trim();
  if (!value) return "";
  if (value.length > AI_IMAGE_CACHE_BODY_MAX_BYTES * 2) return "";
  return value;
}

function parseAiImageDataUrlToBuffer(dataUrl) {
  const safe = sanitizeAiImageDataUrlPayload(dataUrl);
  if (!safe) return null;
  const matched = safe.match(/^data:(image\/[a-z0-9.+-]+);base64,([A-Za-z0-9+/=\s]+)$/i);
  if (!matched) return null;
  const mimeType = String(matched[1] || "").trim().toLowerCase();
  if (!mimeType.startsWith("image/")) return null;
  const base64 = String(matched[2] || "").replace(/\s+/g, "");
  if (!base64 || !/^[A-Za-z0-9+/=]+$/.test(base64)) return null;
  const buffer = Buffer.from(base64, "base64");
  if (!buffer.length) return null;
  return { mimeType, buffer };
}

function getAiImageExtensionByMimeType(mimeType) {
  const safeMime = String(mimeType || "").trim().toLowerCase();
  if (safeMime.includes("png")) return "png";
  if (safeMime.includes("jpeg") || safeMime.includes("jpg")) return "jpg";
  if (safeMime.includes("webp")) return "webp";
  if (safeMime.includes("gif")) return "gif";
  if (safeMime.includes("bmp")) return "bmp";
  if (safeMime.includes("avif")) return "avif";
  if (safeMime.includes("svg")) return "svg";
  return "";
}

function detectAiImageMimeByBuffer(buffer) {
  if (!Buffer.isBuffer(buffer) || !buffer.length) return "";
  if (
    buffer.length >= 8 &&
    buffer[0] === 0x89 &&
    buffer[1] === 0x50 &&
    buffer[2] === 0x4e &&
    buffer[3] === 0x47 &&
    buffer[4] === 0x0d &&
    buffer[5] === 0x0a &&
    buffer[6] === 0x1a &&
    buffer[7] === 0x0a
  ) {
    return "image/png";
  }
  if (buffer.length >= 3 && buffer[0] === 0xff && buffer[1] === 0xd8 && buffer[2] === 0xff) return "image/jpeg";
  if (buffer.length >= 4 && buffer[0] === 0x47 && buffer[1] === 0x49 && buffer[2] === 0x46 && buffer[3] === 0x38) return "image/gif";
  if (buffer.length >= 12 && buffer[0] === 0x52 && buffer[1] === 0x49 && buffer[2] === 0x46 && buffer[3] === 0x46 && buffer[8] === 0x57 && buffer[9] === 0x45 && buffer[10] === 0x42 && buffer[11] === 0x50) {
    return "image/webp";
  }
  if (buffer.length >= 2 && buffer[0] === 0x42 && buffer[1] === 0x4d) return "image/bmp";
  if (buffer.length >= 12 && buffer[4] === 0x66 && buffer[5] === 0x74 && buffer[6] === 0x79 && buffer[7] === 0x70 && buffer[8] === 0x61 && buffer[9] === 0x76 && buffer[10] === 0x69 && buffer[11] === 0x66) {
    return "image/avif";
  }
  if (buffer.length >= 4 && buffer[0] === 0x00 && buffer[1] === 0x00 && buffer[2] === 0x01 && buffer[3] === 0x00) return "image/x-icon";

  const headerText = String(buffer.subarray(0, 512).toString("utf8") || "").trimStart().toLowerCase();
  if (headerText.startsWith("<svg") || (headerText.startsWith("<?xml") && headerText.includes("<svg"))) {
    return "image/svg+xml";
  }
  return "";
}

function looksLikeHtmlTextBuffer(buffer) {
  if (!Buffer.isBuffer(buffer) || !buffer.length) return false;
  const headerText = String(buffer.subarray(0, 512).toString("utf8") || "").replace(/\u0000/g, "").trimStart().toLowerCase();
  return (
    headerText.startsWith("<!doctype html") ||
    headerText.startsWith("<html") ||
    headerText.startsWith("<head") ||
    headerText.startsWith("<body")
  );
}

function getAiImageExtensionByUrl(rawUrl) {
  const safeUrl = sanitizeAiRemoteImageUrl(rawUrl);
  if (!safeUrl) return "";
  try {
    const parsed = new URL(safeUrl);
    const fileName = String(parsed.pathname || "").split("/").pop() || "";
    const matched = fileName.toLowerCase().match(/\.([a-z0-9]{2,5})$/);
    if (!matched) return "";
    const ext = matched[1];
    if (["png", "jpg", "jpeg", "webp", "gif", "bmp", "avif", "svg"].includes(ext)) {
      return ext === "jpeg" ? "jpg" : ext;
    }
    return "";
  } catch {
    return "";
  }
}

function buildAiImageCacheRelativeUrl(fileName) {
  return `/data/ai-image-cache/${encodeURIComponent(fileName)}`;
}

function createAiImageCacheFileName(extension = "png") {
  const safeExt = String(extension || "png")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "") || "png";
  return `ai-${Date.now()}-${Math.random().toString(36).slice(2, 10)}.${safeExt}`;
}

async function pruneAiImageCacheDirectory() {
  let entries = [];
  try {
    entries = await fsp.readdir(AI_IMAGE_CACHE_DIR, { withFileTypes: true });
  } catch {
    return;
  }
  const files = await Promise.all(
    entries
      .filter((entry) => entry.isFile() && /^ai-\d{13}-[a-z0-9]{8}\.[a-z0-9]{2,5}$/i.test(entry.name))
      .map(async (entry) => {
        const absPath = path.join(AI_IMAGE_CACHE_DIR, entry.name);
        try {
          const stat = await fsp.stat(absPath);
          return { name: entry.name, absPath, mtimeMs: Number(stat.mtimeMs || 0) };
        } catch {
          return null;
        }
      })
  );
  const safeFiles = files.filter(Boolean);
  if (safeFiles.length <= AI_IMAGE_CACHE_MAX_FILES) return;
  safeFiles.sort((a, b) => b.mtimeMs - a.mtimeMs);
  const staleFiles = safeFiles.slice(AI_IMAGE_CACHE_MAX_FILES);
  await Promise.all(
    staleFiles.map(async (item) => {
      try {
        await fsp.unlink(item.absPath);
      } catch {
        // ignore stale cache cleanup failure
      }
    })
  );
}

async function writeAiImageCacheFile({ buffer, mimeType = "image/png", sourceUrl = "" }) {
  if (!Buffer.isBuffer(buffer) || !buffer.length) {
    const error = new Error("图片缓存内容为空");
    error.code = "IMAGE_CACHE_EMPTY";
    error.status = 400;
    throw error;
  }
  if (buffer.length > AI_IMAGE_CACHE_MAX_BYTES) {
    const error = new Error(`图片过大（>${Math.floor(AI_IMAGE_CACHE_MAX_BYTES / 1024 / 1024)}MB）`);
    error.code = "IMAGE_CACHE_TOO_LARGE";
    error.status = 413;
    throw error;
  }
  const mimeExtension = getAiImageExtensionByMimeType(mimeType);
  const urlExtension = getAiImageExtensionByUrl(sourceUrl);
  const extension = mimeExtension || urlExtension || "png";
  const fileName = createAiImageCacheFileName(extension);
  await fsp.mkdir(AI_IMAGE_CACHE_DIR, { recursive: true });
  const absPath = path.join(AI_IMAGE_CACHE_DIR, fileName);
  await fsp.writeFile(absPath, buffer);
  await pruneAiImageCacheDirectory();
  return {
    fileName,
    assetUrl: buildAiImageCacheRelativeUrl(fileName),
    mimeType: mimeType || MIME_MAP[`.${extension}`] || "image/png",
    bytes: buffer.length,
  };
}

function sanitizeAiErrorMessage(rawValue, fallback = "请求失败") {
  const text = String(rawValue || "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 420);
  return text || fallback;
}

function writeSseEvent(res, eventName, payload) {
  if (!res || res.writableEnded || res.destroyed) return;
  const safeEvent = String(eventName || "message").replace(/[^a-zA-Z0-9_.-]/g, "") || "message";
  let serialized = "";
  try {
    serialized = typeof payload === "string" ? payload : JSON.stringify(payload ?? {});
  } catch {
    serialized = JSON.stringify({ message: "event_payload_serialize_failed" });
  }
  res.write(`event: ${safeEvent}\n`);
  serialized
    .split(/\r?\n/)
    .forEach((line) => res.write(`data: ${line}\n`));
  res.write("\n");
}

async function readStreamByChunk(readable, onChunk) {
  if (!readable || typeof readable.getReader !== "function") return;
  const reader = readable.getReader();
  const decoder = new TextDecoder("utf-8");
  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      const text = decoder.decode(value, { stream: true });
      if (text) onChunk(text);
    }
    const tail = decoder.decode();
    if (tail) onChunk(tail);
  } finally {
    try {
      reader.releaseLock();
    } catch {
      // ignore
    }
  }
}

function extractOllamaStreamDelta(chunkPayload) {
  if (!chunkPayload || typeof chunkPayload !== "object") return "";
  const content = String(chunkPayload?.message?.content || chunkPayload?.response || "");
  return content;
}

function extractOllamaStreamThinking(chunkPayload) {
  if (!chunkPayload || typeof chunkPayload !== "object") return "";
  return String(chunkPayload?.message?.thinking || chunkPayload?.thinking || "");
}

function extractOpenAiStreamDelta(chunkPayload) {
  if (!chunkPayload || typeof chunkPayload !== "object") return "";
  const firstChoice = Array.isArray(chunkPayload.choices) ? chunkPayload.choices[0] || null : null;
  if (!firstChoice) return "";
  const delta = firstChoice.delta || {};
  if (typeof delta.content === "string") return delta.content;
  if (Array.isArray(delta.content)) {
    return delta.content
      .map((entry) => {
        if (!entry || typeof entry !== "object") return "";
        if (typeof entry.text === "string") return entry.text;
        if (typeof entry.output_text === "string") return entry.output_text;
        return "";
      })
      .filter(Boolean)
      .join("");
  }
  const message = firstChoice.message || {};
  if (typeof message.content === "string") return message.content;
  return "";
}

async function streamAiChatRequest(rawPayload, req, res) {
  const provider = String(rawPayload?.provider || "").trim().toLowerCase() === "network" ? "network" : "ollama";
  const model = sanitizeAiModelName(rawPayload?.model);
  if (!model) {
    writeSseEvent(res, "error", { code: "MODEL_REQUIRED", message: "缺少模型名称", status: 400 });
    return;
  }

  res.writeHead(200, {
    "Content-Type": "text/event-stream; charset=utf-8",
    "Cache-Control": "no-store",
    Connection: "keep-alive",
    "X-Accel-Buffering": "no",
  });
  writeSseEvent(res, "status", {
    phase: "queued",
    provider,
    model,
    message: provider === "ollama" ? "已发送请求，正在等待本地模型响应" : "已发送请求，正在等待模型响应",
  });

  const timeoutMs = resolveAiChatTimeoutMs(provider, model);
  const controller = new AbortController();
  let abortedByTimeout = false;
  const timer = setTimeout(() => {
    abortedByTimeout = true;
    controller.abort();
  }, timeoutMs);
  const onClientClose = () => controller.abort();
  req.on("aborted", onClientClose);
  req.on("close", onClientClose);

  let outputText = "";
  let hasStreamOutput = false;
  const requestStartAt = Date.now();
  const heartbeatTimer = setInterval(() => {
    if (res.writableEnded || res.destroyed) return;
    const elapsedSeconds = Math.max(1, Math.floor((Date.now() - requestStartAt) / 1000));
    if (!hasStreamOutput) {
      writeSseEvent(res, "status", {
        phase: "waiting-first-token",
        provider,
        model,
        message:
          provider === "ollama"
            ? `本地模型思考中，已等待 ${elapsedSeconds} 秒（首轮加载可能较慢）`
            : `模型处理中，已等待 ${elapsedSeconds} 秒`,
      });
      return;
    }
    writeSseEvent(res, "status", {
      phase: "streaming-heartbeat",
      provider,
      model,
      message: `持续生成中，已输出 ${outputText.length} 字`,
    });
  }, 3200);
  try {
    let endpoint = "";
    let headers = {};
    let requestPayload = {};
    if (provider === "ollama") {
      const baseUrl = normalizeAiHttpUrl(rawPayload?.ollamaBaseUrl, AI_OLLAMA_DEFAULT_BASE_URL);
      endpoint = buildAiEndpointUrl(baseUrl, "/api/chat");
      headers = { "Content-Type": "application/json" };
      requestPayload = {
        model,
        messages: buildAiOllamaChatMessages({
          messages: rawPayload?.messages,
          prompt: rawPayload?.prompt,
          systemPrompt: rawPayload?.systemPrompt,
          imageDataUrl: rawPayload?.imageDataUrl,
        }),
        stream: true,
        keep_alive: AI_OLLAMA_KEEP_ALIVE,
      };
    } else {
      const network = rawPayload?.network && typeof rawPayload.network === "object" ? rawPayload.network : {};
      const baseParts = splitAiBaseUrlAndKnownEndpoint(network.baseUrl);
      const baseUrl = baseParts.baseUrl;
      if (!baseUrl) {
        writeSseEvent(res, "error", { code: "NETWORK_BASE_URL_INVALID", message: "网络模型 Base URL 无效", status: 400 });
        return;
      }
      const chatPath =
        baseParts.endpointType === "chat" && baseParts.endpointPath
          ? baseParts.endpointPath
          : normalizeAiHttpPath(network.chatPath, "/chat/completions");
      endpoint = buildAiEndpointUrl(baseUrl, chatPath);
      headers = buildAiRequestHeaders(network.apiKey);
      requestPayload = {
        model,
        messages: buildAiOpenAiChatMessages({
          messages: rawPayload?.messages,
          prompt: rawPayload?.prompt,
          systemPrompt: rawPayload?.systemPrompt,
          imageDataUrl: rawPayload?.imageDataUrl,
        }),
        stream: true,
      };
    }

    const upstreamResponse = await fetch(endpoint, {
      method: "POST",
      headers,
      body: JSON.stringify(requestPayload),
      signal: controller.signal,
      cache: "no-store",
    });

    if (!upstreamResponse.ok) {
      const rawText = await upstreamResponse.text();
      let parsedError = {};
      try {
        parsedError = rawText ? JSON.parse(rawText) : {};
      } catch {
        parsedError = {};
      }
      const message =
        sanitizeAiErrorMessage(parsedError?.error?.message, "") ||
        sanitizeAiErrorMessage(parsedError?.error, "") ||
        sanitizeAiErrorMessage(parsedError?.message, "") ||
        sanitizeAiErrorMessage(rawText, `上游请求失败（${upstreamResponse.status}）`);
      const error = new Error(message);
      error.status = upstreamResponse.status;
      throw error;
    }

    writeSseEvent(res, "status", {
      phase: "streaming",
      provider,
      model,
      message: provider === "ollama" ? "模型已开始生成，正在实时输出" : "已开始接收模型流式输出",
    });

    const responseType = String(upstreamResponse.headers.get("content-type") || "").toLowerCase();
    if (provider === "ollama" || responseType.includes("application/x-ndjson") || responseType.includes("application/jsonl")) {
      let pending = "";
      await readStreamByChunk(upstreamResponse.body, (chunk) => {
        pending += String(chunk || "");
        const lines = pending.split(/\r?\n/);
        pending = lines.pop() || "";
        lines.forEach((rawLine) => {
          const line = String(rawLine || "").trim();
          if (!line) return;
          let payload = null;
          try {
            payload = JSON.parse(line);
          } catch {
            payload = null;
          }
          if (!payload) return;
          const thinking = extractOllamaStreamThinking(payload);
          if (thinking) {
            hasStreamOutput = true;
            writeSseEvent(res, "thinking", { thinking });
          }
          const delta = extractOllamaStreamDelta(payload);
          if (delta) {
            hasStreamOutput = true;
            outputText += delta;
            writeSseEvent(res, "delta", { delta });
          }
        });
      });
      const tailLine = String(pending || "").trim();
      if (tailLine) {
        try {
          const tailPayload = JSON.parse(tailLine);
          const thinking = extractOllamaStreamThinking(tailPayload);
          if (thinking) {
            hasStreamOutput = true;
            writeSseEvent(res, "thinking", { thinking });
          }
          const delta = extractOllamaStreamDelta(tailPayload);
          if (delta) {
            hasStreamOutput = true;
            outputText += delta;
            writeSseEvent(res, "delta", { delta });
          }
        } catch {
          // ignore tail parse failure
        }
      }
    } else if (responseType.includes("text/event-stream")) {
      let sseBuffer = "";
      let streamDone = false;
      const processSseBlock = (rawBlock) => {
        if (streamDone) return;
        const block = String(rawBlock || "").replace(/\r/g, "").trim();
        if (!block) return;
        const lines = block.split("\n");
        const dataLines = [];
        lines.forEach((line) => {
          if (line.startsWith("data:")) dataLines.push(line.slice(5).trimStart());
        });
        if (!dataLines.length) return;
        const dataText = dataLines.join("\n");
        if (dataText === "[DONE]") {
          streamDone = true;
          return;
        }
        let payload = null;
        try {
          payload = JSON.parse(dataText);
        } catch {
          payload = null;
        }
        if (!payload) return;
        const delta = extractOpenAiStreamDelta(payload);
        if (delta) {
          hasStreamOutput = true;
          outputText += delta;
          writeSseEvent(res, "delta", { delta });
        }
      };

      await readStreamByChunk(upstreamResponse.body, (chunk) => {
        if (streamDone) return;
        sseBuffer += String(chunk || "").replace(/\r/g, "");
        while (true) {
          const markerIndex = sseBuffer.indexOf("\n\n");
          if (markerIndex < 0) break;
          const block = sseBuffer.slice(0, markerIndex);
          sseBuffer = sseBuffer.slice(markerIndex + 2);
          processSseBlock(block);
          if (streamDone) break;
        }
      });
      if (!streamDone && sseBuffer.trim()) {
        processSseBlock(sseBuffer);
      }
    } else {
      // JSON fallback: emit as single chunk.
      const rawText = await upstreamResponse.text();
      let payload = null;
      try {
        payload = rawText ? JSON.parse(rawText) : {};
      } catch {
        payload = { rawText };
      }
      const fallbackText =
        provider === "ollama"
          ? String(payload?.message?.content || payload?.response || payload?.text || "").slice(0, 18000)
          : String(extractTextFromOpenAiResponse(payload) || payload?.text || "").slice(0, 18000);
      if (fallbackText) {
        hasStreamOutput = true;
        outputText += fallbackText;
        writeSseEvent(res, "delta", { delta: fallbackText });
      }
    }

    writeSseEvent(res, "done", { provider, model, text: outputText || "模型未返回文本内容" });
  } catch (error) {
    if (!res.writableEnded && !res.destroyed) {
      const message =
        error?.name === "AbortError"
          ? abortedByTimeout
            ? `请求超时（${Math.ceil(timeoutMs / 1000)}秒），模型可能仍在加载，请稍后重试`
            : "请求已取消"
          : sanitizeAiErrorMessage(error?.message, "流式请求失败");
      const statusCode =
        error?.name === "AbortError" ? (abortedByTimeout ? 504 : 408) : Number.isFinite(Number(error?.status)) ? Number(error.status) : 502;
      writeSseEvent(res, "error", {
        status: statusCode,
        message,
        code: error?.name === "AbortError" ? (abortedByTimeout ? "REQUEST_TIMEOUT" : "REQUEST_ABORTED") : "STREAM_FAILED",
      });
    }
  } finally {
    clearInterval(heartbeatTimer);
    clearTimeout(timer);
    req.off("aborted", onClientClose);
    req.off("close", onClientClose);
    if (!res.writableEnded && !res.destroyed) {
      res.end();
    }
  }
}

function decodeHtmlEntities(rawText) {
  const text = String(rawText || "");
  if (!text) return "";
  return text
    .replace(/<!\[CDATA\[([\s\S]*?)\]\]>/g, "$1")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hexCode) => {
      const code = Number.parseInt(hexCode, 16);
      if (!Number.isFinite(code)) return "";
      return String.fromCodePoint(code);
    })
    .replace(/&#(\d+);/g, (_, decCode) => {
      const code = Number.parseInt(decCode, 10);
      if (!Number.isFinite(code)) return "";
      return String.fromCodePoint(code);
    })
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
}

function sanitizeAiWebSearchText(rawValue, maxLength = 300) {
  const max = Number.isFinite(Number(maxLength)) ? Math.max(20, Math.floor(Number(maxLength))) : 300;
  return decodeHtmlEntities(rawValue)
    .replace(/<[^>]*>/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, max);
}

function sanitizeAiWebSearchQuery(rawValue) {
  return sanitizeAiWebSearchText(rawValue, 260);
}

function normalizeAiWebSearchTaskQuery(rawQuery) {
  const safeQuery = sanitizeAiWebSearchQuery(rawQuery);
  if (!safeQuery) return "";
  const quotedTerms = extractAiWebSearchQuotedTerms(safeQuery).filter((term) => term.length >= 2);
  if (quotedTerms.length) {
    return sanitizeAiWebSearchQuery(quotedTerms[0]).slice(0, 180);
  }
  let normalized = safeQuery;
  normalized = normalized.replace(
    /^[请麻烦帮我你可否是否可以]*\s*(检索|搜索|查找|查询|上网查|联网查|联网搜索|联网检索)\s*/i,
    ""
  );
  normalized = normalized.replace(
    /[，,。；;:：]\s*(给出|提供|并给出|请给出|附|附上|包含).{0,50}(来源|链接|出处|参考).*/i,
    ""
  );
  normalized = normalized.replace(/\s*(给出|提供|并给出|请给出).{0,30}(来源|链接|出处|参考).*$/i, "");
  normalized = normalized.replace(/[“”"'`]/g, " ").replace(/\s+/g, " ").trim();
  if (!normalized) return safeQuery.length > 180 ? `${safeQuery.slice(0, 180)}...` : safeQuery;
  return normalized.length > 180 ? `${normalized.slice(0, 180)}...` : normalized;
}

const AI_WEATHER_QUERY_KEYWORDS = [
  "天气",
  "天气预报",
  "气温",
  "温度",
  "降雨",
  "下雨",
  "风力",
  "风速",
  "湿度",
  "weather",
  "forecast",
];

function isLikelyAiWeatherQuery(rawQuery) {
  const text = sanitizeAiWebSearchQuery(rawQuery).toLowerCase();
  if (!text) return false;
  return AI_WEATHER_QUERY_KEYWORDS.some((keyword) => text.includes(String(keyword).toLowerCase()));
}

function extractAiWeatherLocationFromQuery(rawQuery) {
  const text = sanitizeAiWebSearchQuery(rawQuery);
  if (!text) return "";
  const weatherAnchor = /(天气预报|天气|气温|温度|降雨|下雨|风力|风速|湿度|weather|forecast)/i;
  let candidate = text;
  const weatherMatched = weatherAnchor.exec(text);
  if (weatherMatched && Number.isInteger(weatherMatched.index) && weatherMatched.index > 0) {
    candidate = text.slice(0, weatherMatched.index);
  }
  candidate = candidate
    .replace(/[，。！？、,.!?]/g, " ")
    .replace(/\b(today|tomorrow|yesterday|now|current)\b/gi, " ")
    .replace(/今天|今日|明天|明日|昨天|昨日|现在|当前|本周|这周|本月|今年/g, " ")
    .replace(/请你|请问|帮我|看看|看下|查一下|查下|查询|告诉我|一下|如何|怎么样|什么样|情况|历史上的/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!candidate) return "";
  const tokens = candidate.match(/[A-Za-z\u4e00-\u9fff]{2,24}/g) || [];
  if (!tokens.length) return "";
  const ignored = new Set(["天气", "气温", "温度", "预报", "weather", "forecast"]);
  for (let i = tokens.length - 1; i >= 0; i -= 1) {
    const token = sanitizeAiWebSearchText(tokens[i], 24);
    if (!token) continue;
    let normalized = token
      .replace(/^(中国|我国)/, "")
      .replace(/(的|天气|气温|温度|预报)+$/g, "")
      .replace(/^(在|到|去)+/g, "")
      .trim();
    if (!normalized) continue;
    if (ignored.has(normalized.toLowerCase())) continue;
    return sanitizeAiWebSearchText(normalized, 24);
  }
  return "";
}

function buildAiWeatherGeocodeQueryVariants(locationName) {
  const base = sanitizeAiWebSearchText(locationName, 24);
  if (!base) return [];
  const output = [];
  const seen = new Set();
  const add = (rawValue) => {
    const safe = sanitizeAiWebSearchText(rawValue, 24);
    if (!safe) return;
    const key = safe.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    output.push(safe);
  };
  const isMostlyChinese = /^[\u4e00-\u9fff]{2,12}$/.test(base);
  if (isMostlyChinese && !base.endsWith("市")) add(`${base}市`);
  add(base);
  return output.slice(0, 4);
}

function mapOpenMeteoWeatherCodeToText(codeValue) {
  const code = Number(codeValue);
  if (!Number.isFinite(code)) return "未知天气";
  const map = {
    0: "晴朗",
    1: "大部晴朗",
    2: "局部多云",
    3: "阴天",
    45: "有雾",
    48: "冻雾",
    51: "小毛毛雨",
    53: "毛毛雨",
    55: "强毛毛雨",
    56: "冻毛毛雨",
    57: "强冻毛毛雨",
    61: "小雨",
    63: "中雨",
    65: "大雨",
    66: "冻雨",
    67: "强冻雨",
    71: "小雪",
    73: "中雪",
    75: "大雪",
    77: "雪粒",
    80: "阵雨",
    81: "较强阵雨",
    82: "强阵雨",
    85: "阵雪",
    86: "强阵雪",
    95: "雷雨",
    96: "雷雨夹小冰雹",
    99: "雷雨夹冰雹",
  };
  return map[Math.round(code)] || `天气代码 ${Math.round(code)}`;
}

function formatOpenMeteoNumber(value, fractionDigits = 1) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) return "";
  return Number(numeric.toFixed(Math.max(0, Math.min(3, Math.floor(fractionDigits)))));
}

function buildOpenMeteoLocationLabel(location) {
  if (!location || typeof location !== "object") return "";
  const parts = [
    sanitizeAiWebSearchText(location.name, 80),
    sanitizeAiWebSearchText(location.admin1, 80),
    sanitizeAiWebSearchText(location.country, 80),
  ].filter(Boolean);
  return dedupeAiWebSearchStringList(parts, 3).join(" / ");
}

function pickBestOpenMeteoLocation(candidates, locationKeyword = "") {
  const list = Array.isArray(candidates) ? candidates : [];
  if (!list.length) return null;
  const keyword = sanitizeAiWebSearchText(locationKeyword, 30).toLowerCase();
  const ranked = list
    .map((entry) => {
      if (!entry || typeof entry !== "object") return null;
      if (!Number.isFinite(Number(entry.latitude)) || !Number.isFinite(Number(entry.longitude))) return null;
      const name = sanitizeAiWebSearchText(entry.name, 80).toLowerCase();
      const admin1 = sanitizeAiWebSearchText(entry.admin1, 80).toLowerCase();
      const country = sanitizeAiWebSearchText(entry.country, 80).toLowerCase();
      const featureCode = sanitizeAiWebSearchText(entry.feature_code || entry.featureCode, 30).toUpperCase();
      const population = Number(entry.population);
      let score = 0;
      if (keyword) {
        if (name === keyword) score += 120;
        else if (name.includes(keyword)) score += 60;
        if (admin1 === keyword) score += 90;
        else if (admin1.includes(keyword)) score += 35;
      }
      if (country.includes("中国") || String(entry.country_code || "").toUpperCase() === "CN") score += 20;
      if (featureCode === "PPLC") score += 50;
      else if (featureCode.startsWith("PPL")) score += 25;
      if (Number.isFinite(population) && population > 0) {
        score += Math.min(30, Math.log10(population) * 4);
      }
      return { entry, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);
  return ranked.length ? ranked[0].entry : null;
}

function pickOpenMeteoDailyIndex(payload, targetDateKey) {
  const daily = payload?.daily && typeof payload.daily === "object" ? payload.daily : null;
  const times = Array.isArray(daily?.time) ? daily.time : [];
  if (!times.length) return -1;
  const safeTarget = sanitizeAiWebSearchText(targetDateKey, 20);
  if (safeTarget) {
    const exactIndex = times.findIndex((value) => String(value || "").trim() === safeTarget);
    if (exactIndex >= 0) return exactIndex;
  }
  return 0;
}

async function fetchOpenMeteoWeatherSearchResults(query, maxResults = 5, dateIntent = {}) {
  if (!isLikelyAiWeatherQuery(query)) return [];
  const locationName = extractAiWeatherLocationFromQuery(query);
  if (!locationName) return [];
  const geocodeVariants = buildAiWeatherGeocodeQueryVariants(locationName);
  let geocodeCandidates = [];
  let geocodeFailed = false;
  for (const variant of geocodeVariants) {
    const geocodeUrl = `https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(
      variant
    )}&count=8&language=zh&format=json`;
    try {
      const geocodeResult = await fetchAiJson(geocodeUrl, {
        method: "GET",
        headers: { "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)" },
        timeoutMs: AI_WEATHER_REQUEST_TIMEOUT_MS,
      });
      if (!geocodeResult.ok) {
        geocodeFailed = true;
        continue;
      }
      const geocodeData =
        geocodeResult.data && typeof geocodeResult.data === "object" && !("rawText" in geocodeResult.data) ? geocodeResult.data : {};
      const results = Array.isArray(geocodeData.results) ? geocodeData.results : [];
      if (results.length) {
        geocodeCandidates = geocodeCandidates.concat(results);
      }
      if (geocodeCandidates.length >= 8) break;
    } catch {
      geocodeFailed = true;
    }
  }
  if (!geocodeCandidates.length && geocodeFailed) {
    const error = new Error("天气地理编码失败");
    error.code = "WEATHER_GEOCODE_FAILED";
    error.status = 502;
    throw error;
  }
  const selected = pickBestOpenMeteoLocation(geocodeCandidates, locationName);
  if (!selected) return [];
  const latitude = Number(selected.latitude);
  const longitude = Number(selected.longitude);
  const timezone = sanitizeAiWebSearchText(selected.timezone, 80) || "Asia/Shanghai";
  const nowDateIso = formatAiShanghaiDateTimeParts(new Date()).dateIso;
  const targetDateKey = sanitizeAiWebSearchText(dateIntent?.startDateKey, 20) || nowDateIso;
  const forecastUrl = new URL("https://api.open-meteo.com/v1/forecast");
  forecastUrl.searchParams.set("latitude", String(latitude));
  forecastUrl.searchParams.set("longitude", String(longitude));
  forecastUrl.searchParams.set("timezone", timezone);
  forecastUrl.searchParams.set(
    "current",
    "temperature_2m,relative_humidity_2m,precipitation,weather_code,wind_speed_10m"
  );
  forecastUrl.searchParams.set(
    "daily",
    "weather_code,temperature_2m_max,temperature_2m_min,precipitation_probability_max,precipitation_sum"
  );
  forecastUrl.searchParams.set("forecast_days", "10");
  const useExplicitDateRange = sanitizeAiWebSearchTerm(dateIntent?.relativeKeyword || "") === "explicit-date";
  if (useExplicitDateRange && /^\d{4}-\d{2}-\d{2}$/.test(targetDateKey)) {
    forecastUrl.searchParams.set("start_date", targetDateKey);
    forecastUrl.searchParams.set("end_date", targetDateKey);
  }
  const forecastResult = await fetchAiJson(forecastUrl.toString(), {
    method: "GET",
    headers: { "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)" },
    timeoutMs: AI_WEATHER_REQUEST_TIMEOUT_MS,
  });
  if (!forecastResult.ok) {
    const error = new Error(`天气预报获取失败（${forecastResult.status}）`);
    error.code = "WEATHER_FORECAST_FAILED";
    error.status = forecastResult.status;
    throw error;
  }
  const forecastData =
    forecastResult.data && typeof forecastResult.data === "object" && !("rawText" in forecastResult.data) ? forecastResult.data : {};
  const daily = forecastData.daily && typeof forecastData.daily === "object" ? forecastData.daily : {};
  const dailyIndex = pickOpenMeteoDailyIndex(forecastData, targetDateKey);
  const getDailyValue = (key) => {
    const list = Array.isArray(daily?.[key]) ? daily[key] : [];
    if (dailyIndex < 0 || dailyIndex >= list.length) return "";
    return list[dailyIndex];
  };
  const targetDay = Array.isArray(daily?.time) && dailyIndex >= 0 ? String(daily.time[dailyIndex] || "").trim() : targetDateKey;
  const dayWeatherText = mapOpenMeteoWeatherCodeToText(getDailyValue("weather_code"));
  const maxTemp = formatOpenMeteoNumber(getDailyValue("temperature_2m_max"), 1);
  const minTemp = formatOpenMeteoNumber(getDailyValue("temperature_2m_min"), 1);
  const rainProb = formatOpenMeteoNumber(getDailyValue("precipitation_probability_max"), 0);
  const rainSum = formatOpenMeteoNumber(getDailyValue("precipitation_sum"), 1);
  const current = forecastData.current && typeof forecastData.current === "object" ? forecastData.current : {};
  const currentTemp = formatOpenMeteoNumber(current.temperature_2m, 1);
  const currentHumidity = formatOpenMeteoNumber(current.relative_humidity_2m, 0);
  const currentWind = formatOpenMeteoNumber(current.wind_speed_10m, 1);
  const currentPrecip = formatOpenMeteoNumber(current.precipitation, 1);
  const currentCodeText = mapOpenMeteoWeatherCodeToText(current.weather_code);
  const locationLabel = buildOpenMeteoLocationLabel(selected) || locationName;
  const snippetLines = [
    `城市：${locationLabel}`,
    `目标日期：${targetDay || targetDateKey || nowDateIso}`,
    `当日天气：${dayWeatherText}${maxTemp !== "" && minTemp !== "" ? `，${minTemp}~${maxTemp}°C` : ""}${rainProb !== "" ? `，降水概率 ${rainProb}%` : ""}${rainSum !== "" ? `，降水量 ${rainSum}mm` : ""}`,
    `当前实况：${currentCodeText}${currentTemp !== "" ? `，温度 ${currentTemp}°C` : ""}${currentHumidity !== "" ? `，湿度 ${currentHumidity}%` : ""}${currentWind !== "" ? `，风速 ${currentWind}km/h` : ""}${currentPrecip !== "" ? `，降水 ${currentPrecip}mm` : ""}`,
    `数据时间基准：${nowDateIso}（Asia/Shanghai）`,
    "数据源：Open-Meteo（免费）",
  ].filter(Boolean);
  const item = sanitizeAiWebSearchItem({
    title: `${locationLabel}天气（Open-Meteo）`,
    url: "https://open-meteo.com/en/docs",
    snippet: snippetLines.join("；"),
  });
  if (!item) return [];
  return [item].slice(0, Math.max(1, Math.min(8, Number(maxResults) || 5)));
}

const AI_MAP_QUERY_KEYWORDS = [
  "地图",
  "导航",
  "路线",
  "怎么走",
  "路程",
  "距离",
  "经纬度",
  "坐标",
  "附近",
  "周边",
  "route",
  "directions",
  "map",
  "distance",
];

const AI_CN_LOCATION_PREFIXES = [
  "北京市",
  "北京",
  "上海市",
  "上海",
  "天津市",
  "天津",
  "重庆市",
  "重庆",
  "香港",
  "澳门",
  "河北",
  "山西",
  "辽宁",
  "吉林",
  "黑龙江",
  "江苏",
  "浙江",
  "安徽",
  "福建",
  "江西",
  "山东",
  "河南",
  "湖北",
  "湖南",
  "广东",
  "广西",
  "海南",
  "四川",
  "贵州",
  "云南",
  "西藏",
  "陕西",
  "甘肃",
  "青海",
  "宁夏",
  "新疆",
  "内蒙古",
];

function extractCnLocationPrefix(locationKeyword = "") {
  const keyword = sanitizeAiWebSearchText(locationKeyword, 60).replace(/\s+/g, "");
  if (!keyword) return "";
  const matched = AI_CN_LOCATION_PREFIXES
    .slice()
    .sort((a, b) => b.length - a.length)
    .find((prefix) => keyword.startsWith(prefix));
  return matched ? sanitizeAiWebSearchText(matched, 16) : "";
}

function isLikelyAiMapQuery(rawQuery) {
  const text = sanitizeAiWebSearchQuery(rawQuery).toLowerCase();
  if (!text) return false;
  if (AI_MAP_QUERY_KEYWORDS.some((keyword) => text.includes(String(keyword).toLowerCase()))) return true;
  if (/从.{1,24}(到|去|前往).{1,24}/.test(text) && /(怎么走|路线|导航|路程|距离|多久|多远|route|direction|distance)/i.test(text)) {
    return true;
  }
  return false;
}

function normalizeAiMapLocationName(rawValue) {
  return sanitizeAiWebSearchText(rawValue, 60)
    .replace(/[，。！？、,.!?]/g, " ")
    .replace(/\b(route|directions|map|distance|to|from)\b/gi, " ")
    .replace(/(怎么走|路线|导航|路程|距离|经纬度|坐标|附近|周边|在哪里|在哪儿|在哪|地图)+/g, " ")
    .replace(/^(从|在|去|到|前往)+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractAiMapRouteEndpointsFromQuery(rawQuery) {
  const text = sanitizeAiWebSearchQuery(rawQuery);
  if (!text) return null;
  const patterns = [
    /从\s*([A-Za-z0-9\u4e00-\u9fff·（）()\-]{1,36}?)\s*(?:到|去|前往)\s*([A-Za-z0-9\u4e00-\u9fff·（）()\-]{1,36})(?:\s|$|[，。！？,.!?])/i,
    /([A-Za-z0-9\u4e00-\u9fff·（）()\-]{2,36})\s*(?:到|去|前往)\s*([A-Za-z0-9\u4e00-\u9fff·（）()\-]{2,36})\s*(?:怎么走|路线|导航|路程|距离|多久|多远|route|directions|distance)?/i,
  ];
  for (const pattern of patterns) {
    const matched = text.match(pattern);
    if (!matched) continue;
    const origin = normalizeAiMapLocationName(matched[1]);
    const destination = normalizeAiMapLocationName(matched[2]);
    if (!origin || !destination) continue;
    if (origin.toLowerCase() === destination.toLowerCase()) continue;
    return {
      origin: sanitizeAiWebSearchText(origin, 36),
      destination: sanitizeAiWebSearchText(destination, 36),
    };
  }
  return null;
}

function extractAiMapLocationFromQuery(rawQuery) {
  const text = sanitizeAiWebSearchQuery(rawQuery);
  if (!text) return "";
  const routeIntent = extractAiMapRouteEndpointsFromQuery(text);
  if (routeIntent) return "";
  let candidate = text
    .replace(/[，。！？、,.!?]/g, " ")
    .replace(/\b(route|directions|map|distance|to|from)\b/gi, " ")
    .replace(/请你|请问|帮我|看看|看下|查一下|查下|查询|告诉我|一下|如何|怎么样|什么样|情况/g, " ")
    .replace(/(地图|导航|路线|怎么走|路程|距离|经纬度|坐标|附近|周边|在哪里|在哪儿|在哪)+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (!candidate) return "";
  const tokens = candidate.match(/[A-Za-z0-9\u4e00-\u9fff·（）()\-]{2,36}/g) || [];
  if (!tokens.length) return "";
  for (let i = tokens.length - 1; i >= 0; i -= 1) {
    const normalized = normalizeAiMapLocationName(tokens[i]);
    if (!normalized) continue;
    return sanitizeAiWebSearchText(normalized, 36);
  }
  return "";
}

function buildOsmLocationDisplayLabel(location) {
  if (!location || typeof location !== "object") return "";
  const displayName = sanitizeAiWebSearchText(location.displayName || location.display_name, 120);
  if (displayName) return displayName;
  const parts = [
    sanitizeAiWebSearchText(location.name, 80),
    sanitizeAiWebSearchText(location.city, 80),
    sanitizeAiWebSearchText(location.state, 80),
    sanitizeAiWebSearchText(location.country, 80),
  ].filter(Boolean);
  return dedupeAiWebSearchStringList(parts, 4).join(" / ");
}

async function fetchOsmNominatimCandidates(locationKeyword, maxCount = 6) {
  const keyword = sanitizeAiWebSearchText(locationKeyword, 60);
  if (!keyword) return [];
  const limit = Number.isFinite(Number(maxCount)) ? Math.max(1, Math.min(12, Math.floor(Number(maxCount)))) : 6;
  const endpoint = new URL("https://nominatim.openstreetmap.org/search");
  endpoint.searchParams.set("q", keyword);
  endpoint.searchParams.set("format", "jsonv2");
  endpoint.searchParams.set("addressdetails", "1");
  endpoint.searchParams.set("limit", String(limit));
  endpoint.searchParams.set("accept-language", "zh-CN,zh;q=0.9,en;q=0.7");
  const result = await fetchAiJson(endpoint.toString(), {
    method: "GET",
    headers: {
      "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)",
      Accept: "application/json,text/plain,*/*",
    },
    timeoutMs: AI_MAP_REQUEST_TIMEOUT_MS,
  });
  if (!result.ok) {
    const error = new Error(`OpenStreetMap 地理编码失败（${result.status}）`);
    error.code = "MAP_GEOCODE_FAILED";
    error.status = result.status;
    throw error;
  }
  const payload = Array.isArray(result.data) ? result.data : [];
  return payload
    .map((entry) => {
      const lat = Number(entry?.lat);
      const lon = Number(entry?.lon);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
      const displayName = sanitizeAiWebSearchText(entry?.display_name, 160);
      const type = sanitizeAiWebSearchText(entry?.type, 40);
      const className = sanitizeAiWebSearchText(entry?.class, 40);
      const importance = Number(entry?.importance);
      const address = entry?.address && typeof entry.address === "object" ? entry.address : {};
      return {
        latitude: lat,
        longitude: lon,
        displayName,
        type,
        className,
        importance: Number.isFinite(importance) ? importance : 0,
        city: sanitizeAiWebSearchText(address.city || address.town || address.village || "", 80),
        state: sanitizeAiWebSearchText(address.state || address.province || "", 80),
        country: sanitizeAiWebSearchText(address.country || "", 80),
      };
    })
    .filter(Boolean)
    .slice(0, limit);
}

async function fetchOpenMeteoMapGeocodeCandidates(locationKeyword, maxCount = 6) {
  const keyword = sanitizeAiWebSearchText(locationKeyword, 60);
  if (!keyword) return [];
  const limit = Number.isFinite(Number(maxCount)) ? Math.max(1, Math.min(12, Math.floor(Number(maxCount)))) : 6;
  const endpoint = `https://geocoding-api.open-meteo.com/v1/search?name=${encodeURIComponent(
    keyword
  )}&count=${limit}&language=zh&format=json`;
  const result = await fetchAiJson(endpoint, {
    method: "GET",
    headers: {
      "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)",
      Accept: "application/json,text/plain,*/*",
    },
    timeoutMs: AI_MAP_REQUEST_TIMEOUT_MS,
  });
  if (!result.ok) {
    const error = new Error(`Open-Meteo 地理编码失败（${result.status}）`);
    error.code = "MAP_GEOCODE_FAILED";
    error.status = result.status;
    throw error;
  }
  const payload = result.data && typeof result.data === "object" && !("rawText" in result.data) ? result.data : {};
  const list = Array.isArray(payload.results) ? payload.results : [];
  return list
    .map((entry) => {
      const lat = Number(entry?.latitude);
      const lon = Number(entry?.longitude);
      if (!Number.isFinite(lat) || !Number.isFinite(lon)) return null;
      const name = sanitizeAiWebSearchText(entry?.name, 80);
      const admin1 = sanitizeAiWebSearchText(entry?.admin1, 80);
      const country = sanitizeAiWebSearchText(entry?.country, 80);
      const displayName = dedupeAiWebSearchStringList([name, admin1, country], 3).join(" / ");
      const population = Number(entry?.population);
      const importance = Number.isFinite(population) && population > 0 ? Math.min(1, Math.log10(population) / 8) : 0;
      return {
        latitude: lat,
        longitude: lon,
        displayName,
        type: sanitizeAiWebSearchText(entry?.feature_code || "location", 40).toLowerCase(),
        className: "geocoding-open-meteo",
        importance,
        city: name,
        state: admin1,
        country,
      };
    })
    .filter(Boolean)
    .slice(0, limit);
}

function buildAiMapGeocodeQueryVariants(locationKeyword) {
  const keyword = sanitizeAiWebSearchText(locationKeyword, 60);
  if (!keyword) return [];
  const output = [];
  const seen = new Set();
  const add = (rawValue) => {
    const safe = sanitizeAiWebSearchText(rawValue, 60).replace(/\s+/g, " ").trim();
    if (!safe) return;
    const key = safe.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    output.push(safe);
  };
  add(keyword);
  const compact = keyword.replace(/\s+/g, "");
  if (compact && compact !== keyword) add(compact);
  const prefix = extractCnLocationPrefix(compact);
  if (prefix && compact.startsWith(prefix)) {
    const suffix = sanitizeAiWebSearchText(compact.slice(prefix.length), 40);
    if (suffix.length >= 2) {
      add(`${prefix} ${suffix}`);
      add(`${suffix} ${prefix}`);
    }
  }
  const splitTokens = keyword.split(/\s+/g).filter(Boolean);
  if (splitTokens.length >= 2) {
    add(splitTokens[splitTokens.length - 1]);
    add(`${splitTokens[splitTokens.length - 1]} ${splitTokens[0]}`);
  }
  return output.slice(0, 5);
}

async function fetchMapGeocodeCandidates(locationKeyword, maxCount = 6) {
  const variants = buildAiMapGeocodeQueryVariants(locationKeyword);
  if (!variants.length) return [];
  let nominatimFailed = false;
  let openMeteoFailed = false;
  const merged = [];
  for (const variant of variants.slice(0, 3)) {
    try {
      const nominatimResults = await fetchOsmNominatimCandidates(variant, maxCount);
      if (nominatimResults.length) merged.push(...nominatimResults);
    } catch {
      nominatimFailed = true;
    }
    try {
      const openMeteoResults = await fetchOpenMeteoMapGeocodeCandidates(variant, maxCount);
      if (openMeteoResults.length) merged.push(...openMeteoResults);
    } catch {
      openMeteoFailed = true;
    }
    if (merged.length >= maxCount * 2) break;
  }
  const deduped = [];
  const seen = new Set();
  for (const entry of merged) {
    if (!entry || typeof entry !== "object") continue;
    const lat = Number(entry.latitude);
    const lon = Number(entry.longitude);
    if (!Number.isFinite(lat) || !Number.isFinite(lon)) continue;
    const key = `${lat.toFixed(4)},${lon.toFixed(4)}`;
    if (seen.has(key)) continue;
    seen.add(key);
    deduped.push(entry);
    if (deduped.length >= Math.max(6, maxCount * 3)) break;
  }
  if (deduped.length) return deduped;
  if (nominatimFailed || openMeteoFailed) {
    const error = new Error("地图地理编码失败");
    error.code = "MAP_GEOCODE_FAILED";
    error.status = 502;
    throw error;
  }
  return [];
}

function pickBestOsmNominatimLocation(candidates, locationKeyword = "") {
  const list = Array.isArray(candidates) ? candidates : [];
  if (!list.length) return null;
  const keyword = sanitizeAiWebSearchText(locationKeyword, 40).toLowerCase();
  const keywordTokens = dedupeAiWebSearchStringList(
    [
      ...keyword.split(/\s+/g),
      ...(keyword.match(/[\u4e00-\u9fff]{2,}/g) || []),
    ],
    8
  )
    .map((token) => sanitizeAiWebSearchText(token, 20).toLowerCase())
    .filter((token) => token.length >= 2);
  const ranked = list
    .map((entry) => {
      if (!entry || typeof entry !== "object") return null;
      if (!Number.isFinite(Number(entry.latitude)) || !Number.isFinite(Number(entry.longitude))) return null;
      const label = buildOsmLocationDisplayLabel(entry).toLowerCase();
      const cityText = sanitizeAiWebSearchText(entry.city, 80).toLowerCase();
      const stateText = sanitizeAiWebSearchText(entry.state, 80).toLowerCase();
      const countryText = sanitizeAiWebSearchText(entry.country, 80).toLowerCase();
      const type = sanitizeAiWebSearchText(entry.type, 40).toLowerCase();
      const className = sanitizeAiWebSearchText(entry.className, 40).toLowerCase();
      let score = 0;
      if (keyword) {
        if (label === keyword) score += 140;
        else if (label.includes(keyword)) score += 70;
      }
      for (const token of keywordTokens) {
        if (!token) continue;
        if (label.includes(token)) score += 18;
        if (cityText.includes(token)) score += 35;
        if (stateText.includes(token)) score += 28;
        if (countryText.includes(token)) score += 8;
      }
      if (keyword.includes("上海")) {
        if (cityText.includes("上海") || stateText.includes("上海") || label.includes("上海")) score += 120;
        else score -= 45;
      }
      if (keyword.includes("北京")) {
        if (cityText.includes("北京") || stateText.includes("北京") || label.includes("北京")) score += 120;
        else score -= 45;
      }
      if (keyword.includes("深圳")) {
        if (cityText.includes("深圳") || stateText.includes("深圳") || label.includes("深圳")) score += 120;
        else score -= 45;
      }
      if (keyword.includes("广州")) {
        if (cityText.includes("广州") || stateText.includes("广州") || label.includes("广州")) score += 120;
        else score -= 45;
      }
      if (type === "city" || type === "administrative" || type === "town") score += 18;
      if (className === "boundary" || className === "place") score += 10;
      if (label.includes("中国")) score += 8;
      if (Number.isFinite(Number(entry.importance))) score += Math.max(0, Math.min(30, Number(entry.importance) * 30));
      return { entry, score };
    })
    .filter(Boolean)
    .sort((a, b) => b.score - a.score);
  return ranked.length ? ranked[0].entry : null;
}

function formatAiMapDistance(distanceMeters) {
  const meters = Number(distanceMeters);
  if (!Number.isFinite(meters) || meters < 0) return "";
  if (meters < 1000) return `${Math.round(meters)}米`;
  const km = meters / 1000;
  const fixed = km >= 100 ? km.toFixed(0) : km.toFixed(1);
  return `${fixed.replace(/\.0$/, "")}公里`;
}

function formatAiMapDuration(durationSeconds) {
  const seconds = Number(durationSeconds);
  if (!Number.isFinite(seconds) || seconds < 0) return "";
  if (seconds < 60) return `${Math.max(1, Math.round(seconds))}秒`;
  const minutes = Math.round(seconds / 60);
  if (minutes < 60) return `${minutes}分钟`;
  const hours = Math.floor(minutes / 60);
  const remainMinutes = minutes % 60;
  if (!remainMinutes) return `${hours}小时`;
  return `${hours}小时${remainMinutes}分钟`;
}

function buildOpenStreetMapPlaceUrl(location, zoom = 12) {
  const latitude = Number(location?.latitude);
  const longitude = Number(location?.longitude);
  if (!Number.isFinite(latitude) || !Number.isFinite(longitude)) return "https://www.openstreetmap.org";
  const safeZoom = Number.isFinite(Number(zoom)) ? Math.max(3, Math.min(18, Math.floor(Number(zoom)))) : 12;
  return `https://www.openstreetmap.org/?mlat=${latitude.toFixed(6)}&mlon=${longitude.toFixed(
    6
  )}#map=${safeZoom}/${latitude.toFixed(6)}/${longitude.toFixed(6)}`;
}

function buildOpenStreetMapRouteUrl(origin, destination, profile = "car") {
  const originLat = Number(origin?.latitude);
  const originLon = Number(origin?.longitude);
  const destinationLat = Number(destination?.latitude);
  const destinationLon = Number(destination?.longitude);
  if (!Number.isFinite(originLat) || !Number.isFinite(originLon) || !Number.isFinite(destinationLat) || !Number.isFinite(destinationLon)) {
    return "https://www.openstreetmap.org";
  }
  const engine = profile === "foot" ? "fossgis_osrm_foot" : profile === "bike" ? "fossgis_osrm_bike" : "fossgis_osrm_car";
  const routeValue = `${originLat.toFixed(6)},${originLon.toFixed(6)};${destinationLat.toFixed(6)},${destinationLon.toFixed(6)}`;
  const parsed = new URL("https://www.openstreetmap.org/directions");
  parsed.searchParams.set("engine", engine);
  parsed.searchParams.set("route", routeValue);
  return parsed.toString();
}

function buildAmapPlaceUrl(location, locationKeyword = "") {
  const keyword = sanitizeAiWebSearchText(locationKeyword || buildOsmLocationDisplayLabel(location), 120);
  if (keyword) {
    return `https://ditu.amap.com/search?query=${encodeURIComponent(keyword)}`;
  }
  return "https://ditu.amap.com";
}

function buildBaiduPlaceUrl(location, locationKeyword = "") {
  const keyword = sanitizeAiWebSearchText(locationKeyword || buildOsmLocationDisplayLabel(location), 120);
  if (keyword) {
    return `https://map.baidu.com/search/${encodeURIComponent(keyword)}`;
  }
  return "https://map.baidu.com";
}

function buildAmapRouteUrl(origin, destination, profile = "driving") {
  const originLat = Number(origin?.latitude);
  const originLon = Number(origin?.longitude);
  const destinationLat = Number(destination?.latitude);
  const destinationLon = Number(destination?.longitude);
  if (!Number.isFinite(originLat) || !Number.isFinite(originLon) || !Number.isFinite(destinationLat) || !Number.isFinite(destinationLon)) {
    return "https://ditu.amap.com";
  }
  const mode = String(profile || "").toLowerCase() === "walking" ? "walk" : String(profile || "").toLowerCase() === "cycling" ? "ride" : "car";
  const originName = sanitizeAiWebSearchText(buildOsmLocationDisplayLabel(origin), 120) || "起点";
  const destinationName = sanitizeAiWebSearchText(buildOsmLocationDisplayLabel(destination), 120) || "终点";
  return `https://uri.amap.com/navigation?from=${originLon.toFixed(6)},${originLat.toFixed(6)},${encodeURIComponent(
    originName
  )}&to=${destinationLon.toFixed(6)},${destinationLat.toFixed(6)},${encodeURIComponent(
    destinationName
  )}&mode=${mode}&src=991x&coordinate=gaode&callnative=0`;
}

function buildBaiduRouteUrl(origin, destination) {
  const originName = sanitizeAiWebSearchText(buildOsmLocationDisplayLabel(origin), 120) || "起点";
  const destinationName = sanitizeAiWebSearchText(buildOsmLocationDisplayLabel(destination), 120) || "终点";
  const query = `${originName}到${destinationName}`;
  return `https://map.baidu.com/search/${encodeURIComponent(query)}`;
}

function buildPreferredChinaMapPlaceUrl(location, locationKeyword = "") {
  const amapUrl = buildAmapPlaceUrl(location, locationKeyword);
  if (amapUrl) return amapUrl;
  const baiduUrl = buildBaiduPlaceUrl(location, locationKeyword);
  if (baiduUrl) return baiduUrl;
  return buildOpenStreetMapPlaceUrl(location, 12);
}

function buildPreferredChinaMapRouteUrl(origin, destination, profile = "driving") {
  const amapUrl = buildAmapRouteUrl(origin, destination, profile);
  if (amapUrl) return amapUrl;
  return buildOpenStreetMapRouteUrl(origin, destination, profile === "walking" ? "foot" : profile === "cycling" ? "bike" : "car");
}

function buildAmapRouteSearchUrlByName(originName, destinationName) {
  const safeOrigin = sanitizeAiWebSearchText(originName, 80);
  const safeDestination = sanitizeAiWebSearchText(destinationName, 80);
  const query = sanitizeAiWebSearchText(`${safeOrigin}到${safeDestination}`, 180);
  if (!query) return "https://ditu.amap.com";
  return `https://ditu.amap.com/search?query=${encodeURIComponent(query)}`;
}

function buildBaiduRouteSearchUrlByName(originName, destinationName) {
  const safeOrigin = sanitizeAiWebSearchText(originName, 80);
  const safeDestination = sanitizeAiWebSearchText(destinationName, 80);
  const query = sanitizeAiWebSearchText(`${safeOrigin}到${safeDestination}`, 180);
  if (!query) return "https://map.baidu.com";
  return `https://map.baidu.com/search/${encodeURIComponent(query)}`;
}

function isMapLocationLikelyMatch(location, locationKeyword = "") {
  if (!location || typeof location !== "object") return false;
  const keyword = sanitizeAiWebSearchText(locationKeyword, 60).toLowerCase();
  if (!keyword) return true;
  const label = buildOsmLocationDisplayLabel(location).toLowerCase();
  const city = sanitizeAiWebSearchText(location.city, 80).toLowerCase();
  const state = sanitizeAiWebSearchText(location.state, 80).toLowerCase();
  const country = sanitizeAiWebSearchText(location.country, 80).toLowerCase();
  const prefix = extractCnLocationPrefix(keyword).toLowerCase();
  if (prefix && !(label.includes(prefix) || city.includes(prefix) || state.includes(prefix))) {
    return false;
  }
  const compactKeyword = keyword.replace(/\s+/g, "");
  const compactLabel = `${label} ${city} ${state} ${country}`.replace(/\s+/g, "");
  if (compactKeyword.length >= 4) {
    const tail = compactKeyword.slice(-2);
    if (tail && !compactLabel.includes(tail)) {
      return false;
    }
  }
  return true;
}

async function fetchOsrmRouteSummary(origin, destination, profile = "driving") {
  const mode = ["driving", "walking", "cycling"].includes(String(profile || "").toLowerCase())
    ? String(profile).toLowerCase()
    : "driving";
  const originLat = Number(origin?.latitude);
  const originLon = Number(origin?.longitude);
  const destinationLat = Number(destination?.latitude);
  const destinationLon = Number(destination?.longitude);
  if (!Number.isFinite(originLat) || !Number.isFinite(originLon) || !Number.isFinite(destinationLat) || !Number.isFinite(destinationLon)) {
    return null;
  }
  const endpoint = new URL(
    `https://router.project-osrm.org/route/v1/${mode}/${originLon.toFixed(6)},${originLat.toFixed(6)};${destinationLon.toFixed(
      6
    )},${destinationLat.toFixed(6)}`
  );
  endpoint.searchParams.set("overview", "false");
  endpoint.searchParams.set("alternatives", "false");
  endpoint.searchParams.set("steps", "false");
  endpoint.searchParams.set("annotations", "false");
  const result = await fetchAiJson(endpoint.toString(), {
    method: "GET",
    headers: {
      "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)",
      Accept: "application/json,text/plain,*/*",
    },
    timeoutMs: AI_MAP_REQUEST_TIMEOUT_MS,
  });
  if (!result.ok) {
    const error = new Error(`OSRM 路线请求失败（${result.status}）`);
    error.code = "MAP_ROUTE_FAILED";
    error.status = result.status;
    throw error;
  }
  const payload = result.data && typeof result.data === "object" && !("rawText" in result.data) ? result.data : {};
  if (String(payload.code || "").toLowerCase() !== "ok") return null;
  const route = Array.isArray(payload.routes) ? payload.routes[0] : null;
  if (!route || typeof route !== "object") return null;
  const distanceMeters = Number(route.distance);
  const durationSeconds = Number(route.duration);
  if (!Number.isFinite(distanceMeters) || !Number.isFinite(durationSeconds)) return null;
  return {
    profile: mode,
    distanceMeters,
    durationSeconds,
  };
}

function getAiMapProfileLabel(profile) {
  const safe = String(profile || "").toLowerCase();
  if (safe === "walking") return "步行";
  if (safe === "cycling") return "骑行";
  return "驾车";
}

async function fetchOpenStreetMapSearchResults(query, maxResults = 5) {
  if (!isLikelyAiMapQuery(query)) return [];
  const limit = Number.isFinite(Number(maxResults)) ? Math.max(1, Math.min(8, Math.floor(Number(maxResults)))) : 5;
  const nowDateIso = formatAiShanghaiDateTimeParts(new Date()).dateIso;
  const routeIntent = extractAiMapRouteEndpointsFromQuery(query);
  if (routeIntent) {
    let geocodeFailed = false;
    const [originCandidates, destinationCandidates] = await Promise.all(
      [routeIntent.origin, routeIntent.destination].map(async (keyword) => {
        try {
          return await fetchMapGeocodeCandidates(keyword, 8);
        } catch {
          geocodeFailed = true;
          return [];
        }
      })
    );
    const origin = pickBestOsmNominatimLocation(originCandidates, routeIntent.origin);
    const destination = pickBestOsmNominatimLocation(destinationCandidates, routeIntent.destination);
    const originMatched = origin && isMapLocationLikelyMatch(origin, routeIntent.origin);
    const destinationMatched = destination && isMapLocationLikelyMatch(destination, routeIntent.destination);
    if (!origin || !destination || !originMatched || !destinationMatched) {
      const amapSearchUrl = buildAmapRouteSearchUrlByName(routeIntent.origin, routeIntent.destination);
      const baiduSearchUrl = buildBaiduRouteSearchUrlByName(routeIntent.origin, routeIntent.destination);
      const fallbackLines = [
        `起点：${routeIntent.origin}`,
        `终点：${routeIntent.destination}`,
        `导航链接（高德）：${amapSearchUrl}`,
        `备用链接（百度）：${baiduSearchUrl}`,
        geocodeFailed ? "提示：坐标解析受网络影响，已退回按地名导航链接。" : "提示：同名地点较多，已优先提供按地名导航链接。",
        `数据时间基准：${nowDateIso}（Asia/Shanghai）`,
        "数据源：OpenStreetMap/Open-Meteo Geocoding + Amap/Baidu 导航页（免费）",
      ];
      const fallbackItem = sanitizeAiWebSearchItem({
        title: `${routeIntent.origin} 到 ${routeIntent.destination} 路线（地名导航）`,
        url: amapSearchUrl,
        snippet: fallbackLines.join("；"),
      });
      return fallbackItem ? [fallbackItem].slice(0, limit) : [];
    }
    let routeSummary = null;
    let routeError = null;
    try {
      routeSummary = await fetchOsrmRouteSummary(origin, destination, "driving");
    } catch (error) {
      routeError = error;
    }
    if (!routeSummary) {
      try {
        routeSummary = await fetchOsrmRouteSummary(origin, destination, "walking");
      } catch (error) {
        if (!routeError) routeError = error;
      }
    }
    const routeDistance = formatAiMapDistance(routeSummary?.distanceMeters);
    const routeDuration = formatAiMapDuration(routeSummary?.durationSeconds);
    const profileLabel = getAiMapProfileLabel(routeSummary?.profile || "driving");
    const preferredRouteUrl = buildPreferredChinaMapRouteUrl(origin, destination, routeSummary?.profile || "driving");
    const baiduRouteUrl = buildBaiduRouteUrl(origin, destination);
    const osmRouteUrl = buildOpenStreetMapRouteUrl(
      origin,
      destination,
      routeSummary?.profile === "walking" ? "foot" : routeSummary?.profile === "cycling" ? "bike" : "car"
    );
    const snippetLines = [
      `起点：${buildOsmLocationDisplayLabel(origin) || routeIntent.origin}`,
      `终点：${buildOsmLocationDisplayLabel(destination) || routeIntent.destination}`,
      routeDistance ? `${profileLabel}距离：${routeDistance}` : "",
      routeDuration ? `${profileLabel}预计时长：${routeDuration}` : "",
      `导航链接（高德）：${preferredRouteUrl}`,
      `备用链接（百度）：${baiduRouteUrl}`,
      `备用链接（OpenStreetMap）：${osmRouteUrl}`,
      `数据时间基准：${nowDateIso}（Asia/Shanghai）`,
      "数据源：OpenStreetMap/Open-Meteo Geocoding + OSRM（免费）",
    ].filter(Boolean);
    if (!routeSummary && routeError) {
      snippetLines.push("提示：路线摘要服务不可用，已提供可打开的地图路径链接。");
    }
    const item = sanitizeAiWebSearchItem({
      title: `${routeIntent.origin} 到 ${routeIntent.destination} 路线（OpenStreetMap）`,
      url: preferredRouteUrl,
      snippet: snippetLines.join("；"),
    });
    return item ? [item].slice(0, limit) : [];
  }

  const locationKeyword = extractAiMapLocationFromQuery(query);
  if (!locationKeyword) return [];
  let geocodeFailed = false;
  let candidates = [];
  try {
    candidates = await fetchMapGeocodeCandidates(locationKeyword, Math.max(6, limit + 2));
  } catch {
    geocodeFailed = true;
  }
  const selected = candidates.length ? pickBestOsmNominatimLocation(candidates, locationKeyword) : null;
  const selectedMatched = selected && isMapLocationLikelyMatch(selected, locationKeyword);
  if (!selected || !selectedMatched) {
    const fallbackUrl = buildAmapPlaceUrl(null, locationKeyword);
    const fallbackBaiduUrl = buildBaiduPlaceUrl(null, locationKeyword);
    const fallbackLines = [
      `地点：${locationKeyword}`,
      `地图链接（高德）：${fallbackUrl}`,
      `备用链接（百度）：${fallbackBaiduUrl}`,
      geocodeFailed ? "提示：坐标解析受网络影响，已退回按地名打开地图。" : "提示：同名地点较多，已优先提供按地名打开地图。",
      `数据时间基准：${nowDateIso}（Asia/Shanghai）`,
      "数据源：Amap/Baidu 地图页（免费）",
    ];
    const fallbackItem = sanitizeAiWebSearchItem({
      title: `${locationKeyword}地图位置（地名搜索）`,
      url: fallbackUrl,
      snippet: fallbackLines.join("；"),
    });
    return fallbackItem ? [fallbackItem].slice(0, limit) : [];
  }
  const preferredPlaceUrl = buildPreferredChinaMapPlaceUrl(selected, locationKeyword);
  const baiduPlaceUrl = buildBaiduPlaceUrl(selected, locationKeyword);
  const osmPlaceUrl = buildOpenStreetMapPlaceUrl(selected, 12);
  const snippetLines = [
    `地点：${buildOsmLocationDisplayLabel(selected) || locationKeyword}`,
    `坐标：${selected.latitude.toFixed(6)}, ${selected.longitude.toFixed(6)}`,
    selected.type ? `类型：${selected.type}` : "",
    selected.city ? `城市：${selected.city}` : "",
    selected.country ? `国家：${selected.country}` : "",
    `地图链接（高德）：${preferredPlaceUrl}`,
    `备用链接（百度）：${baiduPlaceUrl}`,
    `备用链接（OpenStreetMap）：${osmPlaceUrl}`,
    `数据时间基准：${nowDateIso}（Asia/Shanghai）`,
    "数据源：OpenStreetMap/Open-Meteo Geocoding（免费）",
  ].filter(Boolean);
  const item = sanitizeAiWebSearchItem({
    title: `${locationKeyword}地图位置（OpenStreetMap）`,
    url: preferredPlaceUrl,
    snippet: snippetLines.join("；"),
  });
  return item ? [item].slice(0, limit) : [];
}

const AI_ON_THIS_DAY_TRUSTED_HOSTS = ["baike.baidu.com", "timeanddate.com", "wikipedia.org", "history.com", "gov.cn"];

function normalizeAiHostname(rawUrl) {
  try {
    const parsed = new URL(String(rawUrl || "").trim());
    return String(parsed.hostname || "")
      .trim()
      .toLowerCase()
      .replace(/^www\./, "");
  } catch {
    return "";
  }
}

function isTrustedOnThisDayUrl(rawUrl) {
  const hostname = normalizeAiHostname(rawUrl);
  if (!hostname) return false;
  return AI_ON_THIS_DAY_TRUSTED_HOSTS.some((trustedHost) => hostname === trustedHost || hostname.endsWith(`.${trustedHost}`));
}

function getOnThisDayEnglishMonthSlug(monthNumber) {
  const month = Number(monthNumber);
  const map = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
  if (!Number.isInteger(month) || month < 1 || month > 12) return "";
  return map[month - 1] || "";
}

function getAiOnThisDayDateParts(dateIntent = {}) {
  const explicitDateKey = sanitizeAiWebSearchText(dateIntent?.startDateKey, 20);
  const fallbackDateKey = formatAiShanghaiDateTimeParts(new Date()).dateIso;
  const dateKey = /^\d{4}-\d{2}-\d{2}$/.test(explicitDateKey) ? explicitDateKey : fallbackDateKey;
  const matched = dateKey.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!matched) {
    return {
      dateKey: fallbackDateKey,
      year: Number(fallbackDateKey.slice(0, 4)),
      month: Number(fallbackDateKey.slice(5, 7)),
      day: Number(fallbackDateKey.slice(8, 10)),
      monthSlug: getOnThisDayEnglishMonthSlug(Number(fallbackDateKey.slice(5, 7))),
    };
  }
  return {
    dateKey,
    year: Number(matched[1]),
    month: Number(matched[2]),
    day: Number(matched[3]),
    monthSlug: getOnThisDayEnglishMonthSlug(Number(matched[2])),
  };
}

function buildTrustedOnThisDaySources(dateParts) {
  const month = Number(dateParts?.month);
  const day = Number(dateParts?.day);
  const monthSlug = sanitizeAiWebSearchText(dateParts?.monthSlug, 20).toLowerCase();
  if (!Number.isInteger(month) || !Number.isInteger(day) || !monthSlug) return [];
  return [
    {
      source: "timeanddate",
      label: "Timeanddate",
      url: `https://www.timeanddate.com/on-this-day/${monthSlug}/${day}`,
      parseEvents: true,
    },
    {
      source: "baike.baidu.com",
      label: "百度百科",
      url: `https://baike.baidu.com/item/${month}月${day}日`,
      parseEvents: false,
    },
  ];
}

function extractHtmlTitleAndDescription(rawHtml) {
  const html = String(rawHtml || "");
  if (!html) return { title: "", description: "" };
  const titleMatch = html.match(/<title>([\s\S]*?)<\/title>/i);
  const descMatch = html.match(/<meta[^>]*name=["']description["'][^>]*content=["']([^"']+)["']/i);
  return {
    title: sanitizeAiWebSearchText(titleMatch ? titleMatch[1] : "", 220),
    description: sanitizeAiWebSearchText(descMatch ? descMatch[1] : "", 420),
  };
}

function extractOnThisDayEventsFromHtml(rawHtml, maxCount = 5) {
  const html = decodeHtmlEntities(String(rawHtml || ""));
  if (!html) return [];
  const output = [];
  const seen = new Set();
  const safeLimit = Number.isFinite(Number(maxCount)) ? Math.max(1, Math.floor(Number(maxCount))) : 5;
  const addEvent = (rawValue) => {
    const safe = sanitizeAiWebSearchText(rawValue, 240)
      .replace(/\s+/g, " ")
      .trim();
    if (!safe || safe.length < 14) return;
    if (/^(logo|calendar|javascript|width|height|svg|icon)$/i.test(safe)) return;
    const key = safe.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    output.push(safe);
  };

  const yearLinePattern = /(?:^|[>\s])((?:1[0-9]{3}|20[0-2][0-9]))\s*[–-]\s*([^<\n\r]{12,260})/g;
  let matched;
  while ((matched = yearLinePattern.exec(html)) !== null) {
    addEvent(`${matched[1]}年：${matched[2]}`);
    if (output.length >= safeLimit) break;
  }
  if (!output.length) {
    const altPattern = /((?:1[0-9]{3}|20[0-2][0-9])年[^<\n\r]{8,220})/g;
    while ((matched = altPattern.exec(html)) !== null) {
      addEvent(matched[1]);
      if (output.length >= safeLimit) break;
    }
  }
  return output.slice(0, safeLimit);
}

function buildTrustedOnThisDaySnippet(dateParts, sourceLabel, description, events) {
  const month = Number(dateParts?.month);
  const day = Number(dateParts?.day);
  const dateKey = sanitizeAiWebSearchText(dateParts?.dateKey, 20);
  const lines = [];
  lines.push(`日期：${dateKey || `${month}月${day}日`}`);
  lines.push(`来源：${sanitizeAiWebSearchText(sourceLabel, 60)}`);
  if (description) lines.push(`简介：${sanitizeAiWebSearchText(description, 260)}`);
  if (Array.isArray(events) && events.length) {
    lines.push(`事件摘录：${events.slice(0, 4).join("；")}`);
  } else {
    lines.push("可进入链接查看当日历史事件列表。");
  }
  return sanitizeAiWebSearchText(lines.join("；"), 520);
}

async function fetchTrustedOnThisDayResults(plan = {}, maxResults = 5) {
  if (!plan || plan.onThisDayIntent !== true) return [];
  const dateParts = getAiOnThisDayDateParts(plan.dateIntent);
  const sources = buildTrustedOnThisDaySources(dateParts);
  if (!sources.length) return [];

  const settled = await Promise.allSettled(
    sources.map(async (config) => {
      const response = await fetchAiJson(config.url, {
        method: "GET",
        headers: {
          "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)",
          Accept: "text/html,application/xhtml+xml",
        },
        timeoutMs: AI_ON_THIS_DAY_SOURCE_TIMEOUT_MS,
      });
      if (!response.ok) {
        const error = new Error(`${config.source} 请求失败（${response.status}）`);
        error.code = "TRUSTED_ON_THIS_DAY_FAILED";
        error.status = response.status;
        throw error;
      }
      const rawHtml = String(response.text || response?.data?.rawText || "");
      const meta = extractHtmlTitleAndDescription(rawHtml);
      const events = config.parseEvents ? extractOnThisDayEventsFromHtml(rawHtml, 5) : [];
      const snippet = buildTrustedOnThisDaySnippet(dateParts, config.label, meta.description, events);
      return sanitizeAiWebSearchItem({
        title: meta.title || `${config.label}：${dateParts.month}月${dateParts.day}日`,
        url: config.url,
        snippet,
      });
    })
  );

  const output = [];
  for (const entry of settled) {
    if (entry.status !== "fulfilled") continue;
    if (!entry.value) continue;
    output.push(entry.value);
  }
  const limited = dedupeAiWebSearchStringList(output.map((item) => item.url), 20);
  if (!limited.length) return [];
  const deduped = [];
  const urlSet = new Set();
  for (const item of output) {
    if (!item || typeof item !== "object") continue;
    const key = String(item.url || "").toLowerCase();
    if (!key || urlSet.has(key)) continue;
    urlSet.add(key);
    deduped.push(item);
  }
  return deduped.slice(0, Math.max(1, Math.min(8, Number(maxResults) || 5)));
}

const AI_WEB_SEARCH_STOP_WORDS = new Set([
  "请",
  "请你",
  "请问",
  "帮我",
  "告诉我",
  "看一下",
  "查一下",
  "搜一下",
  "搜索",
  "联网",
  "上网",
  "一下",
  "一下吧",
  "一下吗",
  "看看",
  "这部",
  "这条",
  "这个",
  "电影",
  "新闻",
  "消息",
  "内容",
  "信息",
  "相关",
  "关于",
  "这",
  "那",
  "的",
  "了",
  "吗",
  "呢",
  "啊",
  "呀",
  "吧",
  "多少",
  "是",
  "是什么",
  "有没有",
  "能否",
  "是否",
  "please",
  "search",
  "online",
  "web",
  "internet",
]);
const AI_WEB_SEARCH_STOP_WORD_LIST = Array.from(AI_WEB_SEARCH_STOP_WORDS).sort((a, b) => b.length - a.length);
const AI_WEB_SEARCH_REGION_AUTO = "auto";
const AI_WEB_SEARCH_REGION_MAINLAND_CHINA = "mainland-china";
const AI_WEB_SEARCH_REGION_OUTSIDE_MAINLAND_CHINA = "outside-mainland-china";
const AI_WEB_SEARCH_ENGINE_BING = "bing-rss";
const AI_WEB_SEARCH_ENGINE_BAIDU = "baidu-html";
const AI_WEB_SEARCH_ENGINE_GOOGLE = "google-html";
const AI_WEB_SEARCH_MAINLAND_DEFAULT_ENGINES = [AI_WEB_SEARCH_ENGINE_BING, AI_WEB_SEARCH_ENGINE_BAIDU];
const AI_WEB_SEARCH_OUTSIDE_MAINLAND_DEFAULT_ENGINES = [AI_WEB_SEARCH_ENGINE_BING, AI_WEB_SEARCH_ENGINE_GOOGLE];
const AI_WEB_SEARCH_FRESHNESS_QUERY_PATTERN =
  /(最新|最近|近期|刚刚|刚发布|最新发布|最新消息|latest|recent|newest|most\s+recent|up[-\s]?to[-\s]?date|just\s+(updated|released))/i;
const AI_WEB_SEARCH_FRESHNESS_SIGNAL_PATTERN =
  /(刚刚|分钟(?:前)?|小时(?:前)?|today|yesterday|hours?\s+ago|minutes?\s+ago|updated|发布于|published|更新于|recently)/i;

function normalizeAiWebSearchResultUrl(rawValue) {
  const raw = decodeHtmlEntities(String(rawValue || "")).trim();
  if (!raw) return "";
  let normalized = raw;
  if (normalized.startsWith("//")) normalized = `https:${normalized}`;
  try {
    const parsed = new URL(normalized);
    if (!/^https?:$/i.test(parsed.protocol)) return "";
    parsed.hash = "";
    return parsed.toString();
  } catch {
    return /^https?:\/\//i.test(normalized) ? normalized : "";
  }
}

function sanitizeAiWebSearchTerm(rawValue) {
  return sanitizeAiWebSearchText(rawValue, 64)
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function sanitizeAiWebSearchEngineId(rawValue) {
  const safe = sanitizeAiWebSearchTerm(rawValue).replace(/_/g, "-");
  if (!safe || safe === "auto") return "";
  if (["bing", "bing-rss", "msn"].includes(safe)) return AI_WEB_SEARCH_ENGINE_BING;
  if (["baidu", "baidu-html"].includes(safe)) return AI_WEB_SEARCH_ENGINE_BAIDU;
  if (["google", "google-html"].includes(safe)) return AI_WEB_SEARCH_ENGINE_GOOGLE;
  return "";
}

function sanitizeAiWebSearchEnginePreference(rawValue) {
  const safe = sanitizeAiWebSearchTerm(rawValue).replace(/_/g, "-");
  if (!safe || ["auto", "default", "region"].includes(safe)) return "";
  return sanitizeAiWebSearchEngineId(safe);
}

function sanitizeAiWebSearchRegionHint(rawValue) {
  const safe = sanitizeAiWebSearchTerm(rawValue).replace(/_/g, "-");
  if (!safe || safe === AI_WEB_SEARCH_REGION_AUTO) return AI_WEB_SEARCH_REGION_AUTO;
  if (
    [
      "mainland",
      "mainland-cn",
      "mainland-china",
      "china-mainland",
      "cn-mainland",
      "zh-cn-mainland",
      "cn",
      "china",
    ].includes(safe)
  ) {
    return AI_WEB_SEARCH_REGION_MAINLAND_CHINA;
  }
  if (
    [
      "outside-mainland",
      "outside-mainland-cn",
      "outside-mainland-china",
      "overseas",
      "global",
      "intl",
      "international",
      "non-mainland",
    ].includes(safe)
  ) {
    return AI_WEB_SEARCH_REGION_OUTSIDE_MAINLAND_CHINA;
  }
  return AI_WEB_SEARCH_REGION_AUTO;
}

function normalizeAiWebSearchBoolean(rawValue, defaultValue = false) {
  if (rawValue === undefined || rawValue === null || rawValue === "") return defaultValue === true;
  if (typeof rawValue === "boolean") return rawValue;
  if (typeof rawValue === "number") return Number.isFinite(rawValue) && rawValue !== 0;
  const safe = sanitizeAiWebSearchTerm(rawValue);
  if (!safe) return defaultValue === true;
  if (["1", "true", "yes", "on", "enable", "enabled"].includes(safe)) return true;
  if (["0", "false", "no", "off", "disable", "disabled"].includes(safe)) return false;
  return defaultValue === true;
}

function hasAiWebSearchFreshnessIntent(query) {
  const safeQuery = sanitizeAiWebSearchQuery(query);
  if (!safeQuery) return false;
  return AI_WEB_SEARCH_FRESHNESS_QUERY_PATTERN.test(safeQuery);
}

function dedupeAiWebSearchEngineIds(rawList, maxCount = 6) {
  const limit = Number.isFinite(Number(maxCount)) ? Math.max(1, Math.floor(Number(maxCount))) : 6;
  const output = [];
  const seen = new Set();
  for (const entry of Array.isArray(rawList) ? rawList : []) {
    const engineId = sanitizeAiWebSearchEngineId(entry);
    if (!engineId || seen.has(engineId)) continue;
    seen.add(engineId);
    output.push(engineId);
    if (output.length >= limit) break;
  }
  return output;
}

function inferAiWebSearchRegion(rawPayload = {}) {
  const explicitRegion =
    sanitizeAiWebSearchRegionHint(rawPayload?.regionHint || rawPayload?.region || rawPayload?.webSearchRegion || rawPayload?.regionPreference);
  if (explicitRegion !== AI_WEB_SEARCH_REGION_AUTO) return explicitRegion;
  const timezone = sanitizeAiWebSearchText(rawPayload?.clientTimeZone || "", 80);
  const locale = sanitizeAiWebSearchText(rawPayload?.clientLocale || rawPayload?.clientLanguage || rawPayload?.locale || "", 40).toLowerCase();
  if (/^asia\/(shanghai|urumqi)$/i.test(timezone)) return AI_WEB_SEARCH_REGION_MAINLAND_CHINA;
  if (locale.startsWith("zh-cn") || locale === "zh-hans-cn") return AI_WEB_SEARCH_REGION_MAINLAND_CHINA;
  return AI_WEB_SEARCH_REGION_OUTSIDE_MAINLAND_CHINA;
}

function buildAiWebSearchDefaultGeneralEngineOrder(region) {
  if (region === AI_WEB_SEARCH_REGION_MAINLAND_CHINA) return [...AI_WEB_SEARCH_MAINLAND_DEFAULT_ENGINES];
  return [...AI_WEB_SEARCH_OUTSIDE_MAINLAND_DEFAULT_ENGINES];
}

function resolveAiWebSearchGeneralEnginePlan(rawPayload = {}) {
  const region = inferAiWebSearchRegion(rawPayload);
  const preferredEngine = sanitizeAiWebSearchEnginePreference(
    rawPayload?.enginePreference || rawPayload?.webSearchEnginePreference || ""
  );
  const strictEngineRequested = normalizeAiWebSearchBoolean(
    rawPayload?.strictEngine ?? rawPayload?.strictEngineMode ?? rawPayload?.enforcePreferredEngine,
    true
  );
  const allowEngineFallback = normalizeAiWebSearchBoolean(
    rawPayload?.allowEngineFallback ?? rawPayload?.engineFallback ?? rawPayload?.allowFallback,
    true
  );
  const strictEngineApplied = strictEngineRequested && Boolean(preferredEngine);
  const defaultEngines = buildAiWebSearchDefaultGeneralEngineOrder(region);
  const configuredEngines = preferredEngine
    ? strictEngineApplied || !allowEngineFallback
      ? dedupeAiWebSearchEngineIds([preferredEngine], 1)
      : dedupeAiWebSearchEngineIds([preferredEngine, ...defaultEngines], 3)
    : dedupeAiWebSearchEngineIds(defaultEngines, 3);
  return {
    region,
    preferredEngine: preferredEngine || "auto",
    configuredEngines,
    strictEngineRequested,
    strictEngineApplied,
    allowEngineFallback,
  };
}

function normalizeAiWebSearchIntentToken(rawValue) {
  const safe = sanitizeAiWebSearchTerm(rawValue);
  if (!safe) return "";
  return safe
    .replace(/^[的是在与及和请]+/g, "")
    .replace(/(是多少|是什么|怎么样|如何|多少|吗|呢|吧|啊|呀|是)$/g, "")
    .trim();
}

function stripAiWebSearchStopWords(rawValue) {
  let text = sanitizeAiWebSearchTerm(rawValue);
  if (!text) return "";
  for (const stopWord of AI_WEB_SEARCH_STOP_WORD_LIST) {
    const safeStopWord = sanitizeAiWebSearchTerm(stopWord);
    if (!safeStopWord || safeStopWord.length < 2) continue;
    text = text.replace(new RegExp(escapeAiWebSearchRegExp(safeStopWord), "gi"), " ");
  }
  return text.replace(/\s+/g, " ").trim();
}

function escapeAiWebSearchRegExp(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function tokenizeAiWebSearchQuery(query) {
  const text = sanitizeAiWebSearchQuery(query);
  if (!text) return [];
  const tokens = text.match(/[A-Za-z0-9][A-Za-z0-9._:+/-]{1,}|[\u4e00-\u9fff]{2,}/g) || [];
  const output = [];
  const seen = new Set();
  for (const token of tokens) {
    const safe = sanitizeAiWebSearchTerm(token);
    if (!safe || safe.length < 2) continue;
    if (seen.has(safe)) continue;
    seen.add(safe);
    output.push(safe);
  }
  return output;
}

function extractAiWebSearchQuotedTerms(query) {
  const text = sanitizeAiWebSearchQuery(query);
  if (!text) return [];
  const output = [];
  const seen = new Set();
  const patterns = [
    /《([^《》]{1,60})》/g,
    /「([^「」]{1,60})」/g,
    /『([^『』]{1,60})』/g,
    /[“"]([^“”"]{1,60})[”"]/g,
    /[‘']([^‘’']{1,60})[’']/g,
  ];
  for (const pattern of patterns) {
    let match;
    while ((match = pattern.exec(text)) !== null) {
      const safe = sanitizeAiWebSearchTerm(match[1]);
      if (!safe || safe.length < 2 || seen.has(safe)) continue;
      seen.add(safe);
      output.push(safe);
    }
  }
  return output;
}

function formatAiWebSearchDateKey(dateValue) {
  if (!(dateValue instanceof Date) || Number.isNaN(dateValue.getTime())) return "";
  const year = dateValue.getFullYear();
  const month = String(dateValue.getMonth() + 1).padStart(2, "0");
  const day = String(dateValue.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function createAiWebSearchDateFromParts(yearValue, monthValue, dayValue) {
  const year = Number(yearValue);
  const month = Number(monthValue);
  const day = Number(dayValue);
  if (!Number.isInteger(year) || year < 1900 || year > 2100) return null;
  if (!Number.isInteger(month) || month < 1 || month > 12) return null;
  if (!Number.isInteger(day) || day < 1 || day > 31) return null;
  const date = new Date(year, month - 1, day);
  if (
    Number.isNaN(date.getTime()) ||
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }
  return date;
}

function shiftAiWebSearchDate(baseDate, offsetDays) {
  if (!(baseDate instanceof Date) || Number.isNaN(baseDate.getTime())) return null;
  const offset = Number(offsetDays);
  const safeOffset = Number.isFinite(offset) ? Math.floor(offset) : 0;
  const shifted = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());
  shifted.setDate(shifted.getDate() + safeOffset);
  return shifted;
}

function buildAiWebSearchDateTerms(dateKey) {
  const matched = String(dateKey || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!matched) return [];
  const year = matched[1];
  const month = Number(matched[2]);
  const day = Number(matched[3]);
  if (!Number.isInteger(month) || !Number.isInteger(day)) return [];
  return [
    `${year}-${matched[2]}-${matched[3]}`,
    `${year}/${matched[2]}/${matched[3]}`,
    `${year}.${matched[2]}.${matched[3]}`,
    `${year}-${month}-${day}`,
    `${year}/${month}/${day}`,
    `${year}.${month}.${day}`,
    `${year}年${month}月${day}日`,
    `${month}月${day}日`,
  ];
}

function dedupeAiWebSearchStringList(inputValues, maxCount = 80) {
  const output = [];
  const seen = new Set();
  const limit = Number.isFinite(Number(maxCount)) ? Math.max(1, Math.floor(Number(maxCount))) : 80;
  for (const rawValue of Array.isArray(inputValues) ? inputValues : []) {
    const safe = sanitizeAiWebSearchText(rawValue, 260);
    if (!safe) continue;
    const key = safe.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    output.push(safe);
    if (output.length >= limit) break;
  }
  return output;
}

function buildAiWebSearchDateIntent(query) {
  const safeQuery = sanitizeAiWebSearchQuery(query);
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const explicitDateKeys = [];
  const explicitDateSet = new Set();
  const addExplicitDate = (yearValue, monthValue, dayValue) => {
    const date = createAiWebSearchDateFromParts(yearValue, monthValue, dayValue);
    if (!date) return;
    const key = formatAiWebSearchDateKey(date);
    if (!key || explicitDateSet.has(key)) return;
    explicitDateSet.add(key);
    explicitDateKeys.push(key);
  };

  const fullDatePattern = /(20\d{2})[-/.年](1[0-2]|0?[1-9])[-/.月](3[01]|[12]\d|0?[1-9])(?:日|号)?/g;
  let fullMatch;
  while ((fullMatch = fullDatePattern.exec(safeQuery)) !== null) {
    addExplicitDate(fullMatch[1], fullMatch[2], fullMatch[3]);
  }
  const compactDatePattern = /\b(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\b/g;
  let compactMatch;
  while ((compactMatch = compactDatePattern.exec(safeQuery)) !== null) {
    addExplicitDate(compactMatch[1], compactMatch[2], compactMatch[3]);
  }
  const monthDayPattern = /(1[0-2]|0?[1-9])月(3[01]|[12]\d|0?[1-9])(?:日|号)?/g;
  let monthDayMatch;
  while ((monthDayMatch = monthDayPattern.exec(safeQuery)) !== null) {
    addExplicitDate(today.getFullYear(), monthDayMatch[1], monthDayMatch[2]);
  }

  let startDate = null;
  let endDate = null;
  let searxTimeRange = "";
  let relativeKeyword = "";
  let freshnessIntent = hasAiWebSearchFreshnessIntent(safeQuery);

  if (explicitDateKeys.length) {
    const sorted = [...explicitDateKeys].sort();
    const first = sorted[0];
    const last = sorted[sorted.length - 1];
    const firstDate = createAiWebSearchDateFromParts(first.slice(0, 4), first.slice(5, 7), first.slice(8, 10));
    const lastDate = createAiWebSearchDateFromParts(last.slice(0, 4), last.slice(5, 7), last.slice(8, 10));
    startDate = firstDate;
    endDate = lastDate || firstDate;
    if (startDate && endDate) {
      const diffDays = Math.max(
        0,
        Math.floor(
          (new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate()).getTime() -
            new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()).getTime()) /
            (24 * 60 * 60 * 1000)
        )
      );
      if (diffDays <= 1) searxTimeRange = "day";
      else if (diffDays <= 45) searxTimeRange = "month";
      else searxTimeRange = "year";
    }
    relativeKeyword = "explicit-date";
  } else if (/(今天|今日|today)/i.test(safeQuery)) {
    startDate = today;
    endDate = today;
    searxTimeRange = "day";
    relativeKeyword = "today";
  } else if (/(昨天|昨日|yesterday)/i.test(safeQuery)) {
    startDate = shiftAiWebSearchDate(today, -1);
    endDate = startDate;
    searxTimeRange = "day";
    relativeKeyword = "yesterday";
  } else if (/(明天|明日|tomorrow)/i.test(safeQuery)) {
    startDate = shiftAiWebSearchDate(today, 1);
    endDate = startDate;
    searxTimeRange = "day";
    relativeKeyword = "tomorrow";
  } else if (/(本周|这周|本星期|这星期|本礼拜|这礼拜|最近7天|近7天|过去7天|7天内)/i.test(safeQuery)) {
    startDate = shiftAiWebSearchDate(today, -6);
    endDate = today;
    searxTimeRange = "month";
    relativeKeyword = "week";
  } else if (/(本月|这个月|当月|最近30天|近30天|过去30天|30天内)/i.test(safeQuery)) {
    startDate = new Date(today.getFullYear(), today.getMonth(), 1);
    endDate = today;
    searxTimeRange = "month";
    relativeKeyword = "month";
  } else if (/(今年|本年|this\s+year)/i.test(safeQuery)) {
    startDate = new Date(today.getFullYear(), 0, 1);
    endDate = today;
    searxTimeRange = "year";
    relativeKeyword = "year";
  }

  if (!freshnessIntent && ["today", "yesterday", "tomorrow", "week", "month"].includes(relativeKeyword)) {
    freshnessIntent = true;
  }
  if (!freshnessIntent && startDate && endDate) {
    const dayMs = 24 * 60 * 60 * 1000;
    const endAnchor = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());
    const todayAnchor = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const recencyDays = Math.abs(Math.floor((todayAnchor.getTime() - endAnchor.getTime()) / dayMs));
    if (recencyDays <= 45) freshnessIntent = true;
  }

  if (freshnessIntent && !startDate) {
    startDate = shiftAiWebSearchDate(today, -30);
    endDate = today;
    if (!searxTimeRange) searxTimeRange = "month";
    if (!relativeKeyword) relativeKeyword = "freshness";
  } else if (freshnessIntent && !searxTimeRange) {
    searxTimeRange = "month";
  }

  const startDateKey = formatAiWebSearchDateKey(startDate);
  const endDateKey = formatAiWebSearchDateKey(endDate || startDate);
  const implicitFreshnessRange = freshnessIntent && !explicitDateKeys.length && relativeKeyword === "freshness";
  const absoluteDateTerms = [];
  if (explicitDateKeys.length) {
    for (const key of explicitDateKeys.sort()) {
      absoluteDateTerms.push(...buildAiWebSearchDateTerms(key));
    }
  } else if (startDateKey && !implicitFreshnessRange) {
    absoluteDateTerms.push(...buildAiWebSearchDateTerms(startDateKey));
    if (endDateKey && endDateKey !== startDateKey) {
      absoluteDateTerms.push(...buildAiWebSearchDateTerms(endDateKey));
      absoluteDateTerms.push(`${startDateKey} ${endDateKey}`);
      absoluteDateTerms.push(`${startDateKey} 到 ${endDateKey}`);
    }
  }

  return {
    hasTimeIntent: Boolean(startDateKey),
    relativeKeyword,
    startDateKey,
    endDateKey,
    explicitDateKeys: explicitDateKeys.sort(),
    absoluteDateTerms: dedupeAiWebSearchStringList(absoluteDateTerms, 16),
    searxTimeRange: ["day", "month", "year"].includes(searxTimeRange) ? searxTimeRange : "",
    freshnessIntent,
    implicitFreshnessRange,
  };
}

function buildAiWebSearchDateEnhancedQueries(query, dateIntent = {}) {
  const safeQuery = sanitizeAiWebSearchQuery(query);
  if (!safeQuery) return [];
  const output = [];
  const seen = new Set();
  const comparableBaseQuery = sanitizeAiWebSearchTerm(safeQuery);
  const hasTermInBaseQuery = (rawTerm) => {
    const safeTerm = sanitizeAiWebSearchTerm(rawTerm);
    if (!safeTerm) return false;
    return comparableBaseQuery.includes(safeTerm);
  };
  const add = (rawValue) => {
    const safe = sanitizeAiWebSearchQuery(rawValue);
    if (!safe || seen.has(safe)) return;
    seen.add(safe);
    output.push(safe);
  };

  const implicitFreshnessRange = dateIntent?.implicitFreshnessRange === true;
  const terms = implicitFreshnessRange ? [] : Array.isArray(dateIntent.absoluteDateTerms) ? dateIntent.absoluteDateTerms.slice(0, 3) : [];
  if (terms.length) {
    const extraTerms = terms.filter((term) => !hasTermInBaseQuery(term));
    if (extraTerms.length >= 1) add(`${safeQuery} ${extraTerms[0]}`);
    if (extraTerms.length >= 2) add(`${safeQuery} ${extraTerms[0]} ${extraTerms[1]}`);
  }
  const startDateKey = sanitizeAiWebSearchQuery(dateIntent.startDateKey || "");
  const endDateKey = sanitizeAiWebSearchQuery(dateIntent.endDateKey || "");
  if (!implicitFreshnessRange && startDateKey && endDateKey && startDateKey !== endDateKey) {
    add(`${safeQuery} ${startDateKey} ${endDateKey}`);
  }
  if (startDateKey && /(今天|今日|昨天|昨日|明天|明日|today|yesterday|tomorrow)/i.test(safeQuery)) {
    add(safeQuery.replace(/今天|今日|昨天|昨日|明天|明日|today|yesterday|tomorrow/gi, startDateKey));
  }
  if (isLikelyAiWeatherQuery(safeQuery)) {
    const location = extractAiWeatherLocationFromQuery(safeQuery);
    if (location) {
      add(`${location}天气预报`);
      add(`${location}天气`);
      if (startDateKey) add(`${location}天气 ${startDateKey}`);
    }
  }
  if (/(历史上的今天|on\s+this\s+day)/i.test(safeQuery)) {
    const matched = String(startDateKey || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (matched) {
      const monthNum = Number(matched[2]);
      const dayNum = Number(matched[3]);
      add(`${monthNum}月${dayNum}日 历史上的今天`);
      add(`${monthNum}月${dayNum}日 历史事件`);
      add(`${matched[1]}-${matched[2]}-${matched[3]} 历史上的今天`);
    }
  }
  if (dateIntent?.freshnessIntent === true) {
    const now = new Date();
    const yearText = String(now.getFullYear());
    const monthText = `${now.getMonth() + 1}月`;
    if (!hasTermInBaseQuery("最新")) add(`${safeQuery} 最新`);
    if (!hasTermInBaseQuery("更新")) add(`${safeQuery} 更新`);
    if (!hasTermInBaseQuery(yearText)) add(`${safeQuery} ${yearText}`);
    add(`${safeQuery} ${yearText} ${monthText}`);
    add(`${safeQuery} 最新发布`);
    if (/\bopenai\b/i.test(safeQuery)) {
      add("openai api models");
      add("openai api models latest");
      add("site:openai.com api models");
      add("site:platform.openai.com docs models");
    }
  }
  add(safeQuery);
  if (output.length <= 6) return output;
  const compact = output.slice(0, 5);
  if (!compact.includes(safeQuery)) compact.push(safeQuery);
  return compact.slice(0, 6);
}

function normalizeAiWebSearchSearxBaseUrl(rawValue) {
  const source = String(rawValue || "").trim();
  if (!source) return "";
  const candidate = /^https?:\/\//i.test(source) ? source : `https://${source}`;
  try {
    const parsed = new URL(candidate);
    if (!/^https?:$/i.test(parsed.protocol)) return "";
    if (/\.onion$/i.test(parsed.hostname || "")) return "";
    parsed.search = "";
    parsed.hash = "";
    const safePath = String(parsed.pathname || "").replace(/\/+$/g, "");
    return `${parsed.origin}${safePath && safePath !== "/" ? safePath : ""}`;
  } catch {
    return "";
  }
}

function parseAiWebSearchSearxEnvInstances(rawValue) {
  return dedupeAiWebSearchStringList(
    String(rawValue || "")
      .split(/[\n,;]+/g)
      .map((entry) => normalizeAiWebSearchSearxBaseUrl(entry))
      .filter(Boolean),
    30
  );
}

function getAiWebSearchDefaultSearxInstances() {
  const envInstances = parseAiWebSearchSearxEnvInstances(AI_WEB_SEARCH_SEARX_INSTANCES_ENV);
  const defaults = dedupeAiWebSearchStringList(
    AI_WEB_SEARCH_SEARX_DEFAULT_INSTANCES.map((entry) => normalizeAiWebSearchSearxBaseUrl(entry)).filter(Boolean),
    30
  );
  return dedupeAiWebSearchStringList([...envInstances, ...defaults], 40);
}

function parseAiWebSearchSearxSpaceInstances(rawPayload) {
  const output = [];
  const add = (rawUrl) => {
    const normalized = normalizeAiWebSearchSearxBaseUrl(rawUrl);
    if (!normalized) return;
    output.push(normalized);
  };

  const instancesNode = rawPayload?.instances;
  if (Array.isArray(instancesNode)) {
    for (const entry of instancesNode) {
      if (typeof entry === "string") add(entry);
      else if (entry && typeof entry === "object") add(entry.url || entry.baseUrl || entry.instance || "");
    }
  } else if (instancesNode && typeof instancesNode === "object") {
    for (const [rawUrl, meta] of Object.entries(instancesNode)) {
      const networkType = sanitizeAiWebSearchTerm(meta?.network_type || meta?.networkType || meta?.network?.type || "");
      if (networkType && /(tor|onion|i2p)/i.test(networkType)) continue;
      if (meta && typeof meta === "object" && typeof meta.error === "string" && meta.error.trim()) continue;
      add(rawUrl);
      if (meta && typeof meta === "object") {
        add(meta.url || meta.baseUrl || meta.instance || "");
      }
    }
  }

  return dedupeAiWebSearchStringList(output, 120);
}

async function loadAiWebSearchSearxInstancePool() {
  const nowMs = Date.now();
  if (aiWebSearchSearxPoolCache.expiresAtMs > nowMs && aiWebSearchSearxPoolCache.instances.length) {
    return aiWebSearchSearxPoolCache.instances;
  }

  const defaultPool = getAiWebSearchDefaultSearxInstances();
  let mergedPool = [...defaultPool];
  if (AI_WEB_SEARCH_SEARX_SPACE_INSTANCES_URL) {
    try {
      const result = await fetchAiJson(AI_WEB_SEARCH_SEARX_SPACE_INSTANCES_URL, {
        method: "GET",
        headers: { "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)" },
        timeoutMs: AI_WEB_SEARCH_SEARX_SPACE_TIMEOUT_MS,
      });
      if (result.ok && result.data && typeof result.data === "object" && !("rawText" in result.data)) {
        const remotePool = parseAiWebSearchSearxSpaceInstances(result.data);
        if (remotePool.length) mergedPool = dedupeAiWebSearchStringList([...remotePool, ...defaultPool], 120);
      }
    } catch {
      // Ignore searx.space fetch failures and keep built-in defaults.
    }
  }

  const finalPool = mergedPool.length ? mergedPool : defaultPool;
  aiWebSearchSearxPoolCache = {
    expiresAtMs: nowMs + AI_WEB_SEARCH_SEARX_POOL_CACHE_MS,
    instances: finalPool,
  };
  return finalPool;
}

function hashAiWebSearchText(rawValue) {
  const text = String(rawValue || "");
  let hash = 0;
  for (let i = 0; i < text.length; i += 1) {
    hash = (hash * 31 + text.charCodeAt(i)) >>> 0;
  }
  return hash >>> 0;
}

function pickAiWebSearchSearxInstances(pool, query, limit = 2) {
  const safePool = dedupeAiWebSearchStringList(pool, 120);
  if (!safePool.length) return [];
  const safeLimit = Number.isFinite(Number(limit)) ? Math.max(1, Math.floor(Number(limit))) : 2;
  if (safePool.length <= safeLimit) return safePool;
  const picked = [];
  const offset = hashAiWebSearchText(query) % safePool.length;
  for (let i = 0; i < safeLimit; i += 1) {
    picked.push(safePool[(offset + i) % safePool.length]);
  }
  return picked;
}

function parseAiWebSearchJsonResults(jsonPayload, maxResults = 5) {
  const limit = Math.max(1, Math.min(8, Number(maxResults) || 5));
  const rawResults = Array.isArray(jsonPayload?.results) ? jsonPayload.results : [];
  const output = [];
  for (const rawItem of rawResults) {
    const snippetParts = [
      rawItem?.content,
      rawItem?.snippet,
      rawItem?.description,
      rawItem?.publishedDate,
      rawItem?.publishedDateTime,
      rawItem?.published,
    ].filter(Boolean);
    const item = sanitizeAiWebSearchItem({
      title: rawItem?.title || rawItem?.title_html || "",
      url: rawItem?.url || rawItem?.link || "",
      snippet: snippetParts.join(" "),
    });
    if (!item) continue;
    output.push(item);
    if (output.length >= limit) break;
  }
  return output;
}

async function fetchSearxSearchResults(baseUrl, query, maxResults = 5, options = {}) {
  const safeBaseUrl = normalizeAiWebSearchSearxBaseUrl(baseUrl);
  if (!safeBaseUrl) {
    const error = new Error("SearXNG 实例地址无效");
    error.code = "SEARX_BASE_URL_INVALID";
    throw error;
  }
  const safeQuery = sanitizeAiWebSearchQuery(query);
  if (!safeQuery) return [];
  const safeTimeRange = sanitizeAiWebSearchTerm(options.timeRange || "");
  const timeRange = ["day", "month", "year"].includes(safeTimeRange) ? safeTimeRange : "";
  const buildSearchUrl = (formatValue) => {
    const parsed = new URL(safeBaseUrl);
    parsed.pathname = `${String(parsed.pathname || "").replace(/\/+$/g, "")}/search`;
    parsed.search = "";
    parsed.searchParams.set("q", safeQuery);
    parsed.searchParams.set("language", "zh-CN");
    parsed.searchParams.set("safesearch", "0");
    parsed.searchParams.set("format", formatValue);
    if (timeRange) parsed.searchParams.set("time_range", timeRange);
    return parsed.toString();
  };

  let lastError = null;
  try {
    const jsonResult = await fetchAiJson(buildSearchUrl("json"), {
      method: "GET",
      headers: { "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)" },
      timeoutMs: AI_WEB_SEARCH_ENGINE_TIMEOUT_MS,
    });
    if (jsonResult.ok) {
      const parsed = jsonResult.data && typeof jsonResult.data === "object" && !("rawText" in jsonResult.data) ? jsonResult.data : {};
      const jsonItems = parseAiWebSearchJsonResults(parsed, maxResults);
      if (jsonItems.length) return jsonItems;
    } else {
      const jsonError = new Error(`SearXNG JSON 搜索失败（${jsonResult.status}）`);
      jsonError.code = "SEARX_JSON_FAILED";
      jsonError.status = jsonResult.status;
      lastError = jsonError;
    }
  } catch (error) {
    lastError = error;
  }

  try {
    const rssResult = await fetchAiJson(buildSearchUrl("rss"), {
      method: "GET",
      headers: {
        "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)",
        Accept: "application/rss+xml,application/xml,text/xml;q=0.9,*/*;q=0.8",
      },
      timeoutMs: AI_WEB_SEARCH_ENGINE_TIMEOUT_MS,
    });
    if (rssResult.ok) {
      const rssItems = parseBingRssResults(rssResult.text, maxResults);
      if (rssItems.length) return rssItems;
    } else {
      const rssError = new Error(`SearXNG RSS 搜索失败（${rssResult.status}）`);
      rssError.code = "SEARX_RSS_FAILED";
      rssError.status = rssResult.status;
      lastError = rssError;
    }
  } catch (error) {
    lastError = error;
  }

  if (lastError) throw lastError;
  return [];
}

function buildAiWebSearchFocusedQuery(query, requiredTerms = []) {
  let focused = sanitizeAiWebSearchQuery(query);
  if (!focused) return "";
  for (const term of requiredTerms) {
    const safe = sanitizeAiWebSearchTerm(term);
    if (!safe) continue;
    focused = focused.replace(new RegExp(escapeAiWebSearchRegExp(safe), "gi"), " ");
  }
  focused = stripAiWebSearchStopWords(focused);
  focused = focused.replace(/[^\u4e00-\u9fffA-Za-z0-9]+/g, " ");
  return focused.replace(/\s+/g, " ").trim();
}

function dedupeAiWebSearchTerms(inputTerms, maxCount = 8) {
  const output = [];
  const seen = new Set();
  for (const rawTerm of Array.isArray(inputTerms) ? inputTerms : []) {
    const cleaned = stripAiWebSearchStopWords(rawTerm);
    const term = normalizeAiWebSearchIntentToken(cleaned);
    if (!term || term.length < 2) continue;
    if (AI_WEB_SEARCH_STOP_WORDS.has(term)) continue;
    if (seen.has(term)) continue;
    seen.add(term);
    output.push(term);
    if (output.length >= maxCount) break;
  }
  return output;
}

function buildAiWebSearchPlan(query) {
  const safeQuery = sanitizeAiWebSearchQuery(query);
  const onThisDayIntent = /(历史上的今天|on\s+this\s+day)/i.test(safeQuery);
  const mapIntent = isLikelyAiMapQuery(safeQuery);
  const dateIntent = buildAiWebSearchDateIntent(safeQuery);
  const requiredTerms = extractAiWebSearchQuotedTerms(safeQuery);
  const focusedQuery = buildAiWebSearchFocusedQuery(safeQuery, requiredTerms);
  const focusedTokens = dedupeAiWebSearchTerms(tokenizeAiWebSearchQuery(focusedQuery), 8);
  const rawTokens = dedupeAiWebSearchTerms(tokenizeAiWebSearchQuery(safeQuery).filter((token) => token.length <= 10), 8);
  const dateTokens = dedupeAiWebSearchTerms(dateIntent.absoluteDateTerms, 6);
  const optionalTerms = dedupeAiWebSearchTerms(
    [...focusedTokens, ...rawTokens, ...dateTokens].filter((token) => !requiredTerms.includes(token)),
    10
  );

  const candidateQueries = [];
  const seen = new Set();
  const addCandidate = (rawValue) => {
    const safe = sanitizeAiWebSearchQuery(rawValue);
    if (!safe || seen.has(safe)) return;
    seen.add(safe);
    candidateQueries.push(safe);
  };

  const intentTerms = optionalTerms.slice(0, 3);
  if (requiredTerms.length && intentTerms.length) {
    addCandidate(`${requiredTerms.join(" ")} ${intentTerms.join(" ")}`);
  }
  if (requiredTerms.length) addCandidate(requiredTerms.join(" "));
  const hasExplicitSearchOperator = /\b(site|inurl|intitle|filetype):/i.test(safeQuery);
  const isOpenAiApiQuery = !hasExplicitSearchOperator && /\bopenai\b/i.test(safeQuery) && /(api|models?|模型)/i.test(safeQuery);
  if (isOpenAiApiQuery) {
    addCandidate("openai api models latest");
    addCandidate("openai api models list");
    addCandidate("site:openai.com api models");
    addCandidate("site:platform.openai.com docs models");
  }
  const dateEnhancedQueries = buildAiWebSearchDateEnhancedQueries(safeQuery, dateIntent);
  for (const queryVariant of dateEnhancedQueries) {
    addCandidate(queryVariant);
  }
  if (mapIntent) {
    const routeIntent = extractAiMapRouteEndpointsFromQuery(safeQuery);
    if (routeIntent) {
      addCandidate(`${routeIntent.origin} 到 ${routeIntent.destination} 路线`);
      addCandidate(`${routeIntent.origin} 到 ${routeIntent.destination} 距离`);
      addCandidate(`${routeIntent.origin} 到 ${routeIntent.destination} 导航`);
    } else {
      const locationIntent = extractAiMapLocationFromQuery(safeQuery);
      if (locationIntent) {
        addCandidate(`${locationIntent} 地图`);
        addCandidate(`${locationIntent} 坐标`);
      }
    }
  }
  if (intentTerms.length) addCandidate(intentTerms.join(" "));
  addCandidate(safeQuery);
  if (!candidateQueries.length && safeQuery) addCandidate(safeQuery);

  return {
    query: safeQuery,
    onThisDayIntent,
    mapIntent,
    requiredTerms,
    optionalTerms,
    dateIntent,
    candidateQueries: candidateQueries.slice(0, 6),
  };
}

function sanitizeAiWebSearchItem(rawItem) {
  if (!rawItem || typeof rawItem !== "object") return null;
  const title = sanitizeAiWebSearchText(rawItem.title, 220);
  const snippet = sanitizeAiWebSearchText(rawItem.snippet, 520);
  const url = normalizeAiWebSearchResultUrl(rawItem.url).slice(0, 1000);
  if (!/^https?:\/\//i.test(url)) return null;
  if (!title && !snippet) return null;
  return {
    title: title || "未命名结果",
    url,
    snippet,
  };
}

function extractXmlTagValue(blockText, tagName) {
  const text = String(blockText || "");
  const safeTag = String(tagName || "").trim();
  if (!text || !safeTag) return "";
  const match = text.match(new RegExp(`<${safeTag}>([\\s\\S]*?)</${safeTag}>`, "i"));
  return match ? match[1] : "";
}

function parseBingRssResults(xmlText, maxResults = 5) {
  const limit = Math.max(1, Math.min(8, Number(maxResults) || 5));
  const itemBlocks = String(xmlText || "").match(/<item>[\s\S]*?<\/item>/gi) || [];
  const results = [];
  for (const block of itemBlocks) {
    const item = sanitizeAiWebSearchItem({
      title: extractXmlTagValue(block, "title"),
      url: decodeHtmlEntities(extractXmlTagValue(block, "link")),
      snippet: extractXmlTagValue(block, "description"),
    });
    if (!item) continue;
    results.push(item);
    if (results.length >= limit) break;
  }
  return results;
}

async function fetchBingSearchResults(query, maxResults = 5, options = {}) {
  const safeTimeoutMs = Number.isFinite(Number(options?.timeoutMs))
    ? Math.max(2000, Math.min(AI_WEB_SEARCH_TIMEOUT_MS, Math.floor(Number(options.timeoutMs))))
    : AI_WEB_SEARCH_ENGINE_TIMEOUT_MS;
  const safeQuery = sanitizeAiWebSearchQuery(query);
  if (!safeQuery) return [];
  const marketRaw = sanitizeAiWebSearchText(options?.market || "", 16).replace(/[^A-Za-z-]/g, "");
  const market = /^[A-Za-z]{2}-[A-Za-z]{2}$/.test(marketRaw) ? marketRaw : "zh-CN";
  const countryCode = market.split("-")[1] || "CN";
  const forceEnglish = /^en-/i.test(market) || !/[\u4e00-\u9fff]/.test(safeQuery);
  const endpoint = `https://www.bing.com/search?format=rss&mkt=${encodeURIComponent(market)}&setlang=${encodeURIComponent(
    market
  )}&cc=${encodeURIComponent(countryCode)}&ensearch=${forceEnglish ? "1" : "0"}&q=${encodeURIComponent(safeQuery)}`;
  const result = await fetchAiJson(endpoint, {
    method: "GET",
    headers: { "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)" },
    timeoutMs: safeTimeoutMs,
  });
  if (!result.ok) {
    const error = new Error(result?.text || `Bing 搜索失败（${result.status}）`);
    error.code = "BING_SEARCH_FAILED";
    error.status = result.status;
    throw error;
  }
  return parseBingRssResults(result.text, maxResults);
}

function hasBaiduSafetyVerificationPage(htmlText) {
  const html = String(htmlText || "");
  if (!html) return false;
  return /百度安全验证|验证码|网络不给力，请稍后重试/i.test(html) && /(mkdjump|安全验证|security)/i.test(html);
}

function parseBaiduHtmlResults(htmlText, maxResults = 5) {
  const limit = Math.max(1, Math.min(8, Number(maxResults) || 5));
  const html = String(htmlText || "");
  if (!html || hasBaiduSafetyVerificationPage(html)) return [];
  const output = [];
  const seen = new Set();
  const pushResult = (rawTitle, rawUrl, rawSnippet) => {
    const item = sanitizeAiWebSearchItem({
      title: rawTitle,
      url: rawUrl,
      snippet: rawSnippet,
    });
    if (!item) return false;
    const key = item.url.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    output.push(item);
    return output.length >= limit;
  };
  const resultBlocks = html.match(/<div class="result c-container[\s\S]*?(?=<div class="result c-container|$)/gi) || [];
  for (const block of resultBlocks) {
    const matchedMu = block.match(/\bmu=["']([^"']+)["']/i);
    const matchedH3Anchor = block.match(/<h3[^>]*>[\s\S]*?<a[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/i);
    const matchedAbstract = block.match(/<[^>]*class=["'][^"']*c-abstract[^"']*["'][^>]*>([\s\S]*?)<\/[^>]+>/i);
    const matchedFallbackSnippet = block.match(
      /<div[^>]*class=["'][^"']*(?:c-color-text|content-right_|c-line-clamp)[^"']*["'][^>]*>([\s\S]*?)<\/div>/i
    );
    const rawUrl =
      normalizeAiWebSearchResultUrl(matchedMu?.[1] || "") ||
      normalizeAiWebSearchResultUrl(matchedH3Anchor?.[1] || "");
    const rawTitle = matchedH3Anchor?.[2] || "";
    const rawSnippet = matchedAbstract?.[1] || matchedFallbackSnippet?.[1] || "";
    if (pushResult(rawTitle, rawUrl, rawSnippet)) return output;
  }

  // Fallback parser for template variants that do not expose `result c-container`.
  const h3Pattern = /<h3[^>]*>\s*<a[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>\s*<\/h3>/gi;
  let h3Match;
  while ((h3Match = h3Pattern.exec(html)) !== null) {
    const anchorUrl = normalizeAiWebSearchResultUrl(h3Match[1]);
    const title = h3Match[2] || "";
    const snippetWindow = html.slice(h3Match.index, Math.min(html.length, h3Match.index + 2600));
    const matchedSnippet = snippetWindow.match(
      /<[^>]*class=["'][^"']*(?:c-abstract|c-color-text|c-line-clamp)[^"']*["'][^>]*>([\s\S]*?)<\/[^>]+>/i
    );
    const snippet = matchedSnippet?.[1] || "";
    if (pushResult(title, anchorUrl, snippet)) break;
  }

  return output;
}

async function fetchBaiduSearchResults(query, maxResults = 5, options = {}) {
  const safeTimeoutMs = Number.isFinite(Number(options?.timeoutMs))
    ? Math.max(2000, Math.min(AI_WEB_SEARCH_TIMEOUT_MS, Math.floor(Number(options.timeoutMs))))
    : AI_WEB_SEARCH_ENGINE_TIMEOUT_MS;
  const safeQuery = sanitizeAiWebSearchQuery(query);
  if (!safeQuery) return [];
  const requestedCount = Math.max(8, Math.min(20, Number(maxResults) * 2 || 10));
  const endpoint = `https://www.baidu.com/s?wd=${encodeURIComponent(safeQuery)}&ie=utf-8&rn=${requestedCount}`;
  const result = await fetchAiJson(endpoint, {
    method: "GET",
    headers: {
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      Accept: "text/html,application/xhtml+xml",
      "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
      Referer: "https://www.baidu.com/",
    },
    timeoutMs: safeTimeoutMs,
  });
  if (!result.ok) {
    const error = new Error(result?.text || `Baidu 搜索失败（${result.status}）`);
    error.code = "BAIDU_SEARCH_FAILED";
    error.status = result.status;
    throw error;
  }
  if (hasBaiduSafetyVerificationPage(result.text)) {
    const verificationError = new Error("Baidu 触发安全验证，暂不可用");
    verificationError.code = "BAIDU_SEARCH_VERIFICATION_REQUIRED";
    verificationError.status = 503;
    throw verificationError;
  }
  return parseBaiduHtmlResults(result.text, maxResults);
}

function normalizeGoogleHtmlResultUrl(rawUrl) {
  const safeUrl = decodeHtmlEntities(String(rawUrl || "")).trim();
  if (!safeUrl) return "";
  try {
    const parsed = new URL(safeUrl, "https://www.google.com");
    if (parsed.pathname === "/url") {
      const target = parsed.searchParams.get("q") || parsed.searchParams.get("url") || "";
      return normalizeAiWebSearchResultUrl(target);
    }
    if (parsed.pathname === "/search" || parsed.pathname === "/imgres" || parsed.pathname === "/setprefs") {
      return "";
    }
    const normalized = normalizeAiWebSearchResultUrl(parsed.toString());
    if (!normalized) return "";
    const host = parsed.hostname.toLowerCase();
    if ((host === "google.com" || host.endsWith(".google.com")) && !/\/(maps|travel|news|finance)/i.test(parsed.pathname || "")) {
      return "";
    }
    return normalized;
  } catch {
    return "";
  }
}

function hasGoogleBotProtectionPage(htmlText) {
  const html = String(htmlText || "");
  if (!html) return false;
  return /(unusual traffic|detected unusual traffic|sorry\/index|not a robot|captcha)/i.test(html);
}

function parseGoogleHtmlResults(htmlText, maxResults = 5) {
  const limit = Math.max(1, Math.min(8, Number(maxResults) || 5));
  const html = String(htmlText || "");
  if (!html || hasGoogleBotProtectionPage(html)) return [];
  const output = [];
  const seen = new Set();
  const anchorPattern = /<a[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let match;
  while ((match = anchorPattern.exec(html)) !== null) {
    const url = normalizeGoogleHtmlResultUrl(match[1]);
    if (!url) continue;
    const anchorHtml = match[2] || "";
    const titleMatch = anchorHtml.match(/<h3[^>]*>([\s\S]*?)<\/h3>/i);
    if (!titleMatch) continue;
    const snippetWindow = html.slice(match.index, Math.min(html.length, match.index + 3200));
    const snippetMatch = snippetWindow.match(/<div[^>]*class=["'][^"']*(?:VwiC3b|s3v9rd|IsZvec)[^"']*["'][^>]*>([\s\S]*?)<\/div>/i);
    const item = sanitizeAiWebSearchItem({
      title: titleMatch[1],
      url,
      snippet: snippetMatch?.[1] || "",
    });
    if (!item) continue;
    const key = item.url.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    output.push(item);
    if (output.length >= limit) break;
  }
  return output;
}

async function fetchGoogleSearchResults(query, maxResults = 5, options = {}) {
  const safeTimeoutMs = Number.isFinite(Number(options?.timeoutMs))
    ? Math.max(2000, Math.min(AI_WEB_SEARCH_TIMEOUT_MS, Math.floor(Number(options.timeoutMs))))
    : AI_WEB_SEARCH_ENGINE_TIMEOUT_MS;
  const safeQuery = sanitizeAiWebSearchQuery(query);
  if (!safeQuery) return [];
  const requestedCount = Math.max(8, Math.min(20, Number(maxResults) * 2 || 10));
  const language = /[\u4e00-\u9fff]/.test(safeQuery) ? "zh-CN" : "en";
  const endpoint = `https://www.google.com/search?hl=${encodeURIComponent(language)}&num=${requestedCount}&q=${encodeURIComponent(safeQuery)}`;
  const result = await fetchAiJson(endpoint, {
    method: "GET",
    headers: {
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
      Accept: "text/html,application/xhtml+xml",
      "Accept-Language": language === "zh-CN" ? "zh-CN,zh;q=0.9,en;q=0.8" : "en-US,en;q=0.9",
    },
    timeoutMs: safeTimeoutMs,
  });
  if (!result.ok) {
    const error = new Error(result?.text || `Google 搜索失败（${result.status}）`);
    error.code = "GOOGLE_SEARCH_FAILED";
    error.status = result.status;
    throw error;
  }
  if (hasGoogleBotProtectionPage(result.text)) {
    const blockedError = new Error("Google 暂时触发风控，无法返回结果");
    blockedError.code = "GOOGLE_SEARCH_BLOCKED";
    blockedError.status = 503;
    throw blockedError;
  }
  return parseGoogleHtmlResults(result.text, maxResults);
}

function normalizeDuckDuckGoHtmlResultUrl(rawUrl) {
  const safeUrl = decodeHtmlEntities(String(rawUrl || "")).trim();
  if (!safeUrl) return "";
  let candidate = safeUrl;
  if (candidate.startsWith("//")) candidate = `https:${candidate}`;
  if (candidate.startsWith("/")) candidate = `https://html.duckduckgo.com${candidate}`;
  try {
    const parsed = new URL(candidate);
    if (/duckduckgo\.com$/i.test(parsed.hostname) && parsed.pathname === "/l/") {
      const redirectUrl = parsed.searchParams.get("uddg");
      if (redirectUrl) {
        try {
          return decodeURIComponent(redirectUrl);
        } catch {
          return redirectUrl;
        }
      }
    }
    return parsed.toString();
  } catch {
    return candidate;
  }
}

function parseDuckDuckGoHtmlResults(htmlText, maxResults = 5) {
  const limit = Math.max(1, Math.min(8, Number(maxResults) || 5));
  const html = String(htmlText || "");
  if (!html) return [];
  const output = [];
  const seen = new Set();
  const anchorPattern = /<a[^>]*class=["'][^"']*result__a[^"']*["'][^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let match;
  while ((match = anchorPattern.exec(html)) !== null) {
    const url = normalizeDuckDuckGoHtmlResultUrl(match[1]);
    const title = sanitizeAiWebSearchText(match[2], 220);
    const textWindow = html.slice(match.index, Math.min(html.length, match.index + 2200));
    const snippetMatch = textWindow.match(/<(?:a|div)[^>]*class=["'][^"']*result__snippet[^"']*["'][^>]*>([\s\S]*?)<\/(?:a|div)>/i);
    const snippet = sanitizeAiWebSearchText(snippetMatch ? snippetMatch[1] : "", 520);
    const item = sanitizeAiWebSearchItem({ title, url, snippet });
    if (!item) continue;
    const key = item.url.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    output.push(item);
    if (output.length >= limit) break;
  }
  return output;
}

async function fetchDuckDuckGoSearchResults(query, maxResults = 5) {
  const endpoint = `https://html.duckduckgo.com/html/?kl=cn-zh&kp=-2&q=${encodeURIComponent(query)}`;
  const result = await fetchAiJson(endpoint, {
    method: "GET",
    headers: {
      "User-Agent": "Mozilla/5.0 (compatible; 991x-ai-assistant/1.0)",
      Accept: "text/html,application/xhtml+xml",
    },
    timeoutMs: AI_WEB_SEARCH_ENGINE_TIMEOUT_MS,
  });
  if (!result.ok) {
    const error = new Error(result?.text || `DuckDuckGo 搜索失败（${result.status}）`);
    error.code = "DUCKDUCKGO_SEARCH_FAILED";
    error.status = result.status;
    throw error;
  }
  return parseDuckDuckGoHtmlResults(result.text, maxResults);
}

function getAiWebSearchGeneralEngineFetcher(engineId) {
  const safeEngine = sanitizeAiWebSearchEngineId(engineId);
  if (safeEngine === AI_WEB_SEARCH_ENGINE_BING) return fetchBingSearchResults;
  if (safeEngine === AI_WEB_SEARCH_ENGINE_BAIDU) return fetchBaiduSearchResults;
  if (safeEngine === AI_WEB_SEARCH_ENGINE_GOOGLE) return fetchGoogleSearchResults;
  return null;
}

function pruneAiWebSearchEngineHealthCache(nowMs = Date.now()) {
  for (const [engineId, entry] of aiWebSearchEngineHealthCache.entries()) {
    if (!entry || !Number.isFinite(Number(entry.expiresAtMs)) || Number(entry.expiresAtMs) <= nowMs) {
      aiWebSearchEngineHealthCache.delete(engineId);
    }
  }
  while (aiWebSearchEngineHealthCache.size > 24) {
    const firstKey = aiWebSearchEngineHealthCache.keys().next().value;
    if (!firstKey) break;
    aiWebSearchEngineHealthCache.delete(firstKey);
  }
}

function readAiWebSearchEngineHealth(engineId, nowMs = Date.now()) {
  const safeEngine = sanitizeAiWebSearchEngineId(engineId);
  if (!safeEngine) return null;
  const entry = aiWebSearchEngineHealthCache.get(safeEngine);
  if (!entry || !Number.isFinite(Number(entry.expiresAtMs)) || Number(entry.expiresAtMs) <= nowMs) {
    aiWebSearchEngineHealthCache.delete(safeEngine);
    return null;
  }
  aiWebSearchEngineHealthCache.delete(safeEngine);
  aiWebSearchEngineHealthCache.set(safeEngine, entry);
  return {
    engine: safeEngine,
    healthy: entry.healthy === true,
    latencyMs: Number.isFinite(Number(entry.latencyMs)) ? Math.max(1, Math.floor(Number(entry.latencyMs))) : 0,
    checkedAtMs: Number.isFinite(Number(entry.checkedAtMs)) ? Math.floor(Number(entry.checkedAtMs)) : nowMs,
    reason: sanitizeAiWebSearchText(entry.reason || "", 120),
  };
}

function writeAiWebSearchEngineHealth(record, ttlMs = AI_WEB_SEARCH_ENGINE_HEALTH_CACHE_MS) {
  const safeEngine = sanitizeAiWebSearchEngineId(record?.engine);
  if (!safeEngine) return;
  const ttl = Number.isFinite(Number(ttlMs))
    ? Math.max(10000, Math.min(1000 * 60 * 60, Math.floor(Number(ttlMs))))
    : AI_WEB_SEARCH_ENGINE_HEALTH_CACHE_MS;
  const nowMs = Date.now();
  pruneAiWebSearchEngineHealthCache(nowMs);
  aiWebSearchEngineHealthCache.delete(safeEngine);
  aiWebSearchEngineHealthCache.set(safeEngine, {
    healthy: record?.healthy === true,
    latencyMs: Number.isFinite(Number(record?.latencyMs)) ? Math.max(1, Math.floor(Number(record.latencyMs))) : 0,
    checkedAtMs: nowMs,
    reason: sanitizeAiWebSearchText(record?.reason || "", 120),
    expiresAtMs: nowMs + ttl,
  });
  pruneAiWebSearchEngineHealthCache(nowMs);
}

async function probeAiWebSearchEngine(engineId) {
  const safeEngine = sanitizeAiWebSearchEngineId(engineId);
  const fetcher = getAiWebSearchGeneralEngineFetcher(safeEngine);
  if (!safeEngine || typeof fetcher !== "function") {
    return {
      engine: safeEngine,
      healthy: false,
      latencyMs: 0,
      checkedAtMs: Date.now(),
      reason: "UNSUPPORTED_ENGINE",
      fromCache: false,
    };
  }
  const startedAtMs = Date.now();
  const probeQuery = sanitizeAiWebSearchQuery(AI_WEB_SEARCH_ENGINE_PROBE_QUERY) || "openai";
  try {
    const probeResults = await fetcher(probeQuery, 1, {
      timeoutMs: AI_WEB_SEARCH_ENGINE_PROBE_TIMEOUT_MS,
    });
    const latencyMs = Math.max(1, Date.now() - startedAtMs);
    const healthy = Array.isArray(probeResults) && probeResults.length > 0;
    const record = {
      engine: safeEngine,
      healthy,
      latencyMs,
      checkedAtMs: Date.now(),
      reason: healthy ? "" : "NO_RESULTS",
      fromCache: false,
    };
    writeAiWebSearchEngineHealth(record);
    return record;
  } catch (error) {
    const latencyMs = Math.max(1, Date.now() - startedAtMs);
    const reason = sanitizeAiWebSearchText(error?.message || "", 120) || "PROBE_FAILED";
    const record = {
      engine: safeEngine,
      healthy: false,
      latencyMs,
      checkedAtMs: Date.now(),
      reason,
      fromCache: false,
    };
    writeAiWebSearchEngineHealth(record);
    return record;
  }
}

async function resolveAiWebSearchHealthyEngineOrder(configuredEngines, preferredEngine = "") {
  const safeConfigured = dedupeAiWebSearchEngineIds(configuredEngines, 6);
  if (!safeConfigured.length) {
    return { healthyEngines: [], health: [] };
  }
  const nowMs = Date.now();
  const preferred = sanitizeAiWebSearchEngineId(preferredEngine);
  const health = await Promise.all(
    safeConfigured.map(async (engineId) => {
      const cached = readAiWebSearchEngineHealth(engineId, nowMs);
      if (cached) return { ...cached, fromCache: true };
      return probeAiWebSearchEngine(engineId);
    })
  );
  const healthByEngine = new Map(health.map((entry) => [entry.engine, entry]));
  const healthyCandidates = safeConfigured.filter((engineId) => healthByEngine.get(engineId)?.healthy === true);
  const latencyComparator = (leftEngine, rightEngine) => {
    const leftLatency = Number.isFinite(Number(healthByEngine.get(leftEngine)?.latencyMs))
      ? Number(healthByEngine.get(leftEngine).latencyMs)
      : Number.MAX_SAFE_INTEGER;
    const rightLatency = Number.isFinite(Number(healthByEngine.get(rightEngine)?.latencyMs))
      ? Number(healthByEngine.get(rightEngine).latencyMs)
      : Number.MAX_SAFE_INTEGER;
    if (leftLatency !== rightLatency) return leftLatency - rightLatency;
    return leftEngine.localeCompare(rightEngine);
  };
  let healthyEngines = [...healthyCandidates].sort(latencyComparator);
  if (preferred && healthyEngines.includes(preferred)) {
    healthyEngines = [preferred, ...healthyEngines.filter((engineId) => engineId !== preferred)];
  }
  return { healthyEngines, health };
}

function mergeAiWebSearchEngineResults(engineOutcomes, plan = {}, maxResults = 5) {
  const outcomes = Array.isArray(engineOutcomes) ? engineOutcomes : [];
  const mergedRawResults = [];
  const sources = [];
  for (const outcome of outcomes) {
    if (!outcome || typeof outcome !== "object") continue;
    const source = sanitizeAiWebSearchTerm(outcome.source || "");
    const results = Array.isArray(outcome.results) ? outcome.results : [];
    if (!results.length) continue;
    if (source) sources.push(source);
    mergedRawResults.push(...results);
  }
  const sourceList = dedupeAiWebSearchEngineIds(sources, 8);
  const filteredResults = filterAiWebSearchResults(mergedRawResults, plan, maxResults);
  const outputResults = filteredResults.length
    ? filteredResults
    : mergedRawResults
        .map((entry) => sanitizeAiWebSearchItem(entry))
        .filter(Boolean)
        .slice(0, Math.max(1, Math.min(8, Number(maxResults) || 5)));
  return {
    sources: sourceList,
    results: outputResults,
  };
}

function buildAiWebSearchComparableText(rawValue) {
  return sanitizeAiWebSearchText(rawValue, 3200)
    .toLowerCase()
    .replace(/\s+/g, " ")
    .trim();
}

function getAiWebSearchHost(rawUrl) {
  const safeUrl = normalizeAiWebSearchResultUrl(rawUrl || "");
  if (!safeUrl) return "";
  try {
    return String(new URL(safeUrl).hostname || "").toLowerCase();
  } catch {
    return "";
  }
}

function countAiWebSearchTermHits(comparableText, terms) {
  if (!comparableText || !Array.isArray(terms) || !terms.length) return 0;
  let hits = 0;
  for (const term of terms) {
    const safeTerm = sanitizeAiWebSearchTerm(term);
    if (!safeTerm) continue;
    if (comparableText.includes(safeTerm)) hits += 1;
  }
  return hits;
}

function extractAiWebSearchDateCandidatesFromText(rawValue, maxCount = 6) {
  const text = sanitizeAiWebSearchText(rawValue, 3200);
  if (!text) return [];
  const output = [];
  const seen = new Set();
  const addDate = (yearValue, monthValue, dayValue) => {
    const date = createAiWebSearchDateFromParts(yearValue, monthValue, dayValue);
    if (!date) return;
    const key = formatAiWebSearchDateKey(date);
    if (!key || seen.has(key)) return;
    seen.add(key);
    output.push(date);
  };
  const fullDatePattern = /(20\d{2})[-/.年](1[0-2]|0?[1-9])[-/.月](3[01]|[12]\d|0?[1-9])(?:日|号)?/g;
  let fullMatch;
  while ((fullMatch = fullDatePattern.exec(text)) !== null) {
    addDate(fullMatch[1], fullMatch[2], fullMatch[3]);
    if (output.length >= maxCount) return output;
  }
  const compactDatePattern = /\b(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])\b/g;
  let compactMatch;
  while ((compactMatch = compactDatePattern.exec(text)) !== null) {
    addDate(compactMatch[1], compactMatch[2], compactMatch[3]);
    if (output.length >= maxCount) return output;
  }
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  if (/(今天|今日|today)/i.test(text)) {
    addDate(today.getFullYear(), today.getMonth() + 1, today.getDate());
  }
  if (/(昨天|昨日|yesterday)/i.test(text)) {
    const yesterday = shiftAiWebSearchDate(today, -1);
    if (yesterday) addDate(yesterday.getFullYear(), yesterday.getMonth() + 1, yesterday.getDate());
  }
  if (/(明天|明日|tomorrow)/i.test(text)) {
    const tomorrow = shiftAiWebSearchDate(today, 1);
    if (tomorrow) addDate(tomorrow.getFullYear(), tomorrow.getMonth() + 1, tomorrow.getDate());
  }
  return output;
}

function scoreAiWebSearchFreshnessSignal(item) {
  const text = `${item?.title || ""} ${item?.snippet || ""} ${item?.url || ""}`;
  const comparable = buildAiWebSearchComparableText(text);
  if (!comparable) return 0;
  let score = 0;
  if (AI_WEB_SEARCH_FRESHNESS_SIGNAL_PATTERN.test(comparable)) score += 8;
  const now = new Date();
  const currentYear = String(now.getFullYear());
  const lastYear = String(now.getFullYear() - 1);
  if (new RegExp(`\\b${escapeAiWebSearchRegExp(currentYear)}\\b`, "i").test(comparable)) score += 10;
  if (new RegExp(`\\b${escapeAiWebSearchRegExp(lastYear)}\\b`, "i").test(comparable)) score += 3;
  const dates = extractAiWebSearchDateCandidatesFromText(comparable, 4);
  if (dates.length) {
    const anchor = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const dayMs = 24 * 60 * 60 * 1000;
    let nearestDays = Number.POSITIVE_INFINITY;
    for (const date of dates) {
      const compared = new Date(date.getFullYear(), date.getMonth(), date.getDate());
      const diffDays = Math.floor((anchor.getTime() - compared.getTime()) / dayMs);
      if (Math.abs(diffDays) < Math.abs(nearestDays)) nearestDays = diffDays;
    }
    const absNearest = Math.abs(nearestDays);
    if (absNearest <= 2) score += 24;
    else if (absNearest <= 7) score += 18;
    else if (absNearest <= 30) score += 12;
    else if (absNearest <= 90) score += 8;
    else if (absNearest <= 365) score += 3;
    else if (absNearest > 730) score -= 10;
    else score -= 4;
  }
  return score;
}

function scoreAiWebSearchHostTermBoost(item, plan = {}) {
  const host = getAiWebSearchHost(item?.url || "");
  if (!host) return 0;
  const requiredTerms = Array.isArray(plan.requiredTerms) ? plan.requiredTerms : [];
  const optionalTerms = Array.isArray(plan.optionalTerms) ? plan.optionalTerms : [];
  const normalizeHostToken = (rawTerm) =>
    sanitizeAiWebSearchTerm(rawTerm)
      .replace(/[^a-z0-9]/g, "")
      .trim();
  let score = 0;
  const seen = new Set();
  for (const rawTerm of requiredTerms) {
    const token = normalizeHostToken(rawTerm);
    if (!token || token.length < 3 || seen.has(token)) continue;
    seen.add(token);
    if (host.includes(token)) score += 8;
  }
  for (const rawTerm of optionalTerms) {
    const token = normalizeHostToken(rawTerm);
    if (!token || token.length < 3 || seen.has(token)) continue;
    seen.add(token);
    if (host.includes(token)) score += 3;
  }
  return Math.min(24, score);
}

function scoreAiWebSearchOfficialDomainBoost(item, plan = {}) {
  const queryText = sanitizeAiWebSearchQuery(plan?.query || "").toLowerCase();
  if (!queryText || !/\bopenai\b/i.test(queryText)) return 0;
  if (!/(api|model|models|模型)/i.test(queryText)) return 0;
  const host = getAiWebSearchHost(item?.url || "");
  if (!host) return 0;
  if (host === "platform.openai.com") return 64;
  if (host === "openai.com" || host.endsWith(".openai.com")) return 54;
  if (host === "help.openai.com" || host.endsWith(".help.openai.com")) return 36;
  if (/^github\.com$/i.test(host) && /github\.com\/openai\//i.test(String(item?.url || ""))) return 20;
  if (/(\.|^)zhihu\.com$/i.test(host) || /(\.|^)csdn\.net$/i.test(host) || /(\.|^)w3cschool\.cn$/i.test(host)) return -10;
  return 0;
}

function isOpenAiApiModelsIntent(rawQuery) {
  const text = sanitizeAiWebSearchQuery(rawQuery).toLowerCase();
  if (!text) return false;
  return /\bopenai\b/.test(text) && /(api|接口)/.test(text) && /(models?|模型|列表|list|latest|最新)/.test(text);
}

function buildAiWebSearchOfficialSeedResults(rawQuery) {
  if (!isOpenAiApiModelsIntent(rawQuery)) return [];
  const candidates = [
    {
      title: "OpenAI API Models",
      url: "https://platform.openai.com/docs/models",
      snippet: "OpenAI 官方模型文档页面，通常包含可用模型说明与能力介绍。",
    },
    {
      title: "OpenAI Models API Reference",
      url: "https://platform.openai.com/docs/api-reference/models/list",
      snippet: "OpenAI 官方 Models API 列表接口文档，可用于程序化查询模型。",
    },
    {
      title: "OpenAI API Pricing",
      url: "https://openai.com/api/pricing/",
      snippet: "OpenAI 官方 API 定价页面，可交叉核对模型系列与成本信息。",
    },
  ];
  return candidates.map((entry) => sanitizeAiWebSearchItem(entry)).filter(Boolean);
}

function scoreAiWebSearchItem(item, plan = {}) {
  const requiredTerms = Array.isArray(plan.requiredTerms) ? plan.requiredTerms : [];
  const optionalTerms = Array.isArray(plan.optionalTerms) ? plan.optionalTerms : [];
  const dateTerms = Array.isArray(plan?.dateIntent?.absoluteDateTerms) ? plan.dateIntent.absoluteDateTerms : [];
  const onThisDayIntent = plan?.onThisDayIntent === true;
  const freshnessIntent = plan?.dateIntent?.freshnessIntent === true;
  const trustedOnThisDayResult = onThisDayIntent && isTrustedOnThisDayUrl(item?.url || "");
  const comparableAll = buildAiWebSearchComparableText(`${item?.title || ""} ${item?.snippet || ""} ${item?.url || ""}`);
  const comparableTitle = buildAiWebSearchComparableText(item?.title || "");
  const requiredHits = countAiWebSearchTermHits(comparableAll, requiredTerms);
  const optionalHits = countAiWebSearchTermHits(comparableAll, optionalTerms);
  const dateHits = countAiWebSearchTermHits(comparableAll, dateTerms);
  const titleRequiredHits = countAiWebSearchTermHits(comparableTitle, requiredTerms);
  const titleOptionalHits = countAiWebSearchTermHits(comparableTitle, optionalTerms);
  const freshnessScore = freshnessIntent ? scoreAiWebSearchFreshnessSignal(item) : 0;
  const hostTermBoost = scoreAiWebSearchHostTermBoost(item, plan);
  const officialDomainBoost = scoreAiWebSearchOfficialDomainBoost(item, plan);

  if (requiredTerms.length && requiredHits === 0) return -1;
  if (onThisDayIntent && dateTerms.length && dateHits === 0) return -1;

  if (!requiredTerms.length && !optionalTerms.length) {
    return 1 + Math.min(6, Math.floor((item?.snippet || "").length / 120)) + freshnessScore + hostTermBoost + officialDomainBoost;
  }

  let score = 0;
  if (!requiredTerms.length && optionalTerms.length && optionalHits === 0) {
    score -= 4;
  }
  score += requiredHits * 30;
  score += optionalHits * 8;
  score += dateHits * 18;
  score += titleRequiredHits * 12;
  score += titleOptionalHits * 4;
  if (trustedOnThisDayResult) score += 36;
  if (dateTerms.length && dateHits > 0) score += 8;
  if (requiredTerms.length && requiredHits >= requiredTerms.length) score += 12;
  if (optionalTerms.length && optionalHits >= Math.min(2, optionalTerms.length)) score += 4;
  score += freshnessScore;
  score += hostTermBoost;
  score += officialDomainBoost;
  if (freshnessIntent && freshnessScore >= 20) score += 6;
  return score;
}

function filterAiWebSearchResults(rawResults, plan = {}, maxResults = 5) {
  const limit = Math.max(1, Math.min(8, Number(maxResults) || 5));
  const input = Array.isArray(rawResults) ? rawResults : [];
  const collectRanked = (scoringPlan) => {
    const ranked = [];
    const seen = new Set();
    for (const rawItem of input) {
      const item = sanitizeAiWebSearchItem(rawItem);
      if (!item) continue;
      const key = item.url.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      const score = scoreAiWebSearchItem(item, scoringPlan);
      if (score < 0) continue;
      ranked.push({ item, score });
    }
    ranked.sort((a, b) => b.score - a.score);
    return ranked.slice(0, limit).map((entry) => entry.item);
  };

  const strict = collectRanked(plan);
  if (strict.length) return strict;
  // Relaxed fallback: keep available network evidence instead of dropping all candidates.
  const relaxedPlan = {
    ...plan,
    requiredTerms: [],
    optionalTerms: [],
  };
  return collectRanked(relaxedPlan);
}

function mergeAiWebSearchWarnings(warnings) {
  const output = [];
  const seen = new Set();
  for (const warning of Array.isArray(warnings) ? warnings : []) {
    const safe = sanitizeAiWebSearchText(warning, 180);
    if (!safe || seen.has(safe)) continue;
    seen.add(safe);
    output.push(safe);
  }
  return sanitizeAiWebSearchText(output.join("；"), 320);
}

function resolveAiWebSearchCachePolicy(rawPayload = {}, plan = {}) {
  const bypassRequested = normalizeAiWebSearchBoolean(
    rawPayload?.bypassCache ?? rawPayload?.cacheBypass ?? rawPayload?.noCache,
    false
  );
  const freshnessIntent = plan?.dateIntent?.freshnessIntent === true;
  const bypassCache = bypassRequested || freshnessIntent;
  const bypassReason = bypassRequested ? "payload-bypass" : freshnessIntent ? "freshness-intent" : "";
  return {
    bypassRequested,
    freshnessIntent,
    bypassCache,
    bypassReason,
  };
}

function collectAiWebSearchAttemptedSources(rawAttempts, maxCount = 20) {
  const output = [];
  const seen = new Set();
  const attempts = Array.isArray(rawAttempts) ? rawAttempts : [];
  const limit = Number.isFinite(Number(maxCount)) ? Math.max(1, Math.floor(Number(maxCount))) : 20;
  for (const attempt of attempts) {
    const source = sanitizeAiWebSearchTerm(attempt?.source || "");
    if (!source || seen.has(source)) continue;
    seen.add(source);
    output.push(source);
    if (output.length >= limit) break;
  }
  return output;
}

function detectAiWebSearchEngineFallbackOccurred(rawSources, rawAttemptedSources, generalEnginePlan = {}) {
  const preferredEngine = sanitizeAiWebSearchEngineId(generalEnginePlan?.preferredEngine || "");
  if (!preferredEngine) return false;
  const usedGeneralEngines = dedupeAiWebSearchEngineIds(rawSources, 8);
  const attemptedGeneralEngines = dedupeAiWebSearchEngineIds(rawAttemptedSources, 8);
  const merged = dedupeAiWebSearchEngineIds([...usedGeneralEngines, ...attemptedGeneralEngines], 8);
  return merged.some((engineId) => engineId !== preferredEngine);
}

function cloneAiWebSearchCacheValue(value) {
  try {
    return JSON.parse(JSON.stringify(value || {}));
  } catch {
    return null;
  }
}

function pruneAiWebSearchResultCache(nowMs = Date.now()) {
  for (const [key, entry] of aiWebSearchResultCache.entries()) {
    if (!entry || !Number.isFinite(Number(entry.expiresAtMs)) || Number(entry.expiresAtMs) <= nowMs) {
      aiWebSearchResultCache.delete(key);
    }
  }
  while (aiWebSearchResultCache.size > AI_WEB_SEARCH_CACHE_MAX_ITEMS) {
    const firstKey = aiWebSearchResultCache.keys().next().value;
    if (!firstKey) break;
    aiWebSearchResultCache.delete(firstKey);
  }
}

function buildAiWebSearchCacheKey(query, maxResults, plan = {}, generalEnginePlan = {}) {
  const safeQuery = sanitizeAiWebSearchQuery(query).toLowerCase();
  const safeLimit = Number.isFinite(Number(maxResults)) ? Math.max(1, Math.min(8, Math.floor(Number(maxResults)))) : 5;
  const dateIntent = plan?.dateIntent && typeof plan.dateIntent === "object" ? plan.dateIntent : {};
  const todayKey = formatAiShanghaiDateTimeParts(new Date()).dateIso;
  const dateAnchor = sanitizeAiWebSearchText(dateIntent.startDateKey || todayKey, 20) || todayKey;
  const relativeKeyword = sanitizeAiWebSearchTerm(dateIntent.relativeKeyword || "");
  const safeRegion = sanitizeAiWebSearchRegionHint(generalEnginePlan?.region);
  const safePreferredEngine =
    sanitizeAiWebSearchEngineId(generalEnginePlan?.preferredEngine) ||
    (sanitizeAiWebSearchTerm(generalEnginePlan?.preferredEngine) === "auto" ? "auto" : "");
  const safeConfiguredEngines = dedupeAiWebSearchEngineIds(generalEnginePlan?.configuredEngines, 6);
  const strictFlag = generalEnginePlan?.strictEngineApplied === true ? "1" : "0";
  const fallbackFlag = generalEnginePlan?.allowEngineFallback === false ? "0" : "1";
  const freshnessFlag = dateIntent?.freshnessIntent === true ? "1" : "0";
  const intentFlags = [
    plan?.onThisDayIntent ? "on-this-day" : "",
    plan?.mapIntent ? "map" : "",
    isLikelyAiWeatherQuery(safeQuery) ? "weather" : "",
  ]
    .filter(Boolean)
    .join(",");
  return [
    safeQuery,
    `n${safeLimit}`,
    `d${dateAnchor}`,
    `r${relativeKeyword || "-"}`,
    `i${intentFlags || "-"}`,
    `g${safeRegion || AI_WEB_SEARCH_REGION_AUTO}`,
    `p${safePreferredEngine || "auto"}`,
    `e${safeConfiguredEngines.join(",") || "-"}`,
    `s${strictFlag}`,
    `f${fallbackFlag}`,
    `h${freshnessFlag}`,
  ].join("|");
}

function readAiWebSearchCache(cacheKey) {
  const safeKey = String(cacheKey || "").trim();
  if (!safeKey) return null;
  const nowMs = Date.now();
  const cached = aiWebSearchResultCache.get(safeKey);
  if (!cached || !Number.isFinite(Number(cached.expiresAtMs)) || Number(cached.expiresAtMs) <= nowMs) {
    aiWebSearchResultCache.delete(safeKey);
    return null;
  }
  aiWebSearchResultCache.delete(safeKey);
  aiWebSearchResultCache.set(safeKey, cached);
  const cloned = cloneAiWebSearchCacheValue(cached.value);
  if (!cloned || typeof cloned !== "object") return null;
  return {
    ...cloned,
    fromCache: true,
  };
}

function writeAiWebSearchCache(cacheKey, payload, ttlMs = AI_WEB_SEARCH_CACHE_TTL_MS) {
  const safeKey = String(cacheKey || "").trim();
  if (!safeKey || !payload || typeof payload !== "object") return;
  const cloned = cloneAiWebSearchCacheValue(payload);
  if (!cloned || typeof cloned !== "object") return;
  const ttl = Number.isFinite(Number(ttlMs)) ? Math.max(10000, Math.floor(Number(ttlMs))) : AI_WEB_SEARCH_CACHE_TTL_MS;
  const nowMs = Date.now();
  pruneAiWebSearchResultCache(nowMs);
  aiWebSearchResultCache.delete(safeKey);
  aiWebSearchResultCache.set(safeKey, {
    expiresAtMs: nowMs + ttl,
    value: cloned,
  });
  pruneAiWebSearchResultCache(nowMs);
}

async function executeAiWebSearchRequest(rawPayload) {
  const query = normalizeAiWebSearchTaskQuery(rawPayload?.query);
  if (!query) {
    const error = new Error("缺少有效搜索关键词");
    error.code = "SEARCH_QUERY_REQUIRED";
    throw error;
  }
  const maxResultsRaw = Number(rawPayload?.maxResults);
  const maxResults = Number.isFinite(maxResultsRaw) ? Math.max(1, Math.min(8, Math.floor(maxResultsRaw))) : 5;
  const plan = buildAiWebSearchPlan(query);
  const officialSeedResults = buildAiWebSearchOfficialSeedResults(query);
  const generalEnginePlan = resolveAiWebSearchGeneralEnginePlan(rawPayload);
  const cachePolicy = resolveAiWebSearchCachePolicy(rawPayload, plan);
  const cacheKey = buildAiWebSearchCacheKey(query, maxResults, plan, generalEnginePlan);
  const sharedMeta = {
    region: generalEnginePlan.region,
    enginePreference: generalEnginePlan.preferredEngine,
    engineOrder: generalEnginePlan.configuredEngines,
    strictEngineRequested: generalEnginePlan.strictEngineRequested === true,
    strictEngineApplied: generalEnginePlan.strictEngineApplied === true,
    allowEngineFallback: generalEnginePlan.allowEngineFallback === true,
    freshnessIntent: cachePolicy.freshnessIntent === true,
    cacheBypass: cachePolicy.bypassCache === true,
    cacheBypassReason: cachePolicy.bypassReason || "",
  };
  if (!cachePolicy.bypassCache) {
    const cachedResponse = readAiWebSearchCache(cacheKey);
    if (cachedResponse) {
      const attemptedSources = collectAiWebSearchAttemptedSources(cachedResponse?.attempts, 24);
      const engineFallbackOccurred = detectAiWebSearchEngineFallbackOccurred(
        cachedResponse?.sources,
        attemptedSources,
        generalEnginePlan
      );
      return {
        ...sharedMeta,
        ...cachedResponse,
        attemptedSources,
        engineFallbackOccurred,
        fromCache: true,
        cacheBypass: false,
        cacheBypassReason: "",
      };
    }
  }
  const attempts = [];
  const warnings = [];
  const finalizeResponse = (payload, options = {}) => {
    const safePayload = payload && typeof payload === "object" ? payload : {};
    const attemptedSources = collectAiWebSearchAttemptedSources(safePayload?.attempts || attempts, 24);
    const engineFallbackOccurred = detectAiWebSearchEngineFallbackOccurred(
      safePayload?.sources,
      attemptedSources,
      generalEnginePlan
    );
    const finalPayload = {
      ...sharedMeta,
      ...safePayload,
      attemptedSources,
      engineFallbackOccurred,
      fromCache: false,
    };
    const ttlMs = Number.isFinite(Number(options.cacheTtlMs)) ? Number(options.cacheTtlMs) : AI_WEB_SEARCH_CACHE_TTL_MS;
    if (cachePolicy.bypassCache !== true && options.writeCache !== false) {
      writeAiWebSearchCache(cacheKey, finalPayload, ttlMs);
    }
    return finalPayload;
  };
  if (plan.onThisDayIntent) {
    try {
      const trustedRawResults = await fetchTrustedOnThisDayResults(plan, maxResults);
      const trustedFiltered = filterAiWebSearchResults(trustedRawResults, plan, maxResults);
      const trustedOutput = trustedFiltered.length ? trustedFiltered : trustedRawResults;
      attempts.push({
        source: "trusted-on-this-day",
        query,
        rawCount: trustedRawResults.length,
        matchedCount: trustedOutput.length,
      });
      if (trustedOutput.length) {
        const dateParts = getAiOnThisDayDateParts(plan.dateIntent);
        const trustedQueryUsed = sanitizeAiWebSearchQuery(`${dateParts.month}月${dateParts.day}日 历史上的今天`) || query;
        return finalizeResponse({
          ...sharedMeta,
          query,
          queryUsed: trustedQueryUsed,
          source: "trusted-on-this-day",
          sources: ["trusted-on-this-day"],
          results: trustedOutput.slice(0, maxResults),
          message: `联网搜索完成，共 ${Math.min(trustedOutput.length, maxResults)} 条结果`,
          warning: mergeAiWebSearchWarnings(warnings),
          attempts,
        });
      }
    } catch (error) {
      const safeWarning =
        sanitizeAiWebSearchText(error?.message || "", 180) || "可信历史站点检索失败，已回退到通用搜索";
      warnings.push(safeWarning);
      attempts.push({
        source: "trusted-on-this-day",
        query,
        rawCount: 0,
        matchedCount: 0,
        error: safeWarning,
      });
    }
  }
  const weatherIntent = isLikelyAiWeatherQuery(query);
  if (weatherIntent) {
    try {
      const weatherRawResults = await fetchOpenMeteoWeatherSearchResults(query, maxResults, plan.dateIntent);
      const weatherFiltered = filterAiWebSearchResults(weatherRawResults, plan, maxResults);
      const weatherOutput = weatherFiltered.length ? weatherFiltered : weatherRawResults;
      attempts.push({
        source: "open-meteo",
        query,
        rawCount: weatherRawResults.length,
        matchedCount: weatherOutput.length,
      });
      if (weatherOutput.length) {
        return finalizeResponse({
          ...sharedMeta,
          query,
          queryUsed: query,
          source: "open-meteo",
          sources: ["open-meteo"],
          results: weatherOutput,
          message: `联网搜索完成，共 ${weatherOutput.length} 条结果`,
          warning: mergeAiWebSearchWarnings(warnings),
          attempts,
        });
      }
      if (weatherRawResults.length) {
        warnings.push("Open-Meteo 返回结果存在，但与关键词匹配度较低");
      }
    } catch (error) {
      const safeWarning =
        sanitizeAiWebSearchText(error?.message || "", 180) || "Open-Meteo 天气查询失败，已回退到通用搜索";
      warnings.push(safeWarning);
      attempts.push({
        source: "open-meteo",
        query,
        rawCount: 0,
        matchedCount: 0,
        error: safeWarning,
      });
    }
  }
  if (plan.mapIntent) {
    try {
      const mapRawResults = await fetchOpenStreetMapSearchResults(query, maxResults);
      const mapFiltered = filterAiWebSearchResults(mapRawResults, plan, maxResults);
      const mapOutput = mapFiltered.length ? mapFiltered : mapRawResults;
      attempts.push({
        source: "openstreetmap",
        query,
        rawCount: mapRawResults.length,
        matchedCount: mapOutput.length,
      });
      if (mapOutput.length) {
        return finalizeResponse({
          ...sharedMeta,
          query,
          queryUsed: query,
          source: "openstreetmap",
          sources: ["openstreetmap"],
          results: mapOutput,
          message: `联网搜索完成，共 ${mapOutput.length} 条结果`,
          warning: mergeAiWebSearchWarnings(warnings),
          attempts,
        });
      }
      if (mapRawResults.length) {
        warnings.push("OpenStreetMap 返回结果存在，但与关键词匹配度较低");
      }
    } catch (error) {
      const safeWarning =
        sanitizeAiWebSearchText(error?.message || "", 180) || "OpenStreetMap 地图查询失败，已回退到通用搜索";
      warnings.push(safeWarning);
      attempts.push({
        source: "openstreetmap",
        query,
        rawCount: 0,
        matchedCount: 0,
        error: safeWarning,
      });
    }
  }

  const searchAndFilter = async (source, searchQuery, fetcher, metadata = {}) => {
    try {
      const rawResults = await fetcher(searchQuery, Math.max(maxResults + 2, maxResults * 2));
      const rankedInput = officialSeedResults.length ? [...rawResults, ...officialSeedResults] : rawResults;
      const filtered = filterAiWebSearchResults(rankedInput, plan, maxResults);
      const attemptRecord = {
        source,
        query: sanitizeAiWebSearchQuery(searchQuery),
        rawCount: rawResults.length,
        matchedCount: filtered.length,
      };
      for (const [key, rawValue] of Object.entries(metadata || {})) {
        const safeValue = sanitizeAiWebSearchText(rawValue, 200);
        if (!safeValue) continue;
        attemptRecord[key] = safeValue;
      }
      attempts.push(attemptRecord);
      if (!filtered.length && rawResults.length) {
        warnings.push(`${source} 返回结果存在，但与关键词匹配度较低`);
      }
      if (officialSeedResults.length && filtered.some((entry) => /(?:^|\.)openai\.com$/i.test(getAiWebSearchHost(entry?.url || "")))) {
        warnings.push("结果已补充官方 OpenAI 来源链接");
      }
      return { results: filtered, failed: false };
    } catch (error) {
      const rawMessage = sanitizeAiWebSearchText(error?.message || "", 180);
      let safeWarning = `${source} 网络请求失败`;
      if (rawMessage) {
        if (rawMessage.toLowerCase() === "fetch failed") {
          safeWarning = `${source} 网络请求失败`;
        } else if (rawMessage.includes("请求超时")) {
          safeWarning = `${source} 请求超时`;
        } else {
          safeWarning = rawMessage;
        }
      }
      warnings.push(safeWarning);
      const failedAttempt = {
        source,
        query: sanitizeAiWebSearchQuery(searchQuery),
        rawCount: 0,
        matchedCount: 0,
        error: safeWarning,
      };
      for (const [key, rawValue] of Object.entries(metadata || {})) {
        const safeValue = sanitizeAiWebSearchText(rawValue, 200);
        if (!safeValue) continue;
        failedAttempt[key] = safeValue;
      }
      attempts.push(failedAttempt);
      if (officialSeedResults.length) {
        const seedFiltered = filterAiWebSearchResults(officialSeedResults, plan, maxResults);
        if (seedFiltered.length) {
          warnings.push(`${source} 检索失败，已返回官方 OpenAI 兜底链接`);
          return { results: seedFiltered, failed: true };
        }
      }
      return { results: [], failed: true };
    }
  };

  const healthResolution = await resolveAiWebSearchHealthyEngineOrder(
    generalEnginePlan.configuredEngines,
    generalEnginePlan.preferredEngine
  );
  const engineHealth = (Array.isArray(healthResolution.health) ? healthResolution.health : []).map((entry) => ({
    engine: sanitizeAiWebSearchEngineId(entry?.engine || ""),
    healthy: entry?.healthy === true,
    latencyMs: Number.isFinite(Number(entry?.latencyMs)) ? Math.max(1, Math.floor(Number(entry.latencyMs))) : 0,
    checkedAtMs: Number.isFinite(Number(entry?.checkedAtMs)) ? Math.floor(Number(entry.checkedAtMs)) : Date.now(),
    fromCache: entry?.fromCache === true,
    reason: sanitizeAiWebSearchText(entry?.reason || "", 120),
  }));
  let activeGeneralEngines = Array.isArray(healthResolution.healthyEngines) && healthResolution.healthyEngines.length
    ? [...healthResolution.healthyEngines]
    : [...generalEnginePlan.configuredEngines];
  if (!activeGeneralEngines.length) {
    warnings.push("未配置可用通用搜索引擎");
  } else if (!healthResolution.healthyEngines?.length) {
    warnings.push("默认搜索引擎健康检查未通过，已尝试直接回退请求");
  }

  for (let queryIndex = 0; queryIndex < plan.candidateQueries.length; queryIndex += 1) {
    const searchQuery = plan.candidateQueries[queryIndex];
    if (!activeGeneralEngines.length) break;
    const failedEngines = new Set();
    const engineOutcomes = [];
    for (const engineId of activeGeneralEngines) {
      const fetcher = getAiWebSearchGeneralEngineFetcher(engineId);
      if (typeof fetcher !== "function") continue;
      const fetchOptions = {};
      if (engineId === AI_WEB_SEARCH_ENGINE_BING && generalEnginePlan.region === AI_WEB_SEARCH_REGION_OUTSIDE_MAINLAND_CHINA) {
        fetchOptions.market = "en-US";
      }
      const outcome = await searchAndFilter(
        engineId,
        searchQuery,
        (safeSearchQuery, resultLimit) => fetcher(safeSearchQuery, resultLimit, fetchOptions),
        fetchOptions
      );
      if (outcome.failed) failedEngines.add(engineId);
      if (outcome.results.length) {
        engineOutcomes.push({
          source: engineId,
          results: outcome.results,
        });
      }
    }

    if (failedEngines.size) {
      activeGeneralEngines = activeGeneralEngines.filter((engineId) => !failedEngines.has(engineId));
      if (!activeGeneralEngines.length) {
        warnings.push("默认搜索引擎均不可用，已停止继续重试");
      }
    }

    const mergedOutcome = mergeAiWebSearchEngineResults(engineOutcomes, plan, maxResults);
    if (mergedOutcome.results.length) {
      const source = mergedOutcome.sources.length > 1 ? "search-fusion" : mergedOutcome.sources[0] || "none";
      return finalizeResponse({
        ...sharedMeta,
        query,
        queryUsed: searchQuery,
        source,
        sources: mergedOutcome.sources,
        results: mergedOutcome.results,
        engineHealth,
        message: `联网搜索完成，共 ${mergedOutcome.results.length} 条结果`,
        warning: mergeAiWebSearchWarnings(warnings),
        attempts,
      });
    }
  }

  warnings.push("未找到与关键词高度匹配的联网结果");
  return finalizeResponse(
    {
      ...sharedMeta,
      query,
      queryUsed: plan.candidateQueries[0] || query,
      source: "none",
      sources: [],
      results: [],
      engineHealth,
      message: "联网搜索未返回有效结果",
      warning: mergeAiWebSearchWarnings(warnings),
      attempts,
    },
    { cacheTtlMs: 60000 }
  );
}

function sanitizeAiAttachmentName(rawValue) {
  return String(rawValue || "").trim().slice(0, 180);
}

function getAiAttachmentFileExtension(fileName) {
  const safeName = sanitizeAiAttachmentName(fileName).toLowerCase();
  const matched = safeName.match(/\.([a-z0-9]{1,8})$/i);
  return matched ? matched[1] : "";
}

function resolveAiAttachmentDocumentExtension(fileName, fileType = "") {
  const extension = getAiAttachmentFileExtension(fileName);
  if (AI_DOCUMENT_EXTENSIONS.has(extension)) return extension;
  const safeType = String(fileType || "")
    .trim()
    .toLowerCase()
    .split(";")[0]
    .trim();
  return AI_DOCUMENT_MIME_TO_EXTENSION.get(safeType) || "";
}

function decodeAiAttachmentBase64ToBuffer(rawBase64) {
  const cleaned = String(rawBase64 || "")
    .trim()
    .replace(/^data:[^;]+;base64,/i, "")
    .replace(/\s+/g, "");
  if (!cleaned) {
    const error = new Error("缺少文件内容");
    error.code = "ATTACHMENT_DATA_REQUIRED";
    throw error;
  }
  const buffer = Buffer.from(cleaned, "base64");
  if (!buffer.length) {
    const error = new Error("文件内容无效");
    error.code = "ATTACHMENT_DATA_INVALID";
    throw error;
  }
  if (Number.isFinite(AI_ATTACHMENT_PARSE_MAX_BYTES) && buffer.length > AI_ATTACHMENT_PARSE_MAX_BYTES) {
    const error = new Error(`文件过大（最大${Math.round(AI_ATTACHMENT_PARSE_MAX_BYTES / 1024 / 1024)}MB）`);
    error.code = "ATTACHMENT_TOO_LARGE";
    throw error;
  }
  return buffer;
}

function normalizeAiAttachmentBinaryBuffer(rawBuffer) {
  let buffer = null;
  if (Buffer.isBuffer(rawBuffer)) {
    buffer = rawBuffer;
  } else if (rawBuffer && ArrayBuffer.isView(rawBuffer)) {
    buffer = Buffer.from(rawBuffer.buffer, rawBuffer.byteOffset, rawBuffer.byteLength);
  } else if (rawBuffer instanceof ArrayBuffer) {
    buffer = Buffer.from(rawBuffer);
  }
  if (!Buffer.isBuffer(buffer)) return null;
  if (!buffer.length) {
    const error = new Error("文件内容无效");
    error.code = "ATTACHMENT_DATA_INVALID";
    throw error;
  }
  if (Number.isFinite(AI_ATTACHMENT_PARSE_MAX_BYTES) && buffer.length > AI_ATTACHMENT_PARSE_MAX_BYTES) {
    const error = new Error(`文件过大（最大${Math.round(AI_ATTACHMENT_PARSE_MAX_BYTES / 1024 / 1024)}MB）`);
    error.code = "ATTACHMENT_TOO_LARGE";
    throw error;
  }
  return buffer;
}

function decodeAiAttachmentPayloadToBuffer(rawPayload) {
  const fromBinary = normalizeAiAttachmentBinaryBuffer(rawPayload?.buffer);
  if (fromBinary) return fromBinary;
  return decodeAiAttachmentBase64ToBuffer(rawPayload?.dataBase64);
}

function sanitizeAiAttachmentExtractedText(rawText, maxLength = AI_ATTACHMENT_PARSE_TEXT_LIMIT) {
  const max = Number.isFinite(Number(maxLength)) ? Math.max(1000, Math.floor(Number(maxLength))) : AI_ATTACHMENT_PARSE_TEXT_LIMIT;
  return decodeHtmlEntities(String(rawText || ""))
    .replace(/\u0000/g, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim()
    .slice(0, max);
}

function getAiAttachmentImageMimeByName(fileName) {
  const ext = String(path.extname(String(fileName || "")).toLowerCase() || "").replace(/^\./, "");
  if (ext === "png") return "image/png";
  if (ext === "jpg" || ext === "jpeg" || ext === "jpe" || ext === "jfif") return "image/jpeg";
  if (ext === "gif") return "image/gif";
  if (ext === "webp") return "image/webp";
  if (ext === "bmp" || ext === "dib") return "image/bmp";
  if (ext === "avif") return "image/avif";
  if (ext === "ico" || ext === "cur") return "image/x-icon";
  if (ext === "svg" || ext === "svgz") return "image/svg+xml";
  if (ext === "tif" || ext === "tiff") return "image/tiff";
  return "";
}

function isAiAttachmentDisplayableImageMime(rawMime) {
  const mime = String(rawMime || "").trim().toLowerCase();
  return [
    "image/png",
    "image/jpeg",
    "image/gif",
    "image/webp",
    "image/bmp",
    "image/svg+xml",
    "image/tiff",
    "image/avif",
    "image/x-icon",
  ].includes(mime);
}

function createAiAttachmentImageCollector() {
  const items = [];
  const seenKeys = new Set();
  let totalBytes = 0;
  const push = (name, mime, buffer, source = "") => {
    if (items.length >= AI_ATTACHMENT_IMAGE_MAX_COUNT) return;
    if (!Buffer.isBuffer(buffer) || !buffer.length) return;
    const safeMime = String(mime || "").trim().toLowerCase();
    if (!isAiAttachmentDisplayableImageMime(safeMime)) return;
    if (buffer.length < 48) return;
    if (buffer.length > AI_ATTACHMENT_IMAGE_MAX_BYTES) return;
    if (totalBytes + buffer.length > AI_ATTACHMENT_IMAGE_TOTAL_MAX_BYTES) return;
    const sourceKey = sanitizeAiAttachmentExtractedText(source || "", 30).toLowerCase();
    const nameKey = sanitizeAiAttachmentExtractedText(name || "", 160).toLowerCase();
    const dedupeKey = `${sourceKey}::${nameKey}::${safeMime}`;
    if (nameKey && seenKeys.has(dedupeKey)) return;
    if (nameKey) seenKeys.add(dedupeKey);
    totalBytes += buffer.length;
    const safeName =
      sanitizeAiAttachmentExtractedText(name || "", 120).replace(/[^\w.\-()\u4e00-\u9fff]/g, "_") ||
      `image-${items.length + 1}.${safeMime.split("/")[1] || "bin"}`;
    items.push({
      name: safeName,
      mime: safeMime,
      bytes: buffer.length,
      source: sanitizeAiAttachmentExtractedText(source, 30) || "",
      dataUrl: `data:${safeMime};base64,${buffer.toString("base64")}`,
    });
  };
  return {
    push,
    list: () => items.slice(),
  };
}

function getZipImageEntryPatternByExtension(extension) {
  const safeExtension = String(extension || "").trim().toLowerCase();
  if (safeExtension === "docx") return /^word\/media\/[^/]+\.[a-z0-9]+$/i;
  if (safeExtension === "xlsx") return /^xl\/media\/[^/]+\.[a-z0-9]+$/i;
  if (safeExtension === "pptx") return /^ppt\/media\/[^/]+\.[a-z0-9]+$/i;
  if (safeExtension === "odt" || safeExtension === "ods" || safeExtension === "odp") return /^Pictures\/[^/]+\.[a-z0-9]+$/i;
  return null;
}

async function extractZipAttachmentImages(buffer, extension) {
  const pattern = getZipImageEntryPatternByExtension(extension);
  if (!pattern) return [];
  let zip = null;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    return [];
  }
  if (!zip) return [];
  const collector = createAiAttachmentImageCollector();
  const entryNames = sortOpenXmlNumberedPath(Object.keys(zip.files).filter((name) => pattern.test(name)));
  for (const entryName of entryNames) {
    if (collector.list().length >= AI_ATTACHMENT_IMAGE_MAX_COUNT) break;
    const entry = zip.files[entryName];
    if (!entry || entry.dir) continue;
    const mime = getAiAttachmentImageMimeByName(entryName);
    if (!mime) continue;
    let content = null;
    try {
      content = await entry.async("nodebuffer");
    } catch {
      content = null;
    }
    if (!Buffer.isBuffer(content) || !content.length) continue;
    collector.push(path.basename(entryName), mime, content, "zip-media");
  }
  return collector.list();
}

function extractBinaryAttachmentImages(buffer) {
  if (!Buffer.isBuffer(buffer) || !buffer.length) return [];
  const collector = createAiAttachmentImageCollector();
  const safeLength = buffer.length;
  const maxScanLength = Math.min(safeLength, 200 * 1024 * 1024);
  const binary = buffer.subarray(0, maxScanLength);
  const pngSignature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  const pngTail = Buffer.from([0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82]);
  const jpegHead = Buffer.from([0xff, 0xd8]);
  const jpegTail = Buffer.from([0xff, 0xd9]);
  const gif87 = Buffer.from("GIF87a");
  const gif89 = Buffer.from("GIF89a");
  const riff = Buffer.from("RIFF");
  const webp = Buffer.from("WEBP");
  const bmpHead = Buffer.from("BM");
  const pushSlice = (start, endExclusive, mime, prefix) => {
    if (!Number.isFinite(start) || !Number.isFinite(endExclusive)) return;
    const begin = Math.max(0, Math.floor(start));
    const end = Math.min(binary.length, Math.floor(endExclusive));
    if (end <= begin) return;
    const size = end - begin;
    if (size < 64 || size > AI_ATTACHMENT_IMAGE_MAX_BYTES) return;
    collector.push(`${prefix}-${begin}.${mime.split("/")[1] || "bin"}`, mime, binary.subarray(begin, end), "binary-scan");
  };

  let cursor = 0;
  while (cursor >= 0 && cursor < binary.length) {
    const start = binary.indexOf(pngSignature, cursor);
    if (start < 0) break;
    const tailStart = binary.indexOf(pngTail, start + pngSignature.length);
    if (tailStart > start) {
      pushSlice(start, tailStart + pngTail.length, "image/png", "png");
      cursor = tailStart + pngTail.length;
    } else {
      cursor = start + pngSignature.length;
    }
    if (collector.list().length >= AI_ATTACHMENT_IMAGE_MAX_COUNT) break;
  }

  cursor = 0;
  while (cursor >= 0 && cursor < binary.length) {
    const start = binary.indexOf(jpegHead, cursor);
    if (start < 0) break;
    const tailStart = binary.indexOf(jpegTail, start + jpegHead.length);
    if (tailStart > start) {
      pushSlice(start, tailStart + jpegTail.length, "image/jpeg", "jpg");
      cursor = tailStart + jpegTail.length;
    } else {
      cursor = start + jpegHead.length;
    }
    if (collector.list().length >= AI_ATTACHMENT_IMAGE_MAX_COUNT) break;
  }

  cursor = 0;
  while (cursor >= 0 && cursor < binary.length) {
    const startRiff = binary.indexOf(riff, cursor);
    if (startRiff < 0 || startRiff + 12 > binary.length) break;
    if (binary.subarray(startRiff + 8, startRiff + 12).equals(webp)) {
      const payloadSize = binary.readUInt32LE(startRiff + 4);
      const totalSize = payloadSize + 8;
      if (totalSize >= 64 && startRiff + totalSize <= binary.length) {
        pushSlice(startRiff, startRiff + totalSize, "image/webp", "webp");
        cursor = startRiff + totalSize;
      } else {
        cursor = startRiff + 4;
      }
    } else {
      cursor = startRiff + 4;
    }
    if (collector.list().length >= AI_ATTACHMENT_IMAGE_MAX_COUNT) break;
  }

  const scanGif = (signature) => {
    let gifCursor = 0;
    while (gifCursor >= 0 && gifCursor < binary.length) {
      const start = binary.indexOf(signature, gifCursor);
      if (start < 0) break;
      const trailer = binary.indexOf(Buffer.from([0x3b]), start + signature.length);
      if (trailer > start + signature.length + 16) {
        pushSlice(start, trailer + 1, "image/gif", "gif");
        gifCursor = trailer + 1;
      } else {
        gifCursor = start + signature.length;
      }
      if (collector.list().length >= AI_ATTACHMENT_IMAGE_MAX_COUNT) break;
    }
  };
  scanGif(gif87);
  if (collector.list().length < AI_ATTACHMENT_IMAGE_MAX_COUNT) scanGif(gif89);

  cursor = 0;
  while (cursor >= 0 && cursor + 6 <= binary.length) {
    const start = binary.indexOf(bmpHead, cursor);
    if (start < 0 || start + 6 > binary.length) break;
    const totalSize = binary.readUInt32LE(start + 2);
    if (totalSize >= 64 && totalSize <= AI_ATTACHMENT_IMAGE_MAX_BYTES && start + totalSize <= binary.length) {
      pushSlice(start, start + totalSize, "image/bmp", "bmp");
      cursor = start + totalSize;
    } else {
      cursor = start + 2;
    }
    if (collector.list().length >= AI_ATTACHMENT_IMAGE_MAX_COUNT) break;
  }

  return collector.list();
}

async function extractAiDocumentImages(buffer, extension) {
  const safeExtension = String(extension || "").trim().toLowerCase();
  if (!safeExtension) return [];
  const openXmlImages = await extractZipAttachmentImages(buffer, safeExtension);
  const output = [];
  const seen = new Set();
  const pushEntry = (entry) => {
    if (!entry || typeof entry !== "object") return;
    const dataUrl = String(entry.dataUrl || "").trim();
    if (!/^data:image\/[a-z0-9.+-]+;base64,/i.test(dataUrl)) return;
    const name = sanitizeAiAttachmentExtractedText(entry.name || "", 120);
    const mime = sanitizeAiAttachmentExtractedText(entry.mime || "", 40);
    const bytes = Number(entry.bytes) || 0;
    const source = sanitizeAiAttachmentExtractedText(entry.source || "", 30);
    const dedupeKey = `${name.toLowerCase()}::${mime.toLowerCase()}::${bytes}`;
    if (name && seen.has(dedupeKey)) return;
    if (name) seen.add(dedupeKey);
    output.push({
      name,
      dataUrl,
      mime,
      bytes,
      source,
    });
  };
  openXmlImages.forEach(pushEntry);

  if (officeParser && typeof officeParser.parseOffice === "function" && ["pptx", "xlsx", "docx", "odp", "ods", "odt"].includes(safeExtension)) {
    try {
      const ast = await officeParser.parseOffice(buffer, {
        ignoreNotes: false,
        putNotesAtLast: false,
        outputErrorToConsole: false,
        extractAttachments: true,
        ocr: false,
        includeRawContent: false,
      });
      const attachmentList = Array.isArray(ast?.attachments) ? ast.attachments : [];
      attachmentList.forEach((attachment, index) => {
        const base64Data = String(attachment?.data || "").trim();
        if (!base64Data) return;
        const mimeType = String(attachment?.mimeType || attachment?.mime || "").trim().toLowerCase();
        const name = String(attachment?.name || attachment?.fileName || `attachment-${index + 1}`).trim();
        const inferredMime = mimeType || getAiAttachmentImageMimeByName(name);
        if (!isAiAttachmentDisplayableImageMime(inferredMime)) return;
        let dataBuffer = null;
        try {
          dataBuffer = Buffer.from(base64Data, "base64");
        } catch {
          dataBuffer = null;
        }
        if (!Buffer.isBuffer(dataBuffer) || !dataBuffer.length) return;
        if (dataBuffer.length > AI_ATTACHMENT_IMAGE_MAX_BYTES) return;
        pushEntry({
          name,
          mime: inferredMime,
          bytes: dataBuffer.length,
          source: "officeparser",
          dataUrl: `data:${inferredMime};base64,${dataBuffer.toString("base64")}`,
        });
      });
    } catch {
      // Ignore officeparser attachment extraction failure and keep other image sources.
    }
  }

  if (["xls", "ppt", "doc"].includes(safeExtension)) {
    const binaryImages = extractBinaryAttachmentImages(buffer);
    binaryImages.forEach(pushEntry);
  }

  return output.slice(0, AI_ATTACHMENT_IMAGE_MAX_COUNT);
}

function isAiDocumentImageExtractionSupported(extension) {
  const safeExtension = String(extension || "").trim().toLowerCase();
  return ["doc", "docx", "xls", "xlsx", "ppt", "pptx", "odt", "ods", "odp"].includes(safeExtension);
}

function isOfficeFallbackSupportedExtension(extension) {
  const safeExtension = String(extension || "").trim().toLowerCase();
  return ["doc", "docx", "xls", "xlsx", "ppt", "pptx", "odt", "ods", "odp"].includes(safeExtension);
}

function isOpenSourceOfficeFallbackSupportedExtension(extension) {
  const safeExtension = String(extension || "").trim().toLowerCase();
  return ["xls", "xlsx", "ppt", "pptx"].includes(safeExtension);
}

function isGenericAttachmentTextCredible(rawText) {
  const text = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
  if (!text) return false;
  const significant = text.replace(/\s+/g, "");
  if (significant.length < 4) return false;
  const usefulChars = (significant.match(/[A-Za-z0-9\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/g) || []).length;
  return usefulChars >= Math.max(3, Math.floor(significant.length * 0.15));
}

function scoreGenericAttachmentText(rawText) {
  const text = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
  if (!text) return Number.NEGATIVE_INFINITY;
  const usefulChars = (text.match(/[A-Za-z0-9\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/g) || []).length;
  const replacementHits = (text.match(/[�﷿]/g) || []).length;
  const lineBreaks = (text.match(/\n/g) || []).length;
  return usefulChars * 2 + Math.min(30, lineBreaks) - replacementHits * 14;
}

function isCredibleAttachmentTextByExtension(extension, rawText) {
  const safeExtension = String(extension || "").trim().toLowerCase();
  if (safeExtension === "xls" || safeExtension === "xlsx" || safeExtension === "ods") {
    return isCredibleSpreadsheetText(rawText);
  }
  if (safeExtension === "ppt" || safeExtension === "pptx" || safeExtension === "odp") {
    return isCrediblePresentationText(rawText);
  }
  return isGenericAttachmentTextCredible(rawText);
}

async function parseOfficeAttachmentByOpenSource(buffer, extension) {
  const safeExtension = String(extension || "").trim().toLowerCase();
  if (!isOpenSourceOfficeFallbackSupportedExtension(safeExtension)) {
    const error = new Error("当前文件类型不支持开源解析兜底");
    error.code = "ATTACHMENT_FALLBACK_UNSUPPORTED";
    throw error;
  }
  const candidates = [];
  const pushCandidate = (rawText, parserLabel) => {
    const text = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
    if (!text) return;
    let score = scoreGenericAttachmentText(text);
    let credible = isGenericAttachmentTextCredible(text);
    if (safeExtension === "xls" || safeExtension === "xlsx") {
      score = scoreSpreadsheetTextQuality(text);
      credible = isCredibleSpreadsheetText(text);
    } else if (safeExtension === "ppt" || safeExtension === "pptx") {
      score = scorePresentationTextQuality(text);
      credible = isCrediblePresentationText(text);
    }
    if (!credible) return;
    candidates.push({
      text,
      parser: sanitizeAiAttachmentExtractedText(parserLabel, 50) || "opensource-office-parser",
      score,
    });
  };

  if (officeParser && typeof officeParser.parseOffice === "function" && safeExtension !== "xls" && safeExtension !== "ppt") {
    try {
      const ast = await officeParser.parseOffice(buffer, {
        ignoreNotes: false,
        putNotesAtLast: false,
        outputErrorToConsole: false,
        extractAttachments: false,
        ocr: false,
        includeRawContent: false,
      });
      const parserText = sanitizeAiAttachmentExtractedText(
        typeof ast?.toText === "function" ? ast.toText() : ast?.text || "",
        AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4
      );
      if (parserText) {
        if (safeExtension === "xlsx") {
          const rows = parseDelimitedRows(parserText, 140, 24);
          if (rows.length) {
            const normalizedRowsText = rows
              .map((row) => row.filter(Boolean).join(" | "))
              .filter(Boolean)
              .join("\n");
            pushCandidate(normalizedRowsText || parserText, "officeparser-xlsx");
          } else {
            pushCandidate(parserText, "officeparser-xlsx");
          }
        } else {
          pushCandidate(parserText, `officeparser-${safeExtension}`);
        }
      }
    } catch {
      // Ignore open-source parser failure and continue with other parsers.
    }
  }

  if (safeExtension === "ppt" && isLikelyCfbCompoundDocument(buffer) && pptToText && typeof pptToText.extractText === "function") {
    try {
      const text = pptToText.extractText(buffer, {
        separator: "\n",
        encoding: "utf8",
      });
      pushCandidate(text, "ppt-to-text-ppt");
    } catch {
      // Ignore ppt parser failure and continue.
    }
  }

  if (!candidates.length) {
    const error = new Error("开源解析兜底未提取到可用文本");
    error.code = "ATTACHMENT_FALLBACK_EMPTY";
    throw error;
  }
  candidates.sort((left, right) => {
    if (left.score !== right.score) return right.score - left.score;
    return String(right.text || "").length - String(left.text || "").length;
  });
  return {
    text: candidates[0].text,
    parser: candidates[0].parser,
  };
}

function getAiOfficeFallbackConvertProfile(extension) {
  const safeExtension = String(extension || "").trim().toLowerCase();
  const isSpreadsheet = safeExtension === "xls" || safeExtension === "xlsx" || safeExtension === "ods";
  if (isSpreadsheet) {
    return {
      convertTo: "csv",
      outputExtensions: [".csv", ".txt"],
      parserSuffix: "csv",
    };
  }
  return {
    convertTo: "txt",
    outputExtensions: [".txt"],
    parserSuffix: "txt",
  };
}

function getAiOfficeFallbackBinaryCandidates() {
  const candidates = [];
  const pushCandidate = (rawPath) => {
    const safePath = String(rawPath || "").trim();
    if (!safePath) return;
    if (!candidates.includes(safePath)) candidates.push(safePath);
  };
  if (AI_LIBREOFFICE_BIN_ENV) {
    AI_LIBREOFFICE_BIN_ENV.split(/[;,]/).forEach((entry) => pushCandidate(entry));
  }
  pushCandidate("soffice");
  pushCandidate("libreoffice");
  pushCandidate("soffice.exe");
  pushCandidate("libreoffice.exe");
  if (process.platform === "win32") {
    const programFiles = String(process.env.ProgramFiles || "C:\\Program Files").trim() || "C:\\Program Files";
    const programFilesX86 = String(process.env["ProgramFiles(x86)"] || "C:\\Program Files (x86)").trim() || "C:\\Program Files (x86)";
    pushCandidate(path.join(programFiles, "LibreOffice", "program", "soffice.exe"));
    pushCandidate(path.join(programFilesX86, "LibreOffice", "program", "soffice.exe"));
  }
  return candidates;
}

async function runAiOfficeFallbackCommand(command, args, options = {}) {
  const timeoutMs = Number.isFinite(Number(options.timeoutMs))
    ? Math.max(1000, Math.floor(Number(options.timeoutMs)))
    : AI_OFFICE_CONVERT_TIMEOUT_MS;
  const cwd = options.cwd || ROOT_DIR;
  return await new Promise((resolve, reject) => {
    let settled = false;
    let timer = null;
    const child = spawn(command, Array.isArray(args) ? args : [], {
      cwd,
      windowsHide: true,
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    const finish = (error, payload) => {
      if (settled) return;
      settled = true;
      if (timer) clearTimeout(timer);
      if (error) {
        reject(error);
      } else {
        resolve(payload);
      }
    };
    child.stdout.on("data", (chunk) => {
      if (!chunk) return;
      stdout += chunk.toString("utf8");
      if (stdout.length > 30000) stdout = stdout.slice(-30000);
    });
    child.stderr.on("data", (chunk) => {
      if (!chunk) return;
      stderr += chunk.toString("utf8");
      if (stderr.length > 30000) stderr = stderr.slice(-30000);
    });
    child.once("error", (error) => finish(error));
    child.once("close", (code, signal) => {
      finish(null, {
        code: Number(code),
        signal: String(signal || ""),
        stdout: sanitizeAiAttachmentExtractedText(stdout, 30000),
        stderr: sanitizeAiAttachmentExtractedText(stderr, 30000),
      });
    });
    timer = setTimeout(() => {
      try {
        child.kill("SIGKILL");
      } catch {
        // Ignore kill failure.
      }
      const timeoutError = new Error(`进程执行超时（>${Math.round(timeoutMs / 1000)}秒）`);
      timeoutError.code = "OFFICE_FALLBACK_TIMEOUT";
      finish(timeoutError);
    }, timeoutMs);
  });
}

async function resolveAiOfficeFallbackBinary() {
  if (!AI_OFFICE_FALLBACK_ENABLED) return "";
  const now = Date.now();
  if (aiOfficeFallbackBinaryCache.expiresAtMs > now) {
    return aiOfficeFallbackBinaryCache.binary || "";
  }
  const candidates = getAiOfficeFallbackBinaryCandidates();
  let lastError = "";
  for (const candidate of candidates) {
    try {
      const probe = await runAiOfficeFallbackCommand(candidate, ["--version"], {
        timeoutMs: Math.min(8000, AI_OFFICE_CONVERT_TIMEOUT_MS),
        cwd: ROOT_DIR,
      });
      const mergedOutput = `${probe.stdout || ""}\n${probe.stderr || ""}`;
      if (probe.code === 0 || /LibreOffice|OpenOffice/i.test(mergedOutput)) {
        aiOfficeFallbackBinaryCache = {
          expiresAtMs: now + AI_OFFICE_FALLBACK_CACHE_MS,
          binary: candidate,
          error: "",
        };
        return candidate;
      }
      lastError = sanitizeAiAttachmentExtractedText(mergedOutput, 220) || `exit=${probe.code}`;
    } catch (error) {
      lastError = sanitizeAiAttachmentExtractedText(error?.message || "", 220) || "命令不可用";
    }
  }
  aiOfficeFallbackBinaryCache = {
    expiresAtMs: now + AI_OFFICE_FALLBACK_CACHE_MS,
    binary: "",
    error: lastError,
  };
  return "";
}

async function parseOfficeAttachmentByLibreOffice(buffer, extension, fileName) {
  const safeExtension = String(extension || "").trim().toLowerCase();
  if (!isOfficeFallbackSupportedExtension(safeExtension)) {
    const error = new Error("当前文件类型不支持Office兜底解析");
    error.code = "ATTACHMENT_FALLBACK_UNSUPPORTED";
    throw error;
  }
  const officeBinary = await resolveAiOfficeFallbackBinary();
  if (!officeBinary) {
    const detail = sanitizeAiAttachmentExtractedText(aiOfficeFallbackBinaryCache.error || "", 120);
    const error = new Error(detail ? `未找到LibreOffice（${detail}）` : "未找到LibreOffice");
    error.code = "ATTACHMENT_FALLBACK_UNAVAILABLE";
    throw error;
  }
  const profile = getAiOfficeFallbackConvertProfile(safeExtension);
  const token = typeof crypto.randomUUID === "function" ? crypto.randomUUID() : crypto.randomBytes(12).toString("hex");
  const safeBaseName =
    sanitizeAiAttachmentExtractedText(path.basename(String(fileName || ""), path.extname(String(fileName || ""))), 50)
      .replace(/[^A-Za-z0-9_\-.]/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_+|_+$/g, "") || "document";
  const tempDir = await fsp.mkdtemp(path.join(os.tmpdir(), "ai-office-"));
  const inputPath = path.join(tempDir, `${safeBaseName}-${token}.${safeExtension || "bin"}`);
  try {
    await fsp.writeFile(inputPath, buffer);
    const convertArgs = [
      "--headless",
      "--invisible",
      "--norestore",
      "--nodefault",
      "--nolockcheck",
      "--nofirststartwizard",
      "--convert-to",
      profile.convertTo,
      "--outdir",
      tempDir,
      inputPath,
    ];
    const conversion = await runAiOfficeFallbackCommand(officeBinary, convertArgs, {
      timeoutMs: AI_OFFICE_CONVERT_TIMEOUT_MS,
      cwd: tempDir,
    });
    if (conversion.code !== 0) {
      const detail = sanitizeAiAttachmentExtractedText(conversion.stderr || conversion.stdout || "", 200);
      const error = new Error(detail ? `LibreOffice转换失败：${detail}` : "LibreOffice转换失败");
      error.code = "ATTACHMENT_FALLBACK_FAILED";
      throw error;
    }
    const entries = await fsp.readdir(tempDir, { withFileTypes: true });
    const outputFiles = entries
      .filter((entry) => entry?.isFile?.())
      .map((entry) => entry.name)
      .filter((name) => {
        if (name === path.basename(inputPath)) return false;
        const ext = path.extname(name).toLowerCase();
        if (profile.outputExtensions.includes(ext)) return true;
        return profile.outputExtensions.length > 1 && [".csv", ".txt", ".tsv"].includes(ext);
      })
      .sort((a, b) => a.localeCompare(b, "en", { numeric: true, sensitivity: "base" }));
    if (!outputFiles.length) {
      const error = new Error("LibreOffice未生成可读取输出文件");
      error.code = "ATTACHMENT_FALLBACK_EMPTY";
      throw error;
    }
    const blocks = [];
    for (const outputName of outputFiles.slice(0, 12)) {
      const outputPath = path.join(tempDir, outputName);
      const raw = await fsp.readFile(outputPath, "utf8");
      const normalized = sanitizeAiAttachmentExtractedText(raw, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
      if (!normalized) continue;
      if (safeExtension === "xls" || safeExtension === "xlsx" || safeExtension === "ods") {
        const rows = parseDelimitedRows(normalized, 120, 24);
        if (rows.length) {
          const tableText = rows
            .map((row) => row.filter(Boolean).join(" | "))
            .filter(Boolean)
            .join("\n");
          const normalizedTableText = sanitizeAiAttachmentExtractedText(tableText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
          if (normalizedTableText) {
            blocks.push(`[工作表 ${sanitizeAiAttachmentExtractedText(path.basename(outputName, path.extname(outputName)), 40) || "N/A"}]`);
            blocks.push(normalizedTableText);
            blocks.push("");
            continue;
          }
        }
      }
      blocks.push(normalized);
      blocks.push("");
    }
    const mergedText = sanitizeAiAttachmentExtractedText(blocks.join("\n"), AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
    if (!mergedText) {
      const error = new Error("LibreOffice未提取到文本");
      error.code = "ATTACHMENT_FALLBACK_EMPTY";
      throw error;
    }
    return {
      text: mergedText,
      parser: `libreoffice-${safeExtension}-${profile.parserSuffix}`,
    };
  } finally {
    await fsp.rm(tempDir, { recursive: true, force: true }).catch(() => {});
  }
}

function sortOpenXmlNumberedPath(paths) {
  return paths.slice().sort((a, b) => {
    const aNum = Number((String(a).match(/(\d+)(?=\.xml$)/i) || [0, 0])[1]);
    const bNum = Number((String(b).match(/(\d+)(?=\.xml$)/i) || [0, 0])[1]);
    if (aNum !== bNum) return aNum - bNum;
    return String(a).localeCompare(String(b), "en", { numeric: true, sensitivity: "base" });
  });
}

function extractWordXmlToText(xmlText) {
  return sanitizeAiAttachmentExtractedText(
    String(xmlText || "")
      .replace(/<w:tab\/>/gi, "\t")
      .replace(/<w:br[^>]*\/>/gi, "\n")
      .replace(/<\/w:p>/gi, "\n")
      .replace(/<\/w:tr>/gi, "\n")
      .replace(/<\/w:tc>/gi, "\t")
      .replace(/<[^>]+>/g, "")
  );
}

function extractPptxXmlToText(xmlText) {
  const xml = String(xmlText || "");
  if (!xml) return "";
  const chunks = [];
  const seen = new Set();
  const pushChunk = (rawText) => {
    const normalized = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
    if (!normalized) return;
    const key = normalized.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    chunks.push(normalized);
  };
  const textTagPatterns = [
    /<a:t[^>]*>([\s\S]*?)<\/a:t>/gi,
    /<a:fld[^>]*>([\s\S]*?)<\/a:fld>/gi,
    /<p:text[^>]*>([\s\S]*?)<\/p:text>/gi,
    /<c:v[^>]*>([\s\S]*?)<\/c:v>/gi,
    /<dc:(?:title|subject|description|creator)[^>]*>([\s\S]*?)<\/dc:[^>]+>/gi,
    /<cp:(?:keywords|category|contentStatus)[^>]*>([\s\S]*?)<\/cp:[^>]+>/gi,
    /<vt:lpwstr[^>]*>([\s\S]*?)<\/vt:lpwstr>/gi,
  ];
  textTagPatterns.forEach((pattern) => {
    let match = pattern.exec(xml);
    while (match) {
      pushChunk(match[1]);
      match = pattern.exec(xml);
    }
  });
  const structuredText = sanitizeAiAttachmentExtractedText(chunks.join("\n"), AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
  if (structuredText) return structuredText;
  return sanitizeAiAttachmentExtractedText(
    xml
      .replace(/<a:tab\/>/gi, "\t")
      .replace(/<a:br[^>]*\/>/gi, "\n")
      .replace(/<\/(?:a:p|a:tr|p:sp|p:txBody|c:pt|c:ser)>/gi, "\n")
      .replace(/<\/a:tc>/gi, "\t")
      .replace(/<[^>]+>/g, " ")
  );
}

function extractOpenDocumentXmlToText(xmlText) {
  return sanitizeAiAttachmentExtractedText(
    String(xmlText || "")
      .replace(/<text:tab\/>/gi, "\t")
      .replace(/<text:line-break\/>/gi, "\n")
      .replace(/<\/text:(?:p|h|list-item)>/gi, "\n")
      .replace(/<\/table:table-row>/gi, "\n")
      .replace(/<\/table:table-cell>/gi, "\t")
      .replace(/<\/draw:page>/gi, "\n\n")
      .replace(/<[^>]+>/g, " ")
  );
}

async function parseOpenDocumentAttachment(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const contentEntry = zip.file("content.xml");
  if (!contentEntry) {
    const error = new Error("OpenDocument 缺少 content.xml");
    error.code = "ATTACHMENT_PARSE_INVALID";
    throw error;
  }
  const xml = await contentEntry.async("string");
  const text = extractOpenDocumentXmlToText(xml);
  if (!text) {
    const error = new Error("OpenDocument 未提取到可用文本");
    error.code = "ATTACHMENT_PARSE_EMPTY";
    throw error;
  }
  return text;
}

async function parseDocxAttachment(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const entryNames = Object.keys(zip.files).filter((name) =>
    /^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/i.test(name)
  );
  const orderedNames = sortOpenXmlNumberedPath(entryNames);
  const blocks = [];
  for (const entryName of orderedNames) {
    const xml = await zip.files[entryName].async("string");
    const text = extractWordXmlToText(xml);
    if (text) blocks.push(text);
  }
  return blocks.join("\n\n");
}

async function parseDocAttachment(buffer) {
  const extractor = new WordExtractor();
  const doc = await extractor.extract(buffer);
  const blocks = [];
  const safePush = (value) => {
    const text = sanitizeAiAttachmentExtractedText(value, AI_ATTACHMENT_PARSE_TEXT_LIMIT);
    if (text) blocks.push(text);
  };
  safePush(doc?.getBody?.());
  safePush(doc?.getFootnotes?.());
  safePush(doc?.getEndnotes?.());
  safePush(doc?.getHeaders?.());
  safePush(doc?.getFooters?.());
  safePush(doc?.getAnnotations?.());
  safePush(doc?.getTextboxes?.());
  return blocks.join("\n\n");
}

function normalizeXlsxCellValue(value) {
  if (value === null || value === undefined) return "";
  return sanitizeAiAttachmentExtractedText(String(value), 200).replace(/\n+/g, " ");
}

function formatXlsxRowsAsMarkdown(rows, maxRows = 60, maxCols = 16) {
  const safeRows = Array.isArray(rows) ? rows.slice(0, Math.max(1, maxRows)) : [];
  if (!safeRows.length) return "(空表)";
  const colCountRaw = safeRows.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0);
  const colCount = Math.min(Math.max(1, colCountRaw), Math.max(1, maxCols));
  const normalizedRows = safeRows.map((row) =>
    Array.from({ length: colCount }, (_, colIndex) => normalizeXlsxCellValue(Array.isArray(row) ? row[colIndex] : ""))
  );
  const header = normalizedRows[0];
  const lines = [];
  lines.push(`| ${header.join(" | ")} |`);
  lines.push(`| ${Array.from({ length: colCount }, () => "---").join(" | ")} |`);
  normalizedRows.slice(1).forEach((row) => {
    lines.push(`| ${row.join(" | ")} |`);
  });
  return lines.join("\n");
}

function parseDelimitedRows(rawText, maxRows = 60, maxCols = 16) {
  const text = String(rawText || "").replace(/\r\n/g, "\n");
  if (!text.trim()) return [];
  const lines = text
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, Math.max(1, maxRows));
  if (!lines.length) return [];
  return lines.map((line) =>
    line
      .split(/\t|,|;|\|/)
      .slice(0, Math.max(1, maxCols))
      .map((cell) => normalizeXlsxCellValue(cell))
  );
}

function buildSpreadsheetTextFromWorkbook(workbook, maxSheets = 6) {
  const sheetNames = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames.slice(0, Math.max(1, maxSheets)) : [];
  const lines = [];
  let nonEmptySheetCount = 0;
  sheetNames.forEach((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return;
    let rows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      raw: false,
      blankrows: false,
      defval: "",
    });
    if (!Array.isArray(rows) || !rows.length) {
      try {
        const csvText = XLSX.utils.sheet_to_csv(sheet, {
          FS: "\t",
          RS: "\n",
          blankrows: false,
        });
        rows = parseDelimitedRows(csvText, 60, 16);
      } catch {
        rows = [];
      }
    }
    if (Array.isArray(rows) && rows.length) nonEmptySheetCount += 1;
    lines.push(`[工作表] ${sanitizeAiAttachmentExtractedText(sheetName, 120) || "未命名"}`);
    lines.push(formatXlsxRowsAsMarkdown(rows, 60, 16));
    lines.push("");
  });
  return {
    text: lines.join("\n"),
    sheetCount: sheetNames.length,
    nonEmptySheetCount,
  };
}

function extractSpreadsheetTextFallback(buffer) {
  const decodeCandidates = [
    { encoding: "utf8", label: "utf8" },
    { encoding: "utf16le", label: "utf16le" },
    { encoding: "latin1", label: "latin1" },
  ];
  let bestText = "";
  for (const candidate of decodeCandidates) {
    let decoded = "";
    try {
      decoded = Buffer.from(buffer || []).toString(candidate.encoding);
    } catch {
      decoded = "";
    }
    if (!decoded) continue;
    let normalized = decoded;
    if (/<table|<tr|<td|<th|<\/table>/i.test(normalized)) {
      normalized = normalized
        .replace(/<\/(tr|p|div|li|h[1-6])>/gi, "\n")
        .replace(/<(td|th)[^>]*>/gi, "\t")
        .replace(/<[^>]+>/g, " ");
    }
    const rows = parseDelimitedRows(normalized, 120, 20);
    const textByRows = rows.length
      ? rows.map((row) => row.filter(Boolean).join(" | ")).filter(Boolean).join("\n")
      : sanitizeAiAttachmentExtractedText(normalized, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
    const safeText = sanitizeAiAttachmentExtractedText(textByRows, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
    if (safeText.length > bestText.length) bestText = safeText;
  }
  return bestText;
}

function scoreSpreadsheetTextQuality(rawText) {
  const text = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
  if (!text) return Number.NEGATIVE_INFINITY;
  const usefulChars = (text.match(/[A-Za-z0-9\u4e00-\u9fff]/g) || []).length;
  const mojibakeHits = (text.match(/[ÃÂÐÑåæçäöü]/g) || []).length;
  const replacementHits = (text.match(/�/g) || []).length;
  const controlHits = (text.match(/[\u0000-\u0008\u000B-\u001F]/g) || []).length;
  const lineBreaks = (text.match(/\n/g) || []).length;
  const score = usefulChars * 2 + Math.min(40, lineBreaks) - mojibakeHits * 6 - replacementHits * 12 - controlHits * 20;
  return score;
}

function isCredibleSpreadsheetText(rawText) {
  const text = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
  if (!text) return false;
  const significant = text.replace(/\s+/g, "");
  if (significant.length < 6) return false;
  const usefulChars = (significant.match(/[A-Za-z0-9\u4e00-\u9fff]/g) || []).length;
  const allowedChars = (
    significant.match(/[A-Za-z0-9\u4e00-\u9fff\u3040-\u30ff.,:;!?(){}\[\]<>+\-*/=_|%￥$#@'"“”‘’、，。；：？！《》【】（）…—]/g) || []
  ).length;
  const suspiciousChars = Math.max(0, significant.length - allowedChars);
  const suspiciousRatio = significant.length ? suspiciousChars / significant.length : 1;
  const replacementHits = (significant.match(/[�﷿]/g) || []).length;
  if (replacementHits >= 2) return false;
  if (suspiciousRatio > 0.2) return false;
  if (usefulChars < Math.max(6, Math.floor(significant.length * 0.2))) return false;
  return true;
}

function parseXlsxAttachment(buffer, extension = "xlsx") {
  const safeExtension = String(extension || "").trim().toLowerCase();
  const readOptionsList = [
    { type: "buffer", raw: false, cellDates: false },
    { type: "buffer", raw: true, cellDates: true, dense: true },
    { type: "buffer", raw: false, cellDates: false, codepage: 936 },
    { type: "buffer", raw: false, cellDates: false, codepage: 65001 },
  ];
  const readErrors = [];
  const candidates = [];
  for (const readOptions of readOptionsList) {
    try {
      const workbook = XLSX.read(buffer, readOptions);
      const built = buildSpreadsheetTextFromWorkbook(workbook, 6);
      const safeText = sanitizeAiAttachmentExtractedText(built?.text || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
      if (safeText && isCredibleSpreadsheetText(safeText)) {
        candidates.push({
          text: safeText,
          parser: `sheetjs-${safeExtension || "xlsx"}`,
          score: scoreSpreadsheetTextQuality(safeText),
        });
      }
    } catch (error) {
      const safeError = sanitizeAiAttachmentExtractedText(error?.message || "", 120);
      if (safeError) readErrors.push(safeError);
    }
  }

  const fallbackText = extractSpreadsheetTextFallback(buffer);
  if (fallbackText && isCredibleSpreadsheetText(fallbackText)) {
    candidates.push({
      text: fallbackText,
      parser: `spreadsheet-fallback-${safeExtension || "xlsx"}`,
      score: scoreSpreadsheetTextQuality(fallbackText),
    });
  }

  if (candidates.length) {
    candidates.sort((left, right) => {
      if (left.score !== right.score) return right.score - left.score;
      return String(right.text || "").length - String(left.text || "").length;
    });
    return {
      text: candidates[0].text,
      parser: candidates[0].parser,
    };
  }

  const error = new Error(readErrors[0] || "表格文件解析失败");
  error.code = "ATTACHMENT_PARSE_FAILED";
  throw error;
}

function normalizePptCandidateText(rawText) {
  const compact = sanitizeAiAttachmentExtractedText(rawText, 1200).replace(/\s+/g, " ").trim().slice(0, 320);
  if (compact.length < 2) return "";
  const significant = compact.replace(/\s+/g, "");
  if (significant.length < 2) return "";
  const hasCjk = /[\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/.test(significant);
  if (!hasCjk && significant.length < 4) return "";
  const usefulChars = (significant.match(/[A-Za-z0-9\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/g) || []).length;
  const replacementHits = (significant.match(/[�﷿]/g) || []).length;
  const mojibakeHits = (significant.match(/[ÃÂÐÑåæçäöü]/g) || []).length;
  const suspiciousChars = (significant.match(/[^A-Za-z0-9\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af.,:;!?(){}\[\]<>+\-*/=_|%￥$#@'"“”‘’、，。；：？！《》【】（）…—]/g) || [])
    .length;
  const suspiciousRatio = significant.length ? suspiciousChars / significant.length : 1;
  if (replacementHits > 0) return "";
  if (mojibakeHits >= 3 && usefulChars < 8) return "";
  if (suspiciousRatio > 0.35) return "";
  if (usefulChars < Math.max(2, Math.floor(significant.length * 0.18))) return "";
  return compact;
}

function scorePresentationTextQuality(rawText) {
  const text = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
  if (!text) return Number.NEGATIVE_INFINITY;
  const significant = text.replace(/\s+/g, "");
  if (!significant) return Number.NEGATIVE_INFINITY;
  const usefulChars = (significant.match(/[A-Za-z0-9\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/g) || []).length;
  const punctuationChars = (
    significant.match(/[.,:;!?(){}\[\]<>+\-*/=_|%￥$#@'"“”‘’、，。；：？！《》【】（）…—]/g) || []
  ).length;
  const replacementHits = (significant.match(/[�﷿]/g) || []).length;
  const controlHits = (significant.match(/[\u0000-\u0008\u000B-\u001F]/g) || []).length;
  const mojibakeHits = (significant.match(/[ÃÂÐÑåæçäöü]/g) || []).length;
  const suspiciousChars = (significant.match(/[^A-Za-z0-9\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af.,:;!?(){}\[\]<>+\-*/=_|%￥$#@'"“”‘’、，。；：？！《》【】（）…—]/g) || [])
    .length;
  const suspiciousRatio = significant.length ? suspiciousChars / significant.length : 1;
  const lineBreaks = (text.match(/\n/g) || []).length;
  let score = usefulChars * 2 + Math.min(60, punctuationChars) + Math.min(40, lineBreaks * 2);
  score -= replacementHits * 16 + controlHits * 20 + mojibakeHits * 8;
  score -= Math.round(Math.max(0, suspiciousRatio - 0.12) * 180);
  if (significant.length >= 20 && usefulChars < Math.floor(significant.length * 0.2)) score -= 120;
  return score;
}

function isCrediblePresentationText(rawText) {
  const text = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 2);
  if (!text) return false;
  const significant = text.replace(/\s+/g, "");
  if (significant.length < 3) return false;
  const usefulChars = (significant.match(/[A-Za-z0-9\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]/g) || []).length;
  const replacementHits = (significant.match(/[�﷿]/g) || []).length;
  const suspiciousChars = (significant.match(/[^A-Za-z0-9\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af.,:;!?(){}\[\]<>+\-*/=_|%￥$#@'"“”‘’、，。；：？！《》【】（）…—]/g) || [])
    .length;
  const suspiciousRatio = significant.length ? suspiciousChars / significant.length : 1;
  if (replacementHits >= 2) return false;
  if (usefulChars < Math.max(2, Math.floor(significant.length * 0.16))) return false;
  if (suspiciousRatio > 0.38) return false;
  return scorePresentationTextQuality(text) > -10;
}

function pickBestPresentationTextCandidate(candidates) {
  const safeCandidates = (Array.isArray(candidates) ? candidates : [])
    .map((candidate) => {
      const text = sanitizeAiAttachmentExtractedText(candidate?.text || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
      const parser = sanitizeAiAttachmentExtractedText(candidate?.parser || "", 60) || "presentation-parser";
      const score = Number.isFinite(Number(candidate?.score))
        ? Number(candidate.score)
        : scorePresentationTextQuality(text);
      return { text, parser, score };
    })
    .filter((candidate) => candidate.text && isCrediblePresentationText(candidate.text));
  if (!safeCandidates.length) return null;
  safeCandidates.sort((left, right) => {
    if (left.score !== right.score) return right.score - left.score;
    return String(right.text || "").length - String(left.text || "").length;
  });
  return safeCandidates[0];
}

function decodeBufferWithEncoding(buffer, encoding) {
  if (!Buffer.isBuffer(buffer) || !buffer.length) return "";
  const safeEncoding = String(encoding || "").trim().toLowerCase();
  try {
    if (safeEncoding === "utf8" || safeEncoding === "utf-8") return buffer.toString("utf8");
    if (safeEncoding === "utf16le" || safeEncoding === "utf-16le") return buffer.toString("utf16le");
    if (safeEncoding === "latin1" || safeEncoding === "binary") return buffer.toString("latin1");
    if (typeof TextDecoder === "function") {
      const decoder = new TextDecoder(safeEncoding, { fatal: false });
      return decoder.decode(buffer);
    }
  } catch {
    // Ignore decode errors and continue with other candidates.
  }
  return "";
}

function decodePptByteTextAtom(atomBuffer, recordType = 0) {
  if (!Buffer.isBuffer(atomBuffer) || atomBuffer.length < 2) return "";
  if (recordType === 4000 || recordType === 4026) {
    const preferred = normalizePptCandidateText(decodeBufferWithEncoding(atomBuffer, "utf-16le"));
    if (preferred) return preferred;
  }
  if (recordType === 4008) {
    const byteOrders = ["utf-8", "gb18030", "windows-1252", "latin1"];
    for (const encoding of byteOrders) {
      const preferred = normalizePptCandidateText(decodeBufferWithEncoding(atomBuffer, encoding));
      if (preferred) return preferred;
    }
  }
  const zeroBytes = atomBuffer.reduce((count, byte) => (byte === 0 ? count + 1 : count), 0);
  const prefersUtf16 = zeroBytes >= Math.floor(atomBuffer.length * 0.14);
  const decodeOrder = prefersUtf16
    ? ["utf-16le", "utf-8", "gb18030", "windows-1252", "latin1"]
    : ["utf-8", "gb18030", "windows-1252", "latin1", "utf-16le"];
  let bestText = "";
  let bestScore = Number.NEGATIVE_INFINITY;
  decodeOrder.forEach((encoding) => {
    const decoded = decodeBufferWithEncoding(atomBuffer, encoding);
    if (!decoded) return;
    const normalized = normalizePptCandidateText(decoded);
    if (!normalized) return;
    const score = scorePresentationTextQuality(normalized);
    if (score > bestScore || (score === bestScore && normalized.length > bestText.length)) {
      bestScore = score;
      bestText = normalized;
    }
  });
  return bestText;
}

function isLikelyPptTextPayload(payload) {
  if (!Buffer.isBuffer(payload) || payload.length < 3) return false;
  let printableCount = 0;
  let zeroCount = 0;
  for (let index = 0; index < payload.length; index += 1) {
    const byte = payload[index];
    if (byte === 0) zeroCount += 1;
    if ((byte >= 32 && byte <= 126) || byte === 9 || byte === 10 || byte === 13 || byte >= 128) printableCount += 1;
  }
  const printableRatio = printableCount / payload.length;
  const zeroRatio = zeroCount / payload.length;
  return printableRatio >= 0.56 || zeroRatio >= 0.18;
}

function extractPptTextAtomsFromRecordStream(streamBuffer, maxCount = 620) {
  if (!Buffer.isBuffer(streamBuffer) || !streamBuffer.length) return [];
  const safeMaxCount = Number.isFinite(Number(maxCount)) ? Math.max(120, Math.floor(Number(maxCount))) : 620;
  const textChunks = [];
  const seen = new Set();
  const pushCandidate = (rawText) => {
    if (textChunks.length >= safeMaxCount) return;
    const normalized = normalizePptCandidateText(rawText);
    if (!normalized) return;
    const key = normalized.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    textChunks.push(normalized);
  };
  const textRecordTypes = new Set([4000, 4008, 4026, 4086, 4097, 4100, 4101]);
  let offset = 0;
  while (offset + 8 <= streamBuffer.length && textChunks.length < safeMaxCount) {
    const verAndInstance = streamBuffer.readUInt16LE(offset);
    const recordType = streamBuffer.readUInt16LE(offset + 2);
    const recordLength = streamBuffer.readUInt32LE(offset + 4);
    const payloadStart = offset + 8;
    const payloadEnd = payloadStart + recordLength;
    if (payloadEnd > streamBuffer.length) break;
    const recordVersion = verAndInstance & 0x000f;
    const isContainer = recordVersion === 0x000f;
    if (!isContainer && recordLength > 1) {
      const payload = streamBuffer.subarray(payloadStart, payloadEnd);
      const shouldDecode = textRecordTypes.has(recordType) || (recordLength <= 1024 && isLikelyPptTextPayload(payload));
      if (shouldDecode) {
        const decoded = decodePptByteTextAtom(payload, recordType);
        if (decoded) pushCandidate(decoded);
      }
    }
    offset = payloadEnd;
  }
  return textChunks;
}

function isLikelyCfbCompoundDocument(buffer) {
  if (!Buffer.isBuffer(buffer) || buffer.length < 8) return false;
  const cfbMagic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  for (let index = 0; index < cfbMagic.length; index += 1) {
    if (buffer[index] !== cfbMagic[index]) return false;
  }
  return true;
}

function extractPptReadableStringsFromBuffer(buffer, maxCount = 420) {
  const safeMaxCount = Number.isFinite(Number(maxCount)) ? Math.max(80, Math.floor(Number(maxCount))) : 420;
  const textChunks = [];
  const seen = new Set();
  const pushCandidate = (rawText) => {
    if (textChunks.length >= safeMaxCount) return;
    const normalized = normalizePptCandidateText(rawText);
    if (!normalized) return;
    const key = normalized.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    textChunks.push(normalized);
  };

  const unicodeMatches =
    buffer.toString("utf16le").match(/[\u4e00-\u9fffA-Za-z0-9][\u4e00-\u9fffA-Za-z0-9 \t,.;:()\-_'"+&@]{1,260}/g) || [];
  unicodeMatches.forEach((entry) => pushCandidate(entry));

  if (textChunks.length < 10) {
    const utf8Matches =
      buffer.toString("utf8").match(/[\u4e00-\u9fffA-Za-z0-9][\u4e00-\u9fffA-Za-z0-9 \t,.;:()\-_'"+&@]{2,260}/g) || [];
    utf8Matches.forEach((entry) => pushCandidate(entry));
  }

  if (textChunks.length < 6) {
    const asciiMatches = buffer.toString("latin1").match(/[A-Za-z][A-Za-z0-9 \t,.;:()\-_'"+&@]{4,260}/g) || [];
    asciiMatches.forEach((entry) => pushCandidate(entry));
  }
  return textChunks;
}

function parsePptAttachment(buffer, options = {}) {
  const includeParser = Boolean(options?.includeParser);
  let sourceBuffer = buffer;
  let parserBase = "binary-ppt-buffer";
  let hasValidCfbDocument = false;
  try {
    const cfb = CFB.read(buffer, { type: "buffer" });
    const pptDocumentStream = (Array.isArray(cfb?.FileIndex) ? cfb.FileIndex : []).find(
      (entry) =>
        entry &&
        entry.type === 2 &&
        /PowerPoint Document/i.test(String(entry.name || "")) &&
        Buffer.isBuffer(entry.content) &&
        entry.content.length > 0
    );
    if (pptDocumentStream?.content?.length) {
      sourceBuffer = Buffer.from(pptDocumentStream.content);
      parserBase = "binary-ppt-cfb";
      hasValidCfbDocument = true;
    }
  } catch {
    // Not a valid CFB stream, fallback to direct buffer scan.
  }
  const candidates = [];
  const pushCandidate = (rawText, parserLabel, scoreBoost = 0) => {
    const text = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
    if (!text) return;
    candidates.push({
      text,
      parser: parserLabel,
      score: scorePresentationTextQuality(text) + Number(scoreBoost || 0),
    });
  };
  const inputLooksLikePpt = hasValidCfbDocument || isLikelyCfbCompoundDocument(buffer);
  if (inputLooksLikePpt && pptToText && typeof pptToText.extractText === "function") {
    try {
      const text = pptToText.extractText(buffer, {
        separator: "\n",
        encoding: "utf8",
      });
      if (text) pushCandidate(text, "ppt-to-text-ppt", 55);
    } catch {
      // Ignore parser errors and continue with binary record extraction.
    }
  }
  const atomChunks = extractPptTextAtomsFromRecordStream(sourceBuffer, 620);
  if (atomChunks.length) pushCandidate(atomChunks.join("\n"), `${parserBase}-records`, 40);
  const allowBinaryScan = inputLooksLikePpt || atomChunks.length >= 2;
  if (allowBinaryScan) {
    const cfbScanChunks = extractPptReadableStringsFromBuffer(sourceBuffer, 520);
    if (cfbScanChunks.length) pushCandidate(cfbScanChunks.join("\n"), `${parserBase}-scan`);
    if (sourceBuffer !== buffer) {
      const directScanChunks = extractPptReadableStringsFromBuffer(buffer, 320);
      if (directScanChunks.length) pushCandidate(directScanChunks.join("\n"), "binary-ppt-buffer-scan");
    }
  }
  const best = pickBestPresentationTextCandidate(candidates);
  if (!best) {
    const error = new Error("PPT 未提取到可用文本");
    error.code = "ATTACHMENT_PARSE_EMPTY";
    throw error;
  }
  if (includeParser) return best;
  return best.text;
}

async function parsePptxAttachment(buffer, options = {}) {
  const includeParser = Boolean(options?.includeParser);
  const parserBase = "openxml-pptx";
  const candidates = [];
  if (officeParser && typeof officeParser.parseOffice === "function") {
    try {
      const ast = await officeParser.parseOffice(buffer, {
        ignoreNotes: false,
        putNotesAtLast: false,
        outputErrorToConsole: false,
        extractAttachments: false,
        ocr: false,
        includeRawContent: false,
      });
      const openSourceText = sanitizeAiAttachmentExtractedText(
        typeof ast?.toText === "function" ? ast.toText() : ast?.text || "",
        AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4
      );
      if (openSourceText && isCrediblePresentationText(openSourceText)) {
        candidates.push({
          text: openSourceText,
          parser: "officeparser-pptx",
          score: scorePresentationTextQuality(openSourceText) + 30,
        });
      }
    } catch {
      // Ignore open-source parser failure and continue with XML parser.
    }
  }
  let zip = null;
  let zipError = null;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch (error) {
    zipError = error;
  }
  if (zip) {
    const entryNames = Object.keys(zip.files);
    const lines = [];
    const consumedEntries = new Set();
    const seenLineKeys = new Set();
    const pushBlock = (label, rawText) => {
      const text = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT);
      if (!text) return;
      const key = `${String(label || "").toLowerCase()}::${text.toLowerCase()}`;
      if (seenLineKeys.has(key)) return;
      seenLineKeys.add(key);
      lines.push(`[${label}]`);
      lines.push(text);
      lines.push("");
    };
    const appendXmlGroup = async (paths, labelPrefix, indexPattern) => {
      const safePaths = Array.isArray(paths) ? paths.slice(0, AI_ATTACHMENT_OPENXML_ENTRY_MAX_COUNT) : [];
      for (const entryPath of safePaths) {
        const entry = zip.files[entryPath];
        if (!entry || entry.dir) continue;
        consumedEntries.add(entryPath);
        const xml = await entry.async("string");
        const text = extractPptxXmlToText(xml);
        if (!text) continue;
        const rawIndex = Number((String(entryPath).match(indexPattern) || [0, 0])[1]);
        const index = Number.isFinite(rawIndex) && rawIndex > 0 ? rawIndex : 0;
        const fallbackName = sanitizeAiAttachmentExtractedText(path.basename(entryPath, ".xml"), 30) || "N/A";
        pushBlock(index ? `${labelPrefix} ${index}` : `${labelPrefix} ${fallbackName}`, text);
      }
    };

    await appendXmlGroup(
      sortOpenXmlNumberedPath(entryNames.filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))),
      "幻灯片",
      /slide(\d+)\.xml$/i
    );
    await appendXmlGroup(
      sortOpenXmlNumberedPath(entryNames.filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(name))),
      "备注",
      /notesSlide(\d+)\.xml$/i
    );
    await appendXmlGroup(
      sortOpenXmlNumberedPath(entryNames.filter((name) => /^ppt\/comments\/comment\d+\.xml$/i.test(name))),
      "评论",
      /comment(\d+)\.xml$/i
    );
    await appendXmlGroup(
      sortOpenXmlNumberedPath(entryNames.filter((name) => /^ppt\/charts\/chart\d+\.xml$/i.test(name))),
      "图表",
      /chart(\d+)\.xml$/i
    );
    await appendXmlGroup(
      sortOpenXmlNumberedPath(entryNames.filter((name) => /^ppt\/diagrams\/data\d+\.xml$/i.test(name))),
      "图示数据",
      /data(\d+)\.xml$/i
    );

    const primaryText = sanitizeAiAttachmentExtractedText(lines.join("\n"), AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
    if (primaryText) {
      candidates.push({
        text: primaryText,
        parser: parserBase,
        score: scorePresentationTextQuality(primaryText),
      });
    }

    if (!primaryText || !isCrediblePresentationText(primaryText)) {
      const fallbackNames = sortOpenXmlNumberedPath(
        entryNames.filter((name) => {
          if (!/^ppt\/.+\.xml$/i.test(name)) return false;
          if (/_rels\//i.test(name)) return false;
          if (/^ppt\/(theme|slideMasters|slideLayouts|notesMasters)\//i.test(name)) return false;
          return true;
        })
      );
      const fallbackLines = [];
      const seenFallbackTexts = new Set();
      for (const entryPath of fallbackNames.slice(0, AI_ATTACHMENT_OPENXML_ENTRY_MAX_COUNT)) {
        if (consumedEntries.has(entryPath)) continue;
        const entry = zip.files[entryPath];
        if (!entry || entry.dir) continue;
        const xml = await entry.async("string");
        const text = extractPptxXmlToText(xml);
        if (!text) continue;
        const key = text.toLowerCase();
        if (seenFallbackTexts.has(key)) continue;
        seenFallbackTexts.add(key);
        const label = sanitizeAiAttachmentExtractedText(path.basename(entryPath, ".xml"), 36) || "xml";
        fallbackLines.push(`[内容 ${label}]`);
        fallbackLines.push(text);
        fallbackLines.push("");
      }
      if (fallbackLines.length) {
        const fallbackText = sanitizeAiAttachmentExtractedText(fallbackLines.join("\n"), AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
        if (fallbackText) {
          candidates.push({
            text: fallbackText,
            parser: `${parserBase}-fullscan`,
            score: scorePresentationTextQuality(fallbackText),
          });
        }
      }
    }
  }

  const best = pickBestPresentationTextCandidate(candidates);
  if (best) {
    if (includeParser) return best;
    return best.text;
  }

  try {
    if (isLikelyCfbCompoundDocument(buffer)) {
      const fallback = parsePptAttachment(buffer, { includeParser: true });
      if (fallback?.text && isCrediblePresentationText(fallback.text)) {
        const parserLabel = `${parserBase}->${sanitizeAiAttachmentExtractedText(fallback.parser || "binary-ppt", 28)}`;
        if (includeParser) return { text: fallback.text, parser: parserLabel };
        return fallback.text;
      }
    }
  } catch {
    // If fallback also fails, report parse-empty below.
  }

  const zipErrorMessage = sanitizeAiAttachmentExtractedText(zipError?.message || "", 120);
  const error = new Error(zipErrorMessage ? `PPTX解析失败：${zipErrorMessage}` : "PPTX 未提取到可用文本");
  error.code = "ATTACHMENT_PARSE_EMPTY";
  throw error;
}

async function parsePdfAttachment(buffer) {
  const parser = new PDFParse({ data: buffer });
  try {
    const parsed = await parser.getText();
    const text = String(parsed?.text || "");
    if (text.trim()) return text;
    if (Array.isArray(parsed?.pages)) {
      const pageText = parsed.pages
        .map((page) => String(page?.text || "").trim())
        .filter(Boolean)
        .join("\n\n");
      if (pageText) return pageText;
    }
    return "";
  } finally {
    if (typeof parser.destroy === "function") {
      await parser.destroy().catch(() => {});
    }
  }
}

function isAiAudioAttachment(fileType, extension) {
  const safeType = String(fileType || "").trim().toLowerCase();
  const safeExt = String(extension || "").trim().toLowerCase();
  return safeType.startsWith("audio/") || AI_AUDIO_EXTENSIONS.has(safeExt);
}

function isAiVideoAttachment(fileType, extension) {
  const safeType = String(fileType || "").trim().toLowerCase();
  const safeExt = String(extension || "").trim().toLowerCase();
  return safeType.startsWith("video/") || AI_VIDEO_EXTENSIONS.has(safeExt);
}

function extractTextFromAudioTranscriptionResponse(data, fallbackRawText = "") {
  if (typeof data?.text === "string" && data.text.trim()) return data.text.trim();
  if (typeof data?.transcript === "string" && data.transcript.trim()) return data.transcript.trim();
  if (typeof data?.output_text === "string" && data.output_text.trim()) return data.output_text.trim();

  if (Array.isArray(data?.segments)) {
    const joined = data.segments
      .map((segment) => String(segment?.text || "").trim())
      .filter(Boolean)
      .join("\n");
    if (joined) return joined;
  }
  if (Array.isArray(data?.results)) {
    const joined = data.results
      .map((entry) => String(entry?.text || entry?.transcript || "").trim())
      .filter(Boolean)
      .join("\n");
    if (joined) return joined;
  }
  const raw = String(fallbackRawText || "").trim();
  if (!raw) return "";
  if (raw.startsWith("{") || raw.startsWith("[")) return "";
  return raw.slice(0, AI_ATTACHMENT_PARSE_TEXT_LIMIT);
}

async function requestAiNetworkAudioTranscription({ buffer, fileName, fileType, transcription }) {
  const config = transcription && typeof transcription === "object" ? transcription : {};
  const network = config.network && typeof config.network === "object" ? config.network : {};
  const baseUrl = normalizeAiHttpUrl(network.baseUrl, "");
  const model = sanitizeAiModelName(config.model || network.audioModel || network.chatModel);
  if (!baseUrl || !model) {
    return {
      parser: "none",
      text: "",
      message: "未配置网络转写，已附加文件",
    };
  }

  const endpoint = buildAiEndpointUrl(baseUrl, normalizeAiHttpPath(network.audioPath, "/audio/transcriptions"));
  if (!endpoint) {
    return {
      parser: "none",
      text: "",
      message: "网络转写地址无效，已附加文件",
    };
  }

  const formData = new FormData();
  const safeType = String(fileType || "").trim() || "application/octet-stream";
  const safeName = sanitizeAiAttachmentName(fileName) || "media.bin";
  formData.set("model", model);
  formData.set("file", new Blob([buffer], { type: safeType }), safeName);

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), AI_ATTACHMENT_TRANSCRIBE_TIMEOUT_MS);
  try {
    const response = await fetch(endpoint, {
      method: "POST",
      headers: buildAiRequestAuthHeaders(network.apiKey),
      body: formData,
      signal: controller.signal,
      cache: "no-store",
    });
    const rawText = await response.text();
    let parsed = null;
    try {
      parsed = rawText ? JSON.parse(rawText) : {};
    } catch {
      parsed = null;
    }
    if (!response.ok) {
      const message =
        parsed?.error?.message ||
        parsed?.error ||
        parsed?.message ||
        rawText ||
        `网络转写失败（${response.status}）`;
      const error = new Error(sanitizeAiAttachmentExtractedText(String(message), 180) || "网络转写失败");
      error.code = "ATTACHMENT_TRANSCRIBE_FAILED";
      throw error;
    }
    const transcriptText = extractTextFromAudioTranscriptionResponse(parsed, rawText);
    return {
      parser: "network-audio-transcribe",
      text: sanitizeAiAttachmentExtractedText(transcriptText, AI_ATTACHMENT_PARSE_TEXT_LIMIT),
      message: transcriptText ? "音视频转写完成" : "转写接口未返回文本，已附加文件",
    };
  } catch (error) {
    return {
      parser: "none",
      text: "",
      message: `转写失败：${sanitizeAiAttachmentExtractedText(error?.message || "未知错误", 140)}`,
    };
  } finally {
    clearTimeout(timer);
  }
}

async function executeAiAttachmentParseRequest(rawPayload) {
  const fileName = sanitizeAiAttachmentName(rawPayload?.name);
  const fileType = String(rawPayload?.type || "").trim().toLowerCase();
  const extension = getAiAttachmentFileExtension(fileName);
  const documentExtension = resolveAiAttachmentDocumentExtension(fileName, fileType);
  const buffer = decodeAiAttachmentPayloadToBuffer(rawPayload);
  const isDocument = Boolean(documentExtension);
  const isAudio = isAiAudioAttachment(fileType, extension);
  const isVideo = isAiVideoAttachment(fileType, extension);

  if (!isDocument && !isAudio && !isVideo) {
    const error = new Error("仅支持解析 pdf/doc/docx/xls/xlsx/ppt/pptx/odt/ods/odp，或音视频转写");
    error.code = "ATTACHMENT_PARSE_UNSUPPORTED";
    throw error;
  }

  if (isAudio || isVideo) {
    const mediaResult = await requestAiNetworkAudioTranscription({
      buffer,
      fileName: fileName || `media.${extension || "bin"}`,
      fileType: fileType || "application/octet-stream",
      transcription: rawPayload?.transcription,
    });
    const transcriptText = sanitizeAiAttachmentExtractedText(mediaResult?.text || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT);
    return {
      name: fileName,
      parser: sanitizeAiAttachmentExtractedText(mediaResult?.parser || "none", 40) || "none",
      extension: extension || (isVideo ? "video" : "audio"),
      text: transcriptText,
      characterCount: transcriptText.length,
      truncated: false,
      message: sanitizeAiAttachmentExtractedText(mediaResult?.message || "媒体已附加", 160) || "媒体已附加",
    };
  }

  let rawText = "";
  let parser = documentExtension || extension;
  let primaryParseError = null;
  try {
    if (documentExtension === "pdf") {
      parser = "pdf-parse";
      rawText = await parsePdfAttachment(buffer);
    } else if (documentExtension === "doc") {
      parser = "word-extractor-doc";
      rawText = await parseDocAttachment(buffer);
    } else if (documentExtension === "docx") {
      parser = "openxml-docx";
      rawText = await parseDocxAttachment(buffer);
    } else if (documentExtension === "xls" || documentExtension === "xlsx") {
      parser = documentExtension === "xls" ? "sheetjs-xls" : "sheetjs-xlsx";
      const spreadsheetResult = parseXlsxAttachment(buffer, documentExtension);
      if (spreadsheetResult && typeof spreadsheetResult === "object") {
        rawText = sanitizeAiAttachmentExtractedText(spreadsheetResult.text || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
        const parserLabel = sanitizeAiAttachmentExtractedText(spreadsheetResult.parser || "", 40);
        if (parserLabel) parser = parserLabel;
      } else {
        rawText = sanitizeAiAttachmentExtractedText(spreadsheetResult || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
      }
    } else if (documentExtension === "ppt") {
      parser = "binary-ppt-cfb";
      const presentationResult = parsePptAttachment(buffer, { includeParser: true });
      if (presentationResult && typeof presentationResult === "object") {
        rawText = sanitizeAiAttachmentExtractedText(presentationResult.text || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
        const parserLabel = sanitizeAiAttachmentExtractedText(presentationResult.parser || "", 40);
        if (parserLabel) parser = parserLabel;
      } else {
        rawText = sanitizeAiAttachmentExtractedText(presentationResult || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
      }
    } else if (documentExtension === "pptx") {
      parser = "openxml-pptx";
      const presentationResult = await parsePptxAttachment(buffer, { includeParser: true });
      if (presentationResult && typeof presentationResult === "object") {
        rawText = sanitizeAiAttachmentExtractedText(presentationResult.text || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
        const parserLabel = sanitizeAiAttachmentExtractedText(presentationResult.parser || "", 40);
        if (parserLabel) parser = parserLabel;
      } else {
        rawText = sanitizeAiAttachmentExtractedText(presentationResult || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
      }
    } else if (documentExtension === "odt" || documentExtension === "ods" || documentExtension === "odp") {
      parser = `opendocument-${documentExtension}`;
      rawText = await parseOpenDocumentAttachment(buffer);
    }
  } catch (error) {
    primaryParseError = error;
  }

  const fallbackErrorMessages = [];
  const primaryTextSnapshot = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
  const hasPrimaryText = Boolean(primaryTextSnapshot);
  const hasCrediblePrimaryText = hasPrimaryText && isCredibleAttachmentTextByExtension(documentExtension, primaryTextSnapshot);
  const canAttemptOpenSourceFallback = isOpenSourceOfficeFallbackSupportedExtension(documentExtension);
  if (!hasCrediblePrimaryText && canAttemptOpenSourceFallback) {
    try {
      const fallbackResult = await parseOfficeAttachmentByOpenSource(buffer, documentExtension);
      if (fallbackResult && typeof fallbackResult === "object") {
        rawText = sanitizeAiAttachmentExtractedText(fallbackResult.text || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
        const parserLabel = sanitizeAiAttachmentExtractedText(fallbackResult.parser || "", 40);
        if (parserLabel) parser = parserLabel;
      } else {
        rawText = sanitizeAiAttachmentExtractedText(fallbackResult || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
      }
    } catch (fallbackError) {
      const fallbackMessage = sanitizeAiAttachmentExtractedText(fallbackError?.message || "", 140);
      if (fallbackMessage) fallbackErrorMessages.push(`开源解析兜底失败：${fallbackMessage}`);
    }
  }

  const openSourceTextSnapshot = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
  const hasOpenSourceFallbackText = Boolean(openSourceTextSnapshot);
  const hasCredibleOpenSourceText =
    hasOpenSourceFallbackText && isCredibleAttachmentTextByExtension(documentExtension, openSourceTextSnapshot);
  const canAttemptLibreOfficeFallback =
    !hasCredibleOpenSourceText && AI_OFFICE_FALLBACK_ENABLED && isOfficeFallbackSupportedExtension(documentExtension);
  if (canAttemptLibreOfficeFallback) {
    try {
      const fallbackResult = await parseOfficeAttachmentByLibreOffice(buffer, documentExtension, fileName);
      if (fallbackResult && typeof fallbackResult === "object") {
        rawText = sanitizeAiAttachmentExtractedText(fallbackResult.text || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
        const parserLabel = sanitizeAiAttachmentExtractedText(fallbackResult.parser || "", 40);
        if (parserLabel) parser = parserLabel;
      } else {
        rawText = sanitizeAiAttachmentExtractedText(fallbackResult || "", AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
      }
    } catch (fallbackError) {
      const fallbackMessage = sanitizeAiAttachmentExtractedText(fallbackError?.message || "", 140);
      if (fallbackMessage) fallbackErrorMessages.push(`LibreOffice兜底失败：${fallbackMessage}`);
    }
  }

  const includeImages = rawPayload?.includeImages !== false;
  let extractedImages = [];
  if (includeImages && isAiDocumentImageExtractionSupported(documentExtension)) {
    try {
      extractedImages = await extractAiDocumentImages(buffer, documentExtension);
    } catch {
      extractedImages = [];
    }
  }

  const rawTextSnapshot = sanitizeAiAttachmentExtractedText(rawText, AI_ATTACHMENT_PARSE_TEXT_LIMIT * 4);
  if (!rawTextSnapshot && !extractedImages.length) {
    const primaryMessage = sanitizeAiAttachmentExtractedText(primaryParseError?.message || "", 140);
    const mergedMessage = [primaryMessage, ...fallbackErrorMessages].filter(Boolean).join("；") || "文档解析失败";
    const wrapped = new Error(mergedMessage);
    wrapped.code = "ATTACHMENT_PARSE_FAILED";
    throw wrapped;
  }

  const normalizedFullText =
    rawTextSnapshot || `文档文本为空，已提取 ${extractedImages.length} 张图片（可在消息中预览）`;
  if (
    rawTextSnapshot &&
    ["xls", "xlsx", "ppt", "pptx"].includes(String(documentExtension || "").toLowerCase()) &&
    !isCredibleAttachmentTextByExtension(documentExtension, normalizedFullText)
  ) {
    const error = new Error("文档提取文本疑似乱码或无效内容");
    error.code = "ATTACHMENT_PARSE_FAILED";
    throw error;
  }
  const truncated = normalizedFullText.length > AI_ATTACHMENT_PARSE_TEXT_LIMIT;
  const text = normalizedFullText.slice(0, AI_ATTACHMENT_PARSE_TEXT_LIMIT);
  const safeImages = extractedImages
    .map((entry) => {
      const dataUrl = String(entry?.dataUrl || "").trim();
      if (!/^data:image\/[a-z0-9.+-]+;base64,/i.test(dataUrl)) return null;
      return {
        name: sanitizeAiAttachmentExtractedText(entry?.name || "", 120),
        dataUrl,
        mime: sanitizeAiAttachmentExtractedText(entry?.mime || "", 40),
        bytes: Number(entry?.bytes) || 0,
      };
    })
    .filter(Boolean);
  const imageHint = safeImages.length ? `，含${safeImages.length}张图片` : "";
  return {
    name: fileName,
    parser,
    extension: documentExtension || extension,
    text,
    characterCount: normalizedFullText.length,
    truncated,
    images: safeImages,
    message: truncated ? `文档解析完成，结果已截断${imageHint}` : `文档解析完成${imageHint}`,
  };
}

function normalizeAiMessageRole(rawRole) {
  const role = String(rawRole || "").trim().toLowerCase();
  if (role === "assistant" || role === "system") return role;
  return "user";
}

function sanitizeAiMessages(rawMessages) {
  if (!Array.isArray(rawMessages)) return [];
  const output = [];
  rawMessages.forEach((entry) => {
    if (!entry || typeof entry !== "object") return;
    const role = normalizeAiMessageRole(entry.role);
    const content = String(entry.content || "").slice(0, 12000).trim();
    if (!content) return;
    output.push({ role, content });
  });
  return output.slice(-20);
}

function sanitizeAiModelName(rawValue) {
  return String(rawValue || "").trim().slice(0, 120);
}

function isLikelyAiImageModelName(modelName) {
  const safe = sanitizeAiModelName(modelName).toLowerCase();
  if (!safe) return false;
  return /(z[-_ ]?image|flash-image|gpt-image|dall[-_ ]?e|sdxl|stable[-_ ]?diffusion|flux|imagen|cogview|image[-_ ]?gen|image[-_ ]?edit|text[-_ ]?to[-_ ]?image)/i.test(
    safe
  );
}

function isLikelyAiReasoningModelName(modelName) {
  const safe = sanitizeAiModelName(modelName).toLowerCase();
  if (!safe) return false;
  return /(deepseek[-_ ]?reasoner|reasoner|reasoning|deepseek[-_ ]?r1|(?:^|[-_.])r1(?:$|[-_.])|o1|o3|qwq|qwq[-_ ]?\d+)/i.test(safe);
}

function resolveAiChatTimeoutMs(provider, modelName) {
  if (String(provider || "").trim().toLowerCase() === "ollama") return AI_OLLAMA_CHAT_TIMEOUT_MS;
  let timeoutMs = AI_REQUEST_TIMEOUT_MS;
  if (isLikelyAiReasoningModelName(modelName)) {
    timeoutMs = Math.max(timeoutMs, AI_REASONING_CHAT_TIMEOUT_MS);
  }
  return timeoutMs;
}

function sanitizeAiSystemPrompt(rawValue) {
  return String(rawValue || "").trim().slice(0, 8000);
}

function padAiDatePart(value) {
  return String(Number(value) || 0).padStart(2, "0");
}

function formatAiShanghaiDateTimeParts(dateValue = new Date()) {
  try {
    const formatter = new Intl.DateTimeFormat("zh-CN", {
      timeZone: "Asia/Shanghai",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
      hour12: false,
      weekday: "short",
    });
    const partMap = {};
    formatter.formatToParts(dateValue).forEach((part) => {
      if (!part || !part.type || part.type === "literal") return;
      partMap[part.type] = part.value;
    });
    const year = String(partMap.year || "").trim();
    const month = String(partMap.month || "").trim();
    const day = String(partMap.day || "").trim();
    const hour = String(partMap.hour || "").trim();
    const minute = String(partMap.minute || "").trim();
    const second = String(partMap.second || "").trim();
    const weekday = String(partMap.weekday || "").trim();
    if (year && month && day && hour && minute && second) {
      return {
        dateIso: `${year}-${month}-${day}`,
        time: `${hour}:${minute}:${second}`,
        weekday: weekday || "",
      };
    }
  } catch {
    // Fall through to local fallback.
  }
  const localDate = new Date(dateValue);
  return {
    dateIso: `${localDate.getFullYear()}-${padAiDatePart(localDate.getMonth() + 1)}-${padAiDatePart(localDate.getDate())}`,
    time: `${padAiDatePart(localDate.getHours())}:${padAiDatePart(localDate.getMinutes())}:${padAiDatePart(localDate.getSeconds())}`,
    weekday: "",
  };
}

function buildAiRuntimeDateGuardPrompt() {
  const now = formatAiShanghaiDateTimeParts(new Date());
  const weekday = now.weekday ? `，${now.weekday}` : "";
  return [
    "时间基准（必须严格遵守）:",
    `- 当前日期（Asia/Shanghai）：${now.dateIso}${weekday}`,
    `- 当前时间（Asia/Shanghai）：${now.time}`,
    "- 若用户提到“今天/昨日/明日/历史上的今天”，必须以上述当前日期换算并输出绝对日期（YYYY-MM-DD）。",
    "- 若用户未主动询问日期/时间，且任务不依赖日期换算，禁止主动在回答中声明“今天是XXXX-XX-XX”。",
    "- 禁止臆测当前日期；不确定时明确说明不确定。",
  ].join("\n");
}

function stripAiDateGuardPrompt(rawText) {
  const safeText = String(rawText || "");
  if (!safeText) return "";
  return safeText
    .replace(/时间基准（必须严格遵守）:[\s\S]*$/g, " ")
    .replace(/当前日期（Asia\/Shanghai）[:：]\s*\d{4}-\d{2}-\d{2}(?:，[^\s]+)?/gi, " ")
    .replace(/当前时间（Asia\/Shanghai）[:：]\s*\d{2}:\d{2}:\d{2}/gi, " ")
    .trim();
}

function shouldInjectAiRuntimeDateGuard({ prompt, messages }) {
  const promptText = stripAiDateGuardPrompt(prompt);
  if (!promptText) return false;
  if (promptText.includes("[联网搜索结果]")) return false;
  const recentUserText = Array.isArray(messages)
    ? messages
        .slice(-6)
        .filter((entry) => normalizeAiMessageRole(entry?.role) === "user")
        .map((entry) => stripAiDateGuardPrompt(entry?.content || ""))
        .join("\n")
    : "";
  const merged = `${recentUserText}\n${promptText}`.toLowerCase();
  const dateIntentPattern =
    /(今天|今日|昨天|昨日|明天|明日|后天|本周|这周|下周|上周|星期几|周几|几号|当前日期|当前时间|现在几点|历史上的今天|today|yesterday|tomorrow|current date|current time|what day is it|what date is it|on this day)/i;
  return dateIntentPattern.test(merged);
}

function buildAiRequestHeaders(apiKey = "") {
  const headers = { "Content-Type": "application/json" };
  const safeKey = String(apiKey || "").trim();
  if (safeKey) headers.Authorization = `Bearer ${safeKey}`;
  return headers;
}

function buildAiRequestAuthHeaders(apiKey = "") {
  const headers = {};
  const safeKey = String(apiKey || "").trim();
  if (safeKey) headers.Authorization = `Bearer ${safeKey}`;
  return headers;
}

function extractTextFromOpenAiResponse(data) {
  const firstChoice = Array.isArray(data?.choices) ? data.choices[0] || null : null;
  if (!firstChoice) return "";
  const message = firstChoice.message || {};
  const content = message.content;
  if (typeof content === "string") return content;
  if (Array.isArray(content)) {
    return content
      .map((part) => {
        if (!part || typeof part !== "object") return "";
        if (typeof part.text === "string") return part.text;
        if (typeof part.output_text === "string") return part.output_text;
        return "";
      })
      .filter(Boolean)
      .join("\n");
  }
  return typeof firstChoice.text === "string" ? firstChoice.text : "";
}

function extractFirstMediaUrlFromPayload(value, mediaType = "image") {
  const extractCandidateFromText = (rawText) => {
    const text = String(rawText || "").trim();
    if (!text) return "";
    if (mediaType !== "video") {
      const inlineDataMatch = text.match(/data:image\/[a-zA-Z0-9.+-]+;base64,[A-Za-z0-9+/=\s]+/i);
      if (inlineDataMatch && inlineDataMatch[0]) {
        return inlineDataMatch[0].replace(/\s+/g, "");
      }
      if (/^data:image\/[a-zA-Z0-9.+-]+;base64,/i.test(text)) return text;
    }

    const urlCandidates = [];
    const markdownImageRegex = /!\[[^\]]*]\((https?:\/\/[^\s)]+)\)/gi;
    let markdownMatch = markdownImageRegex.exec(text);
    while (markdownMatch) {
      if (markdownMatch[1]) urlCandidates.push(markdownMatch[1]);
      markdownMatch = markdownImageRegex.exec(text);
    }
    const genericUrlRegex = /https?:\/\/[^\s"'`<>]+/gi;
    let genericMatch = genericUrlRegex.exec(text);
    while (genericMatch) {
      if (genericMatch[0]) urlCandidates.push(genericMatch[0]);
      genericMatch = genericUrlRegex.exec(text);
    }

    const normalizedCandidates = Array.from(
      new Set(
        urlCandidates
          .map((urlText) => String(urlText || "").trim().replace(/[),.;]+$/g, ""))
          .filter(Boolean)
      )
    );
    if (!normalizedCandidates.length) return "";

    for (const candidate of normalizedCandidates) {
      if (mediaType === "video") {
        if (/\.(mp4|mov|mkv|webm|m3u8)(\?.*)?$/i.test(candidate) || /\/video(s)?\//i.test(candidate)) return candidate;
        continue;
      }
      if (
        /\.(png|jpe?g|webp|gif|bmp|svg)(\?.*)?$/i.test(candidate) ||
        /\/image(s)?\//i.test(candidate) ||
        /[?&](format|fm|ext)=(png|jpe?g|webp|gif|bmp|svg)/i.test(candidate)
      ) {
        return candidate;
      }
    }

    if (mediaType !== "video" && normalizedCandidates.length === 1) {
      return normalizedCandidates[0];
    }
    return "";
  };

  const queue = [value];
  const checked = new Set();
  while (queue.length) {
    const current = queue.shift();
    if (!current || checked.has(current)) continue;
    checked.add(current);
    if (typeof current === "string") {
      const candidate = extractCandidateFromText(current);
      if (candidate) return candidate;
      continue;
    }
    if (Array.isArray(current)) {
      current.forEach((item) => queue.push(item));
      continue;
    }
    if (typeof current === "object") {
      Object.values(current).forEach((item) => queue.push(item));
    }
  }
  return "";
}

function sanitizeAiImageMimeType(rawValue) {
  const safe = String(rawValue || "")
    .trim()
    .toLowerCase();
  if (/^image\/[a-z0-9.+-]+$/.test(safe)) return safe;
  return "image/png";
}

function normalizeAiBase64Payload(rawValue) {
  const safe = String(rawValue || "").trim().replace(/\s+/g, "");
  if (!safe || safe.length < 40) return "";
  if (!/^[A-Za-z0-9+/=]+$/.test(safe)) return "";
  return safe;
}

function buildAiImageDataUrlFromBase64(base64Raw, mimeType = "image/png") {
  const base64 = normalizeAiBase64Payload(base64Raw);
  if (!base64) return "";
  const safeMime = sanitizeAiImageMimeType(mimeType);
  return `data:${safeMime};base64,${base64}`;
}

function extractFirstImageDataUrlFromPayload(value) {
  const queue = [value];
  const checked = new Set();
  while (queue.length) {
    const current = queue.shift();
    if (!current || checked.has(current)) continue;
    checked.add(current);

    if (typeof current === "string") {
      const safe = current.trim();
      if (/^data:image\/[a-zA-Z0-9.+-]+;base64,/i.test(safe)) return safe;
      const inlineDataMatch = safe.match(/data:image\/[a-zA-Z0-9.+-]+;base64,[A-Za-z0-9+/=\s]+/i);
      if (inlineDataMatch && inlineDataMatch[0]) return inlineDataMatch[0].replace(/\s+/g, "");
      if (
        (safe.startsWith("{") || safe.startsWith("[")) &&
        safe.length <= 120000
      ) {
        try {
          const parsed = JSON.parse(safe);
          if (parsed && typeof parsed === "object") queue.push(parsed);
        } catch {
          // ignore non-JSON text
        }
      }
      continue;
    }

    if (Array.isArray(current)) {
      current.forEach((item) => queue.push(item));
      continue;
    }

    if (typeof current === "object") {
      const inlineData =
        current.inline_data && typeof current.inline_data === "object"
          ? current.inline_data
          : current.inlineData && typeof current.inlineData === "object"
            ? current.inlineData
            : null;
      if (inlineData) {
        const inlineUrl = buildAiImageDataUrlFromBase64(
          inlineData.data || inlineData.base64,
          inlineData.mime_type || inlineData.mimeType || "image/png"
        );
        if (inlineUrl) return inlineUrl;
      }

      const directUrl = buildAiImageDataUrlFromBase64(
        current.b64_json || current.base64 || current.image_base64 || current.imageBase64 || current.b64,
        current.mime_type || current.mimeType || "image/png"
      );
      if (directUrl) return directUrl;

      Object.values(current).forEach((item) => queue.push(item));
    }
  }
  return "";
}

function buildAiImageNoMediaErrorMessage(payload, fallbackMessage = "生图请求未返回图片数据") {
  const textHint = sanitizeAiErrorMessage(extractTextFromOpenAiResponse(payload), "");
  if (textHint) {
    return `模型返回了文本而不是图片：${textHint}`;
  }
  const rootData = payload && typeof payload === "object" ? payload : {};
  const firstData = Array.isArray(rootData?.data) && rootData.data[0] && typeof rootData.data[0] === "object" ? rootData.data[0] : {};
  const taskId = sanitizeAiErrorMessage(
    rootData?.task_id || rootData?.taskId || rootData?.request_id || rootData?.requestId || rootData?.id || firstData?.task_id || firstData?.taskId || firstData?.id,
    ""
  );
  const status = sanitizeAiErrorMessage(
    rootData?.status || rootData?.task_status || rootData?.taskStatus || firstData?.status || firstData?.task_status || firstData?.taskStatus,
    ""
  );
  if (taskId || status) {
    const parts = [];
    if (taskId) parts.push(`taskId=${taskId}`);
    if (status) parts.push(`status=${status}`);
    return `${fallbackMessage}（接口返回任务态：${parts.join("，")}）`;
  }
  return fallbackMessage;
}

async function readVisitDailyTotals() {
  try {
    const raw = await fsp.readFile(VISIT_DURATIONS_FILE, "utf8");
    const parsed = JSON.parse(raw);
    const source = parsed && typeof parsed === "object" && parsed.dailyTotals ? parsed.dailyTotals : parsed;
    return sanitizeVisitDailyTotals(source);
  } catch {
    return {};
  }
}

async function writeVisitDailyTotals(totals) {
  const sanitized = sanitizeVisitDailyTotals(totals);
  await fsp.mkdir(DATA_DIR, { recursive: true });
  const payload = JSON.stringify({ dailyTotals: sanitized }, null, 2);
  await fsp.writeFile(VISIT_DURATIONS_FILE, payload, "utf8");
  return sanitized;
}

async function readLocalShortcuts() {
  try {
    const raw = await fsp.readFile(LOCAL_SHORTCUTS_FILE, "utf8");
    return sanitizeLocalShortcuts(JSON.parse(raw));
  } catch {
    return { applications: [], files: [] };
  }
}

async function writeLocalShortcuts(shortcuts) {
  const sanitized = sanitizeLocalShortcuts(shortcuts);
  await fsp.mkdir(DATA_DIR, { recursive: true });
  const payload = JSON.stringify(sanitized, null, 2);
  try {
    await backupCurrentLocalShortcutsIfNeeded(payload);
  } catch (error) {
    const message = error && error.message ? error.message : String(error || "unknown");
    console.error(`[local-shortcuts] history backup failed: ${message}`);
  }
  await fsp.writeFile(LOCAL_SHORTCUTS_FILE, payload, "utf8");
  return sanitized;
}

function formatHistoryStamp(date = new Date()) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, "0")}${String(date.getDate()).padStart(
    2,
    "0"
  )}-${String(date.getHours()).padStart(2, "0")}${String(date.getMinutes()).padStart(2, "0")}${String(
    date.getSeconds()
  ).padStart(2, "0")}-${String(date.getMilliseconds()).padStart(3, "0")}`;
}

async function pruneLocalShortcutHistoryFiles() {
  let entries = [];
  try {
    entries = await fsp.readdir(LOCAL_SHORTCUTS_HISTORY_DIR, { withFileTypes: true });
  } catch {
    return;
  }

  const files = entries
    .filter((entry) => entry.isFile() && /^local-shortcuts-\d{8}-\d{6}-\d{3}\.json$/i.test(entry.name))
    .map((entry) => entry.name)
    .sort((a, b) => b.localeCompare(a, "en", { numeric: true, sensitivity: "base" }));

  if (files.length <= LOCAL_SHORTCUTS_HISTORY_KEEP_LIMIT) return;
  const staleFiles = files.slice(LOCAL_SHORTCUTS_HISTORY_KEEP_LIMIT);
  await Promise.all(
    staleFiles.map(async (fileName) => {
      try {
        await fsp.unlink(path.join(LOCAL_SHORTCUTS_HISTORY_DIR, fileName));
      } catch {
        // Ignore stale history cleanup failure.
      }
    })
  );
}

async function backupCurrentLocalShortcutsIfNeeded(nextPayload) {
  let previousPayload = "";
  try {
    previousPayload = await fsp.readFile(LOCAL_SHORTCUTS_FILE, "utf8");
  } catch (error) {
    if (error && error.code === "ENOENT") return;
    throw error;
  }

  const prev = previousPayload.trim();
  const next = String(nextPayload || "").trim();
  if (!prev || prev === next) return;

  await fsp.mkdir(LOCAL_SHORTCUTS_HISTORY_DIR, { recursive: true });
  const backupName = `local-shortcuts-${formatHistoryStamp()}.json`;
  const backupPath = path.join(LOCAL_SHORTCUTS_HISTORY_DIR, backupName);
  await fsp.writeFile(backupPath, previousPayload, "utf8");
  await pruneLocalShortcutHistoryFiles();
}

async function openLocalPath(rawPath) {
  const targetPath = normalizeAbsoluteLocalPath(rawPath);
  if (!targetPath) {
    const error = new Error("Invalid Path");
    error.code = "INVALID_PATH";
    throw error;
  }

  let stat;
  try {
    stat = await fsp.stat(targetPath);
  } catch {
    const error = new Error("Path Not Found");
    error.code = "PATH_NOT_FOUND";
    throw error;
  }

  if (!stat.isFile() && !stat.isDirectory()) {
    const error = new Error("Unsupported Path Type");
    error.code = "UNSUPPORTED_PATH_TYPE";
    throw error;
  }

  await new Promise((resolve, reject) => {
    const child = spawn("cmd.exe", ["/d", "/s", "/c", "start", "", "/b", targetPath], {
      windowsHide: true,
      detached: true,
      stdio: "ignore",
    });
    child.on("error", (error) => {
      const err = new Error(error?.message || "Failed to launch local path");
      err.code = "OPEN_FAILED";
      reject(err);
    });
    child.unref();
    resolve();
  });

  return targetPath;
}

async function openLocalShortcut(shortcuts, category, id) {
  const list = category === "applications" ? shortcuts.applications : shortcuts.files;
  const item = Array.isArray(list) ? list.find((entry) => entry.id === id) : null;
  if (!item) {
    const error = new Error("Shortcut Not Found");
    error.code = "SHORTCUT_NOT_FOUND";
    throw error;
  }

  await openLocalPath(item.path);
  return item;
}

function sanitizeLocalAgentText(rawValue, maxLength = 240) {
  return String(rawValue || "").replace(/\u0000/g, "").trim().slice(0, maxLength);
}

function sanitizeLocalAgentPrompt(rawValue, maxLength = 2000) {
  return String(rawValue || "")
    .replace(/\u0000/g, "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .trim()
    .slice(0, maxLength);
}

function createLocalAgentActionError(message, code = "LOCAL_AGENT_ERROR", status = 400) {
  const error = new Error(sanitizeLocalAgentText(message, 220) || "本地执行失败");
  error.code = code;
  error.status = status;
  return error;
}

function detectLocalAgentApplyIntent(rawPrompt, rawPayload) {
  if (normalizeAiWebSearchBoolean(rawPayload?.apply ?? rawPayload?.confirm, false)) return true;
  const text = sanitizeLocalAgentPrompt(rawPrompt, 1200).toLowerCase();
  if (!text) return false;
  if (/(?:预览|先看|仅查看|不要执行|先不要|dry[\s-]?run)/i.test(text)) return false;
  return /(?:确认|执行|开始|立即|马上|apply|confirm|do it now|正式执行)/i.test(text);
}

function detectLocalAgentPreviewIntent(rawPrompt, rawPayload) {
  if (normalizeAiWebSearchBoolean(rawPayload?.preview, false)) return true;
  const text = sanitizeLocalAgentPrompt(rawPrompt, 1200).toLowerCase();
  if (!text) return false;
  return /(?:预览|先看|仅查看|不要执行|dry[\s-]?run|simulate)/i.test(text);
}

function extractLocalAgentPathCandidates(rawPrompt, maxCount = 3) {
  const text = sanitizeLocalAgentPrompt(rawPrompt, 2400);
  if (!text) return [];
  const output = [];
  const seen = new Set();
  const collect = (rawValue) => {
    let trimmed = String(rawValue || "")
      .trim()
      .replace(/[（(]\s*(?:先预览|预览|仅预览|preview|dry[\s-]?run|确认执行|执行)\s*[）)]?$/i, "")
      .replace(/\s+(?:先预览|预览|仅预览|preview|dry[\s-]?run|确认执行|执行)\s*$/i, "")
      .replace(/[。；，、,;:：!！?？）)】\]]+$/g, "")
      .trim();
    const safe = normalizeAbsoluteLocalPath(trimmed);
    if (!safe) return;
    const key = safe.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    output.push(safe);
  };

  const quotedPattern = /["“']((?:[A-Za-z]:[\\/]|\\\\)[^"”']+)["”']/g;
  let quotedMatch = quotedPattern.exec(text);
  while (quotedMatch) {
    collect(quotedMatch[1]);
    if (output.length >= maxCount) return output;
    quotedMatch = quotedPattern.exec(text);
  }

  const loosePattern = /((?:[A-Za-z]:[\\/]|\\\\)[^"“”\r\n]+?)(?=\s*(?:[，。；;！？!?]|(?:并且|并|然后|再|来|并请|并帮|并把)\b|$))/g;
  let looseMatch = loosePattern.exec(text);
  while (looseMatch) {
    collect(looseMatch[1]);
    if (output.length >= maxCount) return output;
    looseMatch = loosePattern.exec(text);
  }

  const plainPattern = /((?:[A-Za-z]:[\\/]|\\\\)[^\s"'“”<>|?*\r\n]+)/g;
  let plainMatch = plainPattern.exec(text);
  while (plainMatch) {
    collect(plainMatch[1]);
    if (output.length >= maxCount) return output;
    plainMatch = plainPattern.exec(text);
  }

  return output.slice(0, maxCount);
}

function extractLocalAgentQuotedSegments(rawPrompt, maxCount = 8) {
  const text = sanitizeLocalAgentPrompt(rawPrompt, 2800);
  if (!text) return [];
  const output = [];
  const pattern = /["“']([^"“”']{1,2000})["”']/g;
  let matched = pattern.exec(text);
  while (matched) {
    const candidate = String(matched[1] || "").replace(/\u0000/g, "").slice(0, 2000);
    if (candidate.trim()) output.push(candidate);
    if (output.length >= maxCount) break;
    matched = pattern.exec(text);
  }
  return output.slice(0, maxCount);
}

function extractLocalAgentPayloadPath(rawPayload, promptText = "") {
  const directPath = normalizeAbsoluteLocalPath(rawPayload?.path || rawPayload?.targetPath || "");
  if (directPath) return directPath;
  const candidates = extractLocalAgentPathCandidates(promptText, 2);
  return candidates.length ? candidates[0] : "";
}

function normalizeLocalAgentOpenKeyword(rawValue) {
  let keyword = String(rawValue || "")
    .replace(/\u0000/g, " ")
    .replace(/\r\n/g, " ")
    .replace(/\r/g, " ")
    .trim();
  if (!keyword) return "";
  keyword = keyword
    .replace(/[“”"'`]/g, "")
    .replace(/^[：:，,、;\-\s]+/, "")
    .replace(/[。；;!?！？\s]+$/g, "")
    .replace(/\s+(?:先预览|预览|仅预览|preview|dry[\s-]?run|确认执行|执行)\s*$/i, "")
    .replace(/^(?:第\s*\d{1,2}\s*(?:个|条|项)\s*)/i, "")
    .replace(/^(?:并且|并|然后|再|请)?\s*(?:打开|启动|播放|查看|浏览)\s*/i, "")
    .replace(/^(?:文件|目录|文件夹)\s*/i, "")
    .trim();
  return keyword.slice(0, 180);
}

function extractLocalAgentPreferredCandidateIndex(promptText = "") {
  const text = sanitizeLocalAgentPrompt(promptText, 1200);
  if (!text) return 0;
  const matched = text.match(/第\s*([1-9]\d{0,1})\s*(?:个|条|项)/i);
  if (!matched) return 0;
  const index = Number(matched[1]);
  if (!Number.isFinite(index) || index <= 0) return 0;
  return Math.floor(index);
}

function extractLocalAgentOpenKeyword(rawPayload, promptText = "") {
  const payloadKeyword = normalizeLocalAgentOpenKeyword(
    rawPayload?.keyword || rawPayload?.name || rawPayload?.fileName || rawPayload?.targetName || ""
  );
  if (payloadKeyword) return payloadKeyword;

  const quoted = extractLocalAgentQuotedSegments(promptText, 8)
    .map((entry) => normalizeLocalAgentOpenKeyword(entry))
    .filter((entry) => entry && !normalizeAbsoluteLocalPath(entry));
  if (quoted.length) return quoted[0];

  const text = sanitizeLocalAgentPrompt(promptText, 2000);
  if (!text) return "";
  const matched =
    text.match(
      /(?:打开|启动|播放|查看|浏览|open|launch|play)\s*(?:本地)?(?:文件|目录|文件夹|音乐|视频|图片|文档)?\s*(?:[:：]|为|是)?\s*([^\n\r]+)/i
    ) ||
    text.match(/(?:文件|目录|文件夹)\s*(?:名|名称)?\s*(?:是|为|叫)?\s*([^\n\r]+)/i);
  if (!matched || !matched[1]) return "";
  const candidate = String(matched[1] || "")
    .replace(/\s*(?:并且|并|然后|再|并请|并帮|并把).+$/i, "")
    .trim();
  return normalizeLocalAgentOpenKeyword(candidate);
}

function resolveLocalAgentOpenSearchRoots(rawPayload = {}) {
  const roots = [];
  const pushRoot = (rawValue) => {
    const safePath = normalizeAbsoluteLocalPath(rawValue);
    if (!safePath) return;
    roots.push(safePath);
  };

  const directList = [];
  if (Array.isArray(rawPayload?.searchDirs)) directList.push(...rawPayload.searchDirs);
  if (Array.isArray(rawPayload?.searchRoots)) directList.push(...rawPayload.searchRoots);
  if (Array.isArray(rawPayload?.roots)) directList.push(...rawPayload.roots);
  directList.forEach((entry) => pushRoot(entry));

  const homeDir = os.homedir();
  if (homeDir && path.win32.isAbsolute(homeDir)) {
    [
      path.join(homeDir, "Desktop"),
      path.join(homeDir, "Documents"),
      path.join(homeDir, "Downloads"),
      path.join(homeDir, "Music"),
      path.join(homeDir, "Pictures"),
      path.join(homeDir, "Videos"),
      path.join(homeDir, "OneDrive", "Desktop"),
      path.join(homeDir, "OneDrive", "Documents"),
      path.join(homeDir, "OneDrive", "Pictures"),
      path.join(homeDir, "OneDrive", "Music"),
      path.join(homeDir, "OneDrive", "Videos"),
    ].forEach((entry) => pushRoot(entry));
  }

  try {
    const tempDir = os.tmpdir();
    if (tempDir) pushRoot(tempDir);
  } catch {
    // ignore temp dir failure
  }
  pushRoot(ROOT_DIR);

  const deduped = [];
  const seen = new Set();
  roots.forEach((entry) => {
    const normalized = normalizeAbsoluteLocalPath(entry);
    if (!normalized) return;
    const key = normalized.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    try {
      const stat = fs.statSync(normalized);
      if (!stat.isDirectory()) return;
      deduped.push(normalized);
    } catch {
      // ignore non-existing root
    }
  });
  return deduped.slice(0, 20);
}

function scoreLocalAgentNameMatch(candidateName, keywordRaw) {
  const keyword = normalizeLocalAgentOpenKeyword(keywordRaw).toLowerCase();
  const name = sanitizeLocalAgentText(candidateName, 240).toLowerCase();
  if (!keyword || !name) return 0;
  const nameNoExt = path.parse(name).name || name;
  const keywordNoExt = path.parse(keyword).name || keyword;
  let score = 0;
  if (name === keyword) score = 160;
  else if (nameNoExt === keyword || name === keywordNoExt || nameNoExt === keywordNoExt) score = 148;
  else if (name.startsWith(keyword)) score = 128;
  else if (nameNoExt.startsWith(keywordNoExt)) score = 118;
  else if (name.includes(keyword)) score = 96;
  else if (nameNoExt.includes(keywordNoExt)) score = 86;
  else {
    const keywordParts = keyword.split(/\s+/).filter(Boolean);
    if (keywordParts.length > 1 && keywordParts.every((part) => name.includes(part))) {
      score = 72;
    }
  }
  if (!score) return 0;
  const keywordExt = path.extname(keyword);
  const nameExt = path.extname(name);
  if (keywordExt && keywordExt === nameExt) score += 10;
  if (keywordExt && keywordExt !== nameExt && nameExt) score -= 4;
  return Math.max(1, score);
}

async function collectLocalAgentShortcutMatches(keyword, maxMatches = LOCAL_AGENT_OPEN_SEARCH_MAX_MATCHES) {
  const shortcuts = await readLocalShortcuts().catch(() => ({ applications: [], files: [] }));
  const entries = [
    ...(Array.isArray(shortcuts?.files) ? shortcuts.files : []),
    ...(Array.isArray(shortcuts?.applications) ? shortcuts.applications : []),
  ];
  const output = [];
  entries.forEach((entry) => {
    const safePath = normalizeAbsoluteLocalPath(entry?.path || "");
    if (!safePath) return;
    const name = sanitizeLocalAgentText(entry?.name || inferShortcutName(safePath), 180) || path.basename(safePath);
    const baseName = path.basename(safePath);
    const score = Math.max(scoreLocalAgentNameMatch(name, keyword), scoreLocalAgentNameMatch(baseName, keyword));
    if (!score) return;
    output.push({
      path: safePath,
      name: name || baseName,
      score: score + 12,
      matchType: "shortcut",
      category: entry?.path && /\.(exe|lnk|bat|cmd)$/i.test(String(entry.path)) ? "应用" : "文件",
      source: "shortcut",
    });
  });
  output.sort((a, b) => b.score - a.score || a.path.length - b.path.length || a.path.localeCompare(b.path, "zh-CN"));
  return output.slice(0, Math.max(1, maxMatches));
}

async function collectLocalAgentFileSystemMatches(keyword, searchRoots = [], options = {}) {
  const roots = Array.isArray(searchRoots) ? searchRoots : [];
  const maxDepth = clampAiInteger(options?.maxDepth, LOCAL_AGENT_OPEN_SEARCH_MAX_DEPTH, 1, LOCAL_AGENT_OPEN_SEARCH_MAX_DEPTH);
  const maxDirs = clampAiInteger(options?.maxDirs, LOCAL_AGENT_OPEN_SEARCH_MAX_DIRS, 100, LOCAL_AGENT_OPEN_SEARCH_MAX_DIRS);
  const maxEntries = clampAiInteger(options?.maxEntries, LOCAL_AGENT_OPEN_SEARCH_MAX_ENTRIES, 2000, LOCAL_AGENT_OPEN_SEARCH_MAX_ENTRIES);
  const maxMatches = clampAiInteger(options?.maxMatches, LOCAL_AGENT_OPEN_SEARCH_MAX_MATCHES, 5, LOCAL_AGENT_OPEN_SEARCH_MAX_MATCHES);
  const queue = roots.map((root) => ({ dir: root, depth: 0 }));
  const visited = new Set();
  const matched = [];
  let scannedDirs = 0;
  let scannedEntries = 0;
  let truncated = false;

  while (queue.length) {
    if (scannedDirs >= maxDirs || scannedEntries >= maxEntries) {
      truncated = true;
      break;
    }
    const current = queue.shift();
    const safeDir = normalizeAbsoluteLocalPath(current?.dir || "");
    if (!safeDir) continue;
    const dirKey = safeDir.toLowerCase();
    if (visited.has(dirKey)) continue;
    visited.add(dirKey);
    scannedDirs += 1;

    let entries = [];
    try {
      entries = await fsp.readdir(safeDir, { withFileTypes: true });
    } catch {
      continue;
    }

    for (const entry of entries) {
      scannedEntries += 1;
      if (scannedEntries >= maxEntries) {
        truncated = true;
        break;
      }
      const fullPath = path.join(safeDir, entry.name);
      if (entry.isDirectory()) {
        if (current.depth < maxDepth) {
          queue.push({ dir: fullPath, depth: current.depth + 1 });
        }
      }

      const score = scoreLocalAgentNameMatch(entry.name, keyword);
      if (!score) continue;
      matched.push({
        path: fullPath,
        name: entry.name,
        score: entry.isDirectory() ? score - 8 : score,
        matchType: entry.isDirectory() ? "directory" : "file",
        category: entry.isDirectory() ? "目录" : "文件",
        source: "filesystem",
      });
    }
    if (truncated) break;
  }

  matched.sort((a, b) => b.score - a.score || a.path.length - b.path.length || a.path.localeCompare(b.path, "zh-CN"));
  const deduped = [];
  const seen = new Set();
  for (const item of matched) {
    const safePath = normalizeAbsoluteLocalPath(item?.path || "");
    if (!safePath) continue;
    const key = safePath.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    deduped.push({
      ...item,
      path: safePath,
    });
    if (deduped.length >= maxMatches) break;
  }

  return {
    matches: deduped,
    scannedDirs,
    scannedEntries,
    truncated,
  };
}

function pickLocalAgentOpenCandidate(matches = [], preferredIndex = 0) {
  const list = Array.isArray(matches) ? matches : [];
  if (!list.length) return { candidate: null, confident: false, reason: "empty", selectedIndex: 0 };
  if (preferredIndex > 0 && preferredIndex <= list.length) {
    return {
      candidate: list[preferredIndex - 1],
      confident: true,
      reason: "preferred-index",
      selectedIndex: preferredIndex,
    };
  }
  const first = list[0];
  const second = list[1] || null;
  const gap = second ? first.score - second.score : 999;
  const confident = !second || (first.score >= 150 && gap >= 8) || (first.score >= 135 && gap >= 12) || (first.score >= 110 && gap >= 20);
  return {
    candidate: first,
    confident,
    reason: confident ? "top-score" : "ambiguous",
    selectedIndex: 1,
  };
}

async function executeLocalAgentOpenByKeyword(rawPayload, promptText, options = {}) {
  const keyword = extractLocalAgentOpenKeyword(rawPayload, promptText);
  if (!keyword) {
    throw createLocalAgentActionError("未识别到可打开的文件名，请补充文件名或绝对路径", "LOCAL_AGENT_OPEN_KEYWORD_REQUIRED", 400);
  }

  const searchRoots = resolveLocalAgentOpenSearchRoots(rawPayload);
  const [shortcutMatches, fsResult] = await Promise.all([
    collectLocalAgentShortcutMatches(keyword, LOCAL_AGENT_OPEN_SEARCH_MAX_MATCHES),
    collectLocalAgentFileSystemMatches(keyword, searchRoots, {
      maxDepth: LOCAL_AGENT_OPEN_SEARCH_MAX_DEPTH,
      maxDirs: LOCAL_AGENT_OPEN_SEARCH_MAX_DIRS,
      maxEntries: LOCAL_AGENT_OPEN_SEARCH_MAX_ENTRIES,
      maxMatches: LOCAL_AGENT_OPEN_SEARCH_MAX_MATCHES,
    }),
  ]);
  const fsMatches = Array.isArray(fsResult?.matches) ? fsResult.matches : [];
  const merged = [...shortcutMatches, ...fsMatches]
    .sort((a, b) => b.score - a.score || a.path.length - b.path.length || a.path.localeCompare(b.path, "zh-CN"))
    .slice(0, LOCAL_AGENT_OPEN_SEARCH_MAX_MATCHES);
  const deduped = [];
  const seen = new Set();
  for (const item of merged) {
    const safePath = normalizeAbsoluteLocalPath(item?.path || "");
    if (!safePath) continue;
    const key = safePath.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    deduped.push({
      ...item,
      path: safePath,
      name: sanitizeLocalAgentText(item?.name || path.basename(safePath), 180) || path.basename(safePath),
      category: sanitizeLocalAgentText(item?.category || "", 20) || "文件",
      source: sanitizeLocalAgentText(item?.source || "", 20),
    });
    if (deduped.length >= LOCAL_AGENT_OPEN_SEARCH_MAX_MATCHES) break;
  }

  if (!deduped.length) {
    throw createLocalAgentActionError(
      `未找到与“${sanitizeLocalAgentText(keyword, 80)}”匹配的本地文件`,
      "LOCAL_AGENT_OPEN_NOT_FOUND",
      404
    );
  }

  const preferredIndex = extractLocalAgentPreferredCandidateIndex(promptText);
  const picked = pickLocalAgentOpenCandidate(deduped, preferredIndex);
  const selected = picked.candidate;
  const explicitConfirm = options?.explicitConfirm === true;
  const previewRequested = options?.previewRequested === true;
  const shouldOpen = !previewRequested && selected && (picked.confident || explicitConfirm || preferredIndex > 0);
  const previewItems = deduped.slice(0, LOCAL_AGENT_PREVIEW_LIMIT).map((item, index) => ({
    file: sanitizeLocalAgentText(item.name, 180),
    category: sanitizeLocalAgentText(item.category, 20),
    path: sanitizeLocalAgentText(item.path, 260),
    index: index + 1,
    source: sanitizeLocalAgentText(item.source, 20),
  }));

  if (!shouldOpen) {
    return {
      action: LOCAL_AGENT_ACTION_OPEN,
      actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_OPEN),
      applied: false,
      previewOnly: true,
      targetPath: sanitizeLocalAgentText(selected?.path || "", 260),
      summary: `匹配到 ${deduped.length} 个候选，未自动打开`,
      warning: "结果不唯一，请补充关键词、指定绝对路径，或说“确认执行并打开第1个”",
      details: {
        keyword: sanitizeLocalAgentText(keyword, 120),
        matchedCount: deduped.length,
        selectedIndex: picked.selectedIndex || 0,
        scannedDirectories: Number(fsResult?.scannedDirs) || 0,
        scannedEntries: Number(fsResult?.scannedEntries) || 0,
        truncated: fsResult?.truncated === true,
        previewItems,
      },
      generatedAtMs: Date.now(),
    };
  }

  const openedPath = await openLocalPath(selected.path);
  return {
    action: LOCAL_AGENT_ACTION_OPEN,
    actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_OPEN),
    applied: true,
    previewOnly: false,
    targetPath: sanitizeLocalAgentText(openedPath, 260),
    summary: `已打开：${sanitizeLocalAgentText(selected.name, 160)}`,
    warning: picked.confident ? "" : "候选结果不唯一，已按排名优先打开",
    details: {
      keyword: sanitizeLocalAgentText(keyword, 120),
      matchedCount: deduped.length,
      selectedIndex: picked.selectedIndex || 1,
      selectedPath: sanitizeLocalAgentText(openedPath, 260),
      selectedScore: Number.isFinite(Number(selected.score)) ? Number(selected.score) : 0,
      scannedDirectories: Number(fsResult?.scannedDirs) || 0,
      scannedEntries: Number(fsResult?.scannedEntries) || 0,
      truncated: fsResult?.truncated === true,
      previewItems,
    },
    generatedAtMs: Date.now(),
  };
}

function resolveLocalAgentActionLabel(action) {
  if (action === LOCAL_AGENT_ACTION_SYSTEM_INFO) return "查看系统信息";
  if (action === LOCAL_AGENT_ACTION_ORGANIZE) return "整理文件";
  if (action === LOCAL_AGENT_ACTION_RENAME) return "批量重命名";
  if (action === LOCAL_AGENT_ACTION_EDIT) return "编辑文本文件";
  if (action === LOCAL_AGENT_ACTION_OPEN) return "打开本地文件/目录";
  return "本地执行";
}

function resolveLocalAgentAction(rawPayload, promptText = "", targetPath = "") {
  const explicitAction = sanitizeLocalAgentText(rawPayload?.action, 60).toLowerCase();
  if (
    [
      LOCAL_AGENT_ACTION_SYSTEM_INFO,
      LOCAL_AGENT_ACTION_ORGANIZE,
      LOCAL_AGENT_ACTION_RENAME,
      LOCAL_AGENT_ACTION_EDIT,
      LOCAL_AGENT_ACTION_OPEN,
    ].includes(explicitAction)
  ) {
    return explicitAction;
  }

  const prompt = sanitizeLocalAgentPrompt(promptText, 1200).toLowerCase();
  if (!prompt) return "";
  if (/(?:系统信息|电脑配置|硬件信息|cpu|内存|磁盘|系统版本|os version|system info|computer specs)/i.test(prompt)) {
    return LOCAL_AGENT_ACTION_SYSTEM_INFO;
  }
  if (
    /(?:整理|归类|分类|收纳|organize|sort files|clean up).*(?:文件夹|目录|文件|folder|directory)/i.test(prompt) ||
    /(?:按类型整理|分类整理|自动整理)/i.test(prompt)
  ) {
    return LOCAL_AGENT_ACTION_ORGANIZE;
  }
  if (/(?:重命名|改名|批量改名|rename)/i.test(prompt)) {
    return LOCAL_AGENT_ACTION_RENAME;
  }
  if (/(?:追加|写入|编辑|修改|替换|append|overwrite|replace|edit)/i.test(prompt)) {
    return LOCAL_AGENT_ACTION_EDIT;
  }
  if (/(?:打开|启动|播放|查看|浏览|open|launch|play)/i.test(prompt)) return LOCAL_AGENT_ACTION_OPEN;
  if (targetPath) return LOCAL_AGENT_ACTION_OPEN;
  return "";
}

async function runLocalAgentCommand(command, args = [], options = {}) {
  const timeoutMs = clampAiInteger(options?.timeoutMs, LOCAL_AGENT_COMMAND_TIMEOUT_MS, 1500, 120000);
  return new Promise((resolve) => {
    let stdout = "";
    let stderr = "";
    let settled = false;
    let timedOut = false;
    const child = spawn(command, args, {
      windowsHide: true,
      stdio: ["ignore", "pipe", "pipe"],
      shell: false,
    });
    const timer = setTimeout(() => {
      timedOut = true;
      try {
        child.kill();
      } catch {
        // ignore kill failure
      }
    }, timeoutMs);
    const finalize = (payload) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      resolve(payload);
    };
    if (child.stdout) {
      child.stdout.on("data", (chunk) => {
        if (typeof chunk === "string") stdout += chunk;
        else stdout += Buffer.from(chunk).toString("utf8");
      });
    }
    if (child.stderr) {
      child.stderr.on("data", (chunk) => {
        if (typeof chunk === "string") stderr += chunk;
        else stderr += Buffer.from(chunk).toString("utf8");
      });
    }
    child.on("error", (error) => {
      finalize({
        ok: false,
        status: -1,
        stdout,
        stderr: stderr || sanitizeLocalAgentText(error?.message || "", 280),
        timedOut,
      });
    });
    child.on("close", (code) => {
      finalize({
        ok: code === 0 && !timedOut,
        status: Number.isFinite(Number(code)) ? Number(code) : -1,
        stdout,
        stderr,
        timedOut,
      });
    });
  });
}

async function collectLocalAgentDiskInfoWindows() {
  const psCommand =
    "Get-CimInstance Win32_LogicalDisk -Filter \"DriveType=3\" | Select-Object DeviceID,Size,FreeSpace | ConvertTo-Json -Compress";
  const result = await runLocalAgentCommand("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", psCommand], {
    timeoutMs: Math.min(LOCAL_AGENT_COMMAND_TIMEOUT_MS, 12000),
  });
  if (!result.ok || !result.stdout.trim()) return [];
  let parsed;
  try {
    parsed = JSON.parse(result.stdout.trim());
  } catch {
    return [];
  }
  const list = Array.isArray(parsed) ? parsed : parsed && typeof parsed === "object" ? [parsed] : [];
  return list
    .map((entry) => {
      const device = sanitizeLocalAgentText(entry?.DeviceID || "", 20);
      const size = Number(entry?.Size);
      const free = Number(entry?.FreeSpace);
      if (!device || !Number.isFinite(size) || size <= 0 || !Number.isFinite(free) || free < 0) return null;
      const used = Math.max(0, size - free);
      const usedPercent = size > 0 ? Math.min(100, Math.max(0, Math.round((used / size) * 1000) / 10)) : 0;
      return {
        device,
        sizeBytes: Math.floor(size),
        freeBytes: Math.floor(free),
        usedBytes: Math.floor(used),
        usedPercent,
      };
    })
    .filter(Boolean)
    .slice(0, 24);
}

function formatLocalAgentBytes(rawValue) {
  const value = Number(rawValue);
  if (!Number.isFinite(value) || value < 0) return "0 B";
  if (value < 1024) return `${Math.floor(value)} B`;
  const units = ["KB", "MB", "GB", "TB", "PB"];
  let current = value / 1024;
  let unitIndex = 0;
  while (current >= 1024 && unitIndex < units.length - 1) {
    current /= 1024;
    unitIndex += 1;
  }
  const fixed = current >= 100 ? current.toFixed(0) : current >= 10 ? current.toFixed(1) : current.toFixed(2);
  return `${fixed} ${units[unitIndex]}`;
}

async function executeLocalAgentSystemInfo(promptText = "") {
  const cpus = Array.isArray(os.cpus()) ? os.cpus() : [];
  const totalMemory = os.totalmem();
  const freeMemory = os.freemem();
  const usedMemory = Math.max(0, totalMemory - freeMemory);
  const uptimeSec = Math.max(0, Math.floor(os.uptime()));
  const diskList = process.platform === "win32" ? await collectLocalAgentDiskInfoWindows() : [];
  const lines = [
    `系统：${os.type()} ${os.release()} (${os.arch()})`,
    `主机名：${os.hostname()}`,
    `CPU：${cpus.length} 核 ${sanitizeLocalAgentText(cpus[0]?.model || "", 120) || ""}`.trim(),
    `内存：已用 ${formatLocalAgentBytes(usedMemory)} / 总计 ${formatLocalAgentBytes(totalMemory)}`,
    `运行时长：${Math.floor(uptimeSec / 3600)} 小时 ${Math.floor((uptimeSec % 3600) / 60)} 分钟`,
  ];
  if (diskList.length) {
    diskList.forEach((disk) => {
      lines.push(
        `磁盘 ${disk.device}：已用 ${formatLocalAgentBytes(disk.usedBytes)} / 总计 ${formatLocalAgentBytes(disk.sizeBytes)}（${disk.usedPercent}%）`
      );
    });
  }
  return {
    action: LOCAL_AGENT_ACTION_SYSTEM_INFO,
    actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_SYSTEM_INFO),
    applied: true,
    previewOnly: false,
    targetPath: "",
    summary: "已获取本机系统信息",
    warning: "",
    details: {
      prompt: sanitizeLocalAgentText(promptText, 240),
      platform: process.platform,
      arch: os.arch(),
      hostname: os.hostname(),
      cpuModel: sanitizeLocalAgentText(cpus[0]?.model || "", 160),
      cpuCores: cpus.length,
      totalMemoryBytes: totalMemory,
      freeMemoryBytes: freeMemory,
      uptimeSeconds: uptimeSec,
      disks: diskList,
      textLines: lines,
    },
    generatedAtMs: Date.now(),
  };
}

function resolveLocalAgentCategoryByExtension(extension = "") {
  const ext = String(extension || "").trim().toLowerCase();
  for (const rule of LOCAL_AGENT_CATEGORY_RULES) {
    if (rule.extensions.has(ext)) return rule;
  }
  return { name: "其他", folder: "其他" };
}

function resolveLocalAgentAvailablePath(targetPath, reservedKeys = new Set()) {
  const parsed = path.parse(targetPath);
  let index = 0;
  while (index < 5000) {
    const candidate =
      index === 0 ? targetPath : path.join(parsed.dir, `${parsed.name} (${index})${parsed.ext || ""}`);
    const key = String(candidate || "").toLowerCase();
    if (!key) return "";
    if (!reservedKeys.has(key) && !fs.existsSync(candidate)) {
      reservedKeys.add(key);
      return candidate;
    }
    index += 1;
  }
  return "";
}

async function resolveLocalAgentDirectoryPath(rawPath) {
  const targetPath = normalizeAbsoluteLocalPath(rawPath);
  if (!targetPath) {
    throw createLocalAgentActionError("缺少有效本地路径", "LOCAL_AGENT_PATH_REQUIRED", 400);
  }
  let stat;
  try {
    stat = await fsp.stat(targetPath);
  } catch {
    throw createLocalAgentActionError("目标路径不存在", "LOCAL_AGENT_PATH_NOT_FOUND", 404);
  }
  if (stat.isDirectory()) {
    return {
      targetPath,
      directoryPath: targetPath,
      targetType: "directory",
    };
  }
  if (stat.isFile()) {
    return {
      targetPath,
      directoryPath: path.dirname(targetPath),
      targetType: "file",
    };
  }
  throw createLocalAgentActionError("目标路径类型不支持", "LOCAL_AGENT_UNSUPPORTED_PATH_TYPE", 400);
}

async function collectLocalAgentFiles(directoryPath, options = {}) {
  const recursive = options?.recursive === true;
  const maxDepth = clampAiInteger(options?.maxDepth, LOCAL_AGENT_MAX_SCAN_DEPTH, 1, LOCAL_AGENT_MAX_SCAN_DEPTH);
  const maxFiles = clampAiInteger(options?.maxFiles, LOCAL_AGENT_MAX_SCAN_FILES, 20, LOCAL_AGENT_MAX_SCAN_FILES);
  const queue = [{ dir: directoryPath, depth: 0 }];
  const files = [];
  let scannedDirectories = 0;
  let truncated = false;

  while (queue.length && !truncated) {
    const current = queue.shift();
    scannedDirectories += 1;
    let entries = [];
    try {
      entries = await fsp.readdir(current.dir, { withFileTypes: true });
    } catch {
      continue;
    }
    for (const entry of entries) {
      const absPath = path.join(current.dir, entry.name);
      if (entry.isDirectory()) {
        if (recursive && current.depth < maxDepth) {
          queue.push({ dir: absPath, depth: current.depth + 1 });
        }
        continue;
      }
      if (!entry.isFile()) continue;
      let stat;
      try {
        stat = await fsp.stat(absPath);
      } catch {
        continue;
      }
      files.push({
        path: absPath,
        name: entry.name,
        extension: path.extname(entry.name || "").toLowerCase(),
        sizeBytes: Number.isFinite(Number(stat.size)) ? Math.max(0, Math.floor(Number(stat.size))) : 0,
        relativePath: path.relative(directoryPath, absPath),
      });
      if (files.length >= maxFiles) {
        truncated = true;
        break;
      }
    }
  }
  return {
    files,
    scannedDirectories,
    truncated,
    recursive,
    maxDepth,
    maxFiles,
  };
}

function shouldLocalAgentUseRecursive(rawPrompt, rawPayload) {
  if (normalizeAiWebSearchBoolean(rawPayload?.recursive, false)) return true;
  const text = sanitizeLocalAgentPrompt(rawPrompt, 1200);
  if (!text) return false;
  return /(?:递归|包含子目录|全部子目录|所有子目录|子文件夹|含子文件夹|all subfolders|recursive)/i.test(text);
}

async function executeLocalAgentOrganize(rawPayload, promptText, targetPath, applyChanges) {
  const target = await resolveLocalAgentDirectoryPath(targetPath);
  const recursive = shouldLocalAgentUseRecursive(promptText, rawPayload);
  const scan = await collectLocalAgentFiles(target.directoryPath, {
    recursive,
    maxDepth: LOCAL_AGENT_MAX_SCAN_DEPTH,
    maxFiles: LOCAL_AGENT_MAX_SCAN_FILES,
  });
  const files = Array.isArray(scan.files) ? scan.files : [];
  if (!files.length) {
    return {
      action: LOCAL_AGENT_ACTION_ORGANIZE,
      actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_ORGANIZE),
      applied: false,
      previewOnly: true,
      targetPath: target.directoryPath,
      summary: "目录中没有可整理的文件",
      warning: "",
      details: {
        recursive,
        scannedDirectories: scan.scannedDirectories,
        fileCount: 0,
        planCount: 0,
        previewItems: [],
      },
      generatedAtMs: Date.now(),
    };
  }

  const reservedPaths = new Set();
  const plans = [];
  for (const file of files) {
    const categoryRule = resolveLocalAgentCategoryByExtension(file.extension);
    const destinationDir = path.join(target.directoryPath, categoryRule.folder);
    const destinationBase = path.join(destinationDir, file.name);
    const destinationPath = resolveLocalAgentAvailablePath(destinationBase, reservedPaths);
    if (!destinationPath) continue;
    const sourceKey = String(file.path || "").toLowerCase();
    const destinationKey = String(destinationPath || "").toLowerCase();
    if (!sourceKey || !destinationKey || sourceKey === destinationKey) continue;
    plans.push({
      sourcePath: file.path,
      destinationPath,
      destinationDir,
      category: categoryRule.name,
      fileName: file.name,
      sizeBytes: file.sizeBytes,
      relativePath: file.relativePath,
    });
    if (plans.length >= LOCAL_AGENT_MAX_RENAME_ITEMS) break;
  }

  const previewItems = plans.slice(0, LOCAL_AGENT_PREVIEW_LIMIT).map((entry) => ({
    file: sanitizeLocalAgentText(entry.fileName, 160),
    category: sanitizeLocalAgentText(entry.category, 40),
    from: sanitizeLocalAgentText(entry.sourcePath, 260),
    to: sanitizeLocalAgentText(entry.destinationPath, 260),
    sizeLabel: formatLocalAgentBytes(entry.sizeBytes),
  }));
  if (!plans.length) {
    return {
      action: LOCAL_AGENT_ACTION_ORGANIZE,
      actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_ORGANIZE),
      applied: false,
      previewOnly: true,
      targetPath: target.directoryPath,
      summary: "未生成可执行的整理计划",
      warning: "",
      details: {
        recursive,
        scannedDirectories: scan.scannedDirectories,
        fileCount: files.length,
        planCount: 0,
        previewItems,
      },
      generatedAtMs: Date.now(),
    };
  }

  if (!applyChanges) {
    return {
      action: LOCAL_AGENT_ACTION_ORGANIZE,
      actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_ORGANIZE),
      applied: false,
      previewOnly: true,
      targetPath: target.directoryPath,
      summary: `已生成整理预览，共 ${plans.length} 个文件待处理`,
      warning: "",
      details: {
        recursive,
        scannedDirectories: scan.scannedDirectories,
        fileCount: files.length,
        planCount: plans.length,
        previewItems,
        hint: "如需正式执行，请在下一条指令中加“确认执行”",
      },
      generatedAtMs: Date.now(),
    };
  }

  let movedCount = 0;
  const failedItems = [];
  for (const planItem of plans) {
    try {
      await fsp.mkdir(planItem.destinationDir, { recursive: true });
      await fsp.rename(planItem.sourcePath, planItem.destinationPath);
      movedCount += 1;
    } catch (error) {
      failedItems.push({
        file: sanitizeLocalAgentText(planItem.fileName, 160),
        from: sanitizeLocalAgentText(planItem.sourcePath, 260),
        to: sanitizeLocalAgentText(planItem.destinationPath, 260),
        error: sanitizeLocalAgentText(error?.message || "移动失败", 180),
      });
    }
  }
  const failedCount = failedItems.length;
  return {
    action: LOCAL_AGENT_ACTION_ORGANIZE,
    actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_ORGANIZE),
    applied: true,
    previewOnly: false,
    targetPath: target.directoryPath,
    summary: failedCount
      ? `整理完成，成功 ${movedCount} 个，失败 ${failedCount} 个`
      : `整理完成，共移动 ${movedCount} 个文件`,
    warning: scan.truncated ? "扫描达到上限，可能有部分文件未纳入此次整理" : "",
    details: {
      recursive,
      scannedDirectories: scan.scannedDirectories,
      fileCount: files.length,
      planCount: plans.length,
      movedCount,
      failedCount,
      previewItems,
      failedItems: failedItems.slice(0, LOCAL_AGENT_PREVIEW_LIMIT),
    },
    generatedAtMs: Date.now(),
  };
}

function extractLocalAgentRenamePrefix(rawPayload, promptText) {
  const payloadPrefix = sanitizeLocalAgentText(rawPayload?.prefix, 80).replace(/[\\/:*?"<>|]+/g, "-").trim();
  if (payloadPrefix) return payloadPrefix;
  const text = sanitizeLocalAgentPrompt(promptText, 1600);
  if (!text) return "file";
  const matched =
    text.match(/(?:重命名为|改名为|命名为|rename(?:\s+to)?|prefix)\s*[:：]?\s*["“']?([^"“”'，。,.\s]{1,80})/i) ||
    text.match(/(?:批量重命名|批量改名)\s*["“']?([^"“”'，。,.\s]{1,80})/i);
  if (matched && matched[1]) {
    const safe = sanitizeLocalAgentText(matched[1], 80).replace(/[\\/:*?"<>|]+/g, "-").trim();
    if (safe) return safe;
  }
  return "file";
}

async function executeLocalAgentRename(rawPayload, promptText, targetPath, applyChanges) {
  const target = await resolveLocalAgentDirectoryPath(targetPath);
  const recursive = shouldLocalAgentUseRecursive(promptText, rawPayload);
  const scan = await collectLocalAgentFiles(target.directoryPath, {
    recursive,
    maxDepth: recursive ? LOCAL_AGENT_MAX_SCAN_DEPTH : 1,
    maxFiles: LOCAL_AGENT_MAX_RENAME_ITEMS,
  });
  const files = Array.isArray(scan.files) ? scan.files : [];
  if (!files.length) {
    return {
      action: LOCAL_AGENT_ACTION_RENAME,
      actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_RENAME),
      applied: false,
      previewOnly: true,
      targetPath: target.directoryPath,
      summary: "目录中没有可重命名的文件",
      warning: "",
      details: {
        recursive,
        fileCount: 0,
        planCount: 0,
        previewItems: [],
      },
      generatedAtMs: Date.now(),
    };
  }

  const prefix = extractLocalAgentRenamePrefix(rawPayload, promptText);
  const sortedFiles = files
    .slice()
    .sort((a, b) => String(a.relativePath || a.name || "").localeCompare(String(b.relativePath || b.name || ""), "zh-CN", { numeric: true }));
  const sequenceWidth = Math.max(3, String(sortedFiles.length).length);
  const reservedPaths = new Set();
  const plans = [];
  sortedFiles.forEach((file, index) => {
    if (plans.length >= LOCAL_AGENT_MAX_RENAME_ITEMS) return;
    const ext = path.extname(file.name || "");
    const nextName = `${prefix}-${String(index + 1).padStart(sequenceWidth, "0")}${ext}`;
    const nextBasePath = path.join(path.dirname(file.path), nextName);
    const nextPath = resolveLocalAgentAvailablePath(nextBasePath, reservedPaths);
    if (!nextPath) return;
    const sourceKey = String(file.path || "").toLowerCase();
    const destinationKey = String(nextPath || "").toLowerCase();
    if (!sourceKey || !destinationKey || sourceKey === destinationKey) return;
    plans.push({
      sourcePath: file.path,
      destinationPath: nextPath,
      fromName: file.name,
      toName: path.basename(nextPath),
      relativePath: file.relativePath,
    });
  });

  const previewItems = plans.slice(0, LOCAL_AGENT_PREVIEW_LIMIT).map((entry) => ({
    from: sanitizeLocalAgentText(entry.fromName, 160),
    to: sanitizeLocalAgentText(entry.toName, 160),
    path: sanitizeLocalAgentText(entry.relativePath, 240),
  }));
  if (!plans.length) {
    return {
      action: LOCAL_AGENT_ACTION_RENAME,
      actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_RENAME),
      applied: false,
      previewOnly: true,
      targetPath: target.directoryPath,
      summary: "未生成可执行的重命名计划",
      warning: "",
      details: {
        recursive,
        fileCount: files.length,
        planCount: 0,
        previewItems,
      },
      generatedAtMs: Date.now(),
    };
  }

  if (!applyChanges) {
    return {
      action: LOCAL_AGENT_ACTION_RENAME,
      actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_RENAME),
      applied: false,
      previewOnly: true,
      targetPath: target.directoryPath,
      summary: `已生成重命名预览，共 ${plans.length} 个文件`,
      warning: "",
      details: {
        recursive,
        prefix,
        fileCount: files.length,
        planCount: plans.length,
        previewItems,
        hint: "如需正式执行，请在下一条指令中加“确认执行”",
      },
      generatedAtMs: Date.now(),
    };
  }

  let renamedCount = 0;
  const failedItems = [];
  for (const planItem of plans) {
    try {
      await fsp.rename(planItem.sourcePath, planItem.destinationPath);
      renamedCount += 1;
    } catch (error) {
      failedItems.push({
        from: sanitizeLocalAgentText(planItem.fromName, 160),
        to: sanitizeLocalAgentText(planItem.toName, 160),
        error: sanitizeLocalAgentText(error?.message || "重命名失败", 180),
      });
    }
  }
  const failedCount = failedItems.length;
  return {
    action: LOCAL_AGENT_ACTION_RENAME,
    actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_RENAME),
    applied: true,
    previewOnly: false,
    targetPath: target.directoryPath,
    summary: failedCount
      ? `重命名完成，成功 ${renamedCount} 个，失败 ${failedCount} 个`
      : `重命名完成，共处理 ${renamedCount} 个文件`,
    warning: scan.truncated ? "扫描达到上限，可能有部分文件未纳入此次重命名" : "",
    details: {
      recursive,
      prefix,
      fileCount: files.length,
      planCount: plans.length,
      renamedCount,
      failedCount,
      previewItems,
      failedItems: failedItems.slice(0, LOCAL_AGENT_PREVIEW_LIMIT),
    },
    generatedAtMs: Date.now(),
  };
}

function resolveLocalAgentEditInstruction(rawPayload, promptText) {
  const explicitMode = sanitizeLocalAgentText(rawPayload?.mode || rawPayload?.editMode, 40).toLowerCase();
  const quotedSegments = extractLocalAgentQuotedSegments(promptText, 8).filter((entry) => !normalizeAbsoluteLocalPath(entry));
  const payloadContent = String(rawPayload?.content ?? rawPayload?.text ?? rawPayload?.appendText ?? "")
    .replace(/\u0000/g, "")
    .slice(0, 24000);
  let mode = explicitMode;
  if (!mode) {
    if (/(?:替换|replace\s+.+\s+with)/i.test(promptText)) mode = "replace";
    else if (/(?:覆盖|重写|overwrite)/i.test(promptText)) mode = "overwrite";
    else if (/(?:追加|append|添加到末尾|附加)/i.test(promptText)) mode = "append";
    else if (/(?:写入|编辑|修改|write|edit)/i.test(promptText)) mode = "append";
    else mode = "append";
  }
  if (!["append", "overwrite", "replace"].includes(mode)) mode = "append";

  let content = payloadContent;
  if (!content && quotedSegments.length) {
    content = quotedSegments[quotedSegments.length - 1];
  }
  content = String(content || "").replace(/\u0000/g, "").slice(0, 24000);

  let replaceFrom = String(rawPayload?.replaceFrom ?? "").replace(/\u0000/g, "").slice(0, 6000);
  let replaceTo = String(rawPayload?.replaceTo ?? "").replace(/\u0000/g, "").slice(0, 6000);
  if (mode === "replace" && (!replaceFrom || replaceTo === "")) {
    if (quotedSegments.length >= 2) {
      replaceFrom = replaceFrom || quotedSegments[0];
      if (!replaceTo) replaceTo = quotedSegments[1];
    }
  }
  if (mode === "replace" && !replaceFrom) {
    const matched = sanitizeLocalAgentPrompt(promptText, 2400).match(/(?:把|将)?(.{1,120})替换(?:为|成)(.{1,120})/i);
    if (matched) {
      replaceFrom = replaceFrom || sanitizeLocalAgentText(matched[1], 120);
      if (!replaceTo) replaceTo = sanitizeLocalAgentText(matched[2], 120);
    }
  }

  return {
    mode,
    content,
    replaceFrom,
    replaceTo,
  };
}

function buildLocalAgentTextPreview(rawText, maxLength = LOCAL_AGENT_PREVIEW_LIMIT) {
  const text = String(rawText || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  if (!text) return "";
  const safeMax = clampAiInteger(maxLength, LOCAL_AGENT_PREVIEW_LIMIT, 20, 5000);
  if (text.length <= safeMax) return text;
  return `${text.slice(0, safeMax)}...`;
}

async function executeLocalAgentEdit(rawPayload, promptText, targetPath, applyChanges) {
  const normalizedPath = normalizeAbsoluteLocalPath(targetPath);
  if (!normalizedPath) {
    throw createLocalAgentActionError("缺少有效文件路径", "LOCAL_AGENT_PATH_REQUIRED", 400);
  }
  const extension = path.extname(normalizedPath).toLowerCase();
  if (!LOCAL_AGENT_TEXT_FILE_EXTENSIONS.has(extension)) {
    throw createLocalAgentActionError("当前文件类型不支持文本编辑", "LOCAL_AGENT_EDIT_UNSUPPORTED_FILE", 400);
  }

  const parentDir = path.dirname(normalizedPath);
  let parentStat;
  try {
    parentStat = await fsp.stat(parentDir);
  } catch {
    parentStat = null;
  }
  if (!parentStat || !parentStat.isDirectory()) {
    throw createLocalAgentActionError("文件所在目录不存在", "LOCAL_AGENT_PARENT_NOT_FOUND", 404);
  }

  let exists = false;
  let currentText = "";
  let currentBytes = 0;
  try {
    const stat = await fsp.stat(normalizedPath);
    if (stat.isFile()) {
      exists = true;
      currentBytes = Number.isFinite(Number(stat.size)) ? Math.max(0, Math.floor(Number(stat.size))) : 0;
      if (currentBytes > LOCAL_AGENT_MAX_EDIT_BYTES) {
        throw createLocalAgentActionError(
          `文件过大，超过编辑上限（${formatLocalAgentBytes(LOCAL_AGENT_MAX_EDIT_BYTES)}）`,
          "LOCAL_AGENT_EDIT_TOO_LARGE",
          413
        );
      }
      currentText = await fsp.readFile(normalizedPath, "utf8");
    } else {
      throw createLocalAgentActionError("目标路径不是文件", "LOCAL_AGENT_EDIT_PATH_NOT_FILE", 400);
    }
  } catch (error) {
    if (error?.code === "ENOENT") {
      exists = false;
      currentText = "";
      currentBytes = 0;
    } else if (error && error.code && String(error.code).startsWith("LOCAL_AGENT_")) {
      throw error;
    } else {
      throw createLocalAgentActionError(
        sanitizeLocalAgentText(error?.message || "读取文件失败", 180) || "读取文件失败",
        "LOCAL_AGENT_EDIT_READ_FAILED",
        400
      );
    }
  }

  const instruction = resolveLocalAgentEditInstruction(rawPayload, promptText);
  if (instruction.mode !== "replace" && !instruction.content) {
    throw createLocalAgentActionError("缺少编辑内容，请使用引号提供要写入的文本", "LOCAL_AGENT_EDIT_CONTENT_REQUIRED", 400);
  }

  let nextText = currentText;
  let replaceCount = 0;
  if (instruction.mode === "overwrite") {
    nextText = instruction.content;
  } else if (instruction.mode === "append") {
    nextText = `${currentText}${instruction.content}`;
  } else if (instruction.mode === "replace") {
    if (!exists) {
      throw createLocalAgentActionError("替换模式要求目标文件已存在", "LOCAL_AGENT_EDIT_FILE_REQUIRED", 400);
    }
    if (!instruction.replaceFrom) {
      throw createLocalAgentActionError("替换模式缺少“替换前”文本", "LOCAL_AGENT_EDIT_REPLACE_FROM_REQUIRED", 400);
    }
    replaceCount = currentText.includes(instruction.replaceFrom) ? currentText.split(instruction.replaceFrom).length - 1 : 0;
    nextText = currentText.split(instruction.replaceFrom).join(instruction.replaceTo || "");
  }

  const nextBytes = Buffer.byteLength(nextText, "utf8");
  if (nextBytes > LOCAL_AGENT_MAX_EDIT_BYTES) {
    throw createLocalAgentActionError(
      `修改后内容过大，超过上限（${formatLocalAgentBytes(LOCAL_AGENT_MAX_EDIT_BYTES)}）`,
      "LOCAL_AGENT_EDIT_RESULT_TOO_LARGE",
      413
    );
  }

  const beforePreview = buildLocalAgentTextPreview(currentText);
  const afterPreview = buildLocalAgentTextPreview(nextText);
  if (!applyChanges) {
    return {
      action: LOCAL_AGENT_ACTION_EDIT,
      actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_EDIT),
      applied: false,
      previewOnly: true,
      targetPath: normalizedPath,
      summary: "已生成编辑预览，未写入磁盘",
      warning: instruction.mode === "replace" && replaceCount === 0 ? "未匹配到可替换文本" : "",
      details: {
        mode: instruction.mode,
        fileExists: exists,
        beforeBytes: currentBytes,
        afterBytes: nextBytes,
        replaceCount,
        beforePreview,
        afterPreview,
        hint: "如需正式写入，请在下一条指令中加“确认执行”",
      },
      generatedAtMs: Date.now(),
    };
  }

  let backupPath = "";
  if (exists) {
    backupPath = `${normalizedPath}.bak-${formatHistoryStamp()}`;
    await fsp.copyFile(normalizedPath, backupPath);
  }
  await fsp.writeFile(normalizedPath, nextText, "utf8");
  return {
    action: LOCAL_AGENT_ACTION_EDIT,
    actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_EDIT),
    applied: true,
    previewOnly: false,
    targetPath: normalizedPath,
    summary: exists ? "文件编辑成功并已创建备份" : "文件创建并写入成功",
    warning: instruction.mode === "replace" && replaceCount === 0 ? "未匹配到可替换文本，文件内容未变化" : "",
    details: {
      mode: instruction.mode,
      fileExists: exists,
      beforeBytes: currentBytes,
      afterBytes: nextBytes,
      replaceCount,
      backupPath: sanitizeLocalAgentText(backupPath, 260),
      beforePreview,
      afterPreview,
    },
    generatedAtMs: Date.now(),
  };
}

async function executeLocalAgentOpenPath(targetPath) {
  const normalizedPath = normalizeAbsoluteLocalPath(targetPath);
  if (!normalizedPath) {
    throw createLocalAgentActionError("缺少有效本地路径", "LOCAL_AGENT_PATH_REQUIRED", 400);
  }
  const openedPath = await openLocalPath(normalizedPath);
  return {
    action: LOCAL_AGENT_ACTION_OPEN,
    actionLabel: resolveLocalAgentActionLabel(LOCAL_AGENT_ACTION_OPEN),
    applied: true,
    previewOnly: false,
    targetPath: openedPath,
    summary: "已在本机默认程序中打开目标路径",
    warning: "",
    details: {
      openedPath: sanitizeLocalAgentText(openedPath, 260),
    },
    generatedAtMs: Date.now(),
  };
}

async function executeAiLocalAgentRequest(rawPayload) {
  const payload = rawPayload && typeof rawPayload === "object" ? rawPayload : {};
  const promptText = sanitizeLocalAgentPrompt(payload?.prompt || payload?.query || payload?.text || "", 2000);
  const targetPath = extractLocalAgentPayloadPath(payload, promptText);
  const action = resolveLocalAgentAction(payload, promptText, targetPath);
  if (!action) {
    return {
      action: "",
      actionLabel: "",
      applied: false,
      previewOnly: true,
      targetPath: "",
      skipped: true,
      summary: "未识别到可执行的本地操作",
      warning: "",
      details: {
        prompt: sanitizeLocalAgentText(promptText, 240),
      },
      generatedAtMs: Date.now(),
    };
  }

  const applyIntent = detectLocalAgentApplyIntent(promptText, payload);
  const previewIntent = detectLocalAgentPreviewIntent(promptText, payload);
  const applyChanges = applyIntent && !previewIntent;

  let result;
  if (action === LOCAL_AGENT_ACTION_SYSTEM_INFO) {
    result = await executeLocalAgentSystemInfo(promptText);
  } else if (action === LOCAL_AGENT_ACTION_ORGANIZE) {
    if (!targetPath) {
      throw createLocalAgentActionError("整理文件需要提供本地目录路径", "LOCAL_AGENT_PATH_REQUIRED", 400);
    }
    result = await executeLocalAgentOrganize(payload, promptText, targetPath, applyChanges);
  } else if (action === LOCAL_AGENT_ACTION_RENAME) {
    if (!targetPath) {
      throw createLocalAgentActionError("批量重命名需要提供本地目录路径", "LOCAL_AGENT_PATH_REQUIRED", 400);
    }
    result = await executeLocalAgentRename(payload, promptText, targetPath, applyChanges);
  } else if (action === LOCAL_AGENT_ACTION_EDIT) {
    if (!targetPath) {
      throw createLocalAgentActionError("编辑文件需要提供本地文件路径", "LOCAL_AGENT_PATH_REQUIRED", 400);
    }
    result = await executeLocalAgentEdit(payload, promptText, targetPath, applyChanges);
  } else if (action === LOCAL_AGENT_ACTION_OPEN) {
    if (targetPath) {
      result = await executeLocalAgentOpenPath(targetPath);
    } else {
      result = await executeLocalAgentOpenByKeyword(payload, promptText, {
        explicitConfirm: applyIntent,
        previewRequested: previewIntent,
      });
    }
  } else {
    throw createLocalAgentActionError("不支持的本地操作类型", "LOCAL_AGENT_ACTION_UNSUPPORTED", 400);
  }

  const safeResult = result && typeof result === "object" ? result : {};
  return {
    action,
    actionLabel: resolveLocalAgentActionLabel(action),
    applied: safeResult.applied === true,
    previewOnly: safeResult.previewOnly !== false,
    targetPath: sanitizeLocalAgentText(safeResult.targetPath || targetPath, 260),
    skipped: false,
    summary: sanitizeLocalAgentText(safeResult.summary || "本地执行已完成", 240),
    warning: sanitizeLocalAgentText(safeResult.warning || "", 220),
    details: safeResult.details && typeof safeResult.details === "object" ? safeResult.details : {},
    generatedAtMs: Number.isFinite(Number(safeResult.generatedAtMs)) ? Math.floor(Number(safeResult.generatedAtMs)) : Date.now(),
  };
}

async function listGaozhiMarkdownFiles() {
  let entries = [];
  try {
    entries = await fsp.readdir(GAOZHI_DIR, { withFileTypes: true });
  } catch {
    return [];
  }
  const documents = await Promise.all(
    entries
      .filter((entry) => entry.isFile() && /\.md$/i.test(entry.name))
      .map(async (entry) => {
        const webPath = `gaozhi/${entry.name}`;
        const absPath = path.join(GAOZHI_DIR, entry.name);
        try {
          const stat = await fsp.stat(absPath);
          const birthtimeMs = Number.isFinite(stat.birthtimeMs) ? Math.floor(stat.birthtimeMs) : 0;
          const mtimeMs = Number.isFinite(stat.mtimeMs) ? Math.floor(stat.mtimeMs) : 0;
          const uploadedAtMs = birthtimeMs > 0 ? birthtimeMs : mtimeMs;
          return {
            path: webPath,
            uploadedAtMs,
            uploadedAt: uploadedAtMs > 0 ? new Date(uploadedAtMs).toISOString() : "",
          };
        } catch {
          return { path: webPath, uploadedAtMs: 0, uploadedAt: "" };
        }
      })
  );
  return documents.sort((a, b) => {
    if (b.uploadedAtMs !== a.uploadedAtMs) return b.uploadedAtMs - a.uploadedAtMs;
    return b.path.localeCompare(a.path, "zh-CN", { numeric: true, sensitivity: "base" });
  });
}

function resolveTutorialPostsDirectory() {
  if (fs.existsSync(FENXIANG_POSTS_DIR)) {
    return { absoluteDir: FENXIANG_POSTS_DIR, webPrefix: "fenxiang/_posts/" };
  }
  return { absoluteDir: FENXIANG_POSTS_DIR, webPrefix: "fenxiang/_posts/" };
}

async function listTutorialMarkdownFiles() {
  const { absoluteDir, webPrefix } = resolveTutorialPostsDirectory();
  const documents = [];
  async function walkDirectory(currentDir, relativePrefix = "") {
    let entries = [];
    try {
      entries = await fsp.readdir(currentDir, { withFileTypes: true });
    } catch {
      return;
    }
    await Promise.all(
      entries.map(async (entry) => {
        const nextRelative = relativePrefix ? `${relativePrefix}/${entry.name}` : entry.name;
        const normalizedRelative = nextRelative.replaceAll("\\", "/");
        const absPath = path.join(currentDir, entry.name);
        if (entry.isDirectory()) {
          await walkDirectory(absPath, normalizedRelative);
          return;
        }
        if (!entry.isFile() || !/\.md$/i.test(entry.name)) return;
        const webPath = `${webPrefix}${normalizedRelative}`;
        try {
          const stat = await fsp.stat(absPath);
          const birthtimeMs = Number.isFinite(stat.birthtimeMs) ? Math.floor(stat.birthtimeMs) : 0;
          const mtimeMs = Number.isFinite(stat.mtimeMs) ? Math.floor(stat.mtimeMs) : 0;
          // For tutorial sorting, prioritize modification time so "新/旧" reflects recent edits.
          const uploadedAtMs = mtimeMs > 0 ? mtimeMs : birthtimeMs;
          documents.push({
            path: webPath,
            uploadedAtMs,
            uploadedAt: uploadedAtMs > 0 ? new Date(uploadedAtMs).toISOString() : "",
          });
        } catch {
          documents.push({ path: webPath, uploadedAtMs: 0, uploadedAt: "" });
        }
      })
    );
  }
  await walkDirectory(absoluteDir);
  return documents.sort((a, b) => {
    if (b.uploadedAtMs !== a.uploadedAtMs) return b.uploadedAtMs - a.uploadedAtMs;
    return b.path.localeCompare(a.path, "zh-CN", { numeric: true, sensitivity: "base" });
  });
}

async function fetchOllamaModelTags(baseUrlRaw = AI_OLLAMA_DEFAULT_BASE_URL) {
  const baseUrl = normalizeAiHttpUrl(baseUrlRaw, AI_OLLAMA_DEFAULT_BASE_URL);
  const endpoint = buildAiEndpointUrl(baseUrl, "/api/tags");
  const result = await fetchAiJson(endpoint, { method: "GET", timeoutMs: 15000 });
  if (!result.ok) {
    const error = new Error(result?.data?.error || result?.text || `HTTP ${result.status}`);
    error.code = "OLLAMA_TAGS_FAILED";
    error.status = result.status;
    throw error;
  }
  const modelsRaw = Array.isArray(result?.data?.models) ? result.data.models : [];
  const models = modelsRaw
    .map((entry) => (typeof entry === "string" ? entry : String(entry?.name || "").trim()))
    .filter(Boolean);
  return models;
}

function buildAiOpenAiChatMessages({ messages, prompt, systemPrompt, imageDataUrl }) {
  const safeMessages = sanitizeAiMessages(messages);
  const output = [];
  const safeSystemPrompt = sanitizeAiSystemPrompt(systemPrompt);
  if (safeSystemPrompt) output.push({ role: "system", content: safeSystemPrompt });
  if (shouldInjectAiRuntimeDateGuard({ prompt, messages: safeMessages })) {
    output.push({ role: "system", content: buildAiRuntimeDateGuardPrompt() });
  }
  safeMessages.forEach((entry) => output.push(entry));
  const safePrompt = String(prompt || "").trim().slice(0, 20000);
  if (safePrompt || imageDataUrl) {
    if (imageDataUrl && /^data:image\/[a-zA-Z0-9.+-]+;base64,/i.test(imageDataUrl)) {
      const contentParts = [];
      if (safePrompt) contentParts.push({ type: "text", text: safePrompt });
      contentParts.push({ type: "image_url", image_url: { url: imageDataUrl } });
      output.push({ role: "user", content: contentParts });
    } else if (safePrompt) {
      output.push({ role: "user", content: safePrompt });
    }
  }
  return output;
}

function buildAiOllamaChatMessages({ messages, prompt, systemPrompt, imageDataUrl }) {
  const safeMessages = sanitizeAiMessages(messages);
  const output = [];
  const safeSystemPrompt = sanitizeAiSystemPrompt(systemPrompt);
  if (safeSystemPrompt) output.push({ role: "system", content: safeSystemPrompt });
  if (shouldInjectAiRuntimeDateGuard({ prompt, messages: safeMessages })) {
    output.push({ role: "system", content: buildAiRuntimeDateGuardPrompt() });
  }
  safeMessages.forEach((entry) => output.push(entry));
  const safePrompt = String(prompt || "").trim().slice(0, 20000);
  if (safePrompt || imageDataUrl) {
    const userPayload = { role: "user", content: safePrompt || "请分析这张图片" };
    if (imageDataUrl && /^data:image\/[a-zA-Z0-9.+-]+;base64,/i.test(imageDataUrl)) {
      userPayload.images = [imageDataUrl.replace(/^data:image\/[a-zA-Z0-9.+-]+;base64,/i, "")];
    }
    output.push(userPayload);
  }
  return output;
}

async function executeAiChatRequest(rawPayload) {
  const provider = String(rawPayload?.provider || "").trim().toLowerCase() === "network" ? "network" : "ollama";
  const model = sanitizeAiModelName(rawPayload?.model);
  if (!model) {
    const error = new Error("缺少模型名称");
    error.code = "MODEL_REQUIRED";
    throw error;
  }

  if (provider === "ollama") {
    const baseUrl = normalizeAiHttpUrl(rawPayload?.ollamaBaseUrl, AI_OLLAMA_DEFAULT_BASE_URL);
    const endpoint = buildAiEndpointUrl(baseUrl, "/api/chat");
    const requestPayload = {
      model,
      messages: buildAiOllamaChatMessages({
        messages: rawPayload?.messages,
        prompt: rawPayload?.prompt,
        systemPrompt: rawPayload?.systemPrompt,
        imageDataUrl: rawPayload?.imageDataUrl,
      }),
      stream: false,
      keep_alive: AI_OLLAMA_KEEP_ALIVE,
    };
    const result = await fetchAiJson(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(requestPayload),
      timeoutMs: AI_OLLAMA_CHAT_TIMEOUT_MS,
    });
    if (!result.ok) {
      const error = new Error(result?.data?.error || result?.text || `Ollama请求失败（${result.status}）`);
      error.code = "OLLAMA_CHAT_FAILED";
      error.status = result.status;
      throw error;
    }
    const text =
      String(result?.data?.message?.content || "").trim() ||
      String(result?.data?.response || "").trim() ||
      "模型未返回文本内容";
    return { text, model, provider, raw: result.data };
  }

  const network = rawPayload?.network && typeof rawPayload.network === "object" ? rawPayload.network : {};
  const baseParts = splitAiBaseUrlAndKnownEndpoint(network.baseUrl);
  const baseUrl = baseParts.baseUrl;
  if (!baseUrl) {
    const error = new Error("网络模型 Base URL 无效");
    error.code = "NETWORK_BASE_URL_INVALID";
    throw error;
  }
  const chatPath =
    baseParts.endpointType === "chat" && baseParts.endpointPath
      ? baseParts.endpointPath
      : normalizeAiHttpPath(network.chatPath, "/chat/completions");
  const endpoint = buildAiEndpointUrl(baseUrl, chatPath);
  const requestPayload = {
    model,
    messages: buildAiOpenAiChatMessages({
      messages: rawPayload?.messages,
      prompt: rawPayload?.prompt,
      systemPrompt: rawPayload?.systemPrompt,
      imageDataUrl: rawPayload?.imageDataUrl,
    }),
    stream: false,
  };
  const result = await fetchAiJson(endpoint, {
    method: "POST",
    headers: buildAiRequestHeaders(network.apiKey),
    body: JSON.stringify(requestPayload),
    timeoutMs: resolveAiChatTimeoutMs(provider, model),
  });
  if (!result.ok) {
    const errorMessage =
      result?.data?.error?.message ||
      result?.data?.error ||
      result?.data?.message ||
      result?.text ||
      `网络模型请求失败（${result.status}）`;
    const error = new Error(String(errorMessage).slice(0, 400));
    error.code = "NETWORK_CHAT_FAILED";
    error.status = result.status;
    throw error;
  }
  const text = extractTextFromOpenAiResponse(result.data) || "模型未返回文本内容";
  return { text, model, provider, raw: result.data };
}

async function requestAiImageViaChatCompletion({ model, prompt, systemPrompt, imageDataUrl, network }) {
  const baseParts = splitAiBaseUrlAndKnownEndpoint(network?.baseUrl);
  const baseUrl = baseParts.baseUrl;
  if (!baseUrl) {
    const error = new Error("网络模型 Base URL 无效");
    error.code = "NETWORK_BASE_URL_INVALID";
    throw error;
  }
  const chatPath =
    baseParts.endpointType === "chat" && baseParts.endpointPath
      ? baseParts.endpointPath
      : normalizeAiHttpPath(network?.chatPath, "/chat/completions");
  const endpoint = buildAiEndpointUrl(baseUrl, chatPath);
  const headers = buildAiRequestHeaders(network?.apiKey);
  const messages = buildAiOpenAiChatMessages({
    messages: [],
    prompt: `${String(prompt || "").trim()}\n\n请直接返回图片结果，不要只返回文字描述。`,
    systemPrompt,
    imageDataUrl,
  });
  const attemptPayloads = [
    { modalities: ["image", "text"] },
    { response_modalities: ["IMAGE", "TEXT"] },
    { modalities: ["image"] },
    { response_modalities: ["IMAGE"] },
    {},
  ];
  let lastError = null;
  for (const extras of attemptPayloads) {
    const requestPayload = {
      model,
      messages,
      stream: false,
      ...extras,
    };
    const result = await fetchAiJson(endpoint, {
      method: "POST",
      headers,
      body: JSON.stringify(requestPayload),
      timeoutMs: AI_IMAGE_REQUEST_TIMEOUT_MS,
    });
    if (!result.ok) {
      const message =
        result?.data?.error?.message ||
        result?.data?.error ||
        result?.data?.message ||
        result?.text ||
        `生图请求失败（${result.status}）`;
      const error = new Error(String(message).slice(0, 400));
      error.code = "NETWORK_IMAGE_CHAT_FALLBACK_FAILED";
      error.status = result.status;
      lastError = error;
      continue;
    }

    const imageUrl = String(extractFirstMediaUrlFromPayload(result.data, "image") || "").trim();
    const imageDataUrl = String(extractFirstImageDataUrlFromPayload(result.data) || "").trim();
    if (imageUrl || imageDataUrl) {
      return {
        provider: "network",
        model,
        imageUrl,
        imageDataUrl,
        raw: result.data,
        route: "chat-completions",
      };
    }
    const noMediaError = new Error(buildAiImageNoMediaErrorMessage(result.data, "模型未返回图片数据"));
    noMediaError.code = "NETWORK_IMAGE_CHAT_NO_MEDIA";
    noMediaError.status = 422;
    lastError = noMediaError;
  }
  if (lastError) throw lastError;
  const error = new Error("生图请求失败");
  error.code = "NETWORK_IMAGE_CHAT_FALLBACK_FAILED";
  throw error;
}

function isLikelyAiImageEditModelName(modelName) {
  const safe = sanitizeAiModelName(modelName).toLowerCase();
  if (!safe) return false;
  return /image[-_ ]?edit/i.test(safe);
}

function isAiImageEditsPath(pathText) {
  const safe = normalizeAiHttpPath(pathText, "/images/generations");
  return /\/images\/edits\/?$/i.test(safe);
}

function sanitizeAiImageEditSize(rawValue) {
  const safe = String(rawValue || "")
    .trim()
    .toLowerCase();
  if (!safe) return "";
  if (!/^\d{2,5}x\d{2,5}$/.test(safe)) return "";
  return safe;
}

function sanitizeAiImageEditTaskTypes(rawTaskTypes) {
  if (!Array.isArray(rawTaskTypes)) return [];
  const output = [];
  rawTaskTypes.forEach((entry) => {
    const safe = String(entry || "")
      .trim()
      .toLowerCase();
    if (!safe) return;
    if (!/^[a-z][a-z0-9_-]{0,31}$/.test(safe)) return;
    output.push(safe);
  });
  return output.slice(0, 8);
}

function collectAiImageEditSourceDataUrls(rawPayload) {
  const output = [];
  const seen = new Set();
  const pushItem = (value) => {
    const text = String(value || "").trim();
    if (!text || seen.has(text)) return;
    seen.add(text);
    output.push(text);
  };
  if (Array.isArray(rawPayload?.imageDataUrls)) {
    rawPayload.imageDataUrls.forEach((entry) => {
      pushItem(entry);
    });
  }
  pushItem(rawPayload?.imageDataUrl);
  return output.slice(0, 6);
}

async function collectAiImageEditSourceRemoteUrls(rawPayload) {
  const output = [];
  const seen = new Set();
  const pushUrl = (value) => {
    const safe = sanitizeAiRemoteImageUrl(value);
    if (!safe || seen.has(safe)) return;
    seen.add(safe);
    output.push(safe);
  };
  if (Array.isArray(rawPayload?.imageUrls)) {
    rawPayload.imageUrls.forEach((entry) => {
      pushUrl(entry);
    });
  }
  pushUrl(rawPayload?.imageUrl);
  pushUrl(rawPayload?.url);
  return output.slice(0, 6);
}

async function executeAiImageEditRequest({ model, prompt, network, endpoint, rawPayload }) {
  const dataUrls = collectAiImageEditSourceDataUrls(rawPayload);
  const sourceItems = [];
  dataUrls.forEach((dataUrl, index) => {
    const parsed = parseAiImageDataUrlToBuffer(dataUrl);
    if (!parsed || !parsed.buffer || !parsed.buffer.length) return;
    sourceItems.push({
      buffer: parsed.buffer,
      mimeType: parsed.mimeType || "image/png",
      fileName: `image-${index + 1}.${getAiImageExtensionByMimeType(parsed.mimeType) || "png"}`,
    });
  });

  if (!sourceItems.length) {
    const remoteUrls = await collectAiImageEditSourceRemoteUrls(rawPayload);
    for (let i = 0; i < remoteUrls.length; i += 1) {
      const remoteUrl = remoteUrls[i];
      const fetched = await executeAiImageAssetFetch({ url: remoteUrl });
      sourceItems.push({
        buffer: fetched.buffer,
        mimeType: fetched.contentType || "image/png",
        fileName: `image-${i + 1}.${getAiImageExtensionByMimeType(fetched.contentType) || "png"}`,
      });
    }
  }

  if (!sourceItems.length) {
    const error = new Error(`${sanitizeAiModelName(model) || "当前模型"} 需要至少一张图片附件`);
    error.code = "IMAGE_EDIT_IMAGE_REQUIRED";
    error.status = 400;
    throw error;
  }

  const form = new FormData();
  form.set("prompt", String(prompt || "").trim().slice(0, 12000));
  form.set("model", model);
  sourceItems.forEach((item) => {
    form.append("image", new Blob([item.buffer], { type: item.mimeType }), item.fileName);
  });

  const rawTaskTypes = sanitizeAiImageEditTaskTypes(rawPayload?.taskTypes);
  const taskTypes =
    rawTaskTypes.length > 0
      ? rawTaskTypes
      : sourceItems.length === 2
        ? ["id", "style"]
        : [];
  taskTypes.slice(0, sourceItems.length).forEach((taskType) => {
    form.append("task_types", taskType);
  });

  const size = sanitizeAiImageEditSize(rawPayload?.size);
  if (size) form.set("size", size);

  const stepsRaw = Number(rawPayload?.numInferenceSteps ?? rawPayload?.num_inference_steps);
  if (Number.isFinite(stepsRaw) && stepsRaw > 0) {
    form.set("num_inference_steps", String(Math.floor(stepsRaw)));
  } else {
    form.set("num_inference_steps", "8");
  }

  const cfgScaleRaw = Number(rawPayload?.cfgScale ?? rawPayload?.cfg_scale);
  if (Number.isFinite(cfgScaleRaw) && cfgScaleRaw > 0) {
    form.set("cfg_scale", String(cfgScaleRaw));
  } else {
    form.set("cfg_scale", "1");
  }

  const loraWeightsRaw = rawPayload?.loraWeights ?? rawPayload?.lora_weights;
  if (loraWeightsRaw && typeof loraWeightsRaw === "object") {
    form.set("lora_weights", JSON.stringify(loraWeightsRaw).slice(0, 4000));
  } else if (typeof loraWeightsRaw === "string" && loraWeightsRaw.trim()) {
    form.set("lora_weights", loraWeightsRaw.trim().slice(0, 4000));
  }

  const result = await fetchAiJson(endpoint, {
    method: "POST",
    headers: buildAiRequestAuthHeaders(network?.apiKey),
    body: form,
    timeoutMs: AI_IMAGE_EDIT_TIMEOUT_MS,
  });
  if (!result.ok) {
    const message =
      result?.data?.error?.message ||
      result?.data?.error ||
      result?.data?.message ||
      result?.text ||
      `生图编辑请求失败（${result.status}）`;
    const error = new Error(String(message).slice(0, 400));
    error.code = "NETWORK_IMAGE_EDIT_FAILED";
    error.status = result.status;
    throw error;
  }

  const firstData = Array.isArray(result?.data?.data) ? result.data.data[0] : null;
    const imageUrl = String(
      firstData?.url || firstData?.image_url || firstData?.imageUrl || firstData?.result_url || firstData?.resultUrl || extractFirstMediaUrlFromPayload(result.data, "image") || ""
    ).trim();
    const imageDataUrl =
    buildAiImageDataUrlFromBase64(
      firstData?.b64_json || firstData?.base64 || firstData?.image_base64 || firstData?.imageBase64 || firstData?.b64,
      firstData?.mime_type || firstData?.mimeType || "image/png"
    ) ||
    String(extractFirstImageDataUrlFromPayload(result.data) || "").trim();
  if (!imageUrl && !imageDataUrl) {
    const error = new Error(buildAiImageNoMediaErrorMessage(result.data, "生图编辑请求未返回图片数据"));
    error.code = "NETWORK_IMAGE_EDIT_EMPTY_RESULT";
    error.status = 422;
    throw error;
  }
  return {
    provider: "network",
    model,
    imageUrl,
    imageDataUrl,
    raw: result.data,
    route: "image-edits",
  };
}

async function executeAiImageAssetFetch(rawPayload) {
  const assetUrl = sanitizeAiRemoteImageUrl(rawPayload?.url);
  if (!assetUrl) {
    const error = new Error("图片链接无效，仅支持 http/https");
    error.code = "INVALID_IMAGE_URL";
    error.status = 400;
    throw error;
  }

  const result = await fetchAiBinary(assetUrl, {
    method: "GET",
    headers: {
      Accept: "image/*,*/*;q=0.8",
    },
    timeoutMs: AI_IMAGE_FETCH_TIMEOUT_MS,
    maxBytes: AI_IMAGE_FETCH_MAX_BYTES,
  });
  if (!result.ok) {
    const textLower = String(result.text || "").toLowerCase();
    const blockedByChallenge =
      textLower.includes("just a moment") ||
      textLower.includes("cloudflare") ||
      textLower.includes("cf-chl") ||
      textLower.includes("attention required");
    let message = "";
    if (result.status === 403 || blockedByChallenge) {
      message = "目标站点拒绝访问（403），可能启用了防盗链或人机校验";
    } else if (result.status === 404) {
      message = "图片地址不存在（404）";
    } else if (result.status === 429) {
      message = "目标站点请求过于频繁（429），请稍后重试";
    } else if (result.status >= 500) {
      message = `目标站点服务异常（${result.status}）`;
    } else {
      message = sanitizeAiErrorMessage(result.text, `远程图片获取失败（${result.status}）`);
    }
    const error = new Error(message);
    error.code = "REMOTE_IMAGE_FETCH_FAILED";
    error.status = result.status;
    throw error;
  }
  if (!result.buffer || !result.buffer.length) {
    const error = new Error("远程图片内容为空");
    error.code = "REMOTE_IMAGE_EMPTY";
    error.status = 502;
    throw error;
  }
  const contentTypeHeader = String(result.contentType || "").split(";")[0].trim().toLowerCase();
  if (contentTypeHeader && !contentTypeHeader.startsWith("image/")) {
    const error = new Error("远程资源不是图片");
    error.code = "REMOTE_IMAGE_UNSUPPORTED_TYPE";
    error.status = 415;
    throw error;
  }
  const detectedMime = detectAiImageMimeByBuffer(result.buffer);
  if (!contentTypeHeader && !detectedMime) {
    const error = new Error("远程资源不是可识别图片");
    error.code = "REMOTE_IMAGE_UNRECOGNIZED";
    error.status = 415;
    throw error;
  }
  if (looksLikeHtmlTextBuffer(result.buffer) && detectedMime !== "image/svg+xml") {
    const error = new Error("远程资源返回了网页内容，而不是图片");
    error.code = "REMOTE_IMAGE_HTML_CONTENT";
    error.status = 415;
    throw error;
  }
  const contentType = contentTypeHeader || detectedMime || "image/png";
  return {
    contentType,
    buffer: result.buffer,
  };
}

async function executeAiImageCachePersist(rawPayload) {
  const parsedDataUrl = parseAiImageDataUrlToBuffer(rawPayload?.imageDataUrl);
  if (parsedDataUrl) {
    return writeAiImageCacheFile({
      buffer: parsedDataUrl.buffer,
      mimeType: parsedDataUrl.mimeType,
      sourceUrl: "",
    });
  }

  const sourceUrl = sanitizeAiRemoteImageUrl(rawPayload?.imageUrl || rawPayload?.url);
  if (!sourceUrl) {
    const error = new Error("缺少可缓存图片来源");
    error.code = "IMAGE_CACHE_SOURCE_REQUIRED";
    error.status = 400;
    throw error;
  }
  const fetched = await executeAiImageAssetFetch({ url: sourceUrl });
  return writeAiImageCacheFile({
    buffer: fetched.buffer,
    mimeType: fetched.contentType || "image/png",
    sourceUrl,
  });
}

async function normalizeAiImageResultWithLocalCache(resultPayload) {
  if (!resultPayload || typeof resultPayload !== "object") return resultPayload;
  const imageDataUrl = String(resultPayload.imageDataUrl || "").trim();
  const imageUrl = String(resultPayload.imageUrl || "").trim();
  const existingAssetUrl = String(resultPayload.assetUrl || "").trim();
  if (existingAssetUrl) {
    return { ...resultPayload, imageUrl: existingAssetUrl, imageDataUrl: "" };
  }
  if (!imageDataUrl && !imageUrl) return resultPayload;
  try {
    const cached = await executeAiImageCachePersist({ imageDataUrl, imageUrl });
    const cachedUrl = String(cached?.assetUrl || "").trim();
    if (!cachedUrl) return resultPayload;
    return {
      ...resultPayload,
      assetUrl: cachedUrl,
      imageUrl: cachedUrl,
      imageDataUrl: "",
      cacheFileName: String(cached?.fileName || "").trim(),
      cacheBytes: Number(cached?.bytes || 0),
    };
  } catch {
    return resultPayload;
  }
}

async function executeAiImageRequest(rawPayload) {
  const provider = String(rawPayload?.provider || "").trim().toLowerCase() === "network" ? "network" : "ollama";
  if (provider === "ollama") {
    const error = new Error("本地Ollama目前不直接支持通用生图接口，请切换网络模型");
    error.code = "OLLAMA_IMAGE_UNSUPPORTED";
    throw error;
  }
  const model = sanitizeAiModelName(rawPayload?.model);
  if (!model) {
    const error = new Error("缺少生图模型名称");
    error.code = "IMAGE_MODEL_REQUIRED";
    throw error;
  }
  const prompt = String(rawPayload?.prompt || "").trim().slice(0, 12000);
  if (!prompt) {
    const error = new Error("缺少生图提示词");
    error.code = "IMAGE_PROMPT_REQUIRED";
    throw error;
  }
  const network = rawPayload?.network && typeof rawPayload.network === "object" ? rawPayload.network : {};
  const baseParts = splitAiBaseUrlAndKnownEndpoint(network.baseUrl);
  const baseUrl = baseParts.baseUrl;
  if (!baseUrl) {
    const error = new Error("网络模型 Base URL 无效");
    error.code = "NETWORK_BASE_URL_INVALID";
    throw error;
  }
  const modelLooksLikeEdit = isLikelyAiImageEditModelName(model);
  const hasImageDataInputs = collectAiImageEditSourceDataUrls(rawPayload).length > 0;
  const hasImageUrlInputs = Boolean(
    sanitizeAiRemoteImageUrl(rawPayload?.imageUrl || rawPayload?.url) ||
      (Array.isArray(rawPayload?.imageUrls) && rawPayload.imageUrls.some((entry) => sanitizeAiRemoteImageUrl(entry)))
  );
  const hasEditTaskTypes = sanitizeAiImageEditTaskTypes(rawPayload?.taskTypes).length > 0;
  let imagePath =
    baseParts.endpointType === "image" && baseParts.endpointPath
      ? baseParts.endpointPath
      : normalizeAiHttpPath(network.imagePath, "/images/generations");
  if (modelLooksLikeEdit && /^\/images\/generations\/?$/i.test(imagePath)) {
    imagePath = "/images/edits";
  }
  // If non-edit model accidentally keeps /images/edits from previous config, fall back to generations.
  if (!modelLooksLikeEdit && isAiImageEditsPath(imagePath) && !hasImageDataInputs && !hasImageUrlInputs && !hasEditTaskTypes) {
    imagePath = "/images/generations";
  }
  const endpoint = buildAiEndpointUrl(baseUrl, imagePath);
  const useImageEditsRoute =
    modelLooksLikeEdit || (isAiImageEditsPath(imagePath) && (hasImageDataInputs || hasImageUrlInputs || hasEditTaskTypes));
  if (useImageEditsRoute) {
    const editResult = await executeAiImageEditRequest({
      model,
      prompt,
      network,
      endpoint,
      rawPayload,
    });
    return normalizeAiImageResultWithLocalCache(editResult);
  }
  const requestPayload = {
    model,
    prompt,
    n: 1,
  };
  const canFallbackToChat = !modelLooksLikeEdit && isLikelyAiImageModelName(model);

  try {
    const result = await fetchAiJson(endpoint, {
      method: "POST",
      headers: buildAiRequestHeaders(network.apiKey),
      body: JSON.stringify(requestPayload),
      timeoutMs: AI_IMAGE_REQUEST_TIMEOUT_MS,
    });
    if (!result.ok) {
      const message =
        result?.data?.error?.message ||
        result?.data?.error ||
        result?.data?.message ||
        result?.text ||
        `生图请求失败（${result.status}）`;
      const error = new Error(String(message).slice(0, 400));
      error.code = "NETWORK_IMAGE_FAILED";
      error.status = result.status;
      throw error;
    }
    const firstData = Array.isArray(result?.data?.data) ? result.data.data[0] : null;
    const imageUrl = String(
      firstData?.url || firstData?.image_url || firstData?.imageUrl || firstData?.result_url || firstData?.resultUrl || extractFirstMediaUrlFromPayload(result.data, "image") || ""
    ).trim();
    const imageDataUrl =
      buildAiImageDataUrlFromBase64(
        firstData?.b64_json || firstData?.base64 || firstData?.image_base64 || firstData?.imageBase64 || firstData?.b64,
        firstData?.mime_type || firstData?.mimeType || "image/png"
      ) ||
      String(extractFirstImageDataUrlFromPayload(result.data) || "").trim();
    if (imageUrl || imageDataUrl) {
      return normalizeAiImageResultWithLocalCache({
        provider,
        model,
        imageUrl,
        imageDataUrl,
        raw: result.data,
        route: "image-generations",
      });
    }
    if (!canFallbackToChat) {
      const error = new Error(buildAiImageNoMediaErrorMessage(result.data, "生图请求未返回图片数据"));
      error.code = "NETWORK_IMAGE_EMPTY_RESULT";
      error.status = 422;
      throw error;
    }
  } catch (primaryError) {
    if (!canFallbackToChat) throw primaryError;
    try {
      const fallbackResult = await requestAiImageViaChatCompletion({
        model,
        prompt,
        systemPrompt: rawPayload?.systemPrompt,
        imageDataUrl: rawPayload?.imageDataUrl,
        network,
      });
      return normalizeAiImageResultWithLocalCache(fallbackResult);
    } catch (fallbackError) {
      const primaryMessage = sanitizeAiErrorMessage(primaryError?.message, "");
      const fallbackMessage = sanitizeAiErrorMessage(fallbackError?.message, "");
      const message = [primaryMessage, fallbackMessage].filter(Boolean).join("；");
      const error = new Error(message || "生图请求失败");
      error.code = "NETWORK_IMAGE_FAILED";
      error.status = Number.isFinite(Number(fallbackError?.status))
        ? Number(fallbackError.status)
        : Number.isFinite(Number(primaryError?.status))
          ? Number(primaryError.status)
          : 502;
      throw error;
    }
  }

  const fallbackResult = await requestAiImageViaChatCompletion({
    model,
    prompt,
    systemPrompt: rawPayload?.systemPrompt,
    imageDataUrl: rawPayload?.imageDataUrl,
    network,
  });
  return normalizeAiImageResultWithLocalCache(fallbackResult);
}

async function executeAiConnectionTest(rawPayload) {
  const provider = String(rawPayload?.provider || "").trim().toLowerCase() === "network" ? "network" : "ollama";
  if (provider === "ollama") {
    const models = await fetchOllamaModelTags(rawPayload?.ollamaBaseUrl || AI_OLLAMA_DEFAULT_BASE_URL);
    return {
      provider,
      message: models.length ? `本地Ollama连接成功，检测到 ${models.length} 个模型` : "本地Ollama连接成功，但未检测到模型",
      models,
    };
  }

  const network = rawPayload?.network && typeof rawPayload.network === "object" ? rawPayload.network : {};
  const baseParts = splitAiBaseUrlAndKnownEndpoint(network.baseUrl);
  const baseUrl = baseParts.baseUrl;
  if (!baseUrl) {
    const error = new Error("网络模型 Base URL 无效");
    error.code = "NETWORK_BASE_URL_INVALID";
    throw error;
  }
  const modelsPath =
    baseParts.endpointType === "models" && baseParts.endpointPath ? baseParts.endpointPath : "/models";
  const endpoint = buildAiEndpointUrl(baseUrl, modelsPath);
  const result = await fetchAiJson(endpoint, {
    method: "GET",
    headers: buildAiRequestHeaders(network.apiKey),
    timeoutMs: 15000,
  });
  if (!result.ok) {
    const message =
      result?.data?.error?.message ||
      result?.data?.error ||
      result?.data?.message ||
      result?.text ||
      `连接失败（${result.status}）`;
    const error = new Error(String(message).slice(0, 400));
    error.code = "NETWORK_TEST_FAILED";
    error.status = result.status;
    throw error;
  }
  const modelCount = Array.isArray(result?.data?.data) ? result.data.data.length : 0;
  return {
    provider,
    message: modelCount > 0 ? `网络模型连接成功，检测到 ${modelCount} 个模型` : "网络模型连接成功",
    models: Array.isArray(result?.data?.data)
      ? result.data.data
          .map((item) => String(item?.id || "").trim())
          .filter(Boolean)
          .slice(0, 50)
      : [],
  };
}

function writeAiUpstreamError(res, error, fallbackMessage) {
  if (error && error.code === "REQUEST_TIMEOUT") {
    writeJson(res, 504, { error: error.message || "AI 请求超时" });
    return;
  }
  if (error && error.code === "REQUEST_ABORTED") {
    writeJson(res, 408, { error: error.message || "AI 请求已取消" });
    return;
  }
  const status = Number(error?.status);
  if (Number.isFinite(status) && status >= 400 && status < 600) {
    writeJson(res, status, { error: error?.message || fallbackMessage });
    return;
  }
  writeJson(res, 502, { error: error?.message || fallbackMessage });
}

async function handleApi(req, res, pathname) {
  if (pathname === "/api/ai/ollama/tags") {
    if (req.method !== "GET") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const parsedUrl = new url.URL(req.url || "/", `http://${req.headers.host || `${HOST}:${PORT}`}`);
      const baseUrl = parsedUrl.searchParams.get("baseUrl") || AI_OLLAMA_DEFAULT_BASE_URL;
      const models = await fetchOllamaModelTags(baseUrl);
      writeJson(res, 200, { ok: true, baseUrl: normalizeAiHttpUrl(baseUrl, AI_OLLAMA_DEFAULT_BASE_URL), models });
    } catch (error) {
      writeAiUpstreamError(res, error, "Failed to load Ollama models");
    }
    return true;
  }

  if (pathname === "/api/ai/test-connection") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req, 1024 * 1024 * 4);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const result = await executeAiConnectionTest(parsed);
      writeJson(res, 200, { ok: true, ...result });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else {
        writeAiUpstreamError(res, error, "Connection test failed");
      }
    }
    return true;
  }

  if (pathname === "/api/ai/search") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req, 1024 * 1024 * 2);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const result = await executeAiWebSearchRequest(parsed);
      writeJson(res, 200, { ok: true, ...result });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else if (error && error.code === "SEARCH_QUERY_REQUIRED") {
        writeJson(res, 400, { error: error.message || "Invalid search query" });
      } else {
        writeAiUpstreamError(res, error, "AI web search request failed");
      }
    }
    return true;
  }

  if (pathname === "/api/ai/research") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req, 1024 * 1024 * 2);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const result = await executeAiResearchRequest(parsed);
      writeJson(res, 200, { ok: true, ...result });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else if (error && error.code === "SEARCH_QUERY_REQUIRED") {
        writeJson(res, 400, { error: error.message || "Invalid research query" });
      } else {
        writeAiUpstreamError(res, error, "AI research request failed");
      }
    }
    return true;
  }

  if (pathname === "/api/ai/browser-agent") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req, 1024 * 1024 * 2);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const result = await executeAiBrowserAgentRequest(parsed);
      writeJson(res, 200, { ok: true, ...result });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else if (
        error &&
        ["BROWSER_AGENT_TARGET_REQUIRED", "BROWSER_AGENT_UNAVAILABLE", "BROWSER_AGENT_LAUNCH_FAILED", "BROWSER_AGENT_NAV_FAILED"].includes(
          error.code
        )
      ) {
        writeJson(res, Number(error.status) || 400, { error: error.message || "Invalid browser-agent payload", code: error.code });
      } else {
        writeAiUpstreamError(res, error, "AI browser-agent request failed");
      }
    }
    return true;
  }

  if (pathname === "/api/ai/local-agent") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req, 1024 * 1024 * 2);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const result = await executeAiLocalAgentRequest(parsed);
      writeJson(res, 200, { ok: true, ...result });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else if (error && String(error.code || "").startsWith("LOCAL_AGENT_")) {
        writeJson(res, Number(error.status) || 400, { error: error.message || "Invalid local-agent payload", code: error.code });
      } else if (error && ["INVALID_PATH", "PATH_NOT_FOUND", "UNSUPPORTED_PATH_TYPE"].includes(error.code)) {
        writeJson(res, 400, { error: error.message || "Invalid local path", code: error.code });
      } else {
        writeAiUpstreamError(res, error, "AI local-agent request failed");
      }
    }
    return true;
  }

  if (pathname === "/api/ai/webpage") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req, 1024 * 1024 * 2);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const result = await executeAiWebPageExtractRequest(parsed);
      writeJson(res, 200, { ok: true, ...result });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else if (error && ["WEBPAGE_URL_INVALID", "WEBPAGE_TEXT_EMPTY"].includes(error.code)) {
        writeJson(res, Number(error.status) || 400, { error: error.message || "Invalid webpage payload" });
      } else {
        writeAiUpstreamError(res, error, "AI webpage extract request failed");
      }
    }
    return true;
  }

  if (pathname === "/api/ai/attachments/parse") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const requestUrl = new url.URL(req.url || "/", `http://${req.headers.host || `${HOST}:${PORT}`}`);
      const contentType = String(req.headers["content-type"] || "")
        .trim()
        .toLowerCase();
      let parsed = {};
      if (contentType.startsWith("application/octet-stream")) {
        const binaryBuffer = await readRequestBuffer(req, AI_ATTACHMENT_PARSE_MAX_BYTES);
        const includeImagesRaw = String(requestUrl.searchParams.get("includeImages") || "").trim().toLowerCase();
        parsed = {
          name: sanitizeAiAttachmentName(requestUrl.searchParams.get("name") || ""),
          type: String(requestUrl.searchParams.get("type") || "").trim().toLowerCase(),
          includeImages: includeImagesRaw ? !["0", "false", "no", "off"].includes(includeImagesRaw) : true,
          buffer: binaryBuffer,
        };
      } else {
        const rawBody = await readRequestBody(req, AI_ATTACHMENT_PARSE_REQUEST_MAX_BYTES);
        parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      }
      const result = await executeAiAttachmentParseRequest(parsed);
      writeJson(res, 200, { ok: true, ...result });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else if (
        error &&
        [
          "ATTACHMENT_DATA_REQUIRED",
          "ATTACHMENT_DATA_INVALID",
          "ATTACHMENT_TOO_LARGE",
          "ATTACHMENT_PARSE_UNSUPPORTED",
          "ATTACHMENT_PARSE_FAILED",
          "ATTACHMENT_PARSE_EMPTY",
        ].includes(error.code)
      ) {
        writeJson(res, 400, { error: error.message || "Invalid attachment payload" });
      } else {
        writeAiUpstreamError(res, error, "AI attachment parse request failed");
      }
    }
    return true;
  }

  if (pathname === "/api/ai/chat") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req, 1024 * 1024 * 10);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const requestUrl = new url.URL(req.url || "/", `http://${req.headers.host || `${HOST}:${PORT}`}`);
      const streamFlag = String(requestUrl.searchParams.get("stream") || "")
        .trim()
        .toLowerCase();
      const useStream = streamFlag === "1" || streamFlag === "true" || streamFlag === "yes";
      if (useStream) {
        await streamAiChatRequest(parsed, req, res);
        return true;
      }
      const result = await executeAiChatRequest(parsed);
      writeJson(res, 200, { ok: true, ...result });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else {
        writeAiUpstreamError(res, error, "AI chat request failed");
      }
    }
    return true;
  }

  if (pathname === "/api/ai/image") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req, 1024 * 1024 * 32);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const result = await executeAiImageRequest(parsed);
      writeJson(res, 200, { ok: true, ...result });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else {
        writeAiUpstreamError(res, error, "AI image request failed");
      }
    }
    return true;
  }

  if (pathname === "/api/ai/image/fetch") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req, 1024 * 64);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const result = await executeAiImageAssetFetch(parsed);
      res.writeHead(200, {
        "Content-Type": result.contentType || "image/png",
        "Cache-Control": "no-store",
        "Content-Length": result.buffer.length,
      });
      res.end(result.buffer);
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else {
        writeAiUpstreamError(res, error, "AI image asset fetch failed");
      }
    }
    return true;
  }

  if (pathname === "/api/ai/image/cache") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req, AI_IMAGE_CACHE_BODY_MAX_BYTES);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const result = await executeAiImageCachePersist(parsed);
      writeJson(res, 200, { ok: true, ...result });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else {
        writeAiUpstreamError(res, error, "AI image cache persist failed");
      }
    }
    return true;
  }

  if (pathname === "/api/gaozhi") {
    if (req.method !== "GET") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    const documents = await listGaozhiMarkdownFiles();
    writeJson(res, 200, { files: documents.map((item) => item.path), documents });
    return true;
  }

  if (pathname === "/api/fenxiang") {
    if (req.method !== "GET") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    const documents = await listTutorialMarkdownFiles();
    writeJson(res, 200, { files: documents.map((item) => item.path), documents });
    return true;
  }

  if (pathname === "/api/local-shortcuts") {
    if (req.method === "GET") {
      const shortcuts = await readLocalShortcuts();
      writeJson(res, 200, shortcuts);
      return true;
    }
    if (req.method === "POST") {
      try {
        const rawBody = await readRequestBody(req);
        const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
        const shortcuts = await writeLocalShortcuts(parsed);
        writeJson(res, 200, { ok: true, ...shortcuts });
      } catch (error) {
        if (error && error.code === "PAYLOAD_TOO_LARGE") {
          writeJson(res, 413, { error: "Payload Too Large" });
        } else if (error instanceof SyntaxError) {
          writeJson(res, 400, { error: "Invalid JSON" });
        } else {
          writeJson(res, 500, { error: "Failed to save local shortcuts" });
        }
      }
      return true;
    }
    writeJson(res, 405, { error: "Method Not Allowed" });
    return true;
  }

  if (pathname === "/api/local-open") {
    if (req.method !== "POST") {
      writeJson(res, 405, { error: "Method Not Allowed" });
      return true;
    }
    try {
      const rawBody = await readRequestBody(req);
      const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
      const category = sanitizeLocalShortcutCategory(parsed?.category);
      const id = typeof parsed?.id === "string" ? parsed.id.trim() : "";
      const directPath = normalizeAbsoluteLocalPath(parsed?.path ?? "");
      const hasValidId = /^[a-zA-Z0-9_-]{6,80}$/.test(id);
      if (!category || (!directPath && !hasValidId)) {
        writeJson(res, 400, { error: "Invalid shortcut identifier" });
        return true;
      }

      let item;
      if (directPath) {
        await openLocalPath(directPath);
        item = {
          id: hasValidId ? id : "",
          name: inferShortcutName(directPath),
        };
      } else {
        const shortcuts = await readLocalShortcuts();
        item = await openLocalShortcut(shortcuts, category, id);
      }
      writeJson(res, 200, { ok: true, id: item.id, name: item.name });
    } catch (error) {
      if (error && error.code === "PAYLOAD_TOO_LARGE") {
        writeJson(res, 413, { error: "Payload Too Large" });
      } else if (error instanceof SyntaxError) {
        writeJson(res, 400, { error: "Invalid JSON" });
      } else if (error && error.code === "SHORTCUT_NOT_FOUND") {
        writeJson(res, 404, { error: "Shortcut not found" });
      } else if (error && (error.code === "INVALID_PATH" || error.code === "PATH_NOT_FOUND" || error.code === "UNSUPPORTED_PATH_TYPE")) {
        writeJson(res, 400, { error: error.message || "Invalid path" });
      } else {
        const errorMessage = error && error.message ? error.message : "Failed to open local shortcut";
        console.error("[local-open] Failed:", errorMessage);
        writeJson(res, 500, { error: "Failed to open local shortcut" });
      }
    }
    return true;
  }

  if (pathname === "/api/visit-durations") {
    if (req.method === "GET") {
      const dailyTotals = await readVisitDailyTotals();
      writeJson(res, 200, { dailyTotals });
      return true;
    }

    if (req.method === "POST") {
      try {
        const rawBody = await readRequestBody(req);
        const parsed = rawBody.trim() ? JSON.parse(rawBody) : {};
        const dailyTotals = await writeVisitDailyTotals(parsed?.dailyTotals ?? parsed);
        writeJson(res, 200, { ok: true, dailyTotals });
      } catch (error) {
        if (error && error.code === "PAYLOAD_TOO_LARGE") {
          writeJson(res, 413, { error: "Payload Too Large" });
        } else if (error instanceof SyntaxError) {
          writeJson(res, 400, { error: "Invalid JSON" });
        } else {
          writeJson(res, 500, { error: "Failed to save visit durations" });
        }
      }
      return true;
    }

    writeJson(res, 405, { error: "Method Not Allowed" });
    return true;
  }

  return false;
}

function streamFile(filePath, res) {
  fs.stat(filePath, (statErr, stat) => {
    if (statErr || !stat.isFile()) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not Found");
      return;
    }

    const ext = path.extname(filePath).toLowerCase();
    const contentType = MIME_MAP[ext] || "application/octet-stream";
    res.writeHead(200, {
      "Content-Type": contentType,
      "Content-Length": stat.size,
      "Cache-Control": "no-store",
      "Last-Modified": stat.mtime.toUTCString(),
    });
    const stream = fs.createReadStream(filePath);
    stream.on("error", () => {
      res.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Internal Server Error");
    });
    stream.pipe(res);
  });
}

const server = http.createServer(async (req, res) => {
  const parsedUrl = new url.URL(req.url || "/", `http://${req.headers.host || `${HOST}:${PORT}`}`);
  const pathname = parsedUrl.pathname || "/";

  if (await handleApi(req, res, pathname)) {
    return;
  }

  if (req.method !== "GET" && req.method !== "HEAD") {
    res.writeHead(405, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Method Not Allowed");
    return;
  }

  const filePath = toSafeFilePath(pathname);
  if (!filePath) {
    res.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Forbidden");
    return;
  }

  streamFile(filePath, res);
});

server.listen(PORT, HOST, () => {
  console.log(`991X local server running at http://${HOST}:${PORT}`);
  console.log("告知中心动态API: /api/gaozhi");
  console.log("分享收获动态API: /api/fenxiang");
  console.log("本地快捷API: /api/local-shortcuts");
  console.log("本地打开API: /api/local-open");
  console.log("浏览历时持久化API: /api/visit-durations");
  console.log(
    "AI助手API: /api/ai/ollama/tags | /api/ai/test-connection | /api/ai/search | /api/ai/research | /api/ai/browser-agent | /api/ai/local-agent | /api/ai/webpage | /api/ai/attachments/parse | /api/ai/chat | /api/ai/chat?stream=1 | /api/ai/image | /api/ai/image/fetch | /api/ai/image/cache"
  );
});
