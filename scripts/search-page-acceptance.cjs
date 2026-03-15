const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");
const { chromium } = require("playwright-core");

const PROJECT_ROOT = path.join(__dirname, "..");
const APP_URL = "http://127.0.0.1:9910/";
const ARTIFACT_ROOT = path.join(PROJECT_ROOT, "artifacts", "search-page-acceptance");
const DEFAULT_VIEWPORT = Object.freeze({ width: 1600, height: 960 });

function parseRunOptions(argv = process.argv.slice(2)) {
  const options = {
    viewport: { ...DEFAULT_VIEWPORT },
    runTag: "desktop",
  };
  argv.forEach((arg) => {
    const safe = String(arg || "").trim();
    if (!safe) return;
    if (safe === "--mobile") {
      options.viewport = { width: 390, height: 844 };
      options.runTag = "mobile-390x844";
      return;
    }
    if (safe.startsWith("--viewport=")) {
      const value = safe.slice("--viewport=".length);
      const matched = value.match(/^(\d{2,5})x(\d{2,5})$/i);
      if (!matched) return;
      const width = Number(matched[1]);
      const height = Number(matched[2]);
      if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) return;
      options.viewport = { width, height };
      options.runTag = `${width}x${height}`;
      return;
    }
    if (safe.startsWith("--tag=")) {
      const value = safe.slice("--tag=".length).trim().replace(/[^\w.-]+/g, "-");
      if (value) options.runTag = value;
    }
  });
  return options;
}

function pad2(value) {
  return String(value).padStart(2, "0");
}

function stampNow(date = new Date()) {
  return `${date.getFullYear()}${pad2(date.getMonth() + 1)}${pad2(date.getDate())}-${pad2(date.getHours())}${pad2(date.getMinutes())}${pad2(
    date.getSeconds()
  )}`;
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function resolveChromePath() {
  const candidates = [
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
    "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
  ];
  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) return candidate;
  }
  throw new Error("未找到可用浏览器（Chrome/Edge），无法执行搜索页验收");
}

async function checkServerReadyOnce() {
  try {
    const response = await fetch(APP_URL, { method: "GET" });
    return Boolean(response && response.ok);
  } catch {
    return false;
  }
}

async function waitForServerReady(timeoutMs = 30000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    if (await checkServerReadyOnce()) return;
    await delay(400);
  }
  throw new Error(`服务启动超时（${timeoutMs}ms）: ${APP_URL}`);
}

function startServerProcess() {
  const server = spawn(process.execPath, ["server.js"], {
    cwd: PROJECT_ROOT,
    stdio: ["ignore", "pipe", "pipe"],
    windowsHide: true,
  });
  let bootLog = "";
  const appendLog = (chunk) => {
    const text = String(chunk || "");
    if (!text) return;
    bootLog += text;
    if (bootLog.length > 20000) bootLog = bootLog.slice(-20000);
  };
  server.stdout.on("data", appendLog);
  server.stderr.on("data", appendLog);
  return { server, getBootLog: () => bootLog };
}

async function ensureSearchView(page) {
  await clickMainMenuButton(page, "#menuSearchButton");
  await page.waitForFunction(() => document.body?.dataset?.mainView === "search");
  await page.waitForSelector("#portalSearchCanvas:not([hidden])");
}

async function registerIfNeeded(page, username, password) {
  await clickMainMenuButton(page, "#menuWebsiteButton");
  await page.waitForSelector("#websiteCenter:not([hidden])");
  const loginButton = await page.$("#websiteAuthPlaceholder .feature-login-button");
  if (!loginButton) return { skipped: true };
  await loginButton.click();
  await page.waitForSelector("#notesAuthModal:not([hidden])");
  await page.click("#notesTabRegister");
  await page.fill("#notesRegisterUsername", username);
  await page.fill("#notesRegisterPassword", password);
  await page.fill("#notesRegisterConfirmPassword", password);
  await page.click("#notesRegisterForm button[type='submit']");
  await page.waitForFunction(() => {
    const modal = document.getElementById("notesAuthModal");
    return Boolean(modal && modal.hidden);
  });
  return { skipped: false };
}

async function clickMainMenuButton(page, selector) {
  const ensureVisible = async () => {
    const target = page.locator(selector);
    if (await target.isVisible().catch(() => false)) return true;
    const mobileTrigger = page.locator("#mobileTrigger");
    if (await mobileTrigger.isVisible().catch(() => false)) {
      await mobileTrigger.click();
      await delay(120);
    }
    return target.isVisible().catch(() => false);
  };
  const visible = await ensureVisible();
  if (!visible) throw new Error(`菜单按钮不可见：${selector}`);
  await page.click(selector);
}

async function run() {
  const runOptions = parseRunOptions();
  ensureDir(ARTIFACT_ROOT);
  const runDir = path.join(ARTIFACT_ROOT, `${stampNow()}-${runOptions.runTag}`);
  ensureDir(runDir);

  const checkpoints = [];
  const addCheckpoint = (payload) => checkpoints.push(payload);
  const screenshotOf = (name) => path.join(runDir, name);

  let serverControl = null;
  let browser = null;
  let context = null;
  try {
    const hasExternalServer = await checkServerReadyOnce();
    if (!hasExternalServer) {
      serverControl = startServerProcess();
      await waitForServerReady();
    }

    browser = await chromium.launch({
      executablePath: resolveChromePath(),
      headless: true,
      args: ["--disable-gpu"],
    });
    context = await browser.newContext({
      viewport: runOptions.viewport,
      locale: "zh-CN",
    });
    const page = await context.newPage();

    const dialogLog = [];
    page.on("dialog", async (dialog) => {
      dialogLog.push({
        type: dialog.type(),
        message: String(dialog.message() || "").trim(),
      });
      await dialog.accept();
    });

    await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
    await page.waitForSelector("#menuSearchButton", { state: "attached" });
    await ensureSearchView(page);

    // 1) 引擎下拉键盘交互
    await page.focus("#portalSearchEngineDropdownBtn");
    const beforeEngine = String((await page.textContent("#portalSearchCurrentEngineName")) || "").trim();
    await page.keyboard.press("Enter");
    await page.waitForFunction(() => document.getElementById("portalSearchEngineDropdown")?.classList.contains("open"));
    await delay(320);
    const engineOptionCount = await page.$$eval("#portalSearchEngineDropdownMenu .portal-search-engine-dropdown-item", (nodes) => nodes.length);
    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Enter");
    await delay(120);
    const afterEngine = String((await page.textContent("#portalSearchCurrentEngineName")) || "").trim();
    const dropdownClosed = !(await page.evaluate(() => document.getElementById("portalSearchEngineDropdown")?.classList.contains("open")));
    const s1 = screenshotOf("01-engine-keyboard.png");
    await page.screenshot({ path: s1, fullPage: true });
    addCheckpoint({
      point: "引擎下拉键盘交互",
      expected: "支持 Enter 打开、方向键切换、Enter 选择后关闭",
      actual: `选项数=${engineOptionCount}；切换前=${beforeEngine}，切换后=${afterEngine}，已关闭=${dropdownClosed}`,
      screenshot: s1,
      passed: dropdownClosed && (engineOptionCount <= 1 || beforeEngine !== afterEngine),
    });

    // 2) 空关键词提示（notice）
    await page.fill("#portalSearchInput", "");
    await page.click("#portalSearchButton");
    await page.waitForSelector("#portalSearchNotice:not([hidden])", { timeout: 5000 });
    const emptyNotice = String((await page.textContent("#portalSearchNotice")) || "").trim();
    const emptyPanelVisible = await page.evaluate(() => {
      const panel = document.getElementById("portalSearchResultPanel");
      return Boolean(panel && panel.classList.contains("show"));
    });
    const s2 = screenshotOf("02-empty-state.png");
    await page.screenshot({ path: s2, fullPage: true });
    addCheckpoint({
      point: "空关键词提示",
      expected: "网络模式空关键词时显示 notice，且不展示结果面板",
      actual: `notice=${emptyNotice}；结果面板可见=${emptyPanelVisible}`,
      screenshot: s2,
      passed: emptyNotice.includes("请输入关键词再搜索") && !emptyPanelVisible,
    });

    // 3) 网络搜索触发行为（不展示 loading 面板）
    await page.fill("#portalSearchInput", "acceptance");
    const popupPromise = page.waitForEvent("popup", { timeout: 2500 }).catch(() => null);
    await page.click("#portalSearchButton");
    const popup = await popupPromise;
    if (popup) {
      try {
        await popup.close();
      } catch {
        // ignore popup close errors
      }
    }
    await delay(260);
    const loadingStateVisible = await page.evaluate(() => {
      const node = document.querySelector("#portalSearchResultList .portal-search-result-state.is-loading");
      return Boolean(node);
    });
    const webPanelVisible = await page.evaluate(() => {
      const panel = document.getElementById("portalSearchResultPanel");
      return Boolean(panel && panel.classList.contains("show"));
    });
    const s3 = screenshotOf("03-loading-state.png");
    await page.screenshot({ path: s3, fullPage: true });
    addCheckpoint({
      point: "网络搜索无 loading 面板",
      expected: "网络搜索触发后不显示结果面板 loading 态",
      actual: `loading可见=${loadingStateVisible}；结果面板可见=${webPanelVisible}`,
      screenshot: s3,
      passed: !loadingStateVisible && !webPanelVisible,
    });

    // 4) 设置面板焦点约束
    const accountUsername = `search_acc_${Date.now().toString().slice(-7)}`;
    const accountPassword = "SearchPass#2026";
    await registerIfNeeded(page, accountUsername, accountPassword);
    await ensureSearchView(page);
    await page.waitForSelector("#portalSearchOpenSettings:not([hidden])", { timeout: 8000 });
    await page.click("#portalSearchOpenSettings");
    await page.waitForSelector("#portalSearchSettingsPanel.is-open", { timeout: 5000 });
    for (let i = 0; i < 16; i += 1) {
      await page.keyboard.press("Tab");
    }
    const focusInsidePanel = await page.evaluate(() => {
      const panel = document.getElementById("portalSearchSettingsPanel");
      return Boolean(panel && panel.contains(document.activeElement));
    });
    const s4 = screenshotOf("04-settings-open-focus-trap.png");
    await page.screenshot({ path: s4, fullPage: true });
    await page.keyboard.press("Escape");
    await page.waitForFunction(() => {
      const panel = document.getElementById("portalSearchSettingsPanel");
      return Boolean(panel && panel.hidden);
    });
    const focusReturned = await page.evaluate(() => document.activeElement?.id === "portalSearchOpenSettings");
    const s4b = screenshotOf("05-settings-close-focus-return.png");
    await page.screenshot({ path: s4b, fullPage: true });
    addCheckpoint({
      point: "设置面板焦点约束",
      expected: "Tab 焦点不逃逸，Esc 关闭后焦点回到设置按钮",
      actual: `焦点在面板内=${focusInsidePanel}，焦点回归=${focusReturned}`,
      screenshot: s4,
      screenshot2: s4b,
      passed: focusInsidePanel && focusReturned,
    });

    // 5) 本地索引取消 + 命中高亮/来源标签
    await ensureSearchView(page);
    const fnAvailability = await page.evaluate(() => ({
      build: typeof window.portalSearchBuildLocalIndex,
      cancel: typeof window.portalSearchRequestCancelIndexing,
      perform: typeof window.portalSearchPerformSearch,
      renderMode: typeof window.portalSearchRenderModeSwitch,
    }));
    if (
      fnAvailability.build !== "function" ||
      fnAvailability.cancel !== "function" ||
      fnAvailability.perform !== "function" ||
      fnAvailability.renderMode !== "function"
    ) {
      throw new Error(`搜索页关键函数未暴露：${JSON.stringify(fnAvailability)}`);
    }

    const localAcceptance = await page.evaluate(async () => {
      const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
      const totalFiles = 220;
      const makeFileHandle = (name, text, delayMs) => ({
        kind: "file",
        name,
        async getFile() {
          await wait(delayMs);
          return new File([text], name, {
            type: "text/plain;charset=utf-8",
            lastModified: Date.now(),
          });
        },
      });
      const entries = [makeFileHandle("alpha-main.txt", "alpha body hit alpha beta gamma", 9)];
      for (let i = 1; i < totalFiles; i += 1) {
        entries.push(makeFileHandle(`doc-${i}.txt`, `alpha payload file ${i}`, 5));
      }
      const directoryHandle = {
        kind: "directory",
        name: "mock-search-acceptance",
        async *values() {
          for (const entry of entries) {
            yield entry;
          }
        },
      };

      portalSearchState.supportsEnhancedLocal = true;
      portalSearchState.directoryHandle = directoryHandle;
      portalSearchState.currentMode = "local";
      portalSearchRenderModeSwitch();

      const buildPromise = portalSearchBuildLocalIndex();
      window.setTimeout(() => {
        try {
          portalSearchRequestCancelIndexing();
          window.setTimeout(() => {
            try {
              portalSearchRequestCancelIndexing();
            } catch {
              // ignore secondary cancel errors
            }
          }, 140);
        } catch {
          // ignore cancel errors
        }
      }, 180);
      await buildPromise;

      const statusText = String(document.getElementById("portalSearchLocalStatus")?.textContent || "").trim();
      const indexedCount = Number(portalSearchState.localIndex?.length || 0);
      const cancelDisabled = Boolean(document.getElementById("portalSearchCancelIndexBtn")?.disabled);

      const input = document.getElementById("portalSearchInput");
      if (input) input.value = "alpha";
      portalSearchPerformSearch();
      await wait(140);

      const firstCard = document.querySelector("#portalSearchResultList .portal-search-result-item");
      const highlightCount = firstCard ? firstCard.querySelectorAll(".portal-search-result-highlight").length : 0;
      const sourceTags = firstCard
        ? Array.from(firstCard.querySelectorAll(".portal-search-result-source-tag"))
            .map((node) => String(node.textContent || "").trim())
            .filter(Boolean)
        : [];
      return {
        statusText,
        indexedCount,
        cancelDisabled,
        highlightCount,
        sourceTags,
      };
    });

    const s5 = screenshotOf("06-local-cancel-highlight.png");
    await page.screenshot({ path: s5, fullPage: true });
    addCheckpoint({
      point: "本地索引取消与结果增强",
      expected: "支持取消索引；结果高亮关键词并展示命中来源标签",
      actual: `status=${localAcceptance.statusText}; indexed=${localAcceptance.indexedCount}; cancelDisabled=${localAcceptance.cancelDisabled}; highlights=${localAcceptance.highlightCount}; tags=${localAcceptance.sourceTags.join(
        "、"
      )}`,
      screenshot: s5,
      passed:
        localAcceptance.statusText.includes("索引已取消") &&
        localAcceptance.indexedCount > 0 &&
        localAcceptance.cancelDisabled &&
        localAcceptance.highlightCount > 0 &&
        localAcceptance.sourceTags.length > 0,
    });

    const summary = {
      runAt: new Date().toISOString(),
      url: APP_URL,
      checkpoints,
      allPassed: checkpoints.every((item) => item.passed),
      dialogLog,
    };
    const summaryPath = path.join(runDir, "summary.json");
    fs.writeFileSync(summaryPath, JSON.stringify(summary, null, 2), "utf8");
    console.log(JSON.stringify({ ok: summary.allPassed, runDir, summaryPath, checkpoints }, null, 2));
    if (!summary.allPassed) {
      process.exitCode = 1;
    }
  } finally {
    if (context) {
      try {
        await context.close();
      } catch {
        // ignore
      }
    }
    if (browser) {
      try {
        await browser.close();
      } catch {
        // ignore
      }
    }
    if (serverControl && serverControl.server && !serverControl.server.killed) {
      try {
        serverControl.server.kill();
      } catch {
        // ignore
      }
    }
  }
}

run().catch((error) => {
  console.error("[search-page-acceptance] failed:", error && error.stack ? error.stack : error);
  process.exit(1);
});
