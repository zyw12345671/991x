const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");
const { chromium } = require("playwright-core");

const PROJECT_ROOT = path.join(__dirname, "..");
const APP_URL = "http://127.0.0.1:9910/";
const ARTIFACT_ROOT = path.join(PROJECT_ROOT, "artifacts", "favorites-layout-acceptance");

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
  throw new Error("未找到可用浏览器（Chrome/Edge）");
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

async function importJsonViaTopbarButton(page, payloadObject, filename = "import.json") {
  const fileChooserPromise = page.waitForEvent("filechooser");
  await page.click("#websiteImportButton");
  const fileChooser = await fileChooserPromise;
  await fileChooser.setFiles({
    name: filename,
    mimeType: "application/json",
    buffer: Buffer.from(JSON.stringify(payloadObject, null, 2)),
  });
}

async function run() {
  ensureDir(ARTIFACT_ROOT);
  const runDir = path.join(ARTIFACT_ROOT, stampNow());
  ensureDir(runDir);

  const accountUsername = `fav_layout_${Date.now().toString().slice(-7)}`;
  const accountPassword = "LayoutPass#2026";
  const importPayload = {
    source: "991x-protected-feature-favorites-content",
    feature: "favorites",
    items: Array.from({ length: 31 }, (_, index) => {
      const no = index + 1;
      return {
        title: `布局验收内容${no}`,
        link: `https://layout-check-${no}.example.com`,
      };
    }),
  };

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
      viewport: { width: 1600, height: 960 },
      locale: "zh-CN",
    });
    const page = await context.newPage();

    const dialogLog = [];
    page.on("dialog", async (dialog) => {
      const record = {
        type: dialog.type(),
        message: String(dialog.message() || "").trim(),
      };
      dialogLog.push(record);
      if (dialog.type() === "prompt") {
        await dialog.accept("2");
      } else {
        await dialog.accept();
      }
    });

    await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
    await page.waitForSelector("#menuWebsiteButton", { state: "visible" });

    await page.click("#menuWebsiteButton");
    await page.waitForSelector("#websiteCenter:not([hidden])");
    await page.click("#websiteAuthPlaceholder .feature-login-button");
    await page.waitForSelector("#notesAuthModal:not([hidden])");
    await page.click("#notesTabRegister");
    await page.fill("#notesRegisterUsername", accountUsername);
    await page.fill("#notesRegisterPassword", accountPassword);
    await page.fill("#notesRegisterConfirmPassword", accountPassword);
    await page.click("#notesRegisterForm button[type='submit']");
    await page.waitForFunction(() => {
      const modal = document.getElementById("notesAuthModal");
      return Boolean(modal && modal.hidden);
    });

    await page.click("#menuFavoritesButton");
    await page.waitForFunction(() => document.body?.dataset?.mainView === "favorites");
    await page.waitForSelector("#websiteWorkspace:not([hidden])");

    await importJsonViaTopbarButton(page, importPayload, "favorites-layout-31.json");
    await page.waitForFunction(() => {
      const cards = document.querySelectorAll("#websiteCollectionGrid .website-collection-card");
      return cards.length >= 31;
    });

    const metrics = await page.evaluate(() => {
      const canvas = document.getElementById("websiteCollectionCanvas");
      const grid = document.getElementById("websiteCollectionGrid");
      if (!canvas || !grid) return null;
      const canvasStyle = window.getComputedStyle(canvas);
      const gridStyle = window.getComputedStyle(grid);
      const columnCount = String(gridStyle.gridTemplateColumns || "")
        .split(/\s+/)
        .filter(Boolean).length;
      const visibleRows = Number.parseFloat(String(canvasStyle.getPropertyValue("--website-collection-visible-rows") || "").trim() || "0");
      const rowHeight = Number.parseFloat(String(canvasStyle.getPropertyValue("--website-collection-row-height") || "").trim() || "0");
      const rowGap = Number.parseFloat(String(canvasStyle.getPropertyValue("--website-collection-row-gap") || "").trim() || "0");
      const totalCards = grid.querySelectorAll(".website-collection-card").length;
      const clientHeight = grid.clientHeight;
      const scrollHeight = grid.scrollHeight;
      const hasVerticalScrollbar = scrollHeight > clientHeight + 1;
      const overflowY = gridStyle.overflowY;
      return {
        currentView: document.body?.dataset?.mainView || "",
        columnCount,
        visibleRows,
        rowHeight,
        rowGap,
        totalCards,
        clientHeight,
        scrollHeight,
        hasVerticalScrollbar,
        overflowY,
        expectedVisibleCount: columnCount * visibleRows,
      };
    });
    if (!metrics) throw new Error("无法读取网络内容布局指标");

    const screenshotPath = path.join(runDir, "favorites-layout-desktop.png");
    await page.screenshot({ path: screenshotPath, fullPage: true });

    const assertions = [];
    const pushAssert = (name, passed, detail) => {
      assertions.push({ name, passed: Boolean(passed), detail: String(detail || "") });
    };

    pushAssert("当前视图应为 favorites", metrics.currentView === "favorites", `currentView=${metrics.currentView}`);
    pushAssert("桌面列数应为 5", metrics.columnCount === 5, `columnCount=${metrics.columnCount}`);
    pushAssert("可视行数应为 6", metrics.visibleRows === 6, `visibleRows=${metrics.visibleRows}`);
    pushAssert(
      "导入后卡片数量至少 31",
      metrics.totalCards >= 31,
      `totalCards=${metrics.totalCards}`
    );
    pushAssert(
      "可视容量应为 30（5x6）",
      metrics.expectedVisibleCount === 30,
      `expectedVisibleCount=${metrics.expectedVisibleCount}`
    );
    pushAssert(
      "超过 30 条后应出现纵向滚动",
      metrics.hasVerticalScrollbar,
      `clientHeight=${metrics.clientHeight}, scrollHeight=${metrics.scrollHeight}, overflowY=${metrics.overflowY}`
    );

    const allPassed = assertions.every((item) => item.passed);
    const summary = {
      runAt: new Date().toISOString(),
      appUrl: APP_URL,
      accountUsername,
      screenshot: screenshotPath,
      metrics,
      assertions,
      allPassed,
    };
    const summaryPath = path.join(runDir, "summary.json");
    fs.writeFileSync(summaryPath, `${JSON.stringify(summary, null, 2)}\n`, "utf8");

    if (!allPassed) {
      throw new Error(`网络内容布局验收失败：${summaryPath}`);
    }
    console.log("网络内容布局验收：通过");
    console.log(`截图: ${screenshotPath}`);
    console.log(`报告: ${summaryPath}`);
    console.log(JSON.stringify(summary, null, 2));
  } finally {
    try {
      if (context) await context.close();
    } catch {
      // ignore
    }
    try {
      if (browser) await browser.close();
    } catch {
      // ignore
    }
    if (serverControl && serverControl.server) {
      const { server, getBootLog } = serverControl;
      if (!server.killed) {
        server.kill("SIGTERM");
        await Promise.race([
          new Promise((resolve) => server.once("exit", resolve)),
          delay(1500),
        ]);
        if (!server.killed) {
          try {
            server.kill("SIGKILL");
          } catch {
            // ignore
          }
        }
      }
      if (server.exitCode && server.exitCode !== 0) {
        const bootLog = getBootLog();
        if (bootLog) {
          console.error("服务日志（异常退出）:");
          console.error(bootLog);
        }
      }
    }
  }
}

run().catch((error) => {
  console.error("网络内容布局验收失败:", error && error.stack ? error.stack : error);
  process.exit(1);
});
