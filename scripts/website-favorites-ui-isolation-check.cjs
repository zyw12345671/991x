const fs = require("fs");
const path = require("path");
const { spawn } = require("child_process");
const { chromium } = require("playwright-core");

const PROJECT_ROOT = path.join(__dirname, "..");
const APP_URL = "http://127.0.0.1:9910/";
const SCREENSHOT_ROOT = path.join(PROJECT_ROOT, "artifacts", "website-favorites-ui-isolation");
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
      options.runTag = "mobile";
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
  throw new Error("未找到可用浏览器（Chrome/Edge），无法执行实机截图验收");
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForServerReady(timeoutMs = 30000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    try {
      const response = await fetch(APP_URL, { method: "GET" });
      if (response.ok) return;
    } catch {
      // ignore
    }
    await delay(400);
  }
  throw new Error(`服务启动超时（${timeoutMs}ms）: ${APP_URL}`);
}

async function checkServerReadyOnce() {
  try {
    const response = await fetch(APP_URL, { method: "GET" });
    return Boolean(response && response.ok);
  } catch {
    return false;
  }
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

function listCardTitlesByCanvas(page) {
  return page.$$eval("#websiteCollectionGrid .website-collection-card .website-collection-link strong", (nodes) =>
    nodes.map((node) => String(node.textContent || "").trim()).filter(Boolean)
  );
}

async function waitForCardTitle(page, title, timeoutMs = 6000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const titles = await listCardTitlesByCanvas(page);
    if (titles.includes(title)) return;
    await delay(120);
  }
  throw new Error(`未在卡片区找到标题：${title}`);
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

async function expectTopbar(page, expectedTitle, expectedSubtitle) {
  await page.waitForFunction(
    ({ title, subtitle }) => {
      const t = document.getElementById("topbarTitle");
      const s = document.getElementById("topbarSubtitle");
      return (
        !!t &&
        !!s &&
        String(t.textContent || "").trim() === String(title || "").trim() &&
        String(s.textContent || "").trim() === String(subtitle || "").trim()
      );
    },
    { title: expectedTitle, subtitle: expectedSubtitle }
  );
}

async function run() {
  const runOptions = parseRunOptions();
  ensureDir(SCREENSHOT_ROOT);
  const runStamp = stampNow();
  const runDir = path.join(SCREENSHOT_ROOT, `${runStamp}-${runOptions.runTag}`);
  ensureDir(runDir);

  const accountUsername = `iso_${Date.now().toString().slice(-7)}`;
  const accountPassword = "IsoPass#2026";
  const websiteItemTitle = "隔离站点A";
  const websiteItemUrl = "https://site-a.example.com";
  const favoriteItemTitle = "隔离内容A";
  const favoriteItemUrl = "https://content-a.example.com";

  const checkpoints = [];
  const addCheckpoint = (point, expected, actual, screenshotPath, passed) => {
    checkpoints.push({
      point,
      expected,
      actual,
      screenshot: screenshotPath,
      passed: Boolean(passed),
    });
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
    const executablePath = resolveChromePath();
    browser = await chromium.launch({
      executablePath,
      headless: true,
      args: ["--disable-gpu"],
    });
    context = await browser.newContext({
      viewport: runOptions.viewport,
      locale: "zh-CN",
    });
    const page = await context.newPage();
    const dialogMessages = [];
    page.on("dialog", async (dialog) => {
      dialogMessages.push(String(dialog.message() || "").trim());
      await dialog.accept();
    });
    await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
    await page.waitForSelector("#menuWebsiteButton", { state: "visible" });

    // 1) 网站区入口（未登录）截图
    await page.click("#menuWebsiteButton");
    await expectTopbar(page, "网站收藏", "收录喜欢的网站");
    await page.waitForSelector("#websiteCenter:not([hidden])");
    const s1 = path.join(runDir, "01-website-view-before-login.png");
    await page.screenshot({ path: s1, fullPage: true });
    addCheckpoint(
      "网站区入口",
      "顶部标题为“网站收藏”，并显示登录入口",
      "标题正确，网站区已打开",
      s1,
      true
    );

    // 2) 注册并解锁
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
    await page.waitForSelector("#websiteWorkspace:not([hidden])");
    const s2 = path.join(runDir, "02-after-register-unlocked.png");
    await page.screenshot({ path: s2, fullPage: true });
    addCheckpoint(
      "登录解锁",
      "注册成功后网站区工作台可见",
      `注册账号 ${accountUsername} 成功，工作台已显示`,
      s2,
      true
    );

    // 3) 在网站区新增卡片
    await page.click("#websiteCollectionAddButton");
    await page.waitForSelector("#websiteAddModal:not([hidden])");
    await page.fill("#websiteAddNameInput", websiteItemTitle);
    await page.fill("#websiteAddUrlInput", websiteItemUrl);
    await page.click("#websiteAddSubmitButton");
    await page.waitForFunction(() => {
      const modal = document.getElementById("websiteAddModal");
      return Boolean(modal && modal.hidden);
    });
    await waitForCardTitle(page, websiteItemTitle);
    const s3 = path.join(runDir, "03-website-added-item.png");
    await page.screenshot({ path: s3, fullPage: true });
    addCheckpoint(
      "网站区新增",
      "网站区出现“隔离站点A”卡片",
      "新增成功并可见",
      s3,
      true
    );

    // 4) 切换到网络内容并校验标题
    await page.click("#menuFavoritesButton");
    await expectTopbar(page, "网络内容", "收藏喜欢的网络内容");
    const s4 = path.join(runDir, "04-favorites-view-initial.png");
    await page.screenshot({ path: s4, fullPage: true });
    addCheckpoint(
      "网络内容区入口",
      "顶部标题为“网络内容”与副标题“收藏喜欢的网络内容”",
      "标题与副标题正确",
      s4,
      true
    );

    // 5) 在网络内容区新增卡片
    await page.click("#websiteCollectionAddButton");
    await page.waitForSelector("#websiteAddModal:not([hidden])");
    await page.fill("#websiteAddNameInput", favoriteItemTitle);
    await page.fill("#websiteAddUrlInput", favoriteItemUrl);
    await page.click("#websiteAddSubmitButton");
    await page.waitForFunction(() => {
      const modal = document.getElementById("websiteAddModal");
      return Boolean(modal && modal.hidden);
    });
    await waitForCardTitle(page, favoriteItemTitle);
    const s5 = path.join(runDir, "05-favorites-added-item.png");
    await page.screenshot({ path: s5, fullPage: true });
    addCheckpoint(
      "网络内容区新增",
      "网络内容区出现“隔离内容A”卡片",
      "新增成功并可见",
      s5,
      true
    );

    // 6) 回网站区，验证隔离
    await page.click("#menuWebsiteButton");
    await expectTopbar(page, "网站收藏", "收录喜欢的网站");
    await waitForCardTitle(page, websiteItemTitle);
    const websiteTitles = await listCardTitlesByCanvas(page);
    const websiteIsolationPassed = websiteTitles.includes(websiteItemTitle) && !websiteTitles.includes(favoriteItemTitle);
    const s6 = path.join(runDir, "06-back-to-website-isolation.png");
    await page.screenshot({ path: s6, fullPage: true });
    addCheckpoint(
      "网站区回切隔离",
      "网站区只看到网站数据，不出现网络内容数据",
      `当前标题: ${websiteTitles.join(" / ") || "（空）"}`,
      s6,
      websiteIsolationPassed
    );
    if (!websiteIsolationPassed) throw new Error(`网站区隔离失败：${JSON.stringify(websiteTitles)}`);

    // 7) 回网络内容区，验证隔离
    await page.click("#menuFavoritesButton");
    await expectTopbar(page, "网络内容", "收藏喜欢的网络内容");
    await waitForCardTitle(page, favoriteItemTitle);
    const favoritesTitles = await listCardTitlesByCanvas(page);
    const favoritesIsolationPassed = favoritesTitles.includes(favoriteItemTitle) && !favoritesTitles.includes(websiteItemTitle);
    const s7 = path.join(runDir, "07-back-to-favorites-isolation.png");
    await page.screenshot({ path: s7, fullPage: true });
    addCheckpoint(
      "网络内容区回切隔离",
      "网络内容区只看到自身数据，不出现网站区数据",
      `当前标题: ${favoritesTitles.join(" / ") || "（空）"}`,
      s7,
      favoritesIsolationPassed
    );
    if (!favoritesIsolationPassed) throw new Error(`网络内容区隔离失败：${JSON.stringify(favoritesTitles)}`);

    // 8) 跨区导入阻断：favorites <- website export shape
    const crossWebsiteToFavoritesPayload = {
      source: "991x-protected-feature-website-collection",
      feature: "website",
      items: [{ title: "跨区站点X", link: "https://cross-x.example.com" }],
    };
    const dialogCountBeforeCross1 = dialogMessages.length;
    await importJsonViaTopbarButton(page, crossWebsiteToFavoritesPayload, "cross-website-to-favorites.json");
    await delay(700);
    const favoritesAfterCrossImport = await listCardTitlesByCanvas(page);
    const crossImportBlocked1 = !favoritesAfterCrossImport.includes("跨区站点X");
    const crossImportDialog1 = dialogMessages.slice(dialogCountBeforeCross1).some((message) => message.includes("文件类型与当前功能不匹配"));
    const s8 = path.join(runDir, "08-cross-import-blocked-favorites.png");
    await page.screenshot({ path: s8, fullPage: true });
    addCheckpoint(
      "跨区导入阻断（网络内容区）",
      "网络内容区不能导入网站收藏导出JSON，且弹出错误提示",
      `当前标题: ${favoritesAfterCrossImport.join(" / ") || "（空）"}；提示弹窗: ${crossImportDialog1 ? "已出现" : "未出现"}`,
      s8,
      crossImportBlocked1 && crossImportDialog1
    );
    if (!crossImportBlocked1) throw new Error("跨区导入阻断失败：favorites 导入了 website 文件");
    if (!crossImportDialog1) throw new Error("跨区导入阻断失败：favorites 未出现错误提示弹窗");

    // 9) 跨区导入阻断：website <- favorites export shape
    await page.click("#menuWebsiteButton");
    await expectTopbar(page, "网站收藏", "收录喜欢的网站");
    const crossFavoritesToWebsitePayload = {
      source: "991x-protected-feature-favorites-content",
      feature: "favorites",
      items: [{ title: "跨区内容Y", link: "https://cross-y.example.com" }],
    };
    const dialogCountBeforeCross2 = dialogMessages.length;
    await importJsonViaTopbarButton(page, crossFavoritesToWebsitePayload, "cross-favorites-to-website.json");
    await delay(700);
    const websiteAfterCrossImport = await listCardTitlesByCanvas(page);
    const crossImportBlocked2 = !websiteAfterCrossImport.includes("跨区内容Y");
    const crossImportDialog2 = dialogMessages.slice(dialogCountBeforeCross2).some((message) => message.includes("文件类型与当前功能不匹配"));
    const s9 = path.join(runDir, "09-cross-import-blocked-website.png");
    await page.screenshot({ path: s9, fullPage: true });
    addCheckpoint(
      "跨区导入阻断（网站区）",
      "网站区不能导入网络内容导出JSON，且弹出错误提示",
      `当前标题: ${websiteAfterCrossImport.join(" / ") || "（空）"}；提示弹窗: ${crossImportDialog2 ? "已出现" : "未出现"}`,
      s9,
      crossImportBlocked2 && crossImportDialog2
    );
    if (!crossImportBlocked2) throw new Error("跨区导入阻断失败：website 导入了 favorites 文件");
    if (!crossImportDialog2) throw new Error("跨区导入阻断失败：website 未出现错误提示弹窗");

    // 10) 导出文件名前缀区分
    const websiteDownloadPromise = page.waitForEvent("download");
    await page.click("#websiteExportButton");
    const websiteDownload = await websiteDownloadPromise;
    const websiteFilename = websiteDownload.suggestedFilename();
    await page.click("#menuFavoritesButton");
    await expectTopbar(page, "网络内容", "收藏喜欢的网络内容");
    const favoritesDownloadPromise = page.waitForEvent("download");
    await page.click("#websiteExportButton");
    const favoritesDownload = await favoritesDownloadPromise;
    const favoritesFilename = favoritesDownload.suggestedFilename();
    const filenameRulePass = /^website-collection-/.test(websiteFilename) && /^favorites-content-/.test(favoritesFilename);
    const s10 = path.join(runDir, "10-export-filename-patterns.png");
    await page.screenshot({ path: s10, fullPage: true });
    addCheckpoint(
      "导出命名隔离",
      "网站区与网络内容区导出文件名前缀不同",
      `website: ${websiteFilename} | favorites: ${favoritesFilename}`,
      s10,
      filenameRulePass
    );
    if (!filenameRulePass) throw new Error(`导出命名不符合预期: ${websiteFilename} / ${favoritesFilename}`);

    const summary = {
      runAt: new Date().toISOString(),
      appUrl: APP_URL,
      accountUsername,
      screenshotDir: runDir,
      viewport: runOptions.viewport,
      runTag: runOptions.runTag,
      allPassed: checkpoints.every((item) => item.passed),
      checkpointCount: checkpoints.length,
      checkpoints,
    };
    const summaryPath = path.join(runDir, "summary.json");
    fs.writeFileSync(summaryPath, `${JSON.stringify(summary, null, 2)}\n`, "utf8");
    console.log("UI隔离验收：全部完成");
    console.log(`截图目录: ${runDir}`);
    console.log(`总结报告: ${summaryPath}`);
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
  console.error("UI隔离验收失败:", error && error.stack ? error.stack : error);
  process.exit(1);
});
