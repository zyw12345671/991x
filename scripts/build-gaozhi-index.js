const fs = require("node:fs/promises");
const path = require("node:path");

const ROOT_DIR = path.resolve(__dirname, "..");
const GAOZHI_DIR = path.join(ROOT_DIR, "gaozhi");
const OUT_FILE = path.join(GAOZHI_DIR, "index.json");

async function build() {
  let entries = [];
  try {
    entries = await fs.readdir(GAOZHI_DIR, { withFileTypes: true });
  } catch (error) {
    console.error("读取 gaozhi 目录失败：", error.message);
    process.exit(1);
  }

  const documents = await Promise.all(
    entries
      .filter((entry) => entry.isFile() && /\.md$/i.test(entry.name))
      .map(async (entry) => {
        const absPath = path.join(GAOZHI_DIR, entry.name);
        let uploadedAtMs = 0;
        try {
          const stat = await fs.stat(absPath);
          const birthtimeMs = Number.isFinite(stat.birthtimeMs) ? Math.floor(stat.birthtimeMs) : 0;
          const mtimeMs = Number.isFinite(stat.mtimeMs) ? Math.floor(stat.mtimeMs) : 0;
          uploadedAtMs = birthtimeMs > 0 ? birthtimeMs : mtimeMs;
        } catch {
          uploadedAtMs = 0;
        }
        return {
          path: `gaozhi/${entry.name}`,
          uploadedAtMs,
          uploadedAt: uploadedAtMs > 0 ? new Date(uploadedAtMs).toISOString() : "",
        };
      })
  );

  documents.sort((a, b) => {
    if (b.uploadedAtMs !== a.uploadedAtMs) return b.uploadedAtMs - a.uploadedAtMs;
    return b.path.localeCompare(a.path, "zh-CN", { numeric: true, sensitivity: "base" });
  });
  const files = documents.map((item) => item.path);

  const output = `${JSON.stringify({ files, documents }, null, 2)}\n`;
  await fs.writeFile(OUT_FILE, output, "utf8");
  console.log(`已生成 ${path.relative(ROOT_DIR, OUT_FILE)}，共 ${files.length} 条文档记录。`);
}

build();
