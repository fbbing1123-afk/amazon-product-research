import { chromium } from "playwright";
import { writeFile } from "node:fs/promises";

function parseArgs(argv) {
  const args = {};
  for (let index = 2; index < argv.length; index += 2) {
    const key = argv[index];
    if (!key?.startsWith("--")) continue;
    args[key.slice(2)] = argv[index + 1];
  }
  return args;
}

function firstNumber(text) {
  const match = String(text || "").replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function splitKeyword(text) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  const match = normalized.match(/^(.+?)([\u3400-\u9fff].*)$/);
  if (!match) return { keyword: normalized, translation: "" };
  return {
    keyword: match[1].trim(),
    translation: match[2].trim(),
  };
}

async function waitForResults(page, timeoutMs = 120_000) {
  const deadline = Date.now() + timeoutMs;
  let lastState = "";
  while (Date.now() < deadline) {
    const state = await page.evaluate(() => {
      const body = document.body.innerText || "";
      const match = body.match(/搜索结果数\s*[:：]\s*([0-9,]+)/);
      return {
        count: match ? Number(match[1].replace(/,/g, "")) : 0,
        masks: [...document.querySelectorAll(".el-loading-mask")]
          .filter((node) => node.offsetWidth || node.offsetHeight).length,
        error: body.includes("处理请求出现错误"),
      };
    });
    lastState = JSON.stringify(state);
    if (state.count > 0 && state.masks === 0) return state.count;
    await page.waitForTimeout(1000);
  }
  throw new Error(`等待关键词结果超时：${lastState}`);
}

async function readPageRows(page) {
  return page.evaluate(() => {
    return [...document.querySelectorAll(".el-table__body tr")].map((row) => {
      return [...row.children].map((cell) => (cell.innerText || "").replace(/\s+/g, " ").trim());
    });
  });
}

async function main() {
  const args = parseArgs(process.argv);
  const asins = String(args.asins || "").split(",").map((item) => item.trim()).filter(Boolean);
  const output = args.output;
  if (!asins.length || !output) {
    throw new Error("缺少 --asins 或 --output");
  }

  const version = await (await fetch("http://127.0.0.1:9223/json/version")).json();
  const browser = await chromium.connectOverCDP(version.webSocketDebuggerUrl, { timeout: 20_000 });
  const pages = browser.contexts()[0].pages();
  const page = pages.find((item) => item.url().includes("sellersprite.com"));
  if (!page) throw new Error("没有找到卖家精灵页面");
  await page.bringToFront();

  const queryUrl = `https://www.sellersprite.com/v3/traffic/extend/asin?q=${encodeURIComponent(asins.join(","))}&marketId=1&date=`;
  await page.goto(queryUrl, { waitUntil: "domcontentloaded", timeout: 60_000 });
  const queryButton = page.locator("button.el-button--primary", { hasText: "立即查询" }).last();
  await queryButton.waitFor({ state: "visible", timeout: 30_000 });
  await queryButton.click({ force: true });
  const currentVariant = page.locator("button", { hasText: "用当前变体拓词" }).last();
  await currentVariant.waitFor({ state: "visible", timeout: 30_000 });
  await currentVariant.click({ force: true });
  const resultCount = await waitForResults(page);

  const outputRows = [];
  const seen = new Set();
  for (let pageNumber = 1; pageNumber <= 100; pageNumber += 1) {
    const rows = await readPageRows(page);
    for (const cells of rows) {
      if (cells.length < 14) continue;
      const { keyword, translation } = splitKeyword(cells[2]);
      if (!keyword || seen.has(keyword.toLowerCase())) continue;
      seen.add(keyword.toLowerCase());
      const relatedText = cells[4] || "";
      const relatedMatch = relatedText.match(/相关产品\s*[:：]\s*([0-9,]+)/);
      outputRows.push({
        "关键词": keyword,
        "关键词翻译": translation,
        "相关产品数": relatedMatch ? Number(relatedMatch[1].replace(/,/g, "")) : firstNumber(relatedText),
        "月搜索量": cells[7] || "",
        "流量占比": cells[3] || "",
        "ABA周排名": cells[6] || "",
        "月购买量": cells[8] || "",
        "展示量/点击量": cells[9] || "",
        "SPR": cells[10] || "",
        "标题密度": cells[11] || "",
        "需供比/商品数": cells[12] || "",
        "广告竞品数": cells[13] || "",
        "ABA集中度": cells[14] || "",
        "PPC竞价": cells[15] || "",
      });
    }

    const nextButton = page.locator(".el-pagination button.btn-next").last();
    if (!(await nextButton.count())) break;
    const disabled = await nextButton.getAttribute("disabled");
    const className = (await nextButton.getAttribute("class")) || "";
    if (disabled !== null || className.includes("is-disabled")) break;
    const before = (await readPageRows(page))[0]?.[1] || "";
    await nextButton.click({ force: true });
    for (let attempt = 0; attempt < 30; attempt += 1) {
      await page.waitForTimeout(300);
      const current = (await readPageRows(page))[0]?.[1] || "";
      if (current !== before) break;
    }
  }

  await writeFile(
    output,
    JSON.stringify({ asins, result_count: resultCount, rows: outputRows }, null, 2),
    "utf8",
  );
  console.log(JSON.stringify({ output, result_count: resultCount, scraped: outputRows.length }, null, 2));
  browser._shouldCloseConnectionOnClose = false;
  await browser.close();
}

main().catch((error) => {
  console.error(error?.stack || String(error));
  process.exit(1);
});
