import { getDebuggerAddress } from "../../shared/browser_runtime.mjs";
import { chromium } from "./chart-runtime/node_modules/playwright/index.mjs";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import sharp from "./chart-runtime/node_modules/sharp/lib/index.js";

const OVERVIEW_SELECTOR = ".fullPanoramaBox";
const DEFAULT_TIMEOUT_MS = 30_000;
const FIXED_CAPTURE_WAIT_MS = 20_000;

const MARKETPLACES = {
  US: { origin: "https://www.amazon.com", postalCode: "10001", captureWidth: 1896, captureHeight: 821 },
  CA: { origin: "https://www.amazon.ca", postalCode: "M4Y1M7", captureWidth: 1898, captureHeight: 921 },
  JP: { origin: "https://www.amazon.co.jp", postalCode: "100-0001", captureWidth: 1898, captureHeight: 921 },
};

function parseArgs(argv) {
  const args = {};
  for (let index = 2; index < argv.length; index += 1) {
    const key = argv[index];
    if (!key.startsWith("--")) continue;
    args[key.slice(2)] = argv[index + 1];
    index += 1;
  }
  return args;
}

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function safeFileName(value) {
  return normalizeText(value)
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 80);
}

function stage(message) {
  console.error(`[us-category-capture] ${new Date().toISOString()} ${message}`);
}

async function withTimeout(task, timeoutMs, label) {
  let timer;
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${label} timed out after ${timeoutMs}ms`)), timeoutMs);
  });
  try {
    return await Promise.race([Promise.resolve().then(task), timeout]);
  } finally {
    clearTimeout(timer);
  }
}

async function connectBrowser(debuggerAddress) {
  const version = await withTimeout(
    () => fetch(`http://${debuggerAddress}/json/version`).then((response) => response.json()),
    3000,
    "CDP version probe",
  );
  stage(`连接项目 Chrome：${version.Browser || debuggerAddress}`);
  return chromium.connectOverCDP(`http://${debuggerAddress}`, { timeout: 15_000 });
}

function compactPostal(value) {
  return normalizeText(value).replace(/[^a-z0-9]/gi, "").toUpperCase();
}

function postalMatches(value, postalCode) {
  const haystack = compactPostal(value);
  const full = compactPostal(postalCode);
  const sector = full.replace(/[A-Z]{2}$/, "");
  return [full, sector].filter((item) => item.length >= 4).some((item) => haystack.includes(item));
}

async function amazonLocationText(page) {
  return normalizeText(await page.locator("#glow-ingress-line2, #nav-global-location-slot").first().textContent({ timeout: 3000 }).catch(() => ""));
}

async function waitForPostalLocation(page, postalCode, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  let text = "";
  while (Date.now() < deadline) {
    text = await amazonLocationText(page);
    if (postalMatches(text, postalCode)) return text;
    const confirmed = normalizeText((await page.locator("#GLUXHiddenSuccessSelectedAddressPlaceholder, #GLUXZipConfirmationValue").allTextContents()).join(" "));
    if (postalMatches(confirmed, postalCode)) return confirmed;
    await page.waitForTimeout(400);
  }
  return text;
}

async function ensureAmazonPostalCode(page, origin, postalCode, timeoutMs) {
  await page.goto(origin, { waitUntil: "domcontentloaded", timeout: timeoutMs });
  const current = await amazonLocationText(page);
  if (postalMatches(current, postalCode)) return { status: "already-set", text: current };

  for (const selector of ["#sp-cc-accept", "#sp-cc-rejectall-link"]) {
    const button = page.locator(selector).first();
    if (await button.isVisible({ timeout: 800 }).catch(() => false)) {
      await button.click({ timeout: 3000, noWaitAfter: true });
      await page.waitForTimeout(700);
      break;
    }
  }

  const entry = page.locator("#nav-global-location-popover-link, #glow-ingress-block, #nav-global-location-slot").first();
  if (!(await entry.count())) throw new Error("没有找到 Amazon 配送地址入口");
  await entry.click({ timeout: 5000, force: true, noWaitAfter: true });
  await page.locator("#GLUXZipUpdateInput_0, #GLUXZipUpdateInput, input[name='zipCode'], input[name='postalCode']").filter({ visible: true }).first().waitFor({ state: "visible", timeout: 7000 });
  const canadianFirst = page.locator("#GLUXZipUpdateInput_0").filter({ visible: true }).first();
  if (await canadianFirst.count()) {
    const compact = compactPostal(postalCode);
    const canadianSecond = page.locator("#GLUXZipUpdateInput_1").filter({ visible: true }).first();
    await canadianFirst.fill(compact.slice(0, 3), { timeout: 5000 });
    await canadianSecond.fill(compact.slice(3), { timeout: 5000 });
  } else {
    const postalInput = page.locator("#GLUXZipUpdateInput, input[name='zipCode'], input[name='postalCode'], input[placeholder*='postal' i], input[aria-label*='postal' i]").filter({ visible: true }).first();
    await postalInput.waitFor({ state: "visible", timeout: 7000 });
    await postalInput.fill(postalCode, { timeout: 5000 });
  }
  const apply = page.locator("#GLUXZipUpdate input[type='submit'], #GLUXZipUpdate input.a-button-input, #GLUXZipUpdate button, #GLUXZipUpdate").filter({ visible: true }).first();
  await apply.click({ timeout: 5000, force: true, noWaitAfter: true });
  await page.waitForTimeout(1200);

  const confirm = page.locator("#GLUXConfirmClose, button[name='glowDoneButton'], input[name='glowDoneButton']").filter({ visible: true }).first();
  if (await confirm.count()) await confirm.click({ timeout: 5000, force: true, noWaitAfter: true }).catch(() => {});
  await waitForPostalLocation(page, postalCode, 10_000);
  try {
    await page.reload({ waitUntil: "domcontentloaded", timeout: timeoutMs });
  } catch (error) {
    // 日本站提交邮编后可能自行刷新；与主动刷新重叠时等待当前导航完成。
    if (!/ERR_ABORTED|frame was detached/i.test(error?.message || String(error))) throw error;
    await page.waitForLoadState("domcontentloaded", { timeout: timeoutMs }).catch(() => {});
  }
  const verified = await waitForPostalLocation(page, postalCode, 8000);
  if (!postalMatches(verified, postalCode)) {
    throw new Error(`Amazon 邮编未切换到 ${postalCode}，当前显示：${verified || "空"}`);
  }
  return { status: "success", text: verified };
}

async function firstProduct(page, origin, keyword, timeoutMs) {
  const searchUrl = `${origin}/s?k=${encodeURIComponent(keyword)}`;
  await page.goto(searchUrl, { waitUntil: "domcontentloaded", timeout: timeoutMs });
  if (/\/ap\/signin/i.test(page.url())) throw new Error("Amazon 登录已失效");
  const cards = page.locator('.s-main-slot [data-component-type="s-search-result"][data-asin]');
  await cards.first().waitFor({ state: "visible", timeout: timeoutMs });
  const count = await cards.count();
  for (let index = 0; index < count; index += 1) {
    const card = cards.nth(index);
    // 类目调研使用第一个自然商品，跳过广告和 Amazon 品牌推荐卡，避免关键词歧义导致类目跑偏。
    if (await card.locator(".puis-sponsored-label-text").count()) continue;
    const cardPrefix = normalizeText((await card.innerText().catch(() => "")).slice(0, 120));
    if (/^(Sponsored|Featured from Amazon brands)\b/i.test(cardPrefix)) continue;
    const asin = normalizeText(await card.getAttribute("data-asin"));
    const link = card.locator('a.s-line-clamp-3[href], a[href*="/dp/"]').first();
    if (!(await link.count())) continue;
    const href = await link.getAttribute("href", { timeout: 3000 }).catch(() => null);
    if (!asin || !href) continue;
    return {
      asin,
      title: normalizeText(await link.innerText().catch(() => "")),
      href: new URL(href, origin).href,
    };
  }
  throw new Error(`搜索结果中没有可打开的商品：${keyword}`);
}

async function resolveBestSellerCategory(page, product, timeoutMs) {
  await page.goto(product.href, { waitUntil: "domcontentloaded", timeout: timeoutMs });
  if (/\/ap\/signin/i.test(page.url())) throw new Error("Amazon 登录已失效");
  await page.waitForTimeout(2500);

  // 商品详情里的 Best Sellers Rank 通常先列大类，再列更具体的小类；选择最后一个有效类目。
  const links = await page.locator('a[href*="/gp/bestsellers/"]').evaluateAll((elements) => {
    const seen = new Set();
    return elements
      .map((element) => ({
        text: (element.textContent || "").replace(/\s+/g, " ").trim(),
        href: element.href,
      }))
      .filter((item) => {
        if (!item.text || !item.href || seen.has(item.href)) return false;
        seen.add(item.href);
        return /\/gp\/bestsellers\//i.test(item.href);
      });
  });
  if (!links.length) throw new Error(`商品页没有找到 Best Sellers 类目：${product.asin}`);

  const leafCandidates = links.filter((item) => !/\/gp\/bestsellers\/?(?:$|[?#])/i.test(item.href));
  return (leafCandidates.length ? leafCandidates : links).at(-1);
}

async function waitForOverview(page, timeoutMs) {
  const overview = page.locator(OVERVIEW_SELECTOR).first();
  await overview.waitFor({ state: "visible", timeout: timeoutMs });
  await overview.scrollIntoViewIfNeeded({ timeout: timeoutMs });
  await page.waitForTimeout(FIXED_CAPTURE_WAIT_MS);

  const state = await overview.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    const text = (element.innerText || "").replace(/\s+/g, " ").trim();
    const canvases = Array.from(element.querySelectorAll("canvas"));
    return {
      width: rect.width,
      height: rect.height,
      text,
      canvasCount: canvases.length,
      canvasesReady: canvases.every((canvas) => canvas.width > 0 && canvas.height > 0),
    };
  });
  if (state.width < 300 || state.height < 120 || !/月总销量/.test(state.text)) {
    throw new Error("Softime 面板出现但指标未加载完成");
  }
  return { overview, state };
}

async function captureOverview(page, context, outputPath, captureWidth, captureHeight) {
  // 首张示例图包含类目树、Best Sellers 标题和完整插件面板，因此从类目标题处截取固定视口。
  const localizedTitle = page.locator("h1.a-size-large.a-spacing-medium.a-text-bold").first();
  const title = await localizedTitle.count()
    ? localizedTitle
    : page.locator("h1").filter({ hasText: /Best Sellers in|売れ筋ランキング/ }).last();
  if (!(await title.count())) throw new Error("没有找到 Best Sellers 类目标题");
  await title.evaluate((element) => {
    const top = element.getBoundingClientRect().top + window.scrollY;
    window.scrollTo(0, top);
  });
  await page.waitForTimeout(500);

  const cdp = await context.newCDPSession(page);
  try {
    const result = await cdp.send("Page.captureScreenshot", {
      format: "png",
      fromSurface: true,
      captureBeyondViewport: false,
    });
    const fullImage = Buffer.from(result.data, "base64");
    const metadata = await sharp(fullImage).metadata();
    const width = Math.min(captureWidth, metadata.width || 0);
    const height = Math.min(captureHeight, metadata.height || 0);
    if (width < 1200 || height < 600) throw new Error("项目 Chrome 视口尺寸不足");
    const cropped = sharp(fullImage).extract({ left: 0, top: 0, width, height });
    const [png, stats] = await Promise.all([cropped.clone().png().toBuffer(), cropped.stats()]);
    const meaningful = stats.channels.some((channel) => (
      channel.min < 245 || channel.max - channel.min > 8 || channel.stdev > 2
    ));
    if (!meaningful) throw new Error("Softime 截图为空白");
    await writeFile(outputPath, png);
    return { width, height };
  } finally {
    await cdp.detach().catch(() => {});
  }
}

async function prepareWindow(page, context) {
  const cdp = await context.newCDPSession(page);
  const original = await cdp.send("Browser.getWindowForTarget");
  await cdp.send("Browser.setWindowBounds", {
    windowId: original.windowId,
    bounds: { left: 0, top: 0, width: 1920, height: 1080, windowState: "normal" },
  });
  await page.bringToFront();
  await page.setViewportSize({ width: 1920, height: 1080 });
  return async () => {
    try {
      await cdp.send("Browser.setWindowBounds", {
        windowId: original.windowId,
        bounds: original.bounds,
      });
    } finally {
      await cdp.detach().catch(() => {});
    }
  };
}

async function main() {
  const args = parseArgs(process.argv);
  const debuggerAddress = args["debugger-address"] || getDebuggerAddress();
  const timeoutMs = Number(args["timeout-ms"] || DEFAULT_TIMEOUT_MS);
  const site = String(args.site || "US").toUpperCase();
  const marketplace = MARKETPLACES[site];
  if (!marketplace) throw new Error(`不支持的站点：${site}`);
  const postalCode = args["postal-code"] || marketplace.postalCode;
  const outputDir = path.resolve(args["output-dir"] || "outputs/us_category_captures");
  const keywords = JSON.parse(args.keywords || "[]");
  if (!Array.isArray(keywords) || !keywords.length) throw new Error("缺少 --keywords JSON 数组");
  await mkdir(outputDir, { recursive: true });

  const browser = await connectBrowser(debuggerAddress);
  const results = [];
  try {
    const context = browser.contexts()[0];
    if (!context) throw new Error("项目 Chrome 没有可用上下文");
    const targetHost = new URL(marketplace.origin).hostname;
    const page = context.pages().find((candidate) => {
      try {
        return new URL(candidate.url()).hostname.endsWith(targetHost);
      } catch {
        return false;
      }
    }) || await context.newPage();
    const restoreWindow = await prepareWindow(page, context);
    try {
      const postal = await ensureAmazonPostalCode(page, marketplace.origin, postalCode, timeoutMs);
      stage(`站点与邮编：${site} ${postal.text || postalCode}`);
      for (let index = 0; index < keywords.length; index += 1) {
        const keyword = normalizeText(keywords[index]);
        const outputPath = path.join(outputDir, `${String(index + 1).padStart(2, "0")}_${safeFileName(keyword)}.png`);
        stage(`${index + 1}/${keywords.length} 搜索：${keyword}`);
        try {
          const product = await firstProduct(page, marketplace.origin, keyword, timeoutMs);
          stage(`首个商品：${product.asin} ${product.title}`);
          const category = await resolveBestSellerCategory(page, product, timeoutMs);
          stage(`类目：${category.text} ${category.href}`);
          await page.goto(category.href, { waitUntil: "domcontentloaded", timeout: timeoutMs });
          const overview = await waitForOverview(page, timeoutMs);
          const image = await captureOverview(
            page,
            context,
            outputPath,
            marketplace.captureWidth,
            marketplace.captureHeight,
          );
          results.push({
            keyword,
            status: "ok",
            product,
            category,
            resolvedCategoryUrl: page.url(),
            outputPath,
            image,
            overview: {
              canvasCount: overview.state.canvasCount,
              canvasesReady: overview.state.canvasesReady,
              text: overview.state.text,
            },
            site,
            postalCode,
          });
        } catch (error) {
          const message = error?.message || String(error);
          stage(`失败：${keyword}：${message}`);
          results.push({ keyword, status: "error", error: message });
        }
      }
    } finally {
      await restoreWindow().catch(() => {});
    }
  } finally {
    browser._shouldCloseConnectionOnClose = false;
    await browser.close().catch(() => {});
  }

  const reportPath = path.join(outputDir, "capture_results.json");
  await writeFile(reportPath, JSON.stringify(results, null, 2), "utf8");
  console.log(JSON.stringify({ reportPath, results }, null, 2));
  if (results.some((item) => item.status !== "ok")) process.exitCode = 2;
}

main().catch((error) => {
  console.error(error?.stack || error);
  process.exitCode = 1;
});
