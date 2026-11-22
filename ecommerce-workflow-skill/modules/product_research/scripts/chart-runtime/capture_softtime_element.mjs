import { chromium } from "playwright";
import sharp from "sharp";


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


function normalize(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}


function parseMonthlySales(text) {
  const match = normalize(text).match(/月总销量\s*[:：]?\s*([0-9][0-9,]*)\b/);
  return match ? Number(match[1].replace(/,/g, "")) : null;
}


async function main() {
  const args = parseArgs(process.argv);
  const url = args.url;
  const output = args.output;
  const debuggerAddress = args["debugger-address"] || "127.0.0.1:9223";
  const timeoutMs = Number(args["timeout-ms"] || 60000);
  if (!url || !output) {
    throw new Error("Missing --url or --output");
  }

  const browser = await chromium.connectOverCDP(`http://${debuggerAddress}`, {
    timeout: 20000,
  });
  try {
    const context = browser.contexts()[0];
    const page = await context.newPage();
    await page.setViewportSize({ width: 2400, height: 1200 });
    await page.goto(url, { waitUntil: "domcontentloaded", timeout: timeoutMs });
    const overview = page.locator(".fullPanoramaBox").first();
    await overview.waitFor({ state: "visible", timeout: timeoutMs });
    await overview.scrollIntoViewIfNeeded({ timeout: timeoutMs });
    await page.waitForTimeout(20000);
    const text = await overview.innerText({ timeout: 5000 });
    if (!text.includes("月总销量")) {
      throw new Error("Softtime overview does not contain monthly sales");
    }
    const canvases = await overview.locator("canvas").count();
    if (canvases < 12) {
      throw new Error(`Softtime overview has only ${canvases} canvases`);
    }

    // 元素级截图由 Playwright 直接处理 DPR 和裁剪，不再手工换算整页坐标。
    await overview.screenshot({
      path: output,
      type: "png",
      animations: "disabled",
      caret: "hide",
      timeout: timeoutMs,
    });
    const stats = await sharp(output).stats();
    const meaningful = stats.channels.some(
      (channel) =>
        channel.min < 245 ||
        channel.max - channel.min > 8 ||
        channel.stdev > 2,
    );
    if (!meaningful) {
      throw new Error("Softtime element screenshot is blank");
    }
    console.log(
      JSON.stringify({
        url: page.url(),
        monthly_sales: parseMonthlySales(text),
        canvases,
        screenshot_blank: false,
      }),
    );
  } finally {
    browser._shouldCloseConnectionOnClose = false;
    await browser.close().catch(() => {});
  }
}


main().catch((error) => {
  console.error(error?.stack || String(error));
  process.exit(1);
});
