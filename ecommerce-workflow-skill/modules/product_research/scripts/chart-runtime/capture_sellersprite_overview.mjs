import { chromium } from "playwright";
import { writeFile } from "node:fs/promises";
import sharp from "sharp";

const FIXED_CAPTURE_WAIT_MS = 20_000;
const CAPTURE_LEFT_PAD_CSS_PX = 16;
const CAPTURE_TOP_PAD_CSS_PX = 80;
const CAPTURE_EDGE_PAD_PX = 4;
const CAPTURE_MIN_TOP_CSS_PX = 120;
const CATEGORY_OVERVIEW_TARGET_ASPECT = 1529 / 487;

function parseArgs(argv) {
  const args = {};
  for (let i = 2; i < argv.length; i += 1) {
    const key = argv[i];
    if (!key.startsWith("--")) {
      continue;
    }
    args[key.slice(2)] = argv[i + 1];
    i += 1;
  }
  return args;
}

function parseMonthlySales(text) {
  const compact = String(text || "").replace(/\s+/g, " ");
  const match = compact.match(/月总销量\s*[:：]?\s*([0-9][0-9,]*)\b/);
  return match ? Number(match[1].replace(/,/g, "")) : null;
}

function normalize(text) {
  return String(text || "").replace(/\s+/g, " ").trim();
}

function compactPostal(text) {
  return normalize(text)
    .replace(/[\u200c-\u200f\ufeff]/g, "")
    .replace(/[^a-z0-9]/gi, "")
    .toUpperCase();
}

function postalNeedles(text) {
  const full = compactPostal(text);
  const sector = full.replace(/[A-Z]{2}$/, "");
  return [...new Set([full, sector].filter((value) => value.length >= 4))];
}

function postalMatches(text, postalCode) {
  const haystack = compactPostal(text);
  return postalNeedles(postalCode).some((needle) => haystack.includes(needle));
}

async function withTimeout(task, timeoutMs, label) {
  let timer;
  const operation = Promise.resolve().then(task);
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`${label} timed out after ${timeoutMs}ms`)), timeoutMs);
  });
  try {
    return await Promise.race([operation, timeout]);
  } finally {
    clearTimeout(timer);
  }
}

function stage(message) {
  console.error(`[category-overview] ${new Date().toISOString()} ${message}`);
}

async function connectBrowserWithRetry(debuggerAddress) {
  let lastError = null;
  for (let attempt = 1; attempt <= 3; attempt += 1) {
    try {
      // 先探测 /json/version，区分端口不存在与浏览器 WebSocket 握手卡住。
      const versionResponse = await withTimeout(
        () => fetch(`http://${debuggerAddress}/json/version`),
        2000,
        "CDP version probe",
      );
      if (!versionResponse.ok) {
        throw new Error(`CDP version probe returned HTTP ${versionResponse.status}`);
      }
      const version = await withTimeout(() => versionResponse.json(), 2000, "CDP version response");
      const socketUrl = version.webSocketDebuggerUrl;
      if (!socketUrl) {
        throw new Error("CDP version response has no webSocketDebuggerUrl");
      }
      stage(`CDP connection attempt ${attempt}/3`);
      return await withTimeout(
        // 直接连接浏览器 WebSocket，绕开项目浏览器 HTTP 入口偶发的握手超时。
        () => chromium.connectOverCDP(socketUrl, { timeout: 30000 }),
        35000,
        "CDP connection",
      );
    } catch (error) {
      lastError = error;
      stage(`CDP connection attempt ${attempt}/3 failed: ${error?.message || error}`);
      if (attempt < 3) {
        await new Promise((resolve) => setTimeout(resolve, 800));
      }
    }
  }
  throw new Error(`CDP connection failed after two retries: ${lastError?.message || lastError}`);
}

async function locationText(page) {
  return normalize(await page.locator("#glow-ingress-line2, #nav-global-location-slot").first().textContent({ timeout: 3000 }).catch(() => ""));
}

async function successText(page) {
  return normalize((await page.locator("#GLUXHiddenSuccessSelectedAddressPlaceholder, #GLUXZipConfirmationValue").allTextContents()).join(" "));
}

async function waitForPostalLocation(page, postalCode, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  let current = "";
  while (Date.now() < deadline) {
    current = await locationText(page);
    if (postalMatches(current, postalCode)) {
      return current;
    }
    await page.waitForTimeout(Math.min(500, Math.max(50, deadline - Date.now())));
  }
  return current;
}

async function closeVisibleLocationConfirmation(page) {
  await page.locator("#GLUXConfirmClose").evaluateAll((elements) => {
    const visible = elements.find((element) => {
      const rect = element.getBoundingClientRect();
      const style = getComputedStyle(element);
      return rect.width > 0 && rect.height > 0 && style.display !== "none" && style.visibility !== "hidden";
    });
    if (visible) {
      visible.click();
    }
  }).catch(() => {});
}

async function captureOverviewWithCDP(page, context, output, options = {}) {
  const fromSurface = Boolean(options.fromSurface);
  const writeBlank = Boolean(options.writeBlank);
  const debugFullOutput = options.debugFullOutput || "";
  const overview = page.locator(".fullPanoramaBox").first();
  let box = null;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    await page.evaluate(() => {
      document.querySelector(".fullPanoramaBox")?.scrollIntoView({ block: "start", inline: "nearest" });
    });
    await page.waitForTimeout(500);
    await page.evaluate((minimumTop) => {
      const element = document.querySelector(".fullPanoramaBox");
      if (!element) return;
      const rect = element.getBoundingClientRect();
      // 顶部至少留出 120 CSS 像素，避免浏览器合成截图时吞掉插件最上方指标行。
      if (rect.top < minimumTop) {
        window.scrollBy(0, rect.top - minimumTop);
      }
    }, CAPTURE_MIN_TOP_CSS_PX);
    await page.waitForTimeout(200);
    box = await overview.boundingBox();
    const viewport = page.viewportSize();
    if (box && viewport && box.y >= -1 && box.y + box.height <= viewport.height) {
      break;
    }
  }
  if (!box || box.width <= 0 || box.height <= 0) {
    throw new Error("Softtime market overview has no valid screenshot bounds.");
  }
  const viewport = page.viewportSize();
  if (!viewport || box.y < -1 || box.y + box.height > viewport.height) {
    throw new Error(`Softtime overview is outside viewport: y=${box.y}, height=${box.height}`);
  }
  stage(`screenshot bounds: x=${box.x}, y=${box.y}, width=${box.width}, height=${box.height}`);
  const deviceScale = await page.evaluate(() => window.devicePixelRatio || 1);
  // 先截完整视口，再按元素边界裁剪。直接裁剪 CDP surface 在窄窗口下会把未绘制区域当成白色。
  const cdp = await context.newCDPSession(page);
  const floatingToolStyle = await page.evaluate(() => {
    const tool = document.querySelector("#sorftime-logo");
    if (!tool) return null;
    const original = tool.getAttribute("style");
    tool.style.setProperty("display", "none", "important");
    return original;
  });
  try {
    let fullScreenshot;
    try {
      // 优先使用 Playwright 原生截图，避免大型 Canvas 页面阻塞 CDP 截图。
      fullScreenshot = await withTimeout(
        () => page.screenshot({
          type: "png",
          animations: "disabled",
          caret: "hide",
        }),
        60000,
        "Playwright screenshot",
      );
    } catch (playwrightError) {
      stage(`Playwright screenshot failed, falling back to CDP: ${playwrightError?.message || playwrightError}`);
      const result = await withTimeout(() => cdp.send("Page.captureScreenshot", {
        format: "png",
        fromSurface,
        captureBeyondViewport: false,
      }), 45000, "CDP screenshot");
      fullScreenshot = Buffer.from(result.data, "base64");
    }
    if (debugFullOutput) {
      await writeFile(debugFullOutput, fullScreenshot);
    }
    const metadata = await sharp(fullScreenshot).metadata();
    // Playwright 返回 CSS 像素，而截图是设备像素；左侧额外留出安全边界，避免最左卡片和边框被裁掉。
    const captureX = Math.max(0, box.x - CAPTURE_LEFT_PAD_CSS_PX);
    const captureY = Math.max(0, box.y - CAPTURE_TOP_PAD_CSS_PX);
    const left = Math.max(0, Math.round(captureX * deviceScale));
    const top = Math.max(0, Math.round(captureY * deviceScale));
    const width = Math.min(
      Math.max(1, Math.round((box.x + box.width - captureX) * deviceScale)),
      (metadata.width || 0) - left,
    );
    const height = Math.min(
      Math.max(1, Math.round((box.y + box.height - captureY) * deviceScale)),
      (metadata.height || 0) - top,
    );
    if (width <= 0 || height <= 0) {
      throw new Error(`Softtime screenshot bounds exceed viewport: ${metadata.width}x${metadata.height}`);
    }
    stage(
      `crop bounds: left=${left}, top=${top}, width=${width}, height=${height}, `
      + `screenshot=${metadata.width}x${metadata.height}, dpr=${deviceScale}, boxY=${box.y}`,
    );
    // 分阶段处理裁切、去白边和补边距，避免 Sharp 在同一个懒加载管线里产生边界冲突。
    const extracted = await sharp(fullScreenshot)
      .extract({ left, top, width, height })
      .png()
      .toBuffer();
    const trimmed = await sharp(extracted)
      .trim({ background: "#ffffff", threshold: 12 })
      .png()
      .toBuffer();
    const trimmedMetadata = await sharp(trimmed).metadata();
    const normalizedHeight = Math.min(
      trimmedMetadata.height || 0,
      Math.max(1, Math.round((trimmedMetadata.width || 1) / CATEGORY_OVERVIEW_TARGET_ASPECT)),
    );
    // 只保留三行市场概览，裁掉下方商品榜标题，使所有类目图片显示比例统一。
    const normalized = await sharp(trimmed)
      .extract({
        left: 0,
        top: 0,
        width: trimmedMetadata.width || 1,
        height: normalizedHeight,
      })
      .png()
      .toBuffer();
    const cropped = sharp(normalized)
      // 再去掉白边后补回极小边距，确保左侧边框不会紧贴图片边缘。
      .extend({
        top: CAPTURE_EDGE_PAD_PX,
        bottom: CAPTURE_EDGE_PAD_PX,
        left: CAPTURE_EDGE_PAD_PX,
        right: CAPTURE_EDGE_PAD_PX,
        background: "#ffffff",
      });
    const [pngBuffer, stats] = await Promise.all([
      cropped.clone().png().toBuffer(),
      cropped.stats(),
    ]);
    const hasMeaningfulPixels = stats.channels.some((channel) => (
      channel.min < 245 || channel.max - channel.min > 8 || channel.stdev > 2
    ));
    const blank = !hasMeaningfulPixels;
    if (!blank || writeBlank) {
      await writeFile(output, pngBuffer);
    }
    return { blank, width, height, fromSurface };
  } finally {
    await page.evaluate((original) => {
      const tool = document.querySelector("#sorftime-logo");
      if (!tool) return;
      if (original === null) tool.removeAttribute("style");
      else tool.setAttribute("style", original);
    }, floatingToolStyle).catch(() => {});
    await cdp.detach().catch(() => {});
  }
}

async function prepareCaptureWindow(page, context) {
  const cdp = await context.newCDPSession(page);
  const original = await cdp.send("Browser.getWindowForTarget");
  const bounds = original.bounds || {};
  await cdp.send("Browser.setWindowBounds", {
    windowId: original.windowId,
    bounds: { left: 0, top: 0, width: 1920, height: 1080, windowState: "normal" },
  });
  await page.bringToFront();
  await page.waitForTimeout(750);
  await page.setViewportSize({ width: 1920, height: 1080 });
  return async () => {
    try {
      await cdp.send("Browser.setWindowBounds", {
        windowId: original.windowId,
        bounds: {
          left: bounds.left,
          top: bounds.top,
          width: bounds.width,
          height: bounds.height,
          windowState: bounds.windowState || "normal",
        },
      });
    } finally {
      await cdp.detach().catch(() => {});
    }
  };
}

async function ensureAmazonPostalCode(page, postalCode) {
  const code = String(postalCode || "").trim();
  if (!code) {
    return { status: "skipped", text: "" };
  }
  const currentLocation = await locationText(page);
  if (postalMatches(currentLocation, code)) {
    return { status: "already-set", text: currentLocation };
  }
  for (const selector of ["#sp-cc-accept", "#sp-cc-rejectall-link"]) {
    try {
      const button = page.locator(selector).first();
      if (await button.isVisible({ timeout: 800 })) {
        await button.click({ timeout: 3000, noWaitAfter: true });
        await page.waitForTimeout(1000);
        break;
      }
    } catch {
      // Cookie controls are not always present.
    }
  }
  const postalInput = page.locator("#GLUXZipUpdateInput, input[name='zipCode'], input[name='postalCode'], input[placeholder*='zip' i], input[placeholder*='postal' i], input[aria-label*='zip' i], input[aria-label*='postal' i], input[aria-label*='Postleitzahl' i]").filter({ visible: true }).first();
  if (!(await postalInput.isVisible({ timeout: 5000 }).catch(() => false))) {
    const entry = page.locator("#nav-global-location-popover-link, #glow-ingress-block, #nav-global-location-slot").first();
    if (!(await entry.count())) {
      return { status: "failed", error: "Amazon location entry not found", text: await locationText(page) };
    }
    await entry.click({ timeout: 3000, force: true, noWaitAfter: true });
  }
  await postalInput.waitFor({ state: "visible", timeout: 5000 });
  await postalInput.fill(code, { timeout: 3000 });
  const apply = page.locator("#GLUXZipUpdate").first();
  await apply.click({ timeout: 3000, force: true, noWaitAfter: true });
  await page.waitForTimeout(1500);
  const confirmed = await successText(page);
  if (postalMatches(confirmed, code)) {
    await closeVisibleLocationConfirmation(page);
    return { status: "success", text: confirmed };
  }
  const current = await waitForPostalLocation(page, code, 5000);
  if (postalMatches(current, code)) {
    return { status: "success", text: current };
  }
  const errors = normalize((await page.locator("#GLUXZipError, #GLUXZipServerError").allTextContents()).join(" "));
  return {
    status: "failed",
    error: errors || `Amazon location text did not include target postal code ${code}`,
    text: current,
  };
}

async function main() {
  const args = parseArgs(process.argv);
  const url = args.url;
  const output = args.output;
  const debuggerAddress = args["debugger-address"] || "127.0.0.1:9223";
  const timeoutMs = Number(args["timeout-ms"] || 30000);
  const allowPartial = args["allow-partial"] === "1" || args["allow-partial"] === "true";
  const postalCode = args["postal-code"] || "";
  const debugFullOutput = args["debug-full-output"] || "";

  if (!url || !output) {
    throw new Error("Missing required --url or --output");
  }

  stage(`connecting to CDP ${debuggerAddress}`);
  const browser = await connectBrowserWithRetry(debuggerAddress);
  try {
    const context = browser.contexts()[0] || await browser.newContext({ viewport: { width: 1920, height: 1080 } });
    const targetHost = new URL(url).hostname;
    const existingPage = context.pages().find((candidate) => {
      try {
        return new URL(candidate.url()).hostname === targetHost && !/\/ap\/signin/i.test(candidate.url());
      } catch {
        return false;
      }
    });
    const page = existingPage || await withTimeout(() => context.newPage(), 5000, "new Amazon tab");
    const ownsPage = !existingPage;
    const restoreCaptureWindow = await prepareCaptureWindow(page, context);
    try {
      stage(`using ${existingPage ? "existing" : "new"} Amazon tab`);
      await withTimeout(() => page.setViewportSize({ width: 1920, height: 1080 }), 5000, "set viewport");
      stage("navigating to category");
      await withTimeout(() => page.goto(url, { waitUntil: "domcontentloaded", timeout: timeoutMs }), timeoutMs + 5000, "Amazon navigation");
      if (/\/ap\/signin/i.test(page.url())) {
        throw new Error("Amazon login required: page redirected to /ap/signin.");
      }
      stage("checking Amazon postal code");
      const postalResult = await withTimeout(() => ensureAmazonPostalCode(page, postalCode), Math.min(15000, timeoutMs), "Amazon postal-code check");
      if (!["skipped", "success", "already-set"].includes(postalResult.status)) {
        throw new Error(`Amazon delivery postal code was not confirmed: ${postalResult.status} ${postalResult.error || ""} ${postalResult.text || ""}`);
      }
      if (["success", "already-set"].includes(postalResult.status)) {
        stage(`postal code accepted: ${postalResult.text || postalResult.status}; reloading for verification`);
        await withTimeout(() => page.goto(url, { waitUntil: "domcontentloaded", timeout: timeoutMs }), timeoutMs + 5000, "Amazon reload");
        const verifiedLocation = await withTimeout(() => waitForPostalLocation(page, postalCode, 5000), 7000, "Amazon postal-code verification");
        if (!postalMatches(verifiedLocation, postalCode)) {
          throw new Error(`Amazon postal code verification failed after reload: ${verifiedLocation || "unknown"}`);
        }
        stage(`postal code verified: ${verifiedLocation}`);
      }
      stage("waiting for Softtime overview region");
      await withTimeout(() => page.waitForSelector(".fullPanoramaBox", { state: "visible", timeout: timeoutMs }), timeoutMs + 3000, "Softtime overview region");
      const overview = page.locator(".fullPanoramaBox");
      await overview.scrollIntoViewIfNeeded({ timeout: timeoutMs });
      const readinessStarted = Date.now();
      stage(`waiting fixed ${FIXED_CAPTURE_WAIT_MS}ms before capture`);
      // 插件区域出现后固定等待满 20 秒，避免结构先出现但 Canvas 尚未完成绘制。
      await page.waitForTimeout(FIXED_CAPTURE_WAIT_MS);
      const complete = await page.evaluate((selector) => {
        const el = document.querySelector(selector);
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const text = (el.innerText || "").replace(/\s+/g, " ");
        const canvases = Array.from(el.querySelectorAll("canvas"));
        const required = [
          /月总销量/,
          /月总销售额/,
          /Listing.*垄断/,
          /卖家.*垄断/,
          /品牌.*垄断|品牌销量垄断/,
          /平均售价/,
          /平均评价数量/,
          /平均星级/,
          /A\+视频占比/,
        ];
        return rect.width >= 300
          && rect.height >= 120
          && required.every((pattern) => pattern.test(text))
          && canvases.length >= 12
          && canvases.every((canvas) => canvas.width > 0 && canvas.height > 0);
      }, ".fullPanoramaBox");
      if (!complete && !allowPartial) {
        throw new Error("Softtime market overview is incomplete after the fixed 20-second wait.");
      }
      const text = await withTimeout(() => overview.innerText({ timeout: 3000 }), 5000, "Softtime overview text");
      let captureResult = null;
      let lastCaptureError = null;
      let blankAttempts = 0;
      const captureRetryDeadline = Date.now() + Math.min(3000, timeoutMs);
      // Canvas 尺寸有效并不等于已经完成绘制。对纯白截图短暂重试，避免把空白图写入 Excel。
      while (Date.now() < captureRetryDeadline) {
        try {
          captureResult = await captureOverviewWithCDP(page, context, output, {
            debugFullOutput,
          });
          if (!captureResult.blank) {
            break;
          }
          blankAttempts += 1;
          stage(`blank overview capture detected; retry ${blankAttempts}`);
        } catch (error) {
          lastCaptureError = error;
          stage(`overview capture retry: ${error?.message || error}`);
        }
        const remaining = Math.max(0, captureRetryDeadline - Date.now());
        if (!remaining) {
          break;
        }
        await page.waitForTimeout(Math.min(750, remaining));
      }
      // 最后再用 surface 截图兜底。即便页面确实没有更多像素，也保留最终截图，避免整行没有图片。
      if (!captureResult || captureResult.blank) {
        try {
          captureResult = await captureOverviewWithCDP(page, context, output, {
            fromSurface: true,
            writeBlank: true,
          });
        } catch (error) {
          lastCaptureError = error;
        }
      }
      if (!captureResult && lastCaptureError) {
        throw lastCaptureError;
      }
      stage(`capture complete: ${output}`);
      console.log(JSON.stringify({
        url: page.url(),
        monthly_sales: parseMonthlySales(text),
        complete,
        waited_ms: Date.now() - readinessStarted,
        text: text.replace(/\s+/g, " ").trim(),
        screenshot_blank: Boolean(captureResult?.blank),
        screenshot_from_surface: Boolean(captureResult?.fromSurface),
        amazon_delivery_zip_status: postalResult.status,
        amazon_delivery_location_text: postalResult.text || "",
      }));
    } finally {
      await restoreCaptureWindow().catch(() => {});
      if (ownsPage) {
        await page.close().catch(() => {});
      }
    }
  } finally {
    // 这是通过 connectOverCDP 接入的远程浏览器；关闭客户端连接，不关闭用户浏览器本身。
    if (typeof browser.close === "function") {
      browser._shouldCloseConnectionOnClose = false;
      await browser.close().catch(() => {});
    }
  }
}

main().catch((error) => {
  console.error(error && error.stack ? error.stack : String(error));
  process.exit(1);
});
