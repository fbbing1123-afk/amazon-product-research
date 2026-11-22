import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const { chromium } = require(
  "../product_research/scripts/chart-runtime/node_modules/playwright"
);

const ASIN_PATTERN = /^[A-Z0-9]{10}$/;

function parseArgs(argv) {
  const args = {
    debuggerAddress: "127.0.0.1:9223",
    outputDir: path.resolve("outputs", "competitor_product_cards"),
    viewportWidth: 380,
    viewportHeight: 1100,
    timeoutMs: 15000,
    urlIncludes: "/s?",
    asins: [],
  };
  for (let index = 0; index < argv.length; index += 1) {
    const name = argv[index];
    const value = argv[index + 1];
    if (name === "--asin" && value) {
      args.asins.push(value.trim().toUpperCase());
      index += 1;
    } else if (name === "--asins" && value) {
      args.asins.push(...value.split(/[\s,;/|]+/).map((item) => item.trim().toUpperCase()).filter(Boolean));
      index += 1;
    } else if (name === "--debugger-address" && value) {
      args.debuggerAddress = value;
      index += 1;
    } else if (name === "--output-dir" && value) {
      args.outputDir = path.resolve(value);
      index += 1;
    } else if (name === "--viewport-width" && value) {
      args.viewportWidth = Number.parseInt(value, 10);
      index += 1;
    } else if (name === "--viewport-height" && value) {
      args.viewportHeight = Number.parseInt(value, 10);
      index += 1;
    } else if (name === "--timeout-ms" && value) {
      args.timeoutMs = Number.parseInt(value, 10);
      index += 1;
    } else if (name === "--url-includes" && value) {
      args.urlIncludes = value;
      index += 1;
    }
  }
  args.asins = [...new Set(args.asins)];
  if (!args.asins.length || args.asins.some((asin) => !ASIN_PATTERN.test(asin))) {
    throw new Error("Pass one or more valid ASINs with --asin or --asins.");
  }
  return args;
}

function findTargetPage(browser, urlIncludes) {
  const pages = browser.contexts().flatMap((context) => context.pages());
  return (
    pages.find((page) => page.url().includes(urlIncludes)) ||
    pages.find((page) => /amazon\.[^/]+\/s\?/.test(page.url()))
  );
}

async function hideCaptureOverlays(page) {
  return page.addStyleTag({
    content: `
      #navbar,
      #nav-progressive-subnav,
      #nav-flyout-ewc,
      .nav-ewcFlyout,
      #sorftime-logo,
      #sorftime-logo-wrap,
      .toolBoxWrap-logo {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
      }
    `,
  });
}

function productCard(page, asin) {
  return page.locator(`.s-main-slot .s-result-item[data-asin="${asin}"]`).first();
}

async function waitForCard(page, asin, timeoutMs) {
  const card = productCard(page, asin);
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (await card.count()) {
      await card.waitFor({ state: "visible", timeout: Math.max(1000, deadline - Date.now()) });
      return card;
    }
    await page.evaluate(() => window.scrollBy(0, Math.max(600, window.innerHeight * 0.8)));
    await page.waitForTimeout(400);
  }
  throw new Error(`Amazon search result card not found for ${asin}.`);
}

async function waitForPluginMetrics(card, timeoutMs) {
  await card
    .waitFor({ state: "visible", timeout: timeoutMs })
    .then(() =>
      card.evaluate(
        (element, limit) =>
          new Promise((resolve) => {
            const ready = () => /Listing|ASIN|FBA|AI Analysis|AI \u5206\u6790/.test(element.innerText || "");
            if (ready()) return resolve(true);
            const observer = new MutationObserver(() => {
              if (ready()) {
                observer.disconnect();
                resolve(true);
              }
            });
            observer.observe(element, { childList: true, subtree: true, characterData: true });
            setTimeout(() => {
              observer.disconnect();
              resolve(false);
            }, limit);
          }),
        timeoutMs,
      ),
    )
    .catch(() => false);
}

async function captureCard(page, card, file) {
  const bounds = await card.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    return {
      width: rect.width,
      height: rect.height,
    };
  });
  if (bounds.width < 100 || bounds.height < 100) {
    throw new Error(`Invalid product card bounds: ${JSON.stringify(bounds)}`);
  }

  const originalStyle = await card.evaluate((element, frozenBounds) => {
    const cssText = element.style.cssText;
    for (const [name, value] of Object.entries({
      width: `${frozenBounds.width}px`,
      minWidth: `${frozenBounds.width}px`,
      maxWidth: `${frozenBounds.width}px`,
      height: `${frozenBounds.height}px`,
      minHeight: `${frozenBounds.height}px`,
      maxHeight: `${frozenBounds.height}px`,
      overflow: "hidden",
    })) {
      element.style.setProperty(
        name.replace(/[A-Z]/g, (letter) => `-${letter.toLowerCase()}`),
        value,
        "important",
      );
    }
    return cssText;
  }, bounds);

  try {
    return await card.screenshot({
      path: file,
      type: "png",
      animations: "disabled",
      caret: "hide",
      scale: "css",
      timeout: 15000,
    });
  } finally {
    await card.evaluate((element, cssText) => {
      element.style.cssText = cssText;
    }, originalStyle).catch(() => {});
  }
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  await mkdir(args.outputDir, { recursive: true });

  const browser = await chromium.connectOverCDP(`http://${args.debuggerAddress}`);
  const page = findTargetPage(browser, args.urlIncludes);
  if (!page) {
    throw new Error("No Amazon collection search page is open in the project browser.");
  }

  const originalViewport = page.viewportSize();
  let overlayStyle;
  const results = [];
  try {
    await page.bringToFront();
    await page.setViewportSize({ width: args.viewportWidth, height: args.viewportHeight });
    overlayStyle = await hideCaptureOverlays(page);
    await page.waitForTimeout(600);

    for (let index = 0; index < args.asins.length; index += 1) {
      const asin = args.asins[index];
      const card = await waitForCard(page, asin, args.timeoutMs);
      await card.scrollIntoViewIfNeeded({ timeout: args.timeoutMs });
      await waitForPluginMetrics(card, args.timeoutMs);
      await page.waitForTimeout(500);

      const file = path.join(
        args.outputDir,
        `${String(index + 1).padStart(2, "0")}_${asin}.png`,
      );
      const screenshot = await captureCard(page, card, file);
      results.push({
        asin,
        status: "success",
        file,
        width: screenshot.readUInt32BE(16),
        height: screenshot.readUInt32BE(20),
      });
    }

    await writeFile(
      path.join(args.outputDir, "manifest.json"),
      JSON.stringify({ page: page.url(), cards: results }, null, 2),
      "utf8",
    );
  } finally {
    if (overlayStyle) {
      await overlayStyle.evaluate((element) => element.remove()).catch(() => {});
    }
    if (originalViewport) {
      await page.setViewportSize(originalViewport).catch(() => {});
    }
  }

  console.log(JSON.stringify({ outputDir: args.outputDir, cards: results }, null, 2));
  process.exit(0);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
