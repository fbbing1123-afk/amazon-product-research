import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";

const require = createRequire(import.meta.url);
const runtimeDir = path.dirname(fileURLToPath(import.meta.url));

function requireFromOptionalRoot(moduleName, root) {
  if (!root) return null;
  try {
    return createRequire(path.join(root, "noop.js"))(moduleName);
  } catch (error) {
    if (process.env.CODEX_DEBUG_NODE_RESOLVE) {
      console.error(`Failed to load ${moduleName} from CODEX_NODE_MODULES=${root}: ${error.message}`);
    }
    return null;
  }
}

function loadModule(moduleName) {
  try {
    return require(moduleName);
  } catch {
    const fallback = requireFromOptionalRoot(moduleName, process.env.CODEX_NODE_MODULES);
    if (fallback) return fallback;
    throw new Error(
      `Cannot load ${moduleName}. CODEX_NODE_MODULES=${process.env.CODEX_NODE_MODULES || ""}. Run npm install or pnpm install in ${runtimeDir}.`,
    );
  }
}

function resolveModule(modulePath) {
  try {
    return require.resolve(modulePath);
  } catch {
    if (process.env.CODEX_NODE_MODULES) {
      try {
        return createRequire(path.join(process.env.CODEX_NODE_MODULES, "noop.js")).resolve(modulePath);
      } catch {
        // Fall through to the distribution-friendly error below.
      }
    }
    throw new Error(
      `Cannot resolve ${modulePath}. Run npm install or pnpm install in ${runtimeDir}.`,
    );
  }
}

const { chromium } = loadModule("playwright");
const vendoredEchartsPath = path.join(runtimeDir, "vendor", "echarts.min.js");
const echartsPath = await fs
  .access(vendoredEchartsPath)
  .then(() => vendoredEchartsPath)
  .catch(() => resolveModule("echarts/dist/echarts.min.js"));

const SUMMARY_PATH = path.resolve(
  process.argv[2] ?? "../../work/market_analysis/analysis/summary.json",
);
const OUTPUT_DIR = path.resolve(process.argv[3] ?? "../../work/custom_template/echarts_section3_charts");
const SCALE = 3;

const DISPLAY_IMAGES = {
  B10: [210, 281],
  D10: [190, 281],
  F10: [190, 281],
  B11: [210, 253],
  D11: [190, 253],
  F11: [190, 253],
  B12: [501, 253],
  B13: [501, 253],
  B14: [501, 253],
};

const TEXT = "#17324D";
const MUTED = "#64748B";
const GRID = "#E5E7EB";
const BLUE = "#4472C4";
const ORANGE = "#ED7D31";
const GREEN = "#00A050";
const COLORS = ["#4472C4", "#ED7D31", "#FFC000", "#70AD47", "#5B9BD5", "#A5A5A5"];
const FONT = '"Microsoft YaHei", "Noto Sans SC", Arial, sans-serif';

function num(value, fallback = 0) {
  if (value === null || value === undefined || value === "") return fallback;
  const parsed = Number(String(value).replace(/[$€£¥￥,%\s,]/g, ""));
  return Number.isFinite(parsed) ? parsed : fallback;
}

function fieldValue(row, baseName, exactNames = []) {
  for (const name of exactNames) {
    if (Object.prototype.hasOwnProperty.call(row, name)) return row[name];
  }
  if (Object.prototype.hasOwnProperty.call(row, baseName)) return row[baseName];
  const prefix = `${baseName}(`;
  for (const [key, value] of Object.entries(row)) {
    if (String(key).startsWith(prefix)) return value;
  }
  return undefined;
}

function metricValue(row, metric) {
  if (!metric) return 1;
  if (metric === "\u6708\u9500\u552e\u989d") {
    return num(fieldValue(row, "\u6708\u9500\u552e\u989d", [
      "\u6708\u9500\u552e\u989d($)",
      "\u6708\u9500\u552e\u989d(€)",
      "\u6708\u9500\u552e\u989d(£)",
      "\u6708\u9500\u552e\u989d(¥)",
      "\u6708\u9500\u552e\u989d(￥)",
    ]));
  }
  if (metric === "\u4ef7\u683c") {
    return num(fieldValue(row, "\u4ef7\u683c", [
      "\u4ef7\u683c($)",
      "\u4ef7\u683c(€)",
      "\u4ef7\u683c(£)",
      "\u4ef7\u683c(¥)",
      "\u4ef7\u683c(￥)",
    ]));
  }
  return num(row[metric]);
}

function inferCurrencySymbol(records) {
  for (const row of records) {
    for (const key of Object.keys(row)) {
      const match = String(key).match(/\(([$€£¥￥])\)/);
      if (match && (String(key).includes("\u4ef7\u683c") || String(key).includes("\u9500\u552e\u989d"))) {
        return match[1];
      }
    }
  }
  return "$";
}

function grouped(records, key, metric = null) {
  const map = new Map();
  for (const row of records) {
    const label = String(row[key] || "\u7a7a\u767d");
    map.set(label, (map.get(label) || 0) + metricValue(row, metric));
  }
  return [...map.entries()].sort((a, b) => b[1] - a[1]);
}

function groupedRecentYears(records, metric = null, yearCount = 6) {
  const currentYear = new Date().getFullYear();
  const years = Array.from({ length: yearCount }, (_, index) => String(currentYear - index));
  const map = new Map(years.map((year) => [year, 0]));
  for (const row of records) {
    const raw = String(row["\u4e0a\u67b6\u65f6\u95f4"] || "");
    const year = /^\d{4}/.test(raw) ? raw.slice(0, 4) : "";
    if (map.has(year)) {
      map.set(year, map.get(year) + metricValue(row, metric));
    }
  }
  return [...map.entries()];
}

const BRAND_ANALYSIS_TOP_N = 10;

function topNWithOther(rows, n = 5) {
  const top = rows.slice(0, n);
  const other = rows.slice(n).reduce((sum, [, value]) => sum + value, 0);
  if (other > 0) top.push(["\u5176\u5b83", other]);
  return top;
}

function truncateLabel(label, maxChars) {
  const text = String(label || "\u7a7a\u767d").trim() || "\u7a7a\u767d";
  if (text.length <= maxChars) return text;
  return `${text.slice(0, Math.max(1, maxChars - 1))}…`;
}

function formatTotal(value, decimals = 0) {
  return Number(value || 0).toLocaleString("en-US", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  });
}

function niceAxis(value, floor = null) {
  const baseValue = floor ? Math.max(value, floor) : value;
  if (baseValue <= 0) return 1;
  const padded = baseValue * 1.08;
  const mag = 10 ** Math.floor(Math.log10(padded));
  const norm = padded / mag;
  if (norm <= 2) return 2 * mag;
  if (norm <= 3) return 3 * mag;
  if (norm <= 5) return 5 * mag;
  return 10 * mag;
}

function pieOption({ title, value, rows, compact = false, width }) {
  const maxName = compact && rows.length > 8 ? 9 : width <= 190 ? 12 : 16;
  const total = rows.reduce((sum, [, itemValue]) => sum + itemValue, 0) || 1;
  const data = rows
    .filter(([, itemValue]) => itemValue > 0)
    .map(([name, itemValue]) => ({
      name: `${truncateLabel(name, maxName)} ${Math.round((itemValue / total) * 100)}%`,
      value: itemValue,
    }));
  // 品牌较多时为图例预留独立下半区，避免圆环与首行图例重叠。
  const denseLegend = compact && data.length > 8;
  return {
    backgroundColor: "#FFFFFF",
    color: COLORS,
    animation: false,
    textStyle: { fontFamily: FONT, color: TEXT },
    title: {
      text: title,
      subtext: value,
      left: "center",
      top: compact ? 4 : 8,
      itemGap: 1,
      textStyle: {
        color: TEXT,
        fontSize: compact ? 12 : 13,
        fontWeight: 700,
        overflow: "truncate",
        width: width - 12,
      },
      subtextStyle: {
        color: GREEN,
        fontSize: compact ? 15 : 16,
        fontWeight: 800,
      },
    },
    tooltip: { show: false },
    legend: {
      type: "plain",
      left: 4,
      right: 4,
      top: denseLegend ? 151 : undefined,
      bottom: denseLegend ? undefined : compact ? 3 : 8,
      itemWidth: denseLegend ? 7 : 9,
      itemHeight: denseLegend ? 7 : 9,
      itemGap: denseLegend ? 5 : compact ? 7 : 8,
      textStyle: {
        color: TEXT,
        fontFamily: FONT,
        fontSize: denseLegend ? 9 : compact ? 10 : 11,
        lineHeight: denseLegend ? 13 : compact ? 14 : 16,
      },
    },
    series: [
      {
        type: "pie",
        radius: denseLegend ? [28, 43] : compact ? ["31%", "48%"] : ["43%", "68%"],
        center: ["50%", denseLegend ? 95 : compact ? "44%" : "45%"],
        avoidLabelOverlap: true,
        percentPrecision: 0,
        stillShowZeroSum: false,
        itemStyle: { borderColor: "#FFFFFF", borderWidth: 1.5 },
        label: {
          show: !compact,
          position: "inside",
          formatter: "{d}%",
          color: "#FFFFFF",
          fontSize: compact ? 9 : 11,
          fontFamily: FONT,
          fontWeight: 500,
        },
        labelLine: { show: false },
        data,
      },
    ],
  };
}

function comboOption({ title, rows, barMetric, lineMetric, barName, lineName, barMaxOverride = null }) {
  const lineValues = rows.map((row) => metricValue(row, lineMetric));
  const barValues = rows.map((row) => metricValue(row, barMetric));
  const barMax = barMaxOverride ?? niceAxis(Math.max(...barValues, 0));
  const lineMax = niceAxis(Math.max(...lineValues, 0));
  const categories = rows.map((_row, index) => String(index + 1));
  const dense = categories.length > 80;
  const xLabelInterval = dense ? Math.max(0, Math.ceil(categories.length / 12) - 1) : 3;
  return {
    backgroundColor: "#FFFFFF",
    color: [ORANGE, BLUE],
    animation: false,
    textStyle: { fontFamily: FONT, color: TEXT },
    title: {
      text: title,
      left: "center",
      top: 4,
      textStyle: { color: TEXT, fontSize: 13, fontWeight: 700 },
    },
    grid: { left: 45, right: 43, top: 38, bottom: 42 },
    legend: {
      left: "center",
      bottom: 3,
      itemWidth: 20,
      itemHeight: 8,
      itemGap: 18,
      textStyle: { color: TEXT, fontFamily: FONT, fontSize: 11 },
      data: [barName, lineName],
    },
    tooltip: { show: false },
    xAxis: {
      type: "category",
      data: categories,
      boundaryGap: true,
      axisTick: { show: false },
      axisLine: { lineStyle: { color: "#CBD5E1" } },
      axisLabel: { color: MUTED, fontSize: 9, interval: xLabelInterval, margin: 8 },
    },
    yAxis: [
      {
        type: "value",
        min: 0,
        max: barMax,
        splitNumber: 5,
        axisLabel: {
          color: MUTED,
          fontSize: 9,
          formatter: "{value}",
        },
        axisLine: { show: false },
        axisTick: { show: false },
        splitLine: { lineStyle: { color: GRID, width: 1 } },
      },
      {
        type: "value",
        min: 0,
        max: lineMax,
        splitNumber: 5,
        axisLabel: {
          color: MUTED,
          fontSize: 9,
          formatter: "{value}",
        },
        axisLine: { show: false },
        axisTick: { show: false },
        splitLine: { show: false },
      },
    ],
    series: [
      {
        name: barName,
        type: "bar",
        yAxisIndex: 0,
        data: barValues,
        barMaxWidth: dense ? 4 : 7,
        itemStyle: { color: ORANGE, borderRadius: [2, 2, 0, 0] },
      },
      {
        name: lineName,
        type: "line",
        yAxisIndex: 1,
        data: lineValues,
        smooth: 0.18,
        symbol: "circle",
        symbolSize: dense ? 2.5 : 4,
        lineStyle: { color: BLUE, width: dense ? 1.8 : 2.4 },
        itemStyle: { color: BLUE, borderColor: "#FFFFFF", borderWidth: 1 },
      },
    ],
  };
}

async function renderOption(browser, cell, option) {
  const [width, height] = DISPLAY_IMAGES[cell];
  const context = await browser.newContext({
    viewport: { width, height },
    deviceScaleFactor: SCALE,
  });
  const page = await context.newPage();
  await page.setContent(`<!doctype html>
    <html>
      <head>
        <meta charset="utf-8" />
        <style>
          html, body { margin: 0; padding: 0; width: ${width}px; height: ${height}px; background: #fff; overflow: hidden; }
          #chart { width: ${width}px; height: ${height}px; background: #fff; }
        </style>
      </head>
      <body><div id="chart"></div></body>
    </html>`);
  await page.addScriptTag({ path: echartsPath });
  await page.evaluate((chartOption) => {
    const el = document.getElementById("chart");
    const chart = window.echarts.init(el, null, { renderer: "svg" });
    chart.setOption(chartOption);
  }, option);
  await page.locator("#chart").screenshot({
    path: path.join(OUTPUT_DIR, `${cell}_hires.png`),
    omitBackground: false,
  });
  await context.close();
}

async function findBrowserExecutable() {
  const candidates = [
    process.env.CHROME_EXECUTABLE,
    // 兼容同事在用户目录安装的浏览器，直接复用，无需重复下载。
    process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, "Google/Chrome/Application/chrome.exe"),
    process.env.LOCALAPPDATA && path.join(process.env.LOCALAPPDATA, "Microsoft/Edge/Application/msedge.exe"),
    "C:/Program Files/Google/Chrome/Application/chrome.exe",
    "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
    "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
    "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
  ].filter(Boolean);
  for (const candidate of candidates) {
    try {
      await fs.access(candidate);
      return candidate;
    } catch {
      // Try the next installed browser candidate.
    }
  }
  return null;
}

async function main() {
  const summary = JSON.parse(await fs.readFile(SUMMARY_PATH, "utf8"));
  const records = summary.source_records || [];
  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  const currencySymbol = inferCurrencySymbol(records);
  const revenueTitle = `\u6708\u9500\u552e\u989d(${currencySymbol})`;
  const priceTitle = `\u4ef7\u683c(${currencySymbol})`;

  const brandCount = grouped(records, "\u54c1\u724c");
  const brandSales = grouped(records, "\u54c1\u724c", "\u6708\u9500\u91cf");
  const brandRevenue = grouped(records, "\u54c1\u724c", "\u6708\u9500\u552e\u989d");
  const recentYearCount = groupedRecentYears(records);
  const recentYearSales = groupedRecentYears(records, "\u6708\u9500\u91cf");
  const recentYearRevenue = groupedRecentYears(records, "\u6708\u9500\u552e\u989d");

  const specs = {
    B10: pieOption({
      title: "\u54c1\u724c (\u8ba1\u6570)",
      value: formatTotal(records.length),
      rows: topNWithOther(brandCount, BRAND_ANALYSIS_TOP_N),
      compact: true,
      width: DISPLAY_IMAGES.B10[0],
    }),
    D10: pieOption({
      title: "\u6708\u9500\u91cf (\u6c42\u548c)",
      value: formatTotal(brandSales.reduce((sum, [, value]) => sum + value, 0)),
      rows: topNWithOther(brandSales, BRAND_ANALYSIS_TOP_N),
      compact: true,
      width: DISPLAY_IMAGES.D10[0],
    }),
    F10: pieOption({
      title: revenueTitle,
      value: formatTotal(Math.trunc(brandRevenue.reduce((sum, [, value]) => sum + value, 0))),
      rows: topNWithOther(brandRevenue, BRAND_ANALYSIS_TOP_N),
      compact: true,
      width: DISPLAY_IMAGES.F10[0],
    }),
    B11: pieOption({
      title: "\u8fd16\u5e74\u4ea7\u54c1\u6570",
      value: formatTotal(recentYearCount.reduce((sum, [, value]) => sum + value, 0)),
      rows: recentYearCount,
      compact: true,
      width: DISPLAY_IMAGES.B11[0],
    }),
    D11: pieOption({
      title: "\u8fd16\u5e74\u6708\u9500\u91cf\u5360\u6bd4",
      value: formatTotal(recentYearSales.reduce((sum, [, value]) => sum + value, 0)),
      rows: recentYearSales,
      compact: true,
      width: DISPLAY_IMAGES.D11[0],
    }),
    F11: pieOption({
      title: `\u8fd16\u5e74\u6708\u9500\u552e\u989d(${currencySymbol})`,
      value: formatTotal(Math.trunc(recentYearRevenue.reduce((sum, [, value]) => sum + value, 0))),
      rows: recentYearRevenue,
      compact: true,
      width: DISPLAY_IMAGES.F11[0],
    }),
    B12: comboOption({
      title: "Review\u6570\u91cf\u5206\u6790",
      rows: [...records].sort((a, b) => num(b["\u8bc4\u5206\u6570"]) - num(a["\u8bc4\u5206\u6570"])),
      barMetric: "\u8bc4\u5206\u6570",
      lineMetric: "\u6708\u9500\u91cf",
      barName: "\u8bc4\u5206\u6570",
      lineName: "\u6708\u9500\u91cf",
    }),
    B13: comboOption({
      title: "\u8bc4\u5206\u5206\u6790",
      rows: [...records].sort((a, b) => num(b["\u8bc4\u5206"]) - num(a["\u8bc4\u5206"])),
      barMetric: "\u8bc4\u5206",
      lineMetric: "\u6708\u9500\u91cf",
      barName: "\u8bc4\u5206",
      lineName: "\u6708\u9500\u91cf",
      barMaxOverride: 5,
    }),
    B14: comboOption({
      title: "\u4ef7\u683c\u5206\u6790",
      rows: [...records].sort((a, b) => metricValue(b, "\u4ef7\u683c") - metricValue(a, "\u4ef7\u683c")),
      barMetric: "\u4ef7\u683c",
      lineMetric: "\u6708\u9500\u91cf",
      barName: priceTitle,
      lineName: "\u6708\u9500\u91cf",
    }),
  };

  const executablePath = await findBrowserExecutable();
  const browser = await chromium.launch({ headless: true, executablePath });
  try {
    for (const [cell, option] of Object.entries(specs)) {
      await renderOption(browser, cell, option);
      console.log(`rendered ${cell}`);
    }
  } finally {
    await browser.close();
  }
}

await main();

