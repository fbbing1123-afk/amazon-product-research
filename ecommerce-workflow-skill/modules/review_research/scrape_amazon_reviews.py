from __future__ import annotations

import argparse
import json
import math
import re
import shutil
import sys
import tempfile
import time
import urllib.request
import unicodedata
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import urlencode

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


MODULES_DIR = Path(__file__).resolve().parents[1]
if str(MODULES_DIR) not in sys.path:
    sys.path.insert(0, str(MODULES_DIR))

from shared.browser_runtime import get_debugger_address
from shared.marketplaces import (
    amazon_location_text,
    amazon_postal_code_matches,
    ensure_amazon_postal_code,
    get_marketplace,
)
from shared.image_conversion import restore_unloaded_floating_images
from shared.runtime_monitor import monitor_event
from review_research.pain_phrase_analysis import (
    SemanticGroup,
    aggregate_semantic_groups,
    analyze_pain_phrases,
    normalize_semantic_groups,
    semantic_group_for_phrase,
)


SUMMARY_SHEET = "评论抓取摘要"
DETAIL_SHEET = "商品评论"
PAIN_POINT_SHEET = "差评痛点"
PAIN_FREQUENCY_SHEET = "差评词频分析"
PAIN_FREQUENCY_SECTION = "差评词频分析"
DEFAULT_TRANSLATION_CACHE = Path(__file__).resolve().parents[2] / "resources" / "review_translations.json"
DEFAULT_PAIN_GROUP_CACHE = Path(__file__).resolve().parents[2] / "resources" / "pain_phrase_groups.json"
ASIN_HEADER_PRIORITY = (
    {"竞品asin"},
    {"商品asin", "asin"},
    {"子asin"},
    {"父asin"},
)
ASIN_PATTERN = re.compile(r"(?<![A-Z0-9])([A-Z0-9]{10})(?![A-Z0-9])", re.I)
REVIEW_PATH_PATTERN = re.compile(
    r"/(?:dp|gp/product|product-reviews|portal/customer-reviews)/([A-Z0-9]{10})(?:[/?#]|$)",
    re.I,
)
SORT_OPTIONS = {"recent": "recent", "helpful": "helpful"}
REVIEWER_TYPE_OPTIONS = {
    "default": None,
    "all": "all_reviews",
    "verified": "avp_only_reviews",
}
STAR_OPTIONS = {
    "all_stars",
    "positive",
    "critical",
    "five_star",
    "four_star",
    "three_star",
    "two_star",
    "one_star",
}


EXTRACT_REVIEWS_JS = r"""() => {
  const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
  const bodyText = normalize(document.body?.innerText || '');
  const url = location.href;
  const lowerUrl = url.toLowerCase();
  const login = lowerUrl.includes('/ap/signin')
    || !!document.querySelector('input[type="password"], input[name="password"]');
  const captcha = !!document.querySelector('#captchacharacters, form[action*="validateCaptcha" i]')
    || /enter the characters you see|geben sie die angezeigten zeichen|captcha/i.test(bodyText);
  const pageNotFound = /page not found/i.test(document.title || '')
    || /sorry!? we couldn.?t find that page/i.test(bodyText);
  const reviewNodes = [...document.querySelectorAll('[data-hook="review"]')];
  const productTitle = normalize(
    document.querySelector('[data-hook="product-link"]')?.textContent
    || document.querySelector('.product-title')?.textContent
    || document.querySelector('#cm_cr-product_info .a-size-large')?.textContent
    || ''
  );
  const reviews = reviewNodes.map((node) => {
    const text = (selector) => normalize(node.querySelector(selector)?.textContent || '');
    const attr = (selector, name) => node.querySelector(selector)?.getAttribute(name) || '';
    const starText = text('[data-hook="review-star-rating"] .a-icon-alt')
      || text('[data-hook="cmps-review-star-rating"] .a-icon-alt')
      || text('.review-rating .a-icon-alt')
      || attr('[data-hook="review-star-rating"]', 'aria-label');
    const titleParts = [...node.querySelectorAll('[data-hook="review-title"] span')]
      .map((part) => normalize(part.textContent))
      .filter((part) => part && !/(out of 5 stars|von 5 sternen|sur 5 etoiles|su 5 stelle)/i.test(part));
    const imageUrls = [...node.querySelectorAll('[data-hook="review-image-tile"] img, .review-image-tile-section img')]
      .map((image) => image.currentSrc || image.src || image.getAttribute('data-src') || '')
      .filter(Boolean);
    const reviewLink = node.querySelector('[data-hook="review-title"]')?.href
      || node.querySelector('a[href*="customer-reviews"]')?.href
      || '';
    return {
      reviewId: node.id || node.getAttribute('data-review-id') || '',
      reviewerName: text('.a-profile-name'),
      ratingText: starText,
      title: titleParts.at(-1) || text('[data-hook="review-title"]'),
      body: text('[data-hook="review-body"]'),
      dateText: text('[data-hook="review-date"]'),
      variant: text('[data-hook="format-strip"]'),
      verifiedText: text('[data-hook="avp-badge"]'),
      vineText: text('[data-hook="vine-review-badge"]') || text('.cr-vine-review-badge'),
      helpfulText: text('[data-hook="helpful-vote-statement"]'),
      imageUrls,
      hasVideo: !!node.querySelector('video, [data-hook="video-block"], [class*="video" i]'),
      reviewUrl: reviewLink,
    };
  });
  const noReviews = /no customer reviews|no reviews|noch keine kundenrezensionen|keine rezensionen|0 global ratings|0 globale bewertungen/i.test(bodyText);
  return {
    url,
    login,
    captcha,
    pageNotFound,
    noReviews,
    productTitle,
    reviews,
    hasShowMore: !!document.querySelector('[data-hook="show-more-button"]'),
    filters: {
      sortBy: document.querySelector('#sort-order-dropdown')?.value || '',
      reviewerType: document.querySelector('#reviewer-type-dropdown')?.value || '',
      starFilter: document.querySelector('#star-count-dropdown')?.value || '',
    },
    pageTitle: document.title || '',
  };
}"""


@dataclass
class AsinResult:
    asin: str
    review_source_asin: str = ""
    star_filter: str = "all_stars"
    reviewer_type: str = "default"
    sort_by: str = "recent"
    product_title: str = ""
    product_url: str = ""
    first_review_url: str = ""
    requested_pages: int = 0
    scanned_pages: int = 0
    reviews: list[dict[str, Any]] = field(default_factory=list)
    status: str = "success"
    error: str = ""
    fallback_reason: str = ""


class AmazonAuthenticationRequired(RuntimeError):
    pass


def require_sync_playwright():
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as exc:
        raise RuntimeError(
            "未安装 Python Playwright。请使用项目现有 Python 环境运行，或执行 python -m pip install playwright。"
        ) from exc
    return sync_playwright


def normalize_header(value: Any) -> str:
    return re.sub(r"[\s_\-]+", "", str(value or "")).lower()


def normalize_asin(value: Any) -> str | None:
    text = str(value or "").strip().upper()
    path_match = REVIEW_PATH_PATTERN.search(text)
    if path_match:
        return path_match.group(1).upper()
    if re.fullmatch(r"[A-Z0-9]{10}", text):
        return text
    match = ASIN_PATTERN.search(text)
    return match.group(1).upper() if match else None


def choose_source_sheet(workbook):
    for worksheet in workbook.worksheets:
        if worksheet.title.endswith("数据源"):
            return worksheet
    for name in ("数据源", "US", "DE", "UK", "Sheet1"):
        if name in workbook.sheetnames:
            return workbook[name]
    return workbook.worksheets[0]


def extract_asins_from_workbook(source: Path) -> list[str]:
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        worksheet = choose_source_sheet(workbook)
        header_row = None
        asin_columns: list[int] = []
        for row_index, row in enumerate(
            worksheet.iter_rows(min_row=1, max_row=min(10, worksheet.max_row), values_only=True), start=1
        ):
            normalized = [normalize_header(value) for value in row]
            columns = []
            for aliases in ASIN_HEADER_PRIORITY:
                columns = [index for index, value in enumerate(normalized) if value in aliases]
                if columns:
                    break
            if columns:
                header_row = row_index
                asin_columns = columns
                break
        if header_row is None:
            raise ValueError(f"工作表“{worksheet.title}”前 10 行中找不到 ASIN 或竞品ASIN 表头。")

        found: list[str] = []
        seen: set[str] = set()
        for row in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
            for column in asin_columns:
                asin = normalize_asin(row[column] if column < len(row) else None)
                if asin and asin not in seen:
                    seen.add(asin)
                    found.append(asin)
        return found
    finally:
        workbook.close()


def extract_competitor_asins_from_workbook(source: Path) -> list[str]:
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        worksheet = next(
            (item for item in workbook.worksheets if item.title.endswith("竞品分析")),
            None,
        )
        if worksheet is None:
            return []
        found: list[str] = []
        seen: set[str] = set()
        for row in worksheet.iter_rows(
            min_row=1,
            max_row=min(5, worksheet.max_row),
            values_only=False,
        ):
            for cell in row:
                values = [cell.value]
                hyperlink = getattr(cell, "hyperlink", None)
                if hyperlink and hyperlink.target:
                    values.append(hyperlink.target)
                for value in values:
                    asin = normalize_asin(value)
                    if asin and asin not in seen:
                        seen.add(asin)
                        found.append(asin)
        return found
    finally:
        workbook.close()


def extract_review_asins_from_workbook(source: Path) -> list[str]:
    competitor_asins = extract_competitor_asins_from_workbook(source)
    return competitor_asins or extract_asins_from_workbook(source)


def extract_product_context_from_workbook(source: Path) -> dict[str, str]:
    """读取数据源中的标题和卖点，供评论商品一致性校验使用。"""
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        worksheet = choose_source_sheet(workbook)
        header_row = None
        headers: dict[str, int] = {}
        for row_index, row in enumerate(
            worksheet.iter_rows(min_row=1, max_row=min(10, worksheet.max_row), values_only=True),
            start=1,
        ):
            normalized = [normalize_header(value) for value in row]
            if any(value in {"asin", "商品asin", "竞品asin", "子asin"} for value in normalized):
                header_row = row_index
                headers = {value: index for index, value in enumerate(normalized) if value}
                break
        if header_row is None:
            return {}
        asin_column = next(
            (
                headers.get(alias)
                for alias in ("竞品asin", "商品asin", "asin", "子asin")
                if headers.get(alias) is not None
            ),
        )
        if asin_column is None:
            return {}
        title_column = next(
            (
                headers.get(alias)
                for alias in ("商品标题", "标题", "商品名称")
                if headers.get(alias) is not None
            ),
        )
        selling_points_column = next(
            (
                headers.get(alias)
                for alias in ("产品卖点", "卖点", "五点")
                if headers.get(alias) is not None
            ),
        )
        contexts: dict[str, str] = {}
        for row in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
            asin = normalize_asin(row[asin_column] if asin_column < len(row) else None)
            if not asin:
                continue
            values = []
            if title_column is not None and title_column < len(row):
                values.append(str(row[title_column] or ""))
            if selling_points_column is not None and selling_points_column < len(row):
                values.append(str(row[selling_points_column] or ""))
            contexts[asin] = " ".join(value for value in values if value).strip()
        return contexts
    finally:
        workbook.close()


def extract_parent_asins_from_workbook(source: Path) -> dict[str, str]:
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        worksheet = choose_source_sheet(workbook)
        header_row = None
        asin_column = None
        parent_column = None
        for row_index, row in enumerate(
            worksheet.iter_rows(min_row=1, max_row=min(10, worksheet.max_row), values_only=True),
            start=1,
        ):
            normalized = [normalize_header(value) for value in row]
            primary_columns = [
                index
                for aliases in ASIN_HEADER_PRIORITY[:-1]
                for index, value in enumerate(normalized)
                if value in aliases
            ]
            parent_columns = [
                index for index, value in enumerate(normalized) if value == "父asin"
            ]
            if primary_columns and parent_columns:
                header_row = row_index
                asin_column = primary_columns[0]
                parent_column = parent_columns[0]
                break
        if header_row is None or asin_column is None or parent_column is None:
            return {}

        mapping: dict[str, str] = {}
        for row in worksheet.iter_rows(min_row=header_row + 1, values_only=True):
            asin = normalize_asin(row[asin_column] if asin_column < len(row) else None)
            parent = normalize_asin(row[parent_column] if parent_column < len(row) else None)
            if asin and parent and asin != parent:
                mapping.setdefault(asin, parent)
        return mapping
    finally:
        workbook.close()


def parse_rating(value: Any) -> float | None:
    match = re.search(r"([0-5](?:[.,]\d+)?)", str(value or ""))
    if not match:
        return None
    return float(match.group(1).replace(",", "."))


def parse_helpful_count(value: Any) -> int:
    text = str(value or "").strip()
    match = re.search(r"([\d.,]+)", text)
    if match:
        digits = re.sub(r"\D", "", match.group(1))
        return int(digits) if digits else 0
    if re.search(r"\bone\b|\beine?\b|\buna?\b", text, re.I):
        return 1
    return 0


MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12,
    "januar": 1, "februar": 2, "maerz": 3, "marz": 3, "märz": 3, "april": 4, "mai": 5, "juni": 6,
    "juli": 7, "august": 8, "september": 9, "oktober": 10, "november": 11, "dezember": 12,
}


def parse_review_date(value: Any) -> date | None:
    text = str(value or "").strip().lower()
    normalized = text.replace("ä", "ae")
    for month_name, month in MONTHS.items():
        candidate = month_name.replace("ä", "ae")
        if candidate not in normalized:
            continue
        us_match = re.search(rf"{re.escape(candidate)}\s+(\d{{1,2}}),?\s+(\d{{4}})", normalized)
        if us_match:
            return date(int(us_match.group(2)), month, int(us_match.group(1)))
        eu_match = re.search(rf"(\d{{1,2}})\.?\s+{re.escape(candidate)}\s+(\d{{4}})", normalized)
        if eu_match:
            return date(int(eu_match.group(2)), month, int(eu_match.group(1)))
    iso_match = re.search(r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})", normalized)
    if iso_match:
        return date(int(iso_match.group(1)), int(iso_match.group(2)), int(iso_match.group(3)))
    return None


def normalize_review(
    raw: dict[str, Any],
    *,
    asin: str,
    review_source_asin: str | None = None,
    page_number: int,
    source_url: str,
) -> dict[str, Any]:
    return {
        "asin": asin,
        "review_source_asin": review_source_asin or asin,
        "review_id": str(raw.get("reviewId") or "").strip(),
        "reviewer_name": str(raw.get("reviewerName") or "").strip(),
        "rating": parse_rating(raw.get("ratingText")),
        "title": str(raw.get("title") or "").strip(),
        "body": str(raw.get("body") or "").strip(),
        "date": parse_review_date(raw.get("dateText")),
        "date_text": str(raw.get("dateText") or "").strip(),
        "variant": str(raw.get("variant") or "").strip(),
        "verified": bool(str(raw.get("verifiedText") or "").strip()),
        "vine": bool(str(raw.get("vineText") or "").strip()),
        "vine_text": str(raw.get("vineText") or "").strip(),
        "helpful_count": parse_helpful_count(raw.get("helpfulText")),
        "image_urls": list(dict.fromkeys(str(url) for url in raw.get("imageUrls") or [] if url)),
        "has_video": bool(raw.get("hasVideo")),
        "review_url": str(raw.get("reviewUrl") or "").strip(),
        "page_number": page_number,
        "source_url": source_url,
    }


def review_is_trusted(review: dict[str, Any]) -> bool:
    return bool(review.get("verified") or review.get("vine"))


def review_matches_star_filter(review: dict[str, Any], star_filter: str) -> bool:
    rating = review.get("rating")
    if rating is None or star_filter in {"all_stars", "all"}:
        return star_filter in {"all_stars", "all"}
    if star_filter == "positive":
        return 4 <= rating <= 5
    if star_filter == "critical":
        return 1 <= rating <= 3
    if star_filter == "five_star":
        return rating == 5
    if star_filter == "four_star":
        return rating == 4
    if star_filter == "three_star":
        return rating == 3
    if star_filter == "two_star":
        return rating == 2
    if star_filter == "one_star":
        return rating == 1
    return True


# Amazon 父体评论池偶尔会把其他国家或其他商品的评论混入当前站点页面。
# 通过评论卡片中的“Reviewed in ...”地点做第一层一致性校验，避免污染竞品分析。
REVIEW_LOCATION_PATTERNS = {
    "US": (
        "reviewed in the united states",
        "reviewed in united states",
        "reviewed in usa",
        "bewertet in den vereinigten staaten",
    ),
    "UK": (
        "reviewed in the united kingdom",
        "reviewed in united kingdom",
        "bewertet im vereinigten königreich",
        "évalué au royaume-uni",
    ),
    "DE": (
        "reviewed in germany",
        "bewertet in deutschland",
        "évalué en allemagne",
    ),
    "CA": (
        "reviewed in canada",
        "bewertet in kanada",
        "évalué au canada",
    ),
    "FR": (
        "reviewed in france",
        "bewertet in frankreich",
        "évalué en france",
    ),
}


def review_matches_marketplace(review: dict[str, Any], site: str) -> bool:
    """检查评论卡片显示的评价国家是否与当前调研站点一致。"""
    date_text = str(review.get("date_text") or "").strip().lower()
    if not date_text:
        return True
    target_code = get_marketplace(site).code
    target_patterns = REVIEW_LOCATION_PATTERNS.get(target_code, ())
    if any(pattern in date_text for pattern in target_patterns):
        return True
    known_patterns = tuple(
        pattern
        for patterns in REVIEW_LOCATION_PATTERNS.values()
        for pattern in patterns
    )
    # 无法识别地点时保留，避免因 Amazon 文案变化误删合法评论；识别为其他站点时剔除。
    return not any(pattern in date_text for pattern in known_patterns)


# Amazon 合并父体时可能把完全不同商品的评论放进同一个评论池。
# 只拦截带有强烈异品特征的评论，普通且无法确定的评论继续保留。
PRODUCT_MISMATCH_SIGNATURES = (
    "toilet",
    "tiles",
    "grout",
    "metal end",
    "lenski gauge",
    "night light",
    "remote control",
    "remote-",
    "newborn",
    "diaper change",
    "nursery",
    "kids light",
    "my granddaughter",
    "my little daughter",
    "for the baby",
    "baby to sleep",
    "many colours",
    "colorful",
    "lamp and white noise",
)


def review_matches_product_context(review: dict[str, Any], product_context: str) -> bool:
    """过滤标题和正文中出现多个强异品特征的合并评论。"""
    if not product_context.strip():
        return True
    review_text = " ".join(
        str(review.get(key) or "") for key in ("title", "body")
    ).lower()
    context_text = product_context.lower()
    suspicious = [
        phrase
        for phrase in PRODUCT_MISMATCH_SIGNATURES
        if phrase in review_text and phrase not in context_text
    ]
    return not suspicious


def reviewer_filter_label(reviewer_type: str, trusted_only: bool) -> str:
    if reviewer_type == "verified":
        return "Verified Purchase only"
    if reviewer_type == "default":
        return "Top reviews；保留 Verified Purchase 或 Vine 绿标" if trusted_only else "Top reviews"
    if trusted_only:
        return "All reviewers；保留 Verified Purchase 或 Vine 绿标"
    return "All reviewers"


def build_review_url(
    base_url: str,
    asin: str,
    *,
    sort_by: str,
    reviewer_type: str,
    star_filter: str,
    endpoint: str = "product-reviews",
) -> str:
    query_values = {
        "filterByStar": star_filter,
        "sortBy": SORT_OPTIONS[sort_by],
    }
    reviewer_value = REVIEWER_TYPE_OPTIONS[reviewer_type]
    if reviewer_value:
        query_values["reviewerType"] = reviewer_value
    if endpoint == "portal/customer-reviews":
        query = urlencode({"ie": "UTF8", **query_values})
        return f"{base_url}/{endpoint}/{asin}?{query}#reviews-filter-bar"
    query = urlencode(query_values)
    return f"{base_url}/{endpoint}/{asin}/?{query}"


def primary_review_source(
    asin: str,
    parent_asins: dict[str, str] | None,
) -> tuple[str, str, str]:
    del parent_asins
    return asin, "product-reviews", ""


def extract_parent_asin_from_product_html(value: Any, child_asin: str) -> str | None:
    html = str(value or "")
    child_asin = normalize_asin(child_asin) or ""
    candidates: list[str] = []
    for pattern in (
        r'["\']parentAsin["\']\s*:\s*["\']([A-Z0-9]{10})["\']',
        r'parentAsin=([A-Z0-9]{10})(?:[&"\'])',
    ):
        candidates.extend(re.findall(pattern, html, flags=re.I))
    for candidate in candidates:
        parent_asin = normalize_asin(candidate)
        if parent_asin and parent_asin != child_asin:
            return parent_asin
    return None


def resolve_current_parent_asin(page, *, asin: str, site: str) -> str | None:
    marketplace = get_marketplace(site)
    page.goto(
        f"{marketplace.amazon_base_url}/dp/{asin}?th=1",
        wait_until="domcontentloaded",
        timeout=60_000,
    )
    if page_is_login(page):
        raise AmazonAuthenticationRequired(
            "Amazon 在商品页要求重新登录。请在项目专用浏览器完成登录后重试。"
        )
    return extract_parent_asin_from_product_html(page.content(), asin)


def product_title_from_page_title(value: Any) -> str:
    text = str(value or "").strip()
    text = re.sub(r"^Amazon\.[^:]+:\s*(?:Customer reviews|Kundenrezensionen)\s*:\s*", "", text, flags=re.I)
    return text


def page_is_login(page) -> bool:
    return bool(
        page.evaluate(
            """() => location.pathname.includes('/ap/signin')
              || !!document.querySelector('input[type="password"], input[name="password"]')"""
        )
    )


def amazon_login_has_real_challenge(page) -> bool:
    try:
        return bool(
            page.evaluate(
                r"""() => {
                  const isVisible = (el) => {
                    if (!el) return false;
                    const style = getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    return style.display !== 'none'
                      && style.visibility !== 'hidden'
                      && rect.width > 0
                      && rect.height > 0;
                  };
                  const text = document.body?.innerText || '';
                  const hasCredentialInput = [...document.querySelectorAll(
                    'input[type="password"], input[type="email"], #ap_password, #ap_email'
                  )].some(isVisible);
                  return hasCredentialInput
                    || !!document.querySelector('#captchacharacters, form[action*="validateCaptcha" i]')
                    || /captcha|robot check|security check|verify your identity/i.test(text);
                }"""
            )
        )
    except Exception:
        return False


def amazon_account_page_ready(page) -> bool:
    try:
        return bool(
            page.evaluate(
                r"""() => {
                  const path = location.pathname.toLowerCase();
                  const text = (document.body?.innerText || '').replace(/\s+/g, ' ');
                  const accountText = (
                    document.querySelector('#nav-link-accountList-nav-line-1')?.innerText
                    || document.querySelector('#nav-link-accountList')?.innerText
                    || ''
                  ).trim();
                  if (path.includes('/ap/signin')) return false;
                  // “Mein Konto/Your Account”在未登录页面也会出现，必须优先检查顶部问候语。
                  if (/^(?:hello|hallo)[,，]?\s*(?:sign\s*in|anmelden|einloggen|login)\b/i.test(accountText)) {
                    return false;
                  }
                  if (/^您好[，,]?\s*(?:请登录|登录)\b/i.test(accountText)) {
                    return false;
                  }
                  const signedInGreeting = /^(?:hello|hallo)[,，]?\s*[^\n]+/i.test(accountText)
                    || /^您好[，,]?\s*[^\n]+/i.test(accountText);
                  return signedInGreeting
                    || /Your Orders|Meine Bestellungen|Ihre Bestellungen|订单/i.test(text);
                }"""
            )
        )
    except Exception:
        return False


def click_amazon_login_shell(page) -> bool:
    if amazon_login_has_real_challenge(page):
        return False
    try:
        selectors = (
            '#nav-link-accountList',
            'a[data-nav-ref="nav_signin"]',
            'a[href*="/ap/signin"]',
            'input#continue',
            'button#continue',
            'input#signInSubmit',
        )
        for selector in selectors:
            candidate = page.locator(selector).first
            if candidate.count() and candidate.is_visible(timeout=700):
                candidate.click(timeout=10_000)
                page.wait_for_timeout(1_500)
                return True
    except Exception:
        pass
    return False


def recover_amazon_login_shell(page, *, site: str, return_url: str | None = None) -> bool:
    marketplace = get_marketplace(site)
    target_url = return_url or marketplace.amazon_account_url
    probe_urls = (marketplace.amazon_account_url, marketplace.amazon_base_url + "/")

    for probe_url in probe_urls:
        try:
            page.goto(probe_url, wait_until="domcontentloaded", timeout=60_000)
            page.wait_for_timeout(2_000)
            if amazon_account_page_ready(page):
                if target_url != page.url:
                    page.goto(target_url, wait_until="domcontentloaded", timeout=60_000)
                    page.wait_for_timeout(1_000)
                return not page_is_login(page)
            if amazon_login_has_real_challenge(page):
                return False
            if click_amazon_login_shell(page):
                page.wait_for_timeout(2_000)
                if amazon_account_page_ready(page):
                    if target_url != page.url:
                        page.goto(target_url, wait_until="domcontentloaded", timeout=60_000)
                        page.wait_for_timeout(1_000)
                    return not page_is_login(page)
        except Exception:
            continue
    return False


def preflight_amazon(page, *, site: str, postal_code: str | None) -> None:
    marketplace = get_marketplace(site)
    monitor_event("amazon_preflight_started", site=site, postal_code=postal_code)
    page.goto(marketplace.amazon_account_url, wait_until="domcontentloaded", timeout=60_000)
    page.wait_for_timeout(1000)
    if not amazon_account_page_ready(page) and not recover_amazon_login_shell(
        page,
        site=site,
        return_url=marketplace.amazon_account_url,
    ):
        raise AmazonAuthenticationRequired(
            "Amazon 自动恢复登录态失败。仅当页面要求输入账号密码、验证码或安全验证时，"
            "才需要在项目专用浏览器中手动处理后重试。"
        )
    if postal_code:
        location = ensure_amazon_postal_code(page, postal_code)
        if location.get("status") not in {"success", "already-set"}:
            raise RuntimeError(
                f"Amazon 邮编未确认成功，已停止评论抓取：{location.get('error') or location.get('text') or location}"
            )
    monitor_event("amazon_preflight_finished", site=site, postal_code=postal_code)


def collect_reviews_for_asin(
    page,
    *,
    asin: str,
    site: str,
    postal_code: str | None,
    review_source_asin: str | None = None,
    review_endpoint: str = "product-reviews",
    fallback_reason: str = "",
    max_pages: int,
    max_reviews: int,
    sort_by: str,
    reviewer_type: str,
    trusted_only: bool,
    star_filter: str,
    product_context: str = "",
    preflight_complete: bool = False,
) -> AsinResult:
    marketplace = get_marketplace(site)
    query_asin = review_source_asin or asin
    result = AsinResult(
        asin=asin,
        review_source_asin=query_asin,
        star_filter=star_filter,
        reviewer_type=reviewer_type,
        sort_by=sort_by,
        product_url=f"{marketplace.amazon_base_url}/dp/{asin}",
        requested_pages=max_pages,
        fallback_reason=fallback_reason,
    )
    seen: set[str] = set()
    review_url = build_review_url(
        marketplace.amazon_base_url,
        query_asin,
        sort_by=sort_by,
        reviewer_type=reviewer_type,
        star_filter=star_filter,
        endpoint=review_endpoint,
    )
    result.first_review_url = review_url
    started_at = time.monotonic()
    monitor_event(
        "review_job_started",
        asin=asin,
        query_asin=query_asin,
        star_filter=star_filter,
        candidate_limit=max_reviews,
        max_pages=max_pages,
    )
    try:
        if not preflight_complete:
            preflight_amazon(page, site=site, postal_code=postal_code)
        page.goto(review_url, wait_until="domcontentloaded", timeout=60_000)
        if postal_code:
            location_text = amazon_location_text(page)
            if location_text and not amazon_postal_code_matches(location_text, postal_code):
                page.reload(wait_until="domcontentloaded", timeout=60_000)
                location_text = amazon_location_text(page)
                if location_text and not amazon_postal_code_matches(location_text, postal_code):
                    result.status = "failed"
                    result.error = (
                        "Amazon 评论页左上角配送位置与目标邮编不一致："
                        f"{location_text}，目标邮编 {postal_code}"
                    )
                    return result
        page.wait_for_function(
            r"""() => {
              const body = (document.body?.innerText || '').replace(/\s+/g, ' ');
              return document.querySelectorAll('[data-hook="review"]').length > 0
                || document.querySelector('#captchacharacters, form[action*="validateCaptcha" i]')
                || document.querySelector('input[type="password"], input[name="password"]')
                || /page not found/i.test(document.title || '')
                || /no customer reviews|no reviews|noch keine kundenrezensionen|keine rezensionen|0 global ratings|0 globale bewertungen/i.test(body);
            }""",
            timeout=60_000,
        )
        monitor_event(
            "review_page_ready",
            asin=asin,
            star_filter=star_filter,
            duration_seconds=round(time.monotonic() - started_at, 1),
        )
    except Exception as exc:
        result.status = "failed"
        result.error = f"评论页打开或等待首批评论失败：{exc}"
        monitor_event(
            "review_job_failed",
            asin=asin,
            star_filter=star_filter,
            phase="open_page",
            duration_seconds=round(time.monotonic() - started_at, 1),
            error=str(exc),
        )
        return result

    for batch_number in range(1, max_pages + 1):
        try:
            state = page.evaluate(EXTRACT_REVIEWS_JS)
        except Exception as exc:
            result.status = "partial" if result.reviews else "failed"
            result.error = f"第 {batch_number} 批评论解析失败：{exc}"
            break

        if state.get("login"):
            raise AmazonAuthenticationRequired(
                "Amazon 在评论页要求重新登录。请在项目专用浏览器完成登录后重试。"
            )
        if state.get("captcha"):
            result.status = "partial" if result.reviews else "blocked"
            result.error = f"加载第 {batch_number} 批评论时出现 Amazon 验证码或机器人验证。"
            break
        if state.get("pageNotFound"):
            result.status = "not_found"
            result.error = f"评论来源 ASIN {query_asin} 的评论页不存在。"
            break

        filters = state.get("filters") or {}
        expected_filters = {
            "sortBy": SORT_OPTIONS[sort_by],
            "starFilter": star_filter,
        }
        reviewer_value = REVIEWER_TYPE_OPTIONS[reviewer_type]
        if reviewer_value:
            expected_filters["reviewerType"] = reviewer_value
        mismatches = {
            key: (filters.get(key), expected)
            for key, expected in expected_filters.items()
            if filters.get(key) and filters.get(key) != expected
        }
        if mismatches:
            result.status = "failed"
            result.error = f"Amazon 页面筛选条件未正确应用：{mismatches}"
            break

        result.scanned_pages = batch_number
        result.product_title = result.product_title or str(state.get("productTitle") or "").strip()
        result.product_title = result.product_title or product_title_from_page_title(state.get("pageTitle"))
        new_seen_count = 0
        candidate_limit_reached = False
        for raw in state.get("reviews") or []:
            review = normalize_review(
                raw,
                asin=asin,
                review_source_asin=query_asin,
                page_number=batch_number,
                source_url=str(state.get("url") or review_url),
            )
            if not review_matches_star_filter(review, star_filter):
                continue
            if not review_matches_marketplace(review, site):
                monitor_event(
                    "review_filtered_marketplace_mismatch",
                    asin=asin,
                    star_filter=star_filter,
                    review_id=review.get("review_id", ""),
                    date_text=review.get("date_text", ""),
                )
                continue
            if not review_matches_product_context(review, product_context):
                monitor_event(
                    "review_filtered_product_mismatch",
                    asin=asin,
                    star_filter=star_filter,
                    review_id=review.get("review_id", ""),
                )
                continue
            identity = review["review_id"] or "|".join(
                str(review.get(key) or "") for key in ("reviewer_name", "date_text", "title", "body")
            )
            if not identity or identity in seen:
                continue
            seen.add(identity)
            new_seen_count += 1
            if trusted_only and not review_is_trusted(review):
                if len(seen) >= max_reviews:
                    candidate_limit_reached = True
                    break
                continue
            result.reviews.append(review)
            if len(seen) >= max_reviews:
                candidate_limit_reached = True
                break

        print(
            f"[review_research] {asin} [{star_filter}] batch={batch_number} "
            f"candidates={len(seen)}/{max_reviews}, trusted={len(result.reviews)}",
            flush=True,
        )
        monitor_event(
            "review_batch_scanned",
            asin=asin,
            star_filter=star_filter,
            batch=batch_number,
            candidates=len(seen),
            trusted=len(result.reviews),
            new_candidates=new_seen_count,
        )

        if candidate_limit_reached or len(seen) >= max_reviews:
            break
        if state.get("noReviews") and not result.reviews:
            result.status = "no_reviews"
            break
        if not state.get("reviews"):
            result.status = "partial" if result.reviews else "no_reviews_found"
            result.error = f"第 {batch_number} 批没有识别到评论卡片。"
            break
        if new_seen_count == 0 and batch_number > 1:
            result.status = "insufficient_reviews" if result.reviews else "no_reviews_found"
            result.error = f"第 {batch_number} 批没有新增评论，已停止继续加载。"
            break
        if batch_number >= max_pages:
            if len(seen) < max_reviews:
                result.status = "batch_limit_reached"
                result.error = (
                    f"达到 {max_pages} 个加载批次上限，页面扫描到 {len(seen)} 条候选评论，"
                    f"可信筛选后保留 {len(result.reviews)} 条，尚未达到候选上限 {max_reviews} 条。"
                )
            break
        show_more = page.locator('[data-hook="show-more-button"]:visible')
        if not show_more.count():
            if len(seen) < max_reviews:
                result.status = "insufficient_reviews"
                result.error = (
                    f"页面扫描到 {len(seen)} 条候选评论，可信筛选后保留 "
                    f"{len(result.reviews)} 条，Show 10 more reviews 按钮已消失，"
                    f"少于候选上限 {max_reviews} 条。"
                )
            break
        previous_count = len(state.get("reviews") or [])
        try:
            monitor_event(
                "review_more_click_started",
                asin=asin,
                star_filter=star_filter,
                next_batch=batch_number + 1,
                current_dom_reviews=previous_count,
            )
            show_more.last.scroll_into_view_if_needed(timeout=5_000)
            show_more.last.click(timeout=10_000)
            page.wait_for_function(
                "count => document.querySelectorAll('[data-hook=review]').length > count",
                arg=previous_count,
                timeout=30_000,
            )
            monitor_event(
                "review_more_click_finished",
                asin=asin,
                star_filter=star_filter,
                next_batch=batch_number + 1,
            )
        except Exception as exc:
            result.status = "partial" if result.reviews else "failed"
            result.error = f"点击 Show 10 more reviews 加载第 {batch_number + 1} 批失败：{exc}"
            monitor_event(
                "review_job_failed",
                asin=asin,
                star_filter=star_filter,
                phase="show_more",
                batch=batch_number + 1,
                duration_seconds=round(time.monotonic() - started_at, 1),
                error=str(exc),
            )
            break
    monitor_event(
        "review_job_finished",
        asin=asin,
        query_asin=query_asin,
        star_filter=star_filter,
        status=result.status,
        candidates=len(seen),
        trusted=len(result.reviews),
        batches=result.scanned_pages,
        duration_seconds=round(time.monotonic() - started_at, 1),
    )
    return result


def collect_reviews(
    *,
    asins: Iterable[str],
    parent_asins: dict[str, str] | None,
    site: str,
    postal_code: str | None,
    debugger_address: str,
    max_pages: int,
    max_reviews: int | None,
    max_positive_reviews: int,
    max_critical_reviews: int,
    sort_by: str,
    reviewer_type: str,
    trusted_only: bool,
    star_filters: Iterable[str],
    product_contexts: dict[str, str] | None = None,
) -> list[AsinResult]:
    sync_playwright = require_sync_playwright()
    with sync_playwright() as playwright:
        browser = playwright.chromium.connect_over_cdp(f"http://{debugger_address}")
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        preflight_page = context.new_page()
        try:
            preflight_amazon(preflight_page, site=site, postal_code=postal_code)
            preflight_page.close()
            jobs = [(asin, star_filter) for asin in asins for star_filter in star_filters]
            results = []
            parent_asins = parent_asins or {}
            product_contexts = product_contexts or {}
            for index, (asin, star_filter) in enumerate(jobs, start=1):
                review_limit = review_limit_for_filter(
                    star_filter,
                    max_reviews=max_reviews,
                    max_positive_reviews=max_positive_reviews,
                    max_critical_reviews=max_critical_reviews,
                )
                print(
                    f"[review_research] job {index}/{len(jobs)}: "
                    f"{asin} [{star_filter}], candidate_limit={review_limit}",
                    flush=True,
                )
                monitor_event(
                    "review_job_dispatch",
                    job_index=index,
                    job_count=len(jobs),
                    asin=asin,
                    star_filter=star_filter,
                    candidate_limit=review_limit,
                )
                source_asin, review_endpoint, source_reason = primary_review_source(
                    asin,
                    parent_asins,
                )
                # 每个评论任务使用独立页面，避免上一个 ASIN 的导航状态影响下一个任务。
                page = context.new_page()
                try:
                    current = collect_reviews_for_asin(
                        page,
                        asin=asin,
                        review_source_asin=source_asin,
                        review_endpoint=review_endpoint,
                        fallback_reason=source_reason,
                        site=site,
                        postal_code=postal_code,
                        max_pages=max_pages,
                        max_reviews=review_limit,
                        sort_by=sort_by,
                        reviewer_type=reviewer_type,
                        trusted_only=trusted_only,
                        star_filter=star_filter,
                        product_context=product_contexts.get(asin, ""),
                        preflight_complete=True,
                    )
                finally:
                    page.close()
                results.append(current)
                if current.status == "blocked":
                    for pending_asin, pending_filter in jobs[index:]:
                        results.append(
                            AsinResult(
                                asin=pending_asin,
                                review_source_asin=pending_asin,
                                star_filter=pending_filter,
                                reviewer_type=reviewer_type,
                                sort_by=sort_by,
                                product_url=f"{get_marketplace(site).amazon_base_url}/dp/{pending_asin}",
                                requested_pages=max_pages,
                                status="not_attempted",
                                error="前一个 ASIN 出现 Amazon 验证码，已停止后续抓取。",
                            )
                        )
                    break
        finally:
            if not preflight_page.is_closed():
                preflight_page.close()
            # 退出 sync_playwright 上下文即可断开 CDP 客户端，不能关闭用户的项目浏览器。
    return results


def review_limit_for_filter(
    star_filter: str,
    *,
    max_reviews: int | None = None,
    max_positive_reviews: int = 50,
    max_critical_reviews: int = 100,
) -> int:
    if max_reviews is not None:
        return max_reviews
    if star_filter in {"positive", "five_star", "four_star"}:
        return max_positive_reviews
    return max_critical_reviews


def style_header(worksheet, row: int, columns: int) -> None:
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for cell in worksheet[row][:columns]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def set_hyperlink(cell, url: str) -> None:
    if not url:
        return
    cell.hyperlink = url
    cell.font = Font(color="0563C1", underline="single")


def prepare_sheet(workbook, name: str):
    if name in workbook.sheetnames:
        del workbook[name]
    return workbook.create_sheet(name)


def load_review_translations(path: Path | None) -> dict[str, str]:
    if path is None or not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if isinstance(payload, dict) and isinstance(payload.get("translations"), dict):
        payload = payload["translations"]
    if not isinstance(payload, dict):
        raise ValueError("评论翻译缓存必须是 JSON 对象，或包含 translations 对象。")
    return {
        str(source).strip(): str(translation).strip()
        for source, translation in payload.items()
        if str(source).strip() and str(translation).strip()
    }


def write_missing_review_translations(path: Path, missing: Iterable[str]) -> None:
    unique = list(dict.fromkeys(text.strip() for text in missing if text.strip()))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {
                "target_language": "zh-CN",
                "review_scope": "Critical reviews only",
                "positive_review_translation": "skipped",
                "missing_count": len(unique),
                "translations": {text: "" for text in unique},
            },
            ensure_ascii=False,
            indent=2,
        ) + "\n",
        encoding="utf-8",
    )


def load_pain_phrase_groups(path: Path | None) -> dict[str, SemanticGroup]:
    if path is None or not path.exists():
        return {}
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    if isinstance(payload, dict) and isinstance(payload.get("groups"), dict):
        payload = payload["groups"]
    if not isinstance(payload, dict):
        raise ValueError("痛点分组缓存必须是 JSON 对象，或包含 groups 对象。")
    groups: dict[str, SemanticGroup] = {}
    for phrase, group in payload.items():
        if not isinstance(group, dict):
            continue
        english = str(group.get("group_en") or "").strip()
        chinese = str(group.get("group_zh") or "").strip()
        if str(phrase).strip() and english and chinese:
            groups[str(phrase).strip()] = SemanticGroup(english, chinese)
    return groups


def write_missing_pain_phrase_groups(path: Path, missing: Iterable[str]) -> None:
    unique = list(dict.fromkeys(text.strip() for text in missing if text.strip()))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(
            {
                "missing_count": len(unique),
                "groups": {
                    phrase: {"group_en": "", "group_zh": ""}
                    for phrase in unique
                },
            },
            ensure_ascii=False,
            indent=2,
        ) + "\n",
        encoding="utf-8",
    )


def estimated_text_row_height(value: str, column_width: int = 72) -> float:
    visual_lines = 0
    for line in value.splitlines() or [""]:
        units = sum(2 if unicodedata.east_asian_width(character) in {"W", "F"} else 1 for character in line)
        visual_lines += max(1, math.ceil(units / column_width))
    return max(54.0, min(409.0, visual_lines * 15.0 + 8.0))


def original_review_image_url(url: str) -> str:
    text = str(url or "").strip()
    if not text:
        return ""
    base, separator, query = text.partition("?")
    base = re.sub(r"\._[^/]+_(?=\.[A-Za-z0-9]{2,5}$)", "", base)
    return base + (separator + query if separator else "")


def download_review_image(url: str, output: Path) -> None:
    request = urllib.request.Request(
        original_review_image_url(url),
        headers={"User-Agent": "Mozilla/5.0"},
    )
    with urllib.request.urlopen(request, timeout=15) as response:
        output.write_bytes(response.read())


def create_review_image_contact_sheet(urls: list[str], output: Path) -> int:
    from PIL import Image as PillowImage
    from PIL import ImageOps

    images = []
    for index, url in enumerate(urls, start=1):
        downloaded = output.with_name(f"{output.stem}_{index}.img")
        try:
            download_review_image(url, downloaded)
            with PillowImage.open(downloaded) as image:
                images.append(ImageOps.exif_transpose(image).convert("RGB"))
        except Exception:
            continue
    if not images:
        return 0

    tile_width, tile_height = 120, 100
    columns = 1 if len(images) == 1 else 2 if len(images) <= 4 else 3
    rows = (len(images) + columns - 1) // columns
    canvas = PillowImage.new("RGB", (tile_width * columns, tile_height * rows), "white")
    for index, image in enumerate(images):
        thumbnail = ImageOps.contain(image, (tile_width - 6, tile_height - 6))
        x = (index % columns) * tile_width + (tile_width - thumbnail.width) // 2
        y = (index // columns) * tile_height + (tile_height - thumbnail.height) // 2
        canvas.paste(thumbnail, (x, y))
    canvas.save(output, format="PNG")
    return len(images)


def add_pain_point_sheet(
    workbook,
    results: list[AsinResult],
    image_dir: Path,
    translations: dict[str, str] | None = None,
    custom_groups: dict[str, SemanticGroup] | None = None,
    trusted_only: bool = True,
) -> tuple[list[ExcelImage], list[str], list[str]]:
    worksheet = prepare_sheet(workbook, PAIN_POINT_SHEET)
    headers = [
        "竞品ASIN", "实际筛选", "星级", "评论日期", "属性/型号",
        "差评词/短语", "归一化痛点（中文）", "Helpful数", "图片数量",
        "评论图片", "评价内容", "评论链接",
    ]
    worksheet.append(headers)
    style_header(worksheet, 1, len(headers))
    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False
    embedded_images: list[ExcelImage] = []
    missing_translations: list[str] = []
    translations = translations or {}
    normalized_custom_groups = normalize_semantic_groups(custom_groups)

    critical_reviews = [
        (result, review)
        for result in results
        if result.star_filter == "critical"
        for review in result.reviews
    ]
    critical_review_data = [review for _, review in critical_reviews]
    terms_by_review, _ = analyze_pain_phrases(critical_review_data)
    frequency_rows: list[tuple[str, str, str, Any]] = []
    missing_groups: list[str] = []
    grouped_by_source: dict[
        tuple[str, str, str], tuple[list[dict[str, Any]], list[list[str]]]
    ] = {}
    for (result, review), terms in zip(critical_reviews, terms_by_review):
        key = (result.asin, result.review_source_asin or result.asin, result.reviewer_type)
        reviews, grouped_terms = grouped_by_source.setdefault(key, ([], []))
        reviews.append(review)
        grouped_terms.append(terms)
    for (asin, source_asin, actual_filter), (reviews, grouped_terms) in grouped_by_source.items():
        grouped_rows, current_missing = aggregate_semantic_groups(
            reviews,
            grouped_terms,
            custom_groups=custom_groups,
            phrase_translations=translations,
        )
        frequency_rows.extend(
            (asin, source_asin, actual_filter, item) for item in grouped_rows
        )
        missing_groups.extend(current_missing)
    missing_groups = list(dict.fromkeys(missing_groups))

    image_dir.mkdir(parents=True, exist_ok=True)
    for (result, review), pain_terms in zip(critical_reviews, terms_by_review):
        content = "\n".join(part for part in (review.get("title"), review.get("body")) if part)
        normalized_pain_terms = [
            group.chinese
            if (
                group := semantic_group_for_phrase(
                    phrase,
                    normalized_custom_groups,
                    review_text=content,
                )
            )
            else "待归类"
            for phrase in pain_terms
        ]
        translation = translations.get(content, "")
        display_content = f"{content}\n({translation})" if translation else content
        if content and not translation:
            missing_translations.append(content)
        image_urls = list(review.get("image_urls") or [])
        content_height = estimated_text_row_height(display_content)
        worksheet.append(
            [
                result.asin,
                reviewer_filter_label(result.reviewer_type, trusted_only),
                review.get("rating"),
                review.get("date"),
                review.get("variant") or "/",
                "\n".join(pain_terms) or "/",
                "\n".join(normalized_pain_terms) or "/",
                review.get("helpful_count"),
                len(image_urls),
                None,
                display_content,
                review.get("review_url"),
            ]
        )
        row = worksheet.max_row
        set_hyperlink(worksheet.cell(row, 12), str(review.get("review_url") or ""))
        if image_urls:
            contact_sheet = image_dir / f"{result.asin}_{review.get('review_id') or row}.png"
            embedded_count = create_review_image_contact_sheet(image_urls, contact_sheet)
            worksheet.cell(row, 9).value = embedded_count
            if embedded_count:
                excel_image = ExcelImage(str(contact_sheet))
                scale = min(240 / excel_image.width, 300 / excel_image.height, 1)
                excel_image.width *= scale
                excel_image.height *= scale
                worksheet.add_image(excel_image, f"J{row}")
                embedded_images.append(excel_image)
                worksheet.row_dimensions[row].height = max(90, excel_image.height * 0.75, content_height)
            else:
                worksheet.cell(row, 10).value = None
        else:
            worksheet.row_dimensions[row].height = content_height

    worksheet.auto_filter.ref = worksheet.dimensions
    widths = {
        "A": 14, "B": 22, "C": 8, "D": 13, "E": 28, "F": 28,
        "G": 24, "H": 12, "I": 10, "J": 35, "K": 72, "L": 42,
    }
    for letter, width in widths.items():
        worksheet.column_dimensions[letter].width = width
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )
    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = thin_border
            if cell.row >= 2:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    for cell in worksheet["D"][1:]:
        cell.number_format = "yyyy-mm-dd"

    if PAIN_FREQUENCY_SHEET in workbook.sheetnames:
        del workbook[PAIN_FREQUENCY_SHEET]
    frequency_sheet = workbook[SUMMARY_SHEET]
    frequency_headers = [
        "竞品ASIN", "实际筛选", "关键词/短语",
        "中文翻译", "评论覆盖数", "覆盖率", "涉及型号",
    ]
    section_title_row = frequency_sheet.max_row + 2
    for column in range(1, len(frequency_headers) + 1):
        frequency_sheet.cell(
            section_title_row,
            column,
            PAIN_FREQUENCY_SECTION if column == 1 else None,
        )
    style_header(frequency_sheet, section_title_row, len(frequency_headers))
    frequency_header_row = section_title_row + 1
    for column, value in enumerate(frequency_headers, start=1):
        frequency_sheet.cell(frequency_header_row, column, value)
    style_header(frequency_sheet, frequency_header_row, len(frequency_headers))
    for asin, source_asin, actual_filter, item in frequency_rows:
        frequency_sheet.append(
            [
                asin,
                reviewer_filter_label(actual_filter, trusted_only),
                item.phrase,
                item.translation,
                item.review_count,
                item.review_ratio,
                "\n".join(item.variants) or "/",
            ]
        )
    for letter, width in {
        "A": 14, "B": 22, "C": 34,
        "D": 28, "E": 14, "F": 14, "G": 32,
    }.items():
        frequency_sheet.column_dimensions[letter].width = max(
            frequency_sheet.column_dimensions[letter].width or 0,
            width,
        )
    for row in frequency_sheet.iter_rows(
        min_row=section_title_row,
        max_row=frequency_sheet.max_row,
        min_col=1,
        max_col=len(frequency_headers),
    ):
        for cell in row:
            cell.border = thin_border
            if cell.row > frequency_header_row:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
    for row in range(frequency_header_row + 1, frequency_sheet.max_row + 1):
        frequency_sheet.cell(row, 6).number_format = "0.0%"
    return embedded_images, missing_translations, missing_groups


def write_review_sheets(
    source: Path,
    output: Path,
    results: list[AsinResult],
    *,
    site: str,
    translation_cache: Path | None = DEFAULT_TRANSLATION_CACHE,
    translation_report: Path | None = None,
    pain_group_cache: Path | None = DEFAULT_PAIN_GROUP_CACHE,
    pain_group_report: Path | None = None,
    trusted_only: bool = True,
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    if source.resolve() != output.resolve():
        shutil.copy2(source, output)
    workbook = load_workbook(output, rich_text=True)
    recovered_source_images = restore_unloaded_floating_images(workbook, output)
    try:
        summary = prepare_sheet(workbook, SUMMARY_SHEET)
        details = prepare_sheet(workbook, DETAIL_SHEET)

        summary_headers = [
            "抓取时间", "站点", "竞品ASIN", "评论类型", "实际筛选",
            "抓取状态", "抓取评论数",
            "Verified Purchase数量", "绿标评论数量", "含图片评论数", "含视频评论数",
            "商品链接", "错误/说明",
        ]
        summary.append(summary_headers)
        style_header(summary, 1, len(summary_headers))
        now = datetime.now()
        for result in results:
            summary.append(
                [
                    now,
                    site,
                    result.asin,
                    "Positive reviews" if result.star_filter == "positive" else "Critical reviews" if result.star_filter == "critical" else result.star_filter,
                    reviewer_filter_label(result.reviewer_type, trusted_only),
                    result.status,
                    len(result.reviews),
                    sum(1 for review in result.reviews if review.get("verified")),
                    sum(1 for review in result.reviews if review.get("vine")),
                    sum(1 for review in result.reviews if review.get("image_urls")),
                    sum(1 for review in result.reviews if review.get("has_video")),
                    result.product_url,
                    result.error,
                ]
            )
            row = summary.max_row
            set_hyperlink(summary.cell(row, 12), result.product_url)

        detail_headers = [
            "站点", "竞品ASIN", "实际筛选", "评论日期", "星级",
            "评论类型", "评论标题", "评论正文", "购买变体",
            "Verified Purchase", "绿标评论", "Helpful数", "图片URL", "含视频", "评论链接",
        ]
        details.append(detail_headers)
        style_header(details, 1, len(detail_headers))
        for result in results:
            for review in result.reviews:
                details.append(
                    [
                        site,
                        result.asin,
                        reviewer_filter_label(result.reviewer_type, trusted_only),
                        review["date"],
                        review["rating"],
                        "Positive reviews" if result.star_filter == "positive" else "Critical reviews" if result.star_filter == "critical" else result.star_filter,
                        review["title"],
                        review["body"],
                        review["variant"],
                        "是" if review["verified"] else "否",
                        "是" if review["vine"] else "否",
                        review["helpful_count"],
                        "\n".join(review["image_urls"]),
                        "是" if review["has_video"] else "否",
                        review["review_url"],
                    ]
                )
                row = details.max_row
                set_hyperlink(details.cell(row, 15), review["review_url"])

        summary.freeze_panes = "A2"
        summary.auto_filter.ref = summary.dimensions
        details.freeze_panes = "A2"
        details.auto_filter.ref = details.dimensions
        summary.sheet_view.showGridLines = False
        details.sheet_view.showGridLines = False
        summary.column_dimensions["A"].width = 20
        summary.column_dimensions["C"].width = 14
        summary.column_dimensions["E"].width = 46 if trusted_only else 22
        summary.column_dimensions["F"].width = 22
        summary.column_dimensions["L"].width = 34
        summary.column_dimensions["M"].width = 48
        for column in range(1, len(summary_headers) + 1):
            letter = get_column_letter(column)
            if summary.column_dimensions[letter].width == 13:
                summary.column_dimensions[letter].width = 14
        detail_widths = {
            "A": 9, "B": 14, "C": 46 if trusted_only else 22, "D": 13, "E": 8,
            "F": 18, "G": 36, "H": 72, "I": 30, "J": 18, "K": 12,
            "L": 13, "M": 48, "N": 10, "O": 42,
        }
        for letter, width in detail_widths.items():
            details.column_dimensions[letter].width = width
        for worksheet in (summary, details):
            for row in worksheet.iter_rows(min_row=2):
                for cell in row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
        for cell in summary["A"][1:]:
            cell.number_format = "yyyy-mm-dd hh:mm"
        for cell in details["D"][1:]:
            cell.number_format = "yyyy-mm-dd"
        translations = load_review_translations(translation_cache)
        custom_groups = load_pain_phrase_groups(pain_group_cache)
        with tempfile.TemporaryDirectory(prefix="review-images-") as temp_dir:
            embedded_images, missing_translations, missing_groups = add_pain_point_sheet(
                workbook,
                results,
                Path(temp_dir),
                translations,
                custom_groups,
                trusted_only,
            )
            try:
                workbook.save(output)
                if translation_report is not None:
                    write_missing_review_translations(translation_report, missing_translations)
                if pain_group_report is not None:
                    write_missing_pain_phrase_groups(pain_group_report, missing_groups)
            finally:
                for image in [*embedded_images, *recovered_source_images]:
                    reference = getattr(image, "ref", None)
                    if hasattr(reference, "close"):
                        reference.close()
    finally:
        workbook.close()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Collect Amazon product reviews into an existing research workbook.")
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--site", default="US")
    parser.add_argument("--amazon-postal-code")
    parser.add_argument("--debugger-address", default=get_debugger_address())
    parser.add_argument("--asin", action="append", default=[], help="Explicit ASIN; repeat to collect more than one")
    parser.add_argument("--asin-limit", type=int)
    parser.add_argument("--max-pages", type=int, default=10)
    parser.add_argument(
        "--max-reviews",
        type=int,
        help="Compatibility override that applies one limit to every review type",
    )
    parser.add_argument("--max-positive-reviews", type=int, default=50)
    parser.add_argument("--max-critical-reviews", type=int, default=100)
    parser.add_argument("--sort-by", choices=tuple(SORT_OPTIONS), default="helpful")
    parser.add_argument("--reviewer-type", choices=tuple(REVIEWER_TYPE_OPTIONS), default="default")
    parser.add_argument(
        "--trusted-only",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Only keep reviews marked Verified Purchase or Vine after loading Top reviews",
    )
    parser.add_argument("--star-filter", choices=("both", *tuple(sorted(STAR_OPTIONS))), default="both")
    parser.add_argument("--translation-cache", type=Path, default=DEFAULT_TRANSLATION_CACHE)
    parser.add_argument("--translation-report", type=Path)
    parser.add_argument("--pain-group-cache", type=Path, default=DEFAULT_PAIN_GROUP_CACHE)
    parser.add_argument("--pain-group-report", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    review_limits = [args.max_positive_reviews, args.max_critical_reviews]
    if args.max_reviews is not None:
        review_limits.append(args.max_reviews)
    if args.max_pages < 1 or any(limit < 1 for limit in review_limits):
        raise ValueError("--max-pages 和评论数量上限必须大于 0。")
    explicit = [normalize_asin(value) for value in args.asin]
    asins = [asin for asin in explicit if asin] if args.asin else extract_review_asins_from_workbook(args.source)
    asins = list(dict.fromkeys(asins))
    if args.asin_limit is not None:
        if args.asin_limit < 1:
            raise ValueError("--asin-limit 必须大于 0。")
        asins = asins[: args.asin_limit]
    if not asins:
        raise ValueError("没有找到可抓取的 ASIN。")
    marketplace = get_marketplace(args.site)
    postal_code = args.amazon_postal_code or marketplace.postal_code
    star_filters = ("positive", "critical") if args.star_filter == "both" else (args.star_filter,)
    print(
        f"[review_research] site={marketplace.code}, asins={len(asins)}, filters={','.join(star_filters)}, "
        f"reviewer={args.reviewer_type}, trusted_only={args.trusted_only}, "
        f"sort={args.sort_by}, max_pages={args.max_pages}, "
        f"positive_candidate_limit={args.max_reviews if args.max_reviews is not None else args.max_positive_reviews}, "
        f"critical_candidate_limit={args.max_reviews if args.max_reviews is not None else args.max_critical_reviews}"
    )
    results = collect_reviews(
        asins=asins,
        parent_asins=extract_parent_asins_from_workbook(args.source),
        site=marketplace.code,
        postal_code=postal_code,
        debugger_address=args.debugger_address,
        max_pages=args.max_pages,
        max_reviews=args.max_reviews,
        max_positive_reviews=args.max_positive_reviews,
        max_critical_reviews=args.max_critical_reviews,
        sort_by=args.sort_by,
        reviewer_type=args.reviewer_type,
        trusted_only=args.trusted_only,
        star_filters=star_filters,
        product_contexts=extract_product_context_from_workbook(args.source),
    )
    translation_report = args.translation_report or args.output.with_name(
        f"{args.output.stem}_missing_review_translations.json"
    )
    pain_group_report = args.pain_group_report or args.output.with_name(
        f"{args.output.stem}_missing_pain_phrase_groups.json"
    )
    write_review_sheets(
        args.source,
        args.output,
        results,
        site=marketplace.code,
        translation_cache=args.translation_cache,
        translation_report=translation_report,
        pain_group_cache=args.pain_group_cache,
        pain_group_report=pain_group_report,
        trusted_only=args.trusted_only,
    )
    print(args.output)
    print(f"[review_research] missing_translations={translation_report}")
    print(f"[review_research] missing_pain_phrase_groups={pain_group_report}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
