from __future__ import annotations

import re
import time
import urllib.parse
from dataclasses import dataclass
from typing import Any


@dataclass(frozen=True)
class Marketplace:
    code: str
    seller_sprite_labels: tuple[str, ...]
    amazon_domain: str
    postal_code: str
    currency_symbol: str

    @property
    def amazon_base_url(self) -> str:
        return f"https://www.{self.amazon_domain}"

    @property
    def amazon_account_url(self) -> str:
        return f"{self.amazon_base_url}/gp/css/homepage.html"

    def amazon_search_url(self, keyword: str) -> str:
        return f"{self.amazon_base_url}/s?k={urllib.parse.quote(keyword)}"


MARKETPLACES: dict[str, Marketplace] = {
    "US": Marketplace(
        code="US",
        seller_sprite_labels=("\u7f8e\u56fd\u7ad9", "\u7f8e\u56fd"),
        amazon_domain="amazon.com",
        postal_code="10001",
        currency_symbol="$",
    ),
    "DE": Marketplace(
        code="DE",
        seller_sprite_labels=("\u5fb7\u56fd\u7ad9", "\u5fb7\u56fd"),
        amazon_domain="amazon.de",
        postal_code="10115",
        currency_symbol="€",
    ),
    "UK": Marketplace(
        code="UK",
        seller_sprite_labels=("\u82f1\u56fd\u7ad9", "\u82f1\u56fd"),
        amazon_domain="amazon.co.uk",
        postal_code="B100AB",
        currency_symbol="£",
    ),
    "CA": Marketplace(
        code="CA",
        seller_sprite_labels=("\u52a0\u62ff\u5927\u7ad9", "\u52a0\u62ff\u5927"),
        amazon_domain="amazon.ca",
        postal_code="M4Y1M7",
        currency_symbol="CA$",
    ),
    "FR": Marketplace(
        code="FR",
        seller_sprite_labels=("\u6cd5\u56fd\u7ad9", "\u6cd5\u56fd"),
        amazon_domain="amazon.fr",
        postal_code="75000",
        currency_symbol="€",
    ),
}

SITE_ALIASES = {
    "USA": "US",
    "AMERICA": "US",
    "\u7f8e\u56fd": "US",
    "\u7f8e\u56fd\u7ad9": "US",
    "GERMANY": "DE",
    "\u5fb7\u56fd": "DE",
    "\u5fb7\u56fd\u7ad9": "DE",
    "GB": "UK",
    "GREAT BRITAIN": "UK",
    "UNITED KINGDOM": "UK",
    "BRITAIN": "UK",
    "\u82f1\u56fd": "UK",
    "\u82f1\u56fd\u7ad9": "UK",
    "CANADA": "CA",
    "CA": "CA",
    "\u52a0\u62ff\u5927": "CA",
    "\u52a0\u62ff\u5927\u7ad9": "CA",
    "FRANCE": "FR",
    "\u6cd5\u56fd": "FR",
    "\u6cd5\u56fd\u7ad9": "FR",
}


def normalize_site_code(value: str | None) -> str:
    raw = str(value or "US").strip()
    code = SITE_ALIASES.get(raw.upper(), SITE_ALIASES.get(raw, raw.upper()))
    if code not in MARKETPLACES:
        supported = ", ".join(sorted(MARKETPLACES))
        raise ValueError(f"Unsupported marketplace site '{value}'. Supported sites: {supported}")
    return code


def get_marketplace(value: str | None) -> Marketplace:
    return MARKETPLACES[normalize_site_code(value)]


def amazon_domains() -> tuple[str, ...]:
    return tuple(sorted({marketplace.amazon_domain for marketplace in MARKETPLACES.values()}))


def amazon_host_matches(host: str | None, marketplace: Marketplace) -> bool:
    normalized = (host or "").lower()
    return normalized in {marketplace.amazon_domain, f"www.{marketplace.amazon_domain}"}


def url_belongs_to_amazon(url: str, domains: tuple[str, ...] | None = None) -> bool:
    host = (urllib.parse.urlparse(str(url or "")).hostname or "").lower()
    return any(host in {domain, f"www.{domain}"} for domain in (domains or amazon_domains()))


def normalize_amazon_url(url: str, site: str | None) -> str:
    marketplace = get_marketplace(site)
    text = str(url or "").strip()
    if not text:
        return text
    if text.startswith("/"):
        return f"{marketplace.amazon_base_url}{text}"
    parsed = urllib.parse.urlparse(text)
    if parsed.scheme not in {"http", "https"}:
        return text
    host = (parsed.hostname or "").lower()
    if not amazon_host_matches(host, marketplace):
        raise ValueError(
            f"Amazon URL marketplace mismatch: target site is {marketplace.code} ({marketplace.amazon_domain}), "
            f"but workbook URL uses {host or text}."
        )
    return text


def ensure_sellersprite_site(page: Any, site: str | None) -> dict[str, str]:
    marketplace = get_marketplace(site)
    target_labels = marketplace.seller_sprite_labels
    all_labels = tuple(label for item in MARKETPLACES.values() for label in item.seller_sprite_labels)

    # 切换站点或刚打开工具页时，页面可能仍在进行一次前端导航；先等 DOM 稳定再读取站点控件。
    try:
        page.wait_for_load_state("domcontentloaded", timeout=15_000)
    except Exception:
        pass

    state = None
    for attempt in range(4):
        try:
            state = page.evaluate(
                r"""({targetLabels, allLabels}) => {
          const normalize = (text) => String(text || '').replace(/\s+/g, ' ').trim();
          const isVisible = (el) => {
            if (!el) return false;
            const style = getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            return style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
          };
          const keywordInput = [...document.querySelectorAll('input, textarea')].find(el =>
            isVisible(el) && /ASIN|关键词|keyword/i.test(el.getAttribute('placeholder') || '')
          );
          const inputRect = keywordInput?.getBoundingClientRect();
          const marketSelect = document.querySelector('.market-select');
          const marketInput = marketSelect?.querySelector('input.el-input__inner');
          if (marketSelect && isVisible(marketSelect) && marketInput && isVisible(marketInput)) {
            const currentText = normalize(marketInput.value || marketInput.getAttribute('value'));
            const currentIcon = marketSelect.querySelector('use')?.getAttribute('xlink:href') || '';
            const iconSite = currentIcon.replace(/^#icon-/, '').toUpperCase();
            if (targetLabels.some(label => currentText === label || currentText.startsWith(label + ' ')) || iconSite === targetLabels[0]) {
              return { status: 'already-set', current: currentText || iconSite };
            }
            const trigger = marketInput.closest('.el-select, .el-input, .market-select') || marketInput;
            trigger.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
            trigger.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
            trigger.click();
            return { status: 'opened', current: currentText || iconSite };
          }
          // 兼容卖家精灵新版：站点选择器不再带 .market-select 外层，直接使用可见的 el-input。
          const standaloneMarketInput = [...document.querySelectorAll('input.el-input__inner')]
            .find(el => {
              if (!isVisible(el)) return false;
              const currentText = normalize(el.value || el.getAttribute('value'));
              return allLabels.some(label => currentText === label || currentText.startsWith(label + ' '));
            });
          if (standaloneMarketInput) {
            const currentText = normalize(standaloneMarketInput.value || standaloneMarketInput.getAttribute('value'));
            if (targetLabels.some(label => currentText === label || currentText.startsWith(label + ' '))) {
              return { status: 'already-set', current: currentText };
            }
            const trigger = standaloneMarketInput.closest('.el-select, [role="combobox"], .el-input') || standaloneMarketInput;
            trigger.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
            trigger.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
            trigger.click();
            return { status: 'opened', current: currentText };
          }
          const candidates = [...document.querySelectorAll('input, .el-select, [role="combobox"], [class*="site" i], button, div, span')]
            .filter(isVisible)
            .map(el => ({ el, text: normalize(el.value || el.innerText || el.textContent), rect: el.getBoundingClientRect() }))
            .filter(item => allLabels.some(label => item.text === label || item.text.startsWith(label + ' ')))
            .filter(item => !inputRect || (Math.abs(item.rect.top - inputRect.top) < 120 && item.rect.left < inputRect.left));
          candidates.sort((a, b) => (a.rect.width * a.rect.height) - (b.rect.width * b.rect.height));
          const current = candidates[0];
          if (!current) return { status: 'not-found', current: '' };
          if (targetLabels.some(label => current.text === label || current.text.startsWith(label + ' '))) {
            return { status: 'already-set', current: current.text };
          }
          const trigger = current.el.closest('.el-select, [role="combobox"], button, [class*="site" i]') || current.el;
          trigger.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, view: window }));
          trigger.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true, view: window }));
          trigger.click();
          return { status: 'opened', current: current.text };
        }""",
                {"targetLabels": list(target_labels), "allLabels": list(all_labels)},
            )
            break
        except Exception as exc:
            error_text = str(exc)
            is_navigation_race = any(
                marker in error_text
                for marker in (
                    "Execution context was destroyed",
                    "navigation",
                    "Target page, context or browser has been closed",
                )
            )
            if not is_navigation_race or attempt == 3:
                raise
            try:
                page.wait_for_load_state("domcontentloaded", timeout=15_000)
            except Exception:
                pass
            page.wait_for_timeout(700 + attempt * 300)
    if state is None:
        raise RuntimeError("SellerSprite marketplace selector could not be evaluated.")
    if state.get("status") == "already-set":
        return {"status": "already-set", "site": marketplace.code, "label": str(state.get("current") or "")}
    if state.get("status") == "not-found":
        raise RuntimeError("Could not find the SellerSprite marketplace selector near the keyword/ASIN input.")

    page.wait_for_timeout(500)
    pattern = re.compile(rf"^\s*(?:{'|'.join(re.escape(label) for label in target_labels)})\s*$")
    options = page.locator('[role="option"], .el-select-dropdown__item, .el-dropdown-menu__item, li')
    clicked = False
    for index in range(options.count() - 1, -1, -1):
        option = options.nth(index)
        try:
            if option.is_visible(timeout=300) and pattern.match(option.inner_text(timeout=500).strip()):
                option.click(timeout=5000)
                clicked = True
                break
        except Exception:
            continue
    if not clicked:
        candidates = page.get_by_text(pattern)
        for index in range(candidates.count() - 1, -1, -1):
            candidate = candidates.nth(index)
            try:
                if candidate.is_visible(timeout=300):
                    candidate.click(timeout=5000)
                    clicked = True
                    break
            except Exception:
                continue
    if not clicked:
        raise RuntimeError(f"SellerSprite marketplace option was not found: {target_labels[0]}")
    try:
        page.wait_for_load_state("domcontentloaded", timeout=15_000)
    except Exception:
        pass
    page.wait_for_timeout(1200)
    return {"status": "changed", "site": marketplace.code, "label": target_labels[0]}


def _matches_amazon_postal_code(text: str, postal_code: str) -> bool:
    compact_text = re.sub(r"[^A-Z0-9]", "", str(text or "").upper())
    compact_postal = re.sub(r"[^A-Z0-9]", "", str(postal_code or "").upper())
    sector = re.sub(r"[A-Z]{2}$", "", compact_postal)
    needles = tuple(dict.fromkeys(value for value in (compact_postal, sector) if len(value) >= 4))
    return any(needle in compact_text for needle in needles)


def amazon_postal_code_matches(text: str, postal_code: str) -> bool:
    return _matches_amazon_postal_code(text, postal_code)


def amazon_location_text(page: Any) -> str:
    location = page.locator("#glow-ingress-line2, #nav-global-location-slot").first
    if not location.count():
        return ""
    try:
        return str(location.text_content(timeout=3000) or "").strip()
    except Exception:
        return ""


def _first_visible_locator(page: Any, selector: str, *, timeout: int = 300) -> Any | None:
    candidates = page.locator(selector)
    try:
        count = candidates.count()
    except Exception:
        return None
    for index in range(count):
        candidate = candidates.nth(index)
        try:
            if candidate.is_visible(timeout=timeout):
                return candidate
        except Exception:
            continue
    return None


def _locator_label(locator: Any) -> str:
    values: list[str] = []
    for getter, argument in (
        ("input_value", None),
        ("text_content", None),
        ("get_attribute", "aria-label"),
    ):
        try:
            method = getattr(locator, getter)
            value = method(timeout=500) if argument is None else method(argument, timeout=500)
        except Exception:
            continue
        if value:
            values.append(str(value))
    return " ".join(values).strip()


def _first_visible_locator_matching(
    page: Any,
    selector: str,
    pattern: re.Pattern[str],
) -> Any | None:
    candidates = page.locator(selector)
    try:
        count = candidates.count()
    except Exception:
        return None
    for index in range(count):
        candidate = candidates.nth(index)
        try:
            if candidate.is_visible(timeout=200) and pattern.search(_locator_label(candidate)):
                return candidate
        except Exception:
            continue
    return None


def _amazon_success_text(page: Any) -> str:
    candidates = page.locator(
        "#GLUXHiddenSuccessSelectedAddressPlaceholder, #GLUXZipConfirmationValue"
    )
    try:
        count = candidates.count()
    except Exception:
        return ""
    for index in range(count):
        try:
            text = str(candidates.nth(index).text_content(timeout=500) or "").strip()
        except Exception:
            continue
        if text:
            return text
    return ""


def _amazon_confirmation_button(page: Any) -> Any | None:
    confirmation = _first_visible_locator(
        page,
        "#GLUXConfirmClose, #GLUXConfirmClose-announce",
    )
    if confirmation is not None:
        return confirmation
    return _first_visible_locator(
        page,
        "button[name='glowDoneButton'], input[name='glowDoneButton']",
    )


def _amazon_visible_error(page: Any) -> str:
    candidates = page.locator(
        "#GLUXZipError, #GLUXZipUnavailable, #GLUXAddressSetError, "
        ".a-alert-inline-error, .glow-toaster-error"
    )
    try:
        count = candidates.count()
    except Exception:
        return ""
    for index in range(count):
        candidate = candidates.nth(index)
        try:
            if not candidate.is_visible(timeout=100):
                continue
            text = str(candidate.text_content(timeout=500) or "").strip()
            if text:
                return re.sub(r"\s+", " ", text)
        except Exception:
            continue
    return ""


def _wait_for_amazon_postal_state(
    page: Any,
    postal_code: str,
    *,
    timeout_ms: int,
) -> tuple[str, str]:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        location_text = amazon_location_text(page)
        if _matches_amazon_postal_code(location_text, postal_code):
            return "header", location_text
        success_text = _amazon_success_text(page)
        if _matches_amazon_postal_code(success_text, postal_code):
            return "confirmation", success_text
        error_text = _amazon_visible_error(page)
        if error_text:
            return "error", error_text
        page.wait_for_timeout(250)
    return "timeout", amazon_location_text(page)


def ensure_amazon_postal_code(page: Any, postal_code: str) -> dict[str, str]:
    postal_code = str(postal_code or "").strip()
    if not postal_code:
        return {"status": "skipped", "text": ""}
    try:
        current_text = amazon_location_text(page)
        if _matches_amazon_postal_code(current_text, postal_code):
            return {"status": "already-set", "text": current_text}

        success_text = _amazon_success_text(page)
        if _matches_amazon_postal_code(success_text, postal_code):
            confirmation = _amazon_confirmation_button(page)
            if confirmation is not None:
                confirmation.click(timeout=5000)
                state, text = _wait_for_amazon_postal_state(page, postal_code, timeout_ms=12000)
                if state == "header":
                    return {"status": "success", "text": text}

        for selector in ("#sp-cc-accept", "#sp-cc-rejectall-link"):
            try:
                button = _first_visible_locator(page, selector, timeout=800)
                if button is not None:
                    button.click(timeout=3000)
                    page.wait_for_timeout(1000)
                    break
            except Exception:
                continue

        entry = None
        for selector in (
            "#nav-global-location-popover-link",
            "#glow-ingress-block",
            "#nav-global-location-slot",
        ):
            entry = _first_visible_locator(page, selector, timeout=1000)
            if entry is not None:
                break
        if entry is None:
            return {
                "status": "failed",
                "error": "Amazon location entry not found",
                "text": amazon_location_text(page),
            }
        entry.click(timeout=10000)
        page.wait_for_timeout(700)

        change = _first_visible_locator(page, "#GLUXChangePostalCodeLink", timeout=500)
        if change is None:
            change = _first_visible_locator_matching(
                page,
                "button, input[type='button'], input[type='submit'], .a-button-text, a",
                re.compile(
                    r"Change address|Change postal|Die Adresse ändern|Adresse ändern|"
                    r"Lieferadresse ändern|Modifier l'adresse|Changer l'adresse|更改地址|配送地址",
                    re.I,
                ),
            )
        if change is not None:
            change.click(timeout=5000)
            page.wait_for_timeout(700)

        postal_input = _first_visible_locator(
            page,
            "#GLUXZipUpdateInput, input[name='zipCode'], input[name='postalCode'], "
            "input[placeholder*='zip' i], input[placeholder*='postal' i], "
            "input[aria-label*='zip' i], input[aria-label*='postal' i], "
            "input[aria-label*='Postleitzahl' i]",
            timeout=1000,
        )
        if postal_input is None:
            return {
                "status": "failed",
                "error": "Amazon postal-code input not visible",
                "text": amazon_location_text(page),
            }
        postal_input.fill(postal_code, timeout=10000)
        page.wait_for_timeout(300)

        apply_button = _first_visible_locator(
            page,
            "#GLUXZipUpdate input.a-button-input, #GLUXZipUpdate button",
            timeout=500,
        )
        if apply_button is None:
            apply_button = _first_visible_locator_matching(
                page,
                "#GLUXZipUpdate, #GLUXZipUpdate-announce, "
                "[id^='a-popover-content-'] input[type='submit'], "
                "[id^='a-popover-content-'] button, "
                "[id^='a-popover-content-'] .a-button-input, "
                "[id^='a-popover-content-'] .a-button-text",
                re.compile(
                    r"^(Apply|Anwenden|Übernehmen|Speichern|Weiter|Appliquer|"
                    r"Valider|Enregistrer|应用|确定|确认)$",
                    re.I,
                ),
            )
        if apply_button is None:
            return {
                "status": "failed",
                "error": "Amazon postal-code apply button not found",
                "text": amazon_location_text(page),
            }
        apply_button.click(timeout=10000)

        state, text = _wait_for_amazon_postal_state(page, postal_code, timeout_ms=15000)
        if state == "confirmation":
            confirmation = _amazon_confirmation_button(page)
            if confirmation is None:
                return {
                    "status": "failed",
                    "error": "Amazon postal-code confirmation button not visible",
                    "text": text,
                }
            confirmation.click(timeout=10000)
            state, text = _wait_for_amazon_postal_state(page, postal_code, timeout_ms=12000)
        elif state == "header":
            confirmation = _amazon_confirmation_button(page)
            if confirmation is not None:
                confirmation.click(timeout=5000)

        if state == "header":
            return {"status": "success", "text": text}
        if state == "error":
            return {"status": "failed", "error": text, "text": amazon_location_text(page)}
        result = {
            "status": "failed",
            "error": f"Amazon location text did not include target postal code {postal_code}",
            "text": text,
        }
        if result.get("status") == "failed" and "location text did not include" in str(result.get("error") or ""):
            # Amazon 有时已保存新地址，但顶部文字要刷新后才更新。
            try:
                page.reload(wait_until="domcontentloaded", timeout=30_000)
                page.wait_for_timeout(1200)
                location = page.locator("#glow-ingress-line2, #nav-global-location-slot").first
                text = str(location.text_content(timeout=3000) or "").strip()
                if _matches_amazon_postal_code(text, postal_code):
                    return {"status": "success", "text": text}
            except Exception:
                pass
        return result
    except Exception as exc:
        return {"status": "failed", "error": str(exc), "text": ""}
