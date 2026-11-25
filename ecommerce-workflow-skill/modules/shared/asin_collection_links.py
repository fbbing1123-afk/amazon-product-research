"""按实测站点规则，从商品数据源直接生成Amazon ASIN集合链接。"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import re
from urllib.parse import urlparse

from openpyxl import load_workbook

RULES = json.loads((Path(__file__).resolve().parents[2] / 'resources/asin_collection_sites.json').read_text(encoding='utf-8'))
SITES = RULES['sites']
ASIN_PATTERN = re.compile(r'[A-Z0-9]{10}')
ALIASES = {'GB': 'UK', 'USA': 'US', '澳大利亚': 'AU', '澳大利亚站': 'AU'}
for code, item in SITES.items():
    for alias in (code, item['label'], item['label'] + '站', item['domain'], 'www.' + item['domain']):
        ALIASES[alias.upper()] = code
HOST_TO_SITE = {host: code for code, item in SITES.items() for host in (item['domain'], 'www.' + item['domain'])}
LINK_HEADERS = {'商品详情页链接', '商品链接', '产品链接', '详情页链接'}


def normalize_site(value):
    code = ALIASES.get(str(value or '').strip().upper())
    if code is None:
        raise ValueError('请指定已验证的站点：' + ', '.join(SITES))
    return code


def normalize_asins(values):
    """去空白、转大写、保序去重；非空错误ASIN不能被悄悄丢弃。"""
    asins, seen = [], set()
    for index, value in enumerate(values, 1):
        asin = str(value if value is not None else '').strip().upper()
        if not asin:
            continue
        if not ASIN_PATTERN.fullmatch(asin):
            raise ValueError(f'第{index}项不是十位ASIN：{asin}')
        if asin not in seen:
            seen.add(asin)
            asins.append(asin)
    if not asins:
        raise ValueError('没有可生成链接的ASIN')
    return asins


def build_collection_url(asins, site):
    """与卖家精灵可见结果一致，使用冒号和竖线，不重复百分号编码。"""
    code = normalize_site(site)
    selected = normalize_asins(asins)
    return 'https://www.' + SITES[code]['domain'] + RULES['path_and_query_prefix'] + RULES['separator'].join(selected)


def source_sheet(book, requested=None):
    if requested:
        matches = [s for s in book if s.title.strip() == requested.strip()]
        if len(matches) != 1:
            raise ValueError(f'找不到数据源页：{requested}')
        return matches[0]
    candidates = [s for s in book if s.title.strip().endswith('数据源') and not any(word in s.title for word in ('评价', '评论'))]
    if len(candidates) != 1:
        raise ValueError('商品数据源页缺失或不唯一，请通过--sheet明确指定')
    return candidates[0]


def infer_workbook_site(book, source=None):
    """源商品链接优先，其次竞品链接；没有站点证据时使用美国站。"""
    from shared.excel_inspect import find_source_sheet

    if source is None:
        source = find_source_sheet(book)
    competitors = [s for s in book if s.title.strip().endswith('竞品分析')]
    groups = [('source_product_links', [source] if source is not None else []),
              ('competitor_links', competitors)]
    for origin, sheets in groups:
        sites = set()
        for sheet in sheets:
            columns = None
            if origin == 'source_product_links':
                columns = {c.column for c in sheet[1] if re.sub(r'\s+', '', str(c.value or '')) in LINK_HEADERS}
            rows = sheet.iter_rows(min_row=2, max_row=sheet.max_row) if columns is not None else sheet.iter_rows(max_row=min(5, sheet.max_row))
            for row in rows:
                for cell in row:
                    if columns is not None and cell.column not in columns:
                        continue
                    values = [str(cell.value or '')]
                    link = getattr(cell, 'hyperlink', None)
                    if link and link.target:
                        values.append(link.target)
                    for value in values:
                        for url in re.findall(r'https?://[^\s\"<>]+', value):
                            host = (urlparse(url).hostname or '').lower()
                            if host in HOST_TO_SITE:
                                sites.add(HOST_TO_SITE[host])
        if len(sites) > 1:
            raise ValueError('商品链接混用多个站点，请明确本次站点')
        if sites:
            return next(iter(sites)), origin
    return 'US', 'default_us'


def collection_from_workbook(workbook, site=None, sheet_name=None):
    """只读原工作簿，不依据销量、子体销量、父ASIN或当前筛选状态删商品。"""
    workbook = Path(workbook)
    fingerprint = hashlib.sha256(workbook.read_bytes()).hexdigest()
    book = load_workbook(workbook, data_only=True, rich_text=True)
    try:
        sheet = source_sheet(book, sheet_name)
        headers = {cell.column: re.sub(r'\s+', '', str(cell.value or '')) for cell in sheet[1]}
        asin_columns = [column for column, label in headers.items() if label.upper() == 'ASIN']
        if len(asin_columns) != 1:
            raise ValueError('数据源必须有唯一的ASIN列，不能用父ASIN列代替')
        asin_column = asin_columns[0]
        link_columns = [column for column, label in headers.items() if label in LINK_HEADERS]
        values, rows_by_asin, domains = [], {}, set()
        blank_rows = []
        for row in sheet.iter_rows(min_row=2):
            cell = row[asin_column - 1]
            asin = str(cell.value if cell.value is not None else '').strip().upper()
            if not asin:
                blank_rows.append(cell.row)
                continue
            if not ASIN_PATTERN.fullmatch(asin):
                raise ValueError(f'{sheet.title}!{cell.coordinate}不是十位ASIN：{asin}')
            values.append(asin)
            rows_by_asin.setdefault(asin, []).append(cell.row)
            for column in link_columns:
                link = row[column - 1]
                urls = [str(link.value or '')]
                if link.hyperlink:
                    urls.append(str(link.hyperlink.target or ''))
                for url in urls:
                    parsed = urlparse(url.strip())
                    if parsed.scheme.lower() in ('http', 'https'):
                        if parsed.hostname not in HOST_TO_SITE:
                            raise ValueError(f'{sheet.title}!{link.coordinate}的商品链接域名不在已验证站点中')
                        domains.add(HOST_TO_SITE[parsed.hostname])
        asins = normalize_asins(values)
        requested_site = normalize_site(site) if site else None
        if len(domains) > 1:
            raise ValueError('数据源商品链接混用多个站点，请先明确本次站点的数据源')
        detected = next(iter(domains), None)
        if requested_site and detected and requested_site != detected:
            raise ValueError(f'指定站点{requested_site}与数据源商品链接站点{detected}不一致')
        inferred_site, inferred_origin = infer_workbook_site(book, source=sheet) if not requested_site and not detected else (None, None)
        selected_site = requested_site or detected or inferred_site
        site_origin = 'explicit' if requested_site else ('source_product_links' if detected else inferred_origin)
        return {'schema': 'amazon-asin-collection-link.v1', 'source_workbook': str(workbook.resolve()),
                'source_sha256': fingerprint, 'source_sheet': sheet.title, 'asin_column': asin_column,
                'site': selected_site, 'site_source': site_origin,
                'source_asin_rows': len(values), 'asin_count': len(asins), 'duplicate_rows': len(values) - len(asins),
                'blank_asin_rows': blank_rows, 'asins': asins, 'source_rows': rows_by_asin,
                'url': build_collection_url(asins, selected_site), 'rules_verified_on': RULES['observed_on']}
    finally:
        book.close()
