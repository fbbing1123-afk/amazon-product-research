from __future__ import annotations

import io
import math
import re
import statistics
from copy import copy, deepcopy
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SOURCE_SHEET = "数据源"
REPORT_SHEETS = ["报表总览", "销量分层", "价格段分析", "评分数分析", "Top ASIN", "异常样本", "数据分析报表"]
PRICE_BUCKET_RANGES = ["0-20", "20-40", "40-60", "60-80", "80-100", "100+"]
CURRENCY_NAMES = {"$": "美元", "€": "欧元", "£": "英镑", "¥": "日元", "￥": "人民币"}
SOURCE_FILTER_HEADERS = ("销量分层", "评分数区间", "价格段")
EXCLUDED_FROM_ANALYSIS = "未纳入市场分析"
UNKNOWN_PRICE_BUCKET = "未知价格"
SLICER_SHEET_NAME = "切片联动"
SLICER_TABLE_HEADER_ROW = 12
SLICER_SOURCE_ROW_OFFSET = SLICER_TABLE_HEADER_ROW - 1


def find_source_sheet(wb):
    for ws in wb.worksheets:
        # 合并评价来源同样以“数据源”结尾，不能被当作商品销售数据。
        if ws.title.endswith(SOURCE_SHEET) and ws.title not in (SOURCE_SHEET, '评价数据源'):
            return ws
    if SOURCE_SHEET in wb.sheetnames:
        return wb[SOURCE_SHEET]
    for name in ("US", "Sheet1"):
        if name in wb.sheetnames:
            return wb[name]
    return None


def infer_product_name(wb):
    for ws in wb.worksheets:
        if ws.title.endswith("市场调研"):
            product_name = ws.title[: -len("市场调研")].strip("-_ ")
            if product_name:
                return product_name
    return None


def canonicalize_product_sheet_names(wb, product_name: str | None) -> None:
    """把复用模板中的旧产品名前缀统一为本次产品名。"""
    normalized = str(product_name or "").strip("-_ ")
    if not normalized:
        return
    for suffix in ("市场调研", "市场分析", "竞品分析", SOURCE_SHEET):
        target_name = f"{normalized}-{suffix}"
        if target_name in wb.sheetnames:
            continue
        worksheet = next(
            (sheet for sheet in wb.worksheets if sheet.title.endswith(suffix)),
            None,
        )
        if worksheet is not None:
            worksheet.title = target_name


def ensure_canonical_source_sheet_name(wb, ws):
    if ws.title != SOURCE_SHEET:
        return ws
    product_name = infer_product_name(wb)
    if not product_name:
        return ws
    canonical_name = f"{product_name}-{SOURCE_SHEET}"
    if canonical_name not in wb.sheetnames:
        ws.title = canonical_name
        return ws
    return wb[canonical_name]


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = str(value).strip()
    text = re.sub(r"[$€£¥￥,\s]", "", text)
    text = text.replace("USD", "").replace("EUR", "").replace("GBP", "").replace("JPY", "").replace("CNY", "")
    text = text.replace("%", "")
    if text in {"", "-", "--", "N/A", "无"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def price_bucket_order(symbol: str, definitions: Optional[List[Dict[str, Any]]] = None) -> List[str]:
    """返回价格段顺序；未传动态定义时保留旧规则作为兼容兜底。"""
    if definitions is not None:
        return [str(item["label"]) for item in definitions]
    return [f"{symbol}{name}" for name in PRICE_BUCKET_RANGES]


def header_label(name: str, symbol: str) -> str:
    return f"{name}({symbol})"


def currency_name(symbol: str) -> str:
    return CURRENCY_NAMES.get(symbol, symbol)


def infer_currency_symbol(rows: List[Dict[str, Any]]) -> str:
    symbols = ("€", "$", "£", "¥", "￥")
    for row in rows:
        for key in row.keys():
            text = str(key or "")
            match = re.search(r"\(([$€£¥￥])\)", text)
            if match and ("价格" in text or "销售额" in text):
                return match.group(1)
    for row in rows:
        for value in row.values():
            text = str(value or "")
            for symbol in symbols:
                if symbol in text:
                    return symbol
    return "$"


def first_field_value(row: Dict[str, Any], base_name: str, exact_names: Optional[List[str]] = None) -> Any:
    for name in exact_names or []:
        if name in row:
            return row.get(name)
    if base_name in row:
        return row.get(base_name)
    prefix = f"{base_name}("
    for key, value in row.items():
        text = str(key or "")
        if text.startswith(prefix):
            return value
    return None


def median(values: Iterable[float]) -> float:
    nums = [v for v in values if v is not None]
    return float(statistics.median(nums)) if nums else 0.0


def avg(values: Iterable[float]) -> float:
    nums = [v for v in values if v is not None]
    return float(sum(nums) / len(nums)) if nums else 0.0


def whole(value: Any) -> int:
    number = to_number(value)
    return int(number) if number is not None else 0


def rating(value: Any) -> Optional[float]:
    number = to_number(value)
    return round(number, 1) if number is not None else None


def _quantile(values: List[float], probability: float) -> float:
    ordered = sorted(values)
    if not ordered:
        return 0.0
    position = (len(ordered) - 1) * probability
    lower = math.floor(position)
    upper = math.ceil(position)
    if lower == upper:
        return float(ordered[lower])
    weight = position - lower
    return float(ordered[lower] + (ordered[upper] - ordered[lower]) * weight)


def _nice_price_step(raw_step: float) -> float:
    """把计算出的价格步长取整为容易阅读的货币单位。"""
    if raw_step <= 0:
        return 1.0
    magnitude = 10 ** math.floor(math.log10(raw_step))
    normalized = raw_step / magnitude
    multiplier = 1.0 if normalized <= 1.5 else 2.0 if normalized <= 3 else 5.0 if normalized <= 7 else 10.0
    return multiplier * magnitude


def _price_display_step(prices: List[float]) -> float:
    # 使用中位价格决定展示精度，避免少数高价异常值把低价区间取整得过粗。
    reference = statistics.median(prices) if prices else 1.0
    if reference < 20:
        return 1.0
    if reference < 100:
        return 5.0
    if reference < 300:
        return 10.0
    if reference < 1000:
        return 25.0
    return 50.0


def _format_price_bound(value: float, step: float) -> str:
    decimals = 2 if step < 1 else 1 if not float(step).is_integer() else 0
    text = f"{value:.{decimals}f}"
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text or "0"


def _price_label(symbol: str, lower: float, upper: float, step: float) -> str:
    lower_text = _format_price_bound(lower, step)
    upper_text = _format_price_bound(upper, step)
    if lower_text == upper_text:
        return f"{symbol}{lower_text}"
    return f"{symbol}{lower_text}-{upper_text}"


def _adaptive_price_step(values: List[float], display_step: float) -> float:
    """根据价格离散程度计算连续价格带的展示步长。"""
    span = values[-1] - values[0]
    q1 = _quantile(values, 0.25)
    q3 = _quantile(values, 0.75)
    iqr = q3 - q1
    raw_step = 2 * iqr / (len(values) ** (1 / 3)) if iqr > 0 else span / max(3, math.sqrt(len(values)))
    return max(_nice_price_step(raw_step), _nice_price_step(display_step))


def _build_price_segment(
    symbol: str,
    raw_lower: float,
    raw_upper: float,
    next_lower: Optional[float],
    step: float,
) -> Dict[str, Any]:
    lower = math.floor(raw_lower / step) * step
    upper = math.ceil(raw_upper / step) * step
    return {
        "label": _price_label(symbol, raw_lower, raw_upper, step),
        "lower": raw_lower,
        "upper_exclusive": next_lower,
    }


def _subdivide_price_segment(
    symbol: str,
    raw_lower: float,
    raw_upper: float,
    next_lower: Optional[float],
    step: float,
    display_step: float,
) -> List[Dict[str, Any]]:
    """把断层内部过宽的连续价格带细分，不按样本数合并。"""
    if raw_upper - raw_lower <= step * 1.5:
        return [_build_price_segment(symbol, raw_lower, raw_upper, next_lower, display_step)]

    segments: List[Dict[str, Any]] = []
    lower = math.floor(raw_lower / step) * step
    while lower < raw_upper:
        upper = min(raw_upper, lower + step)
        is_last = upper >= raw_upper
        segment_lower = raw_lower if not segments else lower
        segment_upper = upper
        segments.append({
            "label": _price_label(symbol, segment_lower, segment_upper, display_step),
            "lower": segment_lower,
            "upper_exclusive": next_lower if is_last else upper,
        })
        lower = upper
    return segments


def build_dynamic_price_buckets(prices: List[float], symbol: str = "$") -> List[Dict[str, Any]]:
    """按价格断层分组，并细分断层内部过宽的连续价格带。"""
    values = sorted(float(price) for price in prices if price is not None and float(price) >= 0)
    if not values:
        return []

    display_step = _price_display_step(values)
    if values[0] == values[-1]:
        return [{
            "label": _price_label(symbol, values[0], values[-1], display_step),
            "lower": values[0],
            "upper_exclusive": None,
        }]

    positive_gaps = [right - left for left, right in zip(values, values[1:]) if right > left]
    typical_gap = statistics.median(positive_gaps) if positive_gaps else 0.0
    break_indexes: List[int] = []
    if typical_gap > 0:
        minimum_gap = max(display_step, typical_gap * 3)
        for index, (left, right) in enumerate(zip(values, values[1:])):
            gap = right - left
            relative_gap = gap / max(abs(left), display_step)
            if gap >= minimum_gap and relative_gap >= 0.10:
                break_indexes.append(index)

    step = _adaptive_price_step(values, display_step)
    if break_indexes:
        segments: List[Dict[str, Any]] = []
        start_index = 0
        for break_index in [*break_indexes, len(values) - 1]:
            end_index = break_index
            raw_lower = values[start_index]
            raw_upper = values[end_index]
            next_lower = values[end_index + 1] if end_index < len(values) - 1 else None
            segments.extend(_subdivide_price_segment(symbol, raw_lower, raw_upper, next_lower, step, display_step))
            start_index = end_index + 1
        return segments

    # 没有明显价格断层时，用四分位距计算连续价格的自适应显示步长。
    start = math.floor(values[0] / step) * step
    end = math.ceil(values[-1] / step) * step
    if end <= start:
        end = start + step

    segments = []
    lower = start
    while lower < end:
        upper = min(end, lower + step)
        segments.append({
            "label": _price_label(symbol, lower, upper, step),
            "lower": lower,
            "upper_exclusive": None if upper >= end else upper,
        })
        lower = upper
    return segments


def bucket_price(
    price: Optional[float],
    symbol: str = "$",
    definitions: Optional[List[Dict[str, Any]]] = None,
) -> str:
    if price is None:
        return UNKNOWN_PRICE_BUCKET
    if definitions:
        number = float(price)
        for definition in definitions:
            upper_exclusive = definition.get("upper_exclusive")
            if upper_exclusive is None or number < upper_exclusive:
                return str(definition["label"])
        return str(definitions[-1]["label"])
    if price <= 20:
        return f"{symbol}0-20"
    if price <= 40:
        return f"{symbol}20-40"
    if price < 60:
        return f"{symbol}40-60"
    if price < 80:
        return f"{symbol}60-80"
    if price < 100:
        return f"{symbol}80-100"
    return f"{symbol}100+"


def bucket_reviews(reviews: Optional[float]) -> str:
    if reviews is None:
        return "未知评分数"
    if reviews <= 30:
        return "≤30评"
    if reviews <= 50:
        return "31-50评"
    if reviews <= 100:
        return "51-100评"
    if reviews <= 300:
        return "101-300评"
    if reviews <= 1000:
        return "301-1000评"
    return "1000评+"


def summarize_group(rows: List[Dict[str, Any]]) -> Dict[str, float]:
    sales = [to_number(r.get("月销量")) or 0 for r in rows]
    prices = [to_number(r.get("价格")) for r in rows if to_number(r.get("价格")) is not None]
    reviews = [to_number(r.get("评分数")) for r in rows if to_number(r.get("评分数")) is not None]
    return {
        "sample_count": len(rows),
        "total_sales": float(sum(sales)),
        "avg_sales": avg(sales),
        "median_sales": median(sales),
        "avg_price": avg(prices),
        "median_price": median(prices),
        "median_reviews": median(reviews),
    }


def compute_metrics(rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    symbol = infer_currency_symbol(rows)
    cleaned = []
    review_records = []
    for row in rows:
        sales = to_number(row.get("月销量"))
        reviews = to_number(row.get("评分数"))
        # 评分数分析不能依赖月销量；缺销量的高评分数产品仍需进入对应区间。
        if sales is None and reviews is None:
            continue
        item = dict(row)
        price = to_number(first_field_value(row, "价格", ["价格($)", "价格(€)", "价格(£)", "价格(¥)", "价格(￥)", "价格"]))
        revenue = to_number(first_field_value(row, "月销售额", ["月销售额($)", "月销售额(€)", "月销售额(£)", "月销售额(¥)", "月销售额(￥)", "月销售额"]))
        if revenue is None and price is not None and sales is not None:
            revenue = sales * price
        item["月销量"] = sales if sales is not None else 0.0
        item["月销售额"] = revenue or 0
        item["价格"] = price
        item["评分数"] = reviews
        item["评分"] = to_number(row.get("评分"))
        item["上架天数"] = to_number(row.get("上架天数"))
        item["销量评分比"] = item["月销量"] / item["评分数"] if item["评分数"] and item["评分数"] > 0 else item["月销量"]
        review_records.append(item)
        if sales is not None:
            cleaned.append(item)

    ranked = sorted(cleaned, key=lambda r: r["月销量"], reverse=True)
    total_sales = sum(r["月销量"] for r in ranked)
    total_revenue = sum(r["月销售额"] for r in ranked)
    revenues = [r["月销售额"] for r in ranked]
    prices = [r["价格"] for r in ranked if r["价格"] is not None]
    reviews = [r["评分数"] for r in review_records if r["评分数"] is not None]
    sales = [r["月销量"] for r in ranked]

    def share(n: int) -> float:
        return sum(r["月销量"] for r in ranked[:n]) / total_sales if total_sales else 0

    layers = {
        "Top 10": ranked[:10],
        "第11-20名": ranked[10:20],
        "第21-50名": ranked[20:50],
        "第51名以后": ranked[50:],
    }
    for layer_name, layer_rows in layers.items():
        for row in layer_rows:
            row["_sales_layer"] = layer_name

    dynamic_price_buckets = build_dynamic_price_buckets(prices, symbol)
    price_groups: Dict[str, List[Dict[str, Any]]] = {
        str(definition["label"]): []
        for definition in dynamic_price_buckets
    }
    review_groups: Dict[str, List[Dict[str, Any]]] = {}
    for row in ranked:
        price_bucket = bucket_price(row.get("价格"), symbol, dynamic_price_buckets)
        row["_price_bucket"] = price_bucket
        price_groups.setdefault(price_bucket, []).append(row)
    for row in review_records:
        review_groups.setdefault(bucket_reviews(row.get("评分数")), []).append(row)
    if any(row.get("价格") is None for row in ranked):
        price_groups[UNKNOWN_PRICE_BUCKET] = [row for row in ranked if row.get("价格") is None]

    low_review_sales = [r["月销量"] for r in ranked if r.get("评分数") is not None and r["评分数"] <= 30]
    low_review_median = median(low_review_sales)

    anomalies = []
    for row in ranked:
        reasons = []
        review_count = row.get("评分数")
        days = row.get("上架天数")
        if review_count is not None and review_count <= 30 and row["月销量"] >= max(80, low_review_median):
            reasons.append("低评分数高销量")
        if row.get("销量评分比", 0) >= 8 and row["月销量"] >= 50:
            reasons.append("月销量/评分数比值偏高")
        if days is not None and days <= 180 and row["月销量"] >= 80:
            reasons.append("上架时间短但销量较高")
        rank_text = str(row.get("搜索排名") or "")
        if "广告位" in rank_text and row["月销量"] >= 80:
            reasons.append("广告位高销量需复核自然承接")
        if reasons:
            anomalies.append({
                "ASIN": row.get("ASIN"),
                "品牌": row.get("品牌"),
                "商品标题": row.get("商品标题"),
                "月销量": row.get("月销量"),
                "价格": row.get("价格"),
                "评分数": row.get("评分数"),
                "评分": row.get("评分"),
                "上架天数": row.get("上架天数"),
                "销量评分比": row.get("销量评分比"),
                "搜索排名": row.get("搜索排名"),
                "异常原因": "；".join(reasons),
                "复核建议": "检查广告、站外、品牌流量、老链接换壳、评价沉淀是否匹配",
            })

    return {
        "sample_count": len(ranked),
        "currency_symbol": symbol,
        "currency_name": currency_name(symbol),
        "price_bucket_order": price_bucket_order(symbol, dynamic_price_buckets),
        "total_sales": float(total_sales),
        "total_revenue": float(total_revenue),
        "median_revenue": median(revenues),
        "avg_sales": avg(sales),
        "median_sales": median(sales),
        "avg_price": avg(prices),
        "median_price": median(prices),
        "avg_reviews": avg(reviews),
        "median_reviews": median(reviews),
        "top5_share": share(5),
        "top10_share": share(10),
        "top20_share": share(20),
        "ranked": ranked,
        "sales_layers": {k: summarize_group(v) for k, v in layers.items()},
        "price_buckets": {k: summarize_group(v) for k, v in price_groups.items()},
        "review_buckets": {k: summarize_group(v) for k, v in review_groups.items()},
        "anomalies": anomalies[:20],
    }


def read_rows(ws) -> List[Dict[str, Any]]:
    headers = [cell.value for cell in ws[1]]
    rows = []
    for source_row, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        item = {
            headers[i]: row[i]
            for i in range(min(len(headers), len(row)))
            if headers[i] and str(headers[i]).strip() not in SOURCE_FILTER_HEADERS
        }
        if any(v is not None and v != "" for v in item.values()):
            item["_source_row"] = source_row
            rows.append(item)
    return rows


def write_source_filter_columns(
    ws,
    rows: List[Dict[str, Any]],
    metrics: Dict[str, Any],
    *,
    header_row: int = 1,
    source_row_offset: int = 0,
    minimum_new_column: int = 1,
) -> None:
    header_to_column = {
        str(ws.cell(header_row, column).value).strip(): column
        for column in range(1, ws.max_column + 1)
        if ws.cell(header_row, column).value not in (None, "")
    }
    next_column = max(ws.max_column + 1, minimum_new_column)
    filter_columns: Dict[str, int] = {}
    for header in SOURCE_FILTER_HEADERS:
        column = header_to_column.get(header)
        if column is None:
            column = next_column
            next_column += 1
        filter_columns[header] = column

    style_source_column = max(1, min(filter_columns.values()) - 1)
    for header, column in filter_columns.items():
        cell = ws.cell(header_row, column, header)
        source = ws.cell(header_row, style_source_column)
        if source.has_style:
            cell._style = copy(source._style)
        cell.font = copy(source.font)
        cell.fill = copy(source.fill)
        cell.border = copy(source.border)
        cell.alignment = copy(source.alignment)
        cell.number_format = source.number_format
        cell.protection = copy(source.protection)
        ws.column_dimensions[get_column_letter(column)].width = max(
            14, ws.column_dimensions[get_column_letter(column)].width or 0
        )

    ranked_by_source_row = {
        row.get("_source_row"): row
        for row in metrics.get("ranked", [])
        if row.get("_source_row") is not None
    }
    currency_symbol = metrics.get("currency_symbol", "$")
    source_rows = {row.get("_source_row") for row in rows if row.get("_source_row") is not None}
    for row_number in range(header_row + 1, ws.max_row + 1):
        source_row_number = row_number - source_row_offset
        ranked_row = ranked_by_source_row.get(source_row_number)
        if ranked_row is None:
            values = (EXCLUDED_FROM_ANALYSIS,) * len(SOURCE_FILTER_HEADERS)
        else:
            values = (
                ranked_row.get("_sales_layer", EXCLUDED_FROM_ANALYSIS),
                bucket_reviews(ranked_row.get("评分数")),
                ranked_row.get("_price_bucket") or bucket_price(ranked_row.get("价格"), currency_symbol),
            )
        for header, value in zip(SOURCE_FILTER_HEADERS, values):
            ws.cell(row_number, filter_columns[header], value)

    # Clear stale helper values below records when a source sheet became shorter before a rerun.
    for row_number in range(header_row + 1, ws.max_row + 1):
        if row_number - source_row_offset in source_rows:
            continue
        for column in filter_columns.values():
            ws.cell(row_number, column, None)


def create_source_slicer_sheet(wb, source_ws, rows: List[Dict[str, Any]], metrics: Dict[str, Any]):
    if SLICER_SHEET_NAME in wb.sheetnames:
        del wb[SLICER_SHEET_NAME]
    sheet_index = wb.worksheets.index(source_ws) + 1
    ws = wb.create_sheet(SLICER_SHEET_NAME, sheet_index)

    for source_row in source_ws.iter_rows():
        for source_cell in source_row:
            target = ws.cell(source_cell.row + SLICER_SOURCE_ROW_OFFSET, source_cell.column, source_cell.value)
            if source_cell.has_style:
                target._style = copy(source_cell._style)
            target.number_format = source_cell.number_format
            target.protection = copy(source_cell.protection)
            target.alignment = copy(source_cell.alignment)
            if source_cell.hyperlink:
                target._hyperlink = copy(source_cell.hyperlink)
                target._hyperlink.ref = target.coordinate
            if source_cell.comment:
                target.comment = copy(source_cell.comment)

    for key, dimension in source_ws.column_dimensions.items():
        target_dimension = ws.column_dimensions[key]
        target_dimension.width = dimension.width
        target_dimension.hidden = dimension.hidden
        target_dimension.bestFit = dimension.bestFit
        target_dimension.outlineLevel = dimension.outlineLevel
    for row_number, dimension in source_ws.row_dimensions.items():
        target_dimension = ws.row_dimensions[row_number + SLICER_SOURCE_ROW_OFFSET]
        target_dimension.height = dimension.height
        target_dimension.hidden = dimension.hidden
        target_dimension.outlineLevel = dimension.outlineLevel

    for merged_range in source_ws.merged_cells.ranges:
        ws.merge_cells(
            start_row=merged_range.min_row + SLICER_SOURCE_ROW_OFFSET,
            start_column=merged_range.min_col,
            end_row=merged_range.max_row + SLICER_SOURCE_ROW_OFFSET,
            end_column=merged_range.max_col,
        )

    rebuilt_source_images = []
    for source_image in getattr(source_ws, "_images", []):
        marker = getattr(getattr(source_image, "anchor", None), "_from", None)
        if marker is None:
            rebuilt_source_images.append(source_image)
            continue
        raw_image = source_image._data()
        source_replacement = XLImage(io.BytesIO(raw_image))
        source_replacement.width = source_image.width
        source_replacement.height = source_image.height
        source_replacement.anchor = deepcopy(source_image.anchor)
        rebuilt_source_images.append(source_replacement)

        image = XLImage(io.BytesIO(raw_image))
        image.width = source_image.width
        image.height = source_image.height
        target_coord = ws.cell(
            marker.row + 1 + SLICER_SOURCE_ROW_OFFSET,
            marker.col + 1,
        ).coordinate
        ws.add_image(image, target_coord)
    source_ws._images = rebuilt_source_images

    write_source_filter_columns(
        ws,
        rows,
        metrics,
        header_row=SLICER_TABLE_HEADER_ROW,
        source_row_offset=SLICER_SOURCE_ROW_OFFSET,
        minimum_new_column=13,
    )
    for header in SOURCE_FILTER_HEADERS:
        for column in range(1, ws.max_column + 1):
            if ws.cell(SLICER_TABLE_HEADER_ROW, column).value == header:
                ws.column_dimensions[get_column_letter(column)].hidden = True
                break

    title_end_column = min(12, max(6, ws.max_column))
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=title_end_column)
    title = ws.cell(1, 1, "切片联动")
    title.fill = PatternFill("solid", fgColor="E58A00")
    title.font = Font(color="FFFFFF", bold=True, size=14)
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    panel_specs = ((1, 4, "销量分层"), (5, 8, "评分数分层"), (9, 12, "价格段分层"))
    for start_column, end_column, label in panel_specs:
        ws.merge_cells(start_row=2, start_column=start_column, end_row=10, end_column=start_column)
        label_cell = ws.cell(2, start_column, label)
        label_cell.fill = PatternFill("solid", fgColor="FFF200")
        label_cell.font = Font(color="000000", bold=True)
        label_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row_number in range(2, 11):
            for column in range(start_column + 1, end_column + 1):
                ws.cell(row_number, column).fill = PatternFill("solid", fgColor="FFFBE6")
    for row_number in range(2, 11):
        ws.row_dimensions[row_number].height = 22
    ws.row_dimensions[11].height = 8
    ws.freeze_panes = f"A{SLICER_TABLE_HEADER_ROW + 1}"
    ws.sheet_view.showGridLines = False
    return ws


def style_sheet(ws, freeze="A2"):
    ws.freeze_panes = freeze
    ws.sheet_view.showGridLines = False
    widths = {"A": 18, "B": 18, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 18, "I": 20, "J": 28, "K": 42, "L": 42}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def style_header(cell, fill="1F4E78"):
    cell.font = Font(color="FFFFFF", bold=True)
    cell.fill = PatternFill("solid", fgColor=fill)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def write_row(ws, row_num: int, values: List[Any], header: bool = False, style_headers: bool = True):
    for col_num, value in enumerate(values, 1):
        cell = ws.cell(row_num, col_num, value)
        if header and style_headers:
            style_header(cell)
        elif isinstance(value, int):
            cell.number_format = "0"
        elif isinstance(value, float):
            cell.number_format = "0.0"


def add_chart(ws, title: str, anchor: str, min_col: int, min_row: int, max_row: int, cat_col: int = 1):
    if max_row <= min_row:
        return
    chart = BarChart()
    chart.y_axis.title = None
    data = Reference(ws, min_col=min_col, min_row=min_row, max_row=max_row)
    cats = Reference(ws, min_col=cat_col, min_row=min_row + 1, max_row=max_row)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.x_axis.title = None
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True
    chart.dataLabels.showSerName = False
    chart.dataLabels.showCatName = False
    chart.dataLabels.showLegendKey = False
    ws.add_chart(chart, anchor)


def pct(v: float) -> str:
    return f"{v:.1%}"


def money(v: float, symbol: str = "$") -> str:
    return f"{symbol}{whole(v):,}"


def clean_old_report_sheets(wb):
    for name in REPORT_SHEETS:
        if name in wb.sheetnames:
            del wb[name]


def create_overview_sheet(wb, metrics: Dict[str, Any]):
    ws = wb.create_sheet("报表总览", 0)
    style_sheet(ws, "A12")
    symbol = metrics.get("currency_symbol", "$")
    ws.merge_cells("A1:L1")
    ws["A1"] = "市场数据分析报表"
    ws["A1"].font = Font(bold=True, size=20, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="17365D")
    ws["A1"].alignment = Alignment(horizontal="center")

    concentration = "高" if metrics["top10_share"] >= 0.55 else "中" if metrics["top10_share"] >= 0.35 else "低"
    opportunity = "谨慎进入" if concentration == "高" else "可重点评估"
    if metrics["sample_count"] < 20:
        opportunity = "样本偏少，建议补充数据后再决策"
    conclusions = [
        f"样本共 {metrics['sample_count']} 个链接，总月销量 {whole(metrics['total_sales']):,}，总月销售额 {money(metrics['total_revenue'], symbol)}，说明该方向存在真实需求。",
        f"平均月销量 {whole(metrics['avg_sales'])}，中位月销量 {whole(metrics['median_sales'])}；如果均值明显高于中位数，说明市场受少数头部链接拉动。",
        f"Top10 销量占比 {pct(metrics['top10_share'])}，头部集中度判断为：{concentration}。",
        f"中位价格 {whole(metrics['median_price'])} {metrics.get('currency_name', currency_name(symbol))}，建议优先研究销量排名第 11-20 名与第 21-50 名的中腰部样本。",
        f"最终建议：{opportunity}。重点复核低评分高销量、销量/评分数比值异常、上架时间短但销量高的 ASIN。",
    ]
    ws["A3"] = "结论摘要"
    ws["A3"].font = Font(bold=True, size=14, color="1F4E78")
    for idx, text in enumerate(conclusions, 4):
        ws.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=12)
        ws.cell(idx, 1, text).fill = PatternFill("solid", fgColor="EAF2F8")
        ws.cell(idx, 1).alignment = Alignment(wrap_text=True, vertical="top")

    write_row(ws, 11, ["指标", "值", "指标", "值"], True)
    support = [
        ["总月销量", whole(metrics["total_sales"]), "样本数", whole(metrics["sample_count"])],
        ["总月销售额", whole(metrics["total_revenue"]), "平均月销量", whole(metrics["avg_sales"])],
        ["中位月销量", whole(metrics["median_sales"]), header_label("平均价格", symbol), whole(metrics["avg_price"])],
        [header_label("中位价格", symbol), whole(metrics["median_price"]), "平均评分数", whole(metrics["avg_reviews"])],
        ["中位评分数", whole(metrics["median_reviews"]), "Top5销量占比", pct(metrics["top5_share"])],
        ["Top10销量占比", pct(metrics["top10_share"]), "Top20销量占比", pct(metrics["top20_share"])],
    ]
    for row_num, row in enumerate(support, 12):
        write_row(ws, row_num, row)
    return ws


def create_layer_sheet(wb, metrics: Dict[str, Any]):
    ws = wb.create_sheet("销量分层")
    style_sheet(ws)
    symbol = metrics.get("currency_symbol", "$")
    write_row(ws, 1, ["分层", "样本数", "总月销量", "平均月销量", "中位月销量", header_label("平均价格", symbol), "中位评分数"], True)
    for row_num, (name, data) in enumerate(metrics["sales_layers"].items(), 2):
        write_row(ws, row_num, [name, whole(data["sample_count"]), whole(data["total_sales"]), whole(data["avg_sales"]), whole(data["median_sales"]), whole(data["avg_price"]), whole(data["median_reviews"])])
    add_chart(ws, "销量排名分层总月销量", "I2", 3, 1, 1 + len(metrics["sales_layers"]))
    return ws


def create_price_sheet(wb, metrics: Dict[str, Any]):
    ws = wb.create_sheet("价格段分析")
    style_sheet(ws)
    symbol = metrics.get("currency_symbol", "$")
    write_row(ws, 1, ["价格段", "样本数", "总月销量", "平均月销量", "中位月销量", header_label("中位价格", symbol)], True)
    order = metrics.get("price_bucket_order", price_bucket_order(symbol)) + [UNKNOWN_PRICE_BUCKET]
    row_num = 2
    for name in order:
        if name not in metrics["price_buckets"]:
            continue
        data = metrics["price_buckets"][name]
        write_row(ws, row_num, [name, whole(data["sample_count"]), whole(data["total_sales"]), whole(data["avg_sales"]), whole(data["median_sales"]), whole(data["median_price"])])
        row_num += 1
    add_chart(ws, "价格段总月销量", "I2", 3, 1, row_num - 1)
    return ws


def create_review_sheet(wb, metrics: Dict[str, Any]):
    ws = wb.create_sheet("评分数分析")
    style_sheet(ws)
    write_row(ws, 1, ["评分数区间", "样本数", "总月销量", "平均月销量", "中位月销量", "中位评分数"], True)
    order = ["≤30评", "31-50评", "51-100评", "101-300评", "301-1000评", "1000评+", "未知评分数"]
    row_num = 2
    for name in order:
        if name not in metrics["review_buckets"]:
            continue
        data = metrics["review_buckets"][name]
        write_row(ws, row_num, [name, whole(data["sample_count"]), whole(data["total_sales"]), whole(data["avg_sales"]), whole(data["median_sales"]), whole(data["median_reviews"])])
        row_num += 1
    add_chart(ws, "评分数区间总月销量", "I2", 3, 1, row_num - 1)
    return ws


def create_top_sheet(wb, metrics: Dict[str, Any]):
    ws = wb.create_sheet("Top ASIN")
    style_sheet(ws)
    symbol = metrics.get("currency_symbol", "$")
    write_row(ws, 1, ["ASIN", "品牌", "月销量", header_label("价格", symbol), "评分数", "评分", "上架天数", "搜索排名", "标题"], True)
    for row_num, row in enumerate(metrics["ranked"][:30], 2):
        write_row(ws, row_num, [
            row.get("ASIN"), row.get("品牌"), whole(row.get("月销量")), whole(row.get("价格")), whole(row.get("评分数")), rating(row.get("评分")), whole(row.get("上架天数")), row.get("搜索排名"), str(row.get("商品标题") or "")[:100]
        ])
    add_chart(ws, "Top ASIN 月销量", "K2", 3, 1, min(31, len(metrics["ranked"]) + 1))
    return ws


def create_anomaly_sheet(wb, metrics: Dict[str, Any]):
    ws = wb.create_sheet("异常样本")
    style_sheet(ws)
    symbol = metrics.get("currency_symbol", "$")
    write_row(ws, 1, ["ASIN", "品牌", "月销量", header_label("价格", symbol), "评分数", "评分", "销量/评分数", "上架天数", "异常原因", "复核建议", "标题"], True)
    for row_num, row in enumerate(metrics["anomalies"], 2):
        write_row(ws, row_num, [
            row.get("ASIN"), row.get("品牌"), whole(row.get("月销量")), whole(row.get("价格")), whole(row.get("评分数")), rating(row.get("评分")), whole(row.get("销量评分比")), whole(row.get("上架天数")), row.get("异常原因"), row.get("复核建议"), str(row.get("商品标题") or "")[:100]
        ])
    return ws


def market_analysis_sheet_name(wb) -> str:
    for name in wb.sheetnames:
        if name.endswith("市场分析"):
            return name
    for name in wb.sheetnames:
        if name.endswith("市场调研"):
            return name[:-4] + "市场分析"
    return "市场分析"


def section_title(ws, row: int, title: str):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    cell = ws.cell(row, 1, title)
    cell.font = Font(bold=True, size=13, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="left", vertical="center")


def build_sales_conclusion(metrics: Dict[str, Any]) -> str:
    concentration = "高" if metrics["top10_share"] >= 0.55 else "中" if metrics["top10_share"] >= 0.35 else "低"
    return f"小结：Top10销量占比{pct(metrics['top10_share'])}，头部集中度{concentration}；平均月销量{whole(metrics['avg_sales'])}，中位月销量{whole(metrics['median_sales'])}，建议重点看中腰部链接是否具备可追赶空间。"


def build_review_conclusion(metrics: Dict[str, Any]) -> str:
    low = metrics["review_buckets"].get("≤30评", {})
    low_median = whole(low.get("median_sales", 0))
    low_avg = whole(low.get("avg_sales", 0))
    return f"小结：低评分数样本平均月销量{low_avg}，中位月销量{low_median}；如果均值明显高于中位数，说明低评起量更多来自少数异常样本，不能简单判断为低评好打。"


def build_price_conclusion(metrics: Dict[str, Any]) -> str:
    buckets = metrics["price_buckets"]
    best_name = ""
    best_sales = -1
    for name, data in buckets.items():
        if name == UNKNOWN_PRICE_BUCKET:
            continue
        sales = whole(data.get("total_sales", 0))
        if sales > best_sales:
            best_name = name
            best_sales = sales
    return f"小结：当前销量贡献最高的价格段是{best_name}，总月销量{best_sales}；新品切入时应结合该价格段的销量、评分门槛和产品差异化承接判断。"


def set_value(ws, row: int, col: int, value: Any):
    cell = ws.cell(row, col)
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            cell = ws.cell(merged_range.min_row, merged_range.min_col)
            break
    cell.value = value
    if isinstance(value, int):
        cell.number_format = "0"
    elif isinstance(value, float):
        cell.number_format = "0.0"


def ensure_rows_before_merged_title(ws, title_row: int, required_data_rows: int, available_data_rows: int):
    rows_to_insert = required_data_rows - available_data_rows
    if rows_to_insert <= 0:
        return 0
    template_row = max(1, title_row - 1)
    template_height = ws.row_dimensions[template_row].height
    template_dimension = ws.row_dimensions[template_row]
    template_styles = [copy(ws.cell(template_row, column)._style) for column in range(1, ws.max_column + 1)]

    affected_ranges = [
        merged_range
        for merged_range in list(ws.merged_cells.ranges)
        if merged_range.min_row >= title_row
    ]
    affected_cells = []
    for merged_range in affected_ranges:
        cell = ws.cell(merged_range.min_row, merged_range.min_col)
        affected_cells.append({
            "range": merged_range,
            "value": cell.value,
            "style": copy(cell._style),
            "number_format": cell.number_format,
            "alignment": copy(cell.alignment),
        })
        ws.unmerge_cells(str(merged_range))

    ws.insert_rows(title_row, rows_to_insert)
    for row in range(title_row, title_row + rows_to_insert):
        ws.row_dimensions[row].height = template_height
        ws.row_dimensions[row].hidden = template_dimension.hidden
        ws.row_dimensions[row].outlineLevel = template_dimension.outlineLevel
        ws.row_dimensions[row].collapsed = template_dimension.collapsed
        for column, style in enumerate(template_styles, 1):
            ws.cell(row, column)._style = copy(style)

    for item in affected_cells:
        merged_range = item["range"]
        min_row = merged_range.min_row + rows_to_insert
        max_row = merged_range.max_row + rows_to_insert
        range_string = f"{get_column_letter(merged_range.min_col)}{min_row}:{get_column_letter(merged_range.max_col)}{max_row}"
        ws.merge_cells(range_string)
        cell = ws.cell(min_row, merged_range.min_col)
        cell.value = item["value"]
        cell._style = item["style"]
        cell.number_format = item["number_format"]
        cell.alignment = item["alignment"]

    return rows_to_insert


def find_row_with_value(ws, value: Any, start_row: int, end_row: int, col: int = 1) -> Optional[int]:
    for row in range(start_row, end_row + 1):
        if ws.cell(row, col).value == value:
            return row
    return None


def clear_table_values(ws, first_row: int, last_row: int, first_column: int = 1, last_column: int = 7) -> None:
    """清除模板动态表格的旧内容，但保留既有样式、边框和行高。"""
    for row in range(first_row, last_row + 1):
        for column in range(first_column, last_column + 1):
            cell = ws.cell(row, column)
            if isinstance(cell, MergedCell):
                continue
            cell.value = None
            cell.hyperlink = None


def create_merged_market_analysis_sheet(wb, metrics: Dict[str, Any]):
    sheet_name = market_analysis_sheet_name(wb)
    ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.create_sheet(sheet_name, 1)
    symbol = metrics.get("currency_symbol", "$")
    for chart in list(getattr(ws, "_charts", [])):
        ws._charts.remove(chart)

    if ws["A1"].value is None:
        ws["A1"] = "市场数据分析"

    # 只填模板预留的数据位，不重设行高、列宽、合并单元格、颜色、字体。
    label_values = {
        "C3": header_label("平均价格", symbol),
        "E3": header_label("中位价格", symbol),
        "A6": header_label("总销售额", symbol),
        "C6": header_label("平均销售额", symbol),
        "E6": header_label("中位销售额", symbol),
        "F14": header_label("平均价格", symbol),
        "F26": header_label("平均价格", symbol),
        "F40": header_label("中位价格", symbol),
    }
    for coordinate, value in label_values.items():
        cell = ws[coordinate]
        set_value(ws, cell.row, cell.column, value)

    overview_values = {
        "B3": whole(metrics["sample_count"]),
        "D3": whole(metrics["avg_price"]),
        "F3": whole(metrics["median_price"]),
        "B4": pct(metrics["top5_share"]),
        "D4": pct(metrics["top10_share"]),
        "F4": pct(metrics["top20_share"]),
        "B5": whole(metrics["total_sales"]),
        "D5": whole(metrics["avg_sales"]),
        "F5": whole(metrics["median_sales"]),
        "B6": whole(metrics["total_revenue"]),
        "D6": whole(metrics["total_revenue"] / metrics["sample_count"] if metrics["sample_count"] else 0),
        "F6": whole(metrics["median_revenue"]),
    }
    for coordinate, value in overview_values.items():
        cell = ws[coordinate]
        set_value(ws, cell.row, cell.column, value)

    concentration = "高" if metrics["top10_share"] >= 0.55 else "中" if metrics["top10_share"] >= 0.35 else "低"
    opportunity = "谨慎进入" if concentration == "高" else "可重点评估"
    ws["A8"] = f"样本共 {metrics['sample_count']} 个链接，总月销量 {whole(metrics['total_sales']):,}，总月销售额 {money(metrics['total_revenue'], symbol)}，Top10销量占比 {pct(metrics['top10_share'])}，头部集中度为{concentration}，建议：{opportunity}。"
    ws["A20"] = build_sales_conclusion(metrics)
    ws["A34"] = build_review_conclusion(metrics)

    layer_start = 14
    clear_table_values(ws, layer_start + 1, layer_start + 4)
    for row_num, (name, data) in enumerate(metrics["sales_layers"].items(), layer_start + 1):
        values = [name, whole(data["sample_count"]), whole(data["total_sales"]), whole(data["avg_sales"]), whole(data["median_sales"]), whole(data["avg_price"]), whole(data["median_reviews"])]
        for col_num, value in enumerate(values, 1):
            set_value(ws, row_num, col_num, value)
    add_chart(ws, "", "A12", 3, layer_start, layer_start + len(metrics["sales_layers"]))

    review_start = 26
    review_order = ["≤30评", "31-50评", "51-100评", "101-300评", "301-1000评", "1000评+", "未知评分数"]
    review_data_end = review_start + len(review_order)
    if "未知评分数" in metrics["review_buckets"]:
        # 模板把最后一行做成了合并占位行；有未知评分样本时展开并沿用上一行样式。
        merged_review_row = f"A{review_data_end}:G{review_data_end}"
        if merged_review_row in {str(item) for item in ws.merged_cells.ranges}:
            ws.unmerge_cells(merged_review_row)
            for column in range(1, 8):
                ws.cell(review_data_end, column)._style = copy(ws.cell(review_data_end - 1, column)._style)
    clear_table_values(ws, review_start + 1, review_start + len(review_order))
    row_num = review_start + 1
    for name in review_order:
        if row_num > review_start + 6 and name not in metrics["review_buckets"]:
            continue
        data = metrics["review_buckets"].get(name, {})
        values = [
            name,
            whole(data.get("sample_count", 0)),
            whole(data.get("total_sales", 0)),
            whole(data.get("avg_sales", 0)),
            whole(data.get("median_sales", 0)),
            whole(data.get("avg_price", 0)),
            whole(data.get("median_reviews", 0)),
        ]
        for col_num, value in enumerate(values, 1):
            set_value(ws, row_num, col_num, value)
        row_num += 1
    add_chart(ws, "", "A24", 3, review_start, row_num - 1)

    price_start = 40
    price_order = metrics.get("price_bucket_order", price_bucket_order(symbol))
    price_title_row = find_row_with_value(ws, "3.价格分析", 46, 55) or 46
    available_price_rows = max(0, price_title_row - price_start - 1)
    inserted_price_rows = ensure_rows_before_merged_title(ws, price_title_row, len(price_order), available_price_rows)
    current_price_title_row = price_title_row + inserted_price_rows
    clear_table_values(ws, price_start + 1, current_price_title_row - 1)
    price_conclusion_row = price_title_row + inserted_price_rows + 1
    ws.cell(price_conclusion_row, 1).value = build_price_conclusion(metrics)

    row_num = price_start + 1
    for name in price_order:
        data = metrics["price_buckets"].get(name, {})
        values = [name, whole(data.get("sample_count", 0)), whole(data.get("total_sales", 0)), whole(data.get("avg_sales", 0)), whole(data.get("median_sales", 0)), whole(data.get("median_price", 0)), whole(data.get("median_reviews", 0))]
        for col_num, value in enumerate(values, 1):
            set_value(ws, row_num, col_num, value)
        row_num += 1
    add_chart(ws, "", "A38", 3, price_start, row_num - 1)
    return ws


def apply_borders(wb):
    return None


def clean_old_report_sheets(wb):
    for name in list(wb.sheetnames):
        if name in REPORT_SHEETS:
            del wb[name]


def write_report_sheets(wb, metrics: Dict[str, Any]):
    clean_old_report_sheets(wb)
    create_merged_market_analysis_sheet(wb, metrics)
    create_anomaly_sheet(wb, metrics)
    apply_borders(wb)


def generate_report(
    input_path: str | Path,
    output_path: str | Path,
    include_source_slicers: bool = False,
    product_name: str | None = None,
) -> Path:
    input_path = Path(input_path)
    output_path = Path(output_path)
    wb = load_workbook(input_path, rich_text=True)
    canonicalize_product_sheet_names(wb, product_name)
    source_ws = find_source_sheet(wb)
    if source_ws is None:
        raise ValueError("找不到数据源 sheet：需要 产品名-数据源、数据源、US 或 Sheet1")
    source_ws = ensure_canonical_source_sheet_name(wb, source_ws)
    rows = read_rows(source_ws)
    metrics = compute_metrics(rows)
    if include_source_slicers:
        create_source_slicer_sheet(wb, source_ws, rows, metrics)
    write_report_sheets(wb, metrics)
    if source_ws.title != SOURCE_SHEET and SOURCE_SHEET in wb.sheetnames and len(wb.worksheets) > 1:
        del wb[SOURCE_SHEET]
    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="生成亚马逊市场数据分析报表")
    parser.add_argument("input")
    parser.add_argument("output")
    parser.add_argument("--source-slicers", action="store_true", help="在数据源增加分层辅助列，供后续创建切片器")
    parser.add_argument("--product-name", help="本次产品名称，用于替换复用模板中的旧 Sheet 前缀")
    args = parser.parse_args()
    out = generate_report(
        args.input,
        args.output,
        include_source_slicers=args.source_slicers,
        product_name=args.product_name,
    )
    print(out)
