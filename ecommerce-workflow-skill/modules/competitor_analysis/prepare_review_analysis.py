from __future__ import annotations

import argparse
import hashlib
import json
import random
import re
from collections import Counter
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


ASIN_PATTERN = re.compile(r"(?<![A-Z0-9])([A-Z0-9]{10})(?![A-Z0-9])", re.I)
POSITIVE_REVIEWS = "Positive reviews"
CRITICAL_REVIEWS = "Critical reviews"
SCHEMA_VERSION = "competitor-analysis.v1"
ALLOWED_REVIEW_STATUSES = {"success", "insufficient_reviews", "no_reviews"}


def normalize_text(value: Any) -> str:
    return str(value or "").strip()


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def machine_payload(analysis: dict[str, Any]) -> dict[str, Any]:
    return {
        "schema_version": analysis.get("schema_version"),
        "source_workbook": analysis.get("source_workbook"),
        "source_digest": analysis.get("source_digest"),
        "sample_size": analysis.get("sample_size"),
        "products": [
            {key: value for key, value in product.items() if key != "advantages"}
            for product in analysis.get("products", [])
        ],
    }


def machine_digest(analysis: dict[str, Any]) -> str:
    encoded = json.dumps(
        machine_payload(analysis),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def extract_asin(value: Any) -> str | None:
    match = ASIN_PATTERN.search(normalize_text(value).upper())
    return match.group(1) if match else None


def find_sheet(workbook, exact_name: str, suffix: str | None = None):
    if exact_name in workbook.sheetnames:
        return workbook[exact_name]
    if suffix:
        for worksheet in workbook.worksheets:
            if worksheet.title.endswith(suffix):
                return worksheet
    raise ValueError(f"工作簿缺少“{exact_name}”工作表。")


def header_columns(worksheet) -> dict[str, int]:
    return {
        normalize_text(cell.value): cell.column
        for cell in worksheet[1]
        if normalize_text(cell.value)
    }


def header_column(headers: dict[str, int], *names: str) -> int | None:
    return next((headers[name] for name in names if name in headers), None)


def find_label_row(worksheet, label: str, columns: tuple[int, ...] = (1, 2)) -> int:
    for row in range(1, worksheet.max_row + 1):
        if any(normalize_text(worksheet.cell(row, column).value) == label for column in columns):
            return row
    raise ValueError(f"工作表“{worksheet.title}”中找不到“{label}”行。")


def competitor_columns(worksheet) -> list[dict[str, Any]]:
    link_row = find_label_row(worksheet, "详情页链接")
    products: list[dict[str, Any]] = []
    for column in range(3, worksheet.max_column + 1):
        cell = worksheet.cell(link_row, column)
        target = cell.hyperlink.target if cell.hyperlink else cell.value
        asin = extract_asin(target) or extract_asin(cell.value)
        if asin:
            products.append({"asin": asin, "column": cell.column_letter})
    if not products:
        title_asins = []
        for row in range(1, min(worksheet.max_row, 2) + 1):
            for cell in worksheet[row]:
                title_asins.extend(ASIN_PATTERN.findall(normalize_text(cell.value).upper()))
        for index, asin in enumerate(dict.fromkeys(title_asins)):
            products.append({"asin": asin, "column": get_column_letter(index + 3)})
    if not products:
        raise ValueError("竞品分析页的“详情页链接”行和标题行中都没有找到 ASIN。")
    return products


def competitor_context(worksheet, column: str) -> dict[str, Any]:
    rating_row = find_label_row(worksheet, "评分/评论数")
    selling_points_row = find_label_row(worksheet, "产品卖点")
    comparison_start_row = find_label_row(worksheet, "对比维度")
    evaluation_row = find_label_row(worksheet, "评价分析")
    comparisons = []
    for row in range(comparison_start_row, evaluation_row):
        dimension = normalize_text(worksheet.cell(row, 2).value)
        value = normalize_text(worksheet[f"{column}{row}"].value)
        if dimension or value:
            comparisons.append({"dimension": dimension, "value": value})
    return {
        "rating_and_review_count": normalize_text(worksheet[f"{column}{rating_row}"].value),
        "selling_points": normalize_text(worksheet[f"{column}{selling_points_row}"].value),
        "comparison_dimensions": comparisons,
    }


def summary_groups(summary, details) -> list[dict[str, Any]]:
    summary_headers = header_columns(summary)
    detail_headers = header_columns(details)
    required_summary = {"评论类型", "实际筛选", "抓取状态", "抓取评论数"}
    required_detail = {"评论类型", "评论标题", "评论正文", "购买变体", "评论链接"}
    if missing := required_summary - summary_headers.keys():
        raise ValueError(f"评论抓取摘要缺少列：{', '.join(sorted(missing))}")
    if missing := required_detail - detail_headers.keys():
        raise ValueError(f"商品评论缺少列：{', '.join(sorted(missing))}")
    summary_asin_column = header_column(summary_headers, "竞品ASIN", "ASIN")
    detail_asin_column = header_column(detail_headers, "竞品ASIN", "ASIN")
    if summary_asin_column is None:
        raise ValueError("评论抓取摘要缺少“竞品ASIN”列。")
    if detail_asin_column is None:
        raise ValueError("商品评论缺少“竞品ASIN”列。")

    groups: list[dict[str, Any]] = []
    for row in range(2, summary.max_row + 1):
        review_type = normalize_text(summary.cell(row, summary_headers["评论类型"]).value)
        if not review_type:
            break
        asin = extract_asin(summary.cell(row, summary_asin_column).value)
        if not asin:
            raise ValueError(f"评论抓取摘要第 {row} 行的 ASIN 无效。")
        count = int(summary.cell(row, summary_headers["抓取评论数"]).value or 0)
        status = normalize_text(summary.cell(row, summary_headers["抓取状态"]).value)
        if status not in ALLOWED_REVIEW_STATUSES:
            raise ValueError(
                f"评论抓取摘要第 {row} 行状态为“{status or '/'}”，"
                "不能用于竞品分析；请先完成或修复评论抓取。"
            )
        if status == "no_reviews" and count:
            raise ValueError(f"评论抓取摘要第 {row} 行状态为 no_reviews，但抓取评论数为 {count}。")
        review_source_column = summary_headers.get("评论来源ASIN")
        review_source_asin = (
            extract_asin(summary.cell(row, review_source_column).value)
            if review_source_column is not None
            else asin
        )
        actual_filter = normalize_text(summary.cell(row, summary_headers["实际筛选"]).value)
        if not actual_filter:
            raise ValueError(f"评论抓取摘要第 {row} 行缺少实际筛选。")
        groups.append(
            {
                "asin": asin,
                "review_source_asin": review_source_asin,
                "actual_filter": actual_filter,
                "review_type": review_type,
                "status": status,
                "count": count,
                "reviews": [],
            }
        )

    detail_row = 2
    for group in groups:
        for _ in range(group["count"]):
            if detail_row > details.max_row:
                raise ValueError("商品评论行数少于评论抓取摘要记录的数量。")
            actual_type = normalize_text(details.cell(detail_row, detail_headers["评论类型"]).value)
            actual_asin = extract_asin(details.cell(detail_row, detail_asin_column).value)
            if actual_asin != group["asin"]:
                raise ValueError(
                    f"商品评论第 {detail_row} 行 ASIN 为“{actual_asin or '/'}”，"
                    f"与摘要中的“{group['asin']}”不一致。"
                )
            if actual_type != group["review_type"]:
                raise ValueError(
                    f"商品评论第 {detail_row} 行类型为“{actual_type}”，"
                    f"与摘要中的“{group['review_type']}”不一致。"
                )
            group["reviews"].append(
                {
                    "title": normalize_text(details.cell(detail_row, detail_headers["评论标题"]).value),
                    "body": normalize_text(details.cell(detail_row, detail_headers["评论正文"]).value),
                    "variant": normalize_text(details.cell(detail_row, detail_headers["购买变体"]).value),
                    "review_url": normalize_text(details.cell(detail_row, detail_headers["评论链接"]).value),
                }
            )
            detail_row += 1
    return groups


def assign_critical_pain_rows(groups: list[dict[str, Any]], pain_sheet) -> None:
    headers = header_columns(pain_sheet)
    required = {"归一化痛点（中文）"}
    if missing := required - headers.keys():
        raise ValueError(f"差评痛点页缺少列：{', '.join(sorted(missing))}")
    asin_column = header_column(headers, "竞品ASIN", "ASIN")
    if asin_column is None:
        raise ValueError("差评痛点页缺少“竞品ASIN”列。")
    rows_by_asin: dict[str, list[list[str]]] = {}
    for pain_row in range(2, pain_sheet.max_row + 1):
        asin = extract_asin(pain_sheet.cell(pain_row, asin_column).value)
        if not asin:
            continue
        raw = normalize_text(pain_sheet.cell(pain_row, headers["归一化痛点（中文）"]).value)
        labels = []
        for item in raw.splitlines():
            label = item.strip()
            if label and label not in {"/", "待归类"} and label not in labels:
                labels.append(label)
        rows_by_asin.setdefault(asin, []).append(labels)

    for group in groups:
        if group["review_type"] != CRITICAL_REVIEWS:
            continue
        pain_rows = rows_by_asin.get(group["asin"], [])
        if len(pain_rows) < group["count"]:
            raise ValueError(
                f"差评痛点中 ASIN {group['asin']} 的行数少于摘要记录的差评数量。"
            )
        group["pain_rows"] = pain_rows[:group["count"]]


def random_sample(reviews: list[dict[str, str]], asin: str, sample_size: int) -> list[dict[str, str]]:
    if len(reviews) <= sample_size:
        return reviews
    fingerprint = "\n".join(f"{item['title']}\n{item['body']}" for item in reviews)
    seed = int.from_bytes(hashlib.sha256(f"{asin}\n{fingerprint}".encode("utf-8")).digest()[:8], "big")
    return random.Random(seed).sample(reviews, sample_size)


def percent_text(ratio: float) -> str:
    percent = ratio * 100
    if abs(percent - round(percent)) < 1e-9:
        return f"{int(round(percent))}%"
    return f"{percent:.1f}%"


def disadvantages_for_group(group: dict[str, Any]) -> list[dict[str, Any]]:
    pain_rows = group.get("pain_rows", [])
    if not pain_rows:
        return []
    counts = Counter(label for labels in pain_rows for label in set(labels))
    total = len(pain_rows)
    return [
        {
            "label": label,
            "review_count": count,
            "coverage_rate": count / total,
            "display": f"{label}（{percent_text(count / total)}）",
        }
        for label, count in sorted(counts.items(), key=lambda item: (-item[1], item[0]))
    ]


def prepare_review_analysis(workbook_path: Path, output: Path, sample_size: int = 25) -> dict[str, Any]:
    if sample_size < 1:
        raise ValueError("sample_size 必须大于 0。")
    workbook = load_workbook(workbook_path, data_only=False)
    try:
        competitor_sheet = find_sheet(workbook, "竞品分析", "竞品分析")
        summary = find_sheet(workbook, "评论抓取摘要")
        details = find_sheet(workbook, "商品评论")
        pain_sheet = find_sheet(workbook, "差评痛点")
        products = competitor_columns(competitor_sheet)
        groups = summary_groups(summary, details)
        assign_critical_pain_rows(groups, pain_sheet)

        by_asin: dict[str, list[dict[str, Any]]] = {}
        for group in groups:
            by_asin.setdefault(group["asin"], []).append(group)

        report_products = []
        for product in products:
            asin = product["asin"]
            context = competitor_context(competitor_sheet, product["column"])
            product_groups = by_asin.get(asin, [])
            grouped_types = [group["review_type"] for group in product_groups]
            for required_type in (POSITIVE_REVIEWS, CRITICAL_REVIEWS):
                if grouped_types.count(required_type) != 1:
                    raise ValueError(
                        f"竞品 ASIN {asin} 必须且只能有一行 {required_type} 摘要，"
                        f"当前为 {grouped_types.count(required_type)} 行。"
                    )
            positive = next(
                (group for group in product_groups if group["review_type"] == POSITIVE_REVIEWS),
                None,
            )
            critical = next(
                (group for group in product_groups if group["review_type"] == CRITICAL_REVIEWS),
                None,
            )
            positive_reviews = positive["reviews"] if positive else []
            report_products.append(
                {
                    "asin": asin,
                    "column": product["column"],
                    **context,
                    "has_review_data": bool(product_groups),
                    "positive_review_count": len(positive_reviews),
                    "positive_sample": random_sample(positive_reviews, asin, sample_size),
                    "advantages": [],
                    "critical_review_count": critical["count"] if critical else 0,
                    "disadvantages": disadvantages_for_group(critical) if critical else [],
                }
            )

        report = {
            "schema_version": SCHEMA_VERSION,
            "source_workbook": str(workbook_path.resolve()),
            "source_digest": file_sha256(workbook_path),
            "sample_size": sample_size,
            "products": report_products,
            "overall_summary": "",
            "improvements": [],
        }
        report["machine_digest"] = machine_digest(report)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
        return report
    finally:
        workbook.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="为竞品评价分析抽取每个 ASIN 的好评样本和差评覆盖率。")
    parser.add_argument("--workbook", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--sample-size", type=int, default=25)
    args = parser.parse_args()
    report = prepare_review_analysis(args.workbook, args.output, args.sample_size)
    for product in report["products"]:
        print(
            f"{product['asin']}: positive_sample={len(product['positive_sample'])}, "
            f"critical={product['critical_review_count']}, disadvantages={len(product['disadvantages'])}"
        )
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
