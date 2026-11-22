from __future__ import annotations

import json
import tempfile
import unittest
from datetime import date
from pathlib import Path
from unittest.mock import patch
from urllib.parse import parse_qs, urlparse

from openpyxl import Workbook, load_workbook
from PIL import Image as PillowImage

from review_research.scrape_amazon_reviews import (
    AsinResult,
    DETAIL_SHEET,
    PAIN_FREQUENCY_SHEET,
    PAIN_POINT_SHEET,
    SUMMARY_SHEET,
    build_review_url,
    collect_reviews_for_asin,
    extract_asins_from_workbook,
    extract_parent_asin_from_product_html,
    extract_parent_asins_from_workbook,
    normalize_review,
    parse_helpful_count,
    parse_rating,
    parse_review_date,
    primary_review_source,
    product_title_from_page_title,
    review_limit_for_filter,
    review_is_trusted,
    reviewer_filter_label,
    original_review_image_url,
    write_review_sheets,
)
from review_research.backfill_review_translations import refresh_workbook
from review_research.backfill_pain_phrase_groups import refresh_workbook as refresh_pain_groups
from review_research.pain_phrase_analysis import (
    SemanticGroup,
    aggregate_semantic_groups,
    analyze_pain_phrases,
)
from review_research.run import calculate_review_command_timeout


class ReviewResearchTests(unittest.TestCase):
    def test_candidate_limit_stops_before_loading_more_reviews(self):
        class LocatorThatMustNotBeUsed:
            def count(self):
                raise AssertionError("达到候选评论上限后不应继续查找或点击 Show 10 more reviews")

        class FakePage:
            url = "https://www.amazon.com/product-reviews/B0ABCDEFGH/"

            def __init__(self):
                self.evaluate_count = 0

            def goto(self, *_args, **_kwargs):
                return None

            def wait_for_function(self, *_args, **_kwargs):
                return None

            def evaluate(self, _script):
                self.evaluate_count += 1
                return {
                    "url": self.url,
                    "reviews": [
                        {
                            "reviewId": "R1",
                            "ratingText": "5.0 out of 5 stars",
                            "title": "普通评论一",
                            "body": "普通评论一",
                        },
                        {
                            "reviewId": "R2",
                            "ratingText": "5.0 out of 5 stars",
                            "title": "可信评论",
                            "body": "可信评论",
                            "verifiedText": "Verified Purchase",
                        },
                        {
                            "reviewId": "R3",
                            "ratingText": "4.0 out of 5 stars",
                            "title": "普通评论二",
                            "body": "普通评论二",
                        },
                    ],
                    "filters": {
                        "sortBy": "helpful",
                        "starFilter": "positive",
                    },
                }

            def locator(self, _selector):
                return LocatorThatMustNotBeUsed()

        page = FakePage()
        with patch("review_research.scrape_amazon_reviews.preflight_amazon"):
            result = collect_reviews_for_asin(
                page,
                asin="B0ABCDEFGH",
                site="US",
                postal_code=None,
                max_pages=20,
                max_reviews=3,
                sort_by="helpful",
                reviewer_type="default",
                trusted_only=True,
                star_filter="positive",
            )

        self.assertEqual(page.evaluate_count, 1)
        self.assertEqual([review["review_id"] for review in result.reviews], ["R2"])
        self.assertEqual(result.status, "success")
        self.assertEqual(result.error, "")

    def test_completed_batch_preflight_is_not_repeated(self):
        class EmptyLocator:
            def count(self):
                return 0

        class FakePage:
            url = "https://www.amazon.com/product-reviews/B0ABCDEFGH/"

            def goto(self, *_args, **_kwargs):
                return None

            def wait_for_function(self, *_args, **_kwargs):
                return None

            def evaluate(self, _script):
                return {
                    "url": self.url,
                    "reviews": [
                        {
                            "reviewId": "R1",
                            "ratingText": "1.0 out of 5 stars",
                            "title": "可信差评",
                            "body": "可信差评",
                            "verifiedText": "Verified Purchase",
                        }
                    ],
                    "filters": {
                        "sortBy": "helpful",
                        "starFilter": "critical",
                    },
                }

            def locator(self, _selector):
                return EmptyLocator()

        with patch("review_research.scrape_amazon_reviews.preflight_amazon") as preflight:
            collect_reviews_for_asin(
                FakePage(),
                asin="B0ABCDEFGH",
                site="US",
                postal_code=None,
                max_pages=10,
                max_reviews=100,
                sort_by="helpful",
                reviewer_type="default",
                trusted_only=True,
                star_filter="critical",
                preflight_complete=True,
            )

        preflight.assert_not_called()

    def test_uses_separate_default_review_limits(self):
        self.assertEqual(review_limit_for_filter("positive"), 50)
        self.assertEqual(review_limit_for_filter("five_star"), 50)
        self.assertEqual(review_limit_for_filter("critical"), 100)
        self.assertEqual(review_limit_for_filter("one_star"), 100)
        self.assertEqual(review_limit_for_filter("positive", max_reviews=30), 30)
        self.assertEqual(review_limit_for_filter("critical", max_reviews=30), 30)

    def test_review_command_timeout_covers_all_asin_filter_jobs(self):
        self.assertEqual(calculate_review_command_timeout(5, "both"), 6_600)
        self.assertEqual(calculate_review_command_timeout(3, "critical"), 2_400)

    def test_review_command_timeout_keeps_existing_minimum(self):
        self.assertEqual(calculate_review_command_timeout(1, "both"), 1_800)
        self.assertEqual(calculate_review_command_timeout(0, "both"), 1_800)

    def test_extracts_and_counts_negative_phrases(self):
        reviews = [
            {"title": "Not durable", "body": "It stopped working after one week."},
            {"title": "Waste money", "body": "Stopped working after the second use."},
        ]
        terms_by_review, frequency_rows = analyze_pain_phrases(reviews)
        self.assertIn("not durable", terms_by_review[0])
        self.assertIn("waste money", terms_by_review[1])
        stopped = next(item for item in frequency_rows if item.phrase == "stopped working")
        self.assertEqual(stopped.review_count, 2)
        self.assertEqual(stopped.review_ratio, 1)

    def test_semantic_groups_use_review_union_and_deduplicate_variants(self):
        reviews = [
            {"variant": "Model A"},
            {"variant": "Model A"},
            {"variant": "Model B"},
        ]
        terms_by_review = [
            ["sticky surface", "sticky touch", "stopped working", "not sure it's worth it"],
            ["sticky feeling", "never worked"],
            ["washing machine broke"],
        ]
        rows, missing = aggregate_semantic_groups(reviews, terms_by_review)

        self.assertEqual(missing, [])
        stickiness = next(item for item in rows if item.phrase == "surface stickiness")
        self.assertEqual(stickiness.translation, "表面黏腻")
        self.assertEqual(stickiness.review_count, 2)
        self.assertEqual(stickiness.review_ratio, 2 / 3)
        self.assertEqual(stickiness.variants, ("Model A",))
        failure = next(item for item in rows if item.phrase == "product failure")
        self.assertEqual(failure.review_count, 3)
        self.assertEqual(failure.variants, ("Model A", "Model B"))
        value = next(item for item in rows if item.phrase == "poor value for money")
        self.assertEqual(value.review_count, 1)

    def test_extracts_unique_asins_from_data_source(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "source.xlsx"
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "demo-数据源"
            worksheet.append(["商品标题", "ASIN", "父ASIN", "月销量"])
            worksheet.append(["A", "B0ABCDEFGH", "B0PARENT01", 10])
            worksheet.append(["B", "https://www.amazon.com/dp/B0JKLMN0PQ", "B0PARENT02", 20])
            worksheet.append(["duplicate", "B0ABCDEFGH", "B0PARENT03", 30])
            worksheet.append(["invalid", "B0TOO-LONG1", "B0PARENT04", 40])
            workbook.save(path)
            workbook.close()

            self.assertEqual(extract_asins_from_workbook(path), ["B0ABCDEFGH", "B0JKLMN0PQ"])
            self.assertEqual(
                extract_parent_asins_from_workbook(path),
                {
                    "B0ABCDEFGH": "B0PARENT01",
                    "B0JKLMN0PQ": "B0PARENT02",
                },
            )

    def test_normalizes_localized_review_fields(self):
        review = normalize_review(
            {
                "reviewId": "R123",
                "reviewerName": "Ada",
                "ratingText": "4,0 von 5 Sternen",
                "title": "Gut",
                "body": "Solide.",
                "dateText": "Rezension aus Deutschland vom 3. März 2026",
                "variant": "Farbe: Blau",
                "verifiedText": "Verifizierter Kauf",
                "vineText": "Vine Customer Review of Free Product",
                "helpfulText": "Eine Person fand diese Informationen hilfreich",
                "imageUrls": ["https://example.test/a.jpg", "https://example.test/a.jpg"],
                "hasVideo": False,
                "reviewUrl": "https://example.test/review/R123",
            },
            asin="B0ABCDEFGH",
            page_number=2,
            source_url="https://example.test/page/2",
        )
        self.assertEqual(review["rating"], 4.0)
        self.assertEqual(review["date"], date(2026, 3, 3))
        self.assertEqual(review["helpful_count"], 1)
        self.assertTrue(review["verified"])
        self.assertTrue(review["vine"])
        self.assertEqual(len(review["image_urls"]), 1)

    def test_parsers_and_review_url(self):
        self.assertEqual(parse_rating("5.0 out of 5 stars"), 5.0)
        self.assertEqual(parse_helpful_count("1,234 people found this helpful"), 1234)
        self.assertEqual(parse_review_date("Reviewed in the United Kingdom on 28 August 2026"), date(2026, 8, 28))
        url = build_review_url(
            "https://www.amazon.com",
            "B0ABCDEFGH",
            sort_by="helpful",
            reviewer_type="verified",
            star_filter="positive",
        )
        self.assertNotIn("pageNumber", url)
        self.assertIn("sortBy=helpful", url)
        self.assertIn("reviewerType=avp_only_reviews", url)
        self.assertIn("filterByStar=positive", url)
        default_url = build_review_url(
            "https://www.amazon.com",
            "B0ABCDEFGH",
            sort_by="helpful",
            reviewer_type="default",
            star_filter="critical",
        )
        self.assertIn("sortBy=helpful", default_url)
        self.assertIn("filterByStar=critical", default_url)
        self.assertNotIn("reviewerType", default_url)
        portal_url = build_review_url(
            "https://www.amazon.com",
            "B0PARENT01",
            sort_by="helpful",
            reviewer_type="all",
            star_filter="critical",
            endpoint="portal/customer-reviews",
        )
        portal_query = parse_qs(urlparse(portal_url).query)
        self.assertEqual(portal_query["ie"], ["UTF8"])
        self.assertEqual(portal_query["filterByStar"], ["critical"])
        self.assertEqual(portal_query["reviewerType"], ["all_reviews"])
        self.assertEqual(portal_query["sortBy"], ["helpful"])
        self.assertEqual(
            product_title_from_page_title("Amazon.com: Customer reviews: Demo Product"),
            "Demo Product",
        )
        self.assertEqual(
            primary_review_source("B0CHILD001", {"B0CHILD001": "B0PARENT01"})[:2],
            ("B0CHILD001", "product-reviews"),
        )
        self.assertEqual(
            primary_review_source("B0CHILD001", {"B0CHILD001": "B0PARENT01"})[2],
            "",
        )
        self.assertEqual(
            primary_review_source("B0CHILD001", {})[:2],
            ("B0CHILD001", "product-reviews"),
        )
        self.assertEqual(
            extract_parent_asin_from_product_html(
                '<script>{"parentAsin": "B0PARENT01"}</script>',
                "B0CHILD001",
            ),
            "B0PARENT01",
        )
        self.assertEqual(
            extract_parent_asin_from_product_html(
                'parentAsin=B0PARENT02&landingAsin=B0CHILD001',
                "B0CHILD001",
            ),
            "B0PARENT02",
        )
        self.assertEqual(
            original_review_image_url("https://m.media-amazon.com/images/I/demo._SY88_.jpg"),
            "https://m.media-amazon.com/images/I/demo.jpg",
        )
        self.assertTrue(review_is_trusted({"verified": True, "vine": False}))
        self.assertTrue(review_is_trusted({"verified": False, "vine": True}))
        self.assertFalse(review_is_trusted({"verified": False, "vine": False}))
        self.assertEqual(
            reviewer_filter_label("default", True),
            "Top reviews；保留 Verified Purchase 或 Vine 绿标",
        )

    def test_writes_summary_and_detail_sheets(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "output.xlsx"
            workbook = Workbook()
            workbook.active.title = "数据源"
            workbook.save(source)
            workbook.close()
            review_image = Path(temp_dir) / "review.png"
            PillowImage.new("RGB", (40, 30), "red").save(review_image)
            second_review_image = Path(temp_dir) / "review-2.png"
            PillowImage.new("RGB", (30, 40), "blue").save(second_review_image)
            review = normalize_review(
                {
                    "reviewId": "R1",
                    "ratingText": "5.0 out of 5 stars",
                    "title": "Works",
                    "body": "Good product",
                    "dateText": "Reviewed in the United States on August 20, 2026",
                    "verifiedText": "Verified Purchase",
                    "vineText": "Vine Customer Review of Free Product",
                    "helpfulText": "2 people found this helpful",
                    "imageUrls": [review_image.as_uri(), second_review_image.as_uri()],
                    "reviewUrl": "https://example.test/review/R1",
                },
                asin="B0ABCDEFGH",
                page_number=1,
                source_url="https://example.test/page/1",
            )
            result = AsinResult(
                asin="B0ABCDEFGH",
                star_filter="positive",
                reviewer_type="verified",
                sort_by="helpful",
                product_title="Demo",
                product_url="https://www.amazon.com/dp/B0ABCDEFGH",
                first_review_url="https://www.amazon.com/product-reviews/B0ABCDEFGH/",
                requested_pages=1,
                scanned_pages=1,
                reviews=[review],
            )
            write_review_sheets(source, output, [result], site="US")

            checked = load_workbook(output, data_only=True)
            try:
                self.assertIn(SUMMARY_SHEET, checked.sheetnames)
                self.assertIn(DETAIL_SHEET, checked.sheetnames)
                self.assertIn(PAIN_POINT_SHEET, checked.sheetnames)
                self.assertNotIn(PAIN_FREQUENCY_SHEET, checked.sheetnames)
                self.assertEqual(checked[SUMMARY_SHEET]["C2"].value, "B0ABCDEFGH")
                self.assertEqual(checked[SUMMARY_SHEET]["D2"].value, "Positive reviews")
                self.assertEqual(checked[SUMMARY_SHEET]["E2"].value, "Verified Purchase only")
                self.assertEqual(checked[SUMMARY_SHEET]["G2"].value, 1)
                self.assertEqual(checked[SUMMARY_SHEET]["I2"].value, 1)
                self.assertEqual(checked[SUMMARY_SHEET].column_dimensions["F"].width, 22)
                details = checked[DETAIL_SHEET]
                self.assertEqual(
                    [cell.value for cell in details[1]],
                    [
                        "站点", "竞品ASIN", "实际筛选", "评论日期", "星级",
                        "评论类型", "评论标题", "评论正文", "购买变体",
                        "Verified Purchase", "绿标评论", "Helpful数", "图片URL", "含视频", "评论链接",
                    ],
                )
                self.assertEqual(details["B2"].value, "B0ABCDEFGH")
                self.assertEqual(details["C2"].value, "Verified Purchase only")
                self.assertEqual(details["F2"].value, "Positive reviews")
                self.assertEqual(details["H2"].value, "Good product")
                self.assertEqual(details["K2"].value, "是")
            finally:
                checked.close()

            critical = AsinResult(
                asin="B0ABCDEFGH",
                star_filter="critical",
                reviewer_type="verified",
                sort_by="helpful",
                product_title="Demo",
                product_url="https://www.amazon.com/dp/B0ABCDEFGH",
                reviews=[review],
            )
            translation_report = Path(temp_dir) / "missing_review_translations.json"
            pain_group_report = Path(temp_dir) / "missing_pain_phrase_groups.json"
            write_review_sheets(
                source,
                output,
                [critical],
                site="US",
                translation_cache=Path(temp_dir) / "empty_cache.json",
                translation_report=translation_report,
                pain_group_cache=Path(temp_dir) / "empty_groups.json",
                pain_group_report=pain_group_report,
            )
            checked = load_workbook(output, data_only=True)
            try:
                pain_points = checked[PAIN_POINT_SHEET]
                self.assertEqual(
                    [cell.value for cell in pain_points[1]],
                    [
                        "竞品ASIN", "实际筛选", "星级", "评论日期",
                        "属性/型号", "差评词/短语", "归一化痛点（中文）", "Helpful数",
                        "图片数量", "评论图片", "评价内容", "评论链接",
                    ],
                )
                self.assertEqual(pain_points["A2"].value, "B0ABCDEFGH")
                self.assertEqual(pain_points["B2"].value, "Verified Purchase only")
                self.assertEqual(pain_points["E2"].value, "/")
                self.assertTrue(pain_points["F2"].value)
                self.assertEqual(pain_points["G2"].value, "待归类")
                self.assertIn("Good product", pain_points["K2"].value)
                self.assertEqual(pain_points["I2"].value, 2)
                self.assertIsNone(pain_points["J2"].value)
                self.assertIsNone(pain_points["J2"].hyperlink)
                self.assertIsNotNone(pain_points["L2"].hyperlink)
                self.assertEqual(pain_points["A1"].border.left.style, "thin")
                self.assertEqual(pain_points["A2"].border.bottom.style, "thin")
                self.assertEqual(len(pain_points._images), 1)
                frequency = checked[SUMMARY_SHEET]
                frequency_title_row = next(
                    row for row in range(1, frequency.max_row + 1)
                    if frequency.cell(row, 1).value == PAIN_FREQUENCY_SHEET
                )
                frequency_header_row = frequency_title_row + 1
                frequency_data_row = frequency_header_row + 1
                self.assertEqual(
                    [frequency.cell(frequency_header_row, column).value for column in range(1, 8)],
                    [
                        "竞品ASIN", "实际筛选", "关键词/短语",
                        "中文翻译", "评论覆盖数", "覆盖率", "涉及型号",
                    ],
                )
                self.assertEqual(frequency.cell(frequency_data_row, 1).value, "B0ABCDEFGH")
                self.assertEqual(frequency.cell(frequency_data_row, 2).value, "Verified Purchase only")
                self.assertEqual(frequency.cell(frequency_data_row, 5).value, 1)
                self.assertEqual(frequency.cell(frequency_data_row, 6).value, 1)
                self.assertEqual(frequency.cell(frequency_data_row, 7).value, "/")
            finally:
                checked.close()

            missing = json.loads(translation_report.read_text(encoding="utf-8"))
            content = "Works\nGood product"
            self.assertEqual(missing["missing_count"], 1)
            self.assertIn(content, missing["translations"])
            missing_groups = json.loads(pain_group_report.read_text(encoding="utf-8"))
            self.assertEqual(missing_groups["missing_count"], 1)
            self.assertIn("works", missing_groups["groups"])
            self.assertEqual(refresh_workbook(output, {content: "很好用", "works": "有效"}), 2)
            translated = load_workbook(output, data_only=True)
            try:
                self.assertEqual(translated[PAIN_POINT_SHEET]["K2"].value, f"{content}\n(很好用)")
                self.assertGreaterEqual(translated[PAIN_POINT_SHEET].row_dimensions[2].height, 54)
                self.assertEqual(len(translated[PAIN_POINT_SHEET]._images), 1)
                summary = translated[SUMMARY_SHEET]
                title_row = next(
                    row for row in range(1, summary.max_row + 1)
                    if summary.cell(row, 1).value == PAIN_FREQUENCY_SHEET
                )
                self.assertEqual(summary.cell(title_row + 2, 4).value, "有效")
            finally:
                translated.close()

            row_count, remaining = refresh_pain_groups(
                output,
                {"works": SemanticGroup("product effectiveness", "产品效果")},
            )
            self.assertEqual(row_count, 1)
            self.assertEqual(remaining, [])
            regrouped = load_workbook(output, data_only=True)
            try:
                summary = regrouped[SUMMARY_SHEET]
                title_row = next(
                    row for row in range(1, summary.max_row + 1)
                    if summary.cell(row, 1).value == PAIN_FREQUENCY_SHEET
                )
                data_row = title_row + 2
                self.assertEqual(summary.cell(data_row, 1).value, "B0ABCDEFGH")
                self.assertEqual(summary.cell(data_row, 2).value, "Verified Purchase only")
                self.assertEqual(summary.cell(data_row, 3).value, "product effectiveness")
                self.assertEqual(summary.cell(data_row, 4).value, "产品效果")
                self.assertEqual(summary.cell(data_row, 5).value, 1)
                self.assertEqual(regrouped[PAIN_POINT_SHEET]["G2"].value, "产品效果")
                self.assertEqual(len(regrouped[PAIN_POINT_SHEET]._images), 1)
            finally:
                regrouped.close()

    def test_translation_report_contains_only_critical_reviews(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "source.xlsx"
            output = Path(temp_dir) / "output.xlsx"
            translation_report = Path(temp_dir) / "missing_review_translations.json"
            workbook = Workbook()
            workbook.active.title = "数据源"
            workbook.save(source)
            workbook.close()

            positive_review = normalize_review(
                {
                    "reviewId": "R-POSITIVE",
                    "ratingText": "5.0 out of 5 stars",
                    "title": "Great toy",
                    "body": "The kids loved it.",
                    "verifiedText": "Verified Purchase",
                    "reviewUrl": "https://example.test/review/R-POSITIVE",
                },
                asin="B0ABCDEFGH",
                page_number=1,
                source_url="https://example.test/positive",
            )
            critical_review = normalize_review(
                {
                    "reviewId": "R-CRITICAL",
                    "ratingText": "1.0 out of 5 stars",
                    "title": "Stopped working",
                    "body": "It broke after one day.",
                    "verifiedText": "Verified Purchase",
                    "reviewUrl": "https://example.test/review/R-CRITICAL",
                },
                asin="B0ABCDEFGH",
                page_number=1,
                source_url="https://example.test/critical",
            )
            common = {
                "asin": "B0ABCDEFGH",
                "reviewer_type": "all",
                "sort_by": "helpful",
                "product_title": "Demo",
                "product_url": "https://www.amazon.com/dp/B0ABCDEFGH",
            }
            write_review_sheets(
                source,
                output,
                [
                    AsinResult(star_filter="positive", reviews=[positive_review], **common),
                    AsinResult(star_filter="critical", reviews=[critical_review], **common),
                ],
                site="US",
                translation_cache=Path(temp_dir) / "empty_cache.json",
                translation_report=translation_report,
                pain_group_cache=Path(temp_dir) / "empty_groups.json",
                pain_group_report=Path(temp_dir) / "missing_pain_groups.json",
            )

            report = json.loads(translation_report.read_text(encoding="utf-8"))
            self.assertEqual(report["review_scope"], "Critical reviews only")
            self.assertEqual(report["positive_review_translation"], "skipped")
            self.assertEqual(report["missing_count"], 1)
            self.assertIn("Stopped working\nIt broke after one day.", report["translations"])
            self.assertNotIn("Great toy\nThe kids loved it.", report["translations"])


if __name__ == "__main__":
    unittest.main()
