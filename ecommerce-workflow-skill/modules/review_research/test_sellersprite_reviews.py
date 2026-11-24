"""验证评论源文件归属、数量、原始空值及异常记录，避免分析错用数据。"""

from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock

from openpyxl import Workbook

from .download_sellersprite_reviews import NamedDownload, asin_from_input, continue_after_user_check, export_status, read_export


class SellerSpriteExportTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "reviews.xlsx"
        self.asin = "B012345678"

    def make_export(self, rows, declared=None, asin=None, headers=None):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = f"{asin or self.asin}-Reviews({len(rows) if declared is None else declared})"
        sheet.append(headers or ["ASIN", "标题", "内容", "VP评论", "Vine Voice评论", "星级", "评论链接", "评论时间"])
        for row in rows:
            sheet.append(row)
        workbook.create_sheet("Note").append(["此页不是评论数据"])
        workbook.save(self.path)
        workbook.close()

    def row(self, review_id="R123456", **changes):
        data = {"asin": None, "title": "示例标题", "body": "示例正文", "vp": "Y", "vine": "", "rating": 5,
                "url": f"https://www.amazon.com/portal/customer-reviews/srp/-/{review_id}/ref=test", "date": "2026-11-24"}
        data.update(changes)
        return list(data.values())

    def test_input_site_and_dedicated_product_path(self):
        self.assertEqual(asin_from_input("b012345678"), self.asin)
        self.assertEqual(asin_from_input("https://www.amazon.com/title/dp/B012345678?th=1"), self.asin)
        for invalid in ["https://www.amazon.ca/dp/B012345678", "https://www.amazon.com.evil.test/dp/B012345678", "https://www.amazon.com/s?k=B012345678"]:
            with self.subTest(invalid=invalid), self.assertRaises(ValueError):
                asin_from_input(invalid)

    def test_blank_asin_and_variant_are_preserved(self):
        self.make_export([self.row(), self.row("R234567", asin="B087654321", vp="", vine="Y", rating=1)])
        result = read_export(self.path, self.asin)
        self.assertEqual(result["review_count"], 2)
        self.assertIsNone(result["reviews"][0]["source_row_asin"])
        self.assertIsNone(result["reviews"][0]["raw"]["ASIN"])
        self.assertEqual(result["reviews"][1]["source_row_asin"], "B087654321")
        self.assertEqual(result["rating_counts"], {"1": 1, "5": 1})
        self.assertEqual(result["trusted_count"], 2)
        self.assertFalse(result["all_amazon_reviews_confirmed"])
        self.assertEqual(result["blank_source_asin_count"], 1)
        self.assertFalse(any("原始 ASIN 为空" in warning for warning in result["warnings"]))
        self.assertTrue(any("其它子体" in warning for warning in result["warnings"]))

    def test_wrong_product_export_rejected(self):
        self.make_export([self.row()], asin="B087654321")
        with self.assertRaisesRegex(ValueError, "错商品"):
            read_export(self.path, self.asin)

    def test_us_export_preserves_international_review_links(self):
        url = "https://www.amazon.de/gp/customer-reviews/R123456?ASIN=" + self.asin
        self.make_export([self.row(url=url)])
        result = read_export(self.path, self.asin)
        self.assertEqual(url, result["reviews"][0]["review_url"])
        self.assertEqual(self.asin, result["reviews"][0]["asin"])
        self.make_export([self.row(url=url.replace("amazon.de", "amazon.de.example.com"))])
        with self.assertRaises(ValueError):
            read_export(self.path, self.asin)

    def test_missing_columns_rejected(self):
        self.make_export([], headers=["标题", "内容"])
        with self.assertRaisesRegex(ValueError, "表头"):
            read_export(self.path, self.asin)

    def test_count_mismatch_and_duplicates_rejected(self):
        for rows, declared, message in [([self.row()], 2, "数量不一致"), ([self.row(), self.row()], 2, "重复评论")]:
            with self.subTest(message=message):
                self.make_export(rows, declared=declared)
                with self.assertRaisesRegex(ValueError, message):
                    read_export(self.path, self.asin)

    def test_invalid_rating_link_and_flag_rejected(self):
        for changes in [{"rating": 6}, {"rating": "未知"}, {"url": "https://example.com/R1"}, {"vp": "未知"}]:
            with self.subTest(changes=changes):
                self.make_export([self.row(**changes)])
                with self.assertRaises(ValueError):
                    read_export(self.path, self.asin)

    def test_empty_body_kept_for_user_review(self):
        self.make_export([self.row(body="")])
        result = read_export(self.path, self.asin)
        self.assertEqual(result["review_count"], 1)
        self.assertEqual(result["reviews"][0]["body"], "")
        self.assertTrue(any("正文为空" in item for item in result["warnings"]))

    def test_empty_export_does_not_mean_no_amazon_reviews(self):
        self.make_export([])
        result = read_export(self.path, self.asin)
        self.assertEqual(result["review_count"], 0)
        self.assertFalse(result["all_amazon_reviews_confirmed"])
        self.assertTrue(result["warnings"])

    def test_translation_columns_and_missing_values(self):
        base = ["ASIN", "标题", "内容", "VP评论", "Vine Voice评论", "星级", "评论链接", "评论时间"]
        self.make_export([self.row()])
        result = read_export(self.path, self.asin)
        self.assertEqual(export_status(result, "downloaded"), "incomplete_translation")
        for body, expected in [("中文正文", "downloaded"), ("", "incomplete_translation")]:
            with self.subTest(expected=expected):
                self.make_export([self.row() + ["中文标题", body]], headers=base + ["标题(翻译)", "内容(翻译)"])
                result = read_export(self.path, self.asin)
                self.assertEqual(export_status(result, "downloaded"), expected)
                self.assertEqual(result["reviews"][0]["title_translated"], "中文标题")

    def test_only_continue_after_checkbox_already_checked(self):
        page = Mock(url="https://www.amazon.com/checkpoint/continue")
        checkbox, button = Mock(), Mock()
        checkbox.count.return_value = button.count.return_value = 1
        page.locator.return_value = checkbox
        page.get_by_role.return_value = button
        for checked, enabled, expected in [(False, False, False), (False, True, False), (True, False, False), (True, True, True)]:
            with self.subTest(checked=checked, enabled=enabled):
                checkbox.is_checked.return_value = checked
                button.is_enabled.return_value = enabled
                button.click.reset_mock()
                self.assertEqual(continue_after_user_check(page), expected)
                self.assertEqual(button.click.call_count, int(expected))
                checkbox.click.assert_not_called()
                checkbox.check.assert_not_called()

    def test_download_ignores_other_tabs_and_requires_completion(self):
        download = NamedDownload(Mock(), Path(self.directory.name))
        download.frame_id = "target-frame"
        download.on_begin({"frameId": "another-frame", "guid": "other", "suggestedFilename": "other.xlsx"})
        self.assertIsNone(download.guid)
        download.on_begin({"frameId": "target-frame", "guid": "target", "suggestedFilename": "reviews.xlsx"})
        download.on_progress({"guid": "other", "state": "completed"})
        with self.assertRaisesRegex(RuntimeError, "尚未完成"):
            download.completed_path()
        self.path.write_bytes(b"test")
        download.on_progress({"guid": "target", "state": "completed"})
        self.assertEqual(download.completed_path(), self.path.resolve())

    def test_download_rejects_bad_name_and_wrong_directory(self):
        download = NamedDownload(Mock(), Path(self.directory.name))
        download.frame_id = "target-frame"
        download.on_begin({"frameId": "target-frame", "guid": "target", "suggestedFilename": "../wrong.xlsx"})
        with self.assertRaisesRegex(ValueError, "文件名"):
            download.completed_path()
        download.error = None
        download.filename = "reviews.xlsx"
        download.on_progress({"guid": "target", "state": "completed", "filePath": str(self.path.parent.parent / "wrong.xlsx")})
        with self.assertRaisesRegex(ValueError, "输出目录"):
            download.completed_path()

    def test_download_restores_browser_defaults_when_work_fails(self):
        page = Mock()
        page.context.new_cdp_session.return_value.send.return_value = {"frameTree": {"frame": {"id": "target-frame"}}}
        session = page.context.browser.new_browser_cdp_session.return_value
        with self.assertRaisesRegex(RuntimeError, "下载中断"):
            with NamedDownload(page, Path(self.directory.name)):
                raise RuntimeError("下载中断")
        session.send.assert_any_call("Browser.setDownloadBehavior", {
            "behavior": "allow", "downloadPath": str(Path(self.directory.name).resolve()), "eventsEnabled": True,
        })
        session.send.assert_called_with("Browser.setDownloadBehavior", {"behavior": "default", "eventsEnabled": False})
        session.detach.assert_called_once()


if __name__ == "__main__":
    unittest.main()
