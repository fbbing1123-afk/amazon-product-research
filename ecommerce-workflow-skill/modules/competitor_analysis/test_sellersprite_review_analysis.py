"""验证新版评论筛选、跨商品隔离及动态单元格回填。"""

from copy import deepcopy
import json
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, PatternFill

from competitor_analysis.sellersprite_review_analysis import (
    SCHEMA, agent_prompt, combine, fill_workbook, load_packet, make_packet, prepare, validate_result,
)
from review_research.download_sellersprite_reviews import helpful_count
from competitor_analysis.review_mentions import MENTION_SCOPE, format_report_point, mention_stats


class SellerSpriteAnalysisTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.export = self.root / "export.xlsx"
        self.asin = "B012345678"
        self.headers = ["ASIN", "标题", "内容", "VP评论", "Vine Voice评论", "星级", "评论链接", "评论时间",
                        "赞同数", "标题(翻译)", "内容(翻译)"]

    def row(self, number, rating=5, helpful=0, vp="Y", vine="", asin=None, title="真实标题", body="真实评价"):
        return [asin, title, body, vp, vine, rating, f"https://www.amazon.com/gp/customer-reviews/R{number}",
                "2026-11-24", helpful, "中文标题" if title else "", "中文评价" if body else ""]

    def source(self, rows, missing_helpful=False):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = f"{self.asin}-Reviews({len(rows)})"
        sheet.append([h for h in self.headers if not missing_helpful or h != "赞同数"])
        for row in rows:
            sheet.append(row[:8] + row[9:] if missing_helpful else row)
        workbook.save(self.export)
        workbook.close()
        return make_packet(self.export, self.asin)

    def result(self, packet, kind, points=None):
        return {"schema": SCHEMA, "asin": packet["asin"], "packet_digest": packet["packet_digest"],
                "kind": kind, "reviewed_numbers": [r["number"] for r in packet["samples"][kind]],
                "points": points or [], "no_supported_points_reason": "样本没有明确支持的结论",
                "evidence_scope": MENTION_SCOPE}

    def save(self, name, payload):
        path = self.root / name
        path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        return path

    def test_or_filter_before_sort_and_independent_star_groups(self):
        packet = self.source([
            self.row(1, 5, 900, vp="", vine=""),
            self.row(2, 4, 3, vp="", vine="Y"),
            self.row(3, 5, 9, vp="Y", vine="Y"),
            self.row(4, 3, 40), self.row(5, 1, 20, vp="", vine="Y"),
            self.row(6, 2, 50), self.row(7, 1, 999, vp="", vine=""),
        ])
        self.assertEqual(packet["trusted_count"], 5)
        self.assertEqual([r["review_id"] for r in packet["samples"]["advantages"]], ["R3", "R2"])
        self.assertEqual([r["review_id"] for r in packet["samples"]["disadvantages"]], ["R6", "R4", "R5"])

    def test_limits_30_and_100_are_after_numeric_descending_sort(self):
        packet = self.source([self.row(i, rating=5 if i <= 40 else 2, helpful=i) for i in range(1, 151)])
        self.assertEqual(packet["eligible_counts"], {"advantages": 40, "disadvantages": 110})
        self.assertEqual([r["review_id"] for r in packet["samples"]["advantages"]], [f"R{i}" for i in range(40, 10, -1)])
        self.assertEqual([r["review_id"] for r in packet["samples"]["disadvantages"]], [f"R{i}" for i in range(150, 50, -1)])
        self.assertEqual([r["number"] for r in packet["samples"]["disadvantages"]], list(range(1, 101)))
        points = validate_result(packet, "disadvantages", self.result(
            packet, "disadvantages", [{"text": "配件缺失", "review_numbers": [1]}]))
        self.assertEqual(format_report_point(points[0], "disadvantages"), "配件缺失（1%）")

    def test_ties_keep_source_order_and_empty_asin_remains_empty(self):
        packet = self.source([self.row(1, helpful=None), self.row(2, helpful="0"), self.row(3, helpful="1,200")])
        rows = packet["samples"]["advantages"]
        self.assertEqual([r["review_id"] for r in rows], ["R3", "R1", "R2"])
        self.assertEqual(rows[0]["helpful_count"], 1200)
        self.assertTrue(all(r["source_row_asin"] is None for r in rows))

    def test_invalid_helpful_and_missing_column_do_not_silently_sort_as_zero(self):
        for value in (-1, "1k", "1,2", 2.5, True, "unknown", float("inf"), float("nan")):
            with self.subTest(value=value), self.assertRaises(ValueError):
                helpful_count(value)
        self.assertEqual(helpful_count(None), 0)
        with self.assertRaisesRegex(ValueError, "缺少赞同数"):
            self.source([self.row(1)], missing_helpful=True)

    def test_empty_text_is_not_backfilled_with_a_lower_ranked_review(self):
        packet = self.source([self.row(1, helpful=999, title="", body="")] +
                             [self.row(i, helpful=i) for i in range(2, 34)])
        self.assertEqual(len(packet["samples"]["advantages"]), 30)
        self.assertEqual(packet["samples"]["advantages"][0]["review_id"], "R1")
        result = self.result(packet, "advantages", [{"text": "编造内容", "review_numbers": [1]}])
        with self.assertRaisesRegex(ValueError, "没有原始文本"):
            validate_result(packet, "advantages", result)

    def test_prepare_keeps_original_and_emits_isolated_numbered_prompt(self):
        self.source([self.row(1, rating=1, body="本商品独有差评"), self.row(2, body="只属于好评")])
        original = self.export.read_bytes()
        first = prepare(self.export, self.asin, self.root / "runs")
        second = prepare(self.export, self.asin, self.root / "runs")
        self.assertNotEqual(first, second)
        self.assertEqual(self.export.read_bytes(), original)
        packet = load_packet(first / "packet.json")
        prompt = agent_prompt(packet, "disadvantages")
        self.assertIn("每点精简在 15 个字以内", prompt)
        self.assertIn('"序号": 1', prompt)
        self.assertIn("本商品独有差评", prompt)
        self.assertNotIn("只属于好评", prompt)
        self.assertIn("列全明确提及该问题的所有评论序号", prompt)
        self.assertIn(MENTION_SCOPE, prompt)

    def test_wrong_asin_digest_kind_and_incomplete_reviewing_are_rejected(self):
        packet = self.source([self.row(1, rating=1)])
        result = self.result(packet, "disadvantages", [{"text": "配件缺失", "review_numbers": [1]}])
        for field, value in (("asin", "B087654321"), ("packet_digest", "old"),
                             ("kind", "advantages"), ("reviewed_numbers", []), ("reviewed_numbers", [True])):
            changed = deepcopy(result)
            changed[field] = value
            with self.subTest(field=field), self.assertRaises(ValueError):
                validate_result(packet, "disadvantages", changed)

    def test_length_evidence_duplicates_and_empty_result_are_rejected(self):
        packet = self.source([self.row(1, rating=2)])
        for points in ([{"text": "长" * 16, "review_numbers": [1]}],
                       [{"text": "配件缺失", "review_numbers": [2]}],
                       [{"text": "配件缺失", "review_numbers": [True]}],
                       [{"text": "配件缺失", "review_numbers": []}],
                       [{"text": "配件缺失", "review_numbers": [1, 1]}],
                       [{"text": "配件缺失50%", "review_numbers": [1]}],
                       [{"text": "1. 配件缺失", "review_numbers": [1]}],
                       [{"text": "配件缺失", "review_numbers": [1]}] * 2):
            with self.subTest(points=points), self.assertRaises(ValueError):
                validate_result(packet, "disadvantages", self.result(packet, "disadvantages", points))
        result = self.result(packet, "disadvantages")
        result["no_supported_points_reason"] = ""
        with self.assertRaisesRegex(ValueError, "必须说明原因"):
            validate_result(packet, "disadvantages", result)

    def test_exact_15_chars_and_evidence_mapping_pass(self):
        packet = self.source([self.row(1, rating=2)])
        points = [{"text": "字" * 15, "review_numbers": [1]}]
        result = validate_result(packet, "disadvantages", self.result(packet, "disadvantages", points))
        self.assertEqual(result[0]["review_ids"], ["R1"])
        self.assertEqual(format_report_point(result[0], "disadvantages"), "字" * 15 + "（100%）")

    def test_negative_mentions_use_selected_denominator_including_blank_and_allow_multiple_issues(self):
        packet = self.source([self.row(1, rating=1), self.row(2, rating=2),
                              self.row(3, rating=3, title="", body=""), self.row(4, rating=5),
                              self.row(5, rating=1, vp="", vine="")])
        points = validate_result(packet, "disadvantages", self.result(packet, "disadvantages", [
            {"text": "配件缺失", "review_numbers": [1, 2]},
            {"text": "没有说明书", "review_numbers": [1, 2]},
        ]))
        self.assertEqual([(p["mention_count"], p["sample_count"], p["mention_percent"]) for p in points],
                         [(2, 3, 66.7), (2, 3, 66.7)])
        self.assertEqual(format_report_point(points[0], "disadvantages"), "配件缺失（66.7%）")
        good = validate_result(packet, "advantages", self.result(packet, "advantages", [
            {"text": "容易使用", "review_numbers": [1]}]))
        self.assertNotIn("mention_percent", good[0])
        self.assertEqual(format_report_point(good[0], "advantages"), "容易使用")

    def test_legacy_representative_evidence_must_be_reanalysed_before_counting(self):
        packet = self.source([self.row(1, rating=1)])
        result = self.result(packet, "disadvantages", [{"text": "配件缺失", "review_numbers": [1]}])
        result.pop("evidence_scope")
        with self.assertRaisesRegex(ValueError, "未确认全部提及"):
            validate_result(packet, "disadvantages", result)

    def test_mention_rounding_invalid_counts_and_tampered_statistics(self):
        self.assertEqual(mention_stats([1], 16)["mention_percent"], 6.3)
        for evidence, count in (([], 0), ([1, 1], 2), ([3], 2), ([True], 2)):
            with self.subTest(evidence=evidence, count=count), self.assertRaises(ValueError):
                mention_stats(evidence, count)
        point = {"text": "配件缺失", "review_numbers": [1], **mention_stats([1], 2)}
        point["mention_percent"] = 100
        with self.assertRaisesRegex(ValueError, "有效提及统计"):
            format_report_point(point, "disadvantages")

    def test_changed_source_or_packet_is_rejected(self):
        packet = self.source([self.row(1, rating=2)])
        path = self.save("packet.json", packet)
        changed = deepcopy(packet)
        changed["samples"]["disadvantages"][0]["body"] = "其它产品缓存"
        self.save("packet.json", changed)
        with self.assertRaisesRegex(ValueError, "已变化"):
            load_packet(path)
        self.save("packet.json", packet)
        self.source([self.row(1, rating=2, body="新的评论正文")])
        with self.assertRaisesRegex(ValueError, "已变化"):
            load_packet(path)

    def test_empty_filtered_samples_cannot_invent_points(self):
        packet = self.source([self.row(1, vp="", vine="")])
        for kind in ("advantages", "disadvantages"):
            self.assertEqual(validate_result(packet, kind, self.result(packet, kind)), [])
            with self.assertRaisesRegex(ValueError, "证据序号"):
                validate_result(packet, kind, self.result(packet, kind, [{"text": "凭空分析", "review_numbers": [1]}]))

    def test_other_variant_asin_is_preserved_without_cross_file_lookup(self):
        packet = self.source([self.row(1, asin="B087654321")])
        self.assertEqual(packet["samples"]["advantages"][0]["source_row_asin"], "B087654321")
        self.assertTrue(any("其它子体" in message for message in packet["warnings"]))

    def make_report(self, shifted):
        path = self.root / "report.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "产品竞品分析"
        for row, label in ((4, "详情页链接"), (9 + shifted, "优点"), (10 + shifted, "缺点")):
            sheet.cell(row, 2, label)
        sheet["C4"] = self.asin
        sheet["C4"].hyperlink = f"https://www.amazon.com/dp/{self.asin}"
        sheet["D4"] = "B087654321"
        sheet["D4"].hyperlink = "https://www.amazon.com/dp/B087654321"
        sheet.merge_cells("A1:E1")
        sheet["A1"] = "原有标题"
        sheet.column_dimensions["C"].width = 42
        for row in (9 + shifted, 10 + shifted):
            sheet.row_dimensions[row].height = 105
            for column in (3, 4):
                cell = sheet.cell(row, column, "原有内容")
                cell.fill = PatternFill("solid", fgColor="FFEEAA")
                cell.alignment = Alignment(wrap_text=True)
        workbook.create_sheet("其它页")["A1"] = "保持原样"
        workbook.save(path)
        workbook.close()
        with ZipFile(path, "a") as archive:
            archive.writestr("xl/slicers/local-preservation-check.xml", b"preserve-native-part")
        return path

    def test_fill_uses_shifted_labels_keeps_styles_other_asin_and_native_parts(self):
        packet = self.source([self.row(1), self.row(2, rating=1)])
        analysis = combine(self.save("packet.json", packet),
                           self.save("good.json", self.result(packet, "advantages", [{"text": "容易使用", "review_numbers": [1]}])),
                           self.save("bad.json", self.result(packet, "disadvantages", [{"text": "配件缺失", "review_numbers": [1]}])))
        workbook = self.make_report(shifted=4)
        before = workbook.read_bytes()
        output = self.root / "filled.xlsx"
        result = fill_workbook(workbook, output, analysis)
        self.assertEqual(result["cells"], {"advantages": "C13", "disadvantages": "C14"})
        self.assertEqual(workbook.read_bytes(), before)
        with ZipFile(workbook) as a, ZipFile(output) as b:
            changed = [name for name in a.namelist() if a.read(name) != b.read(name)]
            self.assertEqual(changed, ["xl/worksheets/sheet1.xml"])
        original, filled = load_workbook(workbook), load_workbook(output)
        try:
            sheet = filled["产品竞品分析"]
            self.assertEqual(sheet["C13"].value, "容易使用")
            self.assertEqual(sheet["C14"].value, "配件缺失（100%）")
            self.assertEqual(sheet["D14"].value, "原有内容")
            self.assertEqual(sheet["C13"]._style, original["产品竞品分析"]["C13"]._style)
            self.assertEqual(sheet.row_dimensions[13].height, 105)
            self.assertEqual(str(sheet.merged_cells), "A1:E1")
        finally:
            original.close()
            filled.close()
        with self.assertRaisesRegex(ValueError, "独立输出"):
            fill_workbook(workbook, workbook, analysis)
        with self.assertRaisesRegex(ValueError, "独立输出"):
            fill_workbook(workbook, output, analysis)

    def test_wrong_target_product_does_not_generate_output(self):
        packet = self.source([self.row(1)])
        workbook = self.make_report(0)
        analysis = {"asin": "B000000000", "source_file": str(self.export), "advantages": [], "disadvantages": []}
        output = self.root / "wrong.xlsx"
        with self.assertRaisesRegex(ValueError, "详情页列"):
            fill_workbook(workbook, output, analysis)
        self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
