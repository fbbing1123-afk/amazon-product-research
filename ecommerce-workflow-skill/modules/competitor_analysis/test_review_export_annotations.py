"""验证评论原表加列的证据对齐、完整保留和防覆盖行为。"""

from copy import deepcopy
import json
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, PatternFill

from competitor_analysis.annotate_review_export import HEADERS, annotate_export
from competitor_analysis.sellersprite_review_analysis import SCHEMA, combine, make_packet
from review_research.download_sellersprite_reviews import read_export


class ReviewExportAnnotationTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.export = self.root / "export.xlsx"
        self.asin = "B012345678"
        self.headers = ["ASIN", "标题", "内容", "VP评论", "Vine Voice评论", "星级", "评论链接", "评论时间", "赞同数"]

    def source(self, rows):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = f"{self.asin}-Reviews({len(rows)})"
        sheet.append(self.headers)
        for index, row in enumerate(rows, 1):
            rating, helpful, vp, vine = row
            sheet.append([None, f"标题{index}", f"正文{index}", vp, vine, rating,
                          f"https://www.amazon.com/gp/customer-reviews/R{index}", "2026-11-24", helpful])
        for cell in sheet[1]:
            cell.fill = PatternFill("solid", fgColor="FFE98A00")
        for row in range(2, sheet.max_row + 1):
            sheet.cell(row, 3).alignment = Alignment(horizontal="left")
            sheet.row_dimensions[row].height = 24
        workbook.create_sheet("Note")["A1"] = "保留原始说明"
        workbook.save(self.export)
        workbook.close()
        with ZipFile(self.export, "a") as archive:
            archive.writestr("xl/drawings/native-preservation.xml", b"preserve-native-drawing")
        return make_packet(self.export, self.asin)

    def analysis(self, packet, good=None, bad=None):
        self.packet_path = self.root / "packet.json"
        self.packet_path.write_text(json.dumps(packet, ensure_ascii=False), encoding="utf-8")
        paths = []
        for kind, points in (("advantages", good or []), ("disadvantages", bad or [])):
            result = {"schema": SCHEMA, "asin": self.asin, "packet_digest": packet["packet_digest"], "kind": kind,
                      "reviewed_numbers": [r["number"] for r in packet["samples"][kind]], "points": points,
                      "no_supported_points_reason": "没有明确结论"}
            path = self.root / f"{kind}.json"
            path.write_text(json.dumps(result, ensure_ascii=False), encoding="utf-8")
            paths.append(path)
        return combine(self.packet_path, *paths)

    def test_multiple_reasons_stay_on_original_row_and_all_reviews_remain(self):
        packet = self.source([(1, 1, "Y", ""), (1, 9, "", "Y"), (5, 10, "Y", ""), (2, 100, "", "")])
        analysis = self.analysis(packet,
            good=[{"text": "使用方便", "review_numbers": [1]}],
            bad=[{"text": "配件缺失", "review_numbers": [1, 2]}, {"text": "螺母无法拧紧", "review_numbers": [2]}])
        before = self.export.read_bytes()
        output = self.root / "annotated.xlsx"
        result = annotate_export(packet, analysis, output)
        self.assertEqual(result["review_rows"], 4)
        source, annotated = read_export(self.export, self.asin), read_export(output, self.asin)
        self.assertEqual(annotated["headers"], self.headers + HEADERS)
        self.assertNotIn("分析入选", annotated["headers"])
        self.assertNotIn("分析编号", annotated["headers"])
        self.assertEqual(annotated["reviews"][0]["raw"]["差评原因"], "配件缺失\n螺母无法拧紧")
        self.assertEqual(annotated["reviews"][1]["raw"]["差评原因"], "配件缺失")
        self.assertEqual(annotated["reviews"][2]["raw"]["优点归纳"], "使用方便")
        self.assertTrue(all(not annotated["reviews"][3]["raw"][key] for key in HEADERS))
        self.assertEqual(self.export.read_bytes(), before)
        self.assertEqual([r["review_id"] for r in source["reviews"]], [r["review_id"] for r in annotated["reviews"]])
        self.assertTrue(all(r["source_row_asin"] is None for r in annotated["reviews"]))
        with ZipFile(self.export) as a, ZipFile(output) as b:
            self.assertEqual(a.namelist(), b.namelist())
            changed = {name for name in a.namelist() if a.read(name) != b.read(name)}
            self.assertEqual(changed, {"xl/worksheets/sheet1.xml", "xl/styles.xml"})
        original, workbook = load_workbook(self.export), load_workbook(output)
        try:
            self.assertEqual(workbook.sheetnames, original.sheetnames)
            sheet, old = workbook.worksheets[0], original.worksheets[0]
            self.assertEqual(sheet.auto_filter.ref, "A1:L5")
            self.assertEqual(sheet.freeze_panes, "A2")
            self.assertEqual(sheet["J1"]._style, old["A1"]._style)
            self.assertTrue(sheet["K2"].alignment.wrap_text)
            self.assertGreater(sheet.row_dimensions[2].height, 24)
            self.assertEqual(sheet.row_dimensions[5].height, old.row_dimensions[5].height)
            for row in range(1, old.max_row + 1):
                for column in range(1, old.max_column + 1):
                    self.assertEqual(sheet.cell(row, column).value, old.cell(row, column).value)
                    self.assertEqual(sheet.cell(row, column)._style, old.cell(row, column)._style)
        finally:
            original.close()
            workbook.close()

    def test_unselected_reviews_stay_blank_and_selected_uncited_reviews_have_explanations(self):
        packet = self.source([(5, number, "Y", "") for number in range(31)] + [(1, number, "Y", "") for number in range(101)])
        analysis = self.analysis(packet, good=[{"text": "容易使用", "review_numbers": [1]}],
                                 bad=[{"text": "配件缺失", "review_numbers": [1]}])
        output = self.root / "limits.xlsx"
        annotate_export(packet, analysis, output)
        rows = read_export(output, self.asin)["reviews"]
        self.assertTrue(all(not rows[0]["raw"][key] for key in HEADERS))
        self.assertTrue(all(not rows[31]["raw"][key] for key in HEADERS))
        self.assertEqual(rows[1]["raw"]["分析说明"], "已分析，未被结论引用")

    def test_empty_export_still_has_headers_without_fabricated_reviews(self):
        packet = self.source([])
        analysis = self.analysis(packet)
        output = self.root / "empty.xlsx"
        annotate_export(packet, analysis, output)
        result = read_export(output, self.asin)
        self.assertEqual(result["review_count"], 0)
        self.assertEqual(result["headers"], self.headers + HEADERS)

    def test_wrong_source_or_result_and_overwrite_are_rejected(self):
        packet = self.source([(1, 10, "Y", "")])
        analysis = self.analysis(packet)
        with self.assertRaisesRegex(ValueError, "独立输出"):
            annotate_export(packet, analysis, self.export)
        changed = deepcopy(analysis)
        changed["asin"] = "B087654321"
        with self.assertRaisesRegex(ValueError, "不匹配"):
            annotate_export(packet, changed, self.root / "wrong.xlsx")
        self.source([(1, 20, "Y", "")])
        with self.assertRaisesRegex(ValueError, "源文件已变化"):
            annotate_export(packet, analysis, self.root / "stale.xlsx")

    def test_annotation_columns_do_not_duplicate(self):
        packet = self.source([(1, 10, "Y", "")])
        output = self.root / "annotated.xlsx"
        annotate_export(packet, self.analysis(packet), output)
        with self.assertRaisesRegex(ValueError, "独立输出"):
            annotate_export(packet, self.analysis(packet), output)
        annotated_packet = make_packet(output, self.asin)
        with self.assertRaisesRegex(ValueError, "已含分析列"):
            annotate_export(annotated_packet, self.analysis(annotated_packet), self.root / "duplicated.xlsx")


if __name__ == "__main__":
    unittest.main()
