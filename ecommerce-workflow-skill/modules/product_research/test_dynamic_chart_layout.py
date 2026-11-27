from copy import copy
from pathlib import Path
import tempfile
import unittest
import zipfile

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, TwoCellAnchor
from openpyxl.styles import Alignment, Border, PatternFill, Side
from PIL import Image as PillowImage

from modules.keyword_backfill.seller_sprite_keyword_tool import (
    KeywordRow, TemplateInfo, write_selected_keywords,
)
from modules.product_research.scripts.fill_custom_template import ensure_category_rows
from modules.product_research.scripts.insert_sellersprite_screenshots import (
    CHARTS, insert_screenshots, locate_chart_regions,
)
from modules.shared.worksheet_rows import insert_rows_preserving_merges


class DynamicChartLayoutTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.picture = self.root / "chart.png"
        PillowImage.new("RGB", (800, 200), "white").save(self.picture)

    def fixture(self):
        wb = Workbook()
        ws = wb.active
        ws.title = "产品-市场调研"
        ws["A2"] = "1. 关键词搜索量及趋势"
        ws["A4"] = "关键词"
        ws["A8"] = "2. 所属平台分类及常用分类节点"
        ws["A9"] = "一级类目"
        ws["A12"] = "3.市场数据分析"
        for row in (5, 6, 7, 10, 11):
            ws.row_dimensions[row].height = 36
            ws.merge_cells(start_row=row, end_row=row, start_column=6, end_column=8)
        for row, label in enumerate(CHARTS.values(), 21):
            ws.cell(row, 1, label)
            ws.row_dimensions[row].height = 195
            ws.row_dimensions[row].outlineLevel = 1
            ws.row_dimensions[row].thickBot = True
            for col in range(1, 10):
                cell = ws.cell(row, col)
                cell.fill = PatternFill("solid", fgColor="DAEEF3")
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = Border(*(Side(style="thin"),) * 4)
            ws.merge_cells(f"B{row}:F{row}")
            ws.merge_cells(f"G{row}:I{row}")
        ws.row_dimensions[30].hidden = True
        wb.create_sheet("其它页")["A1"] = "保留"
        return wb

    def expand(self, ws):
        # 调用两个模块的真实扩容入口：关键词增加三行，类目增加两行。
        info = TemplateInfo(ws, "产品", [], 4, 1)
        selected = [KeywordRow([], f"keyword {i}", "词", 1, 10) for i in range(6)]
        write_selected_keywords(info, selected)
        ensure_category_rows(ws, 4, 13, 15)

    def save(self, wb, name):
        path = self.root / name
        wb.save(path)
        return path

    def insert(self, source, name):
        output = self.root / name
        result = insert_screenshots(source, output, {key: self.picture for key in CHARTS})
        return output, result

    def test_both_modules_move_all_existing_formats(self):
        wb = self.fixture()
        ws = wb.active
        styles = {(r, c): copy(ws.cell(r, c)._style) for r in range(12, 24) for c in range(1, 10)}
        dimensions = {r: dict(ws.row_dimensions[r]) for r in (21, 22, 23, 30)}
        self.expand(ws)
        self.assertEqual([26, 27, 28], [a.min_row for a in locate_chart_regions(ws).values()])
        for (row, col), style in styles.items():
            self.assertEqual(style, ws.cell(row + 5, col)._style, (row, col))
        for row, dim in dimensions.items():
            self.assertEqual(dim, dict(ws.row_dimensions[row + 5]))
            self.assertEqual(row + 5, ws.row_dimensions[row + 5].index)
        self.assertEqual(36, ws.row_dimensions[8].height)
        self.assertEqual(36, ws.row_dimensions[15].height)
        self.assertIn("G26:I26", ws.merged_cells)

    def test_string_one_cell_and_two_cell_anchors_move_after_reload(self):
        wb = self.fixture()
        for anchor in ("B21", OneCellAnchor(_from=AnchorMarker(col=1, row=21)),
                       TwoCellAnchor(_from=AnchorMarker(col=1, row=22), to=AnchorMarker(col=5, row=23))):
            image = Image(str(self.picture))
            image.anchor = anchor
            wb.active.add_image(image)
        self.expand(wb.active)
        path = self.save(wb, "anchors.xlsx")
        loaded = load_workbook(path)
        self.addCleanup(loaded.close)
        anchors = [i.anchor for i in loaded.active._images]
        self.assertEqual([25, 26, 27], [a._from.row for a in anchors])
        self.assertEqual(28, anchors[2].to.row)

    def test_insert_before_screenshots_preserves_all_other_parts(self):
        wb = self.fixture()
        self.expand(wb.active)
        source = self.save(wb, "expanded.xlsx")
        # 用未知扩展部件模拟 WPS 原生部件，必须逐字节保留。
        with zipfile.ZipFile(source, "a") as archive:
            archive.writestr("xl/cellimages.xml", b"preserve-native-extension")
        output, result = self.insert(source, "placed.xlsx")
        self.assertEqual(["B26:F26", "B27:F27", "B28:F28"], list(result.values()))
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(output) as after:
            changed = [n for n in before.namelist() if before.read(n) != after.read(n)]
            self.assertEqual({"xl/worksheets/sheet1.xml", "xl/worksheets/_rels/sheet1.xml.rels",
                              "[Content_Types].xml"}, set(changed))
        loaded = load_workbook(output)
        self.addCleanup(loaded.close)
        self.assertEqual([25, 26, 27], [i.anchor._from.row for i in loaded.active._images])

    def test_insert_after_screenshots_moves_three_images(self):
        source = self.save(self.fixture(), "source.xlsx")
        placed, _ = self.insert(source, "placed.xlsx")
        wb = load_workbook(placed)
        self.addCleanup(wb.close)
        self.expand(wb.active)
        self.assertEqual([25, 26, 27], [i.anchor._from.row for i in wb.active._images])
        self.assertEqual([195] * 3, [wb.active.row_dimensions[r].height for r in (26, 27, 28)])

    def test_repeated_insertion_keeps_other_image_and_reuses_media(self):
        wb = self.fixture()
        wb.active.add_image(Image(str(self.picture)), "B19")
        source = self.save(wb, "source.xlsx")
        first, _ = self.insert(source, "first.xlsx")
        second, _ = self.insert(first, "second.xlsx")
        with zipfile.ZipFile(first) as before, zipfile.ZipFile(second) as after:
            self.assertEqual(set(before.namelist()), set(after.namelist()))
        wb = load_workbook(second)
        self.addCleanup(wb.close)
        self.assertEqual([18, 20, 21, 22], [i.anchor._from.row for i in wb.active._images])

    def test_missing_or_ambiguous_labels_fail_without_output(self):
        for value in (None, "Review数量分析"):
            wb = self.fixture()
            if value is None:
                wb.active["A21"] = None
            else:
                wb.active["A31"] = value
            source = self.save(wb, "invalid.xlsx")
            with self.assertRaisesRegex(ValueError, "必须唯一"):
                self.insert(source, "absent.xlsx")
            self.assertFalse((self.root / "absent.xlsx").exists())

    def test_missing_image_region_does_not_use_conclusion_area(self):
        wb = self.fixture()
        wb.active.unmerge_cells("B21:F21")
        source = self.save(wb, "invalid-region.xlsx")
        with self.assertRaisesRegex(ValueError, "合并图表区域"):
            self.insert(source, "absent.xlsx")
        self.assertFalse((self.root / "absent.xlsx").exists())

    def test_original_size_floats_at_region_start_even_for_short_row(self):
        wb = self.fixture()
        wb.active.row_dimensions[21].height = 45
        source = self.save(wb, "short.xlsx")
        placed, _ = self.insert(source, "original-size.xlsx")
        wb = load_workbook(placed)
        self.addCleanup(wb.close)
        self.assertEqual(45, wb.active.row_dimensions[21].height)
        for row,image in zip((20,21,22),wb.active._images):
            anchor = image.anchor
            self.assertIsInstance(anchor,OneCellAnchor)
            self.assertEqual((1,row,0,0),(anchor._from.col,anchor._from.row,
                                        anchor._from.colOff,anchor._from.rowOff))
            self.assertEqual((800*9525,200*9525),(anchor.ext.cx,anchor.ext.cy))
            self.assertEqual(self.picture.read_bytes(),image._data())
        self.assertIsNone(wb.active["G21"].value)

    def test_no_growth_is_noop_and_crossing_merge_expands(self):
        wb = self.fixture()
        ws = wb.active
        ws.merge_cells("B16:F18")
        before = [str(a) for a in ws.merged_cells.ranges]
        insert_rows_preserving_merges(ws, 17, 0, 10)
        self.assertEqual(before, [str(a) for a in ws.merged_cells.ranges])
        insert_rows_preserving_merges(ws, 17, 2, 10)
        self.assertIn("B16:F20", ws.merged_cells)


if __name__ == "__main__":
    unittest.main()
