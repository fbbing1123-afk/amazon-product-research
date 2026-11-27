from __future__ import annotations

import tempfile
import unittest
import zipfile
import asyncio
from unittest.mock import AsyncMock, Mock, patch
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PillowImage, ImageDraw

from modules.product_research.scripts.fill_custom_template import (
    apply_category_share_display, collect_category_summaries, fill_category_section,
    set_category_monthly_sales, write_missing_category_translation_report,
)
from modules.product_research.scripts.refresh_category_shares import (
    collect_updates, refresh_category_shares,
)
from modules.product_research.scripts.capture_sellersprite_categories import screenshot_clip, stable_screenshot
from modules.product_research.scripts.refresh_sellersprite_category_images import refresh_images
from modules.product_research.scripts.capture_category_overviews import iter_category_rows
import json


class CategoryShareTests(unittest.TestCase):
    def ranked_fixture(self, categories):
        source = Workbook()
        sheet = source.active
        sheet.title = '数据源'
        sheet.append(['类目路径', '大类目', '月销量'])
        for count in range(1, categories + 1):
            for _ in range(count):
                # 小类目1销量最高，检验展示顺序依据商品数量而不是销量。
                sheet.append([f'Electronics:Category{count}', 'Electronics', 100000 if count == 1 else 1])
                sheet.cell(sheet.max_row, 1).hyperlink = f'https://www.amazon.com/b?node={count}'
        sheet.append(['', 'Electronics', 1])
        report = Workbook()
        research = report.active
        research.title = '市场调研'
        research['A1'] = '2. 所属平台分类及常用分类节点'
        research['A2'], research['I2'] = '一级类目', '月总销量'
        for row in range(3, 15):
            research.cell(row, 1, '旧类目')
            research.cell(row, 11, '旧末级').hyperlink = 'https://www.amazon.com/b?node=999'
        research['A15'], research['G16'] = '3. 数据分析概览', '保留下方内容'
        return report, source

    def test_only_top_ten_are_displayed_and_scheduled_with_full_sample_shares(self):
        report, source = self.ranked_fixture(12)
        source_values = list(source.active.values)
        fill_category_section(report, source)
        research = report.active
        self.assertEqual([f'Category{count}' for count in range(12, 2, -1)],
                         [str(research.cell(row, 2).value).splitlines()[0] for row in range(3, 13)])
        self.assertEqual('12', list(iter_category_rows(research))[0]['url'].rstrip('/').split('/')[-1])
        self.assertEqual(10, len(list(iter_category_rows(research))))
        self.assertTrue(all(research.cell(row, column).value is None
                            for row in (13, 14) for column in range(1, 12)))
        self.assertEqual('3. 数据分析概览', research['A15'].value)
        self.assertEqual('保留下方内容', research['G16'].value)
        self.assertEqual(source_values, list(source.active.values))
        self.assertEqual(12, len(collect_category_summaries(source.active)))
        _, _, _, updates = collect_updates(report, source)
        self.assertAlmostEqual(12 / 78, updates[0]['share'])
        self.assertLess(sum(update['share'] for update in updates), 1)
        self.assertIn('15.4%', research['I3'].value)
        report.close()
        source.close()

    def test_ten_or_fewer_categories_are_all_displayed_and_empty_source_has_no_task(self):
        for categories in (0, 3, 10):
            with self.subTest(categories=categories):
                report, source = self.ranked_fixture(categories)
                fill_category_section(report, source)
                self.assertEqual(categories, len(list(iter_category_rows(report.active))))
                report.close()
                source.close()

    def test_rebuilding_old_layout_removes_category_images_but_preserves_lower_image_and_size(self):
        report, source = self.ranked_fixture(12)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            image_path = root / 'picture.png'
            with PillowImage.new('RGB', (80, 40), 'blue') as image:
                image.save(image_path)
            for row in (3, 14, 16):
                report.active.add_image(ExcelImage(str(image_path)), f'F{row}')
            report.active.row_dimensions[14].height = 140
            report.active.column_dimensions['F'].width = 40
            original, output = root / 'original.xlsx', root / 'output.xlsx'
            report.save(original)
            report.close()
            report = load_workbook(original)
            fill_category_section(report, source)
            report.save(output)
            report.close()
            checked = load_workbook(output)
            self.assertEqual(1, len(checked.active._images))
            self.assertEqual(15, checked.active._images[0].anchor._from.row)
            self.assertEqual(140, checked.active.row_dimensions[14].height)
            self.assertEqual(40, checked.active.column_dimensions['F'].width)
            checked.close()
        source.close()

    def test_missing_translations_only_request_displayed_categories(self):
        report, source = self.ranked_fixture(12)
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / 'translations.json'
            write_missing_category_translation_report(source, output)
            names = {item['english'] for item in json.loads(output.read_text(encoding='utf-8'))['missing_translations']}
            self.assertNotIn('Category1', names)
            self.assertNotIn('Category2', names)
            self.assertIn('Category12', names)
        report.close()
        source.close()

    def test_waits_until_loading_animation_stops(self):
        session=Mock(send=AsyncMock(side_effect=[{'data':s} for s in ['转动1','转动2','静止','静止','静止']]))
        with patch('modules.product_research.scripts.capture_sellersprite_categories.asyncio.sleep',AsyncMock()):
            result=asyncio.run(stable_screenshot(session,{}))
        self.assertEqual(result['data'],'静止')
        self.assertEqual(session.send.await_count,5)

    def test_unfinished_screenshot_is_not_accepted(self):
        async def blocked(*args):
            await asyncio.Future()
        with self.assertRaisesRegex(TimeoutError,'加载动画'):
            asyncio.run(stable_screenshot(Mock(send=blocked),{},timeout=0.02))

    def test_clip_converts_browser_zoom_without_resizing_output(self):
        for zoom in (0.5,0.8,0.9,1,1.25,1.5,2):
            clip=screenshot_clip({'x':300,'y':20,'width':1600,'height':940},
                                 {'x':0,'y':400},{'visualViewport':{'zoom':zoom}})
            self.assertAlmostEqual(300*zoom,clip['x'])
            self.assertAlmostEqual(420*zoom,clip['y'])
            self.assertAlmostEqual(1600,clip['width']*clip['scale'])
            self.assertAlmostEqual(940,clip['height']*clip['scale'])

    def test_clip_supports_legacy_metrics_and_rejects_invalid_zoom(self):
        box = {'x': 20, 'y': 50, 'width': 1600, 'height': 940}
        scroll = {'x': 100, 'y': 250}
        clip = screenshot_clip(box, scroll, {
            'visualViewport': {}, 'layoutViewport': {'clientWidth': 1800},
            'cssLayoutViewport': {'clientWidth': 2000},
        })
        self.assertAlmostEqual(108, clip['x'])
        self.assertAlmostEqual(270, clip['y'])
        self.assertAlmostEqual(1600, clip['width'] * clip['scale'])
        for zoom in (0, -1, float('nan'), float('inf')):
            with self.assertRaises(ValueError):
                screenshot_clip(box, scroll, {'visualViewport': {'zoom': zoom}})

    def test_sales_writer_migrates_old_format_and_preserves_zero(self):
        cell=Workbook().active['I3']
        cell.value=30524
        cell.number_format='#,##0"\n数量占比：91.2%"'
        set_category_monthly_sales(cell,0)
        self.assertEqual('0\n数量占比：91.2%',cell.value)
        self.assertEqual('General',cell.number_format)
        set_category_monthly_sales(cell,12345)
        self.assertEqual('12,345\n数量占比：91.2%',cell.value)
        apply_category_share_display(cell,0.5004)
        set_category_monthly_sales(cell,54321)
        self.assertEqual('54,321\n数量占比：50.0%',cell.value)
        self.assertEqual('D9EAF7',cell.fill.fgColor.rgb[-6:])
        for missing in ('',None,'/'):
            cell.value=missing
            apply_category_share_display(cell,0)
            self.assertEqual('/\n数量占比：0.0%',cell.value)

    def test_image_refresh_keeps_native_size_and_other_parts(self):
        report,source=self.fixture()
        data=report.create_sheet('数据源')
        for row in source.active:
            for cell in row:
                data.cell(cell.row,cell.column,cell.value)
                if cell.hyperlink:
                    data.cell(cell.row,cell.column).hyperlink=cell.hyperlink.target
        with tempfile.TemporaryDirectory() as directory:
            root=Path(directory)
            picture=root/'native.png'
            with PillowImage.new('RGB',(1600,940),'white') as pixels:
                ImageDraw.Draw(pixels).rectangle((10,10,1590,930),fill='blue')
                pixels.save(picture)
            results={}
            for row,node in ((3,22),(4,11)):
                image=ExcelImage(str(picture))
                image.width,image.height=300,170
                report.active.add_image(image,f'F{row}')
                url=f'https://www.amazon.com/gp/bestsellers/electronics/{node}'
                results[url]={'status':'ok','monthly_sales':node*100,'image':str(picture)}
            before,after=root/'before.xlsx',root/'after.xlsx'
            report.save(before)
            one = {next(iter(results)): next(iter(results.values()))}
            with self.assertRaisesRegex(ValueError, '所有类目'):
                refresh_images(before, root/'rejected.xlsx', one)
            partial = root/'partial.xlsx'
            refresh_images(before, partial, one, allow_partial=True)
            checked = load_workbook(partial, rich_text=True)
            self.assertEqual('42,095\n数量占比：66.7%', checked.active['I4'].value)
            self.assertEqual(checked.active._images[1].anchor.ext.cx, 300*9525)
            checked.close()
            with zipfile.ZipFile(before) as old, zipfile.ZipFile(partial) as new:
                self.assertEqual(old.read('xl/media/image2.png'), new.read('xl/media/image2.png'))
            refresh_images(before,after,results)
            book=load_workbook(after,rich_text=True)
            self.assertEqual('2,200\n数量占比：33.3%',book.active['I3'].value)
            self.assertEqual('1,100\n数量占比：66.7%',book.active['I4'].value)
            self.assertTrue(all((i.anchor.ext.cx,i.anchor.ext.cy)==(1600*9525,940*9525) for i in book.active._images))
            book.close()
            with zipfile.ZipFile(before) as old,zipfile.ZipFile(after) as new:
                allowed={'xl/worksheets/sheet1.xml','xl/styles.xml','xl/drawings/drawing1.xml','xl/media/image1.png','xl/media/image2.png'}
                self.assertEqual(old.namelist(),new.namelist())
                self.assertTrue({n for n in old.namelist() if old.read(n)!=new.read(n)}<=allowed)
            report.close()
            source.close()

    def fixture(self):
        source = Workbook()
        data = source.active
        data.title = "数据源"
        data.append(["类目路径", "大类目", "月销量"])
        for node, name, sales in [(11, "A", 1), (11, "A", 2), (22, "B", 9999)]:
            data.append([f"Electronics:{name}", "Electronics", sales])
            data.cell(data.max_row, 1).hyperlink = f"https://www.amazon.com/b?node={node}"
        data.append(["", "Electronics", 10000])
        report = Workbook()
        sheet = report.active
        sheet.title = "市场调研"
        sheet["A1"] = "2. 所属平台分类及常用分类节点\n注：蓝色为常见类目 红色为可探索类目"
        sheet["A2"] = "一级类目"
        sheet["I2"] = "月总销量"
        # 报告的行序与数据源相反，必须按节点匹配。
        for row, name, node in [(3, "B", 22), (4, "A", 11)]:
            sheet.cell(row, 1, "Electronics")
            sheet.cell(row, 2, name).hyperlink = f"https://www.amazon.com/gp/bestsellers/electronics/{node}"
            sheet.cell(row, 9, 42095)
        sheet["A5"] = "3. 数据分析概览"
        report.create_sheet("其它页")["A1"] = "保留"
        return report, source

    def test_counts_not_sales_and_invalid_rows_excluded(self):
        report, source = self.fixture()
        _, _, _, updates = collect_updates(report, source)
        self.assertAlmostEqual(1 / 3, updates[0]["share"])
        self.assertAlmostEqual(2 / 3, updates[1]["share"])
        self.assertEqual([1, 2], [u["count"] for u in updates])
        self.assertAlmostEqual(1, sum(s.product_share for s in collect_category_summaries(source.active)))

    def test_unknown_node_does_not_guess_by_name(self):
        report, source = self.fixture()
        report.active["B3"].hyperlink = "https://www.amazon.com/gp/bestsellers/electronics/999"
        with self.assertRaisesRegex(ValueError, "无法唯一匹配"):
            collect_updates(report, source)

    def test_template_and_later_sales_refresh_keep_display(self):
        report, source = self.fixture()
        fill_category_section(report, source)
        sheet = report.active
        self.assertNotIn("蓝色", sheet["A1"].value)
        self.assertIn("注：数量占比按有效类目样本计算", sheet["A1"].value)
        set_category_monthly_sales(sheet["I3"], 12345)
        self.assertEqual('12,345\n数量占比：66.7%', sheet["I3"].value)
        self.assertEqual('General', sheet["I3"].number_format)
        self.assertTrue(sheet["I3"].alignment.wrap_text)
        self.assertEqual("D9EAF7", sheet["I3"].fill.fgColor.rgb[-6:])
        # 真实零销量不转成缺失值，100% 和 0% 均可显示。
        sheet["I3"] = 0
        for share, color in ((0, "FCE4D6"), (0.5, "FCE4D6"), (1, "D9EAF7")):
            apply_category_share_display(sheet["I3"], share)
            self.assertEqual(f'0\n数量占比：{share:.1%}', sheet["I3"].value)
            self.assertEqual(color, sheet["I3"].fill.fgColor.rgb[-6:])

    def test_zip_refresh_preserves_values_other_parts_and_newline(self):
        report, source = self.fixture()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            before, data, after = [root / name for name in ("report.xlsx", "data.xlsx", "output.xlsx")]
            report.save(before)
            source.save(data)
            refresh_category_shares(before, data, after)
            with zipfile.ZipFile(before) as old, zipfile.ZipFile(after) as new:
                self.assertEqual(old.namelist(), new.namelist())
                changed = {name for name in old.namelist() if old.read(name) != new.read(name)}
                self.assertEqual({"xl/worksheets/sheet1.xml", "xl/styles.xml"}, changed)
            reopened = load_workbook(after, rich_text=True)
            self.assertEqual('42,095\n数量占比：33.3%', reopened.active["I3"].value)
            self.assertEqual('s', reopened.active["I3"].data_type)
            self.assertEqual('42,095\n数量占比：66.7%', reopened.active["I4"].value)
            self.assertEqual("FCE4D6", reopened.active["I3"].fill.fgColor.rgb[-6:])
            self.assertEqual("D9EAF7", reopened.active["I4"].fill.fgColor.rgb[-6:])
            reopened.close()
            with self.assertRaisesRegex(ValueError, "不能覆盖"):
                refresh_category_shares(before, data, before)


if __name__ == "__main__":
    unittest.main()
