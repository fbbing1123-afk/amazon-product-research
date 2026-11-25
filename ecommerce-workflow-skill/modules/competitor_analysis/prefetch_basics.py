"""复用评论导出商品页取基础证据，数据源优先，单独写入竞品数值。"""
from copy import copy
from pathlib import Path
import re
import tempfile
from xml.dom import minidom
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from competitor_analysis.collect_basic_facts import (
    ALL_FIELDS, PANEL, asin_from_input, normalize_alexa_questions,
    normalize_selling_points, parse_plugin_panel, read_source, resolve_product,
)
from competitor_analysis.annotate_review_export import worksheet_part, children, element
from competitor_analysis.populate_basics import competitor_asins
from market_analysis.excel_market_report import find_source_sheet


def source_candidates(workbook, asins):
    """显式选商品源页，避免把合并后的评价数据源误认为商品数据。"""
    book = load_workbook(workbook, read_only=True, data_only=True)
    try:
        sheet = find_source_sheet(book)
        if sheet is None:
            raise ValueError('缺少商品数据源')
        name = sheet.title
    finally:
        book.close()
    return read_source(workbook, asins, name, ALL_FIELDS)


async def collect_same_page(page, asin, candidates):
    """不新建CDP连接，不再次导航；仅缺失字段读取当前商品。"""
    source = resolve_product(asin, candidates, fields=ALL_FIELDS)
    needed = [key for key, item in source['fields'].items() if item['source']=='missing']
    observation = {'url':page.url, 'observed_at':None, 'fields':{}}
    if any(key in needed for key in ('monthly_sales','price','listed_at')):
        # 两次稳定读取，避免插件刚出现时返回占位值。
        previous = None
        for _ in range(20):
            observation = parse_plugin_panel(await page.locator(PANEL).inner_text(timeout=3000), asin, page.url)
            signature = [observation['fields'][key]['value'] for key in needed if key != 'selling_points']
            if signature == previous and all(v is not None for v in signature):
                break
            previous = signature
            await page.wait_for_timeout(500)
    if 'selling_points' in needed:
        root = page.locator('#feature-bullets')
        value = None
        if await root.count()==1 and await root.is_visible():
            expand = page.get_by_role('button', name='See more About this item', exact=True)
            if await expand.count()==1 and await expand.is_visible() and await expand.get_attribute('aria-expanded')!='true':
                await expand.click(timeout=3000)
            if re.search('About this item', await root.inner_text(timeout=3000), re.I):
                value = normalize_selling_points(await root.locator('li').all_inner_texts())
        observation['fields']['selling_points'] = {'value':value,'raw':value,'label':'About this item','source':'amazon_product_page'}
    if asin_from_input(page.url)!=asin:
        raise ValueError('读取竞品证据时页面ASIN发生变化')
    result = resolve_product(asin, candidates, lambda *_: observation, ALL_FIELDS)
    # 推荐问题只作为维度参考，不代表产品支持该功能，也不作为必填数值。
    questions = page.locator('#dpx-rex-nice-widget-container')
    value = None
    if await questions.count()==1 and await questions.is_visible():
        value = normalize_alexa_questions(await questions.locator('#dpx-rex-nile-inline-default-pills-container li').all_inner_texts())
    result['alexa_questions'] = value
    result['page_url'] = page.url
    return result


def refresh_basic_cells(workbook, output, products):
    """按当前ASIN列、字段标签写三项基础数值；中文卖点/对比保留给语义阶段。"""
    if output.exists() or workbook.resolve()==output.resolve():
        raise ValueError('竞品基础数据必须另存新文件')
    book = load_workbook(workbook, rich_text=True)
    try:
        sheet = next(s for s in book if s.title.endswith('竞品分析'))
        asins, updates = competitor_asins(sheet), {}
        labels = {'monthly_sales':('月销量','销量'), 'price':('价格($)','价格()','价格'), 'listed_at':('上架时间',)}
        for asin, product in products.items():
            if asin not in asins or product.get('asin') != asin:
                raise ValueError('竞品基础数据商品归属不一致')
            link_rows = [c.row for row in sheet.iter_rows(min_col=1,max_col=2) for c in row if str(c.value or '').strip()=='详情页链接']
            if len(link_rows)!=1:
                raise ValueError('详情页链接行不唯一')
            links = [c for c in sheet[link_rows[0]] if asin in str(c.value or '')]
            if len(links)!=1:
                raise ValueError('竞品列与详情页链接不一致，停止写入')
            column = links[0].column
            for field, names in labels.items():
                matches = [c for row in sheet.iter_rows(min_col=1,max_col=2) for c in row if str(c.value or '').strip() in names]
                value = product.get('fields',{}).get(field,{}).get('value')
                if len(matches)!=1 or value is None:
                    raise ValueError(f'{asin}的{names[0]}缺失或定位不唯一')
                updates[f'{get_column_letter(column)}{matches[0].row}'] = value
        name = sheet.title
    finally:
        book.close()
    output.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(workbook) as archive, tempfile.TemporaryDirectory(dir=output.parent) as folder:
        part = worksheet_part(archive,name)
        doc = minidom.parseString(archive.read(part))
        rows = {r.getAttribute('r'):r for r in doc.getElementsByTagName('row')}
        for address,value in updates.items():
            row = rows[re.search(r'\d+$',address)[0]]
            cell = next((c for c in children(row,'c') if c.getAttribute('r')==address),None)
            if cell is None:
                cell = element(row,'c',{'r':address})
                col = column_index_from_string(re.match('[A-Z]+',address)[0])
                following = next((c for c in children(row,'c') if column_index_from_string(re.match('[A-Z]+',c.getAttribute('r'))[0])>col),None)
                row.insertBefore(cell,following) if following else row.appendChild(cell)
            for child in list(cell.childNodes):
                cell.removeChild(child)
            if isinstance(value,(int,float)):
                cell.setAttribute('t','n')
                node = element(cell,'v')
                node.appendChild(doc.createTextNode(str(value)))
            else:
                cell.setAttribute('t','inlineStr')
                node = element(cell,'is')
                text = element(node,'t')
                text.appendChild(doc.createTextNode(str(value)))
                node.appendChild(text)
            cell.appendChild(node)
        staged = Path(folder)/'basic.xlsx'
        with ZipFile(staged,'w') as target:
            for info in archive.infolist():
                target.writestr(copy(info),doc.toxml(encoding='utf-8') if info.filename==part else archive.read(info.filename))
        checked = load_workbook(staged,rich_text=True)
        try:
            if any(checked[name][cell].value!=value for cell,value in updates.items()):
                raise ValueError('竞品基础数据回读验证失败')
        finally:
            checked.close()
        staged.replace(output)
    return {'output':str(output),'cells':list(updates),'asins':list(products)}
