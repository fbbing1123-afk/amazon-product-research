"""统一截取卖家精灵类目概览，原尺寸浮动放图并保留逐项检查点。"""
from __future__ import annotations

import argparse
import asyncio
import base64
import json
import math
import re
import sys
import time
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image as PILImage, ImageStat
from playwright.async_api import async_playwright

MODULES = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(MODULES))
from product_research.scripts.capture_category_overviews import iter_category_rows, clear_existing_category_images, insert_overview_image
from product_research.scripts.fill_custom_template import set_category_monthly_sales

PANEL = '#seller-sprite-best-sellers-box'
CAPTURE_VERSION = 'category-panel-v2-zoom-aware'
CAPTURE_WIDTH = 1600


def read_sales(text):
    match = re.search(r'(?:^|\n)月销量\s*\n\s*([\d,]+)(?:\n|$)', text)
    return int(match[1].replace(',', '')) if match else None


def screenshot_clip(box, scroll, metrics):
    """DOM是CSS像素，CDP裁剪使用DIP；转换浏览器90%等页面缩放。"""
    zoom = metrics.get('visualViewport', {}).get('zoom')
    if zoom is None:
        zoom = metrics['layoutViewport']['clientWidth'] / metrics['cssLayoutViewport']['clientWidth']
    if not math.isfinite(zoom) or zoom <= 0:
        raise ValueError('无法获取有效浏览器缩放比例，停止截图')
    return {'x': (box['x'] + scroll['x']) * zoom, 'y': (box['y'] + scroll['y']) * zoom,
            'width': box['width'] * zoom, 'height': box['height'] * zoom, 'scale': 1 / zoom}


async def capture(page, url, image, postal_code):
    original_viewport = page.viewport_size
    session = await page.context.new_cdp_session(page)
    target, original_style = None, None
    try:
        window = await session.send('Browser.getWindowForTarget')
        await session.send('Browser.setWindowBounds', {'windowId':window['windowId'],'bounds':{'windowState':'normal'}})
        await page.bring_to_front()
        await page.set_viewport_size({'width': 2100, 'height': 1400})
        if page.url != url:
            await page.goto(url, wait_until='domcontentloaded', timeout=45000)
        if any(part in page.url for part in ('/checkpoint/', '/ap/signin', 'validateCaptcha')):
            raise RuntimeError('Amazon需要人工验证或登录，页面已保留')
        location = await page.locator('#glow-ingress-line2').inner_text(timeout=5000)
        if postal_code not in location:
            raise RuntimeError(f'配送邮编不符：{location}')
        panel = page.locator(PANEL)
        await panel.wait_for(state='visible', timeout=30000)
        expand = panel.get_by_text('展开', exact=True)
        if await expand.count() and await expand.is_visible():
            await expand.click()
        # 统一截图前的页面排版，不缩放已生成的图片。
        target = panel.locator('.best-sellers-page')
        await target.wait_for(state='visible', timeout=10000)
        original_style = await target.get_attribute('style')
        await target.evaluate('(e,w)=>{e.style.width=w+"px";e.style.maxWidth="none";window.dispatchEvent(new Event("resize"));}', CAPTURE_WIDTH)
        # 保留插件出现后固定等待20秒的策略。
        await page.wait_for_timeout(20000)
        refreshed, deadline = False, time.monotonic() + 120
        while True:
            text = await panel.inner_text(timeout=5000)
            if '需要登录' in text:
                raise RuntimeError('卖家精灵插件未登录')
            sales = read_sales(text)
            # 产品特征无数据时，插件会初始化空图；其它12项仍须绘制完成。
            charts_ready = await target.locator('.overview').evaluate('''e => {
                const charts = Array.from(e.querySelectorAll('[_echarts_instance_]'));
                return charts.length >= 13 && charts.every(chart =>
                    chart.classList.contains('product-feature') ||
                    Array.from(chart.querySelectorAll('canvas')).some(c => c.width > 0 && c.height > 0));
            }''')
            if sales is not None and charts_ready:
                break
            if time.monotonic() > deadline:
                raise TimeoutError('类目样本或概览图未在120秒内完整加载')
            refresh = panel.get_by_text('立即刷新', exact=True)
            if not refreshed and await refresh.count() and await refresh.is_visible():
                await refresh.click()
                refreshed = True
            print('[类目] 等待月销量及全部概览图加载', flush=True)
            await page.wait_for_timeout(10000)
        tip = panel.get_by_text('知道了', exact=True)
        if await tip.count() and await tip.is_visible():
            await tip.click()
        # 加载等待期间用户可能再次最小化，实际截图前重新唤醒渲染。
        await session.send('Browser.setWindowBounds', {'windowId':window['windowId'],'bounds':{'windowState':'normal'}})
        await page.bring_to_front()
        # Amazon导航菜单可能留下灰色遮罩，按正常点击关闭后再截插件。
        cover = page.locator('#nav-cover')
        if await cover.count() and await cover.is_visible():
            await cover.click(timeout=5000)
            await cover.wait_for(state='hidden', timeout=5000)
        await target.evaluate("e=>e.scrollIntoView({block:'start',inline:'nearest'})")
        await page.mouse.move(0, 0)
        await page.wait_for_timeout(1200)
        box = await target.bounding_box(timeout=5000)
        if not box or abs(box['width'] - CAPTURE_WIDTH) > 2 or box['height'] < 200:
            raise ValueError('插件图表边界不完整或未达到统一宽度')
        scroll = await page.evaluate('({x:scrollX,y:scrollY})')
        metrics = await session.send('Page.getLayoutMetrics')
        print(f"[类目] 采集边界 {box['width']:.0f}×{box['height']:.0f}，浏览器缩放 {metrics['visualViewport'].get('zoom')}",flush=True)
        shot = await asyncio.wait_for(session.send('Page.captureScreenshot', {
            'format':'png', 'fromSurface':True, 'captureBeyondViewport':True,
            'clip':screenshot_clip(box,scroll,metrics),
        }), timeout=20)
        content = base64.b64decode(shot['data'])
        with PILImage.open(BytesIO(content)) as preview:
            stats = ImageStat.Stat(preview.convert('RGB'))
            if max(stats.stddev) < 3:
                raise ValueError('插件截图接近空白')
            if max(stats.mean) < 160:
                raise ValueError('插件白底面板明显变暗，请检查页面遮罩后重试')
            size = preview.size
        # 直接检验浏览器输出，不通过缩放图片来掩盖坐标或像素比错误。
        if any(abs(actual - expected) > 1 for actual, expected in zip(size, (box['width'], box['height']))):
            raise ValueError(f'截图尺寸与面板不符：{size}，停止写入')
        image.write_bytes(content)
        return {'monthly_sales':sales,'image':str(image.resolve()),'pixels':list(size),'text':text,
                'url':page.url,'capture_version':CAPTURE_VERSION,'bounds':box,
                'browser_zoom':metrics['visualViewport'].get('zoom'),
                'captured_at':datetime.now(timezone.utc).isoformat()}
    finally:
        if target is not None and await target.count():
            await target.evaluate('(e,s)=>s===null?e.removeAttribute("style"):e.setAttribute("style",s)', original_style)
        if original_viewport:
            await page.set_viewport_size(original_viewport)
        else:
            await session.send('Emulation.clearDeviceMetricsOverride')
        await session.detach()


async def run(args):
    args.image_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = args.image_dir / 'manifest.json'
    results = json.loads(manifest_path.read_text(encoding='utf-8')) if manifest_path.exists() else {}
    wb = load_workbook(args.input, rich_text=True)
    ws = next(ws for ws in wb if ws.title.endswith('市场调研'))
    items = list(iter_category_rows(ws))[args.start:args.start + args.limit]
    failed = False
    try:
        async with async_playwright() as runtime:
            browser = await runtime.chromium.connect_over_cdp(args.cdp, timeout=20000)
            context = browser.contexts[0]
            for item in items:
                url, started = item['url'], time.monotonic()
                print(f"[类目] 开始 {item['label']}，第{item['row']}行", flush=True)
                try:
                    prior = results.get(url,{})
                    if prior.get('status') == 'ok' and prior.get('capture_version') == CAPTURE_VERSION and Path(prior['image']).exists():
                        result = prior
                    else:
                        page = next((p for p in context.pages if p.url == url),None)
                        if page is None:
                            page = await context.new_page()
                        node = re.search(r'/(\d+)(?:/|$)',url)[1]
                        result = await capture(page,url,args.image_dir/f'{node}.png',args.postal_code)
                        result['status'] = 'ok'
                    if not args.capture_only:
                        # 用户自行缩放并嵌入；此处保持原始像素和原单元格布局。
                        clear_existing_category_images(ws,{item['row']})
                        insert_overview_image(ws,item['row'],Path(result['image']),*result['pixels'])
                        set_category_monthly_sales(ws.cell(item['row'],9),result['monthly_sales'])
                    print(f"[类目] 完成：{result['pixels']}，月销量{result['monthly_sales']}",flush=True)
                except Exception as error:
                    failed = True
                    result = {'status':'failed','error':str(error)}
                    print(f"[类目] 失败：{type(error).__name__}: {error}",flush=True)
                result.update({'row':item['row'],'label':item['label'],'elapsed_seconds':round(time.monotonic()-started,1)})
                results[url] = result
                manifest_path.write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
                if not args.capture_only:
                    args.output.parent.mkdir(parents=True,exist_ok=True)
                    temporary = args.output.with_suffix('.part.xlsx')
                    wb.save(temporary)
                    temporary.replace(args.output)
                    wb.close()
                    wb = load_workbook(args.output,rich_text=True)
                    ws = next(ws for ws in wb if ws.title.endswith('市场调研'))
    finally:
        wb.close()
    return int(failed)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input',type=Path,required=True)
    parser.add_argument('--output',type=Path)
    parser.add_argument('--capture-only',action='store_true')
    parser.add_argument('--image-dir',type=Path,required=True)
    parser.add_argument('--postal-code',default='10001')
    parser.add_argument('--cdp',default='http://127.0.0.1:9223')
    parser.add_argument('--start',type=int,default=0)
    parser.add_argument('--limit',type=int,default=5)
    args = parser.parse_args()
    if not args.capture_only and (not args.output or args.output.resolve() == args.input.resolve()):
        parser.error('写表需要独立--output，或用--capture-only仅截图')
    if not 1 <= args.limit <= 5 or args.start < 0:
        parser.error('单批处理1到5个类目，start须为非负整数')
    return asyncio.run(run(args))


if __name__ == '__main__':
    raise SystemExit(main())
