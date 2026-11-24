"""从已生成的卖家精灵市场报告截取三张图表，保留真实样本口径。"""
import argparse
import asyncio
import base64
import json
import re
from pathlib import Path
from urllib.parse import urlparse

from PIL import Image, ImageStat
from playwright.async_api import async_playwright


async def capture(args):
    parsed = urlparse(args.report_url)
    if parsed.hostname != 'www.sellersprite.com' or not re.fullmatch(r'/v3/custom-market-research/\d+/\d+', parsed.path):
        raise ValueError('需要卖家精灵自定义市场报告的准确网址')
    args.output_dir.mkdir(parents=True, exist_ok=True)
    manifest = {'report_url': args.report_url, 'loaded_count': args.loaded_count, 'charts': {}}
    async with async_playwright() as runtime:
        browser = await runtime.chromium.connect_over_cdp(args.cdp, timeout=20000)
        context = browser.contexts[0]
        page = next((p for p in context.pages if p.url == args.report_url), None)
        if page is None:
            page = await context.new_page()
            await page.goto(args.report_url, wait_until='domcontentloaded', timeout=45000)
        session = await context.new_cdp_session(page)
        try:
            for key, label in [('review', '评分数分布'), ('rating', '评分值分布'), ('price', '价格分布')]:
                heading = page.locator('strong').filter(has_text=re.compile('^' + label + '$'))
                await heading.wait_for(state='attached', timeout=30000)
                section = heading.locator('xpath=ancestor::*[.//canvas][1]')
                text = await section.inner_text(timeout=5000)
                match = re.search(r'样本商品数\s*:\s*(\d+)', text)
                if not match:
                    raise ValueError(f'{label}缺少实际样本数')
                canvas = section.locator('canvas')
                if await canvas.count() != 1:
                    raise ValueError(f'{label}没有唯一的图表画布')
                await canvas.evaluate("e=>e.scrollIntoView({block:'center'})")
                await page.wait_for_timeout(1000)
                box = await canvas.bounding_box(timeout=5000)
                scroll = await page.evaluate('({x:scrollX,y:scrollY})')
                if not box or min(box['width'],box['height']) < 100:
                    raise ValueError(f'{label}尚未绘制')
                shot = await asyncio.wait_for(session.send('Page.captureScreenshot', {
                    'format':'png', 'captureBeyondViewport':True, 'fromSurface':True,
                    'clip':{'x':box['x']+scroll['x'],'y':box['y']+scroll['y'],
                            'width':box['width'],'height':box['height'],'scale':1},
                }), timeout=20)
                image = args.output_dir / f'{key}.png'
                image.write_bytes(base64.b64decode(shot['data']))
                with Image.open(image) as pixels:
                    if max(ImageStat.Stat(pixels.convert('RGB')).stddev) < 4:
                        raise ValueError(f'{label}图像为空白')
                    size = pixels.size
                manifest['charts'][key] = {'label':label,'sample_count':int(match[1]),'image':str(image.resolve()),'pixels':size}
                print(f'{label}：样本 {match[1]}，图像 {size}',flush=True)
                (args.output_dir/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
        finally:
            await session.detach()
    return manifest


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--report-url',required=True)
    parser.add_argument('--loaded-count',type=int)
    parser.add_argument('--output-dir',type=Path,required=True)
    parser.add_argument('--cdp',default='http://127.0.0.1:9223')
    asyncio.run(capture(parser.parse_args()))


if __name__ == '__main__':
    main()
