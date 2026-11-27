"""从当前商品页下载主图原文件，按ASIN校验，不依赖商品数据源或截图。"""
from __future__ import annotations

import argparse
import asyncio
import hashlib
from io import BytesIO
import json
from pathlib import Path
import re
import sys
import tempfile
from urllib.parse import urlparse

from PIL import Image

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from competitor_analysis.collect_basic_facts import asin_from_input
from review_research.download_sellersprite_reviews import now_iso, write_json

IMAGE_HOSTS = {'m.media-amazon.com', 'images-na.ssl-images-amazon.com',
               'images-eu.ssl-images-amazon.com', 'images-fe.ssl-images-amazon.com'}


sys.path.insert(0, str((Path(__file__).resolve().parent / "..").resolve()))
from shared.browser_runtime import get_cdp_url
from shared.amazon_access import (
    AmazonAccessRequired, ensure_amazon_access_async, goto_amazon_async, is_amazon_access_url, pending_amazon_access_page,
)

def valid_image_url(url):
    """只使用商品主图节点实际给出的亚马逊图片地址，不推测或改造网址。"""
    parsed = urlparse(str(url or ''))
    return (parsed.scheme == 'https' and parsed.hostname in IMAGE_HOSTS
            and not parsed.username and not parsed.password
            and '/images/I/' in parsed.path and parsed.port in (None, 443))


def image_candidates(observation):
    """优先原图链接，再取主图动态尺寸中最大的候选，最后保留当前显示图。"""
    choices = []
    if observation.get('hires'):
        choices.append(observation['hires'])
    try:
        dynamic = json.loads(observation.get('dynamic') or '{}')
    except (TypeError, ValueError):
        dynamic = {}
    if isinstance(dynamic, dict):
        sizes = [(url, size) for url, size in dynamic.items()
                 if isinstance(size, list) and len(size) == 2
                 and all(type(n) in (int, float) and n > 0 for n in size)]
        choices.extend(url for url, _ in sorted(sizes, key=lambda item: item[1][0]*item[1][1], reverse=True))
    choices.extend([observation.get('current'), observation.get('src')])
    return list(dict.fromkeys(url for url in choices if valid_image_url(url)))


def validate_asset(asset, asin):
    """回填前核验商品归属、文件指纹与真实像素，拒绝错图或被替换的文件。"""
    if (asset.get('status') != 'ok' or asset.get('asin') != asin
            or asin_from_input(asset.get('page_url', '')) != asin
            or asset.get('observed_asin') != asin
            or not valid_image_url(asset.get('image_url'))):
        raise ValueError(f'{asin}主图的网页归属无效')
    payload = Path(asset['image_path']).read_bytes()
    if hashlib.sha256(payload).hexdigest() != asset.get('sha256'):
        raise ValueError(f'{asin}主图文件已变化')
    with Image.open(BytesIO(payload)) as picture:
        if list(picture.size) != [asset.get('width'), asset.get('height')] or picture.format not in ('JPEG', 'PNG'):
            raise ValueError(f'{asin}主图格式或像素不一致')
        picture.verify()
    return payload


async def selected_asin(page):
    """Amazon可能重复渲染购买表单，只有全部ASIN值一致才接受。"""
    values = await page.locator('#ASIN').evaluate_all('els => els.map(el => el.value)')
    distinct = set(values)
    return next(iter(distinct)) if values and len(distinct) == 1 else None


async def collect_main_image(page, asin, directory, timeout_ms=15000, postal_code=None):
    """处理确认后读取当前选中ASIN主图，不重新触发评论下载。"""
    asin = asin_from_input(asin)
    if is_amazon_access_url(page.url):
        await ensure_amazon_access_async(page, expected_url=f'https://www.amazon.com/dp/{asin}')
    if asin_from_input(page.url) != asin:
        raise ValueError('主图页面ASIN与目标不一致')
    await ensure_amazon_access_async(page, expected_url=f'https://www.amazon.com/dp/{asin}')
    hero = page.locator('#landingImage')
    await hero.wait_for(state='visible', timeout=timeout_ms)
    if await hero.count() != 1:
        raise ValueError('商品主图节点不唯一')
    observed_asin = await selected_asin(page)
    if observed_asin != asin:
        raise ValueError('页面当前选中的商品与目标ASIN不一致')
    observation = await hero.evaluate('''el => ({hires:el.getAttribute('data-old-hires'),
        dynamic:el.getAttribute('data-a-dynamic-image'),src:el.getAttribute('src'),current:el.currentSrc,alt:el.alt})''')
    postal = page.locator('#glow-ingress-line2')
    postal_text = (await postal.inner_text()).strip() if await postal.count() == 1 else ''
    if postal_code and postal_code not in postal_text:
        raise ValueError(f'主图页面配送邮编不符：{postal_text}')
    rating_node = page.locator('#acrPopover')
    count_node = page.locator('#acrCustomerReviewText')
    rating_text = await rating_node.first.get_attribute('title') if await rating_node.count() else ''
    count_values = set(await count_node.all_inner_texts())
    count_text = next(iter(count_values)).strip() if len(count_values) == 1 else ''
    rating = re.search(r'([0-5](?:\.\d+)?)\s+out of\s+5', rating_text or '')
    count = re.fullmatch(r'\(?([\d,]+)\)?(?:\s+(?:ratings?|reviews?))?', count_text, re.I)
    errors = []
    best = None
    for url in image_candidates(observation):
        response = None
        try:
            response = await page.context.request.get(url, timeout=timeout_ms, headers={'Accept':'image/jpeg,image/png'})
            if not response.ok or not valid_image_url(response.url):
                raise ValueError(f'图片响应异常：{response.status}')
            raw = await response.body()
            if len(raw) > 25*1024*1024:
                raise ValueError('图片文件超出25MB限制')
            with Image.open(BytesIO(raw)) as picture:
                picture.load()
                width, height = picture.size
                if min(width, height) < 100:
                    raise ValueError('主图尺寸异常或为占位图')
                if picture.format in ('JPEG', 'PNG'):
                    payload, extension = raw, 'jpg' if picture.format == 'JPEG' else 'png'
                else:
                    # Excel不支持的格式仅转为无损PNG，保持全部像素，不重采样。
                    buffer = BytesIO()
                    picture.save(buffer, format='PNG')
                    payload, extension = buffer.getvalue(), 'png'
            candidate = (width*height, payload, extension, width, height, url, hashlib.sha256(raw).hexdigest())
            if best is None or candidate[0] > best[0]:
                best = candidate
            # 原图链接返回足够清晰的文件后停止，避免下载所有缩略图。
            if max(width, height) >= 1000:
                break
        except Exception as error:
            errors.append(str(error))
        finally:
            if response is not None:
                await response.dispose()
    if best is None:
        raise ValueError('商品页主图获取失败：'+'；'.join(errors))
    if asin_from_input(page.url) != asin or await selected_asin(page) != asin:
        raise ValueError('下载主图期间页面商品发生变化')
    _, payload, extension, width, height, url, original_hash = best
    digest = hashlib.sha256(payload).hexdigest()
    directory = Path(directory).resolve()
    directory.mkdir(parents=True, exist_ok=True)
    path = directory/f'{asin}-{digest[:16]}.{extension}'
    if path.exists() and path.read_bytes() != payload:
        raise ValueError('主图输出路径存在不同内容')
    path.write_bytes(payload)
    result = {'status':'ok', 'asin':asin, 'observed_asin':observed_asin, 'page_url':page.url,
              'image_url':url, 'image_path':str(path), 'sha256':digest, 'download_sha256':original_hash,
              'width':width, 'height':height, 'high_resolution':max(width,height)>=1000,
              'observed_at':now_iso(), 'postal_code_text':postal_text, 'alt':observation.get('alt'),
              'rating':float(rating[1]) if rating else None,
              'review_count':int(count[1].replace(',','')) if count else None}
    validate_asset(result, asin)
    return result


async def collect_batch(asins, output_dir, endpoint, postal_code='10001'):
    """独立图片入口，目标可以完全不在商品数据源里，失败页保留给用户。"""
    from playwright.async_api import async_playwright
    output_dir.mkdir(parents=True, exist_ok=True)
    directory = Path(tempfile.mkdtemp(prefix='product-images-', dir=output_dir)).resolve()
    manifest = {'schema':'competitor-product-images.v1', 'products':{}}
    async with async_playwright() as runtime:
        browser = await runtime.chromium.connect_over_cdp(endpoint, timeout=15000)
        context = browser.contexts[0]
        for asin in asins:
            pending = pending_amazon_access_page(context)
            if pending is not None:
                try:
                    await ensure_amazon_access_async(pending)
                except AmazonAccessRequired as error:
                    manifest['products'][asin] = {'asin':asin, 'status':'needs_user_action', 'error':str(error)}
                    write_json(directory/'images.json', manifest)
                    break
            page = next((p for p in context.pages if re.search(r'/(?:dp|gp/product)/'+asin+r'(?:[/?]|$)', p.url)), None)
            created = page is None
            if created:
                page = await context.new_page()
            try:
                if created:
                    await goto_amazon_async(page, f'https://www.amazon.com/dp/{asin}')
                result = await collect_main_image(page, asin, directory, postal_code=postal_code)
                if created:
                    await page.close()
            except Exception as error:
                result = {'asin':asin, 'status':'needs_user_action' if isinstance(error, AmazonAccessRequired) else 'failed',
                          'error':str(error), 'page_url':page.url}
            manifest['products'][asin] = result
            write_json(directory/'images.json', manifest)
            print(json.dumps(result, ensure_ascii=False), flush=True)
            if result['status'] == 'needs_user_action':
                break
    print(f'图片清单：{directory / "images.json"}', flush=True)
    return 0 if all(a['status']=='ok' for a in manifest['products'].values()) else 2


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('items', nargs='+', help='ASIN或美国站商品链接，不要求在数据源中')
    parser.add_argument('--output-dir', type=Path, required=True)
    parser.add_argument('--cdp', default=get_cdp_url())
    parser.add_argument('--postal-code', default='10001')
    args = parser.parse_args()
    return asyncio.run(collect_batch(list(dict.fromkeys(asin_from_input(v) for v in args.items)), args.output_dir, args.cdp,
                                   args.postal_code))


if __name__ == '__main__':
    raise SystemExit(main())
