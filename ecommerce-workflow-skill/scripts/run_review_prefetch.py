"""先预取评论，下载等待期间截图和填数；单一浏览器连接、单一工作簿写入顺序。"""
from __future__ import annotations

import argparse
import asyncio
import hashlib
import json
from pathlib import Path
import sys
import tempfile
import time

from openpyxl import load_workbook
from playwright.async_api import async_playwright

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'modules'))
from competitor_analysis.populate_basics import competitor_asins
from competitor_analysis.prefetch_basics import source_candidates, collect_same_page, refresh_basic_cells
from competitor_analysis.collect_product_images import collect_main_image
from competitor_analysis.refresh_product_images import refresh_product_images
from product_research.scripts.capture_category_overviews import iter_category_rows
from product_research.scripts.capture_sellersprite_categories import capture
from product_research.scripts.populate_market_summary import refresh_summary
from product_research.scripts.refresh_sellersprite_category_images import refresh_images
from review_research.download_sellersprite_reviews import asin_from_input, now_iso, write_json
from review_research.prefetch_sellersprite_reviews import DownloadRouter, prefetch_one
from shared.browser_runtime import get_cdp_url
from shared.amazon_access import AmazonAccessRequired, pending_amazon_access_page
from shared.quality_gates import validate_category_research, validate_zip_package


def read_plan(args):
    """读取已有布局检查点；试跑允许显式选择少量竞品和类目，默认处理全部。"""
    book = load_workbook(args.workbook, rich_text=True)
    try:
        competitors = competitor_asins(next(s for s in book if s.title.endswith('竞品分析')))
        asins = list(dict.fromkeys(asin_from_input(a) for a in args.asins)) if args.asins else competitors
        if not asins or not set(asins) <= set(competitors):
            raise ValueError('预取商品必须来自当前竞品页')
        categories = list(iter_category_rows(next(s for s in book if s.title.endswith('市场调研'))))
        end = args.category_start + args.category_limit if args.category_limit is not None else None
        selected = [] if args.skip_categories else categories[args.category_start:end]
        if not selected and not args.skip_categories:
            raise ValueError('当前类目范围为空')
        return asins, selected, len(selected) != len(categories)
    finally:
        book.close()


def overlap_evidence(events):
    """用真实开始/下载完成时间证明重叠，不推算未经对照实测的提速百分比。"""
    starts, spans = {}, []
    for item in events:
        if item['event'] == 'review_started':
            starts[item['asin']] = item['seconds']
        elif item['event'] == 'review_file_completed' and item['asin'] in starts:
            spans.append((starts[item['asin']], item['seconds'], item['asin']))
    return [{**item, 'pending_asins': [asin for start, end, asin in spans if start <= item['seconds'] <= end]}
            for item in events if item['event'] in ('metrics_started', 'metrics_finished', 'basics_write_started', 'basics_write_finished', 'category_capture')
            and any(start <= item['seconds'] <= end for start, end, _ in spans)]


async def execute(args):
    asins, categories, partial = read_plan(args)
    candidates = await asyncio.to_thread(source_candidates, args.workbook, asins)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    directory = Path(tempfile.mkdtemp(prefix='review-prefetch-', dir=args.output_dir)).resolve()
    print(f'试跑目录：{directory}', flush=True)
    manifest = {'schema': 'review-prefetch.v1', 'input': str(args.workbook.resolve()),
                'input_sha256': hashlib.sha256(args.workbook.read_bytes()).hexdigest(),
                'started_at': now_iso(), 'asins': asins, 'categories': categories,
                'reviews': {}, 'screenshots': {}, 'errors': [], 'status': 'running',
                'skip_metrics': args.skip_metrics, 'skip_categories': args.skip_categories}
    events, began = [], time.monotonic()
    def log(event, **fields):
        entry = {'event': event, 'seconds': round(time.monotonic()-began, 3), **fields}
        events.append(entry)
        with (directory/'events.jsonl').open('a', encoding='utf-8') as handle:
            handle.write(json.dumps(entry, ensure_ascii=False) + '\n')
        print(json.dumps(entry, ensure_ascii=False), flush=True)
    def checkpoint():
        write_json(directory/'manifest.json', manifest)
    checkpoint()
    ui_lock, ready, blocked = asyncio.Lock(), asyncio.Event(), asyncio.Event()
    launch_limit = args.max_downloads or 2
    manifest['max_downloads'] = launch_limit
    slots = asyncio.Semaphore(launch_limit)
    prepared, initial_prepared, products = asyncio.Event(), asyncio.Event(), {}
    def on_prepared(asin, facts):
        products[asin] = facts
        write_json(directory/'basic-facts.json', {'products':list(products.values())})
        log('product_prepared', asin=asin, status=facts.get('status'))
        if len(products) >= min(launch_limit, len(asins)):
            initial_prepared.set()
        if len(products)==len(asins):
            prepared.set()
    async def collect_facts(page, asin):
        # 复用正在导出的商品页，主图独立取证，基础字段失败也保留已取到的图片。
        try:
            asset = await collect_main_image(page, asin, directory/'product-images')
        except Exception as error:
            asset = {'asin':asin, 'status':'failed', 'error':str(error)}
        try:
            facts = await collect_same_page(page, asin, candidates[asin])
        except Exception as error:
            facts = {'asin':asin, 'status':'incomplete', 'error':str(error)}
        facts['main_image'] = asset
        return facts
    (directory/'categories').mkdir()
    try:
        async with async_playwright() as runtime:
            browser = await runtime.chromium.connect_over_cdp(args.cdp, timeout=20000)
            context = browser.contexts[0]
            async with DownloadRouter(browser, directory/'downloads', log) as router:
                async def review_job(asin):
                    # 开页和导出共用槽位；已有商品保存完成后才补下一个。
                    async with slots:
                        result = await prefetch_one(context, router, asin, directory, ui_lock, ready, blocked,
                                                    log, args.download_timeout, args.postal_code,
                                                    collect_facts=collect_facts, on_prepared=on_prepared)
                        manifest['reviews'][asin] = result
                        checkpoint()
                async def reviews():
                    try:
                        await asyncio.gather(*(review_job(asin) for asin in asins))
                    finally:
                        # 首个启动失败也不能令本地工作无限等待。
                        ready.set()
                        prepared.set()
                        initial_prepared.set()
                async def metrics():
                    # 首个导出启动即可计算本地指标；竞品回填等待本批证据齐备。
                    await ready.wait()
                    market_input = args.workbook
                    if not args.skip_metrics:
                        log('metrics_started')
                        result = await asyncio.to_thread(refresh_summary, args.workbook, directory/'01-metrics.xlsx', directory/'market')
                        manifest['metrics'] = result
                        market_input = Path(result['output'])
                        log('metrics_finished', sample_count=result['sample_count'])
                        checkpoint()
                    await prepared.wait()
                    # 同一写入流程先整体指标再竞品数值，原中文卖点和对比不被原文覆盖。
                    valid = {asin:fact for asin,fact in products.items() if all(fact.get('fields',{}).get(k,{}).get('value') is not None
                             for k in ('monthly_sales','price','listed_at'))}
                    if valid:
                        log('basics_write_started')
                        manifest['basics'] = await asyncio.to_thread(refresh_basic_cells,market_input,directory/'01-basics.xlsx',valid)
                        market_input = Path(manifest['basics']['output'])
                        log('basics_write_finished',asins=list(valid))
                        checkpoint()
                    assets = {asin:fact['main_image'] for asin,fact in products.items() if fact.get('main_image',{}).get('status')=='ok'}
                    if assets:
                        manifest['product_images'] = await asyncio.to_thread(
                            refresh_product_images, market_input, directory/'01-product-images.xlsx', assets,
                            display_mode=getattr(args,'product_image_display','original'), allow_partial=True)
                        log('product_images_written',asins=list(assets))
                        checkpoint()
                async def screenshots():
                    await initial_prepared.wait()
                    for item in categories:
                        async with ui_lock:
                            if blocked.is_set():
                                break
                            pending_page = pending_amazon_access_page(context)
                            if pending_page is not None:
                                blocked.set()
                                await pending_page.bring_to_front()
                                log('amazon_access_paused', url=pending_page.url)
                                break
                            log('category_started', row=item['row'], url=item['url'])
                            page = next((p for p in context.pages if p.url == item['url']), None)
                            created, keep_page = page is None, False
                            if page is None:
                                page = await context.new_page()
                            try:
                                node = item['url'].rstrip('/').split('/')[-1]
                                result = await capture(page, item['url'], directory/f'categories/{node}.png', args.postal_code,
                                                       on_capture=lambda: log('category_capture', row=item['row'], url=item['url']))
                                result.update(status='ok', row=item['row'], label=item['label'])
                                log('category_finished', row=item['row'])
                            except Exception as error:
                                result = {'status': 'failed', 'error': str(error), 'row': item['row']}
                                if isinstance(error, AmazonAccessRequired) or any(part in page.url for part in ('/checkpoint/', '/ap/signin', 'validateCaptcha')) or '未登录' in str(error):
                                    blocked.set()
                                    keep_page = True
                                    result['status'] = 'needs_user_action'
                                log('category_failed', row=item['row'], error=str(error))
                            finally:
                                # 只清理本轮创建的任务页，人工验证页保留，用户原有页面不自动关闭。
                                if created and not keep_page and not page.is_closed():
                                    await page.close()
                            manifest['screenshots'][item['url']] = result
                            checkpoint()
                # 所有任务结束前不恢复下载设置，局部失败也等待其它已启动任务安全收尾。
                results = await asyncio.gather(reviews(), metrics(), screenshots(), return_exceptions=True)
                manifest['errors'] = [str(result) for result in results if isinstance(result, BaseException)]
    except Exception as error:
        manifest['errors'].append(str(error))
        log('stage_failed', error=str(error))
    checkpoint()
    capture_ok = (len(manifest['screenshots']) == len(categories)
                  and all(r['status']=='ok' for r in manifest['screenshots'].values()))
    if categories and (manifest.get('metrics') or args.skip_metrics) and capture_ok:
        try:
            log('image_write_started')
            market_input = (Path(manifest['product_images']['output']) if manifest.get('product_images') else
                            Path(manifest['basics']['output']) if manifest.get('basics') else
                            args.workbook if args.skip_metrics else directory/'01-metrics.xlsx')
            result = await asyncio.to_thread(refresh_images, market_input, directory/'02-market-checkpoint.xlsx',
                                             manifest['screenshots'], allow_partial=partial)
            gates = [validate_zip_package(Path(result['output'])), validate_category_research(Path(result['output']))]
            manifest['market_checkpoint'] = result
            manifest['gates'] = [g.as_dict() for g in gates]
            log('image_write_finished')
        except Exception as error:
            manifest['errors'].append(str(error))
    reviews_ok = len(manifest['reviews']) == len(asins) and all(
        r['status'] in ('downloaded', 'incomplete_translation') for r in manifest['reviews'].values())
    manifest['access_paused'] = blocked.is_set()
    manifest['remaining_asins'] = [asin for asin in asins if manifest['reviews'].get(asin, {}).get('status')
                                   not in ('downloaded', 'incomplete_translation')]
    manifest['overlap_evidence'] = overlap_evidence(events)
    manifest['products'] = products
    facts_ok = len(products)==len(asins) and all(p.get('status')=='complete' for p in products.values())
    images_ok = len(products)==len(asins) and all(p.get('main_image',{}).get('status')=='ok' for p in products.values()) and bool(manifest.get('product_images'))
    market_ok = bool(manifest.get('market_checkpoint')) if categories else (args.skip_metrics or bool(manifest.get('metrics')))
    manifest['status'] = ('completed_with_warnings' if any(r['status']=='incomplete_translation' for r in manifest['reviews'].values())
                          else 'completed') if reviews_ok and facts_ok and images_ok and capture_ok and market_ok and not manifest['errors'] else 'incomplete'
    manifest['elapsed_seconds'], manifest['finished_at'] = round(time.monotonic()-began, 3), now_iso()
    checkpoint()
    print(json.dumps({'directory':str(directory), 'status':manifest['status'], 'elapsed_seconds':manifest['elapsed_seconds'],
                      'overlap_events':len(manifest['overlap_evidence'])}, ensure_ascii=False), flush=True)
    return 1 if manifest['status']=='incomplete' else 0


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--workbook', type=Path, required=True)
    parser.add_argument('--output-dir', type=Path, required=True)
    parser.add_argument('--asins', nargs='+')
    parser.add_argument('--category-start', type=int, default=0)
    parser.add_argument('--category-limit', type=int)
    parser.add_argument('--skip-categories', action='store_true', help='定点恢复时复用已完成的类目截图')
    parser.add_argument('--skip-metrics', action='store_true', help='定点恢复时复用已完成的整体指标')
    parser.add_argument('--max-downloads', type=int, default=2, help='同时开页及导出的商品上限，默认2；完成后再补下一个')
    parser.add_argument('--download-timeout', type=float, default=600, help='每商品总等待上限；连续180秒无新进度提前退出')
    parser.add_argument('--postal-code', default='10001')
    parser.add_argument('--cdp', default=get_cdp_url())
    parser.add_argument('--product-image-display', choices=['fit','original'], default='original', help='默认原尺寸浮动放置，控制表内显示尺寸')
    args = parser.parse_args()
    if (args.max_downloads is not None and args.max_downloads < 1) or args.category_start < 0 or (args.category_limit is not None and args.category_limit < 1):
        parser.error('并发下载数须为正数，类目起点非负且数量为正数')
    if not 0 < args.download_timeout < float('inf'):
        parser.error('下载超时须为有限正数')
    return asyncio.run(execute(args))


if __name__ == '__main__':
    raise SystemExit(main())
