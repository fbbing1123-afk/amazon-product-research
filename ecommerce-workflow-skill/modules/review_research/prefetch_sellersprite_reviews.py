"""单一CDP会话内预取多商品评论，下载等待不占用截图操作锁。"""
from __future__ import annotations

import asyncio
from pathlib import Path
import re
import time
from shared.amazon_access import (
    AmazonAccessRequired, ensure_amazon_access_async, goto_amazon_async, is_amazon_access_url, pending_amazon_access_page,
)

from review_research.download_sellersprite_reviews import (
    NamedDownload, NeedsUserAction, PANEL, asin_from_input, export_status, now_iso, read_export, write_json,
)


class DownloadRouter:
    """整批使用固定目录，按发起页面与GUID路由，不重设其它商品的下载路径。"""
    def __init__(self, browser, destination, log=None):
        self.browser = browser
        self.destination = destination.resolve()
        self.trackers = {}
        self.session = None
        self.log = log

    async def __aenter__(self):
        self.destination.mkdir(parents=True, exist_ok=True)
        self.session = await self.browser.new_browser_cdp_session()
        self.session.on('Browser.downloadWillBegin', self.on_begin)
        self.session.on('Browser.downloadProgress', self.on_progress)
        try:
            await self.session.send('Browser.setDownloadBehavior', {
                'behavior': 'allow', 'downloadPath': str(self.destination), 'eventsEnabled': True})
        except Exception:
            await self.session.detach()
            raise
        return self

    async def register(self, page, asin):
        session = await page.context.new_cdp_session(page)
        try:
            frame_id = (await session.send('Page.getFrameTree'))['frameTree']['frame']['id']
        finally:
            await session.detach()
        tracker = NamedDownload(page, self.destination)
        tracker.frame_id, tracker.asin = frame_id, asin
        if frame_id in self.trackers:
            raise ValueError('同一商品页不能同时启动重复下载')
        self.trackers[frame_id] = tracker
        return tracker

    def on_begin(self, event):
        tracker = self.trackers.get(event.get('frameId'))
        if tracker is not None:
            tracker.on_begin(event)
            if tracker.filename and not tracker.filename.upper().startswith(tracker.asin + '-'):
                tracker.error = '下载文件名与当前分析商品ASIN不一致'

    def on_progress(self, event):
        for tracker in self.trackers.values():
            previous = tracker.state
            tracker.on_progress(event)
            if tracker.state == 'completed' and previous != 'completed' and self.log:
                self.log('review_file_completed', asin=tracker.asin)

    async def __aexit__(self, *args):
        pending = [t.asin for t in self.trackers.values() if getattr(t, 'started', False) and t.state not in ('completed','canceled')]
        try:
            if pending:
                # 用户要求未导出完成的页保留；后续完成仍落原目录，不能取消下载或改到别处。
                await self.session.send('Browser.setDownloadBehavior', {
                    'behavior':'allow','downloadPath':str(self.destination),'eventsEnabled':False})
                if self.log:
                    self.log('downloads_pending', asins=pending, download_directory=str(self.destination))
            else:
                await self.session.send('Browser.setDownloadBehavior', {'behavior':'default','eventsEnabled':False})
        finally:
            await self.session.detach()


async def check_product(page, asin, postal_code, log=None):
    """处理确认后核验目标、邮编和插件，不重复启动导出。"""
    try:
        await ensure_amazon_access_async(page, expected_url=f'https://www.amazon.com/dp/{asin}', log=log)
    except AmazonAccessRequired as error:
        await page.bring_to_front()
        raise NeedsUserAction(str(error)) from error
    if asin_from_input(page.url) != asin:
        raise ValueError('商品页ASIN与当前下载目标不一致')
    location = await page.locator('#glow-ingress-line2').inner_text(timeout=10000)
    if postal_code not in location:
        raise ValueError(f'配送邮编不符：{location}')
    panel = page.locator(PANEL)
    await panel.wait_for(state='visible', timeout=45000)
    text = await panel.inner_text(timeout=5000)
    if '需要登录' in text:
        await page.bring_to_front()
        raise NeedsUserAction('请先登录当前商品页的卖家精灵插件后定点重试')
    found = re.search(r'ASIN\s*[:：]\s*([A-Z0-9]{10})\b', text, re.I)
    if not found or found[1].upper() != asin:
        raise ValueError('卖家精灵插件ASIN与当前下载目标不一致')
    return panel


async def save_export(path, asin, run_dir, seconds):
    """验证原始文件后保存归属记录；已下载文件可据此恢复，无需再次抓取。"""
    payload = await asyncio.to_thread(read_export, path, asin)
    payload.update(downloaded_at=now_iso(), translation_requested=True, download_seconds=seconds)
    json_path = run_dir / asin / 'reviews.json'
    json_path.parent.mkdir(parents=True, exist_ok=True)
    write_json(json_path, payload)
    return {'asin': asin, 'status': export_status(payload, 'downloaded'), 'source_file': str(path),
            'source_sha256': payload['source_sha256'], 'reviews_file': str(json_path),
            'review_count': payload['review_count'], 'translation': payload['translation'],
            'warnings': payload['warnings'], 'download_seconds': seconds}


async def prefetch_one(context, router, asin, run_dir, ui_lock, started_event, blocked_event,
                       log, timeout=600, postal_code='10001', page=None, collect_facts=None, on_prepared=None):
    tracker, complete, prepared = None, False, False
    try:
        async with ui_lock:
            if blocked_event.is_set():
                return {'asin': asin, 'status': 'not_attempted'}
            pending_page = pending_amazon_access_page(context)
            if pending_page is not None:
                await pending_page.bring_to_front()
                await ensure_amazon_access_async(pending_page, log=log)
            if page is None:
                page = await context.new_page()
                await goto_amazon_async(page, f'https://www.amazon.com/dp/{asin}', log=log)
                log('product_opened', asin=asin)
            # 插件首次注入需要前台页，启动完成后才允许转为后台下载。
            await page.bring_to_front()
            panel = await check_product(page, asin, postal_code, log=log)
            button = panel.get_by_text('评论下载', exact=True)
            await button.wait_for(state='visible', timeout=15000)
            if await button.count() != 1:
                raise ValueError('评论下载入口不唯一')
            await button.hover(timeout=10000)
            translate = page.get_by_role('checkbox', name='翻译标题和内容', exact=True)
            label = page.get_by_text('翻译标题和内容', exact=True)
            await label.wait_for(state='visible', timeout=10000)
            if not await translate.is_checked():
                await label.click(timeout=10000)
            if not await translate.is_checked():
                raise ValueError('翻译标题和内容未勾选')
            tracker = await router.register(page, asin)
            await button.click(timeout=10000)
            tracker.started = True
            began = time.monotonic()
            log('review_started', asin=asin)
            started_event.set()
            if collect_facts:
                try:
                    facts = await collect_facts(page, asin)
                except Exception as error:
                    facts = {'asin':asin,'status':'incomplete','error':str(error)}
                if on_prepared:
                    on_prepared(asin, facts)
                    prepared = True
        # 下载/翻译期间释放前台操作锁；截图可以使用其它页，当前商品页保留。
        deadline, last_progress, last_heartbeat = began + timeout, '', began
        last_activity = began
        while tracker.state != 'completed':
            if is_amazon_access_url(page.url):
                async with ui_lock:
                    await page.bring_to_front()
                    panel = await check_product(page, asin, postal_code, log=log)
                log('review_access_resumed', asin=asin)
            if tracker.error:
                raise ValueError(tracker.error)
            if tracker.state == 'canceled' or page.is_closed():
                raise RuntimeError('商品页关闭或评论下载被取消')
            if time.monotonic() >= deadline:
                raise TimeoutError(f'评论下载超过{timeout:g}秒，未重复点击')
            try:
                text = await panel.inner_text(timeout=2000)
                progress = re.search(r'下载中\s*\.{0,3}\s*\d+%', text)
                if progress and progress[0] != last_progress:
                    last_progress = progress[0]
                    last_activity = time.monotonic()
                    log('review_progress', asin=asin, progress=last_progress)
            except Exception:
                # 后台页偶发DOM超时不能冒充下载失败，以浏览器下载事件为准。
                pass
            if time.monotonic() - last_activity >= min(180, timeout):
                raise TimeoutError(f'评论下载连续{min(180,timeout):g}秒无新进度，保留页面及已成功项')
            if time.monotonic() - last_heartbeat >= 20:
                log('review_waiting', asin=asin, waiting_seconds=round(time.monotonic()-began, 1))
                last_heartbeat = time.monotonic()
            await asyncio.sleep(0.3)
        path = tracker.completed_path()
        async with ui_lock:
            # 文件完成与页面恢复分别核验，不能把错商品的插件状态带到源文件归属。
            await check_product(page, asin, postal_code, log=log)
        if asin_from_input(page.url) != asin:
            raise ValueError('下载期间商品页被切换，拒绝归属不明的文件')
        seconds = round(time.monotonic() - began, 3)
        result = await save_export(path, asin, run_dir, seconds)
        complete = True
        log('review_finished', asin=asin, status=result['status'], count=result['review_count'])
        return result
    except Exception as error:
        # 导航超时或插件等待失败也可能由验证页引起，不能继续开启下一批。
        if (page is not None and not page.is_closed() and is_amazon_access_url(page.url)
                and not isinstance(error, (AmazonAccessRequired, NeedsUserAction))):
            error = NeedsUserAction('Amazon需要登录或人工验证，当前页面已保留')
        if isinstance(error, AmazonAccessRequired):
            error = NeedsUserAction(str(error))
        if isinstance(error, NeedsUserAction) or (tracker is not None and getattr(tracker, 'started', False)
                                                and tracker.state not in ('completed', 'canceled')):
            blocked_event.set()
        status = 'needs_user_action' if isinstance(error, NeedsUserAction) else 'failed'
        log('review_failed', asin=asin, status=status, error=str(error))
        return {'asin': asin, 'status': status, 'error': str(error)}
    finally:
        if on_prepared and not prepared:
            on_prepared(asin, {'asin':asin,'status':'incomplete','error':'商品取数或导出尚未启动成功'})
        # 只在文件保存、归属与内容校验成功后关闭；未完成和出错页面保留给用户。
        if complete and page is not None and not page.is_closed():
            await page.close()
