"""验证确认恢复、目标归属、有限重试和下载中途恢复；不访问真实Amazon。"""
import asyncio
import os
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import AsyncMock, Mock, patch

from shared import amazon_checkpoint as checkpoint
from shared.amazon_access import (
    AmazonAccessRequired, ensure_amazon_access, ensure_amazon_access_async, goto_amazon, goto_amazon_async,
)


TARGET = 'https://www.amazon.com/dp/B012345678'
GATE = 'https://www.amazon.com/checkpoint/continue'


class Control:
    def __init__(self, page, button=False):
        self.page, self.button, self.first = page, button, self

    def is_visible(self):
        return self.page.controls_visible

    def is_checked(self):
        return self.page.checked

    def is_enabled(self):
        return True

    def check(self, **kwargs):
        self.page.checked = True

    def click(self, **kwargs):
        self.page.clicks += 1
        if self.page.destination:
            self.page.navigate(self.page.destination)


class Page:
    def __init__(self, url=GATE, destination=TARGET):
        self.url, self.destination = url, destination
        self.now, self.clicks, self.checked = 0, 0, False
        self.controls_visible, self.captcha, self.closed = True, False, False
        self.events = []
        self.main_frame = object()
        self.activated = False
        self.session = Mock()
        self.session.send.return_value = {'windowId':1,'bounds':{'windowState':'normal'}}
        self.context = SimpleNamespace(new_cdp_session=Mock(return_value=self.session))

    def bring_to_front(self):
        self.activated = True

    def on(self, name, callback):
        self.events.append(callback)

    def navigate(self, url):
        self.url = url
        for callback in self.events:
            callback(self.main_frame)

    def is_closed(self):
        return self.closed

    def evaluate(self, script):
        return {'body':True, 'captcha':self.captcha} if script == checkpoint._PROBE else self.captcha

    def locator(self, selector):
        return Control(self)

    def get_by_role(self, *args, **kwargs):
        return Control(self, True)

    def wait_for_timeout(self, milliseconds):
        self.now += milliseconds/1000


class AsyncControl(Control):
    async def is_visible(self):
        return super().is_visible()

    async def is_checked(self):
        return super().is_checked()

    async def is_enabled(self):
        return super().is_enabled()

    async def check(self, **kwargs):
        super().check(**kwargs)

    async def click(self, **kwargs):
        super().click(**kwargs)


class AsyncPage(Page):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.session = SimpleNamespace(send=AsyncMock(return_value={
            'windowId':1,'bounds':{'windowState':'normal'}}),detach=AsyncMock())
        self.context = SimpleNamespace(new_cdp_session=AsyncMock(return_value=self.session))

    async def bring_to_front(self):
        self.activated = True

    async def evaluate(self, script):
        return super().evaluate(script)

    def locator(self, selector):
        return AsyncControl(self)

    def get_by_role(self, *args, **kwargs):
        return AsyncControl(self, True)

    async def wait_for_timeout(self, milliseconds):
        super().wait_for_timeout(milliseconds)
        await asyncio.sleep(0)


class SyncTests(unittest.TestCase):
    def run_handler(self, page, **kwargs):
        with patch.object(checkpoint, 'time', SimpleNamespace(monotonic=lambda:page.now)):
            return checkpoint.handle_checkpoint(page, log=Mock(), **kwargs)

    def test_checked_and_restored_once(self):
        page = Page()
        result = self.run_handler(page)
        self.assertEqual(result['status'], 'cleared')
        self.assertTrue(page.checked)
        self.assertEqual(page.clicks, 1)

    def test_minimized_checkpoint_is_activated_before_check(self):
        page = Page()
        page.session.send.return_value = {'windowId':1,'bounds':{'windowState':'minimized'}}
        original = Control.check
        def check(control, **kwargs):
            self.assertTrue(control.page.activated)
            return original(control, **kwargs)
        with patch.object(Control, 'check', check):
            self.assertEqual(self.run_handler(page)['status'], 'cleared')
        page.session.send.assert_any_call('Browser.setWindowBounds', {
            'windowId':1,'bounds':{'windowState':'normal'}})
        page.session.detach.assert_called_once()

    def test_failed_check_records_step_and_does_not_submit(self):
        page = Page()
        with patch.object(Control,'check',side_effect=TimeoutError()):
            result = self.run_handler(page,timeout=0.5)
        self.assertEqual(result['error'],'check:TimeoutError')
        self.assertEqual(page.clicks,0)

    def test_submission_limit_and_no_repeat_until_navigation(self):
        page = Page(destination=None)
        self.assertEqual(self.run_handler(page, timeout=5)['status'], 'unresolved')
        self.assertEqual(page.clicks, 2)
        self.assertTrue(self.run_handler(page)['cached'])
        self.assertEqual(page.clicks, 2)
        page.navigate(GATE)
        page.destination = TARGET
        self.assertEqual(self.run_handler(page)['status'], 'cleared')
        self.assertEqual(page.clicks, 3)

    def test_missing_controls_finish_within_budget(self):
        page = Page()
        page.controls_visible = False
        self.assertEqual(self.run_handler(page, timeout=0.5)['status'], 'unresolved')
        self.assertLessEqual(page.now, 0.5)
        self.assertEqual(page.clicks, 0)

    def test_unsupported_pages_never_click(self):
        for url, status in [(TARGET, 'normal'), ('https://www.amazon.com/ap/signin', 'login_required'),
                            ('https://www.amazon.com/errors/validateCaptcha', 'other_verification'),
                            ('https://amazon.com.evil.example/checkpoint/continue', 'other_site')]:
            with self.subTest(url=url):
                page = Page(url=url)
                self.assertEqual(self.run_handler(page)['status'], status)
                self.assertEqual(page.clicks, 0)

    def test_captcha_on_checkpoint_is_preserved(self):
        page = Page()
        page.captcha = True
        self.assertEqual(self.run_handler(page)['status'], 'other_verification')
        self.assertEqual(page.clicks, 0)

    def test_wrong_target_after_clear_is_reentered_once(self):
        for destination in ('https://www.amazon.com/', 'https://www.amazon.com/dp/B087654321'):
            page = Page(destination=destination)
            page.goto = Mock(side_effect=lambda url,**kwargs:page.navigate(url))
            with patch.object(checkpoint, 'time', SimpleNamespace(monotonic=lambda:page.now)):
                result = ensure_amazon_access(page, expected_url=TARGET, log=Mock())
            self.assertEqual(result['status'],'cleared')
            self.assertTrue(result['target_reentered'])
            self.assertEqual(page.url,TARGET)
            page.goto.assert_called_once_with(TARGET,wait_until='domcontentloaded',timeout=45000)

    def test_reentry_cannot_loop_when_redirect_still_wrong(self):
        for second_gate in (False,True):
            page = Page(destination='https://www.amazon.com/')
            page.goto = Mock(side_effect=lambda url,**kwargs:page.navigate(GATE if second_gate else 'https://www.amazon.com/'))
            with patch.object(checkpoint,'time',SimpleNamespace(monotonic=lambda:page.now)):
                with self.assertRaises(AmazonAccessRequired):
                    ensure_amazon_access(page,expected_url=TARGET,log=Mock())
            page.goto.assert_called_once()

    def test_normal_wrong_page_and_missing_target_do_not_trigger_reentry(self):
        page = Page(url='https://www.amazon.com/')
        page.goto = Mock()
        with self.assertRaises(AmazonAccessRequired):
            ensure_amazon_access(page,expected_url=TARGET,log=Mock())
        page.goto.assert_not_called()
        page = Page(destination='https://www.amazon.com/')
        page.goto = Mock()
        with patch.object(checkpoint,'time',SimpleNamespace(monotonic=lambda:page.now)):
            self.assertEqual(ensure_amazon_access(page,log=Mock())['status'],'cleared')
        page.goto.assert_not_called()

    def test_reentry_stops_on_navigation_failure_or_login(self):
        for failure in ('timeout','login'):
            page = Page(destination='https://www.amazon.com/')
            page.goto = Mock(side_effect=TimeoutError() if failure == 'timeout' else
                             lambda url,**kwargs:page.navigate('https://www.amazon.com/ap/signin'))
            with patch.object(checkpoint,'time',SimpleNamespace(monotonic=lambda:page.now)):
                with self.assertRaises(AmazonAccessRequired):
                    ensure_amazon_access(page,expected_url=TARGET,log=Mock())
            page.goto.assert_called_once()

    def test_reentry_never_navigates_to_external_or_challenge_target(self):
        for target in ('https://example.com/','https://www.amazon.com/ap/signin',GATE):
            page = Page(destination='https://www.amazon.com/')
            page.goto = Mock()
            with patch.object(checkpoint,'time',SimpleNamespace(monotonic=lambda:page.now)):
                with self.assertRaises(AmazonAccessRequired):
                    ensure_amazon_access(page,expected_url=target,log=Mock())
            page.goto.assert_not_called()

    def test_target_matching_keeps_identity_and_allows_tracking(self):
        self.assertTrue(checkpoint.same_amazon_target(TARGET+'?ref=abc', TARGET))
        self.assertTrue(checkpoint.same_amazon_target('https://amazon.com/gp/product/B012345678', TARGET))
        self.assertTrue(checkpoint.same_amazon_target('https://www.amazon.com/s?ref=abc&k=vacuum+bag',
                                                     'https://www.amazon.com/s?k=vacuum%20bag'))
        self.assertTrue(checkpoint.same_amazon_target('https://www.amazon.com/Best-Sellers/zgbs/arts-crafts/123/ref=abc',
                                                     'https://www.amazon.com/gp/bestsellers/123'))
        for actual in ('https://www.amazon.ca/dp/B012345678', 'https://www.amazon.com/',
                       'https://www.amazon.com/dp/B087654321', 'https://amazon.com.evil.example/dp/B012345678'):
            self.assertFalse(checkpoint.same_amazon_target(actual, TARGET))
        self.assertFalse(checkpoint.same_amazon_target('https://www.amazon.com/s?k=other',
                                                      'https://www.amazon.com/s?k=vacuum'))

    def test_collection_separator_reencoding_keeps_all_asins(self):
        expected = 'https://www.amazon.com/s?rh=p_78:B012345678|B087654321'
        self.assertTrue(checkpoint.same_amazon_target(
            'https://www.amazon.com/s?rh=p_78:B012345678%257CB087654321',expected))
        for actual in ('https://www.amazon.com/s?rh=p_78:B012345678',
                       'https://www.amazon.com/s?rh=p_78:B087654321%257CB012345678',
                       'https://www.amazon.com/s?rh=p_78:B012345678%257CB000000000'):
            self.assertFalse(checkpoint.same_amazon_target(actual,expected))

    def test_navigation_timeout_at_checkpoint_can_recover(self):
        page = Page()
        page.goto = Mock(side_effect=TimeoutError('导航超时'))
        with patch.object(checkpoint, 'time', SimpleNamespace(monotonic=lambda:page.now)):
            self.assertEqual(goto_amazon(page, TARGET, log=Mock())['status'], 'cleared')
        self.assertEqual(page.goto.call_count, 1)

    def test_ordinary_navigation_error_is_not_suppressed(self):
        page = Page(url=TARGET)
        page.goto = Mock(side_effect=TimeoutError('普通导航超时'))
        with self.assertRaises(TimeoutError):
            goto_amazon(page, TARGET)

    def test_log_omits_query_and_fragment(self):
        page = Page(url=GATE+'?token=secret#private')
        log = Mock()
        with patch.object(checkpoint, 'time', SimpleNamespace(monotonic=lambda:page.now)):
            checkpoint.handle_checkpoint(page, log=log)
        self.assertNotIn('secret', str(log.call_args_list))
        self.assertNotIn('private', str(log.call_args_list))

    def test_handler_duration_does_not_override_batch_log_clock(self):
        page = Page()
        log = Mock()
        with patch.object(checkpoint, 'time', SimpleNamespace(monotonic=lambda:page.now)):
            result = checkpoint.handle_checkpoint(page, log=log)
        self.assertTrue(all('seconds' not in call.kwargs for call in log.call_args_list))
        self.assertEqual(log.call_args.kwargs['handling_seconds'], result['seconds'])

    def test_basic_collector_retains_failed_page_and_stops_next_request(self):
        from competitor_analysis.collect_basic_facts import PluginCollector, resolve_product
        page = Mock(url=GATE)
        collector = PluginCollector('fixture')
        collector.page = page
        with patch('competitor_analysis.collect_basic_facts.goto_amazon', side_effect=AmazonAccessRequired('超时')) as navigate:
            with collector:
                result = resolve_product('B012345678', {}, collector)
                self.assertEqual(result['status'], 'needs_user_action')
                self.assertTrue(result['access_paused'])
                with self.assertRaises(AmazonAccessRequired):
                    collector('B087654321', ['price'])
        navigate.assert_called_once()
        page.close.assert_not_called()

    def test_sync_unfinished_export_keeps_original_download_destination(self):
        from review_research.download_sellersprite_reviews import NamedDownload
        download = NamedDownload(Mock(), Path('fixture-downloads'))
        download.session = Mock()
        download.started = True
        download.state = 'inProgress'
        download.__exit__(AmazonAccessRequired, AmazonAccessRequired('验证未恢复'), None)
        download.session.send.assert_called_once_with('Browser.setDownloadBehavior', {
            'behavior':'allow', 'downloadPath':str(download.destination), 'eventsEnabled':False})
        download.session.detach.assert_called_once()


class AsyncTests(unittest.IsolatedAsyncioTestCase):
    async def test_background_checkpoint_is_activated_before_check(self):
        page = AsyncPage()
        original = AsyncControl.check
        async def check(control, **kwargs):
            self.assertTrue(control.page.activated)
            return await original(control, **kwargs)
        with patch.object(AsyncControl,'check',check), patch.object(
                checkpoint,'time',SimpleNamespace(monotonic=lambda:page.now)):
            result = await checkpoint.handle_checkpoint_async(page,log=Mock())
        self.assertEqual(result['status'],'cleared')
        page.session.detach.assert_awaited_once()

    async def test_concurrent_handlers_submit_once(self):
        page = AsyncPage()
        with patch.object(checkpoint, 'time', SimpleNamespace(monotonic=lambda:page.now)):
            results = await asyncio.gather(*(checkpoint.handle_checkpoint_async(page, log=Mock()) for _ in range(2)))
        self.assertEqual(page.clicks, 1)
        self.assertEqual({r['status'] for r in results}, {'cleared','normal'})

    async def test_async_timeout_recovery_checks_target(self):
        page = AsyncPage()
        page.goto = AsyncMock(side_effect=TimeoutError('导航超时'))
        with patch.object(checkpoint, 'time', SimpleNamespace(monotonic=lambda:page.now)):
            self.assertEqual((await goto_amazon_async(page, TARGET, log=Mock()))['status'], 'cleared')
        self.assertEqual(page.clicks, 1)

    async def test_async_failure_is_cached_and_closed_page_reported(self):
        page = AsyncPage(destination=None)
        with patch.object(checkpoint, 'time', SimpleNamespace(monotonic=lambda:page.now)):
            self.assertEqual((await checkpoint.handle_checkpoint_async(page, timeout=3, log=Mock()))['status'], 'unresolved')
            self.assertTrue((await checkpoint.handle_checkpoint_async(page, log=Mock()))['cached'])
        self.assertEqual(page.clicks, 2)
        closed = AsyncPage()
        closed.closed = True
        self.assertEqual((await checkpoint.handle_checkpoint_async(closed, log=Mock()))['status'], 'closed')

    async def test_async_wrong_target_is_reentered_once(self):
        page = AsyncPage(destination='https://www.amazon.com/')
        page.goto = AsyncMock(side_effect=lambda url,**kwargs:page.navigate(url))
        with patch.object(checkpoint, 'time', SimpleNamespace(monotonic=lambda:page.now)):
            result = await ensure_amazon_access_async(page,expected_url=TARGET,log=Mock())
        self.assertTrue(result['target_reentered'])
        self.assertEqual(result['status'],'cleared')
        page.goto.assert_awaited_once()

    async def test_async_reentry_stops_after_another_wrong_redirect(self):
        page = AsyncPage(destination='https://www.amazon.com/')
        page.goto = AsyncMock(side_effect=lambda url,**kwargs:page.navigate('https://www.amazon.com/dp/B087654321'))
        with patch.object(checkpoint, 'time', SimpleNamespace(monotonic=lambda:page.now)):
            with self.assertRaises(AmazonAccessRequired):
                await ensure_amazon_access_async(page, expected_url=TARGET, log=Mock())
        page.goto.assert_awaited_once()

    async def test_async_reentry_handles_one_new_checkpoint_and_returns_cleared(self):
        page = AsyncPage(destination='https://www.amazon.com/')
        def navigate(url,**kwargs):
            page.destination = TARGET
            page.navigate(GATE)
            raise TimeoutError('导航已到新的确认页')
        page.goto = AsyncMock(side_effect=navigate)
        with patch.object(checkpoint,'time',SimpleNamespace(monotonic=lambda:page.now)):
            result = await ensure_amazon_access_async(page,expected_url=TARGET,log=Mock())
        self.assertEqual(result['status'],'cleared')
        self.assertEqual(result['target_reentry_status'],'cleared')
        self.assertEqual(page.url,TARGET)
        page.goto.assert_awaited_once()

    async def test_async_normal_wrong_page_never_reenters(self):
        page = AsyncPage(url='https://www.amazon.com/')
        page.goto = AsyncMock()
        with self.assertRaises(AmazonAccessRequired):
            await ensure_amazon_access_async(page,expected_url=TARGET,log=Mock())
        page.goto.assert_not_awaited()


@unittest.skipUnless(os.environ.get('AMAZON_CHECKPOINT_BROWSER_TESTS') == '1', '需要显式开启隔离Chrome测试')
class BrowserTests(unittest.TestCase):
    def test_sync_gate_reenters_product_keyword_collection_and_category(self):
        from playwright.sync_api import sync_playwright
        targets = [TARGET,'https://www.amazon.com/s?k=vacuum+bag',
                   'https://www.amazon.com/s?rh=p_78:B012345678|B087654321',
                   'https://www.amazon.com/gp/bestsellers/electronics/123']
        with sync_playwright() as runtime:
            browser = runtime.chromium.launch(channel='chrome',headless=True)
            try:
                for target in targets:
                    with self.subTest(target=target):
                        context = browser.new_context()
                        requests = []
                        def route(request):
                            requests.append(request.request.url)
                            html = ('<input type="checkbox" id="continueCheck"><button onclick="location.href=\'/\'">Continue shopping</button>'
                                    if '/checkpoint/' in request.request.url else '<body>已加载测试目标</body>')
                            request.fulfill(status=200,content_type='text/html; charset=utf-8',body=html)
                        context.route('**/*',route)
                        page = context.new_page()
                        page.goto(GATE)
                        result = ensure_amazon_access(page,expected_url=target,log=Mock())
                        self.assertTrue(result['target_reentered'])
                        self.assertTrue(checkpoint.same_amazon_target(page.url,target))
                        self.assertEqual(sum(checkpoint.same_amazon_target(u,target) for u in requests),1)
                        context.close()
            finally:
                browser.close()

    def test_explicit_missing_metric_is_distinct_from_loading_and_unfinished_charts(self):
        asyncio.run(self.overview_scenario())

    async def overview_scenario(self):
        from playwright.async_api import async_playwright
        from product_research.scripts.capture_sellersprite_categories import overview_status
        async with async_playwright() as runtime:
            browser = await runtime.chromium.launch(channel='chrome',headless=True)
            try:
                page = await browser.new_page()
                cases = [('--',False,False,False,True),('',False,False,False,False),
                         ('加载中',False,False,False,False),('123',False,False,False,False),
                         ('--',True,False,False,False),('--',False,True,False,False),
                         ('--',False,False,True,False)]
                for value, animated, progress, broken, expected in cases:
                    with self.subTest(value=value,animated=animated,progress=progress,broken=broken):
                        charts = ''.join(f'<div _echarts_instance_="{i}">{"" if broken and i==0 else "<canvas width=100 height=60></canvas>"}</div>' for i in range(11))
                        spinner = '<div role="progressbar" style="width:10px;height:10px"></div>' if progress else ''
                        style = 'animation:loading 1s infinite' if animated else ''
                        await page.set_content('<style>@keyframes loading {from {opacity:0.2} to {opacity:1}}</style>'
                            '<div id="target"><div class="overview">'+charts+
                            '<div class="product-feature" _echarts_instance_="feature"></div>'+
                            f'<div style="{style}"><div>近30天售出件数</div><div>{value}</div>{spinner}'
                            '<div class="sold-out-box" _echarts_instance_="sold"></div></div></div></div>')
                        state = await overview_status(page.locator('#target'))
                        self.assertEqual(state['ready'],expected,state)
                        if expected:
                            self.assertIn('近30天售出件数',state['unavailable_metrics'])
            finally:
                await browser.close()

    def test_sync_isolated_chrome_gate_and_target(self):
        from playwright.sync_api import sync_playwright
        with sync_playwright() as runtime:
            browser = runtime.chromium.launch(channel='chrome', headless=True)
            try:
                context = browser.new_context()
                def route(request):
                    if '/checkpoint/' in request.request.url:
                        html = '<input type="checkbox" id="continueCheck"><button onclick="if(document.querySelector(\'#continueCheck\').checked)location.href=\'/dp/B012345678\'">Continue shopping</button>'
                    else:
                        html = '<body><input id="ASIN" value="B012345678"><div id="glow-ingress-line2">New York 10001</div></body>'
                    request.fulfill(status=200, content_type='text/html; charset=utf-8', body=html)
                context.route('**/*', route)
                page = context.new_page()
                page.goto(GATE)
                self.assertEqual(ensure_amazon_access(page, expected_url=TARGET, log=Mock())['status'], 'cleared')
                self.assertEqual(page.locator('#ASIN').input_value(), 'B012345678')
            finally:
                browser.close()

    def test_async_review_export_recovers_without_second_download(self):
        for landing in (TARGET,'https://www.amazon.com/'):
            with self.subTest(landing=landing):
                asyncio.run(self.review_scenario(landing))

    async def review_scenario(self,landing):
        from playwright.async_api import async_playwright
        from review_research.prefetch_sellersprite_reviews import prefetch_one
        async with async_playwright() as runtime:
            browser = await runtime.chromium.launch(channel='chrome', headless=True)
            try:
                context = await browser.new_context()
                exports = []
                await context.expose_function('recordExportClick',lambda:exports.append('clicked'))
                async def route(request):
                    if '/checkpoint/' in request.request.url:
                        html = f'<input type="checkbox" id="continueCheck"><button onclick="location.href=\'{landing}\'">Continue shopping</button>'
                    else:
                        html = '''<div id="glow-ingress-line2">New York 10001</div>
                        <div id="seller-sprite-extension-quick-view-listing">ASIN: B012345678
                        <button onclick="recordExportClick()">评论下载</button></div><label><input type="checkbox" checked>翻译标题和内容</label>'''
                    await request.fulfill(status=200, content_type='text/html; charset=utf-8', body=html)
                await context.route('**/*', route)
                page = await context.new_page()
                await page.goto(TARGET)
                tracker = Mock(state='inProgress', error=None, completed_path=Mock(return_value=Path('source.xlsx')))
                router = Mock(register=AsyncMock(return_value=tracker))
                async def collect(*args):
                    # 导出已经点击后模拟跳转；恢复期间仅派发完成事件。
                    await page.goto(GATE)
                    return {'asin':'B012345678'}
                events = []
                def log(event, **fields):
                    events.append(event)
                    if event == 'review_access_resumed':
                        tracker.state = 'completed'
                with tempfile.TemporaryDirectory() as folder, patch(
                        'review_research.prefetch_sellersprite_reviews.save_export',
                        AsyncMock(return_value={'status':'downloaded','review_count':3})) as save:
                    result = await prefetch_one(context, router, 'B012345678', Path(folder), asyncio.Lock(),
                                                asyncio.Event(), asyncio.Event(), log, page=page, collect_facts=collect)
                self.assertEqual(result['status'], 'downloaded', result)
                self.assertEqual(events.count('review_started'), 1)
                self.assertEqual(exports,['clicked'])
                self.assertIn('review_access_resumed', events)
                self.assertEqual('checkpoint_target_reentry' in events,landing != TARGET)
                router.register.assert_awaited_once()
                save.assert_awaited_once()
                self.assertTrue(page.is_closed())
            finally:
                await browser.close()


if __name__ == '__main__':
    unittest.main()
