"""核验并发导出归属、退出恢复及真实重叠证据，禁止串商品或虚报提速。"""
import asyncio
from pathlib import Path
import tempfile
import json
from types import SimpleNamespace
import unittest
from contextlib import ExitStack
from unittest.mock import AsyncMock, Mock, patch

from review_research.prefetch_sellersprite_reviews import DownloadRouter, save_export, prefetch_one
from scripts.run_review_prefetch import overlap_evidence, execute
from review_research.download_sellersprite_reviews import NeedsUserAction


class RouterTests(unittest.IsolatedAsyncioTestCase):
    async def test_unfinished_export_failure_stops_new_product_opening(self):
        with tempfile.TemporaryDirectory() as folder:
            page = Mock(url='https://www.amazon.com/dp/B012345678', bring_to_front=AsyncMock(),
                        is_closed=Mock(return_value=False), close=AsyncMock())
            page.get_by_role.return_value = Mock(is_checked=AsyncMock(return_value=True))
            page.get_by_text.return_value = Mock(wait_for=AsyncMock())
            button = Mock(wait_for=AsyncMock(), count=AsyncMock(return_value=1), hover=AsyncMock(), click=AsyncMock())
            panel = Mock(get_by_text=Mock(return_value=button))
            router = Mock(register=AsyncMock(return_value=Mock(state='inProgress', error='下载连接中断')))
            context = Mock(pages=[])
            blocked = asyncio.Event()
            with patch('review_research.prefetch_sellersprite_reviews.check_product', AsyncMock(return_value=panel)):
                result = await prefetch_one(context, router, 'B012345678', Path(folder), asyncio.Lock(),
                                            asyncio.Event(), blocked, Mock(), page=page)
                next_result = await prefetch_one(context, router, 'B087654321', Path(folder), asyncio.Lock(),
                                                 asyncio.Event(), blocked, Mock())
            self.assertEqual(result['status'], 'failed')
            self.assertTrue(blocked.is_set())
            self.assertEqual(next_result['status'], 'not_attempted')
            context.new_page.assert_not_called()
            page.close.assert_not_awaited()

    async def test_page_closes_only_after_file_validation_succeeds(self):
        for fail in (False,True):
            with self.subTest(fail=fail),tempfile.TemporaryDirectory() as folder:
                order=[]
                page=Mock(url='https://www.amazon.com/dp/B012345678',bring_to_front=AsyncMock(),
                          is_closed=Mock(return_value=False),close=AsyncMock(side_effect=lambda:order.append('close')))
                page.get_by_role.return_value=Mock(is_checked=AsyncMock(return_value=True))
                page.get_by_text.return_value=Mock(wait_for=AsyncMock())
                button=Mock(wait_for=AsyncMock(),count=AsyncMock(return_value=1),hover=AsyncMock(),click=AsyncMock())
                panel=Mock(get_by_text=Mock(return_value=button))
                tracker=Mock(state='completed',completed_path=Mock(return_value=Path(folder)/'source.xlsx'))
                router=Mock(register=AsyncMock(return_value=tracker))
                async def save(*args):
                    order.append('validate')
                    if fail:
                        raise ValueError('归属错误')
                    return {'status':'downloaded','review_count':1}
                with patch('review_research.prefetch_sellersprite_reviews.check_product',AsyncMock(return_value=panel)),patch(
                        'review_research.prefetch_sellersprite_reviews.save_export',side_effect=save):
                    result=await prefetch_one(Mock(pages=[]),router,'B012345678',Path(folder),asyncio.Lock(),
                                              asyncio.Event(),asyncio.Event(),Mock(),page=page)
                self.assertEqual(order,['validate'] if fail else ['validate','close'])
                self.assertEqual(result['status'],'failed' if fail else 'downloaded')

    async def test_pending_export_keeps_destination_and_is_not_canceled(self):
        with tempfile.TemporaryDirectory() as folder:
            session=Mock(send=AsyncMock(),detach=AsyncMock())
            browser=Mock(new_browser_cdp_session=AsyncMock(return_value=session))
            async with DownloadRouter(browser,Path(folder)) as router:
                router.trackers['one']=Mock(asin='B012345678',started=True,state='inProgress')
            session.send.assert_awaited_with('Browser.setDownloadBehavior',{
                'behavior':'allow','downloadPath':str(Path(folder).resolve()),'eventsEnabled':False})
            self.assertFalse(any(c.args[0]=='Browser.cancelDownload' for c in session.send.call_args_list))

    async def test_all_startups_fail_without_blocking_local_work(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root/'input.xlsx'
            source.write_bytes(b'fixture')
            args = SimpleNamespace(workbook=source, output_dir=root/'out', cdp='fixture',
                                   max_downloads=2, download_timeout=1, postal_code='10001',
                                   skip_metrics=False, skip_categories=True)
            page = Mock(goto=AsyncMock())
            browser = Mock(contexts=[Mock(new_page=AsyncMock(return_value=page))])
            runtime = Mock(chromium=Mock(connect_over_cdp=AsyncMock(return_value=browser)))
            manager = AsyncMock()
            manager.__aenter__.return_value = runtime
            local = Mock(return_value={'sample_count':2,'output':str(root/'metrics.xlsx')})
            async def failed(*args, **kwargs):
                return {'asin':args[2], 'status':'failed'}
            with patch('scripts.run_review_prefetch.read_plan', return_value=(['B012345678'], [], True)), patch(
                    'scripts.run_review_prefetch.source_candidates', return_value={'B012345678':{}}), patch(
                    'scripts.run_review_prefetch.async_playwright', return_value=manager), patch(
                    'scripts.run_review_prefetch.DownloadRouter', return_value=AsyncMock()), patch(
                    'scripts.run_review_prefetch.prefetch_one', side_effect=failed), patch(
                    'scripts.run_review_prefetch.refresh_summary', local):
                self.assertEqual(await asyncio.wait_for(execute(args), 3), 1)
            local.assert_called_once()
            manifest = json.loads(next((root/'out').glob('*/manifest.json')).read_text(encoding='utf-8'))
            self.assertEqual(manifest['status'], 'incomplete')
            self.assertEqual(manifest['metrics']['sample_count'], 2)

    async def test_save_download_creates_product_directory_without_redownloading(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            payload = {'source_sha256':'fingerprint', 'review_count':1, 'translation':{}, 'warnings':[]}
            with patch('review_research.prefetch_sellersprite_reviews.read_export', return_value=payload), patch(
                    'review_research.prefetch_sellersprite_reviews.export_status', return_value='downloaded'):
                result = await save_export(root/'original.xlsx', 'B012345678', root, 1)
            self.assertTrue((root/'B012345678/reviews.json').is_file())
            self.assertEqual(result['source_file'], str(root/'original.xlsx'))

    async def test_interleaved_downloads_keep_product_and_original_name(self):
        with tempfile.TemporaryDirectory() as folder:
            session = Mock(send=AsyncMock(), detach=AsyncMock())
            browser = Mock(new_browser_cdp_session=AsyncMock(return_value=session))
            log = Mock()
            async with DownloadRouter(browser, Path(folder), log) as router:
                trackers = []
                for frame, asin in [('one', 'B012345678'), ('two', 'B087654321')]:
                    frames = Mock(send=AsyncMock(return_value={'frameTree': {'frame': {'id': frame}}}), detach=AsyncMock())
                    page = Mock(context=Mock(new_cdp_session=AsyncMock(return_value=frames)))
                    trackers.append(await router.register(page, asin))
                    router.on_begin({'frameId': frame, 'guid': frame, 'suggestedFilename': asin+'-reviews.xlsx'})
                router.on_begin({'frameId': 'foreign', 'guid': 'foreign', 'suggestedFilename': 'other.xlsx'})
                # 第二个商品先完成，不能误记到第一个商品。
                for frame, tracker in [('two', trackers[1]), ('one', trackers[0])]:
                    (Path(folder)/tracker.filename).write_bytes(b'fixture')
                    router.on_progress({'guid': frame, 'state': 'completed'})
                    self.assertEqual(tracker.completed_path().name, tracker.asin+'-reviews.xlsx')
                    if frame == 'two':
                        self.assertIsNone(trackers[0].state)
                self.assertEqual([c.kwargs['asin'] for c in log.call_args_list], ['B087654321', 'B012345678'])
            settings = [c.args[1]['behavior'] for c in session.send.call_args_list if c.args[0]=='Browser.setDownloadBehavior']
            self.assertEqual(settings, ['allow', 'default'])
            session.detach.assert_awaited_once()

    async def test_wrong_asin_filename_rejected_and_failure_restores_settings(self):
        with tempfile.TemporaryDirectory() as folder:
            session = Mock(send=AsyncMock(), detach=AsyncMock())
            browser = Mock(new_browser_cdp_session=AsyncMock(return_value=session))
            with self.assertRaisesRegex(RuntimeError, '中断'):
                async with DownloadRouter(browser, Path(folder)) as router:
                    frames = Mock(send=AsyncMock(return_value={'frameTree': {'frame': {'id': 'one'}}}), detach=AsyncMock())
                    page = Mock(context=Mock(new_cdp_session=AsyncMock(return_value=frames)))
                    tracker = await router.register(page, 'B012345678')
                    router.on_begin({'frameId':'one', 'guid':'guid', 'suggestedFilename':'B087654321-reviews.xlsx'})
                    with self.assertRaisesRegex(ValueError, 'ASIN不一致'):
                        tracker.completed_path()
                    raise RuntimeError('中断')
            session.send.assert_any_await('Browser.setDownloadBehavior', {'behavior':'default', 'eventsEnabled':False})


class EvidenceTests(unittest.TestCase):
    def test_requires_real_completion_and_capture_within_download(self):
        events = [
            {'event':'review_started', 'asin':'A', 'seconds':10},
            {'event':'review_started', 'asin':'failed', 'seconds':11},
            {'event':'metrics_finished', 'seconds':12},
            {'event':'category_started', 'seconds':13},
            {'event':'category_capture', 'seconds':20},
            {'event':'review_file_completed', 'asin':'A', 'seconds':30},
            {'event':'category_capture', 'seconds':40},
        ]
        result = overlap_evidence(events)
        self.assertEqual([r['seconds'] for r in result], [12,20])
        self.assertTrue(all(r['pending_asins']==['A'] for r in result))


class BoundedPrefetchTests(unittest.IsolatedAsyncioTestCase):
    async def run_scenario(self, challenge=False):
        """真实调度与预取逻辑，页面和文件I/O替换为可控边界。"""
        with tempfile.TemporaryDirectory() as folder, ExitStack() as stack:
            root = Path(folder)
            source = root/'source.xlsx'
            source.write_bytes(b'fixture')
            asins = ['B000000001', 'B000000002', 'B000000003', 'B000000004']
            args = SimpleNamespace(workbook=source, output_dir=root/'out', cdp='fixture', max_downloads=None,
                                   download_timeout=5, postal_code='10001', skip_metrics=False, skip_categories=False)
            category = {'url':'https://www.amazon.com/gp/bestsellers/1', 'row':10, 'label':'样本类目'}
            pages, opened, trackers = [], [], {}
            first_pair, all_opened, metrics_done, screenshot_done = (asyncio.Event() for _ in range(4))
            loop = asyncio.get_running_loop()
            max_live = 0
            context = Mock(pages=pages)
            async def new_page():
                page = Mock(url='about:blank', context=context, bring_to_front=AsyncMock(),
                            is_closed=Mock(return_value=False))
                async def close(*args, **kwargs):
                    page.is_closed.return_value = True
                page.close = AsyncMock(side_effect=close)
                async def goto(url, **kwargs):
                    nonlocal max_live
                    asin = url.rsplit('/',1)[-1]
                    opened.append(asin)
                    page.url = 'https://www.amazon.com/checkpoint/continue' if challenge and asin==asins[1] else url
                    max_live = max(max_live, sum('/dp/' in p.url and not p.is_closed() for p in pages))
                    if len(opened)==2:
                        first_pair.set()
                    if len(opened)==4:
                        all_opened.set()
                page.goto = AsyncMock(side_effect=goto)
                page.get_by_role.return_value = Mock(is_checked=AsyncMock(return_value=True))
                page.get_by_text.return_value = Mock(wait_for=AsyncMock())
                pages.append(page)
                return page
            context.new_page = AsyncMock(side_effect=new_page)
            browser = Mock(contexts=[context])
            manager = AsyncMock()
            manager.__aenter__.return_value = Mock(chromium=Mock(connect_over_cdp=AsyncMock(return_value=browser)))
            router = AsyncMock()
            router.__aenter__.return_value = router
            async def register(page, asin):
                tracker = Mock(state='inProgress', error=None, started=False,
                               completed_path=Mock(return_value=root/(asin+'.xlsx')))
                trackers[asin] = tracker
                return tracker
            router.register.side_effect = register
            async def check(page, *args):
                if '/checkpoint/' in page.url:
                    raise NeedsUserAction('人工验证')
                button = Mock(wait_for=AsyncMock(), count=AsyncMock(return_value=1), hover=AsyncMock(), click=AsyncMock())
                return Mock(get_by_text=Mock(return_value=button), inner_text=AsyncMock(return_value='下载中...50%'))
            async def facts(page, asin, *args):
                return {'asin':asin,'status':'complete','fields':{k:{'value':1} for k in ('monthly_sales','price','listed_at')}}
            def metrics(*args):
                loop.call_soon_threadsafe(metrics_done.set)
                return {'sample_count':4,'output':str(root/'metrics.xlsx')}
            async def capture(*args, **kwargs):
                self.assertTrue(any(t.state=='inProgress' for t in trackers.values()))
                kwargs['on_capture']()
                screenshot_done.set()
                return {'image':str(root/'category.png')}
            patches = {
                'scripts.run_review_prefetch.read_plan': Mock(return_value=(asins,[category],False)),
                'scripts.run_review_prefetch.source_candidates': Mock(return_value={a:{} for a in asins}),
                'scripts.run_review_prefetch.async_playwright': Mock(return_value=manager),
                'scripts.run_review_prefetch.DownloadRouter': Mock(return_value=router),
                'scripts.run_review_prefetch.collect_main_image': AsyncMock(return_value={'status':'ok'}),
                'scripts.run_review_prefetch.collect_same_page': AsyncMock(side_effect=facts),
                'scripts.run_review_prefetch.refresh_summary': Mock(side_effect=metrics),
                'scripts.run_review_prefetch.refresh_basic_cells': Mock(return_value={'output':str(root/'basics.xlsx')}),
                'scripts.run_review_prefetch.refresh_product_images': Mock(return_value={'output':str(root/'images.xlsx')}),
                'scripts.run_review_prefetch.capture': AsyncMock(side_effect=capture),
                'scripts.run_review_prefetch.refresh_images': Mock(return_value={'output':str(root/'final.xlsx')}),
                'scripts.run_review_prefetch.validate_zip_package': Mock(return_value=Mock(as_dict=Mock(return_value={}))),
                'scripts.run_review_prefetch.validate_category_research': Mock(return_value=Mock(as_dict=Mock(return_value={}))),
                'review_research.prefetch_sellersprite_reviews.check_product': AsyncMock(side_effect=check),
                'review_research.prefetch_sellersprite_reviews.save_export': AsyncMock(return_value={'status':'downloaded','review_count':3}),
            }
            for name, value in patches.items():
                stack.enter_context(patch(name, value))
            task = asyncio.create_task(execute(args))
            try:
                await asyncio.wait_for(first_pair.wait(), 2)
                await asyncio.wait_for(metrics_done.wait(), 2)
                self.assertEqual(opened, asins[:2])
                if challenge:
                    trackers[asins[0]].state = 'completed'
                    self.assertEqual(await asyncio.wait_for(task, 3), 1)
                    self.assertEqual(opened, asins[:2])
                    patches['scripts.run_review_prefetch.capture'].assert_not_awaited()
                    pages[1].close.assert_not_awaited()
                else:
                    await asyncio.wait_for(screenshot_done.wait(), 2)
                    self.assertEqual(opened, asins[:2])
                    for tracker in trackers.values():
                        tracker.state = 'completed'
                    await asyncio.wait_for(all_opened.wait(), 2)
                    # 等最后一个注册完成，避免只看到导航完成而尚未点击导出。
                    while len(trackers)<4:
                        await asyncio.sleep(0)
                    for tracker in trackers.values():
                        tracker.state = 'completed'
                    self.assertEqual(await asyncio.wait_for(task, 3), 0)
                    self.assertLessEqual(max_live, 2)
                    self.assertEqual(len(trackers), 4)
                manifest = json.loads(next((root/'out').glob('*/manifest.json')).read_text(encoding='utf-8'))
                self.assertEqual(manifest['max_downloads'], 2)
                self.assertEqual(manifest['access_paused'], challenge)
                self.assertEqual(manifest['remaining_asins'], asins[1:] if challenge else [])
            finally:
                if not task.done():
                    task.cancel()
                await asyncio.gather(task, return_exceptions=True)

    async def test_opening_is_bounded_and_local_work_overlaps_first_pair(self):
        await self.run_scenario()

    async def test_challenge_stops_later_products_and_screenshots_but_keeps_local_work(self):
        await self.run_scenario(challenge=True)


if __name__ == '__main__':
    unittest.main()
