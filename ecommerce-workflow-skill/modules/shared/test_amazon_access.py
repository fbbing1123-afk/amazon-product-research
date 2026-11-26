"""验证人工处理状态跨阶段传递，类目受阻后不继续请求下一个页面。"""
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import AsyncMock, Mock, patch

from openpyxl import Workbook
from shared.amazon_access import AmazonAccessRequired, is_amazon_access_url, pending_amazon_access_page
from product_research.scripts import capture_sellersprite_categories as categories


class AccessTests(unittest.TestCase):
    def test_only_amazon_challenge_and_login_routes_are_blocked(self):
        for url in ('https://www.amazon.com/checkpoint/continue', 'https://www.amazon.de/errors/validateCaptcha',
                    'https://www.amazon.co.uk/ap/signin', 'https://www.amazon.ca/ap/cvf'):
            self.assertTrue(is_amazon_access_url(url))
        for url in ('https://example.com/checkpoint/continue', 'https://www.amazon.com/dp/B012345678',
                    'https://amazon.com.example.com/ap/signin'):
            self.assertFalse(is_amazon_access_url(url))

    def test_closed_challenge_page_is_not_pending(self):
        page = Mock(url='https://www.amazon.com/checkpoint/continue', is_closed=Mock(return_value=True))
        self.assertIsNone(pending_amazon_access_page(Mock(pages=[page])))


class CategoryAccessTests(unittest.IsolatedAsyncioTestCase):
    async def test_existing_challenge_prevents_category_navigation(self):
        pending = Mock(url='https://www.amazon.com/checkpoint/continue', is_closed=Mock(return_value=False),
                       bring_to_front=AsyncMock())
        page = Mock(context=Mock(pages=[pending]))
        with self.assertRaises(AmazonAccessRequired):
            await categories.capture(page, 'https://www.amazon.com/gp/bestsellers/1', Path('unused.png'), '10001')
        page.goto.assert_not_called()
        page.context.new_cdp_session.assert_not_called()
        pending.bring_to_front.assert_awaited_once()

    async def test_category_challenge_saves_checkpoint_and_stops_batch(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root/'source.xlsx'
            book = Workbook()
            book.active.title = '样本-市场调研'
            book.save(source)
            book.close()
            args = SimpleNamespace(input=source, image_dir=root/'images', capture_only=True,
                                   start=0, limit=5, cdp='fixture', postal_code='10001')
            page = Mock(url='https://www.amazon.com/checkpoint/continue')
            context = Mock(pages=[], new_page=AsyncMock(return_value=page))
            manager = AsyncMock()
            manager.__aenter__.return_value = Mock(chromium=Mock(connect_over_cdp=AsyncMock(
                return_value=Mock(contexts=[context]))))
            items = [{'url':f'https://www.amazon.com/gp/bestsellers/{i}', 'row':10+i, 'label':str(i)} for i in (1,2)]
            with patch.object(categories, 'iter_category_rows', return_value=items), \
                 patch.object(categories, 'async_playwright', return_value=manager), \
                 patch.object(categories, 'capture', side_effect=AmazonAccessRequired('人工验证')) as capture:
                self.assertEqual(await categories.run(args), 1)
            capture.assert_awaited_once()
            context.new_page.assert_awaited_once()
            page.close.assert_not_called()
            result = json.loads((args.image_dir/'manifest.json').read_text(encoding='utf-8'))
            self.assertEqual(list(result), [items[0]['url']])
            self.assertEqual(result[items[0]['url']]['status'], 'needs_user_action')


if __name__ == '__main__':
    unittest.main()
