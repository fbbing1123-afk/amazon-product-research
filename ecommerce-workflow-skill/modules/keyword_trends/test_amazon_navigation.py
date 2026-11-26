"""验证正常搜索少访问、验证页保留以及定点补取不会丢失已成功结果。"""
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import Mock, patch

from keyword_trends import sellersprite_playwright_runner as runner


class AmazonNavigationTests(unittest.TestCase):
    def extract(self, postal_status='already-set', redirect=None, challenge=False, timeout=False):
        amazon = Mock(url='about:blank')
        def navigate(url, **kwargs):
            amazon.url = 'https://www.amazon.com/checkpoint/continue' if challenge else url
            if timeout:
                raise TimeoutError('页面加载超时')
        amazon.goto.side_effect = navigate
        amazon.evaluate.return_value = False
        seller = Mock(context=Mock(pages=[], new_page=Mock(return_value=amazon)))
        def postal(*args):
            if redirect:
                amazon.url = redirect
            return {'status':postal_status, 'text':'New York 10001'}
        with patch.object(runner, 'check_amazon_login_state', return_value={'status':'logged-in'}), \
             patch.object(runner, 'ensure_amazon_postal_code', side_effect=postal), \
             patch.object(runner, 'extract_count_from_amazon_text', return_value={'status':'success','count':123,'text':'123 results'}), \
             patch.object(runner, 'sleep'):
            if challenge:
                with self.assertRaises(runner.ManualActionRequired):
                    runner.extract_amazon_result(seller, 'vacuum bag')
                amazon.close.assert_not_called()
                seller.bring_to_front.assert_not_called()
            else:
                result = runner.extract_amazon_result(seller, 'vacuum bag')
                self.assertEqual(result['amazonResultCount'], 123)
                amazon.close.assert_called_once()
        return amazon

    def test_correct_postal_code_uses_one_search_request(self):
        self.assertEqual(self.extract().goto.call_count, 1)

    def test_tracking_parameters_and_equivalent_keyword_encoding_do_not_reload(self):
        page = self.extract(redirect='https://www.amazon.com/s?ref=abc&k=vacuum+bag')
        self.assertEqual(page.goto.call_count, 1)

    def test_postal_change_or_leaving_search_requires_one_return(self):
        self.assertEqual(self.extract(postal_status='success').goto.call_count, 2)
        self.assertEqual(self.extract(redirect='https://www.amazon.com/').goto.call_count, 2)

    def test_challenge_does_not_probe_account_or_close_page(self):
        self.assertEqual(self.extract(challenge=True).goto.call_count, 1)

    def test_navigation_timeout_at_challenge_still_pauses_and_retains_page(self):
        self.assertEqual(self.extract(challenge=True, timeout=True).goto.call_count, 1)

    def test_existing_challenge_prevents_new_tab(self):
        pending = Mock(url='https://www.amazon.com/checkpoint/continue', is_closed=Mock(return_value=False))
        seller = Mock(context=Mock(pages=[pending]))
        with self.assertRaises(runner.ManualActionRequired):
            runner.extract_amazon_result(seller, 'next keyword')
        seller.context.new_page.assert_not_called()
        pending.bring_to_front.assert_called_once()

    def test_visible_challenge_at_product_url_stops_login_probe(self):
        page = Mock(url='https://www.amazon.com/s?k=example', evaluate=Mock(return_value=True))
        with self.assertRaises(runner.ManualActionRequired):
            runner.require_amazon_login(page)
        page.goto.assert_not_called()

    def test_retry_keeps_successful_checkpoint_when_next_keyword_needs_user(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder)/'keywords.json'
            path.write_text(json.dumps({'results':[{'keyword':str(i),'amazonResultStatus':'failed'} for i in range(3)]}), encoding='utf-8')
            manager = Mock(__enter__=Mock(return_value=Mock()), __exit__=Mock(return_value=False))
            with patch.object(runner, 'require_sync_playwright', return_value=Mock(return_value=manager)), \
                 patch.object(runner, 'connect_cdp', return_value=(Mock(), Mock(), Mock())), \
                 patch.object(runner, 'extract_amazon_result', side_effect=[
                     {'amazonResultStatus':'success','amazonResultCount':10}, runner.ManualActionRequired('验证')
                 ]) as extract:
                with self.assertRaises(runner.ManualActionRequired):
                    runner.retry_failed_amazon_results(path)
            self.assertEqual(extract.call_count, 2)
            results = json.loads(path.read_text(encoding='utf-8'))['results']
            self.assertEqual(results[0]['amazonResultCount'], 10)
            self.assertEqual(results[2]['amazonResultStatus'], 'failed')


if __name__ == '__main__':
    unittest.main()
