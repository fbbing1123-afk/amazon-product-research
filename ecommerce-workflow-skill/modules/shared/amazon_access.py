"""识别Amazon访问状态，并在业务明确调用时处理已知确认页。"""
from urllib.parse import urlparse

from shared.marketplaces import url_belongs_to_amazon
from shared.amazon_checkpoint import (
    handle_checkpoint, handle_checkpoint_async, same_amazon_target, _emit as emit_checkpoint_event,
)


class AmazonAccessRequired(RuntimeError):
    """当前Amazon页面需要用户完成验证或登录。"""


CHALLENGE_SCRIPT = r'''() => {
  const visible = el => {
    const box = el.getBoundingClientRect();
    const style = getComputedStyle(el);
    return box.width > 0 && box.height > 0 && style.display !== 'none' && style.visibility !== 'hidden';
  };
  return [...document.querySelectorAll('#captchacharacters, #continueCheck, form[action*="validateCaptcha" i]')].some(visible);
}'''


def is_amazon_access_url(url):
    if not url_belongs_to_amazon(url):
        return False
    path = urlparse(url).path.lower()
    return any(part in path for part in ('/checkpoint/', 'validatecaptcha', '/ap/signin', '/ap/cvf'))


def pending_amazon_access_page(context):
    """先查看已有标签的网址，避免验证尚未处理时继续打开新页面。"""
    return next((page for page in context.pages if not page.is_closed() and is_amazon_access_url(page.url)), None)


def amazon_challenge_visible(page):
    return is_amazon_access_url(page.url) or bool(page.evaluate(CHALLENGE_SCRIPT))


async def amazon_challenge_visible_async(page):
    return is_amazon_access_url(page.url) or bool(await page.evaluate(CHALLENGE_SCRIPT))


def _require_result(page, result, expected_url):
    if result['status'] not in {'normal', 'cleared'}:
        raise AmazonAccessRequired(f"Amazon访问受阻：{result['status']}，页面及进度保留")
    if expected_url and not same_amazon_target(page.url, expected_url):
        raise AmazonAccessRequired('确认页后目标地址不符，保留页面，拒绝读取其它商品或首页')
    return result


def _should_reenter(page, result, expected_url, allowed):
    """只恢复本次确认后丢失的原任务，不替普通页面纠正任意错误导航。"""
    return (allowed and result['status'] == 'cleared' and expected_url
            and url_belongs_to_amazon(expected_url) and not is_amazon_access_url(expected_url)
            and not same_amazon_target(page.url, expected_url))


def _navigate(page, url, timeout):
    """导航超时若已落在确认页，交给后续状态检查；其它错误原样保留。"""
    try:
        page.goto(url, wait_until='domcontentloaded', timeout=timeout)
    except Exception:
        if not is_amazon_access_url(page.url):
            raise


async def _navigate_async(page, url, timeout):
    try:
        await page.goto(url, wait_until='domcontentloaded', timeout=timeout)
    except Exception:
        if not is_amazon_access_url(page.url):
            raise


def ensure_amazon_access(page, *, expected_url=None, log=None, allow_target_reentry=True):
    """处理已知确认后再检查其它访问限制；正常页面不产生点击。"""
    result = (handle_checkpoint(page, log=log) if is_amazon_access_url(page.url)
              else {'status': 'normal' if url_belongs_to_amazon(page.url) else 'other_site', 'attempts': 0})
    _require_result(page, result, None)
    if page.is_closed() or amazon_challenge_visible(page):
        raise AmazonAccessRequired('Amazon需要登录或其它验证，页面及进度保留')
    if _should_reenter(page, result, expected_url, allow_target_reentry):
        emit_checkpoint_event(log, 'checkpoint_target_reentry', page)
        try:
            _navigate(page, expected_url, 45000)
            # 本轮最多重进一次；若再次走错页或遇验证，保留现场，不递归重开。
            recovered = ensure_amazon_access(page, expected_url=expected_url, log=log, allow_target_reentry=False)
        except AmazonAccessRequired:
            raise
        except Exception as error:
            raise AmazonAccessRequired(f'确认页后重进原目标失败：{type(error).__name__}，页面及进度保留') from error
        result = {**result, 'target_reentered': True, 'target_reentry_status': recovered['status']}
    return _require_result(page, result, expected_url)


async def ensure_amazon_access_async(page, *, expected_url=None, log=None, allow_target_reentry=True):
    """异步业务在原有操作锁内调用，保持下载路由和浏览器连接不变。"""
    result = (await handle_checkpoint_async(page, log=log) if is_amazon_access_url(page.url)
              else {'status': 'normal' if url_belongs_to_amazon(page.url) else 'other_site', 'attempts': 0})
    _require_result(page, result, None)
    if page.is_closed() or await amazon_challenge_visible_async(page):
        raise AmazonAccessRequired('Amazon需要登录或其它验证，页面及进度保留')
    if _should_reenter(page, result, expected_url, allow_target_reentry):
        emit_checkpoint_event(log, 'checkpoint_target_reentry', page)
        try:
            await _navigate_async(page, expected_url, 45000)
            recovered = await ensure_amazon_access_async(
                page, expected_url=expected_url, log=log, allow_target_reentry=False)
        except AmazonAccessRequired:
            raise
        except Exception as error:
            raise AmazonAccessRequired(f'确认页后重进原目标失败：{type(error).__name__}，页面及进度保留') from error
        result = {**result, 'target_reentered': True, 'target_reentry_status': recovered['status']}
    return _require_result(page, result, expected_url)


def goto_amazon(page, url, *, timeout=45000, log=None):
    """导航超时若实际停在确认页，仍处理该页；普通导航错误原样上抛。"""
    _navigate(page, url, timeout)
    return ensure_amazon_access(page, expected_url=url, log=log)


async def goto_amazon_async(page, url, *, timeout=45000, log=None):
    """异步导航和确认恢复共用目标核验，不额外重载成功页面。"""
    await _navigate_async(page, url, timeout)
    return await ensure_amazon_access_async(page, expected_url=url, log=log)
