"""只读识别Amazon验证页，供各阶段暂停访问并保留人工处理页面。"""
from urllib.parse import urlparse

from shared.marketplaces import url_belongs_to_amazon


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
