"""处理已知Amazon确认页，供同步及异步业务直接调用，不另开浏览器。"""
from __future__ import annotations

import asyncio
import json
import math
import re
import time
from datetime import datetime, timezone
from urllib.parse import urlsplit, urlunsplit, parse_qs
from weakref import WeakKeyDictionary

from shared.marketplaces import url_belongs_to_amazon


_STATES = WeakKeyDictionary()
_LOCKS = WeakKeyDictionary()
_PROBE = r'''() => {
  const visible = e => {
    const r = e.getBoundingClientRect(), s = getComputedStyle(e);
    return r.width > 0 && r.height > 0 && s.display !== 'none' && s.visibility !== 'hidden';
  };
  return {body: !!document.body, captcha: [...document.querySelectorAll(
    '#captchacharacters, form[action*="validateCaptcha" i], iframe[src*="recaptcha" i], iframe[src*="hcaptcha" i]'
  )].some(visible)};
}'''


def is_checkpoint_url(url):
    return url_belongs_to_amazon(url) and urlsplit(url).path.lower().rstrip('/') == '/checkpoint/continue'


def same_amazon_target(actual, expected):
    """接受同站商品规范链接及跟踪参数，拒绝首页、其它商品和其它站点。"""
    if not url_belongs_to_amazon(actual) or not url_belongs_to_amazon(expected):
        return False
    left, right = urlsplit(actual), urlsplit(expected)
    if (left.hostname or '').removeprefix('www.') != (right.hostname or '').removeprefix('www.'):
        return False
    product = r'/(?:dp|gp/product)/([A-Z0-9]{10})(?:[/?]|$)'
    wanted = re.search(product, right.path, re.I)
    if wanted:
        observed = re.search(product, left.path, re.I)
        return bool(observed and observed[1].upper() == wanted[1].upper())
    category = r'/(?:gp/bestsellers|zgbs)/(?:[^?#]*?/)?(\d+)(?:/ref=.*)?/?$'
    wanted_category = re.search(category, right.path, re.I)
    if wanted_category:
        observed_category = re.search(category, left.path, re.I)
        return bool(observed_category and observed_category[1] == wanted_category[1])
    clean = lambda path: path.split('/ref=')[0].rstrip('/')
    if clean(left.path) != clean(right.path):
        return False
    query, required = parse_qs(left.query), parse_qs(right.query)
    def value(items, key):
        values = items.get(key)
        if key == 'rh' and values and all(v.startswith('p_78:') for v in values):
            # Amazon集合页可能把竖线再编码一层；只归一该分隔符，保留ASIN顺序和数量。
            return [re.sub('%7c', '|', v, flags=re.I) for v in values]
        return values
    return all(value(query,key) == value(required,key) for key in ('k', 'rh') if key in required)


def _route_kind(page):
    if page.is_closed():
        return 'closed'
    if not url_belongs_to_amazon(page.url):
        return 'other_site'
    path = urlsplit(page.url).path.lower()
    if '/ap/signin' in path or '/ap/cvf' in path:
        return 'login_required'
    if 'validatecaptcha' in path:
        return 'other_verification'
    return None


def _kind(page):
    route = _route_kind(page)
    if route:
        return route
    probe = page.evaluate(_PROBE)
    return ('other_verification' if probe['captcha'] else 'checkpoint' if is_checkpoint_url(page.url)
            else 'normal' if probe['body'] else 'loading')


async def _kind_async(page):
    route = _route_kind(page)
    if route:
        return route
    probe = await page.evaluate(_PROBE)
    return ('other_verification' if probe['captcha'] else 'checkpoint' if is_checkpoint_url(page.url)
            else 'normal' if probe['body'] else 'loading')


def _state(page):
    if page not in _STATES:
        state = {'navigation': 0, 'failed_navigation': None, 'result': None}
        _STATES[page] = state
        def navigated(frame):
            if frame == page.main_frame:
                state['navigation'] += 1
        page.on('framenavigated', navigated)
    return _STATES[page]


def _emit(log, event, page, **fields):
    parsed = urlsplit(page.url)
    fields.update(url=urlunsplit((parsed.scheme, parsed.netloc, parsed.path, '', '')))
    if log:
        # 业务日志的seconds是批次时间，处理耗时不能覆盖它造成时间倒退。
        if 'seconds' in fields:
            fields['handling_seconds'] = fields.pop('seconds')
        log(event, **fields)
    else:
        print(json.dumps({'time': datetime.now(timezone.utc).isoformat(), 'event': event, **fields},
                         ensure_ascii=False), flush=True)


def _finish(page, state, status, attempts, began, log, error=''):
    result = {'status': status, 'attempts': attempts, 'seconds': round(time.monotonic()-began, 3)}
    if error:
        result['error'] = error
    if status != 'cleared':
        state.update(failed_navigation=state['navigation'], result=result)
    _emit(log, 'checkpoint_finished', page, **result)
    return result


def _validate(timeout, max_attempts):
    if not math.isfinite(timeout) or timeout <= 0 or type(max_attempts) is not int or max_attempts < 1:
        raise ValueError('确认页时限须为有限正数，提交次数须为正整数')


def _activate(page):
    """后台或最小化窗口可能暂停动画帧，点击前恢复目标页的正常渲染。"""
    session = page.context.new_cdp_session(page)
    try:
        window = session.send('Browser.getWindowForTarget')
        if window['bounds']['windowState'] == 'minimized':
            session.send('Browser.setWindowBounds', {
                'windowId': window['windowId'], 'bounds': {'windowState': 'normal'}})
        page.bring_to_front()
    finally:
        session.detach()


async def _activate_async(page):
    """只唤醒当前确认页，不新建浏览器、不关闭其它标签。"""
    session = await page.context.new_cdp_session(page)
    try:
        window = await session.send('Browser.getWindowForTarget')
        if window['bounds']['windowState'] == 'minimized':
            await session.send('Browser.setWindowBounds', {
                'windowId': window['windowId'], 'bounds': {'windowState': 'normal'}})
        await page.bring_to_front()
    finally:
        await session.detach()


def handle_checkpoint(page, *, timeout=15, max_attempts=2, log=None):
    """处理一次确认页；同一失败文档不重复提交，刷新或导航后允许再试。"""
    _validate(timeout, max_attempts)
    if not is_checkpoint_url(page.url):
        return {'status': _kind(page), 'attempts': 0}
    state = _state(page)
    if state['failed_navigation'] == state['navigation']:
        return {**state['result'], 'cached': True}
    began = time.monotonic()
    deadline, attempts, normal_checks, last_click, error = began+timeout, 0, 0, -math.inf, ''
    activated, phase = False, 'detect'
    _emit(log, 'checkpoint_detected', page)
    while time.monotonic() < deadline:
        try:
            kind = _kind(page)
            normal_checks = normal_checks+1 if kind == 'normal' else 0
            if normal_checks >= 2:
                return _finish(page, state, 'cleared', attempts, began, log)
            if kind not in {'normal', 'loading', 'checkpoint'}:
                return _finish(page, state, kind, attempts, began, log)
            if kind == 'checkpoint' and attempts < max_attempts:
                checkbox = page.locator('#continueCheck').first
                button = page.get_by_role('button', name='Continue shopping', exact=True).first
                if checkbox.is_visible() and button.is_visible():
                    if not activated:
                        phase = 'activate'
                        _activate(page)
                        activated = True
                    remaining = lambda cap: max(1, min(cap, (deadline-time.monotonic())*1000))
                    if not checkbox.is_checked():
                        phase = 'check'
                        checkbox.check(timeout=remaining(1000))
                        _emit(log, 'checkpoint_checked', page)
                    if button.is_enabled() and time.monotonic()-last_click >= 2 and time.monotonic() < deadline:
                        phase = 'submit'
                        attempts += 1
                        last_click = time.monotonic()
                        button.click(timeout=remaining(1500), no_wait_after=True)
                        _emit(log, 'checkpoint_submitted', page, attempts=attempts)
                        phase = 'wait_target'
        except Exception as exc:
            error = f'{phase}:{type(exc).__name__}'
        if page.is_closed():
            return _finish(page, state, 'closed', attempts, began, log)
        page.wait_for_timeout(max(0, min(250, (deadline-time.monotonic())*1000)))
    return _finish(page, state, 'unresolved', attempts, began, log, error)


async def handle_checkpoint_async(page, *, timeout=15, max_attempts=2, log=None):
    """异步处理与同页互斥；等待期间继续分发已有下载事件。"""
    _validate(timeout, max_attempts)
    lock = _LOCKS.setdefault(page, asyncio.Lock())
    async with lock:
        if not is_checkpoint_url(page.url):
            return {'status': await _kind_async(page), 'attempts': 0}
        state = _state(page)
        if state['failed_navigation'] == state['navigation']:
            return {**state['result'], 'cached': True}
        began = time.monotonic()
        deadline, attempts, normal_checks, last_click, error = began+timeout, 0, 0, -math.inf, ''
        activated, phase = False, 'detect'
        _emit(log, 'checkpoint_detected', page)
        while time.monotonic() < deadline:
            try:
                kind = await _kind_async(page)
                normal_checks = normal_checks+1 if kind == 'normal' else 0
                if normal_checks >= 2:
                    return _finish(page, state, 'cleared', attempts, began, log)
                if kind not in {'normal', 'loading', 'checkpoint'}:
                    return _finish(page, state, kind, attempts, began, log)
                if kind == 'checkpoint' and attempts < max_attempts:
                    checkbox = page.locator('#continueCheck').first
                    button = page.get_by_role('button', name='Continue shopping', exact=True).first
                    if await checkbox.is_visible() and await button.is_visible():
                        if not activated:
                            phase = 'activate'
                            await _activate_async(page)
                            activated = True
                        remaining = lambda cap: max(1, min(cap, (deadline-time.monotonic())*1000))
                        if not await checkbox.is_checked():
                            phase = 'check'
                            await checkbox.check(timeout=remaining(1000))
                            _emit(log, 'checkpoint_checked', page)
                        if await button.is_enabled() and time.monotonic()-last_click >= 2 and time.monotonic() < deadline:
                            phase = 'submit'
                            attempts += 1
                            last_click = time.monotonic()
                            await button.click(timeout=remaining(1500), no_wait_after=True)
                            _emit(log, 'checkpoint_submitted', page, attempts=attempts)
                            phase = 'wait_target'
            except Exception as exc:
                error = f'{phase}:{type(exc).__name__}'
            if page.is_closed():
                return _finish(page, state, 'closed', attempts, began, log)
            await page.wait_for_timeout(max(0, min(250, (deadline-time.monotonic())*1000)))
        return _finish(page, state, 'unresolved', attempts, began, log, error)
