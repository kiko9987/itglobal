# -*- coding: utf-8 -*-
"""드라이브 요청 한도(403 userRateLimitExceeded) 재시도 + 방문 사진 자동 재저장 (2026-10-27 L-04377)."""
import sys

sys.path.insert(0, '.')

from googleapiclient.errors import HttpError

import dashboard.utils.google_drive as gd


class _Resp(dict):
    def __init__(self, status):
        super().__init__()
        self.status = status
        self.reason = ''


def _err(status, reason):
    return HttpError(_Resp(status), f'{{"error": {{"errors": [{{"reason": "{reason}"}}]}}}}'.encode(), uri='u')


def test_retriable_classification():
    assert gd._drive_retriable(_err(403, 'userRateLimitExceeded'))
    assert gd._drive_retriable(_err(403, 'rateLimitExceeded'))
    assert gd._drive_retriable(_err(429, 'x')) and gd._drive_retriable(_err(503, 'x'))
    assert not gd._drive_retriable(_err(403, 'insufficientFilePermissions'))   # 권한 오류는 재시도 안 함
    assert not gd._drive_retriable(_err(404, 'notFound'))


class _Req:
    def __init__(self, fn):
        self.fn = fn

    def execute(self):
        return self.fn()


class _Files:
    def __init__(self, fails):
        self.fails, self.calls = fails, 0

    def list(self, **k):
        def run():
            self.calls += 1
            if self.calls <= self.fails:
                raise _err(403, 'userRateLimitExceeded')
            return {'files': []}
        return _Req(run)

    def create(self, **k):
        return _Req(lambda: {'id': 'NEW', 'webViewLink': 'L'})


def test_find_or_create_folder_retries_rate_limit(monkeypatch):
    f = _Files(fails=2)
    monkeypatch.setattr(gd, '_get_drive_service', lambda: type('S', (), {'files': lambda self: f})())
    import time
    monkeypatch.setattr(time, 'sleep', lambda s: None)
    assert gd.find_or_create_folder('현장사진', 'P')['id'] == 'NEW' and f.calls == 3


def test_find_or_create_folder_gives_up_on_permission(monkeypatch):
    class _F(_Files):
        def list(self, **k):
            def run():
                self.calls += 1
                raise _err(403, 'insufficientFilePermissions')
            return _Req(run)
    f = _F(0)
    monkeypatch.setattr(gd, '_get_drive_service', lambda: type('S', (), {'files': lambda self: f})())
    assert gd.find_or_create_folder('현장사진', 'P') is None and f.calls == 1


def test_photo_retry_schedules_then_asks_reupload(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    import dashboard.utils.redis_client as rcm
    cnt = {}

    class _R:
        def incr(self, k):
            cnt[k] = cnt.get(k, 0) + 1
            return cnt[k]

        def expire(self, k, t):
            pass
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': _R()})())
    timers = []

    class _T:
        def __init__(self, delay, fn, args=()):
            timers.append(delay)
            self.daemon = False

        def start(self):
            pass
    monkeypatch.setattr(sb.threading, 'Timer', _T)
    posted = []
    c = type('C', (), {'chat_postMessage': lambda self, **k: posted.append(k['text'])})()
    ev = {'channel': 'C', 'thread_ts': '1.0', 'ts': '2.0'}
    assert sb._retry_visit_photos_later(c, ev, 'L-04377') == 1
    assert timers == [90] and '자동으로 다시 저장' in posted[0]
    sb._retry_visit_photos_later(c, ev, 'L-04377')                 # 두 번째 — 안내 반복 없음
    assert timers == [90, 300] and len(posted) == 1
    sb._retry_visit_photos_later(c, ev, 'L-04377')                 # 세 번째 — 다시 올려 달라고
    assert timers == [90, 300] and '다시 올려 주세요' in posted[1]
