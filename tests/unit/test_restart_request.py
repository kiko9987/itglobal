# -*- coding: utf-8 -*-
"""서버 재시작 승인 요청(restart_request) 테스트 (2026-10-28).

- 요청: 사장님 DM 에 [지금 재시작][취소] 카드, 대기 중이면 같은 카드에 이유 덧붙임.
- 승인·취소: 사장님만, 30분 지난 요청은 만료, 이미 처리된 요청은 다시 못 누름.
- 실행 중·승인된 요청이 있으면 새 요청을 받지 않음(워치독이 처리 중).
"""
import sys

import pytest

sys.path.insert(0, '.')

from dashboard.services import restart_request as rr

ADMIN = 'U_ADMIN'


class FakeClient:
    def __init__(self):
        self.posts, self.updates = [], []

    def chat_postMessage(self, **kw):
        self.posts.append(kw)
        return {'channel': 'D_ADMIN', 'ts': '111.222'}

    def chat_update(self, **kw):
        self.updates.append(kw)


@pytest.fixture
def env(monkeypatch, tmp_path):
    monkeypatch.setenv('SLACK_ADMIN_CHANNEL', ADMIN)
    monkeypatch.setattr(rr, 'last_server_start', lambda: '2026-10-28 08:00:00')
    monkeypatch.setattr(rr, 'commits_since', lambda since: ['abc1234 fix: 첫 수정', 'def5678 feat: 둘째'])
    monkeypatch.setattr(rr, 'frontend_build_needed', lambda: False)
    return tmp_path / 'restart_request.json'


def _buttons(blocks):
    return [e['action_id'] for b in blocks if b['type'] == 'actions' for e in b['elements']]


def test_request_posts_card_with_two_buttons(env):
    c = FakeClient()
    req = rr.request_restart('문서 변환 봇 반영', c, path=env, now=1000)
    assert req['status'] == 'pending' and req['channel'] == 'D_ADMIN' and req['ts'] == '111.222'
    assert c.posts[0]['channel'] == ADMIN
    blocks = c.posts[0]['blocks']
    assert _buttons(blocks) == [rr.APPROVE_ACTION, rr.CANCEL_ACTION]
    text = blocks[0]['text']['text']
    assert '문서 변환 봇 반영' in text and '반영 대기 커밋 2개' in text and 'abc1234' in text
    assert rr.load_request(env)['id'] == req['id']


def test_second_request_merges_into_same_card(env):
    c = FakeClient()
    first = rr.request_restart('첫 이유', c, path=env, now=1000)
    second = rr.request_restart('둘째 이유', c, path=env, now=1100)
    assert second['id'] == first['id'] and len(c.posts) == 1      # 새 카드 없음
    assert second['reasons'] == ['첫 이유', '둘째 이유']
    assert c.updates and c.updates[-1]['ts'] == '111.222'


def test_expired_pending_request_is_replaced(env):
    c = FakeClient()
    rr.request_restart('옛 이유', c, path=env, now=1000)
    new = rr.request_restart('새 이유', c, path=env, now=1000 + rr.PENDING_TTL_SEC + 1)
    assert len(c.posts) == 2 and new['reasons'] == ['새 이유']


def test_build_warning_shown(env, monkeypatch):
    monkeypatch.setattr(rr, 'frontend_build_needed', lambda: True)
    c = FakeClient()
    rr.request_restart('화면 수정', c, path=env, now=1000)
    assert '빌드 먼저' in c.posts[0]['blocks'][0]['text']['text']


def test_only_admin_can_approve(env):
    c = FakeClient()
    req = rr.request_restart('이유', c, path=env, now=1000)
    msg = rr.decide(req['id'], 'U_OTHER', True, c, path=env, now=1010)
    assert '대표님만' in msg and rr.load_request(env)['status'] == 'pending'


def test_approve_marks_approved_and_removes_buttons(env):
    c = FakeClient()
    req = rr.request_restart('이유', c, path=env, now=1000)
    assert rr.decide(req['id'], ADMIN, True, c, path=env, now=1010) == ''
    saved = rr.load_request(env)
    assert saved['status'] == 'approved' and saved['approved_at'] == 1010
    last = c.updates[-1]['blocks']
    assert _buttons(last) == [] and '승인됨' in last[0]['text']['text']
    # 다시 누르면 이미 처리됨
    assert '이미' in rr.decide(req['id'], ADMIN, True, c, path=env, now=1020)


def test_cancel(env):
    c = FakeClient()
    req = rr.request_restart('이유', c, path=env, now=1000)
    assert rr.decide(req['id'], ADMIN, False, c, path=env, now=1010) == ''
    assert rr.load_request(env)['status'] == 'cancelled'
    assert '취소됨' in c.updates[-1]['blocks'][0]['text']['text']


def test_approving_expired_request_marks_expired(env):
    c = FakeClient()
    req = rr.request_restart('이유', c, path=env, now=1000)
    msg = rr.decide(req['id'], ADMIN, True, c, path=env, now=1000 + rr.PENDING_TTL_SEC + 5)
    assert '만료' in msg and rr.load_request(env)['status'] == 'expired'


def test_wrong_request_id_rejected(env):
    c = FakeClient()
    rr.request_restart('이유', c, path=env, now=1000)
    assert '이미 처리' in rr.decide('otherid', ADMIN, True, c, path=env, now=1010)


def test_new_request_refused_while_approved(env):
    c = FakeClient()
    req = rr.request_restart('이유', c, path=env, now=1000)
    rr.decide(req['id'], ADMIN, True, c, path=env, now=1010)
    with pytest.raises(RuntimeError):
        rr.request_restart('또 요청', c, path=env, now=1020)


def test_reads_bom_file_written_by_watchdog(env):
    # 워치독(PowerShell 5.1 Set-Content -Encoding UTF8)은 BOM 을 붙여 쓴다
    env.write_bytes('﻿{"id": "x1", "status": "done", "reasons": ["가"]}'.encode('utf-8'))
    assert rr.load_request(env)['status'] == 'done'
    assert rr.active_request(env) is None


class _AnyApp:
    """slack_bolt App 대역 — 어떤 데코레이터든 받아서 action 핸들러만 모은다."""
    def __init__(self):
        self.actions = {}

    def action(self, name):
        def deco(fn):
            self.actions[name] = fn
            return fn
        return deco

    def __getattr__(self, _name):
        return lambda *a, **k: (lambda fn: fn)


def test_main_bot_buttons_wired(env, monkeypatch):
    import time as _t
    from dashboard.blueprints import slack_bot as sb
    app = _AnyApp()
    sb._register_handlers(app)
    assert rr.APPROVE_ACTION in app.actions and rr.CANCEL_ACTION in app.actions

    calls = []
    monkeypatch.setattr(rr, 'decide', lambda rid, uid, approve, client: calls.append((rid, uid, approve)) or '')
    body = {'user': {'id': ADMIN}, 'actions': [{'value': 'req1'}], 'channel': {'id': 'D_ADMIN'}}
    app.actions[rr.APPROVE_ACTION](ack=lambda: None, body=body, client=FakeClient())
    app.actions[rr.CANCEL_ACTION](ack=lambda: None, body=body, client=FakeClient())
    for _ in range(100):
        if len(calls) == 2:
            break
        _t.sleep(0.01)
    assert sorted(calls) == [('req1', ADMIN, False), ('req1', ADMIN, True)]


def test_last_server_start_reads_rotated_logs(monkeypatch, tmp_path):
    (tmp_path / 'itglobal_dashboard.log').write_text(
        '2026-10-28 08:10:00.1 [INFO] waitress:449 - Serving on http://0.0.0.0:5000\n', encoding='utf-8')
    (tmp_path / 'itglobal_dashboard.log.1').write_text(
        '2026-10-27 13:38:54.1 [INFO] waitress:449 - Serving on http://0.0.0.0:5000\n', encoding='utf-8')
    monkeypatch.setattr(rr, '_LOG_DIR', tmp_path)
    assert rr.last_server_start() == '2026-10-28 08:10:00'
