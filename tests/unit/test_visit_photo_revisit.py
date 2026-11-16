# -*- coding: utf-8 -*-
"""방문 사진 자동 완료·재방문 카드·List (2026-11-16 상태 머신 점검 A9·A14).

① 사진 자동 완료는 [방문 완료] 버튼이 살아 있는 진행 중 카드에서만 — 취소 카드는 그대로.
② 재방문인데 방문 List 에 행이 없으면(완료·취소로 지워짐) update 대신 add 웹훅.
③ 재방문 때 옛 카드가 이미 끝난 방문이면 새 카드를 보냄(진행 중 카드는 예전처럼 고침).
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb


class _RC:
    def __init__(self):
        self.d = {}

    def get(self, k):
        return self.d.get(k)

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.d:
            return False
        self.d[k] = v
        return True

    def setex(self, k, ttl, v):
        self.d[k] = v

    def delete(self, *ks):
        for k in ks:
            self.d.pop(k, None)


@pytest.fixture
def rc(monkeypatch):
    import dashboard.utils.redis_client as rcm
    r = _RC()
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': r})())
    return r


_ACTIVE = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': '>*방문일* : 11/15\n>이름 : 홍길동'}},
           {'type': 'actions', 'elements': [{'action_id': 'visit_edit_info'}, {'action_id': 'visit_complete'},
                                            {'action_id': 'visit_cancel'}]}]
_CANCELLED = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': '⚪ 방문 취소'}},
              {'type': 'actions', 'elements': [{'action_id': 'visit_uncancel'}]}]


class _Client:
    def __init__(self, blocks):
        self.blocks, self.updates, self.posts = blocks, [], []

    def conversations_replies(self, **k):
        return {'messages': [{'text': '카드', 'blocks': self.blocks}]}

    def chat_update(self, **k):
        self.updates.append(k)

    def chat_postMessage(self, **k):
        self.posts.append(k)
        return {'ts': '2.2'}

    def chat_postEphemeral(self, **k):
        pass

    def users_info(self, **k):
        return {'user': {}}


# ── ① 사진 자동 완료 ──

@pytest.fixture
def auto(monkeypatch, rc):
    hooks = []
    monkeypatch.setattr(sb, '_visit_not_yet_reached', lambda *a: False)
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', lambda *a, **k: hooks.append(a[0]))
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda *a: 'JW')

    def _run(blocks, status='방문 예약'):
        monkeypatch.setattr(sb, '_find_lead_by_no', lambda *a, **k: {'상태': status, '방문 예정일': '2026-11-15'})
        c = _Client(blocks)
        try:
            sb._auto_complete_visit_from_photo_locked(c, 'C_VISIT', '1.1', 'L-04300', 'U1')
        except Exception:
            pass                                                   # 시트 마커 등 뒷단계는 범위 밖
        return c
    return type('A', (), {'run': staticmethod(_run), 'hooks': hooks})


def test_photo_completes_active_card(auto):
    c = auto.run(_ACTIVE)
    assert auto.hooks == ['SLACK_VISIT_COMPLETE_WEBHOOK_URL'] and c.updates


def test_photo_on_cancelled_card_does_nothing(auto):
    c = auto.run(_CANCELLED)
    assert auto.hooks == [] and c.updates == []


def test_photo_when_sheet_says_cancelled(auto):
    c = auto.run(_ACTIVE, status='방문 취소')
    assert auto.hooks == [] and c.updates == []


def test_photo_when_cancel_flag_set(auto, rc):
    rc.d['visit_cancelled:L-04300'] = '1'
    c = auto.run(_ACTIVE)
    assert auto.hooks == [] and c.updates == []


def test_photo_after_construction_confirmed_still_completes(auto):
    """공사 확정 뒤 다녀온 방문도 카드가 살아 있으면 완료 (시트 상태는 보호 규칙이 지킴)."""
    auto.run(_ACTIVE, status='공사 확정')
    assert auto.hooks == ['SLACK_VISIT_COMPLETE_WEBHOOK_URL']


# ── ② List add/update ──

@pytest.fixture
def lst(monkeypatch, rc):
    sent = []
    monkeypatch.setenv('SLACK_LIST_WEBHOOK_URL', 'https://add')
    monkeypatch.setenv('SLACK_LIST_UPDATE_WEBHOOK_URL', 'https://update')

    class _Resp:
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

        def read(self):
            return b''
    monkeypatch.setattr(sb.urllib.request, 'urlopen', lambda req, timeout=5: sent.append(req.full_url) or _Resp())
    return sent


def _post(action='visit'):
    sb._post_to_slack_list(_Client([]), {'리드 No': 'L-04300', '고객 연락처': '010-1234-5678'},
                           {'visit_date': '2026-11-20', 'visit_address': '', 'consultation': '', 'estimate': ''},
                           channel='', message_ts='', action=action)


def test_revisit_with_row_gone_adds(lst, rc, monkeypatch):
    rc.d['slack_list_posted:L-04300'] = '1'
    monkeypatch.setattr(sb, '_vlist_sync_row_contact', lambda *a: False)
    _post()
    assert lst == ['https://add']


def test_revisit_with_row_present_updates(lst, rc, monkeypatch):
    rc.d['slack_list_posted:L-04300'] = '1'
    monkeypatch.setattr(sb, '_vlist_sync_row_contact', lambda *a: True)
    _post()
    assert lst == ['https://update']


def test_price_keeps_old_rule(lst, rc, monkeypatch):
    rc.d['slack_list_posted:L-04300'] = '1'
    monkeypatch.setattr(sb, '_vlist_sync_row_contact', lambda *a: (_ for _ in ()).throw(AssertionError()))
    _post('price')
    assert lst == ['https://update']


# ── ③ 재방문 카드 ──

@pytest.mark.parametrize('old_blocks,expect', [(_CANCELLED, 'new'), (_ACTIVE, 'update')])
def test_revisit_card(monkeypatch, rc, old_blocks, expect):
    monkeypatch.setenv('SLACK_VISIT_CHANNEL', 'C_VISIT')
    monkeypatch.setattr(sb, '_visit_slack_handler', object())
    monkeypatch.setattr(sb, '_visit_slack_app', None)
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda *a: 'JW')
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda *a, **k: {})
    rc.d['visit_notice_msg:L-04300'] = 'C_VISIT|1.1'
    c = _Client(old_blocks)
    ch, ts = sb._post_visit_notice(c, 'L-04300', '온라인', 'U1', '2026-11-20', '홍길동', '010', '주소', '상담')
    if expect == 'new':
        assert c.posts and not c.updates and ts == '2.2' and rc.d['visit_notice_msg:L-04300'] == 'C_VISIT|2.2'
    else:
        assert c.updates and not c.posts and ts == '1.1'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
