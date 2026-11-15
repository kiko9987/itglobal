# -*- coding: utf-8 -*-
"""5차 보안 점검(2026-11-15) — 슬랙 봇 입구 '사무실 직원만' 문지기, 오래된 선입금 카드 ⚠️, 결재함 ⚠️.

#1 전자결재 바로가기·목록이 누구에게나 열려 거래처 계좌·계산서 글·공사 연락처가 보였다.
#3·#4 계산서·상담·공사 카드 버튼, /공사확정에 사람 확인이 없었다.
#2 그룹웨어 결재함에서 승인·이체 완료할 때 ⚠️ 가 안 보였다. #7 오래된 선입금 경고가 카드에 안 남았다.
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb
from dashboard.services import e_approval as ea
from dashboard.services import hr
from tests.unit.test_e_approval import (  # noqa: F401 — 픽스처 재사용
    FIELDS, KIKO, MGR, SB, _no_final_by_default, se, tmp_db)
from tests.unit.test_fund_dup_guards import _age


@pytest.fixture
def people(monkeypatch):
    monkeypatch.setattr(sb, '_STAFF_CACHE', {})
    cards = {'U_STAFF': {'access': 'staff'}, 'U_FIELD': {'access': 'field'}}
    monkeypatch.setattr(hr, 'find_by_slack', lambda uid: cards.get(uid))
    monkeypatch.setattr(sb, '_settlement_checker_ids', lambda: ('U_CHECK',))
    return cards


class _GateApp:
    def __init__(self):
        self.mw = None

    def middleware(self, fn):
        self.mw = fn
        return fn


class _Client:
    def __init__(self):
        self.calls = []

    def views_open(self, **k):
        self.calls.append(('views_open', k))

    def chat_postEphemeral(self, **k):
        self.calls.append(('ephemeral', k))

    def conversations_open(self, **k):
        return {'channel': {'id': 'D1'}}

    def chat_postMessage(self, **k):
        self.calls.append(('dm', k))


def _run(gate, body):
    acks, nexts = [], []
    client = _Client()
    r = gate(body=body, next=lambda: nexts.append(1), client=client,
             ack=lambda *a, **k: acks.append(k) or 'RESP')
    return r, acks, nexts, client


@pytest.fixture
def gate(people):
    app = _GateApp()
    sb.install_staff_gate(app, '시험봇', '전자결재')
    return app.mw


def test_staff_and_checker_pass(gate):
    for uid in ('U_STAFF', 'U_CHECK'):
        r, acks, nexts, _ = _run(gate, {'type': 'block_suggestion', 'user': {'id': uid}})
        assert nexts == [1] and acks == []


def test_events_and_view_closed_pass(gate):
    for body in ({'type': 'event_callback', 'event': {'type': 'app_home_opened', 'user': 'U_FIELD'}},
                 {'type': 'view_closed', 'user': {'id': 'U_FIELD'}}):
        r, acks, nexts, _ = _run(gate, body)
        assert nexts == [1]


def test_field_options_empty(gate):
    r, acks, nexts, _ = _run(gate, {'type': 'block_suggestion', 'user': {'id': 'U_FIELD'}})
    assert nexts == [] and acks == [{'options': []}] and r == 'RESP'


def test_field_submit_replaced_with_locked_view(gate):
    r, acks, nexts, _ = _run(gate, {'type': 'view_submission', 'user': {'id': 'U_FIELD'}})
    assert nexts == [] and acks[0]['response_action'] == 'update'
    assert '전자결재는 사무실 직원만' in acks[0]['view']['blocks'][0]['text']['text']


def test_resigned_or_unlinked_command_refused(gate):
    r, acks, nexts, _ = _run(gate, {'command': '/공사확정', 'user_id': 'U_GONE'})
    assert nexts == [] and ':lock:' in acks[0]['text']


def test_field_shortcut_gets_locked_modal(gate):
    import time
    r, acks, nexts, client = _run(gate, {'type': 'shortcut', 'user': {'id': 'U_FIELD'}, 'trigger_id': 'T1'})
    assert nexts == [] and acks == [{}]
    for _ in range(50):
        if client.calls:
            break
        time.sleep(0.02)
    assert client.calls and client.calls[0][0] == 'views_open'


def test_field_button_in_channel_gets_ephemeral(gate):
    import time
    r, acks, nexts, client = _run(gate, {'type': 'block_actions', 'user': {'id': 'U_FIELD'},
                                         'channel': {'id': 'C1'}, 'trigger_id': 'T1'})
    assert nexts == []
    for _ in range(50):
        if client.calls:
            break
        time.sleep(0.02)
    assert client.calls[0][0] == 'ephemeral'


def test_particle_follows_word(people):
    assert '이 기능은 사무실' in sb._slack_staff_denied('U_FIELD', '이 기능')
    assert '전자결재는 사무실' in sb._slack_staff_denied('U_FIELD', '전자결재')
    assert '이 명령은 사무실' in sb._slack_staff_denied('U_FIELD')


def test_cache_does_not_store_lookup_failure(people, monkeypatch):
    def boom(uid):
        raise RuntimeError('db down')
    monkeypatch.setattr(hr, 'find_by_slack', boom)
    assert '실패' in sb._slack_staff_denied('U_STAFF')
    monkeypatch.setattr(hr, 'find_by_slack', lambda uid: {'access': 'staff'})
    assert sb._slack_staff_denied('U_STAFF') is None


def test_every_business_bot_installs_gate():
    """새 봇·새 버튼이 생겨도 입구에서 막히게 — 업무 봇 초기화마다 문지기를 다는지."""
    src = open('dashboard/blueprints/slack_bot.py', encoding='utf-8').read()
    for app_var in ('_slack_app', '_project_slack_app', '_eapproval_slack_app', '_as_slack_app',
                    '_invoice_slack_app', '_payment_slack_app'):
        assert f'install_staff_gate({app_var},' in src, app_var


# ── #7 오래된 선입금 → 카드 ⚠️ ─────────────────────────────────────

def test_old_prepaid_marked_on_card(se):
    p = se.process_new_submit(KIKO, dict(FIELDS))
    se.apply_action('transfer', p['id'], SB)
    _age(p['id'], ea.DUP_WINDOW_DAYS + 30)
    late = ea.create(ea.FORM_FUND, MGR, '강성환',
                     dict(FIELDS, grn='R', origin_channel='CINV', origin_ts='1700000000.000777'))
    assert ea.extra(late).get('old_prepaid') == p['id']
    lines = ea.fund_warning_lines(late)
    assert any('선입금' in x and ea.title(p) in x for x in lines)
    plain = ea.create(ea.FORM_FUND, MGR, '강성환', dict(FIELDS))   # 계산서 없이 올린 건은 표시 안 함
    assert 'old_prepaid' not in ea.extra(plain)


# ── #2 결재함 상세에 ⚠️ ──────────────────────────────────────────

def test_groupware_detail_shows_warnings_to_approvers_only(se, monkeypatch):
    from dashboard.services import gw_approval as gw
    import dashboard.services.gw_home as gh
    monkeypatch.setattr(gh, '_card_link', lambda row: '')
    a = se.process_new_submit(KIKO, dict(FIELDS))
    b = se.process_new_submit(MGR, dict(FIELDS))          # 같은 거래처·금액 → ⚠️
    assert ea.fund_warning_lines(ea.get(b['id']))
    approver = {'see_all': True, 'slack_id': KIKO, 'email': 'k@test.invalid'}
    other = {'see_all': False, 'slack_id': 'U_OTHER', 'email': 'o@test.invalid'}
    monkeypatch.setattr(gw, '_can_view', lambda row, v: True)
    d = gw.detail(b['id'], approver)
    assert d['warnings'] and any('같은 거래처·금액' in w for w in d['warnings'])
    assert gw.detail(b['id'], other)['warnings'] == []
