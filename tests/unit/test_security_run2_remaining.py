# -*- coding: utf-8 -*-
"""보안 점검 2차 남은 항목 (2026-11-13).

#8 슬랙 작성자 이름: 인사카드 연결(슬랙 ID) 우선, 미연결 계정이 직원 이름·이니셜을 표시 이름으로 쓰면 비움
#6 플렉스 옛 결재 문서: 생년월일이 다른 동명이인이 있으면 입사일 이후 문서만
#5 실시간 소켓: 연결 뒤에도 세션 재확인 → 끝난 세션의 연결은 끊음
"""
import sys
from contextlib import contextmanager
from datetime import datetime, timedelta

sys.path.insert(0, '.')

import pytest

from dashboard.blueprints import slack_helpers as sh
from dashboard.services import gw_approval as ga
from dashboard.services import hr
from dashboard.utils import socket_session_guard as sg


# ── #8 작성자 이름 ─────────────────────────────────────────

class _Client:
    def __init__(self, email, display):
        self.p = {'email': email, 'display_name': display, 'real_name': ''}

    def users_info(self, user):
        return {'ok': True, 'user': {'profile': self.p}}


@pytest.fixture
def staff(monkeypatch):
    monkeypatch.setenv('SALES_EMAILS', '{}')
    cards = {'U_MW': {'name': '박민우'}}
    monkeypatch.setattr(hr, 'find_by_slack', lambda uid: cards.get(uid))
    monkeypatch.setattr(hr, 'list_people', lambda include_resigned=False: [{'name': '박민우'}, {'name': '박정우'}])
    monkeypatch.setattr(sh, '_lookup_initial_by_name', lambda n: '')
    monkeypatch.setattr(sh, '_load_initials_from_config', lambda: {'박민우': 'MW', '박정우': 'JW'})


def test_linked_card_name_wins_over_display(staff):
    assert sh._slack_user_to_korean_name(_Client('mw@x', '박정우'), 'U_MW') == '박민우'


@pytest.mark.parametrize('display', ['박정우', 'JW', 'jw'])
def test_unlinked_impersonation_blanked(staff, display):
    assert sh._slack_user_to_korean_name(_Client('guest@x', display), 'U_GUEST') == ''


def test_unlinked_unique_display_kept(staff):
    assert sh._slack_user_to_korean_name(_Client('guest@x', '외부 협력사'), 'U_GUEST') == '외부 협력사'


# ── #6 플렉스 동명이인 ────────────────────────────────────

def _p(**k):
    return dict({'id': 1, 'name': '홍길동', 'birth_date': '1990-01-01', 'hire_date': '2025-03-01'}, **k)


@pytest.mark.parametrize('others,expected', [
    ([], ''),                                                                   # 동명이인 없음
    ([{'id': 2, 'name': '홍길동', 'birth_date': '1990-01-01'}], ''),            # 같은 사람 재입사
    ([{'id': 2, 'name': '홍길동', 'birth_date': '1985-05-05'}], '2025-03-01'),  # 다른 사람 → 입사일 이후만
])
def test_archive_since(monkeypatch, others, expected):
    monkeypatch.setattr(hr, 'same_name_people', lambda n: [_p()] + others)
    assert ga._archive_since(_p()) == expected


def test_archive_since_without_hire_date_hides_all(monkeypatch):
    monkeypatch.setattr(hr, 'same_name_people', lambda n: [{'id': 2, 'birth_date': '1985-05-05'}])
    assert ga._archive_since(_p(hire_date='')) == '9999'


class _Cur:
    def __init__(self, row):
        self.row = row

    def fetchone(self):
        return self.row


def test_archive_detail_respects_since(monkeypatch):
    row = ('F-1', 't', 'form', '완료', '홍길동', '2024-06-01 10:00', '', '', '', '{}')

    @contextmanager
    def fake():
        yield type('C', (), {'execute': lambda _s, *a: _Cur(row)})()
    monkeypatch.setattr(ga, '_aconn', fake)
    v = {'see_all': False, 'name': '홍길동', 'archive_since': '2025-03-01'}
    assert ga.archive_detail('F-1', v) is None                 # 동명이인의 입사 전 문서
    assert ga.archive_detail('F-1', dict(v, archive_since=''))['doc_no'] == 'F-1'


# ── #5 소켓 세션 재확인 ────────────────────────────────────

class _DB:
    def __init__(self, users):
        self.users = users

    def get_user_by_email(self, e):
        return self.users.get(e)


NOW = datetime(2026, 11, 13, 17, 0)


def test_session_still_valid_rules():
    lt = (NOW - timedelta(hours=1)).isoformat()
    db = _DB({'a@x': {'is_active': True}, 'off@x': {'is_active': False}, 'res@x': {'is_resigned': True},
              'inv@x': {'session_invalidate_after': (NOW - timedelta(minutes=10)).strftime('%Y-%m-%d %H:%M:%S')}})
    assert sg.session_still_valid('a@x', lt, now=NOW, db=db)
    assert not sg.session_still_valid('a@x', (NOW - timedelta(hours=5)).isoformat(), now=NOW, db=db)   # 4시간 초과
    for e in ('off@x', 'res@x', 'inv@x', 'gone@x'):
        assert not sg.session_still_valid(e, lt, now=NOW, db=db)


def test_db_error_keeps_connection():
    class Boom:
        def get_user_by_email(self, e):
            raise RuntimeError('db')
    assert sg.session_still_valid('a@x', NOW.isoformat(), now=NOW, db=Boom())


def test_sweep_disconnects_only_ended_sessions():
    sg._conns.clear()
    sg.register('s1', 'a@x', (NOW - timedelta(hours=1)).isoformat())
    sg.register('s2', 'res@x', (NOW - timedelta(hours=1)).isoformat())
    cut = []
    db = _DB({'a@x': {'is_active': True}, 'res@x': {'is_resigned': True}})
    assert sg.sweep(cut.append, now=NOW, db=db) == ['s2']
    assert cut == ['s2'] and set(sg.tracked()) == {'s1'}
    sg.unregister('s1')
