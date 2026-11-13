# -*- coding: utf-8 -*-
"""보안 점검 2차 남은 항목 (2026-11-13).

#8 슬랙 작성자 이름: 인사카드 연결(슬랙 ID) 우선, 미연결 계정이 직원 이름·이니셜을 표시 이름으로 쓰면 비움
#6 플렉스 옛 결재 문서: 이름 + 본인 근무 기간(11-14, 동명이인 근무 기간은 숨김)
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


@pytest.mark.parametrize('others,mine,blocked', [
    ([], [('2025-03-01', '9999-12-31')], []),                                    # 동명이인 없음 → 본인 근무 기간
    ([{'id': 2, 'name': '홍길동', 'birth_date': '1990-01-01', 'hire_date': '2020-01-01', 'resign_date': '2021-12-31'}],
     [('2025-03-01', '9999-12-31'), ('2020-01-01', '2021-12-31')], []),          # 같은 사람 재입사 → 예전 기간도 본인
    ([{'id': 2, 'name': '홍길동', 'birth_date': '1985-05-05', 'hire_date': '2019-01-01', 'last_work_date': '2024-12-31'}],
     [('2025-03-01', '9999-12-31')], [('2019-01-01', '2024-12-31')]),             # 다른 사람 → 그 사람 기간은 숨김
    ([{'id': 2, 'name': '홍길동', 'birth_date': '1985-05-05'}],
     [('2025-03-01', '9999-12-31')], [('0000-00-00', '9999-12-31')]),             # 기간 모르는 동명이인 → 전부 숨김
])
def test_archive_windows(monkeypatch, others, mine, blocked):
    monkeypatch.setattr(hr, 'same_name_people', lambda n: [_p()] + others)
    assert ga._archive_windows(_p()) == {'mine': mine, 'others': blocked}


def test_archive_windows_without_hire_date_hides_all(monkeypatch):
    monkeypatch.setattr(hr, 'same_name_people', lambda n: [])
    w = ga._archive_windows(_p(hire_date=''))
    assert w == {'mine': [], 'others': []} and ga._window_sql(w) == ('0', [])
    assert not ga._in_windows('2025-06-01 10:00', w)


class _Cur:
    def __init__(self, row):
        self.row = row

    def fetchone(self):
        return self.row


def _detail(monkeypatch, created_at, w):
    row = ('F-1', 't', 'form', '완료', '홍길동', created_at, '', '', '', '{}')

    @contextmanager
    def fake():
        yield type('C', (), {'execute': lambda _s, *a: _Cur(row)})()
    monkeypatch.setattr(ga, '_aconn', fake)
    return ga.archive_detail('F-1', {'see_all': False, 'name': '홍길동', 'archive_windows': w})


def test_archive_detail_respects_windows(monkeypatch):
    w = {'mine': [('2025-03-01', '9999-12-31')], 'others': []}
    assert _detail(monkeypatch, '2024-06-01 10:00', w) is None             # 입사 전 = 같은 이름의 옛 작성자
    assert _detail(monkeypatch, '2025-03-01 09:00', w)['doc_no'] == 'F-1'  # 입사 당일부터
    w2 = {'mine': [('2020-01-01', '9999-12-31')], 'others': [('2023-01-01', '2023-12-31')]}
    assert _detail(monkeypatch, '2023-06-01 10:00', w2) is None            # 동명이인 근무 기간과 겹침
    assert _detail(monkeypatch, '2024-06-01 10:00', w2)['doc_no'] == 'F-1'
    w3 = {'mine': [('2020-01-01', '2021-12-31')], 'others': []}
    assert _detail(monkeypatch, '2021-12-31 18:00', w3)['doc_no'] == 'F-1'  # 퇴사일 그날까지


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
