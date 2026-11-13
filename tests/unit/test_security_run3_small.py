# -*- coding: utf-8 -*-
"""보안 점검 3차 작은 항목 (2026-11-14).

#3 방문 일정 캔버스: 고객 주소·문의의 마크다운 링크·멘션·줄바꿈 중화
#6 조합형 표시 이름: 인사카드 미연결 계정이 '박용구,미정' 처럼 묶어 써도 직원 이름으로 기록되지 않음
#5 플렉스 옛 결재 문서: 목록 SQL 이 근무 기간 창으로 거른다 (실제 SQLite)
"""
import sqlite3
from contextlib import contextmanager

import pytest

from dashboard.blueprints import slack_helpers as sh
from dashboard.services import gw_approval as ga
from dashboard.services import hr
from dashboard.services.lead_helpers import canvas_md_safe
from dashboard.services.visit_canvas_sync import _render_item


# ── #3 캔버스 ──────────────────────────────────────────────

def test_canvas_md_safe_breaks_links_mentions_and_lines():
    out = canvas_md_safe('역삼동 (2층) [견적서](https://x.example) ![](@U1) <https://y|z> - 가짜\r\n끝')
    assert '](' not in out and '![' not in out and '<' not in out and '>' not in out
    assert '\n' not in out and '\r' not in out and ' ' not in out
    assert out.startswith('역삼동 (2층) ')   # 보통 괄호·글자는 그대로


@pytest.mark.parametrize('platform', ['채널톡', '거래처'])
def test_render_item_neutralizes_customer_text(platform):
    lead = {'리드 No': 'L-00001', '플랫폼': platform, '방문 예정일': '2026-11-20',
            '고객 연락처': '010-1234-5678', '방문 주소': '서울 [지도](https://x.example)\n- 가짜 줄',
            '문의 내용': '천장형 ![](@U123) <!channel>', '온라인 상담자': '고광일'}
    out = _render_item(lead, {'고광일': 'KIKO'})
    assert '](' not in out and '![' not in out and '<' not in out and '\n' not in out
    assert '010-1234-5678' in out and ' / ' in out   # 배정 파서가 쓰는 구분자·연락처는 그대로


# ── #6 조합형 표시 이름 ────────────────────────────────────

class _Client:
    def __init__(self, display):
        self.p = {'email': 'guest@x', 'display_name': display, 'real_name': ''}

    def users_info(self, user):
        return {'ok': True, 'user': {'profile': self.p}}


@pytest.fixture
def staff(monkeypatch):
    monkeypatch.setenv('SALES_EMAILS', '{}')
    monkeypatch.setattr(hr, 'find_by_slack', lambda uid: None)
    monkeypatch.setattr(hr, 'list_people', lambda include_resigned=False: [{'name': '박용구'}, {'name': '박정우'}])
    monkeypatch.setattr(sh, '_lookup_initial_by_name', lambda n: '')


@pytest.mark.parametrize('display', ['박용구,미정', '미정/박정우', ' 외부 , 박용구 '])
def test_combined_display_name_blanked(staff, display):
    assert sh._slack_user_to_korean_name(_Client(display), 'U_GUEST') == ''


def test_unrelated_combined_name_kept(staff):
    assert sh._slack_user_to_korean_name(_Client('외부,협력사'), 'U_GUEST') == '외부,협력사'


# ── #5 플렉스 목록 SQL ─────────────────────────────────────

def test_archive_list_sql_uses_windows(monkeypatch):
    conn = sqlite3.connect(':memory:')
    conn.execute(ga._ARCHIVE_DDL)
    docs = [('F-1', '2019-05-01 10:00'),   # 같은 이름의 옛 작성자(입사 전)
            ('F-2', '2023-06-01 10:00'),   # 동명이인 근무 기간과 겹침
            ('F-3', '2025-06-01 10:00')]   # 본인
    conn.executemany("INSERT INTO flex_approval_archive (doc_no, title, form, status, writer, created_at) "
                     "VALUES (?, 't', '지출결의', '완료', '홍길동', ?)", docs)

    @contextmanager
    def fake():
        yield conn
    monkeypatch.setattr(ga, '_aconn', fake)
    v = {'see_all': False, 'name': '홍길동',
         'archive_windows': {'mine': [('2020-01-01', '9999-12-31')], 'others': [('2023-01-01', '2023-12-31')]}}
    r = ga.archive_list(v)
    assert [i['doc_no'] for i in r['items']] == ['F-3'] and r['total'] == 1 and r['forms'] == ['지출결의']
    assert ga.archive_list(dict(v, see_all=True))['total'] == 3
