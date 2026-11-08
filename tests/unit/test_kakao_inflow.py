# -*- coding: utf-8 -*-
"""홈페이지 카카오 버튼 클릭(/track/kakao-click) ↔ 채널톡 카카오톡 문의 유입 짝짓기 (2026-11-08 L-04501)."""
import json
from unittest.mock import patch

import pytest
from flask import Flask

from dashboard.blueprints import call_track as ct
from dashboard.blueprints import channeltalk as cht

T0 = 1_794_137_000.0


class _FakeRedis:
    def __init__(self):
        self.kv, self.h, self.z = {}, {}, {}

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.kv:
            return None
        self.kv[k] = v
        return True

    def pipeline(self):
        return self

    def hset(self, k, key=None, value=None, mapping=None):
        d = self.h.setdefault(k, {})
        if mapping:
            d.update(mapping)
        if key is not None:
            d[key] = value

    def hsetnx(self, k, key, value):
        d = self.h.setdefault(k, {})
        if key in d:
            return 0
        d[key] = value
        return 1

    def expire(self, *a):
        pass

    def zadd(self, k, m):
        self.z.setdefault(k, {}).update(m)

    def zremrangebyscore(self, *a):
        pass

    def execute(self):
        pass

    def zrangebyscore(self, k, lo, hi):
        return [m for m, s in self.z.get(k, {}).items() if lo <= s <= hi]

    def hgetall(self, k):
        return self.h.get(k, {})


@pytest.fixture
def fake():
    f = _FakeRedis()
    with patch.object(ct, 'get_redis_client') as g:
        g.return_value.redis = f
        yield f


def _click(fake, cid, ts, inflow='', ref=''):
    fake.h[f'kakao_click:{cid}'] = {'inflow': inflow, 'ref': ref, 'page': '/', 'ts': str(int(ts))}
    fake.z.setdefault(ct.KAKAO_CLICK_INDEX, {})[cid] = ts


# ── 비콘 엔드포인트 ──────────────────────────────────────────

def test_beacon_records_kakao_click_separately_from_calls(fake):
    app = Flask(__name__)
    app.register_blueprint(ct.call_track_bp)
    c = app.test_client()
    r = c.post('/track/kakao-click', data=json.dumps({'inflow': 'naver/시스템에어컨/피부과 에어컨', 'btn': 'itg-contact-item'}),
               headers={'Origin': 'https://www.itg-aircon.com', 'Content-Type': 'text/plain;charset=UTF-8'})
    assert r.status_code == 200 and r.get_json()['ok']
    assert len(fake.z[ct.KAKAO_CLICK_INDEX]) == 1
    assert ct.CALL_CLICK_INDEX not in fake.z            # 전화 클릭 집계에 섞이지 않음
    r2 = c.post('/track/kakao-click', data='{}', headers={'Origin': 'https://evil.example'})
    assert r2.status_code == 403


# ── 마커 규칙 (홈페이지 폼 '유입경로' 칸과 같은 규칙) ──────────────

@pytest.mark.parametrize('rec, marker', [
    ({'inflow': 'naver/시스템에어컨/피부과 에어컨', 'ref': 'm.search.naver.com'}, '네이버/시스템에어컨/피부과에어컨'),
    ({'inflow': 'daangn/web_install/D'}, '당근/web_install/D'),
    ({'inflow': '', 'ref': 'm.search.naver.com'}, 'ref/m.search.naver.com'),
    ({'inflow': '', 'ref': ''}, 'direct'),
])
def test_marker(rec, marker):
    assert ct.kakao_click_marker(rec) == marker


# ── 짝짓기 ────────────────────────────────────────────────

def test_match_click_minutes_before_message(fake):
    _click(fake, 'a1', T0 - 120, inflow='google/냉난방기설치')
    out = ct.claim_kakao_click(T0, 'chat:x')
    assert out['status'] == 'match' and out['marker'] == '구글/냉난방기설치' and out['cid'] == 'a1'
    assert fake.h['kakao_click:a1']['claim'] == 'chat:x'


def test_claimed_click_not_reused(fake):
    _click(fake, 'a1', T0 - 60, inflow='google')
    assert ct.claim_kakao_click(T0, 'chat:1')['status'] == 'match'
    assert ct.claim_kakao_click(T0 + 30, 'chat:2')['status'] == 'none'


def test_no_click_in_window_is_none(fake):
    _click(fake, 'old', T0 - ct.KAKAO_MATCH_BEFORE - 60, inflow='google')   # 너무 오래전
    out = ct.claim_kakao_click(T0, 'chat:x')
    assert out['status'] == 'none' and out['marker'] == ''


def test_different_inflows_overlap_is_ambiguous(fake):
    _click(fake, 'a1', T0 - 300, inflow='google')
    _click(fake, 'a2', T0 - 100, inflow='daangn/web_install/D')
    out = ct.claim_kakao_click(T0, 'chat:x')
    assert out['status'] == 'ambiguous' and out['n'] == 2 and not out['marker']
    assert 'claim' not in fake.h['kakao_click:a1'] and 'claim' not in fake.h['kakao_click:a2']


def test_same_inflow_overlap_takes_latest(fake):
    _click(fake, 'a1', T0 - 300, ref='m.search.naver.com')
    _click(fake, 'a2', T0 - 100, ref='m.search.naver.com')
    out = ct.claim_kakao_click(T0, 'chat:x')
    assert out['status'] == 'match' and out['cid'] == 'a2' and out['marker'] == 'ref/m.search.naver.com'


def test_tag_lead(fake):
    _click(fake, 'a1', T0 - 60, inflow='google')
    ct.tag_kakao_click('a1', 'L-09999')
    assert fake.h['kakao_click:a1']['lead'] == 'L-09999'


# ── 채널톡 쪽 ─────────────────────────────────────────────

def test_only_kakao_chats_are_matched(fake):
    _click(fake, 'a1', T0 - 60, inflow='google')
    web = {'id': 'c1', 'mediumProfile': {'mediumName': 'web'}}
    assert cht._kakao_inflow(web, int(T0 * 1000)) == {}
    assert 'claim' not in fake.h['kakao_click:a1']
    kakao = {'id': 'c2', 'mediumProfile': {'mediumName': 'appKakao'}}
    assert cht._kakao_inflow(kakao, int(T0 * 1000))['status'] == 'match'


def _card(kakao):
    uc = {'mediumProfile': {'mediumName': 'appKakao'}}
    return cht._new_chat_card(uc, {'name': '리본 399'}, '피부과 견적', int(T0 * 1000), lead_no='L-09999', kakao=kakao)['text']


def test_card_matched_shows_channel_and_keyword():
    t = _card({'status': 'match', 'marker': '네이버/시스템에어컨/피부과에어컨'})
    assert '새 문의 접수 알림 - 온라인 (카카오톡 · 네이버)*' in t
    assert '유입 경로 : 홈페이지 카카오 버튼' in t
    assert '검색 키워드 : 피부과에어컨' in t


def test_card_unmatched_and_ambiguous():
    t = _card({'status': 'none', 'n': 0})
    assert '온라인 (카카오톡)*' in t and '홈페이지 카카오 버튼 기록 없음' in t
    t2 = _card({'status': 'ambiguous', 'n': 2})
    assert '확인 불가' in t2 and '2건' in t2


def test_card_without_kakao_info_unchanged():
    t = _card(None)
    assert '유입 경로' not in t and '온라인 (카카오톡)*' in t


def test_register_lead_puts_marker_first_line():
    captured = {}

    def _append(leads):
        captured.update(leads[0])
        return ['L-09999']
    with patch('dashboard.services.lead_sync._append_leads_to_main', _append):
        uc = {'mediumProfile': {'mediumName': 'appKakao'}}
        no = cht._register_chat_lead(uc, {'name': 'x'}, '피부과 견적', int(T0 * 1000), inflow_marker='구글/냉난방기설치')
    assert no == 'L-09999'
    assert captured['문의 내용'] == '[유입:구글/냉난방기설치]\n피부과 견적'
    from dashboard.services.lead_helpers import split_inflow_marker
    assert split_inflow_marker(captured['문의 내용']) == ('구글/냉난방기설치', '피부과 견적')
