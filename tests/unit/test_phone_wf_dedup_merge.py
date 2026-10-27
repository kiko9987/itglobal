# -*- coding: utf-8 -*-
"""전화 워크플로 같은 번호 병합 회귀 테스트 (2026-10-27).

- L-04129: 공사 확정·방문 완료 리드에 새 방문이 병합 → 상태가 '방문 예약'으로 내려가고 이력 덮어씀
- L-04264: 병합이 상담 내용을 덮어써 부재중·유선 상담 이력 소실
- L-04264: '부천로86번길 34'와 '부천로 86번길 66'을 같은 주소로 봐 다른 현장이 병합
시트·Slack 미접촉 — 순수 함수 + 시트 manager mock.
"""
import sys

sys.path.insert(0, '.')

import pandas as pd
import pytest
import dashboard.services.lead_sync as ls


# ── 주소 core: N번길까지 ────────────────────────────────────────────────────

@pytest.mark.parametrize('a,b,same', [
    ('부천 원미구 부천로86번길 34 Proud APT', '부천시 원미구 심곡동 부천로 86번길 66 1층', False),
    ('부천 원미구 부천로86번길 34 Proud APT', '부천로 86번길 34 2층', True),
    ('중구 다동길43 한외빌딩 10층 하나핀크', '중구 다동길 43 한외빌딩 10층 하나핀크', True),
    ('강남구 선릉로76길 40 신호빌딩 1층', '강남구 선릉로76길 40', True),
    ('강남구 선릉로76길 40', '강남구 선릉로76길 42', False),
    ('성남 수정구 등자로 56', '등자로 56 3층', True),
])
def test_addr_core_beonbil(a, b, same):
    assert (ls._addr_core(a) == ls._addr_core(b)) is same


def test_addr_core_values():
    assert ls._addr_core('부천로86번길 34') == '부천로86번길34'
    assert ls._addr_core('선릉로76길 40') == '선릉로76길40'
    assert ls._addr_core('도산대로67길 7 파크빌딩') == '도산대로67길7'


def test_normalize_address_key_beonbil():
    assert ls._normalize_address_key('부천 원미구 부천로86번길 34 Proud APT') == '부천 원미구 부천로86번길 34'
    assert ls._normalize_address_key('남양주 진건읍 진건오남로 77 심미에셈빌 1층') == '남양주 진건읍 진건오남로 77'


# ── 상담 내용 이어 붙이기 ───────────────────────────────────────────────────

@pytest.fixture
def no_initial(monkeypatch):
    import dashboard.blueprints.slack_helpers as sh
    monkeypatch.setattr(sh, '_to_initial', lambda n: {'빈승정': 'SJ', '박용구': 'YG'}.get(n, ''))


def test_append_keeps_history(no_initial):
    old = '[10.15 14:51 MJ · 부재중] 통화 요청\n─────────\n[10.16 14:03 JSH · 유선 상담] 천1 견적 문의'
    got = ls._append_workflow_consult(old, '천장형1대 신규설치', {'온라인 상담자': '빈승정', '상태': '방문 예약'})
    assert got.startswith(old + '\n─────────\n[')
    assert got.endswith('SJ · 방문 예약] 천장형1대 신규설치')


@pytest.mark.parametrize('old', ['', '-', None])
def test_append_empty_old(no_initial, old):
    assert ls._append_workflow_consult(old, '새 내용', {}) == '새 내용'


def test_append_idempotent(no_initial):
    old = '1차 ─────────\n[10.23 13:19 YG · 방문 예약] 전력 정리'
    assert ls._append_workflow_consult(old, '전력 정리', {'온라인 상담자': '박용구'}) == old


def test_append_unknown_person(no_initial):
    got = ls._append_workflow_consult('옛 내용', '새', {'온라인 상담자': '@외부인', '상태': ''})
    assert got.endswith('외부인 · -] 새')


class _Mgr:
    def __init__(self):
        self.bodies = []
        mgr = self

        class _Req:
            def __init__(self, body):
                mgr.bodies.append(body)

            def execute(self):
                return {}

        class _Vals:
            def batchUpdate(self, spreadsheetId, body):
                return _Req(body)

        class _SS:
            def values(self):
                return _Vals()

        class _Svc:
            def spreadsheets(self):
                return _SS()
        self.service = _Svc()


def test_update_existing_appends_k(no_initial):
    m = _Mgr()
    wf = {'상태': '방문 예약', '방문 예정일': '2026-10-20', '상담 내용': '천장형1대 신규설치',
          '온라인 상담자': '빈승정'}
    ex = pd.Series({'상담 내용': '[10.16 14:03 JSH · 유선 상담] 천1 견적 문의'})
    ls._update_existing_lead_from_workflow(m, {'sheet_id': 'S', 'sheet_name': 'T'}, 10, wf, existing_row=ex)
    cells = {d['range']: d['values'][0][0] for d in m.bodies[0]['data']}
    k = cells["'T'!K10"]
    assert k.startswith('[10.16 14:03 JSH · 유선 상담] 천1 견적 문의\n─────────\n[')
    assert k.endswith('SJ · 방문 예약] 천장형1대 신규설치')
    assert cells["'T'!D10"] == '방문 예약'


def test_update_existing_without_row_overwrites_as_before():
    m = _Mgr()
    ls._update_existing_lead_from_workflow(m, {'sheet_id': 'S', 'sheet_name': 'T'}, 10,
                                           {'상담 내용': '새 내용'})
    assert m.bodies[0]['data'][0]['values'][0][0] == '새 내용'


# ── 병합 대상 가드 (sync 본문 소스 확인) ─────────────────────────────────────

def test_dedup_skips_done_or_confirmed_leads_source():
    import inspect
    src = inspect.getsource(ls)
    i = src.index("[SYNC/전화WF/DEDUP-GUARD] 같은 번호")
    j = src.index("_update_existing_lead_from_workflow(\n", i)
    block = src[i:j]
    assert "('공사 확정', '방문 완료')" in block and 'consult_visit_done' in block
    assert 'existing_row=main_df.loc[existing_sheet_row - 2]' in src[j:j + 300]
