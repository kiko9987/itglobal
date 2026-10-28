# -*- coding: utf-8 -*-
"""2026-10-28 5차 — 문의 유입·A/S·캔버스 배정·정합성 점검 추가."""
import json
import sys
from datetime import date, datetime

sys.path.insert(0, '.')

import pytest


def test_pending_payload_keeps_previous_leads_as_list():
    from dashboard.services import lead_sync as ls
    lead = {'고객명': '홍길동', '_meta_previous_leads': [
        {'lead_no': 'L-00001', 'consult_dt': datetime(2026, 10, 1, 9, 0), 'consult_time': '2026.10.01. 09:00'}]}
    raw = ls._serialize_pending_payload('homepage', lead)
    back = json.loads(raw)['lead']['_meta_previous_leads']
    assert isinstance(back, list) and back[0]['lead_no'] == 'L-00001'


def test_repeat_section_tolerates_old_string_payload():
    from dashboard.services import lead_sync as ls
    # 옛 대기열에 글자로 저장된 값 — 예외 없이 처리
    assert isinstance(ls._build_repeat_section({'_meta_previous_leads': "[{'lead_no': 'L-1'}]"}), str)
    s = ls._build_repeat_section({'_meta_previous_leads': json.dumps([{'lead_no': 'L-00009', 'consult_time': 'x'}])})
    assert 'L-00009' in s


def test_lead_inflow_gaps_business_days():
    from dashboard.services.invariant_checks import lead_inflow_gaps
    leads = [{'플랫폼': '당근', '상담 시간': '2026.10.23. 10:00'},      # 금
             {'플랫폼': '홈페이지', '상담 시간': '2026.10.27. 10:00'},   # 화
             {'플랫폼': '전화', '상담 시간': '2026.10.28. 09:00'}]
    gaps = {p: g for p, _, g in lead_inflow_gaps(leads, today=date(2026, 10, 28))}   # 수
    assert gaps == {'당근': 2}          # 월·화 두 영업일 비었음, 홈페이지는 화요일 이후 0일


def test_overpay_lingering_flags_old_negative_unpaid():
    from dashboard.services.invariant_checks import check_overpay_lingering
    recs = [
        {'프로젝트 코드': 'R4182-SJ', '미수금': -5000,
         '잔금_메모': '2026/10/17\n하나\n입금4,950,000원\n⚠️ 과입금: 초과 5,000원'},
        {'프로젝트 코드': 'G1', '미수금': -5000, '잔금_메모': '2026/10/26\n⚠️ 과입금'},   # 아직 7일 안 됨
        {'프로젝트 코드': 'G2', '미수금': -1},                                               # 1원 반올림
    ]
    v = check_overpay_lingering({'records': recs, 'today': date(2026, 10, 28)})
    assert [x['title'] for x in v] == ['R4182-SJ'] and '11일째' in v[0]['detail']


def test_as_create_stops_when_number_unknown(monkeypatch):
    from dashboard.services import as_service as asv
    monkeypatch.setattr(asv, '_get_sheet', lambda: (None, 's', 'n'))

    def _boom():
        raise RuntimeError('읽기 실패')
    monkeypatch.setattr(asv, 'get_next_as_no', _boom)
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: (_ for _ in ()).throw(RuntimeError('no redis')))
    assert asv.create_as_row('G1', 'a', 'w', '', 'r', 'JW') == ('', None)


def test_as_number_uses_watermark(monkeypatch):
    from dashboard.services import as_service as asv
    store = {'as_no_seq': '12'}

    class R:
        def set(self, k, v, nx=False, ex=None):
            if nx and k in store:
                return False
            store[k] = str(v)
            return True

        def get(self, k):
            return store.get(k)

        def delete(self, k):
            store.pop(k, None)
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': R()})())
    monkeypatch.setattr(asv, '_get_sheet', lambda: (None, 's', 'n'))
    monkeypatch.setattr(asv, 'get_next_as_no', lambda: 'AS-0010')   # 시트는 10까지, 워터마크는 12
    got = {}
    monkeypatch.setattr(asv, '_append_as_row', lambda m, s, n, as_no, *a: (got.setdefault('no', as_no), 5))
    no, row = asv.create_as_row('G1', 'a', 'w', '', 'r', 'JW')
    assert no == 'AS-0013' and row == 5 and store['as_no_seq'] == '13' and 'as_no_lock' not in store


def test_canvas_unreadable_initials_and_korean_names():
    from dashboard.services import visit_assignment_sync as vas
    i2n = {'JW': '박정우', 'MJ': '박민재'}
    assert vas._extract_lead_initials('박민재 (SJ) 010-1234-5678 / 주소', i2n) == ['MJ']
    assert vas._unread_initial_tokens('JH (SJ) 010-1234-5678 / 주소', i2n) == ['JH']
    assert vas._unread_initial_tokens('jw (SJ) 010-1234-5678', i2n) == []        # 소문자도 명단에 있으면 OK
