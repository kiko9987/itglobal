# -*- coding: utf-8 -*-
"""2026-11-06 새 프로젝트 수동 입력 — 최근 방문 리드 대조."""
import sys
from datetime import date

sys.path.insert(0, '.')

from dashboard.services.lead_project_match import match_visited_leads

TODAY = date(2026, 11, 6)


def _lead(no, status='방문 완료', visit='2026-11-04', phone='010-1234-5678',
          addr='부천 오정구 석천로 397 3층', platform='전화', consult='2026-11-01 10:00'):
    return {'리드 No': no, '상태': status, '방문 예정일': visit, '고객 연락처': phone,
            '방문 주소': addr, '플랫폼': platform, '상담 시간': consult}


def _m(leads, registered=(), **kw):
    return [x['리드 No'] for x in match_visited_leads(leads, set(registered), today=TODAY, **kw)]


def test_phone_and_address_match():
    leads = [_lead('L-1')]
    assert _m(leads, phone='01012345678') == ['L-1']
    # 층·상호가 달라도 도로명+번지가 같으면 같은 현장
    assert _m(leads, address='부천시 오정구 석천로 397 부천테크노파크 1층') == ['L-1']
    r = match_visited_leads(leads, set(), phone='010-1234-5678', address='석천로 397', today=TODAY)
    assert r[0]['_match'] == '연락처·주소' and r[0]['_visit_date'] == '2026-11-04'


def test_excludes_registered_confirmed_unvisited_old():
    leads = [
        _lead('L-1'),                                    # 이미 프로젝트에 연결
        _lead('L-2', status='공사 확정'),
        _lead('L-3', status='유선 상담'),
        _lead('L-4', status='방문 예약', visit='2026-11-20'),   # 아직 안 간 예약
        _lead('L-5', visit='2026-03-01'),                 # 180일 넘음
        _lead('ETC-abc'),                                 # 리드 번호가 아님(불러오기 대상 아님)
        _lead('L-6', status='방문 예약', visit='2026-11-02'),   # 지난 예약 = 방문한 것
        _lead('L-7', status='견적 제출', visit=''),        # 날짜 없으면 상담 시간
    ]
    assert _m(leads, registered={'L-1'}, phone='010-1234-5678') == ['L-6', 'L-7']


def test_partner_phone_not_matched():
    # 거래처 담당자 번호는 방문이 수십 건 — 연락처로는 안 맞추고 주소로만
    leads = [_lead('L-1', platform='거래처'), _lead('L-2', addr='강남구 테헤란로 1')]
    assert _m(leads, phone='010-1234-5678') == ['L-2']                     # 리드 쪽이 거래처
    assert _m(leads, phone='010-1234-5678', inflow='거래처') == []          # 입력한 유입이 거래처
    assert _m(leads, address='석천로 397', inflow='거래처') == ['L-1']


def test_order_and_limit():
    leads = [_lead(f'L-{i}', visit=f'2026-10-{10 + i:02d}', addr='다른로 1') for i in range(5)]
    leads.append(_lead('L-both', visit='2026-09-01'))
    got = _m(leads, phone='010-1234-5678', address='석천로 397')
    # 둘 다 맞는 것 먼저, 그다음 최근 방문 순, 3건까지
    assert got == ['L-both', 'L-4', 'L-3']


def test_nothing_to_compare():
    assert _m([_lead('L-1')]) == []
    assert _m([_lead('L-1')], phone='1234', address='') == []
    assert _m([_lead('L-1')], address='주소 모름') == []
