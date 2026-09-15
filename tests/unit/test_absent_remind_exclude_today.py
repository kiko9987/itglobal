# -*- coding: utf-8 -*-
"""부재중/미완료/견적요청 리마인드 — 오늘 접수분 제외 회귀 테스트 (2026-09-15).

방금 온 문의(당일 접수)는 아직 응대 전이라 '다시 연락' 대상이 아님 → 수집 상한=오늘 제외.
어제(직전 영업일) 접수 미처리는 계속 포함. 실행일 무관하도록 모듈 헬퍼로 날짜 생성.
"""
import sys
sys.path.insert(0, '.')

from datetime import date

from dashboard.services import absent_remind as ar


def _lead(dt: date, status: str, consultant: str = '', sales: str = ''):
    return {
        '리드 No': 'L-x',
        '상담 시간': dt.strftime('%Y.%m.%d. 10:00'),
        '상태': status,
        '온라인 상담자': consultant,
        '영업 담당자': sales,
        '고객명': 't',
        '고객 연락처': '010-0000-0000',
    }


def test_excludes_today_includes_prev_business_day(monkeypatch):
    today = date.today()
    prev = ar._previous_business_day(today)   # 직전 영업일 (cutoff 안, today 미만)

    leads = [
        _lead(today, '상담 대기'),                       # 오늘 미완료 → 제외
        _lead(prev, '상담 대기'),                        # 어제 미완료 → 포함
        _lead(today, '부재중'),                          # 오늘 부재중 → 제외
        _lead(prev, '부재중', consultant='홍길동'),       # 어제 부재중 → 포함
        _lead(today, '견적 요청', consultant='김철수'),   # 오늘 견적요청 → 제외
        _lead(prev, '견적 요청', consultant='김철수'),    # 어제 견적요청 → 포함
    ]
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'get_lead_records', lambda: leads)

    unassigned, retry, quote, callback = ar.collect_absent_leads()

    def _dates(items):
        return [ar._lead_date(l) for l in items]

    # A. 미완료 — 오늘 제외, 어제 포함
    assert today not in _dates(unassigned)
    assert prev in _dates(unassigned)

    # B. 부재중 — 오늘 제외, 어제 포함
    retry_all = [l for v in retry.values() for l in v]
    assert today not in _dates(retry_all)
    assert prev in _dates(retry_all)

    # C. 견적요청 — 오늘 제외, 어제 포함
    quote_all = [l for v in quote.values() for l in v]
    assert today not in _dates(quote_all)
    assert prev in _dates(quote_all)
