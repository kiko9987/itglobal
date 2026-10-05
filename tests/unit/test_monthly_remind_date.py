# -*- coding: utf-8 -*-
"""계산서 마감 월간 리마인드 발송일 — 매월 8일(마감 10일 이틀 전), 주말·공휴일이면 직전 영업일.
2026-10-05 사용자 결정(10일 당일 안내는 발행할 시간이 없음)."""
import sys
from datetime import date

sys.path.insert(0, '.')

import pytest

import dashboard.services.invoice_collection_remind as icr


@pytest.mark.parametrize('ym,expected', [
    ((2026, 10), date(2026, 10, 8)),   # 목 — 그대로
    ((2026, 11), date(2026, 11, 6)),   # 8일 일요일 → 금
    ((2027, 2), date(2027, 2, 5)),     # 8일 설 연휴 → 직전 영업일 금
    ((2027, 5), date(2027, 5, 7)),     # 8일 토요일 → 금
])
def test_monthly_remind_date(ym, expected):
    assert icr.monthly_remind_date(*ym) == expected


def test_send_only_on_remind_day(monkeypatch):
    sent = []
    monkeypatch.setattr(icr, '_post', lambda text, tag: sent.append(tag) or {'ok': True})
    monkeypatch.setattr(icr, '_load_recs', lambda: [])
    r = icr.send_monthly_invoice_remind(today=date(2026, 10, 7))
    assert r['skipped'].startswith('발송일 아님') and sent == []
    icr.send_monthly_invoice_remind(today=date(2026, 10, 8))
    assert sent == ['월간']
