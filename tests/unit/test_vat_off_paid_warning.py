# -*- coding: utf-8 -*-
"""'VAT 별도 → 없음' 금액 수정 요청 시 이미 입금된 프로젝트 경고 (2026-10-07 경영지원 요청).

계산서 발행 예정이던 건을 고객이 입금 후 현금거래로 바꾸는 경우 — 요청 카드에 받은 금액·
계산서 상태를 보여 반영 전에 확인. 실례 R4163-TH: 계약금 200,000원 입금 후 VAT 없음 요청.
"""
import sys
sys.path.insert(0, '.')

from dashboard.blueprints.slack_bot import _vat_off_paid_warning as w

R4163 = {'프로젝트 코드': 'R4163-TH', '부가세': True, '총액 1': 2500000,
         '계약금': 200000, '중도금': 0, '잔금': 0, '계약금 계산서': 'N입금'}


def test_warns_when_paid_and_vat_turned_off():
    out = w(R4163, {'부가세': False})
    assert '이미 입금된 프로젝트' in out
    assert '계약금 200,000원(계산서 N입금)' in out


def test_lists_every_paid_stage_with_invoice_status():
    p = dict(R4163, 중도금='1,000,000', **{'중도금 계산서': '발행'})
    out = w(p, {'부가세': 'FALSE'})
    assert '계약금 200,000원' in out and '중도금 1,000,000원(계산서 발행)' in out


def test_no_warning_cases():
    assert w(dict(R4163, 계약금=0), {'부가세': False}) == ''          # 입금 없음
    assert w(dict(R4163, 부가세=False), {'부가세': False}) == ''      # 원래 VAT 없음
    assert w(R4163, {'부가세': True}) == ''                           # 별도 유지/켜기
    assert w(R4163, {'총액 1': 2400000}) == ''                        # 금액만 변경
    assert w(dict(R4163, 부가세='TRUE'), {'부가세': 'FALSE'}) != ''   # 문자열 값도 인식
