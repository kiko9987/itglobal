# -*- coding: utf-8 -*-
"""공사 금액 요청 ✅ 반영 무결성 가드 테스트 (2026-08-07).

경영지원이 PM 사이트/구글 시트에 금액을 반영하지 않고 ✅ 만 누르면, 요청자에게
'완료' DM 이 잘못 나가고 카드도 완료로 바뀜. _amount_request_applied 로 요청 필드가
실제로 바뀌었는지 확인 → 안 바뀌었으면 호출부가 완료 처리 skip + 경고.

판정 규칙: 요청(updates) 필드 중 하나라도 요청 전(before)과 달라졌으면 반영됨(True).
실제값이 요청값과 달라도(경영지원 재량 조정) '변경'이면 반영 인정. proj 없으면
None(확인 불가 → 차단 안 함). 전부 그대로면 False(미반영 → 경고).
"""
import sys
sys.path.insert(0, '.')

from dashboard.blueprints.slack_bot import _amount_request_applied, _amt_int


# ── _amt_int 파싱 ──
def test_amt_int_parses():
    assert _amt_int('6,800,000') == 6800000
    assert _amt_int(7000000.0) == 7000000
    assert _amt_int('') == 0
    assert _amt_int(None) == 0
    assert _amt_int('-') == 0


# ── 미반영 감지 (핵심: ✅만 누른 케이스) ──
def test_not_applied_amount_unchanged():
    before = {'총액 1': 6800000, '부가세': True}
    updates = {'총액 1': 7000000}
    proj = {'총액 1': 6800000.0, '부가세': True}  # PM 반영 안 됨 → 그대로
    assert _amount_request_applied(proj, before, updates) is False


def test_not_applied_vat_unchanged():
    before = {'총액 1': 5000000, '부가세': True}
    updates = {'부가세': False}
    proj = {'총액 1': 5000000, '부가세': True}  # VAT 그대로
    assert _amount_request_applied(proj, before, updates) is False


# ── 반영됨 감지 ──
def test_applied_amount_changed():
    before = {'총액 1': 6800000, '부가세': True}
    updates = {'총액 1': 7000000}
    proj = {'총액 1': 7000000.0, '부가세': True}  # PM 반영됨
    assert _amount_request_applied(proj, before, updates) is True


def test_applied_even_if_value_differs_from_request():
    """경영지원이 요청값(7M)과 다른 값(7.2M)으로 조정 반영해도 '변경'이면 인정."""
    before = {'총액 1': 6800000, '부가세': True}
    updates = {'총액 1': 7000000}
    proj = {'총액 1': 7200000, '부가세': True}
    assert _amount_request_applied(proj, before, updates) is True


def test_applied_vat_changed():
    before = {'총액 1': 5000000, '부가세': True}
    updates = {'부가세': False}
    proj = {'총액 1': 5000000, '부가세': False}
    assert _amount_request_applied(proj, before, updates) is True


def test_applied_when_either_field_changed():
    """총액·부가세 동시 요청 중 하나만 바뀌어도 반영 인정(부분 반영도 진행 중으로 봄)."""
    before = {'총액 1': 5000000, '부가세': True}
    updates = {'총액 1': 6000000, '부가세': False}
    proj = {'총액 1': 6000000, '부가세': True}  # 총액만 바뀜
    assert _amount_request_applied(proj, before, updates) is True


# ── 확인 불가 (proj 없음) → None (차단 안 함) ──
def test_none_when_project_missing():
    assert _amount_request_applied(None, {'총액 1': 1}, {'총액 1': 2}) is None
