# -*- coding: utf-8 -*-
"""invariant_checks 순수 검사·헬퍼 단위 테스트 (redis/네트워크 불요)."""
from datetime import datetime

from dashboard.services import invariant_checks as ic


# ── 헬퍼 ──────────────────────────────────────────────
def test_parse_amount():
    assert ic._parse_amount('1,000,000원') == 1000000.0
    assert ic._parse_amount('₩ 50,000') == 50000.0
    assert ic._parse_amount('') == 'empty'
    assert ic._parse_amount('-') == 'empty'
    assert ic._parse_amount(None) == 'empty'
    assert ic._parse_amount('abc') is None
    assert ic._parse_amount('3827') == 3827.0
    assert ic._parse_amount('-5000') == -5000.0
    # 빈 숫자셀(pandas NaN)·불리언은 'empty'(금액 아님) — 오탐 방지
    assert ic._parse_amount('nan') == 'empty'
    assert ic._parse_amount(float('nan')) == 'empty'
    assert ic._parse_amount('None') == 'empty'


def test_recent_date():
    now = datetime(2026, 9, 11)
    assert ic._recent_date('26/09/10', now, 45) is True
    assert ic._recent_date('2026.09.01', now, 45) is True
    assert ic._recent_date('26/06/01', now, 45) is False   # 오래됨
    assert ic._recent_date('26/12/25', now, 45) is False   # 미래
    assert ic._recent_date('메모뿐', now, 45) is False      # 파싱 불가
    assert ic._recent_date('', now, 45) is False
    assert ic._recent_date('09/10', now, 45) is True        # 연도 없음 → 올해 가정


# ── check_phantom_codes ──────────────────────────────
def test_phantom_codes():
    ctx = {'card_codes': {'G0001-AA', 'G0002-BB', ''}, 'valid_codes': {'G0001-AA'}}
    vios = ic.check_phantom_codes(ctx)
    keys = {v['title'] for v in vios}
    assert keys == {'G0002-BB'}          # 유효코드·빈코드 제외
    assert vios[0]['check'] == 'phantom_code'


# ── check_orphan_settlement ──────────────────────────
def _octx(records, payment_card_codes, now=None):
    return {
        'records': records,
        'payment_card_codes': set(payment_card_codes),
        'now': now or datetime(2026, 9, 11),
        'recent_days': 45,
    }


def test_orphan_settlement_flags_missing_card():
    r = {'프로젝트 코드': 'G4026-MS', '미수금': '0', '계약금': '1,000,000',
         '중도금': '0', '잔금': '2,000,000', '수금 날짜': '26/09/10'}
    vios = ic.check_orphan_settlement(_octx([r], payment_card_codes=set()))
    assert len(vios) == 1 and vios[0]['title'] == 'G4026-MS'


def test_orphan_settlement_ok_when_card_exists():
    r = {'프로젝트 코드': 'G4026-MS', '미수금': '0', '계약금': '1,000,000',
         '잔금': '2,000,000', '수금 날짜': '26/09/10'}
    vios = ic.check_orphan_settlement(_octx([r], payment_card_codes={'G4026-MS'}))
    assert vios == []


def test_orphan_settlement_skips_outstanding_old_and_unpaid():
    base = {'프로젝트 코드': 'G1-AA', '수금 날짜': '26/09/10'}
    outstanding = {**base, '미수금': '500,000', '계약금': '1,000,000'}   # 미수금 남음
    unpaid = {**base, '미수금': '0', '계약금': '0', '중도금': '0', '잔금': '0'}  # 입금 0
    old = {'프로젝트 코드': 'G2-BB', '미수금': '0', '잔금': '1,000,000',
           '수금 날짜': '26/01/01'}                                        # 오래됨
    vios = ic.check_orphan_settlement(_octx([outstanding, unpaid, old], set()))
    assert vios == []


# ── check_amount_anomaly ─────────────────────────────
def test_amount_anomaly_detects_bad_values():
    records = [
        {'프로젝트 코드': 'G1-AA', '잔금': 'abc'},              # 파싱 불가
        {'프로젝트 코드': 'G2-BB', '계약금': '-5000'},          # 음수
        {'프로젝트 코드': 'G3-CC', '총액 1': '999999999999999'},  # 범위 초과
        {'프로젝트 코드': 'G4-DD', '잔금': '1,000,000', '총액 1': '', '부가세': '-'},  # 정상/빈값
    ]
    vios = ic.check_amount_anomaly({'records': records})
    by_code = {v['title'] for v in vios}
    assert by_code == {'G1-AA', 'G2-BB', 'G3-CC'}   # G4-DD 는 위반 없음
    assert any('파싱 불가' in v['detail'] for v in vios)
    assert any('음수' in v['detail'] for v in vios)
    assert any('범위 초과' in v['detail'] for v in vios)


def test_amount_anomaly_skips_records_without_code():
    vios = ic.check_amount_anomaly({'records': [{'잔금': 'abc'}]})
    assert vios == []
