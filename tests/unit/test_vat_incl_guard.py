# -*- coding: utf-8 -*-
"""계산서 요청 'VAT 포함' 정리 (2026-11-09 SB 요청).

7/24~11/9 세금계산서 요청 342건 중 포함 22건: 고객이 부가세 포함 둥근 금액으로 낸 돈(정상) 14 ·
1.1로 딱 나누어떨어짐(별도로 보낼 수 있었음) 7 · 공급가를 포함으로(덜 발행) 1(G3955-MS).
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb


@pytest.mark.parametrize('amt, out', [
    ('4290000', (4_290_000, 3_900_000)),      # R4250-SJ — 별도 3,900,000 으로 바꿈
    ('4,290,000', (4_290_000, 3_900_000)),
    ('5000000', None),                        # 고객이 낸 둥근 금액 — 공급가가 안 떨어짐(포함 그대로)
    ('0', None), ('', None),
])
def test_vat_incl_to_supply(amt, out):
    assert sb._vat_incl_to_supply(amt, 'incl') == out


def test_vat_incl_to_supply_ignores_separate():
    assert sb._vat_incl_to_supply('4290000', 'sep') is None


@pytest.fixture
def proj(monkeypatch):
    import dashboard.services.project_service as ps
    state = {'recs': [], 'iss': {'issued_gross': 0}}
    monkeypatch.setattr(ps, 'get_project_records', lambda: state['recs'])
    monkeypatch.setattr(sb, '_project_issued_invoice', lambda code, stage=None: state['iss'])
    return state


def test_supply_sent_as_included_warns(proj):
    # G3955-MS: 총액1(공급가) 250,000원을 VAT 포함으로 → 227,273 + 부가세로 덜 발행될 뻔
    proj['recs'] = [{'프로젝트 코드': 'G3955-MS', '총액 1': 250_000, '총액 2': 275_000}]
    w = sb._invoice_amount_warning('G3955-MS', '250000', 'incl')
    assert '250,000원은 공급가' in w and "'VAT 별도'로 바꾸세요" in w


def test_remaining_supply_sent_as_included_warns(proj):
    # 이미 일부 발행 → 창에 남은 공급가가 채워진 걸 포함으로 바꾼 경우
    proj['recs'] = [{'프로젝트 코드': 'R1-X', '총액 1': 9_100_000, '총액 2': 10_010_000}]
    proj['iss'] = {'issued_gross': 5_500_000, 'remaining_supply': 4_100_000}
    assert '4,100,000원은 공급가' in sb._invoice_amount_warning('R1-X', '4100000', 'incl')


@pytest.mark.parametrize('code, t1, t2, amt', [
    ('R4223-SJ', 9_100_000, 10_010_000, '3000000'),   # 고객이 포함 금액으로 낸 돈 — 정상
    ('R3642-SH', 13_909_090, 15_300_000, '15300000'),  # 총액2 그대로 — 정상
    ('V0-X', 250_000, 250_000, '250000'),             # 부가세 없는 공사 — 볼 것 없음
])
def test_normal_included_passes(proj, code, t1, t2, amt):
    proj['recs'] = [{'프로젝트 코드': code, '총액 1': t1, '총액 2': t2}]
    assert sb._invoice_amount_warning(code, amt, 'incl') == ''


def test_separate_supply_not_warned_by_new_rule(proj):
    proj['recs'] = [{'프로젝트 코드': 'G3955-MS', '총액 1': 250_000, '총액 2': 275_000}]
    assert sb._invoice_amount_warning('G3955-MS', '250000', 'sep') == ''
