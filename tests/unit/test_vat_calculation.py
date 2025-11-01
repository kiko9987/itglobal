#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
부가세 계산 테스트
Google Sheets ROUND 수식과 Python 계산 일치 검증
"""
import sys
sys.path.insert(0, '.')

import pytest
from decimal import Decimal
from dashboard.services.project_service import _calculate_total2


def test_vat_calculation_google_sheets_compatibility():
    """
    부가세 계산이 Google Sheets FLOOR + 끝자리 1/9원 보정과 동일한지 검증

    Google Sheets 수식:
    =IF(R:R=TRUE, Q:Q+FLOOR(Q:Q*0.1,1) + IF(MOD(Q:Q+FLOOR(Q:Q*0.1,1), 10)=1, -1, IF(MOD(Q:Q+FLOOR(Q:Q*0.1,1), 10)=9, 1, 0)), Q:Q)

    FLOOR/ROUNDDOWN는 절사 (0.5 이상도 내림)
    끝자리 1원 → -1원 보정, 9원 → +1원 보정
    """
    # Case 1: 15005 × 0.1 = 1500.5 → 1500 (절사), 16505 끝자리 5 → 보정 없음
    total1 = 15005
    total2 = _calculate_total2(total1, vat_flag=True)
    assert total2 == 16505, f"Expected 16505 (15005 + 1500, 끝자리 5 보정 없음), got {total2}"

    # Case 2: 15015 × 0.1 = 1501.5 → 1501 (절사), 16516 끝자리 6 → 보정 없음
    total1 = 15015
    total2 = _calculate_total2(total1, vat_flag=True)
    assert total2 == 16516, f"Expected 16516 (15015 + 1501, 끝자리 6 보정 없음), got {total2}"

    # Case 3: 15025 × 0.1 = 1502.5 → 1502 (절사), 16527 끝자리 7 → 보정 없음
    total1 = 15025
    total2 = _calculate_total2(total1, vat_flag=True)
    assert total2 == 16527, f"Expected 16527 (15025 + 1502, 끝자리 7 보정 없음), got {total2}"


def test_vat_calculation_edge_cases():
    """부가세 계산 엣지 케이스"""

    # Case 1: 0원
    total1 = 0
    total2 = _calculate_total2(total1, vat_flag=True)
    assert total2 == 0, f"Expected 0, got {total2}"

    # Case 2: 부가세 미포함
    total1 = 15005
    total2 = _calculate_total2(total1, vat_flag=False)
    assert total2 == 15005, f"Expected 15005, got {total2}"

    # Case 3: 1원 (최소값) - 끝자리 1 보정으로 0이 됨
    total1 = 1
    total2 = _calculate_total2(total1, vat_flag=True)
    assert total2 == 0, f"Expected 0 (1 + 0 - 1원 보정), got {total2}"

    # Case 4: 5원 (0.5 절사) - 끝자리 5 보정 없음
    total1 = 5
    total2 = _calculate_total2(total1, vat_flag=True)
    assert total2 == 5, f"Expected 5 (5 + 0, 끝자리 5 보정 없음), got {total2}"


def test_vat_vs_python_round():
    """
    Python round() (은행가 반올림) vs FLOOR (절사) + 끝자리 1/9원 보정 차이 검증
    """
    from decimal import ROUND_DOWN

    # 15005 × 0.1 = 1500.5
    total1 = 15005

    # Python round() (은행가 반올림): 1500 (짝수로)
    python_vat = round(total1 * 0.1)
    assert python_vat == 1500, f"Python round(): {python_vat}"

    # FLOOR (절사): 1500 (0.5도 내림)
    floor_vat = int((Decimal(str(total1)) * Decimal('0.1')).quantize(
        Decimal('1'),
        rounding=ROUND_DOWN
    ))
    assert floor_vat == 1500, f"FLOOR: {floor_vat}"

    # 함수 결과는 FLOOR + 끝자리 보정과 일치해야 함
    total2 = _calculate_total2(total1, vat_flag=True)
    expected = total1 + 1500  # FLOOR 결과 (16505, 끝자리 5 → 보정 없음)
    assert total2 == expected, f"Expected {expected} (FLOOR + 끝자리 보정), got {total2}"


def test_vat_calculation_vs_google_sheets_examples():
    """실제 Google Sheets 예시 값들과 비교 (끝자리 1/9원 보정 포함)"""

    test_cases = [
        # (총액1, 예상 총액2, 설명)
        (10005, 11005, "10005 × 0.1 = 1000.5 → 1000 (절사), 11005 끝자리 5 → 보정 없음"),
        (99995, 109994, "99995 × 0.1 = 9999.5 → 9999 (절사), 109994 끝자리 4 → 보정 없음"),
        (100000, 110000, "100000 × 0.1 = 10000, 110000 끝자리 0 → 보정 없음"),
        (123455, 135800, "123455 × 0.1 = 12345.5 → 12345 (절사), 135800 끝자리 0 → 보정 없음"),
        (500000, 550000, "500000 × 0.1 = 50000, 550000 끝자리 0 → 보정 없음"),
        # 끝자리 1 보정 테스트 (총액2가 -1원 조정됨)
        (10001, 11000, "10001 × 0.1 = 1000.1 → 1000 (절사), 11001 끝자리 1 → -1원 보정 = 11000"),
        # 끝자리 9 보정 테스트 (총액2가 +1원 조정됨)
        (14090909, 15500000, "14090909 × 0.1 = 1409090.9 → 1409090 (절사), 15499999 끝자리 9 → +1원 보정 = 15500000"),
    ]

    for total1, expected_total2, description in test_cases:
        total2 = _calculate_total2(total1, vat_flag=True)

        assert total2 == expected_total2, \
            f"{description}\n총액1={total1}: 총액2={total2} (예상: {expected_total2})"


if __name__ == '__main__':
    pytest.main([__file__, '-v'])
