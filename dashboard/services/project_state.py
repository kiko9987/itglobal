# -*- coding: utf-8 -*-
"""공사 상태 판정 한 곳 (2026-11-16 상태 머신 점검 묶음 E).

공사현황 시트엔 상태 열이 없고 `수금 관련 특이사항`에 '공사 취소'가 들어 있는지로 취소를 본다. 이 판정이
10곳 넘게 흩어져 있었고, 리컨사일러만 부분 문자열('공사 취소' in)을 써서 '공사취소'(붙여 씀)를 다르게 봤다.
판정은 모두 이 함수로 한다.
"""
import re

CANCEL_RE = re.compile(r'공사\s*취소')


def spec_cancelled(spec) -> bool:
    """`수금 관련 특이사항` 문구가 공사 취소 표시인지."""
    return bool(CANCEL_RE.search(str(spec or '')))


def is_cancelled(rec) -> bool:
    """공사 레코드(dict)가 취소 상태인지."""
    return spec_cancelled((rec or {}).get('수금 관련 특이사항'))
