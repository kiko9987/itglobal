# -*- coding: utf-8 -*-
"""배정 주소 fallback — 짧은 상호명 정확 일치 매칭 회귀 테스트 (2026-08-11 ETC-4feb23).

전화 없는 캔버스2 배정 라인은 주소 substring 으로 매칭하는데, 오탐 방지용
`len >= 8` 가드 때문에 '알만 고양점'(6자) 같은 짧은 상호명이 막혀 미매칭됐음.
→ 정확 일치(공백무시, 4자+)는 길이 무관 매칭 허용. 8자 미만 substring 은 계속 차단.
"""
import sys
sys.path.insert(0, '.')

from dashboard.services import visit_assignment_sync as vas


def _cand(lead_no, addr, status='방문 예약'):
    return {'리드 No': lead_no, '방문 주소': addr, '상태': status,
            '상담 시간': '2026.08.10. 17:35', '방문 예정일': '2026-08-11'}


def test_short_poi_exact_match():
    # '알만 고양점' 6자 — 정확 일치로 매칭 (기존 8자 가드에 막혔던 케이스)
    a = {'address': '알만 고양점'}
    r = vas._resolve_lead_for_assignment(a, {}, [_cand('ETC-4feb23', '알만 고양점')])
    assert r and r.get('리드 No') == 'ETC-4feb23'


def test_exact_match_ignores_whitespace():
    a = {'address': '알만고양점'}  # 공백 없음
    r = vas._resolve_lead_for_assignment(a, {}, [_cand('ETC-x', '알만 고양점')])
    assert r and r.get('리드 No') == 'ETC-x'


def test_too_short_no_match():
    # 4자 미만은 정확 일치라도 오탐 위험 → 매칭 안 함
    a = {'address': '1층'}
    assert vas._resolve_lead_for_assignment(a, {}, [_cand('ETC-y', '1층')]) is None


def test_short_substring_still_blocked():
    # 부분 일치는 8자+ 만 — '서울 강남'(5자)이 긴 주소의 부분이어도 매칭 안 함(오탐 방지)
    a = {'address': '서울 강남'}
    cand = [_cand('ETC-z', '서울 강남구 테헤란로 152 강남파이낸스센터')]
    assert vas._resolve_lead_for_assignment(a, {}, cand) is None


def test_long_substring_still_matches():
    # 기존 8자+ substring 매칭은 그대로 동작
    a = {'address': '테헤란로 152 강남파이낸스센터'}
    cand = [_cand('ETC-w', '서울 강남구 테헤란로 152 강남파이낸스센터')]
    r = vas._resolve_lead_for_assignment(a, {}, cand)
    assert r and r.get('리드 No') == 'ETC-w'


# ── 전화 suffix 매칭 — 안심번호·앞자리 탈락 내성 (2026-08-11 ETC-590dbb) ──
def test_phone_suffix_ansim_number():
    # 시트 0507-1384-3577(안심번호) ↔ 캔버스 07-1384-3577 (앞 05 탈락)
    phone_map = {'050713843577': [_cand('ETC-590dbb', '인천 계양구 벌말로 553')]}
    a = {'phone_digits': '0713843577'}
    r = vas._resolve_lead_for_assignment(a, phone_map, [])
    assert r and r.get('리드 No') == 'ETC-590dbb'


def test_phone_suffix_leading_zero():
    # 010 앞0 탈락: 시트 01012345678 ↔ 캔버스 1012345678
    phone_map = {'01012345678': [_cand('L-x', '주소 여덟자이상 테스트')]}
    a = {'phone_digits': '1012345678'}
    r = vas._resolve_lead_for_assignment(a, phone_map, [])
    assert r and r.get('리드 No') == 'L-x'


def test_phone_no_false_suffix_across_prefix():
    # 다른 국번(070 vs 010), 뒤 8자리 같아도 full suffix 아니면 매칭 안 함 (오탐 방지)
    phone_map = {'01012345678': [_cand('L-a', '주소 여덟자이상 테스트')]}
    a = {'phone_digits': '07012345678'}  # 070-1234-5678
    assert vas._resolve_lead_for_assignment(a, phone_map, []) is None
