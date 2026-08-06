# -*- coding: utf-8 -*-
"""캔버스2 온라인 당번 파싱 회귀 테스트 (2026-08-03 SB 등기발송 오파싱).

온라인 섹션에 '/' 필드가 있는 태스크 라인(예: 'JK 점심때 (SB) … / 등기발송 2건')이
있으면, 파서가 그 안의 모든 이니셜(SB)을 온라인 당번으로 오수집 → commit 시 SB 가
잘못된 온라인 당번 DM 을 받았음. '/' 태스크 라인은 이니셜 수집 skip.
_load_initial_maps 를 monkeypatch (네트워크 없이 파서 로직만 검증).
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import visit_assignment_sync as vas


def test_online_duty_skips_task_line(monkeypatch):
    """온라인 섹션의 '/' 태스크 라인(등기발송)은 online_duty 에서 제외 → JK 만."""
    monkeypatch.setattr(vas, '_load_initial_maps',
                        lambda: ({'JK': '강정권', 'SB': '조성복'}, {}))
    html = (
        '<p>온라인</p><p>JK</p>'
        '<p>JK 점심때 (SB) 8월 4일 / - / 근처 우체국 / 등기발송 2건</p>'
        '<p>휴무</p>'
    )
    r = vas.parse_assignment_canvas_full(html)
    assert r['online_duty'] == ['JK']


def test_online_duty_pure_initials_and_shifts(monkeypatch):
    """순수 이니셜/시간대 라인은 정상 파싱 (회귀 방지)."""
    monkeypatch.setattr(vas, '_load_initial_maps',
                        lambda: ({'SD': '서동', 'MS': '문성'}, {}))
    r = vas.parse_assignment_canvas_full(
        '<p>온라인</p><p>SD(오전)+MS(오후)</p><p>휴무</p>')
    assert set(r['online_duty']) == {'SD', 'MS'}
    assert r['online_duty_shifts'] == {'SD': '오전', 'MS': '오후'}


def test_online_shared_marker(monkeypatch):
    """'공용업무 전인원 제외 없음' = online_shared True, 당번은 비움 (2026-08-04)."""
    monkeypatch.setattr(vas, '_load_initial_maps',
                        lambda: ({'JK': '강정권'}, {}))
    r = vas.parse_assignment_canvas_full(
        '<p>온라인</p><p>공용업무 전인원 제외 없음</p><p>휴무</p>')
    assert r['online_shared'] is True
    assert r['online_duty'] == []  # 텍스트만 → 특정 당번 없음


def test_online_shared_variants(monkeypatch):
    """공용/전인원/제외없음 개별 키워드도 각각 감지."""
    monkeypatch.setattr(vas, '_load_initial_maps',
                        lambda: ({'JK': '강정권'}, {}))
    for txt in ('공용', '전 인원', '제외없음', '전인원 공용'):
        r = vas.parse_assignment_canvas_full(
            f'<p>온라인</p><p>{txt}</p><p>휴무</p>')
        assert r['online_shared'] is True, txt


def test_online_shared_false_when_duty_assigned(monkeypatch):
    """특정 당번(JK)만 있으면 online_shared False (공용 아님)."""
    monkeypatch.setattr(vas, '_load_initial_maps',
                        lambda: ({'JK': '강정권'}, {}))
    r = vas.parse_assignment_canvas_full(
        '<p>온라인</p><p>JK</p><p>휴무</p>')
    assert r['online_shared'] is False
    assert r['online_duty'] == ['JK']


def test_online_prefix_noise_no_swallow(monkeypatch):
    """'광고OFF (JSH(오전)+SD(오후))' — 앞 대문자 노이즈(OFF)가 뒤 이니셜을 삼키지 않음.

    OFF 가 [A-Z]{2,4} 로 잡혀 'JSH(오전' 을 시프트로 통째 삼켜 JSH 가 누락되던 버그
    (2026-08-06 08-07 온라인 당번). 시프트 charset 에 대문자·괄호 제외로 fix.
    """
    monkeypatch.setattr(vas, '_load_initial_maps',
                        lambda: ({'JSH': '정성훈', 'SD': '서동'}, {}))
    r = vas.parse_assignment_canvas_full(
        '<p>온라인</p><p>광고OFF (JSH(오전)+SD(오후))</p><p>휴무</p>')
    assert set(r['online_duty']) == {'JSH', 'SD'}
    assert r['online_duty_shifts'] == {'JSH': '오전', 'SD': '오후'}


def test_online_shared_with_initials_coexist(monkeypatch):
    """공용 표기 + 이니셜 공존 시: shared True 이면서 당번도 수집."""
    monkeypatch.setattr(vas, '_load_initial_maps',
                        lambda: ({'JK': '강정권'}, {}))
    r = vas.parse_assignment_canvas_full(
        '<p>온라인</p><p>공용 (JK 오전 지원)</p><p>휴무</p>')
    assert r['online_shared'] is True
    assert 'JK' in r['online_duty']


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
