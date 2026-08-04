# -*- coding: utf-8 -*-
"""POI 결합-후보(G1) 지점명 이중 부착 방지 회귀 테스트 (2026-08-04 L-03451·L-03486).

_enrich_with_poi 의 결합-후보 경로는 공백형 상호('남양가 양꼬치')를 카카오 정식
place_name('남양가양꼬치 마곡점')으로 replace 해 지점명을 보강함. 그런데 raw 끝에
이미 지점명이 있으면('힙춘향 마라 계양점') replace 시 뒤 지점명이 남아 중복됨
('힙춘향마라 계양점 계양점'). POI 정식명이 이미 verified 에 있으면(공백 무시) skip.
_search_poi 를 monkeypatch (네트워크 없이 매치 로직 검증).
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services import address_resolver as ar


def test_no_duplicate_branch_when_already_present(monkeypatch):
    """raw 끝에 지점명 있으면 POI replace 로 이중부착 금지 (L-03486 힙춘향마라 계양점)."""
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('힙춘향마라 계양점', '인천 계양구 봉오대로 712')])
    r = ar._enrich_with_poi('인천 계양구 봉오대로 712 1층 힙춘향 마라 계양점',
                            '인천 계양구 봉오대로 712 1층 힙춘향 마라 계양점')
    assert '계양점 계양점' not in r
    assert r.count('계양점') == 1


def test_no_duplicate_branch_full_name_present(monkeypatch):
    """정식명 전체가 이미 있으면 skip (L-03451 케이플라워마트 대화점 대화점)."""
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('케이플라워마트 대화점', '고양 일산서구 대화로 362')])
    r = ar._enrich_with_poi('고양 일산서구 대화로 362 케이플라워마트 대화점',
                            '고양 일산서구 대화로 362 케이플라워마트 대화점')
    assert '대화점 대화점' not in r


def test_branch_still_appended_when_absent(monkeypatch):
    """지점명이 아직 없으면 정상 보강 (회귀 방지 — L-03475 남양가양꼬치 마곡점)."""
    monkeypatch.setattr(ar, '_search_poi',
                        lambda q: [('남양가양꼬치 마곡점', '서울 강서구 마곡중앙로 161-8')])
    r = ar._enrich_with_poi('강서구 마곡중앙로 161-8 그랑트윈타워 A동 남양가 양꼬치',
                            '강서구 마곡중앙로 161-8 그랑트윈타워 A동 남양가 양꼬치')
    assert '남양가양꼬치 마곡점' in r


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
