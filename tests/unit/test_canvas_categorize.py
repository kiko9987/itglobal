# -*- coding: utf-8 -*-
"""캔버스1 카테고리 분류 회귀 테스트 (2026-08-05, L-03565 채널톡 drop 사고).

_categorize 가 온라인 플랫폼을 화이트리스트(당근·홈페이지·카카오톡·전화)로만
처리해 채널톡·숨고·큐플레이스·메일 방문이 None → 캔버스1에서 통째 drop 됐음.
catch-all 로 전환 — 거래처/소개·기타(ETC) 외 전부 '온라인 방문'.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.services.visit_canvas_sync import _categorize


def _lead(platform='', lead_no='L-00001'):
    return {'리드 No': lead_no, '플랫폼': platform}


def test_channeltalk_is_online():
    """채널톡 = 온라인 방문 (버그 케이스, L-03565)."""
    assert _categorize(_lead('채널톡')) == '온라인 방문'


def test_other_online_platforms():
    """숨고·큐플레이스·메일도 온라인 방문 (화이트리스트 누락분)."""
    for p in ('숨고', '큐플레이스', '메일'):
        assert _categorize(_lead(p)) == '온라인 방문', p


def test_known_online_platforms_regression():
    """기존 온라인 플랫폼 회귀 방지."""
    for p in ('당근', '홈페이지', '카카오톡', '전화'):
        assert _categorize(_lead(p)) == '온라인 방문', p


def test_partner_and_intro():
    assert _categorize(_lead('거래처')) == '거래처'
    assert _categorize(_lead('소개')) == '거래처'


def test_etc_by_leadno_or_platform():
    assert _categorize(_lead('전화', lead_no='ETC-abc123')) == '기타'
    assert _categorize(_lead('기타')) == '기타'


def test_empty_platform_defaults_online():
    """플랫폼 빈값도 drop 하지 않고 온라인 방문 (유효 방문은 무조건 노출)."""
    assert _categorize(_lead('')) == '온라인 방문'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
