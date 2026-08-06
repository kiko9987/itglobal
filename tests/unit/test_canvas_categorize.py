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


# ── _render_item 온라인 이니셜 오부착 회귀 (L-03565 채널톡 '(KIKO)') ──
from dashboard.services.visit_canvas_sync import _render_item


def _full(platform, lead_no='L-00001'):
    return {'리드 No': lead_no, '플랫폼': platform, '방문 예정일': '2026-08-06',
            '고객 연락처': '010-0000-0000', '방문 주소': '성남 중원구 둔촌대로 560',
            '상담 내용': '천장형 견적', '온라인 상담자': '고광일'}


def test_render_channeltalk_no_initial():
    """채널톡 = 온라인 → 이니셜 prefix 없음 (버그 케이스)."""
    out = _render_item(_full('채널톡'), {'고광일': 'KIKO'})
    assert '(KIKO)' not in out and not out.strip().startswith('(')


def test_render_other_online_no_initial():
    """숨고·큐플레이스·메일도 온라인 → 이니셜 없음."""
    for p in ('숨고', '큐플레이스', '메일', '당근'):
        out = _render_item(_full(p), {'고광일': 'KIKO'})
        assert '(KIKO)' not in out, p


def test_render_partner_keeps_initial():
    """거래처 = 이니셜 prefix 유지 (온라인 상담자 기준)."""
    out = _render_item(_full('거래처'), {'고광일': 'KIKO'})
    assert out.strip().startswith('(')


def test_render_etc_keeps_initial():
    """ETC 리드(기타 취급) = 이니셜 prefix 유지."""
    out = _render_item(_full('전화', lead_no='ETC-abc123'), {'고광일': 'KIKO'})
    assert out.strip().startswith('(')


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
