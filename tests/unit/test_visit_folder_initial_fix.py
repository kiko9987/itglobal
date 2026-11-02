# -*- coding: utf-8 -*-
"""방문 폴더 이니셜 자동 교정 — (미정) + 온라인 상담자 폴백 이니셜 (2026-11-02 L-04421)."""
import sys
sys.path.insert(0, '.')

import pytest

from dashboard.blueprints import slack_bot as sb


@pytest.fixture
def renamed(monkeypatch):
    calls = []
    import dashboard.utils.google_drive as gd
    monkeypatch.setattr(gd, 'rename_folder', lambda fid, name: calls.append((fid, name)) or True)
    return calls


def _lead(sales='-', online='고광일', plat='메일'):
    return {'플랫폼': plat, '영업 담당자': sales, '온라인 상담자': online}


def _ini(name):
    return sb._to_initial(name)


def test_online_fallback_initial_renamed_after_sales_assigned(renamed):
    """L-04421: 영업 담당자 없을 때 '(KiKO)' 로 만든 폴더 → 박용구 배정 후 사진 업로드 시 교정."""
    cur = f'메일 ({_ini("고광일")}) 영등포구 양평로21길 10 롯데양평빌딩 26.11.02'
    out = sb._maybe_fix_folder_initial('F1', cur, _lead(sales='박용구'))
    assert out == f'메일 ({_ini("박용구")}) 영등포구 양평로21길 10 롯데양평빌딩 26.11.02'
    assert renamed == [('F1', out)]


def test_no_prefix_platform_also_fixed(renamed):
    cur = f'({_ini("고광일")}) 강서구 화곡로 1 26.11.02'
    out = sb._maybe_fix_folder_initial('F1', cur, _lead(sales='조성헌', plat='홈페이지'))
    assert out.startswith(f'({_ini("조성헌")}) ')


def test_still_no_sales_keeps_name(renamed):
    cur = f'메일 ({_ini("고광일")}) 주소 26.11.02'
    assert sb._maybe_fix_folder_initial('F1', cur, _lead()) == cur
    assert renamed == []


def test_manual_other_initial_untouched(renamed):
    """손으로 다른 이니셜로 바꾼 폴더는 그대로."""
    cur = f'메일 ({_ini("박정우")}) 주소 26.11.02'
    assert sb._maybe_fix_folder_initial('F1', cur, _lead(sales='박용구')) == cur
    assert renamed == []


def test_online_is_also_sales_untouched(renamed):
    cur = f'({_ini("박용구")}) 주소 26.11.02'
    assert sb._maybe_fix_folder_initial('F1', cur, _lead(sales='박용구', online='박용구')) == cur
    assert renamed == []


def test_partner_lead_untouched(renamed):
    """거래처류는 원래 온라인 상담자 기준 — 영업 담당자가 있어도 안 바꿈."""
    cur = f'거래처 ({_ini("고광일")}) 주소 26.11.02'
    assert sb._maybe_fix_folder_initial('F1', cur, _lead(sales='박용구', plat='거래처')) == cur
    assert renamed == []


def test_initial_only_matches_first_paren(renamed):
    """주소·상호 쪽 괄호는 보지 않음."""
    cur = f'메일 ({_ini("박정우")}) 주소 ({_ini("고광일")}) 26.11.02'
    assert sb._maybe_fix_folder_initial('F1', cur, _lead(sales='박용구')) == cur


def test_undecided_still_fixed(renamed):
    out = sb._maybe_fix_folder_initial('F1', '메일 (미정) 주소 26.11.02', _lead(sales='박용구'))
    assert out == f'메일 ({_ini("박용구")}) 주소 26.11.02'
