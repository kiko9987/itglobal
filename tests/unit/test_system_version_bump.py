# -*- coding: utf-8 -*-
"""시스템이 올린 버전은 저장한 사람 화면을 막지 않는다 (2026-11-20 R4297-YG 계기).

박용구 대표가 사업자명을 저장(v1) → 거래처명 자동 채우기가 '기타'→'한국파이프콘베어'로 쓰며 v2 →
대표 화면은 v1 이라 다음 저장 두 번이 '다른 사용자가 먼저 수정했습니다'로 거절됐다.
"""
import inspect
import sys

sys.path.insert(0, '.')

import pytest

from dashboard.services import system_version_bump as sb


class _R:
    def __init__(self):
        self.d = {}

    def set(self, k, v, ex=None):
        self.d[k] = v

    def get(self, k):
        return self.d.get(k)


@pytest.fixture
def r(monkeypatch):
    rr = _R()
    monkeypatch.setattr(sb, '_rc', lambda: rr)
    return rr


def test_untouched_field_passes_and_is_dropped(r):
    sb.record('R4297-YG', 2, {'거래처명': ('기타', '한국파이프콘베어')})
    data = {'거래처명': '기타', '견적서 및 계약서 폴더 경로': '1Rc5', '_version': '1'}
    assert sb.rebase('R4297-YG', '1', '2', data) is True
    assert '거래처명' not in data                       # 옛 값으로 되돌리지 않게 뺀다
    assert data['견적서 및 계약서 폴더 경로'] == '1Rc5'   # 사람이 고친 칸은 그대로


def test_already_new_value_or_absent_passes(r):
    sb.record('R1-AA', 5, {'거래처명': ('기타', '강원포장')})
    assert sb.rebase('R1-AA', 4, 5, {'거래처명': '강원포장'}) is True
    assert sb.rebase('R1-AA', 4, 5, {'시공자': '아이티'}) is True
    assert sb.rebase('R1-AA', 4, 5, {'거래처명': '-'}) is False   # 화면이 '-'(모름)로 바꿈 = 다른 값 → 충돌


def test_field_changed_differently_is_real_conflict(r):
    sb.record('R1-AA', 5, {'거래처명': ('기타', '강원포장')})
    data = {'거래처명': '대성크레인'}
    assert sb.rebase('R1-AA', 4, 5, data) is False
    assert data == {'거래처명': '대성크레인'}            # 막을 땐 아무것도 안 건드린다


def test_human_version_in_between_is_real_conflict(r):
    sb.record('R1-AA', 5, {'거래처명': ('기타', '강원포장')})   # v6 은 사람이 고침(기록 없음)
    assert sb.rebase('R1-AA', 4, 6, {'시공자': '아이티'}) is False
    assert sb.rebase('R1-AA', 5, 5, {}) is False                  # 같은 버전은 여기 올 일 없음
    assert sb.rebase('R1-AA', 'x', 5, {}) is False


def test_dash_and_blank_are_same(r):
    sb.record('R2-BB', 3, {'거래처명': ('-', '강원포장')})
    assert sb.rebase('R2-BB', 2, 3, {'거래처명': ''}) is True


def test_wired_into_lock_and_partner_backfill():
    from dashboard.blueprints import projects as pj
    from dashboard.services import partner_names as pn
    assert 'data=None' in str(inspect.signature(pj._check_optimistic_lock_update))
    assert '_sys_rebase(' in inspect.getsource(pj._check_optimistic_lock_update)
    assert 'data=data' in inspect.getsource(pj.update_project)
    assert '_sysbump(' in inspect.getsource(pn)
