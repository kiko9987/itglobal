# -*- coding: utf-8 -*-
"""슬랙 모달 현장 목록은 시트 로드를 기다리지 않는다 (2026-11-17 계약서 작성 요청).

캐시 TTL 5분 < 주기 갱신 10분이라 캐시가 비는 구간이 있고, 그때 들어온 요청이 시트 전체
로드(8~15초)를 떠안았다. 슬랙은 3초만 기다려 "연결하는 데 문제가 발생했습니다"가 떴다.
"""
import sys
import time

sys.path.insert(0, '.')

import pandas as pd
import pytest

from dashboard.services import project_service as ps


@pytest.fixture
def no_cache(monkeypatch):
    """캐시는 늘 비어 있다. 시트 로드로 넘어가면 calls 에 흔적이 남는다
    (load_data 는 @handle_error 가 예외를 삼키므로 raises 가 아니라 흔적으로 본다)."""
    calls = []
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'test-sheet')
    monkeypatch.setattr(ps, 'smart_get', lambda *a, **kw: None)

    def _slow_load(*a, **kw):
        calls.append(time.time())
        raise RuntimeError('시트 전체 로드 — 8~15초 걸리는 자리')

    monkeypatch.setattr(ps, 'get_sheets_manager', _slow_load)
    monkeypatch.setattr(ps, '_last_df', None)
    monkeypatch.setattr(ps, '_last_df_at', 0.0)
    return calls


DF = pd.DataFrame([{'프로젝트 코드': 'G4049-JW', '사업자명': '강원포장'}])


def test_stale_ok_answers_from_last_snapshot(no_cache, monkeypatch):
    monkeypatch.setattr(ps, '_last_df', DF)
    monkeypatch.setattr(ps, '_last_df_at', time.time() - 30)
    out = ps.load_data(stale_ok=True)
    assert out is DF and not no_cache          # 로드를 안 탔다


def test_stale_too_old_falls_back_to_load(no_cache, monkeypatch):
    """너무 묵은 자료는 쓰지 않는다 — 주기 갱신이 멈춘 채 하루를 넘기지 않게."""
    monkeypatch.setattr(ps, '_last_df', DF)
    monkeypatch.setattr(ps, '_last_df_at', time.time() - ps.STALE_MAX_AGE - 10)
    ps.load_data(stale_ok=True)
    assert no_cache, '너무 묵었으면 로드로 가야 한다'


def test_without_stale_ok_still_loads(no_cache):
    """PM·시트 쓰기 같은 다른 경로는 그대로 — 기다려서라도 최신을 받는다."""
    ps.load_data()
    assert no_cache


def test_cache_hit_refreshes_snapshot(monkeypatch):
    """캐시로 응답할 때마다 직전 자료도 같이 갱신해 둔다."""
    monkeypatch.setattr(ps, '_last_df', None)
    monkeypatch.setattr(ps, '_last_df_at', 0.0)
    monkeypatch.setattr(ps, 'smart_get', lambda *a, **kw: DF)
    assert ps.load_data() is DF
    assert ps._last_df is DF and ps._last_df_at > 0


def test_empty_frame_is_not_remembered(monkeypatch):
    """빈 결과를 스냅샷으로 들고 있으면 목록이 영영 빈칸이 된다."""
    monkeypatch.setattr(ps, '_last_df', None)
    monkeypatch.setattr(ps, '_last_df_at', 0.0)
    ps._remember_df(pd.DataFrame())
    ps._remember_df(None)
    assert ps._last_df is None
