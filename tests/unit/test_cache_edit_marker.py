# -*- coding: utf-8 -*-
"""PM 편집 부분 갱신이 진행 중이던 옛 전체 로드에 덮이지 않는지 (2026-10-07 R4163-TH).

16:20:02 시작한 시트 로드(23초)가 16:20:10 PM 부분 갱신을 16:20:25 에 덮어써 캐시가 편집 전
값으로 회귀 → 금액 수정 ✅ 확인이 '미반영' 오탐. 원인 ①부분 갱신이 무효화 마커를 안 남김
②마커 TTL 10초 < 로드 시간. 테스트 Redis = db15 (루트 conftest).
"""
import sys
import time

sys.path.insert(0, '.')

import pandas as pd

from dashboard.utils import smart_cache_manager as scm
from dashboard.utils.smart_cache_manager import (
    CacheStrategy, smart_get, smart_set, smart_set_invalidation_marker, get_smart_cache,
)

KEY = 'test_edit_marker_sheet'


def _cleanup():
    r = get_smart_cache().redis
    r.delete(f'cache:{KEY}', f'invalidation:{KEY}')


def test_marker_outlives_slow_sheet_load():
    assert scm.INVALIDATION_MARKER_TTL >= 60   # 시트 API timeout 60s 이상


def test_stale_inflight_load_cannot_overwrite_edit():
    _cleanup()
    try:
        load_started = time.time() - 15          # 편집 15초 전에 시작한 전체 로드
        smart_set_invalidation_marker(KEY)        # 편집 시각 마커
        smart_set(KEY, 'edited', CacheStrategy.CRITICAL_DATA)
        smart_set(KEY, 'stale', CacheStrategy.CRITICAL_DATA, fetched_at=load_started)
        assert smart_get(KEY, CacheStrategy.CRITICAL_DATA) == 'edited'
        assert get_smart_cache().redis.ttl(f'invalidation:{KEY}') > 60
        # 편집 뒤 시작한 로드는 정상 저장
        smart_set(KEY, 'fresh', CacheStrategy.CRITICAL_DATA, fetched_at=time.time())
        assert smart_get(KEY, CacheStrategy.CRITICAL_DATA) == 'fresh'
    finally:
        _cleanup()


def test_update_project_in_cache_sets_marker_before_write(monkeypatch):
    from dashboard.services import project_service as ps
    calls = []
    df = pd.DataFrame([{'프로젝트 코드': 'R4163-TH', '부가세': True, '총액 1': 2500000}])
    monkeypatch.setattr(ps, 'smart_get', lambda *a, **k: df)
    monkeypatch.setattr(ps, 'smart_set', lambda key, value, *a, **k: calls.append(('set', key)))
    monkeypatch.setattr(scm, 'smart_set_invalidation_marker', lambda key, *a, **k: calls.append(('marker', key)))
    assert ps.update_project_in_cache('R4163-TH', {'부가세': 'FALSE'}) is True
    assert calls == [('marker', 'current_sheet_data'), ('set', 'current_sheet_data')]
    assert bool(df.loc[0, '부가세']) is False
