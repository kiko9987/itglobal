"""백그라운드 시트 미리 받기 — 경과 비율 계산(2026-11-15).

예전엔 Redis 의 '남은' TTL 을 전체 TTL 로 보고 계산해 경과 비율이 늘 1.0 이었다
→ 30초 확인마다 시트 4천 행을 통째로 다시 받아 사이트 전체가 45초마다 10~15초씩 느려졌다.
"""
from dashboard.utils import background_prefetch as bp
from dashboard.utils.smart_cache_manager import CacheStrategy


class _FakeCache:
    def __init__(self, remaining):
        self.remaining = remaining

    def _get_cache_item_info(self, key):
        if self.remaining is None:
            return None
        return {'ttl': self.remaining}


def _run(monkeypatch, remaining, threshold=bp.PREFETCH_THRESHOLD):
    pf = bp.BackgroundPrefetch()
    pf.register_prefetch_job('current_sheet_data', lambda: None, CacheStrategy.CRITICAL_DATA, threshold)
    calls = []
    monkeypatch.setattr(bp, 'get_smart_cache', lambda: _FakeCache(remaining))
    monkeypatch.setattr(pf, '_execute_prefetch', lambda key, info: calls.append(key))
    pf._check_and_prefetch()
    return calls


def test_fresh_cache_is_not_refetched(monkeypatch):
    # 방금 저장(남은 290초 / 전체 300초) → 경과 3% — 다시 받지 않음(예전엔 여기서도 받았음)
    assert _run(monkeypatch, 290) == []


def test_refetch_after_threshold(monkeypatch):
    # 남은 100초 → 경과 200/300 = 67% ≥ 60% → 받음
    assert _run(monkeypatch, 100) == ['current_sheet_data']


def test_just_before_threshold_waits(monkeypatch):
    # 남은 130초 → 경과 170/300 = 57% < 60% → 아직
    assert _run(monkeypatch, 130) == []


def test_missing_cache_fetches_now(monkeypatch):
    assert _run(monkeypatch, None) == ['current_sheet_data']


def test_longer_ttl_than_strategy_counts_as_fresh(monkeypatch):
    # 다른 경로가 더 긴 TTL 로 저장했으면 경과 0 으로 본다(음수 비율 없음)
    assert _run(monkeypatch, 1200) == []


def test_threshold_means_about_three_minutes():
    assert bp._full_ttl(CacheStrategy.CRITICAL_DATA) * bp.PREFETCH_THRESHOLD == 180
