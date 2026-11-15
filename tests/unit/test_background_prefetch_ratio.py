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


# ── 리드 시트: TEMPORARY(60초) 기준 ──────────────────────────────────
def _run_leads(monkeypatch, remaining):
    pf = bp.BackgroundPrefetch()
    pf.register_prefetch_job('online_leads_sheet_data', lambda: None, CacheStrategy.TEMPORARY,
                             bp.LEADS_PREFETCH_THRESHOLD)
    calls = []
    monkeypatch.setattr(bp, 'get_smart_cache', lambda: _FakeCache(remaining))
    monkeypatch.setattr(pf, '_execute_prefetch', lambda key, info: calls.append(key))
    pf._check_and_prefetch()
    return calls


def test_leads_fresh_waits(monkeypatch):
    assert _run_leads(monkeypatch, 50) == []          # 10/60 경과


def test_leads_refetched_before_expiry(monkeypatch):
    assert _run_leads(monkeypatch, 30) == ['online_leads_sheet_data']   # 30/60 경과 ≥ 0.4


# ── 캐시 워밍 주기 갱신: 미리 받기가 맡은 공사현황 시트는 건너뜀 ─────────
def _covers_with(monkeypatch, running=True, job=True, disabled=False, backoff=None):
    from dashboard.utils import cache_preloader as cp
    pf = bp.BackgroundPrefetch()
    pf._running = running
    if job:
        pf.register_prefetch_job('current_sheet_data', lambda: None)
        pf._prefetch_jobs['current_sheet_data']['is_disabled'] = disabled
        pf._prefetch_jobs['current_sheet_data']['backoff_until'] = backoff
    monkeypatch.setattr(bp, 'get_background_prefetch', lambda: pf)
    return cp._prefetch_covers('current_sheet_data')


def test_preloader_skips_when_prefetch_running(monkeypatch):
    assert _covers_with(monkeypatch) is True


def test_preloader_refreshes_when_prefetch_stopped(monkeypatch):
    assert _covers_with(monkeypatch, running=False) is False
    assert _covers_with(monkeypatch, job=False) is False
    assert _covers_with(monkeypatch, disabled=True) is False


def test_preloader_refreshes_when_prefetch_backing_off(monkeypatch):
    import time
    assert _covers_with(monkeypatch, backoff=time.time() + 120) is False
