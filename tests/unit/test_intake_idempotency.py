# -*- coding: utf-8 -*-
"""인입 확인 멱등 헬퍼 회귀 테스트 (orphan claim 방지).

핵심 계약:
  1. 정상 1회 커밋:  go → mark_done → 재확인은 'done' 차단
  2. 커밋 중단(시트 timeout·재시작): go → (커밋 유실, mark_done 안 함) →
     락 TTL 만료 후 재확인이 다시 'go' (영구 차단 X)  ← 2026-08-26 G3742-JW 사고의 핵심
  3. 커밋 실패: go → release → 즉시 재확인 'go' (락 만료 안 기다림)
  4. 동시 확인(락 유효 중): 두번째는 'busy'
  5. Redis 불가(rc=None)·예외: 항상 'go' (기존 동작 유지, 차단 안 함)

fakeredis 로 SETNX·TTL·get/delete 시맨틱만 검증 — 슬랙/시트 미접촉.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.blueprints.slack_bot import (
    _intake_acquire, _intake_mark_done, _intake_release, _intake_key_pair,
)

IID, PC, STG = 'intake_abc', 'G3742-JW', '잔금'


class FakeRedis:
    """SETNX(nx)·EX(만료)·get·delete 만 흉내내는 최소 가짜 레디스.

    만료는 실시간 대신 tick(n) 으로 결정적으로 진행 → sleep 없이 TTL 검증.
    """
    def __init__(self):
        self.store = {}      # key -> value
        self.expire = {}     # key -> 남은 tick 수 (None=무기한)
        self.raise_on = set()  # 이 op 이름이면 예외 (예: {'set'})

    def _expired(self, key):
        return key in self.store and self.expire.get(key) is not None and self.expire[key] <= 0

    def _sweep(self, key):
        if self._expired(key):
            self.store.pop(key, None)
            self.expire.pop(key, None)

    def get(self, key):
        if 'get' in self.raise_on:
            raise RuntimeError('redis down')
        self._sweep(key)
        return self.store.get(key)

    def set(self, key, val, nx=False, ex=None):
        if 'set' in self.raise_on:
            raise RuntimeError('redis down')
        self._sweep(key)
        if nx and key in self.store:
            return None            # 이미 존재 → SETNX 실패
        self.store[key] = val
        self.expire[key] = ex
        return True

    def delete(self, key):
        if 'delete' in self.raise_on:
            raise RuntimeError('redis down')
        self.store.pop(key, None)
        self.expire.pop(key, None)
        return 1

    def tick(self, n=1):
        """모든 TTL 을 n 만큼 진행."""
        for k in list(self.expire):
            if self.expire[k] is not None:
                self.expire[k] -= n


def test_normal_commit_then_blocked():
    rc = FakeRedis()
    assert _intake_acquire(rc, IID, PC, STG) == 'go'
    _intake_mark_done(rc, IID, PC, STG)
    # 재확인(슬랙 재시도·더블클릭)은 완료마커로 차단
    assert _intake_acquire(rc, IID, PC, STG) == 'done'
    # 락은 해제돼 남지 않음
    _, lock_key = _intake_key_pair(IID, PC, STG)
    assert rc.get(lock_key) is None


def test_interrupted_commit_recovers_after_lock_expires():
    """커밋이 중간에 끊겨(mark_done 미호출) 완료마커가 없으면,
    락 TTL 만료 후 재확인이 다시 'go' 여야 한다 (orphan 영구차단 방지)."""
    rc = FakeRedis()
    assert _intake_acquire(rc, IID, PC, STG, lock_ttl=120) == 'go'
    # ── 여기서 커밋이 시트 timeout/재시작으로 유실됐다고 가정 (mark_done 안 함) ──
    # 락 유효 중엔 재시도가 'busy' (동시 이중커밋 방지)
    assert _intake_acquire(rc, IID, PC, STG, lock_ttl=120) == 'busy'
    # 락 TTL(120) 경과
    rc.tick(120)
    # 이제 재확인이 다시 열린다 → 실제로 재기록 가능
    assert _intake_acquire(rc, IID, PC, STG, lock_ttl=120) == 'go'


def test_failed_commit_release_allows_immediate_retry():
    rc = FakeRedis()
    assert _intake_acquire(rc, IID, PC, STG) == 'go'
    _intake_release(rc, IID, PC, STG)        # 커밋 실패 → 즉시 락 해제
    # 락 만료를 기다리지 않고 바로 재시도 가능
    assert _intake_acquire(rc, IID, PC, STG) == 'go'


def test_concurrent_confirm_is_busy():
    rc = FakeRedis()
    assert _intake_acquire(rc, IID, PC, STG) == 'go'    # 첫 확인 진행중
    assert _intake_acquire(rc, IID, PC, STG) == 'busy'  # 동시 두번째 → 대기


def test_done_marker_survives_but_lock_expires():
    """완료마커(done)는 하루 유지, 락은 짧게 만료 — done 이 우선 차단."""
    rc = FakeRedis()
    _intake_acquire(rc, IID, PC, STG)
    _intake_mark_done(rc, IID, PC, STG)
    rc.tick(120)                              # 락은 사라질 시간
    assert _intake_acquire(rc, IID, PC, STG) == 'done'   # 그래도 done 이 막는다


def test_redis_none_never_blocks():
    # 레디스 불가 시 항상 진행 (기존 동작). mark_done/release 는 무해.
    assert _intake_acquire(None, IID, PC, STG) == 'go'
    _intake_mark_done(None, IID, PC, STG)
    _intake_release(None, IID, PC, STG)
    assert _intake_acquire(None, IID, PC, STG) == 'go'


def test_redis_exception_degrades_to_go():
    rc = FakeRedis()
    rc.raise_on.add('get')     # get 에서 예외
    assert _intake_acquire(rc, IID, PC, STG) == 'go'   # 차단하지 않고 진행
    rc.raise_on.clear()
    rc.raise_on.add('set')     # 락 SETNX 에서 예외
    assert _intake_acquire(rc, IID, PC, STG) == 'go'


def test_different_stage_independent():
    rc = FakeRedis()
    _intake_acquire(rc, IID, PC, '중도금')
    _intake_mark_done(rc, IID, PC, '중도금')
    # 같은 프로젝트라도 단계가 다르면 독립 (잔금은 아직 가능)
    assert _intake_acquire(rc, IID, PC, '잔금') == 'go'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
