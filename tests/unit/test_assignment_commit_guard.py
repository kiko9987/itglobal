# -*- coding: utf-8 -*-
"""일정 확정(commit) 연타/중복 실행 가드 회귀 테스트 (2026-09-15).

배경: JW 가 캔버스2 편집 중 /일정확정 을 여러 번 누르자, commit 들이 경합하며
편집 중인 캔버스를 매번 다시 읽어 앞서 보낸 DM 을 '완전 제거'로 지우는 race →
그날 방문 담당 DM 이 전량 유실. 가드로 실행 중·직후 쿨다운 재실행을 즉시 거절.

pure 단위 테스트 — Redis/Slack/시트 미접촉 (get_redis_client·_commit_locked mock).
"""
import sys
sys.path.insert(0, '.')

import pytest
import dashboard.services.visit_assignment_sync as vas


class _FakeRedis:
    def __init__(self, acquire_ok):
        self._acquire_ok = acquire_ok
        self.calls = []

    def set(self, key, val, nx=False, ex=None):
        self.calls.append(('set', key, nx, ex))
        # nx=True 면 이미 점유 중일 때 None(거짓) 반환 (실제 redis-py 동작)
        return True if self._acquire_ok else None

    def expire(self, key, ttl):
        self.calls.append(('expire', key, ttl))
        return True

    def delete(self, key):
        self.calls.append(('delete', key))
        return 1


def _patch_redis(monkeypatch, fake):
    import dashboard.utils.redis_client as rcmod
    monkeypatch.setattr(rcmod, 'get_redis_client',
                        lambda: type('C', (), {'redis': fake})())


def test_busy_guard_rejects_without_running_body(monkeypatch):
    """실행 중(가드 점유)이면 본체 미진입하고 즉시 거절."""
    fake = _FakeRedis(acquire_ok=False)
    _patch_redis(monkeypatch, fake)
    # 본체가 호출되면 실패시킴 — 거절 경로는 본체를 안 타야 함
    monkeypatch.setattr(vas, '_commit_locked',
                        lambda: pytest.fail('본체(_commit_locked)가 호출됨 — 가드 실패'))

    res = vas.commit()
    assert res['ok'] is False
    assert '처리 중' in res['reason'] or '방금' in res['reason']
    # 쿨다운 전환/삭제도 하지 않음 (애초에 우리가 점유한 게 아니므로)
    assert not any(c[0] in ('expire', 'delete') for c in fake.calls)


def test_acquire_runs_body_then_cooldown(monkeypatch):
    """가드 획득 성공 시 본체 실행하고, 완료 후 쿨다운 TTL 로 전환."""
    fake = _FakeRedis(acquire_ok=True)
    _patch_redis(monkeypatch, fake)
    sentinel = {'ok': True, 'marker': 'ran'}
    monkeypatch.setattr(vas, '_commit_locked', lambda: sentinel)

    res = vas.commit()
    assert res is sentinel                       # 본체 결과 그대로 반환
    # 획득(set nx) 후 완료 시 쿨다운(expire) 전환
    assert ('set', vas._COMMIT_GUARD_KEY, True, vas._COMMIT_INFLIGHT_TTL) in fake.calls
    assert ('expire', vas._COMMIT_GUARD_KEY, vas._COMMIT_COOLDOWN_TTL) in fake.calls


def test_body_exception_still_sets_cooldown(monkeypatch):
    """본체가 예외를 던져도 finally 에서 쿨다운 전환 (가드 영구 점유 방지)."""
    fake = _FakeRedis(acquire_ok=True)
    _patch_redis(monkeypatch, fake)

    def _boom():
        raise RuntimeError('boom')
    monkeypatch.setattr(vas, '_commit_locked', _boom)

    with pytest.raises(RuntimeError):
        vas.commit()
    assert ('expire', vas._COMMIT_GUARD_KEY, vas._COMMIT_COOLDOWN_TTL) in fake.calls


def test_redis_down_bypasses_guard(monkeypatch):
    """Redis 장애 시 가드를 우회하고 본체 실행 (가용성 우선)."""
    import dashboard.utils.redis_client as rcmod

    def _raise():
        raise ConnectionError('redis down')
    monkeypatch.setattr(rcmod, 'get_redis_client', _raise)
    sentinel = {'ok': True, 'marker': 'ran-nogaurd'}
    monkeypatch.setattr(vas, '_commit_locked', lambda: sentinel)

    res = vas.commit()
    assert res is sentinel   # 가드 우회하고 정상 실행


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
