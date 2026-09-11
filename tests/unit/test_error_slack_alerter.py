# -*- coding: utf-8 -*-
"""error_slack_alerter._AlertGate 큐레이션 로직 단위 테스트.

네트워크·Redis 없이(redis_getter=None → in-memory 폴백) 제어 가능한 시계로
분류/억제 규칙만 검증한다.
"""
import logging

from dashboard.utils.error_slack_alerter import (
    _AlertGate, SPIKE_THRESHOLD, GLOBAL_RATE_MAX,
    RECURRING_COOLDOWN_SEC, CRITICAL_COOLDOWN_SEC, MIN_KNOWN_FOR_NEW,
)


class _Clock:
    def __init__(self, t=1000.0):
        self.t = t

    def __call__(self):
        return self.t

    def tick(self, dt):
        self.t += dt


def _rec(msg, name='dashboard.blueprints.x', level=logging.ERROR,
         lineno=10, exc_info=None):
    r = logging.LogRecord(name=name, level=level, pathname='x.py',
                          lineno=lineno, msg=msg, args=(), exc_info=exc_info,
                          func='f')
    return r


def _gate(clock=None):
    return _AlertGate(now=(clock or _Clock()), redis_getter=None)


# ── denylist ────────────────────────────────────────────────
def test_denies_noise_logger_prefix():
    g = _gate()
    assert g.decide(_rec('boom', name='engineio.server')) is None
    assert g.decide(_rec('boom', name='waitress.queue')) is None


def test_denies_noise_message():
    g = _gate()
    assert g.decide(_rec('Invalid session xyz')) is None
    assert g.decide(_rec('I/O operation on closed file')) is None


def test_below_error_level_ignored():
    g = _gate()
    assert g.decide(_rec('just a warning', level=logging.WARNING)) is None


# ── transient: 급증 시에만 ──────────────────────────────────
def test_transient_ignored_until_spike():
    clock = _Clock()
    g = _gate(clock)
    msg = 'Google Sheets 연결 오류: TimeoutError - The read operation timed out'
    # 임계값 직전까지는 전부 무시
    for _ in range(SPIKE_THRESHOLD - 1):
        assert g.decide(_rec(msg)) is None
    # 임계값 도달 순간 1건 급증 경보
    d = g.decide(_rec(msg))
    assert d is not None and d['kind'] == 'spike'
    assert d['window_count'] >= SPIKE_THRESHOLD
    # 직후 동일 유형은 쿨다운으로 억제
    assert g.decide(_rec(msg)) is None


# ── normal: new vs recurring + dedup ────────────────────────
def test_new_signature_after_warmup():
    clock = _Clock()
    g = _gate(clock)
    # 서로 다른 라인의 일반 에러 MIN_KNOWN 개로 seen 워밍업 (전부 recurring)
    for i in range(MIN_KNOWN_FOR_NEW):
        d = g.decide(_rec(f'일반 에러 A{i}', lineno=100 + i))
        assert d is not None and d['kind'] == 'recurring'
    # 워밍업 후 처음 보는 시그니처 → new
    d = g.decide(_rec('완전히 새로운 에러', lineno=999))
    assert d is not None and d['kind'] == 'new'


def test_recurring_dedup_within_cooldown():
    clock = _Clock()
    g = _gate(clock)
    first = g.decide(_rec('반복되는 일반 에러', lineno=50))
    assert first is not None
    # 같은 시그니처 즉시 재발 → 억제
    assert g.decide(_rec('반복되는 일반 에러', lineno=50)) is None
    # 쿨다운 경과 후 재알림
    clock.tick(RECURRING_COOLDOWN_SEC + 1)
    assert g.decide(_rec('반복되는 일반 에러', lineno=50)) is not None


# ── critical: 무조건(짧은 dedup) ────────────────────────────
def test_critical_always_even_if_transient_text():
    clock = _Clock()
    g = _gate(clock)
    msg = 'TimeoutError 지만 CRITICAL 이면 통과'
    d = g.decide(_rec(msg, level=logging.CRITICAL))
    assert d is not None and d['kind'] == 'critical'
    # 120초 내 동일 → 억제
    assert g.decide(_rec(msg, level=logging.CRITICAL)) is None
    clock.tick(CRITICAL_COOLDOWN_SEC + 1)
    assert g.decide(_rec(msg, level=logging.CRITICAL)) is not None


# ── 전역 rate cap ───────────────────────────────────────────
def test_global_rate_cap_suppresses_extra():
    clock = _Clock()
    g = _gate(clock)
    sent = 0
    # 서로 다른 시그니처를 다수 투입(각자 쿨다운·new/recurring 무관하게 전역 상한 확인)
    for i in range(GLOBAL_RATE_MAX + 4):
        d = g.decide(_rec(f'서로 다른 에러 {i}', lineno=200 + i))
        if d is not None:
            sent += 1
    assert sent == GLOBAL_RATE_MAX
    # 상한 초과분은 suppressed 카운터로 누적되어 다음 허용 경보에 표기됨
    clock.tick(61)  # rate 윈도우 경과
    d = g.decide(_rec('윈도우 경과 후 에러', lineno=300))
    assert d is not None
    assert d['suppressed'] >= 4
