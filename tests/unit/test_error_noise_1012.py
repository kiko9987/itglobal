"""2026-10-12 Sentry 오류 점검 회귀 테스트.

- [취소 되돌리기] 버튼마다 NameError(_bg) — 리팩터 잔재 한 줄(18f9205 이후)
- 리드 강등 List fallback 의 old_lead 미정의
- 서버 재시작 때 Pub/Sub 루프가 닫힌 연결 예외를 ERROR 로 남기던 잡음
"""
import builtins
import logging
import symtable
import time
from pathlib import Path
from types import SimpleNamespace

from dashboard.utils.cache_invalidation import CacheInvalidationService

ROOT = Path(__file__).resolve().parents[2] / 'dashboard'

# 의도적으로 globals() 검사로 감싼 이름 (실행돼도 예외 없음)
_ALLOWED = {('utils/error_handler.py', 'request')}


def _undefined_names(path: Path):
    src = path.read_text(encoding='utf-8-sig')
    if 'import *' in src:
        return []
    top = symtable.symtable(src, str(path), 'exec')
    known = set(dir(builtins)) | {'__file__', '__name__', '__doc__', '__spec__', '__path__',
                                  '__package__', '__builtins__', '__loader__'}
    known |= {s.get_name() for s in top.get_symbols() if s.is_assigned() or s.is_imported()}

    def _global_assigns(t):
        for s in t.get_symbols():
            if s.is_declared_global() and s.is_assigned():
                known.add(s.get_name())
        for c in t.get_children():
            _global_assigns(c)
    _global_assigns(top)

    found = []

    def _walk(t):
        for s in t.get_symbols():
            if s.is_referenced() and s.is_global() and not s.is_declared_global() \
                    and s.get_name() not in known:
                found.append((t.get_name(), s.get_name()))
        for c in t.get_children():
            _walk(c)
    for c in top.get_children():
        _walk(c)
    return found


def test_no_undefined_global_names_in_functions():
    """함수 안에서 정의 안 된 이름을 읽는 코드(NameError 잠복)가 없어야 한다."""
    problems = []
    for path in ROOT.rglob('*.py'):
        rel = path.relative_to(ROOT).as_posix()
        for func, name in _undefined_names(path):
            if (rel, name) not in _ALLOWED:
                problems.append(f'{rel}: {func}() -> {name}')
    assert problems == []


def _make_service(pubsub):
    svc = CacheInvalidationService.__new__(CacheInvalidationService)
    svc.redis = SimpleNamespace(redis=SimpleNamespace(pubsub=lambda: pubsub, ping=lambda: True))
    svc._pubsub = None
    svc._subscriber_thread = None
    svc._handlers = []
    svc._running = True
    return svc


class _FakePubSub:
    def __init__(self, on_get):
        self._on_get = on_get

    def subscribe(self, *_a):
        pass

    def get_message(self, **_kw):
        return self._on_get()

    def close(self):
        pass


def test_pubsub_close_during_shutdown_is_not_error(caplog):
    holder = {}

    def _on_get():
        holder['svc']._running = False  # stop() 이 먼저 플래그를 내리고 연결을 닫음
        raise ValueError('I/O operation on closed file.')

    svc = _make_service(_FakePubSub(_on_get))
    holder['svc'] = svc
    with caplog.at_level(logging.DEBUG, logger='dashboard.utils.cache_invalidation'):
        svc._subscriber_loop()
    assert not [r for r in caplog.records if r.levelno >= logging.ERROR]


def test_pubsub_unexpected_error_while_running_still_logged(caplog, monkeypatch):
    monkeypatch.setattr(time, 'sleep', lambda _s: None)
    calls = {'n': 0}
    holder = {}

    def _on_get():
        calls['n'] += 1
        if calls['n'] == 1:
            raise RuntimeError('boom')
        holder['svc']._running = False
        return None

    svc = _make_service(_FakePubSub(_on_get))
    holder['svc'] = svc
    with caplog.at_level(logging.DEBUG, logger='dashboard.utils.cache_invalidation'):
        svc._subscriber_loop()
    errors = [r for r in caplog.records if r.levelno >= logging.ERROR]
    assert len(errors) == 1 and 'boom' in errors[0].getMessage()
