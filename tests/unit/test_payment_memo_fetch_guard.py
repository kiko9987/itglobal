# -*- coding: utf-8 -*-
"""입금 폴러 — 메모 fetch 실패 시 사이클 skip + 스레드별 service (2026-10-06 14:00 사고).

핀 리마인드(자금 이동 스캔)와 폴러가 프로세스 싱글톤 service 를 동시에 써서 SSL 충돌 →
폴러가 빈 메모로 진행해 1786행 baseline phash 를 '' 로 비움 → 다음 사이클 대량 재baseline.
"""
import ssl
import sys
import threading
from unittest.mock import MagicMock

sys.path.insert(0, '.')

from dashboard.services import payment_sync as ps


class _FakeRedis:
    def __init__(self):
        self.writes = []

    def exists(self, k):
        return True  # baseline 완료 상태

    def hgetall(self, k):
        return {'u': '1000', 'v': '0', 'w': '0', 'aa': 'false', 'x': '0',
                'u_phash': 'abc', 'v_phash': '', 'w_phash': ''}

    def hset(self, *a, **kw):
        self.writes.append(('hset', a, kw))

    def expire(self, *a, **kw):
        self.writes.append(('expire', a))

    def get(self, k):
        return None


def _svc(notes_side_effect=None, notes_value=None):
    svc = MagicMock()
    row = ['G4139-MJ'] + [''] * 19 + [1000, 0, 0] + [''] * 7
    svc.spreadsheets.return_value.values.return_value.get.return_value.execute.return_value = {
        'values': [row]}
    g = svc.spreadsheets.return_value.get.return_value.execute
    if notes_side_effect is not None:
        g.side_effect = notes_side_effect
    else:
        g.return_value = notes_value
    return svc


def _run(monkeypatch, svc):
    fr = _FakeRedis()
    monkeypatch.setattr(ps, '_get_payment_service', lambda: svc)
    monkeypatch.setattr(ps, 'get_redis_client', lambda: MagicMock(redis=fr))
    resets = []
    monkeypatch.setattr(ps, '_reset_payment_service', lambda: resets.append(1))
    result = {'processed': 0, 'sent': 0, 'errors': 0}
    out = ps._sync_payments_locked(result, 'sid', '공사 현황', 'C1', 'xoxb-test')
    return out, fr, resets


def test_memo_fetch_ssl_error_skips_cycle_without_touching_baseline(monkeypatch):
    svc = _svc(notes_side_effect=ssl.SSLError('[SSL: WRONG_VERSION_NUMBER] wrong version number'))
    out, fr, resets = _run(monkeypatch, svc)
    assert out == {'processed': 0, 'sent': 0, 'errors': 0}
    assert fr.writes == []        # baseline phash 를 '' 로 덮어쓰지 않음
    assert resets == [1]          # SSL → service 재생성


def test_memo_fetch_other_error_also_skips(monkeypatch):
    svc = _svc(notes_side_effect=RuntimeError('quota'))
    out, fr, resets = _run(monkeypatch, svc)
    assert fr.writes == [] and resets == []


def test_empty_notes_response_after_baseline_skips(monkeypatch):
    svc = _svc(notes_value={'sheets': [{'data': [{'rowData': []}]}]})
    out, fr, _ = _run(monkeypatch, svc)
    assert fr.writes == []


def test_payment_service_is_per_thread(monkeypatch):
    import googleapiclient.discovery as disc
    from google.oauth2 import service_account
    monkeypatch.setattr(service_account.Credentials, 'from_service_account_file',
                        classmethod(lambda cls, *a, **kw: object()))
    monkeypatch.setattr(disc, 'build', lambda *a, **kw: object())
    ps._reset_payment_service()
    try:
        a1 = ps._get_payment_service()
        a2 = ps._get_payment_service()
        box = {}
        t = threading.Thread(target=lambda: box.setdefault('b', ps._get_payment_service()))
        t.start(); t.join()
        assert a1 is a2                      # 같은 스레드는 재사용
        assert box['b'] is not None and box['b'] is not a1   # 다른 스레드는 별개 인스턴스
        ps._reset_payment_service()
        assert ps._get_payment_service() is not a1           # 리셋은 호출 스레드 것만 재생성
    finally:
        ps._reset_payment_service()
