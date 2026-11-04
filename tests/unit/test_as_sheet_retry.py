"""A/S 시트 읽기 — 끊긴 연결(SSLEOFError)에서 바로 실패하지 않고 재연결 후 다시 읽는지 (2026-11-05)."""
import ssl
import threading

from dashboard.services import as_service as asv
from dashboard.utils.google_sheets import GoogleSheetsManager


class _Req:
    def __init__(self, fn):
        self._fn = fn

    def execute(self):
        return self._fn()


class _Svc:
    def __init__(self, fn):
        self._fn = fn
        self.ranges = []

    def spreadsheets(self):
        return self

    def values(self):
        return self

    def get(self, spreadsheetId, range, valueRenderOption):
        self.ranges.append(range)
        return _Req(self._fn)


class _Mgr:
    """실제 GoogleSheetsManager._execute_with_retry 를 그대로 돌리는 가짜 매니저."""
    MAX_RETRIES = GoogleSheetsManager.MAX_RETRIES
    INITIAL_BACKOFF = 0
    MAX_BACKOFF = 0

    def __init__(self, rows):
        def _eof():
            raise ssl.SSLEOFError(8, 'EOF occurred in violation of protocol (_ssl.c:2427)')
        self.service = _Svc(_eof)                        # 밤새 끊긴 연결
        self._fresh = _Svc(lambda: {'values': rows})
        self._lock = threading.RLock()
        self.resets = 0

    def reset_service(self):
        self.resets += 1
        self.service = self._fresh
        return True

    _execute_with_retry = GoogleSheetsManager._execute_with_retry


def test_list_as_recovers_from_ssl_eof(monkeypatch):
    mgr = _Mgr([['AS-0002', 'G1', '주소', '', '', '요청', 'JW', '', '', '', '', '접수 완료', ''],
                ['AS-0001', 'G2', '주소', '', '', '요청', 'JW', '', '', '', '', '요청', '']])
    monkeypatch.setattr(asv, '_get_sheet', lambda: (mgr, 'sid', 'A/S 관리'))
    items = asv.list_as()
    assert [x['No'] for x in items] == ['AS-0002', 'AS-0001']
    assert mgr.resets == 1 and mgr._fresh.ranges == ['A/S 관리!A2:M']   # 새 연결로 다시 읽음


def test_next_as_no_recovers_from_ssl_eof(monkeypatch):
    mgr = _Mgr([['No'], ['AS-0007'], ['AS-0012']])
    monkeypatch.setattr(asv, '_get_sheet', lambda: (mgr, 'sid', 'A/S 관리'))
    assert asv.get_next_as_no() == 'AS-0013'
    assert mgr.resets == 1
