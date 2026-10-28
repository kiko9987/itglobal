# -*- coding: utf-8 -*-
"""2026-10-28 4차 — 대기 중 버전·코드 변경 Redis 연결·메모 지움·리드 API 주소·인증 빈틈."""
import sys

sys.path.insert(0, '.')

import pytest


class _R:
    """간단 Redis 흉내 (scan_iter·renamenx·get/set/ttl/delete)."""
    def __init__(self, d=None):
        self.d = dict(d or {})

    def scan_iter(self, match='*', count=None):
        import fnmatch
        return [k for k in list(self.d) if fnmatch.fnmatch(k, match)]

    def renamenx(self, a, b):
        if b in self.d:
            return False
        self.d[b] = self.d.pop(a)
        return True

    def get(self, k):
        return self.d.get(k)

    def set(self, k, v, ex=None):
        self.d[k] = v

    def setex(self, k, ttl, v):
        self.d[k] = v

    def ttl(self, k):
        return 100

    def delete(self, k):
        self.d.pop(k, None)


def test_rekey_moves_only_exact_code_tokens():
    from dashboard.services.project_rekey import rekey_project_redis
    r = _R({
        'invoice_thread:G4153-JW': 'h', 'payment_slack:ts:G4153-JW:잔금': '1.2',
        'license_state:G4153-JW': 'x', 'card_field_snap:G4153-JWX': 'other',   # 다른 코드 — 안 건드림
        'project_card_msg:G4153-JW': 'C1|9.9', 'project_thread:C1|9.9': 'G4153-JW',
        'invariant_seen:abc:G4153-JW': '1',                                        # 기록 — 안 건드림
    })
    moved = rekey_project_redis('G4153-JW', 'G4153-YM', rc=r)
    assert 'invoice_thread:G4153-YM' in r.d and 'payment_slack:ts:G4153-YM:잔금' in r.d
    assert 'license_state:G4153-YM' in r.d and 'card_field_snap:G4153-JWX' in r.d
    assert r.d['project_thread:C1|9.9'] == 'G4153-YM'
    assert 'invariant_seen:abc:G4153-JW' in r.d
    assert len(moved) == 4


def test_rekey_does_not_overwrite_existing_target():
    from dashboard.services.project_rekey import rekey_project_redis
    r = _R({'license_state:G1-AA': 'old', 'license_state:G1-BB': 'keep'})
    assert rekey_project_redis('G1-AA', 'G1-BB', rc=r) == []
    assert r.d['license_state:G1-BB'] == 'keep'


def test_pending_version_counts_as_current(monkeypatch):
    import dashboard.blueprints.projects as pj
    r = _R()
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': r})())
    monkeypatch.setattr(pj, '_version_col_index', lambda m: 1)
    row = ['G1', '5']
    # 시트는 아직 5, 방금 내 저장이 6 으로 대기 중 → 화면이 6 으로 다시 저장해도 통과
    pj._pending_version_set('G1', 6)
    v, err = pj._check_optimistic_lock_update(None, 's', 'n', 'G1', 2, row, '6')
    assert err is None and v == 7
    # 시트에 6 이 쓰이면 대기 표시 지움
    pj._pending_version_clear('G1', 6)
    assert pj._pending_version_get('G1') is None


class _NoteMgr:
    def __init__(self, notes):
        self.notes = notes

    def get_cell_note(self, sid, sn, addr):
        return self.notes.get(addr, '')


def test_batch_memo_clear_only_when_seen(monkeypatch):
    import dashboard.blueprints.projects as pj
    monkeypatch.setattr(pj, 'payment_field_column', lambda f: 'X')
    mgr = _NoteMgr({'X5': '옛 메모'})
    # 본 메모를 지움 → 지움 대상
    _r, failed, upd, _o = pj._prepare_batch_memo_updates(
        [{'field_name': '잔금', 'memo': None, 'original_memo': '옛 메모'}], 5, mgr, 's', 'n', 'G1')
    assert failed == 0 and upd == {'X5': ''}
    # original 없이 빈값 → 예전처럼 건너뜀 (G3823 사고 방지)
    _r, failed, upd, _o = pj._prepare_batch_memo_updates(
        [{'field_name': '잔금', 'memo': ''}], 5, mgr, 's', 'n', 'G1')
    assert upd == {}


def test_lead_write_api_under_api_prefix():
    from dashboard.blueprints.leads import leads_api_bp
    assert leads_api_bp.url_prefix == '/api/leads'


def test_bearer_token_no_longer_bypasses():
    from dashboard.utils.security_middleware import SecurityMiddleware
    assert SecurityMiddleware._validate_jwt_token(object.__new__(SecurityMiddleware), 'anything') is False
