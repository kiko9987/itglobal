# -*- coding: utf-8 -*-
"""공사 카드 스레드 연결이 옛 번호로 남아도 등록증은 지금 번호로 (2026-11-20 R4297-YG 계기).

11-16 번호 당기기(R4298-YG → R4297-YG) 뒤 project_thread 연결 8개가 옛 번호로 남았고,
대표님이 올린 사업자등록증을 없는 R4298-YG 폴더에 넣으려다 세 번 실패했다.
"""
import sys

sys.path.insert(0, '.')

import pytest


class _R:
    def __init__(self, d=None, h=None):
        self.d = dict(d or {})
        self.h = dict(h or {})

    def get(self, k):
        return self.d.get(k)

    def set(self, k, v, ex=None):
        self.d[k] = v

    def ttl(self, k):
        return 100

    def hget(self, name, k):
        return self.h.get(k)

    def hgetall(self, name):
        return self.h

    def scan_iter(self, match='*', count=None):
        import fnmatch
        return [k for k in list(self.d) if fnmatch.fnmatch(k, match)]

    def renamenx(self, a, b):
        if b in self.d:
            return False
        self.d[b] = self.d.pop(a)
        return True

    def hset(self, name, k, v):
        self.h[k] = v


@pytest.fixture
def r(monkeypatch):
    rr = _R({'project_thread:C1|1.1': 'R4298-YG'}, {'R4298-YG': 'R4297-YG'})
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('X', (), {'redis': rr})())
    from dashboard.services import project_rekey as pr
    monkeypatch.setattr(pr, '_rc', lambda rc=None: rr)
    return rr


def test_thread_link_follows_rekey_and_heals(r):
    from dashboard.services.business_license_handler import resolve_project_from_thread
    assert resolve_project_from_thread('C1', '1.1') == 'R4297-YG'
    assert r.d['project_thread:C1|1.1'] == 'R4297-YG'          # 연결도 고쳐 둠


def test_rekey_updates_every_thread_link_with_old_code(r):
    from dashboard.services.project_rekey import rekey_project_redis
    r.d['project_thread:C1|2.2'] = 'G4301-YG'
    r.d['project_thread:C1|3.3'] = 'G4301-YGX'                 # 다른 코드는 그대로
    rekey_project_redis('G4301-YG', 'G4300-YG', rc=r)
    assert r.d['project_thread:C1|2.2'] == 'G4300-YG'
    assert r.d['project_thread:C1|3.3'] == 'G4301-YGX'
