# -*- coding: utf-8 -*-
"""거래처 탭 캐시 10분 재구성 (2026-10-22 — 원래 매일 07:00 이라 새 줄이 다음 날 반영).

- 지우기+쓰기를 한 트랜잭션(pipeline)으로 → 재구성 중 조회가 빈 캐시를 만나지 않음.
- 새로 들어온 상호의 공사는 '등록증 없음' 판정 캐시를 비워 바로 계산서 요청 가능.
- 처음 만들 때(빈 캐시)는 전부 새 상호라 판정 캐시를 건드리지 않음.
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.services.partner_status_sync as ps


class _Pipe:
    def __init__(self, r):
        self.r, self.ops = r, []

    def delete(self, k):
        self.ops.append(('delete', k))

    def hset(self, k, mapping):
        self.ops.append(('hset', k, dict(mapping)))

    def sadd(self, k, *v):
        self.ops.append(('sadd', k, set(v)))

    def execute(self):
        self.r.executed.append(list(self.ops))
        for op in self.ops:
            if op[0] == 'delete':
                self.r.data.pop(op[1], None)
            elif op[0] == 'hset':
                self.r.data.setdefault(op[1], {}).update(op[2])
            else:
                self.r.data.setdefault(op[1], set()).update(op[2])


class _R:
    def __init__(self, names=()):
        self.data = {ps._NAME_SET_KEY: set(names)} if names else {}
        self.executed = []

    def smembers(self, k):
        return set(self.data.get(k) or ())

    def pipeline(self, transaction=True):
        assert transaction
        return _Pipe(self)


ROWS = [
    ['등록번호', '상호', '대표자', '주소', '업태', '종목', '이메일'],
    ['105-86-00001', '주식회사 기존상사', '김', '서울', '도매', '기계', 'old@x.com'],
    ['566-87-02004', '주식회사 디자인애드올', '박종찬', '서울 강남', '전문디자인업', '인테리어', 'ahalldesign@naver.com'],
]


@pytest.fixture
def env(monkeypatch):
    class _Vals:
        def get(self, **k):
            return type('E', (), {'execute': lambda self: {'values': ROWS}})()

    svc = type('S', (), {'spreadsheets': lambda self: type('SS', (), {'values': lambda self: _Vals()})()})()
    monkeypatch.setattr(ps, '_sheet', lambda: (type('M', (), {'service': svc})(), 'SID'))
    import dashboard.services.project_service as psv
    monkeypatch.setattr(psv, 'get_project_records', lambda: [
        {'프로젝트 코드': 'G4058-YG', '사업자명': '주식회사 디자인애드올'},
        {'프로젝트 코드': 'R1-X', '사업자명': '주식회사 기존상사'},
        {'프로젝트 코드': 'R2-X', '사업자명': '-'},
    ])
    import dashboard.services.business_license_handler as blh
    cleared = []
    monkeypatch.setattr(blh, 'invalidate_license_state', lambda code: cleared.append(code))
    import dashboard.utils.redis_client as rcm
    holder = {}

    def _use(names=()):
        holder['r'] = _R(names)
        monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': holder['r']})())
        return holder['r']
    return _use, cleared


def test_new_partner_clears_license_verdict(env):
    use, cleared = env
    r = use(names={ps._norm_name('주식회사 기존상사'), ps._norm_name('상호')})
    res = ps.rebuild_partner_caches()
    assert res['added'] == 1 and cleared == ['G4058-YG']          # 새로 넣은 줄의 공사만
    key = ps._norm_name('주식회사 디자인애드올')
    assert key in r.data[ps._NAME_SET_KEY]
    assert r.data[ps._EMAIL_CACHE_KEY][key] == 'ahalldesign@naver.com'


def test_rebuild_is_one_transaction(env):
    use, _ = env
    r = use(names={ps._norm_name('주식회사 기존상사')})
    ps.rebuild_partner_caches()
    assert len(r.executed) == 1                                  # 지우기·쓰기 모두 한 번의 EXEC
    kinds = [op[0] for op in r.executed[0]]
    assert kinds.count('delete') == 4 and 'sadd' in kinds


def test_first_build_does_not_clear_all(env):
    use, cleared = env
    use()                                                        # 빈 캐시(첫 구성·Redis 초기화)
    res = ps.rebuild_partner_caches()
    assert res['added'] == 0 and cleared == []


def test_no_change_no_clear(env):
    use, cleared = env
    names = {ps._norm_name(x[1]) for x in ROWS}
    use(names=names)
    assert ps.rebuild_partner_caches()['added'] == 0 and cleared == []
