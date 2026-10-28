# -*- coding: utf-8 -*-
"""공사 코드가 바뀐 뒤 옛 카드 버튼·입금 지정 기록의 옛 코드 → 지금 코드 (2026-10-28)."""
import sys

sys.path.insert(0, '.')


class _R:
    def __init__(self):
        self.d, self.h = {}, {}

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

    def hset(self, n, k, v):
        self.h.setdefault(n, {})[k] = v

    def hget(self, n, k):
        return self.h.get(n, {}).get(k)

    def hdel(self, n, k):
        self.h.get(n, {}).pop(k, None)


def test_rekey_records_alias_and_resolves_chain():
    from dashboard.services.project_rekey import rekey_project_redis, resolve_project_code
    r = _R()
    rekey_project_redis('G4153-JW', 'G4153-YM', rc=r)
    rekey_project_redis('G4153-YM', 'R4153-YM', rc=r)
    assert resolve_project_code('G4153-JW', rc=r) == 'R4153-YM'      # 두 번 바뀐 옛 카드
    assert resolve_project_code('G4153-YM', rc=r) == 'R4153-YM'
    assert resolve_project_code('R4153-YM', rc=r) == 'R4153-YM'      # 지금 코드는 그대로
    assert resolve_project_code('G9999-AA', rc=r) == 'G9999-AA'      # 바뀐 적 없는 코드


def test_revert_back_to_old_code_is_live_again():
    from dashboard.services.project_rekey import rekey_project_redis, resolve_project_code
    r = _R()
    rekey_project_redis('G1-AA', 'G1-BB', rc=r)
    rekey_project_redis('G1-BB', 'G1-AA', rc=r)                       # 되돌림
    assert resolve_project_code('G1-AA', rc=r) == 'G1-AA'
    assert resolve_project_code('G1-BB', rc=r) == 'G1-AA'


def test_resolve_survives_redis_error():
    from dashboard.services.project_rekey import resolve_project_code

    class _Bad:
        def hget(self, *a):
            raise ConnectionError('down')
    assert resolve_project_code('G1-AA', rc=_Bad()) == 'G1-AA'


def test_intake_designation_codes_updated(monkeypatch):
    import dashboard.services.project_rekey as pr
    import dashboard.blueprints.slack_bot as sb
    monkeypatch.setattr(pr, 'resolve_project_code', lambda c, rc=None: {'G4153-JW': 'G4153-YM'}.get(c, c))
    d = {'designation': {
        'project_code': 'G4153-JW', 'stage': '잔금',
        'project_option': {'text': {'type': 'plain_text', 'text': 'G4153-JW · 대연 · 송파'}, 'value': 'G4153-JW'},
        'splits': [{'project_code': 'G4153-JW', 'project_option': {'text': {'text': 'G4153-JW'}, 'value': 'G4153-JW'}},
                   {'project_code': 'R1-AA', 'project_option': None}]}}
    out = sb._intake_fix_codes(d)['designation']
    assert out['project_code'] == 'G4153-YM'
    assert out['project_option']['value'] == 'G4153-YM' and out['project_option']['text']['text'].startswith('G4153-YM')
    assert out['splits'][0]['project_code'] == 'G4153-YM' and out['splits'][0]['project_option']['value'] == 'G4153-YM'
    assert out['splits'][1]['project_code'] == 'R1-AA'
    assert sb._cur_code('-') == '-' and sb._cur_code('') == '' and sb._cur_code(' G4153-JW ') == 'G4153-YM'
