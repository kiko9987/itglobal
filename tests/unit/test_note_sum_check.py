# -*- coding: utf-8 -*-
"""수금 칸 값 ≠ 메모 입금 합 정합성 점검 (2026-11-16 상태 머신 점검)."""
import sys

sys.path.insert(0, '.')

import pytest

from dashboard.services import invariant_checks as ic
from dashboard.services.payment_note_sum import cell_note_mismatch, note_amount_total

F2L = {'계약금': 'V', '중도금': 'W', '잔금': 'X'}


@pytest.mark.parametrize('note,expect', [
    ('하나,11/05, 04:43\n255******31304\n입금3,573,029원\n신한14322878\n카드 실결제 3,625,600원\n수수료 52,571원', 3625600),
    ('2026/08/19 09:23\n입금 3,595,000원\n다모아백화점\n출금 -1,529,000원 (2대 반품 환불, 순액 2,066,000)', 2066000),
    ('2026/09/04\n입금 3,000,000원\n라은정\n· 통합 3,800,000원 분할\n\n하나,09/08, 15:00\n255******31304\n출금3,000,000원\n'
     '라은정\n\n2026/09/08\n입금 3,000,000원\n라은정', 3000000),
    ('일시 04/30, 08:45\n입금 1,805,549원\n계좌번호 255******31304', 1805549),
    ('하나,10/28, 14:33\n입금500,000원\n심민경\n\n하나,11/03, 14:59\n출금500,000원\n심민경', 0),
    ('입금일: \n입금자: ', None),                                            # 옛 빈 틀
    ('⚠ 입금 지정 이상징후 이력 입금 9,999원\n2026/11/01\n입금 100,000원', 100000),
    ('◆ 부담자 A\n2026/11/01\n입금 100,000원\n\n2026-11-02 R>G 매출이동 (하나 → 기업 법인 간 이체 · 처리 SB)', 100000),
    ('', None), (None, None),
])
def test_note_amount_total(note, expect):
    assert note_amount_total(note) == expect


def test_cell_note_mismatch():
    assert cell_note_mismatch(3625600, '입금3,573,029원\n카드 실결제 3,625,600원') is None
    assert cell_note_mismatch('₩3,000,000', '입금 2,000,000원') == 2000000
    assert cell_note_mismatch(0, '입금일: ') is None


def _rows(v, note, code='G4300-JW'):
    return [(4301, code, {'계약금': (v, note), '중도금': ('', ''), '잔금': ('', '')})]


def test_mismatch_rows_and_check_only_new():
    m = ic.note_sum_mismatches(_rows(1000000, '입금 500,000원'), F2L)
    assert len(m) == 1 and m[0][1] == 'G4300-JW' and 'V4301' in m[0][2] and '500,000' in m[0][2]
    ctx = {'note_sum_mismatches': m, 'known_note_sum': set()}
    assert [v['title'] for v in ic.check_note_sum(ctx)] == ['G4300-JW']
    ctx['known_note_sum'] = {m[0][0]}
    assert ic.check_note_sum(ctx) == []
    assert ic.check_note_sum in ic.INVARIANTS


class _RC:
    def __init__(self):
        self.sets = {}

    def exists(self, k):
        return k in self.sets

    def sadd(self, k, *v):
        self.sets.setdefault(k, set()).update(v)

    def smembers(self, k):
        return self.sets.get(k, set())

    def scan_iter(self, *a, **k):
        return iter(())


def _ctx(monkeypatch, reads, rc):
    import dashboard.utils.redis_client as rcm
    import dashboard.services.project_service as ps
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': rc})())
    monkeypatch.setattr(ps, 'get_project_records', lambda *a, **k: [])
    monkeypatch.setattr(ic, '_read_sheet_cells', lambda: [])
    monkeypatch.setattr(ic, '_read_sheet_formulas', lambda: [])
    monkeypatch.setattr(ic, 'main_field_to_letter', lambda: F2L, raising=False)
    import dashboard.utils.google_sheets as gs
    monkeypatch.setattr(gs, 'main_field_to_letter', lambda: F2L)
    it = iter(reads)
    monkeypatch.setattr(ic, '_read_stage_cells_with_notes', lambda: next(it))
    monkeypatch.setattr(ic.time, 'sleep', lambda s: None)
    return ic.build_context()


def test_first_run_seeds_known(monkeypatch):
    rc = _RC()
    ctx = _ctx(monkeypatch, [_rows(1000000, '입금 500,000원')], rc)
    assert ic.check_note_sum(ctx) == []                                  # 처음 있던 건 기준
    assert any(k.startswith('notesum:') for k in rc.sets[ic._NOTE_KNOWN_KEY])


def test_new_mismatch_rechecked(monkeypatch):
    rc = _RC()
    rc.sets[ic._NOTE_KNOWN_KEY] = {'__seeded__'}
    # 첫 읽기엔 어긋남, 다시 읽으니 맞음(입금 기록 도중) → 알림 없음
    ctx = _ctx(monkeypatch, [_rows(1000000, '입금 500,000원'), _rows(1000000, '입금 1,000,000원')], rc)
    assert ic.check_note_sum(ctx) == []
    # 다시 읽어도 그대로 → 알림
    ctx = _ctx(monkeypatch, [_rows(1000000, '입금 500,000원'), _rows(1000000, '입금 500,000원')], rc)
    assert len(ic.check_note_sum(ctx)) == 1


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
