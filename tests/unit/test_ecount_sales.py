"""이카운트 매출전표 목록 (보고서 모드, 2026-10-30) — 차액 계산·장부·[넣음] 표시."""
from datetime import datetime

import pytest

from dashboard.services import ecount_sales as ecs


def _rec(code, t1, t2=None, vat=True, confirmed='2026-10-01', **kw):
    r = {'프로젝트 코드': code, '총액 1': t1, '총액 2': t2 if t2 is not None else t1,
         '부가세': vat, '공사 확정': confirmed, '유입 구분': '거래처', '현장 주소': '성남 분당구 1'}
    r.update(kw)
    return r


def test_target_uses_total2_minus_total1():
    # 부가세 포함 금액으로 정한 공사 — 10% 반올림(345,454)이 아니라 총액 2 − 총액 1
    assert ecs.target_of(_rec('G0014-YG', 3454545, 3800000)) == (3454545, 345455)
    assert ecs.target_of(_rec('G4147-JW', 12362000, 12362000, vat=False)) == (12362000, 0)
    # 총액 2 가 아직 비면 10%
    assert ecs.target_of(_rec('G9', 1000000, 0)) == (1000000, 100000)
    assert ecs.target_of(_rec('G9', 0, 0)) == (0, 0)


def test_compute_pending_known_cases_from_10_29():
    """10-29 대조에서 찾은 빠진 차액 5건이 그대로 잡힌다."""
    recs = [
        _rec('G4152-JW', 11242000, 11242000, vat=False),
        _rec('R4163-TH', 2500000, 2500000, vat=False, confirmed='2026-10-03'),
        _rec('G4208-MJ', 2800000, 3080000, confirmed='2026-10-21'),
        _rec('R4180-YM', 3100000, 3410000, confirmed='2026-10-14', **{'유입 구분': '채널톡'}),
        _rec('G4076-YG', 33800000, 37180000, confirmed='2026-09-04'),   # 맞음 → 안 뜸
        _rec('G0001-XX', 500000, 550000, confirmed=''),                 # 확정 전 → 안 뜸
    ]
    ledger = {
        'G4152-JW': (11242000, 1124200),
        'R4163-TH': (2500000, 250000),
        'G4208-MJ': (2800000, 0),
        'G4076-YG': (33800000, 3380000),
    }
    items = ecs.compute_pending(recs, ledger, {'G4152-JW': '2026-10-26', 'R4163-TH': '2026-10-07',
                                               'G4208-MJ': '2026-10-22'})
    got = {e['code']: (e['kind'], e['supply'], e['vat'], e['date']) for e in items}
    assert got == {
        'G4152-JW': ('부가세 변경', 0, -1124200, '2026-10-26'),
        'R4163-TH': ('부가세 변경', 0, -250000, '2026-10-07'),
        'G4208-MJ': ('부가세 변경', 0, 280000, '2026-10-22'),
        'R4180-YM': ('공사 확정', 3100000, 310000, '2026-10-14'),
    }
    # 날짜순
    assert [e['code'] for e in items] == ['R4163-TH', 'R4180-YM', 'G4208-MJ', 'G4152-JW']


def test_classify_amount_cancel_resume():
    assert ecs._classify(True, (100, 10), (200, 20)) == '금액 변경'
    assert ecs._classify(True, (100, 10), (0, 0)) == '공사 취소'
    assert ecs._classify(True, (0, 0), (100, 10)) == '공사 재개'
    assert ecs._classify(False, (0, 0), (100, 10)) == '공사 확정'
    items = ecs.compute_pending([_rec('R4181-JSH', 0, 0)], {'R4181-JSH': (2200000, 220000)})
    assert (items[0]['kind'], items[0]['supply'], items[0]['vat']) == ('공사 취소', -2200000, -220000)


def test_cancel_clears_confirmed_date():
    """취소는 공사 확정 날짜를 바로 지운다(총액 1=0 은 나중) — 날짜가 비면 장부 금액 전액 마이너스."""
    r = _rec('R4181-JSH', 2200000, 2420000, confirmed='', **{'수금 관련 특이사항': '공사 취소 - 고객변심'})
    items = ecs.compute_pending([r], {'R4181-JSH': (2200000, 220000)})
    assert (items[0]['kind'], items[0]['supply'], items[0]['vat']) == ('공사 취소', -2200000, -220000)
    assert '적요: 취소 사유 : 고객변심' in ecs.item_text(items[0])
    # 이미 취소를 적었거나(0) 장부에 없는 미확정 공사는 안 뜸
    assert ecs.compute_pending([r], {'R4181-JSH': (0, 0)}) == []
    assert ecs.compute_pending([r], {}) == []
    # 재개 — 날짜·금액 복원
    back = ecs.compute_pending([_rec('R4181-JSH', 2200000, 2420000)], {'R4181-JSH': (0, 0)})
    assert (back[0]['kind'], back[0]['supply'], back[0]['vat']) == ('공사 재개', 2200000, 220000)


def test_new_project_with_zero_amount_waits():
    """확정됐지만 금액이 아직 0이면 넣을 게 없다 — 금액이 들어오면 '공사 확정'으로 뜬다."""
    assert ecs.compute_pending([_rec('G5000-AA', 0, 0)], {}) == []


def test_blocks_and_done_mark():
    items = ecs.compute_pending([_rec('R4180-YM', 3100000, 3410000, confirmed='2026-10-14')], {})
    blocks = ecs.build_blocks(items)
    assert '넣을 것 1건' in blocks[0]['text']['text']
    sec = blocks[2]
    assert sec['accessory']['action_id'] == 'ecount_sales_done'
    t = sec['text']['text']
    assert '*R4180-YM* · 공사 확정 · 10/14' in t
    assert '공급가: 3,100,000원 · 부가세: 310,000원 · 합계: 3,410,000원' in t
    assert '적요: 성남 분당구 1' in t
    done = ecs.mark_done_blocks(blocks, items[0]['key'], 'SB', datetime(2026, 10, 30, 10, 5))
    assert 'accessory' not in done[2]
    assert ':white_check_mark: SB 넣음 10/30 10:05' in done[2]['text']['text']
    assert done[0] == blocks[0]


def test_dm_blocks_amount_then_sales():
    items = ecs.compute_pending([_rec('R4180-YM', 3100000, 3410000, confirmed='2026-10-14')], {})
    blocks = ecs.build_blocks(items, [{'summary': '`R4068-TH` 부가세 있음→없음', 'permalink': ''}])
    assert blocks[0]['text']['text'].startswith(':pencil2: *금액 반영 대기')
    assert blocks[1] == {'type': 'divider'}
    assert '이카운트 매출전표 — 넣을 것 1건' in blocks[2]['text']['text']
    assert blocks[4]['accessory']['action_id'] == 'ecount_sales_done'
    assert ecs.build_blocks([], []) == []


def test_block_cap():
    recs = [_rec(f'G{i:04d}-AA', 1000000, 1100000) for i in range(ecs.MAX_ITEMS + 3)]
    blocks = ecs.build_blocks(ecs.compute_pending(recs, {}))
    assert len(blocks) <= 50
    assert '나머지 3건' in blocks[-1]['elements'][0]['text']


def test_split_project_warns():
    e = ecs.compute_pending([_rec('G4148-JW', 16042000, 17046200,
                                  **{'잔금 부담자': '임대인 · 이희승 · 122-15-41208 · 6,887,100 · 계산서'})], {})
    assert e[0]['split'] is True
    assert '부담자별로 나눠' in ecs.item_text(e[0])
    assert '부담자별로' not in ecs.item_text(ecs.compute_pending([_rec('G1', 100, 110)], {})[0])


@pytest.fixture
def tmp_db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    db = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(ecs, '_db', lambda: db)
    monkeypatch.setattr(ecs, '_table_ready', False)
    return db


def test_ledger_record_once_and_seed(tmp_db):
    assert ecs.seed_baseline([('G1', 100, 10, ''), ('G2', 200, 0, '')], by='seed') == 2
    assert ecs.seed_baseline([('G1', 999, 99, '')], by='seed') == 0      # 이미 있는 공사는 건너뜀
    sent, target = (100, 10), (150, 15)
    key = ecs.entry_key('G1', sent, target)
    assert ecs.record_entry('G1', 50, 5, '금액 변경', key, by='slack:SB') is True
    assert ecs.record_entry('G1', 50, 5, '금액 변경', key, by='slack:SB') is False   # 두 번 눌림
    assert ecs.ledger_totals() == {'G1': (150, 15), 'G2': (200, 0)}
    assert ecs.compute_pending([_rec('G1', 150, 165)], ecs.ledger_totals()) == []
