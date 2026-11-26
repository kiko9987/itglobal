"""차량 유류비·하이패스 ← SB 법인카드 월 지출결의서(2026-11-26). 이름·카드 번호는 모두 가짜."""
import openpyxl
import pytest

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import vehicle_card_costs as vcc
from dashboard.services import vehicle_manage as vm


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, va, vm, hr):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    gw._invalidate()
    yield d


def _book(path, sheets):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        ws.append([None, '2026년 법인카드 지출결의서'])
        ws.append([None, '종류', 'NO', '카드번호', '차량번호', '차량종류', '사용액'])
        for r in rows:
            ws.append([None] + list(r))
    wb.save(path)
    return str(path)


SHEET_10 = [
    ('유류비', 1, '1111', '231허6854', '레이', 392000),
    (None, 2, '2222', '박민우', None, 416000),             # 사람 이름 = MW 스파크
    (None, 3, '3333', '81루 4904', '포터', 317162),          # 끝 4자리 오타도
    (None, 4, '4444', '퇴사자갑', None, 50000),              # 못 찾음 → 안 넣고 남김
    (None, 5, '5555', '03모8021', '쏘렌토', 0),              # 0원은 건너뜀
    ('유류비 합계', None, None, None, None, 1175162),
    ('하이패스', 1, 'B000001', '231허 6854', '레이', 56650),
    (None, 2, 'B000002', '143고9687', '카니발', 600),
    ('하이패스 합계', None, None, None, None, 57250),
    ('종류', 'NO', '카드번호', '부서', '사용자', '사용액'),
    ('개인지급', 1, '9999', '시공팀', '가나다', 1161811),     # 차량 아님 — 읽지 않음
]


def test_parse_and_save_monthly_rows(db, tmp_path):
    path = _book(tmp_path / '2026년 법인카드 월 지출결의서(SB).xlsx', {'2610': SHEET_10, '메모': []})
    res = vcc.run(path=path)
    assert res['status'] == 'ok' and res['added'] == 5 and res['changed'] == 0
    assert [(u['label'], u['amount']) for u in res['unmatched']] == [('퇴사자갑', 50000)]
    rows = {(r['plate'], r['kind']): r for r in vm.costs(limit=50)}
    assert rows[('231허6854', '유류')]['amount'] == 392000 and rows[('231허6854', '유류')]['day'] == '2026-10-31'
    assert rows[('231허6854', '유류')]['item'] == '10월 유류비 (월 합계)' and rows[('231허6854', '유류')]['source'] == 'card'
    assert rows[('53우1783', '유류')]['amount'] == 416000
    assert rows[('817루4904', '유류')]['amount'] == 317162
    assert rows[('231허6854', '통행료')]['amount'] == 56650 and rows[('143고9687', '통행료')]['amount'] == 600
    assert not any(r['plate'] == '03모8021' for r in vm.costs(limit=50))


def test_rerun_updates_changed_amount_only(db, tmp_path):
    path = _book(tmp_path / 'a.xlsx', {'2610': SHEET_10})
    vcc.run(path=path)
    changed = [list(r) for r in SHEET_10]
    changed[0][5] = 400000                                   # SB 가 금액을 고침
    path2 = _book(tmp_path / 'b.xlsx', {'2610': changed})
    res = vcc.run(path=path2)
    assert res['added'] == 0 and res['changed'] == 1
    r = [x for x in vm.costs(limit=50) if x['plate'] == '231허6854' and x['kind'] == '유류']
    assert len(r) == 1 and r[0]['amount'] == 400000


def test_month_of_sheet_names():
    assert vcc._month_of('2602').isoformat() == '2026-02-28'
    assert vcc._month_of('2412').isoformat() == '2024-12-31'
    assert vcc._month_of('메모') is None and vcc._month_of('2613') is None
