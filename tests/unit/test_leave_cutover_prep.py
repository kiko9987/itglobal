"""휴가 전환 준비(2026-11-18) — 플렉스 가져오기 스크립트의 잔여 대조·앞날 휴가 옮기기.

플렉스는 신청이 잡힐 때 휴가 날짜의 달 몫에서 잔여를 뺀다(11-18 월별 현황으로 확인). 그래서 받은 달 안의 휴가는
SNAPSHOT 잔여에서 이미 빠져 있어 'flex:' 줄을 올리고 거기서 떼고, 다음 달 이후 휴가는 평소처럼 뗀다.
"""
import sys
from datetime import date

import pytest

from dashboard.services import leave as lv

sys.path.insert(0, 'scripts')
import leave_import_flex as imp  # noqa: E402

KTH = 'U0CFWU7KW8Y'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(lv, '_db', lambda: d)
    monkeypatch.setattr(lv, '_table_ready', False)
    monkeypatch.delenv('LEAVE_APPROVER_IDS', raising=False)
    monkeypatch.setattr(imp, 'SNAPSHOT', [('김태현', KTH, '2022-05-23', 7)])
    monkeypatch.setattr(imp, 'FLEX_DIR', tmp_path)          # 진짜 플렉스 폴더(촉진 기록 JSON)를 읽지 않게
    with d._get_connection() as conn:
        conn.execute('''CREATE TABLE IF NOT EXISTS gw_leave_history (flex_event_id TEXT PRIMARY KEY, person_id INTEGER,
            name TEXT NOT NULL, type_label TEXT, start TEXT NOT NULL, "end" TEXT, days REAL, unit TEXT, status TEXT,
            reason TEXT, approval TEXT, comment TEXT, requested_at TEXT)''')
        rows = [
            ('f1', '연차', '2026-12-24', '2026-12-24', 1, 'DAY', '승인'),          # 받은 달(12월) 안 → 잔여에서 이미 빠짐
            ('f2', '연차', '2027-01-15', '2027-01-15', 0.5, 'HALF_DAY_AM', '승인'),  # 다음 달 이후 → 평소처럼 뗌
            ('f3', '연차', '2026-12-28', '2026-12-28', 1, 'DAY', '승인 대기'),     # 플렉스에서 먼저 끝낼 것
            ('f4', '군소집훈련', '2026-12-10', '2026-12-10', 1, 'DAY', '승인'),    # 연차 아님 → 손으로
            ('f5', '연차', '2026-12-30', '2026-12-30', 1, 'DAY', '취소'),          # 옮길 것 없음
            ('f6', '연차', '2026-11-20', '2026-11-20', 1, 'DAY', '승인'),          # 받은 날 전 날짜 → 대상 아님
        ]
        for fid, t, s, e, n, u, st in rows:
            conn.execute('INSERT INTO gw_leave_history VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)',
                         (fid, 1, '김태현', t, s, e, n, u, st, '', '', '', '2026-11-20 10:00'))
        conn.commit()
    return d


def _run(*args):
    return imp.main(['--summary', 'none.xlsx', '--as-of', '2026-12-02', *args])


def test_covered_is_by_month_of_as_of():
    assert imp.covered('2026-12-31', date(2026, 12, 2))
    assert not imp.covered('2027-01-01', date(2026, 12, 2))
    assert not imp.covered('2026-12-01', date(2026, 11, 10))      # 11월에 받은 잔여엔 12월 휴가가 안 빠져 있음


def test_compare_left_blank_is_zero():
    snap = [('조성헌', 'U', '2025-04-22', 0), ('김태현', KTH, '2022-05-23', 9), ('없는사람', 'X', '2026-01-01', 1)]
    assert imp.compare_left(snap, {'조성헌': None, '김태현': 7.0}) == [('김태현', 9.0, 7.0), ('없는사람', 1.0, None)]


def test_summary_left_reads_any_month_column(tmp_path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '구성원'
    ws.append(['이름', '상태', '입사일', '연차 잔여(26-12)'])
    ws.append(['김태현', '재직', '2022-05-23', 7])
    ws.append(['조성헌', '재직', '2025-04-22', None])
    ws.append(['퇴사자', '퇴사', '2020-01-01', 3])
    p = tmp_path / 's.xlsx'
    wb.save(p)
    col, left = imp.summary_left(str(p))
    assert col == '연차 잔여(26-12)' and left == {'김태현': 7.0, '조성헌': None}


def test_preview_writes_nothing(db, capsys):
    assert _run() == 0
    assert lv.staff_by_name('김태현') is None
    out = capsys.readouterr().out
    assert '2026-12-24' in out and '플렉스 잔여에서 이미 빠짐' in out and '다음 달 이후' in out
    assert '✋' in out and '승인 대기' in out and '군소집훈련' in out


def test_apply_moves_future_leave_and_keeps_balance(db):
    assert _run('--apply') == 0
    reqs = lv.requests_of('김태현')
    assert sorted((r['start'], r['ltype'], r['days'], r['status'], r['reason']) for r in reqs) == [
        ('2026-12-24', 'annual', 1.0, 'approved', 'flex:f1'),
        ('2027-01-15', 'half_am', 0.5, 'approved', 'flex:f2'),
    ]
    g = [x for x in lv.grants('김태현') if x['source'] == 'flex:김태현'][0]
    assert g['days'] == 8.0                        # 7 + 12월 휴가 1(잔여에서 이미 빠져 있던 것)
    # 플렉스 잔여 7 그대로 + 1월 반차 0.5 만 새로 빠짐
    assert lv.balance('김태현', date(2026, 12, 3)) == 6.5
    # 겹침 검사가 이제 플렉스 휴가를 본다(해지 뒤 같은 날 또 신청 못 함)
    assert lv.overlaps('김태현', date(2026, 12, 24), date(2026, 12, 24))
    # 승인 단계 기록(DM·공지 없음)
    assert [s['step'] for s in lv.steps(reqs[0]['id'])] == ['request', 'approve_1', 'approve_2']


def test_apply_twice_is_idempotent(db):
    assert _run('--apply') == 0
    assert _run('--apply') == 0
    assert len(lv.requests_of('김태현')) == 2
    g = [x for x in lv.grants('김태현') if x['source'] == 'flex:김태현'][0]
    assert g['days'] == 8.0 and g['used'] == 1.5
    assert lv.balance('김태현', date(2026, 12, 3)) == 6.5


def test_apply_stops_when_summary_differs(db, tmp_path, capsys):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '구성원'
    ws.append(['이름', '상태', '연차 잔여(26-12)'])
    ws.append(['김태현', '재직', 9])
    p = tmp_path / 's.xlsx'
    wb.save(p)
    assert imp.main(['--summary', str(p), '--as-of', '2026-12-02', '--apply']) == 1
    assert lv.staff_by_name('김태현') is None              # 아무것도 안 씀
    assert 'SNAPSHOT 7 / 엑셀 9' in capsys.readouterr().out
    assert imp.main(['--summary', str(p), '--as-of', '2026-12-02', '--apply', '--allow-diff']) == 0
    assert lv.staff_by_name('김태현') is not None


def test_cancelled_in_flex_after_move_is_flagged(db, capsys):
    assert _run('--apply') == 0
    with db._get_connection() as conn:
        conn.execute("UPDATE gw_leave_history SET status='취소 승인' WHERE flex_event_id='f1'")
        conn.commit()
    capsys.readouterr()
    assert _run() == 0
    assert '플렉스에선 지금 취소 승인' in capsys.readouterr().out



def test_apply_imports_flex_promotions(db, tmp_path, capsys):
    import json
    (tmp_path / '원본').mkdir()
    (tmp_path / '원본' / 'flex_promotion_20261202.json').write_text(json.dumps({'users': {'U1': '김태현'}, 'docs': [
        {'u': 'U1', 'task': 'tk', 'to': '2027-05-22', 'boostedAt': '2026-11-22T23:00:00', 'days': 7,
         'status': 'SHOULD_USER_WRITE', 'userBy': '2026-12-03', 'mgrBy': '2027-03-23', 'planned': [], 'canceled': False}]},
        ensure_ascii=False), encoding='utf-8')
    assert _run('--apply') == 0
    out = capsys.readouterr().out
    assert '④ 연차 촉진 기록 ← flex_promotion_20261202.json' in out and '김태현 ~2027-05-22 촉구 2026-11-23' in out
    from dashboard.services import leave_promo as lp
    recs = lp.records('김태현')
    assert len(recs) == 1 and recs[0]['source'] == 'flex:tk' and recs[0]['noticed_at'].startswith('2026-11-23 08:00')
    assert recs[0]['status'] == lp.NOTICED                                          # 모르는 플렉스 상태 = 진행 중
    assert all(c['name'] != '김태현' for c in lp.candidates(lv._d('2026-12-02')))   # 같은 촉구를 또 안 보냄


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
