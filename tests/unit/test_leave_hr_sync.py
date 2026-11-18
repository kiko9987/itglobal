# -*- coding: utf-8 -*-
"""사람 정본 = 인사카드 (2026-11-16 상태 머신 점검 묶음 G).

① 휴가 장부는 인사카드를 따라간다 — 퇴사면 연차 발생·슬랙 신청에서 빠지고, 입사일은 인사카드 값.
② 인사카드를 '퇴사'로 바꾸면 휴가 명단·PM 로그인도 같이, 권한 명단에 남아 있으면 관리자 알림.
③ 지난 날짜 신청은 오늘까지 살아 있는 연차에서만(소멸 연차 부활 막기).
④ 휴가 기능을 켜는 날 한 사람에게 연차가 몰려 생기면 넣지 않고 보류.
⑤ 켜기 전 점검 스크립트(읽기 전용).
"""
import sqlite3
import sys
from datetime import date

sys.path.insert(0, '.')

import pytest
from cryptography.fernet import Fernet

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import leave as lv

ADMIN = 'kiko@itg-aircon.com'
MW, MW_MAIL = 'U0BJ9M1MRJ4', 'mw@itg-aircon.com'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, lv):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.delenv('LEAVE_APPROVER_IDS', raising=False)
    gw._invalidate()
    yield d
    gw._invalidate()


def _person(**kw):
    vals = {'name': '박민우', 'status': '재직', 'hire_date': '2022-02-03', 'work_email': MW_MAIL,
            'login_email': MW_MAIL, 'slack_id': MW}
    vals.update(kw)
    return hr.create_person(vals, ADMIN)


def _staff(hire='2022-02-03', accrue_from='2026-11-11'):
    lv.upsert_staff('박민우', hire, MW, accrue_from=accrue_from)


# ── ① 읽을 때 인사카드 기준 ──

def test_staff_follows_hr_card(db):
    _staff()
    p = _person()
    assert [s['name'] for s in lv.staff()] == ['박민우'] and lv.staff_by_slack(MW)
    with db._get_connection() as conn:                                  # 플렉스 재이관 등 hook 밖에서 퇴사
        conn.execute("UPDATE hr_people SET status='퇴사' WHERE id=?", (p['id'],))
        conn.commit()
    assert lv.staff() == [] and lv.staff_by_slack(MW) is None            # 연차 발생·슬랙 신청에서 빠짐
    assert [s['name'] for s in lv.staff(active_only=False)] == ['박민우']  # 장부 원본은 그대로


def test_hire_date_from_hr_card(db):
    _staff(hire='2022-02-03')
    _person(hire_date='2021-02-03')
    assert lv.staff()[0]['hire_date'] == '2021-02-03'


def test_rehire_card_wins_over_old_resigned_card(db):
    _staff()
    old = _person(status='퇴사')
    assert lv.staff() == []
    _person(name='박민우', status='재직')                                   # 같은 슬랙 ID 로 재입사
    assert [s['name'] for s in lv.staff()] == ['박민우'] and old


def test_no_hr_table_keeps_old_behaviour(db):
    _staff()
    assert [s['name'] for s in lv.staff()] == ['박민우']                  # 인사카드 표 없음 → 명단 그대로


# ── ② 인사카드 '퇴사' → 휴가 명단·PM 로그인·권한 명단 알림 ──

def test_resign_on_hr_card_propagates(db, caplog, monkeypatch):
    _staff()
    monkeypatch.setattr(db, 'update_user_status', lambda e, s: (True, 'ok'))
    p = _person()
    with caplog.at_level('ERROR'):
        hr.update_person(p['id'], {'status': '퇴사', 'resign_date': '2026-11-30'}, ADMIN)
    with db._get_connection() as conn:
        assert conn.execute('SELECT active FROM leave_staff WHERE slack_id=?', (MW,)).fetchone()[0] == 0
    assert not any('권한' in r.getMessage() for r in caplog.records)     # 권한 명단에 없는 사람
    hr.update_person(p['id'], {'status': '재직'}, ADMIN)                  # 잘못 눌러 되돌림
    with db._get_connection() as conn:
        assert conn.execute('SELECT active FROM leave_staff WHERE slack_id=?', (MW,)).fetchone()[0] == 1


def test_resign_marks_pm_user(db, monkeypatch):
    calls = []
    monkeypatch.setattr(db, 'update_user_status', lambda email, st: calls.append((email, st)) or (True, 'ok'))
    p = _person(work_email='mw.work@itg-aircon.com')
    hr.update_person(p['id'], {'status': '퇴사'}, ADMIN)
    assert sorted(calls) == [('mw.work@itg-aircon.com', '퇴사'), (MW_MAIL, '퇴사')]
    calls.clear()
    hr.update_person(p['id'], {'department': '영업'}, ADMIN)              # 다른 칸 수정은 그대로
    assert calls == []


def test_resign_of_privileged_person_alerts(db, caplog, monkeypatch):
    monkeypatch.setattr(db, 'update_user_status', lambda e, s: (True, 'ok'))
    sb = 'U0BHC2JV7U5'                                                     # 정산 승인 기본 명단(SB)
    monkeypatch.delenv('SLACK_SETTLEMENT_CHECKER_ID', raising=False)
    p = _person(name='황샛별', slack_id=sb, work_email='sb@itg-aircon.com', login_email='sb@itg-aircon.com')
    with caplog.at_level('ERROR'):
        hr.update_person(p['id'], {'status': '퇴사'}, ADMIN)
    msgs = [r.getMessage() for r in caplog.records if '명단에 있어요' in r.getMessage()]
    assert msgs and '정산 승인' in msgs[0]


def test_hire_date_change_copies_to_leave_staff(db):
    _staff()
    p = _person()
    hr.update_person(p['id'], {'hire_date': '2021-05-01'}, ADMIN)
    with db._get_connection() as conn:
        assert conn.execute('SELECT hire_date FROM leave_staff WHERE slack_id=?', (MW,)).fetchone()[0] == '2021-05-01'


# ── ③ 지난 날짜 신청 ──

def test_past_dated_request_cannot_use_expired_grant(db, monkeypatch):
    _staff()
    lv.add_grant('박민우', 2, date(2026, 1, 1), date(2026, 10, 31), 'old', '소멸된 연차')
    lv.add_grant('박민우', 1, date(2026, 10, 1), date(2027, 9, 30), 'live', '살아 있는 연차')
    oct20 = date(2026, 10, 20)                                            # 화요일
    real_date = lv.date

    class _D(real_date):
        @classmethod
        def today(cls):
            return real_date(2026, 11, 16)
    monkeypatch.setattr(lv, 'date', _D)
    errs = lv.validate('박민우', lv.ANNUAL, '', oct20, oct20 + lv.timedelta(days=1))   # 이틀
    assert 'start' in errs and '소멸' in errs['start']                    # 첫날 기준이면 3일 → 통과했었다
    assert lv.validate('박민우', lv.ANNUAL, '', oct20, oct20) == {}       # 하루는 살아 있는 연차로
    r = lv.create('박민우', MW, MW, '박민우', lv.ANNUAL, '', oct20, oct20)
    with db._get_connection() as conn:
        gid = conn.execute('SELECT grant_id FROM leave_uses WHERE request_id=?', (r['id'],)).fetchone()[0]
        assert conn.execute('SELECT source FROM leave_grants WHERE id=?', (gid,)).fetchone()[0] == 'live'


def test_future_request_unchanged(db):
    _staff()
    lv.add_grant('박민우', 3, date(2026, 11, 1), date(2027, 10, 31), 'g', '연차')
    d = date(2026, 12, 1)
    assert lv.validate('박민우', lv.ANNUAL, '', d, d + lv.timedelta(days=1)) == {}


# ── ④ 켜는 날 몰려 생기는 연차 보류 ──

def test_accrue_holds_burst_and_allows_backfill(db):
    lv.upsert_staff('김신입', '2026-01-05', 'UNEW')                        # accrue_from 비어 있음 → 입사일부터
    _staff()                                                              # 이관된 사람(accrue_from 11-11)
    held = []
    made = lv.accrue(date(2026, 11, 16), held=held)
    assert [h['name'] for h in held] == ['김신입'] and held[0]['count'] == 10
    assert all(g['name'] != '김신입' for g in made)
    assert lv.grants('김신입') == []
    made = lv.accrue(date(2026, 11, 16), allow_backfill=True)
    assert len([g for g in made if g['name'] == '김신입']) == 10
    held = []
    lv.accrue(date(2026, 12, 6), held=held)                               # 다음 달 1줄은 평소처럼
    assert held == [] and len(lv.grants('김신입')) == 11


def test_run_accrual_alerts_on_hold(db, monkeypatch, caplog):
    from dashboard.blueprints import slack_leave as sl
    monkeypatch.setattr(lv, 'enabled', lambda: True)
    monkeypatch.setattr(sl, '_client', lambda: None)
    lv.upsert_staff('김신입', '2026-01-05', 'UNEW')
    with caplog.at_level('ERROR'):
        assert sl.run_accrual(date(2026, 11, 16)) == 0
    assert any('보류' in r.getMessage() and '김신입' in r.getMessage() for r in caplog.records)


# ── ⑤ 켜기 전 점검 스크립트 ──

def test_cutover_check_reads_only(db, tmp_path):
    _staff()
    lv.upsert_staff('김신입', '2026-01-05', 'UNEW')
    _person()
    with db._get_connection() as conn:
        conn.execute('''CREATE TABLE IF NOT EXISTS gw_leave_history (flex_event_id TEXT PRIMARY KEY, person_id INTEGER,
            name TEXT NOT NULL, type_label TEXT, start TEXT NOT NULL, "end" TEXT, days REAL, unit TEXT, status TEXT,
            reason TEXT, approval TEXT, comment TEXT, requested_at TEXT)''')
        conn.execute("INSERT INTO gw_leave_history VALUES ('e1', 1, '박민우', '연차', '2026-11-20', '2026-11-20', 1, "
                     "'DAY', '승인', '', '', '', '2026-11-05')")
        # 잔여를 읽은 날(11-10) 뒤에 신청된 것 — 그 잔여엔 아직 안 빠짐(11-18 R3: 날짜가 아니라 신청 시각 기준)
        conn.execute("INSERT INTO gw_leave_history VALUES ('e2', 1, '박민우', '연차', '2026-11-12', '2026-11-12', 1, "
                     "'DAY', '승인', '', '', '', '2026-11-11 09:00')")
        conn.commit()
    sys.path.insert(0, 'scripts')
    import leave_cutover_check as cc
    ro = cc._ro(tmp_path / 'users.db')
    res = cc.check(ro, date(2026, 11, 16), date(2026, 11, 10))
    assert ('김신입', '인사카드에 같은 슬랙 ID 없음') in res['mismatch']
    assert res['no_accrue_from'] == ['김신입']
    assert ('김신입', 10, 10, True) in res['accrue']
    # ④ 신청 시각이 11-10 뒤인 e2 만(e1 은 11-05 신청이라 11-10 잔여에 이미 빠짐)
    assert [(r['name'], r['start']) for r in res['flex_after']] == [('박민우', '2026-11-12')]
    # ⑤ 오늘(11-16) 뒤 날짜로 확정된 e1 — 아직 장부로 안 옮김
    assert [(r['flex_event_id'], r['moved']) for r in res['flex_future']] == [('e1', False)]
    assert res['requests'] == {}
    with pytest.raises(sqlite3.OperationalError):                         # 읽기 전용 연결
        ro.execute("UPDATE leave_staff SET active=0")


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
