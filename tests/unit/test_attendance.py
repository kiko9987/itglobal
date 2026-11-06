"""시공팀 출퇴근·조출/연장 정산 (2026-11-06) — SB 정산표(2026-08·10)와 같은 값이 나오는지."""
from datetime import date

import pytest

from dashboard.services import attendance as at


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    from dashboard.services import leave as lv
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(at, '_db', lambda: d)
    monkeypatch.setattr(at, '_table_ready', False)
    monkeypatch.setattr(lv, '_db', lambda: d)
    monkeypatch.setattr(lv, '_table_ready', False)
    for k in ('ATTENDANCE_STAFF', 'ATTENDANCE_HOURLY_WAGE', 'ATTENDANCE_ENABLED', 'ATTENDANCE_MANAGER_IDS'):
        monkeypatch.delenv(k, raising=False)
    return d


def m(hm):
    h, mm = hm.split(':')
    return int(h) * 60 + int(mm)


@pytest.mark.parametrize('cin,cout,early,enight,late,lnight', [
    ('06:12', '19:23', 90, 0, 120, 0),     # 8/4 김태현: 78분→90, 113분→120
    ('07:21', '21:35', 0, 0, 240, 0),      # 8/5: 9분→0, 245분→240
    ('07:02', '22:06', 30, 0, 270, 0),     # 8/8: 28→30, 연장 270(22시 상한)·야간 6분→0
    ('05:31', '21:19', 90, 30, 240, 0),    # 8/10: 06시 이전 29분→야간 30, 조출 90
    ('05:03', '16:22', 90, 60, 0, 0),      # 10/21: 야간 57→60(빨간칸) + 90
    ('07:30', '17:30', 0, 0, 0, 0),
])
def test_calc_matches_sb_sheet(cin, cout, early, enight, late, lnight):
    c = at.calc(m(cin), m(cout))
    assert (c['early'], c['early_night'], c['late'], c['late_night']) == (early, enight, late, lnight)


def test_calc_late_night_and_meal():
    c = at.calc(m('07:00'), m('23:40'))        # 22시 이후 100분 → 야간 90
    assert c['late'] == 270 and c['late_night'] == 90 and c['meal']
    assert not at.calc(m('07:00'), m('21:40'))['meal']    # 250분 → 240
    assert at.r30(14) == 0 and at.r30(15) == 30 and at.r30(44) == 30 and at.r30(45) == 60


SAMPLE = ('20261001\t074020\t6703129000607\t김태현\t1\t아이티글로벌\t기사\t\t14\t일반근무\t0\n'
          '20261001\t173233\t6703129000607\t김태현\t4\t아이티글로벌\t기사\t\t14\t일반근무\t0\n'
          '20261026\t063300\t0103588275485\t김종연\t1\t아이티글로벌\t기사\n'
          '20261027\t044800\t0103588275485\t김종연\t4\t아이티글로벌\t기사\n'
          '20261027\t044800\t0103588275485\t김종연\t1\t아이티글로벌\t기사\n'
          '20261027\t174700\t0103588275485\t김종연\t4\t아이티글로벌\t기사\n'
          '잘못된 줄\n')


def test_parse_and_ingest_once(db, tmp_path):
    rows = at.parse_export(SAMPLE)
    assert len(rows) == 6 and rows[0] == {'sabun': '6703129000607', 'name': '김태현',
                                          'ts': '2026-10-01 07:40:20', 'mode': 1}
    f = tmp_path / 'ERPExport_20261027.txt'
    f.write_bytes(SAMPLE.encode('cp949'))
    assert at.ingest_dir(str(tmp_path)) == 6
    assert at.ingest_dir(str(tmp_path)) == 0          # 매시 다시 읽어도 같은 줄은 한 번


def test_next_day_tag_counted_once(db):
    """8/26 김종연: 퇴근을 안 찍고 다음 날 04:48 태그가 퇴근·출근 둘로 잡힘 → 그날은 퇴근 없음(확인 필요)."""
    at.add_events(at.parse_export(SAMPLE))
    r26 = at.day_record('김종연', '2026-10-26')
    assert r26['in'] == m('06:33') and r26['out'] is None and at.needs_check(r26) == '퇴근 기록 없음'
    r27 = at.day_record('김종연', '2026-10-27')
    assert r27['in'] == m('04:48') and r27['out'] == m('17:47')


def test_report_fills_missing_and_summary(db, monkeypatch):
    monkeypatch.setenv('ATTENDANCE_HOURLY_WAGE', '15165')
    at.add_events(at.parse_export(SAMPLE))
    r = at.add_report('김종연', '2026-10-26', '', '19:10', '현장', 'U1', '김종연')
    assert at.day_record('김종연', '2026-10-26')['out'] is None        # 승인 전엔 안 들어감
    assert at.decide_report(r['id'], True, '황샛별')['status'] == at.APPROVED
    assert at.decide_report(r['id'], True, '황샛별') is None           # 두 번 눌러도 한 번
    rec = at.day_record('김종연', '2026-10-26')
    assert rec['out'] == m('19:10') and rec['out_src'] == 'report'
    s = at.summary('2026-10', ('김태현', '김종연'))
    k = s['김종연']
    # 26일: 조출 57→60 · 연장 100→90 / 27일: 야간 72→60 · 조출 90 · 연장 17→30
    assert k['early_total'] == 60 + 60 + 90 and k['late_total'] == 90 + 30 and k['night'] == 60
    assert k['pay_normal'] == int((60 + 90 + 90 + 30) / 60 * 15165 * 1.5 + 0.5)
    assert k['pay_night'] == int(1 * 15165 * 2 + 0.5)
    assert not k['checks']
    assert '확인 필요' not in at.summary_text('2026-10', s)
    assert any(d['reported'] for d in k['days'])


def test_late_arrival_overtime_flagged(db):
    """10/17 김태현 10:02~22:06 — 연장은 계산하되 사람이 정하게 '확인 필요'."""
    at.add_events(at.parse_export('20261017\t100200\tS1\t김태현\t1\n20261017\t220600\tS1\t김태현\t4\n'))
    v = at.summary('2026-10', ('김태현',))['김태현']
    assert v['late_total'] == 270 and [c['why'] for c in v['checks']] == ['늦게 출근한 날의 연장']
    assert at.missing_days('김태현', date(2026, 10, 18), back=1) == []    # 본인에겐 묻지 않는다


def test_sb_pay_formula():
    """10월 김태현: 일반 30.5h·야간 1.0h → 693,799 + 30,330 (SB 표)."""
    assert int(30.5 * 15165 * 1.5 + 0.5) == 693799 and int(1.0 * 15165 * 2 + 0.5) == 30330


def test_missing_days_skip_weekend_and_reported(db):
    at.add_events(at.parse_export(SAMPLE))
    got = {x['day']: x['why'] for x in at.missing_days('김종연', date(2026, 10, 28), back=3)}
    assert got == {'2026-10-26': '퇴근 기록 없음'}         # 25일(일)은 기록 없어도 안 묻고, 27일은 다 찍힘
    assert at.mark_asked('김종연', '2026-10-26') and not at.mark_asked('김종연', '2026-10-26')
    at.add_report('김종연', '2026-10-26', '', '19:10', '', 'U1', '김종연')
    assert at.missing_days('김종연', date(2026, 10, 28), back=3) == []   # 보고 대기 중이면 다시 안 묻는다


def test_manager_ids_test_mode(db, monkeypatch):
    assert at.manager_ids() == ('U04UL2ZLJAX',)               # 켜기 전엔 KiKO 만
    monkeypatch.setenv('ATTENDANCE_MANAGER_IDS', 'UA,UB')
    assert at.manager_ids() == ('UA', 'UB')
