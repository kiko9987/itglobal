"""회사 쉬는 날(2026-11-25) — 법정 공휴일 + 임시공휴일·회사 지정, 걸린 휴가 다시 세기. 이름은 모두 가짜."""
from datetime import date

import pytest

from dashboard.services import company_holidays as ch
from dashboard.services import leave as lv

JW, YG, KIM = 'U051ENK8AAY', 'U050XNG5VBR', 'U0TESTKIM01'
TODAY = date(2026, 11, 25)
BOSS = ('boss@example.com', '경영가')


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(lv, '_db', lambda: d)
    monkeypatch.setattr(lv, '_table_ready', False)
    monkeypatch.delenv('LEAVE_APPROVER_IDS', raising=False)
    ch.invalidate()
    yield d
    ch.invalidate()


def _approved(start, end=None, ltype=lv.ANNUAL, left=5.0):
    lv.upsert_staff('김가짜', '2022-02-03', KIM, accrue_from='2026-11-06')
    if not lv.grants('김가짜'):
        lv.add_grant('김가짜', left, date(2026, 11, 1), date(2027, 12, 31), 'test:김가짜', '시험')
    r = lv.create('김가짜', KIM, KIM, '김가짜', ltype, '', start, end)
    r = lv.transition(r['id'], 'approve', JW, '박정우')
    return lv.transition(r['id'], 'approve', YG, '박용구')


def _add(start, end='', name='창립기념일', kind=ch.KIND_COMPANY, yearly=False, **k):
    return ch.save('add', {'name': name, 'kind': kind, 'start': start, 'end': end, 'yearly': yearly},
                   *BOSS, today=TODAY, notify=False, **k)


def test_legal_days_and_labor_day(db):
    lg = ch.legal(2026)
    assert lg[date(2026, 5, 1)] == '노동절' and lg[date(2026, 12, 25)] == '크리스마스'      # 기독탄신일 → 크리스마스(11-25)
    assert lv.is_workday(date(2026, 12, 25)) is False and lv.is_workday(date(2026, 12, 28)) is True
    rows = ch.year_list(2026)['entries']
    seol = [e for e in rows if e['name'] == '설날']
    assert len(seol) == 1 and (seol[0]['start'], seol[0]['end'], seol[0]['days']) == ('2026-02-16', '2026-02-18', 3)
    assert all(not e['editable'] for e in rows)


def test_add_company_holiday_refunds_leave(db):
    r = _approved(date(2026, 12, 28), date(2026, 12, 30))            # 월~수 3일
    assert r['days'] == 3 and lv.balance('김가짜', date(2026, 12, 1)) == 2
    pv = ch.preview('add', {'name': '연말 휴무', 'kind': ch.KIND_COMPANY, 'start': '2026-12-29'}, today=TODAY)
    assert pv['count'] == 1 and pv['affected'][0]['new_days'] == 2 and pv['affected'][0]['delta'] == -1
    res = _add('2026-12-29', name='연말 휴무', expect=1)
    assert res['count'] == 1
    assert lv.get(r['id'])['days'] == 2 and lv.balance('김가짜', date(2026, 12, 1)) == 3
    assert lv.is_workday(date(2026, 12, 29)) is False and lv.holiday_name(date(2026, 12, 29)) == '연말 휴무'
    log = [s for s in lv.steps(r['id']) if s['step'] == 'holiday']
    assert log and '3일 → 2일' in log[0]['reason']
    # 새로 내는 휴가도 그날을 빼고 센다
    assert lv.count_days(lv.ANNUAL, date(2026, 12, 28), date(2026, 12, 30)) == 2


def test_delete_puts_days_back_or_refuses_when_short(db):
    r = _approved(date(2026, 12, 28), date(2026, 12, 30), left=3.0)
    hid = _add('2026-12-29', name='연말 휴무')['id']
    assert lv.get(r['id'])['days'] == 2 and lv.balance('김가짜', date(2026, 12, 1)) == 1
    other = lv.create('김가짜', KIM, KIM, '김가짜', lv.ANNUAL, '', date(2026, 12, 21), None)   # 남은 1일을 씀(대기)
    assert other and lv.balance('김가짜', date(2026, 12, 1)) == 0
    with pytest.raises(ValueError, match='연차가 모자라'):
        ch.save('delete', {}, *BOSS, hid=hid, today=TODAY, notify=False)
    assert ch.get(hid) and lv.get(r['id'])['days'] == 2                       # 아무것도 안 바뀜
    lv.transition(other['id'], 'cancel', KIM, '김가짜')
    ch.save('delete', {}, *BOSS, hid=hid, today=TODAY, notify=False)
    assert ch.get(hid) is None and lv.get(r['id'])['days'] == 3 and lv.balance('김가짜', date(2026, 12, 1)) == 0
    assert lv.is_workday(date(2026, 12, 29)) is True


def test_expect_mismatch_and_half_day(db):
    r = _approved(date(2026, 12, 29), ltype=lv.HALF_AM)
    with pytest.raises(ValueError, match='다시 확인'):
        _add('2026-12-29', expect=0)
    _add('2026-12-29', expect=1)
    assert lv.get(r['id'])['days'] == 0


def test_yearly_and_summary(db):
    _add('2026-08-07', name='창립기념일', yearly=True)
    assert lv.holiday_name(date(2027, 8, 7)) == '창립기념일' and lv.holiday_name(date(2025, 8, 7)) == ''
    y = ch.year_list(2027)
    row = next(e for e in y['entries'] if e['name'] == '창립기념일')
    assert row['editable'] and row['start'] == '2027-08-07' and row['yearly']
    assert y['by_kind'][ch.KIND_COMPANY] == 1 and y['total'] == y['by_kind'][ch.KIND_LEGAL] + 1
    assert any(m['name'] == '창립기념일' for m in ch.month_days(2027, 8))
    # 법정 공휴일과 겹치는 날은 원래 쉬는 날 — 다시 셀 휴가 없음, 회사 지정 개수에도 안 듦
    _approved(date(2026, 12, 24), date(2026, 12, 28))
    assert ch.preview('add', {'name': '성탄 휴무', 'kind': ch.KIND_COMPANY, 'start': '2026-12-25'}, today=TODAY)['count'] == 0


def test_validation(db):
    bad = [({'name': '', 'kind': 'company', 'start': '2026-12-01'}, '이름'),
           ({'name': 'x', 'kind': 'legal', 'start': '2026-12-01'}, '구분'),
           ({'name': 'x', 'kind': 'company', 'start': '2026-12-05', 'end': '2026-12-01'}, '앞서'),
           ({'name': 'x', 'kind': 'company', 'start': '2026-07-01', 'end': '2026-08-15'}, '31일'),
           ({'name': 'x', 'kind': 'company', 'start': '2026-12-30', 'end': '2027-01-02', 'yearly': True}, '한 해')]
    for data, msg in bad:
        with pytest.raises(ValueError, match=msg):
            ch.save('add', data, *BOSS, today=TODAY, notify=False)
    h = ch.save('add', {'name': ' 대체  공휴일 ', 'kind': 'temp', 'start': '2026-12-01', 'yearly': True}, *BOSS,
                today=TODAY, notify=False)
    got = ch.get(h['id'])
    assert got['name'] == '대체 공휴일' and got['yearly'] == 0                  # 임시공휴일은 그해만
    with pytest.raises(LookupError):
        ch.save('edit', {'name': 'y', 'kind': 'temp', 'start': '2026-12-02'}, *BOSS, hid=999, today=TODAY, notify=False)


def test_dm_text_particles():
    a = {'period': '12/29(화)', 'type_label': '연차', 'old_days': 3.0, 'new_days': 2.0, 'delta': -1.0, 'deducts': True}
    t = ch._dm_text(a, '연말 휴무', 'add')
    assert '「연말 휴무」가 회사 쉬는 날로' in t and '3일 → 2일' in t and '연차 1일이 돌아왔어요' in t
    assert '「창립기념일」이 취소되어' in ch._dm_text(a, '창립기념일', 'delete')
