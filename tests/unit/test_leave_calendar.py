# -*- coding: utf-8 -*-
"""휴가·토요 휴무 → 구글 캘린더(11-25 부장님). 구글은 가짜 서비스로 대신한다(운영 캘린더로 아무것도 안 나감)."""
import sys
from datetime import date

sys.path.insert(0, '.')

import pytest
from cryptography.fernet import Fernet

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.services import leave_calendar as lc
from dashboard.services import sat_off as so

ADMIN = 'kiko@itg-aircon.com'
JW, YG, CO = 'UJW', 'UYG', 'UCO'
A, B = 'UA0001', 'UB0001'
TODAY = date(2026, 11, 20)


class _Req:
    def __init__(self, fn):
        self.fn = fn

    def execute(self):
        return self.fn()


class _HttpErr(Exception):
    def __init__(self, status):
        super().__init__(f'http {status}')
        self.resp = type('R', (), {'status': status})()


class FakeEvents:
    def __init__(self):
        self.items = {}        # id → 일정
        self.calls = []

    def list(self, calendarId, privateExtendedProperty, **kw):
        k, v = privateExtendedProperty.split('=')
        rows = [e for e in self.items.values() if e.get('status') != 'cancelled'
                and ((e.get('extendedProperties') or {}).get('private') or {}).get(k) == v]
        return _Req(lambda: {'items': [dict(e) for e in rows]})

    def insert(self, calendarId, body):
        def run():
            self.calls.append(('insert', body['id']))
            if body['id'] in self.items:
                raise _HttpErr(409)
            self.items[body['id']] = dict(body)
            return body
        return _Req(run)

    def update(self, calendarId, eventId, body):
        def run():
            self.calls.append(('update', eventId))
            self.items[eventId] = dict(body, id=eventId)
            return body
        return _Req(run)

    def delete(self, calendarId, eventId):
        def run():
            self.calls.append(('delete', eventId))
            self.items[eventId] = dict(self.items[eventId], status='cancelled')
        return _Req(run)


class FakeSvc:
    def __init__(self):
        self.ev = FakeEvents()

    def events(self):
        return self.ev


@pytest.fixture
def env(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, lv):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('LEAVE_APPROVER_IDS', f'{JW},{YG}')
    monkeypatch.setenv('SAT_OFF_COORDINATOR_IDS', CO)
    monkeypatch.setenv('GW_LEAVE_CALENDAR_ID', 'cal-test')
    monkeypatch.setattr(lc.time, 'sleep', lambda s: None)
    monkeypatch.setattr(lc, '_alerted_on', None)
    gw._invalidate()
    for name, uid, dept in (('영업가', A, '기술영업'), ('영업나', B, '기술영업'), ('빈대리', CO, '기술영업'),
                            ('박정우', JW, 'CEO'), ('박용구', YG, 'CEO')):
        hr.create_person({'name': name, 'status': '재직', 'hire_date': '2022-05-23', 'department': dept,
                          'slack_id': uid, 'work_email': f'{uid.lower()}@itg-aircon.com'}, ADMIN)
        lv.upsert_staff(name, '2022-05-23', uid, accrue_from='2026-11-11')
    fake = type('S', (), {'posts': [], 'conversations_open': lambda self, users: {'channel': {'id': 'D'}},
                          'chat_postMessage': lambda self, **k: {'ts': '1.1'}})()
    monkeypatch.setattr(so, '_client', lambda: fake)
    monkeypatch.setattr(so, '_run_bg', lambda fn: fn())
    yield FakeSvc()
    gw._invalidate()


def _leave(name, ltype, start, end, status=lv.APPROVED, reason='개인 사정'):
    with lv._conn() as conn:
        cur = conn.execute('INSERT INTO leave_requests (name, ltype, start, "end", days, reason, status) '
                           'VALUES (?, ?, ?, ?, 1, ?, ?)', (name, ltype, start, end, reason, status))
        return cur.lastrowid


def _titles(svc):
    return sorted((e['start']['date'], e['summary']) for e in svc.ev.items.values() if e.get('status') != 'cancelled')


def test_only_approved_leave_without_type_or_reason(env):
    _leave('영업가', lv.HALF_PM, '2026-11-27', '2026-11-27')
    _leave('영업나', lv.SICK, '2026-12-01', '2026-12-02', reason='감기')
    _leave('빈대리', lv.HALF_AM, '2026-12-03', '2026-12-03', status=lv.WAIT_1)       # 승인 전 → 안 올림
    _leave('박정우', lv.HALF_AM, '2026-12-04', '2026-12-04', status=lv.CANCELLED)    # 취소 → 안 올림
    r = lc.sync(TODAY, svc=env)
    assert r['added'] == 2 and not r['errors']
    assert _titles(env) == [('2026-11-27', '영업가 오후 반차'), ('2026-12-01', '영업나 휴가')]
    sick = [e for e in env.ev.items.values() if e['summary'] == '영업나 휴가'][0]
    assert sick['end']['date'] == '2026-12-03'            # 종일 일정 끝은 다음 날
    assert '병가' not in str(sick) and '감기' not in str(sick)


def test_second_run_changes_nothing_then_follows_cancel_and_edit(env):
    a = _leave('영업가', lv.ANNUAL, '2026-12-07', '2026-12-08')
    b = _leave('영업나', lv.HALF_AM, '2026-12-09', '2026-12-09')
    lc.sync(TODAY, svc=env)
    env.ev.calls.clear()
    assert lc.sync(TODAY, svc=env) == {'added': 0, 'changed': 0, 'removed': 0, 'errors': []}
    assert env.ev.calls == []
    with lv._conn() as conn:
        conn.execute('UPDATE leave_requests SET "end"=? WHERE id=?', ('2026-12-07', a))       # 하루 줄임
        conn.execute('UPDATE leave_requests SET status=? WHERE id=?', (lv.CANCELLED, b))      # 취소
    r = lc.sync(TODAY, svc=env)
    assert (r['changed'], r['removed']) == (1, 1)
    assert _titles(env) == [('2026-12-07', '영업가 휴가')]
    # 취소했다 다시 승인되면(같은 ID 가 지운 채로 남음) 되살린다
    with lv._conn() as conn:
        conn.execute('UPDATE leave_requests SET status=? WHERE id=?', (lv.APPROVED, b))
    r = lc.sync(TODAY, svc=env)
    assert r['added'] == 1 and not r['errors']
    assert ('2026-12-09', '영업나 오전 반차') in _titles(env)


def test_sat_off_only_after_confirmed(env):
    so.toggle(date(2026, 12, 5), '영업가', True, A, '영업가', TODAY)
    so.toggle(date(2026, 12, 12), '영업나', True, B, '영업나', TODAY)
    lc.sync(TODAY, svc=env)
    assert _titles(env) == []                              # 작성 중인 달은 안 올림
    so.submit('2026-12', CO, '빈대리', TODAY)
    lc.sync(TODAY, svc=env)
    assert _titles(env) == []                              # 결재 대기도 안 올림
    so.decide('2026-12', True, JW, '박정우')
    lc.sync(TODAY, svc=env)
    assert _titles(env) == [('2026-12-05', '영업가 토요 휴무'), ('2026-12-12', '영업나 토요 휴무')]
    assert all(e.get('colorId') == lc.SAT_COLOR for e in env.ev.items.values())


def test_full_off_day_is_one_event(env):
    so.set_full_off(date(2026, 12, 26), True, '박정우', '연말')
    so.submit('2026-12', CO, '빈대리', TODAY)
    so.decide('2026-12', True, JW, '박정우')
    lc.sync(TODAY, svc=env)
    assert _titles(env) == [('2026-12-26', '토요 전체 휴무 (연말)')]


def test_leaves_hand_made_events_alone(env):
    env.ev.items['manual1'] = {'id': 'manual1', 'summary': '회식', 'start': {'date': '2026-12-10'},
                               'end': {'date': '2026-12-11'}}
    lc.sync(TODAY, svc=env)
    assert env.ev.items['manual1'].get('status') != 'cancelled'


def test_off_without_calendar_id(env, monkeypatch):
    monkeypatch.setenv('GW_LEAVE_CALENDAR_ID', '')
    assert lc.sync(TODAY, svc=env) == {'skipped': 'off'}


def test_no_permission_alerts_once_a_day(env, monkeypatch):
    class Denied(FakeSvc):
        def events(self):
            ev = super().events()
            ev.list = lambda **kw: _Req(lambda: (_ for _ in ()).throw(_HttpErr(403)))
            return ev
    seen = []
    monkeypatch.setattr(lc.logger, 'error', lambda msg, **kw: seen.append(msg))
    monkeypatch.setattr(lc.logger, 'warning', lambda msg, **kw: None)
    lc.sync(TODAY, svc=Denied())
    lc.sync(TODAY, svc=Denied())
    assert len(seen) == 1 and '공유' in seen[0]


def test_event_ids_follow_google_rules():
    import re
    for key in ('lv:1', 'so:2026-12-05:영업가', 'sf:2026-12-26'):
        assert re.fullmatch(r'[a-v0-9]{5,1024}', lc._eid(key))
