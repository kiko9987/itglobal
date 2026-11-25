# -*- coding: utf-8 -*-
"""토요일 근무 (2026-11-20 부장님: 시공은 토요일 근무, 영업은 토요일 스케줄 근무).

- 부서(인사카드)가 시공이면 토요일도 휴가 일수에 든다(플렉스 근무 규칙 '시공' 월~토).
- 영업이면 휴가 첫날·마지막 날로 고른 토요일만 든다(근무표 토요일). 기간 가운데 낀 토요일은 빠진다.
- 그 밖(경영지원·미화·대표·부서 없음)은 월~금. 일요일·공휴일은 누구나 빠진다.
- 플렉스 기록과 같아야 한다: 박민재 조의 10-06~10-10 = 4일(10-09 한글날 빠짐, 10-10 토 셈).
"""
import sys
from datetime import date

sys.path.insert(0, '.')

import pytest
from cryptography.fernet import Fernet

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.services import leave_export

ADMIN = 'kiko@itg-aircon.com'
SIGONG, SALES, OFFICE = 'U0SIGONG01', 'U0SALES001', 'U0OFFICE01'


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
    for name, uid, dept in (('시공직원', SIGONG, '기술시공(내부)'), ('영업직원', SALES, '기술영업'),
                            ('사무직원', OFFICE, '경영지원')):
        hr.create_person({'name': name, 'status': '재직', 'hire_date': '2022-05-23', 'department': dept,
                          'slack_id': uid, 'work_email': f'{uid.lower()}@itg-aircon.com'}, ADMIN)
        lv.upsert_staff(name, '2022-05-23', uid, accrue_from='2026-11-11')
        lv.add_grant(name, 15, date(2026, 11, 10), date(2027, 5, 22), f'flex:{name}', '플렉스 잔여')
    yield d
    gw._invalidate()


def test_work_pattern_follows_hr_department(db):
    assert lv.work_pattern('시공직원') == lv.WORK_SAT
    assert lv.work_pattern('영업직원') == lv.WORK_SAT_SCHEDULE
    assert lv.work_pattern('사무직원') == lv.WORK_WEEKDAY
    assert lv.work_pattern('없는사람') == lv.WORK_WEEKDAY


def test_single_saturday(db):
    sat = date(2026, 11, 21)
    assert lv.count_days(lv.ANNUAL, sat, sat, '시공직원') == 1
    assert lv.count_days(lv.ANNUAL, sat, sat, '영업직원') == 1        # 토요일을 골라 쉼 = 근무 토요일
    assert lv.count_days(lv.ANNUAL, sat, sat, '사무직원') == 0
    assert lv.count_days(lv.HALF_AM, sat, None, '시공직원') == 0.5
    assert lv.count_days(lv.ANNUAL, date(2026, 11, 22), None, '시공직원') == 0   # 일요일은 누구나 빠짐
    assert lv.validate('시공직원', lv.ANNUAL, '', sat, sat) == {}
    assert '쉬는 날' in lv.validate('사무직원', lv.ANNUAL, '', sat, sat)['start']


def test_saturday_inside_range(db):
    """목~화(토 낀 기간): 시공은 토요일 셈, 영업은 가운데 토요일이라 안 셈, 사무는 안 셈."""
    a, b = date(2026, 11, 19), date(2026, 11, 24)
    assert lv.count_days(lv.ANNUAL, a, b, '시공직원') == 5
    assert lv.count_days(lv.ANNUAL, a, b, '영업직원') == 4
    assert lv.count_days(lv.ANNUAL, a, b, '사무직원') == 4
    # 영업이 토요일까지 쉬면(마지막 날 = 토) 그 토요일은 셈
    assert lv.count_days(lv.ANNUAL, date(2026, 11, 19), date(2026, 11, 21), '영업직원') == 3


def test_flex_parity_park_minjae_mourning(db):
    """플렉스: 10-06(화)~10-10(토) 조의 = 06·07·08·10 → 4일. 10-09 한글날은 누구나 빠짐."""
    a, b = date(2026, 10, 6), date(2026, 10, 10)
    assert lv.count_days(lv.EVENT, a, b, '영업직원') == 4
    assert lv.count_days(lv.EVENT, a, b, '사무직원') == 3
    # 플렉스: 강민석(영업) 08-14(금)~08-18(화) = 2일 — 15(토) 광복절·17(월) 대체공휴일·16 일요일 빠짐
    assert lv.count_days(lv.ANNUAL, date(2026, 8, 14), date(2026, 8, 18), '영업직원') == 2
    # 플렉스: 김종연(시공) 2025-06-10~06-21 = 11일(토 14·21 셈, 일 빠짐)
    assert lv.count_days(lv.ANNUAL, date(2025, 6, 10), date(2025, 6, 21), '시공직원') == 11


def test_saturday_leave_deducts_annual_and_counts_toward_limits(db):
    sat = date(2026, 11, 28)
    before = lv.balance('시공직원', sat)
    r = lv.create('시공직원', SIGONG, SIGONG, '시공직원', lv.ANNUAL, '', sat, sat)
    assert r and r['days'] == 1 and lv.balance('시공직원', sat) == before - 1
    # 경조사 한도(한 해 누적)도 영업 근무 토요일을 같은 기준으로 센다
    lv.create('영업직원', SALES, SALES, '영업직원', lv.EVENT, 'mourn1', date(2026, 10, 6), date(2026, 10, 10))
    assert lv.used_in_year('영업직원', lv.EVENT, 2026, 'mourn1') == 4
    assert '이미 신청한 4일' in lv.validate('영업직원', lv.EVENT, 'mourn1',
                                                      date(2026, 12, 1), date(2026, 12, 2))['end']


def test_year_split_keeps_original_saturday_boundary(db):
    """해를 넘는 기간을 해마다 나눠 셀 때도 영업 토요일은 원래 첫날·마지막 날 기준."""
    a, b = date(2026, 12, 26), date(2027, 1, 4)        # 토 ~ 월
    assert lv.days_in_year(lv.SICK, a, b, 2026, '영업직원') == 5    # 26(토, 첫날)·28·29·30·31
    assert lv.days_in_year(lv.SICK, a, b, 2027, '영업직원') == 1    # 1/1 신정·1/2(토, 가운데)·1/3 일 빠짐 → 1/4
    assert lv.days_in_year(lv.SICK, a, b, 2027, '시공직원') == 2    # 시공은 1/2(토)도 셈


# ── 남은 날 취소 (11-24 부장님) ─────────────────────────────────────────
def _approved(name, uid, a, b):
    r = lv.create(name, uid, uid, name, lv.ANNUAL, '', a, b)
    with lv._conn() as conn:
        conn.execute('UPDATE leave_requests SET status=? WHERE id=?', (lv.APPROVED, r['id']))
    return lv.get(r['id'])


def test_cut_rest_on_sunday_does_not_count_middle_saturday(db):
    """영업 목~화 휴가(가운데 토요일은 안 셈 = 4일)를 일요일에 취소 — 남는 건 목·금 2일, 마지막 날은 금요일.
    예전엔 '어제'(토)가 새 마지막 날이 되어 토요일까지 3일이 빠졌다."""
    r = _approved('영업직원', SALES, date(2026, 11, 19), date(2026, 11, 24))
    assert r['days'] == 4
    sun = date(2026, 11, 22)
    assert lv.cancel_mode(r, SALES, sun) == 'rest'
    r = lv.cut_rest(r['id'], SALES, '영업직원', today=sun)
    assert r['end'] == '2026-11-20' and r['days'] == 2
    assert lv.count_days(lv.ANNUAL, lv._d(r['start']), lv._d(r['end']), '영업직원') == 2   # 다시 세도 같음
    assert '11월 20일 (금)까지만 휴가' in lv.cut_notice_text(r, sun)


def test_cut_rest_keeps_counted_saturday_for_sigong_and_skips_weekend_for_office(db):
    r = _approved('시공직원', SIGONG, date(2026, 11, 19), date(2026, 11, 24))
    r = lv.cut_rest(r['id'], SIGONG, '시공직원', today=date(2026, 11, 22))
    assert r['end'] == '2026-11-21' and r['days'] == 3                 # 시공은 토요일도 근무일 → 토요일까지 남김
    r = _approved('사무직원', OFFICE, date(2026, 11, 19), date(2026, 11, 25))
    r = lv.cut_rest(r['id'], OFFICE, '사무직원', today=date(2026, 11, 23))   # 월요일에 취소
    assert r['end'] == '2026-11-20' and r['days'] == 2                 # 주말로 끝나지 않고 금요일까지


def test_cut_rest_not_offered_when_only_uncounted_days_passed(db):
    """토요일(첫날)부터인 사무직 휴가를 월요일에 취소 — 지난 날(토·일)은 센 날이 없으니 전부 취소."""
    r = lv.create('사무직원', OFFICE, OFFICE, '사무직원', lv.ANNUAL, '', date(2026, 11, 21), date(2026, 11, 25))
    with lv._conn() as conn:
        conn.execute('UPDATE leave_requests SET status=? WHERE id=?', (lv.APPROVED, r['id']))
    r = lv.get(r['id'])
    assert lv.cancel_mode(r, OFFICE, date(2026, 11, 23)) == 'all'
    assert lv.cut_rest(r['id'], OFFICE, '사무직원', today=date(2026, 11, 23)) is None


# ── 일수 미리보기 (11-24 부장님) ────────────────────────────────────────
def test_days_preview_explains_saturdays(db):
    a, b = date(2026, 11, 19), date(2026, 11, 24)            # 목~화, 21일 토
    p = lv.days_preview('영업직원', lv.ANNUAL, a, b)
    assert p['days'] == 4 and p['lines'] == ['휴가 일수: 4일',
                                             '11월 21일 (토) — 기간 가운데 낀 토요일은 근무가 없는 날로 봐서 빠져요.']
    p = lv.days_preview('영업직원', lv.ANNUAL, a, date(2026, 11, 21))
    assert p['days'] == 3 and p['lines'][1] == '11월 21일 (토) — 휴가 첫날이나 마지막 날인 토요일은 근무일로 봐서 들어가요.'
    p = lv.days_preview('시공직원', lv.ANNUAL, date(2026, 11, 14), b)
    assert p['days'] == 9 and p['lines'][1] == '11월 14일 (토), 11월 21일 (토) — 토요일도 근무일이라 들어가요.'
    p = lv.days_preview('사무직원', lv.ANNUAL, a, b)
    assert p['days'] == 4 and p['lines'][1] == '11월 21일 (토) — 토요일은 근무일이 아니라 빠져요.'
    # 공휴일·반차·일요일
    p = lv.days_preview('사무직원', lv.ANNUAL, date(2026, 12, 24), date(2026, 12, 25))
    assert p['lines'] == ['휴가 일수: 1일', '12월 25일 (금) — 쉬는 날(크리스마스)이라 빠져요.']
    assert lv.days_preview('시공직원', lv.HALF_AM, date(2026, 11, 21), None)['lines'] == [
        '휴가 일수: 0.5일', '11월 21일 (토) — 토요일도 근무일이라 들어가요.']
    assert lv.days_preview('시공직원', lv.ANNUAL, date(2026, 11, 22), None)['lines'] == ['휴가 일수: 0일']
    # 날짜가 덜 골라졌거나 거꾸로면 비움
    assert lv.days_preview('영업직원', lv.ANNUAL, None, b) == {}
    assert lv.days_preview('영업직원', lv.ANNUAL, b, a) == {}


def test_days_preview_follows_confirmed_schedule(db):
    """기술영업은 스케줄 근무표가 확정된 달이면 그 표대로 — 쉬는 토요일은 첫날이어도 빠지고, 근무 토요일은 가운데여도 듦."""
    from dashboard.services import sat_off
    with lv._conn() as conn:
        conn.execute('INSERT INTO sat_off_months (month, status) VALUES (?,?)', ('2026-11', sat_off.CONFIRMED))
        conn.execute('INSERT INTO sat_off_days (day, name) VALUES (?,?)', ('2026-11-21', '영업직원'))
    p = lv.days_preview('영업직원', lv.ANNUAL, date(2026, 11, 21), date(2026, 11, 23))
    assert p['days'] == 1 and p['lines'][1] == '11월 21일 (토) — 스케줄 근무표에서 쉬는 토요일이라 빠져요.'
    p = lv.days_preview('영업직원', lv.ANNUAL, date(2026, 11, 27), date(2026, 11, 30))
    assert p['days'] == 3 and p['lines'][1] == '11월 28일 (토) — 스케줄 근무표에서 근무하는 토요일이라 들어가요.'
    assert p['days'] == lv.count_days(lv.ANNUAL, date(2026, 11, 27), date(2026, 11, 30), '영업직원')


def test_slack_leave_view_shows_days_preview(db):
    """슬랙 신청 창 — 날짜 칸은 고를 때마다 창을 다시 그리고(dispatch_action), 고른 기간의 일수·토요일 설명을 보인다."""
    from dashboard.blueprints import slack_leave as sl
    v = sl.leave_view({'name': '영업직원'})
    pickers = [b for b in v['blocks'] if b.get('block_id') in ('lv_start', 'lv_end')]
    assert len(pickers) == 2 and all(b['dispatch_action'] for b in pickers)
    assert not any('휴가 일수' in str(b) for b in v['blocks'])               # 날짜 전엔 없음
    v = sl.leave_view({'name': '영업직원'}, {'type': lv.ANNUAL, 'start': '2026-11-19', 'end': '2026-11-24'})
    ctx = [b for b in v['blocks'] if b['type'] == 'context' and '휴가 일수' in b['elements'][0]['text']]
    assert ctx[0]['elements'][0]['text'] == ('*휴가 일수: 4일*\n'
                                             '• 11월 21일 (토) — 기간 가운데 낀 토요일은 근무가 없는 날로 봐서 빠져요.')
    ids = [b.get('block_id') for b in v['blocks'] if b.get('block_id')]
    assert ids == ['lv_type', 'lv_start', 'lv_end', 'lv_reason']               # 칸 순서는 그대로
