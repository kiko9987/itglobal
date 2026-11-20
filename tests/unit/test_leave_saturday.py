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
