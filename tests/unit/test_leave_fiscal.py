"""회계연도 전환(2027-01-01, 11-18 부장님 확정)·퇴직 정산."""
from datetime import date

import pytest

from dashboard.services import leave as lv

# 설계 문서(docs/designs/leave-cutover-2026-12.md 3. 정산표)의 18명: (이름, 입사일, 2027-01-01 비례분, 2028-01-01 부여)
TABLE = [
    ('박용구', '2007-01-01', 24, 25), ('박정우', '2019-08-07', 7.5, 18), ('박민우', '2022-02-03', 15.5, 17),
    ('김종연', '2022-05-23', 10.5, 17), ('김태현', '2022-05-23', 10.5, 17), ('황샛별', '2023-04-04', 12, 16),
    ('권태숙', '2023-07-24', 7.5, 16), ('권태훈', '2024-04-01', 12.5, 16), ('주영민', '2024-04-08', 12, 16),
    ('빈승정', '2024-08-12', 6.5, 16), ('박민재', '2024-10-02', 4, 16), ('고광일', '2025-01-01', 15, 16),
    ('조성헌', '2025-04-22', 10.5, 15), ('강정권', '2026-04-01', 11.5, 15), ('권진현', '2026-11-02', 2.5, 15),
    ('이상효', '2026-11-02', 2.5, 15), ('이종원', '2026-11-03', 2.5, 15), ('김민욱', '2026-11-09', 2.5, 15),
]


def _fiscal(name, hire, today):
    return {g['valid_from'].year: g['days'] for g in lv.fiscal_grants({'name': name, 'hire_date': hire}, today)}


@pytest.mark.parametrize('name,hire,y27,y28', TABLE)
def test_fiscal_table_matches_design(name, hire, y27, y28):
    assert lv.FISCAL_FROM == date(2027, 1, 1)
    assert _fiscal(name, hire, date(2026, 12, 14)) == {}                 # 12-15 전엔 아직
    assert _fiscal(name, hire, date(2026, 12, 15)) == {2027: y27}        # 12-15 부터 미리
    assert _fiscal(name, hire, date(2028, 1, 1)) == {2027: y27, 2028: y28}


def test_no_anniversary_grants_from_2027_and_new_hire_in_2027():
    p = {'name': '주영민', 'hire_date': '2024-04-08', 'accrue_from': '2026-11-06'}
    srcs = [g['source'] for g in lv.due_grants(p, date(2027, 6, 1))]
    assert srcs == ['fiscal:주영민:2027']                                # 2027-04-08 기념일 부여 없음
    n = {'name': '신입', 'hire_date': '2027-03-01', 'accrue_from': '2027-03-01'}
    g = lv.due_grants(n, date(2028, 1, 1))
    assert [x['source'].split(':')[0] for x in g] == ['monthly'] * 10 + ['fiscal']
    assert g[-1]['days'] == 13 and g[-1]['valid_from'] == date(2028, 1, 1)   # 15 × 306/365 = 12.57 → 13


def test_hire_basis_respects_day_after_rule():
    p = {'name': '김종연', 'hire_date': '2022-05-23'}
    at = lambda d: [g['valid_from'] for g in lv.hire_basis_grants(p, d) if g['valid_from'] >= date(2027, 1, 1)]  # noqa: E731
    assert at(date(2027, 5, 22)) == []                                   # 5년을 꽉 채우고 그날 퇴직 → 그 해 몫 없음
    assert at(date(2027, 5, 23)) == [date(2027, 5, 23)]


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(lv, '_db', lambda: d)
    monkeypatch.setattr(lv, '_table_ready', False)
    monkeypatch.setenv('LEAVE_APPROVER_IDS', 'UJW,UYG')
    lv.upsert_staff('김종연', '2022-05-23', 'UKJ', accrue_from='2026-12-03')
    lv.add_grant('김종연', 7, date(2026, 12, 2), date(2027, 5, 22), 'flex:김종연', '플렉스 잔여')
    return d


def test_settlement_design_example(db):
    # 설계 흐름 6: 2027-06-30 퇴사 → 입사일 기준 17(2027-05-23) / 회계연도 실제 10.5 → 차액 6.5
    lv.accrue(date(2027, 6, 30))
    r = lv.create('김종연', 'UKJ', 'UKJ', '김종연', lv.ANNUAL, '', date(2027, 6, 7), date(2027, 6, 8), 'x')
    lv.create('김종연', 'UKJ', 'UKJ', '김종연', lv.ANNUAL, '', date(2027, 7, 6), date(2027, 7, 6), 'x')
    s = lv.resign_settlement('김종연', date(2027, 6, 30))
    assert s['ok'] and s['basis_total'] == 17 and s['actual_total'] == 10.5 and s['short'] == 6.5
    assert s['unused'] == 8.5 and len(s['later']) == 1                        # 7/6 휴가는 퇴직일 뒤 → 취소 안내
    assert r


def test_settlement_no_short_and_before_fiscal(db):
    lv.accrue(date(2027, 3, 1))
    s = lv.resign_settlement('김종연', date(2027, 5, 22))                    # 기념일 전날 퇴직 → 입사일 기준 0
    assert s['basis_total'] == 0 and s['actual_total'] == 10.5 and s['short'] == 0   # 더 준 10.5 는 깎지 않음
    s = lv.resign_settlement('김종연', date(2026, 12, 20))
    assert s['before_fiscal'] and s['short'] == 0 and s['unused'] == 7
