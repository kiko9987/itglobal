"""배우자 출산휴가 법 규칙(2026-11-25 부장님: 만들어 두고 아직 켜지 않음 — 박민재 님까지는 회사 재량).
남녀고용평등법 제18조의2: 20일, 출산예정일 50일 전부터, 출산한 날부터 120일이 지나면 못 씀, 3회까지 나눠 씀."""
from datetime import date

import pytest

from dashboard.services import leave as lv

KIKO, SB, MW, JW, YG = 'UKIKO', 'USB', 'UMW', 'UJW', 'UYG'
SB_T = lv.SPOUSE_BIRTH


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(lv, '_db', lambda: d)
    monkeypatch.setattr(lv, '_table_ready', False)
    monkeypatch.setenv('LEAVE_APPROVER_IDS', f'{JW},{YG}')
    monkeypatch.setattr(lv, 'manager_ids', lambda: (SB, KIKO))
    lv.upsert_staff('박민우', '2022-02-03', MW, accrue_from='2026-11-01')
    return d


def _take(day):
    return lv.create('박민우', MW, MW, '박민우', SB_T, '', day, day, '배우자 출산')


def test_off_by_default_keeps_old_way(db):
    """꺼져 있으면 지금과 같다 — 기간은 받은 대로(비우면 언제든), 나눠 쓰기 제한 없음."""
    assert lv.SPOUSE_RULES_ON is False
    g, err = lv.custom_grant('박민우', SB_T, 20, date(2026, 11, 25), None, '', KIKO, '고광일',
                             today=date(2026, 11, 25), birth=date(2026, 11, 20))
    assert err == '' and g['valid_to'] == lv.ANYTIME and g['birth_date'] is None
    for d in (2, 3, 4, 7):                                     # 4번째도 막지 않음
        assert _take(date(2026, 12, d))
    assert lv.custom_left('박민우', SB_T, date(2026, 12, 8)) == 16


def test_window_from_birth_and_due():
    b = date(2026, 7, 31)
    assert lv.spouse_window(birth=b) == (date(2026, 6, 11), date(2026, 11, 27))      # 출산일을 1일째로 120일째
    assert lv.spouse_window(birth=b, due=date(2026, 8, 5)) == (date(2026, 6, 16), date(2026, 11, 27))
    assert lv.spouse_window(due=date(2026, 8, 5)) == (date(2026, 6, 16), date(2026, 12, 2))   # 출산 전 — 예정일 기준
    with pytest.raises(ValueError):
        lv.spouse_window()


def test_on_sets_window_and_limits_three_splits(db, monkeypatch):
    monkeypatch.setattr(lv, 'SPOUSE_RULES_ON', True)
    g, err = lv.custom_grant('박민우', SB_T, None, None, None, '', KIKO, '고광일', today=date(2026, 11, 25))
    assert '출산일' in err
    g, err = lv.custom_grant('박민우', SB_T, None, date(2026, 1, 1), None, '', KIKO, '고광일',
                             today=date(2026, 11, 25), birth=date(2026, 11, 20))
    assert err == '' and g['days'] == 20 and g['birth_date'] == '2026-11-20'
    assert (g['valid_from'], g['valid_to']) == ('2026-10-01', '2027-03-19')        # 받은 기간은 무시
    assert _take(date(2026, 12, 2)) and _take(date(2026, 12, 9))
    r = lv.create('박민우', MW, MW, '박민우', SB_T, '', date(2026, 12, 14), date(2026, 12, 18), '배우자 출산')
    assert r and r['days'] == 5                                                     # 3번째(여러 날 한 번에)
    errs = lv.validate('박민우', SB_T, '', date(2026, 12, 21), date(2026, 12, 21))
    assert '3번까지' in errs['start'] and '이미 3번' in errs['start']
    assert lv.validate('박민우', SB_T, '', date(2027, 3, 22), date(2027, 3, 22))     # 기한 뒤도 막힘(기간 밖)


def test_on_but_old_grant_without_birth_is_not_limited(db, monkeypatch):
    """켜도 출산일 없이 준 예전 부여(박민재 님처럼 '언제든')는 나눠 쓰기를 막지 않는다."""
    lv.custom_grant('박민우', SB_T, 20, date(2026, 11, 25), None, '', KIKO, '고광일', today=date(2026, 11, 25))
    monkeypatch.setattr(lv, 'SPOUSE_RULES_ON', True)
    for d in (2, 3, 4, 7, 8):
        assert _take(date(2026, 12, d))
