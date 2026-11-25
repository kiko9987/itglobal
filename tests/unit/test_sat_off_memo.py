# -*- coding: utf-8 -*-
"""스케줄 근무 휴무 메모(2026-11-25 부장님) — 엑셀 「격주휴무 신청서」 메모란 대신.
각자 꼭 쉬어야 하는 토요일·못 쉬는 토요일과 사유를 먼저 적어 두면 다른 매니저가 보고 양보한다. 이름은 모두 가짜."""
import sys
from datetime import date

sys.path.insert(0, '.')

import pytest

from dashboard.services import sat_off as so
from tests.unit.test_sat_off import A, B, C, CO, DEC, JW, S5, S12, TODAY, YG, _confirm, slack  # noqa: F401


def test_must_memo_picks_and_everyone_sees(slack):
    r = so.write_memo(S12, '영업가', so.MEMO_MUST, ' 가족  여행 ', A, '영업가', TODAY)
    assert r['picked'] and not r['warn'] and r['reason'] == '가족 여행'
    assert '영업가' in so.picks(DEC)['2026-12-12']
    v = so.view(DEC, B, '영업나', TODAY)                      # 다른 매니저도 봄
    assert [(m['name'], m['kind'], m['reason']) for m in v['memos']] == [('영업가', 'must', '가족 여행')]
    assert v['can_memo'] and not v['can_memo_others'] and not v['can_month_note']
    # 같은 사람·같은 토요일은 고쳐 씀
    so.write_memo(S12, '영업가', so.MEMO_MUST, '결혼식', A, '영업가', TODAY)
    assert [m['reason'] for m in so.memos(DEC)['items']] == ['결혼식']


def test_must_memo_when_full_keeps_memo_and_asks_to_yield(slack):
    so.toggle(S12, '영업나', True, B, '영업나', TODAY)
    so.toggle(S12, '영업다', True, C, '영업다', TODAY)          # 4명 중 절반(2명) 마감
    r = so.write_memo(S12, '영업가', so.MEMO_MUST, '예비군', A, '영업가', TODAY)
    assert not r['picked'] and '양보' in r['warn']
    assert '영업가' not in so.picks(DEC)['2026-12-12'] and so.memos(DEC)['items'][0]['reason'] == '예비군'
    so.toggle(S12, '영업다', False, C, '영업다', TODAY)          # 영업다 양보
    so.toggle(S12, '영업가', True, A, '영업가', TODAY)
    assert sorted(so.picks(DEC)['2026-12-12']) == ['영업가', '영업나']


def test_cant_memo_and_permissions(slack):
    r = so.write_memo(S5, '영업나', so.MEMO_CANT, '현장 핸들링', B, '영업나', TODAY)
    assert not r['picked'] and '영업나' not in so.picks(DEC)['2026-12-05']
    with pytest.raises(PermissionError):
        so.write_memo(S5, '영업나', so.MEMO_MUST, '대신', A, '영업가', TODAY)      # 남의 메모는 안 됨
    so.write_memo(S5, '영업다', so.MEMO_MUST, '돌잔치', CO, '빈대리', TODAY)        # 취합 담당은 대신 적음
    with pytest.raises(PermissionError):
        so.delete_memo(S5, '영업다', A, '영업가')
    so.delete_memo(S5, '영업나', B, '영업나')                                       # 본인 것
    so.delete_memo(S5, '영업다', CO, '빈대리')                                      # 취합 담당
    assert so.memos(DEC)['items'] == []
    assert '영업다' in so.picks(DEC)['2026-12-05']                                  # 메모를 지워도 휴무 이름은 그대로
    with pytest.raises(LookupError):
        so.delete_memo(S5, '영업나', B, '영업나')


def test_month_note(slack):
    with pytest.raises(PermissionError):
        so.set_month_note(DEC, '전체휴무 소통', A, '영업가')
    so.set_month_note(DEC, '26일 전체휴무 대표님 소통 예정', CO, '빈대리')
    assert so.memos(DEC)['month_note']['reason'] == '26일 전체휴무 대표님 소통 예정'
    assert so.view(DEC, CO, '빈대리', TODAY)['can_month_note']
    so.set_month_note(DEC, '  ', YG, '박용구')                                     # 빈 글 = 지움
    assert so.memos(DEC)['month_note'] is None


def test_confirmed_month_and_validation(slack):
    so.toggle(S5, '영업다', True, C, '영업다', TODAY)
    _confirm()
    r = so.write_memo(S12, '영업가', so.MEMO_MUST, '해외여행', A, '영업가', TODAY)
    assert not r['picked'] and '변경 신청' in r['warn']
    for bad, msg in (('', '사유'), ('가' * 61, '60자')):
        with pytest.raises(so.SatOffError, match=msg):
            so.write_memo(S12, '영업가', so.MEMO_MUST, bad, A, '영업가', TODAY)
    with pytest.raises(so.SatOffError, match='지난 토요일'):
        so.write_memo(S5, '영업가', so.MEMO_MUST, '여행', A, '영업가', date(2026, 12, 8))
    with pytest.raises(so.SatOffError, match='종류'):
        so.write_memo(S12, '영업가', 'month', '여행', A, '영업가', TODAY)
    # 승인권자는 확정된 달에도 꼭 쉬어야 하는 날을 넣으면 이름까지 바로
    r = so.write_memo(S12, '영업나', so.MEMO_MUST, '결혼식', JW, '박정우', TODAY)
    assert r['picked']
