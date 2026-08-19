# -*- coding: utf-8 -*-
"""캔버스2 배정 → lead 매칭 disambiguation 회귀 테스트.

배경: 반복 방문 현장(무전화·동일 주소)에서 캔버스2 라인이 옛 lead 로 잘못 매칭.
과천 서울랜드: ETC-d8b768(07-23~24 지난) vs ETC-04915c(07-30~31 현행) 둘 다
'과천 광명로 181 서울랜드' 무전화. 주소 fallback 이 첫 매치(옛것)를 반환해
현행 방문이 v9 DM 에서 누락됨 (2026-07-29).

수정: _prefer_active_recent — ①지난방문(end<today) 제외 → ②활성 → ③최신.
phone 경로(_pick_lead_for_phone)·주소 fallback(_resolve_lead_for_assignment) 공용.
pure 함수만 — Slack/시트 미접촉 (date.today() 상대 계산으로 날짜 비의존).
"""
import sys
sys.path.insert(0, '.')

from datetime import date, timedelta
import pytest
import dashboard.services.visit_assignment_sync as vas

_TODAY = date.today()
_PAST = f'{(_TODAY - timedelta(days=5)).isoformat()}'
_PAST_END = f'{(_TODAY - timedelta(days=6)).isoformat()}~{(_TODAY - timedelta(days=5)).isoformat()}'
_FUT = f'{(_TODAY + timedelta(days=3)).isoformat()}'
_FUT_RANGE = f'{(_TODAY + timedelta(days=1)).isoformat()}~{(_TODAY + timedelta(days=2)).isoformat()}'
_ADDR = '과천 광명로 181 서울랜드'


def _lead(no, vd, status='방문 예약', addr=_ADDR):
    return {'리드 No': no, '방문 예정일': vd, '상태': status, '방문 주소': addr}


class TestPreferActiveRecent:
    def test_excludes_past_visit(self):
        """지난 방문 lead 는 현행 lead 있으면 선택 안 됨 (순서 무관)."""
        old = _lead('ETC-old', _PAST_END)
        cur = _lead('ETC-cur', _FUT_RANGE)
        assert vas._prefer_active_recent([old, cur])['리드 No'] == 'ETC-cur'
        assert vas._prefer_active_recent([cur, old])['리드 No'] == 'ETC-cur'  # 순서 무관

    def test_single_returns_itself(self):
        one = _lead('ETC-one', _PAST_END)  # 지난 방문 1개뿐이면 그거라도 반환
        assert vas._prefer_active_recent([one])['리드 No'] == 'ETC-one'

    def test_both_future_picks_latest(self):
        """둘 다 현행이면 최신(뒤쪽) 우선 — 기존 2026-07-27 동작 유지."""
        a = _lead('ETC-a', _FUT)
        b = _lead('ETC-b', _FUT_RANGE)
        assert vas._prefer_active_recent([a, b])['리드 No'] == 'ETC-b'

    def test_empty(self):
        assert vas._prefer_active_recent([]) is None

    def test_prefers_active_over_completed(self):
        """현행 중 활성(방문 예약) 우선, 완료 상태 배제."""
        done = _lead('ETC-done', _FUT, status='방문 완료')
        active = _lead('ETC-act', _FUT_RANGE, status='방문 예약')
        assert vas._prefer_active_recent([active, done])['리드 No'] == 'ETC-act'


class TestResolveAddressFallback:
    def test_repeat_visit_picks_current(self):
        """서울랜드 재현: 주소 fallback 이 현행(07-30~31) lead 선택."""
        old = _lead('ETC-d8b768', _PAST_END)
        cur = _lead('ETC-04915c', _FUT_RANGE)
        addr_candidates = [old, cur]  # 옛것이 앞 (시트 순서)
        a = {'phone_digits': '', 'address': _ADDR}
        picked = vas._resolve_lead_for_assignment(a, {}, addr_candidates)
        assert picked is not None and picked['리드 No'] == 'ETC-04915c'


class TestPhoneDateDisambig:
    """번호+주소가 동일한 반복 방문(다른 날짜)은 캔버스 라인 날짜로 회차 구분.

    ETC-986341(08-20) vs ETC-eccc5c(08-29): 010-3751-3157·관악구 광장빌딩 동일.
    주소로 못 가르니 _prefer_active_recent 이 '최신(뒤쪽)' = 08-29 를 골라 08-20
    라인이 엉뚱한 lead 에 붙음 (2026-08-18). 라인 날짜로 정확한 회차 선택.
    """

    @staticmethod
    def _md(vd):
        m, d = int(vd[5:7]), int(vd[8:10])
        return (m, d)

    def test_line_date_picks_matching_visit(self):
        early = _lead('ETC-early', _FUT_RANGE)   # today+1~+2 (앞쪽)
        late = _lead('ETC-late', _FUT)           # today+3 (뒤쪽 = 기본 선택)
        # 라인 날짜 없으면 기존대로 최신(late)
        assert vas._pick_lead_for_phone([early, late])['리드 No'] == 'ETC-late'
        # 라인 날짜가 early 와 같으면 early — 뒤쪽 기본값을 날짜가 뒤집음
        assert vas._pick_lead_for_phone(
            [early, late], _ADDR, self._md(_FUT_RANGE))['리드 No'] == 'ETC-early'
        # 라인 날짜가 late 와 같으면 late
        assert vas._pick_lead_for_phone(
            [early, late], _ADDR, self._md(_FUT))['리드 No'] == 'ETC-late'

    def test_line_date_no_match_falls_back(self):
        """라인 날짜가 어느 후보와도 안 맞으면 기존 로직(최신)으로 폴백."""
        early = _lead('ETC-early', _FUT_RANGE)
        late = _lead('ETC-late', _FUT)
        assert vas._pick_lead_for_phone(
            [early, late], _ADDR, (1, 1))['리드 No'] == 'ETC-late'

    def test_line_md_parse(self):
        assert vas._line_md_from_assignment(
            {'raw': '(MW) 8월 20일 / 010-3751-3157 / 광장빌딩 / 공사'}) == (8, 20)
        assert vas._line_md_from_assignment(
            {'raw': '(당) 8월 19~21일 / ... '}) == (8, 19)   # 범위는 시작일
        assert vas._line_md_from_assignment({'raw': '주소만 / 내용'}) is None


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
