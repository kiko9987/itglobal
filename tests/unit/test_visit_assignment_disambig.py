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


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
