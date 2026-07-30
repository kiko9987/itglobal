# -*- coding: utf-8 -*-
"""K열 재상담 헤더 제거 회귀 테스트 (Slack List 상담 컬럼 노출 방지).

배경: 방문 수정/취소/완료 경로(_trigger_visit_list_webhook)가 K열 상담 내용을
raw(헤더 포함)로 Slack List 에 전송 → `[MM.DD HH:MM 이니셜 · 상태]` 헤더가 List
상담 컬럼에 그대로 노출(L-03421, 2026-07-30). _extract_latest_consult_content 로
최신 회차 content 만 전달하도록 수정한 것의 회귀 보호.

핵심: 상태가 '방문 예약' 처럼 공백 포함 다중어절이어도 헤더가 제거돼야 함.
pure 함수만 — Redis/Slack 미접촉.
"""
import sys
sys.path.insert(0, '.')

import pytest
from dashboard.blueprints.slack_bot import (
    _extract_latest_consult_content, _parse_consultation_entries,
)


class TestHeaderStrip:
    @pytest.mark.parametrize('text, expected', [
        # L-03421 실사례 — 다중어절 상태 '방문 예약'
        ('[07.28 16:09 JK · 방문 예약] 이자카야 냉난방 견적 요청',
         '이자카야 냉난방 견적 요청'),
        # 단일어절 상태
        ('[07.20 10:00 YG · 부재중] 첫 상담 내용', '첫 상담 내용'),
        # 재상담 누적 — 최신 회차만
        ('[07.20 10:00 YG · 유선] 1회차 내용\n─────────\n'
         '[07.24 14:00 JK · 방문 예약] 2회차 최신 내용', '2회차 최신 내용'),
        # 인라인 ─── divider + 빈 방문완료 회차 (L-03371, 2026-07-30):
        #   최신(방문완료)이 빈 content → 직전 non-empty 회차 반환
        ('[07.24 15:20 SD · 방문 예약] 스40 설치 문의 ─── [07.30 10:51 MS · 방문 완료]',
         '스40 설치 문의'),
        # 인라인 divider + 최신 회차에 content 있음 → 최신 반환
        ('[07.24 15:20 SD · 방문 예약] 1회차 ─── [07.30 10:51 MS · 유선] 2회차 내용',
         '2회차 내용'),
        # 라인 divider + 빈 최신 회차 → 직전 non-empty
        ('[07.27 09:08 JSH · 방문 예약] 50평 식당\n─────────\n[07.30 10:18 MS · 방문 완료]',
         '50평 식당'),
    ])
    def test_strips_header(self, text, expected):
        assert _extract_latest_consult_content(text) == expected

    def test_inline_divider_parsed_as_two(self):
        """인라인 ' ─── ' divider 도 회차 분할 (표준 9칸 라인과 동일)."""
        e = _parse_consultation_entries(
            '[07.24 15:20 SD · 방문 예약] 문의 ─── [07.30 10:51 MS · 방문 완료]')
        assert len(e) == 2
        assert e[0]['content'] == '문의'
        assert e[1]['status'] == '방문 완료'
        assert e[1]['content'] == ''

    def test_headerless_unchanged(self):
        """헤더 없는 옛 형식은 원문 그대로 (no-op)."""
        plain = '이자카야에 설치하는 천장형 에어컨\n주방 1\n홀 2'
        assert _extract_latest_consult_content(plain) == plain

    def test_empty(self):
        assert _extract_latest_consult_content('') == ''
        assert _extract_latest_consult_content(None) == ''

    def test_multiword_status_parsed(self):
        """'방문 예약' 상태가 status 로 정확히 파싱 (content 로 새지 않음)."""
        e = _parse_consultation_entries('[07.28 16:09 JK · 방문 예약] 본문')
        assert len(e) == 1
        assert e[0]['status'] == '방문 예약'
        assert e[0]['content'] == '본문'
        assert e[0]['ini'] == 'JK'


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
