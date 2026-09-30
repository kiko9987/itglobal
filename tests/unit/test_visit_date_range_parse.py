# -*- coding: utf-8 -*-
"""방문 예정일 범위 파싱 회귀 테스트 (2026-09-29 ETC-908d6c).

버그: '2026-09-29~10-02'(끝=월-일, 연 생략)에서 파서가 '10'을 일(日)로 오인하고
월은 시작 월(09)로 채워 2026-09-10(과거)으로 계산 → 캔버스1·DM 필터에서 '지난
방문'으로 오제외. 종료일의 연·월을 각각 optional 로 두어 MM-DD 형식을 인식.

pure 파서 테스트(날짜 비의존) + _fetch_visit_leads 포함 테스트(today mock).
"""
import sys
sys.path.insert(0, '.')

from datetime import date
import pandas as pd
import pytest

import dashboard.services.visit_assignment_sync as vas
import dashboard.services.visit_canvas_sync as vcs


class TestParseVisitDateEnd:
    @pytest.mark.parametrize('vd,expected', [
        ('2026-09-29~10-02', date(2026, 10, 2)),   # 끝=월-일(연 생략) ← 버그 케이스
        ('2026-09.29~9.30', date(2026, 9, 30)),    # 끝=월.일(점 구분)
        ('2026-01-29~30', date(2026, 1, 30)),      # 끝=일만
        ('2026-09-29~2026-10-02', date(2026, 10, 2)),  # 끝=전체 날짜
        ('2026-09-29', date(2026, 9, 29)),         # 단일
    ])
    def test_end(self, vd, expected):
        assert vas._parse_visit_date_end(vd) == expected

    def test_start_unaffected(self):
        assert vas._parse_visit_date_start('2026-09-29~10-02') == date(2026, 9, 29)


class TestCanvasDateDisplay:
    """캔버스1 표시 — 월 넘김 범위는 끝 월 명시 (2026-09-30 ETC-afaa5d '9월 30~10일' 오표기)."""
    @pytest.mark.parametrize('vd,expected', [
        ('2026-09-30~10-01', '9월 30일~10월 1일'),     # 버그 케이스 (옛: '9월 30~10일')
        ('2026-09-29~-10-02', '9월 29일~10월 2일'),    # '~-' 오타형
        ('2026-12-30~2027-01-02', '12월 30일~1월 2일'),
        ('2026-09-29~30', '9월 29~30일'),               # 같은 달
        ('2026-09.29~9.30', '9월 29~30일'),             # 점 구분, 같은 달
        ('2026-09-29~2026-09-30', '9월 29~30일'),
        ("'2026-10-02", '10월 2일'),                     # 단일(escape)
        ('2026-09.29', '9월 29일'),
        ('', '-'),
        ('미정', '미정'),
    ])
    def test_display(self, vd, expected):
        assert vcs._fmt_visit_date(vd) == expected


class _FakeRedis:
    def get(self, *_a, **_k):
        return None


def test_fetch_visit_leads_includes_month_day_range_end(monkeypatch):
    """끝이 월-일(연 생략) 범위인 방문 예약이 캔버스1 소스에 포함돼야 함.

    today=09-20 기준: 버그 파싱(09-10)이면 과거로 제외, 정상 파싱(10-02)이면 포함.
    """
    df = pd.DataFrame([{
        '리드 No': 'ETC-rangeend', '상태': '방문 예약',
        '방문 예정일': '2026-09-29~10-02', '고객명': '모랫말로',
        '고객 연락처': '010-0000-0000', '방문 주소': '인천 영종구 모랫말로 6-8',
        '상담 내용': '철거', '온라인 상담자': '박민우', '플랫폼': '기타',
    }])
    monkeypatch.setattr('dashboard.services.lead_service.load_leads_data',
                        lambda *a, **k: df)
    monkeypatch.setattr('dashboard.utils.redis_client.get_redis_client',
                        lambda: type('C', (), {'redis': _FakeRedis()})())

    class _FixedDate(date):
        @classmethod
        def today(cls):
            return date(2026, 9, 20)
    monkeypatch.setattr(vcs, 'date', _FixedDate)

    leads = vcs._fetch_visit_leads()
    assert any(l['리드 No'] == 'ETC-rangeend' for l in leads)


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
