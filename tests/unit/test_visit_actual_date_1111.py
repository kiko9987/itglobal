"""방문 완료 시 실제 방문일 반영 (2026-11-11).

- 예정 기간 안 완료 → 완료일로 좁힘 (기존 2026-09-30 규칙)
- 예정일보다 먼저 완료 → 완료일로 (앞당겨 다녀온 것, L-04286)
- 예정보다 늦은 완료 → 그대로 (늦게 간 것인지 버튼만 늦은 것인지 모름)
"""
import sys
from datetime import datetime as _real_dt

import pytest

import dashboard.blueprints.slack_bot as sb


class _FixedDT(_real_dt):
    @classmethod
    def now(cls, tz=None):
        return cls(2026, 10, 19, 15, 0)


def _run(monkeypatch, visit_date):
    writes = []
    lead = {'리드 No': 'L-09999', '상담 내용': '[10.15 10:00 YG · 방문 예약] 테스트',
            '상태': '방문 예약', '방문 예정일': visit_date}
    monkeypatch.setattr(sb, 'datetime', _FixedDT)
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda _no: dict(lead))
    monkeypatch.setattr(sb, '_update_lead_dispatch', lambda no, upd: writes.append(upd))
    sb._mark_visit_complete_on_sheet('L-09999', initial='MW', dt_str='10.19 15:00')
    assert writes and writes[0]['상태'] == '방문 완료'
    return writes[0].get('방문 예정일')


def test_completed_before_planned_date_uses_completion_day(monkeypatch):
    assert _run(monkeypatch, '2026-10-20') == "'2026-10-19"
    assert _run(monkeypatch, '2026-10-21~23') == "'2026-10-19"


def test_completed_within_range_narrows(monkeypatch):
    assert _run(monkeypatch, '2026-10-18~21') == "'2026-10-19"


def test_same_day_or_late_completion_unchanged(monkeypatch):
    assert _run(monkeypatch, '2026-10-19') is None
    assert _run(monkeypatch, '2026-10-17') is None          # 늦은 완료 — 그대로
    assert _run(monkeypatch, '2026-10-15~17') is None


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
