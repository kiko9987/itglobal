from datetime import datetime
from dashboard.services import visit_plan as vp, vehicle_assign as va


def test_past_auto_send(monkeypatch):
    # 11-25(수) 16:51 확정 → 다음 근무일 11-26 의 16:30 자동 발송은 이미 지남
    assert vp._past_auto_send('2026-11-26', datetime(2026, 11, 25, 16, 51))
    assert not vp._past_auto_send('2026-11-26', datetime(2026, 11, 25, 16, 29))
    assert not vp._past_auto_send('2026-11-27', datetime(2026, 11, 25, 17, 0))     # 모레 일정은 내일 16:30 에
    assert vp._past_auto_send('2026-11-26', datetime(2026, 11, 26, 8, 0))          # 당일


def test_assign_editors(monkeypatch):
    from dashboard.services import hr
    monkeypatch.setattr(hr, 'is_admin_email', lambda e: e == 'boss@example.com')
    monkeypatch.setattr(va, 'assign_editors', lambda: ['helper@example.com'])
    assert va.can_assign('boss@example.com') and va.can_assign('Helper@example.com') and not va.can_assign('x@example.com')
