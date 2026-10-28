# -*- coding: utf-8 -*-
"""2026-10-29 점검 — 조용한 유실·알림 누락·새해 날짜."""
import logging
import sys
from datetime import date

sys.path.insert(0, '.')

import dashboard.blueprints.sms_inbound as si
from dashboard.services.sms_intake import year_for_md
from dashboard.utils.error_slack_alerter import _AlertGate, ALWAYS_ALERT


def _rec(msg, **extra):
    r = logging.LogRecord(name='dashboard.services.x', level=logging.ERROR, pathname='x.py',
                          lineno=10, msg=msg, args=(), exc_info=None, func='f')
    for k, v in extra.items():
        setattr(r, k, v)
    return r


def test_force_alert_bypasses_transient_suppression():
    g = _AlertGate(now=lambda: 1000.0, redis_getter=None)
    msg = '[LEADS] 리드 L-00001 업데이트 실패: 최대 재시도 횟수 초과 HTTP 502'
    assert g.decide(_rec(msg)) is None                    # 예전처럼 한 번뿐이면 묻힘
    assert g.decide(_rec(msg + ' ', **ALWAYS_ALERT)) is not None


def test_year_for_md_new_year():
    assert year_for_md(12, 31, today=date(2027, 1, 4)) == 2026      # 12/31 입금을 1/4 에 기록
    assert year_for_md(1, 2, today=date(2027, 1, 4)) == 2027
    assert year_for_md(10, 28, today=date(2026, 10, 29)) == 2026
    assert year_for_md(10, 30, today=date(2026, 10, 29)) == 2026     # 하루 여유(자정 무렵)
    assert year_for_md(2, 30, today=date(2026, 10, 29)) == 2026      # 없는 날짜 → 올해 그대로


def test_visit_range_crossing_new_year():
    from dashboard.services.visit_assignment_sync import _parse_visit_date_end
    assert _parse_visit_date_end('2026-12-30~01-02') == date(2027, 1, 2)
    assert _parse_visit_date_end('2026-10-23~28') == date(2026, 10, 28)
    assert _parse_visit_date_end('2026-12-30~2027-01-02') == date(2027, 1, 2)
    assert _parse_visit_date_end('2026-10-24~20') == date(2026, 10, 20)     # 같은 달 오기 — 경고가 잡도록 그대로


class _FakeRedis:
    def __init__(self):
        self.kv = {}

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.kv:
            return None
        self.kv[k] = v
        return True

    def get(self, k):
        return self.kv.get(k)

    def delete(self, k):
        self.kv.pop(k, None)


DEP = '[Web발신]\n하나,10/02, 15:00\n255******31304\n입금3,014,000원\n잔액12,345,678원\n대연이엔지주식회'


def test_card_failure_releases_dedup_so_other_phone_can_retry(monkeypatch):
    fr = _FakeRedis()
    monkeypatch.setattr(si, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    calls = []
    monkeypatch.setattr(si, '_post_intake_card', lambda *a, **k: calls.append(1) or len(calls) > 1)
    first = si.ingest_deposit(DEP, source='sms:sb')
    assert first['status'] == 'card_failed'
    second = si.ingest_deposit(DEP, source='sms:jw')       # 다른 폰이 같은 문자 → 이번엔 카드
    assert second['status'] == 'ok' and len(calls) == 2
    assert si.ingest_deposit(DEP, source='sms:yg')['status'] == 'duplicate'
