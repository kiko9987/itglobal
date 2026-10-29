# -*- coding: utf-8 -*-
"""3묶음 운영 감시 (2026-10-30) — 시트 쓰기 재시도 순서·유실, 예약 작업 감시, 인증서 알림 기준."""
import sys
sys.path.insert(0, '.')

import json
import time
from datetime import timedelta
from types import SimpleNamespace

import pytest

from dashboard.services import sheet_write_queue as q
from dashboard.services import sync_scheduler as ss


class _Redis:
    def __init__(self):
        self.lists = {}

    def rpush(self, k, v):
        self.lists.setdefault(k, []).append(v)

    def lrem(self, k, n, v):
        lst = self.lists.get(k, [])
        if v in lst:
            lst.remove(v)

    def lindex(self, k, i):
        lst = self.lists.get(k, [])
        return lst[i] if len(lst) > i else None


@pytest.fixture
def redis(monkeypatch):
    r = _Redis()
    monkeypatch.setattr(q, '_rc', lambda: r)
    monkeypatch.setattr(q, '_bump_processed_and_maybe_rollup', lambda: None)
    return r


def _op(op_type='t', created=None):
    op = {'op_id': 'abcdef12-0000', 'op_type': op_type, 'payload': {'tag': 'G1-XX'},
          'meta': {}, 'created_at': created or time.time(), 'attempts': 0, 'last_error': ''}
    return op, json.dumps(op)


def test_retry_in_place_keeps_op_in_processing_until_done(redis):
    """실패해도 processing 에 남은 채 그 자리에서 재시도 — 대기 중 재시작돼도 부팅 복구가 되살린다."""
    op, raw = _op()
    redis.rpush(q.PROCESSING_KEY, raw)
    calls, seen_processing = [], []

    def handler(p):
        calls.append(1)
        if len(calls) < 3:
            raise RuntimeError('일시 오류')

    def sleep(_s):
        seen_processing.append(raw in redis.lists[q.PROCESSING_KEY])
    assert q._run_op(op, raw, handler, sleep=sleep) is True
    assert len(calls) == 3 and seen_processing == [True, True]
    assert redis.lists[q.PROCESSING_KEY] == []
    assert q.QUEUE_KEY not in redis.lists          # 큐 맨 뒤로 다시 넣지 않음(순서 보존)


def test_deadletter_after_max_attempts(redis, monkeypatch):
    sent = []
    monkeypatch.setattr(q, '_notify_admin_deadletter', lambda op, exc: sent.append(op['attempts']))
    op, raw = _op()
    redis.rpush(q.PROCESSING_KEY, raw)

    def handler(p):
        raise RuntimeError('계속 실패')
    assert q._run_op(op, raw, handler, sleep=lambda s: None) is False
    assert redis.lists[q.PROCESSING_KEY] == []
    assert len(redis.lists[q.FAILED_KEY]) == 1 and sent == [q.MAX_ATTEMPTS]


def test_oldest_pending_age_uses_requeued_at(redis):
    old, raw = _op(created=time.time() - 3600)
    redis.rpush(q.QUEUE_KEY, raw)
    assert 3500 < q.oldest_pending_age() < 3700
    old['requeued_at'] = time.time() - 5
    redis.lists[q.QUEUE_KEY] = [json.dumps(old)]
    assert q.oldest_pending_age() < 60


def test_oldest_pending_age_empty(redis):
    assert q.oldest_pending_age() == 0.0


# ── 예약 작업 감시 ──
class _Sched:
    def __init__(self, jobs):
        self.jobs = jobs

    def get_jobs(self):
        return self.jobs

    def get_job(self, jid):
        return next((j for j in self.jobs if j.id == jid), None)


def _job(jid, seconds=None, trigger=None):
    return SimpleNamespace(id=jid, trigger=trigger or SimpleNamespace(interval=timedelta(seconds=seconds)))


@pytest.fixture
def sched(monkeypatch):
    sent = []
    monkeypatch.setattr(ss, '_notify_admin', lambda key, text: sent.append((key, text)))
    monkeypatch.setattr(ss, '_notify_admin_bg', lambda key, text: sent.append((key, text)))
    monkeypatch.setattr(ss, '_job_last_done', {})
    monkeypatch.setattr(ss, '_job_skip_streak', {})
    monkeypatch.setattr(ss, '_fail_streaks', {})
    return sent


def test_stale_interval_jobs(monkeypatch, sched):
    now = time.time()
    monkeypatch.setattr(ss, '_scheduler', _Sched([_job('payment_sync', 30), _job('refund_match_scan', 600)]))
    monkeypatch.setattr(ss, '_scheduler_started_at', now - 3600)
    ss._job_last_done.update({'payment_sync': now - 20, 'refund_match_scan': now - 40 * 60})
    assert ss._stale_interval_jobs(now) == [('refund_match_scan', 40)]
    ss._job_last_done['refund_match_scan'] = now - 25 * 60          # 주기 3배(30분) 안
    assert ss._stale_interval_jobs(now) == []


def test_max_instances_streak_alerts_then_resets(monkeypatch, sched):
    from apscheduler.events import EVENT_JOB_EXECUTED, EVENT_JOB_MAX_INSTANCES
    monkeypatch.setattr(ss, '_scheduler', _Sched([_job('payment_sync', 30), _job('workflow_phone_sync', 10)]))
    ev = lambda code, j='payment_sync': SimpleNamespace(code=code, job_id=j)
    for _ in range(3):
        ss._on_job_event(ev(EVENT_JOB_MAX_INSTANCES))
    assert sched == []                                               # 30초×3 = 1분 반 — 아직 조용
    ss._on_job_event(ev(EVENT_JOB_MAX_INSTANCES))                    # 2분
    assert sched and sched[0][0] == 'job_stuck_payment_sync'
    for _ in range(11):                                              # 10초 주기는 12회(2분)부터
        ss._on_job_event(ev(EVENT_JOB_MAX_INSTANCES, 'workflow_phone_sync'))
    assert len(sched) == 1
    ss._on_job_event(ev(EVENT_JOB_MAX_INSTANCES, 'workflow_phone_sync'))
    assert len(sched) == 2
    ss._on_job_event(ev(EVENT_JOB_EXECUTED))
    assert 'payment_sync' not in ss._job_skip_streak and 'payment_sync' in ss._job_last_done


def test_missed_cron_job_alerts_but_not_interval(monkeypatch, sched):
    from apscheduler.events import EVENT_JOB_MISSED
    from apscheduler.triggers.cron import CronTrigger
    monkeypatch.setattr(ss, '_scheduler', _Sched([_job('daily_backup', trigger=CronTrigger(hour=3, minute=15)),
                                                  _job('payment_sync', 30)]))
    ss._on_job_event(SimpleNamespace(code=EVENT_JOB_MISSED, job_id='payment_sync'))
    assert sched == []
    ss._on_job_event(SimpleNamespace(code=EVENT_JOB_MISSED, job_id='daily_backup', scheduled_run_time='03:15'))
    assert sched and sched[0][0] == 'job_missed_daily_backup'


def test_fail_streak(sched):
    for _ in range(2):
        ss._fail_streak('refund_match', True, '반환 짝짓기 실패')
    assert sched == []
    ss._fail_streak('refund_match', True, '반환 짝짓기 실패')
    assert len(sched) == 1
    ss._fail_streak('refund_match', False, '')
    assert 'refund_match' not in ss._fail_streaks


def test_ops_watch_queue_alerts(monkeypatch, sched):
    monkeypatch.setattr(ss, '_scheduler', _Sched([]))
    monkeypatch.setattr(q, 'worker_alive', lambda: False)
    monkeypatch.setattr(q, 'oldest_pending_age', lambda: 900.0)
    monkeypatch.setattr(q, 'get_stats', lambda: {'pending': 4, 'processing': 1, 'failed': 0})
    ss._safe_ops_watch()
    assert {k for k, _ in sched} == {'sheet_queue_worker_dead', 'sheet_queue_stale'}


def test_stagger_offsets_avoid_payment_poll_ticks():
    """무거운 10분 작업 첫 실행은 30초 입금 폴러 회차 사이(15초 어긋남)."""
    import inspect
    src = inspect.getsource(ss.start_scheduler)
    import re
    for m in re.findall(r"_stagger\(([\d.]+)\)", src):
        assert (float(m) * 60) % 30 == 15, m


def test_cert_alert_days_relative_to_lifetime():
    c90 = {'notBefore': 'Oct  1 00:00:00 2026 GMT', 'notAfter': 'Dec 30 00:00:00 2026 GMT'}
    c45 = {'notBefore': 'Oct  1 00:00:00 2026 GMT', 'notAfter': 'Nov 15 00:00:00 2026 GMT'}
    assert ss._cert_alert_days(c90) == 18
    assert ss._cert_alert_days(c45) == 9
    assert ss._cert_alert_days({}) == 18


def test_invariant_sheet_write_failed_reminds_daily():
    from datetime import datetime
    from dashboard.services.invariant_checks import check_sheet_write_failed
    op = {'op_id': 'x1', 'op_type': 'project_update_sheet', 'payload': {'project_code': 'G4100-SJ'},
          'created_at': time.time(), 'last_error': '503'}
    v1 = check_sheet_write_failed({'queue_failed': [op], 'now': datetime(2026, 10, 30)})
    v2 = check_sheet_write_failed({'queue_failed': [op], 'now': datetime(2026, 10, 31)})
    assert len(v1) == 1 and v1[0]['title'] == 'G4100-SJ'
    assert v1[0]['key'] != v2[0]['key']                  # 날마다 다시 알림
    assert check_sheet_write_failed({'queue_failed': []}) == []


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
