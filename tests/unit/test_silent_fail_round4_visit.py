"""방문 일정 조용한 실패 점검(11-24) — 당일 배정이 '처리 중'에 묶이던 것, 방문봇이 없어 DM 을 통째로 건너뛰던 것,
확정 결과를 'done' 먼저 쓰고 뒤 단계를 붙여 화면이 못 보던 것. 이름·번호는 모두 가짜."""
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime, timedelta

import pytest

import dashboard.services.visit_assignment_sync as vas
import dashboard.utils.redis_client as rcmod
from dashboard.services import vehicle_assign as va
from dashboard.services import visit_plan as vp
from test_visit_plan import DAY, LEADS, STAFF, _NowThread, _lead, _RC, _Redis, db  # noqa: F401  (db = 픽스처)

TODAY = date.today().isoformat()


def _errors(monkeypatch, mod):
    """mod.logger.error 호출을 모은다 — [(글, force_alert)]."""
    got = []
    orig = mod.logger.error

    def fake(msg, *a, **k):
        got.append((str(msg), bool((k.get('extra') or {}).get('force_alert'))))
        return orig(msg, *a, **k)
    monkeypatch.setattr(mod.logger, 'error', fake)
    return got


def _urgent_env(monkeypatch, dm=None):
    leads = [{**_lead('L-7', '', '서준'), '방문 예정일': TODAY, '플랫폼': '기타'}]
    monkeypatch.setattr(vp, 'can_plan', lambda e: e == 'p@example.com')
    monkeypatch.setattr(va, 'online_duty', lambda d: [])
    wrote, sent = {}, []
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'update_lead', lambda no, vals: wrote.update({no: vals}) or {'success': True})
    monkeypatch.setattr(vas, '_update_slack_list_managers', lambda a, pm, sheet_leads=None: (1, 0, []))
    monkeypatch.setattr(vas, '_get_visit_client', lambda: None)
    monkeypatch.setattr(va, '_dm', dm or (lambda sid, text: sent.append((sid, text)) or True))
    import dashboard.services.visit_canvas_sync as vcs
    monkeypatch.setattr(vcs, 'rebuild_canvas_async', lambda: None)
    monkeypatch.setattr(vp, '_urgent_stale_noted', set())
    return leads, wrote, sent


def _urgent_row(lead_no='L-7'):
    with vp._conn() as conn:
        r = conn.execute('SELECT * FROM visit_plan_urgent WHERE day = ? AND lead_no = ?', (TODAY, lead_no)).fetchone()
    return dict(r) if r else None


# ── 1. 당일 배정 ─────────────────────────────────────────────────────────
def test_urgent_work_exception_marks_failed(db, monkeypatch):
    """try 밖 단계(DM 보낼 사람 찾기 등)에서 예외가 나도 'running' 에 묶이지 않고 failed + ERROR(force_alert, 리드 번호)."""
    leads, wrote, _ = _urgent_env(monkeypatch)
    monkeypatch.setattr(vp, '_slack_of', lambda code, name: (_ for _ in ()).throw(RuntimeError('명단 못 읽음')))
    errs = _errors(monkeypatch, vp)
    res = vp.urgent_assign(TODAY, 'L-7', ['TH'], 'p@example.com', leads, background=False)
    row = _urgent_row()
    assert row['state'] == 'failed'
    assert '시트엔 담당자를 적었지만' in res['message'] and res['dm'] == 0 and res['dm_failed'] == 1
    assert wrote == {'L-7': {'영업 담당자': '서준'}}
    assert any('L-7' in m and force for m, force in errs)
    # failed 는 화면에 '실패'로, 코스에 붙지 않음
    d = vp.load(TODAY, leads, 'p@example.com')
    assert d['urgent']['L-7']['state'] == 'failed'
    assert not any('L-7' in c['leads'] for c in d['plan']['courses'])


def test_urgent_work_exception_before_sheet(db, monkeypatch):
    """시트 전에 터지면 '다시 해 주세요'."""
    leads, _, _ = _urgent_env(monkeypatch)
    monkeypatch.setattr(vp, '_urgent_steps', lambda *a, **k: (_ for _ in ()).throw(RuntimeError('boom')))
    res = vp.urgent_assign(TODAY, 'L-7', ['TH'], 'p@example.com', leads, background=False)
    assert res['message'] == '당일 배정을 끝내지 못했어요. 다시 해 주세요.' and _urgent_row()['state'] == 'failed'


def _put_running(minutes_ago, by='jk@example.com'):
    at = (datetime.now() - timedelta(minutes=minutes_ago)).strftime('%Y-%m-%d %H:%M:%S')
    with vp._conn() as conn:
        conn.execute("INSERT INTO visit_plan_urgent (day, lead_no, members, by, at, state) VALUES (?,?,?,?,?,'running')",
                     (TODAY, 'L-7', 'YM', by, at))


def test_urgent_stale_running_can_retry(db, monkeypatch):
    """서버 재시작 등으로 끊긴 'running'(5분 넘음) — 화면은 'stale', 코스에 안 붙고, 다시 배정하면 지우고 새로. ERROR 는 한 번."""
    leads, wrote, sent = _urgent_env(monkeypatch)
    errs = _errors(monkeypatch, vp)
    _put_running(minutes_ago=10)
    d = vp.load(TODAY, leads, 'p@example.com')
    assert d['urgent']['L-7']['state'] == 'stale'
    assert not any('L-7' in c['leads'] for c in d['plan']['courses'])
    vp.load(TODAY, leads, 'p@example.com')
    stale_logs = [m for m, force in errs if '끊겼어요' in m and force]
    assert len(stale_logs) == 1 and 'L-7' in stale_logs[0]                  # 화면을 여러 번 열어도 한 번만
    res = vp.urgent_assign(TODAY, 'L-7', ['TH'], 'p@example.com', leads, background=False)
    assert res == {'names': '서준', 'dm': 1, 'dm_failed': 0}
    row = _urgent_row()
    assert row['state'] == 'done' and row['members'] == 'TH' and row['by'] == 'p@example.com'
    assert len([m for m, _ in errs if '끊겼어요' in m]) == 1                  # 다시 배정할 때 또 알리지 않음(같은 줄)


def test_urgent_fresh_running_still_blocks(db, monkeypatch):
    """아직 5분 안 된 'running' 은 그대로 막는다(두 사람이 동시에 누른 경우)."""
    leads, _, _ = _urgent_env(monkeypatch)
    _put_running(minutes_ago=1)
    assert vp.load(TODAY, leads, 'p@example.com')['urgent']['L-7']['state'] == 'running'
    with pytest.raises(ValueError, match='방금 다른 사람'):
        vp.urgent_assign(TODAY, 'L-7', ['TH'], 'p@example.com', leads, background=False)
    assert _urgent_row()['state'] == 'running' and _urgent_row()['members'] == 'YM'


def test_urgent_state_bad_time_is_stale():
    assert vp._urgent_state({'state': 'running', 'at': ''}) == 'stale'
    assert vp._urgent_state({'state': None}) == 'done'
    assert vp._urgent_state({'state': 'failed', 'at': ''}) == 'failed'


def test_urgent_insert_db_error_not_masked(db, monkeypatch):
    """INSERT 가 겹침(IntegrityError)이 아닌 DB 오류면 '다른 사람이 했어요'로 덮지 않고 로그 + 그대로 올린다."""
    leads, _, _ = _urgent_env(monkeypatch)
    errs = _errors(monkeypatch, vp)
    real = vp._conn

    class _Wrap:
        def __init__(self, c):
            self.c = c

        def execute(self, sql, *a):
            if sql.startswith('INSERT INTO visit_plan_urgent'):
                raise sqlite3.OperationalError('database is locked')
            return self.c.execute(sql, *a)

    @contextmanager
    def fake_conn():
        with real() as c:
            yield _Wrap(c)
    monkeypatch.setattr(vp, '_conn', fake_conn)
    with pytest.raises(sqlite3.OperationalError):
        vp.urgent_assign(TODAY, 'L-7', ['TH'], 'p@example.com', leads, background=False)
    assert any('당일 배정 기록 실패' in m for m, _ in errs)


# ── 2. 방문봇 client 없음 ────────────────────────────────────────────────
def test_no_visit_client_counts_failed(monkeypatch):
    """방문봇 client 가 없으면 보낼 사람 수만큼 실패로 세고 ERROR(force_alert) — 전엔 'DM 0명' 성공처럼 보였다."""
    fr = _Redis()
    monkeypatch.setattr(vas, '_get_visit_client', lambda: None)
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: _RC(fr))
    monkeypatch.setattr(vas, '_email_from_initial', lambda ini, itn: f'{ini.lower()}@example.com')
    errs = _errors(monkeypatch, vas)
    itn = {p['code']: p['name'] for p in STAFF}
    target = date.today() + timedelta(days=3)
    # TH 는 이전 DM 과 같음(보낼 필요 없음), LSH 는 빠짐(재배정 안내 대상)
    fr.set(f'dm_v9_msg:TH:{target.isoformat()}', '{"channel": "D1", "ts": "1", "lead_nos": ["L-3"], "memo": ""}')
    fr.set(f'dm_v9_msg:LSH:{target.isoformat()}', '{"channel": "D2", "ts": "2", "lead_nos": ["L-9"], "memo": ""}')
    a = [{'lead': _lead('L-1'), 'assign': ['YM', 'JK'], 'memo': ''}, {'lead': _lead('L-3'), 'assign': ['TH'], 'memo': ''}]
    r = vas._send_dms_for_next_visit(a, {}, itn, ['SJ'], [], addr_candidates=[], target_date=target)
    assert r['visit_bot_missing'] is True
    assert r['visit_mgr_failed'] == 2 and r['visit_mgr_sent'] == 0          # YM·JK (TH 는 같아서 뺌)
    assert r['deassign_failed'] == 1 and r['online_duty_failed'] == 1
    assert any('방문봇 연결이 없어' in m and force for m, force in errs)
    s = vp._summary({'dm': r})
    assert s['bot_missing'] is True and s['dm_failed'] == 2 and s['duty_failed'] == 1 and s['deassign_failed'] == 1


def test_summary_without_bot_flag():
    s = vp._summary({'dm': {'visit_mgr_sent': 2}})
    assert s['bot_missing'] is False and s['duty_failed'] == 0 and s['deassign_failed'] == 0


# ── 3. 확정 'done' 은 한 번, 뒤 단계까지 끝낸 뒤 ───────────────────────────
def _confirm_env(monkeypatch, seen):
    monkeypatch.setattr(vas, 'commit_plan', lambda parsed, target_date=None: {'ok': True, 'dm': {'visit_mgr_sent': 1}})
    monkeypatch.setattr(vp.threading, 'Thread', _NowThread)
    monkeypatch.setattr(vp, '_staff', lambda: [dict(p) for p in STAFF])
    finishes = []
    real = vp._finish
    monkeypatch.setattr(vp, '_finish', lambda day, state, result, by='', plan=None:
                        finishes.append(state) or real(day, state, result, by, plan))

    def req(day=None, by=''):
        seen.append(('req', vp.run_state(DAY)['state']))
        return 0

    def notice(day, text, by=''):
        seen.append(('notice', vp.run_state(DAY)['state']))
        return 2

    def audit(res, by):
        seen.append(('audit', vp.run_state(DAY)['state']))
        return [{'lead_no': 'L-9'}]
    monkeypatch.setattr(va, 'request_assignment', req)
    monkeypatch.setattr(va, 'notify_plan_changed', notice)
    monkeypatch.setattr(vp, '_audit', audit)
    return finishes


BASE = {'courses': [{'members': ['YM'], 'leads': ['L-1']}, {'members': ['TH'], 'leads': ['L-3']}],
        'online': [], 'online_shared': False}
CHANGED = {'courses': [{'members': ['YM', 'LSH'], 'leads': ['L-1']}, {'members': ['TH'], 'leads': ['L-3']}],
           'online': [], 'online_shared': False}


def test_confirm_done_written_once_after_all_steps(db, monkeypatch):
    seen = []
    finishes = _confirm_env(monkeypatch, seen)
    vp.start_confirm(DAY, BASE, 'p@example.com', LEADS, lambda: LEADS)
    seen.clear()
    finishes.clear()
    vp.start_confirm(DAY, CHANGED, 'p@example.com', LEADS, lambda: LEADS)    # 사람이 바뀜 → 차량 배정 변경 알림
    assert seen == [('req', 'running'), ('notice', 'running'), ('audit', 'running')]   # 그 사이엔 'running'
    assert finishes == ['done']                                                         # 'done' 은 한 번만
    st = vp.run_state(DAY)
    assert st['state'] == 'done' and st['result']['car_notice'] == 2 and st['result']['cards'] == [{'lead_no': 'L-9'}]
    assert not any(k in st['result'] for k in ('car_request_failed', 'car_notice_failed', 'cards_error'))
    assert vp.confirmed_info(DAY)['confirmed_at']


def test_confirm_after_steps_fail_are_marked(db, monkeypatch):
    """뒤 단계가 실패해도 'done' 은 남고, 실패한 단계는 결과에 표시 + ERROR(전엔 _audit 실패면 car_notice 도 DB 에 안 남음)."""
    seen = []
    _confirm_env(monkeypatch, seen)
    vp.start_confirm(DAY, BASE, 'p@example.com', LEADS, lambda: LEADS)
    boom = lambda *a, **k: (_ for _ in ()).throw(RuntimeError('boom'))
    monkeypatch.setattr(va, 'request_assignment', boom)
    monkeypatch.setattr(vp, '_audit', boom)
    errs = _errors(monkeypatch, vp)
    vp.start_confirm(DAY, CHANGED, 'p@example.com', LEADS, lambda: LEADS)
    st = vp.run_state(DAY)
    assert st['state'] == 'done'
    r = st['result']
    assert r['car_request_failed'] is True and r['cards_error'] is True and r['car_notice'] == 2
    assert 'cards' not in r
    assert sum('차량 배정 요청 DM 실패' in m or '누락 점검 실패' in m for m, _ in errs) == 2
    # 변경 알림만 실패
    monkeypatch.setattr(va, 'request_assignment', lambda day=None, by='': 0)
    monkeypatch.setattr(vp, '_audit', lambda res, by: [])
    monkeypatch.setattr(va, 'notify_plan_changed', boom)
    vp.start_confirm(DAY, BASE, 'p@example.com', LEADS, lambda: LEADS)
    r = vp.run_state(DAY)['result']
    assert r['car_notice_failed'] is True and 'car_notice' not in r and r['cards'] == []


def test_confirm_finish_failure_alerts(db, monkeypatch):
    """마지막 결과 저장이 실패하면 ERROR(force_alert) — 화면은 'running' 이 5분 뒤 끊김으로."""
    seen = []
    _confirm_env(monkeypatch, seen)
    real = vp._finish

    def fin(day, state, result, by='', plan=None):
        if state == 'done':
            raise sqlite3.OperationalError('disk I/O error')
        return real(day, state, result, by, plan)
    monkeypatch.setattr(vp, '_finish', fin)
    errs = _errors(monkeypatch, vp)
    vp.start_confirm(DAY, BASE, 'p@example.com', LEADS, lambda: LEADS)
    assert vp.run_state(DAY)['state'] == 'running'
    assert any('확정 결과 저장 실패' in m and force for m, force in errs)
