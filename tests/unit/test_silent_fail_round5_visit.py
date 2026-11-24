"""방문 일정 조용한 실패 점검 5차(11-24) — 차량 배정 DM 실패 수, 슬랙 List 를 못 고친 이유·List 에 없는 리드,
당일 배정 List 결과, 온라인 당번 저장 실패, 확정 스레드 바깥 try·끊김 알림, 사장님 DM 실패, Redis DM 기록,
당번 DM·재배정 안내 실패 수, 코스에서 뺀 현장(담당자 그대로)·'재배정' 안내 문구. 이름·번호는 모두 가짜."""
import json
from datetime import date, datetime, timedelta

import dashboard.services.visit_assignment_sync as vas
import dashboard.utils.redis_client as rcmod
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import visit_canvas_audit as vca
from dashboard.services import visit_plan as vp
from test_silent_fail_round4_visit import BASE, TODAY, _confirm_env, _errors, _urgent_env, _urgent_row
from test_visit_plan import DAY, LEADS, STAFF, _Client, _lead, _NowThread, _RC, _Redis, db  # noqa: F401  (db = 픽스처)

ASSIGNER = 'sj@example.com'


def _warnings(monkeypatch, mod):
    got = []
    orig = mod.logger.warning

    def fake(msg, *a, **k):
        got.append(str(msg))
        return orig(msg, *a, **k)
    monkeypatch.setattr(mod.logger, 'warning', fake)
    return got


# ── 1. 차량 배정 요청·변경 알림 — 못 보낸 수 ────────────────────────────────
def _car_env(monkeypatch, ok=False):
    monkeypatch.setattr(va, 'CAR_DM_ON', True)
    monkeypatch.setattr(va, 'assigners', lambda: [ASSIGNER])
    monkeypatch.setattr(hr, 'is_admin_email', lambda e: e == ASSIGNER)     # 배정 DM 은 관리자인 배정 담당에게만(11-25)
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'slack_id': 'USJ'} if e == ASSIGNER else None)
    sent = []
    monkeypatch.setattr(va, '_dm', lambda sid, text: sent.append((sid, text)) or ok)
    return sent


def test_sent_count_is_int():
    s = va.SentCount(2, ['a@example.com'])
    assert s == 2 and json.dumps({'n': s}) == '{"n": 2}' and va.failed_of(s) == ['a@example.com']
    assert va.failed_of(0) == [] and va.failed_of(None) == []          # 가짜(int)·None 도 그대로


def test_request_assignment_counts_failed(db, monkeypatch):
    """배정 담당에게 다 못 보내면 0 + failed + ERROR, 요청 시각은 안 남김(다음 확정 때 다시)."""
    sent = _car_env(monkeypatch, ok=False)
    errs = _errors(monkeypatch, va)
    r = va.request_assignment(DAY)
    assert r == 0 and va.failed_of(r) == [ASSIGNER] and len(sent) == 1
    assert any('배정 담당 DM 못 보냄' in m for m, _ in errs)
    assert not (va._day_info(DAY) or {}).get('requested_at')
    # 보낼 필요가 없는 0(배정 담당이 확정) — failed 없음
    r = va.request_assignment(DAY, ASSIGNER)
    assert r == 0 and va.failed_of(r) == []


def test_assign_dm_skips_non_admin_assigner(db, monkeypatch):
    """배정 담당 명단에 있어도 관리자가 아니면(고칠 수 없으니) '차량 배정해 주세요' DM 을 보내지 않는다(11-25)."""
    sent = _car_env(monkeypatch, ok=True)
    monkeypatch.setattr(va, 'assigners', lambda: [ASSIGNER, 'staff9@example.com'])
    monkeypatch.setattr(hr, 'find_by_email', lambda e: {'slack_id': 'U' + e[:3]})
    assert va.request_assignment(DAY) == 1 and [s for s, _ in sent] == ['Usj@']


def test_notify_plan_changed_no_target_vs_failed(db, monkeypatch):
    """차량 배정을 안 한 날의 0(대상 없음)과 다 실패한 0 을 구별."""
    _car_env(monkeypatch, ok=False)
    monkeypatch.setattr(va, '_day_rows', lambda day: {})
    r = va.notify_plan_changed(DAY, '추가 서준')
    assert r == 0 and va.failed_of(r) == []
    monkeypatch.setattr(va, '_day_rows', lambda day: {'11가1111': {'people': ['TH']}})
    r = va.notify_plan_changed(DAY, '추가 서준')
    assert r == 0 and va.failed_of(r) == [ASSIGNER]


def test_confirm_marks_car_dm_failed(db, monkeypatch):
    seen = []
    _confirm_env(monkeypatch, seen)
    vp.start_confirm(DAY, BASE, 'p@example.com', LEADS, lambda: LEADS)
    monkeypatch.setattr(va, 'request_assignment', lambda day=None, by='': va.SentCount(0, [ASSIGNER]))
    monkeypatch.setattr(va, 'notify_plan_changed', lambda day, text, by='': va.SentCount(0, [ASSIGNER]))
    changed = {'courses': [{'members': ['YM', 'LSH'], 'leads': ['L-1']}, {'members': ['TH'], 'leads': ['L-3']}],
               'online': [], 'online_shared': False}
    vp.start_confirm(DAY, changed, 'p@example.com', LEADS, lambda: LEADS)
    r = vp.run_state(DAY)['result']
    assert r['car_request_failed'] is True and r['car_request_failed_n'] == 1
    assert r['car_notice'] == 0 and r['car_notice_failed'] is True and r['car_notice_failed_n'] == 1


# ── 2. 슬랙 List 담당자 칸 ──────────────────────────────────────────────
class _Resp:
    def __init__(self, data):
        self.data = data


class _ListClient:
    """slackLists.items.list 를 쪽마다 — 첫 쪽 L-1, 다음 쪽 L-2."""
    def __init__(self, pages=None, boom=False):
        self.calls, self.boom = [], boom
        self.pages = pages or [
            {'ok': True, 'items': [{'id': 'Rec1', 'fields': [{'column_id': vas._LIST_COL_LEAD, 'text': 'L-1'}]}],
             'response_metadata': {'next_cursor': 'c2'}},
            {'ok': True, 'items': [{'id': 'Rec2', 'fields': [{'column_id': vas._LIST_COL_LEAD, 'text': 'L-2'}]}],
             'response_metadata': {'next_cursor': ''}},
        ]

    def api_call(self, method, http_verb='GET', params=None):
        self.calls.append(dict(params or {}))
        if self.boom:
            raise RuntimeError('ratelimited')
        return _Resp(self.pages[len(self.calls) - 1])


def _list_env(monkeypatch, client):
    monkeypatch.setenv('SLACK_VISIT_LIST_ID', 'F1')
    monkeypatch.setenv('SLACK_VISIT_BOT_TOKEN', 'x')
    monkeypatch.setattr(vas, '_get_visit_client', lambda: client)
    monkeypatch.setattr(vas, '_get_list_manager_option_map', lambda lid: {'YG': 'O1'})
    monkeypatch.setattr(vas, '_load_initial_maps', lambda: ({'YG': '임산'}, {}))
    monkeypatch.setattr(vas, '_load_addr_candidates', lambda: [])
    monkeypatch.setattr(vas, '_resolve_lead_for_assignment', lambda p, pm, ac: {'리드 No': p['lead']})
    sent = []

    class _U:
        def __init__(self, body):
            sent.append(body)

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

        def read(self):
            return b'{"ok": true}'
    import urllib.request
    monkeypatch.setattr(urllib.request, 'urlopen', lambda req, timeout=10: _U(json.loads(req.data.decode())))
    return sent


def test_list_pages_with_cursor_and_counts_not_found(monkeypatch):
    """cursor 로 끝까지 읽어 둘째 쪽 리드도 고치고, List 에 없는 리드는 .not_found 로(전엔 조용히 건너뜀)."""
    c = _ListClient()
    sent = _list_env(monkeypatch, c)
    res = vas._update_slack_list_managers(
        [{'lead': 'L-1', 'assign': ['YG']}, {'lead': 'L-2', 'assign': ['YG']}, {'lead': 'L-3', 'assign': ['YG']}],
        {}, sheet_leads=[{}])
    ok, fail, missing = res                                 # 3-튜플 그대로 풀림
    assert (ok, fail, missing) == (2, 0, [])
    assert vas.list_result_info(res) == ('', ['L-3'])
    assert [b['cells'][0]['row_id'] for b in sent] == ['Rec1', 'Rec2']
    assert len(c.calls) == 2 and c.calls[1]['cursor'] == 'c2' and 'cursor' not in c.calls[0]


def test_list_skipped_reasons_are_errors(monkeypatch):
    """통째로 못 고친 경우마다 이유 + ERROR — 전엔 모두 (0,0,[]) 라 성공과 구별이 안 됐다."""
    errs = _errors(monkeypatch, vas)
    a = [{'lead': 'L-1', 'assign': ['YG']}]
    monkeypatch.delenv('SLACK_VISIT_LIST_ID', raising=False)
    assert vas.list_result_info(vas._update_slack_list_managers(a, {}))[0] == 'List 설정 없음'
    _list_env(monkeypatch, None)
    assert vas.list_result_info(vas._update_slack_list_managers(a, {}))[0] == '방문봇 연결 없음'
    _list_env(monkeypatch, _ListClient())
    monkeypatch.setattr(vas, '_get_list_manager_option_map', lambda lid: {})
    assert vas.list_result_info(vas._update_slack_list_managers(a, {}))[0] == '담당자 선택지 조회 실패'
    _list_env(monkeypatch, _ListClient(boom=True))
    res = vas._update_slack_list_managers(a, {})
    assert tuple(res) == (0, 0, []) and vas.list_result_info(res)[0] == 'List 항목 조회 실패'
    assert len(errs) == 4
    _list_env(monkeypatch, _ListClient(pages=[{'ok': False, 'error': 'invalid_auth'}]))
    assert vas.list_result_info(vas._update_slack_list_managers(a, {}))[0] == 'List 항목 조회 실패'


def test_list_info_plain_tuple():
    assert vas.list_result_info((1, 0, [])) == ('', [])


def test_apply_assignments_carries_list_info(monkeypatch):
    import dashboard.services.lead_service as ls
    import dashboard.services.visit_canvas_sync as vcs
    monkeypatch.setattr(vas, '_load_initial_maps', lambda: ({p['code']: p['name'] for p in STAFF}, {}))
    monkeypatch.setattr(ls, 'update_lead', lambda no, vals: {'success': True})
    monkeypatch.setattr(vcs, '_fetch_visit_leads', lambda: [])
    monkeypatch.setattr(vcs, 'rebuild_canvas_async', lambda: None)
    monkeypatch.setattr(vas, '_update_slack_list_managers', lambda a, pm, sheet_leads=None:
                        vas.ListUpdateResult(1, 0, [], not_found=['L-3']))
    monkeypatch.setattr(vas, '_send_dms_for_next_visit', lambda *a, **k: {'dm_history_unread': True})
    res = vas._apply_assignments({'assignments': [{'lead': _lead('L-1'), 'assign': ['YM']}],
                                  'online_duty': [], 'off_duty': []})
    assert res['list_skipped'] == '' and res['list_not_found'] == ['L-3']
    s = vp._summary(res)
    assert s['list_not_found'] == ['L-3'] and s['list_skipped'] == '' and s['dm_history_unread'] is True


# ── 3. 당일 배정 — List 결과를 버리지 않음 ─────────────────────────────────
def test_urgent_list_skipped_shown(db, monkeypatch):
    leads, _, _ = _urgent_env(monkeypatch)
    monkeypatch.setattr(vas, '_update_slack_list_managers', lambda a, pm, sheet_leads=None:
                        vas.ListUpdateResult(0, 0, [], skipped='방문봇 연결 없음'))
    errs = _errors(monkeypatch, vp)
    res = vp.urgent_assign(TODAY, 'L-7', ['TH'], 'p@example.com', leads, background=False)
    assert res['list_failed'] == '방문봇 연결 없음' and res['dm'] == 1
    row = _urgent_row()
    assert row['state'] == 'done' and json.loads(row['result'])['list_failed'] == '방문봇 연결 없음'
    assert any('슬랙 List 담당자 못 고침' in m and 'L-7' in m for m, _ in errs)


def test_urgent_list_not_found_missing_and_error(db, monkeypatch):
    leads, _, _ = _urgent_env(monkeypatch)
    monkeypatch.setattr(vas, '_update_slack_list_managers', lambda a, pm, sheet_leads=None:
                        vas.ListUpdateResult(0, 0, [], not_found=['L-7']))
    assert vp.urgent_assign(TODAY, 'L-7', ['TH'], 'p@example.com', leads, background=False)['list_failed'] \
        == 'List 에 이 현장 행이 없음'
    with vp._conn() as conn:
        conn.execute('DELETE FROM visit_plan_urgent')
    monkeypatch.setattr(vas, '_update_slack_list_managers', lambda a, pm, sheet_leads=None:
                        (0, 0, [{'lead_no': 'L-7', 'names': ['서준']}]))
    assert '선택지 없음(서준)' in vp.urgent_assign(TODAY, 'L-7', ['TH'], 'p@example.com', leads, background=False)['list_failed']
    with vp._conn() as conn:
        conn.execute('DELETE FROM visit_plan_urgent')
    monkeypatch.setattr(vas, '_update_slack_list_managers', lambda *a, **k: (_ for _ in ()).throw(RuntimeError('x')))
    import dashboard.services.visit_canvas_sync as vcs
    monkeypatch.setattr(vcs, 'rebuild_canvas_async', lambda: (_ for _ in ()).throw(RuntimeError('y')))
    res = vp.urgent_assign(TODAY, 'L-7', ['TH'], 'p@example.com', leads, background=False)
    assert res['list_failed'] == '오류' and res['canvas_failed'] is True and _urgent_row()['state'] == 'done'


def test_urgent_list_ok_no_flag(db, monkeypatch):
    leads, _, _ = _urgent_env(monkeypatch)          # 가짜 = 보통 3-튜플 (1,0,[])
    res = vp.urgent_assign(TODAY, 'L-7', ['TH'], 'p@example.com', leads, background=False)
    assert res == {'names': '서준', 'dm': 1, 'dm_failed': 0}


# ── 5. 온라인 당번 저장 실패 ─────────────────────────────────────────────
def test_confirm_duty_save_failed(db, monkeypatch):
    seen = []
    _confirm_env(monkeypatch, seen)
    monkeypatch.setattr(va, 'record_online_duty', lambda *a, **k: (_ for _ in ()).throw(RuntimeError('locked')))
    errs = _errors(monkeypatch, vp)
    vp.start_confirm(DAY, BASE, 'p@example.com', LEADS, lambda: LEADS)
    st = vp.run_state(DAY)
    assert st['state'] == 'done' and st['result']['duty_save_failed'] is True
    assert any('온라인 당번 저장 실패' in m for m, _ in errs)


# ── 6. 확정 스레드 바깥 try · 끊김 알림 ─────────────────────────────────────
def test_confirm_outer_try_after_commit(db, monkeypatch):
    """시트·DM 뒤 정리(_summary 등)에서 예외 → failed + ERROR(force_alert), 기록 없이 끊기지 않음."""
    seen = []
    _confirm_env(monkeypatch, seen)
    monkeypatch.setattr(vp, '_summary', lambda res: (_ for _ in ()).throw(KeyError('x')))
    errs = _errors(monkeypatch, vp)
    vp.start_confirm(DAY, BASE, 'p@example.com', LEADS, lambda: LEADS)
    st = vp.run_state(DAY)
    assert st['state'] == 'failed' and '시트·DM 은 끝났지만' in st['result']['message']
    assert any('도중에 멈췄어요' in m and '시트·DM 뒤' in m and force for m, force in errs)


def test_confirm_outer_try_before_commit(db, monkeypatch):
    """이전 확정 읽기(try 밖)에서 예외 → failed 기록 시도 + ERROR. 기록도 못 하면 ERROR 만."""
    finishes = []
    monkeypatch.setattr(vp, '_row', lambda day: (_ for _ in ()).throw(RuntimeError('db')))
    monkeypatch.setattr(vp, '_finish', lambda day, state, result, by='', plan=None: finishes.append((state, result)))
    errs = _errors(monkeypatch, vp)
    vp._run_confirm(DAY, BASE, 'p@example.com', lambda: LEADS)
    assert finishes == [('failed', {'message': '확정하지 못했어요. 잠시 뒤 다시 해 주세요.'})]
    assert any('시트·DM 전' in m and force for m, force in errs)
    monkeypatch.setattr(vp, '_finish', lambda *a, **k: (_ for _ in ()).throw(RuntimeError('db2')))
    vp._run_confirm(DAY, BASE, 'p@example.com', lambda: LEADS)
    assert any('확정 실패 기록도 못 함' in m for m, _ in errs)


def test_confirm_value_error_logged(db, monkeypatch):
    monkeypatch.setattr(vas, 'commit_plan', lambda parsed, target_date=None: {'ok': False, 'reason': '⏳ 처리 중입니다'})
    monkeypatch.setattr(vp.threading, 'Thread', _NowThread)
    warns = _warnings(monkeypatch, vp)
    vp.start_confirm(DAY, BASE, 'p@example.com', LEADS, lambda: LEADS)
    st = vp.run_state(DAY)
    assert st['state'] == 'failed' and st['result']['message'] == '처리 중입니다'
    assert any('확정 안 됨' in m and '처리 중입니다' in m for m in warns)


def test_confirm_stale_alerts_once(db, monkeypatch):
    monkeypatch.setattr(vp, '_confirm_stale_noted', set())
    errs = _errors(monkeypatch, vp)
    vp.save(DAY, BASE, 'p@example.com', LEADS)
    at = (datetime.now() - timedelta(minutes=10)).strftime('%Y-%m-%d %H:%M:%S')
    with vp._conn() as conn:
        conn.execute("UPDATE visit_plan_day SET run_state = 'running', run_at = ?, run_by = 'p@example.com' WHERE day = ?",
                     (at, DAY))
    assert vp.run_state(DAY)['state'] == 'stale'
    assert vp.load(DAY, LEADS, 'p@example.com')['run']['state'] == 'stale'
    stale = [m for m, force in errs if '확정이 처리 중에 끊겼어요' in m and force]
    assert len(stale) == 1 and DAY in stale[0] and '박기획' in stale[0]


# ── 7. 사장님 DM 실패 ─────────────────────────────────────────────────────
def test_audit_admin_dm_failed(monkeypatch):
    monkeypatch.setattr(vca, 'check_cards_vs_canvas1', lambda today=None: [{'lead_no': 'L-9', 'name': '-', 'reason': 'x'}])
    monkeypatch.setattr(vca, 'send_admin_dm', lambda t: False)
    monkeypatch.setattr(vp, '_name_of_email', lambda e: e)
    errs = _errors(monkeypatch, vp)
    cards = vp._audit({'list_missing': []}, 'p@example.com')
    assert cards == [{'lead_no': 'L-9', 'name': '-', 'reason': 'x'}] and cards.admin_dm_failed is True
    assert any('사장님께 DM 으로 못 보냈어요' in m for m, _ in errs)
    monkeypatch.setattr(vca, 'send_admin_dm', lambda t: True)
    assert vp._audit({'list_missing': []}, 'p@example.com').admin_dm_failed is False


def test_confirm_admin_dm_failed_in_result(db, monkeypatch):
    seen = []
    _confirm_env(monkeypatch, seen)
    cards = vp._Cards([{'lead_no': 'L-9'}])
    cards.admin_dm_failed = True
    monkeypatch.setattr(vp, '_audit', lambda res, by: cards)
    vp.start_confirm(DAY, BASE, 'p@example.com', LEADS, lambda: LEADS)
    r = vp.run_state(DAY)['result']
    assert r['cards'] == [{'lead_no': 'L-9'}] and r['admin_dm_failed'] is True


def test_send_admin_dm_unconfigured_is_error(monkeypatch):
    monkeypatch.delenv('VISIT_AUDIT_ALERT_CHANNEL', raising=False)
    monkeypatch.delenv('SLACK_ADMIN_CHANNEL', raising=False)
    got = []
    monkeypatch.setattr(vca.logger, 'error', lambda msg, *a, **k: got.append(str(msg)))
    assert vca.send_admin_dm('x') is False and any('미설정' in m for m in got)


# ── 8·9a·10. DM — Redis 기록, 당번·재배정 안내 실패 수, '재배정' 문구 ────────────
class _FailClient(_Client):
    def __init__(self, fail=()):
        super().__init__()
        self.fail = set(fail)

    def chat_postMessage(self, channel, text, **k):
        if channel in self.fail:
            self.posted.append({'channel': channel, 'text': text, 'failed': True})
            return {'ok': False, 'error': 'channel_not_found'}
        return super().chat_postMessage(channel, text, **k)


def _dm_env(monkeypatch, client, redis=None, email=None):
    monkeypatch.setattr(vas, '_get_visit_client', lambda: client)
    if redis is None:
        monkeypatch.setattr(rcmod, 'get_redis_client', lambda: (_ for _ in ()).throw(ConnectionError('redis down')))
    else:
        monkeypatch.setattr(rcmod, 'get_redis_client', lambda: _RC(redis))
    monkeypatch.setattr(vas, '_email_from_initial', email or (lambda ini, itn: f'{ini.lower()}@example.com'))
    monkeypatch.setattr(vas, '_get_visit_card_permalink', lambda c, n: '')
    return {p['code']: p['name'] for p in STAFF}


def test_redis_down_marks_history_unread(monkeypatch):
    fc = _Client()
    itn = _dm_env(monkeypatch, fc, redis=None)
    warns = _warnings(monkeypatch, vas)
    target = date.today() + timedelta(days=3)
    r = vas._send_dms_for_next_visit([{'lead': _lead('L-1'), 'assign': ['YM'], 'memo': ''}], {}, itn, [], [],
                                     addr_candidates=[], target_date=target)
    assert r['dm_history_unread'] is True and r['visit_mgr_sent'] == 1
    assert any('이전 DM 기록을 못 읽어' in m for m in warns)
    assert vp._summary({'dm': r})['dm_history_unread'] is True


def test_redis_scan_fail_marks_history_unread(monkeypatch):
    class _BadScan(_Redis):
        def scan_iter(self, match='*', count=100):
            raise ConnectionError('scan')
    itn = _dm_env(monkeypatch, _Client(), redis=_BadScan())
    r = vas._send_dms_for_next_visit([{'lead': _lead('L-1'), 'assign': ['YM'], 'memo': ''}], {}, itn, [], [],
                                     addr_candidates=[], target_date=date.today() + timedelta(days=3))
    assert r['dm_history_unread'] is True


def test_redis_ok_no_flag(monkeypatch):
    itn = _dm_env(monkeypatch, _Client(), redis=_Redis())
    r = vas._send_dms_for_next_visit([{'lead': _lead('L-1'), 'assign': ['YM'], 'memo': ''}], {}, itn, [], [],
                                     addr_candidates=[], target_date=date.today() + timedelta(days=3))
    assert r['dm_history_unread'] is False and r['online_duty_failed'] == 0 and r['deassign_failed'] == 0


def test_duty_and_deassign_failures_counted(monkeypatch):
    """당번 DM 이메일 매핑 실패·응답 실패, 재배정 안내 응답 실패를 센다 + ERROR 한 번."""
    fr = _Redis()
    fc = _FailClient(fail={'U_sj', 'U_lsh'})
    itn = _dm_env(monkeypatch, fc, redis=fr,
                  email=lambda ini, itn: '' if ini == 'JK' else f'{ini.lower()}@example.com')
    errs = _errors(monkeypatch, vas)
    target = date.today() + timedelta(days=3)
    fr.set(f'dm_v9_msg:LSH:{target.isoformat()}', '{"channel": "D2", "ts": "2", "lead_nos": ["L-9"], "memo": ""}')
    r = vas._send_dms_for_next_visit([{'lead': _lead('L-1'), 'assign': ['YM'], 'memo': ''}], {}, itn, ['SJ', 'JK'], [],
                                     addr_candidates=[], target_date=target)
    assert r['online_duty_failed'] == 2 and r['online_duty_sent'] == 0
    assert r['deassign_failed'] == 1 and r['deassign_sent'] == 0 and r['visit_mgr_sent'] == 1
    dm_err = [m for m, _ in errs if 'DM 못 보냄' in m]
    assert len(dm_err) == 1 and '당번 2명' in dm_err[0] and '재배정 안내 1명' in dm_err[0]
    s = vp._summary({'dm': r})
    assert s['duty_failed'] == 2 and s['deassign_failed'] == 1 and s['bot_missing'] is False


def test_deassign_email_missing_counted(monkeypatch):
    fr = _Redis()
    itn = _dm_env(monkeypatch, _Client(), redis=fr,
                  email=lambda ini, itn: '' if ini == 'LSH' else f'{ini.lower()}@example.com')
    target = date.today() + timedelta(days=3)
    fr.set(f'dm_v9_msg:LSH:{target.isoformat()}', '{"channel": "D2", "ts": "2", "lead_nos": ["L-9"], "memo": ""}')
    r = vas._send_dms_for_next_visit([{'lead': _lead('L-1'), 'assign': ['YM'], 'memo': ''}], {}, itn, [], [],
                                     addr_candidates=[], target_date=target)
    assert r['deassign_failed'] == 1


def _deassign_text(monkeypatch, prev_leads, now):
    fr, fc = _Redis(), _Client()
    itn = _dm_env(monkeypatch, fc, redis=fr)
    target = date.today() + timedelta(days=3)
    fr.set(f'dm_v9_msg:LSH:{target.isoformat()}',
           json.dumps({'channel': 'D2', 'ts': '2', 'lead_nos': prev_leads, 'memo': ''}))
    vas._send_dms_for_next_visit(now, {}, itn, [], [], addr_candidates=[], target_date=target)
    return next(p['text'] for p in fc.posted if p['channel'] == 'U_lsh')


def test_deassign_text_moved_vs_dropped(monkeypatch):
    """빠진 매니저 안내 — 다른 매니저에게 간 현장만 '재배정', 아무에게도 안 준 현장은 '일정에서 빠졌습니다'."""
    now = [{'lead': _lead('L-1'), 'assign': ['YM'], 'memo': ''}]
    t = _deassign_text(monkeypatch, ['L-1'], now)                     # L-1 은 YM 에게 감
    assert '1건이 다른 매니저에게 재배정되었습니다' in t and '일정에서 빠' not in t
    t = _deassign_text(monkeypatch, ['L-9'], now)                     # L-9 는 아무 코스에도 없음
    assert '1건이 일정에서 빠졌습니다' in t and '재배정' not in t
    t = _deassign_text(monkeypatch, ['L-1', 'L-9'], now)
    assert '2건이 바뀌었습니다(다른 매니저에게 재배정 1건 · 일정에서 빠짐 1건)' in t
    assert 'L-1 · -* — 다른 매니저에게 재배정' in t and 'L-9 · -* — 일정에서 빠짐' in t


# ── 10. 다시 확정 때 코스에서 뺀 현장 — 담당자 그대로(시트는 안 건드림) ───────────
def test_left_assigned_listed(db, monkeypatch):
    prev = {'courses': [{'members': ['YM'], 'leads': ['L-1']}, {'members': ['TH'], 'leads': ['L-3']},
                        {'members': ['JK'], 'leads': ['L-4']}], 'online': [], 'online_shared': False}
    now = {'courses': [{'members': ['YM'], 'leads': ['L-1']}], 'online': [], 'online_shared': False}
    # L-3 = 담당자 '서준' 그대로 → 표시, L-4 = 담당자 없음 → 뺌
    assert vp._left_assigned(DAY, prev, now, LEADS) == [{'lead_no': 'L-3', 'who': '서준'}]
    assert vp._left_assigned(DAY, None, now, LEADS) == []
    # 다른 날로 옮긴 현장은 뺀다(그날 방문 목록에 없음)
    moved = [dict(l, **({'방문 예정일': TODAY} if l['리드 No'] == 'L-3' else {})) for l in LEADS]
    assert vp._left_assigned(DAY, prev, now, moved) == []


def test_confirm_left_assigned_in_result(db, monkeypatch):
    seen = []
    _confirm_env(monkeypatch, seen)
    vp.start_confirm(DAY, BASE, 'p@example.com', LEADS, lambda: LEADS)
    assert vp.run_state(DAY)['result']['left_assigned'] == []
    only_ym = {'courses': [{'members': ['YM'], 'leads': ['L-1']}], 'online': [], 'online_shared': False}
    vp.start_confirm(DAY, only_ym, 'p@example.com', LEADS, lambda: LEADS)
    assert vp.run_state(DAY)['result']['left_assigned'] == [{'lead_no': 'L-3', 'who': '서준'}]
