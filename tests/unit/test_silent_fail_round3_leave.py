"""조용한 실패 점검 3차(11-24) — 연차 촉구 DM 결과·입사·퇴사 할 일 [✅ 완료]·할 일 DM.

12. 촉구(근로기준법 61조 촉진의 첫 단계): 기록은 '보냄'으로 먼저 남고 DM 은 뒤에서 간다 → DM 결과를
    notice_dm('sent'·'failed')에 남겨 촉진 탭이 '⚠ 촉구 DM 못 보냄'을, 자동 촉구 뒤 경영지원 요약이 'DM 못 보냄'을 띄운다.
25. 입사·퇴사 할 일 [✅ 완료]: 뒤 스레드 예외를 잡아 ERROR(+force_alert) + 누른 사람에게 안내.
    새 할 일·기한 알림 DM 을 받을 사람 누구에게도 못 보냈으면(봇 토큰 없음 포함) ERROR.
"""
import logging
import os
import types
from datetime import date, timedelta

import pytest

from dashboard.services import leave as lv
from test_gw_leave_admin import mgr  # noqa: F401
from test_gw_leave_apply import ADMIN, FakeSlack, _as, client, slack  # noqa: F401


def _alerts(caplog, *needles):
    """ERROR 이상 + force_alert 이고 문구가 다 들어 있는 기록."""
    return [r for r in caplog.records
            if r.levelno >= logging.ERROR and getattr(r, 'force_alert', False)
            and all(n in r.getMessage() for n in needles)]


# ── 12. 연차 촉구 DM 결과 ──────────────────────────────────────────────
@pytest.fixture
def lp(mgr, monkeypatch):  # noqa: F811
    from dashboard.services import leave_promo
    monkeypatch.setattr(leave_promo, '_ready', False)
    return leave_promo


class _DmFails(FakeSlack):
    """정해 둔 사람에게 가는 DM 만 실패(나머지는 FakeSlack 그대로)."""

    def __init__(self, *bad):
        super().__init__()
        self.bad = set(bad)

    def conversations_open(self, users):
        if users in self.bad:
            raise RuntimeError('cannot_dm_bot')
        return super().conversations_open(users)


def _row(client, pid):  # noqa: F811
    return next(r for r in client.get('/api/leave/promotions').get_json()['rows'] if r['id'] == pid)


def _send(client, name='박민우'):  # noqa: F811
    to = next(g['valid_to'] for g in lv.grants(name))
    j = client.post('/api/leave/promotions/send', json={'name': name, 'kind': 'annual', 'grant_to': to}).get_json()
    assert j['success'], j
    return j['id']


def test_notice_dm_sent_is_recorded(client, mgr, lp):  # noqa: F811
    _as(client, ADMIN)
    pid = _send(client)
    assert lp.get(pid)['notice_dm'] == 'sent'
    assert _row(client, pid)['notice_dm'] == 'sent'


def test_notice_dm_failure_recorded_alerted_and_shown(client, mgr, lp, monkeypatch, caplog):  # noqa: F811
    from dashboard.blueprints import slack_leave
    caplog.set_level(logging.WARNING)
    _as(client, ADMIN)
    monkeypatch.setattr(slack_leave, '_client', lambda: _DmFails('UMW'))
    pid = _send(client)
    assert lp.get(pid)['notice_dm'] == 'failed'
    assert _row(client, pid)['notice_dm'] == 'failed'                    # 촉진 탭이 '⚠ 촉구 DM 못 보냄'
    assert _alerts(caplog, '촉구 DM 을 못 보냄', '박민우', '직접 전해')


def test_notice_without_bot_token_records_failed(client, mgr, lp, monkeypatch, caplog):  # noqa: F811
    from dashboard.blueprints import slack_leave
    caplog.set_level(logging.WARNING)
    _as(client, ADMIN)
    monkeypatch.setattr(slack_leave, '_client', lambda: None)
    pid = _send(client)
    assert lp.get(pid)['notice_dm'] == 'failed'
    assert _alerts(caplog, '토큰이 없어 촉구 DM', '박민우')
    assert lp.get(pid)['designate_dm'] is None                            # 지정 통보 칸은 안 건드림


def test_old_records_have_blank_notice_dm(client, mgr, lp):  # noqa: F811
    """열을 나중에 붙여도 옛 기록(플렉스 등)은 빈칸 — 촉진 탭에 경고가 안 뜬다."""
    _as(client, ADMIN)
    pid = _send(client)
    with lv._conn() as c:
        c.execute('UPDATE leave_promotions SET notice_dm=NULL WHERE id=?', (pid,))
    assert lp.view(lp.get(pid))['notice_dm'] == ''


def test_daily_summary_marks_people_dm_failed(slack, mgr, lp, monkeypatch, caplog):  # noqa: F811
    """자동 촉구 — 촉구 DM 이 요약보다 먼저 가므로, 요약에서 DM 못 받은 사람을 구별한다."""
    from dashboard.blueprints import slack_leave
    caplog.set_level(logging.WARNING)
    monkeypatch.setenv('LEAVE_ENABLED', 'true')
    fake = _DmFails('UMW')                                                # 박민우만 실패, 이직원·경영지원은 됨
    monkeypatch.setattr(slack_leave, '_client', lambda: fake)
    to = lv._d(next(g['valid_to'] for g in lv.grants('박민우')))
    day = lp.schedule(lp.ANNUAL, to)['notice_from'] + timedelta(days=1)
    out = slack_leave.run_promotion(day)
    by = {r['name']: r for r in out['sent']}
    assert by['박민우']['notice_dm'] == 'failed' and by['이직원']['notice_dm'] == 'sent'
    summary = next(p['text'] for p in fake.to('UKIKO') if '오늘 촉구를 보냈어요' in p['text'])
    lines = {ln.split(' — ')[0].lstrip('• '): ln for ln in summary.splitlines() if ln.startswith('• ')}
    assert '촉구 DM 못 보냄' in lines['박민우'] and '직접 전해' in lines['박민우']
    assert '못 보냄' not in lines['이직원']
    assert _alerts(caplog, '촉구 DM 을 못 보냄', '박민우')


def test_promo_table_shows_notice_dm_warning():
    js = open(os.path.join('dashboard', 'static', 'js', 'gw-leave.js'), encoding='utf-8').read()
    assert "r.notice_dm === 'failed'" in js and '⚠ 촉구 DM 못 보냄 — 직접 전해 주세요' in js


# ── 25. 입사·퇴사 할 일 ───────────────────────────────────────────────
from dashboard.services import hr_tasks as ht  # noqa: E402
from test_hr_tasks import _person, db  # noqa: E402,F401


class _Eph:
    """전자결재 봇 client 대신 — DM·카드 고치기·ephemeral 을 적어 둔다. fail=True 면 DM 이 모두 실패."""

    def __init__(self, fail=False):
        self.calls, self.fail = [], fail

    def conversations_open(self, users):
        if self.fail:
            raise RuntimeError('cannot_dm_bot')
        return {'channel': {'id': 'D' + users}}

    def chat_postMessage(self, **kw):
        self.calls.append(('post', kw['channel'], kw.get('text')))
        return {'ts': str(len(self.calls))}

    def chat_update(self, **kw):
        self.calls.append(('update', kw['channel']))

    def chat_postEphemeral(self, **kw):
        self.calls.append(('eph', kw['channel'], kw['user'], kw['text']))


class _SyncThread:
    """뒤 스레드를 바로 돌린다(시험)."""

    def __init__(self, target, daemon=None):
        self.target = target

    def start(self):
        self.target()


@pytest.fixture
def done_btn(db, monkeypatch):  # noqa: F811
    """[✅ 완료] 처리 함수와 할 일 하나. 전자결재 봇 토큰은 없게(누른 앱의 client 를 씀)."""
    from dashboard.blueprints import slack_eapproval as se
    from dashboard.blueprints import slack_hr_tasks as sh
    from dashboard.services import certificates as ct
    monkeypatch.delenv('SLACK_EAPPROVAL_BOT_TOKEN', raising=False)
    monkeypatch.setattr(ct, 'can_handle', lambda uid: uid == 'UHANDLER')
    monkeypatch.setattr(se, '_name', lambda uid: '황샛별')
    monkeypatch.setattr(sh, 'threading', types.SimpleNamespace(Thread=_SyncThread))
    handlers = {}

    class _App:
        def action(self, aid):
            def deco(fn):
                handlers[aid] = fn
                return fn
            return deco
    sh.register_hr_task_handlers(_App())
    _person('A', '새사람', hire_date='2026-12-01')
    ht.sync()
    tid = ht.open_tasks()[0]['id']

    def press(client):
        body = {'user': {'id': 'UHANDLER'}, 'channel': {'id': 'DUHANDLER'}, 'actions': [{'value': str(tid)}]}
        handlers[ht.A_DONE](lambda: None, body, client)
    return types.SimpleNamespace(tid=tid, press=press, sh=sh)


def test_done_button_ok_no_alert(done_btn, caplog):
    caplog.set_level(logging.WARNING)
    c = _Eph()
    done_btn.press(c)
    assert ht.get(done_btn.tid)['status'] == ht.DONE
    assert not [x for x in c.calls if x[0] == 'eph']
    assert not _alerts(caplog)


def test_done_button_complete_error_alerts_and_tells_user(done_btn, monkeypatch, caplog):
    caplog.set_level(logging.WARNING)

    def boom(*a, **kw):
        raise RuntimeError('database is locked')
    monkeypatch.setattr(ht, 'complete', boom)
    c = _Eph()
    done_btn.press(c)                                                     # 예외가 밖으로 새지 않음
    assert _alerts(caplog, '완료 처리 실패', f'#{done_btn.tid}', '저장 안 됨')
    assert _alerts(caplog)[0].exc_info
    eph = [x for x in c.calls if x[0] == 'eph']
    assert eph == [('eph', 'DUHANDLER', 'UHANDLER', '완료 처리를 못 했어요. 잠시 뒤 다시 눌러 주세요.')]
    assert ht.get(done_btn.tid)['status'] == ht.OPEN


def test_done_button_refresh_error_after_save(done_btn, monkeypatch, caplog):
    """완료는 저장됐는데 카드 고치기에서 터짐 — '못 했어요'가 아니라 저장은 됐다고, 다시 누르면 카드만 고친다."""
    caplog.set_level(logging.WARNING)
    real_refs = ht.refs

    def boom(row):
        raise ValueError('bad dm_refs')
    monkeypatch.setattr(ht, 'refs', boom)
    c = _Eph()
    done_btn.press(c)
    assert ht.get(done_btn.tid)['status'] == ht.DONE
    assert _alerts(caplog, '완료 처리 실패', '저장 됨')
    eph = [x[3] for x in c.calls if x[0] == 'eph']
    assert eph == ['완료는 저장했는데 카드를 못 고쳤어요. 잠시 뒤 다시 눌러 주세요.']
    monkeypatch.setattr(ht, 'refs', real_refs)
    caplog.clear()
    done_btn.press(_Eph())                                                # 다시 눌러도 조용히(두 번 눌러도 한 번)
    assert not _alerts(caplog)


def test_notify_new_all_failed_is_error_partial_is_not(db, monkeypatch, caplog):  # noqa: F811
    from dashboard.blueprints import slack_hr_tasks as sh
    caplog.set_level(logging.WARNING)
    _person('A', '새사람', hire_date='2026-12-01')
    rows = ht.sync()
    assert sh.notify_new(_Eph(fail=True), rows) == 0
    assert len(_alerts(caplog, '새 할 일 DM 을 아무에게도 못 보냄', '새사람')) == len(rows)
    # 받을 사람 중 한 명이라도 받으면 WARNING 만
    caplog.clear()
    monkeypatch.setattr(ht, 'handler_ids', lambda: ('UGOOD', 'UBAD'))

    class _Half(_Eph):
        def conversations_open(self, users):
            if users == 'UBAD':
                raise RuntimeError('cannot_dm_bot')
            return super().conversations_open(users)
    assert sh.notify_new(_Half(), rows) == len(rows)
    assert not _alerts(caplog)
    assert any(r.levelno == logging.WARNING and 'UBAD' in r.getMessage() for r in caplog.records)


def test_notify_new_no_receivers_is_quiet(db, monkeypatch, caplog):  # noqa: F811
    from dashboard.blueprints import slack_hr_tasks as sh
    caplog.set_level(logging.WARNING)
    monkeypatch.setattr(ht, 'handler_ids', lambda: ())
    _person('A', '새사람', hire_date='2026-12-01')
    assert sh.notify_new(_Eph(fail=True), ht.sync()) == 0
    assert not _alerts(caplog)


def test_sync_without_bot_token_is_error(db, monkeypatch, caplog):  # noqa: F811
    from dashboard.blueprints import slack_hr_tasks as sh
    caplog.set_level(logging.WARNING)
    monkeypatch.delenv('SLACK_EAPPROVAL_BOT_TOKEN', raising=False)
    _person('A', '새사람', hire_date='2026-12-01')
    assert sh.sync_and_notify() == 2
    assert _alerts(caplog, '토큰이 없어 새 할 일 DM', '새사람')


def test_run_daily_reminder_all_failed_is_error(db, monkeypatch, caplog):  # noqa: F811
    from dashboard.blueprints import slack_eapproval as se
    from dashboard.blueprints import slack_hr_tasks as sh
    caplog.set_level(logging.WARNING)
    _person('A', '새사람', hire_date='2026-12-01')
    ht.sync()                                                             # 새 할 일 DM 은 이미 갔다고
    monkeypatch.setattr(se, '_notify', lambda client, uid, text: False)
    sh.run_daily(date(2026, 12, 7), client=_Eph())                       # 월요일, 서류 기한 12/8
    assert _alerts(caplog, '기한 알림 1건을 아무에게도 못 보냄')
    caplog.clear()
    monkeypatch.setattr(se, '_notify', lambda client, uid, text: True)
    sh.run_daily(date(2026, 12, 7), client=_Eph())
    assert not _alerts(caplog)


def test_run_daily_without_client_on_workday_is_error(db, monkeypatch, caplog):  # noqa: F811
    from dashboard.blueprints import slack_hr_tasks as sh
    caplog.set_level(logging.WARNING)
    monkeypatch.delenv('SLACK_EAPPROVAL_BOT_TOKEN', raising=False)
    _person('A', '새사람', hire_date='2026-12-01')
    ht.sync()
    sh.run_daily(date(2026, 12, 7))
    assert _alerts(caplog, '토큰이 없어 기한 알림 1건')
    caplog.clear()
    sh.run_daily(date(2026, 12, 12))                                      # 토요일 — 알릴 날이 아님
    assert not _alerts(caplog)
