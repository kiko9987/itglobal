"""조용한 실패 점검 2차(11-24) — 전자결재·휴가 상태 저장 뒤 단계와 꼭 가야 하는 알림.

- 전자결재 apply_action·process_edit: 상태를 저장한 뒤 단계가 실패해도 ERROR(+force_alert)만 남기고
  나머지 단계(카드 갱신·감사 로그)는 계속, 화면엔 성공(None / row).
- 반려·보완 요청·남이 한 취소 알림 DM 실패 → ERROR. 보통 알림(완료 등)은 예전처럼 WARNING.
- 휴가 반려·남이 한 취소 DM, #휴가_일정 취소·기간 줄임 표시 실패 → ERROR.
"""
import logging
from datetime import date

import pytest  # noqa: F401

from dashboard.services import e_approval as ea
from dashboard.services import leave as lv
from test_e_approval import FIELDS, JW, KIKO, MGR, SB, YG, _no_final_by_default, se, tmp_db  # noqa: F401
from test_leave import MW, _seed, db, slack  # noqa: F401


def _alerts(caplog, *needles):
    """ERROR 이상 + force_alert 이고 문구가 다 들어 있는 기록."""
    return [r for r in caplog.records
            if r.levelno >= logging.ERROR and getattr(r, 'force_alert', False)
            and all(n in r.getMessage() for n in needles)]


def _spy_audit(mod, monkeypatch):
    got = []
    monkeypatch.setattr(mod, '_audit', lambda row, action, name, reason='': got.append((action, row['id'])))
    return got


def _boom(*a, **kw):
    raise RuntimeError('boom')


# ── 전자결재 ─────────────────────────────────────────────────────────
def _eap_dm_fails_for(fake, uid, monkeypatch):
    """uid 에게 가는 DM 만 실패. 돌려주는 dict 의 'on' 을 False 로 하면 다시 된다."""
    state = {'on': True}

    def conv(**kw):
        fake.calls.append(('conversations_open', kw))
        if state['on'] and kw['users'] == uid:
            raise RuntimeError('cannot_dm_bot')
        return {'channel': {'id': 'D-' + kw['users']}}
    monkeypatch.setattr(fake, 'conversations_open', conv, raising=False)
    return state


def _dm_texts(fake, uid):
    return [kw.get('text', '') for kw in fake.of('chat_postMessage') if kw.get('channel') == 'D-' + uid]


def test_eap_reject_dm_failure_alerts_and_rest_continues(se, monkeypatch, caplog):
    caplog.set_level(logging.WARNING)
    audits = _spy_audit(se, monkeypatch)
    row = se.process_new_submit(MGR, dict(FIELDS))
    _eap_dm_fails_for(se._fake, MGR, monkeypatch)
    n_upd = len(se._fake.of('chat_update'))
    assert se.apply_action('reject', row['id'], KIKO, '견적 다시') is None
    assert ea.get(row['id'])['status'] == ea.REJECTED
    errs = _alerts(caplog, '알림 DM 실패', MGR, '직접 알려')
    assert errs and '반려됐어요' in errs[0].getMessage()
    assert len(se._fake.of('chat_update')) > n_upd                    # 카드 갱신 계속
    assert ('reject', row['id']) in audits                             # 감사 로그 계속


def test_eap_plain_notify_failure_stays_warning(se, monkeypatch, caplog):
    """이체 완료 같은 보통 알림은 예전처럼 WARNING — 카드·스레드로도 보인다."""
    caplog.set_level(logging.WARNING)
    row = se.process_new_submit(MGR, dict(FIELDS))
    assert se.apply_action('approve', row['id'], KIKO) is None
    _eap_dm_fails_for(se._fake, MGR, monkeypatch)
    assert se.apply_action('transfer', row['id'], SB) is None
    assert not _alerts(caplog, '알림 DM 실패')
    assert any(r.levelno == logging.WARNING and '알림 DM 실패' in r.getMessage() for r in caplog.records)


def test_eap_cancel_by_other_alerts_but_self_cancel_sends_nothing(se, monkeypatch, caplog):
    caplog.set_level(logging.WARNING)
    a = se.process_new_submit(MGR, dict(FIELDS))
    b = se.process_new_submit(MGR, dict(FIELDS))
    _eap_dm_fails_for(se._fake, MGR, monkeypatch)
    assert se.apply_action('cancel', a['id'], KIKO) is None            # 승인자가 취소 → 요청자에게 꼭
    assert _alerts(caplog, '알림 DM 실패', '요청이 취소됐어요')
    caplog.clear()
    n = len([c for c in se._fake.calls if c[0] == 'conversations_open'])
    assert se.apply_action('cancel', b['id'], MGR) is None             # 본인 취소 → 알림 없음
    assert len([c for c in se._fake.calls if c[0] == 'conversations_open']) == n
    assert not _alerts(caplog, '알림 DM 실패')


def test_eap_revise_requester_dm_failures_alert(se, monkeypatch, caplog):
    caplog.set_level(logging.WARNING)
    # 첫 보완 요청 — 요청자 DM([✏️ 수정하기]) 이 안 가면 ERROR
    a = se.process_new_submit(MGR, dict(FIELDS))
    fail = _eap_dm_fails_for(se._fake, MGR, monkeypatch)
    assert se.apply_action('revise', a['id'], KIKO, '금액 확인') is None
    assert _alerts(caplog, 'DM 실패', MGR, 'requester')
    fail['on'] = False
    caplog.clear()
    # 두 번째 보완 요청 — 예전 DM 이 있어 짧은 알림. 이것도 안 가면 ERROR
    b = se.process_new_submit(MGR, dict(FIELDS))
    assert se.apply_action('revise', b['id'], KIKO, '금액 확인') is None
    assert not _alerts(caplog, 'DM 실패')
    ea.resubmit(b['id'], MGR, '강성환', {'amount': 123000})
    fail['on'] = True
    assert se.apply_action('revise', b['id'], KIKO, '계좌 확인') is None
    assert _alerts(caplog, '알림 DM 실패', '보완 요청이 왔어요')


def test_eap_approve_next_dm_step_failure_keeps_refresh_and_audit(se, monkeypatch, caplog):
    caplog.set_level(logging.WARNING)
    row = se.process_new_submit(MGR, dict(FIELDS))
    audits = _spy_audit(se, monkeypatch)
    monkeypatch.setattr(se, '_dm_round', _boom)
    n_upd = len(se._fake.of('chat_update'))
    assert se.apply_action('approve', row['id'], KIKO) is None
    assert ea.get(row['id'])['status'] == ea.APPROVED
    assert _alerts(caplog, f"#{row['id']}", '다음 처리자 DM')
    assert len(se._fake.of('chat_update')) > n_upd
    assert ('approve', row['id']) in audits


def test_eap_transfer_remember_payee_failure_is_independent(se, monkeypatch, caplog):
    caplog.set_level(logging.WARNING)
    row = se.process_new_submit(MGR, dict(FIELDS))
    assert se.apply_action('approve', row['id'], KIKO) is None
    audits = _spy_audit(se, monkeypatch)
    monkeypatch.setattr(ea, 'remember_payee', _boom)
    assert se.apply_action('transfer', row['id'], SB) is None
    assert ea.get(row['id'])['status'] == ea.DONE
    assert _alerts(caplog, f"#{row['id']}", '거래처 계좌 기억')
    assert any('이체가 끝났어요' in t for t in _dm_texts(se._fake, MGR))   # 요청자 알림은 갔다
    assert ('transfer', row['id']) in audits


def test_eap_branch_refresh_and_audit_failures_each_alert(se, monkeypatch, caplog):
    caplog.set_level(logging.WARNING)
    row = se.process_new_submit(MGR, dict(FIELDS))
    audits = []
    monkeypatch.setattr(se, '_thread_log', _boom)                     # ① 분기 블록
    monkeypatch.setattr(se, 'refresh', _boom)                         # ③ 카드 갱신

    def audit_boom(row_, action, name, reason=''):
        audits.append(action)
        raise RuntimeError('audit db locked')
    monkeypatch.setattr(se, '_audit', audit_boom)                     # ④ 감사 로그
    assert se.apply_action('reject', row['id'], KIKO, '견적 다시') is None
    assert ea.get(row['id'])['status'] == ea.REJECTED
    assert _alerts(caplog, '반려 기록·요청자 알림')
    assert _alerts(caplog, '카드·DM 새로 그리기')
    assert _alerts(caplog, '감사 로그', 'audit db locked')
    assert audits == ['reject']                                        # 앞이 실패해도 감사 로그까지 갔다


def test_eap_web_act_returns_success_when_after_step_fails(se, monkeypatch, caplog):
    """그룹웨어 결재함(gw_approval.act)도 상태 저장 뒤 실패로 500 '다시 눌러 주세요'가 나지 않는다."""
    caplog.set_level(logging.WARNING)
    row = se.process_new_submit(MGR, dict(FIELDS))
    monkeypatch.setattr(se, 'refresh', _boom)
    assert se.apply_action('approve', row['id'], KIKO, expected_rev=0) is None
    assert ea.get(row['id'])['status'] == ea.APPROVED
    assert _alerts(caplog, '카드·DM 새로 그리기')


def test_eap_process_edit_after_steps_independent(se, monkeypatch, caplog):
    caplog.set_level(logging.WARNING)
    row = se.process_new_submit(MGR, dict(FIELDS))
    assert se.apply_action('revise', row['id'], KIKO, '금액 확인') is None
    audits = _spy_audit(se, monkeypatch)
    monkeypatch.setattr(se, 'refresh', _boom)
    monkeypatch.setattr(se, '_thread_log', _boom)
    out = se.process_edit(MGR, row['id'], {'amount': 123000}, [])
    assert out and out['id'] == row['id'] and out['status'] == ea.PENDING and out['amount'] == 123000
    assert _alerts(caplog, '수정 스레드 기록') and _alerts(caplog, '카드·DM 새로 그리기')
    assert any('수정돼서 다시 올라왔어요' in t for t in _dm_texts(se._fake, KIKO))   # 승인자 다시 알림은 갔다
    assert ('resubmit', row['id']) in audits


# ── 휴가 ─────────────────────────────────────────────────────────────
def _leave_dm_fails_for(c, uid, monkeypatch):
    real = c.conversations_open

    def conv(users):
        if users == uid:
            raise RuntimeError('cannot_dm_bot')
        return real(users)
    monkeypatch.setattr(c, 'conversations_open', conv)


def _approved_leave(sl, day_from, day_to):
    r = lv.create('박민우', MW, MW, '박민우', lv.ANNUAL, '', day_from, day_to)
    sl.start(r)
    assert sl.apply_action('approve', r['id'], JW) is None
    assert sl.apply_action('approve', r['id'], YG) is None
    r = lv.get(r['id'])
    assert r['status'] == lv.APPROVED and r.get('announce_ts')
    return r


def test_leave_reject_dm_failure_alerts(slack, monkeypatch, caplog):
    sl, c = slack
    caplog.set_level(logging.WARNING)
    _seed(left=3)
    r = lv.create('박민우', MW, MW, '박민우', lv.ANNUAL, '', date(2026, 12, 7), date(2026, 12, 7))
    sl.start(r)
    _leave_dm_fails_for(c, MW, monkeypatch)
    assert sl.apply_action('reject', r['id'], JW, '현장 일정') is None
    assert lv.get(r['id'])['status'] == lv.REJECTED
    errs = _alerts(caplog, '알림 DM 실패', MW, '직접 알려')
    assert errs and '반려됐어요' in errs[0].getMessage()


def test_leave_admin_cancel_alerts_dm_and_announce_but_tells_approvers(slack, monkeypatch, caplog):
    sl, c = slack
    caplog.set_level(logging.WARNING)
    _seed(left=3)
    r = _approved_leave(sl, date(2026, 12, 7), date(2026, 12, 7))
    _leave_dm_fails_for(c, MW, monkeypatch)
    real_upd = c.chat_update

    def upd(channel, ts, text, blocks=None):
        if channel == lv.ANNOUNCE_CHANNEL_DEFAULT:
            raise RuntimeError('message_not_found')
        return real_upd(channel, ts, text, blocks)
    monkeypatch.setattr(c, 'chat_update', upd)
    n = len(c.posts)
    assert sl.apply_action('cancel', r['id'], KIKO) is None            # 경영지원이 남의 휴가 취소
    assert lv.get(r['id'])['status'] == lv.CANCELLED
    assert _alerts(caplog, '#휴가_일정 취소 표시 실패', '채널 글을 직접 고쳐 주세요')
    assert _alerts(caplog, '알림 DM 실패', MW, '취소됐어요')
    told = {p[0] for p in c.posts[n:] if '취소됐어요' in p[1]}
    assert {f'D-{JW}', f'D-{YG}'} <= told                              # 뒤 단계(승인자 알림)는 계속


def test_leave_own_cancel_approver_dm_failure_stays_warning(slack, monkeypatch, caplog):
    """본인이 취소하면 본인 DM 은 없고, 승인자 알림 실패는 WARNING(승인자 카드가 '취소'로 바뀌어 보인다)."""
    sl, c = slack
    caplog.set_level(logging.WARNING)
    _seed(left=3)
    r = _approved_leave(sl, date(2027, 1, 25), date(2027, 1, 25))
    _leave_dm_fails_for(c, JW, monkeypatch)
    assert sl.apply_action('cancel', r['id'], MW) is None
    assert lv.get(r['id'])['status'] == lv.CANCELLED
    assert not _alerts(caplog, '알림 DM 실패')
    assert any(p[0] == f'D-{YG}' and '취소됐어요' in p[1] for p in c.posts)


def test_leave_cut_rest_announce_failure_alerts_and_tells_approvers(slack, monkeypatch, caplog):
    sl, c = slack
    caplog.set_level(logging.WARNING)
    _seed(left=5)
    r = _approved_leave(sl, date(2026, 11, 9), date(2026, 11, 13))
    wed = date(2026, 11, 11)
    real_cut = lv.cut_rest
    monkeypatch.setattr(lv, 'can_act', lambda *a, **kw: True)
    monkeypatch.setattr(lv, 'cancel_mode', lambda *a, **kw: 'rest')
    monkeypatch.setattr(lv, 'cut_rest', lambda rid, uid, actor: real_cut(rid, uid, actor, today=wed))
    real_upd = c.chat_update

    def upd(channel, ts, text, blocks=None):
        if channel == lv.ANNOUNCE_CHANNEL_DEFAULT:
            raise RuntimeError('message_not_found')
        return real_upd(channel, ts, text, blocks)
    monkeypatch.setattr(c, 'chat_update', upd)
    n = len(c.posts)
    assert sl.apply_action('cancel', r['id'], MW) is None
    assert lv.get(r['id'])['end'] == '2026-11-10'
    assert _alerts(caplog, '#휴가_일정 기간 줄임 표시 실패', '채널 글을 직접 고쳐 주세요')
    told = {p[0] for p in c.posts[n:] if '오늘부터 취소' in p[1]}
    assert {f'D-{JW}', f'D-{YG}'} <= told
