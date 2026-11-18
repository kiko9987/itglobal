"""그룹웨어에서 본인 휴가 신청·취소 (2026-11-15, KiKO: "PM 사이트에서도 휴가를 쓰게").

슬랙 휴가 앱과 같은 길로 간다:
- 신청 = 슬랙 신청 창과 같은 칸(종류·경조사 종류·날짜·마지막 날·사유), 같은 검사(leave.validate),
  같은 장부 함수(leave.create), 같은 알림(slack_leave.start → 신청자 DM 카드·JW→YG 승인 카드·확정 시 #휴가_일정).
- 취소 = 슬랙 [신청 취소]·[남은 날 취소] 버튼과 같은 함수(slack_leave.apply_action).
- 열기 규칙도 같다(leave.is_open — LEAVE_ENABLED 전엔 LEAVE_PILOT_IDS 만).

사람 구분은 슬랙 ID 하나로 한다(휴가 장부가 슬랙 ID 로 사람을 찾음).
- 구글 로그인: 로그인 메일로 슬랙에 물어본 ID 가 인사카드 ID 와 같을 때만(gw_approval.slack_id_status, 11-13 보안 점검).
- 현장 세션: 1회용 링크를 누른 슬랙 계정(session.user.gw_slack_id)이 지금 인사카드 ID 와 같을 때만.
직원 화면은 본인 휴가만 신청·취소. 경영지원은 휴가 관리 기능(등록·고치기·취소·연차 조정, 11-18).
"""
from __future__ import annotations

import threading
from datetime import date
from typing import Dict, Optional, Tuple

from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)
_LOG = '[GW휴가]'

MSG_NOT_OPEN = '휴가 신청은 아직 준비 중이에요. 지금은 플렉스에서 신청해 주세요.'
MSG_NOT_LISTED = '휴가 명단에 아직 없어요. 경영지원에 문의해 주세요.'
MSG_NO_SLACK = '인사카드에 슬랙 계정이 연결되지 않아 휴가를 신청할 수 없어요. 경영지원에 문의해 주세요.'
MSG_SLACK_DOWN = '슬랙 계정을 확인하지 못했어요. 잠시 뒤 다시 해 주세요.'
MSG_RELOGIN = '로그인 정보가 오래됐어요. 슬랙 휴가 앱에서 [그룹웨어 열기]로 다시 들어와 주세요.'


def _person_and_sid(user: dict) -> Tuple[Optional[dict], str, str]:
    """(인사카드, 인정한 슬랙 ID, 못 쓰는 이유). 이유가 있으면 ID 는 ''."""
    if user.get('gw_field'):
        p = hr.get_person(int(user.get('gw_person_id') or 0))
        if not p or p.get('status') == hr.STATUS_RESIGNED:
            return None, '', MSG_NOT_LISTED
        link_sid = (user.get('gw_slack_id') or '').strip()
        if not link_sid:
            return p, '', MSG_RELOGIN            # 이 기능 전에 만든 현장 세션 — 누른 슬랙 계정을 모름
        if link_sid != (p.get('slack_id') or '').strip():
            logger.warning(f'{_LOG} 현장 세션 슬랙 ID 와 인사카드가 다름 — 거절 (#{p["id"]})')
            return p, '', MSG_RELOGIN
        # 휴대폰 등록(관리자가 숫자 발급)은 인사카드 슬랙 ID 를 그대로 믿는다 → 그 ID 의 휴가 장부 이름이
        # 인사카드 이름과 같을 때만(카드에 남의 슬랙 ID 를 넣어 그 사람 휴가를 건드리지 못하게, 11-15)
        staff = lv.staff_by_slack(link_sid)
        if staff and staff['name'] != p['name']:
            logger.warning(f'{_LOG} 현장 세션 슬랙 ID 의 휴가 장부 이름이 인사카드와 다름 — 거절 (#{p["id"]})')
            return p, '', MSG_NOT_LISTED
        return p, link_sid, ''
    email = (user.get('email') or '').strip().lower()
    p = hr.find_by_email(email) if email else None
    if not p:
        return None, '', MSG_NOT_LISTED
    from dashboard.services.gw_approval import slack_id_status
    sid, state = slack_id_status(email, p.get('slack_id') or '')
    if state == 'unavailable':
        return p, '', MSG_SLACK_DOWN
    if not sid:
        return p, '', MSG_NO_SLACK
    return p, sid, ''


def who(user: dict) -> Dict:
    """신청할 수 있는지와 그 근거 — {uid, name(휴가 장부 이름), why(못 하면 문구)}."""
    p, sid, why = _person_and_sid(user or {})
    if why:
        return {'uid': '', 'name': '', 'why': why, 'person': p}
    if not lv.is_open(sid):
        return {'uid': sid, 'name': '', 'why': MSG_NOT_OPEN, 'person': p}
    staff = lv.staff_by_slack(sid)
    if not staff:
        return {'uid': sid, 'name': '', 'why': MSG_NOT_LISTED, 'person': p}
    return {'uid': sid, 'name': staff['name'], 'why': '', 'person': p}


def form_info(user: dict, today: Optional[date] = None) -> Dict:
    """신청 창에 필요한 것 — 슬랙 신청 창과 같은 종류·경조사 세부·안내 문구."""
    w = who(user)
    out = {'can_apply': not w['why'], 'why': w['why']}
    if w['why']:
        return out
    out.update({
        'name': w['name'],
        'balance': lv.balance(w['name'], today or date.today()),
        'types': [{'value': k, 'label': t} for k, t in lv.type_options()],
        'half_types': list(lv.HALF_TYPES),
        'event': lv.EVENT,
        'event_subs': [{'value': k, 'label': f'{v[0]} ({v[1]}일)'} for k, v in lv.EVENT_SUBS.items()],
        'needs_proof': [k for k, _ in lv.type_options() if lv.needs_proof(k)],
    })
    return out


def _run_bg(fn) -> None:
    """슬랙 알림은 뒤에서(화면이 슬랙 응답을 기다리지 않게). 시험에서는 바로 돌도록 바꿔 끼운다."""
    threading.Thread(target=fn, daemon=True).start()


def apply(user: dict, data: dict) -> Tuple[Optional[dict], Dict[str, str]]:
    """신청 — (만든 줄, {}) 또는 (None, {칸: 오류 문구}). 칸 이름은 슬랙 창과 같다(type·sub·start·end·reason)."""
    w = who(user)
    if w['why']:
        return None, {'_': w['why']}
    name, uid = w['name'], w['uid']
    vals = {k: str((data or {}).get(k) or '').strip() for k in ('type', 'sub', 'start', 'end', 'reason')}
    if not vals['reason']:
        return None, {'reason': '사유를 적어 주세요.'}
    if len(vals['reason']) > 200:
        return None, {'reason': '사유는 200자까지 적을 수 있어요.'}
    ltype, sub = vals['type'], vals['sub']
    start = lv._d(vals['start'])
    end = start if ltype in lv.HALF_TYPES else (lv._d(vals['end']) or start)
    errs = lv.validate(name, ltype, sub, start, end)
    if errs:
        return None, errs
    row = lv.create(name, uid, uid, name, ltype, sub if ltype == lv.EVENT else '', start, end, vals['reason'])
    if not row:
        return None, {'start': '그 사이 남은 연차가 바뀌었어요. 다시 확인해 주세요.'}
    logger.info(f'{_LOG} #{row["id"]} 신청(웹) {name} {ltype} {row["start"]}~{row["end"]} ({row["days"]}일)')

    def _bg():
        try:
            from dashboard.blueprints import slack_leave
            slack_leave.start(lv.get(row['id']))
        except Exception as exc:
            logger.error(f'{_LOG} 신청 알림 실패 (#{row["id"]}): {exc}', exc_info=True, extra={'force_alert': True})
    _run_bg(_bg)
    return row, {}


def cancel(user: dict, rid: int) -> Optional[str]:
    """본인 휴가 취소 — 성공이면 None, 아니면 보여 줄 문구. 슬랙 [신청 취소]·[남은 날 취소]와 같은 함수."""
    p, sid, why = _person_and_sid(user or {})
    if why:
        return why
    row = lv.get(int(rid))
    if not row or sid not in (row.get('slack_id'), row.get('submitter_id')):
        return '본인 휴가만 취소할 수 있어요.'
    from dashboard.blueprints import slack_leave
    return slack_leave.apply_action('cancel', row['id'], sid)


def _vals(data: dict) -> Tuple[Dict[str, str], Optional[date], Optional[date]]:
    vals = {k: str((data or {}).get(k) or '').strip() for k in ('type', 'sub', 'start', 'end', 'reason')}
    start = lv._d(vals['start'])
    end = start if vals['type'] in lv.HALF_TYPES else (lv._d(vals['end']) or start)
    return vals, start, end


def edit(user: dict, rid: int, data: dict) -> Tuple[Optional[dict], Dict[str, str], str]:
    """본인 고치기(승인 대기)·변경 신청(승인된 휴가) — (줄, 오류, 'edit'|'change'). 슬랙 [✏️] 창과 같은 함수."""
    p, sid, why = _person_and_sid(user or {})
    if why:
        return None, {'_': why}, ''
    row = lv.get(int(rid))
    if not row or sid not in (row.get('slack_id'), row.get('submitter_id')):
        return None, {'_': '본인 휴가만 고칠 수 있어요.'}, ''
    mode = lv.own_edit_mode(row, sid)
    if not mode:
        msg = '이미 변경 신청이 승인을 기다리고 있어요.' if lv.pending_change(row['id']) else '이 휴가는 지금 고칠 수 없어요.'
        return None, {'_': msg}, ''
    vals, start, end = _vals(data)
    if len(vals['reason']) > 200:
        return None, {'reason': '사유는 200자까지 적을 수 있어요.'}, mode
    fn = lv.change_request if mode == 'change' else lv.edit_request
    name = row['name']
    new, errs = fn(row['id'], vals['type'], vals['sub'], start, end, vals['reason'], sid, name)
    if errs:
        return None, errs, mode

    def _bg():
        try:
            from dashboard.blueprints import slack_leave
            if mode == 'change':
                slack_leave.after_change(new)
            else:
                slack_leave.after_edit(new, sid, False)
        except Exception as exc:
            logger.error(f'{_LOG} 고치기 알림 실패 (#{new["id"]}): {exc}', exc_info=True, extra={'force_alert': True})
    _run_bg(_bg)
    logger.info(f'{_LOG} #{new["id"]} {"변경 신청" if mode == "change" else "고침"}(웹) {name} {new["start"]}~{new["end"]}')
    return new, {}, mode


def manager_sid(user: dict) -> str:
    """경영지원(휴가 관리자)이면 그 사람 슬랙 ID, 아니면 ''."""
    try:
        _p, sid, why = _person_and_sid(user or {})
    except Exception:
        return ''
    return sid if (not why and sid and sid in lv.manager_ids()) else ''


def _actor(user: dict, sid: str) -> str:
    p, _s, _w = _person_and_sid(user or {})
    return (p or {}).get('name') or (user or {}).get('name') or sid


def admin_edit(user: dict, rid: int, data: dict) -> Tuple[Optional[dict], Dict[str, str]]:
    sid = manager_sid(user)
    if not sid:
        return None, {'_': '경영지원만 고칠 수 있어요.'}
    vals, start, end = _vals(data)
    actor = _actor(user, sid)
    row, errs = lv.edit_request(int(rid), vals['type'], vals['sub'], start, end, vals['reason'], sid, actor, as_admin=True)
    if errs:
        return None, errs

    def _bg():
        try:
            from dashboard.blueprints import slack_leave
            slack_leave.after_edit(row, sid, True)
        except Exception as exc:
            logger.error(f'{_LOG} 경영지원 고침 알림 실패 (#{row["id"]}): {exc}', exc_info=True, extra={'force_alert': True})
    _run_bg(_bg)
    logger.info(f'{_LOG} #{row["id"]} 경영지원 고침({actor}) {row["name"]} {row["start"]}~{row["end"]}')
    return row, {}


def admin_cancel(user: dict, rid: int) -> Optional[str]:
    """경영지원 취소(끝난 휴가 포함) — 슬랙 카드와 같은 함수."""
    sid = manager_sid(user)
    if not sid:
        return '경영지원만 취소할 수 있어요.'
    from dashboard.blueprints import slack_leave
    return slack_leave.apply_action('cancel', int(rid), sid)


def adjust(user: dict, name: str, data: dict) -> Tuple[Optional[dict], str]:
    sid = manager_sid(user)
    if not sid:
        return None, '경영지원만 조정할 수 있어요.'
    actor = _actor(user, sid)
    g, err = lv.adjust(name, (data or {}).get('days'), str((data or {}).get('reason') or ''), sid, actor)
    if err:
        return None, err

    def _bg():
        try:
            from dashboard.blueprints import slack_leave
            slack_leave.after_adjust(name, g, sid)
        except Exception as exc:
            logger.error(f'{_LOG} 연차 조정 알림 실패 ({name}): {exc}', exc_info=True)
    _run_bg(_bg)
    logger.info(f'{_LOG} 연차 조정 {name} {g["days"]:+g}일 ({actor})')
    return g, ''


def admin_register(user: dict, name: str, data: dict) -> Tuple[Optional[dict], Dict[str, str]]:
    sid = manager_sid(user)
    if not sid:
        return None, {'_': '경영지원만 등록할 수 있어요.'}
    vals, start, end = _vals(data)
    if len(vals['reason']) > 200:
        return None, {'reason': '사유는 200자까지 적을 수 있어요.'}
    actor = _actor(user, sid)
    row, errs = lv.admin_register(name, vals['type'], vals['sub'], start, end, vals['reason'], sid, actor)
    if errs:
        return None, errs

    def _bg():
        try:
            from dashboard.blueprints import slack_leave
            slack_leave.after_admin_register(row, sid)
        except Exception as exc:
            logger.error(f'{_LOG} 경영지원 등록 알림 실패 (#{row["id"]}): {exc}', exc_info=True, extra={'force_alert': True})
    _run_bg(_bg)
    logger.info(f'{_LOG} #{row["id"]} 경영지원 등록({actor}) {name} {row["ltype"]} {row["start"]}~{row["end"]}')
    return row, {}


def admin_info(user: dict, person: dict) -> Dict:
    """직원별 휴가 창의 경영지원 칸 — 쓸 수 있는지·종류·조정 기록·신청 줄(고치기·취소용)."""
    sid = manager_sid(user)
    if not sid:
        return {'can': False}
    from dashboard.services import leave_view
    name = leave_view.leave_name(person)
    if not name:
        return {'can': False, 'why': '휴가 장부에 연결되지 않은 직원이에요.'}
    rows = {r['id']: r for r in lv.requests_of(name, limit=500)}
    editable = [rid for rid, r in rows.items() if r['status'] in lv.HOLDING]
    return {
        'can': True, 'name': name,
        'types': [{'value': k, 'label': v} for k, v in lv.type_options() + lv.admin_type_options()],
        'admin_types': [{'value': k, 'label': v} for k, v in lv.admin_type_options()],
        'balance': lv.balance(name),
        'event': lv.EVENT, 'half_types': list(lv.HALF_TYPES),
        'event_subs': [{'value': k, 'label': f'{v[0]} ({v[1]}일)'} for k, v in lv.EVENT_SUBS.items()],
        'editable': editable,
        'rows': {str(rid): {'type': r['ltype'], 'sub': r.get('sub') or '', 'start': r['start'], 'end': r['end'],
                            'reason': r.get('reason') or ''} for rid, r in rows.items() if rid in editable},
        'adjustments': [{'days': a['days'], 'note': a.get('note') or '', 'at': (a.get('created_at') or '')[:16],
                         'valid_to': a['valid_to']} for a in lv.adjustments(name)],
    }


def annotate(summary: dict, user: dict, today: Optional[date] = None) -> dict:
    """내 휴가 요약의 슬랙(장부) 기록마다 본인 취소 단추 정보 — cancel: 'all'|'rest'|'' 와 확인 문구."""
    today = today or date.today()
    try:
        _p, sid, why = _person_and_sid(user or {})
    except Exception as exc:
        logger.warning(f'{_LOG} 취소 단추 판단 실패: {exc}')
        sid, why = '', 'error'
    rows = {}
    staff = lv.staff_by_slack(sid) if sid and not why else None
    if staff:
        rows = {r['id']: r for r in lv.requests_of(staff['name'], limit=500)}
    for e in summary.get('entries') or []:
        mode, emode, r = '', '', None
        if e.get('source') == 'slack' and e.get('id') in rows:
            r = rows[e['id']]
            mode = lv.own_cancel_mode(r, sid, today)
            emode = lv.own_edit_mode(r, sid, today)
        e['cancel'] = mode
        e['cancel_label'] = lv.cancel_label(mode) if mode else ''
        e['cancel_confirm'] = lv.cancel_confirm_text(mode) if mode else ''
        # 고치기(11-18): 'edit' 승인 대기 / 'change' 승인된 시작 전 휴가 — 창을 채울 지금 값도 같이
        e['edit'] = emode
        e['edit_label'] = lv.edit_label(emode) if emode else ''
        if emode and r:
            e['edit_vals'] = {'type': r['ltype'], 'sub': r.get('sub') or '', 'start': r['start'], 'end': r['end'],
                              'reason': r.get('reason') or ''}
        if r and r.get('replaces'):
            e['change_of'] = r['replaces']
    return summary
