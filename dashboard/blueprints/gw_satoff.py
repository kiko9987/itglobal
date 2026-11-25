"""그룹웨어 「토요일 휴무」(2026-11-20 부장님) — 기술영업 격주휴무를 달력에서 고르고 결재받는다.
규칙은 services/sat_off.py. 사람 구분은 휴가와 같이 슬랙 ID(leave_web._person_and_sid)로 한다.
상태 바꾸는 API 는 /api/ 로 시작해야 보안 미들웨어를 통과한다.
"""
from datetime import date

from flask import Blueprint, jsonify, redirect, render_template, request, session

from dashboard.blueprints.hr import BUG_ERRORS, _hr_admin, _is_admin, _is_field, _me, gw_page
from dashboard.services import hr
from dashboard.services import sat_off as so
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

gw_satoff_bp = Blueprint('gw_satoff', __name__)
# 개발 중 문(dev_preview_gate)은 뗐다 — 시공·미화를 뺀 모두에게 연다(11-25 부장님). 고르기·결재는 그대로 역할별
CLOSED_DEPTS = ('시공', '미화')   # 부서 이름에 들어 있으면 안 보임(기술시공(내부)·미화). 현장 세션은 groupware 문지기가 따로 막음


def _who():
    """(이름, 슬랙 ID, 못 쓰는 이유)."""
    from dashboard.services import leave_web
    p, sid, why = leave_web._person_and_sid(session.get('user') or {})
    return ((p or {}).get('name') or ''), sid, why


def can_view_email(email: str) -> bool:
    """볼 수 있나(메뉴·화면·API 공용) — 시공·미화를 뺀 재직자(11-25 부장님). 기술영업·취합 담당·승인권자·전사 휴가 보기는
    부서와 상관없이. 고르기·결재 같은 고치는 일은 서비스가 역할로 다시 확인한다(그 밖의 직원은 보기만)."""
    if not email:
        return False
    if hr.can_view_all_leave(email):
        return True
    p = hr.find_by_email(email)
    if not p or p.get('status') == hr.STATUS_RESIGNED:
        return False
    dept = p.get('department') or ''
    if dept == so.MEMBER_DEPT or (p.get('slack_id') or '') in so.manager_ids():
        return True
    return not any(x in dept for x in CLOSED_DEPTS)


def _allowed(name: str, sid: str) -> bool:
    r = so.role(sid, name)
    return r['member'] or r['manager'] or r['approver'] or _hr_admin() or can_view_email(_me())


def _err(e: Exception):
    if isinstance(e, BUG_ERRORS):   # KeyError·IndexError·UnicodeError·JSONDecodeError = 코드 버그(11-24 조용한 실패 점검)
        logger.error(f'[토요휴무] 처리 실패(코드 오류): {type(e).__name__}: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '처리하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    if isinstance(e, PermissionError):
        return jsonify({'success': False, 'message': str(e)}), 403
    if isinstance(e, LookupError):
        return jsonify({'success': False, 'message': str(e)}), 404
    if isinstance(e, ValueError):
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.error(f'[토요휴무] 처리 실패: {e}', exc_info=True)
    return jsonify({'success': False, 'message': '처리하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500


def _actor():
    """상태를 바꾸는 요청 — 슬랙 계정이 확인된 사람만."""
    name, sid, why = _who()
    if why or not sid:
        raise PermissionError(why or '슬랙 계정을 확인하지 못했어요.')
    return name, sid


def _month(v) -> str:
    m = so.parse_month(v or '')
    if not m:
        raise ValueError('달 형식은 YYYY-MM 이에요.')
    return m


def _day(v):
    try:
        return date.fromisoformat(str(v or ''))
    except ValueError:
        return None


@gw_satoff_bp.route('/leave/saturday')
@gw_page
def page_satoff():
    if not can_view_email(_me()):
        return redirect('/leave')
    return render_template('gw/leave_saturday.html')


@gw_satoff_bp.route('/api/satoff', methods=['GET'])
@gw_page
def api_view():
    name, sid, why = _who()
    if not _allowed(name, sid):
        return jsonify({'success': False, 'message': '스케줄 근무를 볼 수 없는 계정이에요.'}), 403
    try:
        m = so.parse_month(request.args.get('month') or '') or date.today().strftime('%Y-%m')
        data = so.view(m, sid if not why else '', name, admin=_is_admin() and not _is_field())
    except Exception as e:
        return _err(e)
    data['why'] = why
    data['can_full_off'] = _is_admin() and not _is_field()      # 전체 휴무 지정·풀기 — 관리자(11-25)
    return jsonify({'success': True, **data})


@gw_satoff_bp.route('/api/satoff/toggle', methods=['POST'])
@gw_page
def api_toggle():
    body = request.get_json(silent=True) or {}
    try:
        name, sid = _actor()
        res = so.toggle(_day(body.get('day')), (body.get('name') or name).strip(), bool(body.get('on')), sid, name,
                        admin=_is_admin() and not _is_field())     # 관리자는 확정된 달도 바로 고침(11-25)
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, **res})


@gw_satoff_bp.route('/api/satoff/submit', methods=['POST'])
@gw_page
def api_submit():
    body = request.get_json(silent=True) or {}
    try:
        name, sid = _actor()
        m = _month(body.get('month'))
        so.submit(m, sid, name)
    except Exception as e:
        return _err(e)
    logger.info(f'[토요휴무] {m} 결재 올림 by {name}')
    so.notify('submit', m, name)
    return jsonify({'success': True})


@gw_satoff_bp.route('/api/satoff/decide', methods=['POST'])
@gw_page
def api_decide():
    body = request.get_json(silent=True) or {}
    approve = bool(body.get('approve'))
    note = (body.get('note') or '').strip()
    try:
        name, sid = _actor()
        m = _month(body.get('month'))
        so.decide(m, approve, sid, name, note)
    except Exception as e:
        return _err(e)
    logger.info(f'[토요휴무] {m} {"승인" if approve else "돌려보냄"} by {name}')
    so.notify('approve' if approve else 'return', m, name, note=note)
    return jsonify({'success': True, 'affected': so.affected_leaves(m) if approve else []})


@gw_satoff_bp.route('/api/satoff/change', methods=['POST'])
@gw_page
def api_change():
    body = request.get_json(silent=True) or {}
    try:
        name, sid = _actor()
        m = _month(body.get('month'))
        c = so.request_change(m, (body.get('name') or name).strip(), body.get('add') or [], body.get('remove') or [],
                              body.get('reason') or '', sid, name)
    except Exception as e:
        return _err(e)
    if c['status'] == so.CH_WAIT:
        so.notify('change_request', m, name, change=c)
    else:
        so.notify('change_approve', m, name, change=c)
    return jsonify({'success': True, 'change': c})


@gw_satoff_bp.route('/api/satoff/change/<int:cid>/decide', methods=['POST'])
@gw_page
def api_change_decide(cid):
    body = request.get_json(silent=True) or {}
    approve = bool(body.get('approve'))
    note = (body.get('note') or '').strip()
    try:
        name, sid = _actor()
        c = so.decide_change(cid, approve, sid, name, note)
    except Exception as e:
        return _err(e)
    so.notify('change_approve' if approve else 'change_reject', c['month'], name, change=c, note=note)
    return jsonify({'success': True, 'change': c})


@gw_satoff_bp.route('/api/satoff/worked', methods=['POST'])
@gw_page
def api_worked():
    """휴무 출근 기록(on=true)·되돌리기(on=false) — 구두 승인 뒤 본인이(11-24 부장님). 결재 단계·승인권자 알림 없음."""
    body = request.get_json(silent=True) or {}
    on = body.get('on', True) is not False
    try:
        name, sid = _actor()
        target = (body.get('name') or name).strip()
        day = _day(body.get('day'))
        if on:
            res = so.mark_worked(day, target, body.get('verbal_by') or '', body.get('reason') or '', sid, name)
        else:
            res = so.undo_worked(day, target, sid, name)
        m = res['day'][:7]
    except Exception as e:
        return _err(e)
    logger.info(f'[토요휴무] {res["day"]} {target} 휴무 출근 {"기록" if on else "되돌림"} by {name}')
    try:
        if so.affected_leaves(m):     # 휴가 일수가 달라지는 신청이 있을 때만 경영지원에 알림
            so.notify('worked' if on else 'worked_undo', m, name,
                      change={'name': target, 'label': so.kday(day)})
    except Exception as e:            # 기록은 저장됨 — 알림 확인만 실패
        logger.error(f'[토요휴무] 휴무 출근 뒤 휴가 일수 확인 실패 ({target} {res["day"]}): {e}', exc_info=True,
                     extra={'force_alert': True})
    return jsonify({'success': True, **res})


@gw_satoff_bp.route('/api/satoff/memo', methods=['POST'])
@gw_page
def api_memo():
    """휴무 메모 쓰기(on 생략·true)·지우기(on=false) — 꼭 쉬어야 해요·못 쉬어요 + 사유(11-25 부장님)."""
    body = request.get_json(silent=True) or {}
    on = body.get('on', True) is not False
    try:
        name, sid = _actor()
        target = name                    # 메모는 로그인한 본인 것만(11-25 부장님) — 화면이 보낸 이름은 쓰지 않음
        admin = _is_admin() and not _is_field()
        if on:
            res = so.write_memo(_day(body.get('day')), target, body.get('kind') or '', body.get('reason') or '', sid, name,
                                admin=admin)
        else:
            res = so.delete_memo(_day(body.get('day')), target, sid, name, admin=admin)
    except Exception as e:
        return _err(e)
    logger.info(f'[토요휴무] {res["day"]} {target} 메모 {"적음" if on else "지움"} by {name}')
    return jsonify({'success': True, **res})


@gw_satoff_bp.route('/api/satoff/month-note', methods=['POST'])
@gw_page
def api_month_note():
    body = request.get_json(silent=True) or {}
    try:
        name, sid = _actor()
        res = so.set_month_note(_month(body.get('month')), body.get('text') or '', sid, name,
                                admin=_is_admin() and not _is_field())
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, **res})


@gw_satoff_bp.route('/api/satoff/full', methods=['POST'])
@gw_page
def api_full_off():
    """전체 휴무 지정(on=true)·풀기(on=false) — 관리자(admin·super_admin)만(11-25 부장님)."""
    if not _is_admin() or _is_field():
        return jsonify({'success': False, 'message': '전체 휴무는 관리자만 정할 수 있어요.'}), 403
    body = request.get_json(silent=True) or {}
    on = body.get('on', True) is not False
    try:
        name, _sid = _who()[:2]
        res = so.set_full_off(_day(body.get('day')), on, name or _me(), body.get('note') or '')
    except Exception as e:
        return _err(e)
    logger.info(f'[토요휴무] {res["day"]} 전체 휴무 {"지정" if on else "풂"} by {name or _me()} ({res["count"]}명)')
    m = res['day'][:7]
    try:
        if so.month_row(m)['status'] == so.CONFIRMED and so.affected_leaves(m):
            so.notify('full_off', m, name, change={'label': so.kday(_day(body.get('day'))), 'how': '지정' if on else '풂'})
    except Exception as e:            # 저장은 됨 — 알림 확인만 실패
        logger.error(f'[토요휴무] 전체 휴무 뒤 휴가 일수 확인 실패 ({res["day"]}): {e}', exc_info=True, extra={'force_alert': True})
    return jsonify({'success': True, **res})


@gw_satoff_bp.route('/api/satoff/change/<int:cid>/cancel', methods=['POST'])
@gw_page
def api_change_cancel(cid):
    try:
        name, sid = _actor()
        c = so.cancel_change(cid, sid, name)
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, 'change': c})
