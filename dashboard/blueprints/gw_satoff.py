"""그룹웨어 「토요일 휴무」(2026-11-20 부장님) — 기술영업 격주휴무를 달력에서 고르고 결재받는다.
규칙은 services/sat_off.py. 사람 구분은 휴가와 같이 슬랙 ID(leave_web._person_and_sid)로 한다.
상태 바꾸는 API 는 /api/ 로 시작해야 보안 미들웨어를 통과한다.
"""
from datetime import date

from flask import Blueprint, jsonify, redirect, render_template, request, session

from dashboard.blueprints.hr import BUG_ERRORS, _hr_admin, _me, dev_preview_gate, gw_page
from dashboard.services import hr
from dashboard.services import sat_off as so
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

gw_satoff_bp = Blueprint('gw_satoff', __name__)
gw_satoff_bp.before_request(dev_preview_gate)   # 개발 중 — 부장님만(11-21)


def _who():
    """(이름, 슬랙 ID, 못 쓰는 이유)."""
    from dashboard.services import leave_web
    p, sid, why = leave_web._person_and_sid(session.get('user') or {})
    return ((p or {}).get('name') or ''), sid, why


def can_view_email(email: str) -> bool:
    """메뉴 보이기용(가벼운 확인) — 기술영업·취합 담당·승인권자·경영지원·전사 휴가 보기. 실제 권한은 API 가 다시 확인."""
    if not email:
        return False
    if hr.can_view_all_leave(email):
        return True
    p = hr.find_by_email(email)
    if not p:
        return False
    return p.get('department') == so.MEMBER_DEPT or (p.get('slack_id') or '') in so.manager_ids()


def _allowed(name: str, sid: str) -> bool:
    r = so.role(sid, name)
    return r['member'] or r['manager'] or r['approver'] or _hr_admin() or hr.can_view_all_leave(_me())


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
        return jsonify({'success': False, 'message': '토요일 휴무는 기술영업과 결재하는 분만 볼 수 있어요.'}), 403
    try:
        m = so.parse_month(request.args.get('month') or '') or date.today().strftime('%Y-%m')
        data = so.view(m, sid if not why else '', name)
    except Exception as e:
        return _err(e)
    data['why'] = why
    return jsonify({'success': True, **data})


@gw_satoff_bp.route('/api/satoff/toggle', methods=['POST'])
@gw_page
def api_toggle():
    body = request.get_json(silent=True) or {}
    try:
        name, sid = _actor()
        res = so.toggle(_day(body.get('day')), (body.get('name') or name).strip(), bool(body.get('on')), sid, name)
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


@gw_satoff_bp.route('/api/satoff/change/<int:cid>/cancel', methods=['POST'])
@gw_page
def api_change_cancel(cid):
    try:
        name, sid = _actor()
        c = so.cancel_change(cid, sid, name)
    except Exception as e:
        return _err(e)
    return jsonify({'success': True, 'change': c})
