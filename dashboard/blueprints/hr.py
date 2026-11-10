"""그룹웨어 인사 화면 — 내 정보·구성원(인사카드)·그룹웨어 설정.

정식 오픈 전까지는 시범 명단(그룹웨어 새 화면을 쓸 수 있는 사람)만 들어온다.
API 는 /api/ 아래(CSRF 경로 규칙). 권한은 services/hr.py 가 다시 확인한다.
"""
import json
from functools import wraps

from flask import Blueprint, jsonify, redirect, render_template, request, session

from dashboard.auth import login_required
from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

hr_bp = Blueprint('hr', __name__)


def _me() -> str:
    return ((session.get('user') or {}).get('email') or '').lower()


def _is_admin() -> bool:
    return (session.get('user') or {}).get('permission_level') in ('admin', 'super_admin')


def _is_field() -> bool:
    return bool((session.get('user') or {}).get('gw_field'))


def _hr_admin() -> bool:
    """HR 관리자인가 — 현장 세션은 절대 아님."""
    return not _is_field() and hr.is_hr_admin(_me())


def _my_person():
    """로그인한 사람의 인사카드 — 현장 세션은 인사카드 번호로, 그 외는 이메일로."""
    user = session.get('user') or {}
    if user.get('gw_field'):
        p = hr.get_person(int(user.get('gw_person_id') or 0))
        return p if p and p.get('status') != hr.STATUS_RESIGNED else None
    return hr.find_by_email(_me())


def gw_page(f):
    """그룹웨어 화면 — 로그인 + 새 화면을 쓸 수 있는 사람(정식 오픈 전엔 시범 명단).
    현장 세션은 링크 발급 때 이미 확인했고, 다닐 수 있는 경로는 groupware 문지기가 막는다."""
    @login_required
    @wraps(f)
    def wrapper(*a, **kw):
        if not (session.get('user') or {}).get('gw_field') and not gw.can_use(_me()):
            if request.path.startswith('/api/'):
                return jsonify({'success': False, 'message': '아직 열리지 않은 화면입니다.'}), 403
            return redirect('/projects')
        return f(*a, **kw)
    return wrapper


def _fail(e: Exception):
    if isinstance(e, PermissionError):
        return jsonify({'success': False, 'message': str(e)}), 403
    if isinstance(e, LookupError):
        return jsonify({'success': False, 'message': str(e)}), 404
    if isinstance(e, (ValueError, hr.SsnKeyMissing)):
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.error(f'[HR] 처리 실패: {e}', exc_info=True)
    return jsonify({'success': False, 'message': '처리하지 못했습니다. 잠시 후 다시 해 주세요.'}), 500


# ── 화면 ──────────────────────────────────────────────────────────────
@hr_bp.route('/hr/me')
@gw_page
def page_me():
    p = _my_person()
    return render_template('gw/hr_me.html', person_id=p['id'] if p else None, fields=hr.FIELDS)


@hr_bp.route('/hr/people')
@gw_page
def page_people():
    if not _hr_admin():
        return redirect('/hr/me')
    return render_template('gw/hr_people.html', fields=hr.FIELDS, statuses=hr.STATUSES)


@hr_bp.route('/gw/admin')
@gw_page
def page_admin():
    if not _is_admin():
        return redirect('/hr/me')
    return render_template('gw/admin.html')


# ── 인사 API ──────────────────────────────────────────────────────────
@hr_bp.route('/api/hr/people', methods=['GET'])
@gw_page
def api_list():
    if not _hr_admin():
        return jsonify({'success': False, 'message': '구성원 목록은 경영지원·대표만 볼 수 있습니다.'}), 403
    return jsonify({'success': True, 'people': hr.list_people(include_resigned=request.args.get('resigned') == '1')})


@hr_bp.route('/api/hr/people/<int:pid>', methods=['GET'])
@gw_page
def api_card(pid):
    me = '' if _is_field() else _me()      # 현장 세션은 관리자 권한 판단에 이메일을 쓰지 않음
    if not _hr_admin():
        mine = _my_person()
        if not mine or mine['id'] != pid:
            return jsonify({'success': False, 'message': '본인 인사카드만 볼 수 있습니다.'}), 403
    c = hr.card(pid, me)
    if not c:
        return jsonify({'success': False, 'message': '없는 구성원입니다.'}), 404
    if _hr_admin():
        c['access_log'] = hr.access_log(pid)
    return jsonify({'success': True, 'person': c})


@hr_bp.route('/api/hr/people/<int:pid>', methods=['PUT'])
@gw_page
def api_update(pid):
    try:
        c = hr.update_person(pid, request.get_json(silent=True) or {}, _me())
        logger.info(f'[HR] 인사카드 수정 #{pid} by {_me()}')
        return jsonify({'success': True, 'person': c})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/hr/people', methods=['POST'])
@gw_page
def api_create():
    try:
        c = hr.create_person(request.get_json(silent=True) or {}, _me())
        logger.info(f'[HR] 구성원 추가 #{c["id"]} by {_me()}')
        return jsonify({'success': True, 'person': c})
    except Exception as e:
        return _fail(e)


@hr_bp.route('/api/hr/people/<int:pid>/ssn', methods=['POST'])
@gw_page
def api_reveal_ssn(pid):
    try:
        full = hr.reveal_ssn(pid, _me())
        logger.info(f'[HR] 주민번호 전체 보기 #{pid} by {_me()}')
        return jsonify({'success': True, 'ssn': full})
    except Exception as e:
        return _fail(e)


# ── 그룹웨어 설정 API(관리자) ──────────────────────────────────────────
@hr_bp.route('/api/gw/roles', methods=['GET'])
@gw_page
def api_roles():
    if not _is_admin():
        return jsonify({'success': False, 'message': '관리자만 볼 수 있습니다.'}), 403
    return jsonify({'success': True, 'hr_admins': hr.hr_admins(), 'leave_viewers': hr.leave_viewers(),
                    'pilots': gw.list_pilots(), 'open_all': gw.is_open_all()})


@hr_bp.route('/api/gw/roles', methods=['POST'])
@gw_page
def api_set_roles():
    """{"key": "hr_admins"|"leave_viewers", "emails": [...]}"""
    if not _is_admin():
        return jsonify({'success': False, 'message': '관리자만 바꿀 수 있습니다.'}), 403
    data = request.get_json(silent=True) or {}
    key = data.get('key')
    if key not in ('hr_admins', 'leave_viewers'):
        return jsonify({'success': False, 'message': '바꿀 수 없는 항목입니다.'}), 400
    emails = sorted({str(e).strip().lower() for e in (data.get('emails') or []) if str(e).strip()})
    bad = [e for e in emails if '@' not in e]
    if bad:
        return jsonify({'success': False, 'message': f'이메일 형식이 아닙니다: {", ".join(bad)}'}), 400
    if key == 'hr_admins' and not emails:
        return jsonify({'success': False, 'message': '인사카드 관리자는 한 명 이상 있어야 합니다.'}), 400
    gw.set_setting(key, json.dumps(emails), _me())
    logger.info(f'[GW] {key} = {emails} by {_me()}')
    return jsonify({'success': True, key: emails})
