"""그룹웨어 새 틀 — 화면 틀 고르기와 시범 명단 API.

- 모든 화면이 쓰는 modern_base.html 이 `gw_layout` 값으로 머리글만 바꿔 끼운다.
  명단에 없는 사람에겐 gw_layout=False → 기존 머리글 그대로.
- API 는 /api/ 아래(CSRF 경로 규칙).
"""
from flask import Blueprint, jsonify, request, session

from dashboard.auth import admin_required, login_required
from dashboard.services import groupware as gw
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

groupware_bp = Blueprint('groupware', __name__)


@groupware_bp.app_context_processor
def inject_groupware_layout():
    """화면을 그릴 때 새 틀을 쓸지와 메뉴를 넣어 준다. 실패하면 기존 틀."""
    try:
        user = session.get('user') or {}
        email = user.get('email', '')
        if not gw.can_use(email):
            return {'gw_layout': False}
        if session.get('gw_off'):
            # 시범 사용자가 「예전 화면」으로 돌린 상태 — 사용자 메뉴에 「새 화면으로」만 보임
            return {'gw_layout': False, 'gw_can_switch': True}
        from dashboard.services import hr
        return {
            'gw_layout': True,
            'gw_nav': gw.build_nav(request.path, user.get('permission_level', ''), hr.is_hr_admin(email)),
        }
    except Exception as e:
        logger.warning(f'[GW] 틀 고르기 실패 → 기존 틀: {e}')
        return {'gw_layout': False}


@groupware_bp.route('/api/gw/layout', methods=['POST'])
@login_required
def set_layout():
    """본인 화면 틀 바꾸기 — {"new": true|false}. 새 틀은 쓸 수 있는 사람만."""
    data = request.get_json(silent=True) or {}
    want_new = bool(data.get('new'))
    email = (session.get('user') or {}).get('email', '')
    if want_new and not gw.can_use(email):
        return jsonify({'success': False, 'message': '아직 새 화면을 쓸 수 없는 계정입니다.'}), 403
    if want_new:
        session.pop('gw_off', None)
    else:
        session['gw_off'] = True
    session.modified = True
    return jsonify({'success': True, 'new': want_new})


@groupware_bp.route('/api/gw/pilots', methods=['GET'])
@admin_required
def get_pilots():
    return jsonify({'success': True, 'pilots': gw.list_pilots(), 'open_all': gw.is_open_all()})


@groupware_bp.route('/api/gw/pilots', methods=['POST'])
@admin_required
def change_pilots():
    """{"email": "...", "on": true|false} 또는 {"open_all": true|false}."""
    data = request.get_json(silent=True) or {}
    by = (session.get('user') or {}).get('email', 'unknown')
    try:
        if 'open_all' in data:
            gw.set_open_all(bool(data['open_all']), by)
            logger.info(f'[GW] 전체 켜기={bool(data["open_all"])} by {by}')
        else:
            email = data.get('email', '')
            if data.get('on'):
                gw.add_pilot(email, by)
            else:
                gw.remove_pilot(email)
            logger.info(f'[GW] 시범 명단 {"추가" if data.get("on") else "제외"} {email} by {by}')
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    return jsonify({'success': True, 'pilots': gw.list_pilots(), 'open_all': gw.is_open_all()})
