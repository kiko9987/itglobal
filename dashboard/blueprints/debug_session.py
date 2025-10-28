"""
임시 세션 디버깅 엔드포인트
"""

from flask import Blueprint, session, jsonify

debug_bp = Blueprint('debug', __name__, url_prefix='/api/debug')

@debug_bp.route('/session')
def debug_session():
    """현재 세션 정보 반환"""
    return jsonify({
        'session': dict(session),
        'user': session.get('user', {}),
        'user_permission_level': session.get('user', {}).get('permission_level'),
        'user_email': session.get('user', {}).get('email'),
        'login_time': session.get('login_time')
    })