"""
Stats Blueprint - 매출 통계 페이지
모던 시스템으로 재구현
"""

from flask import Blueprint, render_template, jsonify, session
from functools import wraps
import logging

# Blueprint 생성
stats_bp = Blueprint('stats', __name__, url_prefix='')

logger = logging.getLogger(__name__)

def login_required(f):
    """로그인 필요 데코레이터"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            return jsonify({'error': 'Unauthorized'}), 401
        return f(*args, **kwargs)
    return decorated_function

def get_user_role():
    """현재 사용자의 역할 반환"""
    user = session.get('user', {})
    email = user.get('email', '')

    # 최고관리자 확인
    if email == 'ceo@itglobal.com':
        return 'supreme_admin'

    # 관리자 확인
    admin_emails = ['admin@itglobal.com', 'kiko@itglobal.com']
    if email in admin_emails:
        return 'admin'

    return 'user'

@stats_bp.route('/stats')
@login_required
def stats_page():
    """통계 페이지"""
    try:
        user_role = get_user_role()
        user_email = session.get('user', {}).get('email', '')

        return render_template('modern_stats.html',
                               user_role=user_role,
                               user_email=user_email,
                               user=session.get('user', {}))
    except Exception as e:
        logger.error(f"통계 페이지 로드 오류: {str(e)}")
        return f"통계 페이지를 로드할 수 없습니다: {str(e)}", 500

@stats_bp.route('/api/user/role')
@login_required
def get_user_role_api():
    """사용자 역할 정보 반환"""
    try:
        user_role = get_user_role()
        return jsonify({'success': True, 'role': user_role})
    except Exception as e:
        logger.error(f"사용자 역할 조회 오류: {str(e)}")
        return jsonify({'success': False, 'role': 'user'}), 500
