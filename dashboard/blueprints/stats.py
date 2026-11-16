"""
Stats Blueprint - 매출 통계 페이지
모던 시스템으로 재구현
"""

from flask import Blueprint, render_template, jsonify, session
import logging
from dashboard.auth import login_required
from dashboard.utils.error_helpers import generate_error_id

# Blueprint 생성
stats_bp = Blueprint('stats', __name__, url_prefix='')

logger = logging.getLogger(__name__)

# 로그인 확인은 다른 화면과 같은 공용 login_required — 예전 자체 데코레이터는 페이지 요청에도
#   영어 JSON 401 만 돌려줘 세션이 끝난 직원이 로그인 화면으로 못 가고, 세션 만료·퇴사 확인도 안 했다(최종 점검 11-16)

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
        # 오류 원문은 서버 로그에만 — 화면에 내부 경로·예외 내용이 그대로 보였음(최종 점검 11-16)
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 통계 페이지 로드 오류: {e}", exc_info=True)
        return "통계를 불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.", 500

@stats_bp.route('/api/user/role')
@login_required
def get_user_role_api():
    """사용자 역할 정보 반환"""
    try:
        user_role = get_user_role()
        return jsonify({'success': True, 'role': user_role})
    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 사용자 역할 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'role': 'user',
            'error_id': error_id
        }), 500
