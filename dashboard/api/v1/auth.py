"""
API v1 - 인증 관련 엔드포인트
CSRF 토큰, 세션 검증 등
"""

import logging
from flask import Blueprint, jsonify, session
from dashboard.utils.security_middleware import csrf_protection

logger = logging.getLogger(__name__)

# Blueprint 생성
auth_bp = Blueprint('auth_api', __name__, url_prefix='/api/v1/auth')


@auth_bp.route('/csrf-token', methods=['GET'])
def get_csrf_token():
    """
    CSRF 토큰 조회 엔드포인트

    Returns:
        JSON: { "success": true, "data": {"token": "csrf_token_value"} }

    Note:
        - API 클라이언트나 SPA에서 사용
        - 세션이 있어야 토큰 생성 가능
        - 생성된 토큰은 X-CSRF-Token 헤더로 전송 필요

    Example:
        GET /api/v1/auth/csrf-token
        Response: {
            "success": true,
            "data": {"token": "abc123..."}
        }

        // 사용 예시
        const response = await fetch('/api/v1/auth/csrf-token');
        const { success, data } = await response.json();
        if (success) {
            const token = data.token;
            // 토큰 사용
            fetch('/api/v1/projects', {
                method: 'POST',
                headers: {
                    'X-CSRF-Token': token,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(projectData)
            });
        }
    """
    try:
        # 세션 ID 가져오기 (없으면 새로 생성)
        session_id = session.get('session_id')
        if not session_id:
            # 세션 없으면 경고
            logger.warning("CSRF 토큰 요청 시 세션 없음")
            return jsonify({
                'success': False,
                'error': '세션이 필요합니다. 먼저 로그인해주세요.',
                'code': 'SESSION_REQUIRED'
            }), 401

        # CSRF 토큰 생성
        csrf_token = csrf_protection.generate_token(session_id)

        return jsonify({
            'success': True,
            'data': {
                'token': csrf_token
            }
        }), 200

    except Exception as e:
        logger.error(f"CSRF 토큰 생성 실패: {e}")
        return jsonify({
            'success': False,
            'error': 'CSRF 토큰 생성 중 오류가 발생했습니다.',
            'code': 'TOKEN_GENERATION_FAILED'
        }), 500


@auth_bp.route('/session', methods=['GET'])
def check_session():
    """
    세션 유효성 확인 엔드포인트

    Returns:
        JSON: { "success": true, "data": {"authenticated": true/false, "user": {...}} }

    Note:
        - 현재 세션이 유효한지 확인
        - 로그인 상태 체크용
    """
    try:
        if 'user' in session:
            return jsonify({
                'success': True,
                'data': {
                    'authenticated': True,
                    'user': {
                        'email': session['user'].get('email'),
                        'name': session['user'].get('name'),
                        'role': session['user'].get('permission_level', 'user')
                    }
                }
            }), 200
        else:
            return jsonify({
                'success': True,
                'data': {
                    'authenticated': False
                }
            }), 200

    except Exception as e:
        logger.error(f"세션 확인 실패: {e}")
        return jsonify({
            'success': False,
            'error': '세션 확인 중 오류가 발생했습니다.',
            'code': 'SESSION_CHECK_FAILED'
        }), 500
