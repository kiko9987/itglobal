"""
Authentication Blueprint
사용자 인증 관련 라우트들을 관리하는 블루프린트
"""

import os
import uuid
from datetime import datetime
from flask import Blueprint, render_template, redirect, url_for, session, request, jsonify
from dashboard.utils.logging_config import get_logger
from dashboard.utils.error_helpers import generate_error_id

# 인증 시스템 import
from dashboard.auth import user_manager, login_required, admin_required, get_user_role
from dashboard.google_oauth import get_google_oauth, is_oauth_configured
from dashboard.utils.google_calendar import (
    create_flow as create_calendar_flow,
    save_credentials as save_calendar_credentials,
    CalendarConfigurationError,
)

logger = get_logger(__name__)

# 블루프린트 생성
auth_bp = Blueprint('auth', __name__)

@auth_bp.route('/login')
def login_page():
    """로그인 페이지 (Google OAuth 전용)"""
    if 'user' in session:
        return redirect(url_for('projects.project_list'))

    # OAuth가 설정되지 않은 경우 오류 표시
    oauth_enabled = is_oauth_configured()
    if not oauth_enabled:
        logger.error("Google OAuth가 설정되지 않았습니다.")

    return render_template('login.html', oauth_enabled=oauth_enabled)


@auth_bp.route('/auth/google')
def google_login():
    """구글 OAuth 로그인 시작"""
    try:
        if not is_oauth_configured():
            return redirect(url_for('auth.login_page', message='oauth_not_configured'))

        authorization_url, state = get_google_oauth().get_authorization_url(
            redirect_uri=url_for('auth.google_callback', _external=True)
        )

        if not authorization_url:
            return redirect(url_for('auth.login_page', message='oauth_error'))

        session['oauth_state'] = state
        return redirect(authorization_url)

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] Google OAuth 시작 오류: {str(e)}", exc_info=True)
        return redirect(url_for('auth.login_page', message='oauth_error'))

@auth_bp.route('/auth/callback')
def google_callback():
    """구글 OAuth 콜백"""
    try:
        if not is_oauth_configured():
            return redirect(url_for('auth.login_page', message='oauth_not_configured'))

        # 상태 검증
        state = request.args.get('state')
        if state != session.get('oauth_state'):
            logger.warning("OAuth state 불일치")
            return redirect(url_for('auth.login_page', message='oauth_error'))

        # 에러 체크
        error = request.args.get('error')
        if error:
            logger.warning(f"OAuth 에러: {error}")
            return redirect(url_for('auth.login_page', message='oauth_cancelled'))

        # 인증 코드 가져오기
        code = request.args.get('code')
        if not code:
            return redirect(url_for('auth.login_page', message='oauth_error'))

        # 사용자 정보 가져오기 (credentials도 함께)
        user_info, credentials = get_google_oauth().get_user_info(
            code, state,
            redirect_uri=url_for('auth.google_callback', _external=True)
        )

        if not user_info:
            return redirect(url_for('auth.login_page', message='domain_not_allowed'))

        # 사용자 인증 및 자동 등록
        user = user_manager.authenticate_google_user(user_info)
        if not user:
            return redirect(url_for('auth.login_page', message='login_failed'))

        # 세션 생성 - 최소 정보만 저장 (보안)
        session['user'] = {
            'email': user['email'],
            'name': user['name'],
            'permission_level': user.get('permission_level', 'viewer'),
            'is_active': user.get('is_active', True),
            'auth_method': user.get('auth_method', 'google'),
            'picture': user.get('picture', '')
            # password_hash는 저장하지 않음
        }
        session['login_time'] = datetime.now().isoformat()
        session['session_id'] = str(uuid.uuid4())  # CSRF 토큰 생성을 위한 세션 ID
        session.permanent = True

        # 세션 타임아웃 설정 (프론트엔드와 동기화)
        session_timeout_hours = int(os.getenv('SESSION_TIMEOUT_HOURS', 4))
        session['session_timeout_minutes'] = session_timeout_hours * 60

        # OAuth state 정리
        session.pop('oauth_state', None)

        # 캘린더 credentials 저장 (로그인과 동시에 캘린더 연동)
        if credentials:
            try:
                save_calendar_credentials(credentials)
                logger.info(f"[Calendar] 로그인 시 캘린더 토큰 저장 완료: {user['email']}")
            except Exception as e:
                logger.warning(f"[Calendar] 토큰 저장 실패 (로그인은 성공): {e}")

        logger.info(f"Google OAuth 로그인 성공: {user['email']}")
        return redirect(url_for('projects.project_list'))

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] Google OAuth 콜백 오류: {str(e)}", exc_info=True)
        return redirect(url_for('auth.login_page', message='oauth_error'))

@auth_bp.route('/logout')
def logout():
    """로그아웃 처리"""
    session.clear()
    return redirect(url_for('auth.login_page', message='logout_success'))


@auth_bp.route('/api/auth/ping', methods=['POST'])
@login_required
def refresh_session():
    """세션 연장 전용 API - 사용자 활동 시 login_time 갱신"""
    try:
        session['login_time'] = datetime.now().isoformat()
        session.modified = True
        logger.debug(f"세션 연장: {session.get('user', {}).get('email', 'unknown')}")

        # 세션 타임아웃 정보를 프론트엔드에 전달
        session_timeout_minutes = session.get('session_timeout_minutes')
        if not session_timeout_minutes:
            # 세션에 없으면 환경변수에서 계산
            session_timeout_hours = int(os.getenv('SESSION_TIMEOUT_HOURS', 4))
            session_timeout_minutes = session_timeout_hours * 60
            session['session_timeout_minutes'] = session_timeout_minutes
            session.modified = True

        return jsonify({
            'success': True,
            'message': '세션이 연장되었습니다',
            'session_timeout_minutes': session_timeout_minutes
        })
    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 세션 연장 실패: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '세션을 연장할 수 없습니다',
            'error_id': error_id
        }), 500


@auth_bp.route('/auth/google/calendar')
@login_required
def google_calendar_auth():
    """Google Calendar OAuth 시작"""
    try:
        flow = create_calendar_flow(
            redirect_uri=url_for('auth.google_calendar_callback', _external=True)
        )
    except CalendarConfigurationError as exc:
        logger.error(f"[Calendar OAuth] 설정 오류: {exc}")
        return redirect(url_for('projects.project_list', message='calendar_config_error'))

    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='consent'
    )
    # state 값을 세션에 저장 (CSRF 방지)
    session['calendar_oauth_state'] = state
    session.modified = True  # 세션 변경 강제 저장
    return redirect(authorization_url)


@auth_bp.route('/auth/google/calendar/callback')
@login_required
def google_calendar_callback():
    """Google Calendar OAuth 콜백"""
    error = request.args.get('error')
    if error:
        logger.warning(f"[Calendar OAuth] 사용자가 접근을 거부했습니다: {error}")
        return redirect(url_for('projects.project_list', message='calendar_oauth_cancelled'))

    state = request.args.get('state')
    stored_state = session.get('calendar_oauth_state')

    # state 검증 (CSRF 방지)
    if not state or state != stored_state:
        logger.warning(f"[Calendar OAuth] state 불일치 - 받은 값: {state}, 저장된 값: {stored_state}")
        return redirect(url_for('projects.project_list', message='calendar_oauth_state_error'))

    # 사용된 state는 제거
    session.pop('calendar_oauth_state', None)

    code = request.args.get('code')
    if not code:
        logger.warning("[Calendar OAuth] code 파라미터 없음")
        return redirect(url_for('projects.project_list', message='calendar_oauth_error'))

    try:
        flow = create_calendar_flow(
            redirect_uri=url_for('auth.google_calendar_callback', _external=True),
            state=state,
        )
        flow.fetch_token(authorization_response=request.url)
        save_calendar_credentials(flow.credentials)
        session.modified = True

        user_email = session.get('user', {}).get('email', 'unknown')
        logger.info(f"[Calendar OAuth] 캘린더 연동 완료: {user_email}")
        return redirect(url_for('projects.project_list', message='calendar_linked'))

    except CalendarConfigurationError as exc:
        logger.error(f"[Calendar OAuth] 설정 오류: {exc}")
        return redirect(url_for('projects.project_list', message='calendar_config_error'))
    except Exception as exc:
        logger.error(f"[Calendar OAuth] 토큰 발급 실패: {exc}")
        return redirect(url_for('projects.project_list', message='calendar_oauth_error'))
