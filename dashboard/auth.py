"""
사용자 인증 시스템
SQLite 기반 사용자 관리 (파일 경쟁 상태 해결, 보안 강화)
Google OAuth 전용 인증 시스템
"""

import json
import os
from datetime import datetime, timedelta
from functools import wraps
from flask import session, request, redirect, url_for, jsonify
from .utils.user_database import get_user_database
from .utils.logging_config import get_logger

logger = get_logger(__name__)

class UserManager:
    def __init__(self, users_file='users.json'):
        # 통합된 데이터베이스 시스템 사용 (instance/users.db)
        self.db = get_user_database()
        self.legacy_users_file = os.path.join(os.path.dirname(__file__), users_file)
        self._migrate_if_needed()

    def _migrate_if_needed(self):
        """기존 JSON 파일이 있으면 SQLite로 마이그레이션"""
        if os.path.exists(self.legacy_users_file):
            success, message = self.db.migrate_from_json(self.legacy_users_file)
            if success:
                print(f"[MIGRATION] {message}")
                # 마이그레이션 성공 시 JSON 파일을 백업으로 이동
                backup_file = self.legacy_users_file + '.backup'
                if not os.path.exists(backup_file):
                    try:
                        os.rename(self.legacy_users_file, backup_file)
                        print(f"[MIGRATION] JSON 파일을 {backup_file}로 백업했습니다.")
                    except Exception as e:
                        print(f"[MIGRATION] 백업 실패: {e}")
            else:
                print(f"[MIGRATION ERROR] {message}")

    # Google OAuth 전용: 비밀번호 기반 인증 제거됨
    # authenticate_user() 메서드 삭제 - Google OAuth만 사용

    def authenticate_google_user(self, google_user_info):
        """구글 OAuth 사용자 인증 및 자동 등록 (SQLite 기반)"""
        return self.db.authenticate_google_user(google_user_info)

    def get_user_by_email(self, email):
        """이메일로 사용자 정보 가져오기 (SQLite 기반)"""
        return self.db.get_user_by_email(email)

    def get_all_users(self):
        """모든 사용자 정보 가져오기 (SQLite 기반)"""
        return self.db.get_all_users()

    # Google OAuth 전용: 비밀번호 기반 사용자 생성 제거됨
    # create_user() 메서드 삭제 - Google OAuth로 자동 등록만 사용

    def update_user_permission(self, email, permission_level):
        """사용자 권한 업데이트 (SQLite 기반)"""
        return self.db.update_user_permission(email, permission_level)

    def toggle_user_status(self, email, is_active):
        """사용자 활성화/비활성화 (SQLite 기반)"""
        return self.db.toggle_user_status(email, is_active)

    def delete_user(self, email):
        """사용자 삭제 (SQLite 기반)"""
        return self.db.delete_user(email)

# 전역 사용자 매니저 인스턴스
user_manager = UserManager()

def _validate_session(require_admin=False):
    """세션 검증 공통 로직"""
    if 'user' not in session:
        if request.is_json:
            return jsonify({'error': '로그인이 필요합니다.', 'redirect': '/login'}), 401
        return redirect(url_for('auth.login_page', message='access_denied'))
    
    # 세션 유효성 검사
    user = session['user']
    if not user.get('email') or not user.get('name'):
        session.clear()
        if request.is_json:
            return jsonify({'error': '세션이 만료되었습니다.', 'redirect': '/login'}), 401
        return redirect(url_for('auth.login_page', message='session_expired'))
    
    # 로그인 시간 기반 세션 검증
    SESSION_TIMEOUT_HOURS = int(os.getenv('SESSION_TIMEOUT_HOURS', 4))  # 기본 4시간 (보안 권장)
    login_time_str = session.get('login_time')
    if login_time_str:
        try:
            login_time = datetime.fromisoformat(login_time_str)
            if datetime.now() - login_time > timedelta(hours=SESSION_TIMEOUT_HOURS):
                session.clear()
                if request.is_json:
                    return jsonify({'error': '세션이 만료되었습니다.', 'redirect': '/login'}), 401
                return redirect(url_for('auth.login_page', message='session_expired'))
        except (ValueError, TypeError) as e:
            # 잘못된 날짜 형식으로 세션 무효화
            logger.warning(f"세션 시간 파싱 실패: {e}")
            session.clear()
            if request.is_json:
                return jsonify({'error': '세션이 만료되었습니다.', 'redirect': '/login'}), 401
            return redirect(url_for('auth.login_page', message='session_expired'))

    # 실시간 DB 확인: 사용자 상태 및 권한 검증
    try:
        db = get_user_database()
        current_user = db.get_user_by_email(user.get('email'))

        # 사용자가 DB에서 삭제됨
        if not current_user:
            session.clear()
            if request.is_json:
                return jsonify({'error': '사용자 계정을 찾을 수 없습니다.', 'redirect': '/login'}), 401
            return redirect(url_for('auth.login_page', message='user_deleted'))

        # 사용자가 비활성화됨
        if not current_user.get('is_active', True):
            session.clear()
            if request.is_json:
                return jsonify({'error': '계정이 비활성화되었습니다.', 'redirect': '/login'}), 401
            return redirect(url_for('auth.login_page', message='account_disabled'))

        # 사용자가 퇴사 처리됨
        if current_user.get('is_resigned', False):
            session.clear()
            if request.is_json:
                return jsonify({'error': '퇴사 처리된 계정입니다.', 'redirect': '/login'}), 401
            return redirect(url_for('auth.login_page', message='account_resigned'))

        # 세션 무효화 타임스탬프 확인 (권한/상태 변경 시 모든 세션 강제 종료)
        session_invalidate_after = current_user.get('session_invalidate_after')
        if session_invalidate_after and login_time_str:
            try:
                # SQLite timestamp를 datetime으로 변환
                invalidate_time = datetime.fromisoformat(session_invalidate_after.replace(' ', 'T'))
                login_time = datetime.fromisoformat(login_time_str)

                # 로그인 시간이 무효화 시점보다 이전이면 세션 종료
                if login_time < invalidate_time:
                    session.clear()
                    if request.is_json:
                        return jsonify({'error': '권한이 변경되어 다시 로그인해야 합니다.', 'redirect': '/login'}), 401
                    return redirect(url_for('auth.login_page', message='session_invalidated'))
            except Exception as e:
                import logging
                logging.warning(f"세션 무효화 타임스탬프 비교 실패: {e}")

        # 세션의 권한 정보를 DB와 동기화 (권한 변경 즉시 반영)
        db_permission = current_user.get('permission_level', 'viewer')
        if session['user'].get('permission_level') != db_permission:
            session['user']['permission_level'] = db_permission
            session.modified = True  # nested dict 변경 시 필수!

    except Exception as e:
        # DB 조회 실패 시 경고 로그만 남기고 계속 진행 (DB 장애 시 전체 서비스 중단 방지)
        import logging
        logging.warning(f"세션 검증 중 DB 조회 실패: {e}")

    # 관리자 권한 검사
    permission_level = session['user'].get('permission_level', '')
    if require_admin and permission_level not in ['admin', 'super_admin']:
        if request.is_json:
            return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
        return redirect(url_for('projects.project_list'))
    
    return None

def login_required(f):
    """로그인 필수 데코레이터"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        validation_result = _validate_session()
        if validation_result:
            return validation_result
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    """관리자 권한 필수 데코레이터"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        validation_result = _validate_session(require_admin=True)
        if validation_result:
            return validation_result
        return f(*args, **kwargs)
    return decorated_function

def get_user_role():
    """사용자 역할 결정 (DB 기반)"""
    if 'user' not in session:
        return 'viewer'

    user = session['user']
    permission = user.get('permission_level', 'viewer')

    # DB 권한만 사용 (user_permissions.json 제거됨)
    # super_admin도 프론트엔드에서는 Admin으로 표시
    if permission in ['admin', 'super_admin']:
        return 'Admin'
    elif permission == 'editor':
        return 'Editor'
    else:
        return 'Viewer'