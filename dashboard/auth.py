"""
사용자 인증 시스템
간단한 파일 기반 사용자 관리 (추후 DB로 확장 가능)
"""

import json
import os
import hashlib
import secrets
from datetime import datetime, timedelta
from functools import wraps
from flask import session, request, redirect, url_for, jsonify

class UserManager:
    def __init__(self, users_file='users.json'):
        self.users_file = os.path.join(os.path.dirname(__file__), users_file)
        self._ensure_users_file()
    
    def _ensure_users_file(self):
        """사용자 파일이 없으면 빈 파일 생성 (OAuth 로그인 시 자동 등록)"""
        if not os.path.exists(self.users_file):
            self._save_users({})
    
    def _hash_password(self, password):
        """비밀번호 해싱"""
        salt = secrets.token_hex(16)
        password_hash = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt.encode('utf-8'), 100000)
        return f"{salt}:{password_hash.hex()}"
    
    def _verify_password(self, stored_password, provided_password):
        """비밀번호 검증"""
        try:
            salt, password_hash = stored_password.split(':')
            provided_hash = hashlib.pbkdf2_hmac('sha256', provided_password.encode('utf-8'), salt.encode('utf-8'), 100000)
            return password_hash == provided_hash.hex()
        except:
            return False
    
    def _load_users(self):
        """사용자 데이터 로드"""
        try:
            with open(self.users_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {}
    
    def _save_users(self, users):
        """사용자 데이터 저장"""
        with open(self.users_file, 'w', encoding='utf-8') as f:
            json.dump(users, f, ensure_ascii=False, indent=2)
    
    def authenticate_user(self, email, password):
        """사용자 인증 (기존 방식)"""
        users = self._load_users()
        user = users.get(email)
        
        if not user:
            return None
        
        if not user.get('is_active', False):
            return None
        
        if not self._verify_password(user['password_hash'], password):
            return None
        
        # 로그인 시간 업데이트
        user['last_login'] = datetime.now().isoformat()
        users[email] = user
        self._save_users(users)
        
        # 비밀번호 해시 제거 후 반환
        safe_user = user.copy()
        del safe_user['password_hash']
        return safe_user
    
    def authenticate_google_user(self, google_user_info):
        """구글 OAuth 사용자 인증 및 자동 등록"""
        email = google_user_info['email'].lower()
        name = google_user_info['name']
        picture = google_user_info.get('picture', '')
        
        users = self._load_users()
        user = users.get(email)
        
        if not user:
            # 첫 번째 사용자인지 확인 (빈 사용자 파일이거나 OAuth 사용자가 없는 경우)
            oauth_users = [u for u in users.values() if u.get('auth_method') == 'google_oauth']
            is_first_oauth_user = len(oauth_users) == 0
            
            # 첫 번째 OAuth 사용자는 관리자 권한, 나머지는 viewer 권한
            permission_level = 'admin' if is_first_oauth_user else 'viewer'
            
            user = {
                'name': name,
                'email': email,
                'password_hash': '',  # Google OAuth 사용자는 비밀번호 없음
                'permission_level': permission_level,
                'is_active': True,
                'auth_method': 'google_oauth',
                'picture': picture,
                'created_at': datetime.now().isoformat(),
                'last_login': datetime.now().isoformat()
            }
            users[email] = user
            self._save_users(users)
            
            # 새 사용자 알림 로그
            if is_first_oauth_user:
                print(f"🎉 첫 번째 Google 사용자 등록: {name} ({email}) - 관리자 권한 자동 부여!")
            else:
                print(f"✅ 새 Google 사용자 등록: {name} ({email}) - viewer 권한")
            
        else:
            # 기존 사용자 로그인 시간 및 정보 업데이트
            if not user.get('is_active', False):
                return None
            
            user['last_login'] = datetime.now().isoformat()
            user['name'] = name  # 이름이 변경되었을 수 있음
            user['picture'] = picture  # 프로필 사진 업데이트
            user['auth_method'] = 'google_oauth'
            users[email] = user
            self._save_users(users)
        
        # 안전한 사용자 정보 반환
        safe_user = user.copy()
        if 'password_hash' in safe_user:
            del safe_user['password_hash']
        
        return safe_user
    
    def get_user_by_email(self, email):
        """이메일로 사용자 정보 가져오기"""
        users = self._load_users()
        user = users.get(email)
        if user:
            safe_user = user.copy()
            if 'password_hash' in safe_user:
                del safe_user['password_hash']
            return safe_user
        return None
    
    def get_all_users(self):
        """모든 사용자 정보 가져오기 (관리자용)"""
        users = self._load_users()
        safe_users = []
        for email, user in users.items():
            safe_user = user.copy()
            if 'password_hash' in safe_user:
                del safe_user['password_hash']
            safe_users.append(safe_user)
        return safe_users
    
    def create_user(self, name, email, password, permission_level='viewer'):
        """새 사용자 생성"""
        users = self._load_users()
        
        if email in users:
            return False, "이미 존재하는 사용자입니다."
        
        users[email] = {
            'name': name,
            'email': email,
            'password_hash': self._hash_password(password),
            'permission_level': permission_level,
            'is_active': True,
            'created_at': datetime.now().isoformat(),
            'last_login': None
        }
        
        self._save_users(users)
        return True, "사용자가 생성되었습니다."
    
    def update_user_permission(self, email, permission_level):
        """사용자 권한 업데이트"""
        users = self._load_users()
        
        if email not in users:
            return False, "사용자를 찾을 수 없습니다."
        
        users[email]['permission_level'] = permission_level
        self._save_users(users)
        return True, "권한이 업데이트되었습니다."
    
    def toggle_user_status(self, email, is_active):
        """사용자 활성화/비활성화"""
        users = self._load_users()
        
        if email not in users:
            return False, "사용자를 찾을 수 없습니다."
        
        users[email]['is_active'] = is_active
        self._save_users(users)
        status = "활성화" if is_active else "비활성화"
        return True, f"사용자가 {status}되었습니다."
    
    def delete_user(self, email):
        """사용자 삭제"""
        users = self._load_users()
        
        if email not in users:
            return False, "사용자를 찾을 수 없습니다."
        
        del users[email]
        self._save_users(users)
        return True, "사용자가 삭제되었습니다."

# 전역 사용자 매니저 인스턴스
user_manager = UserManager()

def _validate_session(require_admin=False):
    """세션 검증 공통 로직"""
    if 'user' not in session:
        if request.is_json:
            return jsonify({'error': '로그인이 필요합니다.', 'redirect': '/login'}), 401
        return redirect(url_for('login_page', message='access_denied'))
    
    # 세션 유효성 검사
    user = session['user']
    if not user.get('email') or not user.get('name'):
        session.clear()
        if request.is_json:
            return jsonify({'error': '세션이 만료되었습니다.', 'redirect': '/login'}), 401
        return redirect(url_for('login_page', message='session_expired'))
    
    # 로그인 시간 기반 세션 검증
    login_time_str = session.get('login_time')
    if login_time_str:
        try:
            login_time = datetime.fromisoformat(login_time_str)
            if datetime.now() - login_time > timedelta(hours=8):
                session.clear()
                if request.is_json:
                    return jsonify({'error': '세션이 만료되었습니다.', 'redirect': '/login'}), 401
                return redirect(url_for('login_page', message='session_expired'))
        except:
            session.clear()
            if request.is_json:
                return jsonify({'error': '세션이 만료되었습니다.', 'redirect': '/login'}), 401
            return redirect(url_for('login_page', message='session_expired'))
    
    # 관리자 권한 검사
    if require_admin and session['user'].get('permission_level') != 'admin':
        if request.is_json:
            return jsonify({'error': '관리자 권한이 필요합니다.'}), 403
        return redirect(url_for('project_list'))
    
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

def _load_user_permissions():
    """사용자 권한 설정 로드"""
    try:
        permissions_file = os.path.join(os.path.dirname(__file__), 'user_permissions.json')
        with open(permissions_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {"admin_users": [], "editor_users": []}

def get_user_role():
    """사용자 역할 결정 (설정 파일 기반)"""
    if 'user' not in session:
        return 'viewer'
    
    user = session['user']
    permission = user.get('permission_level', 'viewer')
    name = user.get('name', '')
    
    # 설정 파일에서 권한 목록 로드
    permissions = _load_user_permissions()
    admin_users = permissions.get('admin_users', [])
    editor_users = permissions.get('editor_users', [])
    
    # DB 권한을 우선으로 하되, 설정 파일의 목록도 체크
    if permission == 'admin' or name in admin_users:
        return 'Admin'
    elif permission == 'editor' or name in editor_users:
        return 'Editor'
    else:
        return 'Viewer'