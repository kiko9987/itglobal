"""
JWT 토큰 기반 인증 시스템
기존 세션 기반 인증을 보완하는 토큰 시스템
"""

import jwt
import secrets
from datetime import datetime, timedelta, timezone
from typing import Optional, Dict, Any
from flask import request, jsonify, current_app
from functools import wraps
import logging

logger = logging.getLogger(__name__)

class JWTManager:
    """JWT 토큰 관리자"""
    
    def __init__(self, secret_key: str, algorithm: str = 'HS256'):
        self.secret_key = secret_key
        self.algorithm = algorithm
        self.access_token_expire = 60 * 60  # 1시간
        self.refresh_token_expire = 24 * 60 * 60 * 7  # 7일
        
    def generate_tokens(self, user_data: Dict[str, Any]) -> Dict[str, str]:
        """액세스 토큰과 리프레시 토큰 생성"""
        now = datetime.now(timezone.utc)
        
        # 액세스 토큰 페이로드
        access_payload = {
            'user_id': user_data.get('email'),
            'name': user_data.get('name'),
            'permission_level': user_data.get('permission_level', 'viewer'),
            'iat': now,
            'exp': now + timedelta(seconds=self.access_token_expire),
            'type': 'access'
        }
        
        # 리프레시 토큰 페이로드
        refresh_payload = {
            'user_id': user_data.get('email'),
            'iat': now,
            'exp': now + timedelta(seconds=self.refresh_token_expire),
            'type': 'refresh',
            'jti': secrets.token_urlsafe(32)  # 토큰 고유 ID
        }
        
        # 토큰 생성
        access_token = jwt.encode(access_payload, self.secret_key, algorithm=self.algorithm)
        refresh_token = jwt.encode(refresh_payload, self.secret_key, algorithm=self.algorithm)
        
        return {
            'access_token': access_token,
            'refresh_token': refresh_token,
            'token_type': 'Bearer',
            'expires_in': self.access_token_expire
        }
    
    def verify_token(self, token: str, token_type: str = 'access') -> Optional[Dict[str, Any]]:
        """토큰 검증"""
        try:
            payload = jwt.decode(
                token, 
                self.secret_key, 
                algorithms=[self.algorithm],
                options={'verify_exp': True}
            )
            
            # 토큰 타입 확인
            if payload.get('type') != token_type:
                logger.warning(f"잘못된 토큰 타입: {payload.get('type')} (expected: {token_type})")
                return None
            
            return payload
            
        except jwt.ExpiredSignatureError:
            logger.warning("만료된 토큰")
            return None
        except jwt.InvalidTokenError as e:
            logger.warning(f"유효하지 않은 토큰: {str(e)}")
            return None
        except Exception as e:
            logger.error(f"토큰 검증 오류: {str(e)}")
            return None
    
    def refresh_access_token(self, refresh_token: str) -> Optional[Dict[str, str]]:
        """리프레시 토큰으로 새 액세스 토큰 생성"""
        payload = self.verify_token(refresh_token, 'refresh')
        if not payload:
            return None
        
        # 사용자 정보로 새 액세스 토큰 생성
        user_data = {
            'email': payload.get('user_id'),
            'name': payload.get('name', ''),
            'permission_level': payload.get('permission_level', 'viewer')
        }
        
        # 새 액세스 토큰만 생성
        now = datetime.now(timezone.utc)
        access_payload = {
            'user_id': user_data['email'],
            'name': user_data['name'],
            'permission_level': user_data['permission_level'],
            'iat': now,
            'exp': now + timedelta(seconds=self.access_token_expire),
            'type': 'access'
        }
        
        access_token = jwt.encode(access_payload, self.secret_key, algorithm=self.algorithm)
        
        return {
            'access_token': access_token,
            'token_type': 'Bearer',
            'expires_in': self.access_token_expire
        }

# 전역 JWT 매니저 인스턴스
_jwt_manager = None

def init_jwt_manager(secret_key: str):
    """JWT 매니저 초기화"""
    global _jwt_manager
    _jwt_manager = JWTManager(secret_key)

def jwt_required(optional: bool = False, admin_required: bool = False):
    """JWT 토큰 필수 데코레이터"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            auth_header = request.headers.get('Authorization')
            
            if not auth_header:
                if optional:
                    request.jwt_user = None
                    return f(*args, **kwargs)
                return jsonify({
                    'error': 'Authorization header missing',
                    'code': 'MISSING_AUTH_HEADER'
                }), 401
            
            # Bearer 토큰 파싱
            try:
                token_type, token = auth_header.split(' ', 1)
                if token_type.lower() != 'bearer':
                    raise ValueError("Invalid token type")
            except ValueError:
                return jsonify({
                    'error': 'Invalid authorization header format',
                    'code': 'INVALID_AUTH_HEADER'
                }), 401
            
            # 토큰 검증
            if not _jwt_manager:
                logger.error("JWT manager not initialized")
                return jsonify({
                    'error': 'Authentication system not available',
                    'code': 'AUTH_SYSTEM_ERROR'
                }), 500
            
            payload = _jwt_manager.verify_token(token)
            if not payload:
                return jsonify({
                    'error': 'Invalid or expired token',
                    'code': 'INVALID_TOKEN'
                }), 401
            
            # 관리자 권한 확인
            if admin_required and payload.get('permission_level') != 'admin':
                return jsonify({
                    'error': 'Admin permission required',
                    'code': 'INSUFFICIENT_PERMISSIONS'
                }), 403
            
            # 요청에 사용자 정보 추가
            request.jwt_user = {
                'email': payload.get('user_id'),
                'name': payload.get('name'),
                'permission_level': payload.get('permission_level'),
                'token_exp': payload.get('exp')
            }
            
            return f(*args, **kwargs)
        
        return decorated_function
    return decorator

def get_current_jwt_user() -> Optional[Dict[str, Any]]:
    """현재 JWT 사용자 정보 반환"""
    return getattr(request, 'jwt_user', None)

def create_jwt_tokens(user_data: Dict[str, Any]) -> Dict[str, str]:
    """JWT 토큰 생성"""
    if not _jwt_manager:
        raise RuntimeError("JWT manager not initialized")
    return _jwt_manager.generate_tokens(user_data)

def refresh_jwt_token(refresh_token: str) -> Optional[Dict[str, str]]:
    """JWT 토큰 새로고침"""
    if not _jwt_manager:
        raise RuntimeError("JWT manager not initialized")
    return _jwt_manager.refresh_access_token(refresh_token)

def get_token_info(token: str) -> Optional[Dict[str, Any]]:
    """토큰 정보 조회 (검증 없이)"""
    try:
        # 검증 없이 디코딩 (주의: 검증된 용도로만 사용)
        payload = jwt.decode(token, options={"verify_signature": False})
        return payload
    except Exception as e:
        logger.error(f"토큰 정보 조회 오류: {str(e)}")
        return None

# Flask 라우트용 헬퍼 함수들
def create_token_response(user_data: Dict[str, Any]) -> Dict[str, Any]:
    """토큰 응답 생성"""
    tokens = create_jwt_tokens(user_data)
    return {
        'success': True,
        'message': 'Authentication successful',
        'data': {
            'user': {
                'email': user_data.get('email'),
                'name': user_data.get('name'),
                'permission_level': user_data.get('permission_level')
            },
            'tokens': tokens
        }
    }

def create_refresh_response(refresh_token: str) -> Dict[str, Any]:
    """토큰 새로고침 응답 생성"""
    new_tokens = refresh_jwt_token(refresh_token)
    if not new_tokens:
        return {
            'success': False,
            'error': 'Invalid refresh token',
            'code': 'INVALID_REFRESH_TOKEN'
        }
    
    return {
        'success': True,
        'message': 'Token refreshed successfully',
        'data': {
            'tokens': new_tokens
        }
    }