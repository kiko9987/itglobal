"""
보안 유틸리티 - CSRF, XSS, 입력값 검증
"""

import secrets
import hashlib
import re
import html
from datetime import datetime, timedelta
from typing import Optional, Dict, Any, List
from flask import session, request, jsonify
from functools import wraps
import logging

logger = logging.getLogger(__name__)

class CSRFProtection:
    """CSRF 보호 시스템"""
    
    def __init__(self, secret_key: str):
        self.secret_key = secret_key
        self.token_lifetime = 3600  # 1시간
    
    def generate_csrf_token(self) -> str:
        """CSRF 토큰 생성"""
        timestamp = str(int(datetime.now().timestamp()))
        random_bytes = secrets.token_urlsafe(32)
        
        # 타임스탬프와 랜덤 바이트를 결합하여 서명
        message = f"{timestamp}:{random_bytes}"
        import hmac
        signature = hmac.new(
            self.secret_key.encode(), 
            message.encode(), 
            hashlib.sha256
        ).hexdigest()
        
        token = f"{message}:{signature}"
        session['csrf_token'] = token
        session['csrf_generated'] = datetime.now().isoformat()
        
        return token
    
    def validate_csrf_token(self, token: str) -> bool:
        """CSRF 토큰 검증"""
        try:
            if not token:
                return False
            
            # 세션의 토큰과 비교
            session_token = session.get('csrf_token')
            if not session_token or session_token != token:
                return False
            
            # 토큰 파싱
            parts = token.split(':')
            if len(parts) != 3:
                return False
            
            timestamp_str, random_bytes, signature = parts
            
            # 타임스탬프 검증
            try:
                timestamp = int(timestamp_str)
                current_time = int(datetime.now().timestamp())
                
                if current_time - timestamp > self.token_lifetime:
                    return False
            except ValueError:
                return False
            
            # 서명 검증
            message = f"{timestamp_str}:{random_bytes}"
            import hmac
            expected_signature = hmac.new(
                self.secret_key.encode(),
                message.encode(),
                hashlib.sha256
            ).hexdigest()
            
            return secrets.compare_digest(signature, expected_signature)
            
        except Exception as e:
            logger.error(f"CSRF 토큰 검증 오류: {str(e)}")
            return False

class InputValidator:
    """입력값 검증 및 새니타이징"""
    
    @staticmethod
    def sanitize_html(input_str: str) -> str:
        """HTML 태그 이스케이프"""
        if not input_str:
            return ""
        return html.escape(str(input_str))
    
    @staticmethod
    def validate_email(email: str) -> bool:
        """이메일 형식 검증"""
        if not email:
            return False
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    @staticmethod
    def validate_project_code(code: str) -> bool:
        """프로젝트 코드 형식 검증"""
        if not code:
            return False
        pattern = r'^[A-Z]\d{4}-[A-Z]{2,4}$'
        return re.match(pattern, code) is not None
    
    @staticmethod
    def validate_phone(phone: str) -> bool:
        """전화번호 형식 검증"""
        if not phone:
            return False
        # 한국 전화번호 패턴
        pattern = r'^(01[016789]|02|0[3-9]\d)-?\d{3,4}-?\d{4}$'
        return re.match(pattern, phone.replace('-', '')) is not None
    
    @staticmethod
    def validate_amount(amount: str) -> bool:
        """금액 형식 검증"""
        if not amount:
            return False
        try:
            # 쉼표 제거 후 숫자 확인
            clean_amount = str(amount).replace(',', '').replace('₩', '').strip()
            float(clean_amount)
            return True
        except ValueError:
            return False
    
    @staticmethod
    def clean_string(input_str: str, max_length: int = 1000) -> str:
        """문자열 정리 및 길이 제한"""
        if not input_str:
            return ""
        
        # HTML 이스케이프
        clean_str = html.escape(str(input_str))
        
        # 길이 제한
        if len(clean_str) > max_length:
            clean_str = clean_str[:max_length] + "..."
        
        return clean_str.strip()

class SecurityHeaders:
    """보안 헤더 관리"""
    
    @staticmethod
    def get_security_headers() -> Dict[str, str]:
        """보안 헤더 반환"""
        return {
            'X-Content-Type-Options': 'nosniff',
            'X-Frame-Options': 'DENY',
            'X-XSS-Protection': '1; mode=block',
            'Strict-Transport-Security': 'max-age=31536000; includeSubDomains',
            'Content-Security-Policy': (
                "default-src 'self'; "
                "script-src 'self' 'unsafe-inline' https://code.jquery.com https://cdn.socket.io https://cdn.jsdelivr.net https://cdnjs.cloudflare.com https://cdn.datatables.net; "
                "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com https://cdn.jsdelivr.net https://cdnjs.cloudflare.com https://cdn.datatables.net; "
                "font-src 'self' https://fonts.gstatic.com https://cdn.jsdelivr.net https://cdnjs.cloudflare.com; "
                "img-src 'self' data: https:; "
                "connect-src 'self' wss: ws: https://code.jquery.com https://cdn.jsdelivr.net https://cdnjs.cloudflare.com https://cdn.datatables.net https://fonts.googleapis.com https://fonts.gstatic.com;"
            ),
            'Referrer-Policy': 'strict-origin-when-cross-origin',
            'Permissions-Policy': (
                'camera=(), microphone=(), geolocation=(), '
                'payment=(), usb=(), magnetometer=(), gyroscope=()'
            )
        }

# 전역 인스턴스
_csrf_protection = None
_input_validator = InputValidator()
_security_headers = SecurityHeaders()

def init_csrf_protection(secret_key: str):
    """CSRF 보호 초기화"""
    global _csrf_protection
    _csrf_protection = CSRFProtection(secret_key)

def csrf_protect(f):
    """CSRF 보호 데코레이터"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if request.method in ['POST', 'PUT', 'DELETE', 'PATCH']:
            csrf_token = request.headers.get('X-CSRF-Token') or request.form.get('csrf_token')
            
            if not _csrf_protection or not _csrf_protection.validate_csrf_token(csrf_token):
                logger.warning(f"CSRF 토큰 검증 실패 - IP: {request.remote_addr}")
                return jsonify({
                    'error': 'CSRF token validation failed',
                    'code': 'INVALID_CSRF_TOKEN'
                }), 403
        
        return f(*args, **kwargs)
    
    return decorated_function

def validate_input(**validators):
    """입력값 검증 데코레이터"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            data = request.get_json() or {}
            errors = []
            
            for field, validator_type in validators.items():
                value = data.get(field)
                
                if validator_type == 'email' and value:
                    if not _input_validator.validate_email(value):
                        errors.append(f"Invalid email format: {field}")
                
                elif validator_type == 'project_code' and value:
                    if not _input_validator.validate_project_code(value):
                        errors.append(f"Invalid project code format: {field}")
                
                elif validator_type == 'phone' and value:
                    if not _input_validator.validate_phone(value):
                        errors.append(f"Invalid phone format: {field}")
                
                elif validator_type == 'amount' and value:
                    if not _input_validator.validate_amount(value):
                        errors.append(f"Invalid amount format: {field}")
            
            if errors:
                return jsonify({
                    'error': 'Validation failed',
                    'details': errors
                }), 400
            
            # 입력값 새니타이징
            for key, value in data.items():
                if isinstance(value, str):
                    data[key] = _input_validator.clean_string(value)
            
            request._validated_data = data
            return f(*args, **kwargs)
        
        return decorated_function
    return decorator

def get_csrf_token() -> str:
    """CSRF 토큰 가져오기"""
    if not _csrf_protection:
        raise RuntimeError("CSRF protection not initialized")
    return _csrf_protection.generate_csrf_token()

def add_security_headers(response):
    """응답에 보안 헤더 추가"""
    headers = _security_headers.get_security_headers()
    for key, value in headers.items():
        response.headers[key] = value
    return response

# Flask 앱 초기화용 함수
def init_security(app):
    """Flask 앱에 보안 설정 적용"""
    # CSRF 보호 초기화
    init_csrf_protection(app.config['SECRET_KEY'])
    
    # 모든 응답에 보안 헤더 추가
    @app.after_request
    def add_security_headers_to_response(response):
        return add_security_headers(response)