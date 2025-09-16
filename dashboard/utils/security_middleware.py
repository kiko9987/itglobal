"""
보안 미들웨어
- CSRF 보호 강화
- 입력값 검증 및 XSS 방지
- Rate Limiting
- API 보안 강화
"""

import html
import re
import time
import json
import hashlib
from collections import defaultdict, deque
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Set
from functools import wraps
from flask import request, abort, jsonify, session
import logging

logger = logging.getLogger(__name__)

class RateLimiter:
    """속도 제한기"""

    def __init__(self):
        self.requests = defaultdict(deque)
        self.blocked_ips = defaultdict(float)  # IP: 차단 해제 시간

    def is_allowed(self, identifier: str, limit: int = 100, window: int = 3600) -> bool:
        """요청 허용 여부 확인"""
        current_time = time.time()

        # 차단된 IP 확인
        if identifier in self.blocked_ips:
            if current_time < self.blocked_ips[identifier]:
                return False
            else:
                del self.blocked_ips[identifier]

        # 윈도우 시간 내 요청만 유지
        request_times = self.requests[identifier]
        while request_times and request_times[0] < current_time - window:
            request_times.popleft()

        # 제한 확인
        if len(request_times) >= limit:
            # 과도한 요청 시 IP 차단 (1시간)
            self.blocked_ips[identifier] = current_time + 3600
            logger.warning(f"Rate limit exceeded, blocking IP: {identifier}")
            return False

        # 요청 기록
        request_times.append(current_time)
        return True

    def get_remaining_requests(self, identifier: str, limit: int = 100, window: int = 3600) -> int:
        """남은 요청 수 반환"""
        current_time = time.time()
        request_times = self.requests[identifier]

        # 윈도우 시간 내 요청만 카운트
        valid_requests = sum(1 for req_time in request_times if req_time > current_time - window)
        return max(0, limit - valid_requests)


class InputValidator:
    """입력값 검증기"""

    # 위험한 패턴들
    XSS_PATTERNS = [
        r'<script[^>]*>.*?</script>',
        r'javascript:',
        r'on\w+\s*=',
        r'<iframe[^>]*>.*?</iframe>',
        r'<object[^>]*>.*?</object>',
        r'<embed[^>]*>.*?</embed>',
        r'<form[^>]*>.*?</form>',
    ]

    SQL_INJECTION_PATTERNS = [
        r'(\b(SELECT|INSERT|UPDATE|DELETE|DROP|CREATE|ALTER|EXEC|UNION)\b)',
        r'(\-\-|\#|\/\*|\*\/)',
        r'(\b(OR|AND)\b\s+\d+\s*=\s*\d+)',
        r'(\bUNION\b.*\bSELECT\b)',
    ]

    COMMAND_INJECTION_PATTERNS = [
        r'[;&|`$]',
        r'\b(cat|ls|pwd|whoami|id|ps|netstat|ifconfig)\b',
        r'(\.\.\/|\.\.\\)',
    ]

    @classmethod
    def sanitize_string(cls, value: str) -> str:
        """문자열 정화"""
        if not isinstance(value, str):
            return str(value)

        # HTML 태그 이스케이프
        sanitized = html.escape(value)

        # 연속 공백 정규화
        sanitized = re.sub(r'\s+', ' ', sanitized).strip()

        # 길이 제한 (10KB)
        if len(sanitized) > 10240:
            sanitized = sanitized[:10240]

        return sanitized

    @classmethod
    def is_safe_string(cls, value: str) -> bool:
        """문자열 안전성 검사"""
        if not isinstance(value, str):
            return False

        # XSS 패턴 검사
        for pattern in cls.XSS_PATTERNS:
            if re.search(pattern, value, re.IGNORECASE):
                logger.warning(f"XSS pattern detected: {pattern}")
                return False

        # SQL Injection 패턴 검사
        for pattern in cls.SQL_INJECTION_PATTERNS:
            if re.search(pattern, value, re.IGNORECASE):
                logger.warning(f"SQL injection pattern detected: {pattern}")
                return False

        # Command Injection 패턴 검사
        for pattern in cls.COMMAND_INJECTION_PATTERNS:
            if re.search(pattern, value, re.IGNORECASE):
                logger.warning(f"Command injection pattern detected: {pattern}")
                return False

        return True

    @classmethod
    def validate_project_code(cls, project_code: str) -> bool:
        """프로젝트 코드 형식 검증"""
        if not isinstance(project_code, str):
            return False

        # 프로젝트 코드 패턴: 문자1개 + 숫자4개 + '-' + 문자1-3개
        pattern = r'^[A-Z]\d{4}-[A-Z]{1,3}$'
        return bool(re.match(pattern, project_code))

    @classmethod
    def validate_email(cls, email: str) -> bool:
        """이메일 형식 검증"""
        if not isinstance(email, str):
            return False

        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email)) and len(email) <= 254

    @classmethod
    def validate_phone(cls, phone: str) -> bool:
        """전화번호 형식 검증"""
        if not isinstance(phone, str):
            return False

        # 한국 전화번호 패턴
        pattern = r'^(\+82|0)?(01[016789]|02|0[3-9]\d)-?\d{3,4}-?\d{4}$'
        return bool(re.match(pattern, phone.replace('-', '').replace(' ', '')))

    @classmethod
    def validate_amount(cls, amount: Any) -> bool:
        """금액 형식 검증"""
        try:
            if isinstance(amount, str):
                # 콤마 제거 후 숫자 변환
                amount = amount.replace(',', '').replace('원', '').strip()
                amount = float(amount)

            if isinstance(amount, (int, float)):
                return 0 <= amount <= 999999999999  # 1조 미만

            return False
        except (ValueError, TypeError):
            return False


class CSRFProtection:
    """CSRF 보호"""

    def __init__(self, secret_key: str):
        self.secret_key = secret_key

    def generate_token(self, session_id: str) -> str:
        """CSRF 토큰 생성"""
        timestamp = str(int(time.time()))
        data = f"{session_id}:{timestamp}:{self.secret_key}"
        token = hashlib.sha256(data.encode()).hexdigest()
        return f"{timestamp}.{token}"

    def validate_token(self, token: str, session_id: str, max_age: int = 3600) -> bool:
        """CSRF 토큰 검증"""
        try:
            if not token or '.' not in token:
                return False

            timestamp_str, token_hash = token.split('.', 1)
            timestamp = int(timestamp_str)

            # 토큰 만료 확인
            if time.time() - timestamp > max_age:
                return False

            # 토큰 해시 검증
            expected_data = f"{session_id}:{timestamp_str}:{self.secret_key}"
            expected_hash = hashlib.sha256(expected_data.encode()).hexdigest()

            return token_hash == expected_hash

        except (ValueError, TypeError):
            return False


class SecurityMiddleware:
    """보안 미들웨어"""

    def __init__(self, app, secret_key: str):
        self.app = app
        self.rate_limiter = RateLimiter()
        self.csrf_protection = CSRFProtection(secret_key)
        self.validator = InputValidator()

        # 보안 헤더
        self.security_headers = {
            'X-Content-Type-Options': 'nosniff',
            'X-Frame-Options': 'DENY',
            'X-XSS-Protection': '1; mode=block',
            'Referrer-Policy': 'strict-origin-when-cross-origin',
            'Content-Security-Policy': "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline';"
        }

        # 보안 로그
        self.security_events = deque(maxlen=1000)

    def process_request(self):
        """요청 처리"""
        # 1. Rate Limiting
        if not self._check_rate_limit():
            self._log_security_event('RATE_LIMIT_EXCEEDED', request.remote_addr)
            abort(429, description="Too Many Requests")

        # 2. CSRF 검사 (상태 변경 요청)
        if request.method in ['POST', 'PUT', 'DELETE', 'PATCH']:
            if not self._check_csrf_token():
                self._log_security_event('CSRF_TOKEN_INVALID', request.remote_addr)
                abort(403, description="CSRF token validation failed")

        # 3. 입력값 검증
        if request.is_json:
            try:
                data = request.get_json()
                if not self._validate_json_input(data):
                    self._log_security_event('INVALID_INPUT', request.remote_addr)
                    abort(400, description="Invalid input detected")
            except Exception as e:
                logger.error(f"JSON 파싱 오류: {e}")
                abort(400, description="Invalid JSON")

    def process_response(self, response):
        """응답 처리"""
        # 보안 헤더 추가
        for header, value in self.security_headers.items():
            response.headers[header] = value

        # CSRF 토큰 추가 (HTML 응답)
        if response.content_type and 'text/html' in response.content_type:
            if 'user' in session:
                csrf_token = self.csrf_protection.generate_token(session.get('session_id', ''))
                # 응답에 토큰 추가하는 로직 (구현 필요)

        return response

    def _check_rate_limit(self) -> bool:
        """Rate Limiting 확인"""
        identifier = self._get_client_identifier()

        # API 엔드포인트별 다른 제한
        if request.path.startswith('/api/'):
            limit = 1000  # API는 더 관대한 제한
            window = 3600
        else:
            limit = 100   # 일반 페이지는 엄격한 제한
            window = 3600

        return self.rate_limiter.is_allowed(identifier, limit, window)

    def _check_csrf_token(self) -> bool:
        """CSRF 토큰 확인"""
        # API 요청은 별도 검증 (JWT 토큰 등)
        if request.path.startswith('/api/'):
            return self._validate_api_auth()

        # 폼 요청은 CSRF 토큰 검증
        csrf_token = request.headers.get('X-CSRF-Token') or request.form.get('csrf_token')
        if not csrf_token:
            return False

        session_id = session.get('session_id', '')
        return self.csrf_protection.validate_token(csrf_token, session_id)

    def _validate_api_auth(self) -> bool:
        """API 인증 검증"""
        # Authorization 헤더 확인
        auth_header = request.headers.get('Authorization', '')
        if auth_header.startswith('Bearer '):
            token = auth_header[7:]
            # JWT 토큰 검증 로직 (구현 필요)
            return self._validate_jwt_token(token)

        # 세션 기반 인증 확인
        if 'user' in session:
            return True

        return False

    def _validate_jwt_token(self, token: str) -> bool:
        """JWT 토큰 검증 (구현 필요)"""
        # JWT 라이브러리를 사용한 토큰 검증
        try:
            # import jwt
            # payload = jwt.decode(token, self.secret_key, algorithms=['HS256'])
            # return True
            return True  # 임시로 True 반환
        except Exception:
            return False

    def _validate_json_input(self, data: Any) -> bool:
        """JSON 입력값 검증"""
        if isinstance(data, dict):
            for key, value in data.items():
                if not self._validate_field(key, value):
                    return False
        elif isinstance(data, list):
            for item in data:
                if not self._validate_json_input(item):
                    return False
        elif isinstance(data, str):
            return self.validator.is_safe_string(data)

        return True

    def _validate_field(self, field_name: str, value: Any) -> bool:
        """필드별 검증"""
        if value is None:
            return True

        if isinstance(value, str):
            # 모든 문자열은 기본 안전성 검사
            if not self.validator.is_safe_string(value):
                return False

            # 필드별 특수 검증
            if 'email' in field_name.lower():
                return self.validator.validate_email(value)
            elif 'phone' in field_name.lower() or '연락처' in field_name:
                return self.validator.validate_phone(value)
            elif '프로젝트' in field_name and '코드' in field_name:
                return self.validator.validate_project_code(value)
            elif any(keyword in field_name for keyword in ['금액', '수금', '비용']):
                return self.validator.validate_amount(value)

            # 문자열 길이 제한
            return len(value) <= 1000

        elif isinstance(value, (int, float)):
            # 숫자 범위 확인
            return -999999999999 <= value <= 999999999999

        elif isinstance(value, (dict, list)):
            return self._validate_json_input(value)

        return True

    def _get_client_identifier(self) -> str:
        """클라이언트 식별자 생성"""
        # IP 주소 기반 (프록시 환경에서는 X-Forwarded-For 고려)
        ip = request.headers.get('X-Forwarded-For', request.remote_addr)
        if ip:
            ip = ip.split(',')[0].strip()

        # 사용자 세션이 있으면 세션 ID 사용
        if 'user' in session:
            return f"user_{session['user']['id']}"

        return f"ip_{ip}"

    def _log_security_event(self, event_type: str, client: str, details: str = None):
        """보안 이벤트 로깅"""
        event = {
            'timestamp': datetime.now().isoformat(),
            'type': event_type,
            'client': client,
            'path': request.path,
            'method': request.method,
            'user_agent': request.headers.get('User-Agent', ''),
            'details': details
        }

        self.security_events.append(event)
        logger.warning(f"Security event: {event_type} from {client} at {request.path}")

    def get_security_stats(self) -> Dict[str, Any]:
        """보안 통계 반환"""
        now = datetime.now()
        hour_ago = now - timedelta(hours=1)

        recent_events = [
            event for event in self.security_events
            if datetime.fromisoformat(event['timestamp']) > hour_ago
        ]

        event_counts = defaultdict(int)
        for event in recent_events:
            event_counts[event['type']] += 1

        return {
            'total_events_last_hour': len(recent_events),
            'event_types': dict(event_counts),
            'blocked_ips': len(self.rate_limiter.blocked_ips),
            'active_rate_limits': len(self.rate_limiter.requests)
        }


# 데코레이터 함수들
def rate_limit(limit: int = 100, window: int = 3600):
    """Rate Limiting 데코레이터"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            # 미들웨어가 있으면 미들웨어에서 처리
            if hasattr(request, 'security_middleware'):
                identifier = request.security_middleware._get_client_identifier()
                if not request.security_middleware.rate_limiter.is_allowed(identifier, limit, window):
                    abort(429, description="Too Many Requests")

            return f(*args, **kwargs)
        return decorated_function
    return decorator


def validate_input(*field_rules):
    """입력값 검증 데코레이터"""
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if request.is_json:
                data = request.get_json()
                validator = InputValidator()

                for field_name, rule in field_rules:
                    if field_name in data:
                        value = data[field_name]

                        if rule == 'required' and not value:
                            abort(400, description=f"Required field missing: {field_name}")
                        elif rule == 'email' and not validator.validate_email(value):
                            abort(400, description=f"Invalid email format: {field_name}")
                        elif rule == 'project_code' and not validator.validate_project_code(value):
                            abort(400, description=f"Invalid project code format: {field_name}")
                        elif rule == 'safe_string' and not validator.is_safe_string(value):
                            abort(400, description=f"Unsafe string detected: {field_name}")

            return f(*args, **kwargs)
        return decorated_function
    return decorator


# 전역 미들웨어 인스턴스 (초기화 필요)
security_middleware = None


def init_security_middleware(app, secret_key: str):
    """보안 미들웨어 초기화"""
    global security_middleware
    security_middleware = SecurityMiddleware(app, secret_key)

    @app.before_request
    def before_request():
        if security_middleware:
            security_middleware.process_request()

    @app.after_request
    def after_request(response):
        if security_middleware:
            return security_middleware.process_response(response)
        return response


def get_security_middleware() -> Optional[SecurityMiddleware]:
    """보안 미들웨어 인스턴스 반환"""
    return security_middleware