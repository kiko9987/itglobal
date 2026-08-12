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
import os
import secrets
from collections import defaultdict, deque
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Set
from functools import wraps
from flask import request, abort, jsonify, session, g
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
        r'[;`$]',  # &, | 제거 (회사명 등에서 사용)
        r'\b(cat|ls|pwd|whoami|id|ps|netstat|ifconfig)\b',
        r'(\.\.\/|\.\.\\)',
    ]

    # 비즈니스 필드에서 허용할 특수문자
    BUSINESS_SAFE_CHARS = ['&', '|', '-', '·', '/', '(', ')', '+', '.', ',', ':', ' ']

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
        """전화번호 형식 검증 (유연한 패턴)"""
        if not isinstance(phone, str):
            return False

        # 공백, 하이픈 제거
        clean_phone = phone.replace('-', '').replace(' ', '').replace('(', '').replace(')', '')

        # 다양한 전화번호 패턴 허용
        # 1) 국제번호: +82-10-1234-5678
        # 2) 휴대폰: 010-1234-5678
        # 3) 지역번호: 02-123-4567, 031-123-4567
        # 4) 대표번호: 1588-1234, 1577-1234
        # 5) 내선: 123-4567
        patterns = [
            r'^\+82\d{9,11}$',  # 국제번호
            r'^0\d{8,10}$',  # 일반 전화번호 (서울 02 포함, 9~11자리)
            r'^1[5-9]\d{2}\d{4}$',  # 대표번호
            r'^\d{7,8}$',  # 내선번호
        ]

        return any(re.match(pattern, clean_phone) for pattern in patterns)

    @classmethod
    def validate_amount(cls, amount: Any) -> bool:
        """금액 형식 검증 (9999억까지)"""
        try:
            if isinstance(amount, str):
                # 콤마 제거 후 숫자 변환
                amount = amount.replace(',', '').replace('원', '').strip()

                # 빈 문자열은 허용 (삭제/초기화)
                if amount == '':
                    return True

                amount = float(amount)

            if isinstance(amount, (int, float)):
                return 0 <= amount <= 999999999999  # 9999억 (억 단위)

            return False
        except (ValueError, TypeError):
            return False


class CSRFProtection:
    """CSRF 보호 (HMAC 기반 - 더 안전한 방식)"""

    def __init__(self, secret_key: str):
        self.secret_key = secret_key
        self.token_lifetime = 3600  # 1시간

    def generate_token(self, session_id: str) -> str:
        """
        CSRF 토큰 생성 (HMAC 기반)

        Args:
            session_id: 세션 ID

        Returns:
            CSRF 토큰 문자열 (timestamp:random_bytes:signature 형식)
        """
        import secrets
        import hmac

        timestamp = str(int(time.time()))
        random_bytes = secrets.token_urlsafe(32)

        # 타임스탬프, 랜덤 바이트, 세션 ID를 결합하여 HMAC 서명
        message = f"{timestamp}:{random_bytes}:{session_id}"
        signature = hmac.new(
            self.secret_key.encode(),
            message.encode(),
            hashlib.sha256
        ).hexdigest()

        token = f"{timestamp}:{random_bytes}:{signature}"
        return token

    def validate_token(self, token: str, session_id: str, max_age: int = 3600) -> bool:
        """
        CSRF 토큰 검증 (HMAC 기반)

        Args:
            token: 검증할 토큰
            session_id: 세션 ID
            max_age: 토큰 유효 시간 (초)

        Returns:
            검증 성공 여부
        """
        import secrets
        import hmac

        try:
            if not token:
                return False

            # 토큰 파싱
            parts = token.split(':')
            if len(parts) != 3:
                return False

            timestamp_str, random_bytes, signature = parts

            # 타임스탬프 검증
            timestamp = int(timestamp_str)
            current_time = int(time.time())

            if current_time - timestamp > max_age:
                logger.debug(f"CSRF 토큰 만료: {current_time - timestamp}초 경과")
                return False

            # HMAC 서명 검증
            message = f"{timestamp_str}:{random_bytes}:{session_id}"
            expected_signature = hmac.new(
                self.secret_key.encode(),
                message.encode(),
                hashlib.sha256
            ).hexdigest()

            # timing attack 방지를 위한 constant-time 비교
            return secrets.compare_digest(signature, expected_signature)

        except (ValueError, TypeError) as e:
            logger.warning(f"CSRF 토큰 검증 오류: {e}")
            return False


class SecurityMiddleware:
    """보안 미들웨어"""

    # HTML/특수문자가 허용되는 필드 패턴
    # 패턴 기반 매칭으로 유지보수 용이성 향상
    HTML_ALLOWED_FIELD_PATTERNS = [
        '내용',        # '공사 내용', '작업 내용', '추가 내용' 등
        '특이사항',    # '수금 관련 특이사항', '공사 특이사항', '기타 특이사항' 등
        '주소',        # '현장 주소', '배송 주소', '사업장 주소' 등
        '경로',        # '폴더 경로', '파일 경로' 등
        '_메모',       # 모든 메모 필드 (계약금_메모, 중도금_메모, 잔금_메모 등)
        '비고',        # '비고', '특이 비고' 등
        '설명',        # '상세 설명', '작업 설명' 등
    ]

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
            'Content-Security-Policy': (
                "default-src 'self'; "
                "script-src 'self' 'unsafe-inline' http://localhost:5173 https://cdn.jsdelivr.net https://cdn.datatables.net https://code.jquery.com https://cdnjs.cloudflare.com; "
                "style-src 'self' 'unsafe-inline' http://localhost:5173 https://cdn.jsdelivr.net https://cdn.datatables.net https://cdnjs.cloudflare.com https://use.fontawesome.com; "
                "img-src 'self' data: https:; "
                "font-src 'self' https://cdnjs.cloudflare.com https://use.fontawesome.com data:; "
                "connect-src 'self' http://localhost:5173 ws://localhost:5173 ws://localhost:5000 https:; "
                "worker-src 'self' blob:;"
            )
        }

        # 보안 로그
        self.security_events = deque(maxlen=1000)

    def process_request(self):
        """요청 처리"""
        # 0. 정적 리소스·SocketIO polling·health check는 rate limit 대상 아님
        # (한 페이지 로드에 다수 요청 발생 + SocketIO polling 5초 주기 = 시간당 720건)
        _p = request.path
        if (_p.startswith('/static/') or _p.startswith('/favicon')
                or _p.startswith('/socket.io/')  # SocketIO polling 제외
                or _p.startswith('/api/health')  # health check 제외 (모니터링용)
                or _p.endswith('.css') or _p.endswith('.js')
                or _p.endswith('.map') or _p.endswith('.ico')
                or _p.endswith('.woff') or _p.endswith('.woff2')
                or _p.endswith('.png') or _p.endswith('.jpg')
                or _p.endswith('.svg')):
            return None

        # 1. Rate Limiting
        if not self._check_rate_limit():
            self._log_security_event('RATE_LIMIT_EXCEEDED', request.remote_addr)
            return jsonify({'error': 'Too Many Requests'}), 429

        # 1.5. Slack webhook 우회 — 슬랙은 자체 signing secret으로 서명 검증함
        if request.path.startswith('/slack/'):
            return None  # CSRF/세션 검증 skip, slack_bolt가 직접 서명 검증

        # 1.6. 채널톡 webhook 우회 — 채널톡 자체 X-Signature 서명 검증 사용
        if request.path.startswith('/channeltalk/'):
            return None

        # 1.7. SMS 인입 webhook 우회 — 폰 SMS 포워딩(로그인 세션 없음).
        #      기기별 토큰(SMS_INBOUND_TOKENS)으로 sms_inbound 라우트가 자체 인증.
        if request.path.startswith('/sms/'):
            return None

        # 2. CSRF/인증 검사 (상태 변경 요청)
        if request.method in ['POST', 'PUT', 'DELETE', 'PATCH']:
            if request.path.startswith('/api/') or request.path.startswith('/admin/api/'):
                # API 엔드포인트: 세션 기반 인증 검증
                if not self._validate_api_auth():
                    self._log_security_event('API_AUTH_FAILED', request.remote_addr)
                    return jsonify({'error': 'Authentication required'}), 401
            else:
                # 폼 요청: CSRF 토큰 검증
                if not self._check_csrf_token():
                    self._log_security_event('CSRF_TOKEN_INVALID', request.remote_addr)
                    return jsonify({'error': 'CSRF token validation failed'}), 403

        # 3. 입력값 검증 (POST/PUT/PATCH/DELETE 요청만)
        if request.method in ['POST', 'PUT', 'PATCH', 'DELETE'] and request.is_json:
            try:
                # 에러 정보 초기화 (flask.g를 사용하여 thread-safe하게)
                g.validation_error = None
                data = request.get_json()
                if not self._validate_json_input(data):
                    self._log_security_event('INVALID_INPUT', request.remote_addr)

                    # 구체적인 에러 메시지 생성
                    validation_error = getattr(g, 'validation_error', None)
                    if validation_error:
                        error_response = {
                            'error': validation_error.get('message', 'Invalid input detected'),
                            'field': validation_error.get('field'),
                            'message': validation_error.get('message', 'Invalid input detected')
                        }
                    else:
                        error_response = {'error': 'Invalid input detected'}

                    return jsonify(error_response), 400
            except Exception as e:
                logger.error(f"JSON 파싱 오류: {e}")
                return jsonify({'error': 'Invalid JSON'}), 400

    def process_response(self, response):
        """응답 처리"""
        # 보안 헤더 추가
        for header, value in self.security_headers.items():
            response.headers[header] = value

        # CSRF 토큰 추가 (HTML 응답)
        if response.content_type and 'text/html' in response.content_type:
            if 'user' in session:
                csrf_token = self.csrf_protection.generate_token(session.get('session_id', ''))

                # HTML 응답에 메타 태그 자동 주입
                try:
                    data = response.get_data(as_text=True)

                    # <head> 태그를 찾아서 메타 태그 삽입
                    meta_tag = f'<meta name="csrf-token" content="{csrf_token}">'

                    if '<head>' in data:
                        # <head> 태그 바로 다음에 삽입
                        data = data.replace('<head>', f'<head>\n    {meta_tag}', 1)
                        response.set_data(data)
                        logger.debug(f"CSRF 토큰 메타 태그 삽입 완료: {request.path}")
                    else:
                        # <head> 태그가 없으면 로그만 기록
                        logger.debug(f"<head> 태그 없음, CSRF 메타 태그 미삽입: {request.path}")

                except Exception as e:
                    logger.warning(f"CSRF 메타 태그 삽입 실패: {e}")

        return response

    def _check_rate_limit(self) -> bool:
        """Rate Limiting 확인"""
        identifier = self._get_client_identifier()

        # Slack 이벤트/명령은 Slack signing secret 으로 이미 인증됨.
        # Slack 서버가 retry·events 대량 발송 시 1000/시간 쉽게 초과 →
        # slash command 실패 (2026-07-22 /as 사고). rate limit 예외.
        if request.path.startswith('/slack/'):
            return True

        # SMS 인입 webhook 도 rate limit 예외 — 폰 포워딩은 기기 토큰으로 자체 인증,
        # 입금 몰릴 때 다건 연속 인입이 익명 제한(1000/시간)에 걸려 유실되면 안 됨.
        if request.path.startswith('/sms/'):
            return True

        # 인증된 세션 (로그인 성공한 매니저) 은 rate limit 예외 (2026-07-22):
        # 매니저 브라우저 다중 탭 + prefetch + 폴링으로 정상 사용 중에도
        # 1000/시간 초과 → 1시간 IP 차단 사고 발생. 로그인 자체가 신뢰 경계이므로
        # 인증된 세션은 rate limit skip. 익명 트래픽만 제한.
        try:
            from flask import session
            if session.get('user_email') or session.get('logged_in'):
                return True
        except Exception:
            pass

        # 개발 환경에서는 rate limiting 대폭 완화
        is_development = self.app.config.get('DEBUG') or os.getenv('FLASK_ENV') == 'development'

        if is_development:
            # 개발 환경: 매우 관대한 제한
            limit = 10000
            window = 3600
        elif '/api/' in request.path:
            # 프로덕션 API — 20명 매니저 동시 편집·조회 대비. 사용자당 250/분 여유.
            # SocketIO polling은 별도 제외됨.
            limit = 15000
            window = 3600
        else:
            # 프로덕션 일반 페이지 — 다중 탭·프리페치·백그라운드 폴링 대비.
            # 2026-07-22 상향: /projects 접속 자체 차단 사고 대응 (1000→10000).
            # 매니저 브라우저가 여러 폴링 + prefetch 로 시간당 500~800 요청 발생.
            limit = 10000
            window = 3600

        return self.rate_limiter.is_allowed(identifier, limit, window)

    def _check_csrf_token(self) -> bool:
        """CSRF 토큰 확인"""
        # API 요청은 별도 검증 (JWT 토큰 등)
        if request.path.startswith('/api/') or request.path.startswith('/admin/api/'):
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
        logger.debug(f"[VALIDATION] JSON 입력 검증 시작: type={type(data).__name__}")

        if isinstance(data, dict):
            logger.debug(f"[VALIDATION] dict 검증 시작: {len(data)}개 필드")
            for key, value in data.items():
                if not self._validate_field(key, value):
                    logger.warning(f"[VALIDATION] ✗ JSON 검증 실패: key={key}")
                    return False
        elif isinstance(data, list):
            logger.debug(f"[VALIDATION] list 검증 시작: {len(data)}개 아이템")
            for item in data:
                if not self._validate_json_input(item):
                    logger.warning(f"[VALIDATION] ✗ JSON 검증 실패: list item")
                    return False
        elif isinstance(data, str):
            result = self.validator.is_safe_string(data)
            logger.debug(f"[VALIDATION] {'✓' if result else '✗'} string 검증: {repr(data)[:100]}")
            return result

        logger.debug(f"[VALIDATION] ✓ JSON 검증 성공")
        return True

    def _validate_field(self, field_name: str, value: Any) -> bool:
        """필드별 검증"""
        # 수금 확인 필드만 WARNING 레벨로 로깅
        if '수금' in field_name or '확인' in field_name:
            logger.warning(f"[VALIDATION] 필드 검증 시작: field_name={field_name}, value={repr(value)[:100]}, type={type(value).__name__}")

        if value is None:
            if '수금' in field_name or '확인' in field_name:
                logger.warning(f"[VALIDATION] ✓ {field_name}: None 허용")
            return True

        # Boolean 타입은 항상 허용 (int의 서브클래스이므로 먼저 체크)
        if isinstance(value, bool):
            if '수금' in field_name or '확인' in field_name:
                logger.warning(f"[VALIDATION] ✓ {field_name}: Boolean 값 허용 - {value}")
            return True

        if isinstance(value, str):
            # Boolean 문자열('true', 'false')도 허용
            if value.lower() in ('true', 'false'):
                if '수금' in field_name or '확인' in field_name:
                    logger.warning(f"[VALIDATION] ✓ {field_name}: Boolean 문자열 허용 - {value}")
                return True
            # HTML/특수문자 허용 필드는 패턴 매칭으로 확인
            # XSS/SQL/Command 검증을 건너뛰고 길이만 체크
            if any(pattern in field_name for pattern in self.HTML_ALLOWED_FIELD_PATTERNS):
                # 길이 제한만 체크 (5000자)
                result = len(value) <= 5000
                logger.debug(f"[VALIDATION] {'✓' if result else '✗'} {field_name}: HTML 허용 필드, 길이={len(value)}")
                return result

            # 일반 문자열은 기본 안전성 검사
            if not self.validator.is_safe_string(value):
                logger.warning(f"[VALIDATION] ✗ {field_name}: 안전하지 않은 문자열 - value={repr(value)[:100]}")
                g.validation_error = {
                    'field': field_name,
                    'message': f'{field_name} 필드에 허용되지 않는 문자가 포함되어 있습니다'
                }
                return False

            # 필드별 특수 검증
            if 'email' in field_name.lower():
                # 빈 문자열/대시는 선택 입력으로 허용
                if not value or value.strip() in ('', '-'):
                    return True
                result = self.validator.validate_email(value)
                logger.debug(f"[VALIDATION] {'✓' if result else '✗'} {field_name}: 이메일 검증 - {value}")
                if not result:
                    g.validation_error = {
                        'field': field_name,
                        'message': f'{field_name} 필드의 이메일 형식이 올바르지 않습니다'
                    }
                return result
            elif 'phone' in field_name.lower() or '연락처' in field_name:
                # 빈 문자열은 선택 입력으로 허용 (발주처 담당자 등이 모를 수 있음)
                if not value or value.strip() in ('', '-'):
                    return True
                result = self.validator.validate_phone(value)
                logger.debug(f"[VALIDATION] {'✓' if result else '✗'} {field_name}: 전화번호 검증 - {value}")
                if not result:
                    g.validation_error = {
                        'field': field_name,
                        'message': f'{field_name} 필드의 전화번호 형식이 올바르지 않습니다'
                    }
                return result
            elif '프로젝트' in field_name and '코드' in field_name:
                result = self.validator.validate_project_code(value)
                logger.debug(f"[VALIDATION] {'✓' if result else '✗'} {field_name}: 프로젝트 코드 검증 - {value}")
                if not result:
                    g.validation_error = {
                        'field': field_name,
                        'message': f'{field_name} 필드의 프로젝트 코드 형식이 올바르지 않습니다'
                    }
                return result
            # 금액 필드 명시적 리스트 (날짜 제외, '수금' 단독 키워드 제거)
            elif any(keyword in field_name for keyword in ['계약금', '중도금', '잔금', '총액', '비용', '마진', '순익']):
                result = self.validator.validate_amount(value)
                logger.debug(f"[VALIDATION] {'✓' if result else '✗'} {field_name}: 금액 검증 - {value}")
                if not result:
                    g.validation_error = {
                        'field': field_name,
                        'message': f'{field_name} 필드의 금액 형식이 올바르지 않습니다 (0~9999억 범위)'
                    }
                return result

            # 문자열 길이 제한
            result = len(value) <= 1000
            logger.debug(f"[VALIDATION] {'✓' if result else '✗'} {field_name}: 일반 문자열, 길이={len(value)}")
            if not result:
                g.validation_error = {
                    'field': field_name,
                    'message': f'{field_name} 필드가 너무 깁니다 (최대 1000자)'
                }
            return result

        elif isinstance(value, (int, float)):
            # 숫자 범위 확인
            result = -999999999999 <= value <= 999999999999
            logger.debug(f"[VALIDATION] {'✓' if result else '✗'} {field_name}: 숫자 범위 검증 - {value}")
            return result

        elif isinstance(value, (dict, list)):
            result = self._validate_json_input(value)
            logger.debug(f"[VALIDATION] {'✓' if result else '✗'} {field_name}: JSON 검증")
            return result

        logger.debug(f"[VALIDATION] ✓ {field_name}: 기타 타입 허용 - {type(value).__name__}")
        return True

    def _get_client_identifier(self) -> str:
        """클라이언트 식별자 생성"""
        # IP 주소 기반 (프록시 환경에서는 X-Forwarded-For 고려)
        ip = request.headers.get('X-Forwarded-For', request.remote_addr)
        if ip:
            ip = ip.split(',')[0].strip()

        # 사용자 세션이 있으면 이메일 사용 (안전한 접근)
        if 'user' in session:
            user_email = session.get('user', {}).get('email', 'unknown')
            return f"user_{user_email}"

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
csrf_protection = None


def init_security_middleware(app, secret_key: str):
    """보안 미들웨어 초기화"""
    global security_middleware, csrf_protection
    security_middleware = SecurityMiddleware(app, secret_key)
    # CSRF 보호 인스턴스를 전역으로도 제공 (API 엔드포인트에서 직접 사용 가능)
    csrf_protection = security_middleware.csrf_protection

    @app.before_request
    def before_request():
        if security_middleware:
            return security_middleware.process_request()

    @app.after_request
    def after_request(response):
        if security_middleware:
            return security_middleware.process_response(response)
        return response


def get_security_middleware() -> Optional[SecurityMiddleware]:
    """보안 미들웨어 인스턴스 반환"""
    return security_middleware