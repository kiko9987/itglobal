"""
보안 시스템 테스트
CSRF, XSS, 입력값 검증 테스트
"""

import pytest
from utils.security import CSRFProtection, InputValidator, SecurityHeaders
from flask import Flask, session

class TestCSRFProtection:
    """CSRF 보호 테스트"""
    
    def test_generate_csrf_token(self, app):
        """CSRF 토큰 생성 테스트"""
        with app.test_request_context():
            csrf = CSRFProtection('test-secret')
            token = csrf.generate_csrf_token()
            
            assert token is not None
            assert len(token.split(':')) == 3
            assert session.get('csrf_token') == token
    
    def test_validate_csrf_token_success(self, app):
        """CSRF 토큰 검증 성공 테스트"""
        with app.test_request_context():
            csrf = CSRFProtection('test-secret')
            token = csrf.generate_csrf_token()
            
            assert csrf.validate_csrf_token(token) is True
    
    def test_validate_csrf_token_failure(self, app):
        """CSRF 토큰 검증 실패 테스트"""
        with app.test_request_context():
            csrf = CSRFProtection('test-secret')
            
            # 잘못된 토큰
            assert csrf.validate_csrf_token('invalid-token') is False
            
            # 세션에 토큰이 없는 경우
            assert csrf.validate_csrf_token('') is False
    
    def test_csrf_token_expiry(self, app):
        """CSRF 토큰 만료 테스트"""
        with app.test_request_context():
            csrf = CSRFProtection('test-secret')
            csrf.token_lifetime = -1  # 즉시 만료
            
            token = csrf.generate_csrf_token()
            assert csrf.validate_csrf_token(token) is False

class TestInputValidator:
    """입력값 검증 테스트"""
    
    def test_sanitize_html(self):
        """HTML 새니타이징 테스트"""
        validator = InputValidator()
        
        # 기본 HTML 태그 이스케이프
        assert validator.sanitize_html('<script>alert("xss")</script>') == '&lt;script&gt;alert(&quot;xss&quot;)&lt;/script&gt;'
        assert validator.sanitize_html('<img src="x" onerror="alert(1)">') == '&lt;img src=&quot;x&quot; onerror=&quot;alert(1)&quot;&gt;'
        
        # 빈 문자열 처리
        assert validator.sanitize_html('') == ''
        assert validator.sanitize_html(None) == ''
    
    def test_validate_email(self):
        """이메일 형식 검증 테스트"""
        validator = InputValidator()
        
        # 유효한 이메일
        assert validator.validate_email('test@example.com') is True
        assert validator.validate_email('user.name+tag@domain.co.kr') is True
        
        # 유효하지 않은 이메일
        assert validator.validate_email('invalid-email') is False
        assert validator.validate_email('@domain.com') is False
        assert validator.validate_email('user@') is False
        assert validator.validate_email('') is False
        assert validator.validate_email(None) is False
    
    def test_validate_project_code(self):
        """프로젝트 코드 형식 검증 테스트"""
        validator = InputValidator()
        
        # 유효한 프로젝트 코드
        assert validator.validate_project_code('G0001-IT') is True
        assert validator.validate_project_code('A1234-ABCD') is True
        
        # 유효하지 않은 프로젝트 코드
        assert validator.validate_project_code('invalid') is False
        assert validator.validate_project_code('G001-IT') is False  # 숫자가 4자리가 아님
        assert validator.validate_project_code('g0001-IT') is False  # 소문자
        assert validator.validate_project_code('') is False
    
    def test_validate_phone(self):
        """전화번호 형식 검증 테스트"""
        validator = InputValidator()
        
        # 유효한 전화번호
        assert validator.validate_phone('010-1234-5678') is True
        assert validator.validate_phone('02-123-4567') is True
        assert validator.validate_phone('031-123-4567') is True
        
        # 유효하지 않은 전화번호
        assert validator.validate_phone('invalid-phone') is False
        assert validator.validate_phone('010-12-34') is False
        assert validator.validate_phone('') is False
    
    def test_validate_amount(self):
        """금액 형식 검증 테스트"""
        validator = InputValidator()
        
        # 유효한 금액
        assert validator.validate_amount('1000000') is True
        assert validator.validate_amount('1,000,000') is True
        assert validator.validate_amount('₩1,000,000') is True
        assert validator.validate_amount('123.45') is True
        
        # 유효하지 않은 금액
        assert validator.validate_amount('invalid-amount') is False
        assert validator.validate_amount('') is False
    
    def test_clean_string(self):
        """문자열 정리 테스트"""
        validator = InputValidator()
        
        # HTML 이스케이프 및 공백 제거
        assert validator.clean_string('  <script>  ') == '&lt;script&gt;'
        
        # 길이 제한
        long_string = 'a' * 2000
        result = validator.clean_string(long_string, max_length=100)
        assert len(result) <= 103  # "..." 포함
        assert result.endswith('...')

class TestSecurityHeaders:
    """보안 헤더 테스트"""
    
    def test_get_security_headers(self):
        """보안 헤더 반환 테스트"""
        headers = SecurityHeaders.get_security_headers()
        
        # 필수 보안 헤더 확인
        assert 'X-Content-Type-Options' in headers
        assert 'X-Frame-Options' in headers
        assert 'X-XSS-Protection' in headers
        assert 'Content-Security-Policy' in headers
        
        # 헤더 값 확인
        assert headers['X-Content-Type-Options'] == 'nosniff'
        assert headers['X-Frame-Options'] == 'DENY'
        assert 'default-src' in headers['Content-Security-Policy']