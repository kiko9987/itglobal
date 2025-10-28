"""
JWT 인증 시스템 테스트
"""

import pytest
import jwt
from datetime import datetime, timezone, timedelta
from .jwt_auth import JWTManager, create_jwt_tokens, refresh_jwt_token

class TestJWTManager:
    """JWT 매니저 테스트"""
    
    def test_generate_tokens(self):
        """토큰 생성 테스트"""
        manager = JWTManager('test-secret')
        user_data = {
            'email': 'test@example.com',
            'name': '테스트 사용자',
            'permission_level': 'admin'
        }
        
        tokens = manager.generate_tokens(user_data)
        
        # 토큰 구조 확인
        assert 'access_token' in tokens
        assert 'refresh_token' in tokens
        assert 'token_type' in tokens
        assert 'expires_in' in tokens
        
        assert tokens['token_type'] == 'Bearer'
        assert tokens['expires_in'] == 3600  # 1시간
    
    def test_verify_access_token_success(self):
        """액세스 토큰 검증 성공 테스트"""
        manager = JWTManager('test-secret')
        user_data = {
            'email': 'test@example.com',
            'name': '테스트 사용자',
            'permission_level': 'admin'
        }
        
        tokens = manager.generate_tokens(user_data)
        payload = manager.verify_token(tokens['access_token'])
        
        assert payload is not None
        assert payload['user_id'] == user_data['email']
        assert payload['name'] == user_data['name']
        assert payload['permission_level'] == user_data['permission_level']
        assert payload['type'] == 'access'
    
    def test_verify_refresh_token_success(self):
        """리프레시 토큰 검증 성공 테스트"""
        manager = JWTManager('test-secret')
        user_data = {
            'email': 'test@example.com',
            'name': '테스트 사용자',
            'permission_level': 'admin'
        }
        
        tokens = manager.generate_tokens(user_data)
        payload = manager.verify_token(tokens['refresh_token'], 'refresh')
        
        assert payload is not None
        assert payload['user_id'] == user_data['email']
        assert payload['type'] == 'refresh'
        assert 'jti' in payload  # 토큰 고유 ID
    
    def test_verify_token_failure(self):
        """토큰 검증 실패 테스트"""
        manager = JWTManager('test-secret')
        
        # 잘못된 토큰
        assert manager.verify_token('invalid-token') is None
        
        # 다른 시크릿으로 생성된 토큰
        other_manager = JWTManager('other-secret')
        tokens = other_manager.generate_tokens({'email': 'test@example.com'})
        assert manager.verify_token(tokens['access_token']) is None
    
    def test_verify_token_wrong_type(self):
        """잘못된 토큰 타입 검증 테스트"""
        manager = JWTManager('test-secret')
        user_data = {'email': 'test@example.com'}
        
        tokens = manager.generate_tokens(user_data)
        
        # 액세스 토큰을 리프레시 토큰으로 검증 시도
        assert manager.verify_token(tokens['access_token'], 'refresh') is None
        
        # 리프레시 토큰을 액세스 토큰으로 검증 시도
        assert manager.verify_token(tokens['refresh_token'], 'access') is None
    
    def test_refresh_access_token_success(self):
        """액세스 토큰 새로고침 성공 테스트"""
        manager = JWTManager('test-secret')
        user_data = {
            'email': 'test@example.com',
            'name': '테스트 사용자',
            'permission_level': 'admin'
        }
        
        tokens = manager.generate_tokens(user_data)
        new_tokens = manager.refresh_access_token(tokens['refresh_token'])
        
        assert new_tokens is not None
        assert 'access_token' in new_tokens
        assert 'token_type' in new_tokens
        assert 'expires_in' in new_tokens
        
        # 새 액세스 토큰 검증
        payload = manager.verify_token(new_tokens['access_token'])
        assert payload is not None
        assert payload['user_id'] == user_data['email']
    
    def test_refresh_access_token_failure(self):
        """액세스 토큰 새로고침 실패 테스트"""
        manager = JWTManager('test-secret')
        
        # 잘못된 리프레시 토큰
        assert manager.refresh_access_token('invalid-token') is None
        
        # 액세스 토큰으로 새로고침 시도
        user_data = {'email': 'test@example.com'}
        tokens = manager.generate_tokens(user_data)
        assert manager.refresh_access_token(tokens['access_token']) is None
    
    def test_token_expiry(self):
        """토큰 만료 테스트"""
        # 짧은 만료 시간으로 매니저 생성
        manager = JWTManager('test-secret')
        manager.access_token_expire = -1  # 즉시 만료
        
        user_data = {'email': 'test@example.com'}
        tokens = manager.generate_tokens(user_data)
        
        # 만료된 토큰 검증
        assert manager.verify_token(tokens['access_token']) is None

class TestJWTDecorators:
    """JWT 데코레이터 테스트"""
    
    def test_jwt_required_with_valid_token(self, app, client):
        """유효한 토큰으로 보호된 엔드포인트 접근 테스트"""
        from .jwt_auth import jwt_required, init_jwt_manager, create_jwt_tokens
        
        with app.app_context():
            init_jwt_manager(app.config['SECRET_KEY'])
            
            @app.route('/protected')
            @jwt_required()
            def protected():
                from flask import request
                return {'user': request.jwt_user['email']}
            
            # 토큰 생성
            user_data = {'email': 'test@example.com', 'name': '테스트'}
            tokens = create_jwt_tokens(user_data)
            
            # 보호된 엔드포인트 접근
            response = client.get('/protected', headers={
                'Authorization': f"Bearer {tokens['access_token']}"
            })
            
            assert response.status_code == 200
            data = response.get_json()
            assert data['user'] == 'test@example.com'
    
    def test_jwt_required_without_token(self, app, client):
        """토큰 없이 보호된 엔드포인트 접근 테스트"""
        from .jwt_auth import jwt_required, init_jwt_manager
        
        with app.app_context():
            init_jwt_manager(app.config['SECRET_KEY'])
            
            @app.route('/protected')
            @jwt_required()
            def protected():
                return {'message': 'success'}
            
            response = client.get('/protected')
            
            assert response.status_code == 401
            data = response.get_json()
            assert data['code'] == 'MISSING_AUTH_HEADER'
    
    def test_jwt_required_admin_permission(self, app, client):
        """관리자 권한 필요 엔드포인트 테스트"""
        from .jwt_auth import jwt_required, init_jwt_manager, create_jwt_tokens
        
        with app.app_context():
            init_jwt_manager(app.config['SECRET_KEY'])
            
            @app.route('/admin')
            @jwt_required(admin_required=True)
            def admin_only():
                return {'message': 'admin access'}
            
            # 일반 사용자 토큰
            user_tokens = create_jwt_tokens({
                'email': 'user@example.com',
                'permission_level': 'viewer'
            })
            
            response = client.get('/admin', headers={
                'Authorization': f"Bearer {user_tokens['access_token']}"
            })
            
            assert response.status_code == 403
            
            # 관리자 토큰
            admin_tokens = create_jwt_tokens({
                'email': 'admin@example.com',
                'permission_level': 'admin'
            })
            
            response = client.get('/admin', headers={
                'Authorization': f"Bearer {admin_tokens['access_token']}"
            })
            
            assert response.status_code == 200