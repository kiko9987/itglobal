"""
API 응답 시스템 테스트
"""

import pytest
import json
from utils.api_response import APIResponse, APIErrorCode, APIValidator

class TestAPIResponse:
    """API 응답 클래스 테스트"""
    
    def test_success_response(self, app):
        """성공 응답 테스트"""
        with app.app_context():
            response = APIResponse.success(
                data={'test': 'data'}, 
                message='Test success'
            )
            
            assert response.status_code == 200
            data = json.loads(response.data)
            
            assert data['success'] is True
            assert data['message'] == 'Test success'
            assert data['data']['test'] == 'data'
            assert 'timestamp' in data
    
    def test_error_response(self, app):
        """에러 응답 테스트"""
        with app.app_context():
            response = APIResponse.error(
                message='Test error',
                error_code=APIErrorCode.BAD_REQUEST,
                status_code=400
            )
            
            assert response.status_code == 400
            data = json.loads(response.data)
            
            assert data['success'] is False
            assert data['error']['message'] == 'Test error'
            assert data['error']['code'] == APIErrorCode.BAD_REQUEST
            assert 'timestamp' in data
    
    def test_validation_error_response(self, app):
        """유효성 검사 실패 응답 테스트"""
        with app.app_context():
            errors = ['Field is required', 'Invalid format']
            response = APIResponse.validation_error(errors)
            
            assert response.status_code == 400
            data = json.loads(response.data)
            
            assert data['success'] is False
            assert data['error']['code'] == APIErrorCode.VALIDATION_FAILED
            assert data['error']['details'] == errors
    
    def test_not_found_response(self, app):
        """리소스를 찾을 수 없음 응답 테스트"""
        with app.app_context():
            response = APIResponse.not_found('User')
            
            assert response.status_code == 404
            data = json.loads(response.data)
            
            assert data['success'] is False
            assert data['error']['message'] == 'User not found'
            assert data['error']['code'] == APIErrorCode.NOT_FOUND
    
    def test_paginated_response(self, app):
        """페이지네이션 응답 테스트"""
        with app.app_context():
            test_data = [{'id': 1}, {'id': 2}]
            response = APIResponse.paginated(
                data=test_data,
                page=1,
                per_page=10,
                total=25
            )
            
            assert response.status_code == 200
            data = json.loads(response.data)
            
            assert data['success'] is True
            assert data['data'] == test_data
            assert data['metadata']['pagination']['page'] == 1
            assert data['metadata']['pagination']['per_page'] == 10
            assert data['metadata']['pagination']['total'] == 25
            assert data['metadata']['pagination']['total_pages'] == 3
            assert data['metadata']['pagination']['has_next'] is True
            assert data['metadata']['pagination']['has_prev'] is False
    
    def test_created_response(self, app):
        """생성 성공 응답 테스트"""
        with app.app_context():
            response = APIResponse.created({'id': 123}, 'User created')
            
            assert response.status_code == 201
            data = json.loads(response.data)
            
            assert data['success'] is True
            assert data['message'] == 'User created'
            assert data['data']['id'] == 123

class TestAPIValidator:
    """API 유효성 검사 테스트"""
    
    def test_validate_required_fields(self):
        """필수 필드 검증 테스트"""
        data = {'name': 'John', 'age': 30}
        required_fields = ['name', 'email', 'age']
        
        errors = APIValidator.validate_required_fields(data, required_fields)
        
        assert len(errors) == 1
        assert "'email' is required" in errors
    
    def test_validate_field_types(self):
        """필드 타입 검증 테스트"""
        data = {'name': 'John', 'age': '30', 'active': True}
        field_types = {'name': str, 'age': int, 'active': bool}
        
        errors = APIValidator.validate_field_types(data, field_types)
        
        assert len(errors) == 1
        assert "'age' must be of type int" in errors
    
    def test_validate_field_lengths(self):
        """필드 길이 검증 테스트"""
        data = {'username': 'ab', 'description': 'a' * 1000}
        field_lengths = {
            'username': {'min': 3, 'max': 20},
            'description': {'min': 0, 'max': 500}
        }
        
        errors = APIValidator.validate_field_lengths(data, field_lengths)
        
        assert len(errors) == 2
        assert "'username' must be at least 3 characters" in errors
        assert "'description' must be at most 500 characters" in errors

class TestAPIDecorators:
    """API 데코레이터 테스트"""
    
    def test_api_response_decorator_success(self, app, client):
        """API 응답 데코레이터 성공 케이스 테스트"""
        from utils.api_response import api_response
        
        with app.app_context():
            @app.route('/test-success')
            @api_response
            def test_endpoint():
                return {'data': 'test'}
            
            response = client.get('/test-success')
            
            assert response.status_code == 200
            data = json.loads(response.data)
            
            assert data['success'] is True
            assert data['data'] == 'test'
    
    def test_api_response_decorator_error(self, app, client):
        """API 응답 데코레이터 에러 케이스 테스트"""
        from utils.api_response import api_response
        
        with app.app_context():
            @app.route('/test-error')
            @api_response
            def test_endpoint():
                raise ValueError("Test error")
            
            response = client.get('/test-error')
            
            assert response.status_code == 500
            data = json.loads(response.data)
            
            assert data['success'] is False
            assert data['error']['code'] == APIErrorCode.INTERNAL_ERROR
    
    def test_validate_json_request_decorator(self, app, client):
        """JSON 요청 검증 데코레이터 테스트"""
        from utils.api_response import validate_json_request, api_response
        
        with app.app_context():
            @app.route('/test-validation', methods=['POST'])
            @validate_json_request('name', 'email', name=str, age=int)
            @api_response
            def test_endpoint():
                return {'message': 'success'}
            
            # 유효한 요청
            response = client.post('/test-validation', 
                                 json={'name': 'John', 'email': 'john@example.com', 'age': 30})
            assert response.status_code == 200
            
            # 필수 필드 누락
            response = client.post('/test-validation', 
                                 json={'name': 'John'})
            assert response.status_code == 400
            data = json.loads(response.data)
            assert data['error']['code'] == APIErrorCode.VALIDATION_FAILED
            
            # 잘못된 타입
            response = client.post('/test-validation', 
                                 json={'name': 'John', 'email': 'john@example.com', 'age': '30'})
            assert response.status_code == 400