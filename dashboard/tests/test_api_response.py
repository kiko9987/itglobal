"""
API 응답 시스템 테스트

2026-09-11 현행 API 에 맞춰 갱신:
- APIResponse.success/created 의 message 는 data.message 에, 타임스탬프는 meta.timestamp 에 위치.
- error.code 는 Enum 의 .value(문자열). 에러코드명: RESOURCE_NOT_FOUND / VALIDATION_ERROR.
- 페이지네이션은 PaginationHelper.paginated_response 사용(구 APIResponse.paginated 제거됨).
- APIValidator / validate_json_request 는 제거된 API 라 관련 테스트 폐기.
"""

import pytest
import json
from dashboard.api.responses import APIResponse, APIErrorCode, PaginationHelper


class TestAPIResponse:
    """API 응답 클래스 테스트"""

    def test_success_response(self, app):
        """성공 응답 테스트 (message 는 data.message, timestamp 는 meta)"""
        with app.app_context():
            response = APIResponse.success(
                data={'test': 'data'},
                message='Test success'
            )

            assert response.status_code == 200
            data = json.loads(response.data)

            assert data['success'] is True
            assert data['data']['message'] == 'Test success'
            assert data['data']['test'] == 'data'
            assert 'timestamp' in data['meta']

    def test_error_response(self, app):
        """에러 응답 테스트 (code 는 enum .value 문자열)"""
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
            assert data['error']['code'] == APIErrorCode.BAD_REQUEST.value
            assert 'timestamp' in data['meta']

    def test_validation_error_response(self, app):
        """유효성 검사 실패 응답 테스트"""
        with app.app_context():
            errors = ['Field is required', 'Invalid format']
            response = APIResponse.validation_error(errors)

            assert response.status_code == 400
            data = json.loads(response.data)

            assert data['success'] is False
            assert data['error']['code'] == APIErrorCode.VALIDATION_ERROR.value
            assert data['error']['details'] == errors

    def test_not_found_response(self, app):
        """리소스를 찾을 수 없음 응답 테스트"""
        with app.app_context():
            response = APIResponse.not_found('User')

            assert response.status_code == 404
            data = json.loads(response.data)

            assert data['success'] is False
            assert data['error']['message'] == 'User not found'
            assert data['error']['code'] == APIErrorCode.RESOURCE_NOT_FOUND.value

    def test_paginated_response(self, app):
        """페이지네이션 응답 테스트 (PaginationHelper.paginated_response)"""
        with app.app_context():
            test_data = [{'id': 1}, {'id': 2}]
            response = PaginationHelper.paginated_response(
                items=test_data,
                page=1,
                limit=10,
                total=25
            )

            assert response.status_code == 200
            data = json.loads(response.data)

            assert data['success'] is True
            assert data['data']['items'] == test_data
            pg = data['meta']['pagination']
            assert pg['page'] == 1
            assert pg['limit'] == 10
            assert pg['total'] == 25
            assert pg['pages'] == 3
            assert pg['has_next'] is True
            assert pg['has_prev'] is False

    def test_created_response(self, app):
        """생성 성공 응답 테스트"""
        with app.app_context():
            response = APIResponse.created({'id': 123}, 'User created')

            assert response.status_code == 201
            data = json.loads(response.data)

            assert data['success'] is True
            assert data['data']['message'] == 'User created'
            assert data['data']['id'] == 123

    def test_updated_response(self, app):
        """갱신 성공 응답 테스트 (2026-09-11 추가된 헬퍼)"""
        with app.app_context():
            response = APIResponse.updated(message='Updated')

            assert response.status_code == 200
            data = json.loads(response.data)
            assert data['success'] is True
            assert data['data']['message'] == 'Updated'


class TestAPIDecorators:
    """API 데코레이터 테스트"""

    def test_api_response_decorator_success(self, app, client):
        """api_response 데코레이터 성공 케이스 (반환 dict 를 data 로 래핑)"""
        from dashboard.api.responses import api_response

        @app.route('/test-success')
        @api_response
        def test_endpoint():
            return {'value': 'test'}

        response = client.get('/test-success')

        assert response.status_code == 200
        data = json.loads(response.data)

        assert data['success'] is True
        assert data['data']['value'] == 'test'

    def test_api_response_decorator_error(self, app, client):
        """api_response 데코레이터 에러 케이스 (예외 → 500 INTERNAL_ERROR)"""
        from dashboard.api.responses import api_response

        @app.route('/test-error')
        @api_response
        def test_endpoint():
            raise ValueError("Test error")

        response = client.get('/test-error')

        assert response.status_code == 500
        data = json.loads(response.data)

        assert data['success'] is False
        assert data['error']['code'] == APIErrorCode.INTERNAL_ERROR.value
