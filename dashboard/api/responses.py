"""
표준 API 응답 구조 (Standard API Response Format)

이 모듈은 시스템 전체의 표준 API 응답 구조를 정의합니다.
API_STANDARDS.md 문서를 준수합니다.

주요 클래스:
- APIResponse: 표준 응답 생성기 (success, error, created, validation_error 등)
- APIErrorCode: 표준 에러 코드 Enum
- PaginationHelper: 페이지네이션 헬퍼
- APIException: API 예외 기본 클래스

Usage:
    from dashboard.api.responses import APIResponse, APIErrorCode

    # 성공 응답
    return APIResponse.success(data=projects)

    # 생성 성공 (201)
    return APIResponse.created(data=new_project)

    # 검증 오류 (400)
    return APIResponse.validation_error(details=errors)

    # 리소스 없음 (404)
    return APIResponse.not_found(resource="프로젝트", resource_id=code)

참고 문서:
    - dashboard/API_STANDARDS.md: HTTP 상태 코드 규칙, 사용 예시
    - dashboard/tests/test_api_response.py: 테스트 예시
"""

from datetime import datetime
from typing import Any, Dict, Optional, Union, List
from flask import jsonify, g, request
import uuid
from enum import Enum


class APIErrorCode(Enum):
    """Standard API error codes"""
    # Client Errors (4xx)
    VALIDATION_ERROR = "VALIDATION_ERROR"
    RESOURCE_NOT_FOUND = "RESOURCE_NOT_FOUND"
    UNAUTHORIZED = "UNAUTHORIZED"
    FORBIDDEN = "FORBIDDEN"
    CONFLICT = "CONFLICT"
    RATE_LIMITED = "RATE_LIMITED"
    BAD_REQUEST = "BAD_REQUEST"

    # Server Errors (5xx)
    INTERNAL_ERROR = "INTERNAL_ERROR"
    SERVICE_UNAVAILABLE = "SERVICE_UNAVAILABLE"
    DATABASE_ERROR = "DATABASE_ERROR"
    EXTERNAL_SERVICE_ERROR = "EXTERNAL_SERVICE_ERROR"
    TIMEOUT_ERROR = "TIMEOUT_ERROR"

    # Business Logic Errors
    PROJECT_NOT_FOUND = "PROJECT_NOT_FOUND"
    PROJECT_ALREADY_EXISTS = "PROJECT_ALREADY_EXISTS"
    PROJECT_LOCKED = "PROJECT_LOCKED"
    INSUFFICIENT_PERMISSIONS = "INSUFFICIENT_PERMISSIONS"
    INVALID_PROJECT_STATUS = "INVALID_PROJECT_STATUS"

    # User Management Errors
    USER_NOT_FOUND = "USER_NOT_FOUND"
    USER_ALREADY_EXISTS = "USER_ALREADY_EXISTS"
    INVALID_CREDENTIALS = "INVALID_CREDENTIALS"

    # Data Errors
    DATA_NOT_FOUND = "DATA_NOT_FOUND"
    DATA_CORRUPTION = "DATA_CORRUPTION"
    MISSING_REQUIRED_FIELD = "MISSING_REQUIRED_FIELD"


class APIResponse:
    """Standard API response wrapper"""

    @staticmethod
    def _get_request_id() -> str:
        """Get or generate request ID"""
        return getattr(g, 'request_id', str(uuid.uuid4()))

    @staticmethod
    def _create_meta(pagination: Optional[Dict] = None, **extra_meta) -> Dict:
        """Create standard metadata object"""
        meta = {
            "timestamp": datetime.utcnow().isoformat() + "Z",
            "version": "v1",
            "request_id": APIResponse._get_request_id()
        }

        if pagination:
            meta["pagination"] = pagination

        meta.update(extra_meta)
        return meta

    @staticmethod
    def success(
        data: Any = None,
        message: Optional[str] = None,
        status_code: int = 200,
        pagination: Optional[Dict] = None,
        **meta_extra
    ):
        """Create successful response"""
        response_data = {
            "success": True,
            "data": data,
            "error": None,
            "meta": APIResponse._create_meta(pagination=pagination, **meta_extra)
        }

        if message:
            if response_data["data"] is None:
                response_data["data"] = {}
            if isinstance(response_data["data"], dict):
                response_data["data"]["message"] = message

        return jsonify(response_data), status_code

    @staticmethod
    def error(
        error_code: Union[APIErrorCode, str],
        message: str,
        details: Optional[Dict] = None,
        status_code: int = 400,
        **meta_extra
    ):
        """Create error response"""
        if isinstance(error_code, APIErrorCode):
            error_code = error_code.value

        error_details = {
            "code": error_code,
            "message": message
        }

        if details:
            error_details["details"] = details

        response_data = {
            "success": False,
            "data": None,
            "error": error_details,
            "meta": APIResponse._create_meta(**meta_extra)
        }

        return jsonify(response_data), status_code

    @staticmethod
    def created(data: Any = None, message: str = "Resource created successfully", **meta_extra):
        """Create 201 Created response"""
        return APIResponse.success(data=data, message=message, status_code=201, **meta_extra)

    @staticmethod
    def no_content(message: str = "Operation completed successfully"):
        """Create 204 No Content response"""
        return APIResponse.success(message=message, status_code=204)

    @staticmethod
    def validation_error(details: Dict, message: str = "Request validation failed"):
        """Create validation error response"""
        return APIResponse.error(
            APIErrorCode.VALIDATION_ERROR,
            message,
            details=details,
            status_code=400
        )

    @staticmethod
    def not_found(resource: str = "Resource", resource_id: Optional[str] = None):
        """Create not found error response"""
        message = f"{resource} not found"
        if resource_id:
            message += f" (ID: {resource_id})"

        details = {"resource": resource}
        if resource_id:
            details["resource_id"] = resource_id

        return APIResponse.error(
            APIErrorCode.RESOURCE_NOT_FOUND,
            message,
            details=details,
            status_code=404
        )

    @staticmethod
    def unauthorized(message: str = "Authentication required"):
        """Create unauthorized error response"""
        return APIResponse.error(
            APIErrorCode.UNAUTHORIZED,
            message,
            status_code=401
        )

    @staticmethod
    def forbidden(message: str = "Insufficient permissions"):
        """Create forbidden error response"""
        return APIResponse.error(
            APIErrorCode.FORBIDDEN,
            message,
            status_code=403
        )

    @staticmethod
    def conflict(resource: str, message: Optional[str] = None):
        """Create conflict error response"""
        if not message:
            message = f"{resource} already exists or is in use"

        return APIResponse.error(
            APIErrorCode.CONFLICT,
            message,
            details={"resource": resource},
            status_code=409
        )

    @staticmethod
    def internal_error(message: str = "An internal server error occurred"):
        """Create internal server error response"""
        return APIResponse.error(
            APIErrorCode.INTERNAL_ERROR,
            message,
            status_code=500
        )

    @staticmethod
    def service_unavailable(message: str = "Service is temporarily unavailable"):
        """Create service unavailable error response"""
        return APIResponse.error(
            APIErrorCode.SERVICE_UNAVAILABLE,
            message,
            status_code=503
        )


class PaginationHelper:
    """Helper class for pagination metadata"""

    @staticmethod
    def create_pagination_meta(
        page: int,
        limit: int,
        total: int,
        items: List = None
    ) -> Dict:
        """Create pagination metadata"""
        pages = (total + limit - 1) // limit  # Ceiling division
        has_next = page < pages
        has_prev = page > 1

        pagination = {
            "page": page,
            "limit": limit,
            "total": total,
            "pages": pages,
            "has_next": has_next,
            "has_prev": has_prev,
            "next_page": page + 1 if has_next else None,
            "prev_page": page - 1 if has_prev else None
        }

        if items is not None:
            pagination["count"] = len(items)

        return pagination

    @staticmethod
    def paginated_response(
        items: List,
        page: int,
        limit: int,
        total: int,
        message: Optional[str] = None
    ):
        """Create paginated response"""
        pagination = PaginationHelper.create_pagination_meta(page, limit, total, items)

        data = {
            "items": items,
            "summary": {
                "total": total,
                "page": page,
                "limit": limit,
                "pages": pagination["pages"]
            }
        }

        return APIResponse.success(
            data=data,
            message=message,
            pagination=pagination
        )


def api_response(func):
    """Decorator to automatically wrap response in standard format"""
    from functools import wraps

    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            result = func(*args, **kwargs)

            # If result is already a Response object, return as-is
            if hasattr(result, 'status_code'):
                return result

            # If result is a tuple (data, status_code)
            if isinstance(result, tuple) and len(result) == 2:
                data, status_code = result
                return APIResponse.success(data=data, status_code=status_code)

            # If result is just data
            return APIResponse.success(data=result)

        except Exception as e:
            # Log the error
            import logging
            logging.error(f"API error in {func.__name__}: {str(e)}")
            return APIResponse.internal_error(f"Error in {func.__name__}: {str(e)}")

    return wrapper


class APIException(Exception):
    """Base exception class for API errors"""

    def __init__(
        self,
        error_code: Union[APIErrorCode, str],
        message: str,
        details: Optional[Dict] = None,
        status_code: int = 400
    ):
        self.error_code = error_code
        self.message = message
        self.details = details
        self.status_code = status_code
        super().__init__(message)

    def to_response(self):
        """Convert exception to API response"""
        return APIResponse.error(
            self.error_code,
            self.message,
            self.details,
            self.status_code
        )


class ValidationException(APIException):
    """Exception for validation errors"""

    def __init__(self, details: Dict, message: str = "Request validation failed"):
        super().__init__(APIErrorCode.VALIDATION_ERROR, message, details, 400)


class NotFoundException(APIException):
    """Exception for resource not found errors"""

    def __init__(self, resource: str, resource_id: Optional[str] = None):
        message = f"{resource} not found"
        details = {"resource": resource}

        if resource_id:
            message += f" (ID: {resource_id})"
            details["resource_id"] = resource_id

        super().__init__(APIErrorCode.RESOURCE_NOT_FOUND, message, details, 404)


class ConflictException(APIException):
    """Exception for resource conflict errors"""

    def __init__(self, resource: str, message: Optional[str] = None):
        if not message:
            message = f"{resource} already exists or is in use"

        super().__init__(
            APIErrorCode.CONFLICT,
            message,
            {"resource": resource},
            409
        )


class UnauthorizedException(APIException):
    """Exception for authentication errors"""

    def __init__(self, message: str = "Authentication required"):
        super().__init__(APIErrorCode.UNAUTHORIZED, message, None, 401)


class ForbiddenException(APIException):
    """Exception for authorization errors"""

    def __init__(self, message: str = "Insufficient permissions"):
        super().__init__(APIErrorCode.FORBIDDEN, message, None, 403)


# HTTP status code mapping for common scenarios
HTTP_STATUS_CODES = {
    'success': 200,
    'created': 201,
    'accepted': 202,
    'no_content': 204,
    'bad_request': 400,
    'unauthorized': 401,
    'forbidden': 403,
    'not_found': 404,
    'method_not_allowed': 405,
    'conflict': 409,
    'unprocessable_entity': 422,
    'rate_limited': 429,
    'internal_error': 500,
    'service_unavailable': 503
}