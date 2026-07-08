"""
API Middleware for Request/Response Processing
Implements standardized error handling, request tracking, and response formatting
"""

from flask import Flask, request, g, current_app
from datetime import datetime
import uuid
import time
import logging
from typing import Optional
from werkzeug.exceptions import HTTPException
import traceback

from .responses import APIResponse, APIException, APIErrorCode
from dashboard.utils.error_helpers import generate_error_id

try:
    from googleapiclient.errors import HttpError as _GoogleHttpError
except Exception:  # googleapiclient가 없는 환경도 안전 fallback
    _GoogleHttpError = None


class RequestTrackingMiddleware:
    """Middleware to track requests and add metadata to Flask g object"""

    def __init__(self, app: Optional[Flask] = None):
        self.app = app
        if app:
            self.init_app(app)

    def init_app(self, app: Flask):
        """Initialize the middleware with Flask app"""
        app.before_request(self.before_request)
        app.after_request(self.after_request)

    @staticmethod
    def before_request():
        """Execute before each request"""
        # Generate unique request ID
        g.request_id = str(uuid.uuid4())
        g.request_start_time = time.time()
        g.request_timestamp = datetime.utcnow()

        # Add request metadata
        g.request_method = request.method
        g.request_path = request.path
        g.request_endpoint = request.endpoint
        g.request_args = dict(request.args)
        g.request_headers = dict(request.headers)

        # Log request start (if API endpoint)
        if request.path.startswith('/api/'):
            current_app.logger.info(
                f"API Request Started - {g.request_method} {g.request_path} "
                f"[Request ID: {g.request_id}]"
            )

    @staticmethod
    def after_request(response):
        """Execute after each request"""
        # Calculate request duration
        if hasattr(g, 'request_start_time'):
            duration = (time.time() - g.request_start_time) * 1000  # Convert to milliseconds
            g.request_duration = round(duration, 2)

            # Log request completion (if API endpoint)
            if request.path.startswith('/api/'):
                current_app.logger.info(
                    f"API Request Completed - {g.request_method} {g.request_path} "
                    f"[{response.status_code}] [{g.request_duration}ms] "
                    f"[Request ID: {g.request_id}]"
                )

            # Add performance headers for monitoring
            response.headers['X-Request-ID'] = g.request_id
            response.headers['X-Response-Time'] = f"{g.request_duration}ms"

        return response


class APIErrorHandler:
    """Centralized API error handling middleware"""

    def __init__(self, app: Optional[Flask] = None):
        self.app = app
        if app:
            self.init_app(app)

    def init_app(self, app: Flask):
        """Initialize error handlers with Flask app"""
        # Register error handlers
        app.register_error_handler(APIException, self.handle_api_exception)
        app.register_error_handler(HTTPException, self.handle_http_exception)
        app.register_error_handler(Exception, self.handle_general_exception)

        # Register validation error handlers
        try:
            from marshmallow import ValidationError
            app.register_error_handler(ValidationError, self.handle_validation_error)
        except ImportError:
            pass  # Marshmallow not installed

    @staticmethod
    def handle_api_exception(error: APIException):
        """Handle custom API exceptions"""
        current_app.logger.warning(
            f"API Exception: {error.error_code} - {error.message} "
            f"[Request ID: {getattr(g, 'request_id', 'unknown')}]"
        )
        return error.to_response()

    @staticmethod
    def handle_http_exception(error: HTTPException):
        """Handle standard HTTP exceptions"""
        # Only handle API routes with standard responses
        if not request.path.startswith('/api/'):
            return error  # Let Flask handle non-API routes normally

        error_code_mapping = {
            400: APIErrorCode.BAD_REQUEST,
            401: APIErrorCode.UNAUTHORIZED,
            403: APIErrorCode.FORBIDDEN,
            404: APIErrorCode.RESOURCE_NOT_FOUND,
            409: APIErrorCode.CONFLICT,
            429: APIErrorCode.RATE_LIMITED,
            500: APIErrorCode.INTERNAL_ERROR,
            503: APIErrorCode.SERVICE_UNAVAILABLE
        }

        error_code = error_code_mapping.get(error.code, APIErrorCode.INTERNAL_ERROR)

        current_app.logger.warning(
            f"HTTP Exception: {error.code} - {error.description} "
            f"[Request ID: {getattr(g, 'request_id', 'unknown')}]"
        )

        return APIResponse.error(
            error_code=error_code,
            message=error.description or f"HTTP {error.code} Error",
            status_code=error.code
        )

    @staticmethod
    def handle_validation_error(error):
        """Handle Marshmallow validation errors"""
        current_app.logger.warning(
            f"Validation Error: {error.messages} "
            f"[Request ID: {getattr(g, 'request_id', 'unknown')}]"
        )

        return APIResponse.validation_error(
            details=error.messages,
            message="Request validation failed"
        )

    @staticmethod
    def handle_general_exception(error: Exception):
        """Handle unexpected general exceptions"""
        # Only handle API routes with standard responses
        if not request.path.startswith('/api/'):
            raise error  # Re-raise for non-API routes

        request_id = getattr(g, 'request_id', 'unknown')
        error_id = generate_error_id()

        # Google Sheets API 할당량/속도 초과 감지 → 500이 아니라 429 반환
        # (프론트 userFriendlyErrors.js가 429 안내 문구를 이미 지원)
        if _GoogleHttpError is not None and isinstance(error, _GoogleHttpError):
            status = getattr(getattr(error, 'resp', None), 'status', 0) or 0
            err_msg = str(error).lower()
            if status == 429 or (status == 403 and 'quota' in err_msg):
                current_app.logger.warning(
                    f"[{error_id}] Google Sheets 할당량 초과 (status={status}) "
                    f"[Request ID: {request_id}]"
                )
                return APIResponse.error(
                    error_code=APIErrorCode.RATE_LIMITED,
                    message=(
                        "Google 시트 요청이 일시적으로 몰려 처리되지 못했습니다. "
                        "잠시 후 다시 시도해주세요."
                    ),
                    details={"error_id": error_id},
                    status_code=429,
                )

        # Log detailed error information
        current_app.logger.error(
            f"[{error_id}] Unhandled Exception: {type(error).__name__}: {str(error)} "
            f"[Request ID: {request_id}]\n"
            f"Traceback: {traceback.format_exc()}"
        )

        # In development, include error details
        if current_app.config.get('DEBUG', False):
            details = {
                "type": type(error).__name__,
                "traceback": traceback.format_exc().split('\n'),
                "error_id": error_id,
            }
        else:
            # 프로덕션: 매니저가 문의할 때 로그 매칭에 쓸 error_id만 노출
            details = {"error_id": error_id}

        return APIResponse.error(
            error_code=APIErrorCode.INTERNAL_ERROR,
            message=(
                f"예기치 못한 오류가 발생했습니다. 잠시 후 다시 시도해주세요. "
                f"(오류 ID: {error_id})"
            ),
            details=details,
            status_code=500
        )


class CORSMiddleware:
    """CORS handling middleware for API endpoints"""

    def __init__(self, app: Optional[Flask] = None, origins: str = "*"):
        self.origins = origins
        self.app = app
        if app:
            self.init_app(app)

    def init_app(self, app: Flask):
        """Initialize CORS middleware with Flask app"""
        app.after_request(self.add_cors_headers)

    def add_cors_headers(self, response):
        """Add CORS headers to API responses"""
        if request.path.startswith('/api/'):
            response.headers['Access-Control-Allow-Origin'] = self.origins
            response.headers['Access-Control-Allow-Methods'] = 'GET, POST, PUT, PATCH, DELETE, OPTIONS'
            response.headers['Access-Control-Allow-Headers'] = 'Content-Type, Authorization, X-Requested-With'
            response.headers['Access-Control-Allow-Credentials'] = 'true'
            response.headers['Access-Control-Max-Age'] = '86400'  # 24 hours

        return response


class ResponseValidationMiddleware:
    """Middleware to validate API responses match expected format"""

    def __init__(self, app: Optional[Flask] = None, validate_responses: bool = True):
        self.validate_responses = validate_responses
        self.app = app
        if app:
            self.init_app(app)

    def init_app(self, app: Flask):
        """Initialize response validation middleware"""
        if self.validate_responses:
            app.after_request(self.validate_response)

    @staticmethod
    def validate_response(response):
        """Validate that API responses follow standard format"""
        # Only validate JSON API responses
        if (request.path.startswith('/api/') and
            response.content_type and
            'application/json' in response.content_type and
            response.status_code != 204):  # Skip validation for No Content responses

            try:
                import json
                data = json.loads(response.get_data(as_text=True))

                # Check required fields
                required_fields = ['success', 'data', 'error', 'meta']
                missing_fields = [field for field in required_fields if field not in data]

                if missing_fields:
                    current_app.logger.warning(
                        f"API Response missing required fields: {missing_fields} "
                        f"for {request.method} {request.path} "
                        f"[Request ID: {getattr(g, 'request_id', 'unknown')}]"
                    )

                # Check meta field structure
                if 'meta' in data and isinstance(data['meta'], dict):
                    meta_required = ['timestamp', 'version', 'request_id']
                    meta_missing = [field for field in meta_required if field not in data['meta']]

                    if meta_missing:
                        current_app.logger.warning(
                            f"API Response meta missing fields: {meta_missing} "
                            f"for {request.method} {request.path} "
                            f"[Request ID: {getattr(g, 'request_id', 'unknown')}]"
                        )

            except (json.JSONDecodeError, TypeError) as e:
                current_app.logger.warning(
                    f"Failed to validate API response JSON: {str(e)} "
                    f"for {request.method} {request.path} "
                    f"[Request ID: {getattr(g, 'request_id', 'unknown')}]"
                )

        return response


class SecurityHeadersMiddleware:
    """Middleware to add security headers to responses"""

    def __init__(self, app: Optional[Flask] = None):
        self.app = app
        if app:
            self.init_app(app)

    def init_app(self, app: Flask):
        """Initialize security headers middleware"""
        app.after_request(self.add_security_headers)

    @staticmethod
    def add_security_headers(response):
        """Add security headers to all responses"""
        # Basic security headers
        response.headers['X-Content-Type-Options'] = 'nosniff'
        response.headers['X-Frame-Options'] = 'DENY'
        response.headers['X-XSS-Protection'] = '1; mode=block'

        # API-specific headers
        if request.path.startswith('/api/'):
            response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'

        return response


def init_api_middleware(app: Flask, **config):
    """Initialize all API middleware components"""

    # Request tracking
    request_tracking = RequestTrackingMiddleware(app)

    # Error handling
    error_handler = APIErrorHandler(app)

    # CORS handling
    cors_origins = config.get('cors_origins', '*')
    cors_middleware = CORSMiddleware(app, origins=cors_origins)

    # Response validation (only in development)
    validate_responses = config.get('validate_responses', app.config.get('DEBUG', False))
    response_validator = ResponseValidationMiddleware(app, validate_responses=validate_responses)

    # Security headers
    security_headers = SecurityHeadersMiddleware(app)

    app.logger.info("API middleware initialization completed")

    return {
        'request_tracking': request_tracking,
        'error_handler': error_handler,
        'cors': cors_middleware,
        'response_validator': response_validator,
        'security_headers': security_headers
    }