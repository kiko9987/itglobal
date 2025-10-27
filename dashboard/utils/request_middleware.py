"""
Flask 요청 추적 미들웨어
각 요청에 고유 ID 부여 및 성능 추적
"""

import uuid
import time
from datetime import datetime
from flask import request, g, session
from functools import wraps
from typing import Optional

from .logging_config import get_logger, log_business_event


class RequestTracker:
    """요청 추적 및 성능 모니터링 클래스"""

    def __init__(self, app=None):
        self.app = app
        self.logger = get_logger(__name__)

        if app is not None:
            self.init_app(app)

    def init_app(self, app):
        """Flask 앱에 요청 추적 미들웨어 등록"""
        app.before_request(self._before_request)
        app.after_request(self._after_request)
        app.teardown_appcontext(self._teardown_request)

    def _before_request(self):
        """요청 시작 시 실행"""
        # 고유한 요청 ID 생성
        g.request_id = str(uuid.uuid4())[:8]
        g.start_time = time.time()
        g.start_datetime = datetime.now()

        # 요청 정보 로깅
        self.logger.info(
            f"Request started: {request.method} {request.path}",
            request_id=g.request_id,
            method=request.method,
            endpoint=request.endpoint,
            path=request.path,
            remote_addr=request.remote_addr,
            user_agent=request.headers.get('User-Agent', '')[:100]
        )

        # 사용자 정보 (세션에서)
        if session and 'user' in session:
            user_info = session['user']
            self.logger.debug(
                "Authenticated user request",
                request_id=g.request_id,
                user_id=user_info.get('id'),
                user_email=user_info.get('email'),
                user_role=user_info.get('role')
            )

    def _after_request(self, response):
        """요청 완료 시 실행"""
        if hasattr(g, 'start_time'):
            duration = (time.time() - g.start_time) * 1000  # ms 단위

            # 응답 상태에 따른 로그 레벨 결정
            if response.status_code >= 500:
                log_level = 'error'
            elif response.status_code >= 400:
                log_level = 'warning'
            else:
                log_level = 'info'

            # 응답 정보 로깅
            getattr(self.logger, log_level)(
                f"Request completed: {request.method} {request.path} - {response.status_code}",
                request_id=getattr(g, 'request_id', 'unknown'),
                method=request.method,
                endpoint=request.endpoint,
                status_code=response.status_code,
                duration=f"{duration:.2f}",
                content_length=response.content_length
            )

            # 느린 요청 경고 (5초 이상)
            if duration > 5000:
                self.logger.warning(
                    f"Slow request detected: {duration:.2f}ms",
                    request_id=getattr(g, 'request_id', 'unknown'),
                    endpoint=request.endpoint,
                    duration=f"{duration:.2f}"
                )

        return response

    def _teardown_request(self, exception=None):
        """요청 정리 시 실행"""
        if exception:
            self.logger.error(
                f"Request failed with exception: {type(exception).__name__}",
                request_id=getattr(g, 'request_id', 'unknown'),
                exception_type=type(exception).__name__,
                exception_message=str(exception)
            )


def track_business_operation(operation_name: str, **context):
    """비즈니스 작업 추적 데코레이터"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            logger = get_logger(func.__module__)
            start_time = time.time()
            operation_id = str(uuid.uuid4())[:8]

            # 작업 시작 로깅
            logger.info(
                f"Business operation started: {operation_name}",
                operation_id=operation_id,
                operation=operation_name,
                request_id=getattr(g, 'request_id', None),
                **context
            )

            try:
                result = func(*args, **kwargs)
                duration = (time.time() - start_time) * 1000

                # 성공 로깅
                logger.info(
                    f"Business operation completed: {operation_name}",
                    operation_id=operation_id,
                    operation=operation_name,
                    duration=f"{duration:.2f}",
                    status="success",
                    request_id=getattr(g, 'request_id', None)
                )

                # 비즈니스 이벤트 로깅
                log_business_event(
                    operation_name,
                    operation_id=operation_id,
                    duration=duration,
                    status="success",
                    **context
                )

                return result

            except Exception as e:
                duration = (time.time() - start_time) * 1000

                # 실패 로깅
                logger.error(
                    f"Business operation failed: {operation_name}",
                    operation_id=operation_id,
                    operation=operation_name,
                    duration=f"{duration:.2f}",
                    status="error",
                    error_type=type(e).__name__,
                    error_message=str(e),
                    request_id=getattr(g, 'request_id', None)
                )

                # 비즈니스 이벤트 로깅 (실패)
                log_business_event(
                    f"{operation_name}_failed",
                    operation_id=operation_id,
                    duration=duration,
                    status="error",
                    error_type=type(e).__name__,
                    **context
                )

                raise

        return wrapper
    return decorator


def log_database_operation(table_name: str, operation: str):
    """데이터베이스 작업 추적 데코레이터"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            logger = get_logger('database')
            start_time = time.time()

            logger.debug(
                f"Database operation: {operation} on {table_name}",
                table=table_name,
                operation=operation,
                request_id=getattr(g, 'request_id', None)
            )

            try:
                result = func(*args, **kwargs)
                duration = (time.time() - start_time) * 1000

                logger.debug(
                    f"Database operation completed: {operation} on {table_name}",
                    table=table_name,
                    operation=operation,
                    duration=f"{duration:.2f}",
                    status="success",
                    request_id=getattr(g, 'request_id', None)
                )

                return result

            except Exception as e:
                duration = (time.time() - start_time) * 1000

                logger.error(
                    f"Database operation failed: {operation} on {table_name}",
                    table=table_name,
                    operation=operation,
                    duration=f"{duration:.2f}",
                    error_type=type(e).__name__,
                    error_message=str(e),
                    request_id=getattr(g, 'request_id', None)
                )

                raise

        return wrapper
    return decorator


def log_external_api_call(service_name: str, endpoint: str):
    """외부 API 호출 추적 데코레이터"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            logger = get_logger('external_api')
            start_time = time.time()

            logger.info(
                f"External API call: {service_name} {endpoint}",
                service=service_name,
                endpoint=endpoint,
                request_id=getattr(g, 'request_id', None)
            )

            try:
                result = func(*args, **kwargs)
                duration = (time.time() - start_time) * 1000

                logger.info(
                    f"External API call completed: {service_name} {endpoint}",
                    service=service_name,
                    endpoint=endpoint,
                    duration=f"{duration:.2f}",
                    status="success",
                    request_id=getattr(g, 'request_id', None)
                )

                return result

            except Exception as e:
                duration = (time.time() - start_time) * 1000

                logger.error(
                    f"External API call failed: {service_name} {endpoint}",
                    service=service_name,
                    endpoint=endpoint,
                    duration=f"{duration:.2f}",
                    error_type=type(e).__name__,
                    error_message=str(e),
                    request_id=getattr(g, 'request_id', None)
                )

                raise

        return wrapper
    return decorator


# 전역 인스턴스
request_tracker = RequestTracker()