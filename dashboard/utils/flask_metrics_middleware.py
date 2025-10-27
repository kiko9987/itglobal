"""
Flask 메트릭 수집 미들웨어
HTTP 요청, 응답, 성능 지표 자동 수집
"""

import time
from flask import request, g, session
from functools import wraps
from typing import Optional, Dict, Any

from .metrics_collector import get_metrics_collector, MetricType, time_block
from .logging_config import get_logger

# 성능 임계값 상수 (밀리초)
SLOW_REQUEST_THRESHOLD_MS = 5000  # 5초 - 매우 느린 요청
MODERATE_SLOW_REQUEST_THRESHOLD_MS = 1000  # 1초 - 느린 요청


class FlaskMetricsMiddleware:
    """Flask 애플리케이션용 메트릭 수집 미들웨어"""

    def __init__(self, app=None):
        self.collector = get_metrics_collector()
        self.logger = get_logger(__name__)

        if app is not None:
            self.init_app(app)

    def init_app(self, app):
        """Flask 앱에 메트릭 미들웨어 등록"""
        app.before_request(self._before_request)
        app.after_request(self._after_request)
        app.teardown_appcontext(self._teardown_request)

        # 앱 시작 메트릭
        self.collector.increment_counter("app.startup", tags={"app": app.name})

    def _before_request(self):
        """요청 시작 시 메트릭 수집"""
        g.request_start_time = time.time()

        # 요청 카운터
        self.collector.increment_counter(
            "http.requests.total",
            tags={
                "method": request.method,
                "endpoint": request.endpoint or "unknown"
            }
        )

        # 동시 요청 수 (게이지)
        if not hasattr(g, 'concurrent_requests'):
            g.concurrent_requests = 0
        g.concurrent_requests += 1

        self.collector.set_gauge(
            "http.requests.concurrent",
            g.concurrent_requests
        )

        # 사용자별 요청 추적
        if session and 'user' in session:
            user_info = session['user']
            self.collector.increment_counter(
                "http.requests.by_user",
                tags={
                    "user_id": str(user_info.get('id', 'unknown')),
                    "user_email": user_info.get('email', 'unknown')
                }
            )

    def _after_request(self, response):
        """요청 완료 시 메트릭 수집"""
        if hasattr(g, 'request_start_time'):
            # 응답 시간 측정
            duration = (time.time() - g.request_start_time) * 1000  # ms

            # 응답 시간 히스토그램
            self.collector.record_histogram(
                "http.request.duration",
                duration,
                tags={
                    "method": request.method,
                    "endpoint": request.endpoint or "unknown",
                    "status_code": str(response.status_code)
                },
                unit="ms"
            )

            # 상태 코드별 카운터
            self.collector.increment_counter(
                "http.responses.total",
                tags={
                    "method": request.method,
                    "endpoint": request.endpoint or "unknown",
                    "status_code": str(response.status_code),
                    "status_class": f"{response.status_code // 100}xx"
                }
            )

            # 응답 크기
            if response.content_length:
                self.collector.record_histogram(
                    "http.response.size",
                    response.content_length,
                    tags={
                        "method": request.method,
                        "endpoint": request.endpoint or "unknown"
                    },
                    unit="bytes"
                )

            # 느린 요청 경고 메트릭
            if duration > SLOW_REQUEST_THRESHOLD_MS:
                self.collector.increment_counter(
                    "http.requests.slow",
                    tags={
                        "method": request.method,
                        "endpoint": request.endpoint or "unknown",
                        "duration_bucket": "5s+"
                    }
                )
            elif duration > MODERATE_SLOW_REQUEST_THRESHOLD_MS:
                self.collector.increment_counter(
                    "http.requests.slow",
                    tags={
                        "method": request.method,
                        "endpoint": request.endpoint or "unknown",
                        "duration_bucket": "1s+"
                    }
                )

        # 동시 요청 수 감소
        if hasattr(g, 'concurrent_requests'):
            g.concurrent_requests -= 1
            self.collector.set_gauge(
                "http.requests.concurrent",
                g.concurrent_requests
            )

        return response

    def _teardown_request(self, exception=None):
        """요청 정리 시 메트릭 수집"""
        if exception:
            # request context가 있는 경우에만 메트릭 수집
            try:
                from flask import request, has_request_context
                if has_request_context():
                    # 예외 발생 메트릭
                    self.collector.increment_counter(
                        "http.requests.exceptions",
                        tags={
                            "method": request.method,
                            "endpoint": request.endpoint or "unknown",
                            "exception_type": type(exception).__name__
                        }
                    )
            except RuntimeError:
                # request context가 없는 경우 (예: SocketIO disconnect)
                pass


def track_business_metrics(operation: str, **tags):
    """비즈니스 메트릭 추적 데코레이터"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            collector = get_metrics_collector()

            # 작업 시작 메트릭
            start_time = time.time()
            collector.increment_counter(
                f"business.{operation}.started",
                tags=tags
            )

            try:
                result = func(*args, **kwargs)

                # 성공 메트릭
                duration = (time.time() - start_time) * 1000
                collector.increment_counter(
                    f"business.{operation}.completed",
                    tags={**tags, "status": "success"}
                )
                collector.record_timer(
                    f"business.{operation}.duration",
                    duration,
                    tags={**tags, "status": "success"}
                )

                return result

            except Exception as e:
                # 실패 메트릭
                duration = (time.time() - start_time) * 1000
                collector.increment_counter(
                    f"business.{operation}.failed",
                    tags={**tags, "status": "error", "error_type": type(e).__name__}
                )
                collector.record_timer(
                    f"business.{operation}.duration",
                    duration,
                    tags={**tags, "status": "error", "error_type": type(e).__name__}
                )
                raise

        return wrapper
    return decorator


def track_database_metrics(table: str, operation: str):
    """데이터베이스 작업 메트릭 추적"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            collector = get_metrics_collector()

            start_time = time.time()

            # 데이터베이스 작업 시작
            collector.increment_counter(
                "database.operations.total",
                tags={
                    "table": table,
                    "operation": operation
                }
            )

            try:
                result = func(*args, **kwargs)

                # 성공 메트릭
                duration = (time.time() - start_time) * 1000
                collector.record_timer(
                    "database.operation.duration",
                    duration,
                    tags={
                        "table": table,
                        "operation": operation,
                        "status": "success"
                    }
                )

                # 결과가 리스트인 경우 행 수 기록
                if isinstance(result, list):
                    collector.record_histogram(
                        "database.query.rows",
                        len(result),
                        tags={
                            "table": table,
                            "operation": operation
                        }
                    )

                return result

            except Exception as e:
                # 실패 메트릭
                duration = (time.time() - start_time) * 1000
                collector.increment_counter(
                    "database.operations.errors",
                    tags={
                        "table": table,
                        "operation": operation,
                        "error_type": type(e).__name__
                    }
                )
                collector.record_timer(
                    "database.operation.duration",
                    duration,
                    tags={
                        "table": table,
                        "operation": operation,
                        "status": "error",
                        "error_type": type(e).__name__
                    }
                )
                raise

        return wrapper
    return decorator


def track_cache_metrics(cache_name: str):
    """캐시 작업 메트릭 추적"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            collector = get_metrics_collector()

            # 캐시 작업 타입 추정 (함수명 기반)
            func_name = func.__name__.lower()
            if 'get' in func_name:
                operation = 'get'
            elif 'set' in func_name:
                operation = 'set'
            elif 'delete' in func_name or 'invalidate' in func_name:
                operation = 'delete'
            else:
                operation = 'unknown'

            start_time = time.time()

            try:
                result = func(*args, **kwargs)

                # 성공 메트릭
                duration = (time.time() - start_time) * 1000
                collector.record_timer(
                    "cache.operation.duration",
                    duration,
                    tags={
                        "cache": cache_name,
                        "operation": operation,
                        "status": "success"
                    }
                )

                # 캐시 히트/미스 추적 (get 작업의 경우)
                if operation == 'get':
                    if result is not None:
                        collector.increment_counter(
                            "cache.hits",
                            tags={"cache": cache_name}
                        )
                    else:
                        collector.increment_counter(
                            "cache.misses",
                            tags={"cache": cache_name}
                        )

                collector.increment_counter(
                    "cache.operations.total",
                    tags={
                        "cache": cache_name,
                        "operation": operation,
                        "status": "success"
                    }
                )

                return result

            except Exception as e:
                # 실패 메트릭
                duration = (time.time() - start_time) * 1000
                collector.increment_counter(
                    "cache.operations.errors",
                    tags={
                        "cache": cache_name,
                        "operation": operation,
                        "error_type": type(e).__name__
                    }
                )
                collector.record_timer(
                    "cache.operation.duration",
                    duration,
                    tags={
                        "cache": cache_name,
                        "operation": operation,
                        "status": "error",
                        "error_type": type(e).__name__
                    }
                )
                raise

        return wrapper
    return decorator


def track_external_api_metrics(service: str, endpoint: str = None):
    """외부 API 호출 메트릭 추적"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            collector = get_metrics_collector()

            start_time = time.time()
            api_endpoint = endpoint or func.__name__

            # API 호출 시작
            collector.increment_counter(
                "external_api.calls.total",
                tags={
                    "service": service,
                    "endpoint": api_endpoint
                }
            )

            try:
                result = func(*args, **kwargs)

                # 성공 메트릭
                duration = (time.time() - start_time) * 1000
                collector.record_timer(
                    "external_api.call.duration",
                    duration,
                    tags={
                        "service": service,
                        "endpoint": api_endpoint,
                        "status": "success"
                    }
                )

                collector.increment_counter(
                    "external_api.calls.success",
                    tags={
                        "service": service,
                        "endpoint": api_endpoint
                    }
                )

                return result

            except Exception as e:
                # 실패 메트릭
                duration = (time.time() - start_time) * 1000
                collector.increment_counter(
                    "external_api.calls.errors",
                    tags={
                        "service": service,
                        "endpoint": api_endpoint,
                        "error_type": type(e).__name__
                    }
                )
                collector.record_timer(
                    "external_api.call.duration",
                    duration,
                    tags={
                        "service": service,
                        "endpoint": api_endpoint,
                        "status": "error",
                        "error_type": type(e).__name__
                    }
                )
                raise

        return wrapper
    return decorator


# 편의 함수들
def record_user_action(action: str, user_id: str = None, **context):
    """사용자 액션 메트릭 기록"""
    collector = get_metrics_collector()

    tags = {"action": action}
    if user_id:
        tags["user_id"] = user_id

    # 컨텍스트 정보 추가
    for key, value in context.items():
        if isinstance(value, (str, int, float)):
            tags[key] = str(value)

    collector.increment_counter("user.actions", tags=tags)


def record_business_event(event_type: str, **context):
    """비즈니스 이벤트 메트릭 기록"""
    collector = get_metrics_collector()

    tags = {"event_type": event_type}

    # 컨텍스트 정보 추가
    for key, value in context.items():
        if isinstance(value, (str, int, float)):
            tags[key] = str(value)

    collector.increment_counter("business.events", tags=tags)


def record_performance_metric(name: str, value: float, unit: str = "ms", **tags):
    """성능 메트릭 기록"""
    collector = get_metrics_collector()
    collector.record_histogram(f"performance.{name}", value, tags=tags, unit=unit)


# 전역 미들웨어 인스턴스
flask_metrics = FlaskMetricsMiddleware()