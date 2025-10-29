"""
API 표준 응답 구조

모든 API 엔드포인트가 일관된 응답 구조를 반환하도록 헬퍼 함수 제공.
API_STANDARDS.md의 표준을 준수합니다.

Usage:
    from dashboard.utils.api_response import ApiResponse, ErrorCode

    # 성공 응답
    return ApiResponse.success(data=projects, meta={'total': 100})

    # 에러 응답
    return ApiResponse.error(
        ErrorCode.VALIDATION_ERROR,
        "프로젝트 코드가 필요합니다",
        details={'field': 'project_code'}
    )
"""

from typing import Any, Dict, Optional, Union
from datetime import datetime
from flask import jsonify
from enum import Enum
import uuid


class ErrorCode(Enum):
    """표준 에러 코드"""

    # 4xx: 클라이언트 에러
    VALIDATION_ERROR = "VALIDATION_ERROR"
    MISSING_REQUIRED_FIELD = "MISSING_REQUIRED_FIELD"
    INVALID_FORMAT = "INVALID_FORMAT"
    RESOURCE_NOT_FOUND = "RESOURCE_NOT_FOUND"
    ALREADY_EXISTS = "ALREADY_EXISTS"
    PERMISSION_DENIED = "PERMISSION_DENIED"
    AUTHENTICATION_REQUIRED = "AUTHENTICATION_REQUIRED"

    # 5xx: 서버 에러
    INTERNAL_ERROR = "INTERNAL_ERROR"
    DATABASE_ERROR = "DATABASE_ERROR"
    EXTERNAL_SERVICE_ERROR = "EXTERNAL_SERVICE_ERROR"
    SHEETS_API_ERROR = "SHEETS_API_ERROR"
    CALENDAR_API_ERROR = "CALENDAR_API_ERROR"

    # 특수
    PARTIAL_SUCCESS = "PARTIAL_SUCCESS"


class ApiResponse:
    """API 표준 응답 생성기"""

    @staticmethod
    def success(
        data: Any = None,
        message: Optional[str] = None,
        meta: Optional[Dict[str, Any]] = None,
        status_code: int = 200
    ):
        """성공 응답 생성"""
        response_data = {'success': True}

        if message:
            response_data['message'] = message

        if data is not None:
            response_data['data'] = data

        if meta is None:
            meta = {}

        if 'timestamp' not in meta:
            meta['timestamp'] = datetime.now().isoformat()

        if 'request_id' not in meta:
            meta['request_id'] = str(uuid.uuid4())[:8]

        response_data['meta'] = meta

        return jsonify(response_data), status_code

    @staticmethod
    def error(
        error_code: Union[ErrorCode, str],
        message: str,
        details: Optional[Dict[str, Any]] = None,
        status_code: Optional[int] = None
    ):
        """에러 응답 생성"""
        if isinstance(error_code, ErrorCode):
            code_str = error_code.value
        else:
            code_str = error_code

        if status_code is None:
            status_code = ApiResponse._get_http_status_from_error_code(error_code)

        response_data = {
            'success': False,
            'error': {
                'code': code_str,
                'message': message
            },
            'meta': {
                'timestamp': datetime.now().isoformat(),
                'request_id': str(uuid.uuid4())[:8]
            }
        }

        if details:
            response_data['error']['details'] = details

        return jsonify(response_data), status_code

    @staticmethod
    def partial_success(
        results: list,
        success_count: int,
        failed_count: int,
        total_count: int,
        message: Optional[str] = None
    ):
        """부분 성공 응답 (207 Multi-Status)"""
        if message is None:
            message = f"{total_count}개 중 {success_count}개 성공"
            if failed_count > 0:
                message += f", {failed_count}개 실패"

        return ApiResponse.success(
            data={
                'results': results,
                'summary': {
                    'total_count': total_count,
                    'success_count': success_count,
                    'failed_count': failed_count
                }
            },
            message=message,
            status_code=207
        )

    @staticmethod
    def _get_http_status_from_error_code(error_code: Union[ErrorCode, str]) -> int:
        """에러 코드에서 HTTP 상태 코드 자동 판단"""
        if isinstance(error_code, ErrorCode):
            code_str = error_code.value
        else:
            code_str = error_code

        mapping = {
            'VALIDATION_ERROR': 400,
            'MISSING_REQUIRED_FIELD': 400,
            'INVALID_FORMAT': 400,
            'ALREADY_EXISTS': 400,
            'AUTHENTICATION_REQUIRED': 401,
            'PERMISSION_DENIED': 403,
            'RESOURCE_NOT_FOUND': 404,
            'INTERNAL_ERROR': 500,
            'DATABASE_ERROR': 500,
            'EXTERNAL_SERVICE_ERROR': 503,
            'SHEETS_API_ERROR': 503,
            'CALENDAR_API_ERROR': 503,
            'PARTIAL_SUCCESS': 207
        }

        return mapping.get(code_str, 500)


# ===== 하위 호환성 래퍼 (Backward Compatibility) =====
# 기존 코드에서 사용 중인 레거시 이름을 지원합니다.
# 점진적으로 새 이름(ApiResponse, ErrorCode)으로 마이그레이션하되,
# 기존 임포트가 깨지지 않도록 보장합니다.

# 레거시 클래스 alias
APIResponse = ApiResponse
APIErrorCode = ErrorCode


def api_response(
    success: bool = True,
    data: Any = None,
    message: Optional[str] = None,
    error: Optional[str] = None,
    status_code: int = 200
):
    """
    레거시 api_response 함수 래퍼 (하위 호환성)

    기존 코드에서 사용 중인 api_response() 함수 호출을 지원합니다.
    새 코드에서는 ApiResponse.success() 또는 ApiResponse.error()를 직접 사용하세요.

    Args:
        success: 성공 여부
        data: 응답 데이터
        message: 메시지
        error: 에러 메시지 (success=False일 때)
        status_code: HTTP 상태 코드

    Returns:
        Flask jsonify response, status_code
    """
    if success:
        return ApiResponse.success(
            data=data,
            message=message,
            status_code=status_code
        )
    else:
        return ApiResponse.error(
            ErrorCode.INTERNAL_ERROR,
            message=error or message or "An error occurred",
            status_code=status_code
        )
