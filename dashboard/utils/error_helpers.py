"""
에러 처리 헬퍼 함수

보안:
- 에러 메시지 노출 차단 (CWE-209)
- 에러 ID 기반 추적 시스템
"""

import uuid
import logging

logger = logging.getLogger(__name__)


def generate_error_id() -> str:
    """
    에러 추적용 고유 ID 생성

    Returns:
        str: 8자리 UUID (예: "a3f5c2b1")

    Note:
        - 서버 로그에는 상세 에러 정보 저장
        - 클라이언트에는 에러 ID만 반환
        - 사용자가 에러 ID를 제공하면 로그에서 추적 가능
    """
    return str(uuid.uuid4())[:8]


def create_error_response(error: Exception, user_message: str, error_id: str = None,
                         status_code: int = 500) -> dict:
    """
    표준화된 에러 응답 생성

    Args:
        error: 발생한 예외
        user_message: 사용자에게 표시할 친화적 메시지
        error_id: 에러 ID (없으면 자동 생성)
        status_code: HTTP 상태 코드

    Returns:
        dict: {'success': False, 'error': user_message, 'error_id': error_id}

    Example:
        try:
            # 작업 수행
            pass
        except Exception as e:
            error_id = generate_error_id()
            logger.error(f"[{error_id}] 상세 오류: {str(e)}", exc_info=True)
            return jsonify(create_error_response(
                e,
                "작업을 수행할 수 없습니다",
                error_id
            )), 500
    """
    if error_id is None:
        error_id = generate_error_id()

    return {
        'success': False,
        'error': user_message,
        'error_id': error_id
    }


def log_and_create_error_response(error: Exception, user_message: str,
                                  log_message: str = None,
                                  status_code: int = 500) -> tuple:
    """
    에러 로깅 및 응답 생성 (원스톱 헬퍼)

    Args:
        error: 발생한 예외
        user_message: 사용자에게 표시할 친화적 메시지
        log_message: 로그에 남길 메시지 (없으면 user_message 사용)
        status_code: HTTP 상태 코드

    Returns:
        tuple: (response_dict, status_code)

    Example:
        try:
            # 작업 수행
            pass
        except Exception as e:
            return jsonify(*log_and_create_error_response(
                e,
                "프로젝트를 업데이트할 수 없습니다",
                f"프로젝트 업데이트 오류: {project_code}"
            ))
    """
    error_id = generate_error_id()

    # 로그 메시지 결정
    if log_message is None:
        log_message = user_message

    # 상세 에러 로깅 (서버측)
    logger.error(f"[{error_id}] {log_message}: {str(error)}", exc_info=True)

    # 클라이언트 응답 생성
    response = create_error_response(error, user_message, error_id, status_code)

    return response, status_code
