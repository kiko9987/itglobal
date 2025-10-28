"""
커스텀 예외 클래스 정의
"""


class ServiceUnavailable(Exception):
    """
    서비스를 사용할 수 없을 때 발생하는 예외

    주로 Redis, 외부 API 등의 장애 시 사용
    HTTP 503 Service Unavailable 응답으로 매핑됨
    """
    def __init__(self, message="Service temporarily unavailable"):
        self.message = message
        super().__init__(self.message)


class LockAcquisitionError(Exception):
    """
    락 획득 실패 시 발생하는 예외
    """
    def __init__(self, message="Failed to acquire lock"):
        self.message = message
        super().__init__(self.message)


class CacheError(Exception):
    """
    캐시 작업 실패 시 발생하는 예외
    """
    def __init__(self, message="Cache operation failed"):
        self.message = message
        super().__init__(self.message)
