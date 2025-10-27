"""
Development Configuration
개발 환경 전용 설정
"""

from .base import BaseConfig

class DevelopmentConfig(BaseConfig):
    """개발 환경 설정"""

    # Flask 개발 설정
    DEBUG = True
    TESTING = False
    STRICT_ENV_VALIDATION = True  # 프로덕션과 동일하게 strict 검증 (Production Parity)

    # 개발용 보안 완화
    WTF_CSRF_ENABLED = False  # 개발 시 편의를 위해 비활성화
    SESSION_COOKIE_SECURE = False

    # 개발 환경 로깅
    LOG_LEVEL = 'DEBUG'

    # 개발 서버 설정
    DEVELOPMENT_SERVER = True
    VITE_DEV_SERVER_ENABLED = False  # 임시로 빌드 파일 사용 (WebSocket 오류 디버깅)
    VITE_DEV_SERVER_HOST = 'localhost'
    VITE_DEV_SERVER_PORT = 5173

    # API 제한 완화 (개발용)
    GOOGLE_SHEETS_READ_LIMIT = 1000
    GOOGLE_SHEETS_WRITE_LIMIT = 1000

    # 캐시 설정 (개발용 - 빠른 테스트)
    CACHE_TTL_MINUTES = 1  # 1분으로 단축
    BACKGROUND_PREFETCH_ENABLED = True

    # 서비스워커 비활성화 (개발 시 캐시 문제 방지)
    SERVICE_WORKER_ENABLED = False

    # API 모니터링 더 민감하게 (개발용)
    API_ERROR_RATE_THRESHOLD = 0.05  # 5%
    API_SLOW_RESPONSE_THRESHOLD = 5.0  # 5초
    API_ALERT_COOLDOWN_MINUTES = 1  # 1분

    @classmethod
    def init_app(cls, app):
        """개발 환경 초기화"""
        super().init_app(app)

        # 개발 환경 전용 초기화
        import logging
        logging.getLogger('werkzeug').setLevel(logging.WARNING)
        logging.getLogger('urllib3').setLevel(logging.WARNING)
