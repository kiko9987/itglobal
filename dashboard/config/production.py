"""
Production Configuration
프로덕션 환경 전용 설정
"""

from .base import BaseConfig

class ProductionConfig(BaseConfig):
    """프로덕션 환경 설정"""

    # Flask 프로덕션 설정
    DEBUG = False
    TESTING = False

    # 보안 강화
    WTF_CSRF_ENABLED = True
    SESSION_COOKIE_SECURE = True  # HTTPS 필수
    SESSION_COOKIE_HTTPONLY = True

    # 프로덕션 로깅
    LOG_LEVEL = 'WARNING'

    # Vite 프로덕션 설정
    VITE_DEV_SERVER_ENABLED = False

    # API 제한 (프로덕션 - 엄격)
    GOOGLE_SHEETS_READ_LIMIT = 80   # 여유분 확보
    GOOGLE_SHEETS_WRITE_LIMIT = 80

    # 캐시 설정 (프로덕션 - 안정성)
    # 2026-07-07: 10분 → 20분 확장 (20명 동시 사용 대비 Google Sheets API 호출 감소).
    # 편집·취소·재개 시 즉시 부분 무효화로 사용자 피드백 지연 없음.
    # 백그라운드 프리페치가 주기적으로 갱신하니 stale window는 실질 40초 이내.
    CACHE_TTL_MINUTES = 20
    BACKGROUND_PREFETCH_ENABLED = True

    # 서비스워커 활성화 (프로덕션 성능)
    SERVICE_WORKER_ENABLED = True
    SERVICE_WORKER_CACHE_TTL = 86400 * 7  # 7일

    # API 모니터링 (프로덕션 - 안정성 중심)
    API_ERROR_RATE_THRESHOLD = 0.15  # 15%
    API_SLOW_RESPONSE_THRESHOLD = 15.0  # 15초
    API_ALERT_COOLDOWN_MINUTES = 10  # 10분

    # 파일 업로드 제한 (프로덕션 보안)
    MAX_CONTENT_LENGTH = 8 * 1024 * 1024  # 8MB로 제한

    @classmethod
    def init_app(cls, app):
        """프로덕션 환경 초기화"""
        super().init_app(app)

        # 프로덕션 환경 전용 초기화
        import logging
        from logging.handlers import RotatingFileHandler

        # 로그 파일 로테이션 설정
        if not app.debug and not app.testing:
            file_handler = RotatingFileHandler(
                f'{cls.LOG_DIR}/dashboard.log',
                maxBytes=10485760,  # 10MB
                backupCount=10
            )
            file_handler.setFormatter(logging.Formatter(
                '%(asctime)s %(levelname)s: %(message)s '
                '[in %(pathname)s:%(lineno)d]'
            ))
            file_handler.setLevel(logging.WARNING)
            app.logger.addHandler(file_handler)