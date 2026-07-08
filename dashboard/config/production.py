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

    # 서비스워커 사용 안 함 (2026-07-08).
    # 이유: 기존 sw.js의 Cache First 전략이 Vite 해시 파일 변경 시 stale HTML을
    # 서빙해 매니저 브라우저에서 ERR_FAILED / "이 페이지에 연결할 수 없습니다" 발생.
    # sw.js 자체는 unregister 스텁으로 유지 — 기존 브라우저에 남은 등록분을 자동 해제.
    SERVICE_WORKER_ENABLED = False
    SERVICE_WORKER_CACHE_TTL = 86400 * 7  # 7일

    # API 모니터링 (프로덕션 - 안정성 중심)
    API_ERROR_RATE_THRESHOLD = 0.15  # 15%
    API_SLOW_RESPONSE_THRESHOLD = 15.0  # 15초
    API_ALERT_COOLDOWN_MINUTES = 10  # 10분

    # 파일 업로드 제한 (프로덕션 보안)
    MAX_CONTENT_LENGTH = 8 * 1024 * 1024  # 8MB로 제한

    # 정적 자산 캐시 (2026-07-08): Vite hash filename(예: components.tJkYWzVd.js)이라 콘텐츠 변경 시 자동 무효화.
    # 1년 immutable 캐시로 브라우저 재요청 대폭 감소. Flask send_file 응답의 Cache-Control 헤더에 적용.
    SEND_FILE_MAX_AGE_DEFAULT = 60 * 60 * 24 * 365  # 1년

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