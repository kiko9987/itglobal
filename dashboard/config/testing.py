"""
Testing Configuration
테스트 환경 전용 설정
"""

import tempfile
import os
from .base import BaseConfig

class TestingConfig(BaseConfig):
    """테스트 환경 설정"""

    # Flask 테스트 설정
    DEBUG = False
    TESTING = True
    STRICT_ENV_VALIDATION = False

    # 테스트용 보안 완화
    WTF_CSRF_ENABLED = False
    SESSION_COOKIE_SECURE = False

    # 테스트 로깅
    LOG_LEVEL = 'DEBUG'

    # 테스트용 임시 데이터베이스
    DATABASE_PATH = os.path.join(tempfile.gettempdir(), 'test_projects.db')

    # Vite 테스트 설정
    VITE_DEV_SERVER_ENABLED = False

    # API 제한 완화 (테스트용)
    GOOGLE_SHEETS_READ_LIMIT = 9999
    GOOGLE_SHEETS_WRITE_LIMIT = 9999

    # 캐시 비활성화 (테스트 일관성)
    CACHE_TTL_MINUTES = 0  # 캐시 비활성화
    BACKGROUND_PREFETCH_ENABLED = False

    # 서비스워커 비활성화 (테스트 격리)
    SERVICE_WORKER_ENABLED = False

    # API 모니터링 비활성화 (테스트 노이즈 제거)
    API_MONITORING_ENABLED = False

    # 임시 업로드 폴더
    UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'test_uploads')

    # 테스트 API 키 (더미)
    GOOGLE_SHEETS_API_KEY = 'test-api-key'
    GOOGLE_CLIENT_ID = 'test-client-id'
    GOOGLE_CLIENT_SECRET = 'test-client-secret'

    @classmethod
    def init_app(cls, app):
        """테스트 환경 초기화"""
        # 부모 클래스 초기화 스킵 (임시 디렉토리 사용)
        os.makedirs(cls.UPLOAD_FOLDER, exist_ok=True)

        # 테스트 환경 전용 초기화
        import logging
        logging.disable(logging.CRITICAL)  # 테스트 중 로그 비활성화
