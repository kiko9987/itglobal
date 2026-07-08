"""
Base Configuration
모든 환경에서 공통으로 사용되는 기본 설정
"""

import os
from datetime import timedelta

class BaseConfig:
    """기본 설정 클래스 - 모든 환경 공통"""

    # Flask 기본 설정
    SECRET_KEY = os.environ.get('SECRET_KEY', 'dev-secret-key-change-in-production')
    STRICT_ENV_VALIDATION = True
    WTF_CSRF_ENABLED = True

    # 세션 설정
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = 'Lax'
    PERMANENT_SESSION_LIFETIME = timedelta(hours=24)

    # Redis 세션 설정
    SESSION_TYPE = 'redis'
    SESSION_PERMANENT = True
    SESSION_USE_SIGNER = True  # 세션 쿠키 서명
    SESSION_KEY_PREFIX = 'itg_session:'  # Redis 키 prefix
    SESSION_REDIS = None  # 런타임에 설정됨 (__init__.py에서)

    # 보안 설정
    SESSION_COOKIE_SECURE = False  # HTTPS에서만 True로 설정

    # 데이터베이스 설정
    DATABASE_PATH = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        'instance',
        'projects.db'
    )

    # Google API 설정
    GOOGLE_SHEETS_API_KEY = os.environ.get('GOOGLE_SHEETS_API_KEY', '')
    GOOGLE_CLIENT_ID = os.environ.get('GOOGLE_CLIENT_ID', '')
    GOOGLE_CLIENT_SECRET = os.environ.get('GOOGLE_CLIENT_SECRET', '')

    # Google OAuth 설정
    GOOGLE_OAUTH_ALLOWED_DOMAIN = os.environ.get('GOOGLE_OAUTH_ALLOWED_DOMAIN', 'itg-aircon.com')
    GOOGLE_OAUTH_REDIRECT_URIS = os.environ.get(
        'GOOGLE_OAUTH_REDIRECT_URIS',
        'http://localhost:5000/auth/callback,http://localhost:5001/auth/callback,http://localhost:5004/auth/callback'
    ).split(',')
    GOOGLE_OAUTH_DEFAULT_REDIRECT_URI = os.environ.get(
        'GOOGLE_OAUTH_DEFAULT_REDIRECT_URI',
        'http://localhost:5004/auth/callback'
    )

    # API 제한 설정 (ver10에서 제안된 중앙 관리)
    GOOGLE_SHEETS_READ_LIMIT = 100  # 분당 읽기 요청 수
    GOOGLE_SHEETS_WRITE_LIMIT = 100  # 분당 쓰기 요청 수

    # 캐시 설정
    CACHE_TTL_MINUTES = 5  # 캐시 TTL (분)
    BACKGROUND_PREFETCH_ENABLED = True

    # Vite 설정
    VITE_DEV_SERVER_URL = 'http://localhost:5173'
    VITE_MANIFEST_PATH = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        'static',
        'dist',
        '.vite',
        'manifest.json'
    )

    # 로깅 설정
    LOG_LEVEL = 'INFO'
    LOG_DIR = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        'logs'
    )

    # 서비스워커 설정 — 2026-07-08 무력화 (production.py 주석 참조).
    SERVICE_WORKER_ENABLED = False
    SERVICE_WORKER_CACHE_TTL = 86400  # 24시간 (초)

    # 업로드 설정
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB
    UPLOAD_FOLDER = os.path.join(
        os.path.dirname(os.path.dirname(__file__)),
        'uploads'
    )

    # API 모니터링 설정
    API_MONITORING_ENABLED = True
    API_ALERT_COOLDOWN_MINUTES = 5
    API_ERROR_RATE_THRESHOLD = 0.1  # 10%
    API_SLOW_RESPONSE_THRESHOLD = 10.0  # 10초

    @classmethod
    def init_app(cls, app):
        """Flask 앱 초기화 시 실행되는 메서드"""
        # 필요한 디렉토리 생성
        os.makedirs(cls.LOG_DIR, exist_ok=True)
        os.makedirs(cls.UPLOAD_FOLDER, exist_ok=True)
        os.makedirs(os.path.dirname(cls.DATABASE_PATH), exist_ok=True)

        # 추가 초기화 로직
        pass
