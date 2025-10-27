"""
Flask Dashboard Application Factory
모든 환경에서 사용되는 중앙 애플리케이션 팩토리
"""

import os
import logging
from flask import Flask
from flask_compress import Compress

# 설정 시스템 가져오기
from dashboard.config import get_config

# 유틸리티 및 서비스 초기화
from dashboard.utils.logging_config import setup_logging, get_logger

# 환경 검증 및 초기화
from dashboard.utils.env_validator import validate_dashboard_env
from dashboard.utils.frontend_helpers import init_frontend_helpers
from dashboard.utils.background_prefetch import init_background_prefetch
from dashboard.utils.api_usage_monitor import setup_api_alerts

logger = get_logger(__name__)

# SocketIO 설정 상수
SOCKETIO_MAX_HTTP_BUFFER_SIZE = 10 * 1024 * 1024  # 10MB - 대용량 데이터 처리
SOCKETIO_PING_TIMEOUT = 10  # 10초 후 응답 없으면 연결 종료
SOCKETIO_PING_INTERVAL = 5  # 5초마다 ping 체크

def create_app(config_name=None, config_overrides=None, enable_socketio=True):
    """
    Flask 앱 팩토리 함수

    다양한 환경(run_server.py, 테스트, WSGI 서버)에서
    앱 인스턴스를 생성할 수 있도록 표준화된 팩토리

    Args:
        config_name: 환경 설정 이름 ('development', 'production', 'testing')
        config_overrides: 추가 설정 딕셔너리
        enable_socketio: SocketIO 사용 여부 (기본값: True)

    Returns:
        tuple: (Flask 앱 인스턴스, SocketIO 인스턴스 또는 None)
    """

    # Flask 앱 인스턴스 생성
    app = Flask(__name__)

    # 1. 로깅 시스템 초기화
    setup_logging()
    logger.info("Flask 앱 팩토리 초기화 시작")

    # 2. 환경 변수 검증 및 설정 로딩
    config_class = get_config(config_name)
    strict_env_validation = getattr(config_class, 'STRICT_ENV_VALIDATION', True)

    try:
        env_result = validate_dashboard_env(strict=strict_env_validation)

        if env_result.get('warnings'):
            for warning in env_result['warnings']:
                logger.warning(f"[ENV WARNING] {warning}")

        if strict_env_validation and not env_result.get('valid', True):
            missing = ", ".join(env_result.get('errors', []))
            raise RuntimeError(f"필수 환경변수가 설정되지 않았습니다: {missing}")

        logger.info("환경변수 검증 완료")
    except Exception as e:
        logger.error(f"환경변수 검증 실패: {e}")
        raise

    # 3. 환경별 설정 적용
    app.config.from_object(config_class)
    config_class.init_app(app)

    logger.info(f"환경 설정 적용: {config_class.__name__}")

    # 4. 추가 설정 오버라이드 적용
    if config_overrides:
        app.config.update(config_overrides)
        logger.info(f"추가 설정 오버라이드 적용: {list(config_overrides.keys())}")

    # 4-1. Flask-Compress 초기화 (gzip 압축)
    try:
        compress = Compress(app)
        app.config['COMPRESS_MIMETYPES'] = [
            'application/json',
            'text/html',
            'text/css',
            'text/javascript',
            'application/javascript'
        ]
        app.config['COMPRESS_MIN_SIZE'] = 500  # 500바이트 이상만 압축
        app.config['COMPRESS_LEVEL'] = 6  # 압축 레벨 (1-9, 기본값 6)
        logger.info("Flask-Compress 초기화 완료 (JSON/HTML/CSS/JS gzip 압축 활성화)")
    except Exception as e:
        logger.error(f"Flask-Compress 초기화 실패: {e}")
        logger.warning("압축 없이 계속 진행")

    # 5. 데이터베이스 초기화 (통합)
    # 사용자 DB와 감사 로그가 instance/users.db로 통합됨
    # dashboard/utils/user_database.py에서 자동 초기화됨
    try:
        from dashboard.utils.user_database import get_user_database
        get_user_database()  # 싱글톤 초기화 (테이블 생성 포함)
        logger.info("통합 데이터베이스 시스템 초기화 완료 (instance/users.db)")
    except Exception as e:
        logger.error(f"데이터베이스 초기화 실패: {e}")
        raise

    # 6. 프론트엔드 헬퍼 서비스 초기화
    try:
        init_frontend_helpers(app)
        logger.info("프론트엔드 헬퍼 시스템 초기화 완료")
    except Exception as e:
        logger.error(f"프론트엔드 헬퍼 초기화 실패: {e}")
        raise

    # 7. API 사용량 모니터 시스템
    try:
        from dashboard.utils.api_usage_monitor import setup_api_monitoring

        # API 모니터링 시스템 설정 (스케줄러 포함)
        setup_api_monitoring(
            auto_start_scheduler=app.config.get('API_MONITORING_ENABLED', True),
            check_interval=app.config.get('API_CHECK_INTERVAL', 300),    # 5분
            report_interval=app.config.get('API_REPORT_INTERVAL', 3600)  # 1시간
        )
        logger.info("API 사용량 모니터링 시스템 초기화 완료")
    except Exception as e:
        logger.error(f"API 사용량 모니터링 시스템 초기화 실패: {e}")
        # 필수 시스템이 아니므로 앱 시작 차단하지 않음
        logger.warning("API 사용량 모니터 없이 계속 진행")

    # 7-2. Google Calendar 역방향 동기화 스케줄러
    try:
        from dashboard.services.calendar_sync_scheduler import start_calendar_sync_scheduler

        # Calendar 동기화 스케줄러 시작 (5분마다 삭제된 이벤트 확인)
        if app.config.get('CALENDAR_SYNC_ENABLED', True):
            start_calendar_sync_scheduler()
            logger.info("Google Calendar 역방향 동기화 스케줄러 초기화 완료")
        else:
            logger.info("CALENDAR_SYNC_ENABLED=False - Calendar 동기화 비활성화")
    except Exception as e:
        logger.error(f"Calendar 동기화 스케줄러 초기화 실패: {e}")
        # 필수 시스템이 아니므로 앱 시작 차단하지 않음
        logger.warning("Calendar 동기화 스케줄러 없이 계속 진행")

    # 7-1. 보안 미들웨어 초기화 (CSRF, Rate Limiting, XSS 방어)
    try:
        from dashboard.utils.security_middleware import init_security_middleware

        init_security_middleware(app, app.config['SECRET_KEY'])
        logger.info("보안 미들웨어 초기화 완료 (CSRF, Rate Limiting, XSS 방어 활성화)")
    except Exception as e:
        logger.error(f"보안 미들웨어 초기화 실패: {e}")
        logger.warning("보안 미들웨어 없이 계속 진행 - 보안 취약점 존재")

    # 8. 백그라운드 서비스 초기화 (Flask 2.2+ 호환)
    _background_services_initialized = False

    @app.before_request
    def initialize_background_services():
        """앱 첫 요청 시에만 백그라운드 서비스 초기화 (한 번만 실행)"""
        nonlocal _background_services_initialized
        if not _background_services_initialized:
            try:
                # BACKGROUND_PREFETCH_ENABLED 설정 확인
                if app.config.get('BACKGROUND_PREFETCH_ENABLED', True):
                    logger.info("백그라운드 프리페치 시스템 초기화 중...")
                    init_background_prefetch()
                    logger.info("백그라운드 프리페치 시스템 초기화 완료")
                else:
                    logger.info("BACKGROUND_PREFETCH_ENABLED=False - 백그라운드 프리페치 비활성화")
                _background_services_initialized = True
            except Exception as e:
                logger.error(f"백그라운드 서비스 초기화 실패: {e}")

    # 9. 프로세스 종료 시 정리 작업 (프리패처는 프로세스 수명과 함께)
    import atexit
    import signal
    import sys

    def cleanup_background_services():
        """프로세스 종료 시에만 백그라운드 서비스 정리"""
        try:
            from dashboard.utils.background_prefetch import stop_background_prefetch
            stop_background_prefetch()
            logger.info("백그라운드 프리페치 시스템이 정상적으로 종료되었습니다.")
        except Exception as e:
            logger.debug(f"백그라운드 프리페치 정리 중 오류: {e}")

        try:
            from dashboard.utils.api_usage_monitor import stop_api_monitoring_scheduler
            stop_api_monitoring_scheduler()
            logger.info("API 모니터링 스케줄러가 정상적으로 종료되었습니다.")
        except Exception as e:
            logger.debug(f"API 모니터링 스케줄러 정리 중 오류: {e}")

        try:
            from dashboard.services.calendar_sync_scheduler import stop_calendar_sync_scheduler
            stop_calendar_sync_scheduler()
            logger.info("Calendar 동기화 스케줄러가 정상적으로 종료되었습니다.")
        except Exception as e:
            logger.debug(f"Calendar 동기화 스케줄러 정리 중 오류: {e}")

    # 프로세스 종료 시에만 정리
    atexit.register(cleanup_background_services)

    # 시그널 핸들러로 명시적 종료 처리 (메인 스레드에서만 실행)
    try:
        def signal_handler(signum, frame):
            logger.info(f"시그널 {signum} 수신, 백그라운드 서비스 정리 중...")
            cleanup_background_services()
            sys.exit(0)

        signal.signal(signal.SIGTERM, signal_handler)
        if hasattr(signal, 'SIGINT'):
            signal.signal(signal.SIGINT, signal_handler)
    except ValueError as e:
        # 멀티스레딩 환경에서는 시그널 핸들러 등록 불가
        logger.warning(f"시그널 핸들러 등록 실패: {e}")
        logger.info("atexit 핸들러로만 정리 작업 수행")

    # 10. 블루프린트 등록 (기존 구조 사용)
    try:
        from dashboard.blueprints.projects import projects_bp
        from dashboard.blueprints.monitoring import monitoring_bp
        from dashboard.blueprints.auth import auth_bp
        from dashboard.blueprints.main import main_bp
        from dashboard.blueprints.admin import admin_bp
        from dashboard.blueprints.users import users_bp
        from dashboard.blueprints.analytics import analytics_bp
        from dashboard.blueprints.stats import stats_bp
        from dashboard.blueprints.data_management import data_management_bp
        from dashboard.blueprints.locks import locks_bp
        from dashboard.blueprints.folders import folders_bp
        from dashboard.blueprints.metadata import metadata_bp
        from dashboard.blueprints.leads import leads_bp
        from dashboard.api.v1.auth import auth_bp as auth_api_bp

        # 기본 블루프린트들 등록
        app.register_blueprint(main_bp)
        app.register_blueprint(auth_bp)
        app.register_blueprint(projects_bp)
        app.register_blueprint(leads_bp)
        app.register_blueprint(monitoring_bp)
        app.register_blueprint(admin_bp)
        app.register_blueprint(users_bp)
        app.register_blueprint(analytics_bp)
        app.register_blueprint(stats_bp)
        app.register_blueprint(data_management_bp)
        app.register_blueprint(locks_bp)
        app.register_blueprint(folders_bp)
        app.register_blueprint(metadata_bp)
        app.register_blueprint(auth_api_bp)  # API v1 인증 엔드포인트

        logger.info("블루프린트 등록 완료 (main, auth, projects, leads, monitoring, admin, users, analytics, stats, data_management, locks, folders, metadata, auth_api)")
    except Exception as e:
        logger.error(f"블루프린트 등록 실패: {e}")
        # 블루프린트는 필수가 아니므로 앱 시작을 차단하지 않음
        logger.warning("블루프린트 없이 계속 진행")

    # FUTURE: 미래 확장 블루프린트들 (계획 중)
    # register_auth_blueprint(app)         # 인증 & 권한
    # register_projects_blueprint(app)     # 프로젝트 관리
    # register_analytics_blueprint(app)    # 데이터 분석 & 리포트
    # register_realtime_blueprint(app)     # 실시간 업데이트
    # register_collaboration_blueprint(app)# 협업 도구 & 댓글
    # register_admin_blueprint(app)        # 관리자 패널
    # register_data_blueprint(app)         # 데이터 관리
    # register_audit_blueprint(app)        # 감사 & 로깅
    # register_dev_blueprint(app)          # 개발 & 디버깅

    # 10. Manifest 상태 로깅 (템플릿 헬퍼는 init_frontend_helpers에서 처리)
    from .utils.frontend_helpers import asset_manager

    # 앱 기동 시 manifest 유무 검사 및 로깅
    try:
        manifest = asset_manager._load_manifest()
        if manifest:
            logger.info(f"Vite manifest 파일 발견: {len(manifest)}개 에셋 로드 가능")
            logger.info("모던 빌드 시스템이 활성화되었습니다")
        else:
            logger.warning("Vite manifest 파일이 없습니다 - 레거시 폴백 모드로 동작")
            logger.warning("npm run build를 실행하여 모던 에셋을 생성하세요")
    except Exception as e:
        logger.error(f"Manifest 파일 확인 중 오류: {e}")
        logger.warning("레거시 폴백 모드로 동작합니다")

    # 템플릿 헬퍼 함수들은 init_frontend_helpers()에서 통합 관리됨 (중복 제거 완료)

    # 11. SocketIO 초기화 (선택적)
    socketio_instance = None
    if enable_socketio:
        try:
            from flask_socketio import SocketIO
            # 긴급 조치: eventlet 메모리 크래시 방지 (SIGSEGV 해결)
            # threading 모드로 강제 설정하여 안정성 확보
            # eventlet의 greenlet 메모리 관리 버그를 우회
            socketio_instance = SocketIO(
                app,
                cors_allowed_origins="*",
                ping_timeout=SOCKETIO_PING_TIMEOUT,
                ping_interval=SOCKETIO_PING_INTERVAL,
                async_mode='threading',  # threading 모드로 강제 (메모리 크래시 방지)
                max_http_buffer_size=SOCKETIO_MAX_HTTP_BUFFER_SIZE,
                engineio_logger=False,  # 로그 부하 감소
                logger=False,           # 로그 부하 감소
                transports=['polling']  # WebSocket 비활성화 (개발 환경 Werkzeug 호환성)
            )
            logger.info(f"SocketIO 초기화 완료 (async_mode={socketio_instance.async_mode}, ping_timeout={SOCKETIO_PING_TIMEOUT}s, ping_interval={SOCKETIO_PING_INTERVAL}s)")

            # 전역 socketio 변수에 할당 (블루프린트에서 import 가능)
            global socketio
            socketio = socketio_instance
            logger.info("SocketIO 인스턴스를 전역으로 export 완료")
        except ImportError:
            logger.warning("flask_socketio 모듈이 없어 SocketIO 사용 불가")
        except Exception as e:
            logger.error(f"SocketIO 초기화 실패: {e}")

    logger.info("Flask 앱 팩토리 초기화 완료")
    return app, socketio_instance

# 편의 함수들
def create_dev_app():
    """개발 환경 앱 생성"""
    return create_app('development')

def create_test_app():
    """테스트 환경 앱 생성 (SocketIO 비활성화)"""
    return create_app('testing', enable_socketio=False)

def create_prod_app():
    """프로덕션 환경 앱 생성"""
    return create_app('production')

# 전역 호환성 유지 인스턴스 (레거시)
# LEGACY: 전역 호환성 유지 (미래 블루프린트 분리 후 제거 예정)
app = None
socketio = None  # SocketIO 인스턴스 (블루프린트에서 import 가능)

def get_app():
    """앱 인스턴스 반환 (전역 호환성, SocketIO 없이)"""
    global app
    if app is None:
        app, _ = create_app(enable_socketio=False)  # 레거시 호환성을 위해 SocketIO 비활성화
    return app
