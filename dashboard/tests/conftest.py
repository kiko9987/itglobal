"""
pytest 설정 파일
테스트용 픽스처 및 설정 정의
"""

import pytest
import tempfile
import os
import sys
from unittest.mock import Mock, patch

# 프로젝트 루트를 시스템 경로에 추가
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

@pytest.fixture
def app():
    """테스트용 Flask 앱 픽스처"""
    from flask import Flask
    from utils.security_middleware import init_security_middleware
    from .jwt_auth import init_jwt_manager
    from utils.user_database import get_user_database

    # 블루프린트 임포트
    from dashboard.blueprints.main import main_bp
    from dashboard.blueprints.auth import auth_bp
    from dashboard.blueprints.projects import projects_bp
    from dashboard.blueprints.monitoring import monitoring_bp
    from dashboard.blueprints.admin import admin_bp
    from dashboard.blueprints.users import users_bp
    from dashboard.blueprints.analytics import analytics_bp
    from dashboard.blueprints.stats import stats_bp
    from dashboard.blueprints.data_management import data_management_bp
    from dashboard.blueprints.locks import locks_bp
    from dashboard.blueprints.folders import folders_bp
    from dashboard.blueprints.metadata import metadata_bp

    # 테스트용 임시 앱 생성
    app = Flask(__name__)
    app.config['TESTING'] = True
    app.config['SECRET_KEY'] = 'test-secret-key'
    app.config['WTF_CSRF_ENABLED'] = False  # 테스트에서 CSRF 비활성화

    # 테스트용 임시 데이터베이스
    temp_db = tempfile.NamedTemporaryFile(delete=False, suffix='.db')
    temp_db.close()

    with app.app_context():
        # 보안 시스템 초기화
        init_security_middleware(app, app.config['SECRET_KEY'])
        init_jwt_manager(app.config['SECRET_KEY'])
        # 통합 DB 초기화 (테이블 자동 생성됨)
        get_user_database(temp_db.name)

        # 블루프린트 등록
        app.register_blueprint(main_bp)
        app.register_blueprint(auth_bp)
        app.register_blueprint(projects_bp)
        app.register_blueprint(monitoring_bp)
        app.register_blueprint(admin_bp)
        app.register_blueprint(users_bp)
        app.register_blueprint(analytics_bp)
        app.register_blueprint(stats_bp)
        app.register_blueprint(data_management_bp)
        app.register_blueprint(locks_bp)
        app.register_blueprint(folders_bp)
        app.register_blueprint(metadata_bp)

    yield app

    # 정리
    try:
        os.unlink(temp_db.name)
    except OSError:
        pass

@pytest.fixture
def client(app):
    """테스트용 클라이언트 픽스처"""
    return app.test_client()

@pytest.fixture
def runner(app):
    """테스트용 CLI 러너 픽스처"""
    return app.test_cli_runner()

@pytest.fixture
def test_user():
    """테스트용 사용자 데이터"""
    return {
        'email': 'test@example.com',
        'name': '테스트 사용자',
        'password': 'test_password_123',
        'permission_level': 'admin'
    }

@pytest.fixture
def mock_google_sheets():
    """Google Sheets API 모킹"""
    with patch('utils.google_sheets.GoogleSheetsManager') as mock:
        mock_instance = Mock()
        mock_instance.get_sheet_data.return_value = Mock()
        mock_instance.validate_connection.return_value = True
        mock.return_value = mock_instance
        yield mock_instance

@pytest.fixture
def sample_project_data():
    """테스트용 프로젝트 데이터"""
    return {
        'project_code': 'G0001-IT',
        'company': '테스트 회사',
        'manager': '테스트 매니저',
        'address': '서울시 테스트구 테스트동',
        'total_amount': 1000000,
        'construction_type': '신규 설치'
    }

@pytest.fixture
def auth_headers(test_user):
    """인증 헤더 픽스처"""
    from .jwt_auth import create_jwt_tokens

    tokens = create_jwt_tokens(test_user)
    return {
        'Authorization': f"Bearer {tokens['access_token']}",
        'Content-Type': 'application/json'
    }

@pytest.fixture(autouse=True)
def bypass_login_required():
    """모든 테스트에서 @login_required 데코레이터 우회"""
    # login_required가 원래 함수를 그대로 반환하도록 mock
    with patch('dashboard.auth.login_required', lambda func: func):
        with patch('dashboard.auth.admin_required', lambda func: func):
            # _validate_session이 None을 반환하도록 (인증 성공)
            with patch('dashboard.auth._validate_session', return_value=None):
                yield