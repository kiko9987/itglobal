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
    from utils.security import init_security
    from utils.jwt_auth import init_jwt_manager
    from utils.database import init_database
    
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
        init_security(app)
        init_jwt_manager(app.config['SECRET_KEY'])
        init_database(temp_db.name)
    
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
    from utils.jwt_auth import create_jwt_tokens
    
    tokens = create_jwt_tokens(test_user)
    return {
        'Authorization': f"Bearer {tokens['access_token']}",
        'Content-Type': 'application/json'
    }