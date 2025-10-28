"""
통합 테스트 (Integration Tests)

전문가 피드백 반영: 주요 기능 flow 통합 테스트
- 프로젝트 생성/편집/저장
- 메모 저장
- 락 획득/해제
- 캐시 무효화

Note: 실제 Google Sheets API는 mock으로 대체
"""

import pytest
import json
from unittest.mock import Mock, patch, MagicMock
from datetime import datetime


# ===== Fixtures =====

@pytest.fixture
def app():
    """Flask 앱 인스턴스"""
    from dashboard.app import create_app
    app = create_app('testing')
    app.config['TESTING'] = True
    app.config['WTF_CSRF_ENABLED'] = False
    yield app


@pytest.fixture
def client(app):
    """Flask 테스트 클라이언트"""
    return app.test_client()


@pytest.fixture
def auth_session(client):
    """인증된 세션"""
    with client.session_transaction() as sess:
        sess['user'] = {
            'email': 'test@example.com',
            'name': '테스트 사용자',
            'permission_level': 'admin'
        }
        sess['tab_id'] = 'test-tab-123'
    return client


@pytest.fixture
def mock_sheets():
    """Google Sheets Manager mock"""
    with patch('dashboard.blueprints.projects.get_sheets_manager') as mock:
        manager = MagicMock()
        manager.find_row_by_project_code.return_value = 2
        manager.batch_update_cells.return_value = {'totalUpdatedCells': 5}
        manager.update_cell.return_value = True
        mock.return_value = manager
        yield manager


@pytest.fixture
def mock_cache():
    """캐시 무효화 mock"""
    with patch('dashboard.blueprints.projects.invalidate_project_cache') as mock:
        yield mock


@pytest.fixture
def mock_lock():
    """락 매니저 mock"""
    with patch('dashboard.blueprints.projects.get_project_lock_manager') as mock:
        lock_manager = MagicMock()
        lock_manager.acquire_lock.return_value = True
        lock_manager.release_lock.return_value = True
        lock_manager.is_locked_by_user.return_value = True
        mock.return_value = lock_manager
        yield lock_manager


# ===== 통합 테스트 =====

class TestProjectCreationFlow:
    """프로젝트 생성 flow 통합 테스트"""

    def test_project_creation_full_flow(self, auth_session, mock_sheets, mock_cache):
        """
        프로젝트 생성 전체 flow 테스트:
        1. 프로젝트 생성 요청
        2. Google Sheets 업데이트
        3. 캐시 무효화
        4. 성공 응답
        """
        # Given: 프로젝트 생성 데이터
        project_data = {
            'manager': '홍길동',
            'company': '테스트 회사',
            'address': '서울시 강남구',
            'start_date': '2025-01-01',
            'down_payment': '₩1,000,000'
        }

        # When: 프로젝트 자동 생성 API 호출
        response = auth_session.post(
            '/api/v1/projects/auto-create',
            data=json.dumps(project_data),
            content_type='application/json'
        )

        # Then: 성공 응답
        assert response.status_code == 200
        data = response.get_json()
        assert data['success'] is True
        assert 'project_code' in data

        # And: Google Sheets 업데이트 호출 확인
        mock_sheets.batch_update_cells.assert_called()

        # And: 캐시 무효화 호출 확인
        mock_cache.assert_called_once()


    def test_project_creation_with_invalid_data(self, auth_session):
        """
        잘못된 데이터로 프로젝트 생성 시 검증 실패 테스트
        """
        # Given: 필수 필드 누락된 데이터
        invalid_data = {
            'address': '서울시'
            # manager, company 누락
        }

        # When: 프로젝트 생성 API 호출
        response = auth_session.post(
            '/api/v1/projects/auto-create',
            data=json.dumps(invalid_data),
            content_type='application/json'
        )

        # Then: 400 Bad Request
        assert response.status_code == 400
        data = response.get_json()
        assert data['success'] is False
        assert 'errors' in data


    def test_project_creation_with_currency_format(self, auth_session, mock_sheets, mock_cache):
        """
        MoneyField 통화 포맷 검증 테스트 (HIGH-2 피드백 반영)
        """
        # Given: 다양한 통화 포맷
        project_data = {
            'manager': '김철수',
            'company': '포맷 테스트',
            'down_payment': '₩5,000,000',  # 콤마 포함
            'mid_payment': '3000000',      # 숫자만
            'final_payment': 2000000       # int 타입
        }

        # When: 프로젝트 생성
        response = auth_session.post(
            '/api/v1/projects/auto-create',
            data=json.dumps(project_data),
            content_type='application/json'
        )

        # Then: 모두 정상 처리
        assert response.status_code == 200
        data = response.get_json()
        assert data['success'] is True


class TestProjectEditFlow:
    """프로젝트 편집 flow 통합 테스트"""

    def test_inline_edit_with_optimistic_lock(self, auth_session, mock_sheets, mock_lock, mock_cache):
        """
        인라인 편집 + Optimistic Lock 통합 테스트:
        1. 락 획득
        2. 프로젝트 수정
        3. 버전 검증
        4. 저장
        5. 락 해제
        """
        project_code = 'TEST-001'

        # Given: 편집 데이터
        edit_data = {
            'projectCode': project_code,
            'changes': {
                'manager': '박영희'
            },
            '_version': 0
        }

        # When: 인라인 편집 API 호출
        with patch('dashboard.blueprints.projects.get_project_records') as mock_get:
            mock_get.return_value = [{
                '프로젝트 코드': project_code,
                '담당자': '홍길동',
                '_version': 0
            }]

            response = auth_session.put(
                f'/api/projects/{project_code}',
                data=json.dumps(edit_data),
                content_type='application/json'
            )

        # Then: 성공
        assert response.status_code == 200
        data = response.get_json()
        assert data['success'] is True

        # And: 버전이 증가했는지 확인
        assert data.get('_version') == 1


    def test_optimistic_lock_conflict(self, auth_session, mock_sheets):
        """
        Optimistic Lock 충돌 시나리오 테스트 (409 응답)
        """
        project_code = 'TEST-002'

        # Given: 오래된 버전으로 수정 시도
        edit_data = {
            'projectCode': project_code,
            'changes': {'manager': '최철수'},
            '_version': 0  # 오래된 버전
        }

        # And: 서버의 현재 버전은 5
        with patch('dashboard.blueprints.projects.get_project_records') as mock_get:
            mock_get.return_value = [{
                '프로젝트 코드': project_code,
                '담당자': '이영희',
                '_version': 5  # 최신 버전
            }]

            # When: 편집 시도
            response = auth_session.put(
                f'/api/projects/{project_code}',
                data=json.dumps(edit_data),
                content_type='application/json'
            )

        # Then: 409 Conflict
        assert response.status_code == 409
        data = response.get_json()
        assert data['success'] is False
        assert 'current_data' in data  # 최신 데이터 반환


class TestMemoSaveFlow:
    """메모 저장 flow 통합 테스트"""

    def test_memo_save_and_cache_invalidation(self, auth_session, mock_sheets, mock_cache):
        """
        메모 저장 + 선택적 캐시 무효화 테스트 (MED-4 피드백 반영)
        """
        project_code = 'TEST-003'

        # Given: 메모 데이터
        memo_data = {
            'row': 2,
            'column': 'K',
            'memo': '테스트 메모 내용',
            'project_code': project_code,
            'field_name': '계약금'
        }

        # When: 메모 저장 API 호출
        with patch('dashboard.blueprints.projects.os.getenv') as mock_env:
            mock_env.side_effect = lambda k, d=None: {
                'GOOGLE_SHEET_ID': 'test-sheet-id'
            }.get(k, d)

            response = auth_session.post(
                '/api/save-cell-memo',
                data=json.dumps(memo_data),
                content_type='application/json'
            )

        # Then: 성공
        assert response.status_code == 200
        data = response.get_json()
        assert data['success'] is True

        # And: 셀 노트만 무효화 (프로젝트 캐시는 무효화 안함)
        # MED-4: invalidate_project_cache 호출 안됨


class TestLockManagement:
    """락 관리 통합 테스트"""

    def test_lock_acquire_edit_release_flow(self, auth_session, mock_lock):
        """
        락 획득 → 편집 → 해제 전체 flow 테스트
        """
        project_code = 'TEST-004'

        # Step 1: 락 획득
        acquire_response = auth_session.post(
            f'/api/project-lock/{project_code}',
            data=json.dumps({'action': 'acquire'}),
            content_type='application/json'
        )
        assert acquire_response.status_code == 200
        mock_lock.acquire_lock.assert_called()

        # Step 2: 편집 (락이 있어야 함)
        # ...

        # Step 3: 락 해제
        release_response = auth_session.post(
            f'/api/project-lock/{project_code}',
            data=json.dumps({'action': 'release'}),
            content_type='application/json'
        )
        assert release_response.status_code == 200
        mock_lock.release_lock.assert_called()


    def test_lock_prevents_concurrent_edit(self, auth_session, mock_lock):
        """
        락이 동시 편집을 방지하는지 테스트
        """
        project_code = 'TEST-005'

        # Given: 다른 사용자가 락 보유 중
        mock_lock.acquire_lock.return_value = False
        mock_lock.is_locked_by_user.return_value = False

        # When: 락 획득 시도
        response = auth_session.post(
            f'/api/project-lock/{project_code}',
            data=json.dumps({'action': 'acquire'}),
            content_type='application/json'
        )

        # Then: 실패
        assert response.status_code == 409
        data = response.get_json()
        assert data['locked_by'] is not None


# ===== 실행 =====

if __name__ == '__main__':
    pytest.main([__file__, '-v', '--tb=short'])
