import pandas as pd
from datetime import datetime, timedelta

from dashboard.services import project_service


def test_can_user_edit_project_admin_allows_edit():
    project = {'수금 관련 특이사항': '', '담당자 이메일': 'owner@example.com'}
    assert project_service.can_user_edit_project(project, 'other@example.com', 'admin')


def test_can_user_edit_project_owner_allows_edit():
    project = {'수금 관련 특이사항': '', '담당자 이메일': 'owner@example.com'}
    assert project_service.can_user_edit_project(project, 'owner@example.com', 'user')


def test_can_user_edit_project_denies_cancelled():
    project = {'수금 관련 특이사항': '공사취소', '담당자 이메일': 'owner@example.com'}
    assert not project_service.can_user_edit_project(project, 'owner@example.com', 'admin')


def test_check_overdue_status_detects_overdue():
    confirm = (datetime.now() - timedelta(days=5)).strftime('%Y-%m-%d')
    project = {'공사 확정': confirm, '계약금 입금일': ''}
    assert project_service.check_overdue_status(project)


def test_check_overdue_status_not_overdue_when_recent():
    confirm = datetime.now().strftime('%Y-%m-%d')
    project = {'공사 확정': confirm, '계약금 입금일': ''}
    assert not project_service.check_overdue_status(project)


def test_next_running_number_increments():
    df = pd.DataFrame({'프로젝트 코드': ['G0001-IT', 'G0002-IT']})
    assert project_service._next_running_number(df) == 3


def test_auto_project_code_uses_dataframe_values():
    df = pd.DataFrame({
        '프로젝트 코드': ['G0001-AA', 'G0002-AA'],
        '프로젝트 담당자': ['서울지점', '서울지점'],
        '담당자 이메일': ['owner@example.com', 'owner@example.com']
    })
    comp_map = {'서울지점': 'S'}
    owner_map = {'owner@example.com': 'IT'}
    # Patch helper maps by injecting into config override via DataFrame values
    df['프로젝트 담당자'] = ['서울지점', '서울지점']
    code = project_service._auto_project_code(df, '서울지점', 'owner@example.com')
    assert code.startswith('S')

