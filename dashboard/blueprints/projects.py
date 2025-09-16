import logging
import os
from datetime import datetime

from flask import Blueprint, render_template, redirect, url_for, request, session, jsonify

from auth import login_required, get_user_role

from dashboard.services.project_service import (
    can_user_edit_project,
    check_overdue_status,
    get_project_records,
    invalidate_project_cache,
    get_sheets_manager,
    get_current_data,
    load_data,
)
from dashboard.utils.field_lock_manager import field_lock_manager

logger = logging.getLogger(__name__)

projects_bp = Blueprint('projects', __name__)


@projects_bp.route('/projects')
@login_required
def project_list():
    user_role = get_user_role()
    user_email = session.get('user', {}).get('email', '')

    try:
        projects_data = get_project_records()
        if projects_data:
            for project in projects_data:
                project['is_cancelled'] = project.get('수금 관련 특이사항') == '공사취소'
                project['can_edit'] = can_user_edit_project(project, user_email, user_role)

                project_code = project.get('프로젝트 코드')
                if project_code:
                    project['lock_status'] = field_lock_manager.get_project_locks(project_code)

                project['is_overdue'] = check_overdue_status(project)
        else:
            projects_data = []
    except Exception as exc:
        logger.error('프로젝트 데이터 로드 실패: %s', exc)
        projects_data = []

    return render_template(
        'project_list_server.html',
        projects=projects_data,
        user_role=user_role,
        user_email=user_email,
        timestamp=datetime.now()
    )


@projects_bp.route('/projects/cancel/<project_code>', methods=['POST'])
@login_required
def cancel_project_server(project_code):
    try:
        user_email = session.get('user', {}).get('email', '')
        user_name = session.get('user', {}).get('name', '')

        projects_data = get_project_records()
        project = next((p for p in projects_data if p.get('프로젝트 코드') == project_code), None)

        if not project:
            logger.warning(f"프로젝트를 찾을 수 없음: {project_code}")
            return redirect(url_for('projects.project_list', error='프로젝트를 찾을 수 없습니다.'))

        if project.get('수금 관련 특이사항') == '공사취소':
            return redirect(url_for('projects.project_list', error='이미 취소된 프로젝트입니다.'))

        manager = get_sheets_manager()
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        sheet_range = '공사 현황!A:AM'

        update_data = {
            'projectCode': project_code,
            '수금 관련 특이사항': '공사취소'
        }

        result = manager.update_project_data(sheet_id, sheet_range, update_data)

        if result['success']:
            invalidate_project_cache(project_code)
            logger.info(f"프로젝트 취소 완료: {project_code} by {user_name}")
            return redirect(url_for('projects.project_list', success=f'프로젝트 {project_code}가 취소되었습니다.'))
        else:
            logger.error(f"프로젝트 취소 실패: {project_code} - {result.get('error')}")
            return redirect(url_for('projects.project_list', error='프로젝트 취소에 실패했습니다.'))

    except Exception as e:
        logger.error(f"프로젝트 취소 처리 오류: {e}")
        return redirect(url_for('projects.project_list', error='서버 오류가 발생했습니다.'))


@projects_bp.route('/projects/resume/<project_code>', methods=['POST'])
@login_required
def resume_project_server(project_code):
    try:
        user_email = session.get('user', {}).get('email', '')
        user_name = session.get('user', {}).get('name', '')

        projects_data = get_project_records()
        project = next((p for p in projects_data if p.get('프로젝트 코드') == project_code), None)

        if not project:
            return redirect(url_for('projects.project_list', error='프로젝트를 찾을 수 없습니다.'))

        if project.get('수금 관련 특이사항') != '공사취소':
            return redirect(url_for('projects.project_list', error='취소되지 않은 프로젝트입니다.'))

        manager = get_sheets_manager()
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        sheet_range = '공사 현황!A:AM'

        update_data = {
            'projectCode': project_code,
            '수금 관련 특이사항': '-'
        }

        result = manager.update_project_data(sheet_id, sheet_range, update_data)

        if result['success']:
            invalidate_project_cache(project_code)
            logger.info(f"프로젝트 재개 완료: {project_code} by {user_name}")
            return redirect(url_for('projects.project_list', success=f'프로젝트 {project_code}가 재개되었습니다.'))
        else:
            logger.error(f"프로젝트 재개 실패: {project_code} - {result.get('error')}")
            return redirect(url_for('projects.project_list', error='프로젝트 재개에 실패했습니다.'))

    except Exception as e:
        logger.error(f"프로젝트 재개 처리 오류: {e}")
        return redirect(url_for('projects.project_list', error='서버 오류가 발생했습니다.'))


@projects_bp.route('/projects/update/<project_code>', methods=['POST'])
@login_required
def update_project_server(project_code):
    try:
        user_email = session.get('user', {}).get('email', '')
        user_name = session.get('user', {}).get('name', '')

        projects_data = get_project_records()
        project = next((p for p in projects_data if p.get('프로젝트 코드') == project_code), None)

        if not project:
            return redirect(url_for('projects.project_list', error='프로젝트를 찾을 수 없습니다.'))

        if project.get('수금 관련 특이사항') == '공사취소':
            return redirect(url_for('projects.project_list', error='취소된 프로젝트는 편집할 수 없습니다.'))

        if not can_user_edit_project(project, user_email, get_user_role()):
            return redirect(url_for('projects.project_list', error='편집 권한이 없습니다.'))

        update_data = {'projectCode': project_code}
        changes = []

        for field_name, new_value in request.form.items():
            if field_name != 'projectCode':
                old_value = project.get(field_name, '')
                if str(old_value) != str(new_value):
                    update_data[field_name] = new_value
                    changes.append({
                        'field': field_name,
                        'old': old_value,
                        'new': new_value
                    })

        if not changes:
            return redirect(url_for('projects.project_list', info='변경사항이 없습니다.'))

        manager = get_sheets_manager()
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        sheet_range = '공사 현황!A:AM'

        result = manager.update_project_data(sheet_id, sheet_range, update_data)

        if result['success']:
            invalidate_project_cache(project_code)
            logger.info(f"프로젝트 업데이트 완료: {project_code} by {user_name}")
            return redirect(url_for('projects.project_list', success=f'프로젝트 {project_code}가 업데이트되었습니다.'))
        else:
            logger.error(f"프로젝트 업데이트 실패: {project_code} - {result.get('error')}")
            return redirect(url_for('projects.project_list', error='프로젝트 업데이트에 실패했습니다.'))

    except Exception as e:
        logger.error(f"프로젝트 업데이트 처리 오류: {e}")
        return redirect(url_for('projects.project_list', error='서버 오류가 발생했습니다.'))