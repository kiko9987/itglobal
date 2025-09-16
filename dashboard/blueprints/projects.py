import logging
import os
from datetime import datetime

from flask import Blueprint, render_template, redirect, url_for, request, session, jsonify

from auth import login_required, get_user_role

from ..services.project_service import (
    can_user_edit_project,
    check_overdue_status,
    get_project_records,
    invalidate_project_cache,
    get_sheets_manager,
    get_current_data,
    load_data,
)
from ..utils.field_lock_manager import field_lock_manager

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
                project['is_cancelled'] = project.get('?섍툑 愿???뱀씠?ы빆') == '怨듭궗痍⑥냼'
                project['can_edit'] = can_user_edit_project(project, user_email, user_role)

                project_code = project.get('?꾨줈?앺듃 肄붾뱶')
                if project_code:
                    project['lock_status'] = field_lock_manager.get_project_locks(project_code)

                project['is_overdue'] = check_overdue_status(project)
        else:
            projects_data = []
    except Exception as exc:
        logger.error('?꾨줈?앺듃 ?곗씠??濡쒕뱶 ?ㅽ뙣: %s', exc)
        projects_data = []

    return render_template(
        'project_list_server.html',
        projects=projects_data,
        user_role=user_role,
        user_email=user_email,
        timestamp=datetime.now(),
    )


@projects_bp.route('/projects/cancel/<project_code>', methods=['POST'])
@login_required
def cancel_project_server(project_code):
    try:
        user_email = session.get('user', {}).get('email', '')
        user_name = session.get('user', {}).get('name', '')

        projects_data = get_project_records()
        project = next((p for p in projects_data if p.get('?꾨줈?앺듃 肄붾뱶') == project_code), None)
        if not project:
            logger.warning('?꾨줈?앺듃瑜?李얠쓣 ???놁쓬: %s', project_code)
            return redirect(url_for('projects.project_list', error='?꾨줈?앺듃瑜?李얠쓣 ???놁뒿?덈떎.'))

        if project.get('?섍툑 愿???뱀씠?ы빆') == '怨듭궗痍⑥냼':
            return redirect(url_for('projects.project_list', error='?대? 痍⑥냼???꾨줈?앺듃?낅땲??'))

        manager = get_sheets_manager()
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        sheet_range = os.getenv('PROJECT_SHEET_RANGE', '怨듭궗 ?꾪솴!A:AM')

        update_data = {
            'projectCode': project_code,
            '?섍툑 愿???뱀씠?ы빆': '怨듭궗痍⑥냼'
        }

        result = manager.update_project_data(sheet_id, sheet_range, update_data)

        if result['success']:
            from app import save_audit_log  # lazy import to avoid circular dependency
            save_audit_log(
                user_name,
                user_email,
                get_user_role(),
                'CANCEL_PROJECT',
                f'?꾨줈?앺듃 痍⑥냼: {project_code}',
                project_code,
                '?섍툑 愿???뱀씠?ы빆',
                project.get('?섍툑 愿???뱀씠?ы빆', ''),
                '怨듭궗痍⑥냼'
            )
            invalidate_project_cache(project_code)
            logger.info('?꾨줈?앺듃 痍⑥냼 ?꾨즺: %s by %s', project_code, user_name)
            return redirect(url_for('projects.project_list', success=f'?꾨줈?앺듃 {project_code}媛 痍⑥냼?섏뿀?듬땲??'))

        logger.error('?꾨줈?앺듃 痍⑥냼 ?ㅽ뙣: %s - %s', project_code, result.get('error'))
        return redirect(url_for('projects.project_list', error='?꾨줈?앺듃 痍⑥냼???ㅽ뙣?덉뒿?덈떎.'))

    except Exception as exc:
        logger.error('?꾨줈?앺듃 痍⑥냼 泥섎━ ?ㅻ쪟: %s', exc)
        return redirect(url_for('projects.project_list', error='?쒕쾭 ?ㅻ쪟媛 諛쒖깮?덉뒿?덈떎.'))


@projects_bp.route('/projects/resume/<project_code>', methods=['POST'])
@login_required
def resume_project_server(project_code):
    try:
        user_email = session.get('user', {}).get('email', '')
        user_name = session.get('user', {}).get('name', '')

        projects_data = get_project_records()
        project = next((p for p in projects_data if p.get('?꾨줈?앺듃 肄붾뱶') == project_code), None)
        if not project:
            logger.warning('?꾨줈?앺듃瑜?李얠쓣 ???놁쓬: %s', project_code)
            return redirect(url_for('projects.project_list', error='?꾨줈?앺듃瑜?李얠쓣 ???놁뒿?덈떎.'))

        if project.get('?섍툑 愿???뱀씠?ы빆') != '怨듭궗痍⑥냼':
            return redirect(url_for('projects.project_list', error='痍⑥냼?섏? ?딆? ?꾨줈?앺듃?낅땲??'))

        manager = get_sheets_manager()
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        sheet_range = os.getenv('PROJECT_SHEET_RANGE', '怨듭궗 ?꾪솴!A:AM')

        update_data = {
            'projectCode': project_code,
            '?섍툑 愿???뱀씠?ы빆': '-'
        }

        result = manager.update_project_data(sheet_id, sheet_range, update_data)

        if result['success']:
            from app import save_audit_log
            save_audit_log(
                user_name,
                user_email,
                get_user_role(),
                'RESUME_PROJECT',
                f'?꾨줈?앺듃 ?ш컻: {project_code}',
                project_code,
                '?섍툑 愿???뱀씠?ы빆',
                '怨듭궗痍⑥냼',
                '-'
            )
            invalidate_project_cache(project_code)
            logger.info('?꾨줈?앺듃 ?ш컻 ?꾨즺: %s by %s', project_code, user_name)
            return redirect(url_for('projects.project_list', success=f'?꾨줈?앺듃 {project_code}媛 ?ш컻?섏뿀?듬땲??'))

        logger.error('?꾨줈?앺듃 ?ш컻 ?ㅽ뙣: %s - %s', project_code, result.get('error'))
        return redirect(url_for('projects.project_list', error='?꾨줈?앺듃 ?ш컻???ㅽ뙣?덉뒿?덈떎.'))

    except Exception as exc:
        logger.error('?꾨줈?앺듃 ?ш컻 泥섎━ ?ㅻ쪟: %s', exc)
        return redirect(url_for('projects.project_list', error='?쒕쾭 ?ㅻ쪟媛 諛쒖깮?덉뒿?덈떎.'))


@projects_bp.route('/projects/update/<project_code>', methods=['POST'])
@login_required
def update_project_server(project_code):
    try:
        user_email = session.get('user', {}).get('email', '')
        user_name = session.get('user', {}).get('name', '')

        projects_data = get_project_records()
        project = next((p for p in projects_data if p.get('?꾨줈?앺듃 肄붾뱶') == project_code), None)
        if not project:
            return redirect(url_for('projects.project_list', error='?꾨줈?앺듃瑜?李얠쓣 ???놁뒿?덈떎.'))

        if project.get('?섍툑 愿???뱀씠?ы빆') == '怨듭궗痍⑥냼':
            return redirect(url_for('projects.project_list', error='痍⑥냼???꾨줈?앺듃???몄쭛?????놁뒿?덈떎.'))

        if not can_user_edit_project(project, user_email, get_user_role()):
            return redirect(url_for('projects.project_list', error='?몄쭛 沅뚰븳???놁뒿?덈떎.'))

        update_data = {'projectCode': project_code}
        changes = []
        for field_name, new_value in request.form.items():
            if field_name != 'projectCode':
                old_value = project.get(field_name, '')
                if str(old_value) != str(new_value):
                    update_data[field_name] = new_value
                    changes.append({'field': field_name, 'old': old_value, 'new': new_value})

        if not changes:
            return redirect(url_for('projects.project_list', info='蹂寃쎌궗??씠 ?놁뒿?덈떎.'))

        manager = get_sheets_manager()
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        sheet_range = os.getenv('PROJECT_SHEET_RANGE', '怨듭궗 ?꾪솴!A:AM')

        result = manager.update_project_data(sheet_id, sheet_range, update_data)

        if result['success']:
            from app import save_audit_log
            for change in changes:
                save_audit_log(
                    user_name,
                    user_email,
                    get_user_role(),
                    'UPDATE_PROJECT',
                    f"?꾨줈?앺듃 ?꾨뱶 ?섏젙: {change['field']}",
                    project_code,
                    change['field'],
                    change['old'],
                    change['new']
                )

            invalidate_project_cache(project_code)
            logger.info('?꾨줈?앺듃 ?낅뜲?댄듃 ?꾨즺: %s by %s', project_code, user_name)
            return redirect(url_for('projects.project_list', success=f'?꾨줈?앺듃 {project_code}媛 ?낅뜲?댄듃?섏뿀?듬땲??'))

        logger.error('?꾨줈?앺듃 ?낅뜲?댄듃 ?ㅽ뙣: %s - %s', project_code, result.get('error'))
        return redirect(url_for('projects.project_list', error='?꾨줈?앺듃 ?낅뜲?댄듃???ㅽ뙣?덉뒿?덈떎.'))

    except Exception as exc:
        logger.error('?꾨줈?앺듃 ?낅뜲?댄듃 泥섎━ ?ㅻ쪟: %s', exc)
        return redirect(url_for('projects.project_list', error='?쒕쾭 ?ㅻ쪟媛 諛쒖깮?덉뒿?덈떎.'))


