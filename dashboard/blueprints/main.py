"""
메인 블루프린트
루트 라우트와 기본 페이지들
"""

from flask import Blueprint, redirect, url_for, render_template, session

main_bp = Blueprint('main', __name__)


@main_bp.route('/')
def dashboard():
    """메인 페이지 - 프로젝트 관리로 리다이렉트"""
    return redirect(url_for('projects.project_list'))


@main_bp.route('/test-vite')
def test_vite():
    """Vite 개발 서버 테스트 페이지"""
    from ..utils.frontend_helpers import asset_manager

    # Vite 개발 서버 상태 확인
    try:
        from ..app import is_vite_dev_server_running, get_vite_dev_port
        is_dev_mode = is_vite_dev_server_running()
        vite_port = get_vite_dev_port()
    except ImportError:
        is_dev_mode = False
        vite_port = 5173

    js_url = asset_manager.get_js_bundle('project-list')
    css_url = asset_manager.get_css_bundle('main')

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Vite 테스트</title>
        <meta charset="utf-8">
        {f'<link rel="stylesheet" href="{css_url}">' if css_url else ''}
    </head>
    <body>
        <h1>Vite 개발 서버 상태 테스트</h1>
        <ul>
            <li>개발 모드: {'Yes' if is_dev_mode else 'No'}</li>
            <li>Vite 포트: {vite_port}</li>
            <li>CSS URL: {css_url or 'None'}</li>
            <li>JS URL: {js_url or 'None'}</li>
        </ul>
        {f'<script type="module" src="{js_url}"></script>' if js_url else ''}
    </body>
    </html>
    """



