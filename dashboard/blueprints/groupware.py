"""그룹웨어 새 틀 — 화면 틀 고르기와 시범 명단 API.

- 모든 화면이 쓰는 modern_base.html 이 `gw_layout` 값으로 머리글만 바꿔 끼운다.
  명단에 없는 사람에겐 gw_layout=False → 기존 머리글 그대로.
- API 는 /api/ 아래(CSRF 경로 규칙).
"""
import re
import uuid
from datetime import datetime

from flask import Blueprint, jsonify, redirect, render_template_string, request, session

from dashboard.auth import admin_required, login_required
from dashboard.services import groupware as gw
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

groupware_bp = Blueprint('groupware', __name__)

# ── 현장 세션 문지기 ───────────────────────────────────────────────────
# 슬랙 링크로 들어온 현장·미화 직원(session.user.gw_field)은 아래 경로만. 나머지(기존 PM 화면·API)는 전부 막는다.
_FIELD_PAGES = ('/home', '/hr/me', '/leave', '/logout', '/login')
_FIELD_API = [
    (('GET',), re.compile(r'^/api/home$')),              # 홈(현장은 내 휴가만 담김)
    (('GET',), re.compile(r'^/api/hr/people/\d+$')),     # 본인 카드(본인 여부는 화면 쪽에서 다시 확인)
    (('GET',), re.compile(r'^/api/leave/me$')),          # 내 휴가
    (('POST',), re.compile(r'^/api/auth/ping$')),       # 세션 연장
]


def field_path_allowed(method: str, path: str) -> bool:
    if path.startswith('/static/') or path == '/favicon.ico' or path.startswith('/gw/login/'):
        return True
    if path in _FIELD_PAGES:
        return True
    return any(method in methods and pat.match(path) for methods, pat in _FIELD_API)


@groupware_bp.before_app_request
def field_gatekeeper():
    user = session.get('user') or {}
    if not user.get('gw_field'):
        return None
    if field_path_allowed(request.method, request.path):
        return None
    if request.path.startswith('/api/') or request.is_json:
        return jsonify({'success': False, 'message': '현장 계정은 쓸 수 없는 기능입니다.'}), 403
    return redirect('/home')


_LINK_FAIL_PAGE = '''<!doctype html><html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>링크를 쓸 수 없어요</title>
<style>body{font-family:-apple-system,'Segoe UI',sans-serif;background:#f8fafc;margin:0;padding:48px 16px;color:#0f172a}
.box{max-width:420px;margin:0 auto;background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:24px}
h1{font-size:1.15rem;margin:0 0 8px}p{color:#475569;line-height:1.6;margin:0}</style></head>
<body><div class="box"><h1>링크를 쓸 수 없어요</h1>
<p>이미 쓴 링크이거나 10분이 지났어요. 슬랙 휴가 앱에서 [그룹웨어 열기]를 다시 눌러 주세요.</p></div></body></html>'''


@groupware_bp.route('/gw/login/<token>')
def field_login(token):
    """슬랙 휴가 봇이 준 1회용 링크 → 현장 세션."""
    from dashboard.services import gw_login
    ip = (request.headers.get('X-Forwarded-For') or request.remote_addr or '').split(',')[0].strip()
    p = gw_login.consume(token, ip)
    if not p:
        logger.warning(f'[GW-LOGIN] 링크 거절(만료·사용함·없음) ip={ip}')
        return render_template_string(_LINK_FAIL_PAGE), 410
    session.clear()   # 세션 고정 방지
    session['user'] = gw_login.session_user(p)
    session['login_time'] = datetime.now().isoformat()
    session['session_id'] = str(uuid.uuid4())
    session.permanent = True
    logger.info(f'[GW-LOGIN] 현장 로그인 #{p["id"]} ip={ip}')
    return redirect('/home')


@groupware_bp.app_template_global()
def gw_asset(filename: str) -> str:
    """그룹웨어 정적 파일 주소 + 파일 수정 시각 버전.
    cache_buster()는 서버를 켤 때 한 번 정해져서, CSS·JS 만 고치면 브라우저가 옛 파일을 계속 씀(11-10).
    파일 시각을 붙이면 재시작 없이 고친 파일이 바로 반영된다."""
    from flask import current_app, url_for
    try:
        import os
        v = int(os.path.getmtime(os.path.join(current_app.static_folder, filename)))
    except OSError:
        v = 0
    return f"{url_for('static', filename=filename)}?v={v}"


@groupware_bp.app_context_processor
def inject_groupware_layout():
    """화면을 그릴 때 새 틀을 쓸지와 메뉴를 넣어 준다. 실패하면 기존 틀."""
    try:
        user = session.get('user') or {}
        email = user.get('email', '')
        if user.get('gw_field'):
            # 현장 세션은 늘 새 틀(기존 머리글엔 PM 메뉴가 있음), 메뉴는 현장용만
            return {'gw_layout': True, 'gw_field': True,
                    'gw_nav': gw.build_nav(request.path, 'field', field=True)}
        if not gw.can_use(email):
            return {'gw_layout': False}
        if session.get('gw_off'):
            # 시범 사용자가 「예전 화면」으로 돌린 상태 — 사용자 메뉴에 「새 화면으로」만 보임
            return {'gw_layout': False, 'gw_can_switch': True}
        from dashboard.services import hr
        return {
            'gw_layout': True,
            'gw_nav': gw.build_nav(request.path, user.get('permission_level', ''), hr.is_hr_admin(email),
                                   leave_all=hr.can_view_all_leave(email)),
        }
    except Exception as e:
        logger.warning(f'[GW] 틀 고르기 실패 → 기존 틀: {e}')
        return {'gw_layout': False}


@groupware_bp.route('/api/gw/layout', methods=['POST'])
@login_required
def set_layout():
    """본인 화면 틀 바꾸기 — {"new": true|false}. 새 틀은 쓸 수 있는 사람만."""
    data = request.get_json(silent=True) or {}
    want_new = bool(data.get('new'))
    email = (session.get('user') or {}).get('email', '')
    if want_new and not gw.can_use(email):
        return jsonify({'success': False, 'message': '아직 새 화면을 쓸 수 없는 계정입니다.'}), 403
    if want_new:
        session.pop('gw_off', None)
    else:
        session['gw_off'] = True
    session.modified = True
    return jsonify({'success': True, 'new': want_new})


@groupware_bp.route('/api/gw/pilots', methods=['GET'])
@admin_required
def get_pilots():
    return jsonify({'success': True, 'pilots': gw.list_pilots(), 'open_all': gw.is_open_all()})


@groupware_bp.route('/api/gw/pilots', methods=['POST'])
@admin_required
def change_pilots():
    """{"email": "...", "on": true|false} 또는 {"open_all": true|false}."""
    data = request.get_json(silent=True) or {}
    by = (session.get('user') or {}).get('email', 'unknown')
    try:
        if 'open_all' in data:
            gw.set_open_all(bool(data['open_all']), by)
            logger.info(f'[GW] 전체 켜기={bool(data["open_all"])} by {by}')
        else:
            email = data.get('email', '')
            if data.get('on'):
                gw.add_pilot(email, by)
            else:
                gw.remove_pilot(email)
            logger.info(f'[GW] 시범 명단 {"추가" if data.get("on") else "제외"} {email} by {by}')
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    return jsonify({'success': True, 'pilots': gw.list_pilots(), 'open_all': gw.is_open_all()})
