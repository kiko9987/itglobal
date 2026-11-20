"""그룹웨어 새 틀 — 화면 틀 고르기와 시범 명단 API.

- 모든 화면이 쓰는 modern_base.html 이 `gw_layout` 값으로 머리글만 바꿔 끼운다.
  명단에 없는 사람에겐 gw_layout=False → 기존 머리글 그대로.
- API 는 /api/ 아래(CSRF 경로 규칙).
"""
import re
import uuid
from datetime import datetime
from typing import Optional

from flask import Blueprint, g, jsonify, redirect, render_template_string, request, session

from dashboard.auth import admin_required, login_required
from dashboard.services import groupware as gw
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

groupware_bp = Blueprint('groupware', __name__)

# ── 현장 세션 문지기 ───────────────────────────────────────────────────
# 슬랙 링크로 들어온 현장·미화 직원(session.user.gw_field)은 아래 경로만. 나머지(기존 PM 화면·API)는 전부 막는다.
_FIELD_PAGES = ('/home', '/hr/me', '/leave', '/yearend', '/logout', '/login')
_FIELD_API = [
    (('GET',), re.compile(r'^/api/home$')),              # 홈(현장은 내 휴가만 담김)
    (('GET',), re.compile(r'^/api/hr/people/\d+$')),     # 본인 카드(본인 여부는 화면 쪽에서 다시 확인)
    (('GET',), re.compile(r'^/api/leave/me$')),          # 내 휴가(+신청 창 정보)
    (('POST',), re.compile(r'^/api/leave/apply$')),      # 본인 휴가 신청(본인 여부는 leave_web 이 슬랙 ID 로 다시 확인)
    (('POST',), re.compile(r'^/api/leave/\d+/cancel$')), # 본인 휴가 취소(같음)
    (('GET',), re.compile(r'^/api/hr/documents/\d+$')),  # 본인 서류 내려받기(본인 여부는 화면 쪽에서 다시 확인)
    (('POST',), re.compile(r'^/api/hr/me/(contact|family|documents)$')),   # 내 정보 직접 고치기(11-20)
    (('POST',), re.compile(r'^/api/hr/me/cert-request(/\d+/cancel)?$')),      # 증명서 신청·취소(11-20)
    (('GET', 'POST'), re.compile(r'^/api/yearend/me(/submit|/reopen|/documents)?$')),   # 연말정산 본인 것만
    (('DELETE',), re.compile(r'^/api/yearend/me/documents/\d+$')),
    (('POST',), re.compile(r'^/api/auth/ping$')),       # 세션 연장
    (('GET',), re.compile(r'^/api/gw/notify$')),         # 알림 🔔(휴가·연말정산만 담김)
    (('POST',), re.compile(r'^/api/gw/notify/seen$')),
]


def field_path_allowed(method: str, path: str) -> bool:
    if path.startswith('/static/') or path == '/favicon.ico' or path.startswith('/gw/login/') \
            or path == '/manifest.webmanifest' or path == '/gw/device/redeem':
        return True
    if path in _FIELD_PAGES:
        return True
    return any(method in methods and pat.match(path) for methods, pat in _FIELD_API)


# ── 현장 휴대폰 등록(기기 열쇠) — 플렉스 앱처럼 한 번 등록하면 계속(2026-11-15) ─────────────
# 문지기보다 먼저 돈다(같은 블루프린트 안에서 적은 순서). services/gw_device.py 참고.
def _start_field_session(p: dict, slack_id: str, device_id: Optional[int] = None) -> None:
    from dashboard.services import gw_login
    from dashboard.auth import rotate_session_id
    session.clear()   # 세션 고정 방지
    session['user'] = dict(gw_login.session_user(dict(p, link_slack_id=slack_id)), gw_device_id=device_id)
    session['login_time'] = datetime.now().isoformat()
    session['session_id'] = str(uuid.uuid4())
    session.permanent = True
    rotate_session_id()


@groupware_bp.before_app_request
def restore_device_session():
    from dashboard.services import gw_device
    tok = request.cookies.get(gw_device.COOKIE)
    if not tok or request.path.startswith('/static/'):
        return None
    try:
        if request.path == '/logout':
            # 그 휴대폰에서 로그아웃 = 등록 해제(안 그러면 다음 화면에서 바로 다시 들어옴)
            gw_device.revoke_token(tok, 'logout')
            g.gw_device_clear = True
            return None
        user = session.get('user') or {}
        if user and not user.get('gw_field'):
            return None                       # 회사 계정으로 로그인한 브라우저 — 손대지 않음
        info = gw_device.check(tok)
        if not info:
            g.gw_device_clear = True          # 끊김·오래 안 씀·퇴사 → 열쇠 지우고, 그 열쇠로 연 세션도 닫음
            if user.get('gw_device_id'):
                session.clear()
            return None
        p, d = info['person'], info['device']
        if user and user.get('gw_person_id') != p['id']:
            return None                       # 다른 현장 직원이 링크로 들어와 있는 브라우저
        if not user or user.get('gw_device_id') != d['id']:
            _start_field_session(p, d.get('slack_id') or '', d['id'])
            logger.info(f'[GW-DEVICE] 휴대폰으로 현장 세션 #{p["id"]} device#{d["id"]}')
        else:
            # 4시간 세션 끝(SESSION_TIMEOUT_HOURS)에 걸리지 않게 — 한 시간마다 로그인 시각을 당겨 둔다
            try:
                old = datetime.fromisoformat(session.get('login_time') or '')
            except ValueError:
                old = None
            if not old or datetime.now() - old > gw_device.TOUCH_EVERY:
                session['login_time'] = datetime.now().isoformat()
        if gw_device.needs_touch(d):
            gw_device.touch(d['id'], request.remote_addr or '')
            g.gw_device_set = tok             # 쿠키 기한도 다시 90일로
    except Exception as e:
        logger.warning(f'[GW-DEVICE] 휴대폰 열쇠 확인 실패(그냥 진행): {e}')
    return None


@groupware_bp.after_app_request
def write_device_cookie(resp):
    from dashboard.services import gw_device
    from flask import current_app
    if getattr(g, 'gw_device_clear', False):
        resp.delete_cookie(gw_device.COOKIE, path='/')
    elif getattr(g, 'gw_device_set', None):
        resp.set_cookie(gw_device.COOKIE, g.gw_device_set, max_age=gw_device.IDLE_DAYS * 86400, path='/',
                        httponly=True, samesite='Lax', secure=bool(current_app.config.get('SESSION_COOKIE_SECURE')))
    return resp


@groupware_bp.route('/gw/device/redeem', methods=['POST'])
def device_redeem():
    """로그인 화면 「번호로 로그인」 — 6자리 번호 → 이 휴대폰 등록 + 현장 세션.
    로그인 전이라 CSRF 토큰이 없다 → 보안 미들웨어가 같은 사이트에서 보낸 요청인지만 보고 통과시킨다."""
    from dashboard.services import gw_device
    data = request.get_json(silent=True) or {}
    ip = request.remote_addr or ''
    p, token, why = gw_device.redeem(data.get('code', ''), ip, request.headers.get('User-Agent', ''))
    if why:
        return jsonify({'success': False, 'message': why}), 400
    info = gw_device.check(token)
    _start_field_session(p, (p.get('slack_id') or '').strip(), info['device']['id'] if info else None)
    g.gw_device_set = token
    return jsonify({'success': True, 'redirect': '/leave', 'name': p['name']})


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
<link rel="manifest" href="/manifest.webmanifest">
<style>body{font-family:-apple-system,'Segoe UI',sans-serif;background:#f8fafc;margin:0;padding:48px 16px;color:#0f172a}
.box{max-width:420px;margin:0 auto;background:#fff;border:1px solid #e2e8f0;border-radius:12px;padding:24px}
h1{font-size:1.15rem;margin:0 0 8px}p{color:#475569;line-height:1.6;margin:0}</style></head>
<body><div class="box"><h1>링크를 쓸 수 없어요</h1>
<p>이미 쓴 링크이거나 10분이 지났어요. 슬랙 휴가 앱에서 [그룹웨어 열기]를 다시 눌러 주세요.</p></div></body></html>'''


@groupware_bp.route('/gw/login/<token>')
def field_login(token):
    """슬랙 휴가 봇이 준 1회용 링크 → 현장 세션."""
    from dashboard.services import gw_login
    ip = request.remote_addr or ''   # 프록시 처리된 실제 주소 — 손님이 보낸 X-Forwarded-For 첫 값은 믿지 않음 (2026-11-13)
    p = gw_login.consume(token, ip)
    if not p:
        logger.warning(f'[GW-LOGIN] 링크 거절(만료·사용함·없음) ip={ip}')
        return render_template_string(_LINK_FAIL_PAGE), 410
    session.clear()   # 세션 고정 방지
    session['user'] = gw_login.session_user(p)
    session['login_time'] = datetime.now().isoformat()
    session['session_id'] = str(uuid.uuid4())
    session.permanent = True
    from dashboard.auth import rotate_session_id
    rotate_session_id()   # 세션 ID 새로 발급 (2026-11-13)
    logger.info(f'[GW-LOGIN] 현장 로그인 #{p["id"]} ip={ip}')
    return redirect('/leave')   # 현장 직원은 바로 내 휴가(11-15)


# ── 홈 화면 앱(웹 앱 설치, 2026-11-15) ────────────────────────────────
# 휴대폰 '홈 화면에 추가' → 주소창 없이 앱처럼 열린다. 스토어 앱 아님.
# 서비스 워커는 쓰지 않는다 — 2026-07-08 Cache First 워커가 배포마다 옛 화면을 보여 줘서 일부러 끔(sw.js 머리말).
# 요즘 안드로이드 크롬·iOS 사파리는 워커 없이도 설치 정보(manifest)만으로 주소창 없는 앱으로 연다.
# 로그인 없이 읽혀야 한다(브라우저가 쿠키 없이 가져감) — 회사 이름·아이콘만 담는다.
@groupware_bp.route('/manifest.webmanifest')
def web_app_manifest():
    from flask import Response
    import json as _json
    body = {
        'id': '/home',
        'name': '아이티글로벌',
        'short_name': '아이티글로벌',
        'description': '아이티글로벌 그룹웨어',
        'lang': 'ko',
        'start_url': '/home',
        'scope': '/',
        'display': 'standalone',
        'background_color': '#ffffff',
        'theme_color': '#ffffff',
        'icons': [
            {'src': gw_asset('images/favicon-192.png'), 'sizes': '192x192', 'type': 'image/png', 'purpose': 'any'},
            {'src': gw_asset('images/icon-512.png'), 'sizes': '512x512', 'type': 'image/png', 'purpose': 'any'},
            {'src': gw_asset('images/apple-touch-icon.png'), 'sizes': '180x180', 'type': 'image/png', 'purpose': 'any'},
        ],
    }
    resp = Response(_json.dumps(body, ensure_ascii=False), mimetype='application/manifest+json')
    resp.headers['Cache-Control'] = 'no-cache'
    return resp


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


@groupware_bp.route('/api/gw/search', methods=['GET'])
@login_required
def quick_search():
    """빠른 찾기(위쪽 검색창·/ 키) — 프로젝트·리드. ?q=검색어. 현장 세션은 문지기에서 막힘(메뉴만 찾음).
    검색어는 로그에 남기지 않는다."""
    from dashboard.services import gw_search
    result = gw_search.search(request.args.get('q', ''))
    return jsonify({'success': True, **result})


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
