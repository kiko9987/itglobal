"""본문 없는 POST/DELETE 에 JSON 머리글을 붙이면 운영 보안 미들웨어가 400 — 그룹웨어 화면은 붙이지 않는다(2026-11-10).

운영에서 인사카드 [전체 보기]가 '요청을 처리하지 못했습니다 (400)' 로 실패(로그: JSON 파싱 오류). 시험 서버엔 이 미들웨어가 없어 못 잡았다.
"""
import os
import re

from flask import Flask, jsonify, session

from dashboard.utils.security_middleware import SecurityMiddleware

JS = os.path.join(os.path.dirname(__file__), '..', '..', 'dashboard', 'static', 'js')


def _app():
    app = Flask('t')
    app.secret_key = 't'
    mw = SecurityMiddleware(app, 't')
    app.before_request(mw.process_request)

    @app.route('/api/x', methods=['POST', 'DELETE'])
    def x():
        return jsonify({'success': True})

    @app.route('/login-as')
    def login_as():
        session['user'] = {'email': 'kiko@itg-aircon.com'}
        return 'ok'
    return app


def test_middleware_rejects_json_header_without_body_but_accepts_bare():
    c = _app().test_client()
    c.get('/login-as')
    assert c.post('/api/x', headers={'Content-Type': 'application/json'}).status_code == 400   # 예전 방식
    assert c.post('/api/x').status_code == 200                                                   # 고친 방식
    assert c.delete('/api/x').status_code == 200
    assert c.post('/api/x', json={'a': 1}).status_code == 200                                    # 본문 있으면 그대로


def test_groupware_js_helpers_send_json_header_only_with_body():
    """그룹웨어 화면 JS 의 api() 가 본문 없을 때 JSON 머리글을 안 붙이는지(같은 실수 재발 방지)."""
    for name in os.listdir(JS):
        if not (name.startswith('gw-') and name.endswith('.js')):
            continue
        src = open(os.path.join(JS, name), encoding='utf-8').read()
        for m in re.finditer(r"function api\([^)]*\)\s*\{(.*?)\n  \}", src, re.S):
            body = m.group(1)
            if "'Content-Type': 'application/json'" not in body:
                continue
            # 머리글이 본문 조건 안에서만 붙어야 함 — POST 고정(admin: body 있을 때만 POST)도 허용
            ok = re.search(r"if \((opts\.)?body\)\s*\{[^}]*Content-Type", body) or \
                 re.search(r"else if \(opts\.body\)\s*\{[^}]*Content-Type", body) or \
                 "method: body ? 'POST' : 'GET'" in body
            assert ok, f'{name}: 본문 없는 요청에도 JSON 머리글을 붙임'
