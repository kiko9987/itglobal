"""세콤매니저 REST 전송 받기 (2026-11-06) — 같은 PC만 받고, 받은 요청을 그대로 남긴다."""
import json
import os

from flask import Flask

from dashboard.blueprints.secom_inbound import secom_bp


def _app(tmp_path):
    app = Flask(__name__, instance_path=str(tmp_path))
    app.register_blueprint(secom_bp)
    return app


def test_local_request_is_captured(tmp_path):
    c = _app(tmp_path).test_client()
    body = {'at': 1793925000, 'sabun': '6703129000607', 'type': 1}
    r = c.post('/secom/v1/secom/manager/alarm', json=body, environ_base={'REMOTE_ADDR': '127.0.0.1'})
    assert r.status_code == 200 and r.get_json()['success'] is True
    files = os.listdir(tmp_path / 'secom')
    rec = json.loads(open(tmp_path / 'secom' / files[0], encoding='utf-8').read().splitlines()[0])
    assert rec['path'] == 'v1/secom/manager/alarm' and json.loads(rec['body']) == body


def test_root_path_and_get(tmp_path):
    c = _app(tmp_path).test_client()
    assert c.get('/secom', environ_base={'REMOTE_ADDR': '127.0.0.1'}).status_code in (200, 308)
    assert c.get('/secom/', environ_base={'REMOTE_ADDR': '127.0.0.1'}).status_code == 200


def test_outside_or_proxied_request_refused(tmp_path):
    c = _app(tmp_path).test_client()
    assert c.post('/secom/x', json={}, environ_base={'REMOTE_ADDR': '1.2.3.4'}).status_code == 403
    # Caddy 를 거친 바깥 요청은 127.0.0.1 로 보여도 X-Forwarded-For 가 붙어 있다
    r = c.post('/secom/x', json={}, environ_base={'REMOTE_ADDR': '127.0.0.1'},
               headers={'X-Forwarded-For': '1.2.3.4'})
    assert r.status_code == 403
    assert not (tmp_path / 'secom').exists()
