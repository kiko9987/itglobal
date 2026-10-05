# -*- coding: utf-8 -*-
"""홈페이지 전화 클릭 비콘(/track/call-click) — 파싱·Origin·중복·저장."""
import json
from unittest.mock import patch

import pytest
from flask import Flask

from dashboard.blueprints import call_track as ct


class _FakeRedis:
    def __init__(self):
        self.kv, self.h, self.z = {}, {}, {}

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.kv:
            return None
        self.kv[k] = v
        return True

    def pipeline(self):
        return self

    def hset(self, k, mapping):
        self.h.setdefault(k, {}).update(mapping)

    def expire(self, *a):
        pass

    def zadd(self, k, m):
        self.z.setdefault(k, {}).update(m)

    def zremrangebyscore(self, *a):
        pass

    def execute(self):
        pass

    def zrangebyscore(self, k, lo, hi):
        return [m for m, s in self.z.get(k, {}).items() if lo <= s <= hi]

    def hgetall(self, k):
        return self.h.get(k, {})


@pytest.fixture
def env():
    fake = _FakeRedis()
    app = Flask(__name__)
    app.register_blueprint(ct.call_track_bp)
    with patch.object(ct, 'get_redis_client') as g:
        g.return_value.redis = fake
        yield app.test_client(), fake


def _post(client, body, origin='https://www.itg-aircon.com', ua='UA1'):
    return client.post('/track/call-click', data=body if isinstance(body, str) else json.dumps(body),
                       headers={'Origin': origin, 'User-Agent': ua, 'Content-Type': 'text/plain;charset=UTF-8'})


def test_records_click_with_inflow(env):
    client, fake = env
    r = _post(client, {'inflow': 'naver/에어컨설치/사무실 에어컨', 'page': '/', 'btn': 'btn_6', 'mobile': True})
    assert r.status_code == 200 and r.get_json()['ok']
    assert r.headers['Access-Control-Allow-Origin'] == 'https://www.itg-aircon.com'
    rec = next(iter(fake.h.values()))
    assert rec['inflow'] == 'naver/에어컨설치/사무실 에어컨'
    assert rec['mobile'] == '1' and rec['page'] == '/' and rec['ts']
    assert len(fake.z[ct.CALL_CLICK_INDEX]) == 1


def test_same_device_within_minute_counted_once(env):
    client, fake = env
    _post(client, {'inflow': 'google'})
    r = _post(client, {'inflow': 'google'})
    assert r.get_json().get('dup') is True
    assert len(fake.h) == 1
    _post(client, {'inflow': 'google'}, ua='UA2')   # 다른 기기는 별도
    assert len(fake.h) == 2


def test_foreign_origin_rejected(env):
    client, fake = env
    r = _post(client, {'inflow': 'x'}, origin='https://evil.example')
    assert r.status_code == 403 and not fake.h


@pytest.mark.parametrize('body', ['', 'not json', '[1,2]', 'x' * 3000])
def test_bad_body_rejected(env, body):
    client, fake = env
    assert _post(client, body).status_code == 400 and not fake.h


def test_sanitizes_markup_and_length():
    rec = ct.parse_click_payload(json.dumps({'inflow': '<script>alert(1)</script>' + 'a' * 200, 'page': None}).encode())
    assert '<' not in rec['inflow'] and len(rec['inflow']) <= 90
    assert rec['page'] == '/' and rec['mobile'] == '0'


def test_load_call_clicks_reads_period(env):
    client, fake = env
    _post(client, {'inflow': 'daangn/web_install/A', 'mobile': 1})
    rows = ct.load_call_clicks(0, 9e12)
    assert rows and rows[0]['inflow'] == 'daangn/web_install/A'


def test_security_middleware_lets_beacon_through_without_csrf():
    """운영 보안 미들웨어(CSRF·세션) 하에서도 /track/ 비콘은 통과 (sendBeacon 엔 토큰 없음)."""
    from dashboard.utils.security_middleware import init_security_middleware
    fake = _FakeRedis()
    app = Flask(__name__)
    app.secret_key = 'test'
    init_security_middleware(app, 'test')
    app.register_blueprint(ct.call_track_bp)

    @app.route('/other', methods=['POST'])
    def other():
        return 'ok'

    with patch.object(ct, 'get_redis_client') as g:
        g.return_value.redis = fake
        c = app.test_client()
        assert _post(c, {'inflow': 'google'}).status_code == 200
        assert c.post('/other').status_code in (400, 403)   # 다른 POST 는 여전히 CSRF 차단
