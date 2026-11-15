# -*- coding: utf-8 -*-
"""회사 정보 공개 페이지(/company·/llms.txt) 점검 — 2026-11-15

아임웹이 AI 크롤러를 막고 있어 우리 서버에 둔 창구다. 두 가지가 깨지면 안 된다.
  1. 로그인 없이 열려야 한다(AI 봇은 세션이 없다)
  2. 고객 정보·내부 수치가 절대 섞이면 안 된다(로그인 없는 경로이므로)
"""
import json
import re

import pytest
from flask import Flask

from dashboard.blueprints.public_info import public_info_bp, LLMS_TXT


@pytest.fixture()
def client():
    # create_app() 을 쓰지 않는다(스케줄러 이중 기동 방지). 블루프린트만 빈 앱에 올린다.
    app = Flask(__name__)
    app.register_blueprint(public_info_bp)
    return app.test_client()


@pytest.mark.parametrize('path, ctype', [
    ('/company', 'text/html'),
    ('/company.md', 'text/markdown'),
    ('/llms.txt', 'text/plain'),
    ('/robots.txt', 'text/plain'),
])
def test_공개_경로는_로그인_없이_열린다(client, path, ctype):
    r = client.get(path)
    assert r.status_code == 200
    assert ctype in r.headers['Content-Type']
    assert 'charset=utf-8' in r.headers['Content-Type']


def test_구조화_데이터가_유효하다(client):
    html = client.get('/company').get_data(as_text=True)
    m = re.search(r'<script type="application/ld\+json">(.*?)</script>', html, re.S)
    assert m, 'JSON-LD 블록이 없다'
    data = json.loads(m.group(1))          # 깨지면 여기서 실패
    assert 'HVACBusiness' in data['@type']
    assert data['telephone'] == '+82-1899-1377'
    assert '서울특별시' in data['areaServed']


def test_사실이_홈페이지_llms와_어긋나지_않는다():
    # 두 곳이 다르면 AI 답변이 흔들린다. 핵심 숫자만 지킨다.
    assert '20년' in LLMS_TXT and '1만 건' in LLMS_TXT
    assert '17년' not in LLMS_TXT and '9,000' not in LLMS_TXT   # 옛 숫자 유입 방지
    assert '하자보증 2년' in LLMS_TXT
    assert '1899-1377' in LLMS_TXT


def test_로그인_없는_경로에_내부_정보가_없다(client):
    html = client.get('/company').get_data(as_text=True)
    md = client.get('/company.md').get_data(as_text=True)
    금지 = ('고객명', '연락처 :', '리드 No', '수금', '미수금', '매출', '원가', '세콤', '슬랙')
    for 말 in 금지:
        assert 말 not in html, f'공개 페이지에 {말} 이 들어 있다'
        assert 말 not in md, f'공개 마크다운에 {말} 이 들어 있다'


def test_robots는_회사정보만_열고_나머지를_막는다(client):
    body = client.get('/robots.txt').get_data(as_text=True)
    assert 'Allow: /company' in body
    assert 'Allow: /llms.txt' in body
    assert 'Disallow: /' in body          # PM 화면은 색인 금지
