# -*- coding: utf-8 -*-
"""문서변환봇(별도 슬랙 앱) 진입점 테스트 (2026-10-22).

- /slack/docview-events: 토큰 넣기 전 앱 생성 단계의 url_verification 에 challenge 응답,
  그 밖의 요청은 미설정이면 503.
- 메시지 핸들러: 파일은 문서 보기로 넘기고, DM 에 글만 오면 사용 안내, 봇·수정 이벤트는 무시.
"""
import sys

from flask import Flask

sys.path.insert(0, '.')

from dashboard.blueprints import slack_bot as sb
from dashboard.services import doc_viewer as dv


def _client(monkeypatch):
    monkeypatch.setattr(sb, '_docview_slack_handler', None)
    monkeypatch.setattr(sb, '_init_docview_slack_app', lambda: False)
    app = Flask(__name__)
    app.register_blueprint(sb.slack_bp)
    return app.test_client()


def test_url_verification_answered_before_token_is_set(monkeypatch):
    c = _client(monkeypatch)
    r = c.post('/slack/docview-events', json={'type': 'url_verification', 'challenge': 'abc123'})
    assert r.status_code == 200 and r.get_json() == {'challenge': 'abc123'}


def test_other_requests_rejected_until_configured(monkeypatch):
    c = _client(monkeypatch)
    r = c.post('/slack/docview-events', json={'type': 'event_callback', 'event': {}})
    assert r.status_code == 503


class _FakeApp:
    def __init__(self):
        self.handlers = {}

    def event(self, name):
        def deco(fn):
            self.handlers[name] = fn
            return fn
        return deco


def _message_handler():
    app = _FakeApp()
    sb._register_docview_handlers(app)
    return app.handlers['message']


def test_message_handler_routes_files_and_guides_dm_text(monkeypatch):
    started = []
    monkeypatch.setattr(dv, 'start_doc_view',
                        lambda e, c, t: started.append(e) or dv.doc_view_mode(e))
    handler = _message_handler()
    said = []

    # DM 파일 → 문서 보기, 안내 없음
    handler({'channel_type': 'im', 'subtype': 'file_share', 'ts': '1',
             'files': [{'name': '공고.pdf'}]}, said.append, None)
    assert len(started) == 1 and said == []

    # DM 글 → 사용 안내
    handler({'channel_type': 'im', 'text': '안녕', 'ts': '2'}, said.append, None)
    assert said == [dv.DM_GUIDE]

    # 채널 글·DM 수정 이벤트·봇 메시지 → 아무 답 없음
    handler({'channel_type': 'channel', 'text': '안녕', 'ts': '3'}, said.append, None)
    handler({'channel_type': 'im', 'subtype': 'message_changed', 'ts': '4'}, said.append, None)
    handler({'channel_type': 'im', 'bot_id': 'B1', 'text': 'x', 'ts': '5'}, said.append, None)
    assert said == [dv.DM_GUIDE]
    assert len(started) == 4          # 봇 메시지만 문서 보기까지 가지 않음
