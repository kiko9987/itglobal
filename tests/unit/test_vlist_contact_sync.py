# -*- coding: utf-8 -*-
"""방문 List 워크플로(수정·완료·취소)는 행을 '연락처'로 찾는다 — 연락처가 바뀌면 행을 못 찾아
'목록에서 일치하는 기록 없음' 오류 + 행이 남음 (2026-10-14 L-04225 재상담으로 번호 변경 후 취소).
호출 전에 리드 번호로 행을 찾아 연락처를 맞추고, 행이 없으면 워크플로를 부르지 않는다."""
import json
import sys

sys.path.insert(0, '.')

import dashboard.blueprints.slack_bot as sb


def _item(lead, phone, rid='Rec1'):
    return {'id': rid, 'fields': [
        {'column_id': sb._VLIST_COL['lead'], 'text': lead},
        {'column_id': sb._VLIST_CONTACT_COL, 'value': phone, 'phone': [phone]},
    ]}


class _Resp:
    def __init__(self, data):
        self.data = data


def _patch(monkeypatch, items, updates):
    monkeypatch.setenv('SLACK_VISIT_BOT_TOKEN', 'xoxb-test')
    monkeypatch.setenv('SLACK_VISIT_LIST_ID', 'L1')

    class _Client:
        def __init__(self, token):
            pass

        def api_call(self, method, http_verb='GET', params=None):
            assert method == 'slackLists.items.list'
            return _Resp({'items': items})

    monkeypatch.setattr('slack_sdk.WebClient', _Client)

    class _R:
        def __init__(self, body):
            self.body = body

        def read(self):
            return json.dumps({'ok': True}).encode()

        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def _urlopen(req, timeout=0):
        updates.append(json.loads(req.data))
        return _R(req.data)

    monkeypatch.setattr('urllib.request.urlopen', _urlopen)


def test_contact_changed_row_gets_updated(monkeypatch):
    updates = []
    _patch(monkeypatch, [_item('L-04225', '010-2983-0420')], updates)
    assert sb._vlist_sync_row_contact('L-04225', '010-9131-2240') is True
    assert updates == [{'list_id': 'L1', 'cells': [
        {'row_id': 'Rec1', 'column_id': sb._VLIST_CONTACT_COL, 'phone': ['010-9131-2240']}]}]


def test_same_contact_no_update(monkeypatch):
    updates = []
    _patch(monkeypatch, [_item('L-04225', '010-9131-2240')], updates)
    assert sb._vlist_sync_row_contact('L-04225', '01091312240') is True
    assert updates == []


def test_row_missing_returns_false(monkeypatch):
    updates = []
    _patch(monkeypatch, [_item('L-04000', '010-1111-2222')], updates)
    assert sb._vlist_sync_row_contact('L-04225', '010-9131-2240') is False


def test_webhook_skipped_when_row_missing(monkeypatch):
    """행이 이미 없으면(중복 호출·이미 정리됨) 워크플로를 안 불러 '일치 기록 없음' 오류 DM 이 안 생김."""
    sent = []
    monkeypatch.setenv('SLACK_VISIT_CANCEL_WEBHOOK_URL', 'https://hooks.slack.test/cancel')
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda no: {'고객 연락처': '010-9131-2240', '플랫폼': '전화'})
    monkeypatch.setattr(sb, '_vlist_sync_row_contact', lambda no, c: False)
    monkeypatch.setattr('urllib.request.urlopen', lambda req, timeout=0: sent.append(req.full_url))
    monkeypatch.setattr('dashboard.services.visit_canvas_sync.rebuild_canvas_async', lambda: None)
    sb._trigger_visit_list_webhook('SLACK_VISIT_CANCEL_WEBHOOK_URL', 'L-04225', '', '')
    assert sent == []
