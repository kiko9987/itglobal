# -*- coding: utf-8 -*-
"""방문 취소 멱등성 회귀 테스트 (이중 처리 방지).

사고: 2026-09-01 10:55 L-03768 — 매니저가 방문 취소를 수초 간격으로 두 번 제출.
  구 가드(2420d38)는 _find_lead_by_no 의 '상태'=='방문 취소' 로 판별했으나, 이
  상태는 write-behind(Redis 큐→Sheets 지연)라 1차 갱신이 시트에 전파되기 전
  2차가 stale 상태를 읽어 통과 → ① 이미 삭제된 List 항목에 webhook 재발사
  (워크플로 'Select a list item' record_not_found), ② 회색 카드 재포장(취소 헤더
  2겹·코드블록 중첩) 기형 카드.

핵심 계약 (원자적 Redis SETNX 마커 `visit_cancelled:{lead_no}` 로 대체):
  1. 1회차: 시트 update·webhook·카드 chat_update 각 1회.
  2. 2회차(같은 lead_no): **시트가 아직 stale('방문 예약')이어도** 모두 no-op.
  3. uncancel 이 마커를 삭제하면 재취소가 다시 정상 통과.
  4. Redis 불가 시 보수적 진행 (기능 끊지 않음).

fakeredis 로 SETNX 시맨틱만 검증 — 실 슬랙/시트 미접촉. 5초 action lock 은
patch 로 항상 통과시켜 **SETNX 마커 자체**를 격리 검증.
"""
import sys
sys.path.insert(0, '.')

import json
from unittest import mock

import pytest

import dashboard.blueprints.slack_bot as sb

LEAD = 'L-03768'


class FakeRedis:
    """SETNX(nx)·EX·get·delete 만 흉내내는 최소 가짜 레디스."""
    def __init__(self):
        self.store = {}
        self.raise_on = set()

    def set(self, key, val, nx=False, ex=None):
        if 'set' in self.raise_on:
            raise RuntimeError('redis down')
        if nx and key in self.store:
            return None            # 이미 존재 → SETNX 실패
        self.store[key] = val
        return True

    def get(self, key):
        return self.store.get(key)

    def delete(self, key):
        self.store.pop(key, None)
        return 1


def _make_client():
    """conversations_replies / chat_update 만 쓰는 슬랙 client 목."""
    client = mock.MagicMock()
    client.conversations_replies.return_value = {
        'messages': [{
            'blocks': [{
                'type': 'section',
                'text': {'type': 'mrkdwn', 'text': ':bell: 방문 일정 취소 — `L-03768`'},
            }],
            'text': '',
        }],
    }
    return client


def _body():
    return {'user': {'id': 'U_YM'}}


def _view(reason='타업체 공사 진행'):
    return {
        'private_metadata': json.dumps({
            'lead_no': LEAD, 'channel': 'C_VISIT', 'message_ts': '111.222',
        }),
        'state': {'values': {'reason': {'value': {'value': reason}}}},
    }


@pytest.fixture
def patched(monkeypatch):
    """공유 의존성 목 + FakeRedis 주입. side-effect 호출 카운터 반환."""
    rc = FakeRedis()

    fake_client_holder = mock.MagicMock()
    fake_client_holder.redis = rc
    monkeypatch.setattr(
        'dashboard.utils.redis_client.get_redis_client',
        lambda: fake_client_holder,
    )

    # 5초 action lock 은 항상 통과 → SETNX 마커만 격리 검증
    monkeypatch.setattr(sb, '_try_acquire_action_lock', lambda *a, **k: True)

    dispatch = mock.MagicMock()
    webhook = mock.MagicMock()
    monkeypatch.setattr(sb, '_update_lead_dispatch', dispatch)
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', webhook)
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda *a, **k: 'YM')
    # write-behind stale 재현: 취소 후에도 시트는 여전히 '방문 예약' 을 반환
    monkeypatch.setattr(
        sb, '_find_lead_by_no',
        lambda no: {'상태': '방문 예약', '상담 내용': '기존 상담', '방문 예정일': '2026-09-05'},
    )
    monkeypatch.setattr(
        'dashboard.services.visit_assignment_sync.send_visit_cancel_notification',
        lambda *a, **k: None,
    )
    return {'rc': rc, 'dispatch': dispatch, 'webhook': webhook}


def test_double_cancel_second_is_noop(patched):
    """같은 lead_no 연속 2회 취소 → 2회차는 시트·webhook·카드 모두 no-op."""
    client = _make_client()

    sb._process_visit_cancel_confirmed(client, _body(), _view('삼성전자에서 시공하기로 하심'))
    assert patched['dispatch'].call_count == 1
    assert patched['webhook'].call_count == 1
    assert client.chat_update.call_count == 1

    # 2차 취소 — 시트는 여전히 stale('방문 예약')이지만 SETNX 마커가 차단
    sb._process_visit_cancel_confirmed(client, _body(), _view('타업체 공사 진행'))
    assert patched['dispatch'].call_count == 1     # 시트 append no-op
    assert patched['webhook'].call_count == 1      # List webhook 재발사 X
    assert client.chat_update.call_count == 1       # 카드 재포장 X


def test_marker_persists_in_redis(patched):
    client = _make_client()
    sb._process_visit_cancel_confirmed(client, _body(), _view())
    assert patched['rc'].store.get(f'visit_cancelled:{LEAD}') == '1'


def test_uncancel_clears_marker_allows_recancel(patched, monkeypatch):
    """되돌리기(uncancel)가 마커를 지우면 재취소가 다시 정상 통과."""
    client = _make_client()
    sb._process_visit_cancel_confirmed(client, _body(), _view())
    assert patched['dispatch'].call_count == 1

    # uncancel 은 _build_visit_notice_blocks 등 추가 의존성 → 마커 삭제 부분만 검증
    patched['rc'].delete(f'visit_cancelled:{LEAD}')
    assert patched['rc'].store.get(f'visit_cancelled:{LEAD}') is None

    # 마커 삭제 후 재취소 → 다시 처리됨
    sb._process_visit_cancel_confirmed(client, _body(), _view('재방문 후 재취소'))
    assert patched['dispatch'].call_count == 2
    assert patched['webhook'].call_count == 2


def test_redis_down_degrades_to_process(patched):
    """Redis SETNX 예외 시 보수적 진행 (기능 끊지 않음)."""
    patched['rc'].raise_on.add('set')
    client = _make_client()
    sb._process_visit_cancel_confirmed(client, _body(), _view())
    # 마커를 못 세웠어도 취소 자체는 진행
    assert patched['dispatch'].call_count == 1
    assert patched['webhook'].call_count == 1


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
