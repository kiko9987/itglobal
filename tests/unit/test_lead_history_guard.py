# -*- coding: utf-8 -*-
"""상담 이력(K열) 덮어쓰기 막기 (2026-11-16 상태 머신 점검 묶음 F 1단계).

리드를 시트에서 못 읽으면 `{}` 로 보고 이어 쓰던 세 곳 — 방문 취소·방문 [정보 수정]·[기존 lead 연결] —
은 상담 이력을 새 한 줄로 덮어썼다(10-13 Redis 장애 때 재상담·방문 완료에서 12건 난 것과 같은 유형).
이제 못 읽으면 아무것도 쓰지 않고(카드·List·마커도 그대로) 다시 눌러 달라고 안내한다.
"""
import sys

sys.path.insert(0, '.')

import pandas as pd
import pytest

import dashboard.blueprints.slack_bot as sb


class _RC:
    def __init__(self):
        self.d = {}

    def get(self, k):
        return self.d.get(k)

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.d:
            return False
        self.d[k] = v
        return True

    def delete(self, *ks):
        for k in ks:
            self.d.pop(k, None)


@pytest.fixture
def rc(monkeypatch):
    import dashboard.utils.redis_client as rcm
    r = _RC()
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': r})())
    return r


class _Client:
    def __init__(self):
        self.ephemeral, self.updates, self.posts = [], [], []

    def chat_postEphemeral(self, **k):
        self.ephemeral.append(k['text'])

    def chat_update(self, **k):
        self.updates.append(k)

    def chat_postMessage(self, **k):
        self.posts.append(k)

    def conversations_replies(self, **k):
        return {'messages': [{'text': '카드'}]}

    def reactions_add(self, **k):
        pass


def _must_not(name):
    def _f(*a, **k):
        raise AssertionError(f'{name} 를 부르면 안 됨')
    return _f


# ── _find_lead_by_no(strict) ──

def test_find_lead_strict_separates_failure_from_absence(monkeypatch):
    import dashboard.services.lead_service as ls

    def boom(force_refresh=False):
        raise RuntimeError('Redis 장애')
    monkeypatch.setattr(ls, 'load_leads_data', boom)
    assert sb._find_lead_by_no('L-00001') is None                      # 예전 동작 그대로
    with pytest.raises(sb.LeadReadError):
        sb._find_lead_by_no('L-00001', strict=True)

    df = pd.DataFrame([{'리드 No': 'L-00002', '상담 내용': '이력'}])
    monkeypatch.setattr(ls, 'load_leads_data', lambda force_refresh=False: df)
    assert sb._find_lead_by_no('L-00001', strict=True) is None         # 읽었는데 없음
    assert sb._find_lead_by_no('L-00002', strict=True)['상담 내용'] == '이력'

    monkeypatch.setattr(ls, 'load_leads_data', lambda force_refresh=False: pd.DataFrame())
    with pytest.raises(sb.LeadReadError):                              # 빈 시트 = 못 읽음
        sb._find_lead_by_no('L-00002', strict=True)


# ── 방문 취소 ──

def _cancel_view():
    import json
    return {'private_metadata': json.dumps({'lead_no': 'L-04300', 'channel': 'C_VISIT', 'message_ts': '1.2'}),
            'state': {'values': {'reason': {'reason': {'value': '고객 일정 변경'}}}}}


def test_visit_cancel_stops_when_lead_unreadable(monkeypatch, rc):
    monkeypatch.setattr(sb, '_try_acquire_action_lock', lambda *a: True)
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda *a, **k: None)
    monkeypatch.setattr(sb, '_update_lead_dispatch', _must_not('시트 쓰기'))
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', _must_not('List 웹훅'))
    c = _Client()
    sb._process_visit_cancel_confirmed(c, {'user': {'id': 'U_JW'}}, _cancel_view())
    assert len(c.ephemeral) == 1 and '방문 취소를 처리하지 않았습니다' in c.ephemeral[0]
    assert c.updates == []                                              # 카드 그대로(버튼 유지)
    assert 'visit_cancelled:L-04300' not in rc.d                        # 다시 누를 수 있게 마커 해제


def test_visit_cancel_appends_history_when_readable(monkeypatch, rc):
    monkeypatch.setattr(sb, '_try_acquire_action_lock', lambda *a: True)
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda *a, **k: {'상담 내용': '[11.10 10:00 JW · 방문 예약] 첫 상담'})
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda *a: 'JW')
    wrote = []
    monkeypatch.setattr(sb, '_update_lead_dispatch', lambda no, upd: wrote.append(upd))
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', lambda *a, **k: None)
    try:
        sb._process_visit_cancel_confirmed(_Client(), {'user': {'id': 'U_JW'}}, _cancel_view())
    except Exception:
        pass                                                           # 뒤쪽 카드·DM 단계는 이 테스트 범위 밖
    assert wrote and wrote[0]['상담 내용'].startswith('[11.10 10:00 JW · 방문 예약] 첫 상담\n─────────\n[방문 취소')


# ── 방문 [정보 수정] ──

def test_visit_edit_stops_when_lead_unreadable(monkeypatch):
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda *a, **k: None)
    monkeypatch.setattr(sb, '_update_lead_dispatch', _must_not('시트 쓰기'))
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', _must_not('List 웹훅'))
    told = []
    monkeypatch.setattr(sb, '_notify_edit_failure', lambda client, ch, user, no, reason: told.append(reason))
    sb._process_visit_edit_same_platform(
        _Client(), {'user': {'id': 'U_JW'}}, 'L-04300', 'C_VISIT', '1.2',
        '2026-11-20', '', '홍길동', '010-1234-5678', '', '수정한 상담')
    assert told == ['리드를 시트에서 읽지 못해 아무것도 저장하지 않았습니다']


# ── [기존 lead 연결] ──

@pytest.fixture
def link(monkeypatch, rc):
    import dashboard.blueprints.channeltalk as ct
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ct, '_get_chat_lead_no', lambda chat_id: 'L-04400')
    monkeypatch.setattr(ls, 'get_lead_by_no', lambda no: {'리드 No': no, '키워드': '천장형', '온라인 상담자': '김지원'})
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda *a: 'JW')
    greyed = []
    monkeypatch.setattr(sb, '_grey_out_merged_chat_card', lambda *a, **k: greyed.append(a))
    c = _Client()

    def _run():
        sb._link_chat_to_existing_lead(c, 'chat1', 'L-04100', 'C_CT', '9.9', slack_user_id='U_JW')
        return c, greyed
    return _run


def test_link_stops_when_chat_lead_unreadable(monkeypatch, rc, link):
    import dashboard.services.lead_service as ls

    def boom(*a, **k):
        raise sb.LeadReadError('502')
    monkeypatch.setattr(sb, '_find_lead_by_no', boom)
    monkeypatch.setattr(ls, 'update_lead', _must_not('시트 쓰기'))
    c, greyed = link()
    assert '읽지 못해 통합하지 않았습니다' in c.ephemeral[0]
    assert greyed == [] and 'linked_chat:chat1' not in rc.d            # 카드·재시도 마커 그대로


def test_link_stops_when_marking_save_fails(monkeypatch, rc, link):
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda *a, **k: {'상담 내용': '채팅 상담 이력'})
    monkeypatch.setattr(ls, 'update_lead', lambda no, upd: {'success': False, 'message': '시트 502'})
    c, greyed = link()
    assert '시트 저장에 실패해 통합하지 않았습니다' in c.ephemeral[0]
    assert greyed == [] and 'linked_chat:chat1' not in rc.d


def test_link_keeps_history_and_completes(monkeypatch, rc, link):
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda *a, **k: {'상담 내용': '채팅 상담 이력'})
    wrote = []
    monkeypatch.setattr(ls, 'update_lead', lambda no, upd: wrote.append((no, upd)) or {'success': True})
    c, greyed = link()
    assert wrote[0][0] == 'L-04400' and wrote[0][1]['상담 내용'] == '채팅 상담 이력\n\n→ L-04100 로 통합'
    assert greyed and rc.d.get('linked_chat:chat1') == 'L-04100'


def test_link_absent_chat_lead_still_links_without_marking(monkeypatch, rc, link):
    """시트에 채팅 리드가 정말 없으면 지킬 이력이 없다 — 예전처럼 통합은 진행."""
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda *a, **k: None)
    monkeypatch.setattr(ls, 'update_lead', _must_not('없는 리드 시트 쓰기'))
    c, greyed = link()
    assert greyed and rc.d.get('linked_chat:chat1') == 'L-04100' and c.ephemeral == []


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
