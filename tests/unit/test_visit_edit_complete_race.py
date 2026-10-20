# -*- coding: utf-8 -*-
"""[정보 수정] ↔ 방문 완료(버튼·사진) 동시 처리 회귀 테스트 (2026-10-20 L-04289).

사고: 사진 자동완료(11:03:18)와 [정보 수정](11:03:15 시작)이 겹쳐, 정보 수정의 카드 갱신이
늦게 도착해 회색 완료 카드를 '진행 중'으로 덮어씀 + 시트에선 정보 수정 내용이 완료 기록에 덮임.

고침: ①같은 리드 락(visit_edit_lock) 공유 — 완료는 정보 수정이 끝날 때까지 기다림
      ②이미 완료된 리드는 [정보 수정] 시 카드를 활성으로 되살리지 않음
      ③상담 내용 끝 '방문 완료' 마커 보존
Redis·Slack·시트 미접촉(mock).
"""
import sys
import threading
import time

sys.path.insert(0, '.')

import pytest
import dashboard.blueprints.slack_bot as sb


class _FakeRedis:
    """set(nx, ex)·get·delete 만 흉내 — 스레드 안전."""
    def __init__(self):
        self._d = {}
        self._m = threading.Lock()

    def set(self, k, v, nx=False, ex=None):
        with self._m:
            if nx and k in self._d:
                return None
            self._d[k] = v
            return True

    def get(self, k):
        with self._m:
            return self._d.get(k)

    def delete(self, k):
        with self._m:
            return 1 if self._d.pop(k, None) is not None else 0


@pytest.fixture
def fake_rc(monkeypatch):
    fr = _FakeRedis()
    import dashboard.utils.redis_client as rcmod
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    return fr


# ── ③ 상담 내용 완료 마커 보존 ──────────────────────────────────────────────

def test_replace_last_consult_keeps_completion_marker():
    old = '[10.19 10:40 JK · 방문 예약] 기존 유지 및 솔루션 ─── [10.19 11:03 JK · 방문 완료]'
    new = sb._replace_last_consult_content(old, '기존 유지 및 1way 2대 또는 솔루션')
    assert new.startswith('[10.19 10:40 JK · 방문 예약] 기존 유지 및 1way 2대 또는 솔루션')
    assert new.endswith('[10.19 11:03 JK · 방문 완료]')
    assert new.count('방문 완료]') == 1


def test_replace_last_consult_unchanged_without_marker():
    old = '[10.19 10:40 JK · 방문 예약] 기존'
    assert sb._replace_last_consult_content(old, '수정') == '[10.19 10:40 JK · 방문 예약] 수정'


# ── ① 락 대기 ────────────────────────────────────────────────────────────────

def test_acquire_edit_lock_waits_until_released(fake_rc):
    assert sb._acquire_edit_lock('L-1', ttl=30) is fake_rc           # 첫 획득
    assert sb._acquire_edit_lock('L-1', ttl=30) is None              # wait=0 → 즉시 실패(기존 동작)
    threading.Timer(0.6, lambda: sb._release_edit_lock(fake_rc, 'L-1')).start()
    t0 = time.monotonic()
    assert sb._acquire_edit_lock('L-1', ttl=30, wait=5) is fake_rc   # 풀리면 획득
    assert 0.4 < time.monotonic() - t0 < 4


def test_button_complete_waits_for_running_edit(fake_rc, monkeypatch):
    """[정보 수정]이 락을 쥔 동안 [방문 완료] 본체는 시작하지 않고, 풀린 뒤 실행 + 락 해제."""
    started = []
    monkeypatch.setattr(sb, '_process_visit_complete_locked',
                        lambda client, body: started.append(time.monotonic()))
    edit_rc = sb._acquire_edit_lock('L-04289', ttl=30)               # 정보 수정 진행 중
    body = {'actions': [{'value': 'L-04289'}]}
    th = threading.Thread(target=sb._process_visit_complete, args=(None, body))
    th.start()
    time.sleep(0.8)
    assert started == []                                             # 아직 대기
    sb._release_edit_lock(edit_rc, 'L-04289')                        # 정보 수정 끝
    th.join(5)
    assert len(started) == 1
    assert fake_rc.get('visit_edit_lock:L-04289') is None            # 완료 후 락 해제


def test_photo_complete_waits_for_running_edit(fake_rc, monkeypatch):
    started = []
    monkeypatch.setattr(sb, '_auto_complete_visit_from_photo_locked',
                        lambda *a, **k: started.append(1))
    edit_rc = sb._acquire_edit_lock('L-04289', ttl=30)
    th = threading.Thread(target=sb._auto_complete_visit_from_photo,
                          args=(None, 'C1', '1.0', 'L-04289', 'U1'))
    th.start()
    time.sleep(0.8)
    assert started == []
    sb._release_edit_lock(edit_rc, 'L-04289')
    th.join(5)
    assert started == [1]
    assert fake_rc.get('visit_edit_lock:L-04289') is None


# ── ② 완료된 리드는 [정보 수정]이 카드를 되살리지 않음 ────────────────────────

class _FakeClient:
    def __init__(self):
        self.updates = []

    def chat_update(self, **kw):
        self.updates.append(kw)


def _run_edit(monkeypatch, done: bool):
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda ln: {
        '리드 No': ln, '플랫폼': '홈페이지', '방문 예정일': '2026-10-19',
        '상담 내용': '[10.19 10:40 JK · 방문 예약] 기존',
        '온라인 상담자': '강정권', '본인 방문 여부': ''})
    monkeypatch.setattr(sb, '_update_lead_dispatch', lambda *a, **k: None)
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', lambda *a, **k: None)
    monkeypatch.setattr(sb, '_normalize_visit_address_if_verified', lambda a: (a, None))
    monkeypatch.setattr(sb, '_build_visit_notice_blocks', lambda **k: ('t', []))
    monkeypatch.setattr(sb, '_visit_is_completed', lambda ln, k='': done)
    c = _FakeClient()
    sb._process_visit_edit_same_platform(
        c, {}, 'L-04289', 'C053Q2X1NP8', '1.0',
        '2026-10-19', '', '넥스토리아', '010-2567-8790', '인천 연수구 해돋이로 107', '수정')
    return c


def test_edit_on_completed_lead_does_not_resurrect_card(monkeypatch):
    assert _run_edit(monkeypatch, done=True).updates == []


def test_edit_on_active_lead_updates_card(monkeypatch):
    assert len(_run_edit(monkeypatch, done=False).updates) == 1


def test_visit_is_completed_by_flag_or_marker(fake_rc):
    assert sb._visit_is_completed('L-9', '[10.19 10:40 JK · 방문 예약] x') is False
    assert sb._visit_is_completed('L-9', 'x ─── [10.19 11:03 JK · 방문 완료]') is True
    fake_rc.set('visit_auto_completed:L-9', '1')
    assert sb._visit_is_completed('L-9', '') is True
