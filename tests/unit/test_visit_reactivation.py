# -*- coding: utf-8 -*-
"""완료된 리드에 새 방문이 붙을 때 옛 완료 표시가 남는 문제 회귀 테스트 (2026-10-27 L-04129).

사고: 10-01 사진으로 방문 완료(visit_auto_completed 30일) → 10-23 전화 워크플로가 같은 번호로
10-26 방문을 L-04129 에 병합(카드·시트는 다시 '방문 예약'). 완료 표시가 남아 캔버스1에서 빠지고
사진 자동완료·[방문 완료] 버튼도 '이미 완료'로 skip.

고침: ①새 방문 카드가 나갈 때(_post_visit_notice) 완료 표시 해제
      ②K열 '방문 완료' 마커는 마지막 '방문 예약' 회차보다 뒤에 있을 때만 완료로 봄
Redis·Slack·시트 미접촉(mock).
"""
import sys

sys.path.insert(0, '.')

import pytest
import dashboard.blueprints.slack_bot as sb
from dashboard.services.lead_helpers import consult_visit_done


class _FakeRedis:
    def __init__(self):
        self.d = {}

    def get(self, k):
        return self.d.get(k)

    def set(self, k, v, ex=None, nx=False):
        if nx and k in self.d:
            return None
        self.d[k] = v
        return True

    def setex(self, k, ttl, v):
        self.d[k] = v

    def delete(self, *ks):
        return sum(1 for k in ks if self.d.pop(k, None) is not None)


@pytest.fixture
def fake_rc(monkeypatch):
    fr = _FakeRedis()
    import dashboard.utils.redis_client as rcmod
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    return fr


# ── ② K열 마커 판정 ─────────────────────────────────────────────────────────

@pytest.mark.parametrize('text,expect', [
    ('', False),
    ('-', False),
    ('[10.19 10:40 JK · 방문 예약] 실외기 이전', False),
    ('[10.19 10:40 JK · 방문 예약] 실외기 ─── [10.19 11:03 JK · 방문 완료]', True),
    ('[10.01 11:03 SH · 방문 완료]', True),                                  # 옛 형식: 마커만
    ('상담 메모만 있음 ─── [10.01 11:03 SH · 방문 완료]', True),
    # 완료 뒤 재상담으로 새 방문 예약 → 새 방문은 아직 안 끝남
    ('[09.30 10:00 SH · 방문 예약] 1차 ─── [10.01 11:03 SH · 방문 완료] '
     '─── [10.23 13:19 YM · 방문 예약] 2차 방문', False),
    # 2차도 완료
    ('[09.30 10:00 SH · 방문 예약] 1차 ─── [10.01 11:03 SH · 방문 완료] '
     '─── [10.23 13:19 YM · 방문 예약] 2차 ─── [10.26 15:00 YG · 방문 완료]', True),
    # 완료 뒤 부재중 등 다른 회차는 완료 유지
    ('[10.19 10:40 JK · 방문 예약] a ─── [10.19 11:03 JK · 방문 완료] ─── [10.21 09:00 JK · 부재중] b', True),
])
def test_consult_visit_done(text, expect):
    assert consult_visit_done(text) is expect


def test_visit_is_completed_ignores_stale_marker(fake_rc):
    k = ('[09.30 10:00 SH · 방문 예약] 1차 ─── [10.01 11:03 SH · 방문 완료] '
         '─── [10.23 13:19 YM · 방문 예약] 2차')
    assert sb._visit_is_completed('L-T1', k) is False
    fake_rc.d['visit_auto_completed:L-T1'] = '1'
    assert sb._visit_is_completed('L-T1', k) is True


# ── ① 완료 표시 해제 ─────────────────────────────────────────────────────────

def test_clear_visit_completion(fake_rc):
    fake_rc.d.update({'visit_auto_completed:L-T2': '1', 'visit_completed:L-T2': '1',
                      'visit_auto_completed:L-OTHER': '1', 'visit_notice_msg:L-T2': 'C|1'})
    sb._local_flag_nx('visit_auto_completed:L-T2', 60)
    sb._clear_visit_completion('L-T2')
    assert 'visit_auto_completed:L-T2' not in fake_rc.d
    assert 'visit_completed:L-T2' not in fake_rc.d
    assert fake_rc.d['visit_auto_completed:L-OTHER'] == '1'      # 다른 리드 그대로
    assert fake_rc.d['visit_notice_msg:L-T2'] == 'C|1'           # 카드 위치는 유지(같은 카드 갱신)
    assert not sb._local_flag_get('visit_auto_completed:L-T2')


def test_clear_visit_completion_redis_down(monkeypatch):
    import dashboard.utils.redis_client as rcmod
    def _boom():
        raise ConnectionError('redis down')
    monkeypatch.setattr(rcmod, 'get_redis_client', _boom)
    sb._local_flag_nx('visit_auto_completed:L-T3', 60)
    sb._clear_visit_completion('L-T3')                          # 예외 없이
    assert not sb._local_flag_get('visit_auto_completed:L-T3')


class _Client:
    def __init__(self):
        self.updated, self.posted = [], []

    def chat_update(self, **kw):
        self.updated.append(kw)
        return {'ok': True}

    def chat_postMessage(self, **kw):
        self.posted.append(kw)
        return {'ok': True, 'ts': '9.9'}


@pytest.fixture
def notice_env(monkeypatch, fake_rc):
    monkeypatch.setenv('SLACK_VISIT_CHANNEL', 'CVISIT')
    monkeypatch.setattr(sb, '_visit_slack_handler', object())
    monkeypatch.setattr(sb, '_visit_slack_app', None)
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda ln: {})
    monkeypatch.setattr(sb, '_slack_user_to_initial', lambda c, u: 'YM')
    monkeypatch.setattr(sb, '_build_visit_notice_blocks', lambda **kw: ('t', []))
    return fake_rc


def _notice(client, lead_no='L-04129'):
    return sb._post_visit_notice(
        client, lead_no=lead_no, category='온라인', user_id='U1', visit_date='2026-10-26',
        name='고객', contact='010-1234-5678', visit_address='주소', consultation='2차 방문',
        platform='전화')


def test_post_visit_notice_reactivation_clears_flag(notice_env):
    """L-04129 형: 완료된 카드를 같은 카드로 다시 진행 중으로 갱신 → 완료 표시 해제."""
    notice_env.d.update({'visit_auto_completed:L-04129': '1', 'visit_completed:L-04129': '1',
                         'visit_notice_msg:L-04129': 'CVISIT|1.1'})
    c = _Client()
    assert _notice(c) == ('CVISIT', '1.1')
    assert len(c.updated) == 1 and not c.posted
    assert 'visit_auto_completed:L-04129' not in notice_env.d
    assert 'visit_completed:L-04129' not in notice_env.d


def test_post_visit_notice_new_lead_no_flag(notice_env):
    c = _Client()
    assert _notice(c, 'L-NEW') == ('CVISIT', '9.9')
    assert notice_env.d['visit_notice_msg:L-NEW'] == 'CVISIT|9.9'


def test_complete_button_works_after_reactivation(notice_env, monkeypatch):
    """해제 후엔 [방문 완료] 처리가 'already done' 으로 skip 되지 않음."""
    notice_env.d['visit_auto_completed:L-04129'] = '1'
    _notice(_Client())
    class _Reached(Exception):
        pass

    def _stop(*a, **k):
        raise _Reached   # 웹훅 단계 도달 확인 후 중단 — 뒤 단계(시트·캔버스·드라이브)는 안 감
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', _stop)
    body = {'actions': [{'value': 'L-04129'}], 'channel': {'id': 'CVISIT'},
            'message': {'ts': '1.1', 'blocks': [], 'text': ''}, 'user': {'id': 'U1'}}
    with pytest.raises(_Reached):
        sb._process_visit_complete_locked(_Client(), body)


def test_complete_button_still_skips_when_done(fake_rc, monkeypatch):
    """새 방문 없이 이미 완료된 건은 그대로 skip (재클릭 방어 유지)."""
    fake_rc.d['visit_auto_completed:L-04129'] = '1'
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook',
                        lambda *a, **k: pytest.fail('완료된 건인데 webhook 호출'))
    body = {'actions': [{'value': 'L-04129'}], 'channel': {'id': 'CVISIT'},
            'message': {'ts': '1.1', 'blocks': [], 'text': ''}, 'user': {'id': 'U1'}}
    sb._process_visit_complete_locked(_Client(), body)
