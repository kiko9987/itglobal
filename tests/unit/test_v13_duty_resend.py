# -*- coding: utf-8 -*-
"""온라인 당번 v13 DM — 재확정 시 삭제+재발송 회귀 테스트 (2026-07-30).

배경: v9 는 2026-07-26 매니저-date 단위 삭제+재발송으로 바꿨으나 v13(온라인 당번)은
boolean flag(duty_sent)로 skip 만 해 재확정 시 참고 건수·휴무가 갱신 안 됨. v9 사상으로
통일 — dm_v13_msg:{ini}:{date}={channel,ts,hash}, 내용 변경 시에만 삭제+재발송.

_send_dms_for_next_visit 를 스텁(client/redis/resolve/email)으로 구동.
"""
import sys
sys.path.insert(0, '.')

from datetime import date, timedelta
import pytest
import dashboard.services.visit_assignment_sync as vas
import dashboard.utils.redis_client as rcmod

ITN = {'JSH': '조성헌', 'YG': '박용구', 'JW': '박정우', 'MS': '강민석', 'SJ': '빈승정'}
FUT = (date.today() + timedelta(days=2)).isoformat()


class FakeClient:
    def __init__(self):
        self.posted, self.deleted, self._t = [], [], 1000.0

    def users_lookupByEmail(self, email):
        return {'user': {'id': 'U_' + email.split('@')[0]}}

    def chat_postMessage(self, channel, text, **k):
        self._t += 1
        ts = f'{self._t:.1f}'
        self.posted.append({'channel': channel, 'text': text, 'ts': ts})
        return {'ok': True, 'channel': channel, 'ts': ts}

    def chat_delete(self, channel, ts, **k):
        self.deleted.append((channel, ts))
        return {'ok': True}


class FakeRedis:
    def __init__(self):
        self.d = {}

    def get(self, k):
        return self.d.get(k)

    def set(self, k, v, ex=None, nx=False):
        if nx and k in self.d:
            return None
        self.d[k] = v
        return True

    def delete(self, *ks):
        for k in ks:
            self.d.pop(k, None)

    def ttl(self, k):
        return 100

    def scan_iter(self, match='*', count=100):
        import fnmatch
        return [k for k in list(self.d.keys()) if fnmatch.fnmatch(k, match)]


class _FakeRC:
    def __init__(self, r):
        self.redis = r


def _lead(no, mgr=''):
    return {'리드 No': no, '방문 예정일': FUT, '고객명': no + '상호',
            '상태': '방문 예약', '영업 담당자': mgr}


@pytest.fixture
def env(monkeypatch):
    fc, fr = FakeClient(), FakeRedis()
    monkeypatch.setattr(vas, '_get_visit_client', lambda: fc)
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: _FakeRC(fr))
    monkeypatch.setattr(vas, '_email_from_initial', lambda ini, itn: f'{ini.lower()}@test.com')
    monkeypatch.setattr(vas, '_resolve_lead_for_assignment', lambda a, pm, ac: a.get('_lead'))
    return fc, fr


def _v13(fc):
    return [p for p in fc.posted if '온라인 상담 당번' in p['text']]


def _run(assignments, online_duty=('JSH',), off=('TH', 'MJ')):
    return vas._send_dms_for_next_visit(
        assignments, {}, ITN, list(online_duty), list(off),
        addr_candidates=[], online_duty_shifts={'JSH': '서포트'})


def test_new_then_same_then_changed(env):
    fc, fr = env
    a1 = [{'assign': ['YG'], '_lead': _lead('L-1', '박용구')},
          {'assign': ['JW', 'MS'], '_lead': _lead('L-2', '박정우,강민석')}]
    # 1차 — 신규 발송 + dm_v13_msg 저장
    _run(a1)
    assert len(_v13(fc)) == 1
    assert [k for k in fr.d if k.startswith('dm_v13_msg:JSH:')], 'dm_v13_msg 미저장'
    # 2차 — 동일 내용 → skip (재발송 X, 삭제 X)
    _run(a1)
    assert len(_v13(fc)) == 1, '동일 내용인데 재발송됨(스팸)'
    assert not fc.deleted
    # 3차 — combo 변경(lead 추가) → 기존 삭제 + 재발송
    a2 = a1 + [{'assign': ['SJ'], '_lead': _lead('L-3', '빈승정')}]
    _run(a2)
    assert len(_v13(fc)) == 2, '내용 변경인데 재발송 안 됨'
    assert len(fc.deleted) == 1, '기존 v13 삭제 안 됨'


def test_duty_removed_deletes_v13(env):
    fc, fr = env
    a1 = [{'assign': ['YG'], '_lead': _lead('L-1', '박용구')}]
    _run(a1, online_duty=('JSH',))
    assert len(_v13(fc)) == 1
    # 온라인 당번에서 JSH 빠짐 → 기존 v13 삭제 + 키 제거
    _run(a1, online_duty=())
    assert len(fc.deleted) == 1
    assert not [k for k in fr.d if k.startswith('dm_v13_msg:JSH:')]


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))
