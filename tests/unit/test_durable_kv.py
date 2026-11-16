# -*- coding: utf-8 -*-
"""Redis 에만 있던 업무 기록의 DB 사본 + 빈 Redis 감지 (2026-11-16 상태 머신 점검 묶음 D).

- 입금 인입·계산서 요청 카드·금액/매출 0원 요청은 Redis 에 쓰고 지울 때 DB 사본도 같이.
- Redis 세대 표시가 사라졌는데 DB 엔 있으면 = Redis 가 비워짐 → 아직 유효한 사본만 되채움(처리 끝나 지운 건 안 살림).
- Redis 장애 중에도 은행 문자는 받는다(예전엔 HTTP 500).
- pytest 중엔 기본으로 꺼져 운영 users.db 에 쓰지 않는다.
"""
import json
import sys
import time

sys.path.insert(0, '.')

import pytest

from dashboard.utils import durable_kv as dk


class _FakeRedis:
    def __init__(self):
        self.d = {}

    def get(self, k):
        v = self.d.get(k)
        return v[0] if v and (v[1] is None or v[1] > time.time()) else None

    def set(self, k, v, nx=False, ex=None):
        if nx and self.get(k) is not None:
            return False
        self.d[k] = (v, time.time() + ex if ex else None)
        return True

    def setex(self, k, ex, v):
        return self.set(k, v, ex=ex)

    def delete(self, *ks):
        for k in ks:
            self.d.pop(k, None)

    def exists(self, k):
        return int(self.get(k) is not None)


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(dk, '_db', lambda: d)
    monkeypatch.setattr(dk, '_table_ready', False)
    monkeypatch.setenv('DURABLE_KV_TEST', '1')
    return d


def _rows(d):
    with d._get_connection() as conn:
        return {r[0]: r[1] for r in conn.execute('SELECT key, value FROM redis_durable')}


def test_disabled_under_pytest_by_default(db, monkeypatch):
    monkeypatch.delenv('DURABLE_KV_TEST')
    dk.put('sms_intake:a', '{}', 60)
    assert not dk.enabled()
    with db._get_connection() as conn:
        assert not conn.execute("SELECT 1 FROM sqlite_master WHERE name='redis_durable'").fetchone()


def test_only_protected_prefixes_and_not_locks(db):
    dk.put('sms_intake:a', '{"text":"x"}', 60)
    dk.put('project_amount_req:C:1.0:proc', '1', 60)          # 처리 잠금은 사본 안 둠
    dk.put('payment_sync:row:5', '{}', 60)                     # 다시 만들 수 있는 기준값은 대상 아님
    assert set(_rows(db)) == {'sms_intake:a'}


def test_wipe_restores_live_copies_only(db):
    rc = _FakeRedis()
    assert dk.check_and_restore(rc) is None                    # 처음 설치 — 세대 표시만 만듦
    rc.set('sms_intake:open', '{"text":"미처리"}', ex=3600)
    dk.put('sms_intake:open', '{"text":"미처리"}', 3600)
    rc.set('sms_intake:done', '{"text":"처리됨"}', ex=3600)
    dk.put('sms_intake:done', '{"text":"처리됨"}', 3600)
    rc.delete('sms_intake:done')                               # 확인 완료 → Redis·사본 같이 지움
    dk.delete('sms_intake:done')
    dk.put('invoice_card:C:9.9', '{"code":"G1"}', 1)
    with db._get_connection() as conn:                         # 이미 만료된 사본
        conn.execute("UPDATE redis_durable SET expires_at=? WHERE key='invoice_card:C:9.9'", (int(time.time()) - 5,))
        conn.commit()
    assert dk.check_and_restore(rc) is None                    # 평소 — 아무것도 안 함

    wiped = _FakeRedis()                                       # flush·AOF 없는 재생성
    res = dk.check_and_restore(wiped)
    assert res == {'wiped': True, 'restored': 1}
    assert json.loads(wiped.get('sms_intake:open'))['text'] == '미처리'
    assert wiped.get('sms_intake:done') is None                # 처리 끝난 건 되살리지 않음
    assert wiped.get('invoice_card:C:9.9') is None
    assert dk.check_and_restore(wiped) is None                 # 한 번만


def test_epoch_changed_alerts(db, caplog):
    rc = _FakeRedis()
    dk.check_and_restore(rc)
    other = _FakeRedis()
    other.set(dk.EPOCH_KEY, 'old-backup')
    with caplog.at_level('ERROR'):
        assert dk.check_and_restore(other) is None
    assert any('세대 표시가 바뀜' in r.getMessage() for r in caplog.records)


def test_add_once(db):
    assert dk.add_once('sms_intake:seen:x', 60) is True
    assert dk.add_once('sms_intake:seen:x', 60) is False
    with db._get_connection() as conn:
        conn.execute("UPDATE redis_durable SET expires_at=0 WHERE key='sms_intake:seen:x'")
        conn.commit()
    assert dk.add_once('sms_intake:seen:x', 60) is True        # 만료 뒤엔 다시


# ── 은행 문자 수신: Redis 장애 중에도 카드 ──
DEP = '[Web발신]\n하나,10/02, 15:00\n255******31304\n입금3,014,000원\n잔액12,345,678원\n대연이엔지주식회'


def test_sms_inbound_survives_redis_failure(db, monkeypatch):
    from dashboard.blueprints import sms_inbound as si

    class _Broken:
        def set(self, *a, **k):
            raise ConnectionError('Redis 연결 끊김')
    monkeypatch.setattr(si, 'get_redis_client', lambda: type('C', (), {'redis': _Broken()})())
    posted = []
    monkeypatch.setattr(si, '_post_intake_card', lambda iid, text, pv: posted.append(iid) or True)
    first = si.ingest_deposit(DEP, source='sms:sb')
    assert first['status'] == 'ok' and posted                         # 예전엔 예외 → HTTP 500
    assert '잔액' not in json.loads(_rows(db)[f"sms_intake:{first['id']}"])['text']   # 원문은 DB 사본으로
    assert si.ingest_deposit(DEP, source='sms:jw')['status'] == 'duplicate'           # 다른 폰 중복은 DB 표시로


def test_sms_inbound_normal_path_writes_copy(db, monkeypatch):
    from dashboard.blueprints import sms_inbound as si
    fr = _FakeRedis()
    monkeypatch.setattr(si, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    monkeypatch.setattr(si, '_post_intake_card', lambda *a: True)
    r = si.ingest_deposit(DEP, source='sms:sb')
    key = f"sms_intake:{r['id']}"
    assert fr.get(key) and _rows(db)[key] == fr.get(key)


# ── 스케줄러: 비워짐 감지 → 관리자 알림 ──

def test_scheduler_alerts_on_wipe(monkeypatch):
    from dashboard.services import sync_scheduler as ss
    sent = []
    monkeypatch.setattr(ss, '_notify_admin', lambda key, text: sent.append((key, text)))
    monkeypatch.setattr(dk, 'check_and_restore', lambda rc: {'wiped': True, 'restored': 3})
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': _FakeRedis()})())
    monkeypatch.setattr(ss, '_wipe_checked_at', 0.0)
    ss._check_redis_wiped()
    assert sent and sent[0][0] == 'redis_wiped' and '3건' in sent[0][1]
    ss._check_redis_wiped()                                           # 20초 안엔 다시 안 봄
    assert len(sent) == 1


if __name__ == '__main__':
    sys.exit(pytest.main([__file__, '-v']))


# ── 슬랙 봇: 지정 내용 갱신도 사본에, Redis 에서 지우는 곳은 사본도 지움 ──

def test_update_intake_mirrors_designation(db, monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    import dashboard.utils.redis_client as rcm
    fr = _FakeRedis()
    fr.set('sms_intake:i1', json.dumps({'text': '입금'}), ex=3600)
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    sb._update_intake('i1', designation={'code': 'G4100-AA', 'stage': '계약금'})
    assert json.loads(_rows(db)['sms_intake:i1'])['designation']['code'] == 'G4100-AA'


def test_every_redis_delete_of_protected_key_also_deletes_copy():
    """Redis 에서 지우고 사본을 안 지우면, 나중에 Redis 가 비었을 때 처리 끝난 입금·요청이 되살아난다."""
    import re
    from pathlib import Path
    src = Path('dashboard/blueprints/slack_bot.py').read_text(encoding='utf-8').splitlines()
    pat = re.compile(r'\.delete\((f?["\'](sms_intake|invoice_card|project_amount_req|project_cancel_money_req):'
                     r'[^)]*|key|k)\)')
    misses = []
    for i, line in enumerate(src):
        if not pat.search(line) or ':proc' in line or 'seen' in line:
            continue
        ctx = '\n'.join(src[max(0, i - 25):i])
        if line.strip().startswith(('rc.delete(k)', 'rc.delete(key)')) and not re.search(
                r"(project_amount_req|project_cancel_money_req|sms_intake)", ctx):
            continue                                            # 다른 키를 지우는 rc.delete(key)
        if '_dkv.delete(' not in '\n'.join(src[i:i + 4]):
            misses.append(f'{i + 1}: {line.strip()}')
    assert misses == []


def test_pending_request_ttl_is_long():
    import dashboard.blueprints.slack_bot as sb
    assert sb._PENDING_REQ_TTL >= 60 * 60 * 24 * 365
