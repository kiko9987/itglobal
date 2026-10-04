# -*- coding: utf-8 -*-
"""회사 계좌 출금 문자 서버 보관 (슬랙 미노출, 2026-10-04 사용자 결정).

예전엔 출금 문자를 받자마자 버려(not_payment) 원문이 남지 않았다 — 대연(G4125-YM) 법인 간
이체 때 하나 출금 문자가 서버에 왔지만 확인 불가. 법인 간 이체 짝짓기 증거·진단용으로 보관.
"""
import sys
import time

sys.path.insert(0, '.')

import dashboard.blueprints.sms_inbound as si
from dashboard.services.sms_intake import looks_like_withdrawal

HANA_OUT = '[Web발신]\n하나,10/02, 15:00\n255******31304\n출금3,014,000원\n잔액12,345,678원\n대연이엔지주식회'


class _FakeRedis:
    def __init__(self):
        self.kv, self.z = {}, {}

    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.kv:
            return None
        self.kv[k] = v
        return True

    def get(self, k):
        return self.kv.get(k)

    def zadd(self, key, mapping):
        self.z.setdefault(key, {}).update(mapping)

    def zremrangebyscore(self, key, lo, hi):
        z = self.z.get(key, {})
        for m in [m for m, s in z.items() if lo <= s <= hi]:
            z.pop(m)

    def zrangebyscore(self, key, lo, hi):
        z = self.z.get(key, {})
        return [m for m, s in sorted(z.items(), key=lambda x: x[1]) if s >= lo]


def _patch(monkeypatch):
    fr = _FakeRedis()
    monkeypatch.setattr(si, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    monkeypatch.setattr(si, '_post_intake_card', lambda *a, **k: (_ for _ in ()).throw(
        AssertionError('출금은 슬랙 카드 금지')))
    return fr


def test_looks_like_withdrawal():
    assert looks_like_withdrawal(HANA_OUT)
    assert looks_like_withdrawal('출금 200,000원')
    assert not looks_like_withdrawal('입금 3,014,000원')


def test_company_withdrawal_stored_without_balance_no_slack(monkeypatch):
    fr = _patch(monkeypatch)
    r = si.ingest_deposit(HANA_OUT, source='sms:sb')
    assert r['status'] == 'ignored' and r['reason'] == 'outflow_recorded'
    import json
    rec = json.loads(fr.kv[f"sms_outflow:{r['id']}"])
    assert '잔액' not in rec['text'] and '12,345,678' not in rec['text']   # 잔고 비노출
    assert rec['amount'] == 3014000 and rec['partner'] == '대연이엔지주식회'
    assert rec['bank'] == '하나' and rec['acct_code'] == 'R'
    out = si.recent_outflows(since_ts=int(time.time()) - 60)
    assert [o['id'] for o in out] == [r['id']]


def test_duplicate_forward_from_other_phone_skipped(monkeypatch):
    _patch(monkeypatch)
    a = si.ingest_deposit(HANA_OUT, source='sms:sb')
    b = si.ingest_deposit(HANA_OUT, source='sms:yg')      # 다른 폰이 같은 문자
    assert a['reason'] == 'outflow_recorded' and b['status'] == 'duplicate'


def test_personal_account_withdrawal_not_stored(monkeypatch):
    fr = _patch(monkeypatch)
    r = si.ingest_deposit('[Web발신]\n국민 10/02 15:00\n123-45-6789\n출금 50,000원\n홍길동', source='sms:sb')
    assert r == {'status': 'ignored', 'reason': 'not_payment'}
    assert not any(k.startswith('sms_outflow:') for k in fr.kv)
