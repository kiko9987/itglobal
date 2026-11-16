# -*- coding: utf-8 -*-
"""중복 등록 선점(claim) — 2026-11-16 R4291/R4292 사고 재발 방지.

사고: 13:03:37 1차 제출 → 13:04:08 2차 제출(31초 뒤) → 1차가 13:04:14 에야 키를 심어
둘 다 통과. 멱등 키·dedup 키가 똑같았는데도 두 공사가 생겼다. 조회와 저장 사이가
벌어진 구조였기 때문. 이제는 들어오는 순간 SET NX 로 차지한다.
"""
import json

import pytest

from dashboard.blueprints import projects as P


class FakeRedis:
    """set(nx=, ex=)·get·delete 만 쓰는 최소 구현."""

    def __init__(self):
        self.store = {}

    def set(self, key, value, nx=False, ex=None):
        if nx and key in self.store:
            return False
        self.store[key] = value
        return True

    def get(self, key):
        return self.store.get(key)

    def delete(self, key):
        return self.store.pop(key, None) is not None


DATA = {'사업자': '글로벌그룹', '담당자': '빈승정',
        '현장 주소': '구리 경춘로 167 구리수아주퍼스트오피스텔 2층',
        '공사 시작': '2026-11-17', '총액 1': '9400000'}


def test_claim_is_exclusive():
    """두 번째 요청은 같은 키를 차지하지 못한다 — 사고의 핵심."""
    rc = FakeRedis()
    assert P._claim(rc, 'k', 600) is True
    assert P._claim(rc, 'k', 600) is False


def test_release_only_removes_own_pending():
    """결과가 들어간 키는 풀지 않는다 (남의 응답을 지우면 더 큰 사고)."""
    rc = FakeRedis()
    P._claim(rc, 'k', 600)
    P._release_claim(rc, 'k')
    assert rc.get('k') is None

    rc.set('k', 'R4292-SJ')
    P._release_claim(rc, 'k')
    assert rc.get('k') == 'R4292-SJ'


def test_await_returns_result_pending_or_none():
    """기다리기의 세 갈래: 결과 있음 / 아직 처리 중 / 선점자가 포기."""
    rc = FakeRedis()
    rc.set('done', 'R4292-SJ')
    assert P._await_claim(rc, 'done', timeout=0.05) == 'R4292-SJ'

    P._claim(rc, 'busy', 600)
    assert P._await_claim(rc, 'busy', timeout=0.05, interval=0.01) == P._CLAIM_PENDING

    assert P._await_claim(rc, 'gone', timeout=0.05) is None


def test_dedup_fields_needs_core_values():
    assert P._dedup_fields(DATA) == (
        '글로벌그룹', '빈승정', '구리 경춘로 167 구리수아주퍼스트오피스텔 2층',
        '2026-11-17', '9400000')
    # 주소가 비면 판정 못 함 → None (아무 요청도 막지 않는다)
    assert P._dedup_fields({**DATA, '현장 주소': ''}) is None


def test_same_data_same_hash():
    """R4291/R4292 는 다섯 값이 모두 같았다 → 같은 해시 → 두 번째는 막힌다."""
    a = P._dedup_hash(*P._dedup_fields(DATA))
    b = P._dedup_hash(*P._dedup_fields({**DATA, '거래처명': '비해피팜'}))
    assert a == b
    rc = FakeRedis()
    key = f'project_create_dedup:{a}'
    assert P._claim(rc, key, 300) is True
    assert P._claim(rc, f'project_create_dedup:{b}', 300) is False


def test_find_recent_duplicate_ignores_pending_marker(monkeypatch):
    """선점 표시('__pending__')를 공사 코드로 착각하면 안 된다.

    Redis 조회 레이어를 뺀 이유. 남겨 뒀다면 이 값이 코드처럼 응답에 실렸다.
    """
    import pandas as pd
    monkeypatch.setattr('dashboard.services.project_service.load_data',
                        lambda: pd.DataFrame([]))
    assert P._find_recent_duplicate(DATA) is None


def test_find_recent_duplicate_matches_today_row(monkeypatch):
    """세 번째 그물은 그대로 — 오늘 같은 데이터가 이미 있으면 그 코드를 돌려준다."""
    import pandas as pd
    from datetime import datetime
    today = datetime.now().strftime('%Y-%m-%d')
    df = pd.DataFrame([{
        '프로젝트 코드': 'R4292-SJ', '공사 확정': today,
        '사업자': DATA['사업자'], '담당자': DATA['담당자'],
        '현장 주소': DATA['현장 주소'], '공사 시작': DATA['공사 시작'],
        '총액 1': '9400000',
    }])
    monkeypatch.setattr('dashboard.services.project_service.load_data', lambda: df)
    assert P._find_recent_duplicate(DATA) == 'R4292-SJ'


def test_dedup_response_shape(monkeypatch):
    monkeypatch.setattr('dashboard.services.project_service.get_project_records',
                        lambda: [{'프로젝트 코드': 'R4292-SJ', '총액 1': 9400000}])
    r = P._dedup_response('R4292-SJ')
    assert r['success'] is True and r['deduped'] is True
    assert r['project_code'] == 'R4292-SJ'
    assert r['project_data']['총액 1'] == 9400000


def test_second_request_gets_first_response():
    """사고 재현: 1차가 결과를 넣은 뒤 2차가 와도 새 코드를 만들지 않는다."""
    rc = FakeRedis()
    idem = 'project_create_idem:fp-923f7'
    assert P._claim(rc, idem, 600) is True          # 13:03:37 1차
    rc.set(idem, json.dumps({'success': True, 'project_code': 'R4291-SJ'}))
    assert P._claim(rc, idem, 600) is False         # 13:04:08 2차 — 막힘
    got = P._await_claim(rc, idem, timeout=0.05)
    assert json.loads(got)['project_code'] == 'R4291-SJ'
