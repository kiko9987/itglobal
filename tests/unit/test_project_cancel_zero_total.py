# -*- coding: utf-8 -*-
"""공사 취소 시 매출(총액 1) 0 반환 + 스냅샷/재개 복원 회귀 테스트 (2026-09-16).

경영지원 요청: 취소 시 매출 0, 재개 시 원복. 시트/캐시/Redis/감사는 monkeypatch.
"""
import sys
sys.path.insert(0, '.')

import json
import pytest

from dashboard.services import project_slack_actions as psa


class _Mgr:
    def get_field_to_letter(self):
        return {'수금 관련 특이사항': 'AB', '수금 확인': 'AD', '공사 확정': 'S', '총액 1': 'T'}


class _Redis:
    def __init__(self):
        self.store = {}

    def set(self, k, v, ex=None):
        self.store[k] = v

    def get(self, k):
        return self.store.get(k)

    def delete(self, k):
        self.store.pop(k, None)


class _RC:
    def __init__(self, r):
        self.redis = r


def _patch_common(monkeypatch, project, redis, captured):
    monkeypatch.setattr(psa, '_load_project', lambda code: dict(project))
    monkeypatch.setattr(psa, '_get_sheet_context', lambda: (_Mgr(), 'sid', '공사 현황'))
    monkeypatch.setattr(psa, '_find_row_number', lambda m, s, n, c: 100)
    monkeypatch.setattr(psa, '_queue_batch_write',
                        lambda sid, updates, tag=None: captured.__setitem__('updates', updates))
    monkeypatch.setattr(psa, '_audit_log', lambda **k: None)
    monkeypatch.setattr(psa, '_queue_bg_color', lambda *a, **k: None)
    import dashboard.utils.redis_client as rcmod
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: _RC(redis))
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'update_project_in_cache', lambda code, payload: True)
    monkeypatch.setattr(ps, 'invalidate_project_cache', lambda code: None)


def test_cancel_zeroes_total_and_snapshots(monkeypatch):
    redis = _Redis()
    captured = {}
    proj = {'프로젝트 코드': 'G9999-XX', '수금 관련 특이사항': '', '공사 확정': '2026-09-01',
            '수금 확인': True, '총액 1': 3400000}
    _patch_common(monkeypatch, proj, redis, captured)

    res = psa.perform_cancel('G9999-XX', 'SD', reason='고객 변심')
    assert res['ok'] is True
    ranges = {u['range']: u['values'][0][0] for u in captured['updates']}
    # 총액 1(T100) → 0
    assert ranges.get('공사 현황!T100') == 0
    # 특이사항에 사유 병기
    assert '공사 취소 (사유: 고객 변심)' in ranges.get('공사 현황!AB100', '')
    # 스냅샷에 원본 총액 보존
    snap = json.loads(redis.store['project_cancel_snapshot:G9999-XX'])
    assert snap['total_1'] == 3400000
    assert snap['confirmed_date'] == '2026-09-01'
    assert snap['payment_confirmed'] is True


def test_uncancel_restores_total(monkeypatch):
    redis = _Redis()
    redis.store['project_cancel_snapshot:G9999-XX'] = json.dumps({
        'confirmed_date': '2026-09-01', 'payment_confirmed': True, 'total_1': 3400000,
    })
    captured = {}
    proj = {'프로젝트 코드': 'G9999-XX', '수금 관련 특이사항': '공사 취소 (사유: 고객 변심)',
            '공사 확정': '', '수금 확인': False, '총액 1': 0}
    _patch_common(monkeypatch, proj, redis, captured)

    res = psa.perform_uncancel('G9999-XX', 'SD')
    assert res['ok'] is True
    ranges = {u['range']: u['values'][0][0] for u in captured['updates']}
    # 총액 1 원복
    assert ranges.get('공사 현황!T100') == 3400000
    # 특이사항 해제
    assert ranges.get('공사 현황!AB100') == ''
    # 스냅샷 정리됨
    assert 'project_cancel_snapshot:G9999-XX' not in redis.store


def test_cancel_no_reason_plain_spec(monkeypatch):
    redis = _Redis()
    captured = {}
    proj = {'프로젝트 코드': 'G1-XX', '수금 관련 특이사항': '', '공사 확정': '', '수금 확인': False, '총액 1': 100}
    _patch_common(monkeypatch, proj, redis, captured)
    res = psa.perform_cancel('G1-XX', 'SD')  # 사유 없음
    assert res['ok'] is True
    ranges = {u['range']: u['values'][0][0] for u in captured['updates']}
    assert ranges.get('공사 현황!AB100') == '공사 취소'  # 사유 없으면 순수 '공사 취소'
