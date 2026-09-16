# -*- coding: utf-8 -*-
"""공사 취소 — 상태 즉시 + 매출0은 샛별 ✅ 후 별도 + 재개 복원 회귀 테스트 (2026-09-16).

설계: perform_cancel=상태만(총액 유지)+스냅샷, apply_cancel_zero_total=총액0(✅ 후),
perform_uncancel=총액 복원. 시트/캐시/Redis/감사는 monkeypatch.
"""
import sys
sys.path.insert(0, '.')

import json

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
                        lambda sid, updates, tag=None: captured.setdefault('updates', []).extend(updates))
    monkeypatch.setattr(psa, '_audit_log', lambda **k: None)
    monkeypatch.setattr(psa, '_queue_bg_color', lambda *a, **k: None)
    import dashboard.utils.redis_client as rcmod
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: _RC(redis))
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'update_project_in_cache', lambda code, payload: True)
    monkeypatch.setattr(ps, 'invalidate_project_cache', lambda code: None)


def _ranges(captured):
    return {u['range']: u['values'][0][0] for u in captured.get('updates', [])}


def test_cancel_status_only_snapshots_total(monkeypatch):
    """취소는 상태만 즉시 — 총액은 안 건드리고 스냅샷에 원본 보존, 특이사항에 사유."""
    redis = _Redis()
    captured = {}
    proj = {'프로젝트 코드': 'G9999-XX', '수금 관련 특이사항': '', '공사 확정': '2026-09-01',
            '수금 확인': True, '총액 1': 3400000}
    _patch_common(monkeypatch, proj, redis, captured)

    res = psa.perform_cancel('G9999-XX', 'SD', reason='고객 변심')
    assert res['ok'] is True
    r = _ranges(captured)
    assert '공사 현황!AB100' in r and '공사 취소 (사유: 고객 변심)' in r['공사 현황!AB100']
    assert r.get('공사 현황!AD100') == 'FALSE'
    # 총액(T)은 이 단계에서 안 건드림
    assert '공사 현황!T100' not in r
    # 스냅샷에 원본 총액 보존
    snap = json.loads(redis.store['project_cancel_snapshot:G9999-XX'])
    assert snap['total_1'] == 3400000


def test_apply_zero_total_after_check(monkeypatch):
    """샛별 ✅ 후 apply_cancel_zero_total → 총액 1=0 (취소 상태일 때만)."""
    redis = _Redis()
    captured = {}
    proj = {'프로젝트 코드': 'G9999-XX', '수금 관련 특이사항': '공사 취소 (사유: 고객 변심)',
            '공사 확정': '', '수금 확인': False, '총액 1': 3400000}
    _patch_common(monkeypatch, proj, redis, captured)

    res = psa.apply_cancel_zero_total('G9999-XX', 'SB')
    assert res['ok'] is True
    assert _ranges(captured).get('공사 현황!T100') == 0


def test_apply_zero_total_skips_if_not_cancelled(monkeypatch):
    """취소 상태가 아니면(재개됨) 매출 0 처리 skip."""
    redis = _Redis()
    captured = {}
    proj = {'프로젝트 코드': 'G9999-XX', '수금 관련 특이사항': '', '공사 확정': '2026-09-01',
            '수금 확인': True, '총액 1': 3400000}
    _patch_common(monkeypatch, proj, redis, captured)

    res = psa.apply_cancel_zero_total('G9999-XX', 'SB')
    assert res['ok'] is False and res['reason'] == 'not_cancelled'
    assert 'updates' not in captured  # 아무 write 안 함


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
    r = _ranges(captured)
    assert r.get('공사 현황!T100') == 3400000  # 총액 원복
    assert r.get('공사 현황!AB100') == ''       # 특이사항 해제
    assert 'project_cancel_snapshot:G9999-XX' not in redis.store


def test_cancel_no_reason_plain_spec(monkeypatch):
    redis = _Redis()
    captured = {}
    proj = {'프로젝트 코드': 'G1-XX', '수금 관련 특이사항': '', '공사 확정': '', '수금 확인': False, '총액 1': 100}
    _patch_common(monkeypatch, proj, redis, captured)
    res = psa.perform_cancel('G1-XX', 'SD')  # 사유 없음
    assert res['ok'] is True
    assert _ranges(captured).get('공사 현황!AB100') == '공사 취소'
