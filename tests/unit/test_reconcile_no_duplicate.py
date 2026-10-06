# -*- coding: utf-8 -*-
"""리컨사일러 중복 댓글 방지 — 알림 경로가 스냅샷을 직접 맞춤 (2026-10-06 R4029-JK).

9/30 17:02 PM 편집 댓글 → 서버 재시작(17:09)으로 리컨사일러 첫 점검이 17:19 → PM 15분 마커가
이미 만료(17:17) → 같은 변경에 '(시트 직접수정)' 중복 댓글. 마커(시간) 대신 알림 경로가
스냅샷을 새 값으로 맞추면 점검이 늦어도 차이가 없어 재발송 안 함.
"""
import sys

sys.path.insert(0, '.')

import dashboard.services.project_reconcile as pr
import dashboard.services.project_slack_notifier as psn

SPEC = '9/28 : 나라장터로 검수요청서 전송 후 서초 담당자 환료 진행중'


class _FakeRedis:
    def __init__(self):
        self.kv, self.h = {}, {}

    def get(self, k): return self.kv.get(k)
    def set(self, k, v, nx=False, ex=None):
        if nx and k in self.kv:
            return None
        self.kv[k] = v
        return True
    def setex(self, k, t, v): self.kv[k] = v
    def delete(self, k): self.kv.pop(k, None); self.h.pop(k, None)
    def exists(self, k): return k in self.h or k in self.kv
    def hset(self, k, mapping): self.h.setdefault(k, {}).update(mapping)
    def hgetall(self, k): return dict(self.h.get(k, {}))
    def expire(self, k, t): pass
    def scan_iter(self, match='', count=0):
        p = match.rstrip('*')
        return [k for k in list(self.kv) if k.startswith(p)]


def _setup(monkeypatch, fr):
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': fr})())
    monkeypatch.setenv('SLACK_PROJECT_BOT_TOKEN', 'x')
    import dashboard.blueprints.slack_helpers as sh
    monkeypatch.setattr(sh, 'safe_slack_post_url', lambda *a, **k: {'ok': True})


def test_sync_snapshot_fields_only_when_snapshot_exists(monkeypatch):
    fr = _FakeRedis()
    _setup(monkeypatch, fr)
    pr.sync_snapshot_fields('G1', {'시공자': '최태식'})
    assert 'card_field_snap:G1' not in fr.h                 # 베이스라인 전엔 손대지 않음
    fr.hset('card_field_snap:G1', mapping={'시공자': '-'})
    pr.sync_snapshot_fields('G1', {'시공자': '최태식', '공사 확정': None})
    assert fr.h['card_field_snap:G1'] == {'시공자': '최태식', '공사 확정': ''}


def test_pm_comment_then_late_reconcile_no_duplicate(monkeypatch):
    fr = _FakeRedis()
    _setup(monkeypatch, fr)
    code = 'R4029-JK'
    fr.set(f'project_card_msg:{code}', 'C04UL30G09H|1.1')
    # 리컨사일러 베이스라인(편집 전)
    watch = psn._NOTIFY_FIELDS - pr._RECONCILE_EXCLUDE
    before = {f: '' for f in watch}
    before.update({'프로젝트 코드': code, '수금 관련 특이사항': '-'})
    fr.hset(f'card_field_snap:{code}', mapping=before)
    # ① PM 편집 알림(댓글) — 마커는 없다고 가정(=재시작으로 만료된 상황)
    assert psn.notify_project_field_changes(code, [
        {'field_name': '수금 관련 특이사항', 'old_value': '-', 'new_value': SPEC}], latest_data=None)
    assert fr.h[f'card_field_snap:{code}']['수금 관련 특이사항'] == SPEC
    # ② 늦은 리컨사일러 점검 — 시트에는 새 값. 중복 댓글이 나가면 안 됨
    rec = dict(before, **{'프로젝트 코드': code, '수금 관련 특이사항': SPEC})
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: [rec])
    sent = []
    monkeypatch.setattr(psn, 'notify_project_field_changes', lambda *a, **k: sent.append(a) or True)
    monkeypatch.setattr(psn, 'notify_amount_edit_to_settlement', lambda *a, **k: None)
    r = pr.reconcile_project_cards()
    assert sent == [] and r['reflected'] == 0


def test_real_direct_edit_still_reported(monkeypatch):
    """알림 없이 시트만 바뀐 진짜 직접수정은 그대로 감지·댓글."""
    fr = _FakeRedis()
    _setup(monkeypatch, fr)
    code = 'R4029-JK'
    fr.set(f'project_card_msg:{code}', 'C04UL30G09H|1.1')
    watch = psn._NOTIFY_FIELDS - pr._RECONCILE_EXCLUDE
    before = {f: '' for f in watch}
    before['프로젝트 코드'] = code
    fr.hset(f'card_field_snap:{code}', mapping=before)
    rec = dict(before, **{'시공자': '최태식'})
    import dashboard.services.project_service as ps
    monkeypatch.setattr(ps, 'get_project_records', lambda force_refresh=False: [rec])
    sent = []
    monkeypatch.setattr(psn, 'notify_project_field_changes', lambda *a, **k: sent.append(k.get('editor')) or True)
    monkeypatch.setattr(psn, 'notify_amount_edit_to_settlement', lambda *a, **k: None)
    pr.reconcile_project_cards()
    assert sent == ['시트 직접수정']
