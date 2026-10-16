# -*- coding: utf-8 -*-
"""취소된 공사 편집 시 슬랙 원본 카드(회색 '공사 취소')를 일반 카드로 다시 그리지 않음 (2026-10-16).

관리자 취소 건 편집 허용 + 리컨사일러(시트 직접수정) 경로 공통 — 수정 내역은 스레드 댓글로만.
"""
import sys

sys.path.insert(0, '.')

import dashboard.services.project_slack_notifier as psn


class _R:
    def __init__(self):
        self.store = {'project_card_msg:R1-X': 'C1|111.222'}

    def get(self, k):
        return self.store.get(k)

    def set(self, *a, **k):
        return True

    def setex(self, *a, **k):
        return True

    def delete(self, *a, **k):
        return True


def _run(monkeypatch, notes):
    monkeypatch.setenv('SLACK_PROJECT_BOT_TOKEN', 'x')
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': _R()})())
    posted, updated = [], []
    import dashboard.blueprints.slack_helpers as sh
    monkeypatch.setattr(sh, 'safe_slack_post_url', lambda url, payload, token: posted.append(payload) or {'ok': True})
    monkeypatch.setattr(psn.urllib.request, 'urlopen', lambda *a, **k: updated.append(a) or (_ for _ in ()).throw(RuntimeError('no')))
    import dashboard.services.project_reconcile as pr
    monkeypatch.setattr(pr, 'sync_snapshot_fields', lambda *a, **k: None)
    monkeypatch.setattr(psn, '_thread_permalink', lambda *a, **k: '')
    import dashboard.services.business_license_handler as blh
    monkeypatch.setattr(blh, 'verify_license_exists', lambda code: False)
    latest = {'프로젝트 코드': 'R1-X', '사업자명': 'A', '시공자': '김', '수금 관련 특이사항': notes}
    psn.notify_project_field_changes('R1-X', [{'field_name': '시공자', 'old_value': '이', 'new_value': '김'}],
                                     latest_data=latest, editor='수정 : KIKO')
    return posted, updated


def test_cancelled_project_comment_only(monkeypatch):
    posted, updated = _run(monkeypatch, '공사 취소 (사유: 고객 변심)')
    assert len(posted) == 1 and '시공자' in posted[0]['text']
    assert updated == []                      # 원본 카드 chat.update 안 함


def test_active_project_rerenders_card(monkeypatch):
    posted, updated = _run(monkeypatch, '')
    assert len(posted) == 1
    assert len(updated) == 1                  # 정상 공사는 기존대로 원본 카드 갱신 시도
