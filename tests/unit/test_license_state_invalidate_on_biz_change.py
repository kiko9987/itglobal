# -*- coding: utf-8 -*-
"""사업자명·폴더가 바뀌면 등록증 판정 캐시를 비움 (2026-10-22 G4058).

시트에 사업자명을 '-'에서 채웠는데 이전 '등록증 없음' 판정(1시간 캐시)이 남아 계산서 요청이
막힐 뻔했다. PM 편집·시트 직접수정(리컨사일러) 모두 notify_project_field_changes 를 지난다.
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.services.project_slack_notifier as psn


@pytest.fixture
def cleared(monkeypatch):
    import dashboard.utils.redis_client as rcm
    import dashboard.services.business_license_handler as blh
    monkeypatch.setattr(rcm, 'get_redis_client',
                        lambda: type('C', (), {'redis': type('R', (), {'get': lambda self, k: None})()})())
    import dashboard.services.partner_names as pn
    out = []
    monkeypatch.setattr(blh, 'invalidate_license_state', lambda code: out.append(code))
    # 사업자명 변경은 같은 곳 거래처명 묶기도 부름 — 여기선 실제 시트를 읽지 않게 막음
    monkeypatch.setattr(pn, 'apply_place_on_biz_change', lambda code, biz, editor='', old_biz='': 0)
    return out


@pytest.mark.parametrize('field', ['사업자명', '견적서 및 계약서 폴더 경로'])
def test_biz_or_folder_change_clears_license_state(cleared, field):
    psn.notify_project_field_changes('G4058-YG', [{'field_name': field, 'old_value': '-', 'new_value': '새값'}])
    assert cleared == ['G4058-YG']


def test_other_field_keeps_license_state(cleared):
    psn.notify_project_field_changes('G4058-YG', [{'field_name': '시공자', 'old_value': '이', 'new_value': '김'}])
    psn.notify_project_field_changes('G4058-YG', [{'field_name': '사업자명', 'old_value': 'A', 'new_value': 'A'}])
    assert cleared == []
