# -*- coding: utf-8 -*-
"""정보 수정에서 '본인 방문 필수' 켜고 끄기 (2026-10-14 L-04243: 등록 때 켠 걸 끌 방법이 없었음)."""
import sys
sys.path.insert(0, '.')

import dashboard.blueprints.slack_bot as sb

PARTNER = {'플랫폼': '거래처', '온라인 상담자': '박용구', '영업 담당자': '박용구', '본인 방문 여부': '본인 방문 필수'}


def _sel(v):
    return {'self_visit': {'value': {'selected_option': {'value': v}}}}


def test_dropdown_state():
    assert sb._self_visit_checked({}) is None  # 칸 없는 옛 모달 → 안 바꿈
    assert sb._self_visit_checked(_sel('아무나 방문 가능')) is False
    assert sb._self_visit_checked(_sel('본인 방문 필수')) is True
    # 잠깐 배포됐던 체크 칸 모달 제출도 처리
    assert sb._self_visit_checked({'self_visit': {'value': {'selected_options': []}}}) is False
    assert sb._self_visit_checked({'self_visit': {'value': {'selected_options': [{'value': 'yes'}]}}}) is True


def test_modal_has_same_dropdown_as_registration(monkeypatch):
    """등록 워크플로 모달처럼 방문 예정일 바로 아래 '본인 방문 여부' 선택 상자, 현재 값으로 선택."""
    views = []

    class _C:
        def views_open(self, **kw):
            return {'view': {'id': 'V1'}}

        def views_update(self, **kw):
            views.append(kw['view'])

    monkeypatch.setattr(sb, '_find_lead_by_no', lambda no: dict(PARTNER, **{'방문 예정일': '2026-10-15'}))
    sb._open_visit_edit_modal(_C(), 'L-04243', 'C1', '1.1', 'T1')
    ids = [b.get('block_id') for b in views[0]['blocks']]
    assert ids.index('self_visit') == ids.index('visit_date_end') + 1
    el = views[0]['blocks'][ids.index('self_visit')]['element']
    assert views[0]['blocks'][ids.index('self_visit')]['label']['text'] == '본인 방문 여부'
    assert [o['value'] for o in el['options']] == ['아무나 방문 가능', '본인 방문 필수']
    assert el['initial_option']['value'] == '본인 방문 필수'


def test_turn_off_partner_clears_auto_assigned_visitor():
    # 거래처 워크플로가 등록자(M)를 영업 담당자(N)에 자동 복사해 둔 경우 → 비워서 JW 배정 대상
    assert sb._self_visit_updates(PARTNER, False) == {'본인 방문 여부': '아무나 방문 가능', '영업 담당자': '-'}


def test_turn_off_keeps_other_assigned_visitor():
    lead = dict(PARTNER, **{'영업 담당자': '권태훈'})
    assert sb._self_visit_updates(lead, False) == {'본인 방문 여부': '아무나 방문 가능'}


def test_turn_off_online_lead_keeps_visitor():
    lead = dict(PARTNER, **{'플랫폼': '홈페이지'})
    assert sb._self_visit_updates(lead, False) == {'본인 방문 여부': '아무나 방문 가능'}


def test_turn_on_fills_empty_visitor_with_registrant():
    lead = dict(PARTNER, **{'영업 담당자': '-', '본인 방문 여부': '아무나 방문 가능'})
    assert sb._self_visit_updates(lead, True) == {'본인 방문 여부': '본인 방문 필수', '영업 담당자': '박용구'}


def test_edit_writes_self_visit_and_keeps_badge(monkeypatch):
    writes, cards = [], []
    lead = dict(PARTNER, **{'방문 예정일': '2026-10-15', '상담 내용': '천장형 냉난방기 설계 견적'})
    monkeypatch.setattr(sb, '_find_lead_by_no', lambda no: dict(lead))
    monkeypatch.setattr(sb, '_update_lead_dispatch', lambda no, upd: writes.append(upd))
    monkeypatch.setattr(sb, '_trigger_visit_list_webhook', lambda *a, **k: None)
    monkeypatch.setattr(sb, '_normalize_visit_address_if_verified', lambda a: (a, None))
    real_build = sb._build_visit_notice_blocks
    monkeypatch.setattr(sb, '_build_visit_notice_blocks', lambda **kw: cards.append(kw) or real_build(**kw))

    class _C:
        def chat_update(self, **kw):
            pass

    sb._process_visit_edit_same_platform(_C(), {}, 'L-04243', 'C1', '1.1', '2026-10-15', '', '강서한빛의원',
                                         '010-9020-4835', '강서구 금낭화로 135', '천장형 냉난방기 설계 견적',
                                         self_visit=False)
    assert writes[0]['본인 방문 여부'] == '아무나 방문 가능' and writes[0]['영업 담당자'] == '-'
    assert cards[0]['self_visit_by'] == ''
    writes.clear(); cards.clear()
    sb._process_visit_edit_same_platform(_C(), {}, 'L-04243', 'C1', '1.1', '2026-10-15', '', '강서한빛의원',
                                         '010-9020-4835', '강서구 금낭화로 135', '천장형 냉난방기 설계 견적',
                                         self_visit=None)
    assert '본인 방문 여부' not in writes[0]          # 칸 없으면 안 바꿈
    assert cards[0]['self_visit_by'] == '본인'        # 기존 본인 방문 배지 유지
