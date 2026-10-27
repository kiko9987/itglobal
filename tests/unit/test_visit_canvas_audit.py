# -*- coding: utf-8 -*-
"""방문 일정 ↔ 캔버스1·캔버스2 누락 점검 회귀 테스트 (2026-10-27).

슬랙·시트 미접촉 — 판단은 순수 함수, 조회부는 mock.
"""
import sys
from datetime import date

sys.path.insert(0, '.')

import pytest
import dashboard.services.visit_canvas_audit as va

TODAY = date(2026, 10, 27)
C1 = [
    '(YM) 10월 27일 / 010-4327-1111 / 강남구 도산대로67길 7 파크빌딩 4층 / 실외기',   # ETC: 상호 없음
    '10월 28일 / 010-6725-1716 / 강남구 선릉로76길 40 신호빌딩 1층 / 교체설치',
    '(SH) 10월 29일 / - / 인천 미추홀구 인하로 79 2층 한의원(예정) / 현재공실',
]


def _lead(ln='L-1', st='방문 예약', vd='2026-10-28', ph='010-6725-1716',
          addr='강남구 선릉로76길 40 신호빌딩 1층', name='안희환', mgr=''):
    return {'리드 No': ln, '상태': st, '방문 예정일': vd, '고객 연락처': ph,
            '방문 주소': addr, '고객명': name, '영업 담당자': mgr}


# ── 캔버스1 포함 판단 ────────────────────────────────────────────────────────

def test_in_canvas_by_phone():
    assert va.lead_in_canvas_lines(_lead(), C1)


def test_in_canvas_by_phone_with_spaces_in_sheet():
    assert va.lead_in_canvas_lines(_lead(ph='010- 4327-1111', addr='x'), C1)


def test_in_canvas_by_address_when_no_phone():
    assert va.lead_in_canvas_lines(_lead(ph='-', addr='인천 미추홀구 인하로 79 2층'), C1)


def test_not_in_canvas():
    assert not va.lead_in_canvas_lines(_lead(ph='010-0000-1234', addr='부천 옥길로 131'), C1)


# ── 진행 중 카드 → 캔버스1에 없는 사유 ────────────────────────────────────────

@pytest.mark.parametrize('lead,in_c1,expect', [
    (None, False, '시트에 리드가 없음'),
    (_lead(st='방문 완료'), True, "시트 상태는 '방문 완료'"),                    # L-04289 형
    (_lead(vd='2026-10-24~20'), False, '끝 날짜가 시작보다 앞'),                 # L-04341 형
    (_lead(vd='미정'), False, '방문일을 읽을 수 없음'),
    (_lead(vd='2026-10-20'), False, '완료·취소 처리 안 됨'),
    (_lead(vd='2026-10-28'), False, '캔버스1에 없음'),
])
def test_classify_card_reasons(lead, in_c1, expect):
    assert expect in va.classify_card(lead, in_c1, TODAY)


@pytest.mark.parametrize('vd', ['2026-10-28', '2026-10-27', '2026-10-24~27', '2026-10-26~11-02'])
def test_classify_card_ok(vd):
    assert va.classify_card(_lead(vd=vd), True, TODAY) is None


# ── 캔버스1에 있는데 캔버스2에 안 옮긴 건 ─────────────────────────────────────

def test_canvas2_missing_target_day_only():
    leads = [_lead('L-A', vd='2026-10-28'),            # 대상일 → 누락
             _lead('L-B', vd='2026-10-28'),            # 매칭됨 → 제외
             _lead('L-C', vd='2026-10-29'),            # 다른 날 → 제외
             _lead('L-D', vd='2026-10-26~28'),         # 기간에 대상일 포함 → 누락
             _lead('L-E', vd='2026-10-24~20')]         # 날짜 오류 → 제외(카드 점검에서 잡음)
    got = [x['리드 No'] for x in va.canvas2_missing(leads, {'L-B'}, date(2026, 10, 28))]
    assert got == ['L-A', 'L-D']


# ── 진행 중 카드 수집 ────────────────────────────────────────────────────────

def test_active_visit_cards_only_buttons(monkeypatch):
    def _actions(ln):
        return [{'type': 'actions', 'elements': [
            {'action_id': 'visit_edit_info', 'value': ln},
            {'action_id': 'visit_complete', 'value': ln}]}]
    msgs = [
        {'bot_id': 'B', 'ts': '3', 'blocks': _actions('L-NEW')},
        {'bot_id': 'B', 'ts': '2', 'blocks': [{'type': 'section'}]},          # 완료 카드(버튼 없음)
        {'user': 'U1', 'ts': '1.5', 'blocks': _actions('L-HUMAN')},          # 사람 메시지 제외
        {'bot_id': 'B', 'ts': '1', 'blocks': _actions('L-NEW')},             # 옛 카드 중복 제외
    ]
    monkeypatch.setattr(va, '_slack_get', lambda m, p: {'ok': True, 'messages': msgs})
    monkeypatch.setenv('SLACK_VISIT_CHANNEL', 'C1')
    assert va.active_visit_cards() == [{'lead_no': 'L-NEW', 'ts': '3'}]


# ── 메시지 ──────────────────────────────────────────────────────────────────

def test_format_commit_audit_empty():
    assert va.format_commit_audit({'cards': [], 'c2_missing': [], 'c2_unmatched': []}) == ''


def test_format_commit_audit_sections():
    txt = va.format_commit_audit({
        'target': '2026-10-28',
        'cards': [{'lead_no': 'L-04341', 'name': '정영수', 'reason': '방문일 끝 날짜가 시작보다 앞'}],
        'c2_missing': [{'lead_no': 'L-04368', 'name': '안희환', 'addr': '강남구 선릉로76길 40',
                        'visit': '2026-10-28', 'mgr': ''}],
        'c2_unmatched': [{'raw': '(MW) 10월 27일 / 010- 4327-1111 / 강남구 도산대로67길 7', 'phone': ''}],
    })
    assert '캔버스1에 없는 방문 카드 1건' in txt and 'L-04341' in txt
    assert '10월 28일 방문 1건' in txt and 'L-04368' in txt and '담당 미정' in txt
    assert '리드를 못 찾은 줄 1건' in txt and '010- 4327-1111' in txt


# ── 오후 4시 작업 ────────────────────────────────────────────────────────────

def test_daily_skips_non_business_day(monkeypatch):
    import dashboard.services.absent_remind as ar
    monkeypatch.setattr(ar, '_is_business_day', lambda d: False)
    monkeypatch.setattr(va, 'check_cards_vs_canvas1', lambda today=None: pytest.fail('점검하면 안 됨'))
    assert va.run_daily_canvas1_audit() == {'skipped': '영업일 아님'}


def test_daily_sends_only_when_items(monkeypatch):
    import dashboard.services.absent_remind as ar
    monkeypatch.setattr(ar, '_is_business_day', lambda d: True)
    sent = []
    monkeypatch.setattr(va, 'send_admin_dm', lambda t: sent.append(t) or True)
    monkeypatch.setattr(va, 'check_cards_vs_canvas1', lambda today=None: [])
    assert va.run_daily_canvas1_audit() == {'items': 0} and sent == []
    monkeypatch.setattr(va, 'check_cards_vs_canvas1', lambda today=None: [
        {'lead_no': 'L-04341', 'name': '정영수', 'reason': '방문일 끝 날짜가 시작보다 앞'}])
    assert va.run_daily_canvas1_audit()['items'] == 1
    assert len(sent) == 1 and 'L-04341' in sent[0] and '캔버스1에 없는 건 1건' in sent[0]
