# -*- coding: utf-8 -*-
"""2026-10-28 전체 점검 2차(표기 통일) 회귀 테스트."""
import sys

sys.path.insert(0, '.')

import pandas as pd


def test_reason_codes_become_korean():
    from dashboard.blueprints.slack_bot import _reason_ko
    assert _reason_ko('row_not_found') == '시트에서 이 공사를 찾지 못했습니다'
    assert _reason_ko('already_active').startswith('취소된 공사가 아닙니다')
    assert _reason_ko('Failed to generate project code: Company prefix not found.') == '처리 중 오류가 발생했습니다'
    assert _reason_ko('시트 저장 실패') == '시트 저장 실패'


def test_site_head_has_partner_and_split_biz():
    from dashboard.blueprints.slack_bot import _site_head_lines
    assert _site_head_lines({'거래처명': '메이드바이', '사업자명': '(주)에이'}) == [
        '🤝 거래처명 : 메이드바이', '🏢 사업자명 : (주)에이']
    # 분류값 거래처명은 줄 생략, 사업자명 없으면 '-'
    assert _site_head_lines({'거래처명': '온라인', '사업자명': ''}) == ['🏢 사업자명 : -']


def test_inflow_display_online_and_as_same_helper():
    from dashboard.services.lead_helpers import format_inflow_display
    assert format_inflow_display('온라인') == '온라인'
    assert format_inflow_display('전화') == '온라인 (전화)'
    assert format_inflow_display('거래처') == '거래처'


def test_outstanding_under_two_won_is_zero():
    from dashboard.services.project_service import _calculate_outstanding_amount
    row = pd.Series({'총액 1': 10000000, '부가세': False, '계약금': 3333333, '중도금': 3333333, '잔금': 3333333})
    assert _calculate_outstanding_amount(row) == 0
    row2 = pd.Series({'총액 1': 10000000, '부가세': False, '계약금': 3333333, '중도금': 3333333, '잔금': 3333000})
    assert _calculate_outstanding_amount(row2) == 334


def test_as_manager_options_keep_sales_approvers(monkeypatch):
    """A/S 담당자 목록 — 비영업(황샛별·이미영)·KiKO 는 빠지고, 정산 승인자가 된 JW·YG 는 남는다(10-29)."""
    import dashboard.blueprints.slack_bot as sb
    import dashboard.utils.user_database as udb
    users = [{'email': f'{l}@itg-aircon.com', 'name': n} for l, n in
             [('sb', '황샛별'), ('my', '이미영'), ('kiko', '고광일'), ('jw', '박정우'), ('yg', '박용구')]]
    monkeypatch.setattr(udb, 'UserDatabase', lambda: type('D', (), {'get_all_users': lambda self: users})())
    vals = {o['value'] for o in sb._active_manager_options()}
    assert vals == {'jw@itg-aircon.com', 'yg@itg-aircon.com'}


class _C:
    def __init__(self):
        self.eph, self.msg = [], []

    def chat_postEphemeral(self, **k):
        self.eph.append(k)

    def chat_postMessage(self, **k):
        self.msg.append(k)


def test_modal_open_failed_notice_once_per_click():
    from dashboard.blueprints.slack_bot import _notify_modal_open_failed
    c = _C()
    body = {'user': {'id': 'U1'}, 'channel': {'id': 'C1'}, 'trigger_id': 't-1028-a'}
    _notify_modal_open_failed(c, body)
    _notify_modal_open_failed(c, body)          # 같은 클릭 → 한 번만
    assert len(c.eph) == 1 and c.eph[0]['channel'] == 'C1' and '한 번 더' in c.eph[0]['text']
    # 슬래시 명령(command) 형태
    _notify_modal_open_failed(c, {'user_id': 'U2', 'channel_id': 'C2', 'trigger_id': 't-1028-b'})
    assert c.eph[-1]['user'] == 'U2'


def test_as_tell_requester_uses_command_channel_then_dm():
    import json
    from dashboard.blueprints.slack_bot import _as_tell_requester
    c = _C()
    _as_tell_requester(c, {'user': {'id': 'U1'}, 'view': {'private_metadata': json.dumps({'channel_id': 'C9'})}}, 'ok')
    assert c.eph[0]['channel'] == 'C9'
    _as_tell_requester(c, {'user': {'id': 'U1'}, 'view': {'private_metadata': '{}'}}, 'ok')
    assert c.msg[0]['channel'] == 'U1'
