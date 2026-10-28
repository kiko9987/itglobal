# -*- coding: utf-8 -*-
"""2026-10-28 전체 점검 1차 수정 회귀 테스트.

① PM 공사 재개 = 확정일·수금 확인 복원 ② 칸 메모 충돌(그 사이 바뀐 메모 덮지 않음) ③ /공사확정 '플랜트'
④ 계산서 요약 covered 규칙(JS 와 같게) ⑥ 리드 시트 저장 실패 → 예외 ⑨ 금액·수금 칸 서버 권한.
"""
import sys

sys.path.insert(0, '.')

import pytest

from dashboard.blueprints import projects as pj


# ── ① 재개 ───────────────────────────────────────────────────────────
class _Mgr:
    def get_field_to_letter(self):
        return {'수금 관련 특이사항': 'AL', '공사 확정': 'AQ', '수금 확인': 'AE', '총액 1': 'S'}


def test_resume_restores_date_and_payment():
    ups = pj._prepare_resume_updates('공사 현황', 10, _Mgr(), restore_total_1=5000000,
                                     restore_date='2025-03-02', restore_payment=True)
    got = {u['range']: u['values'][0][0] for u in ups}
    assert got['공사 현황!AQ10'] == '2025-03-02'
    assert got['공사 현황!AE10'] == 'TRUE'
    assert got['공사 현황!S10'] == 5000000
    assert got['공사 현황!AL10'] == ''


def test_resume_without_snapshot_uses_today_and_leaves_payment():
    ups = pj._prepare_resume_updates('공사 현황', 10, _Mgr())
    got = {u['range']: u['values'][0][0] for u in ups}
    assert len(got['공사 현황!AQ10']) == 10 and '공사 현황!AE10' not in got


# ── ② 메모 충돌 ───────────────────────────────────────────────────────
def test_memo_norm_ignores_newline_and_trailing_space():
    assert pj._memo_norm('a  \r\nb\n') == pj._memo_norm('a\nb')
    assert pj._memo_norm('a\nb') != pj._memo_norm('a\nb\n\n2026/10/28 / 입금 100원')


class _NoteMgr:
    def __init__(self, notes):
        self.notes = notes

    def get_cell_note(self, sid, sn, addr):
        return self.notes.get(addr, '')


def test_batch_memo_conflict_not_written(monkeypatch):
    monkeypatch.setattr(pj, 'payment_field_column', lambda f: {'잔금': 'X', '계약금': 'V'}[f])
    mgr = _NoteMgr({'X5': '옛 메모\n\n2026/10/28 / 입금 1,000,000원', 'V5': '계약 메모'})
    memos = [
        {'field_name': '잔금', 'memo': '옛 메모\n추가 줄', 'original_memo': '옛 메모'},   # 그 사이 입금 줄 붙음
        {'field_name': '계약금', 'memo': '계약 메모 수정', 'original_memo': '계약 메모'},  # 그대로 → 저장
    ]
    results, failed, to_update, _old = pj._prepare_batch_memo_updates(memos, 5, mgr, 'sid', 'sn', 'G1')
    assert failed == 1 and 'X5' not in to_update and to_update == {'V5': '계약 메모 수정'}
    c = results[0]
    assert c['conflict'] is True and c['success'] is False and '입금 1,000,000원' in c['current_memo']


def test_batch_memo_without_original_keeps_old_behavior(monkeypatch):
    monkeypatch.setattr(pj, 'payment_field_column', lambda f: 'X')
    mgr = _NoteMgr({'X5': '다른 메모'})
    _r, failed, to_update, _o = pj._prepare_batch_memo_updates(
        [{'field_name': '잔금', 'memo': '새 메모'}], 5, mgr, 'sid', 'sn', 'G1')
    assert failed == 0 and to_update == {'X5': '새 메모'}


# ── ⑨ 금액·수금 칸 권한 ───────────────────────────────────────────────
F2I = {'총액 1': 0, '부가세': 1, '계약금': 2, '수금 확인': 3, '시공자': 4}
ROW = [11242000, True, 0, 'FALSE', '남진열']


def test_editor_cannot_change_money():
    msg, status = pj._money_edit_gate(ROW, F2I, {'총액 1': '11,542,000'}, 'editor')
    assert status == 403 and '총액 1' in msg


def test_editor_unchanged_money_values_pass():
    data = {'총액 1': '11,242,000', '부가세': 'true', '계약금': '-', '수금 확인': 'false', '시공자': '정건우'}
    assert pj._money_edit_gate(ROW, F2I, data, 'editor') is None


def test_admin_can_change_money():
    for perm in ('admin', 'super_admin'):
        assert pj._money_edit_gate(ROW, F2I, {'계약금': 1000000, '수금 확인': 'TRUE'}, perm) is None


# ── ④ 계산서 요약 ─────────────────────────────────────────────────────
def _y(stage_vals, amt, collected=False):
    from dashboard.blueprints.slack_bot import _bill_y_summary
    return _bill_y_summary(stage_vals, amt, collected)


def test_y_trailing_dash_after_last_issue_is_uninvoiced():
    # 계약금만 발행, 잔금 '-' 인데 잔금 입금 있음 → 잔금은 통합발행에 포함 아님 = 미발행 (R4112 규칙)
    y = _y({'계약금': '발행', '중도금': '-', '잔금': '-'}, {'계약금': 1000, '잔금': 4000})
    assert y == '잔금 - 미발행'


def test_y_leading_dash_before_issue_is_covered():
    # 잔금에 한 장으로 발행, 앞 단계 '-' = 포함
    y = _y({'계약금': '-', '중도금': '-', '잔금': '발행'}, {'계약금': 1000, '잔금': 4000}, collected=True)
    assert y == '잔금 - 발행완료'


# ── ③ /공사확정 '플랜트' → '플렌트' ───────────────────────────────────
def test_slack_company_plant_maps_to_sheet_value():
    from dashboard.blueprints import slack_bot as sb
    import json
    cfg = json.load(open('dashboard/project_config.json', encoding='utf-8'))['company_prefix_map']
    for opt in sb._PROJECT_COMPANY_OPTIONS:
        assert sb._PROJECT_COMPANY_SHEET_VALUE.get(opt, opt) in cfg


# ── ⑥ 리드 저장 실패 → 예외 ───────────────────────────────────────────
def test_update_lead_dispatch_raises_on_failure(monkeypatch):
    from dashboard.blueprints import slack_bot as sb
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'update_lead', lambda no, up: {'success': False, 'message': '시트 오류'})
    with pytest.raises(RuntimeError, match='시트 오류'):
        sb._update_lead_dispatch('L-1', {'상태': '방문 예약'})
    monkeypatch.setattr(ls, 'update_lead', lambda no, up: {'success': True})
    sb._update_lead_dispatch('L-1', {'상태': '방문 예약'})
