# -*- coding: utf-8 -*-
"""법인 간 이체(자금 이동 요청) — 메모 원장·짝짓기·매출이동 줄 자동 기록 (2026-10-04).

원장 = 시트 메모: '⚠️ 계좌·사업자 불일치 … — 자금 이동 요청' 줄 있는 입금 블록 = 대기,
뒤에 같은 방향 'YYYY-MM-DD R>G 매출이동' 줄 = 완료. 실데이터 G4125-YM(대연) 형식 사용.
"""
import json
import sys

sys.path.insert(0, '.')

from dashboard.services import fund_transfer as ft
from dashboard.services.payment_sync import _parse_notes

REQ = '⚠️ 계좌·사업자 불일치: 하나은행 (글로벌그룹) 입금 / 프로젝트 글로벌(G) — 자금 이동 요청 (지정 YM · 확인 SB)'
DEP_HANA = ('2026/09/30\n하나,09/30, 17:02\n255******31304\n입금 3,014,000원\n대연이엔지주식회\n' + REQ)
TR_LINE = '2026-10-02 R>G 매출이동 (하나 → 기업 법인 간 이체 · 처리 SB)'
GIUP_IN = '[Web발신]\n2026/10/02 15:01\n입금 3,014,000원\n대연이엔지\n452***38801011\n기업'
GIUP_PV = {'amount': 3014000, 'partner': '대연이엔지', 'bank': '기업', 'acct_code': 'G',
           'date_md': '10/02', 'date_year': '2026'}


def _pend(note):
    return ft.find_pending_in_note(note, 'G4125-YM', 4126, '잔금', 'W')


def test_names_match():
    assert ft.names_match('대연이엔지주식회', '대연이엔지')
    assert ft.names_match('(주)대연이엔지', '대연이엔지 주식회사')
    assert not ft.names_match('대연이엔지', '대한이엔지')
    assert not ft.names_match('대', '대연')          # 2글자 미만 판정 안 함


def test_request_block_is_pending():
    p = _pend(DEP_HANA)
    assert len(p) == 1
    r = p[0]
    assert (r['amount'], r['partner'], r['from'], r['to'], r['bank']) == (3014000, '대연이엔지주식회', 'R', 'G', '하나')
    assert r['date'] == '2026/09/30' and r['req_line'] == REQ and r['col'] == 'W'


def test_transfer_line_after_closes_request():
    assert _pend(DEP_HANA + '\n\n' + TR_LINE) == []


def test_wrong_direction_transfer_does_not_close():
    assert len(_pend(DEP_HANA + '\n\n2026-10-02 G>N 매출이동')) == 1


def test_two_requests_one_transfer_leaves_one():
    second = DEP_HANA.replace('09/30', '10/05').replace('3,014,000', '1,000,000')
    p = _pend(DEP_HANA + '\n\n' + TR_LINE + '\n\n' + second)
    assert len(p) == 1 and p[0]['amount'] == 1000000


def test_non_request_warnings_ignored():
    over = '2026/10/02 15:01\n입금 500,000원\n대연이엔지\n452***38801011\n기업\n⚠️ 과입금: 당시 미수금 300,000원 / 초과 200,000원 (지정 YM · 확인 SB)'
    unacked = DEP_HANA.replace('— 자금 이동 요청', '— 지정 시 경고 미표시')
    assert _pend(over) == [] and _pend(unacked) == []


def test_match_deposit():
    reqs = _pend(DEP_HANA)
    assert [c['code'] for c in ft.match_deposit(GIUP_IN, GIUP_PV, reqs)] == ['G4125-YM']
    assert ft.match_deposit(GIUP_IN, dict(GIUP_PV, amount=3000000), reqs) == []      # 금액 다름
    assert ft.match_deposit(GIUP_IN, dict(GIUP_PV, partner='다른회사'), reqs) == []   # 이름 다름
    hana_in = '하나,10/02, 15:01\n255******31304\n입금 3,014,000원\n대연이엔지'
    assert ft.match_deposit(hana_in, dict(GIUP_PV, acct_code='R'), reqs) == []        # 도착 사업자 다름


def test_insert_line_keeps_year_and_marks_transfer():
    later = '2026/10/10 09:00\n입금 500,000원\n대연이엔지\n452***38801011\n기업'
    note = DEP_HANA + '\n\n' + later
    new = ft.insert_transfer_line(note, REQ, TR_LINE)
    blocks = [b for b in new.split('\n\n')]
    assert blocks[1] == TR_LINE and blocks[2] == later        # 요청 블록 바로 뒤, 뒤 입금 보존
    p = _parse_notes(['', '', new], stage_vals={'잔금': 3514000})
    assert (p[0]['date_year'], p[0]['transfer_to'], p[0]['amount']) == ('2026', 'G', 3014000)
    assert p[1]['amount'] == 500000 and p[1]['transfer_to'] == ''
    assert _pend(new) == []


def test_transfer_line_format():
    r = _pend(DEP_HANA)[0]
    assert ft.transfer_line(r, GIUP_IN, GIUP_PV, 'SB') == TR_LINE


class _Mgr:
    def __init__(self, note):
        self.notes = {'W4126': note}

    def find_row_by_project_code(self, *a, **k):
        return 4126

    def get_cell_note(self, sid, sn, cell):
        return self.notes.get(cell, '')

    def update_cell_note(self, sid, sn, cell, n):
        self.notes[cell] = n
        return True


def test_apply_transfer_writes_once_and_rerenders(monkeypatch):
    m = _Mgr(DEP_HANA)
    rer = []
    import dashboard.services.lead_service as ls
    import dashboard.services.payment_sync as ps
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: m)
    monkeypatch.setattr(ps, 'rerender_stage_card', lambda c, s: rer.append((c, s)) or True)
    monkeypatch.setattr(ft, 'invalidate', lambda: None)
    r = _pend(DEP_HANA)[0]
    assert ft.apply_transfer(r, GIUP_IN, GIUP_PV, 'SB') == TR_LINE
    assert m.notes['W4126'] == DEP_HANA + '\n\n' + TR_LINE and rer == [('G4125-YM', '잔금')]
    ft.apply_transfer(r, GIUP_IN, GIUP_PV, 'SB')                     # 재처리
    assert m.notes['W4126'].count(TR_LINE) == 1


def test_intake_card_buttons():
    import dashboard.blueprints.sms_inbound as si
    pv = dict(GIUP_PV)
    plain = si._build_intake_blocks('iid', GIUP_IN, pv)
    assert [e['action_id'] for e in plain[-1]['elements']] == ['payment_intake_open', 'payment_intake_split_open']
    r = _pend(DEP_HANA)[0]
    sug = si._build_intake_blocks('iid', GIUP_IN, pv, transfer={'cands': [r], 'evidence': True})
    acts = next(b for b in sug if b['type'] == 'actions')['elements']
    assert acts[0]['action_id'] == 'payment_intake_transfer_0' and acts[0]['style'] == 'primary'
    assert json.loads(acts[0]['value']) == {'iid': 'iid', 'code': 'G4125-YM', 'stage': '잔금'}
    txt = ' '.join(b['text']['text'] for b in sug if b['type'] == 'section')
    assert '법인 간 이체로 보입니다' in txt and '출금 문자 확인됨' in txt
    pick = si._build_intake_blocks('iid', GIUP_IN, pv, transfer={'cands': [], 'evidence': False})
    assert pick[-1]['elements'][-1]['action_id'] == 'payment_intake_transfer_pick'


def test_done_card_and_reminder_text(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    monkeypatch.setattr(sb, '_resolve_manager_initial', lambda u: 'SB')
    r = _pend(DEP_HANA)[0]
    t = sb._build_intake_transfer_done_blocks(r, 3014000, GIUP_IN, 'U')[0]['text']['text']
    assert '법인 간 이체 처리됨' in t and '하나 (글로벌그룹) → 기업 (글로벌) · 3,014,000원 · 처리 SB' in t
    from dashboard.services.pin_remind import build_pin_remind_text
    rt = build_pin_remind_text({'moves': [r]})
    assert '자금 이동 대기 — 법인 간 이체 필요 (1건)' in rt
    assert 'G4125-YM 잔금 3,014,000원 · 대연이엔지주식회 · 하나(글로벌그룹) → 글로벌 · 2026/09/30 입금' in rt
