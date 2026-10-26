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

def test_no_move_line():
    """'이동 없이 처리' → 요청 아래 '⚠️ 자금 이동 없음' 줄 (2026-10-26 G3982)."""
    from datetime import datetime
    line = ft.no_move_line('그룹으로 발행 2026-10-23 R>G', 'SB', datetime(2026, 10, 23))
    assert line.startswith('⚠️ 자금 이동 없음: ') and line.endswith('(확인 SB 10/23)')
    from dashboard.services.payment_sync import _TRANSFER_RE
    assert not _TRANSFER_RE.search(line)                          # 매출이동으로 읽히지 않음
    later = '2026/10/10 09:00\n입금 500,000원\n대연이엔지\n452***38801011\n기업'
    new = ft.insert_no_move_line(DEP_HANA + '\n\n' + later, REQ, line)
    blocks = new.split('\n\n')
    assert blocks[0] == DEP_HANA + '\n' + line and blocks[1] == later   # 같은 입금 블록, 뒤 입금 보존
    assert _pend(new) == []                                       # 대기에서 빠짐
    p = _parse_notes(['', '', new], stage_vals={'잔금': 3514000})
    assert [x['amount'] for x in p] == [3014000, 500000]          # 금액 파싱 영향 없음
    assert ft.insert_no_move_line('다른 메모', REQ, line) is None


def test_apply_no_move(monkeypatch):
    m = _Mgr(DEP_HANA)
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'get_sheets_manager', lambda: m)
    monkeypatch.setattr(ft, 'invalidate', lambda: None)
    r = _pend(DEP_HANA)[0]
    line = ft.apply_no_move(r, '자금 이동없이 그룹으로 전액발행', 'SB')
    assert m.notes['W4126'] == DEP_HANA + '\n' + line and _pend(m.notes['W4126']) == []


def test_close_by_invoice_same_entity_only(monkeypatch):
    """그 단계 계산서를 돈이 들어온 법인(R) 이름으로 냈으면 이동 필요 없음 — 다른 법인·다른 단계는 그대로."""
    r = _pend(DEP_HANA)[0]                                      # 하나(R) 입금 → 프로젝트 G, 잔금
    closed = []
    monkeypatch.setattr(ft, 'pending_requests', lambda force=False: [r])
    monkeypatch.setattr(ft, 'apply_no_move', lambda req, reason, checker, when=None: closed.append((reason, checker)))
    assert ft.close_by_invoice('G4125-YM', '잔금', 'G') == [] and not closed       # 글로벌 발행 = 옮겨야 함
    assert ft.close_by_invoice('G4125-YM', '계약금', 'R') == [] and not closed     # 다른 단계
    assert ft.close_by_invoice('G4125-YM', '', 'R') == []                           # 단계 모름 = 안 닫음
    assert ft.close_by_invoice('G4125-YM', '계약금,잔금', 'R') == [r]
    assert closed == [('글로벌그룹 이름으로 계산서 발행 — 돈과 계산서 사업자가 같음', '자동')]
    from dashboard.services.itg_accounts import entity_for_bno
    assert entity_for_bno('533-88-02845') == 'R' and entity_for_bno('884-81-01425') == 'G'
    assert entity_for_bno('224-81-28302') == ''


def test_invoice_before_deposit_closed_by_recorded_issuer(monkeypatch):
    """계산서가 입금보다 먼저 나간 경우(G3982: 10/22 그룹 발행 → 10/23 하나 입금) — 발행 법인을 기억해 두고
    정산 알림 전에 대조해 닫는다."""
    store = {}

    class _R:
        def hset(self, k, mapping):
            store.update(mapping)

        def hget(self, k, f):
            return store.get(f)
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': _R()})())
    ft.record_invoice_issuer('G4125-YM', '잔금', 'R')
    assert ft.invoice_issuer('G4125-YM', '잔금') == 'R' and ft.invoice_issuer('G4125-YM', '계약금') == ''
    r = _pend(DEP_HANA)[0]
    closed = []
    monkeypatch.setattr(ft, 'pending_requests', lambda force=False: [r])
    monkeypatch.setattr(ft, 'apply_no_move', lambda req, reason, checker, when=None: closed.append(req['code']))
    assert ft.close_by_recorded_issuers() == [r] and closed == ['G4125-YM']
    ft.record_invoice_issuer('G4125-YM', '잔금', 'G')              # 글로벌로 (수정)발행 → 옮겨야 함
    closed.clear()
    assert ft.close_by_recorded_issuers() == [] and not closed


class _Now:                                 # 백그라운드 스레드를 바로 실행
    def __init__(self, target, daemon=None, args=()):
        self.t, self.a = target, args

    def start(self):
        self.t(*self.a)


class _SlackC:
    def __init__(self, codes_text='G4125-YM · 잔금'):
        self.codes_text, self.posted, self.eph = codes_text, [], []

    def conversations_replies(self, **k):
        return {'messages': [{'text': '분할 2건 기록 완료', 'blocks': [{'text': {'text': self.codes_text}}]}]}

    def chat_postMessage(self, **k):
        self.posted.append(k['text'])

    def chat_postEphemeral(self, **k):
        self.eph.append(k)


def test_reply_prompts_button_regardless_of_wording(monkeypatch):
    """경영지원 댓글이면 문구와 상관없이, 그 카드에 대기 요청이 있으면 본인에게만 [이동 없이 처리] 버튼."""
    import dashboard.blueprints.slack_bot as sb
    r = _pend(DEP_HANA)[0]
    monkeypatch.setattr(ft, 'pending_requests', lambda force=False: [r])
    monkeypatch.setattr(sb.threading, 'Thread', _Now)
    c = _SlackC()
    ev = {'user': sb._SETTLEMENT_CHECKER_ID, 'text': '그룹으로 전액 발행했어요', 'thread_ts': '1.0', 'ts': '2.0'}
    assert sb._prompt_fund_move_on_reply(c, 'CH', ev) is True
    blocks = c.eph[0]['blocks']
    btn = blocks[1]['elements'][0]
    assert btn['action_id'] == 'fund_move_no_move_0' and btn['text']['text'] == '이동 없이 처리'
    v = json.loads(btn['value'])
    assert (v['code'], v['stage'], v['amount'], v['from'], v['reason']) == ('G4125-YM', '잔금', 3014000, 'R', '그룹으로 전액 발행했어요')
    assert 'G4125-YM' in blocks[0]['text']['text'] and not c.posted               # 아직 기록 안 함
    # 경영지원이 아니면 기존 안내 경로(False)
    assert sb._prompt_fund_move_on_reply(_SlackC(), 'CH', dict(ev, user='U_OTHER')) is False


def test_reply_without_pending_gets_old_nudge(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    import dashboard.utils.redis_client as rcm
    monkeypatch.setattr(ft, 'pending_requests', lambda force=False: [])
    monkeypatch.setattr(sb.threading, 'Thread', _Now)
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': type('R', (), {'set': lambda *a, **k: True})()})())
    c = _SlackC()
    ev = {'user': sb._SETTLEMENT_CHECKER_ID, 'text': '확인했습니다', 'thread_ts': '1.0', 'ts': '2.0'}
    assert sb._prompt_fund_move_on_reply(c, 'CH', ev) is True
    assert c.eph[0]['text'] == sb._INTAKE_REPLY_NUDGE and 'blocks' not in c.eph[0]


def test_no_move_button_records_once(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    r = _pend(DEP_HANA)[0]
    state = {'pending': [r]}
    applied = []
    monkeypatch.setattr(ft, 'pending_requests', lambda force=False: list(state['pending']))

    def _apply(req, reason, checker, when=None):
        applied.append((req['code'], reason, checker))
        state['pending'] = []
    monkeypatch.setattr(ft, 'apply_no_move', _apply)
    monkeypatch.setattr(sb, '_resolve_manager_initial', lambda u: 'SB')
    c = _SlackC()
    v = {'code': 'G4125-YM', 'stage': '잔금', 'amount': 3014000, 'from': 'R', 'thread_ts': '1.0', 'reason': '그룹으로 발행'}
    sb._do_fund_move_no_move(c, 'U', 'CH', v)
    assert applied == [('G4125-YM', '그룹으로 발행', 'SB')] and '이동 없이 처리로 기록했습니다' in c.posted[0]
    sb._do_fund_move_no_move(c, 'U', 'CH', v)                    # 두 번 눌러도 한 번만
    assert len(applied) == 1 and len(c.posted) == 1


# 테스트 데이터를 옛 45열 위치(계약금=U·버전=AS 등)로 짰다 — 기본 배치를 45열로 고정해 검사
#   (2026-10-22 거래처명 열 삽입 뒤에도 같은 논리를 검사. 코드는 헤더 지도를 따라감)
import pytest as _pytest  # noqa: E402
pytestmark = _pytest.mark.usefixtures('sheet_layout_45')
