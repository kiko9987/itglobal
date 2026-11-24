"""법인 간 이체 결재(2026-11-16) — 자금 이동 요청 → 매니저 DM → 채워진 창 → 결재 연결·중복 막기."""
import json

import pytest

from dashboard.services import e_approval as ea
from dashboard.services import fund_transfer as ft
from test_e_approval import (  # noqa: F401 — 픽스처 재사용
    KIKO, MGR, SB, _new_view, _no_final_by_default, se, tmp_db)

REQ = {'code': 'R4292-SJ', 'row': 4293, 'stage': '계약금', 'col': 'V', 'amount': 3102000, 'partner': '민영준',
       'bank': '기업', 'from': 'G', 'to': 'R', 'date': '2026/11/16', 'req_line': '⚠️ … 자금 이동 요청'}


@pytest.fixture
def pending(monkeypatch):
    reqs = [dict(REQ)]
    monkeypatch.setattr(ft, 'pending_requests', lambda force=False, strict=False: reqs)
    monkeypatch.setattr(ft, 'invalidate', lambda: None)
    return reqs


def _move_view(meta, content=None, account='255-910014-31304'):
    vals = {'content': {'content': {'value': content if content is not None else ea.move_content(meta['move'])}},
            'payee_name': {'payee_name': {'value': '(주)아이티글로벌그룹'}},
            'bank': {'bank': {'selected_option': {'value': '하나은행'}}},
            'account': {'account': {'value': account}},
            'urgency': {'urgency': {'selected_option': {'value': '일반'}}},
            'memo': {'memo': {'value': None}}}
    return {'id': 'VM', 'private_metadata': json.dumps(meta, ensure_ascii=False), 'state': {'values': vals}}


def test_move_view_fills_everything_and_links(se, pending):
    meta = {}
    se._set_move(meta, ea.open_moves('R4292-SJ')[0])
    view = se._new_view(meta)
    txt = json.dumps(view, ensure_ascii=False)
    assert view['title']['text'] == '법인 간 이체' and '3,102,000원' in txt and '민영준 11/16' in txt
    assert '255-910014-31304' in txt and '하나은행' in txt
    bids = {b.get('block_id') for b in view['blocks']}
    assert not bids & {'project', 'cost', 'evidence', 'grn', 'amount', 'partner', 'invoice'}   # 정해진 값은 숨김
    f, errs = se.validate_new_submit(_move_view(meta))
    assert errs == {}
    assert (f['project_code'], f['cost_category'], f['grn'], f['amount']) == ('R4292-SJ', ea.COST_MOVE, 'G', 3102000)
    assert f['evidence'] == ea.EVIDENCE_MOVE and f['partner'] == '(주)아이티글로벌그룹'
    row = se.process_new_submit(MGR, f)
    assert ea.extra(row)['fund_move']['key'] == ea.move_key(REQ)
    card = ea.card_text(row, ea.steps(row['id']), 'R4292-SJ')
    assert '*옮기는 입금* : R4292-SJ 계약금 입금(민영준 11/16 3,102,000원) · 글로벌 통장 → 글로벌그룹 통장' in card
    assert '법인 간 이체 (원가 아님)' in card and '증빙' not in card
    # 같은 입금으로 또 올리면 막는다(창 다시 열기·직접 입력 둘 다)
    assert ea.open_moves('R4292-SJ') == []
    _, errs = se.validate_new_submit(_move_view(meta))
    assert '이미 있어요' in errs['content']


def test_manual_move_links_matching_request_and_checks_direction(se, pending):
    v = _new_view(project='R4292-SJ', partner='(주)아이티글로벌그룹', amount='3,102,000', grn='R',
                  payee_name='(주)아이티글로벌그룹', bank='하나은행', account='255-910014-31304')
    v['state']['values']['cost'] = {ea.A_COST: {'selected_option': {'value': ea.COST_MOVE}}}
    v['id'] = 'V2'
    f, errs = se.validate_new_submit(v)
    assert 'grn' in errs and '잘못 들어온 법인' in errs['grn']        # 공사 법인(R) 통장에서 R 로는 못 보낸다
    v['state']['values']['grn'] = {'grn': {'selected_option': {'value': 'G'}}}
    f, errs = se.validate_new_submit(v)
    assert errs == {} and f['evidence'] == ea.EVIDENCE_MOVE
    assert f['extra']['fund_move']['key'] == ea.move_key(REQ)          # 금액·방향이 맞는 요청에 자동 연결
    # G 공사 통장 그대로(그룹 → 글로벌 아님)는 grn_mismatch 경고 대신 방향 검사만
    other = _new_view(project='R4292-SJ', amount='10,000', grn='G', payee_name='(주)아이티글로벌그룹',
                      bank='하나은행', account='255-910014-31304')
    other['state']['values']['cost'] = {ea.A_COST: {'selected_option': {'value': ea.COST_MOVE}}}
    other['id'] = 'V3'
    f2, errs2 = se.validate_new_submit(other)
    assert errs2 == {} and 'fund_move' not in (f2.get('extra') or {})   # 맞는 요청이 없으면 연결 없이


def test_cost_pick_switches_view(se, pending):
    meta = {}
    moves = ea.open_moves('R4292-SJ')
    se._set_move(meta, moves[0])
    assert meta['project'] == 'R4292-SJ' and meta['payee']['bank'] == '하나은행'
    meta2 = {}
    se._set_move_manual(meta2, 'R4292-SJ')
    view = se._new_view(meta2)
    txt = json.dumps(view, ensure_ascii=False)
    assert '돈이 잘못 들어온 법인' in txt and '255-910014-31304' in txt
    meta3 = {'move_opts': [ea.move_option(m) for m in moves]}
    blk = next(b for b in se._new_view(meta3)['blocks'] if b.get('block_id') == 'move_pick')
    assert [o['value'] for o in blk['element']['options']] == [moves[0]['key'], ea.MOVE_PICK_MANUAL]
    cost = next(b for b in se._new_view({})['blocks'] if b.get('block_id') == 'cost')
    assert cost['dispatch_action'] and ea.COST_MOVE in [o['value'] for o in cost['element']['options']]


def test_purchase_invoice_picker_has_no_move_category():
    from dashboard.services import purchase_invoice_assign as pinv
    assert ea.COST_MOVE not in [v for v, _ in pinv.COST_CATEGORIES]
    assert pinv.cost_label(ea.COST_MOVE) == '법인 간 이체 (원가 아님)'


def test_notify_dm_once_and_marked_done_after_submit(se, pending):
    assert se.notify_fund_move(MGR, 'R4292-SJ') == 1
    assert se.notify_fund_move(MGR, 'R4292-SJ') == 0                    # 같은 요청엔 한 번
    dm = [kw for kw in se._fake.of('chat_postMessage') if kw['channel'] == 'D-' + MGR][0]
    btn = dm['blocks'][1]['elements'][0]
    assert btn['action_id'] == ea.A_MOVE_OPEN and json.loads(btn['value'])['key'] == ea.move_key(REQ)
    assert '민영준 11/16 3,102,000원' in dm['blocks'][0]['text']['text']
    meta = {}
    se._set_move(meta, json.loads(btn['value']))
    f, errs = se.validate_new_submit(_move_view(meta))
    row = se.process_new_submit(MGR, f)
    upd = next(kw for kw in se._fake.of('chat_update') if kw['channel'] == 'D-' + MGR)
    assert upd['ts'] == json.loads(ea.kv_get('move_dm:' + ea.move_key(REQ)))['ts']
    assert '결재를 올렸어요' in upd['blocks'][0]['text']['text'] and len(upd['blocks']) == 1   # 버튼 걷음
    assert ea.title(row) in upd['blocks'][0]['text']['text']


def test_no_dm_without_move_note(monkeypatch):
    from dashboard.blueprints import slack_bot as sb
    called = []
    monkeypatch.setattr('dashboard.blueprints.slack_eapproval.notify_fund_move', lambda u, c: called.append(c))
    sb._notify_fund_move_bg('UMGR', 'R4292-SJ', ['⚠️ 과입금 …'])
    sb._notify_fund_move_bg('', 'R4292-SJ', ['… — 자금 이동 요청'])
    assert called == []


def test_own_entity_name_matches_transfer_deposit():
    assert ft.is_own_entity_name('(주)아이티글로벌') and ft.is_own_entity_name('아이티글로벌그룹')
    assert ft.is_own_entity_name('아이티글로') and not ft.is_own_entity_name('아이티')
    assert not ft.is_own_entity_name('아이티플레이') and not ft.is_own_entity_name('민영준')
    text = '[Web발신]\n하나 255-910014-31304 입금 3,102,000원 아이티글로벌'
    reqs = [dict(REQ)]
    assert ft.match_deposit(text, {'amount': 3102000, 'partner': '아이티글로벌', 'acct_code': 'R'}, reqs) == reqs
    assert ft.match_deposit(text, {'amount': 3102000, 'partner': '민영준', 'acct_code': 'R'}, reqs) == reqs
    assert ft.match_deposit(text, {'amount': 3102000, 'partner': '홍길동', 'acct_code': 'R'}, reqs) == []
    assert ft.match_deposit(text, {'amount': 3100000, 'partner': '아이티글로벌', 'acct_code': 'R'}, reqs) == []


def test_invoice_result_has_no_reassign_button_when_feature_off(monkeypatch):
    from dashboard.services import purchase_invoice_assign as pinv
    from dashboard.blueprints import slack_bot as sb
    row = {'assigned_at': '2026-11-17 08:15:00', 'assigned_name': '문원', 'cost_category': '제품대',
           'memo': '중고', 'history': '[]', 'date_md': '11/17', 'grn': 'R', 'amount': 1, 'partner': 'x'}
    off = pinv.result_blocks(row, 'R4298-MW', 'C1', '1.1', '💸 #36', with_button=False)
    assert 'pinv_assign_open' not in json.dumps(off) and '#36' in json.dumps(off, ensure_ascii=False)
    assert 'pinv_assign_open' in json.dumps(pinv.result_blocks(row, 'R4298-MW', 'C1', '1.1'))
    import inspect
    assert 'with_button=_pinv_enabled()' in inspect.getsource(sb._post_pinv_result)


def test_reminder_line_shows_approval_state(se, pending):
    from dashboard.services import pin_remind as pr
    import inspect
    src = inspect.getsource(pr)
    assert '결재 안 올라감' in src
    assert ea.find_move_approval(ea.move_key(REQ)) is None
    meta = {}
    se._set_move(meta, ea.open_moves('R4292-SJ')[0])
    f, _ = se.validate_new_submit(_move_view(meta))
    row = se.process_new_submit(MGR, f)
    assert ea.find_move_approval(ea.move_key(REQ))['id'] == row['id']
    se.apply_action('reject', row['id'], KIKO, '다시')
    assert ea.find_move_approval(ea.move_key(REQ)) is None             # 반려되면 다시 올릴 수 있다


def test_manual_move_blocks_when_requests_unreadable(se, monkeypatch):
    """자금 이동 요청을 못 읽으면 '요청 없음'으로 넘기지 않고 제출을 막는다(중복 검사가 꺼지지 않게, 11-24 점검 6번)."""
    def boom(force=False, strict=False):
        if strict:
            raise RuntimeError('sheet 502')
        return []
    monkeypatch.setattr(ft, 'pending_requests', boom)
    v = _new_view(project='R4292-SJ', partner='(주)아이티글로벌그룹', amount='3,102,000', grn='G',
                  payee_name='(주)아이티글로벌그룹', bank='하나은행', account='255-910014-31304')
    v['state']['values']['cost'] = {ea.A_COST: {'selected_option': {'value': ea.COST_MOVE}}}
    v['id'] = 'V9'
    f, errs = se.validate_new_submit(v)
    assert '읽지 못했어요' in errs.get('content', '')
