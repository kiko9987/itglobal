"""보안 점검 4차(2026-11-14) — 매입 계산서 → 자금집행: 두 번 이체·계좌 바뀜·계산서 글 출처·[다시 지정] 막기."""
import json
from datetime import datetime, timedelta

import pytest

from dashboard.services import e_approval as ea
from tests.unit.test_e_approval import (  # noqa: F401 — 픽스처 재사용
    FIELDS, KIKO, MGR, SB, _new_view, _no_final_by_default, se, tmp_db)


def _view(view_id='V1', **over):
    v = _new_view(payee_name='김강원', bank='신한', account='110-511-825643', **over)
    v['id'] = view_id
    return v


# ── ① 같은 거래처·금액 ──────────────────────────────────────────

def test_no_invoice_path_now_warns_once_then_passes(se):
    first = se.process_new_submit(MGR, dict(FIELDS, partner='강원포장', amount=150000))
    # 예전엔 「계산서 없음」이면 중복 검사를 아예 건너뛰었다
    f, errs = se.validate_new_submit(_view())
    assert set(errs) == {'amount'} and ea.title(first) in errs['amount']
    f, errs = se.validate_new_submit(_view())          # 같은 창에서 한 번 더 → 통과(정말 또 보내는 경우)
    assert errs == {}
    second = se.process_new_submit(MGR, f)
    assert ea.extra(second)['dup_of'] == [first['id']]
    card = ea.card_text(ea.get(second['id']), ea.steps(second['id']), 'G4049-JW')
    assert '최근 같은 거래처·금액 결재' in card and ea.title(first) in card


def test_dup_ignores_other_amounts_rejected_and_old(se):
    a = se.process_new_submit(MGR, dict(FIELDS, partner='(주)강원포장', amount=150000))
    se.process_new_submit(MGR, dict(FIELDS, partner='강원포장', amount=150001))      # 금액 다름
    assert [d['id'] for d in ea.find_recent_fund_dups(150000, '강원 포장')] == [a['id']]   # 표기 차이는 같은 곳
    se.apply_action('reject', a['id'], KIKO, '중복')
    assert ea.find_recent_fund_dups(150000, '강원포장') == []                         # 반려는 제외
    b = se.process_new_submit(MGR, dict(FIELDS, partner='강원포장', amount=150000))
    old = (datetime.now() - timedelta(days=ea.DUP_WINDOW_DAYS + 1)).strftime('%Y-%m-%d %H:%M:%S')
    with ea._db()._get_connection() as conn:
        conn.execute('UPDATE approvals SET created_at=? WHERE id=?', (old, b['id']))
        conn.commit()
    assert ea.find_recent_fund_dups(150000, '강원포장') == []                         # 90일 지난 건 제외


def test_resubmit_recomputes_dup_without_logging_it_as_change(se):
    a = se.process_new_submit(MGR, dict(FIELDS, amount=100000))
    b = se.process_new_submit(MGR, dict(FIELDS, amount=200000))
    assert not ea.extra(b).get('dup_of')
    new, changed = ea.resubmit(b['id'], MGR, '강성환', {'amount': 100000})
    assert ea.extra(new)['dup_of'] == [a['id']]
    assert set(changed) == {'amount'}                 # 표시 칸(dup_of)은 '바뀐 칸' 기록에 안 들어감


def test_old_done_prepayment_no_longer_hides_new_invoice(se):
    p = se.process_new_submit(KIKO, dict(FIELDS))     # 승인자 본인 요청 → 바로 이체 대기
    se.apply_action('transfer', p['id'], SB)
    assert ea.find_unlinked_match('G', FIELDS['amount'], FIELDS['partner'])['id'] == p['id']
    old = (datetime.now() - timedelta(days=ea.DUP_WINDOW_DAYS + 1)).strftime('%Y-%m-%d %H:%M:%S')
    with ea._db()._get_connection() as conn:
        conn.execute('UPDATE approvals SET updated_at=? WHERE id=?', (old, p['id']))
        conn.commit()
    assert ea.find_unlinked_match('G', FIELDS['amount'], FIELDS['partner']) is None


# ── ② 계좌 바뀜 ────────────────────────────────────────────────

def test_payee_changed_flag_card_and_transfer_confirm(se):
    ea.remember_payee(FIELDS['partner'], '김강원', '신한', '110-511-825643')
    f = dict(FIELDS, account='999-888-777666')
    ea.mark_payee_saved(f)
    assert f['extra']['payee_changed'] == '신한 …5643' and not f['extra'].get('payee_saved')
    row = se.process_new_submit(KIKO, f)              # 이체 대기
    card = ea.card_text(row, ea.steps(row['id']), 'G4049-JW')
    assert '예전에 보낸 계좌(신한 …5643)와 달라요' in card
    btn = next(b for b in ea.buttons(row) if b['action_id'] == ea.A_TRANSFER)
    assert '예전에 보낸 계좌' in btn['confirm']['text']['text']
    g = dict(FIELDS)                                   # 같은 계좌면 표시 없음
    ea.mark_payee_saved(g)
    assert g['extra'].get('payee_saved') and 'payee_changed' not in g['extra']


def test_same_account_with_listed_bank_name_is_not_changed(tmp_db):
    """11-20 #65 홍일 PVC — 예전에 '기업은행'으로 적은 계좌가 창에서 'IBK기업은행'으로 채워져 같은 계좌인데 ⚠️."""
    ea.remember_payee('홍일P.V.C', '이남일(홍일P.V.C)', '기업은행', '296-105209-04-016')
    f = dict(FIELDS, partner='홍일 PVC', payee_name='이남일(홍일P.V.C)', bank='IBK기업은행', account='296-105209-04-016')
    ea.mark_payee_saved(f)
    assert f['extra'].get('payee_saved') and 'payee_changed' not in f['extra']
    g = dict(f, bank='신한은행', extra={})                   # 은행이 정말 다르면 그대로 ⚠️
    ea.mark_payee_saved(g)
    assert g['extra'].get('payee_changed')
    assert ea.same_bank('기업', 'IBK기업은행') and not ea.same_bank('기업은행', '신한은행')


def test_invoice_picked_evidence_defaults_when_slack_drops_it():
    """11-17~20 #50·51·57·65·66 — 계산서를 고르면 증빙에 '전자계산서'가 보이지만 슬랙이 빈 값을 넘겨 '증빙: -'."""
    view = {'state': {'values': {}}, 'private_metadata': '{"inv": {"ts": "1.2", "ch": "C1", "parsed": {"amount": 1000, "partner": "x", "grn": "G"}}}'}
    assert ea.read_new_request(view)['evidence'] == '전자계산서'
    view['private_metadata'] = '{}'                                 # 계산서 안 고른 선입금은 비운 그대로
    assert ea.read_new_request(view)['evidence'] == ''
    view = {'state': {'values': {'evidence': {'evidence': {'selected_option': {'value': '기타영수증'}}}}},
            'private_metadata': '{"inv": {"ts": "1.2", "ch": "C1", "parsed": {"amount": 1000, "partner": "x", "grn": "G"}}}'}
    assert ea.read_new_request(view)['evidence'] == '기타영수증'      # 직접 고른 값은 그대로


def test_first_time_partner_has_no_change_flag(tmp_db):
    f = dict(FIELDS, partner='처음보는상사')
    ea.mark_payee_saved(f)
    assert 'payee_changed' not in f['extra']


def test_known_accounts_list_replaces_changed_warning(se):
    """11-15 하이엠솔루텍 — 센터마다 계좌가 달라 다른 센터로 보낼 때마다 '예전 계좌와 달라요'가 떴다(#25)."""
    p = '하이엠솔루텍㈜'
    ea.remember_payee(p, '하이엠솔루텍', '우리은행', '1005-401-790194')            # 예전에 천안으로 보냄
    assert ea.add_known_accounts(p, [('우리은행', '1005-401-790194', '천안'), ('우리은행', '1005-201-085715', '강남'),
                                     ('우리은행', '1005-101-247544', '어음'), ('우리은행', '12-3', '짧음')]) == 3
    f = dict(FIELDS, partner='하이엠솔루텍', payee_name='하이엠솔루텍', bank='우리', account='1005201085715')
    ea.mark_payee_saved(f)                                # 다른 센터(강남) — 표기가 달라도(우리·하이픈 없음) 같은 계좌
    x = f['extra']
    assert x['payee_known'] == '강남' and 'payee_changed' not in x and 'payee_unlisted' not in x
    row = se.process_new_submit(KIKO, f)
    card = ea.card_text(row, ea.steps(row['id']), 'G4049-JW')
    # 계좌 경고는 하나도 없어야 한다(계산서 글 미선택 ⚠️ 는 별개 — test_card_warns_when_invoice_post_not_picked)
    assert '✅ 확인된 계좌(강남)' in card
    assert not any(w in card for w in ('예전에 보낸 계좌', '계좌 목록에 없는 계좌', '어음 계좌'))
    btn = next(b for b in ea.buttons(row) if b['action_id'] == ea.A_TRANSFER)
    assert '확인된 계좌(강남)' in btn['confirm']['text']['text']

    g = dict(f, account='1005-999-000000', extra={})     # 목록에 없는 계좌
    ea.mark_payee_saved(g)
    assert g['extra']['payee_unlisted'] and 'payee_known' not in g['extra'] and 'payee_changed' not in g['extra']
    row = se.process_new_submit(KIKO, g)
    assert '계좌 목록에 없는 계좌예요' in ea.card_text(row, ea.steps(row['id']), 'G4049-JW')

    h = dict(f, bank='국민', extra={})                    # 번호가 같아도 은행이 다르면 다른 계좌
    ea.mark_payee_saved(h)
    assert h['extra'].get('payee_unlisted')

    k = dict(f, account='1005-101-247544', extra={})     # 어음 계좌
    ea.mark_payee_saved(k)
    row = se.process_new_submit(KIKO, k)
    assert '어음 계좌예요' in ea.card_text(row, ea.steps(row['id']), 'G4049-JW')
    assert '어음 계좌예요' in next(b for b in ea.buttons(row) if b['action_id'] == ea.A_TRANSFER)['confirm']['text']['text']

    # 목록 없는 거래처는 예전처럼
    other = dict(FIELDS, account='999-888-777666', extra={})
    ea.remember_payee(FIELDS['partner'], '김강원', '신한', '110-511-825643')
    ea.mark_payee_saved(other)
    assert other['extra']['payee_changed'] and 'payee_unlisted' not in other['extra']


def test_edit_to_listed_account_clears_unlisted_warning(se):
    ea.add_known_accounts('하이엠솔루텍', [('우리은행', '1005-201-085715', '강남')])
    base = dict(FIELDS, partner='하이엠솔루텍', payee_name='하이엠솔루텍', bank='우리은행')
    f = dict(base, account='1005-999-000000')
    ea.mark_payee_saved(f)
    row = se.process_new_submit(MGR, f)
    assert ea.extra(row).get('payee_unlisted')
    g = dict(base, account='1005-201-085715', _edit=row['id'])
    ea.mark_payee_saved(g)
    se.process_new_submit(MGR, g)
    x = ea.extra(ea.get(row['id']))
    assert not x.get('payee_unlisted') and x.get('payee_known') == '강남'


def _hiem(monkeypatch):
    monkeypatch.setattr(ea, 'known_partners', lambda: ['하이엠솔루텍', '하이마트', '대성크레인'])
    ea.remember_payee('하이엠솔루텍', '하이엠솔루텍', '우리은행', '1005-201-085715')     # 지난번=강남
    ea.add_known_accounts('하이엠솔루텍', [('우리은행', '1005-201-085715', '강남'),
                                          ('우리은행', '1005-301-085723', '인천'),
                                          ('우리은행', '1005-101-247544', '어음')])


def test_partner_list_shows_each_branch(tmp_db, monkeypatch):
    """11-16 KiKO — 지점마다 계좌가 다른 거래처는 거래처 목록에서 지점을 바로 고른다."""
    _hiem(monkeypatch)
    texts = [o['text']['text'] for o in ea.partner_options('하이엠')['options']]
    assert texts[:4] == ['하이엠솔루텍 · 강남 (우리 …5715)', '하이엠솔루텍 · ⚠️ 어음 계좌 (우리 …7544)',
                         '하이엠솔루텍 · 인천 (우리 …5723)', '하이엠솔루텍 · 목록에 없는 계좌 (직접 입력)']
    assert '하이엠솔루텍' not in texts and texts[-1] == '새 거래처: 하이엠'
    assert [o['text']['text'] for o in ea.partner_options('하이엠 인천')['options']][:1] == ['하이엠솔루텍 · 인천 (우리 …5723)']
    assert ea.partner_options('인천')['options'][0]['text']['text'].endswith('…5723)')
    opts = ea.partner_options('하이')['options']                         # 지점 없는 거래처는 예전처럼
    assert '하이마트' in [o['value'] for o in opts]
    v = ea.partner_options('하이엠 인천')['options'][0]['value']
    assert ea.partner_from_value(v) == '하이엠솔루텍' and ea.branch_from_value(v) == ('하이엠솔루텍', '1005301085723')
    assert ea.branch_from_value('하이마트') is None and ea.branch_from_value('acct:깨짐') is None


def test_invoice_post_assign_modal_picks_branch_not_last_account(se, monkeypatch):
    """11-18 계산서 글 [📍 현장 지정]→「이체 필요」 — 하이엠솔루텍 수원 계산서에 지난번(인천) 계좌가 채워졌다.
    지점 계좌 칸을 보여 주고 지난번 계좌는 채우지 않으며, 고른 지점 계좌로 칸을 다시 채운다."""
    from dashboard.blueprints import slack_bot as sb
    from dashboard.services import purchase_invoice_assign as pinv
    _hiem(monkeypatch)
    monkeypatch.setattr(pinv, 'get_assignment', lambda ch, ts: None)
    monkeypatch.setattr(ea, 'find_unlinked_match', lambda *a, **k: None)
    meta = {'ch': 'CINV', 'ts': '1.1', 'v': pinv.button_value('CINV', '1.1', {
        'date_md': '11/18', 'grn': 'R', 'amount': 28000, 'partner': '하이엠솔루텍'})}
    meta.update(json.loads(meta.pop('v')))
    v = sb._pinv_modal_for(meta, transfer=True)
    txt = json.dumps(v, ensure_ascii=False)
    assert '지점 계좌' in txt and ea.PAYEE_SAVED_NOTE not in txt
    acct = next(b for b in v['blocks'] if (b.get('block_id') or '').startswith('account'))
    assert 'initial_value' not in acct['element']                      # 지난번(강남) 계좌 안 채움
    meta['br'], meta['bpv'] = '1005301085723', 1                       # 인천 고름
    v = sb._pinv_modal_for(meta, transfer=True)
    acct = next(b for b in v['blocks'] if (b.get('block_id') or '').startswith('account'))
    assert acct['block_id'] == 'account_b1' and acct['element']['initial_value'] == '1005301085723'
    br = next(b for b in v['blocks'] if b.get('block_id') == 'pinv_branch')
    assert br['element']['initial_option']['text']['text'] == '하이엠솔루텍 · 인천 (우리 …5723)'
    ea.add_known_accounts('없는거래처', [])
    meta2 = dict(meta, br=None, bpv=0)
    meta2.update(json.loads(pinv.button_value('CINV', '1.2', {'date_md': '', 'grn': 'R', 'amount': 1, 'partner': '대성크레인'})))
    assert 'pinv_branch' not in json.dumps(sb._pinv_modal_for(meta2, transfer=True))   # 지점 없는 거래처는 그대로


def test_branch_pick_fills_that_branch_not_last_one(se, monkeypatch):
    _hiem(monkeypatch)
    meta = {}
    se._set_payee_prefill(meta, '하이엠솔루텍')                          # 상호만 고름 → 지난번(강남) 안 채움
    assert meta['payee'] is None and meta['branch_list']
    head = ea.new_request_view(meta)['blocks']
    assert any('지점마다 계좌가 다른 거래처' in (b.get('text') or {}).get('text', '') for b in head)
    se._set_payee_prefill(meta, '하이엠솔루텍', '1005301085723')   # 인천 고름
    assert meta['payee'] == {'payee_name': '하이엠솔루텍', 'bank': '우리은행', 'account': '1005301085723'}
    view = ea.new_request_view(meta)
    txt = json.dumps(view, ensure_ascii=False)
    assert '✅ 확인된 계좌(인천)' in txt and '하이엠솔루텍 · 인천 (우리 …5723)' in txt
    assert ea.PAYEE_SAVED_NOTE not in txt
    se._set_payee_prefill(meta, '하이엠솔루텍', '')                      # 목록에 없는 계좌(직접 입력)
    assert meta['payee'] is None and '직접 입력' in meta['partner_opt']['text']['text']
    se._set_payee_prefill(meta, '대성크레인')                            # 목록 없는 거래처는 예전처럼
    assert not meta.get('branch_list') and 'partner_opt' not in meta


def test_invoice_partner_with_branches_gets_branch_select(se, monkeypatch):
    _hiem(monkeypatch)
    meta = {'inv': {'ch': 'C1', 'ts': '1.1', 'parsed': {'partner': '하이엠솔루텍', 'amount': 1000, 'grn': 'G'}}}
    se._set_payee_prefill(meta, '하이엠솔루텍')
    blk = next(b for b in ea.new_request_view(meta)['blocks'] if b.get('block_id') == 'branch')
    assert blk['element']['action_id'] == ea.A_BRANCH and len(blk['element']['options']) == 4
    se._set_payee_prefill(meta, '하이엠솔루텍', '1005201085715')
    blk = next(b for b in ea.new_request_view(meta)['blocks'] if b.get('block_id') == 'branch')
    assert blk['element']['initial_option']['text']['text'].startswith('하이엠솔루텍 · 강남')


# ── ③ 계산서 글 작성자 ──────────────────────────────────────────

@pytest.mark.parametrize('author,other', [(SB, False), (MGR, True), ('', True)])
def test_invoice_author_on_card(se, monkeypatch, author, other):
    monkeypatch.setenv('EAPPROVAL_EXECUTOR_DM_IDS', SB)
    row = se.process_new_submit(MGR, dict(FIELDS, origin_channel='CINV', origin_ts='1700000000.000100'))
    msg = {'ts': '1700000000.000100', 'user': author, 'files': []}
    monkeypatch.setattr(se, '_invoice_post', lambda ch, ts: msg)
    refreshed = []
    monkeypatch.setattr(se, 'refresh', lambda r: refreshed.append(r['id']))
    se.post_invoice_files(row)
    x = ea.extra(ea.get(row['id']))
    assert bool(x['inv_by_other']) is other and refreshed == [row['id']]
    card = ea.card_text(ea.get(row['id']), ea.steps(row['id']), 'G4049-JW')
    assert '*계산서 글* :' in card and (('경영지원이 올린 글이 아니에요' in card) is other)


def test_invoice_author_unknown_when_post_unreadable(se, monkeypatch):
    row = se.process_new_submit(MGR, dict(FIELDS, origin_channel='CINV', origin_ts='1700000000.000100'))
    monkeypatch.setattr(se, '_invoice_post', lambda ch, ts: None)
    se.post_invoice_files(row)
    assert 'inv_by' not in ea.extra(ea.get(row['id']))      # 못 읽으면 아무 표시도 꾸미지 않음


# ── 화면 계산서와 실제 계산서 ────────────────────────────────────

def test_invoice_screen_mismatch(se):
    a = {'ch': 'C', 'ts': '1.1', 'parsed': {'grn': 'G', 'amount': 1, 'partner': 'x'}}
    b = dict(a, ts='2.2')
    view = {'state': {'values': {'invoice': {ea.A_INVOICE: {'selected_option': ea.invoice_option(b)}}}}}
    assert se._invoice_screen_mismatch(view, {'origin_ts': '1.1'})          # 화면 B, 올라갈 건 A
    assert se._invoice_screen_mismatch(view, {'origin_ts': '2.2'}) == ''
    none = {'state': {'values': {'invoice': {ea.A_INVOICE: {'selected_option': ea.invoice_option(None)}}}}}
    assert se._invoice_screen_mismatch(none, {'origin_ts': '1.1'})          # 화면 '계산서 없음', 올라갈 건 A
    assert se._invoice_screen_mismatch(none, {}) == ''
    assert se._invoice_screen_mismatch({'state': {'values': {}}}, {'origin_ts': '1.1'}) == ''   # 칸 없는 창


# ── ④ [다시 지정] ──────────────────────────────────────────────

def _body(uid, ch='CINV', ts='1700000000.000100'):
    return {'user': {'id': uid}, 'actions': [{'value': json.dumps({'ch': ch, 'ts': ts})}]}


def test_pinv_deny_when_feature_off(monkeypatch):
    from dashboard.blueprints import slack_bot as sb
    monkeypatch.setenv('PURCHASE_INVOICE_ASSIGN_ENABLED', 'false')
    assert '쓰지 않아요' in sb._pinv_deny(_body(KIKO))


def test_pinv_deny_reassign_only_owner_or_office(tmp_db, monkeypatch):
    from dashboard.blueprints import slack_bot as sb
    from dashboard.services import purchase_invoice_assign as pinv
    monkeypatch.setenv('PURCHASE_INVOICE_ASSIGN_ENABLED', 'true')
    monkeypatch.setattr(sb, '_SETTLEMENT_CHECKER_IDS', (SB,))
    assert sb._pinv_deny(_body('UANY')) == ''                       # 처음 지정은 누구나
    parsed = {'date_md': '11/11', 'grn': 'G', 'amount': 533500, 'partner': '강원포장'}
    pinv.save_assignment('CINV', '1700000000.000100', parsed, 'G4049-JW', '기타비:운반비', '', MGR, '강성환')
    assert sb._pinv_deny(_body(MGR)) == ''                           # 처음 지정한 사람
    assert sb._pinv_deny(_body(SB)) == ''                            # 경영지원
    assert sb._pinv_deny(_body(KIKO)) == ''                          # 승인자
    assert '처음 지정한 사람' in sb._pinv_deny(_body('UANY'))
    view = {'private_metadata': json.dumps({'ch': 'CINV', 'ts': '1700000000.000100'})}
    assert '처음 지정한 사람' in sb._pinv_deny({'user': {'id': 'UANY'}}, view)   # 제출 때도


# ── 11-15 낮음 2건: 동시 제출 · 취소 기록 ──────────────────────────

def test_concurrent_creates_second_card_still_gets_dup_mark(tmp_db, monkeypatch):
    import threading
    import time
    orig = ea.find_recent_fund_dups

    def slow(*a, **k):          # 확인과 저장 사이를 벌려 두 제출이 겹치게
        r = orig(*a, **k)
        time.sleep(0.2)
        return r
    monkeypatch.setattr(ea, 'find_recent_fund_dups', slow)
    out = []
    ts = [threading.Thread(target=lambda: out.append(ea.create(ea.FORM_FUND, MGR, '강성환', dict(FIELDS))))
          for _ in range(2)]
    for t in ts:
        t.start()
    for t in ts:
        t.join()
    marks = sorted(bool(ea.extra(r).get('dup_of')) for r in out)
    assert marks == [False, True]          # 잠금 없으면 둘 다 False


def test_cancel_is_logged_in_invoice_post_thread(se):
    row = se.process_new_submit(MGR, dict(FIELDS, origin_channel='CINV', origin_ts='1700000000.000100'))
    assert se.apply_action('cancel', row['id'], MGR) is None
    logs = [kw for kw in se._inv.of('chat_postMessage')
            if kw.get('thread_ts') == '1700000000.000100' and '요청 취소' in kw.get('text', '')]
    assert len(logs) == 1 and ea.title(row) in logs[0]['text'] and logs[0]['channel'] == 'CINV'


def test_cancel_without_invoice_post_logs_nothing_there(se):
    row = se.process_new_submit(MGR, dict(FIELDS))
    assert se.apply_action('cancel', row['id'], MGR) is None
    assert not [kw for kw in se._inv.of('chat_postMessage') if '요청 취소' in kw.get('text', '')]


# ── 11-15 수정 뒤 점검: 나빠진 곳 ────────────────────────────────

def test_edit_to_correct_account_clears_payee_changed(se):
    ea.remember_payee(FIELDS['partner'], '김강원', '신한', '110-511-825643')
    f = dict(FIELDS, account='999-888-777666')
    assert ea.complete_new_request(f) == {}
    row = se.process_new_submit(MGR, f)
    assert ea.extra(row).get('payee_changed')
    g = dict(FIELDS, _edit=row['id'], extra=dict(ea.extra(row)))   # 수정 창은 예전 extra 를 들고 온다
    g['extra'].pop('payee_changed', None)
    g['extra'].pop('payee_saved', None)
    assert ea.complete_new_request(g) == {}
    new = se.process_new_submit(MGR, g)
    x = ea.extra(ea.get(row['id']))
    assert not x.get('payee_changed') and x.get('payee_saved')
    assert '예전에 보낸 계좌' not in ea.card_text(ea.get(row['id']), ea.steps(row['id']), 'G4049-JW')


def _age(aid, days):
    t = (datetime.now() - timedelta(days=days)).strftime('%Y-%m-%d %H:%M:%S')
    with ea._db()._get_connection() as conn:
        conn.execute('UPDATE approvals SET updated_at=?, created_at=? WHERE id=?', (t, t, aid))
        conn.commit()


def test_old_prepaid_warns_once_on_invoice_submit(se):
    p = se.process_new_submit(KIKO, dict(FIELDS))
    se.apply_action('transfer', p['id'], SB)
    _age(p['id'], ea.DUP_WINDOW_DAYS + 30)
    assert ea.find_unlinked_match('G', FIELDS['amount'], FIELDS['partner']) is None   # 목록은 안 가림
    assert ea.find_old_prepaid('G', FIELDS['amount'], FIELDS['partner'])['id'] == p['id']
    f_inv = dict(FIELDS, origin_channel='CINV', origin_ts='1700000000.000900')
    import dashboard.services.e_approval as _ea
    orig = _ea.read_new_request
    _ea.read_new_request = lambda view: dict(f_inv)
    try:
        view = {'id': 'V9', 'state': {'values': {'amount': {}}}}
        f, errs = se.validate_new_submit(view)
        assert set(errs) == {'amount'} and ea.title(p) in errs['amount'] and '선입금' in errs['amount']
        f, errs = se.validate_new_submit(view)             # 같은 창에서 한 번 더 → 통과
        assert errs == {}
        f, errs = se.validate_new_submit(dict(view, id='V10'))
        assert set(errs) == {'amount'}                      # 다른 창은 다시 물음
    finally:
        _ea.read_new_request = orig


def test_old_prepaid_window(tmp_db):
    a = ea.create(ea.FORM_FUND, MGR, '강성환', dict(FIELDS))
    with ea._db()._get_connection() as conn:
        conn.execute('UPDATE approvals SET status=? WHERE id=?', (ea.DONE, a['id']))
        conn.commit()
    _age(a['id'], 10)
    assert ea.find_old_prepaid('G', FIELDS['amount'], FIELDS['partner']) is None      # 90일 안은 기존 검사가
    _age(a['id'], 200)
    hit = ea.find_old_prepaid('G', FIELDS['amount'], '강원포장')
    assert hit['id'] == a['id'] and '선입금' in ea.old_prepaid_warning(hit)
    assert ea.find_old_prepaid('R', FIELDS['amount'], FIELDS['partner'])['id'] == a['id']   # 사업자를 잘못 골라도(5차)
    _age(a['id'], ea.OLD_PREPAID_DAYS + 5)
    assert ea.find_old_prepaid('G', FIELDS['amount'], FIELDS['partner']) is None      # 1년 넘음


# ── ⑥ 흔한 금액이 겹쳤을 때 — 공사 코드로 좁히기 (2026-11-17 #45 서울크레인) ──────

def test_common_amount_two_candidates_narrowed_by_code(se):
    """크레인 양중 33만 원처럼 흔한 금액이 90일 안에 두 번이면, 코드 없이는 짝을 못 찾았다."""
    a = se.process_new_submit(KIKO, dict(FIELDS, project_code='R4226-YM', grn='R', partner='서울크레인'))
    se.apply_action('transfer', a['id'], SB)                 # 먼저 건은 이체 완료
    b = se.process_new_submit(MGR, dict(FIELDS, project_code='R4256-MJ', grn='R', partner='서울크레인'))
    # 후보가 둘 — 코드를 안 주면 예전처럼 포기한다(제출 창 경고도, 자동 연결도 안 됨)
    assert ea.find_unlinked_match('R', FIELDS['amount'], '서울크레인') is None
    # 계산서 글에 적힌 코드로 좁히면 그 건
    assert ea.find_unlinked_match('R', FIELDS['amount'], '서울크레인', 'R4256-MJ')['id'] == b['id']
    assert ea.find_unlinked_match('R', FIELDS['amount'], '서울크레인(주)', 'r4226-ym')['id'] == a['id']
    # 어느 쪽도 아닌 코드면 좁혀지지 않는다 — 사람이 고른다
    assert ea.find_unlinked_match('R', FIELDS['amount'], '서울크레인', 'R9999-ZZ') is None


def test_single_candidate_ignores_code_mismatch(se):
    """후보가 하나면 코드가 달라도 그대로 — 코드를 잘못 적은 계산서 글도 짝이 지어져야 한다."""
    a = se.process_new_submit(MGR, dict(FIELDS, project_code='G4049-JW'))
    assert ea.find_unlinked_match('G', FIELDS['amount'], FIELDS['partner'], 'G9999-ZZ')['id'] == a['id']


# ── ⑦ 계산서 글을 안 고른 결재 — 카드 경고 줄 ─────────────────────────

def _card(row):
    return ea.card_text(ea.get(row['id']), ea.steps(row['id']), 'G4049-JW')


def test_card_warns_when_invoice_post_not_picked(se):
    a = se.process_new_submit(MGR, dict(FIELDS))              # 증빙=전자계산서, 계산서 글 안 고름
    assert '계산서 글을 고르지 않았어요' in _card(a)


def test_no_warning_for_supply_cost_without_site(se):
    """공통비 비품은 선입금 뒤 계산서를 받는 게 정상 — 경고가 흔해지면 무뎌진다."""
    a = se.process_new_submit(MGR, dict(FIELDS, project_code='공통비', amount=90000))
    assert '계산서 글을 고르지 않았어요' not in _card(a)


def test_no_warning_when_invoice_picked_or_receipt_or_attached(se):
    picked = se.process_new_submit(MGR, dict(FIELDS, origin_channel='CINV', origin_ts='1700000000.000700'))
    assert '계산서 글을 고르지 않았어요' not in _card(picked)
    receipt = se.process_new_submit(MGR, dict(FIELDS, amount=70000, evidence='기타영수증'))
    assert '계산서 글을 고르지 않았어요' not in _card(receipt)
    attached = se.process_new_submit(MGR, dict(FIELDS, amount=80000))
    ea.add_extra_files(attached['id'], [{'name': '계산서.pdf', 'url': 'https://x/1'}])
    assert '계산서 글을 고르지 않았어요' not in _card(attached)


def test_socket_touch_keeps_active_tab_connected(monkeypatch):
    from dashboard.utils import socket_session_guard as sg
    monkeypatch.setattr(sg, '_conns', {})
    start = datetime(2026, 11, 15, 9, 0, 0)
    sg.register('s1', 'a@x.com', start.isoformat(), 'sess-a')
    sg.register('s2', 'b@x.com', start.isoformat(), 'sess-b')
    later = start + timedelta(hours=3)
    assert sg.touch('a@x.com', later.isoformat(), 'sess-a') == 1
    assert sg.touch('a@x.com', start.isoformat(), 'sess-a') == 0          # 더 옛 시각으로 되돌리지 않음

    class _DB:
        def get_user_by_email(self, e):
            return {'is_active': True}
    cut = sg.sweep(lambda sid: None, now=start + timedelta(hours=5), db=_DB())
    assert cut == ['s2']                                         # 화면을 쓴 a 는 유지, 4시간 넘게 연장 없던 b 만 끊음


def test_socket_touch_and_logout_bound_to_session(monkeypatch):
    """5차 점검 #5: 같은 직원의 다른 탭·브라우저가 연장해도, 로그아웃한 세션의 연결은 살아남지 않는다."""
    from dashboard.utils import socket_session_guard as sg
    monkeypatch.setattr(sg, '_conns', {})
    monkeypatch.setattr(sg, '_disconnect', None)
    start = datetime(2026, 11, 15, 9, 0, 0)
    sg.register('old', 'a@x.com', start.isoformat(), 'sess-old')     # 로그아웃할 탭
    sg.register('new', 'a@x.com', start.isoformat(), 'sess-new')     # 계속 쓰는 탭
    later = start + timedelta(hours=3)
    assert sg.touch('a@x.com', later.isoformat(), 'sess-new') == 1   # 쓰는 세션 연결만 연장
    assert sg.tracked()['old']['login_time'] == start.isoformat()
    assert sg.touch('a@x.com', later.isoformat()) == 0               # 세션 ID 없으면 아무것도 연장 안 함
    assert sg.touch('b@x.com', later.isoformat(), 'sess-new') == 0   # 메일이 다르면 안 함

    class _DB:
        def get_user_by_email(self, e):
            return {'is_active': True}
    assert sg.sweep(lambda sid: None, now=start + timedelta(hours=5), db=_DB()) == ['old']

    # 로그아웃: 그 세션 연결을 바로 끊음(끊기 함수 있음)
    monkeypatch.setattr(sg, '_conns', {})
    sg.register('t1', 'a@x.com', start.isoformat(), 'sess-1')
    sg.register('t2', 'a@x.com', start.isoformat(), 'sess-2')
    cut = []
    monkeypatch.setattr(sg, '_disconnect', cut.append)
    assert sg.revoke_session('sess-1') == ['t1'] and cut == ['t1']
    assert set(sg.tracked()) == {'t2'}
    assert sg.revoke_session(None) == [] and sg.revoke_session('') == []

    # 끊기 함수가 아직 없으면 표시만 → 연장도 안 되고 다음 sweep 이 끊음
    monkeypatch.setattr(sg, '_disconnect', None)
    sg.register('t3', 'a@x.com', start.isoformat(), 'sess-3')
    assert sg.revoke_session('sess-3') == ['t3']
    assert sg.touch('a@x.com', later.isoformat(), 'sess-3') == 0
    assert sg.sweep(lambda sid: None, now=start + timedelta(minutes=10), db=_DB()) == ['t3']


def test_logout_and_ping_pass_own_session_id():
    """연결 등록·연장·로그아웃이 자기 세션 ID 를 넘기는지."""
    import inspect
    import dashboard
    from dashboard.blueprints import auth as auth_bp
    assert "revoke_session(session.get('session_id'))" in inspect.getsource(auth_bp.logout)
    assert "session.get('session_id')" in inspect.getsource(auth_bp.refresh_session)
    assert "_ss.get('session_id'))" in inspect.getsource(dashboard)


@pytest.mark.parametrize('raw, shown', [
    ('<@U0ABC123> 확인 부탁', '@직원 확인 부탁'),
    ('<@U0ABC123|sb> 확인', '@sb 확인'),
    ('견적 <https://drive.google.com/x?a=1&amp;b=2|드라이브> 참고', '견적 드라이브 참고'),
    ('<#C0123ABC|영업관리> 에 올림', '#영업관리 에 올림'),
    ('<!channel> 11/14 G 150,000원 강원포장', '11/14 G 150,000원 강원포장'),
])
def test_pin_summary_makes_slack_tokens_readable(raw, shown):
    from dashboard.services import pin_remind as pr
    import dashboard.services.hr as hr
    orig = hr.find_by_slack
    hr.find_by_slack = lambda uid: None
    try:
        line = pr._esc_from_slack(pr._summary_line(raw))
    finally:
        hr.find_by_slack = orig
    assert line == shown and '<' not in line and '&lt;' not in line


def test_login_never_forces_consent(monkeypatch):
    from dashboard import google_oauth as go
    seen = {}

    class _Flow:
        def authorization_url(self, **kw):
            seen.update(kw)
            return 'https://accounts.google.com/x', 'st'
    m = go.GoogleOAuthManager.__new__(go.GoogleOAuthManager) if hasattr(go, 'GoogleOAuthManager') else None
    if m is None:
        pytest.skip('클래스 이름이 다름')
    m.allowed_domain = 'itg-aircon.com'
    monkeypatch.setattr(m, 'create_flow', lambda redirect_uri=None: _Flow(), raising=False)
    monkeypatch.setattr(m, '_has_valid_refresh_token', lambda: False, raising=False)
    m.get_authorization_url()
    assert seen['prompt'] == 'select_account'
