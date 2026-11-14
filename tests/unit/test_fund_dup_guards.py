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


def test_first_time_partner_has_no_change_flag(tmp_db):
    f = dict(FIELDS, partner='처음보는상사')
    ea.mark_payee_saved(f)
    assert 'payee_changed' not in f['extra']


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
