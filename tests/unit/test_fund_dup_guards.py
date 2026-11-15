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
    assert '✅ 확인된 계좌(강남)' in card and '⚠️' not in card
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


def test_socket_touch_keeps_active_tab_connected(monkeypatch):
    from dashboard.utils import socket_session_guard as sg
    monkeypatch.setattr(sg, '_conns', {})
    start = datetime(2026, 11, 15, 9, 0, 0)
    sg.register('s1', 'a@x.com', start.isoformat())
    sg.register('s2', 'b@x.com', start.isoformat())
    later = start + timedelta(hours=3)
    assert sg.touch('a@x.com', later.isoformat()) == 1
    assert sg.touch('a@x.com', start.isoformat()) == 0          # 더 옛 시각으로 되돌리지 않음

    class _DB:
        def get_user_by_email(self, e):
            return {'is_active': True}
    cut = sg.sweep(lambda sid: None, now=start + timedelta(hours=5), db=_DB())
    assert cut == ['s2']                                         # 화면을 쓴 a 는 유지, 4시간 넘게 연장 없던 b 만 끊음


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
