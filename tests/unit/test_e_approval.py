"""전자결재 엔진 (2026-10-30) — 장부·상태 전환·권한·계좌 기억·카드, 슬랙 흐름(가짜 클라이언트)."""
import json

import pytest

from dashboard.services import e_approval as ea

KIKO, SB, JW, YG, MGR = 'U04UL2ZLJAX', 'U0BHC2JV7U5', 'U051ENK8AAY', 'U050XNG5VBR', 'UMGR'


@pytest.fixture
def tmp_db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    db = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(ea, '_db', lambda: db)
    monkeypatch.setattr(ea, '_table_ready', False)
    monkeypatch.setattr(ea, '_suffix_map', lambda: {'강성환': 'SH', '황샛별': 'SB', '박정우': 'JW', '박용구': 'YG'})
    for k in ('EAPPROVAL_APPROVER_IDS', 'EAPPROVAL_EXECUTOR_DM_IDS',
              'SLACK_SETTLEMENT_CHECKER_ID', 'SLACK_SETTLEMENT_CHECKER_EMAIL'):
        monkeypatch.delenv(k, raising=False)
    return db


FIELDS = {'project_code': 'G4049-JW', 'cost_category': '기타비:운반비', 'content': 'FCU 부품',
          'evidence': '전자계산서', 'grn': 'G', 'partner': '(주)강원포장', 'payee_name': '김강원',
          'bank': '신한', 'account': '110-511-825643', 'amount': 533500, 'urgency': '일반', 'memo': ''}


# ── 역할 ────────────────────────────────────────────────────────────
def test_roles_default(tmp_db):
    assert ea.approver_ids() == (KIKO,)
    assert ea.executor_ids() == (SB, KIKO, JW, YG)
    assert ea.executor_dm_ids() == (SB,)


def test_roles_env_override(tmp_db, monkeypatch):
    monkeypatch.setenv('EAPPROVAL_EXECUTOR_DM_IDS', f'{KIKO},{JW}')
    assert ea.executor_dm_ids() == (KIKO, JW)


# ── 생성·전환 ───────────────────────────────────────────────────────
def test_create_by_manager_waits_for_approval(tmp_db):
    row = ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    assert row['status'] == ea.PENDING and row['amount'] == 533500
    assert [s['result'] for s in ea.steps(row['id'])] == ['requested']


def test_create_by_approver_skips_approval(tmp_db):
    row = ea.create(ea.FORM_FUND, KIKO, '고광일', FIELDS)
    assert row['status'] == ea.APPROVED
    assert [s['result'] for s in ea.steps(row['id'])] == ['requested', 'auto']
    assert '요청자가 승인자' in ea.step_line(row, ea.steps(row['id']))


def test_full_flow_and_double_click_guard(tmp_db):
    row = ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    assert ea.transition(row['id'], 'transfer', SB, '황샛별') is None          # 승인 전 이체 불가
    a = ea.transition(row['id'], 'approve', KIKO, '고광일')
    assert a['status'] == ea.APPROVED
    assert ea.transition(row['id'], 'approve', KIKO, '고광일') is None         # 두 번 눌러도 한 번
    d = ea.transition(row['id'], 'transfer', SB, '황샛별')
    assert d['status'] == ea.DONE
    assert ea.transition(row['id'], 'transfer', JW, '박정우') is None          # 다른 사람이 또 눌러도 한 번
    assert ea.transition(row['id'], 'cancel', MGR, '박정우') is None           # 끝난 건 취소 불가
    line = ea.step_line(d, ea.steps(row['id']))
    assert 'KiKO 승인' in line and '경영지원 (SB) 이체' in line and '고광일' not in line


def test_reject_and_cancel(tmp_db):
    r1 = ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    rej = ea.transition(r1['id'], 'reject', KIKO, '고광일', '금액 확인 필요')
    assert rej['status'] == ea.REJECTED
    assert '반려 사유: 금액 확인 필요' in ea.card_text(rej, ea.steps(r1['id']), 'G4049-JW')
    r2 = ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    ea.transition(r2['id'], 'approve', KIKO, '고광일')
    assert ea.transition(r2['id'], 'cancel', MGR, '박정우')['status'] == ea.CANCELLED


def test_permissions(tmp_db):
    row = ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    assert ea.can_act('approve', row, KIKO) and not ea.can_act('approve', row, SB)
    assert not ea.can_act('reject', row, SB)                     # 승인 대기 중 반려는 승인자만
    assert ea.can_act('cancel', row, MGR) and ea.can_act('cancel', row, KIKO)
    assert not ea.can_act('cancel', row, JW)
    row = ea.transition(row['id'], 'approve', KIKO, '고광일')
    for u in (SB, KIKO, JW, YG):
        assert ea.can_act('transfer', row, u)
    assert not ea.can_act('transfer', row, MGR)
    assert ea.can_act('reject', row, SB)                         # 이체 대기 중엔 이체 담당도 반려 가능


# ── 계산서 글 연결 ──────────────────────────────────────────────────
def test_find_unlinked_match_and_link(tmp_db):
    row = ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    m = ea.find_unlinked_match('G', 533500, '강원포장 ')          # (주)·공백 무시
    assert m and m['id'] == row['id']
    assert ea.find_unlinked_match('R', 533500, '강원포장') is None
    assert ea.find_unlinked_match('G', 533000, '강원포장') is None
    ea.link_origin(row['id'], 'CINV', '111.1')
    assert ea.find_unlinked_match('G', 533500, '강원포장') is None
    assert ea.find_by_origin('CINV', '111.1')['id'] == row['id']


def test_find_unlinked_match_ambiguous_returns_none(tmp_db):
    ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    ea.create(ea.FORM_FUND, MGR, '강성환', FIELDS)
    assert ea.find_unlinked_match('G', 533500, '강원포장') is None


def test_rejected_not_matched(tmp_db):
    row = ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    ea.transition(row['id'], 'reject', KIKO, '고광일', 'x')
    assert ea.find_unlinked_match('G', 533500, '강원포장') is None


# ── 계좌 기억·입력 ──────────────────────────────────────────────────
def test_payee_memory(tmp_db):
    assert ea.lookup_payee('강원포장') is None
    ea.remember_payee('㈜강원포장', '김강원', '신한', '110-1')
    assert ea.lookup_payee('주식회사 강원포장')['account'] == '110-1'
    ea.remember_payee('강원포장', '김강원', '국민', '220-2')
    assert ea.lookup_payee('강원포장')['bank'] == '국민'


def test_parse_and_validate():
    assert ea.parse_amount('533,500원') == 533500 and ea.parse_amount('') == 0
    assert ea.normalize_account(' 110-511 825643 ') == '110-511825643'
    assert ea.grn_from_code('R4120-YM') == 'R' and ea.grn_from_code('프로젝트0') == ''
    assert ea.validate_transfer('윤정민', '신한', '110-511-825643') == {}
    assert set(ea.validate_transfer('', '', '12')) == {'payee_name', 'bank', 'account'}


def _new_view(**over):
    vals = {
        'project': {ea.A_PROJECT: {'selected_option': {'value': over.get('project', 'G4049-JW')}}},
        'evidence': {'evidence': {'selected_option': {'value': '수기계산서'}}},
        'cost': {'cost': {'selected_option': {'value': '기타비:운반비'}}},
        'content': {'content': {'value': '운반'}},
        'partner': {ea.A_PARTNER: {'selected_option': {'value': over.get('partner', '강원포장')}}},
        'amount': {'amount': {'value': over.get('amount', '150,000')}},
        'grn': {'grn': {'selected_option': ({'value': over.get('grn', 'G')} if over.get('grn', 'G') else None)}},
        'payee_name': {'payee_name': {'value': over.get('payee_name', '')}},
        'bank': {'bank': {'value': over.get('bank', '')}},
        'account': {'account': {'value': over.get('account', '')}},
        'urgency': {'urgency': {'selected_option': {'value': '긴급'}}},
        'memo': {'memo': {'value': None}},
    }
    return {'private_metadata': json.dumps({'project': over.get('fixed', '')}), 'state': {'values': vals}}


def test_new_request_uses_remembered_account(tmp_db):
    ea.remember_payee('강원포장', '김강원', '신한', '110-511-825643')
    f = ea.read_new_request(_new_view())
    assert ea.complete_new_request(f) == {}
    assert f['grn'] == 'G' and f['amount'] == 150000 and f['account'] == '110-511-825643'
    assert f['urgency'] == '긴급' and f['evidence'] == '수기계산서'


def test_new_request_errors(tmp_db):
    f = ea.read_new_request(_new_view(partner='처음거래처', amount='abc'))
    errs = ea.complete_new_request(f)
    assert set(errs) == {'amount', 'account'}
    f = ea.read_new_request(_new_view(project='프로젝트0', grn='', payee_name='a', bank='b', account='1234-5678'))
    assert set(ea.complete_new_request(f)) == {'grn'}
    f = ea.read_new_request(_new_view(grn='', payee_name='a', bank='b', account='1234-5678'))
    assert set(ea.complete_new_request(f)) == {'grn'}   # G 공사여도 사업자 구분은 직접 골라야(11-01)
    f = ea.read_new_request(_new_view(project='프로젝트0', grn='N', payee_name='a', bank='b', account='1234-5678'))
    assert ea.complete_new_request(f) == {}


def test_fixed_project_from_metadata(tmp_db):
    grn = next(b for b in ea.new_request_view({}, 'R4120-YM')['blocks'] if b.get('block_id') == 'grn')
    assert grn['element']['initial_option']['value'] == 'R' and not grn.get('optional')
    v = _new_view(fixed='R4120-YM', grn='R', payee_name='a', bank='b', account='1234-5678')
    del v['state']['values']['project']
    f = ea.read_new_request(v)
    assert f['project_code'] == 'R4120-YM' and ea.complete_new_request(f) == {} and f['grn'] == 'R'


def test_partner_options_known_and_new(tmp_db):
    ea.remember_payee('(주)강원포장', '김강원', '신한', '110-511-825643')
    def vals(q):
        r = ea.partner_options(q)
        opts = r.get('options') or [o for g in r.get('option_groups', []) for o in g['options']]
        return [o['value'] for o in opts]
    assert vals('') == ['(주)강원포장']
    assert vals('강원') == ['(주)강원포장', 'new:강원']           # (주) 빼고 비교, 아는 곳 먼저
    assert vals('강원포장') == ['(주)강원포장']                   # 같은 곳이면 '새 거래처' 안 띄움
    assert vals('대성크레인') == ['new:대성크레인']               # 처음 보는 곳은 입력 그대로
    assert ea.partner_from_value('new:대성크레인') == '대성크레인'


def test_prefilled_account_shows_note_and_new_block_ids(tmp_db):
    """거래처를 고르면 예전 계좌를 창에 채우고 '한 번 더 확인' 안내(11-01)."""
    payee = {'payee_name': '김강원', 'bank': '신한', 'account': '110-511-825643'}
    v = ea.new_request_view({'partner': '강원포장', 'payee': payee, 'pv': 2})
    txt = json.dumps(v['blocks'], ensure_ascii=False)
    assert ea.PAYEE_SAVED_NOTE in txt
    acc = next(b for b in v['blocks'] if b.get('block_id') == 'account_2')
    assert acc['element']['initial_value'] == '110-511-825643'
    v2 = ea.new_request_view({'partner': '대성크레인', 'payee': None, 'pv': 3})
    assert ea.PAYEE_SAVED_NOTE not in json.dumps(v2['blocks'], ensure_ascii=False)
    assert '처음 보내는 거래처' in json.dumps(v2['blocks'], ensure_ascii=False)
    # 접미사 붙은 칸도 읽고, 기억한 계좌 그대로면 카드에 표시
    ea.remember_payee('강원포장', '김강원', '신한', '110-511-825643')
    view = _new_view()
    st = view['state']['values']
    for k in ('payee_name', 'bank', 'account'):
        st[f'{k}_2'] = {k: {'value': payee[k]}}
        del st[k]
    f = ea.read_new_request(view)
    assert f['account'] == '110-511-825643' and ea.complete_new_request(f) == {}
    row = ea.create(ea.FORM_FUND, MGR, '강성환', f)
    assert '예전 계좌 자동 채움' in ea.card_text(row, ea.steps(row['id']), 'G4049-JW')


def test_new_request_view_fixed_hides_project_select():
    view = ea.new_request_view({'project': 'G4049-JW'}, 'G4049-JW', 'G4049-JW')
    assert 'project' not in {b.get('block_id') for b in view['blocks']}
    view = ea.new_request_view({})
    assert 'project' in {b.get('block_id') for b in view['blocks']}


# ── 카드 ────────────────────────────────────────────────────────────
def test_buttons_by_status_and_role(tmp_db):
    row = ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    ids = lambda r, role='all': [b['action_id'] for b in ea.buttons(r, role)]  # noqa: E731
    assert ids(row) == [ea.A_APPROVE, ea.A_REJECT, ea.A_CANCEL]
    assert ids(row, 'approver') == [ea.A_APPROVE, ea.A_REJECT]
    assert ids(row, 'requester') == [ea.A_CANCEL]
    assert ids(row, 'executor') == []
    row = ea.transition(row['id'], 'approve', KIKO, '고광일')
    assert ids(row) == [ea.A_TRANSFER, ea.A_CANCEL]
    assert ids(row, 'executor') == [ea.A_TRANSFER]
    row = ea.transition(row['id'], 'transfer', SB, '황샛별')
    assert ids(row) == []
    done = ea.card_blocks(row, ea.steps(row['id']), 'G4049-JW')
    assert [b['type'] for b in done] == ['section', 'context']  # 끝나면 버튼·안내 없이 아래 여백만
    assert done[0]['text']['text'].startswith('⠀\n')


def test_card_blocks_spacing(tmp_db):
    row = ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    types = [b['type'] for b in ea.card_blocks(row, ea.steps(row['id']), 'G4049-JW')]
    # 위 여백(본문 첫 줄) · 본문 · 여백 · 버튼 · 안내 · 아래 여백
    assert types == ['section', 'context', 'actions', 'context', 'context']


def test_card_text(tmp_db):
    row = ea.create(ea.FORM_FUND, MGR, '박정우', dict(FIELDS, urgency='긴급'))
    t = ea.card_text(row, ea.steps(row['id']), 'G4049-JW · 광명 양지로 21', 6500000)
    assert f'자금집행 #{row["id"]}' in t and '승인 대기' in t and '긴급' in t
    assert '금액: *533,500원*' in t and '고객 견적(VAT 별도): 6,500,000원' in t
    assert '신한 110-511-825643 (김강원)' in t and '기타비 · 운반비 — FCU 부품' in t


# ── 슬랙 흐름 (가짜 클라이언트) ─────────────────────────────────────
class _Fake:
    def __init__(self):
        self.calls = []
        self.n = 0

    def __getattr__(self, name):
        def _call(**kw):
            self.calls.append((name, kw))
            if name == 'conversations_open':
                return {'channel': {'id': 'D-' + kw['users']}}
            if name == 'chat_postMessage':
                self.n += 1
                return {'ts': f'{self.n}.0'}
            if name == 'chat_getPermalink':
                return {'permalink': 'https://x/p'}
            return {}
        return _call

    def of(self, name):
        return [kw for n, kw in self.calls if n == name]


@pytest.fixture
def se(tmp_db, monkeypatch):
    from dashboard.blueprints import slack_eapproval as se
    fake, inv = _Fake(), _Fake()
    monkeypatch.setenv('SLACK_EAPPROVAL_CHANNEL', 'CEAP')
    monkeypatch.setenv('SLACK_EAPPROVAL_BOT_TOKEN', 'xoxb-test')
    monkeypatch.setattr(se, 'eap_client', lambda: fake)
    monkeypatch.setattr(se, '_invoice_client', lambda: inv)
    names = {KIKO: '고광일', SB: '황샛별', JW: '박정우', MGR: '강성환'}
    monkeypatch.setattr(se, '_name', lambda uid: names.get(uid, uid))
    monkeypatch.setattr(se, '_project_rec', lambda code: {'현장 주소': '광명 양지로 21', '총액 1': '6,500,000'}
                        if code == 'G4049-JW' else None)
    se._fake, se._inv = fake, inv
    return se


def test_start_posts_card_pins_and_dms(se):
    row = se.process_new_submit(MGR, dict(FIELDS))
    f = se._fake
    card = f.of('chat_postMessage')[0]
    assert card['channel'] == 'CEAP' and '광명 양지로 21' in card['blocks'][0]['text']['text']
    assert f.of('pins_add')[0]['channel'] == 'CEAP'
    refs = ea.dm_refs(row)
    assert {(r['user'], r['role']) for r in refs} == {(KIKO, 'approver')}       # 신청자 시작 DM 없음
    assert not [kw for kw in f.of('chat_postMessage') if kw['channel'] == 'D-' + MGR]
    assert row['card_ts'] == '1.0'


def test_approve_then_transfer_flow(se):
    row = se.process_new_submit(MGR, dict(FIELDS))
    assert se.apply_action('transfer', row['id'], SB).startswith('이미 처리된 결재')  # 승인 전 이체 불가
    assert se.apply_action('approve', row['id'], SB) == '승인은 승인자만 할 수 있어요.'
    assert se.apply_action('approve', row['id'], KIKO) is None
    row = ea.get(row['id'])
    assert row['status'] == ea.APPROVED
    assert (SB, 'executor') in {(r['user'], r['role']) for r in ea.dm_refs(row)}   # 이체 담당 DM
    assert se.apply_action('approve', row['id'], KIKO).startswith('이미 처리된 결재')
    assert se.apply_action('transfer', row['id'], MGR) == '이체 완료는 경영지원 담당만 누를 수 있어요.'
    assert se.apply_action('transfer', row['id'], JW) is None                       # 대리 이체
    row = ea.get(row['id'])
    assert row['status'] == ea.DONE
    f = se._fake
    assert any(kw.get('reply_broadcast') for kw in f.of('chat_postMessage'))        # 이체 완료는 채널에도
    assert f.of('pins_remove')[-1]['timestamp'] == row['card_ts']
    updated = {kw['ts'] for kw in f.of('chat_update')}
    assert row['card_ts'] in updated and all(r['ts'] in updated for r in ea.dm_refs(row))


def test_reject_notifies_requester_and_origin(se):
    from dashboard.services import purchase_invoice_assign as pinv  # noqa: F401
    row = se.start_from_invoice(MGR, '강성환', 'CINV', '111.1',
                                {'date_md': '10/29', 'grn': 'G', 'amount': 533500, 'partner': '강원포장'},
                                'G4049-JW', '기타비:운반비', '', '전자계산서', '김강원', '신한', '110-511-825643', '일반')
    assert row['origin_ts'] == '111.1' and ea.lookup_payee('강원포장')['bank'] == '신한'
    assert se.apply_action('reject', row['id'], KIKO, '견적 다시') is None
    assert ea.get(row['id'])['status'] == ea.REJECTED
    dms = [kw for kw in se._fake.of('chat_postMessage') if kw['channel'] == 'D-' + MGR]
    assert any('견적 다시' in kw.get('text', '') for kw in dms)
    assert any('반려' in kw['text'] and kw['thread_ts'] == '111.1' for kw in se._inv.of('chat_postMessage'))


def test_approver_request_goes_straight_to_executor(se):
    row = se.process_new_submit(KIKO, dict(FIELDS))
    assert row['status'] == ea.APPROVED
    assert {(r['user'], r['role']) for r in ea.dm_refs(row)} == {(SB, 'executor')}


def test_card_post_failure_cancels(se, monkeypatch):
    def boom(**kw):
        raise RuntimeError('channel_not_found')
    monkeypatch.setattr(se._fake, 'chat_postMessage', boom, raising=False)
    with pytest.raises(RuntimeError):
        se.process_new_submit(MGR, dict(FIELDS))
    rows = [ea.get(i) for i in range(1, 3) if ea.get(i)]
    assert rows and rows[0]['status'] == ea.CANCELLED


def test_handle_new_submit_validation_errors(se):
    acks = []
    se.handle_new_submit(lambda **kw: acks.append(kw), {'user': {'id': MGR}},
                         _new_view(partner='처음거래처'))
    assert acks[0]['response_action'] == 'errors' and 'account' in acks[0]['errors']
    view = _new_view(partner='처음거래처')                      # 계좌 칸이 account_1 로 바뀐 화면
    st = view['state']['values']
    st['account_1'] = st.pop('account')
    acks.clear()
    se.handle_new_submit(lambda **kw: acks.append(kw), {'user': {'id': MGR}}, view)
    assert 'account_1' in acks[0]['errors'] and 'account' not in acks[0]['errors']


# ── 비품 신청 (11-01) ───────────────────────────────────────────────
def _supply_view(**over):
    vals = {'item': 'A4 용지 박스', 'link': 'https://www.coupang.com/vp/products/1', 'qty': '2',
            'unit_price': '25,000', 'amount': '', 'reason': '용지 소진', 'memo': ''}
    grn = over.pop('grn', 'G')
    vals.update(over)
    state = {k: {k: {'value': v}} for k, v in vals.items()}
    state['grn'] = {'grn': {'selected_option': {'value': grn} if grn else None}}
    return {'state': {'values': state}, 'private_metadata': '{}'}


def test_supply_request_fills_amount_and_stock_project(tmp_db):
    f = ea.read_supply_request(_supply_view())
    assert ea.complete_supply_request(f) == {}
    assert f['amount'] == 50000 and f['project_code'] == '프로젝트0'
    assert f['extra'] == {'link': 'https://www.coupang.com/vp/products/1', 'qty': 2, 'unit_price': 25000,
                          'reason': '용지 소진'}
    f2 = ea.read_supply_request(_supply_view(amount='53,000'))      # 배송비 포함 직접 입력
    ea.complete_supply_request(f2)
    assert f2['amount'] == 53000


def test_supply_request_errors(tmp_db):
    f = ea.read_supply_request(_supply_view(item='', link='쿠팡', qty='0', unit_price='', reason='', grn=''))
    assert set(ea.complete_supply_request(f)) == {'grn', 'item', 'link', 'qty', 'unit_price', 'reason'}


def test_supply_card_and_buttons(tmp_db):
    f = ea.read_supply_request(_supply_view())
    ea.complete_supply_request(f)
    row = ea.create(ea.FORM_SUPPLY, KIKO, '고광일', f)
    assert ea.extra(row)['qty'] == 2 and ea.status_label(row) == ('🔵', '구매 대기')
    t = ea.card_text(row, ea.steps(row['id']), '프로젝트0 · 회사 재고')
    assert f'비품 신청 #{row["id"]}' in t and '품명: A4 용지 박스 × 2개 (단가 25,000원)' in t
    assert '상품 보기' in t and '사유: 용지 소진' in t and '이체 계좌' not in t and '⏳ 구매' in t
    assert '사업자 구분: G (글로벌)' in t and row['grn'] == 'G'
    btn = ea.buttons(row, 'executor')[0]
    assert btn['text']['text'] == '🛒 구매 완료' and btn['action_id'] == ea.A_TRANSFER
    assert '구매하고 카드에서 [구매 완료]' in ea.dm_intro(row, 'executor')
    row = ea.create(ea.FORM_SUPPLY, KIKO, '고광일', {**f, 'memo': '배송 메시지에 직접 수령'})
    t = ea.card_text(row, ea.steps(row['id']), '프로젝트0 · 회사 재고')
    assert '구매 요청사항: 배송 메시지에 직접 수령' in t and '비고' not in t
    blocks = {b['block_id']: b for b in ea.supply_request_view({})['blocks'] if b.get('block_id')}
    assert not blocks['reason']['element']['multiline'] and blocks['memo']['label']['text'] == '구매 요청사항'


def test_extra_column_added_to_existing_table(tmp_db):
    with tmp_db._get_connection() as conn:          # 10-30 배포 때 만든(extra 없는) 표
        conn.execute('CREATE TABLE approvals (id INTEGER PRIMARY KEY AUTOINCREMENT, form TEXT NOT NULL, '
                     'status TEXT NOT NULL, requester_id TEXT, requester_name TEXT, project_code TEXT, '
                     'cost_category TEXT, content TEXT, evidence TEXT, grn TEXT, partner TEXT, payee_name TEXT, '
                     'bank TEXT, account TEXT, amount INTEGER NOT NULL DEFAULT 0, urgency TEXT, memo TEXT, '
                     'origin_channel TEXT, origin_ts TEXT, card_channel TEXT, card_ts TEXT, '
                     "dm_refs TEXT NOT NULL DEFAULT '[]', created_at TIMESTAMP, updated_at TIMESTAMP)")
        conn.commit()
    row = ea.create(ea.FORM_FUND, MGR, '박정우', FIELDS)
    assert ea.extra(row) == {}


def test_supply_flow_purchase_done(se):
    f = ea.read_supply_request(_supply_view())
    ea.complete_supply_request(f)
    row = se.process_supply_submit(MGR, f)
    assert row['status'] == ea.PENDING and row['project_code'] == '프로젝트0'
    assert se.apply_action('approve', row['id'], KIKO) is None
    assert se.apply_action('transfer', row['id'], MGR) == '구매 완료는 경영지원 담당만 누를 수 있어요.'
    assert se.apply_action('transfer', row['id'], SB) is None
    assert ea.get(row['id'])['status'] == ea.DONE
    msgs = se._fake.of('chat_postMessage')
    assert any('구매 완료' in kw.get('text', '') and kw.get('reply_broadcast') for kw in msgs)
    assert any('구매가 끝났어요' in kw.get('text', '') and kw['channel'] == 'D-' + MGR for kw in msgs)
    assert not se._inv.of('chat_postMessage')                     # 계산서 글 없음


def test_handle_supply_submit_validation(se):
    acks = []
    se.handle_supply_submit(lambda **kw: acks.append(kw), {'user': {'id': MGR}}, _supply_view(link='없음'))
    assert acks[0]['response_action'] == 'errors' and 'link' in acks[0]['errors']


def test_dm_is_notice_with_card_link_no_buttons(se):
    """DM 은 알림+카드 링크만, 버튼은 채널 카드에만(11-01)."""
    row = se.process_new_submit(MGR, dict(FIELDS))
    f = se._fake
    card, dms = f.of('chat_postMessage')[0], f.of('chat_postMessage')[1:]
    assert any(b['type'] == 'actions' for b in card['blocks'])                     # 카드엔 승인·반려
    appr = next(kw for kw in dms if kw['channel'] == 'D-' + KIKO)
    assert not any(b['type'] == 'actions' for b in appr['blocks'])
    txt = json.dumps(appr['blocks'], ensure_ascii=False)
    assert '승인을 요청했어요' in txt and 'https://x/p' in txt and '533,500원' in txt
    se.apply_action('approve', row['id'], KIKO)
    upd = [kw for kw in f.of('chat_update') if kw['channel'] == 'D-' + KIKO][-1]   # 처리 뒤 상태로 바뀜
    assert '이체 대기' in json.dumps(upd['blocks'], ensure_ascii=False)


def test_shortcuts_registered_and_match_manifest():
    """입력창 [/] 바로가기 — 핸들러 등록과 manifest callback_id 가 같아야 슬랙에서 열린다."""
    import os
    import yaml
    from dashboard.blueprints import slack_eapproval as se

    class _App:
        def __init__(self):
            self.shortcuts = []

        def shortcut(self, cid):
            self.shortcuts.append(cid)
            return lambda f: f

        def __getattr__(self, name):          # action/view/command/options — 데코레이터만 흉내
            return lambda *a, **k: (lambda f: f)

    app = _App()
    se.register_eapproval_handlers(app)
    assert set(app.shortcuts) == {ea.SC_HUB}
    path = os.path.join(os.path.dirname(__file__), '..', '..', 'docs', 'slack_eapproval_manifest.yml')
    with open(path, encoding='utf-8') as fh:
        man = yaml.safe_load(fh)
    assert {s['callback_id'] for s in man['features']['shortcuts']} == {ea.SC_HUB}


def test_hub_view_is_menu_without_submit():
    v = ea.hub_view()
    assert v['blocks'] == ea.menu_blocks()
    assert 'submit' not in v                      # 고르면 이 창이 그 입력 창으로 바뀐다


class _ViewClient:
    def __init__(self):
        self.calls = []

    def views_open(self, **kw):
        self.calls.append(('open', None))
        return {'view': {'id': 'V-new'}}

    def views_update(self, **kw):
        self.calls.append(('update', kw['view_id']))


def test_hub_buttons_update_same_view(monkeypatch):
    """바로가기 창 안에서 누르면 새 창을 쌓지 않고 그 창을 바꾼다. 채널 메뉴 메시지면 새 창."""
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se.threading, 'Thread', lambda target, daemon=True, args=(): type('T', (), {'start': lambda s: None})())
    in_view = {'container': {'type': 'view'}, 'view': {'id': 'V-hub'}}
    in_msg = {'container': {'type': 'message'}, 'channel': {'id': 'C1'}}
    assert se._hub_view_id(in_view) == 'V-hub' and se._hub_view_id(in_msg) == ''

    c = _ViewClient()
    se.open_new_modal(c, 'T', 'U1', origin='shortcut', view_id='V-hub')
    se.open_form_modal(c, 'T', 'U1', ea.forms.menu_options()[0]['value'], view_id='V-hub')
    assert c.calls == [('update', 'V-hub'), ('update', 'V-hub')]

    c = _ViewClient()
    se.open_new_modal(c, 'T', 'U1', origin='menu')
    assert c.calls == [('open', None)]


def test_my_requests_view_splits_long_text():
    text = '\n'.join(f'• 줄 {i} ' + '가' * 100 for i in range(80))
    v = ea.my_requests_view(text)
    secs = [b for b in v['blocks'] if b['type'] == 'section']
    assert len(secs) > 1 and all(len(b['text']['text']) <= 3000 for b in secs)
    assert v['blocks'][-1]['elements'][0]['action_id'] == ea.A_HUB_BACK


def test_menu_has_both_buttons():
    els = ea.menu_blocks()[1]['elements']
    assert [e['action_id'] for e in els] == [ea.A_NEW, ea.A_NEW_SUPPLY, ea.A_FORM_PICK, ea.A_MY_REQUESTS]
    assert {o['value'] for o in els[2]['options']} == set(ea.forms.FORMS)
    text = ea.menu_blocks()[0]['text']['text']
    assert '회사 재고 자재 포함) → 💸 자금집행' in text and '사무용 비품 → 🛒 비품 신청' in text


def test_my_requests_open_and_recent_closed(se):
    """[📂 내 요청] — 내 것만, 진행 중 + 최근 끝난 것, 누른 사람에게만(11-03)."""
    a = se.process_new_submit(MGR, dict(FIELDS))
    b = se.process_new_submit(MGR, dict(FIELDS))
    se.process_new_submit(JW, dict(FIELDS))                       # 남의 요청은 안 보임
    se.apply_action('reject', b['id'], KIKO, '금액 확인')
    opened, closed = ea.my_requests(MGR)
    assert [r['id'] for r in opened] == [a['id']] and [r['id'] for r in closed] == [b['id']]
    se.post_my_requests(se._fake, {'user': {'id': MGR}, 'channel': {'id': 'CEAP'}}, MGR)
    eph = se._fake.of('chat_postEphemeral')[-1]
    assert eph['user'] == MGR and eph['channel'] == 'CEAP'
    t = eph['text']
    assert f'자금집행 #{a["id"]}' in t and '승인 대기' in t and '진행 중 (1건)' in t
    assert f'자금집행 #{b["id"]}' in t and '반려' in t and '<https://x/p|카드 보기>' in t
    assert '진행 중인 요청이 없어요.' in ea.my_requests_text([], [], {})


# ── 나머지 양식 (11-01) ─────────────────────────────────────────────
from dashboard.services import e_approval_forms as forms  # noqa: E402


def _form_state(form_key, vals, meta=None, pv=1):
    """form_view 가 만든 block_id 그대로 state 를 꾸민다."""
    view = forms.form_view(form_key, {'pv': pv, **(meta or {})})
    st = {}
    for b in view['blocks']:
        if b.get('type') != 'input':
            continue
        el, bid = b['element'], b['block_id']
        keys = [f.key for f in forms.FORMS[form_key].fields]
        key = bid if bid in keys else bid.rsplit('_', 1)[0]      # 'total_2' → 'total', 'deposit_date' 그대로
        v = vals.get(key)
        if el['type'] == 'external_select':
            st[bid] = {el['action_id']: {'selected_option': {'value': v} if v else None}}
        elif el['type'] == 'datepicker':
            st[bid] = {el['action_id']: {'selected_date': v}}
        elif el['type'] == 'static_select':
            st[bid] = {el['action_id']: {'selected_option': {'value': v} if v else None}}
        elif el['type'] == 'multi_static_select':
            st[bid] = {el['action_id']: {'selected_options': [{'value': x} for x in (v or [])]}}
        else:
            st[bid] = {el['action_id']: {'value': v}}
    return {'private_metadata': view['private_metadata'], 'state': {'values': st}}


CONTRACT = {'project': 'G4237-MJ', 'total': '11440000', 'receiver': '윤명림', 'receiver_phone': '010-4312-1990',
            'send': '대면', 'contact': '010-4312-1990', 'deposit': '5720000', 'deposit_date': '2026-10-30'}


def test_contract_form_prefill_read_validate_and_task_flow(tmp_db):
    rec = {'총액 2': '11440000.0', '발주처 담당자': '윤명림', '발주처 연락처': '010-4312-1990', '발주처 이메일': '-'}
    pre = forms.prefill(forms.get(forms.FORM_CONTRACT), rec)
    assert pre == {'total': '11440000', 'receiver': '윤명림', 'receiver_phone': '010-4312-1990'}   # '-' 는 비움
    key, vals = forms.read_form(_form_state(forms.FORM_CONTRACT, CONTRACT))
    assert key == forms.FORM_CONTRACT and vals['total'] == '11440000' and vals['deposit_date'] == '2026-10-30'
    assert forms.validate(key, vals) == {}
    assert set(forms.validate(key, {**vals, 'send': '', 'receiver': ''})) == {'send', 'receiver'}
    row = ea.create(key, MGR, '강성환', {'project_code': 'G4237-MJ', 'amount': forms.amount_of(key, vals),
                                       'grn': 'G', 'extra': {'f': vals, 'title': forms.doc_title(key, vals, '강성환')}})
    assert row['status'] == ea.APPROVED and ea.title(row).startswith('G계약서 작성 요청 #')   # 승인 없이 바로 처리 대기
    assert ea.status_label(row) == ('🔵', '처리 대기')
    labels = [b['text']['text'] for b in ea.buttons(row, 'all')]
    assert labels == ['✅ 처리 완료', '반려', '요청 취소']
    t = ea.card_text(row, ea.steps(row['id']), 'G4237-MJ · 마포구 양화로 45')
    assert '인수자명: 윤명림' in t and '총 금액: 11,440,000원' in t and '⏳ 경영지원 처리' in t
    assert ea.can_act('reject', row, SB) and ea.can_act('transfer', row, SB) and not ea.can_act('transfer', row, MGR)
    done = ea.transition(row['id'], 'transfer', SB, '황샛별')
    assert ea.status_label(done) == ('🟢', '처리 완료')
    assert '✅ 경영지원 (SB) 처리' in ea.step_line(done, ea.steps(row['id']))


def test_cardpay_fee_and_validation():
    vals = {'project': 'R4120-YM', 'card_co': '신한', 'card_no': '1234-5678-9012-3456', 'expiry': '08/29',
            'sales': '1,000,000', 'installment': '3개월'}
    inst = next(f for f in forms.get(forms.FORM_CARDPAY).fields if f.key == 'installment')
    assert inst.kind == 'select' and inst.options[0] == '일시불' and inst.options[-1] == '12개월'
    assert forms.validate(forms.FORM_CARDPAY, vals) == {}
    assert forms.amount_of(forms.FORM_CARDPAY, vals) == 1030000
    assert forms.doc_title(forms.FORM_CARDPAY, vals, 'x') == 'R카드결제 요청'
    assert any('결제 금액: 1,030,000원' in l for l in forms.card_lines(forms.FORM_CARDPAY, vals))
    bad = forms.validate(forms.FORM_CARDPAY, {**vals, 'card_no': '1234', 'expiry': '2029-08'})
    assert set(bad) == {'card_no', 'expiry'}


def test_corpdocs_and_insurance_rules():
    v = {'grn': 'G (글로벌)', 'docs': ['인감증명서', '기타'], 'reason': '입찰', 'due': '2026-11-10',
         'receive': '등기', 'dest': '김포시 …'}
    assert set(forms.validate(forms.FORM_CORPDOCS, v)) == {'other'}
    assert forms.grn_of(forms.FORM_CORPDOCS, v) == 'G'
    year = next(f for f in forms.get(forms.FORM_CORPDOCS).fields if f.key == 'year')
    assert forms._field_options(forms.get(forms.FORM_CORPDOCS), year)[0] == f'{__import__("datetime").date.today().year}년'
    ins = {'kind': '하자보증', 'project': 'G4178-MW', 'total': '9130000', 'rate': '3', 'start': '2026-10-28',
           'years': '3', 'send': '이메일', 'email': ''}
    assert set(forms.validate(forms.FORM_INSURANCE, ins)) == {'email'}


def test_cardexp_prefill_title_and_memory(tmp_db):
    from datetime import date
    pre = forms.prefill(forms.get(forms.FORM_CARDEXP), None, {'card': '4518-1111-2222-3333'}, today=date(2026, 11, 2))
    assert pre == {'card_no': '4518-1111-2222-3333', 'start': '2026-10-01', 'end': '2026-10-31'}
    assert forms.doc_title(forms.FORM_CARDEXP, pre, '박민우') == '2026년 10월 법인카드지출결의서_박민우 매니저'
    ea.remember_personal('card', MGR, '강성환', '', '4518-1111-2222-3333')
    ea.remember_personal('staff', MGR, '강성환', '국민', '123-45-6789')
    assert ea.personal_memory(MGR) == {'card': '4518-1111-2222-3333', 'bank': '국민', 'account': '123-45-6789'}
    assert ea.known_partners() == []                                  # 본인 기억은 거래처 목록에 안 나옴


def test_expense_spend_flow_hides_account_on_card(se):
    vals = {'use_date': '2026-10-31', 'project': 'G4049-JW', 'cost': '기타비 · 운반비', 'amount': '35000',
            'reason': '현장 주차비', 'bank': '국민', 'account': '123-45-6789'}
    assert set(forms.validate(forms.FORM_EXPENSE, vals)) == {'grn'}   # 사업자 구분 필수(11-02)
    vals['grn'] = 'R (글로벌그룹)'                                     # G 현장이어도 고른 값으로
    assert forms.validate(forms.FORM_EXPENSE, vals) == {}
    row = se.process_form_submit(MGR, forms.FORM_EXPENSE, vals)
    assert row['status'] == ea.PENDING and row['cost_category'] == '기타비:운반비' and row['amount'] == 35000
    assert row['grn'] == 'R'
    card = se._fake.of('chat_postMessage')[0]
    txt = json.dumps(card['blocks'], ensure_ascii=False)
    assert '123-45-6789' not in txt and '본인 계좌 (경영지원 DM에서 확인)' in txt
    assert ea.personal_memory(MGR)['account'] == '123-45-6789'
    assert se.apply_action('approve', row['id'], KIKO) is None
    ex_dm = [kw for kw in se._fake.of('chat_postMessage') if kw['channel'] == 'D-' + SB][-1]
    assert '123-45-6789' in json.dumps(ex_dm['blocks'], ensure_ascii=False)   # 경영지원 DM 에만 계좌
    assert se.apply_action('transfer', row['id'], SB) is None
    assert any('개인경비 35,000원' in kw.get('text', '') for kw in se._fake.of('chat_postMessage'))


def test_task_form_submit_goes_straight_to_executor(se):
    row = se.process_form_submit(MGR, forms.FORM_CONTRACT, dict(CONTRACT))
    assert row['status'] == ea.APPROVED
    assert {(r['user'], r['role']) for r in ea.dm_refs(row)} == {(SB, 'executor')}
    dm = [kw for kw in se._fake.of('chat_postMessage') if kw['channel'] == 'D-' + SB][0]
    assert '처리하고 카드에서 [처리 완료]' in json.dumps(dm['blocks'], ensure_ascii=False)
    assert se.apply_action('transfer', row['id'], JW) is None
    msgs = se._fake.of('chat_postMessage')
    assert any('경영지원 (JW) 처리 완료' in kw.get('text', '') and kw.get('reply_broadcast') for kw in msgs)
    assert any('처리가 끝났어요' in kw.get('text', '') and kw['channel'] == 'D-' + MGR for kw in msgs)


def test_form_submit_errors_use_real_block_ids(se):
    acks = []
    view = _form_state(forms.FORM_CONTRACT, {**CONTRACT, 'total': ''}, pv=2)
    se.handle_form_submit(lambda **kw: acks.append(kw), {'user': {'id': MGR}}, view)
    assert acks[0]['response_action'] == 'errors' and 'total_2' in acks[0]['errors']


def test_cardexp_reminder_first_business_day_once(se, monkeypatch):
    from datetime import date

    class _R:
        def __init__(self):
            self.d = {}

        def set(self, k, v, nx=False, ex=None):
            if nx and k in self.d:
                return None
            self.d[k] = v
            return True
    rc = type('RC', (), {'redis': _R()})()
    from dashboard.utils import redis_client
    monkeypatch.setattr(redis_client, 'get_redis_client', lambda: rc)
    assert se.post_cardexp_reminder(date(2026, 12, 2)) is None        # 12/1(화)이 첫 영업일
    assert se.post_cardexp_reminder(date(2026, 12, 1))                # 보냄
    assert se.post_cardexp_reminder(date(2026, 12, 1)) is None        # 같은 달 두 번 안 보냄
    post = se._fake.of('chat_postMessage')[-1]
    assert '11월 법인카드 지출결의서' in post['blocks'][0]['text']['text']
    assert post['blocks'][1]['elements'][0]['value'] == forms.FORM_CARDEXP


def test_form_views_are_valid_shape():
    for k in forms.FORMS:
        v = forms.form_view(k, {'pv': 1})
        ids = [b.get('block_id') for b in v['blocks'] if b.get('block_id')]
        assert len(ids) == len(set(ids)) and len(v['blocks']) <= 100 and len(v['title']['text']) <= 24
        for b in v['blocks']:
            el = b.get('element') or {}
            if el.get('type') in ('static_select', 'multi_static_select'):
                assert 0 < len(el['options']) <= 100
            if 'placeholder' in el:
                assert el['placeholder']['text']


def test_contract_guarantees_show_fields_validate_and_card(tmp_db):
    """계약서 「보증보험」 고르면 그 보증의 보증금율·기간 칸이 생기고 필수가 된다(11-01)."""
    ids0 = [b.get('block_id') for b in forms.form_view(forms.FORM_CONTRACT, {})['blocks']]
    assert 'guarantees' in ids0 and not any(i and i.startswith('g_defect') for i in ids0)
    v = forms.form_view(forms.FORM_CONTRACT, {}, {'guarantees': ['하자보증']})
    ids = [b.get('block_id') for b in v['blocks']]
    assert 'g_defect_rate' in ids and 'g_defect_years' in ids and 'g_contract_rate' not in ids
    assert next(b for b in v['blocks'] if b.get('block_id') == 'guarantees')['dispatch_action'] is True
    vals = {**CONTRACT, 'guarantees': ['하자보증', '계약보증'], 'g_defect_rate': '3', 'g_defect_years': '2'}
    assert set(forms.validate(forms.FORM_CONTRACT, vals)) == {'g_contract_rate', 'g_contract_years'}
    vals.update(g_contract_rate='10', g_contract_years='1')
    assert forms.validate(forms.FORM_CONTRACT, vals) == {}
    assert forms.guarantee_summary(vals) == '하자보증 3% · 2년, 계약보증 10% · 1년'
    lines = forms.card_lines(forms.FORM_CONTRACT, vals)
    assert '보증보험: 하자보증 3% · 2년, 계약보증 10% · 1년' in lines and not any('보증금율' in l for l in lines)


def test_insurance_contract_note(se, monkeypatch):
    monkeypatch.setattr(se, '_flex_contract_codes', lambda: {'R4100-YM'})
    assert '계약서 작성 요청이 아직 없어요' in se.contract_note('G4237-MJ')
    assert 'Flex로 올린 계약서' in se.contract_note('R4100-YM')
    vals = {**CONTRACT, 'guarantees': ['하자보증'], 'g_defect_rate': '3', 'g_defect_years': '2'}
    row = se.process_form_submit(MGR, forms.FORM_CONTRACT, vals)
    note = se.contract_note('G4237-MJ')
    assert ea.title(row) in note and '하자보증 3% · 2년' in note
    v = forms.form_view(forms.FORM_INSURANCE, {'project': 'G4237-MJ', 'extra_note': note})
    assert note in json.dumps(v['blocks'], ensure_ascii=False)


def test_contract_note_shorter_when_project_fixed():
    menu = forms.form_view(forms.FORM_CONTRACT, {})['blocks'][0]['elements'][0]['text']
    assert '현장을 고르면' in menu
    card = forms.form_view(forms.FORM_CONTRACT, {'project': 'G4239-YM', 'extra_note': '✅ 사업자등록증 있음'})
    first = card['blocks'][0]
    assert first['type'] == 'context' and first['elements'][0]['text'] == '✅ 사업자등록증 있음'   # 고정 문구 없음


def test_license_note_follows_invoice_rules(se, monkeypatch):
    """계약서 창 등록증 안내 = 세금계산서 요청과 같은 판정(공사 폴더 > 같은 사업자 > 거래처 탭)."""
    import dashboard.services.business_license_handler as blh
    import dashboard.services.partner_status_sync as pss
    monkeypatch.setattr(se, '_project_rec', lambda code: {'사업자명': '주식회사 엠제이디앤엠'} if code != 'P-' else {'사업자명': '-'})
    states = {'own': {'exists': True, 'source': 'own', 'view_url': 'https://d/f'},
              'partner': {'exists': True, 'source': 'partner', 'view_url': None},
              'none': {'exists': False, 'source': None, 'view_url': None}}
    monkeypatch.setattr(pss, 'get_partner_info', lambda biz: {'bno': '123-45-67890'})
    for k, want in (('own', '✅ 사업자등록증 있음'), ('partner', '거래처 탭에 등록된 거래처'), ('none', '⚠️')):
        monkeypatch.setattr(blh, 'get_license_state', lambda code, k=k: states[k])
        assert want in se.license_note('G4239-YM')
    assert '123-45-67890' in (monkeypatch.setattr(blh, 'get_license_state', lambda c: states['partner'])
                              or se.license_note('G4239-YM'))
    assert '사업자명이 없어요' in se.license_note('P-')


def test_reset_select_changes_only_that_block_id(se):
    """고른 뒤 칸을 처음 상태로 — 그 actions 블록 block_id 만 바뀌고 나머지는 그대로(11-01)."""
    blocks = [{'type': 'section', 'block_id': 's1', 'text': {'type': 'mrkdwn', 'text': '카드'}},
              {'type': 'actions', 'block_id': 'a1', 'elements': [{'type': 'static_select', 'action_id': 'x'}]}]
    body = {'message': {'ts': '9.9', 'text': '[공사 확정] G4239-YM', 'blocks': json.loads(json.dumps(blocks))},
            'channel': {'id': 'CPROJ'}, 'actions': [{'block_id': 'a1'}]}
    c = _Fake()
    se.reset_select(c, body)
    (kw,) = c.of('chat_update')
    assert kw['channel'] == 'CPROJ' and kw['ts'] == '9.9' and kw['text'] == '[공사 확정] G4239-YM'
    assert kw['blocks'][0] == blocks[0] and kw['blocks'][1]['block_id'] != 'a1'
    assert kw['blocks'][1]['elements'] == blocks[1]['elements']


def test_project_card_select_options():
    from dashboard.services.project_slack_notifier import _eapproval_buttons
    import dashboard.blueprints.slack_eapproval as se_mod
    orig = se_mod.is_configured
    se_mod.is_configured = lambda: True
    try:
        (el,) = _eapproval_buttons('G4049-JW')
    finally:
        se_mod.is_configured = orig
    assert el['type'] == 'static_select' and el['action_id'] == ea.A_PROJECT_FORM_PICK
    assert [o['value'] for o in el['options']] == ['contract|G4049-JW', 'insurance|G4049-JW', 'cardpay|G4049-JW']


def test_partner_options_grouped_and_sorted(tmp_db, monkeypatch):
    """빈 클릭=최근 15곳 + 나머지 가나다순, 입력=그 글자로 시작하는 곳 먼저(11-01)."""
    recent = [f'최근{i:02d}' for i in range(15)]
    monkeypatch.setattr(ea, 'known_partners', lambda: recent + ['하나상사', '(주)가람', '나비엔'])
    groups = ea.partner_options('')['option_groups']
    assert [g['label']['text'] for g in groups] == ['최근 거래한 곳', '가나다순']
    assert [o['value'] for o in groups[0]['options']] == recent
    assert [o['value'] for o in groups[1]['options']] == ['(주)가람', '나비엔', '하나상사']
    monkeypatch.setattr(ea, 'known_partners', lambda: ['서울나비', '나비엔하우징', '가나비'])
    assert [o['value'] for o in ea.partner_options('나비')['options']] == ['나비엔하우징', '가나비', '서울나비', 'new:나비']


def test_grn_required_hidden_only_when_invoice_has_it(tmp_db):
    """사업자 구분은 필수 — 계산서 글에 G/R 이 있을 때만 칸을 숨긴다(11-01)."""
    bids = lambda v: {b.get('block_id'): b for b in v['blocks']}   # noqa: E731
    g = bids(ea.new_request_view({}))['grn']
    assert not g.get('optional') and 'initial_option' not in g['element']
    inv = {'ch': 'C1', 'ts': '1.0', 'parsed': {'grn': 'G', 'amount': 1000, 'partner': '강원포장'}}
    assert 'grn' not in bids(ea.new_request_view({'inv': inv}))
    inv['parsed']['grn'] = ''
    assert 'grn' in bids(ea.new_request_view({'inv': inv}))
