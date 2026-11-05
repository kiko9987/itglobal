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
    # 계산서 현장 지정 장부도 임시 DB로 — 안 그러면 작업 폴더 instance/users.db 에 시험 줄이 남는다(11-05)
    from dashboard.services import purchase_invoice_assign as pinv
    monkeypatch.setattr(pinv, '_db', lambda: db)
    monkeypatch.setattr(pinv, '_table_ready', False)
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
    assert ids(row) == [ea.A_APPROVE, ea.A_REJECT, ea.A_REVISE, ea.A_CANCEL]
    assert ids(row, 'approver') == [ea.A_APPROVE, ea.A_REJECT]
    assert ids(row, 'requester') == [ea.A_CANCEL]
    assert ids(row, 'executor') == []
    row = ea.transition(row['id'], 'approve', KIKO, '고광일')
    assert ids(row) == [ea.A_TRANSFER, ea.A_REVISE, ea.A_CANCEL]
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
    assert set(ea.complete_supply_request(f)) == {'item', 'link', 'qty', 'unit_price', 'reason'}
    # 사업자 구분 칸 없음 — 어느 통장인지는 경영지원 판단(11-05)
    assert 'grn' not in {b.get('block_id') for b in ea.supply_request_view({})['blocks']}


def test_supply_card_and_buttons(tmp_db):
    f = ea.read_supply_request(_supply_view())
    ea.complete_supply_request(f)
    row = ea.create(ea.FORM_SUPPLY, KIKO, '고광일', f)
    assert ea.extra(row)['qty'] == 2 and ea.status_label(row) == ('🔵', '구매 대기')
    t = ea.card_text(row, ea.steps(row['id']), '프로젝트 0 · 회사 재고')
    assert f'비품 신청 #{row["id"]}' in t and '품명: A4 용지 박스 × 2개 (단가 25,000원)' in t
    assert '상품 보기' in t and '사유: 용지 소진' in t and '이체 계좌' not in t and '⏳ 구매' in t
    assert '사업자 구분' not in t and not row['grn']
    btn = ea.buttons(row, 'executor')[0]
    assert btn['text']['text'] == '🛒 구매 완료' and btn['action_id'] == ea.A_TRANSFER
    assert '구매하고 카드에서 [구매 완료]' in ea.dm_intro(row, 'executor')
    row = ea.create(ea.FORM_SUPPLY, KIKO, '고광일', {**f, 'memo': '배송 메시지에 직접 수령'})
    t = ea.card_text(row, ea.steps(row['id']), '프로젝트 0 · 회사 재고')
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
    assert v['blocks'] == ea.hub_blocks()
    assert 'submit' not in v                      # 고르면 입력 창이 위에 쌓인다
    # 고르기 칸 없이 양식마다 버튼 — action_id 는 블록 안에서 겹치지 않고, 메뉴 양식이 빠짐없이
    acts = [e for b in v['blocks'] if b['type'] == 'actions' for e in b['elements']]
    assert all(e['type'] == 'button' for e in acts)
    ids = [e['action_id'] for e in acts]
    assert len(ids) == len(set(ids))
    assert {e['value'] for e in acts if e['action_id'].startswith(ea.A_FORM_OPEN)} == \
        {o['value'] for o in ea.forms.menu_options()}
    import re
    pat = re.compile(rf'^{ea.A_FORM_OPEN}(_.+)?$')
    assert all(pat.match(i) for i in ids if i.startswith(ea.A_FORM_OPEN)) and pat.match(ea.A_FORM_OPEN)
    assert not pat.match(ea.A_FORM_PICK)


class _ViewClient:
    def __init__(self):
        self.calls = []

    def views_open(self, **kw):
        self.calls.append(('open', None))
        return {'view': {'id': 'V-new'}}

    def views_update(self, **kw):
        self.calls.append(('update', kw['view_id']))

    def views_push(self, **kw):
        self.calls.append(('push', kw.get('trigger_id') or kw.get('interactivity_pointer')))
        return {'view': {'id': 'V-pushed'}}


def test_hub_buttons_push_over_menu(monkeypatch):
    """메뉴 창 안에서 누르면 그 위에 쌓는다([취소] → 메뉴로). 채널 메뉴 메시지면 새 창."""
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se.threading, 'Thread', lambda target, daemon=True, args=(): type('T', (), {'start': lambda s: None})())
    in_view = {'container': {'type': 'view'}, 'view': {'id': 'V-hub', 'type': 'modal'}, 'trigger_id': 'T1'}
    in_msg = {'container': {'type': 'message'}, 'channel': {'id': 'C1'}, 'trigger_id': 'T2'}
    assert se._hub_view_id(in_view) == 'V-hub' and se._hub_view_id(in_msg) == ''
    assert se._push_kw(in_view) == {'trigger_id': 'T1'} and se._push_kw(in_msg) is None

    c = _ViewClient()
    se.open_new_modal(c, 'T', 'U1', origin='shortcut', push={'trigger_id': 'T1'})
    se.open_form_modal(c, 'T', 'U1', ea.forms.menu_options()[0]['value'], push={'trigger_id': 'T1'})
    assert c.calls == [('push', 'T1'), ('push', 'T1')]

    c = _ViewClient()
    se.open_new_modal(c, 'T', 'U1', origin='menu')
    assert c.calls == [('open', None)]


def test_home_tab_opens_new_modal_not_update(monkeypatch):
    """홈 탭 버튼은 container=view 지만 모달이 아니라 홈을 바꾸면 안 된다 — 새 창으로."""
    from dashboard.blueprints import slack_eapproval as se
    home = {'container': {'type': 'view'}, 'view': {'id': 'V-home', 'type': 'home'}}
    assert se._hub_view_id(home) == '' and se._from_home(home)
    assert ea.home_view()['type'] == 'home' and ea.home_view()['blocks'] == ea.hub_blocks()

    published = []

    class _C:
        def views_publish(self, **kw):
            published.append(kw)

    se.reset_home(_C(), {'user': {'id': 'U1'}})
    acts = [b for b in published[0]['view']['blocks'] if b['type'] == 'actions'][0]
    assert published[0]['user_id'] == 'U1' and acts['block_id'].startswith('eap_sel_')


def test_workflow_step_opens_hub_and_completes(monkeypatch):
    """워크플로 단계 「전자결재 열기」 — 단계 토큰·pointer 로 메뉴 창을 열고 단계를 끝낸다. manifest 키도 같아야."""
    import os
    import yaml
    from dashboard.blueprints import slack_eapproval as se
    fns, opened, done = {}, [], []

    class _App:
        def function(self, cid):
            return lambda f: fns.setdefault(cid, f)

        def __getattr__(self, name):
            return lambda *a, **k: (lambda f: f)

    class _WC:
        def __init__(self, token):
            self.token = token

        def views_open(self, **kw):
            opened.append((self.token, kw['interactivity_pointer'], kw['view']['title']['text']))

    import slack_sdk
    monkeypatch.setattr(slack_sdk, 'WebClient', _WC)
    se.register_eapproval_handlers(_App())
    fns[ea.FN_OPEN_MENU](body={'event': {'bot_access_token': 'xwfp-1'}},
                         inputs={'interactivity': {'interactivity_pointer': 'P1'}},
                         fail=lambda error: done.append(error))
    # 창을 연 직후엔 단계를 끝내지 않는다 — 끝내면 슬랙이 창 버튼을 막는다
    assert opened == [('xwfp-1', 'P1', '전자결재')] and done == []
    assert ea.hub_view().get('notify_on_close') is True
    path = os.path.join(os.path.dirname(__file__), '..', '..', 'docs', 'slack_eapproval_manifest.yml')
    with open(path, encoding='utf-8') as fh:
        man = yaml.safe_load(fh)
    assert ea.FN_OPEN_MENU in man['functions']


def test_step_completes_on_close_and_successful_submit(monkeypatch):
    """워크플로 창: 닫힘·오류 없는 제출이면 단계 끝냄, 오류 ack·일반 창이면 안 끝냄."""
    from dashboard.blueprints import slack_eapproval as se
    done = []
    monkeypatch.setattr(se.threading, 'Thread', lambda target, daemon=True, args=(): type('T', (), {'start': lambda s: target()})())

    class _WC:
        def __init__(self, token):
            pass

        def functions_completeSuccess(self, function_execution_id, outputs):
            done.append(function_execution_id)

    import slack_sdk
    monkeypatch.setattr(slack_sdk, 'WebClient', _WC)
    fbody = {'function_data': {'execution_id': 'Fx1'}, 'bot_access_token': 'xwfp-1'}
    se._ack_completing_step(lambda **k: None, fbody)(response_action='errors', errors={'a': 'b'})
    assert done == []
    se._ack_completing_step(lambda **k: None, fbody)()
    assert done == ['Fx1']
    se._complete_step({'view': {'id': 'V'}})              # 채널 메뉴로 연 일반 창
    assert done == ['Fx1']


def test_workflow_view_buttons_work_without_trigger_id(monkeypatch):
    """워크플로로 연 창의 버튼 요청엔 trigger_id 가 없다(interactivity_pointer 만) — 그 pointer 로 쌓는다."""
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se.threading, 'Thread', lambda target, daemon=True, args=(): type('T', (), {'start': lambda s: None})())
    acts = {}

    class _App:
        def action(self, aid):
            return lambda f: acts.setdefault(aid, f)

        def __getattr__(self, name):
            return lambda *a, **k: (lambda f: f)

    se.register_eapproval_handlers(_App())
    body = {'container': {'type': 'view'}, 'view': {'id': 'V-wf', 'type': 'modal'}, 'user': {'id': 'U1'},
            'interactivity': {'interactivity_pointer': 'P9'},
            'actions': [{'selected_option': {'value': ea.forms.menu_options()[0]['value']}}]}
    c = _ViewClient()
    acts[ea.A_NEW](ack=lambda: None, body=body, client=c)
    acts[ea.A_NEW_SUPPLY](ack=lambda: None, body=body, client=c)
    acts[ea.A_FORM_PICK](ack=lambda: None, body=body, client=c)
    assert c.calls == [('push', 'P9')] * 3


def test_stacked_view_cancel_keeps_step_and_submit_clears():
    """쌓인 창 [취소] → 메뉴로(단계 유지). 전부 닫힘·맨 아래 창 닫힘이면 끝냄. 쌓인 창 제출은 clear."""
    from dashboard.blueprints import slack_eapproval as se
    done, acks = [], []
    orig = se._complete_step
    se._complete_step = lambda body: done.append(body.get('tag'))
    try:
        closed = {}

        class _App:
            def view_closed(self, cid):
                return lambda f: closed.setdefault('h', f)

            def __getattr__(self, name):
                return lambda *a, **k: (lambda f: f)

        se.register_eapproval_handlers(_App())
        h = closed['h']
        h(ack=lambda: None, body={'tag': 'back', 'view': {'previous_view_id': 'V-hub'}, 'is_cleared': False})
        h(ack=lambda: None, body={'tag': 'x', 'view': {'previous_view_id': 'V-hub'}, 'is_cleared': True})
        h(ack=lambda: None, body={'tag': 'root', 'view': {'previous_view_id': None}})
        assert done == ['x', 'root']

        se._ack_completing_step(lambda **k: acks.append(k), {'view': {'previous_view_id': 'V-hub'}})()
        se._ack_completing_step(lambda **k: acks.append(k), {'view': {}})()
        assert acks == [{'response_action': 'clear'}, {}]
    finally:
        se._complete_step = orig


def test_bank_is_select_and_old_names_map():
    """은행은 고르기 칸 — 예전 직접 입력값(29가지)은 항목으로 맞춰 미리 채우고, 애매하면 비운다(11-05)."""
    nb = forms.normalize_bank
    assert nb('국민') == nb('KB국민은행') == nb('국민은행') == 'KB국민은행'
    assert nb('신한은핸') == '신한은행' and nb('IBK기업은행') == nb('기업') == 'IBK기업은행'
    assert nb('농협은행') == nb('농협') == '농협' and nb('sc제일은행') == nb('한국스탠다드차타드은행 SC') == 'SC제일은행'
    assert nb('새마을 금고') == '새마을금고' and nb('정보통신부우체국') == '우체국'
    assert nb('신한은행(농협)') == '' and nb('') == '' and nb('모르는은행') == ''
    el = next(b for b in ea.transfer_input_blocks({'bank': '국민'}) if b['block_id'] == 'bank')['element']
    assert el['type'] == 'static_select' and el['initial_option']['value'] == 'KB국민은행'
    el = next(b for b in ea.transfer_input_blocks({'bank': '신한은행(농협)'}) if b['block_id'] == 'bank')['element']
    assert 'initial_option' not in el
    vals = {'bank_3': {'bank': {'type': 'static_select', 'selected_option': {'value': '농협'}}}}
    assert ea.field_value(vals, 'bank') == '농협'                        # 다시 채운 칸(bank_N)도
    f = next(x for x in forms.get(forms.FORM_EXPENSE).fields if x.key == 'bank')
    assert f.kind == 'select' and f.options == forms.BANKS


def test_contract_form_has_section_dividers():
    """계약서 창은 칸이 많아 묶음마다 구분선·제목 — 대금·인수·발송·보증보험, 비고 앞은 구분선만(11-05)."""
    blocks = forms.form_view(forms.FORM_CONTRACT, {'form': forms.FORM_CONTRACT})['blocks']
    titles = [b['text']['text'] for b in blocks if b['type'] == 'section' and b['text']['text'].startswith('*')]
    assert titles == ['*💰 대금*', '*📦 인수·납기*', '*✉️ 계약서 발송*', '*🛡️ 보증 조건 (계약서 기재)*']
    assert sum(b['type'] == 'divider' for b in blocks) == 5
    ids = [b.get('block_id') for b in blocks if b['type'] == 'input']
    assert ids.index('total') < ids.index('deposit')                 # 총액 → 나눠 받는 금액 순
    # 구분선이 있어도 읽기는 그대로(입력 칸만 읽음)
    assert 'section' not in {f.kind for f in forms.get(forms.FORM_CONTRACT).fields}


def test_insurance_fills_rate_years_from_contract(monkeypatch):
    """보증보험 신청 — 이 현장 계약서에 적은 보증금율·기간을 채운다. 계약서 보증이 하나면 종류도(11-05)."""
    from dashboard.blueprints import slack_eapproval as se
    terms = {'하자보증': ('3', '2'), '계약보증': ('10', '1')}
    monkeypatch.setattr(se, 'contract_terms', lambda code: dict(terms) if code == 'G1' else {})
    v = {'kind': '하자보증'}
    se.fill_insurance_from_contract(v, 'G1')
    assert v['rate'] == '3' and v['years'] == '2'
    v = {'kind': '하자보증', 'rate': '5'}                       # 직접 넣은 값은 그대로
    se.fill_insurance_from_contract(v, 'G1')
    assert v['rate'] == '5' and v['years'] == '2'
    v = {'kind': '계약보증', 'rate': '3', 'years': '2'}        # 종류를 바꿔 고르면 그 종류 값으로
    se.fill_insurance_from_contract(v, 'G1', override=True)
    assert (v['rate'], v['years']) == ('10', '1')
    v = {}
    se.fill_insurance_from_contract(v, 'G1')                     # 둘이면 종류는 고르게 둔다
    assert v == {}
    terms.pop('계약보증')
    se.fill_insurance_from_contract(v, 'G1')                     # 하나면 종류까지
    assert v == {'kind': '하자보증', 'rate': '3', 'years': '2'}
    v = {'kind': '하자보증'}
    se.fill_insurance_from_contract(v, 'X')                      # 계약서 없음 → 그대로
    assert v == {'kind': '하자보증'}
    ins = forms.get(forms.FORM_INSURANCE)
    assert next(f for f in ins.fields if f.key == 'kind').dispatch
    titles = [b['text']['text'] for b in forms.form_view(forms.FORM_CONTRACT, {})['blocks']
              if b['type'] == 'section']
    assert '*🛡️ 보증 조건 (계약서 기재)*' in titles


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
            'send': '대면', 'contact': '010-4312-1990', 'deposit': '5720000', 'deposit_date': '2026-10-30',
            'balance': '5720000'}


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
    assert labels == ['✅ 처리 완료', '반려', '✏️ 수정 요청', '요청 취소']
    t = ea.card_text(row, ea.steps(row['id']), 'G4237-MJ · 마포구 양화로 45')
    assert '인수자명: 윤명림' in t and '총 금액: 11,440,000원' in t and '⏳ 경영지원 처리' in t
    assert ea.can_act('reject', row, SB) and ea.can_act('transfer', row, SB) and not ea.can_act('transfer', row, MGR)
    done = ea.transition(row['id'], 'transfer', SB, '황샛별')
    assert ea.status_label(done) == ('🟢', '처리 완료')
    assert '✅ 경영지원 (SB) 처리' in ea.step_line(done, ea.steps(row['id']))


def test_cardpay_fee_and_validation():
    vals = {'project': 'R4120-YM', 'card_co': '신한', 'card_no': '1234-5678-9012-3456', 'expiry': '08/29',
            'sales': '1,000,000', 'fee': '30,000', 'total': '1,030,000', 'installment': '3개월'}
    inst = next(f for f in forms.get(forms.FORM_CARDPAY).fields if f.key == 'installment')
    assert inst.kind == 'select' and inst.options[0] == '일시불' and inst.options[-1] == '12개월'
    assert forms.validate(forms.FORM_CARDPAY, vals) == {}
    assert forms.amount_of(forms.FORM_CARDPAY, vals) == 1030000
    assert forms.doc_title(forms.FORM_CARDPAY, vals, 'x') == 'R카드결제 요청'
    assert any('결제 요청 금액: 1,030,000원' in l for l in forms.card_lines(forms.FORM_CARDPAY, vals))
    bad = forms.validate(forms.FORM_CARDPAY, {**vals, 'card_no': '1234', 'expiry': '2029-08'})
    assert set(bad) == {'card_no', 'expiry'}
    # 수수료는 3% 고정 아님 — 0(회사·고객 부담)·1.5% 도 되고, 합계는 검수(11-05)
    assert forms.validate(forms.FORM_CARDPAY, {**vals, 'fee': '0', 'total': '1,000,000'}) == {}
    assert forms.validate(forms.FORM_CARDPAY, {**vals, 'fee': '15000', 'total': '1015000'}) == {}
    assert set(forms.validate(forms.FORM_CARDPAY, {**vals, 'total': '1,000,000'})) == {'total'}
    assert set(forms.validate(forms.FORM_CARDPAY, {**vals, 'fee': ''})) == {'fee'}


def test_contract_parts_must_add_up():
    """계약금+중도금+잔금 합 = 총 금액 검수 — 나눠 받는 금액을 안 넣었으면 검사 안 함(11-05)."""
    base = {'project': 'G4049-JW', 'total': '10,000,000', 'receiver': 'a', 'receiver_phone': '010',
            'send': '이메일', 'contact': 'a@b'}
    assert forms.validate(forms.FORM_CONTRACT, base) == {}
    ok = {**base, 'deposit': '3,000,000', 'balance': '7,000,000'}
    assert forms.validate(forms.FORM_CONTRACT, ok) == {}
    bad = forms.validate(forms.FORM_CONTRACT, {**ok, 'balance': '6,000,000'})
    assert set(bad) == {'total'} and '9,000,000원' in bad['total']


def test_expense_reason_last_and_fine_category():
    keys = [f.key for f in forms.get(forms.FORM_EXPENSE).fields]
    assert keys[-1] == 'reason'
    from dashboard.services.purchase_invoice_assign import COST_CATEGORIES
    assert ('기타비:과태료', '기타비 · 과태료') in COST_CATEGORIES


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
    # 사업자 구분 칸 없음 — 경영지원이 판단(11-05, 플렉스도 없었음)
    assert 'grn' not in {f.key for f in forms.get(forms.FORM_EXPENSE).fields}
    assert forms.validate(forms.FORM_EXPENSE, vals) == {}
    row = se.process_form_submit(MGR, forms.FORM_EXPENSE, vals)
    assert row['status'] == ea.PENDING and row['cost_category'] == '기타비:운반비' and row['amount'] == 35000
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
    assert '보증 조건: 하자보증 3% · 2년, 계약보증 10% · 1년' in lines and not any('보증금율' in l for l in lines)


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


# ── 수정 요청 → 수정하기 · 첨부 (11-05) ─────────────────────────────
def _state_from_view(view: dict, override: dict = None) -> dict:
    """창의 처음 값(initial_*)을 제출 state 로 — 사용자가 손대지 않고 일부만 고친 것처럼."""
    st = {}
    for b in view['blocks']:
        if b.get('type') != 'input':
            continue
        el = b['element']
        aid = el.get('action_id')
        if 'initial_value' in el:
            v = {'value': el['initial_value']}
        elif 'initial_option' in el:
            v = {'selected_option': el['initial_option']}
        elif 'initial_options' in el:
            v = {'selected_options': el['initial_options']}
        elif 'initial_date' in el:
            v = {'selected_date': el['initial_date']}
        else:
            v = {}
        st[b['block_id']] = {aid: v}
    for bid, v in (override or {}).items():
        aid = next(iter(st[bid])) if bid in st else bid
        st[bid] = {aid: v}
    return {'state': {'values': st}, 'private_metadata': view['private_metadata']}


def test_revise_then_edit_fund_same_card(se):
    row = se.process_new_submit(MGR, dict(FIELDS))
    card_ts = row['card_ts']
    assert se.apply_action('revise', row['id'], MGR, '계좌 확인') == '수정 요청은 승인자·경영지원만 할 수 있어요.'
    assert se.apply_action('revise', row['id'], SB, '계좌번호가 계산서와 달라요') is None
    row = ea.get(row['id'])
    assert row['status'] == ea.REVISE and ea.status_label(row) == ('🟠', '수정 요청')
    dm = [m for m in se._fake.of('chat_postMessage') if m['channel'] == 'D-' + MGR][-1]
    assert ea.A_EDIT in json.dumps(dm['blocks']) and '계좌번호가 계산서와 달라요' in json.dumps(dm['blocks'], ensure_ascii=False)
    assert not ea.can_act('edit', row, SB) and ea.can_act('edit', row, MGR)
    text = ea.card_text(row, ea.steps(row['id']), 'G4049-JW')
    assert '✏️ 수정 요청 (SB' in text
    # 수정 창 — 올린 값이 채워져 있다
    v = se.edit_view(row)
    assert v['title']['text'] == '자금집행 요청 수정' and v['submit']['text'] == '다시 올리기'
    sub = _state_from_view(v, {'amount': {'value': '550000'}, 'account': {'value': '110-511-000000'}})
    f, errs = se.validate_new_submit(sub)
    assert errs == {} and f['_edit'] == row['id'] and f['content'] == 'FCU 부품' and f['project_code'] == 'G4049-JW'
    new = se.process_new_submit(MGR, f)
    assert new['id'] == row['id'] and new['card_ts'] == card_ts and new['status'] == ea.PENDING
    assert new['amount'] == 550000 and new['account'] == '110-511-000000'
    logs = [m['text'] for m in se._fake.of('chat_postMessage') if m.get('thread_ts') == card_ts]
    assert any('수정해서 다시 올림' in t and '금액 533,500 → 550,000' in t for t in logs)
    # 다시 올린 뒤엔 승인 기록이 새로 — 승인 대기로 보이고, 승인하면 이체 대기
    assert '⏳ KiKO 승인' in ea.step_line(new, ea.steps(new['id']))
    assert se.apply_action('approve', new['id'], KIKO) is None and ea.get(new['id'])['status'] == ea.APPROVED


def test_revise_task_form_and_resubmit_back_to_processing(se):
    from dashboard.services import e_approval_forms as forms
    vals = {'grn': 'G (글로벌)', 'project': '', 'docs': ['인감증명서'], 'year': '', 'other': '',
            'reason': '입찰 제출용', 'cert_memo': '', 'due': '2026-11-10', 'receive': '이메일', 'dest': 'a@b.c'}
    row = se.process_form_submit(MGR, forms.FORM_CORPDOCS, dict(vals))
    assert row['status'] == ea.APPROVED
    assert se.apply_action('revise', row['id'], KIKO, '제출처를 적어 주세요') is None
    v = se.edit_view(ea.get(row['id']))
    assert v['title']['text'].endswith('수정') and 'attach' in [b.get('block_id') for b in v['blocks']]
    sub = _state_from_view(v, {'reason': {'value': '조달청 입찰 제출용'}})
    key, vals2, errs = se.validate_form_submit(sub)
    assert errs == {} and vals2['_edit'] == row['id']
    new = se.process_form_submit(MGR, key, vals2)
    assert new['status'] == ea.APPROVED and ea.extra(new)['f']['reason'] == '조달청 입찰 제출용'
    assert ea.title(new) == ea.title(row)   # 결재 이름은 처음 그대로


def test_supply_edit_prefill(se):
    f = {'content': 'A4 용지', 'link': 'https://x.y/1', 'qty': 2, 'unit_price': 25000, 'amount': 0,
         'reason': '소진', 'memo': ''}
    errs = ea.complete_supply_request(f)
    assert errs == {}
    row = se.process_supply_submit(MGR, f)
    se.apply_action('revise', row['id'], SB, '수량 확인')
    v = se.edit_view(ea.get(row['id']))
    init = {b['block_id']: b['element'].get('initial_value') for b in v['blocks'] if b.get('type') == 'input'}
    assert init['item'] == 'A4 용지' and init['qty'] == '2' and init['link'] == 'https://x.y/1'
    f2 = ea.read_supply_request(_state_from_view(v, {'qty': {'value': '3'}}))
    assert f2['_edit'] == row['id'] and ea.complete_supply_request(f2) == {}
    new = se.process_supply_submit(MGR, f2)
    assert new['amount'] == 75000 and ea.extra(new)['qty'] == 3


def test_attach_block_everywhere_and_read():
    from dashboard.services import e_approval_forms as forms
    assert 'attach' in [b.get('block_id') for b in ea.new_request_view({})['blocks']]
    assert 'attach' in [b.get('block_id') for b in ea.supply_request_view({})['blocks']]
    for k in forms.FORMS:
        assert forms.form_view(k, {'form': k})['blocks'][-1]['element']['type'] == 'file_input'
    view = {'state': {'values': {'attach': {'attach': {'type': 'file_input', 'files': [
        {'id': 'F1', 'name': '계산서.pdf', 'url_private_download': 'https://files/1'}]}}}}}
    assert forms.read_attachments(view) == [{'id': 'F1', 'name': '계산서.pdf', 'url': 'https://files/1'}]


def test_post_files_to_card_thread(se, monkeypatch):
    row = se.process_new_submit(MGR, dict(FIELDS))
    monkeypatch.setattr(se, '_download', lambda url: b'PDF' if url.endswith('ok') else None)
    n = se.post_files(row, [{'name': 'a.pdf', 'url': 'https://f/ok'}, {'name': 'b.pdf', 'url': 'https://f/x'}], '강성환')
    up = se._fake.of('files_upload_v2')
    assert n == 1 and up[0]['thread_ts'] == row['card_ts'] and up[0]['file_uploads'][0]['filename'] == 'a.pdf'
    warn = [m['text'] for m in se._fake.of('chat_postMessage') if '옮기지 못했어요' in (m.get('text') or '')]
    assert warn and 'b.pdf' in warn[0]


class _FakeDrive:
    def __init__(self, files):
        self._files = files

    def files(self):
        return self

    def list(self, **kw):
        self.kw = kw
        return self

    def execute(self):
        return {'files': self._files}


def test_folder_docs_final_first_and_license(se, monkeypatch):
    from dashboard.services import business_license_handler as bl
    files = [  # 드라이브는 최근 고친 순으로 준다
        {'id': 'D1', 'name': '-최종- 에어컨 도면 26.11.04.pdf', 'modifiedTime': '2026-11-05T00:00:00Z'},
        {'id': 'X1', 'name': '(MJ) 송파 26.10.30.xlsx', 'modifiedTime': '2026-10-30T01:00:00Z'},
        {'id': 'P3', 'name': '(MJ) 송파 26.10.30.pdf', 'modifiedTime': '2026-10-30T00:00:00Z'},
        {'id': 'P2', 'name': '(MJ) 송파 26.10.29 -최종 견적-.pdf', 'modifiedTime': '2026-10-29T00:00:00Z'},
        {'id': 'P1', 'name': '-공사확정- 송파 26.10.27.pdf', 'modifiedTime': '2026-10-27T00:00:00Z'},
        {'id': 'L1', 'name': '사업자등록증.pdf', 'modifiedTime': '2026-10-01T00:00:00Z'},
    ]
    monkeypatch.setattr(bl, '_project_folder_id_for', lambda code, fresh=False: 'FOLDER')
    monkeypatch.setattr(bl, '_get_drive', lambda: _FakeDrive(files))
    monkeypatch.setattr(bl, 'get_license_state', lambda code, use_cache=True: {'view_url': 'https://d/lic', 'source': 'own'})
    docs = se.folder_docs('G4049-JW', precheck=True)
    # 등록증 · 확정/최종 견적 · 확정/최종 도면 · 나머지 (엑셀·등록증 파일 제외). G4257-MW 처럼 최종 도면이 더 최근이어도
    assert [d['id'] for d in docs] == ['lic', 'P2', 'P1', 'D1', 'P3']
    assert [d['checked'] for d in docs] == [True, True, False, False, False]   # 확정·최종 견적서 하나만 미리 고름
    assert [d['checked'] for d in se.folder_docs('G4049-JW')] == [False] * 5


def test_folder_docs_block_carry_and_post(se):
    from dashboard.services import e_approval_forms as forms
    docs = [{'id': 'lic', 'name': '사업자등록증', 'note': '이 공사 폴더', 'checked': True},
            {'id': 'P2', 'name': '최종 견적.pdf', 'note': '2026-10-29', 'checked': False}]
    v = forms.form_view(forms.FORM_CONTRACT, {'form': forms.FORM_CONTRACT}, {'_docs': docs})
    blk = next(b for b in v['blocks'] if b.get('block_id') == 'folder_docs')
    assert [o['value'] for o in blk['element']['initial_options']] == ['lic']
    # 보증 고르기로 다시 그려도 지금 고른 것(P2 로 바꿈) 그대로
    view = {'blocks': v['blocks'], 'state': {'values': {'folder_docs': {'folder_docs': {
        'selected_options': [blk['element']['options'][1]]}}}}}
    kept = forms.carry_folder_docs(view)
    assert [o['value'] for o in kept['element']['initial_options']] == ['P2']
    assert forms.read_folder_docs(view) == [{'id': 'P2', 'name': '최종 견적.pdf'}]
    from dashboard.services import business_license_handler as bl
    row = se.process_new_submit(MGR, dict(FIELDS))
    import pytest as _pt
    with _pt.MonkeyPatch.context() as mp:
        mp.setattr(bl, '_project_folder_id_for', lambda code, fresh=False: 'FOLDER123')
        se.post_folder_docs(row, [{'id': 'lic', 'name': '사업자등록증'}, {'id': 'P2', 'name': '최종 견적.pdf'}])
    msg = se._fake.of('chat_postMessage')[-1]
    # 드라이브 웹 링크가 아니라 탐색기로 여는 폴더 링크(공사확정 카드와 같은 itgfolder://) + 파일 이름
    assert msg['thread_ts'] == row['card_ts'] and '<itgfolder://FOLDER123|탐색기에서 열기>' in msg['text']
    assert '• 최종 견적.pdf' in msg['text'] and 'drive.google.com' not in msg['text']


def test_invoice_photo_copied_to_card(se, monkeypatch):
    # 계산서 글 현장 지정(Redis)까지는 안 가게 결재만 만들어 시작
    row = se.start(ea.create(ea.FORM_FUND, MGR, '강성환', dict(FIELDS, origin_channel='CINV', origin_ts='9.0')))
    se._inv.conversations_history = lambda **kw: {'messages': [
        {'ts': '9.0', 'files': [{'name': 'image.png', 'url_private_download': 'https://f/inv'}]}]}
    monkeypatch.setattr(se, '_download', lambda url: b'IMG')
    assert se.post_invoice_files(row) == 1
    up = se._fake.of('files_upload_v2')[-1]
    assert up['thread_ts'] == row['card_ts'] and up['initial_comment'].startswith('📄 계산서')


def test_manifest_has_file_scopes():
    import os
    path = os.path.join(os.path.dirname(__file__), '..', '..', 'docs', 'slack_eapproval_manifest.yml')
    text = open(path, encoding='utf-8').read()
    assert '- files:read' in text and '- files:write' in text
