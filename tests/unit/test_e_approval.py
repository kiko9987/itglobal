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
    assert '고광일 승인' in line and '황샛별 이체' in line


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
        'partner': {'partner': {'value': over.get('partner', '강원포장')}},
        'amount': {'amount': {'value': over.get('amount', '150,000')}},
        'grn': {'grn': {'selected_option': ({'value': over['grn']} if over.get('grn') else None)}},
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
    f = ea.read_new_request(_new_view(project='프로젝트0', payee_name='a', bank='b', account='1234-5678'))
    assert set(ea.complete_new_request(f)) == {'grn'}
    f = ea.read_new_request(_new_view(project='프로젝트0', grn='N', payee_name='a', bank='b', account='1234-5678'))
    assert ea.complete_new_request(f) == {}


def test_fixed_project_from_metadata(tmp_db):
    v = _new_view(fixed='R4120-YM', payee_name='a', bank='b', account='1234-5678')
    del v['state']['values']['project']
    f = ea.read_new_request(v)
    assert f['project_code'] == 'R4120-YM' and ea.complete_new_request(f) == {} and f['grn'] == 'R'


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
    assert len(ea.card_blocks(row, ea.steps(row['id']), 'G4049-JW')) == 1  # 끝나면 버튼·안내 없음


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
    assert {(r['user'], r['role']) for r in refs} == {(KIKO, 'approver'), (MGR, 'requester')}
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
    assert {(r['user'], r['role']) for r in ea.dm_refs(row)} == {(SB, 'executor'), (KIKO, 'requester')}


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
