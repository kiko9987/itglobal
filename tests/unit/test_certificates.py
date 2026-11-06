"""증명서 신청 (2026-11-06) — 장부·창·카드·처리 흐름(DM 만)."""
import json
from datetime import date, timedelta

import pytest

from dashboard.services import certificates as ct


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(ct, '_db', lambda: d)
    monkeypatch.setattr(ct, '_table_ready', False)
    monkeypatch.setenv('EAPPROVAL_EXECUTOR_DM_IDS', 'UHANDLER')
    return d


VALS = {'ctype': '재직증명서', 'purpose': '금융기관 제출', 'dest': '○○은행', 'ssn': ct.SSN_OPTS[0],
        'receive': ct.RECEIVE[0], 'copies': '', 'need_by': '', 'memo': ''}


def _state(**kw):
    v = {}
    for k, val in kw.items():
        if k in ('c_type', 'c_purpose', 'c_ssn', 'c_receive'):
            v[k] = {k: {'selected_option': {'value': val}}}
        elif k == 'c_need':
            v[k] = {k: {'selected_date': val}}
        else:
            v[k] = {k: {'value': val}}
    return {'state': {'values': v}}


def test_view_and_read():
    view = ct.request_view('박민우')
    ids = [b.get('block_id') for b in view['blocks'] if b.get('block_id')]
    assert ids == ['c_type', 'c_purpose', 'c_dest', 'c_ssn', 'c_receive', 'c_copies', 'c_need', 'c_memo']
    vals = ct.read_request(_state(c_type='경력증명서', c_purpose='관공서 제출', c_dest=' 구청 ',
                                  c_ssn=ct.SSN_OPTS[1], c_receive=ct.RECEIVE[1], c_copies='2'))
    assert vals['ctype'] == '경력증명서' and vals['dest'] == '구청' and vals['copies'] == '2'
    assert ct.validate(vals) == {}
    assert 'c_copies' in ct.validate(dict(vals, copies='0'))
    assert 'c_need' in ct.validate(dict(vals, need_by=(date.today() - timedelta(days=1)).isoformat()))


def test_flow_issue_once_and_cards(db):
    row = ct.create('UREQ', '박민우', VALS)
    assert row['status'] == ct.PENDING and row['copies'] == 1
    h = ct.card_blocks(row, 'handler')
    def act(bl):
        return next(b for b in bl if b['type'] == 'actions')
    assert [e['action_id'] for e in act(h)['elements']] == [ct.A_ISSUE, ct.A_REJECT]
    assert [e['action_id'] for e in act(ct.card_blocks(row, 'requester'))['elements']] == [ct.A_CANCEL]
    assert '○○은행' in ct.card_text(row) and '부' not in ct.card_text(row).split('받는 방법')[1].split('\n')[0]
    done = ct.finish(row['id'], ct.ISSUED, '황샛별', '', ['재직증명서.pdf'])
    assert done['status'] == ct.ISSUED and not any(b['type'] == 'actions' for b in ct.card_blocks(done, 'handler'))
    assert ct.finish(row['id'], ct.REJECTED, '황샛별', 'x') is None                     # 두 번 처리 안 됨
    assert '발급 완료' in ct.card_text(done)
    assert ct.mine('UREQ')[0]['id'] == row['id']


def test_original_copy_shows_copies_and_optional_file(db):
    row = ct.create('UREQ', '박민우', dict(VALS, receive=ct.RECEIVE[1], copies='2'))
    assert '2부' in ct.card_text(row)
    fb = next(b for b in ct.issue_view(row)['blocks'] if b.get('block_id') == ct.FILE_BLOCK)
    assert fb['optional'] is True            # 원본만 줄 땐 파일 없이 완료
    pdf = ct.create('UREQ', '박민우', VALS)
    assert next(b for b in ct.issue_view(pdf)['blocks'] if b.get('block_id') == ct.FILE_BLOCK)['optional'] is False


def test_handlers_follow_eapproval_dm_setting(db, monkeypatch):
    assert ct.handler_ids() == ('UHANDLER',)
    assert ct.can_handle('UHANDLER') and not ct.can_handle('UOTHER')


def test_issue_sends_pdf_to_requester(db, monkeypatch):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    row = ct.create('UREQ', '박민우', VALS)
    calls = []

    class C:
        def conversations_open(self, users):
            return {'channel': {'id': 'D' + users}}

        def files_upload_v2(self, **kw):
            calls.append(('upload', kw['channel'], [u['filename'] for u in kw['file_uploads']]))

        def chat_update(self, **kw):
            calls.append(('update', kw['channel']))

        def chat_postMessage(self, **kw):
            calls.append(('post', kw['channel'], kw.get('text')))
            return {'ts': '1'}

    monkeypatch.setattr(se, '_download', lambda url: b'%PDF' if url else None)
    monkeypatch.setattr(se, '_name', lambda uid: '황샛별')
    assert sc.issue(C(), row['id'], 'UHANDLER', [{'name': 'a.pdf', 'url': 'u'}], '') == ''
    assert ('upload', 'DUREQ', ['a.pdf']) in calls
    assert ct.get(row['id'])['status'] == ct.ISSUED
    assert sc.issue(C(), row['id'], 'UHANDLER', [], '') == '이미 처리된 신청이에요.'
    # 파일을 못 보내면 발급 완료로 바꾸지 않는다
    row2 = ct.create('UREQ', '박민우', VALS)
    monkeypatch.setattr(se, '_download', lambda url: None)
    assert '보내지 못했어요' in sc.issue(C(), row2['id'], 'UHANDLER', [{'name': 'b.pdf', 'url': 'u'}], '')
    assert ct.get(row2['id'])['status'] == ct.PENDING


def test_menu_has_cert_button():
    from dashboard.services import e_approval as ea
    hub = json.dumps(ea.hub_blocks(), ensure_ascii=False)
    menu = json.dumps(ea.menu_blocks(), ensure_ascii=False)
    assert ct.A_OPEN in hub and ct.A_OPEN in menu and '사직' not in hub and '휴직' not in hub
