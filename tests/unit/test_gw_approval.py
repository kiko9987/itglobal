"""그룹웨어 결재함·플렉스 보관함(2026-11-10) — 권한·상자·처리 단추(슬랙과 같은 함수)·보관함 가져오기."""
import json

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import e_approval as ea
from dashboard.services import groupware as gw
from dashboard.services import gw_approval as ga
from dashboard.services import gw_login
from dashboard.services import hr

KIKO, JW, MW = 'U04UL2ZLJAX', 'U051ENK8AAY', 'UMW'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, gw_login, ea):
        monkeypatch.setattr(mod, '_db', lambda: d)
        if hasattr(mod, '_table_ready'):
            monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setattr(ga, '_db', lambda: d)
    monkeypatch.setattr(ga, '_archive_ready', False)
    # 슬랙 신원 확인(2026-11-13) — 테스트에선 슬랙 디렉터리 = 인사카드 slack_id 로 흉내
    from dashboard.services import gw_approval as _ga
    _ga._SID_CACHE.clear()
    monkeypatch.setattr(_ga, '_lookup_slack_id', lambda email: (hr.find_by_email(email) or {}).get('slack_id') or '')
    for k in ('EAPPROVAL_EXECUTOR_DM_IDS', 'EAPPROVAL_APPROVER_IDS', 'EAPPROVAL_FINAL_APPROVER_IDS',
              'EAPPROVAL_N_EXECUTOR_IDS', 'SLACK_SETTLEMENT_CHECKER_ID', 'SLACK_SETTLEMENT_CHECKER_EMAIL',
              'EAPPROVAL_FINAL_ENABLED'):
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    gw._invalidate()
    for fid, name, email, sid in (('K', '고광일', 'kiko@itg-aircon.com', KIKO), ('M', '박민우', 'mw@itg-aircon.com', MW),
                                  ('O', '김다른', 'other@itg-aircon.com', 'UOTHER'), ('N', '이연결없음', 'none@itg-aircon.com', '')):
        hr.upsert_from_flex({'flex_id': fid, 'name': name, 'status': '재직', 'login_email': email, 'work_email': email,
                             'slack_id': sid}, [])
    yield d
    gw._invalidate()


def _approval(d, status, requester, name='박민우', form='fund', amount=550000, extra=None, account='110-123-456789'):
    with d._get_connection() as conn:
        ea._ensure_tables(conn)
        cur = conn.execute('INSERT INTO approvals (form, requester_id, requester_name, status, amount, grn, partner, account, '
                           "project_code, content, card_channel, card_ts, extra) VALUES (?,?,?,?,?, 'G', '알만상사', ?, 'G4266-MJ', "
                           "'배관 자재', 'C1', '1700000000.123456', ?)",
                           (form, requester, name, status, amount, account, json.dumps(extra or {}, ensure_ascii=False)))
        conn.commit()
        return cur.lastrowid


def test_viewer_boxes_and_permissions(db):
    mine = _approval(db, ea.PENDING, MW)
    other = _approval(db, ea.PENDING, 'UOTHER', '김다른')
    k, m, o = ga.viewer('kiko@itg-aircon.com'), ga.viewer('mw@itg-aircon.com'), ga.viewer('other@itg-aircon.com')
    assert k['see_all'] and not m['see_all'] and m['linked']
    assert not ga.viewer('none@itg-aircon.com')['linked']

    kl = ga.list_docs(k, box='todo')
    assert [i['id'] for i in kl['items']] == [other, mine] and kl['counts'] == {'todo': 2, 'mine': 0, 'all': 2}
    assert all(i['my_turn'] for i in kl['items'])
    ml = ga.list_docs(m, box='mine')
    assert [i['id'] for i in ml['items']] == [mine] and 'all' not in ml['counts']
    assert ga.list_docs(m, box='all')['box'] == 'todo'                 # 전체는 결재선 사람만
    assert ga.list_docs(k, box='all', q='알만')['items'] and not ga.list_docs(k, box='all', q='없는말')['items']
    assert ga.list_docs(k, box='all', status='done')['items'] == []

    assert ga.detail(other, m) is None                                 # 남의 결재는 못 봄
    d = ga.detail(mine, m)
    assert d['amount'] == 550000 and any(f['label'] == '계좌번호' and f['value'] == '110-123-456789' for f in d['fields'])
    assert [a['key'] for a in d['actions']] == ['cancel']               # 요청자: 취소만
    kd = ga.detail(mine, k)
    assert [a['key'] for a in kd['actions']] == ['approve', 'reject', 'revise', 'cancel']
    assert kd['slack_link'].endswith('/archives/C1/p1700000000123456')
    assert ga.detail(o and other, o)['actions'] == [{'key': 'cancel', 'label': '요청 취소', 'needs_reason': False, 'primary': False}]


def test_form_fields_and_secret(db):
    aid = _approval(db, ea.PENDING, MW, form='expense', amount=33000,
                    extra={'title': '개인경비', 'f': {'reason': '주차비', 'amount': '33000', 'bank': '신한은행', 'account': '110-1'}})
    k = ga.viewer('kiko@itg-aircon.com')
    labels = {f['label']: f['value'] for f in ga.detail(aid, k)['fields']}
    assert any('계좌' in l for l in labels) and '110-1' in labels.values()


def test_act_calls_same_slack_function(db, monkeypatch):
    from dashboard.blueprints import slack_eapproval as se
    calls = []
    monkeypatch.setattr(se, 'apply_action', lambda a, aid, uid, reason='': calls.append((a, aid, uid, reason)) or None)
    aid = _approval(db, ea.PENDING, MW)
    k = ga.viewer('kiko@itg-aircon.com')
    assert ga.act(aid, 'approve', k) is None and calls[-1] == ('approve', aid, KIKO, '')
    assert ga.act(aid, 'reject', k, '  ') == '사유를 적어 주세요.' and len(calls) == 1
    assert ga.act(aid, 'reject', k, '증빙 없음') is None and calls[-1][3] == '증빙 없음'
    assert ga.act(aid, 'hack', k) == '알 수 없는 처리입니다.'
    assert '슬랙 계정' in ga.act(aid, 'approve', ga.viewer('none@itg-aircon.com'))


def test_flex_archive_import_mask_and_scope(db, tmp_path):
    exp = tmp_path / 'flex_workflow_export_20261101.json'
    exp.write_text(json.dumps({'exported_at': '2026-11-01', 'count': 3, 'docs': [
        {'c': '2026-1001', 't': '자금집행요청(전자계산서용)', 's': 'DONE', 'w': '박민우', 'at': '2026-10-01T01:00:00Z',
         'up': '2026-10-01T02:00:00Z', 'att': 1, 'memo': '자재 대금',
         'f': {'총 금액(VAT포함)': '1,100,000', '프로젝트 코드': 'G4100-MW', '상호명': '알만상사', '계좌번호': '110-1'}},
        {'c': '2026-1002', 't': '2026년 06월 법인카드지출결의서_황샛별 매니저', 's': 'DECLINED', 'w': '황샛별',
         'at': '2026-06-30T01:00:00Z', 'up': '', 'att': 0, 'memo': '',
         'f': {'카드번호': '1234-5678-9012-3456', '청구 금액': '50000'}},
        {'c': '2026-1003', 't': 'R계약서 작성 요청', 's': 'IN_PROGRESS', 'w': '김다른', 'at': '2026-09-01T01:00:00Z',
         'up': '', 'att': 0, 'memo': '', 'f': {}},
    ]}, ensure_ascii=False), encoding='utf-8')
    assert ga.latest_export(str(tmp_path)) == str(exp)
    out = ga.import_flex_export(str(exp))
    assert out['count'] == 3 and out['by_status'] == {'완료': 1, '반려': 1, '진행 중(플렉스)': 1}
    ga.import_flex_export(str(exp))                                          # 다시 돌려도 3건
    k, m = ga.viewer('kiko@itg-aircon.com'), ga.viewer('mw@itg-aircon.com')
    all_ = ga.archive_list(k)
    assert all_['total'] == 3 and set(all_['forms']) == {'자금집행요청(전자계산서용)', '법인카드 지출결의', '계약서 작성 요청'}
    first = [i for i in all_['items'] if i['doc_no'] == '2026-1001'][0]
    assert first['amount'] == 1100000 and first['project_code'] == 'G4100-MW' and first['partner'] == '알만상사'
    card = ga.archive_detail('2026-1002', k)
    assert {'label': '카드번호', 'value': '****-3456'} in card['fields']    # 카드번호 뒤 4자리만
    assert [i['doc_no'] for i in ga.archive_list(m)['items']] == ['2026-1001']   # 본인 것만
    assert ga.archive_detail('2026-1002', m) is None
    assert ga.archive_list(k, q='알만')['total'] == 1 and ga.archive_list(k, form='법인카드 지출결의')['total'] == 1
    assert ga.archive_list(k, date_from='2026-09-01', date_to='2026-10-01')['total'] == 2


@pytest.fixture
def client(db, monkeypatch):
    import dashboard.auth as auth
    from flask import Blueprint, session
    real = auth._validate_session
    monkeypatch.setattr(auth, '_validate_session',
                        lambda require_admin=False: real(require_admin) if session.get('user', {}).get('gw_field') else None)
    from dashboard.blueprints.groupware import groupware_bp
    from dashboard.blueprints.hr import hr_bp
    from dashboard.blueprints.gw_approval import gw_approval_bp
    app = Flask('t')
    app.secret_key = 't'
    for bp in (groupware_bp, hr_bp, gw_approval_bp):
        app.register_blueprint(bp)
    a = Blueprint('auth', __name__)
    a.add_url_rule('/login', 'login_page', lambda: 'login')
    app.register_blueprint(a)
    gw.set_open_all(True, 'kiko@itg-aircon.com')
    hr.upsert_from_flex({'flex_id': 'F', 'name': '김현장', 'status': '재직', 'access': 'field', 'slack_id': 'UF'}, [])
    return app.test_client()


def test_api_list_detail_act_export_and_field_blocked(client, db, monkeypatch):
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, 'apply_action', lambda a, aid, uid, reason='': None)
    aid = _approval(db, ea.PENDING, MW)
    with client.session_transaction() as s:
        s['user'] = {'email': 'kiko@itg-aircon.com', 'name': '고광일', 'permission_level': 'super_admin'}
    j = client.get('/api/approval/list?box=todo').get_json()
    assert j['success'] and j['items'][0]['id'] == aid and j['see_all']
    assert client.get(f'/api/approval/{aid}').get_json()['doc']['actions'][0]['key'] == 'approve'
    r = client.post(f'/api/approval/{aid}/act', json={'action': 'approve'})
    assert r.status_code == 200 and r.get_json()['success']
    assert client.post(f'/api/approval/{aid}/act', json={'action': 'reject'}).status_code == 409   # 사유 없음
    x = client.get('/api/approval/export.xlsx?box=all')
    assert x.status_code == 200 and x.data[:2] == b'PK'

    with client.session_transaction() as s:
        s['user'] = {'email': 'other@itg-aircon.com', 'name': '김다른', 'permission_level': 'editor'}
    assert client.get(f'/api/approval/{aid}').status_code == 404

    tok = gw_login.issue('UF').rsplit('/', 1)[1]
    with client.session_transaction() as s:
        s.clear()
    client.get(f'/gw/login/{tok}')
    assert client.get('/approval').headers['Location'].endswith('/home')            # 현장은 결재함 X
    assert client.get('/api/approval/list').status_code == 403
