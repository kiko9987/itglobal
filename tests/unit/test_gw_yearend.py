"""연말정산 자료 모으기(2026-11-10) — 두 방식(일괄제공 동의·PDF)·제출·보완·명단·묶음·결과·권한."""
import io
import zipfile
from datetime import date

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import gw_login
from dashboard.services import hr
from dashboard.services import yearend as ye

ADMIN = 'kiko@itg-aircon.com'
STAFF = 'mw@itg-aircon.com'
Y = 2026


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, gw_login, ye):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('GW_HR_DOCS_DIR', str(tmp_path / 'docs'))
    sent = []
    monkeypatch.setattr(ye, '_dm', lambda sid, text: sent.append((sid, text)) or bool(sid))
    gw._invalidate()
    hr.upsert_from_flex({'flex_id': 'A', 'name': '박민우', 'status': '재직', 'work_email': STAFF, 'login_email': STAFF,
                         'slack_id': 'UMW', 'ssn': '9001011234567'}, [])
    hr.upsert_from_flex({'flex_id': 'B', 'name': '김현장', 'status': '재직', 'access': 'field', 'slack_id': 'UF',
                         'ssn': '8505052234567'}, [])
    hr.upsert_from_flex({'flex_id': 'C', 'name': '이퇴사', 'status': '퇴사'}, [])
    d.sent = sent
    yield d
    gw._invalidate()


def test_current_year():
    assert ye.current_year(date(2026, 11, 10)) == 2026
    assert ye.current_year(date(2027, 2, 1)) == 2026
    assert ye.current_year(date(2027, 4, 1)) == 2027


def test_flow_bulk_and_pdf(db):
    me = hr.find_by_email(STAFF)
    with pytest.raises(PermissionError):
        ye.save_answers(me['id'], Y, {'method': 'bulk'})            # 준비 중엔 못 냄
    ye.set_season(Y, ADMIN, status='open', due_date='2027-01-31')
    assert ye.problems(ye.my_view(me, Y))[0].startswith('간소화 자료를 어떻게')

    ye.save_answers(me['id'], Y, {'method': 'bulk', 'family_status': 'none'})
    assert ye.problems(ye.my_view(me, Y)) == ['홈택스에서 회사 자료 제공에 동의한 뒤 「동의했어요」를 체크해 주세요.']
    ye.save_answers(me['id'], Y, {'bulk_consent': True})
    assert ye.submit(me['id'], Y, me)['status'] == ye.SUBMITTED
    with pytest.raises(PermissionError):
        ye.save_answers(me['id'], Y, {'note': 'x'})                 # 제출 뒤엔 수정하기 먼저
    ye.reopen(me['id'], Y)

    ye.save_answers(me['id'], Y, {'method': 'pdf', 'family_status': 'changed'})
    p = ye.problems(ye.my_view(me, Y))
    assert '홈택스 간소화 자료 PDF를 업로드해 주세요.' in p and '바뀐 부양가족 내용을 적어 주세요.' in p
    ye.upload(me['id'], Y, 'simple', '간소화.pdf', b'%PDF', f'person:{me["id"]}')
    ye.upload(me['id'], Y, 'extra', '안경.jpg', b'JPG', f'person:{me["id"]}')
    ye.save_answers(me['id'], Y, {'family_note': '자녀 출생'})
    v = ye.my_view(me, Y)
    assert len(v['simple_docs']) == 1 and len(v['extra_docs']) == 1 and ye.problems(v) == []
    with pytest.raises(ValueError):
        ye.upload(me['id'], Y, 'result', 'x.pdf', b'%PDF', 'x')      # 결과는 본인이 못 올림
    ye.submit(me['id'], Y, me)


def test_admin_return_done_roster_bundle_result(db):
    me = hr.find_by_email(STAFF)
    ye.set_season(Y, ADMIN, status='open')
    ye.save_answers(me['id'], Y, {'method': 'pdf', 'family_status': 'none'})
    ye.upload(me['id'], Y, 'simple', 'a.pdf', b'%PDF-A', 'x')
    ye.submit(me['id'], Y, me)
    with pytest.raises(PermissionError):
        ye.set_status(me['id'], Y, ye.DONE, STAFF)
    with pytest.raises(ValueError):
        ye.set_status(me['id'], Y, ye.RETURNED, ADMIN, '')
    assert ye.set_status(me['id'], Y, ye.RETURNED, ADMIN, '등본 누락')['status'] == ye.RETURNED
    ye.submit(me['id'], Y, me)                                       # 보완 뒤 다시 제출
    assert ye.set_status(me['id'], Y, ye.DONE, ADMIN)['status'] == ye.DONE

    rows = {r['name']: r for r in ye.admin_rows(Y)}
    assert set(rows) == {'박민우', '김현장'}                          # 퇴사자 제외
    assert rows['박민우']['simple'] == 1 and rows['김현장']['status'] == ye.DRAFT

    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(ye.roster_xlsx(Y, ADMIN)))
    vals = [tuple(r) for r in wb.active.iter_rows(values_only=True)]
    assert ('박민우', '900101-1234567') == vals[1][:2] or ('박민우', '900101-1234567') in [v[:2] for v in vals]
    with pytest.raises(PermissionError):
        ye.roster_xlsx(Y, STAFF)

    z = zipfile.ZipFile(io.BytesIO(ye.bundle_zip(Y, ADMIN)))
    names = z.namelist()
    assert any(n.startswith(f'박민우_{me["id"]}/간소화_') for n in names) and f'{Y}년귀속_연말정산_요약.csv' in names
    actions = [l['action'] for l in hr.access_log(me['id'])]
    assert 'ye_returned' in actions and 'ye_done' in actions

    r = ye.upload_result(me['id'], Y, '원천징수영수증.pdf', b'%PDF-R', ADMIN)
    assert not r['skipped'] and ye.my_view(me, Y)['result_docs'][0]['file_name'] == '원천징수영수증.pdf'
    assert ye.notify_result(me['id'], Y) and db.sent[-1][0] == 'UMW'


def test_notify_pending(db):
    ye.set_season(Y, ADMIN, status='open', due_date='2027-01-31')
    out = ye.notify_pending(Y, ADMIN)
    assert sorted(out['sent']) == ['김현장', '박민우'] and '1월 31일' in db.sent[0][1]


@pytest.fixture
def client(db, monkeypatch):
    import dashboard.auth as auth
    from flask import Blueprint, session
    real = auth._validate_session
    monkeypatch.setattr(auth, '_validate_session',
                        lambda require_admin=False: real(require_admin) if session.get('user', {}).get('gw_field') else None)
    from dashboard.blueprints.groupware import groupware_bp
    from dashboard.blueprints.hr import hr_bp
    app = Flask('t')
    app.secret_key = 't'
    app.register_blueprint(groupware_bp)
    app.register_blueprint(hr_bp)
    a = Blueprint('auth', __name__)
    a.add_url_rule('/login', 'login_page', lambda: 'login')
    app.register_blueprint(a)
    gw.set_open_all(True, ADMIN)
    ye.set_season(Y, ADMIN, status='open')
    return app.test_client()


def test_api_field_staff_admin(client, monkeypatch):
    monkeypatch.setattr(ye, 'current_year', lambda today=None: Y)
    # 현장 직원: 슬랙 링크로 들어와 휴대폰으로 업로드·제출
    tok = gw_login.issue('UF').rsplit('/', 1)[1]
    client.get(f'/gw/login/{tok}')
    assert client.post('/api/yearend/me', json={'method': 'pdf', 'family_status': 'none'}).status_code == 200
    up = client.post('/api/yearend/me/documents', content_type='multipart/form-data',
                     data={'kind': 'simple', 'file': (io.BytesIO(b'%PDF'), 'g.pdf')}).get_json()
    assert up['ye']['simple_docs'] and up['ye']['problems'] == []
    assert client.post('/api/yearend/me/submit').get_json()['ye']['submission']['status'] == 'submitted'
    assert client.get('/api/yearend/admin').status_code == 403                      # 현황판 X
    assert client.get('/api/yearend/admin/roster.xlsx').status_code == 403

    with client.session_transaction() as s:
        s.clear()
        s['user'] = {'email': STAFF, 'name': '박민우', 'permission_level': 'editor'}
    assert client.get('/api/yearend/me').get_json()['ye']['name'] == '박민우'
    assert client.get('/api/yearend/admin').status_code == 403
    assert client.post('/api/yearend/admin/season', json={'status': 'closed'}).status_code == 403

    with client.session_transaction() as s:
        s['user'] = {'email': ADMIN, 'name': '고광일', 'permission_level': 'super_admin'}
    j = client.get('/api/yearend/admin').get_json()
    assert j['season']['status'] == 'open' and {r['name'] for r in j['rows']} == {'박민우', '김현장'}
    r = client.get('/api/yearend/admin/bundle.zip')
    assert r.status_code == 200 and r.headers['Content-Type'] == 'application/zip'
