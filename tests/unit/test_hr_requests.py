"""내 정보 직접 고치기 (2026-11-20) — 연락처·비상연락망·서류 올리기는 본인이 바로, 나머지는 경영지원만."""
import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import hr_requests as hq

STAFF = 'staff@example.com'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, hq):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('GW_HR_DOCS_DIR', str(tmp_path / 'docs'))
    gw._invalidate()
    yield d
    gw._invalidate()


def _me():
    hr.upsert_from_flex({'flex_id': 'F1', 'name': '박민우', 'status': '재직', 'hire_date': '2022-02-03', 'access': 'staff',
                         'work_email': STAFF, 'login_email': STAFF, 'phone': '010-1111-2222', 'address': '김포',
                         'department': '영업'}, [])
    p = hr.find_by_email(STAFF)
    with hr._conn() as conn:
        conn.execute("INSERT INTO hr_families (person_id, name, relation, phone, source) VALUES (?,?,?,?,'flex')",
                     (p['id'], '이상철', '배우자', '010-3333-4444'))
    return hr.get_person(p['id'])


def test_contact_updates_right_away_and_logged(db):
    p = _me()
    with pytest.raises(ValueError, match='바뀐 내용'):
        hq.update_info(p, {'phone': '010-1111-2222', 'address': '김포'}, STAFF)
    with pytest.raises(ValueError, match='전화'):
        hq.update_info(p, {'phone': '123'}, STAFF)
    row = hq.update_info(p, {'phone': '010-5555-6666', 'address': '김포 걸포동', 'department': '경영지원'}, STAFF)
    now = hr.get_person(p['id'])
    assert now['phone'] == '010-5555-6666' and now['address'] == '김포 걸포동'
    assert now['department'] == '영업'                                               # 인사 정보는 못 바꿈
    assert row['status'] == hq.APPLIED and row['payload']['old']['phone'] == '010-1111-2222'
    assert any(l['action'] == 'self_edit' for l in hr.access_log(p['id']))
    # 사람이 고친 카드 — 플렉스 다시 가져오기가 덮어쓰지 않는다
    hr.upsert_from_flex({'flex_id': 'F1', 'name': '박민우', 'phone': '010-1111-2222'}, [])
    assert hr.get_person(p['id'])['phone'] == '010-5555-6666'


def test_family_add_edit_delete(db):
    p = _me()
    old = hr.families(p['id'])[0]
    with pytest.raises(ValueError, match='이름과 관계'):
        hq.update_family(p, [{'name': '김엄마', 'relation': '', 'phone': ''}], STAFF)
    hq.update_family(p, [{'id': old['id'], 'name': '이상철', 'relation': '배우자', 'phone': '010-9999-0000'},
                         {'name': '박아들', 'relation': '자녀', 'phone': '010-1234-5678'}], STAFF)
    fams = {f['name']: f for f in hr.families(p['id'])}
    assert fams['이상철']['phone'] == '010-9999-0000' and fams['박아들']['source'] == 'self'
    hq.update_family(p, [{'id': fams['박아들']['id'], 'name': '박아들', 'relation': '자녀', 'phone': '010-1234-5678'}], STAFF)
    assert [f['name'] for f in hr.families(p['id'])] == ['박아들']
    with pytest.raises(ValueError, match='없는 줄'):
        hq.update_family(p, [{'id': 99999, 'name': 'x', 'relation': 'y'}], STAFF)


def test_doc_upload_goes_straight_in(db):
    p = _me()
    with pytest.raises(ValueError):
        hq.upload_doc(p, '주민등록등본', 'a.exe', b'x', STAFF)
    with pytest.raises(ValueError):
        hq.upload_doc(p, '근로계약서', 'a.pdf', b'x', STAFF)                     # 본인이 올릴 종류가 아님
    hq.upload_doc(p, '주민등록등본', '등본.pdf', b'%PDF-1', STAFF)
    docs = hr.documents(p['id'])
    assert [d['doc_type'] for d in docs] == ['주민등록등본'] and docs[0]['uploaded_by'] == STAFF
    with pytest.raises(PermissionError):
        hr.delete_document(docs[0]['id'], STAFF)                                     # 지우기는 경영지원만
    assert [r['kind'] for r in hq.history(p['id'])] == ['doc']


def test_api_and_field_gate(db, monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints import hr as hr_bp_mod
    from dashboard.blueprints.groupware import field_path_allowed
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    sent = []
    monkeypatch.setattr(hr_bp_mod, '_dm', lambda uids, text: sent.append(text))
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    p = _me()
    other = hr.find_by_email(STAFF)
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(hr_bp_mod.hr_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': STAFF, 'permission_level': 'user'}
        card = cl.get(f"/api/hr/people/{p['id']}").get_json()['person']
        assert card['can_self_edit'] and '주민등록등본' in card['self_doc_types']
        assert cl.post('/api/hr/me/contact', json={'phone': '010-2222-3333'}).get_json()['success']
        assert hr.get_person(p['id'])['phone'] == '010-2222-3333' and '연락처' in sent[-1]
        r = cl.post('/api/hr/me/family', json={'rows': [{'name': '김엄마', 'relation': '모', 'phone': '01022223333'}]})
        assert r.get_json()['success'] and [f['name'] for f in hr.families(p['id'])] == ['김엄마']
        assert cl.post('/api/hr/me/department', json={}).status_code == 404
        assert cl.post('/api/hr/me/contact', json={'phone': '010-2222-3333'}).status_code == 400   # 그대로면 거절
        assert cl.put(f"/api/hr/people/{p['id']}", json={'department': '경영지원'}).status_code != 200
    assert other['id'] == p['id']
    assert field_path_allowed('POST', '/api/hr/me/contact')
    assert field_path_allowed('POST', '/api/hr/me/documents')
    assert not field_path_allowed('POST', '/api/hr/me/department')


def test_admin_can_edit_family_on_card(db, monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints import hr as hr_bp_mod
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    monkeypatch.setattr(hr, 'hr_admins', lambda: ['hr-admin@example.com'])
    p = _me()
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(hr_bp_mod.hr_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': STAFF, 'permission_level': 'user'}
        assert cl.post(f"/api/hr/people/{p['id']}/family", json={'rows': []}).status_code == 403
        with cl.session_transaction() as s:
            s['user'] = {'email': 'hr-admin@example.com', 'permission_level': 'admin'}
        r = cl.post(f"/api/hr/people/{p['id']}/family", json={'rows': [{'name': '김엄마', 'relation': '모', 'phone': '010-1111-0000'}]})
        assert r.get_json()['success'] and [f['name'] for f in hr.families(p['id'])] == ['김엄마']
    assert any(l['action'] == 'edit' and l['detail'] == 'families' for l in hr.access_log(p['id']))
    assert hq.history(p['id']) == []                                                 # 본인 수정 기록엔 안 남김
