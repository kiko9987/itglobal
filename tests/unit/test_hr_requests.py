"""본인 정보 변경 요청 (2026-11-20) — 연락처·비상연락망·서류 요청 → 반영·반려·취소, API·현장 문지기·홈 할 일."""
import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import hr_requests as hq

ADMIN = 'kiko@itg-aircon.com'
STAFF = 'mw@itg-aircon.com'


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
                         'work_email': STAFF, 'login_email': STAFF, 'phone': '010-1111-2222', 'address': '김포'}, [])
    p = hr.find_by_email(STAFF)
    with hr._conn() as conn:
        conn.execute("INSERT INTO hr_families (person_id, name, relation, phone, source) VALUES (?,?,?,?,'flex')",
                     (p['id'], '이상철', '배우자', '010-3333-4444'))
    return hr.get_person(p['id'])


def test_info_request_approve(db):
    p = _me()
    with pytest.raises(ValueError, match='바뀐 내용'):
        hq.request_info(p, {'phone': '01011112222', 'address': '김포'}, STAFF)     # 숫자만 같으면 같은 번호
    with pytest.raises(ValueError, match='전화'):
        hq.request_info(p, {'phone': '123'}, STAFF)
    r = hq.request_info(p, {'phone': '010-5555-6666', 'address': '김포 걸포동', 'personal_email': ''}, STAFF)
    assert r['payload']['new'] == {'phone': '010-5555-6666', 'address': '김포 걸포동'}
    with pytest.raises(ValueError, match='확인 중'):
        hq.request_info(p, {'phone': '010-7777-8888'}, STAFF)
    with pytest.raises(PermissionError):
        hq.approve(r['id'], STAFF)
    assert hq.approve(r['id'], ADMIN)['status'] == hq.APPROVED
    assert hr.get_person(p['id'])['phone'] == '010-5555-6666' and hr.get_person(p['id'])['address'] == '김포 걸포동'
    with pytest.raises(ValueError):
        hq.approve(r['id'], ADMIN)


def test_family_request_add_edit_delete(db):
    p = _me()
    old = hr.families(p['id'])[0]
    with pytest.raises(ValueError, match='이름과 관계'):
        hq.request_family(p, [{'name': '김엄마', 'relation': '', 'phone': ''}], STAFF)
    r = hq.request_family(p, [{'id': old['id'], 'name': '이상철', 'relation': '배우자', 'phone': '010-9999-0000'},
                              {'name': '박아들', 'relation': '자녀', 'phone': '010-1234-5678'}], STAFF)
    hq.approve(r['id'], ADMIN)
    fams = {f['name']: f for f in hr.families(p['id'])}
    assert fams['이상철']['phone'] == '010-9999-0000' and fams['박아들']['source'] == 'self'
    r2 = hq.request_family(p, [{'id': fams['박아들']['id'], 'name': '박아들', 'relation': '자녀', 'phone': '010-1234-5678'}], STAFF)
    hq.approve(r2['id'], ADMIN)
    assert [f['name'] for f in hr.families(p['id'])] == ['박아들']                  # 빠진 사람은 지움
    with pytest.raises(ValueError, match='없는 줄'):
        hq.request_family(p, [{'id': 99999, 'name': 'x', 'relation': 'y'}], STAFF)


def test_doc_request_staged_until_approved_and_cancel_reject(db):
    p = _me()
    with pytest.raises(ValueError):
        hq.request_doc(p, '주민등록등본', 'a.exe', b'x', STAFF)
    r = hq.request_doc(p, '주민등록등본', '등본.pdf', b'%PDF-1', STAFF)
    path = hq.staged_path(r['id'])
    assert path and hr.documents(p['id']) == []                                      # 반영 전엔 서류에 없음
    hq.approve(r['id'], ADMIN)
    docs = hr.documents(p['id'])
    assert [d['doc_type'] for d in docs] == ['주민등록등본'] and docs[0]['uploaded_by'] == STAFF
    import os
    assert not os.path.exists(path)
    r2 = hq.request_doc(p, '급여계좌사본', '통장.jpg', b'jpg', STAFF)
    p2 = hq.staged_path(r2['id'])
    hq.cancel(r2['id'], p['id'], STAFF)
    assert not os.path.exists(p2) and hq.get(r2['id'])['status'] == hq.CANCELLED
    r3 = hq.request_info(p, {'address': '서울'}, STAFF)
    with pytest.raises(ValueError, match='사유'):
        hq.reject(r3['id'], ADMIN, ' ')
    assert hq.reject(r3['id'], ADMIN, '주소를 도로명으로 적어 주세요')['status'] == hq.REJECTED
    assert [x['status'] for x in hq.mine(p['id'])] == ['rejected', 'approved']      # 취소는 안 보임


def test_api_and_field_gate_and_home(db, monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints import hr as hr_bp_mod
    from dashboard.blueprints.groupware import field_path_allowed
    from dashboard.services import gw_home
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(hr_bp_mod, '_dm', lambda uids, text: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    p = _me()
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(hr_bp_mod.hr_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': STAFF, 'permission_level': 'user'}
        card = cl.get(f"/api/hr/people/{p['id']}").get_json()['person']
        assert card['my_requests'] == [] and '주민등록등본' in card['self_doc_types'] and 'change_requests' not in card
        r = cl.post('/api/hr/me/requests/info', json={'phone': '010-2222-3333'}).get_json()
        assert r['success'] and r['requests'][0]['status'] == 'pending'
        assert cl.post('/api/hr/me/requests/info', json={'phone': '010-2222-4444'}).status_code == 400
        rid = r['requests'][0]['id']
        assert cl.post(f'/api/hr/requests/{rid}/approve').status_code == 403            # 직원은 반영 못 함
        items = [i for i in gw_home.todos(None, '', False, True) if i['id'] == 'hr-requests']
        assert items and '1건' in items[0]['title']
        with cl.session_transaction() as s:
            s['user'] = {'email': ADMIN, 'permission_level': 'admin'}
        card = cl.get(f"/api/hr/people/{p['id']}").get_json()['person']
        assert [x['id'] for x in card['change_requests']] == [rid]
        assert cl.post(f'/api/hr/requests/{rid}/reject', json={}).status_code == 400
        assert cl.post(f'/api/hr/requests/{rid}/approve').get_json()['success']
    assert hr.get_person(p['id'])['phone'] == '010-2222-3333'
    assert field_path_allowed('POST', '/api/hr/me/requests/family')
    assert field_path_allowed('POST', '/api/hr/me/requests/12/cancel')
    assert not field_path_allowed('POST', '/api/hr/requests/12/approve')
