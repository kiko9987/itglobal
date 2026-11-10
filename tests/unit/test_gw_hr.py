"""그룹웨어 인사(2026-11-10) — 인사카드 권한·주민번호 암호화·수정·플렉스 이관 반복."""
import re

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import hr

ADMIN = 'kiko@itg-aircon.com'      # 경영지원(기본 HR 관리자) + 첫 시범
STAFF = 'mw@itg-aircon.com'
JW = 'jw@itg-aircon.com'           # 휴가 결재자
SSN = '900101-1234567'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    gw._invalidate()
    yield d
    gw._invalidate()


def _flex(fid='F1', name='박민우', **kw):
    rec = {'flex_id': fid, 'name': name, 'status': '재직', 'hire_date': '2022-02-03', 'department': '영업',
           'work_email': STAFF, 'login_email': STAFF, 'access': 'staff', 'phone': '01012345678', 'ssn': SSN}
    rec.update(kw)
    return rec


def test_ssn_encrypted_masked_and_reveal_logged(db):
    hr.upsert_from_flex(_flex(), [{'apply_date': '2022-02-03', 'department': '영업'}])
    p = hr.find_by_email(STAFF)
    raw = db._get_connection
    with raw() as conn:
        enc = conn.execute('SELECT ssn_enc FROM hr_people WHERE id = ?', (p['id'],)).fetchone()[0]
    assert enc and not re.search(r'\d{6}-?\d{7}', enc)            # 평문 아님
    c = hr.card(p['id'], STAFF)
    assert c['ssn_masked'] == '900101-1******' and 'ssn_enc' not in c
    assert not c['can_reveal_ssn'] and not c['can_edit']
    with pytest.raises(PermissionError):
        hr.reveal_ssn(p['id'], STAFF)
    assert hr.reveal_ssn(p['id'], ADMIN) == SSN
    assert [l['action'] for l in hr.access_log(p['id'])] == ['ssn_view']


def test_no_key_refuses_plain_storage(db, monkeypatch):
    monkeypatch.delenv('GW_HR_KEY')
    with pytest.raises(hr.SsnKeyMissing):
        hr.create_person({'name': '새직원', 'ssn': SSN}, ADMIN)


def test_update_only_admin_and_validation(db):
    hr.upsert_from_flex(_flex(), [])
    pid = hr.find_by_email(STAFF)['id']
    with pytest.raises(PermissionError):
        hr.update_person(pid, {'memo': 'x'}, STAFF)
    with pytest.raises(PermissionError):
        hr.update_person(pid, {'memo': 'x'}, JW)                   # 결재자도 인사카드 수정 불가
    with pytest.raises(ValueError):
        hr.update_person(pid, {'hire_date': '2022-13-40'}, ADMIN)
    with pytest.raises(ValueError):
        hr.update_person(pid, {'status': '휴가'}, ADMIN)
    with pytest.raises(ValueError):
        hr.update_person(pid, {'ssn': '12345'}, ADMIN)
    c = hr.update_person(pid, {'job_rank': '대리', 'ssn': '9001012234567'}, ADMIN)
    assert c['job_rank'] == '대리' and c['ssn_masked'] == '900101-2******'
    assert hr.access_log(pid)[0]['detail'] == 'job_rank,ssn_enc'   # 무엇을 바꿨는지만(값은 안 남김)


def test_reimport_keeps_human_edits(db):
    hr.upsert_from_flex(_flex(), [{'apply_date': '2022-02-03'}])
    pid = hr.find_by_email(STAFF)['id']
    hr.update_person(pid, {'department': '기술영업'}, ADMIN)
    assert hr.upsert_from_flex(_flex(department='영업', job_role='영업'), [{'apply_date': '2022-02-03'}]) == 'update'
    p = hr.get_person(pid)
    assert p['department'] == '기술영업'        # 사람이 고친 칸은 그대로
    assert p['job_role'] == '영업'              # 비어 있던 칸은 채움
    assert len(hr.appointments(pid)) == 1       # 발령 이력은 다시 써도 한 벌


def test_find_by_email_skips_resigned(db):
    hr.upsert_from_flex(_flex('F1', '이종원', status='퇴사', work_email=None, login_email=None), [])
    hr.upsert_from_flex(_flex('F2', '이종원', work_email='ljw@itg-aircon.com', login_email='ljw@itg-aircon.com'), [])
    assert hr.find_by_email('ljw@itg-aircon.com')['flex_id'] == 'F2'
    assert hr.find_by_email('nobody@itg-aircon.com') is None


def test_roles_default_and_setting(db):
    assert hr.is_hr_admin(ADMIN) and not hr.is_hr_admin(JW)
    assert hr.can_view_all_leave(JW) and not hr.can_view_all_leave(STAFF)
    gw.set_setting('hr_admins', '["yg@itg-aircon.com"]', ADMIN)
    assert hr.is_hr_admin('yg@itg-aircon.com') and not hr.is_hr_admin(ADMIN)


# ── API 권한 ───────────────────────────────────────────────────────────
@pytest.fixture
def client(db, monkeypatch):
    import dashboard.auth as auth
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    from dashboard.blueprints.groupware import groupware_bp
    from dashboard.blueprints.hr import hr_bp
    app = Flask('t', template_folder='dashboard/templates')
    app.secret_key = 't'
    app.register_blueprint(groupware_bp)
    app.register_blueprint(hr_bp)
    hr.upsert_from_flex(_flex(), [])
    hr.upsert_from_flex(_flex('F9', '박정우', work_email=JW, login_email=JW), [])
    return app.test_client()


def _as(c, email, perm='editor'):
    with c.session_transaction() as s:
        s['user'] = {'email': email, 'permission_level': perm}


def test_api_permissions(client):
    staff_id = hr.find_by_email(STAFF)['id']
    jw_id = hr.find_by_email(JW)['id']

    _as(client, STAFF)                                   # 시범 명단 밖 → 전부 막힘
    assert client.get('/api/hr/people').status_code == 403
    assert client.get(f'/api/hr/people/{staff_id}').status_code == 403
    assert client.get('/hr/me').status_code == 302

    gw.add_pilot(STAFF, ADMIN)
    gw.add_pilot(JW, ADMIN)
    assert client.get('/api/hr/people').status_code == 403               # 목록은 HR 관리자만
    me = client.get(f'/api/hr/people/{staff_id}').get_json()
    assert me['success'] and 'access_log' not in me['person'] and 'ssn_enc' not in str(me)
    assert client.get(f'/api/hr/people/{jw_id}').status_code == 403      # 남의 카드
    assert client.post(f'/api/hr/people/{staff_id}/ssn').status_code == 403
    assert client.put(f'/api/hr/people/{staff_id}', json={'memo': 'x'}).status_code == 403
    assert client.get('/api/gw/roles').status_code == 403
    assert client.get('/hr/people').status_code == 302                   # → 내 정보

    _as(client, JW, 'admin')                             # 휴가 결재자: 인사카드 못 봄
    assert client.get('/api/hr/people').status_code == 403
    assert client.get(f'/api/hr/people/{staff_id}').status_code == 403

    _as(client, ADMIN, 'super_admin')
    lst = client.get('/api/hr/people').get_json()['people']
    assert {p['name'] for p in lst} == {'박민우', '박정우'}
    assert client.post(f'/api/hr/people/{staff_id}/ssn').get_json()['ssn'] == SSN
    r = client.put(f'/api/hr/people/{staff_id}', json={'phone': '010-9999-0000'}).get_json()
    assert r['success'] and r['person']['phone'] == '010-9999-0000'
    assert client.post('/api/gw/roles', json={'key': 'hr_admins', 'emails': []}).status_code == 400
    assert client.post('/api/gw/roles', json={'key': 'hr_admins', 'emails': [ADMIN, 'Yg@itg-aircon.com']}).get_json()['hr_admins'] == \
        [ADMIN, 'yg@itg-aircon.com']
