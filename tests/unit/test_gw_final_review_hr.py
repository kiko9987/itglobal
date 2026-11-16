"""그룹웨어 최종 점검(11-16) 인사·연말정산 A급 — 본인 인사카드에 관리 메모가 새던 것,
보완 요청 슬랙 DM 실패가 성공처럼 보이던 것."""
import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import yearend as ye

ADMIN = 'kiko@itg-aircon.com'
STAFF = 'mw@itg-aircon.com'
NOSLACK = 'ns@itg-aircon.com'
Y = 2026


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, ye):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('GW_HR_DOCS_DIR', str(tmp_path / 'docs'))
    sent = []
    # 슬랙 ID 가 없으면 실제 _dm 처럼 False
    monkeypatch.setattr(ye, '_dm', lambda sid, text: sent.append((sid, text)) or bool(sid))
    gw._invalidate()
    hr.upsert_from_flex({'flex_id': 'A', 'name': '박민우', 'status': '재직', 'work_email': STAFF, 'login_email': STAFF,
                         'slack_id': 'UMW', 'memo': '연봉 협상 메모'}, [])
    hr.upsert_from_flex({'flex_id': 'B', 'name': '나슬랙', 'status': '재직', 'work_email': NOSLACK,
                         'login_email': NOSLACK}, [])
    d.sent = sent
    yield d
    gw._invalidate()


# ── 1. 인사 메모·관리 전용 칸 ────────────────────────────────────────────
def test_card_hides_admin_fields_from_non_admin(db):
    pid = hr.find_by_email(STAFF)['id']
    mine = hr.card(pid, STAFF)
    for k in hr.ADMIN_ONLY_FIELDS:
        assert k not in mine, k
    assert mine['name'] == '박민우' and mine['access'] == 'staff'     # 현장 배지에 쓰는 칸은 남김
    assert 'memo' not in hr.card(pid, '')                             # 현장 세션(이메일 없음)도 뺌

    full = hr.card(pid, ADMIN)
    assert full['memo'] == '연봉 협상 메모' and full['slack_id'] == 'UMW' and full['flex_id'] == 'A'
    assert full['can_edit']
    assert hr.get_person(pid)['memo'] == '연봉 협상 메모'              # 저장된 값은 그대로


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
    gw.set_open_all(True, ADMIN)
    return app.test_client()


def _as(c, email, perm='editor'):
    with c.session_transaction() as s:
        s.clear()
        s['user'] = {'email': email, 'permission_level': perm}


def test_api_card_memo_only_for_hr_admin(client):
    pid = hr.find_by_email(STAFF)['id']
    _as(client, STAFF)
    me = client.get(f'/api/hr/people/{pid}').get_json()
    assert me['success'] and me['person']['name'] == '박민우'
    for k in ('memo', 'flex_id', 'slack_id', 'login_email', 'updated_by'):
        assert k not in me['person'], k
    assert '연봉 협상 메모' not in str(me)

    _as(client, ADMIN, 'super_admin')
    full = client.get(f'/api/hr/people/{pid}').get_json()['person']
    assert full['memo'] == '연봉 협상 메모' and full['slack_id'] == 'UMW' and 'access_log' in full


# ── 2. 보완 요청 DM 결과 ────────────────────────────────────────────────
def _submitted(email):
    me = hr.find_by_email(email)
    ye.save_answers(me['id'], Y, {'method': 'bulk', 'bulk_consent': True, 'family_status': 'none'})
    ye.submit(me['id'], Y, me)
    return me['id']


def test_return_reports_notified(client, db):
    ye.set_season(Y, ADMIN, status='open')
    ok_id, fail_id = _submitted(STAFF), _submitted(NOSLACK)
    _as(client, ADMIN, 'super_admin')

    r = client.post(f'/api/yearend/admin/{ok_id}/status?year={Y}', json={'status': 'returned', 'reason': '등본 누락'}).get_json()
    assert r['success'] and r['submission']['status'] == ye.RETURNED and r['notified'] is True
    assert db.sent[-1][0] == 'UMW' and '등본 누락' in db.sent[-1][1]     # 슬랙 ID 있는 사람 → 알림 감

    r = client.post(f'/api/yearend/admin/{fail_id}/status?year={Y}', json={'status': 'returned', 'reason': '등본 누락'}).get_json()
    assert r['success'] and r['submission']['status'] == ye.RETURNED and r['notified'] is False   # 저장은 됐고, 알림은 못 감
    assert ye.my_view(hr.get_person(fail_id), Y)['submission']['status'] == ye.RETURNED

    ye.submit(ok_id, Y, hr.get_person(ok_id))                       # 보완 뒤 다시 제출 → 완료
    r = client.post(f'/api/yearend/admin/{ok_id}/status?year={Y}', json={'status': 'done'}).get_json()
    assert r['success'] and 'notified' not in r                       # 완료 처리는 DM 없음
