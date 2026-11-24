"""11-24 조용한 실패 점검(인사) — 증명서 신청이 경영지원에 안 닿을 때·연말정산 묶음에서 파일이 빠질 때."""
import io
import logging
import os
import zipfile

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import certificates as ct
from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import yearend as ye

ADMIN = 'hr-admin@example.com'
Y = 2026
VALS = {'ctype': '재직증명서', 'company': 'itg', 'purpose': '금융기관 제출', 'dest': '○○은행', 'ssn': ct.SSN_OPTS[0],
        'receive': ct.RECEIVE[0], 'copies': '', 'need_by': '', 'memo': ''}


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    for mod in (gw, hr, ct, ye):
        monkeypatch.setattr(mod, '_db', lambda: d)
        monkeypatch.setattr(mod, '_table_ready', False)
    monkeypatch.setenv('GW_HR_KEY', Fernet.generate_key().decode())
    monkeypatch.setenv('GW_HR_DOCS_DIR', str(tmp_path / 'docs'))
    monkeypatch.setenv('EAPPROVAL_EXECUTOR_DM_IDS', 'UHANDLER')
    monkeypatch.setattr(hr, 'hr_admins', lambda: [ADMIN])
    gw._invalidate()
    yield d
    gw._invalidate()


class _Client:
    """fail 에 든 사람에게 보내는 DM 은 실패."""
    def __init__(self, fail=()):
        self.calls, self.fail = [], set(fail)

    def conversations_open(self, users):
        if users in self.fail:
            raise RuntimeError('channel_not_found')
        return {'channel': {'id': 'D' + users}}

    def chat_postMessage(self, **kw):
        self.calls.append(('post', kw['channel'], kw.get('text')))
        return {'ts': '1'}

    def chat_update(self, **kw):
        self.calls.append(('update', kw['channel']))


def _forced_errors(caplog):
    return [r for r in caplog.records if r.levelno >= logging.ERROR and getattr(r, 'force_alert', False)]


# ── 1. 증명서 신청 ───────────────────────────────────────────────────
def test_send_returns_bool(db, monkeypatch):
    from dashboard.blueprints import slack_cert as sc
    row = ct.create('UREQ', '박민우', VALS)
    assert sc._send(_Client(), row, 'UHANDLER', 'handler') is True
    assert sc._send(_Client(fail={'UHANDLER'}), row, 'UHANDLER', 'handler') is False


def test_submit_ok_marks_notified(db, monkeypatch, caplog):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: '박민우')
    with caplog.at_level(logging.INFO):
        row = sc.submit(_Client(), 'UREQ', VALS)
    assert row['handler_notified'] is True and not _forced_errors(caplog)


def test_submit_handler_dm_fails_errors(db, monkeypatch, caplog):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: '박민우')
    c = _Client(fail={'UHANDLER'})
    with caplog.at_level(logging.INFO):
        row = sc.submit(c, 'UREQ', VALS)
    assert row['handler_notified'] is False
    assert ct.get(row['id'])['status'] == ct.PENDING                       # 신청은 저장
    errs = _forced_errors(caplog)
    assert len(errs) == 1 and f"#{row['id']}" in errs[0].getMessage() and '박민우' in errs[0].getMessage() \
        and '재직증명서' in errs[0].getMessage()


def test_submit_no_handlers_configured_errors(db, monkeypatch, caplog):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: '박민우')
    monkeypatch.setattr(ct, 'handler_ids', lambda: ())
    with caplog.at_level(logging.INFO):
        row = sc.submit(_Client(), 'UREQ', VALS)
    assert row['handler_notified'] is False
    assert '설정돼 있지 않음' in _forced_errors(caplog)[0].getMessage()


def test_requester_card_fail_stays_warning(db, monkeypatch, caplog):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: '박민우')
    with caplog.at_level(logging.INFO):
        row = sc.submit(_Client(fail={'UREQ'}), 'UREQ', VALS)
    assert row['handler_notified'] is True
    assert not [r for r in caplog.records if r.levelno >= logging.ERROR]
    assert any(r.levelno == logging.WARNING and 'requester' in r.getMessage() for r in caplog.records)


def test_slack_submit_tells_requester_when_handlers_missed(db, monkeypatch):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: '박민우')
    told = []
    monkeypatch.setattr(se, '_notify', lambda c, uid, text: told.append((uid, text)) or True)
    sc.submit_from_slack(_Client(fail={'UHANDLER'}), 'UREQ', VALS)
    assert told and told[0][0] == 'UREQ' and '경영지원에 직접 말해' in told[0][1]
    told.clear()
    sc.submit_from_slack(_Client(), 'UREQ', VALS)                          # 잘 가면 따로 DM 없음
    assert told == []


def _person():
    hr.upsert_from_flex({'flex_id': 'F1', 'name': '박민우', 'status': '재직', 'hire_date': '2022-02-03', 'department': '영업',
                         'work_email': 'staff@example.com', 'login_email': 'staff@example.com', 'access': 'staff',
                         'ssn': '900101-1234567', 'slack_id': 'UREQ'}, [])
    return hr.find_by_email('staff@example.com')


def _app(monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints import hr as hr_bp_mod
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(hr_bp_mod.hr_bp)
    return app


def test_web_cert_request_reports_not_notified(db, monkeypatch):
    from dashboard.blueprints import slack_eapproval as se
    c = _Client(fail={'UHANDLER'})
    monkeypatch.setattr(se, 'eap_client', lambda: c)
    monkeypatch.setattr(se, '_name', lambda uid: '박민우')
    _person()
    with _app(monkeypatch).test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': 'staff@example.com', 'permission_level': 'user'}
        j = cl.post('/api/hr/me/cert-request', json=VALS).get_json()
        assert j['success'] and j['notified'] is False and '경영지원에 직접 말해' in j['warning']
        assert j['requests'][0]['status'] == 'pending'
        c.fail.clear()
        j = cl.post('/api/hr/me/cert-request', json=VALS).get_json()
        assert j['success'] and j['notified'] is True and 'warning' not in j


def test_gw_hr_js_shows_not_notified_warning():
    src = open(os.path.join(os.path.dirname(__file__), '..', '..', 'dashboard', 'static', 'js', 'gw-hr.js'),
               encoding='utf-8').read()
    assert 'j.notified === false' in src and '경영지원에 직접 말해 주세요' in src


# ── 2. 연말정산 묶음 ─────────────────────────────────────────────────
def _ye_setup():
    hr.upsert_from_flex({'flex_id': 'A', 'name': '박민우', 'status': '재직', 'slack_id': 'UMW',
                         'ssn': '9001011234567'}, [])
    hr.upsert_from_flex({'flex_id': 'B', 'name': '김현장', 'status': '재직', 'access': 'field', 'slack_id': 'UF'}, [])
    people = {p['name']: p for p in hr.list_people(include_resigned=False)}
    ye.set_season(Y, ADMIN, status='open')
    a, b = people['박민우']['id'], people['김현장']['id']
    ye.upload(a, Y, 'simple', '간소화.pdf', b'%PDF-A', 'x')
    gone = ye.upload(a, Y, 'extra', '안경.jpg', b'JPG', 'x')
    ye.upload(b, Y, 'simple', '현장.pdf', b'%PDF-B', 'x')
    return a, b, gone


def _csv_rows(z):
    import csv
    text = z.read(f'{Y}년귀속_연말정산_요약.csv').decode('utf-8-sig')
    return list(csv.reader(io.StringIO(text)))


def test_bundle_all_present_no_missing_list(db, caplog):
    _ye_setup()
    with caplog.at_level(logging.INFO):
        data, miss = ye.bundle_zip_report(Y, ADMIN)
    z = zipfile.ZipFile(io.BytesIO(data))
    assert miss == 0 and ye.MISSING_LIST_NAME not in z.namelist()
    rows = _csv_rows(z)
    assert rows[0][11] == '누락' and all(r[11] == '' for r in rows[1:])
    assert not [r for r in caplog.records if r.levelno >= logging.ERROR]
    assert ye.bundle_zip(Y, ADMIN)[:2] == b'PK'                            # 예전 이름도 그대로


def test_bundle_missing_file_listed_csv_and_error(db, caplog):
    a, b, gone = _ye_setup()
    f = hr.document_file(gone['id'])
    os.remove(f['path'])
    with caplog.at_level(logging.INFO):
        data, miss = ye.bundle_zip_report(Y, ADMIN)
    z = zipfile.ZipFile(io.BytesIO(data))
    names = z.namelist()
    assert miss == 1 and ye.MISSING_LIST_NAME in names
    assert any(n.startswith(f'박민우_{a}/간소화_') for n in names)           # 남은 파일은 그대로
    assert not any('안경' in n for n in names if n != ye.MISSING_LIST_NAME)
    lst = z.read(ye.MISSING_LIST_NAME).decode('utf-8-sig')
    assert '박민우\t추가증빙\t안경.jpg\t파일 없음' in lst
    assert f['path'] not in lst and 'docs' not in lst and '900101' not in lst    # 경로·주민번호 X
    rows = {r[0]: r for r in _csv_rows(z)[1:]}
    assert rows['박민우'][10] == '1' and rows['박민우'][11] == '1' and rows['김현장'][11] == ''
    errs = _forced_errors(caplog)
    assert len(errs) == 1 and '박민우' in errs[0].getMessage() and '1개' in errs[0].getMessage()
    assert f['path'] not in errs[0].getMessage() and '9001011234567' not in errs[0].getMessage()
    with hr._conn() as conn:
        detail = conn.execute("SELECT detail FROM hr_access_log WHERE action = 'ye_bundle' ORDER BY id DESC").fetchone()[0]
    assert '누락 1개' in detail


def test_bundle_route_sets_missing_header(db, monkeypatch):
    a, b, gone = _ye_setup()
    os.remove(hr.document_file(gone['id'])['path'])
    monkeypatch.setattr(ye, 'current_year', lambda today=None: Y)
    with _app(monkeypatch).test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': ADMIN, 'name': '경영지원', 'permission_level': 'super_admin'}
        r = cl.get(f'/api/yearend/admin/bundle.zip?year={Y}')
        assert r.status_code == 200 and r.headers['X-Missing-Files'] == '1'


def test_gw_yearend_js_reads_missing_header():
    src = open(os.path.join(os.path.dirname(__file__), '..', '..', 'dashboard', 'static', 'js', 'gw-yearend.js'),
               encoding='utf-8').read()
    assert "X-Missing-Files" in src and 'id="yeBundle"' in src
