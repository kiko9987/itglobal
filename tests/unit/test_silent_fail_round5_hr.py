"""11-24 조용한 실패 점검 5차(인사) — 서류 파일 못 지움·증명서 바로 발급 예외/유령 줄/중복 발급·
[📎 파일 올려 발급] 일부 실패·반려/취소 스레드·연말정산 명단 확인 필요·상태 충돌."""
import io
import logging
import os

import pytest
from cryptography.fernet import Fernet
from flask import Flask

from dashboard.services import certificates as ct
from dashboard.services import groupware as gw
from dashboard.services import hr
from dashboard.services import yearend as ye

ADMIN = 'hr-admin@example.com'
STAFF = 'staff@example.com'
Y = 2026
VALS = {'ctype': '재직증명서', 'company': 'group', 'purpose': '금융기관 제출', 'dest': '○○은행', 'ssn': ct.SSN_OPTS[0],
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
    # 바로 발급 준비(글꼴·직인) — 진짜 글꼴 대신 빈 파일, PDF 는 시험마다 가짜로
    fonts = tmp_path / 'fonts'
    fonts.mkdir()
    (fonts / 'malgun.ttf').write_bytes(b'x')
    (fonts / 'malgunbd.ttf').write_bytes(b'x')
    monkeypatch.setenv('CERT_FONT_DIR', str(fonts))
    (tmp_path / 'seal.png').write_bytes(b'png')
    monkeypatch.setenv('CERT_SEAL_DIR', str(tmp_path))
    monkeypatch.setattr(hr, 'hr_admins', lambda: [ADMIN])
    gw._invalidate()
    yield d
    gw._invalidate()


def _person(**kw):
    rec = {'flex_id': 'F1', 'name': '박민우', 'status': '재직', 'hire_date': '2022-02-03', 'department': '영업',
           'work_email': STAFF, 'login_email': STAFF, 'access': 'staff', 'ssn': '900101-1234567', 'slack_id': 'UREQ'}
    rec.update(kw)
    hr.upsert_from_flex(rec, [])
    return hr.find_by_email(rec['work_email'])


def _forced_errors(caplog):
    return [r for r in caplog.records if r.levelno >= logging.ERROR and getattr(r, 'force_alert', False)]


def _fake_pdf(*a, **k):
    return b'%PDF-fake'


def _issue_count():
    with ct._conn() as conn:
        return conn.execute('SELECT COUNT(*) FROM cert_issues').fetchone()[0]


# ── 1. 서류 삭제 — 파일을 못 지우면 기록·ERROR·다시 지우기 ─────────────────
def test_delete_document_file_locked_is_recorded_and_retried(db, monkeypatch, caplog):
    p = _person()
    saved = hr.save_document(p['id'], '연말정산 간소화자료', '간소화_박민우.pdf', b'%PDF', ADMIN)
    path = hr.document_file(saved['id'])['path']
    real_remove = os.remove

    def locked(pth):
        raise PermissionError(13, '다른 프로그램이 사용 중', pth)
    monkeypatch.setattr(hr.os, 'remove', locked)
    caplog.set_level(logging.INFO)
    hr.delete_document(saved['id'], ADMIN)
    assert hr.documents(p['id']) == [] and os.path.exists(path)          # 장부는 지워지고 파일은 남음
    errs = _forced_errors(caplog)
    assert errs and f"#{saved['id']}" in errs[0].getMessage() and '연말정산 간소화자료' in errs[0].getMessage()
    assert all('간소화_박민우' not in r.getMessage() and path not in r.getMessage() for r in caplog.records)
    left = hr.undeleted_files()
    assert len(left) == 1 and left[0]['doc_id'] == saved['id'] and left[0]['last_error'].startswith('PermissionError')

    # 매일 09:00 — 아직 잠겨 있으면 ERROR(번호·종류만), 풀리면 지우고 목록에서 뺌
    caplog.clear()
    assert hr.retry_undeleted_files(alert=True) == {'removed': 0, 'left': 1}
    assert _forced_errors(caplog) and path not in _forced_errors(caplog)[0].getMessage()
    monkeypatch.setattr(hr.os, 'remove', real_remove)
    assert hr.retry_undeleted_files(alert=True) == {'removed': 1, 'left': 0}
    assert not os.path.exists(path) and hr.undeleted_files() == []


def test_delete_document_retries_old_leftovers_and_ignores_missing_file(db, monkeypatch, caplog):
    p = _person()
    a = hr.save_document(p['id'], '기타', 'a.pdf', b'A', ADMIN)
    b = hr.save_document(p['id'], '기타', 'b.pdf', b'B', ADMIN)
    pa = hr.document_file(a['id'])['path']
    real_remove = os.remove
    monkeypatch.setattr(hr.os, 'remove', lambda pth: (_ for _ in ()).throw(PermissionError(13, 'x', pth)))
    hr.delete_document(a['id'], ADMIN)
    assert len(hr.undeleted_files()) == 1
    monkeypatch.setattr(hr.os, 'remove', real_remove)
    os.remove(hr.document_file(b['id'])['path'])                        # 파일이 이미 없음 → 조용히
    caplog.clear()
    hr.delete_document(b['id'], ADMIN)                                   # 이참에 a 파일도 다시 지움
    assert not _forced_errors(caplog)
    assert hr.undeleted_files() == [] and not os.path.exists(pa)


def test_daily_hr_job_retries_undeleted_files(monkeypatch):
    from dashboard.blueprints import slack_hr_tasks
    from dashboard.services import sync_scheduler as ss
    called = []
    monkeypatch.setattr(slack_hr_tasks, 'run_daily', lambda: called.append('tasks'))
    monkeypatch.setattr(hr, 'retry_undeleted_files', lambda alert=True: called.append(('retry', alert)))
    ss._safe_hr_tasks()
    assert called == ['tasks', ('retry', True)]


# ── 2·3. 증명서 바로 발급 — 입력 오류만 문구 그대로, 나머지는 ERROR + 일반 문구, 유령 줄 없음 ─────
def test_issue_pdf_input_errors_are_cert_input_error(db):
    p = _person()
    with pytest.raises(ct.CertInputError):
        ct.issue_pdf(p['id'], '이상한 증명서', 'x', ADMIN, company_key='group')
    with pytest.raises(ct.CertInputError):
        ct.issue_pdf(999, '재직증명서', 'x', ADMIN, company_key='group')
    with pytest.raises(ct.CertNotReady) as e:                          # 아이티글로벌 직인 그림 없음
        ct.issue_pdf(p['id'], '재직증명서', 'x', ADMIN, company_key='itg')
    assert isinstance(e.value, RuntimeError) and isinstance(e.value, ct.CertInputError)
    assert _issue_count() == 0


def test_issue_pdf_failure_after_number_leaves_no_ghost_row(db, monkeypatch, caplog):
    p = _person()

    def broken(*a, **k):
        raise KeyError('glyph')
    monkeypatch.setattr(ct, 'render_pdf', broken)
    with pytest.raises(KeyError):
        ct.issue_pdf(p['id'], '재직증명서', 'x', ADMIN, company_key='group')
    assert _issue_count() == 0 and hr.documents(p['id']) == []

    monkeypatch.setattr(ct, 'render_pdf', _fake_pdf)

    def save_fails(*a, **k):
        raise OSError('disk full')
    monkeypatch.setattr(hr, 'save_document', save_fails)
    with pytest.raises(OSError):
        ct.issue_pdf(p['id'], '재직증명서', 'x', ADMIN, company_key='group')
    assert _issue_count() == 0


def test_issue_pdf_failure_after_save_removes_document_too(db, monkeypatch):
    p = _person()
    monkeypatch.setattr(ct, 'render_pdf', _fake_pdf)
    real_conn = ct._conn
    state = {'n': 0}

    from contextlib import contextmanager

    @contextmanager
    def flaky():
        state['n'] += 1
        if state['n'] == 2:                                             # INSERT 다음 UPDATE 에서 실패
            raise RuntimeError('db locked')
        with real_conn() as c:
            yield c
    monkeypatch.setattr(ct, '_conn', flaky)
    with pytest.raises(RuntimeError, match='db locked'):
        ct.issue_pdf(p['id'], '재직증명서', 'x', ADMIN, company_key='group')
    monkeypatch.setattr(ct, '_conn', real_conn)
    assert _issue_count() == 0 and hr.documents(p['id']) == []
    # 다음 발급은 번호가 건너뛰어도(구멍) 정상
    res = ct.issue_pdf(p['id'], '재직증명서', 'x', ADMIN, company_key='group')
    assert res['doc_no'].endswith('-0002') and _issue_count() == 1


@pytest.fixture
def web(db, monkeypatch):
    import dashboard.auth as auth
    from dashboard.blueprints import hr as hr_bp_mod
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    monkeypatch.setattr(gw, 'can_use', lambda e: True)
    app = Flask(__name__)
    app.secret_key = 'x'
    app.register_blueprint(hr_bp_mod.hr_bp)
    with app.test_client() as cl:
        with cl.session_transaction() as s:
            s['user'] = {'email': ADMIN, 'permission_level': 'admin'}
        yield cl


def test_web_certificate_errors(web, monkeypatch, caplog):
    p = _person()
    url = f"/api/hr/people/{p['id']}/certificate"
    # 입력 오류 — 문구 그대로 400, 알림 없음
    r = web.post(url, json={'company': 'group', 'ctype': '이상한'})
    assert r.status_code == 400 and '종류' in r.get_json()['message'] and not _forced_errors(caplog)
    # PDF 만들다 코드 오류(KeyError) — 원문 대신 일반 문구 + ERROR, 유령 줄 없음
    monkeypatch.setattr(ct, 'render_pdf', lambda *a, **k: (_ for _ in ()).throw(KeyError('days')))
    r = web.post(url, json={'company': 'group', 'ctype': '재직증명서', 'purpose': 'x'})
    assert r.status_code == 500 and r.get_json()['message'] == ct.ISSUE_FAILED_MSG and 'days' not in r.get_json()['message']
    assert _forced_errors(caplog) and _issue_count() == 0
    # 주민번호 열쇠 없음 — 'GW_HR_KEY' 이름 대신 안내
    caplog.clear()
    monkeypatch.delenv('GW_HR_KEY')
    r = web.post(url, json={'company': 'group', 'ctype': '재직증명서', 'purpose': 'x', 'ssn': True})
    assert r.status_code == 503 and r.get_json()['message'] == ct.SSN_KEY_MISSING_MSG
    assert 'GW_HR_KEY' not in r.get_json()['message'] and _forced_errors(caplog)


# ── 4. 슬랙 바로 발급 — DM 실패 뒤 다시 누르면 새로 만들지 않고 다시 보내기 ─────────────
class _Client:
    def __init__(self, upload_fails=False):
        self.calls, self.upload_fails = [], upload_fails

    def conversations_open(self, users):
        return {'channel': {'id': 'D' + users}}

    def files_upload_v2(self, **kw):
        if self.upload_fails:
            raise RuntimeError('upload failed')
        self.calls.append(('upload', kw['channel'], [u['filename'] for u in kw['file_uploads']]))

    def chat_update(self, **kw):
        self.calls.append(('update', kw['channel']))

    def chat_postMessage(self, **kw):
        self.calls.append(('post', kw['channel'], kw.get('text')))
        return {'ts': '1'}

    def chat_postEphemeral(self, **kw):
        self.calls.append(('ephemeral', kw['user'], kw.get('text')))


def test_slack_auto_issue_resends_instead_of_reissuing(db, monkeypatch, caplog):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: '황샛별')
    monkeypatch.setattr(ct, 'render_pdf', _fake_pdf)
    p = _person()
    row = ct.create('UREQ', '박민우', VALS)
    msg = sc.auto_issue(_Client(upload_fails=True), row['id'], 'UHANDLER')
    assert '보내기만' in msg and ct.get(row['id'])['status'] == ct.PENDING
    assert _forced_errors(caplog) and _issue_count() == 1
    c = _Client()
    assert sc.auto_issue(c, row['id'], 'UHANDLER') == ''
    assert _issue_count() == 1 and len(hr.documents(p['id'])) == 1          # 새로 발급하지 않음
    ups = [x for x in c.calls if x[0] == 'upload']
    assert ups and ups[0][1] == 'DUREQ' and ups[0][2] == [hr.documents(p['id'])[0]['file_name']]
    assert any(x[0] == 'post' and x[1] == 'DUHANDLER' and '다시 보내기만' in x[2] for x in c.calls)
    assert ct.get(row['id'])['status'] == ct.ISSUED


def test_slack_auto_issue_error_kinds(db, monkeypatch, caplog):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: '황샛별')
    _person()
    row = ct.create('UREQ', '박민우', dict(VALS, ssn=ct.SSN_OPTS[1]))
    monkeypatch.setattr(ct, 'render_pdf', lambda *a, **k: (_ for _ in ()).throw(ValueError('PNG 형식 아님')))
    msg = sc.auto_issue(_Client(), row['id'], 'UHANDLER')
    assert 'PNG' not in msg and '부장님' in msg and _forced_errors(caplog) and _issue_count() == 0
    caplog.clear()
    monkeypatch.delenv('GW_HR_KEY')
    assert sc.auto_issue(_Client(), row['id'], 'UHANDLER') == ct.SSN_KEY_MISSING_MSG and _forced_errors(caplog)
    assert ct.get(row['id'])['status'] == ct.PENDING


# ── 5. [📎 파일 올려 발급] — 일부만 받아지면 못 보낸 파일 안내 ─────────────────
def test_issue_partial_files_tells_handler(db, monkeypatch, caplog):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: '황샛별')
    monkeypatch.setattr(se, '_download', lambda url: b'%PDF' if url == 'ok' else None)
    row = ct.create('UREQ', '박민우', VALS)
    c = _Client()
    caplog.set_level(logging.WARNING)
    msg = sc.issue(c, row['id'], 'UHANDLER', [{'name': 'a.pdf', 'url': 'ok'}, {'name': 'b.pdf', 'url': 'bad'}], '')
    assert 'b.pdf' in msg and 'a.pdf' not in msg and '직접' in msg
    assert ('upload', 'DUREQ', ['a.pdf']) in c.calls and ct.get(row['id'])['status'] == ct.ISSUED
    assert any(r.levelno == logging.WARNING and 'b.pdf' in r.getMessage() for r in caplog.records)


# ── 6. 반려·취소 스레드 — 예외는 ERROR + 누른 사람 안내, 반려 DM 실패는 ERROR ─────────────
class _App:
    def __init__(self):
        self.h = {}

    def action(self, k):
        return lambda fn: self.h.setdefault(('a', k), fn)

    def view(self, k):
        return lambda fn: self.h.setdefault(('v', k), fn)


class _SyncThread:
    def __init__(self, target=None, daemon=None):
        self.target = target

    def start(self):
        self.target()


@pytest.fixture
def handlers(db, monkeypatch):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(sc.threading, 'Thread', _SyncThread)
    monkeypatch.setattr(se, '_name', lambda uid: '황샛별')
    c = _Client()
    monkeypatch.setattr(se, 'eap_client', lambda: c)
    app = _App()
    sc.register_cert_handlers(app)
    return app.h, c


def _reject_body(rid, reason='서류 부족'):
    import json
    view = {'private_metadata': json.dumps({'id': rid}), 'state': {'values': {'r_reason': {'r_reason': {'value': reason}}}}}
    return {'user': {'id': 'UHANDLER'}}, view


def test_reject_thread_error_alerts_and_tells_handler(handlers, monkeypatch, caplog):
    h, c = handlers
    row = ct.create('UREQ', '박민우', VALS)
    monkeypatch.setattr(ct, 'finish', lambda *a, **k: (_ for _ in ()).throw(RuntimeError('db locked')))
    body, view = _reject_body(row['id'])
    h[('v', ct.CB_REJECT)](lambda **k: None, body, view)
    assert _forced_errors(caplog)
    assert any(x[0] == 'post' and x[1] == 'DUHANDLER' and '반려 처리 중 오류' in x[2] for x in c.calls)


def test_reject_dm_failure_is_error(handlers, monkeypatch, caplog):
    h, c = handlers
    row = ct.create('UREQ', '박민우', VALS)
    real_open = c.conversations_open

    def open_(users):
        if users == 'UREQ':
            raise RuntimeError('channel_not_found')
        return real_open(users)
    c.conversations_open = open_
    body, view = _reject_body(row['id'])
    h[('v', ct.CB_REJECT)](lambda **k: None, body, view)
    assert ct.get(row['id'])['status'] == ct.REJECTED
    assert any('UREQ' in r.getMessage() for r in _forced_errors(caplog))


def test_cancel_thread_error_alerts_and_tells_requester(handlers, monkeypatch, caplog):
    h, c = handlers
    row = ct.create('UREQ', '박민우', VALS)
    monkeypatch.setattr(ct, 'finish', lambda *a, **k: (_ for _ in ()).throw(RuntimeError('db locked')))
    body = {'user': {'id': 'UREQ'}, 'actions': [{'value': str(row['id'])}], 'channel': {'id': 'DUREQ'}}
    h[('a', ct.A_CANCEL)](lambda: None, body, c)
    assert _forced_errors(caplog)
    assert any(x[0] == 'ephemeral' and x[1] == 'UREQ' and '취소 중 오류' in x[2] for x in c.calls)


def test_original_pickup_dm_failure_is_error(db, monkeypatch, caplog):
    from dashboard.blueprints import slack_cert as sc
    from dashboard.blueprints import slack_eapproval as se
    monkeypatch.setattr(se, '_name', lambda uid: '황샛별')
    row = ct.create('UREQ', '박민우', dict(VALS, receive=ct.RECEIVE[1]))
    c = _Client()
    c.conversations_open = lambda users: (_ for _ in ()).throw(RuntimeError('x')) if users == 'UREQ' else {'channel': {'id': 'D' + users}}
    assert sc.issue(c, row['id'], 'UHANDLER', [], '') == ''
    assert any('원본' in r.getMessage() for r in _forced_errors(caplog))


# ── 7. 연말정산 — 명단 '확인 필요' 시트·상태 충돌 409 ─────────────────────────
def test_roster_check_sheet_and_error_without_ssn(db, caplog):
    import openpyxl
    _person()
    hr.upsert_from_flex({'flex_id': 'F2', 'name': '김없음', 'status': '재직'}, [])
    hr.upsert_from_flex({'flex_id': 'F3', 'name': '이깨짐', 'status': '재직', 'ssn': '8505052234567'}, [])
    with hr._conn() as conn:
        conn.execute("UPDATE hr_people SET ssn_enc='garbage' WHERE name='이깨짐'")
    wb = openpyxl.load_workbook(io.BytesIO(ye.roster_xlsx(Y, ADMIN)))
    assert ye.ROSTER_CHECK_SHEET in wb.sheetnames
    rows = list(wb[ye.ROSTER_CHECK_SHEET].iter_rows(values_only=True))
    assert rows[0] == ('성명', '사유') and {r[0] for r in rows[1:]} == {'김없음', '이깨짐'}
    assert not any(ch.isdigit() for r in rows[1:] for ch in (r[1] or ''))           # 사유에 주민번호 없음
    errs = _forced_errors(caplog)
    assert len(errs) == 1 and '이깨짐' in errs[0].getMessage() and '김없음' not in errs[0].getMessage()
    assert all('900101' not in r.getMessage() and '850505' not in r.getMessage() for r in caplog.records)


def test_roster_no_check_sheet_when_all_ok(db, caplog):
    import openpyxl
    _person()
    wb = openpyxl.load_workbook(io.BytesIO(ye.roster_xlsx(Y, ADMIN)))
    assert ye.ROSTER_CHECK_SHEET not in wb.sheetnames and not _forced_errors(caplog)


def test_set_status_requires_current_state(db):
    me = _person()
    ye.set_season(Y, ADMIN, status='open')
    ye.save_answers(me['id'], Y, {'method': 'bulk', 'bulk_consent': True, 'family_status': 'none'})
    ye.submit(me['id'], Y, me)
    ye.reopen(me['id'], Y)                                              # 직원이 그 사이 「수정하기」
    with pytest.raises(ye.StatusConflict):
        ye.set_status(me['id'], Y, ye.DONE, ADMIN)
    with pytest.raises(ye.StatusConflict):
        ye.set_status(me['id'], Y, ye.RETURNED, ADMIN, '등본 누락')
    assert ye._sub(Y, me['id'])['status'] == ye.DRAFT                   # 덮지 않음
    ye.submit(me['id'], Y, me)
    with pytest.raises(ye.StatusConflict):                              # 완료 아닌데 완료 취소
        ye.set_status(me['id'], Y, ye.SUBMITTED, ADMIN)
    assert ye.set_status(me['id'], Y, ye.DONE, ADMIN)['status'] == ye.DONE
    assert ye.set_status(me['id'], Y, ye.SUBMITTED, ADMIN)['status'] == ye.SUBMITTED
    with pytest.raises(ValueError):
        ye.set_status(me['id'], Y, 'draft', ADMIN)


def test_status_route_returns_409_on_conflict(web, monkeypatch):
    monkeypatch.setattr(ye, 'current_year', lambda today=None: Y)
    me = _person()
    ye.set_season(Y, ADMIN, status='open')
    r = web.post(f"/api/yearend/admin/{me['id']}/status?year={Y}", json={'status': 'done'})
    assert r.status_code == 409 and r.get_json()['message'] == ye.STATUS_CONFLICT_MSG


def test_yearend_js_slack_wording():
    src = open(os.path.join(os.path.dirname(__file__), '..', '..', 'dashboard', 'static', 'js', 'gw-yearend.js'),
               encoding='utf-8').read()
    assert '슬랙이 연결되지 않아' not in src and '슬랙 알림을 못 보낸' in src and 'e.status === 409' in src
