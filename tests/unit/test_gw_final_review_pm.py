"""그룹웨어 최종 점검(11-16) 예전 PM 화면 A급 — 작업 로그(최고관리자·기간·DB 오류),
매출통계 세션 만료·오류 원문 노출."""
import sqlite3

import pytest
from flask import Blueprint, Flask

ADMIN = 'kiko@itg-aircon.com'
STAFF = 'mw@itg-aircon.com'


# ── 작업 로그 ──────────────────────────────────────────────────────────
@pytest.fixture
def audit_db(tmp_path, monkeypatch):
    from dashboard.utils import user_database as udb
    d = udb.UserDatabase(db_path=str(tmp_path / 'users.db'))
    repo = udb.AuditLogRepository(d)
    with d._get_connection() as conn:
        rows = [
            (ADMIN, 'update', "datetime('now', '-1 hours')"),
            (STAFF, 'update', "datetime('now', '-2 days')"),
            (STAFF, 'update', "datetime('now', '-10 days')"),   # 7일 밖
            (ADMIN, 'update', "datetime('now', '-40 days')"),   # 30일 밖
        ]
        for email, action, ts in rows:
            conn.execute(f'INSERT INTO audit_logs (user_email, action, timestamp) VALUES (?, ?, {ts})',
                         (email, action))
        conn.commit()
    monkeypatch.setattr(udb, 'get_audit_repository', lambda: repo)
    monkeypatch.setattr(udb, 'get_user_database', lambda db_path=None: d)
    return repo


def test_get_recent_logs_filters_by_days(audit_db):
    assert len(audit_db.get_recent_logs(limit=1000)) == 4
    assert len(audit_db.get_recent_logs(limit=1000, days=7)) == 2
    assert len(audit_db.get_recent_logs(limit=1000, days=30)) == 3
    mine = audit_db.get_recent_logs(limit=1000, user_email=STAFF, days=7)
    assert [r['user_email'] for r in mine] == [STAFF]


def test_get_recent_logs_raises_on_db_error():
    from dashboard.utils.user_database import AuditLogRepository

    class Broken:
        def _get_connection(self):
            raise sqlite3.OperationalError('disk I/O error')

    with pytest.raises(sqlite3.OperationalError):
        AuditLogRepository(Broken()).get_recent_logs(limit=10, days=7)


@pytest.fixture
def mon_client(monkeypatch):
    import dashboard.auth as auth
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)
    from dashboard.blueprints.monitoring import monitoring_bp
    app = Flask('t')
    app.secret_key = 't'
    app.register_blueprint(monitoring_bp)
    return app.test_client()


def _as(c, email, perm):
    with c.session_transaction() as s:
        s.clear()
        s['user'] = {'email': email, 'name': 'x', 'permission_level': perm}


@pytest.mark.parametrize('perm', ['admin', 'super_admin'])
def test_admins_see_all_logs(audit_db, mon_client, perm):
    _as(mon_client, ADMIN, perm)
    j = mon_client.get('/api/audit-logs?days=30').get_json()
    assert j['success'] and j['pagination']['total_count'] == 3
    assert {log['user_email'] for log in j['logs']} == {ADMIN, STAFF}


def test_staff_sees_only_own_logs_within_days(audit_db, mon_client):
    _as(mon_client, STAFF, 'editor')
    j = mon_client.get('/api/audit-logs?days=7').get_json()
    assert j['success'] and j['pagination']['total_count'] == 1
    assert j['logs'][0]['user_email'] == STAFF


def test_days_filter_on_api(audit_db, mon_client):
    _as(mon_client, ADMIN, 'super_admin')
    assert mon_client.get('/api/audit-logs?days=7').get_json()['pagination']['total_count'] == 2
    assert mon_client.get('/api/audit-logs?days=90').get_json()['pagination']['total_count'] == 4


def test_db_error_returns_failure(mon_client, monkeypatch):
    from dashboard.utils import user_database as udb

    class Broken:
        def _get_connection(self):
            raise sqlite3.OperationalError('disk I/O error at C:/secret/users.db')

    monkeypatch.setattr(udb, 'get_audit_repository', lambda: udb.AuditLogRepository(Broken()))
    monkeypatch.setattr(udb, 'get_user_database', lambda db_path=None: Broken())
    _as(mon_client, ADMIN, 'super_admin')
    r = mon_client.get('/api/audit-logs?days=7')
    assert r.status_code == 500
    j = r.get_json()
    assert j['success'] is False and '작업 로그를 불러오지 못했습니다' in j['message']
    assert 'secret' not in r.get_data(as_text=True)


# ── 매출통계 ───────────────────────────────────────────────────────────
@pytest.fixture
def stats_app():
    from dashboard.blueprints.stats import stats_bp
    app = Flask('t')
    app.secret_key = 't'
    auth_stub = Blueprint('auth', __name__)
    auth_stub.add_url_rule('/login', 'login_page', lambda: 'login')
    app.register_blueprint(auth_stub)
    app.register_blueprint(stats_bp)
    return app


def test_stats_page_redirects_to_login_when_logged_out(stats_app):
    r = stats_app.test_client().get('/stats')
    assert r.status_code == 302
    assert '/login' in r.headers['Location']


def test_stats_page_error_hides_exception_text(stats_app, monkeypatch):
    import dashboard.auth as auth
    import dashboard.blueprints.stats as stats
    monkeypatch.setattr(auth, '_validate_session', lambda require_admin=False: None)

    def boom(*a, **k):
        raise RuntimeError('TemplateNotFound C:/secret/path/modern_stats.html')

    monkeypatch.setattr(stats, 'render_template', boom)
    c = stats_app.test_client()
    _as(c, ADMIN, 'admin')
    r = c.get('/stats')
    body = r.get_data(as_text=True)
    assert r.status_code == 500
    assert 'secret' not in body and 'TemplateNotFound' not in body
    assert '통계를 불러오지 못했어요' in body
