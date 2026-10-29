# -*- coding: utf-8 -*-
"""2026-10-29 점검 2묶음 — 보안·백업 감시."""
import os
import sys
import time

sys.path.insert(0, '.')


def test_sheet_safe_text_blocks_formula():
    from dashboard.services.lead_helpers import sheet_safe_text
    assert sheet_safe_text('=HYPERLINK("http://x","확인")') == '\'=HYPERLINK("http://x","확인")'
    assert sheet_safe_text('+8210') == "'+8210"
    assert sheet_safe_text('@me') == "'@me"
    assert sheet_safe_text('-') == '-'
    assert sheet_safe_text('홍길동') == '홍길동'
    assert sheet_safe_text(None) is None


def test_backup_stale(tmp_path):
    from dashboard.services.invariant_checks import backup_stale
    now = time.time()
    for kind, age_h in (('redis', 2), ('users_db', 30)):
        d = tmp_path / kind
        d.mkdir()
        f = d / 'x'
        f.write_text('1')
        os.utime(f, (now - age_h * 3600, now - age_h * 3600))
    got = dict(backup_stale(tmp_path, now=now))
    assert 'redis' not in got
    assert got['users_db'] == 30
    assert got['secrets'] is None          # 폴더 없음


def test_backup_failed_items():
    sys.path.insert(0, 'scripts')
    import backup_daily as bd
    assert bd.failed_items({'users_db': 'OK: a', 'redis': 'OK: b', 'secrets': 'OK: c', 'log_cleanup': 'skip'}) == []
    bad = bd.failed_items({'users_db': 'OK: a', 'redis': 'docker cp 실패: x', 'secrets': 'OK: c'})
    assert bad == ['redis: docker cp 실패: x']


def test_admin_routes_require_login():
    from flask import Flask
    from dashboard.blueprints.slack_bot import slack_bp
    app = Flask(__name__)
    app.secret_key = 'test'
    app.register_blueprint(slack_bp)
    c = app.test_client()
    for path in ('/slack/normalize-visit-dates?dry_run=0', '/slack/migrate-etc-to-sheet?dry_run=0',
                 '/slack/migrate-visit-buttons?dry_run=0&days=3650'):
        assert c.get(path).status_code == 403, path


def test_channeltalk_slack_text_escaped():
    from dashboard.blueprints.channeltalk import _thread_reply_text
    t = _thread_reply_text('<!channel> <https://x|확인> & 끝', '<b>홍</b>', 0)
    assert '<!channel>' not in t and '&lt;!channel&gt;' in t and '&amp; 끝' in t and '&lt;b&gt;' in t
