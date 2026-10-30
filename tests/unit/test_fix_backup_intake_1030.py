# -*- coding: utf-8 -*-
"""2026-10-30 — ①users.db 백업을 SQLite 정식 백업으로(WAL 변경분 포함) ②원문 없는 입금 [지정] 안내 창."""
import importlib.util
import sqlite3
from pathlib import Path
from unittest.mock import MagicMock, patch

_spec = importlib.util.spec_from_file_location(
    'backup_daily', Path(__file__).resolve().parents[2] / 'scripts' / 'backup_daily.py')
bd = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(bd)


def test_sqlite_snapshot_includes_uncheckpointed_wal(tmp_path):
    src = tmp_path / 'users.db'
    live = sqlite3.connect(str(src))
    live.execute('PRAGMA journal_mode=WAL')
    live.execute('PRAGMA wal_autocheckpoint=0')        # 변경이 -wal 에만 머물게
    live.execute('CREATE TABLE t (v TEXT)')
    live.executemany('INSERT INTO t VALUES (?)', [(str(i),) for i in range(500)])
    live.commit()
    assert Path(str(src) + '-wal').stat().st_size > 0

    dest = tmp_path / 'out.db'
    bd.sqlite_snapshot(src, dest)                      # 서버가 연결을 쥔 채로 백업
    live.close()
    con = sqlite3.connect(str(dest))
    assert con.execute('SELECT COUNT(*) FROM t').fetchone()[0] == 500
    assert con.execute('PRAGMA journal_mode').fetchone()[0] == 'delete'   # 파일 하나로 열림
    con.close()
    assert not (tmp_path / 'out.db.tmp').exists()


def test_backup_users_db_reports_failure(tmp_path, monkeypatch):
    bad = tmp_path / 'users.db'
    bad.write_bytes(b'not a sqlite file' * 100)
    monkeypatch.setattr(bd, 'USERS_DB_SRC', bad)
    monkeypatch.setattr(bd, 'BACKUP_ROOT', tmp_path / 'backup')
    res = bd.backup_users_db()
    assert res.startswith('실패')
    assert bd.failed_items({'users_db': res, 'redis': 'OK', 'secrets': 'OK'})


def test_intake_open_without_text_shows_notice():
    from dashboard.blueprints import slack_bot as sb
    client = MagicMock()
    client.views_open.return_value = {'view': {'id': 'V1'}}
    body = {'trigger_id': 't'}
    with patch.object(sb, '_load_intake', return_value={}):
        sb._open_payment_intake_modal(client, body, 'gone', 'C1', '1.0')
    view = client.views_update.call_args.kwargs['view']
    assert client.views_update.call_args.kwargs['view_id'] == 'V1'
    assert '이미 처리된 입금' in view['blocks'][0]['text']['text']
    assert 'submit' not in view


def test_intake_view_never_empty_first_block():
    from dashboard.blueprints import slack_bot as sb
    view = sb._build_payment_intake_view('i1', 'C1', '1.0', '')
    assert view['blocks'][0]['text']['text'].strip()
