"""테스트가 운영 users.db 에 쓰지 않게 (2026-11-20 — 전자결재 테스트가 운영 audit_logs 에 가짜 9,815건)."""
import os

import pytest

import dashboard.utils.user_database as udb


def test_default_path_is_temp_under_pytest():
    d = udb.UserDatabase()
    assert d.db_path == os.environ['USER_DB_PATH']
    assert os.path.basename(os.path.dirname(d.db_path)) != 'instance'


def test_prod_path_refused(tmp_path):
    prod_like = tmp_path / 'instance' / 'users.db'
    prod_like.parent.mkdir()
    with pytest.raises(RuntimeError, match='운영 사용자 DB'):
        udb.UserDatabase(db_path=str(prod_like))
    assert not prod_like.exists()


def test_audit_repository_writes_to_temp_db():
    repo = udb.get_audit_repository()
    assert repo.user_db.db_path == os.environ['USER_DB_PATH']
    assert repo.log_action(user_email='slack:테스트', action='EAPPROVAL_TEST', details='격리 확인')
