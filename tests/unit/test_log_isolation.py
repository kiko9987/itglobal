# -*- coding: utf-8 -*-
"""테스트 로그 격리 회귀 테스트 (2026-10-12).

배경: 테스트 실행 기록(가짜 '발송 OK'·'redis down')이 운영 로그
(dashboard/logs/itglobal_dashboard.log)에 섞여, 9/15·10/6 두 번 실제 장애로 오진.
→ 로그 폴더를 LOG_DIR 로 바꿀 수 있게 하고 conftest 가 테스트 시 임시 폴더 지정.
"""
import logging
import logging.handlers
import os
import sys
from pathlib import Path

sys.path.insert(0, '.')

from dashboard.utils import logging_config as lc

_PROD = (Path(lc.__file__).parent.parent / 'logs').resolve()


def _is_under(p: Path, root: Path) -> bool:
    try:
        p.resolve().relative_to(root)
        return True
    except ValueError:
        return False


def test_pytest_log_dir_is_not_prod():
    assert os.environ.get('LOG_DIR')
    assert not _is_under(Path(os.environ['LOG_DIR']), _PROD)
    assert not _is_under(lc.get_log_dir(), _PROD)


def test_get_log_dir_env_and_default(monkeypatch, tmp_path):
    monkeypatch.setenv('LOG_DIR', str(tmp_path))
    assert lc.get_log_dir() == tmp_path
    monkeypatch.delenv('LOG_DIR')
    assert lc.get_log_dir().resolve() == _PROD   # 미설정 = 운영 기본 폴더(동작 불변)


def test_file_handlers_write_outside_prod():
    lc.setup_logging()
    files = [Path(h.baseFilename) for h in logging.getLogger().handlers
             if isinstance(h, logging.FileHandler)]
    assert files, '파일 핸들러가 없음'
    for f in files:
        assert not _is_under(f, _PROD), f'테스트 로그가 운영 폴더에 기록됨: {f}'
