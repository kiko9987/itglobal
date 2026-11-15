# -*- coding: utf-8 -*-
"""워치독 SYSTEM 폴더 (2026-11-15 5차 점검 개선) — 재시작 요청 파일을 사용자 폴더에 두지 않는다."""
import sys
from pathlib import Path

sys.path.insert(0, '.')

import pytest

from dashboard.services import restart_request as rr

ROOT = Path(__file__).resolve().parents[2]


def test_request_path_is_outside_project():
    assert ROOT not in rr.REQUEST_PATH.parents
    assert rr.REQUEST_PATH.name == 'restart_request.json'


def test_save_refuses_to_create_folder(tmp_path):
    with pytest.raises(FileNotFoundError):
        rr.save_request({'id': 'x'}, path=tmp_path / 'missing' / 'restart_request.json')
    assert not (tmp_path / 'missing').exists()


def test_watchdog_trusts_only_sysdir_and_checks_approver():
    src = (ROOT / 'watchdog.ps1').read_text(encoding='utf-8-sig')
    assert "$RequestPath        = Join-Path $SysDir 'restart_request.json'" in src
    assert "Join-Path $PSScriptRoot 'dashboard\logs\restart_request.json'" not in src
    assert "Get-EnvVal 'SLACK_ADMIN_CHANNEL'" in src and '$r.decided_by' in src


def test_installer_locks_folder_and_runs_copy():
    src = (ROOT / 'install_watchdog.ps1').read_text(encoding='utf-8-sig')
    assert '/inheritance:r' in src and '/setowner' in src
    assert "-ProjectRoot" in src and "Join-Path $SysDir 'watchdog.ps1'" in src
