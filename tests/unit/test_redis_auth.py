# -*- coding: utf-8 -*-
"""Redis 비밀번호 적용(2026-10-29) — 일일 백업 redis-cli 인증·BGSAVE 응답 판정."""
import importlib.util
import sys
from pathlib import Path

sys.path.insert(0, '.')

_spec = importlib.util.spec_from_file_location(
    'backup_daily', Path(__file__).resolve().parents[2] / 'scripts' / 'backup_daily.py')
bd = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(bd)


def test_redis_cli_cmd_passes_password_via_env(monkeypatch):
    monkeypatch.setenv('REDIS_PASSWORD', 'pw-123')
    cmd = bd.redis_cli_cmd('BGSAVE')
    assert cmd[:4] == ['docker', 'exec', '-e', 'REDISCLI_AUTH=pw-123']
    assert cmd[-2:] == ['redis-cli', 'BGSAVE'] and '-a' not in cmd


def test_redis_cli_cmd_without_password(monkeypatch):
    monkeypatch.delenv('REDIS_PASSWORD', raising=False)
    assert bd.redis_cli_cmd('BGSAVE') == ['docker', 'exec', bd.REDIS_CONTAINER, 'redis-cli', 'BGSAVE']


def test_bgsave_started():
    assert bd.bgsave_started('Background saving started\n')
    assert bd.bgsave_started('ERR Background save already in progress')
    # 비밀번호 없이 보내면 종료 코드 0 + NOAUTH — 저장 안 됨
    assert not bd.bgsave_started('NOAUTH Authentication required.\n')
    assert not bd.bgsave_started('')
