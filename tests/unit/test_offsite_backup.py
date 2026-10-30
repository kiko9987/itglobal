# -*- coding: utf-8 -*-
"""외부 백업 — 잠그기·풀기, 묶음 내용, 보관 정리, 업로드 흐름, 멈춤 감지 (2026-10-30)."""
import importlib.util
import json
import sys
import time
from datetime import date
from pathlib import Path

import pytest

sys.path.insert(0, '.')
_spec = importlib.util.spec_from_file_location(
    'offsite_backup', Path(__file__).resolve().parents[2] / 'scripts' / 'offsite_backup.py')
ob = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(ob)

PW = 'test-pw-123'
STAMP = '20261030'


def _make_backup(root: Path, stamp=STAMP):
    (root / 'users_db').mkdir(parents=True)
    (root / 'redis').mkdir()
    (root / 'secrets' / stamp).mkdir(parents=True)
    (root / 'users_db' / f'{stamp}.db').write_bytes(b'sqlite-bytes')
    (root / 'redis' / f'{stamp}.rdb').write_bytes(b'REDIS0011' * 100)
    (root / 'secrets' / stamp / '.env').write_text('A=1\n', encoding='utf-8')
    (root / 'secrets' / stamp / 'credentials.json').write_text('{}', encoding='utf-8')


def test_encrypt_roundtrip_and_wrong_password():
    from cryptography.exceptions import InvalidTag
    blob = ob.encrypt(b'hello' * 1000, PW)
    assert blob.startswith(ob.MAGIC) and b'hello' not in blob
    assert ob.decrypt(blob, PW) == b'hello' * 1000
    with pytest.raises(InvalidTag):
        ob.decrypt(blob, 'wrong')
    tampered = blob[:-1] + bytes([blob[-1] ^ 1])
    with pytest.raises(InvalidTag):
        ob.decrypt(tampered, PW)
    with pytest.raises(ValueError):
        ob.encrypt(b'x', '')


def test_pack_and_restore(tmp_path):
    _make_backup(tmp_path / 'b')
    blob = ob.encrypt(ob.pack(tmp_path / 'b', STAMP), PW)
    src = tmp_path / 'x.itgbak'
    src.write_bytes(blob)
    names = ob.restore(str(src), str(tmp_path / 'out'), PW)
    assert set(names) == {'users.db', 'dump.rdb', 'secrets/.env', 'secrets/credentials.json'}
    assert (tmp_path / 'out' / 'secrets' / '.env').read_text(encoding='utf-8') == 'A=1\n'
    assert (tmp_path / 'out' / 'dump.rdb').read_bytes() == b'REDIS0011' * 100


def test_pack_refuses_partial(tmp_path):
    _make_backup(tmp_path / 'b')
    (tmp_path / 'b' / 'redis' / f'{STAMP}.rdb').unlink()
    with pytest.raises(FileNotFoundError):
        ob.pack(tmp_path / 'b', STAMP)


def test_to_trash_retention():
    today = date(2026, 10, 30)
    names = ['itg-backup-20261030.itgbak', 'itg-backup-20260801.itgbak',   # 90일 안
             'itg-backup-20260715.itgbak',                                  # 90일 넘음, 1일 아님 → 정리
             'itg-backup-20260601.itgbak', 'itg-backup-20251001.itgbak',    # 매달 1일, 400일 안 → 유지
             'itg-backup-20250901.itgbak',                                  # 400일 넘음 → 정리
             '메모.txt']                                                    # 다른 파일 손대지 않음
    names += [f'itg-backup-202610{d:02d}.itgbak' for d in range(20, 25)]
    files = [{'id': n, 'name': n} for n in names]
    assert sorted(f['name'] for f in ob.to_trash(files, today)) == [
        'itg-backup-20250901.itgbak', 'itg-backup-20260715.itgbak']
    # 남는 게 MIN_KEEP 미만이면 정리 안 함 (날짜가 틀어진 경우 다 지우는 사고 방지)
    assert ob.to_trash([{'id': 1, 'name': 'itg-backup-20200101.itgbak'}], today) == []


class _Req:
    def __init__(self, fn):
        self.fn = fn

    def execute(self):
        return self.fn()


class _FakeDrive:
    def __init__(self, existing=()):
        self.files_ = {f['id']: dict(f) for f in existing}
        self.trashed, self.uploaded = [], []
        self.n = 0

    def drives(self):
        return self

    def files(self):
        return self

    def list(self, q=None, **kw):
        if q is None:
            return _Req(lambda: {'drives': [{'id': 'D1', 'name': ob.DRIVE_NAME}, {'id': 'D2', 'name': '다른'}]})
        if 'mimeType' in q:
            return _Req(lambda: {'files': [{'id': 'F1'}]})
        return _Req(lambda: {'files': [f for f in self.files_.values() if f['id'] not in self.trashed]})

    def create(self, body=None, media_body=None, **kw):
        def run():
            self.n += 1
            data = media_body.getbytes(0, media_body.size())
            self.uploaded.append((body, data))
            fid = f'new{self.n}'
            self.files_[fid] = {'id': fid, 'name': body['name']}
            return {'id': fid, 'size': str(len(data))}
        return _Req(run)

    def update(self, fileId=None, body=None, **kw):
        return _Req(lambda: self.trashed.append(fileId))


def test_upload_flow(tmp_path, monkeypatch):
    root = tmp_path / 'b'
    _make_backup(root)
    monkeypatch.setenv('BACKUP_ENCRYPT_PASSWORD', PW)
    monkeypatch.delenv('OFFSITE_BACKUP_FOLDER_ID', raising=False)
    old = [{'id': 'o1', 'name': 'itg-backup-20250101.itgbak'},          # 1월 1일이지만 400일 넘음
           {'id': 'same', 'name': f'itg-backup-{STAMP}.itgbak'}]        # 같은 날 앞선 실행분
    old += [{'id': f'k{d}', 'name': f'itg-backup-202610{d:02d}.itgbak'} for d in range(20, 28)]
    drv = _FakeDrive(old)
    res = ob.upload(root, svc=drv, today=date(2026, 10, 30))
    assert res.startswith('OK: itg-backup-20261030.itgbak')
    body, data = drv.uploaded[0]
    assert body['parents'] == ['F1']
    assert ob.decrypt(data, PW)[:2] == b'\x1f\x8b'                       # gzip
    assert set(drv.trashed) == {'same', 'o1'}
    st = json.loads((root / ob.STATUS_FILE).read_text(encoding='utf-8'))
    assert st['name'] == 'itg-backup-20261030.itgbak' and st['id'] == 'new1'
    assert ob.last_success_age_hours(root) < 0.1


def test_upload_skips_without_password(tmp_path, monkeypatch):
    monkeypatch.delenv('BACKUP_ENCRYPT_PASSWORD', raising=False)
    assert ob.upload(tmp_path, svc=object()).startswith('skip')


def test_upload_size_mismatch_raises(tmp_path, monkeypatch):
    root = tmp_path / 'b'
    _make_backup(root)
    monkeypatch.setenv('BACKUP_ENCRYPT_PASSWORD', PW)
    drv = _FakeDrive()
    orig = drv.create
    drv.create = lambda **kw: _Req(lambda: dict(orig(**kw).execute(), size='1'))
    with pytest.raises(RuntimeError):
        ob.upload(root, svc=drv, today=date(2026, 10, 30))
    assert not (root / ob.STATUS_FILE).exists()                          # 실패면 성공 기록 안 남김


def test_backup_daily_offsite_failure_counts():
    spec = importlib.util.spec_from_file_location(
        'backup_daily', Path(__file__).resolve().parents[2] / 'scripts' / 'backup_daily.py')
    bd = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(bd)
    ok = {'users_db': 'OK: a', 'redis': 'OK: b', 'secrets': 'OK: c'}
    assert bd.failed_items(dict(ok, offsite='skip: BACKUP_ENCRYPT_PASSWORD 없음')) == []
    assert bd.failed_items(dict(ok, offsite='OK: x')) == []
    assert bd.failed_items(dict(ok, offsite='실패: RuntimeError: x')) == ['offsite: 실패: RuntimeError: x']


def test_invariant_offsite_stale(tmp_path):
    from dashboard.services import invariant_checks as ic
    now = time.time()
    assert ic.offsite_stale_hours(tmp_path, now=now, enabled=False) is False
    assert ic.offsite_stale_hours(tmp_path, now=now) is None             # 켜졌는데 기록 없음
    (tmp_path / 'offsite_last.json').write_text(json.dumps({'at': now - 3600}), encoding='utf-8')
    assert ic.offsite_stale_hours(tmp_path, now=now) is False
    (tmp_path / 'offsite_last.json').write_text(json.dumps({'at': now - 30 * 3600}), encoding='utf-8')
    assert ic.offsite_stale_hours(tmp_path, now=now) == 30
    for kind in ('redis', 'users_db', 'secrets'):                       # 로컬 백업은 정상
        (tmp_path / kind).mkdir()
        (tmp_path / kind / 'x').write_text('1')
    vs = ic.check_backup_fresh({'backup_root': tmp_path, 'now_ts': now, 'offsite_enabled': True})
    assert [v['title'] for v in vs] == ['offsite']
