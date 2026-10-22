# -*- coding: utf-8 -*-
"""재시작으로 끊긴 방문 사진 배치는 시작 복구에서 바로 복구 (2026-10-22 L-04320).

재시작 1분 전 시작된 배치를 시작 복구가 '3분 이내 = 진행 중일 수 있음'으로 넘겨 5분 뒤에야
복구됐고, 그사이 '사진 저장 중 (0/9)'로 멈춘 줄 안 매니저가 사진을 다시 올렸다.
이 프로세스 시작 전 배치는 진행 중일 수 없으니 나이와 상관없이 복구, 시작 후 배치만 3분 규칙.
"""
import json
import sys
import time

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb


class _R:
    def __init__(self, batches):
        self.data = {k: json.dumps(v) for k, v in batches.items()}

    def scan_iter(self, pattern, count=200):
        return list(self.data)

    def get(self, k):
        return self.data.get(k)

    def delete(self, k):
        self.data.pop(k, None)


@pytest.fixture
def env(monkeypatch):
    import dashboard.utils.google_drive as gd
    import dashboard.utils.redis_client as rcm
    folder = set()
    uploads = []
    monkeypatch.setattr(gd, 'list_folder_filenames', lambda fid: set(folder))

    def _up(fid, name, content, mimetype=''):
        uploads.append(name)
        folder.add(name)
        return {'id': 'x'}
    monkeypatch.setattr(gd, 'upload_file', _up)
    monkeypatch.setattr('urllib.request.urlopen',
                        lambda req, timeout=30: type('Resp', (), {'read': lambda self: b'img'})())
    monkeypatch.setenv('SLACK_VISIT_BOT_TOKEN', '')
    monkeypatch.setenv('SLACK_BOT_TOKEN', '')

    def _run(created):
        r = _R({'photo_batch:C1:1.1': {'folder_id': 'F', 'channel': 'C1', 'thread_ts': '0.1',
                                       'files': [{'name': 'IMG_1.jpg', 'url': 'https://files.slack.com/u1'},
                                                 {'name': 'IMG_2.jpg', 'url': 'https://files.slack.com/u2'}],
                                       'created': created}})
        monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': r})())
        sb._recover_photo_batches()
        return r
    return _run, uploads


def test_batch_from_before_restart_recovered_immediately(env, monkeypatch):
    run, uploads = env
    monkeypatch.setattr(sb, '_PROC_STARTED_AT', time.time() - 30)   # 30초 전 재시작
    r = run(created=time.time() - 60)                               # 재시작 직전(1분 전) 배치
    assert sorted(uploads) == ['IMG_1.jpg', 'IMG_2.jpg'] and r.data == {}


def test_batch_after_start_waits_three_minutes(env, monkeypatch):
    run, uploads = env
    monkeypatch.setattr(sb, '_PROC_STARTED_AT', time.time() - 3600)
    r = run(created=time.time() - 60)                               # 지금 프로세스가 올리는 중일 수 있음
    assert uploads == [] and list(r.data) == ['photo_batch:C1:1.1']
    run(created=time.time() - 200)                                  # 3분 넘으면 복구
    assert sorted(uploads) == ['IMG_1.jpg', 'IMG_2.jpg']
