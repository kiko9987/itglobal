# -*- coding: utf-8 -*-
"""재시작으로 끊긴 방문 사진 배치를 복구가 끝까지 마무리 (2026-10-22 L-04307).

사진 저장 중 재시작 → 복구가 남은 사진은 올렸지만 답글이 '저장 중 (5/10)'으로 남고 방문 완료도
안 됐다. 복구가 ① 최종 답글 ② 스레드→폴더 매핑 ③ 자동 방문 완료까지 하는지,
저장 루프가 이미 마친 배치(done)는 예전처럼 실패 표기 정리만 하는지 확인.
"""
import json
import sys
import time

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb

FILES = [{'name': f'2026102{i}.jpg', 'url': f'https://files.slack.com/f{i}', 'mime': 'image/jpeg'}
         for i in range(10)]


class _R:
    def __init__(self, batch, reply_hash=None, locks=()):
        self.data = {'photo_batch:C1:9.9': json.dumps(batch)}
        self.hashes = {'visit_photo_reply:C1:1.1': dict(reply_hash or {})}
        self.locks = set(locks)

    def scan_iter(self, pattern, count=200):
        return [k for k in self.data if k.startswith('photo_batch:')]

    def get(self, k):
        return self.data.get(k)

    def set(self, k, v, ex=None, **kw):
        self.data[k] = v

    def delete(self, k):
        self.data.pop(k, None)

    def exists(self, k):
        return k in self.locks or k in self.hashes

    def hgetall(self, k):
        return dict(self.hashes.get(k) or {})

    def hset(self, k, mapping):
        self.hashes.setdefault(k, {}).update(mapping)

    def expire(self, k, t):
        pass


class _Client:
    log = []

    def __init__(self, token=None):
        pass

    def chat_update(self, **k):
        _Client.log.append(('update', k))

    def chat_postMessage(self, **k):
        _Client.log.append(('post', k))
        return {'ts': '5.5'}

    def conversations_replies(self, **k):
        return {'messages': [{'text': '', 'blocks': [{'type': 'section', 'text': {
            'type': 'mrkdwn', 'text': ':bell: *새 방문 일정* `L-04307`'}}]}]}


def _batch(**kw):
    b = {'channel': 'C1', 'thread_ts': '1.1', 'folder_id': 'F', 'reply_ts': '2.2', 'files': FILES,
         'created': time.time() - 3600, 'lead_no': 'L-04307', 'user': 'U_MJ',
         'lead_folder_id': 'LF', 'photo_folder_id': 'PF', 'folder_name': '(MJ) 성남 크란츠테크노 26.10.22',
         'location': ''}
    b.update(kw)
    return b


@pytest.fixture
def env(monkeypatch):
    import dashboard.utils.google_drive as gd
    import dashboard.utils.redis_client as rcm
    import slack_sdk
    state = {'folder': {f['name'] for f in FILES}, 'uploads': [], 'completed': [], 'cleared': [],
             'fail_list': False}

    def _list(fid, strict=False):
        if state['fail_list']:
            raise RuntimeError('HttpError 500')
        return set(state['folder'])

    def _up(fid, name, content, mimetype=''):
        state['uploads'].append(name)
        state['folder'].add(name)
        return {'id': 'x'}
    monkeypatch.setattr(gd, 'list_folder_filenames', _list)
    monkeypatch.setattr(gd, 'upload_file', _up)
    monkeypatch.setattr('urllib.request.urlopen',
                        lambda req, timeout=30: type('Resp', (), {'read': lambda self: b'img'})())
    monkeypatch.setattr(slack_sdk, 'WebClient', _Client)
    monkeypatch.setattr(sb, '_visit_folder_location_lines', lambda name, fid: '')
    monkeypatch.setattr(sb, '_auto_complete_visit_from_photo',
                        lambda client, channel, thread_ts, lead_no, event_user_id:
                        state['completed'].append((lead_no, event_user_id)))
    monkeypatch.setattr(sb, '_clear_stale_failure_note', lambda b, rc, token: state['cleared'].append(1))
    monkeypatch.setenv('SLACK_VISIT_BOT_TOKEN', 'xoxb-visit')
    _Client.log = []

    def run(batch, reply_hash=None, locks=()):
        r = _R(batch, reply_hash, locks)
        monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': r})())
        sb._recover_photo_batches()
        return r
    return run, state


def test_interrupted_batch_is_finished(env):
    run, st = env
    r = run(_batch(), reply_hash={'ts': '2.2'})          # 사진은 다 있음, 답글은 '저장 중 (5/10)'
    kind, args = _Client.log[0]
    assert kind == 'update' and args['ts'] == '2.2'
    assert args['text'].startswith(':file_folder: 사진 10장을 드라이브에 저장했습니다.')
    assert '(MJ) 성남 크란츠테크노 26.10.22' in args['text']
    assert r.hashes['visit_photo_reply:C1:1.1']['count'] == '10'
    assert r.hashes['visit_thread:C1:1.1']['lead_folder_id'] == 'LF'
    assert st['completed'] == [('L-04307', 'U_MJ')] and st['cleared'] == []
    assert r.data == {}


def test_recovered_count_shown(env):
    run, st = env
    st['folder'] -= {FILES[6]['name'], FILES[7]['name'], FILES[8]['name'], FILES[9]['name']}
    run(_batch())
    assert len(st['uploads']) == 4
    text = _Client.log[0][1]['text']
    assert '서버 재시작으로 끊겼던 4장은 자동으로 이어서 저장했습니다' in text
    assert st['completed'] == [('L-04307', 'U_MJ')]


def test_done_batch_keeps_old_behavior(env):
    run, st = env
    r = run(_batch(done=True))
    assert st['cleared'] == [1] and st['completed'] == [] and _Client.log == [] and r.data == {}


def test_replied_batch_completes_without_recount(env):
    run, st = env
    r = run(_batch(replied=True), reply_hash={'ts': '2.2', 'count': '10'})
    assert _Client.log == [] and r.hashes['visit_photo_reply:C1:1.1']['count'] == '10'
    assert st['completed'] == [('L-04307', 'U_MJ')]


def test_old_batch_reads_lead_from_card(env):
    run, st = env
    b = _batch()
    for k in ('lead_no', 'user', 'lead_folder_id', 'photo_folder_id', 'folder_name', 'location'):
        b.pop(k)
    run(b)
    assert st['completed'] == [('L-04307', '')]
    assert _Client.log[0][1]['text'] == ':file_folder: 사진 10장을 드라이브에 저장했습니다.'


def test_capture_only_batch_not_completed(env):
    run, st = env
    png = [{'name': 'image.png', 'url': 'https://files.slack.com/p', 'mime': 'image/png'}]
    st['folder'] = {'image.png'}
    run(_batch(files=png))
    assert st['completed'] == []


def test_list_failure_does_not_reupload(env):
    run, st = env
    st['fail_list'] = True
    r = run(_batch())
    assert st['uploads'] == [] and st['completed'] == [] and list(r.data) == ['photo_batch:C1:9.9']


def test_in_process_batch_with_lock_skipped(env, monkeypatch):
    run, st = env
    monkeypatch.setattr(sb, '_PROC_STARTED_AT', time.time() - 7200)
    r = run(_batch(created=time.time() - 600), locks={'visit_photo_lock:L-04307'})
    assert st['completed'] == [] and list(r.data) == ['photo_batch:C1:9.9']
