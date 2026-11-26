# -*- coding: utf-8 -*-
"""첫 메시지가 사진인 새 채팅 — 카드 스레드에 사진을 올린다 (2026-11-26 L-04704 툴립 548).

예전엔 함수 첫머리에서 읽은 thread_ts(None)를 그대로 써서 '2단계 파일 업로드'를 건너뛰었다.
카드엔 '🖼️ 사진 전송'만 뜨고 도면 사진은 어디에도 없었다.
"""
from dashboard.blueprints import channeltalk as cht


def _run(monkeypatch, card_ok=True):
    store = {}
    uploads = []
    monkeypatch.setattr(cht._threads, 'get_thread_ts', lambda cid: store.get(cid))
    monkeypatch.setattr(cht._threads, 'set_thread_ts', lambda cid, ts: store.__setitem__(cid, ts))
    monkeypatch.setattr(cht._threads, 'add_pending', lambda *a: None)
    monkeypatch.setattr(cht, '_slack_channel', lambda: 'C-CT')
    monkeypatch.setattr(cht, '_get_chat_lead_no', lambda cid: '')
    monkeypatch.setattr(cht, '_save_chat_lead_no', lambda *a: None)
    monkeypatch.setattr(cht, '_kakao_inflow', lambda *a, **k: {})
    monkeypatch.setattr(cht, '_is_spam_message', lambda t: False)
    monkeypatch.setattr(cht, '_register_chat_lead', lambda *a, **k: 'L-04704')
    monkeypatch.setattr(cht, '_attach_chat_open_button', lambda *a, **k: None)
    monkeypatch.setattr(cht, '_new_chat_card', lambda *a, **k: {'text': 'card'})
    monkeypatch.setattr(cht, '_slack_post', lambda m, body, **k: {'ok': True, 'ts': '111.222'} if card_ok else None)
    monkeypatch.setattr(cht.threading, 'Thread', lambda target, daemon=None: type('T', (), {'start': lambda s: None})())
    monkeypatch.setattr(cht, '_slack_upload_files',
                        lambda ch, ts, items, comment: uploads.append((ch, ts, [i['name'] for i in items], comment)) or True)
    import dashboard.services.channeltalk_api as api
    monkeypatch.setattr(api, 'get_file_signed_url', lambda cid, key: 'https://signed')
    monkeypatch.setattr(api, 'download_file', lambda url: b'jpeg-bytes')
    entity = {'chatId': 'CH1', 'plainText': '', 'createdAt': 1795674726000,
              'files': [{'id': 'f1', 'key': 'pri-file/x/i_4a6312fd6a0e.jpeg', 'type': 'image',
                         'name': 'i_4a6312fd6a0e.jpeg', 'contentType': 'image/jpeg'}]}
    cht._handle_user_message_locked(entity, {}, {'id': 'CH1', 'name': '툴립 548'}, {'name': '툴립 548'}, 'CH1')
    return uploads


def test_first_message_photo_goes_to_new_card_thread(monkeypatch):
    uploads = _run(monkeypatch)
    assert len(uploads) == 1
    ch, ts, names, _ = uploads[0]
    assert (ch, ts, names) == ('C-CT', '111.222', ['i_4a6312fd6a0e.jpeg'])


def test_card_failed_no_upload_but_no_crash(monkeypatch):
    # 카드가 아직 없으면(재시도 중) 올릴 스레드가 없다 — 조용히 버리지 않고 ERROR 로그만, 예외 없음
    assert _run(monkeypatch, card_ok=False) == []
