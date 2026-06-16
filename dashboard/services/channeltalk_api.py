"""
채널톡 Open API v5 클라이언트 (양방향 슬랙 통합용)

핵심 기능:
- 매니저(운영자)가 채팅방에 메시지 발신
- 채팅방 정보 조회 (state, 고객 정보)
- 매니저 배정/종료 (다음 단계)

인증: x-access-key + x-access-secret 헤더
"""

import json
import os
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Dict, Optional  # noqa: F401

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

BASE_URL = 'https://api.channel.io/open/v5'


def _headers() -> Dict[str, str]:
    return {
        'x-access-key': os.getenv('CHANNELTALK_ACCESS_KEY', '').strip(),
        'x-access-secret': os.getenv('CHANNELTALK_ACCESS_SECRET', '').strip(),
        'Content-Type': 'application/json',
    }


def _request(method: str, path: str, body: Optional[dict] = None) -> Optional[dict]:
    """채널톡 API 호출 공용 래퍼"""
    url = BASE_URL + path
    data = json.dumps(body).encode('utf-8') if body else None
    req = urllib.request.Request(url, data=data, headers=_headers(), method=method)
    try:
        with urllib.request.urlopen(req, timeout=10) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as exc:
        body_text = exc.read().decode(errors='replace')[:300]
        logger.warning(f'[ChannelTalk] {method} {path} → HTTP {exc.code}: {body_text}')
        return None
    except Exception as exc:
        logger.warning(f'[ChannelTalk] {method} {path} → {type(exc).__name__}: {exc}')
        return None


def send_manager_message(chat_id: str, manager_id: str, plain_text: str) -> Optional[dict]:
    """매니저(운영자)가 채팅방에 메시지 발신.

    Args:
        chat_id: userChat.id
        manager_id: 매니저 ID (.env CHANNELTALK_OPERATOR_ID)
        plain_text: 보낼 메시지 본문

    채널톡 입장에서 항상 같은 매니저(kiko=60994)가 보낸 것으로 기록 → 운영자 시트 추가 X.
    """
    if not chat_id or not manager_id or not plain_text:
        return None
    body = {
        'managerId': manager_id,
        'blocks': [{'type': 'text', 'value': plain_text}],
        'plainText': plain_text,
    }
    return _request('POST', f'/user-chats/{chat_id}/messages', body=body)


def get_user_chat(chat_id: str) -> Optional[dict]:
    """채팅방 정보 조회 (state/배정 상태/고객 정보)"""
    if not chat_id:
        return None
    return _request('GET', f'/user-chats/{chat_id}')


def assign_user_chat(chat_id: str, manager_id: str) -> Optional[dict]:
    """채팅방을 매니저에게 배정. 채널톡 자체 '나에게 배정' 동작과 동일."""
    if not chat_id or not manager_id:
        return None
    return _request('PUT', f'/user-chats/{chat_id}/assignee', body={'assigneeId': manager_id})


def close_user_chat(chat_id: str) -> Optional[dict]:
    """채팅방 종료 (해결됨 처리)."""
    if not chat_id:
        return None
    return _request('PUT', f'/user-chats/{chat_id}/close', body={'closeMessage': ''})


def get_file_signed_url(chat_id: str, file_key: str) -> Optional[str]:
    """채널톡 private 파일의 signed URL 받기 (15분 만료, CloudFront 경유).

    Args:
        chat_id: userChat ID
        file_key: file 객체의 key 필드 (예: pri-file/.../i_xxx)

    Returns:
        signed URL or None (실패 시)
    """
    if not chat_id or not file_key:
        return None
    path = f'/user-chats/{chat_id}/messages/file?' + urllib.parse.urlencode({'key': file_key})
    resp = _request('GET', path)
    if resp and isinstance(resp, dict):
        return resp.get('result')
    return None


def download_file(signed_url: str) -> Optional[bytes]:
    """signed URL에서 파일 바이너리 다운로드 (작은 파일 한정)."""
    if not signed_url:
        return None
    try:
        req = urllib.request.Request(signed_url)
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.read()
    except Exception as exc:
        logger.warning(f'[ChannelTalk] 파일 다운로드 실패: {exc}')
        return None
