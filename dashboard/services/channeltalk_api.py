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


import re as _re

# 2026-07-10 CT2: 채널톡 API 가 이메일 주소 포함 메시지를 종종 거부하는 이슈
# (사용자 관측: 매니저가 이메일 주소 답변 시 전송 실패)
# 실패 케이스에서 이메일을 전각 골뱅이(＠)로 자동 치환 후 재시도.
_EMAIL_RE = _re.compile(r'\b([A-Za-z0-9._%+-]+)@([A-Za-z0-9.-]+\.[A-Za-z]{2,})\b')


def _mask_emails(text: str) -> str:
    """이메일 주소의 @ 를 전각 골뱅이(＠) 로 치환.

    고객 눈에는 여전히 이메일로 보이지만 채널톡 API 의 자동 감지·차단은 회피.
    """
    return _EMAIL_RE.sub(r'\1＠\2', text or '')


def _contains_email(text: str) -> bool:
    return bool(_EMAIL_RE.search(text or ''))


# 2026-07-24: 슬랙이 event.text 안 URL/전화/이메일을 자동으로 mrkdwn 링크 문법으로
# 감쌈 (`<tel:010-...|010-...>`, `<mailto:foo@bar|foo@bar>`, `<http://...|display>`).
# 채널톡 API blocks 파서는 이 문법을 못 읽어 HTTP 422 (`block value not parsable`) 반환.
# → 링크 문법을 plain text 로 unescape 후 전송.
_SLACK_LINK_RE = _re.compile(r'<([^>|]+)(?:\|([^>]+))?>')


def _unescape_slack_markup(text: str) -> str:
    """슬랙 mrkdwn 링크 문법 → plain text.

    - `<tel:010-...|010-...>` → `010-...`
    - `<mailto:foo@bar|foo@bar>` → `foo@bar`
    - `<http://x.com|display>` → `display`
    - `<http://x.com>` → `http://x.com`
    - `<@USERID>`, `<#CHAN|name>`, `<!subteam...>` 등 mention 은 raw 유지 (드묾).
    - HTML escape (`&lt;`, `&gt;`, `&amp;`) 해제.
    """
    if not text:
        return text

    def _sub(m):
        target = m.group(1)
        display = m.group(2) or ''
        if target.startswith(('tel:', 'mailto:', 'http://', 'https://')):
            if display:
                return display
            # tel:/mailto: 는 prefix 제거, http(s):// 는 URL 유지
            if target.startswith('tel:'):
                return target[4:]
            if target.startswith('mailto:'):
                return target[7:]
            return target
        return m.group(0)  # mention 등 — raw 유지

    text = _SLACK_LINK_RE.sub(_sub, text)
    text = text.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
    return text


def send_manager_message(chat_id: str, manager_id: str, plain_text: str) -> Optional[dict]:
    """매니저(운영자)가 채팅방에 메시지 발신.

    Args:
        chat_id: userChat.id
        manager_id: 매니저 ID (.env CHANNELTALK_OPERATOR_ID)
        plain_text: 보낼 메시지 본문

    채널톡 입장에서 항상 같은 매니저(kiko=60994)가 보낸 것으로 기록 → 운영자 시트 추가 X.

    Returns:
        성공: 응답 dict (`email_auto_escaped=True` 가 자동 치환 케이스임을 표시)
        실패: None
    """
    if not chat_id or not manager_id or not plain_text:
        return None

    # 2026-07-24: 슬랙 mrkdwn 링크 문법 unescape (전화·이메일·URL 자동 링크화 대응).
    # 채널톡 API 는 `<tel:...>`, `<mailto:...>` 문법을 못 파싱해 HTTP 422 반환.
    plain_text = _unescape_slack_markup(plain_text)

    body = {
        'managerId': manager_id,
        'blocks': [{'type': 'text', 'value': plain_text}],
        'plainText': plain_text,
    }
    resp = _request('POST', f'/user-chats/{chat_id}/messages', body=body)
    if resp is not None:
        return resp

    # 재시도: 이메일 포함 케이스면 자동 치환 후 재요청 (2026-07-10 CT2)
    if _contains_email(plain_text):
        escaped = _mask_emails(plain_text)
        logger.warning(
            f'[ChannelTalk] 최초 전송 실패 + 이메일 감지 → 전각 골뱅이 치환 후 재시도 '
            f'(chat_id={chat_id})'
        )
        body_retry = {
            'managerId': manager_id,
            'blocks': [{'type': 'text', 'value': escaped}],
            'plainText': escaped,
        }
        resp = _request('POST', f'/user-chats/{chat_id}/messages', body=body_retry)
        if resp is not None:
            # 성공 응답에 자동 치환 표시 (호출자가 매니저에게 안내 가능)
            resp['_email_auto_escaped'] = True
            return resp

    return None


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
