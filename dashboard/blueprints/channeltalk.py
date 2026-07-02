"""
채널톡 ↔ 슬랙 양방향 통합 — 2일차

흐름:
- 채널톡 webhook 수신 → 사용자 메시지면 슬랙 채널에 카드/thread reply
- chatId ↔ slack thread_ts 매핑은 Redis에 저장 (channeltalk_threads.py)

엔드포인트:
- POST /channeltalk/events  : 메인 webhook
- GET  /channeltalk/health  : 헬스 체크
"""

import json
import os
import threading
import urllib.parse
import urllib.request
from datetime import datetime
from typing import Any, Dict, Optional  # noqa: F401

from flask import Blueprint, request, jsonify

from dashboard.services import channeltalk_threads as _threads
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


# chat_id 별 처리 직렬화 lock — 사진 + 텍스트가 거의 동시에 webhook으로 들어올 때
# 첫 webhook의 thread_ts 저장 전에 두 번째가 조회해서 둘 다 새 lead로 등록되는 race 차단
_CHAT_LOCKS: Dict[str, threading.Lock] = {}
_CHAT_LOCKS_MUTEX = threading.Lock()


def _get_chat_lock(chat_id: str) -> threading.Lock:
    """chat_id별 lock 반환 (없으면 생성)."""
    with _CHAT_LOCKS_MUTEX:
        if chat_id not in _CHAT_LOCKS:
            _CHAT_LOCKS[chat_id] = threading.Lock()
        return _CHAT_LOCKS[chat_id]

channeltalk_bp = Blueprint('channeltalk', __name__, url_prefix='/channeltalk')


# ─────────────────────────────────────────────────────────────
# 인입 채널(medium) → 라벨 + 헤더 아이콘
# ─────────────────────────────────────────────────────────────
MEDIUM_LABELS = {
    'appKakao': '카카오톡',
    'web': '채널톡',
    'inAppChat': '채널톡',
    'mobile': '모바일',
    'email': '이메일',
    'mobileApp': '앱',
}

# 채널별 헤더 아이콘 (워크스페이스 커스텀 이모지 사용)
MEDIUM_ICONS = {
    'appKakao': ':카카오톡:',
    'web': ':채널톡:',
    'inAppChat': ':채널톡:',
    'mobile': ':iphone:',
    'email': ':email:',
    'mobileApp': ':iphone:',
}


def _slack_channel() -> str:
    return os.getenv('SLACK_CHANNELTALK_CHANNEL', '').strip()


def _slack_token() -> str:
    return os.getenv('SLACK_BOT_TOKEN', '').strip()


# ─────────────────────────────────────────────────────────────
# Slack API 헬퍼
# ─────────────────────────────────────────────────────────────
def _slack_post(api_path: str, body: dict) -> Optional[dict]:
    token = _slack_token()
    if not token:
        return None
    try:
        data = json.dumps(body).encode('utf-8')
        req = urllib.request.Request(
            f'https://slack.com/api/{api_path}',
            data=data,
            headers={
                'Content-Type': 'application/json; charset=utf-8',
                'Authorization': f'Bearer {token}',
            },
        )
        with urllib.request.urlopen(req, timeout=5) as r:
            resp = json.loads(r.read())
            if not resp.get('ok'):
                logger.warning(f'[ChannelTalk/Slack] {api_path} 실패: {resp.get("error")}')
            return resp
    except Exception as exc:
        logger.warning(f'[ChannelTalk/Slack] {api_path} 예외: {exc}')
        return None


# ─────────────────────────────────────────────────────────────
# 슬랙 메시지 양식
# ─────────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────
# 우리 봇이 슬랙→채널톡으로 forward한 메시지 캐시 (echo loop 방지)
# webhook으로 같은 메시지가 되돌아오면 skip
# ─────────────────────────────────────────────────────────────
_OUR_SENT_CACHE: dict = {}  # (chat_id, text) → epoch sec
_OUR_SENT_TTL = 60  # 초


def mark_our_sent(chat_id: str, text: str) -> None:
    """슬랙→채널톡 forward 직후 호출 — echo loop 방지용 캐시."""
    if not chat_id or not text:
        return
    import time as _t
    now = _t.time()
    _OUR_SENT_CACHE[(chat_id, text.strip())] = now
    # 오래된 항목 정리
    cutoff = now - _OUR_SENT_TTL * 2
    for key in list(_OUR_SENT_CACHE):
        if _OUR_SENT_CACHE[key] < cutoff:
            del _OUR_SENT_CACHE[key]


def is_our_sent(chat_id: str, text: str) -> bool:
    """webhook 도착 시 — 우리가 방금 forward한 메시지인지 확인. 일치 시 True + 캐시 제거."""
    if not chat_id or not text:
        return False
    import time as _t
    key = (chat_id, text.strip())
    if key in _OUR_SENT_CACHE:
        ts = _OUR_SENT_CACHE.pop(key)
        if _t.time() - ts < _OUR_SENT_TTL:
            return True
    return False


# (시간 포맷 함수: channeltalk_helpers.py로 이동)
from dashboard.blueprints.channeltalk_helpers import (
    _format_ts, _format_ts_full, _is_spam_message,
    _SPAM_KEYWORDS, _URL_RE,
)


def _attach_chat_open_button(channel: str, thread_ts: str, original_text: str,
                             lead_no: str = '', chat_id: str = '') -> None:
    """카드 발송 후 thread permalink로 이동하는 [💬 채팅 열기] URL 버튼 부착.

    카카오톡 채팅 인입은 thread에서 직접 응대 → 통합 모달이 아니라 thread 이동.
    unfurl_links/media를 명시적으로 False — 슬랙이 permalink를 자동 unfurl하면
    본문이 quoted preview 안에 갇혀 클릭 동작도 실패함.
    """
    try:
        token = _slack_token()
        if not token:
            return
        # thread permalink 조회 (카드 자체가 thread root이므로 ts=thread_ts)
        permalink_url = (
            f'https://slack.com/api/chat.getPermalink'
            f'?channel={channel}&message_ts={thread_ts}'
        )
        req = urllib.request.Request(
            permalink_url,
            headers={'Authorization': f'Bearer {token}'},
        )
        with urllib.request.urlopen(req, timeout=5) as r:
            resp = json.loads(r.read())
        permalink = (resp or {}).get('permalink', '')
        if not permalink:
            return

        # thread_ts + cid 파라미터 추가 — 슬랙이 사이드 패널로 thread 펼침
        # (permalink는 root 메시지 URL이라 그대로 두면 같은 페이지 navigation으로 끝남)
        sep = '&' if '?' in permalink else '?'
        thread_url = f'{permalink}{sep}thread_ts={thread_ts}&cid={channel}'

        # 본문 끝에 [💬 채팅 열기] 마크다운 링크 (thread 이동용) + 빈 줄로 간격
        # 그 아래 actions block에 [📞 상담하기] action button (통합 모달 — 같은 lead 업데이트)
        updated_text = original_text + f"\n\n<{thread_url}|*💬 채팅 열기*>\n⠀"
        _slack_post('chat.update', {
            'channel': channel,
            'ts': thread_ts,
            'text': updated_text,
            'unfurl_links': False,
            'unfurl_media': False,
            'blocks': [
                {
                    'type': 'section',
                    'text': {'type': 'mrkdwn', 'text': updated_text},
                },
                {
                    'type': 'actions',
                    'elements': [
                        {
                            'type': 'button',
                            'text': {'type': 'plain_text', 'text': ':point_right: 상담하기', 'emoji': True},
                            'style': 'primary',
                            'value': lead_no or chat_id or '-',
                            'action_id': 'button_consult',
                        },
                        {
                            'type': 'button',
                            'text': {'type': 'plain_text', 'text': ':link: 기존 lead 연결', 'emoji': True},
                            'value': chat_id or '-',
                            'action_id': 'link_existing_lead',
                        },
                    ],
                },
                {'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': '⠀'}]},
            ],
        })
    except Exception as exc:
        logger.warning(f'[ChannelTalk] 채팅 열기 버튼 추가 실패: {exc}')


def _new_chat_card(user_chat: dict, user: dict, first_message: str,
                   created_ms: int, lead_no: str = '') -> dict:
    """새 채팅 시작 시 표시할 카드 — 다른 인입 카드(당근/홈페이지)와 동일 양식.

    양식:
        접수번호: L-XXXXX
        :카카오톡: 새 문의 접수 알림 - 카카오톡
        --------------------------------------------
        문의시간 : 2026.06.22. 09:32
        이름 / 상호 : 멜론 518
        연락처 : -
        상세 문의 내용 : 저희 회사가 ...
        --------------------------------------------
    """
    customer_name = user.get('name') or user_chat.get('name') or '익명 고객'
    medium = (user_chat.get('mediumProfile') or {}).get('mediumName', '')
    medium_label = MEDIUM_LABELS.get(medium) if medium else '채널톡'
    medium_icon = MEDIUM_ICONS.get(medium) if medium else ':채널톡:'
    if not medium_label:
        medium_label = medium or '채널톡'
    if not medium_icon:
        medium_icon = ':채널톡:'
    ts_str = _format_ts_full(created_ms)
    SEP = '--------------------------------------------'
    # 슬랙 본문 3000자 한도 — 메시지 본문 truncate (스팸 헤더/메타 여유)
    if first_message and len(first_message) > 2400:
        first_message = first_message[:2400] + '\n…(내용이 길어 일부만 표시 — 채널톡 thread 참조)'

    lines = ['⠀']
    if lead_no:
        lines.append(f'접수번호: `{lead_no}`')
    lines.extend([
        f'{medium_icon} *새 문의 접수 알림 - {medium_label}*',
        SEP,
        f'문의시간 : {ts_str}',
        f'이름 / 상호 : {customer_name}',
        f'연락처 : -',
        f'상세 문의 내용 : {first_message}',
        SEP,
    ])
    header_text = '\n'.join(lines)

    return {'text': header_text}


# (스팸 감지 함수: channeltalk_helpers.py로 이동 — 위에서 import)


def _register_chat_lead(user_chat: dict, user: dict, first_message: str,
                        created_ms: int) -> str:
    """새 채팅 인입을 메인 시트에 1건 lead로 등록 (통계용).

    Returns: lead_no (실패 시 빈 문자열)
    """
    try:
        from dashboard.services.lead_sync import _append_leads_to_main
        from datetime import datetime

        customer_name = (user.get('name') or user_chat.get('name')
                         or '익명 고객').strip()
        medium = (user_chat.get('mediumProfile') or {}).get('mediumName', '')
        platform = MEDIUM_LABELS.get(medium) if medium else '채널톡'
        if not platform:
            platform = medium or '채널톡'

        try:
            dt = datetime.fromtimestamp(created_ms / 1000.0)
            consult_time = dt.strftime('%Y.%m.%d. %H:%M')
        except Exception:
            dt = datetime.now()
            consult_time = dt.strftime('%Y.%m.%d. %H:%M')

        # 본문 길이 제한 (시트 한 셀에 너무 길면 가독성 ↓)
        inquiry = (first_message or '').strip()
        if len(inquiry) > 500:
            inquiry = inquiry[:500] + '...'
        if not inquiry:
            inquiry = '-'

        lead = {
            '리드 No': '',
            '상담 시간': consult_time,
            '플랫폼': platform,
            '상태': '상담 대기',
            '방문 예정일': '-',
            '고객 연락처': '-',
            '이메일': '-',
            '고객명': customer_name,
            '방문 주소': '-',
            '문의 내용': inquiry,   # 채팅 인입 원본 = 첫 메시지
            '상담 내용': '',        # 매니저 처리 결과 (초기엔 빈값)
            '키워드': '-',
            '온라인 상담자': '',
            '영업 담당자': '',
            '마지막 연락일': '',
            '_meta_place': '-',
            '_meta_device': '-',
            '_meta_inquiry': inquiry,
            '_meta_consult_dt': dt,
            '_meta_address_level': '',
        }
        lead_nos = _append_leads_to_main([lead])
        return lead_nos[0] if lead_nos else ''
    except Exception as exc:
        logger.warning(f'[ChannelTalk] 시트 등록 실패 (lead은 기록 안 됨): {exc}')
        return ''


def _thread_reply_text(plain_text: str, customer_name: str, created_ms: int) -> str:
    ts_str = _format_ts(created_ms)
    return f':bust_in_silhouette: *{customer_name}* _{ts_str}_\n{plain_text}'


def _thread_reply_text_manager(text: str, manager_name: str, created_ms: int) -> str:
    """매니저(채널톡 직원)가 채널톡 앱에서 직접 답변한 메시지 — thread reply 양식."""
    ts_str = _format_ts(created_ms)
    return f':technologist: *{manager_name}* (채널톡 답변) _{ts_str}_\n{text}'


def _extract_files(entity: dict) -> list:
    """채널톡 entity.files에서 파일 메타데이터 추출.

    Returns:
        [{'id', 'name', 'key', 'type', 'is_image', 'size'}, ...]
    """
    out = []
    for f in (entity.get('files') or []):
        if not f.get('key'):
            continue
        name = f.get('name') or 'file'
        ftype = (f.get('type') or '').lower()
        ct = (f.get('contentType') or '').lower()
        is_image = (
            ftype == 'image'
            or ct.startswith('image/')
            or any(name.lower().endswith(ext) for ext in
                   ('.jpg', '.jpeg', '.png', '.gif', '.webp', '.heic'))
        )
        out.append({
            'id': f.get('id'),
            'name': name,
            'key': f['key'],
            'type': ftype,
            'is_image': is_image,
            'size': f.get('size', 0),
        })
    return out


def _slack_upload_files(channel: str, thread_ts: str,
                        items: list,  # [{'content', 'name'}, ...]
                        initial_comment: str = '') -> bool:
    """슬랙 files.upload_v2로 여러 파일을 한 번에 thread에 업로드 (1번 알림).

    Returns: 전체 성공 여부.
    """
    token = _slack_token()
    if not token or not items:
        return False
    try:
        file_specs = []
        for it in items:
            content = it.get('content')
            name = it.get('name') or 'file'
            if not content:
                continue
            # Step 1: getUploadURLExternal (파일별로)
            params = urllib.parse.urlencode({'filename': name, 'length': len(content)})
            req = urllib.request.Request(
                f'https://slack.com/api/files.getUploadURLExternal?{params}',
                headers={'Authorization': f'Bearer {token}'},
            )
            with urllib.request.urlopen(req, timeout=10) as r:
                data = json.loads(r.read())
            if not data.get('ok'):
                logger.warning(f'[ChannelTalk/Slack] getUploadURL 실패 ({name}): {data.get("error")}')
                continue
            upload_url = data['upload_url']
            file_id = data['file_id']

            # Step 2: POST 바이너리
            req = urllib.request.Request(upload_url, data=content, method='POST',
                                         headers={'Content-Type': 'application/octet-stream'})
            with urllib.request.urlopen(req, timeout=30) as r:
                r.read()
            file_specs.append({'id': file_id, 'title': name})

        if not file_specs:
            return False

        # Step 3: completeUploadExternal — 여러 파일 묶음으로 한 번에 완료 (알림 1회)
        complete_body = {
            'files': file_specs,
            'channel_id': channel,
            'thread_ts': thread_ts,
        }
        if initial_comment:
            complete_body['initial_comment'] = initial_comment
        req = urllib.request.Request(
            'https://slack.com/api/files.completeUploadExternal',
            data=json.dumps(complete_body).encode('utf-8'),
            headers={
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json; charset=utf-8',
            },
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            resp = json.loads(r.read())
        if not resp.get('ok'):
            logger.warning(f'[ChannelTalk/Slack] completeUpload 실패: {resp.get("error")}')
            return False
        return True
    except Exception as exc:
        logger.warning(f'[ChannelTalk/Slack] files.upload 예외: {exc}')
        return False


# ─────────────────────────────────────────────────────────────
# 핸들러
# ─────────────────────────────────────────────────────────────
def _register_pending_lead_if_any(chat_id: str) -> str:
    """매니저 첫 응답 시점에 Redis pending lead → 시트 등록.
    이미 등록됐으면 (Redis 키 없음) skip. 등록 성공 시 lead_no 반환, 실패 시 빈 문자열.

    SETNX 락 — 동시 응답 시 중복 등록 방지.
    """
    if not chat_id:
        return ''
    try:
        from dashboard.utils.redis_client import get_redis_client
        import json as _json
        rc = get_redis_client().redis
        # SETNX 락 (60초 TTL) — 동시 호출 시 한 번만 등록
        lock_key = f'channeltalk_lead_lock:{chat_id}'
        if not rc.set(lock_key, '1', nx=True, ex=60):
            logger.debug(f'[ChannelTalk] lead 등록 락 잡힘 — skip (chat_id={chat_id})')
            return ''
        pending_key = f'channeltalk_pending_lead:{chat_id}'
        pending_raw = rc.get(pending_key)
        if not pending_raw:
            return ''
        pending = _json.loads(
            pending_raw.decode('utf-8') if isinstance(pending_raw, bytes) else pending_raw
        )
        fake_user_chat = {
            'name': pending.get('user_name', ''),
            'mediumProfile': {'mediumName': pending.get('medium', '')},
        }
        fake_user = {'name': pending.get('user_name', '')}
        lead_no_new = _register_chat_lead(
            fake_user_chat, fake_user,
            pending.get('first_message', ''),
            pending.get('created_ms', 0),
        )
        if lead_no_new:
            rc.delete(pending_key)
            logger.info(
                f'[ChannelTalk] 매니저 응답 → lead 등록 (chat_id={chat_id}, lead_no={lead_no_new})'
            )
            # 슬랙 thread에 등록 완료 안내
            channel = _slack_channel()
            thread_ts = _threads.get_thread_ts(chat_id)
            if channel and thread_ts:
                try:
                    _slack_post('chat.postMessage', {
                        'channel': channel,
                        'thread_ts': thread_ts,
                        'text': f':white_check_mark: 시트 등록 완료 — `{lead_no_new}`',
                        'unfurl_links': False,
                    })
                except Exception:
                    pass
            return lead_no_new
    except Exception as exc:
        logger.warning(f'[ChannelTalk] pending lead 등록 실패 (chat_id={chat_id}): {exc}')
    return ''


def _handle_manager_message(payload: dict) -> None:
    """매니저(own_operator 외 채널톡 직원)가 채널톡 앱에서 직접 답변한 메시지를
    슬랙 thread reply로 forward + 미응답 큐에서 제거.
    """
    entity = payload.get('entity') or {}
    refers = payload.get('refers') or {}
    user_chat = refers.get('userChat') or {}
    manager = refers.get('manager') or {}

    chat_id = entity.get('chatId') or user_chat.get('id')
    plain_text = (entity.get('plainText') or '').strip()
    created_ms = entity.get('createdAt') or 0
    manager_name = (manager.get('name') or '매니저').strip() or '매니저'

    files = _extract_files(entity)
    if not chat_id or (not plain_text and not files):
        return

    channel = _slack_channel()
    if not channel:
        return

    thread_ts = _threads.get_thread_ts(chat_id)
    if not thread_ts:
        # thread 매핑 없음 — 매니저가 우리 카드 생성 전에 답변한 드문 케이스
        logger.debug(
            f'[ChannelTalk] manager 답변 — thread 없음, skip (chat_id={chat_id})'
        )
        return

    display_text = plain_text or (
        '🖼️ 사진' if any(f['is_image'] for f in files) else '📎 파일'
    )
    reply_text = _thread_reply_text_manager(display_text, manager_name, created_ms)

    if not files:
        _slack_post('chat.postMessage', {
            'channel': channel,
            'thread_ts': thread_ts,
            'text': reply_text,
            'unfurl_links': False,
        })
    else:
        # 파일 있을 때 — files.upload_v2로 thread에 첨부 (initial_comment=reply_text)
        from dashboard.services.channeltalk_api import get_file_signed_url, download_file
        items = []
        for f in files:
            try:
                signed = get_file_signed_url(chat_id, f['key'])
                if not signed:
                    continue
                content = download_file(signed)
                if content:
                    items.append({'content': content, 'name': f['name']})
            except Exception as exc:
                logger.warning(f'[ChannelTalk/mgr] 파일 처리 실패: {exc}')
        if items:
            _slack_upload_files(channel, thread_ts, items, reply_text)
        else:
            _slack_post('chat.postMessage', {
                'channel': channel,
                'thread_ts': thread_ts,
                'text': reply_text,
                'unfurl_links': False,
            })

    # 매니저가 답변했으니 미응답 큐에서 제거
    _threads.remove_pending(chat_id)

    # 첫 매니저 응답 시점 — 시트에 lead 등록 (스팸 차단 효과)
    _register_pending_lead_if_any(chat_id)

    # 원본 카드(thread_ts)에 ✅ — 다른 영업도 "처리됨" 한눈에 확인
    try:
        _slack_post('reactions.add', {
            'channel': channel,
            'timestamp': thread_ts,
            'name': 'white_check_mark',
        })
    except Exception as exc:
        logger.debug(f'[ChannelTalk] reaction 추가 스킵 (chat_id={chat_id}): {exc}')

    logger.info(
        f'[ChannelTalk] manager 답변 forward 완료 '
        f'(chat_id={chat_id}, by={manager_name}, text={display_text[:30]!r})'
    )


def _handle_user_message(payload: dict) -> None:
    """사용자(고객) 메시지 처리 — 채널톡 → 슬랙.

    chat_id 별 lock으로 직렬화 — 사진/텍스트가 거의 동시에 들어와도
    첫 webhook이 thread_ts 매핑을 저장한 뒤에 두 번째가 진입.
    """
    entity = payload.get('entity') or {}
    refers = payload.get('refers') or {}
    user_chat = refers.get('userChat') or {}
    user = refers.get('user') or {}

    chat_id = entity.get('chatId') or user_chat.get('id')
    if chat_id:
        # chat_id 기반 lock — 첫 메시지의 thread_ts 저장 완료 후 두 번째 처리
        with _get_chat_lock(chat_id):
            _handle_user_message_locked(
                entity, refers, user_chat, user, chat_id,
            )
        return
    _handle_user_message_locked(entity, refers, user_chat, user, chat_id)


def _handle_user_message_locked(entity, refers, user_chat, user, chat_id) -> None:
    """실제 처리 — lock 안에서 호출됨."""
    plain_text = entity.get('plainText') or ''
    created_ms = entity.get('createdAt') or 0
    customer_name = user.get('name') or user_chat.get('name') or '익명 고객'

    # 파일 추출
    files = _extract_files(entity)
    if not chat_id or (not plain_text and not files):
        return

    channel = _slack_channel()
    if not channel:
        logger.warning('[ChannelTalk] SLACK_CHANNELTALK_CHANNEL 미설정')
        return

    thread_ts = _threads.get_thread_ts(chat_id)

    # 표시 텍스트 (파일만 있을 때는 간단히)
    has_files = bool(files)
    display_text = plain_text or ('🖼️ 사진 전송' if any(f['is_image'] for f in files)
                                  else '📎 파일 전송')

    # 1단계: 텍스트 메시지 — 파일이 없을 때만, 또는 새 채팅(카드)일 때만 발송
    # 파일 있는 thread reply는 텍스트 skip → 파일 업로드의 initial_comment로 합침 (알림 1번)
    if not thread_ts:
        # 새 채팅 — lead 등록 정책:
        # - 스팸 의심 → Redis pending만 저장 (매니저 모달 제출 시 등록)
        # - 일반 메시지 → 즉시 시트 등록 (lead_no 순서 보장, 다른 인입과 시간 순)
        lead_no = ''
        is_spam_suspect = _is_spam_message(display_text)

        if not is_spam_suspect:
            # 일반 채팅 — 즉시 시트 등록 → lead_no 부여
            try:
                lead_no = _register_chat_lead(user_chat, user, display_text, created_ms) or ''
            except Exception as exc:
                logger.warning(f'[ChannelTalk] 즉시 lead 등록 실패: {exc}')

        # 스팸 또는 즉시 등록 실패 시 Redis pending 저장 (매니저 응답 시 활용)
        if not lead_no:
            try:
                from dashboard.utils.redis_client import get_redis_client
                import json as _json
                rc = get_redis_client().redis
                rc.set(
                    f'channeltalk_pending_lead:{chat_id}',
                    _json.dumps({
                        'user_name': user.get('name') or user_chat.get('name') or '익명 고객',
                        'medium': (user_chat.get('mediumProfile') or {}).get('mediumName', ''),
                        'first_message': display_text,
                        'created_ms': created_ms,
                        'is_spam_suspect': is_spam_suspect,
                    }, ensure_ascii=False),
                    ex=60 * 60 * 24 * 30,  # 30일 — 채팅 응답 늦어도 보존
                )
            except Exception as exc:
                logger.warning(f'[ChannelTalk] pending lead 저장 실패: {exc}')

        # 새 채팅 카드 발송 (thread root, lead_no 없음)
        card = _new_chat_card(user_chat, user, display_text, created_ms, lead_no=lead_no)
        # 스팸 의심 시 카드 텍스트 머리에 표시
        card_text = card['text']
        if is_spam_suspect:
            card_text = ':no_entry: *스팸 의심* (미응답 알림 skip)\n' + card_text
        resp = _slack_post('chat.postMessage', {
            'channel': channel,
            'text': card_text,
            'unfurl_links': False,
        })
        if resp and resp.get('ok') and resp.get('ts'):
            thread_ts = resp['ts']
            _threads.set_thread_ts(chat_id, thread_ts)
            # 스팸 의심이면 미응답 알림 큐에 등록 안 함 (5분 알림 skip)
            if not is_spam_suspect:
                _threads.add_pending(chat_id, thread_ts)
            else:
                logger.info(
                    f'[ChannelTalk] 스팸 의심 — 미응답 큐 skip (chat_id={chat_id})'
                )
            # 카드 하단에 [💬 채팅 열기] 링크 — 클릭 시 슬랙 thread로 이동
            _attach_chat_open_button(channel, thread_ts, card_text, lead_no=lead_no, chat_id=chat_id)
        else:
            logger.warning(f'[ChannelTalk] 새 카드 발송 실패 (chat_id={chat_id})')
    elif not has_files:
        # 기존 thread + 텍스트만 → reply
        _slack_post('chat.postMessage', {
            'channel': channel,
            'thread_ts': thread_ts,
            'text': _thread_reply_text(plain_text, customer_name, created_ms),
            'unfurl_links': False,
        })

    # 2단계: 파일 있으면 다운로드 → 슬랙 thread에 한 번에 묶어서 영구 업로드 (알림 1회)
    if has_files and thread_ts:
        from dashboard.services.channeltalk_api import get_file_signed_url, download_file
        if plain_text:
            initial_comment = _thread_reply_text(plain_text, customer_name, created_ms)
        else:
            initial_comment = _thread_reply_text(display_text, customer_name, created_ms)

        # 모든 파일을 일단 다운로드해서 메모리에 모음
        items = []
        for f in files:
            try:
                signed = get_file_signed_url(chat_id, f['key'])
                if not signed:
                    logger.warning(f'[ChannelTalk] signed URL 실패 (key={f["key"]})')
                    continue
                content = download_file(signed)
                if not content:
                    logger.warning(f'[ChannelTalk] 다운로드 실패 (name={f["name"]})')
                    continue
                items.append({'content': content, 'name': f['name']})
            except Exception as exc:
                logger.error(f'[ChannelTalk] 파일 다운로드 예외 ({f.get("name")}): {exc}', exc_info=True)

        # 한 번에 묶어서 업로드 (여러 파일 → 한 메시지 → 알림 1회)
        if items:
            ok = _slack_upload_files(channel, thread_ts, items, initial_comment)
            if ok:
                logger.info(f'[ChannelTalk] 파일 {len(items)}건 슬랙 묶음 업로드 완료')
            else:
                logger.warning(f'[ChannelTalk] 파일 {len(items)}건 슬랙 업로드 실패')


# 채팅 종료 dedup — webhook이 closed 이벤트를 여러 번 보낼 수 있음
_CLOSED_CHATS: dict = {}  # chat_id → epoch sec
_CLOSED_TTL = 3600  # 1시간 (같은 chat_id에 1시간 내 중복 종료 메시지 차단)


def _mark_chat_closed(chat_id: str) -> bool:
    """이미 종료 처리했으면 True. 처음이면 False + 캐시 등록."""
    if not chat_id:
        return True
    import time as _t
    now = _t.time()
    # 오래된 항목 정리
    cutoff = now - _CLOSED_TTL * 2
    for k in list(_CLOSED_CHATS):
        if _CLOSED_CHATS[k] < cutoff:
            del _CLOSED_CHATS[k]
    if chat_id in _CLOSED_CHATS:
        return True
    _CLOSED_CHATS[chat_id] = now
    return False


def _handle_chat_closed(payload: dict) -> None:
    """채팅 종료 이벤트 — 슬랙 thread 끝에 종료 카드 추가 (chat_id별 dedup).

    thread_ts 매핑이 없거나 옛 ts(슬랙이 못 찾는 경우) 메시지 발송 자체 skip.
    text에 고객명 포함 — 혹시 메인 채널로 흘러나가도 어느 상담인지 식별 가능.
    """
    refers = payload.get('refers') or {}
    user_chat = refers.get('userChat') or {}
    user = refers.get('user') or {}
    chat_id = user_chat.get('id')
    if not chat_id:
        return

    # 종료된 채팅 — 미응답 알림 큐에서 즉시 제거 (종료 후 5분 알림 방지)
    _threads.remove_pending(chat_id)

    # dedup — 같은 chat_id에 대해 한 번만
    if _mark_chat_closed(chat_id):
        logger.debug(f'[ChannelTalk] 상담 종료 중복 skip (chat_id={chat_id})')
        return

    thread_ts = _threads.get_thread_ts(chat_id)
    if not thread_ts:
        logger.debug(f'[ChannelTalk] 상담 종료 — thread 매핑 없음, skip (chat_id={chat_id})')
        return

    channel = _slack_channel()
    if not channel:
        return

    # 메시지 존재 확인 — 카드 삭제됐으면 종료 알림 자체 skip (메인 채널 본문 노출 방지)
    try:
        token = _slack_token()
        if token:
            check_req = urllib.request.Request(
                f'https://slack.com/api/conversations.replies'
                f'?channel={channel}&ts={thread_ts}&limit=1',
                headers={'Authorization': f'Bearer {token}'},
            )
            with urllib.request.urlopen(check_req, timeout=5) as r:
                check_resp = json.loads(r.read())
            if not check_resp.get('ok') or not check_resp.get('messages'):
                logger.info(
                    f'[ChannelTalk] 카드 삭제됨 — 상담 종료 알림 skip (chat_id={chat_id})'
                )
                return
    except Exception as exc:
        logger.warning(f'[ChannelTalk] 종료 알림 — 메시지 존재 확인 실패: {exc}')

    # 고객명 — thread 매핑이 어긋나 메인 채널로 흘러나갈 가능성 대비 식별자
    customer_name = (user.get('name') or user_chat.get('name') or '익명 고객').strip()
    text = (
        f':white_check_mark: *상담 종료* (채널톡에서 처리됨) — {customer_name}'
    )

    resp = _slack_post('chat.postMessage', {
        'channel': channel,
        'thread_ts': thread_ts,
        'reply_broadcast': False,  # 명시 — 절대 메인 채널로 보내지 않음
        'text': text,
        'unfurl_links': False,
    })
    if resp and resp.get('ok'):
        logger.info(
            f'[ChannelTalk] 상담 종료 thread reply 완료 '
            f'(chat_id={chat_id}, thread_ts={thread_ts}, customer={customer_name})'
        )
    else:
        # thread_ts가 잘못된 경우 등 — 발송 자체 실패 가능
        logger.warning(
            f'[ChannelTalk] 상담 종료 발송 실패 '
            f'(chat_id={chat_id}, thread_ts={thread_ts}, resp={resp})'
        )


# ─────────────────────────────────────────────────────────────
# 엔드포인트
# ─────────────────────────────────────────────────────────────
@channeltalk_bp.route('/health', methods=['GET'])
def health():
    return jsonify({
        'ok': True,
        'service': 'channeltalk-bridge',
        'operator_id': os.getenv('CHANNELTALK_OPERATOR_ID', ''),
        'slack_channel': _slack_channel(),
    })


@channeltalk_bp.route('/events', methods=['POST', 'GET'])
def events():
    if request.method == 'GET':
        return jsonify({'ok': True, 'message': 'channeltalk webhook endpoint'})

    try:
        payload = request.get_json(silent=True) or {}
    except Exception:
        payload = {}

    if not payload:
        return jsonify({'ok': True}), 200

    entity = payload.get('entity') or {}
    refers = payload.get('refers') or {}
    user_chat = refers.get('userChat') or {}

    # 디버그 — entity 핵심 필드 + 파일 관련 필드 dump
    logger.info(
        f"[ChannelTalk/IN] event={payload.get('event')} "
        f"chatType={entity.get('chatType')} personType={entity.get('personType')} "
        f"text={(entity.get('plainText') or '')[:40]!r} "
        f"files={entity.get('files')} attachments={entity.get('attachments')} "
        f"blocks_types={[b.get('type') for b in (entity.get('blocks') or [])]}"
    )

    # 이벤트 분류
    person_type = entity.get('personType', '')
    chat_state = user_chat.get('state', '')
    chat_type = entity.get('chatType', '')

    try:
        # 1. 사용자(고객) 메시지 → 슬랙으로 전달
        if chat_type == 'userChat' and person_type == 'user':
            _handle_user_message(payload)

        # 2. 매니저 메시지 — 우리 봇 forward(=echo loop) 캐시 비교로 구분
        elif chat_type == 'userChat' and person_type == 'manager':
            chat_id_m = entity.get('chatId') or (refers.get('userChat') or {}).get('id')
            text_m = (entity.get('plainText') or '').strip()
            if is_our_sent(chat_id_m, text_m):
                # 우리 봇이 방금 forward한 메시지 — echo loop 방지로 슬랙 forward는 skip
                # 미응답 큐만 제거 (어차피 사용자가 슬랙에서 답변한 거)
                if chat_id_m:
                    _threads.remove_pending(chat_id_m)
                # 단, lead 등록은 트리거 — 슬랙 thread reply도 매니저 응답이므로 시트 등록 필요
                _register_pending_lead_if_any(chat_id_m)
                logger.debug(f'[ChannelTalk] echo skip — lead 등록 시도 (chatId={chat_id_m})')
            else:
                # 본인이든 다른 매니저든 채널톡 앱에서 직접 답변 → 슬랙 thread로 forward
                _handle_manager_message(payload)

        # 3. 채팅 종료 (state=closed)
        if chat_state == 'closed':
            _handle_chat_closed(payload)

    except Exception as exc:
        logger.error(f'[ChannelTalk] 핸들러 예외: {exc}', exc_info=True)

    return jsonify({'ok': True}), 200
