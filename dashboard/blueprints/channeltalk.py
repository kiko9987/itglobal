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
import urllib.parse
import urllib.request
from datetime import datetime
from typing import Any, Dict, Optional  # noqa: F401

from flask import Blueprint, request, jsonify

from dashboard.services import channeltalk_threads as _threads
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

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
def _format_ts(created_ms: int) -> str:
    """epoch ms → '14:32' / 자정 넘으면 '06.15 14:32'"""
    try:
        dt = datetime.fromtimestamp(created_ms / 1000.0)
        return dt.strftime('%m.%d %H:%M')
    except Exception:
        return ''


def _new_chat_card(user_chat: dict, user: dict, first_message: str,
                   created_ms: int, lead_no: str = '') -> dict:
    """새 채팅 시작 시 표시할 카드 (text + blocks)"""
    customer_name = user.get('name') or user_chat.get('name') or '익명 고객'
    medium = (user_chat.get('mediumProfile') or {}).get('mediumName', '')
    # mediumName 없으면 채널톡 웹 위젯 직접 인입 (default)
    medium_label = MEDIUM_LABELS.get(medium) if medium else '채널톡'
    medium_icon = MEDIUM_ICONS.get(medium) if medium else ':채널톡:'
    if not medium_label:
        medium_label = medium or '채널톡'
    if not medium_icon:
        medium_icon = ':채널톡:'
    ts_str = _format_ts(created_ms)

    lead_line = f'*접수번호:* `{lead_no}`\n' if lead_no else ''
    header_text = (
        lead_line
        + f'{medium_icon} *새 상담 요청*\n'
        + f'>*고객* : {customer_name}\n'
        + f'>*인입 채널* : {medium_label}\n'
        + f'>*시각* : {ts_str}\n'
        + f'>*첫 메시지* : {first_message}\n'
        + f'_(아래 thread에서 실시간 대화)_'
    )

    return {'text': header_text}


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
            '상담 내용': inquiry,
            '키워드': '-',
            '온라인 상담자': '',
            '영업 담당자': '',
            '마지막 연락일': '',
            '피드백': '',
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
def _handle_user_message(payload: dict) -> None:
    """사용자(고객) 메시지 처리 — 채널톡 → 슬랙"""
    entity = payload.get('entity') or {}
    refers = payload.get('refers') or {}
    user_chat = refers.get('userChat') or {}
    user = refers.get('user') or {}

    chat_id = entity.get('chatId') or user_chat.get('id')
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
        # 새 채팅 — 시트에 lead 1건 등록 (통계용)
        lead_no = _register_chat_lead(user_chat, user, display_text, created_ms)

        # 새 채팅 카드 발송 (thread root)
        card = _new_chat_card(user_chat, user, display_text, created_ms, lead_no=lead_no)
        resp = _slack_post('chat.postMessage', {
            'channel': channel,
            'text': card['text'],
            'unfurl_links': False,
        })
        if resp and resp.get('ok') and resp.get('ts'):
            thread_ts = resp['ts']
            _threads.set_thread_ts(chat_id, thread_ts)
            # 새 채팅 — 미배정 알림 대기열에 등록 (5분 후 응답 없으면 알림)
            _threads.add_pending(chat_id, thread_ts)
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


def _handle_chat_closed(payload: dict) -> None:
    """채팅 종료 이벤트 — 슬랙 thread 끝에 종료 카드 추가"""
    refers = payload.get('refers') or {}
    user_chat = refers.get('userChat') or {}
    chat_id = user_chat.get('id')
    if not chat_id:
        return

    thread_ts = _threads.get_thread_ts(chat_id)
    if not thread_ts:
        return

    channel = _slack_channel()
    if not channel:
        return

    _slack_post('chat.postMessage', {
        'channel': channel,
        'thread_ts': thread_ts,
        'text': ':white_check_mark: *상담 종료* (채널톡에서 처리됨)',
        'unfurl_links': False,
    })


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

        # 2. 매니저 발신 메시지는 슬랙에 표시 안 함 (loop 방지)
        elif chat_type == 'userChat' and person_type == 'manager':
            # 향후: 슬랙에 자기 답변 echo 표시 옵션 (현재는 무시)
            logger.debug(f'[ChannelTalk] manager 메시지 skip (chatId={entity.get("chatId")})')

        # 3. 채팅 종료 (state=closed)
        if chat_state == 'closed':
            _handle_chat_closed(payload)

    except Exception as exc:
        logger.error(f'[ChannelTalk] 핸들러 예외: {exc}', exc_info=True)

    return jsonify({'ok': True}), 200
