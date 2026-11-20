"""채널톡 봇 입력 창(양식) → 리드 시트·슬랙 카드 반영 (2026-11-19 L-04606 계기).

채널톡 봇 시나리오에 '설치 견적 문의' 입력 창(이름·연락처·설치 장소·문의 내용)이 붙었는데,
고객이 낸 값은 채팅 메시지가 아니라 '양식 제출'로 들어와 우리 웹훅(고객 메시지만 처리)이 놓쳤다.
L-04606: 신축 3개 층 견적인데 카드엔 '견적 문의' 한 줄, 연락처 '-' — 매니저가 연락처를 다시 물었다.

흐름: 웹훅에서 봇·기타 이벤트가 오면 몇 초 뒤 그 대화의 메시지를 API 로 읽어
제출된 입력 창을 찾는다(메시지 id 별 1회). 리드 시트 빈칸을 채우고, 카드의 이름·연락처·문의 줄을
바꾸고, 카드 스레드에 고객이 남긴 내용을 한 번에 올린다.
"""
import re
import threading
import time
from typing import Dict, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_DONE_TTL = 60 * 60 * 24 * 90
_CHECK_LOCK_SEC = 8      # 봇 메시지가 몰려 와도 대화당 한 번만 확인
_CHECK_DELAY_SEC = 3     # 제출 직후 이벤트가 먼저 오는 경우 — 제출값이 저장될 시간을 둔다

# 채널톡이 이름 없는 방문자에게 붙이는 이름 — '리본 369', '펌킨 939'
_ANON_NAME_RE = re.compile(r'^[가-힣A-Za-z]+ \d{3}$')


def _rc():
    from dashboard.utils.redis_client import get_redis_client
    return get_redis_client().redis


def _esc(s: str) -> str:
    return str(s or '').replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def parse_form(form: dict) -> Dict[str, object]:
    """입력 창 값 → {'name','phone','email','address','inquiry','extra':[(label,value)]} (빈 값은 뺌)."""
    out: Dict[str, object] = {}
    extra: List[Tuple[str, str]] = []
    for i in (form or {}).get('inputs') or []:
        v = str(i.get('value') or '').strip()
        if not v:
            continue
        key = str(i.get('bindingKey') or '')
        label = str(i.get('label') or '').strip()
        if key.endswith('mobileNumber') or '연락처' in label or '전화' in label:
            out['phone'] = v
        elif key.endswith('profile.name') or '이름' in label or '상호' in label:
            out['name'] = v
        elif key.endswith('email') or '이메일' in label:
            out['email'] = v
        elif '주소' in label or '장소' in label or '지역' in label:
            out['address'] = v
        elif '문의' in label or '내용' in label:
            out['inquiry'] = v
        else:
            extra.append((label or '기타', v))
    if extra:
        out['extra'] = extra
    return out


def submitted_forms(chat_id: str) -> List[Tuple[str, dict]]:
    """그 대화에서 고객이 제출한 입력 창 [(메시지 id, form)]."""
    from dashboard.services import channeltalk_api as ct
    r = ct._request('GET', f'/user-chats/{chat_id}/messages?sortOrder=asc&limit=100') or {}
    out = []
    for m in r.get('messages') or []:
        f = m.get('form') or {}
        if f.get('submittedAt') and m.get('id'):
            out.append((m['id'], f))
    return out


def _empty(v) -> bool:
    return str(v or '').strip() in ('', '-', 'nan', 'None', '익명 고객')


def lead_updates(current: dict, vals: dict) -> Dict[str, str]:
    """리드 시트에 쓸 칸 — 비어 있는 칸만 채우고, 문의 내용은 덧붙인다(첫 줄 [유입:…] 표시 보존)."""
    from dashboard.services.lead_helpers import normalize_phone
    from dashboard.services.address_resolver import normalize_display
    up: Dict[str, str] = {}
    phone = normalize_phone(vals.get('phone')) if vals.get('phone') else ''
    if phone and _empty(current.get('고객 연락처')):
        up['고객 연락처'] = phone
    name = str(vals.get('name') or '').strip()
    cur_name = str(current.get('고객명') or '').strip()
    if name and (_empty(cur_name) or _ANON_NAME_RE.match(cur_name)):
        up['고객명'] = name
    addr = str(vals.get('address') or '').strip()
    if addr and _empty(current.get('방문 주소')):
        up['방문 주소'] = normalize_display(addr) or addr
    block = form_text_block(vals)
    if block:
        cur = str(current.get('문의 내용') or '').rstrip()
        if block not in cur:
            up['문의 내용'] = (cur + '\n\n' + block) if not _empty(cur) else block
    return up


def form_text_block(vals: dict) -> str:
    """시트 문의 내용에 덧붙일 묶음 — 긴 문의는 ■소제목·줄바꿈(11-09 규칙)."""
    lines = []
    if vals.get('address'):
        lines.append(f"설치 장소: {vals['address']}")
    if vals.get('inquiry'):
        lines.append(f"문의: {vals['inquiry']}")
    for label, v in vals.get('extra') or []:
        lines.append(f'{label}: {v}')
    return ('■ 고객 입력 창\n' + '\n'.join(lines)) if lines else ''


def thread_note(vals: dict, phone: str = '') -> str:
    """카드 스레드에 올릴 글."""
    rows = []
    if vals.get('name'):
        rows.append(f"• 상호: {_esc(vals['name'])}")
    if phone or vals.get('phone'):
        rows.append(f"• 연락처: {_esc(phone or vals['phone'])}")
    if vals.get('address'):
        rows.append(f"• 설치 장소: {_esc(vals['address'])}")
    if vals.get('inquiry'):
        rows.append(f"• 문의: {_esc(vals['inquiry'])}")
    if vals.get('email'):
        rows.append(f"• 이메일: {_esc(vals['email'])}")
    for label, v in vals.get('extra') or []:
        rows.append(f'• {_esc(label)}: {_esc(v)}')
    return ':memo: *고객이 입력 창에 남긴 내용*\n' + '\n'.join(rows) if rows else ''


def patch_card_text(text: str, vals: dict, phone: str = '') -> str:
    """카드 본문의 이름·연락처·문의 줄만 바꾼다 — 유입 경로 줄·버튼 링크는 그대로."""
    if not text:
        return text
    name = str(vals.get('name') or '').strip()
    if name:
        def _name(m):
            cur = m.group(2).strip()
            return m.group(1) + (_esc(name) if (_empty(cur) or _ANON_NAME_RE.match(cur)) else cur)
        text = re.sub(r'(이름 / 상호 : )([^\n]*)', _name, text, count=1)
    if phone:
        text = re.sub(r'(연락처 : )-\s*(?=\n|$)', lambda m: m.group(1) + phone, text, count=1)
    add = []
    if vals.get('address'):
        add.append(f"설치 장소 : {_esc(vals['address'])}")
    if vals.get('inquiry'):
        add.append(f"문의 내용 : {_esc(vals['inquiry'])}")
    if add and '설치 장소 : ' not in text:
        # 버튼으로 고른 '견적 문의' 한 줄을 고객이 적은 내용으로 바꾼다
        text = re.sub(r'문의 내용 : [^\n]*', lambda m: '\n'.join(add), text, count=1)
    return text


def _patch_card(channel: str, ts: str, vals: dict, phone: str) -> bool:
    import os
    from slack_sdk import WebClient
    cli = WebClient(token=os.getenv('SLACK_BOT_TOKEN', '').strip())
    msgs = cli.conversations_history(channel=channel, latest=ts, oldest=ts, inclusive=True,
                                     limit=1).get('messages') or []
    if not msgs:
        return False
    m = msgs[0]
    # 화면에 보이는 건 블록 글 — 메시지의 text 는 슬랙이 줄바꿈을 공백으로 뭉갠 대체 글자라
    #   줄 단위로 바꿀 수 없다. 블록 글을 바꾸고 대체 글자도 그걸로 맞춘다.
    blocks = m.get('blocks') or []
    new_text, changed = None, False
    for b in blocks:
        t = (b.get('text') or {})
        if b.get('type') == 'section' and t.get('type') == 'mrkdwn':
            nt = patch_card_text(t.get('text') or '', vals, phone)
            if nt != t.get('text'):
                t['text'] = nt
                changed = True
            if new_text is None:
                new_text = t['text']
    if not changed:
        return False
    cli.chat_update(channel=channel, ts=ts, text=new_text, blocks=blocks,
                    unfurl_links=False, unfurl_media=False)
    return True


def apply_submitted_forms(chat_id: str) -> int:
    """그 대화에서 새로 제출된 입력 창을 리드·카드·스레드에 반영. 반영한 창 수."""
    if not chat_id:
        return 0
    import os
    from dashboard.services import channeltalk_threads as threads
    rc = _rc()
    done = 0
    for msg_id, form in submitted_forms(chat_id):
        if not rc.set(f'channeltalk_form_done:{chat_id}:{msg_id}', '1', nx=True, ex=_DONE_TTL):
            continue
        vals = parse_form(form)
        if not vals:
            continue
        done += 1
        from dashboard.services.lead_helpers import normalize_phone
        phone = normalize_phone(vals.get('phone')) if vals.get('phone') else ''
        lead_no = rc.get(f'channeltalk_chat_lead:{chat_id}')
        lead_no = lead_no.decode() if isinstance(lead_no, bytes) else (lead_no or '')
        if lead_no:
            try:
                from dashboard.services.lead_service import get_lead_by_no, update_lead
                up = lead_updates(get_lead_by_no(lead_no) or {}, vals)
                if up:
                    res = update_lead(lead_no, up)
                    logger.info(f'[ChannelTalk/양식] {lead_no} 시트 반영 {list(up)} → {res.get("success")}')
            except Exception as exc:
                logger.error(f'[ChannelTalk/양식] {lead_no} 시트 반영 실패: {exc}', exc_info=True)
        channel = os.getenv('SLACK_CHANNELTALK_CHANNEL', '').strip()
        ts = threads.get_thread_ts(chat_id)
        if channel and ts:
            try:
                _patch_card(channel, ts, vals, phone)
            except Exception as exc:
                logger.warning(f'[ChannelTalk/양식] 카드 줄 바꾸기 실패 ({chat_id}): {exc}')
            note = thread_note(vals, phone)
            if note:
                try:
                    from slack_sdk import WebClient
                    WebClient(token=os.getenv('SLACK_BOT_TOKEN', '').strip()).chat_postMessage(
                        channel=channel, thread_ts=ts, text=note,
                        unfurl_links=False, unfurl_media=False)
                except Exception as exc:
                    logger.warning(f'[ChannelTalk/양식] 스레드 글 실패 ({chat_id}): {exc}')
        logger.info(f'[ChannelTalk/양식] 입력 창 반영 (chat_id={chat_id}, lead={lead_no or "-"}, '
                    f'칸={sorted(k for k in vals if k != "extra")})')
    return done


def schedule_check(chat_id: str) -> None:
    """웹훅에서 부른다 — 몇 초 뒤 그 대화의 제출된 입력 창 확인(대화당 몇 초에 한 번만)."""
    import os
    if not chat_id or not os.getenv('CHANNELTALK_ACCESS_KEY', '').strip():
        return
    try:
        if not _rc().set(f'channeltalk_form_check:{chat_id}', '1', nx=True, ex=_CHECK_LOCK_SEC):
            return
    except Exception:
        return

    def _bg():
        time.sleep(_CHECK_DELAY_SEC)
        try:
            apply_submitted_forms(chat_id)
        except Exception as exc:
            logger.warning(f'[ChannelTalk/양식] 확인 실패 ({chat_id}): {exc}')
    threading.Thread(target=_bg, daemon=True).start()
