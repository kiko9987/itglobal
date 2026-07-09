"""
공사 확정 슬랙 알림 모듈 (별도 봇 "공사 현황 알림 봇" 사용).

- POST /api/projects/auto 에서 호출
- 환경변수: SLACK_PROJECT_BOT_TOKEN, SLACK_PROJECT_CHANNEL
- 별도 토큰이므로 기존 SLACK_BOT_TOKEN과 무관 (인입 알림과 분리)
"""

import json
import os
import urllib.request
from typing import Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)


def _money_kr(value) -> str:
    """₩1,000 / 1000 / 1000.0 / '1,000' 모두 받아 '1,000원' 형태로 정규화. 빈값은 '-'."""
    if value is None or value == '' or value == '-':
        return '-'
    # 숫자 타입은 문자 추출을 거치면 float의 '.0' 뒷자리가 붙어 10배로 잘못 계산됨.
    # (2026-07-08 공사 확정 카드 금액이 10배로 표시된 사고 fix)
    if isinstance(value, bool):  # bool도 int 서브클래스라 우선 걸러냄
        return '-'
    if isinstance(value, (int, float)):
        try:
            return f"{int(value):,}원"
        except (ValueError, OverflowError):
            return '-'
    s = str(value).strip()
    # ₩ 와 콤마 제거 후 숫자만 추출
    digits = ''.join(ch for ch in s if ch.isdigit() or ch == '-')
    if not digits or digits == '-':
        return s if s else '-'
    try:
        num = int(digits)
    except ValueError:
        return s
    return f"{num:,}원"


def _val(data: dict, key: str) -> str:
    """data[key]를 표시용 문자열로. 빈값은 '-'."""
    v = data.get(key)
    if v is None:
        return '-'
    s = str(v).strip()
    return s if s else '-'


def _date_val(data: dict, key: str) -> str:
    """날짜 필드 표시용. pandas Timestamp / datetime / 'YYYY-MM-DD HH:MM:SS' 모두 'YYYY-MM-DD'로 절삭."""
    v = data.get(key)
    if v is None:
        return '-'
    # pandas NaT 처리
    try:
        import pandas as _pd
        if _pd.isna(v):
            return '-'
    except Exception:
        pass
    # datetime/Timestamp면 strftime
    if hasattr(v, 'strftime'):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    if not s or s == 'NaT':
        return '-'
    # 문자열에 시간 붙어있으면 앞 10자만
    if len(s) >= 10 and s[4] == '-' and s[7] == '-':
        return s[:10]
    return s


def _format_inflow(value: str) -> str:
    """유입 구분 표기 — 공용 헬퍼 위임 (온라인/거래처/기타 3카테고리 통일)."""
    from dashboard.services.lead_helpers import format_inflow_display
    return format_inflow_display(value)


def _build_message(
    data: dict, code: str, license_attached: bool = False,
    thread_permalink: Optional[str] = None,
) -> str:
    """공사 확정 알림 메시지 본문 생성.

    license_attached: 사업자등록증 첨부 여부(스레드에 첨부돼 Drive 폴더에 저장됐는지).
    thread_permalink: 미첨부일 때 라인 옆에 '(첨부하기)' 링크 삽입용.
    """
    code_safe = code or '-'

    # 부가세 처리 — TRUE/'TRUE'/True 모두 별도 표기
    vat_raw = data.get('부가세')
    vat_is_separate = (
        vat_raw is True
        or (isinstance(vat_raw, str) and vat_raw.strip().upper() in ('TRUE', 'Y', 'YES', '1'))
        or vat_raw == 1
    )
    amount_str = _money_kr(data.get('총액 1'))
    if amount_str == '-':
        amount_line = amount_str
    elif vat_is_separate:
        amount_line = f"{amount_str} (VAT 별도)"
    else:
        # 2026-07-08 부가세 미체크도 명시 표시 (매니저 판단 명확화)
        amount_line = f"{amount_str} (VAT 없음)"

    # 폴더 경로 (AL열) — 값이 있으면 표시, 형태에 따라 링크 처리
    folder_raw = str(data.get('견적서 및 계약서 폴더 경로', '') or '').strip()
    if folder_raw and folder_raw != '-':
        import re as _re
        # 폴더 ID 패턴(20자+ 영숫자·언더스코어·하이픈)이면 itgfolder:// 클릭 링크
        if _re.match(r'^[a-zA-Z0-9_-]{20,}$', folder_raw):
            folder_display = f'<itgfolder://{folder_raw}|탐색기에서 열기>'
        else:
            # 로컬 경로 등은 raw text (매니저가 복사해서 붙여넣기)
            folder_display = folder_raw
        folder_line = f"📁 폴더 경로 : {folder_display}"
    else:
        folder_line = None

    # 이모지는 raw Unicode 사용 — 취소 시 카드가 ``` 코드블록으로 감싸져도 정상 렌더됨
    # (:bell: 같은 shortcode는 코드블록 안에서 텍스트로 남음).
    lines = [
        f"🔔 *[공사 확정 알림]*  `{code_safe}`",
        "--------------------------------------------",
        f"📥 유입 구분 : {_format_inflow(_val(data, '유입 구분'))}",
        f"🏢 사업자명 : {_val(data, '사업자명')}",
        f"📍 현장 주소 : {_val(data, '현장 주소')}",
        f"👤 발주처 담당자 : {_val(data, '발주처 담당자')}",
        f"📞 발주처 연락처 : {_val(data, '발주처 연락처')}",
        f"✉️ 발주처 이메일 : {_val(data, '발주처 이메일')}",
        f"📋 공사 내용 : {_val(data, '공사 내용')}",
        f"🛠️ 도급 구분 : {_val(data, '도급 구분')}",
        f"👷 시공자 : {_val(data, '시공자')}",
        f"💲 공사 금액 : {amount_line}",
        f"📅 공사 시작 : {_date_val(data, '공사 시작')}",
        f"📅 공사 종료 : {_date_val(data, '공사 종료')}",
    ]
    if folder_line:
        lines.append(folder_line)
    lines.append("--------------------------------------------")
    # 2026-07-09 사업자등록증 배지는 카드 구분선 바깥 하단에 별도 라인으로 표시.
    # 스레드에 파일 첨부 → Drive 저장 성공 시 chat.update로 True.
    # 미첨부 상태에서만 옆에 (첨부하기) 스레드 링크 노출.
    if license_attached:
        license_line = "📎 사업자등록증 : ✅ 첨부됨"
    else:
        license_line = "📎 사업자등록증 : ⬜ 미첨부"
        if thread_permalink:
            license_line += f"  <{thread_permalink}|(첨부하기)>"
    lines.append(license_line)
    # 2026-07-09 리드 알림 스타일로 통일 — 상단 여백은 리드와 동일한 U+2800 라인.
    return "⠀\n" + '\n'.join(lines)


def _build_invoice_button_value(data: dict, code: str) -> str:
    """[계산서 요청] 버튼 value — 모달 pre-fill용 프로젝트 정보 JSON."""
    # 총액 1은 숫자만 — float('9600000.0') → 소수점 이하 절삭 후 digits 추출
    # (str(9600000.0)="9600000.0" 에서 '.'만 빼면 "96000000" 이 되는 10x 버그 방지)
    amount_raw = data.get('총액 1', '')
    try:
        amount_digits = str(int(float(str(amount_raw).replace(',', '').strip() or '0')))
        if amount_digits == '0':
            amount_digits = ''
    except (ValueError, TypeError):
        amount_digits = ''
    vat_raw = data.get('부가세')
    vat_sep = (
        vat_raw is True
        or (isinstance(vat_raw, str) and vat_raw.strip().upper() in ('TRUE', 'Y', 'YES', '1'))
        or vat_raw == 1
    )
    payload = {
        'code': code or '',
        'biz': _val(data, '사업자명'),
        'addr': _val(data, '현장 주소'),
        'amt': amount_digits or '',
        'vat': 'sep' if vat_sep else 'incl',
        'email': _val(data, '발주처 이메일'),
    }
    # dash 처리 — '-'는 빈값으로
    for k in ('biz', 'addr', 'email'):
        if payload[k] == '-':
            payload[k] = ''
    return json.dumps(payload, ensure_ascii=False)


def _thread_permalink(channel: str, ts: str) -> str:
    """이 메시지 스레드로 곧장 열리는 슬랙 웹 URL. `?thread_ts&cid`가 있어야 스레드 pane이 열림.
    workspace 도메인은 SLACK_WORKSPACE_DOMAIN 환경변수 (없으면 fallback)."""
    workspace = os.getenv('SLACK_WORKSPACE_DOMAIN', 'w1679118856-ybv966043').strip()
    ts_no_dot = (ts or '').replace('.', '')
    return f'https://{workspace}.slack.com/archives/{channel}/p{ts_no_dot}?thread_ts={ts}&cid={channel}'


def _build_blocks(
    data: dict, code: str, license_attached: bool = False,
    thread_permalink: Optional[str] = None,
) -> list:
    """공사 확정 알림 blocks — section(본문) + actions(3버튼) + 하단 여백 context.

    버튼: [계산서 요청] [내용 수정] [공사 취소]. 매니저가 밖에서 슬랙만으로 수정/취소 가능.
    thread_permalink 지정 시 사업자등록증 미첨부일 때만 스레드 진입 링크 삽입.
    """
    text = _build_message(
        data, code,
        license_attached=license_attached,
        thread_permalink=thread_permalink,
    )
    btn_value = _build_invoice_button_value(data, code)
    blocks: list = [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text': text}},
    ]
    # 2026-07-09 사업자등록증 라인과 액션 버튼 사이 여백 한 줄.
    blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': '⠀'}]})
    blocks.append({
            'type': 'actions',
            'elements': [
                {
                    'type': 'button',
                    'text': {'type': 'plain_text', 'text': '🧾 세금계산서 요청', 'emoji': True},
                    'action_id': 'invoice_request_open',
                    'value': btn_value,
                },
                {
                    'type': 'button',
                    'text': {'type': 'plain_text', 'text': '✏️ 내용 수정', 'emoji': True},
                    'action_id': 'project_edit_open',
                    'value': code,
                },
                {
                    'type': 'button',
                    'text': {'type': 'plain_text', 'text': '❌ 공사 취소', 'emoji': True},
                    'action_id': 'project_cancel_confirm',
                    'value': code,
                    'confirm': {
                        'title': {'type': 'plain_text', 'text': '공사 취소 확인'},
                        'text': {
                            'type': 'plain_text',
                            'text': f'{code} 공사를 취소하시겠습니까?\n'
                                    f'관리 사이트의 [공사 취소] 버튼과 동일하게 처리됩니다.',
                        },
                        'confirm': {'type': 'plain_text', 'text': '공사 취소'},
                        'deny': {'type': 'plain_text', 'text': '아니오'},
                    },
                },
            ],
        })
    # 2026-07-09 리드 알림 스타일과 통일 — 카드 아래 여백.
    blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': '⠀'}]})
    return blocks


def send_project_created_notification(data: dict, code: str) -> bool:
    """공사 확정 알림을 #공사_확정 채널에 발송 (section + [계산서 요청] 버튼).

    Args:
        data: 프로젝트 dict (시트 한국어 헤더 키 — 거래처/사업자명/현장 주소 등)
        code: 프로젝트 코드 (예: 'G3764-MJ')

    Returns:
        True 성공, False 실패 (호출자는 무시해도 무방 — 시트 등록은 이미 완료).
    """
    token = os.getenv('SLACK_PROJECT_BOT_TOKEN', '').strip()
    channel = os.getenv('SLACK_PROJECT_CHANNEL', '').strip()
    if not token:
        logger.debug('[PROJECT/SLACK] SLACK_PROJECT_BOT_TOKEN 미설정 — 알림 스킵')
        return False
    if not channel:
        logger.debug('[PROJECT/SLACK] SLACK_PROJECT_CHANNEL 미설정 — 알림 스킵')
        return False

    # fallback text (알림 미리보기용) — 리드 방식으로 짧은 요약만. blocks 안 텍스트는 별도.
    biz = _val(data, '사업자명')
    text = f"[공사 확정] {code} {biz}".strip()
    blocks = _build_blocks(data, code)
    payload = {
        'channel': channel,
        'text': text,
        'blocks': blocks,
        'unfurl_links': False,
    }

    try:
        req = urllib.request.Request(
            'https://slack.com/api/chat.postMessage',
            data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
            headers={
                'Content-Type': 'application/json; charset=utf-8',
                'Authorization': f'Bearer {token}',
            },
        )
        with urllib.request.urlopen(req, timeout=5) as r:
            resp = json.loads(r.read())
        if resp.get('ok'):
            ts = resp.get('ts', '')
            logger.info(f"[PROJECT/SLACK] 공사 확정 알림 발송 완료: {code} (ts={ts})")
            # 편집 시 스레드 답글용 채널|ts 매핑 저장 (180일)
            if ts:
                try:
                    from dashboard.utils.redis_client import get_redis_client
                    rc = get_redis_client().redis
                    rc.set(
                        f'project_card_msg:{code}',
                        f'{channel}|{ts}',
                        ex=60 * 60 * 24 * 180,
                    )
                    # 2026-07-08 역방향 매핑 (사업자등록증 첨부 등 스레드 이벤트 처리용).
                    # 스레드 답글이 들어오면 channel|ts로 프로젝트 코드를 즉시 조회.
                    rc.set(
                        f'project_thread:{channel}|{ts}',
                        code,
                        ex=60 * 60 * 24 * 180,
                    )
                except Exception as _exc:
                    logger.debug(f'[PROJECT/SLACK] card 매핑 저장 실패 ({code}): {_exc}')
                # 사업자등록증 첨부 유도 링크는 ts를 알아야 만들 수 있어 chat.update으로 삽입.
                try:
                    permalink = _thread_permalink(channel, ts)
                    upd_blocks = _build_blocks(data, code, thread_permalink=permalink)
                    upd_payload = {
                        'channel': channel, 'ts': ts,
                        'text': text, 'blocks': upd_blocks,
                    }
                    upd_req = urllib.request.Request(
                        'https://slack.com/api/chat.update',
                        data=json.dumps(upd_payload, ensure_ascii=False).encode('utf-8'),
                        headers={
                            'Content-Type': 'application/json; charset=utf-8',
                            'Authorization': f'Bearer {token}',
                        },
                    )
                    with urllib.request.urlopen(upd_req, timeout=5) as _r2:
                        _ = json.loads(_r2.read())
                except Exception as _exc:
                    logger.debug(f'[PROJECT/SLACK] 첨부 링크 삽입 실패 ({code}): {_exc}')
            return True
        logger.warning(f"[PROJECT/SLACK] 슬랙 API 실패 ({code}): {resp.get('error')}")
        return False
    except Exception as exc:
        logger.warning(f"[PROJECT/SLACK] 발송 예외 ({code}): {exc}")
        return False


# ─────────────────────────────────────────────────────────────
# 편집 시 스레드 답글 (공사 확정 카드 밑 댓글로 변경 내역 전송)
# ─────────────────────────────────────────────────────────────
# 알림 대상 필드 (audit-worthy) — 잡음 필드는 제외
_NOTIFY_FIELDS = {
    '사업자', '유입 구분', '사업자명', '현장 주소',
    '공사 구분', '기계 분류', '브랜드',
    '공사 시작', '공사 종료', '공사 내용', '도급 구분', '시공자',
    '발주처 담당자', '발주처 연락처', '발주처 이메일',
    '담당자', '프로젝트 코드',
    '총액 1', '부가세',
    '계약금', '중도금', '잔금', '수금 날짜', '계산서',
    '수금 관련 특이사항',
    '공사 확정', 'Lead No',
    # 2026-07-08: 매니저는 공사 확정 후 폴더를 별도로 만들고 ID만 나중에 채우는 경우가
    # 흔함. 이 필드가 여기 없으면 폴더만 편집 시 카드가 stale 상태로 남아 매니저가
    # '탐색기에서 열기' 링크를 못 봄. 답글 히스토리 + 원본 카드 재렌더링 모두 트리거.
    '견적서 및 계약서 폴더 경로',
}


def _fmt_money(v) -> str:
    try:
        n = int(float(str(v).replace(',', '').strip() or '0'))
        return f'{n:,}'
    except (ValueError, TypeError):
        return str(v) if v is not None else '-'


def _fmt_vat(v) -> str:
    """부가세 값 → 사람 읽는 라벨."""
    if v is True or (isinstance(v, str) and v.strip().upper() in ('TRUE', 'Y', 'YES', '1')) or v == 1:
        return 'VAT 별도'
    return 'VAT 체크 안됨'


def _fmt_date(v) -> str:
    """날짜값 → 'YYYY-MM-DD'."""
    if v is None or v == '' or str(v).strip() in ('-', 'NaT'):
        return '-'
    if hasattr(v, 'strftime'):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    if len(s) >= 10 and s[4] == '-' and s[7] == '-':
        return s[:10]
    return s


def _fmt_field(field: str, value) -> str:
    """필드별 표시 포맷."""
    if field in ('총액 1', '계약금', '중도금', '잔금'):
        return _fmt_money(value)
    if field == '부가세':
        return _fmt_vat(value)
    if field in ('공사 시작', '공사 종료', '수금 날짜', '공사 확정'):
        return _fmt_date(value)
    if value is None or str(value).strip() == '':
        return '-'
    return str(value).strip()


def notify_project_field_changes(code: str, field_changes: list, latest_data: dict = None) -> bool:
    """편집된 필드들을 공사 확정 카드 스레드에 답글로 전송 + 원본 카드 최신 데이터로 재렌더링.

    Args:
        code: 프로젝트 코드 (변경 후 코드)
        field_changes: [{'field_name', 'old_value', 'new_value'}, ...]
        latest_data: 편집 후 시트 dict (있으면 원본 카드 blocks 갱신)

    스레드 답글 = 편집 히스토리 로그, 원본 카드 = 항상 최신 스냅샷.
    _NOTIFY_FIELDS에 포함된 필드만 답글에 표시. Redis에 카드 매핑 없으면 skip.
    """
    if not code or not field_changes:
        return False

    # field_changes 항목 구조: {'field_name', 'old_value', 'new_value'}
    # 대상 필드 필터
    relevant = [c for c in field_changes if c.get('field_name') in _NOTIFY_FIELDS]

    # 실제로 값이 다른 것만 (동일값 저장은 노이즈)
    def _neq(c):
        o = _fmt_field(c['field_name'], c.get('old_value'))
        n = _fmt_field(c['field_name'], c.get('new_value'))
        return o != n
    relevant = [c for c in relevant if _neq(c)]
    if not relevant:
        return False

    # Redis에서 채널|ts 조회 (코드가 변경됐으면 old_code로도 시도)
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    except Exception:
        return False

    mapping = rc.get(f'project_card_msg:{code}')
    if not mapping:
        # 코드 재발번된 경우 old_code로 조회
        code_change = next((c for c in field_changes if c.get('field_name') == '프로젝트 코드'), None)
        if code_change and code_change.get('old_value'):
            old_code = str(code_change['old_value']).strip()
            mapping = rc.get(f'project_card_msg:{old_code}')
            if mapping:
                # 새 코드로 재매핑
                try:
                    rc.set(f'project_card_msg:{code}', mapping, ex=60 * 60 * 24 * 180)
                    rc.delete(f'project_card_msg:{old_code}')
                except Exception:
                    pass
    if not mapping:
        logger.debug(f'[PROJECT/SLACK/편집] card 매핑 없음, skip ({code})')
        return False

    try:
        channel, ts = mapping.split('|', 1) if isinstance(mapping, str) else mapping.decode().split('|', 1)
    except Exception:
        return False

    # 본문 조립
    lines = [f'[{code} 데이터 수정 알림]']
    for c in relevant:
        f = c['field_name']
        old_disp = _fmt_field(f, c.get('old_value'))
        new_disp = _fmt_field(f, c.get('new_value'))
        lines.append(f'- {f}: {old_disp} → {new_disp}')
    text = '\n'.join(lines)

    token = os.getenv('SLACK_PROJECT_BOT_TOKEN', '').strip()
    if not token:
        return False

    try:
        payload = {
            'channel': channel,
            'thread_ts': ts,
            'text': text,
            'unfurl_links': False,
        }
        req = urllib.request.Request(
            'https://slack.com/api/chat.postMessage',
            data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
            headers={
                'Content-Type': 'application/json; charset=utf-8',
                'Authorization': f'Bearer {token}',
            },
        )
        with urllib.request.urlopen(req, timeout=5) as r:
            resp = json.loads(r.read())
        if resp.get('ok'):
            logger.info(f'[PROJECT/SLACK/편집] 변경 알림 답글 완료: {code} ({len(relevant)}개 필드)')
            reply_ok = True
        else:
            logger.warning(f'[PROJECT/SLACK/편집] 답글 실패 ({code}): {resp.get("error")}')
            reply_ok = False
    except Exception as exc:
        logger.warning(f'[PROJECT/SLACK/편집] 답글 예외 ({code}): {exc}')
        reply_ok = False

    # 원본 카드도 최신 데이터로 재렌더링 (스레드 답글과 독립적으로 시도)
    if latest_data is not None:
        try:
            # 사업자등록증 상태 유지 — 편집으로 카드가 재렌더링되면서 배지가 뒤집히면 안 됨.
            try:
                from dashboard.services.business_license_handler import verify_license_exists
                license_attached = verify_license_exists(code)
            except Exception:
                license_attached = False
            permalink = _thread_permalink(channel, ts)
            new_blocks = _build_blocks(
                latest_data, code,
                license_attached=license_attached,
                thread_permalink=permalink,
            )
            biz = _val(latest_data, '사업자명')
            new_text = f"[공사 확정] {code} {biz}".strip()
            payload = {
                'channel': channel,
                'ts': ts,
                'text': new_text,
                'blocks': new_blocks,
            }
            req = urllib.request.Request(
                'https://slack.com/api/chat.update',
                data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
                headers={
                    'Content-Type': 'application/json; charset=utf-8',
                    'Authorization': f'Bearer {token}',
                },
            )
            with urllib.request.urlopen(req, timeout=5) as r:
                up_resp = json.loads(r.read())
            if up_resp.get('ok'):
                logger.info(f'[PROJECT/SLACK/편집] 원본 카드 갱신 완료: {code}')
            else:
                logger.warning(f'[PROJECT/SLACK/편집] 원본 카드 갱신 실패 ({code}): {up_resp.get("error")}')
        except Exception as exc:
            logger.warning(f'[PROJECT/SLACK/편집] 원본 카드 갱신 예외 ({code}): {exc}')

    return reply_ok


def refresh_project_card_license(code: str, latest_data: Optional[dict] = None) -> bool:
    """사업자등록증 첨부 상태 변경 시 원본 공사 확정 카드를 chat.update로 갱신.

    스레드에 사업자등록증 파일이 첨부돼 Drive 저장이 성공한 직후 호출.
    latest_data가 없으면 시트에서 최신 프로젝트 데이터를 재조회한다(있으면 그대로 사용).
    """
    if not code:
        return False

    token = os.getenv('SLACK_PROJECT_BOT_TOKEN', '').strip()
    if not token:
        return False

    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    except Exception:
        return False

    mapping = rc.get(f'project_card_msg:{code}')
    if not mapping:
        logger.debug(f'[PROJECT/SLACK/사업자] card 매핑 없음, skip ({code})')
        return False
    try:
        channel, ts = (
            mapping.split('|', 1) if isinstance(mapping, str) else mapping.decode().split('|', 1)
        )
    except Exception:
        return False

    if latest_data is None:
        try:
            from dashboard.services.project_service import get_project_records
            records = get_project_records() or []
            latest_data = next(
                (r for r in records if (r.get('프로젝트 코드') or '').strip() == code),
                None,
            )
        except Exception:
            latest_data = None
    if not latest_data:
        logger.debug(f'[PROJECT/SLACK/사업자] 프로젝트 데이터 조회 실패, skip ({code})')
        return False

    try:
        from dashboard.services.business_license_handler import verify_license_exists
        license_attached = verify_license_exists(code)
    except Exception:
        license_attached = False

    try:
        permalink = _thread_permalink(channel, ts)
        new_blocks = _build_blocks(
            latest_data, code,
            license_attached=license_attached,
            thread_permalink=permalink,
        )
        biz = _val(latest_data, '사업자명')
        new_text = f"[공사 확정] {code} {biz}".strip()
        payload = {
            'channel': channel,
            'ts': ts,
            'text': new_text,
            'blocks': new_blocks,
        }
        req = urllib.request.Request(
            'https://slack.com/api/chat.update',
            data=json.dumps(payload, ensure_ascii=False).encode('utf-8'),
            headers={
                'Content-Type': 'application/json; charset=utf-8',
                'Authorization': f'Bearer {token}',
            },
        )
        with urllib.request.urlopen(req, timeout=5) as r:
            resp = json.loads(r.read())
        if resp.get('ok'):
            logger.info(
                f'[PROJECT/SLACK/사업자] 원본 카드 갱신 완료: {code} '
                f'(license_attached={license_attached})'
            )
            return True
        logger.warning(
            f'[PROJECT/SLACK/사업자] 원본 카드 갱신 실패 ({code}): {resp.get("error")}'
        )
        return False
    except Exception as exc:
        logger.warning(f'[PROJECT/SLACK/사업자] 원본 카드 갱신 예외 ({code}): {exc}')
        return False
