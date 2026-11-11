"""매입 계산서 현장 지정 (전자결재 1단계, 2026-10-30).

경영지원이 #영업_관리에 올리는 매입 계산서 확인 요청
('10/29 G 533,500원 강원포장 회신 바랍니다.' + 계산서 이미지)은 "이 계산서가 어느 현장 건인지"
묻는 글이다. 지금까지는 매니저가 스레드에 코드를 손으로 적어 답했다.

계산서봇이 이 글을 자동 고정할 때 스레드에 [📍 현장 지정] 버튼을 달고, 매니저가 모달에서
프로젝트·원가 분류를 고르면 결과를 장부(users.db purchase_invoice_assignments)에 남기고
스레드에 답한 뒤 고정을 푼다. 장부는 이후 원가(기타비 등)·이카운트 매입 공사코드 지정에 쓴다.

여기는 슬랙 I/O 없는 순수 로직 + 장부만 둔다. 슬랙 핸들러는 blueprints/slack_bot.py.
설계: docs/e_approval_design.md (4-1 입구 A).
"""
from __future__ import annotations

import json
import logging
import re
from datetime import datetime
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

# ── 슬랙 식별자 ───────────────────────────────────────────────────────
ACTION_OPEN = 'pinv_assign_open'        # [📍 현장 지정] / [✏️ 다시 지정]
ACTION_PROJECT = 'pinv_project'         # 모달 프로젝트 external_select
CALLBACK_SUBMIT = 'submit_pinv_assign'  # 모달 제출

# ── 현장 아닌 선택지 ─────────────────────────────────────────────────
STOCK_CODE = '프로젝트0'    # 회사 재고로 사거나 회사로 먼저 사서 나중에 현장에 내보내는 것
COMMON_CODE = '공통비'      # 사무실 등 현장과 무관한 비용
SPECIAL_LABELS = {
    STOCK_CODE: '프로젝트 0 · 회사 재고 (현장 미정 자재)',
    COMMON_CODE: '공통비 · 현장 무관 (사무실 집기, 수선비 등)',
}
# 법인서류 요청 전용 — 서류는 돈이 아니라 '공통비'(비용 분류)가 맞지 않음. 특정 현장이 아니면 이것(11-05)
NO_SITE_CODE = '현장 없음'
NO_SITE_LABEL = '현장 없음 · 입찰·관공서 등 회사 제출용'
_FORM_ONLY_LABELS = {NO_SITE_CODE: NO_SITE_LABEL}

# ── 원가 분류 (공사현황 손익 4칸 + 기타비 세부) ─────────────────────
COST_CATEGORIES = [
    ('기타비:운반비', '기타비 · 운반비'),
    ('기타비:사다리차', '기타비 · 사다리차'),
    ('기타비:크레인', '기타비 · 크레인'),
    ('기타비:폐기물', '기타비 · 폐기물'),
    ('기타비:주차비', '기타비 · 주차비'),   # 11-05 — 개인경비·현장 주차
    ('기타비:과태료', '기타비 · 과태료'),   # 11-05 경영지원 의견 — 주정차 위반 등
    ('기타비:기타', '기타비 · 기타'),
    ('자재비', '자재비'),
    ('제품대', '제품대'),
    ('도급비', '도급비'),
    ('반환:환불', '반환·환불 (원가 아님)'),   # 11-01 과입금 계약금 반환 등 — 플렉스 1년 22건·5,673만
    # 공통비(현장 무관) 전용 — 사무실 집기 등은 공사 원가 분류가 맞지 않았다(11-06 KiKO, 신규 입사자 책상·의자·모니터)
    ('공통비:집기비품', '공통비 · 사무 집기·비품'),
    ('공통비:수선', '공통비 · 사무실 수선'),
    ('공통비:운영', '공통비 · 기타 운영비'),
]
_COST_LABEL = dict(COST_CATEGORIES)


def cost_project_error(cost: str, project_code: str) -> str:
    """공통비 분류는 현장이 「공통비」일 때만 — 공사 현장 원가에 사무실 비용이 섞이지 않게."""
    if (cost or '').startswith('공통비:') and project_code != COMMON_CODE:
        return '공통비 분류는 현장을 「공통비」로 골랐을 때만 써요.'
    return ''


def cost_label(value: str) -> str:
    return _COST_LABEL.get(value or '', value or '-')


# ── 글 해석 ───────────────────────────────────────────────────────────
# '10/29 G 533,500원 강원포장 회신 바랍니다.' — pin_remind._INVOICE_RE 와 같은 구조 판정
# (금액은 콤마 그룹 or 4자리 이상, '원'은 오타가 잦아 선택).
_POST_RE = re.compile(
    # 금액 바로 뒤에 숫자·쉼표가 더 붙으면(533,5000원 오타) 양식이 아닌 것으로 — 앞부분만 읽어 틀린 금액이 되던 것(11-11)
    r'^\s*(\d{1,2})/(\d{1,2})\s+([GRNgrn])\s+((?:\d{1,3}(?:,\d{3})+|\d{4,}))(?![\d,])\s*(?:원|운|왼|웜)?\s*(.*)$',
    re.S)
# 끝에 붙는 회신 요청 꼬리 — 거래처 이름에서 뺀다
_TAIL_RE = re.compile(r'\s*(?:회신|확인)\s*(?:바랍니다|부탁드립니다|부탁드려요|요청드립니다|해\s*주세요)?\.?\s*$')


_CODE_IN_POST_RE = re.compile(r'(?<![A-Za-z0-9])([GRNP]\d{3,5}-[A-Za-z]{1,4})(?![A-Za-z0-9])', re.I)


def split_code(partner: str) -> tuple:
    """'서울크레인 G4160-JW' → ('서울크레인', 'G4160-JW'). 글에 공사 코드를 같이 쓰면 거래처 이름에 섞여
    들어갔다(11-09 #6). 코드가 없으면 (그대로, '')."""
    m = _CODE_IN_POST_RE.search(partner or '')
    if not m:
        return (partner or ''), ''
    rest = (partner[:m.start()] + ' ' + partner[m.end():])
    return re.sub(r'\s{2,}', ' ', rest).strip(' ·,/-'), m.group(1).upper()


def parse_invoice_post(text: str) -> Optional[Dict]:
    """첫 줄에서 날짜·G/R/N·금액·거래처를 읽는다. 양식이 아니면 None.

    거래처는 꼬리('회신 바랍니다')를 뗀 나머지 첫 줄 전체(상호 뒤 설명이 붙는 경우도 그대로 둠).
    """
    first = (text or '').strip().split('\n', 1)[0]
    m = _POST_RE.match(first)
    if not m:
        return None
    month, day, grn, amount_s, rest = m.groups()
    rest = re.sub(r'[*_~`]', '', rest or '').strip()
    rest = _TAIL_RE.sub('', rest).strip()
    try:
        amount = int(amount_s.replace(',', ''))
    except ValueError:
        return None
    partner, code = split_code(rest)
    out = {
        'date_md': f'{int(month):02d}/{int(day):02d}',
        'grn': grn.upper(),
        'amount': amount,
        'partner': partner[:60],
    }
    if code:   # 글에 적힌 공사 코드 — 현장 지정 창에서 미리 고르고, 다르게 고르면 알림(11-09)
        out['code'] = code
    return out


def post_summary(parsed: Dict) -> str:
    """'10/29 G 533,500원 강원포장' — 모달·답글 머리줄."""
    parts = [parsed.get('date_md', ''), parsed.get('grn', ''),
             f"{int(parsed.get('amount') or 0):,}원", parsed.get('partner', '')]
    return ' '.join(p for p in parts if p)


# ── 프로젝트 검색 ─────────────────────────────────────────────────────
def _codenum(code: str) -> int:
    m = re.search(r'\d+', code or '')
    return int(m.group()) if m else 0


def _special_options(query: str, specials=None) -> List[dict]:
    q = (query or '').strip().lower()
    out = []
    # specials=None 이면 자금집행용 둘(프로젝트0·공통비). 양식 전용(현장 없음)은 이름을 넣어 줄 때만
    for code, label in {**SPECIAL_LABELS, **_FORM_ONLY_LABELS}.items():
        if (code not in specials) if specials is not None else (code not in SPECIAL_LABELS):
            continue
        if not q or q in f'{code} {label}'.lower():
            out.append({'text': {'type': 'plain_text', 'text': label[:75]}, 'value': code})
    return out


def build_project_options(recs: List[dict], query: str, manager_name: Optional[str] = None,
                          limit: int = 30, label_fn=None, specials=None) -> List[dict]:
    """현장 지정 모달 프로젝트 목록.

    - 타이핑: 코드·사업자명·주소·거래처명에 들어 있으면 → 최근(코드 큰) 순.
    - 빈 클릭: 내 담당 프로젝트 최근 순(수금 여부 무관 — 부대비용은 공사 끝난 뒤에도 옴).
    - 프로젝트0·공통비는 빈 클릭이면 맨 위, 타이핑이면 맞을 때만.
    - specials: 보여줄 특수 항목 코드(None=전부). 법인서류 요청은 공통비만.
    """
    q = (query or '').strip().lower()
    label_fn = label_fn or (lambda c, b, a: f'{c} | {b} | {a}'[:75])
    matched = []
    for r in recs or []:
        code = str(r.get('프로젝트 코드') or '').strip()
        if not code:
            continue
        biz = str(r.get('사업자명') or '').strip()
        addr = str(r.get('현장 주소') or '').strip()
        if q:
            partner = str(r.get('거래처명') or '').strip()
            if q not in f'{code} {biz} {addr} {partner}'.lower():
                continue
        else:
            if not manager_name or str(r.get('담당자') or '').strip() != manager_name:
                continue
        matched.append((_codenum(code), code, label_fn(code, biz, addr)))
    matched.sort(key=lambda t: t[0], reverse=True)
    special = _special_options(query, specials)
    room = max(0, limit - len(special))
    opts = [{'text': {'type': 'plain_text', 'text': label}, 'value': code}
            for _n, code, label in matched[:room]]
    return special + opts if not q else opts + special


def project_display(code: str, rec: Optional[dict]) -> str:
    """답글용 한 줄 — 'G4049-JW · 광명 양지로 21 …' / '프로젝트0 · 회사 재고'."""
    if code in SPECIAL_LABELS:
        return SPECIAL_LABELS[code].split(' (')[0]
    if code == NO_SITE_CODE:
        return NO_SITE_CODE
    addr = str((rec or {}).get('현장 주소') or '').strip()
    return f'{code} · {addr}' if addr else code


# ── 장부 (users.db) ───────────────────────────────────────────────────
_DDL = '''
    CREATE TABLE IF NOT EXISTS purchase_invoice_assignments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        channel TEXT NOT NULL,
        msg_ts TEXT NOT NULL,
        date_md TEXT,
        grn TEXT,
        amount INTEGER,
        partner TEXT,
        project_code TEXT NOT NULL,
        cost_category TEXT NOT NULL,
        memo TEXT,
        assigned_by TEXT,
        assigned_name TEXT,
        assigned_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        reply_ts TEXT,
        history TEXT,
        UNIQUE(channel, msg_ts)
    )
'''
_table_ready = False


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


def _ensure_table(conn) -> None:
    global _table_ready
    if _table_ready:
        return
    conn.execute(_DDL)
    conn.execute('CREATE INDEX IF NOT EXISTS idx_pinv_project ON purchase_invoice_assignments(project_code)')
    conn.commit()
    _table_ready = True


def get_assignment(channel: str, msg_ts: str) -> Optional[dict]:
    with _db()._get_connection() as conn:
        _ensure_table(conn)
        row = conn.execute(
            'SELECT * FROM purchase_invoice_assignments WHERE channel=? AND msg_ts=?',
            (channel, msg_ts)).fetchone()
        return dict(row) if row else None


def save_assignment(channel: str, msg_ts: str, parsed: Dict, project_code: str,
                    cost_category: str, memo: str, user_id: str, user_name: str) -> Dict:
    """지정 저장(같은 글이면 덮어쓰고 이전 값은 history 에 쌓음). 저장된 행을 돌려준다."""
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with _db()._get_connection() as conn:
        _ensure_table(conn)
        prev = conn.execute(
            'SELECT * FROM purchase_invoice_assignments WHERE channel=? AND msg_ts=?',
            (channel, msg_ts)).fetchone()
        if prev:
            prev = dict(prev)
            hist = json.loads(prev.get('history') or '[]')
            hist.append({k: prev.get(k) for k in (
                'project_code', 'cost_category', 'memo', 'assigned_by', 'assigned_name', 'assigned_at')})
            conn.execute(
                '''UPDATE purchase_invoice_assignments
                   SET project_code=?, cost_category=?, memo=?, assigned_by=?, assigned_name=?,
                       assigned_at=?, history=?
                   WHERE id=?''',
                (project_code, cost_category, memo, user_id, user_name, now,
                 json.dumps(hist, ensure_ascii=False), prev['id']))
        else:
            conn.execute(
                '''INSERT INTO purchase_invoice_assignments
                   (channel, msg_ts, date_md, grn, amount, partner, project_code, cost_category,
                    memo, assigned_by, assigned_name, assigned_at, history)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,'[]')''',
                (channel, msg_ts, parsed.get('date_md'), parsed.get('grn'),
                 int(parsed.get('amount') or 0), parsed.get('partner'), project_code,
                 cost_category, memo, user_id, user_name, now))
        conn.commit()
        row = conn.execute(
            'SELECT * FROM purchase_invoice_assignments WHERE channel=? AND msg_ts=?',
            (channel, msg_ts)).fetchone()
        return dict(row)


def recent_partners(limit: int = 300) -> List[str]:
    """현장 지정된 계산서 글의 거래처(최근 순) — 자금집행 창 상호 검색 목록에 보탠다."""
    with _db()._get_connection() as conn:
        _ensure_table(conn)
        return [r['partner'] for r in conn.execute(
            'SELECT partner FROM purchase_invoice_assignments ORDER BY id DESC LIMIT ?', (limit,)).fetchall()
            if r['partner']]


def set_reply_ts(row_id: int, reply_ts: str) -> None:
    with _db()._get_connection() as conn:
        _ensure_table(conn)
        conn.execute('UPDATE purchase_invoice_assignments SET reply_ts=? WHERE id=?', (reply_ts, row_id))
        conn.commit()


# ── 슬랙 블록 ─────────────────────────────────────────────────────────
def button_value(channel: str, msg_ts: str, parsed: Dict) -> str:
    """버튼 value — 원 글 위치 + 읽은 값(계산서봇은 원 글을 다시 읽지 않아도 되게)."""
    d = {'ch': channel, 'ts': msg_ts, 'd': parsed.get('date_md', ''),
         'g': parsed.get('grn', ''), 'a': int(parsed.get('amount') or 0),
         'p': parsed.get('partner', '')}
    if parsed.get('code'):
        d['c'] = parsed['code']
    return json.dumps(d, ensure_ascii=False)


def parsed_from_value(v: Dict) -> Dict:
    # 예전 버튼은 'c' 가 없고 거래처에 코드가 섞여 있다 — 여기서 떼어 낸다
    partner, code = split_code(v.get('p', ''))
    out = {'date_md': v.get('d', ''), 'grn': v.get('g', ''),
           'amount': int(v.get('a') or 0), 'partner': partner}
    if v.get('c') or code:
        out['code'] = v.get('c') or code
    return out


def parsed_from_row(row: Dict) -> Dict:
    return {k: row.get(k) for k in ('date_md', 'grn', 'amount', 'partner')}


def prompt_blocks(channel: str, msg_ts: str, parsed: Dict) -> List[dict]:
    """계산서 글 스레드에 다는 버튼 메시지."""
    from dashboard.services.e_approval import spaced
    value = button_value(channel, msg_ts, parsed)
    return spaced('어느 현장 건인지 골라 주세요. 고르면 이 글의 고정이 풀립니다.', [{
        'type': 'button', 'action_id': ACTION_OPEN, 'style': 'primary', 'value': value,
        'text': {'type': 'plain_text', 'text': '📍 현장 지정', 'emoji': True},
    }])


def prompt_done_blocks() -> List[dict]:
    """지정 뒤 버튼 메시지 — 버튼을 빼고 한 줄만 남긴다(결과는 아래 답글)."""
    return [{'type': 'context', 'elements': [
        {'type': 'mrkdwn', 'text': '📍 현장 지정 완료 — 결과는 아래 답글에 있어요.'}]}]


def _initial(name) -> str:
    from dashboard.services.e_approval import initial
    return initial(name)


def result_text(row: Dict, project_line: str, extra: str = '') -> str:
    when = str(row.get('assigned_at') or '')[5:16].replace('-', '.')
    memo = (row.get('memo') or '').strip()
    lines = [f"✅ *{project_line}*",
             f"{cost_label(row.get('cost_category'))} · 지정 {_initial(row.get('assigned_name'))} {when}"]
    if memo:
        lines.append(f'메모: {memo}')
    if json.loads(row.get('history') or '[]'):
        lines.append('_(다시 지정됨)_')
    if extra:
        lines.append(extra)
    return '\n'.join(lines)


def result_blocks(row: Dict, project_line: str, channel: str, msg_ts: str, extra: str = '') -> List[dict]:
    from dashboard.services.e_approval import spaced
    value = button_value(channel, msg_ts, parsed_from_row(row))
    return spaced(result_text(row, project_line, extra), [{
        'type': 'button', 'action_id': ACTION_OPEN, 'value': value,
        'text': {'type': 'plain_text', 'text': '✏️ 다시 지정', 'emoji': True},
    }])


ACTION_TRANSFER_TOGGLE = 'pinv_transfer_toggle'   # 「이체 필요」 체크 → 이체 칸 펼침


def transfer_requested(view_state_values: Dict) -> bool:
    sel = ((view_state_values.get('transfer_toggle') or {}).get(ACTION_TRANSFER_TOGGLE) or {}).get('selected_options')
    return bool(sel)


def modal_view(parsed: Dict, meta: Dict, current: Optional[Dict] = None,
               initial_project_option: Optional[dict] = None, transfer: bool = False,
               payee: Optional[Dict] = None, approval: Optional[Dict] = None,
               candidate: Optional[Dict] = None) -> dict:
    """현장 지정 모달.

    current   = 이미 지정된 행(다시 지정 시 기본값)
    transfer  = 「이체 필요」 체크됨 → 증빙·이체 칸 표시 (payee = 기억한 계좌로 미리 채움)
    approval  = 이 계산서 글에서 이미 올린 자금집행 → 체크 대신 진행 상태만
    candidate = 계산서보다 먼저 올린 자금집행 중 G/R·금액·거래처가 같은 것 → 지정하면 연결
    """
    cost_opts = [{'text': {'type': 'plain_text', 'text': label}, 'value': v}
                 for v, label in COST_CATEGORIES]
    proj_el = {
        'type': 'external_select', 'action_id': ACTION_PROJECT, 'min_query_length': 0,
        'placeholder': {'type': 'plain_text', 'text': '코드·상호·주소로 검색 (누르면 내 담당 현장이 최근 순으로 나와요)'},
    }
    if initial_project_option:
        proj_el['initial_option'] = initial_project_option
    cost_el = {'type': 'static_select', 'action_id': 'cost',
               'placeholder': {'type': 'plain_text', 'text': '원가 분류'},
               'options': cost_opts}
    cur_cost = (current or {}).get('cost_category')
    if cur_cost in _COST_LABEL:
        cost_el['initial_option'] = {'text': {'type': 'plain_text', 'text': _COST_LABEL[cur_cost]},
                                     'value': cur_cost}
    memo_el = {'type': 'plain_text_input', 'action_id': 'memo',
               'placeholder': {'type': 'plain_text', 'text': '필요할 때만 (예: 실외기 반입 크레인)'}}
    if (current or {}).get('memo'):
        memo_el['initial_value'] = current['memo']
    return {
        'type': 'modal',
        'callback_id': CALLBACK_SUBMIT,
        'title': {'type': 'plain_text', 'text': '현장 지정'},
        'submit': {'type': 'plain_text', 'text': '지정'},
        'close': {'type': 'plain_text', 'text': '취소'},
        'private_metadata': json.dumps(meta),
        'blocks': [
            {'type': 'section', 'text': {'type': 'mrkdwn', 'text': f'*{post_summary(parsed)}*'}},
            *([{'type': 'context', 'elements': [{'type': 'mrkdwn', 'text':
                f"📝 계산서 글에 적힌 현장: *{parsed['code']}* — 다른 현장을 고르면 한 번 더 확인해요."}]}]
              if parsed.get('code') else []),
            {'type': 'input', 'block_id': 'project', 'label': {'type': 'plain_text', 'text': '현장'},
             'element': proj_el},
            {'type': 'input', 'block_id': 'cost', 'label': {'type': 'plain_text', 'text': '원가 분류'},
             'element': cost_el},
            {'type': 'input', 'block_id': 'memo', 'optional': True,
             'label': {'type': 'plain_text', 'text': '메모'}, 'element': memo_el},
            *_transfer_blocks(transfer, payee, approval, candidate),
        ],
    }


def _transfer_blocks(transfer: bool, payee: Optional[Dict], approval: Optional[Dict],
                     candidate: Optional[Dict]) -> List[dict]:
    from dashboard.services import e_approval as ea
    if approval:
        label = ea.STATUS_LABELS.get(approval.get('status'), ('', ''))
        return [{'type': 'context', 'elements': [{'type': 'mrkdwn',
                 'text': f"💸 {ea.title(approval)} {label[0]} {label[1]} — 이체 요청은 이미 올라가 있어요."}]}]
    if candidate:
        return [{'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': (
            f"💡 {ea.title(candidate)} ({ea.initial(candidate.get('requester_name'))} · {int(candidate.get('amount') or 0):,}원)"
            '과 같은 건으로 보여요. 지정하면 이 계산서를 그 결재에 연결합니다.')}]}]
    opt = {'text': {'type': 'plain_text', 'text': '이체 필요 — 자금집행 결재 올리기'}, 'value': 'yes'}
    toggle = {'type': 'checkboxes', 'action_id': ACTION_TRANSFER_TOGGLE, 'options': [opt]}
    if transfer:
        toggle['initial_options'] = [opt]
    blocks = [{'type': 'actions', 'block_id': 'transfer_toggle', 'elements': [toggle]}]
    if transfer:
        ev = [{'text': {'type': 'plain_text', 'text': l}, 'value': v} for v, l in ea.EVIDENCE_TYPES]
        blocks.append({'type': 'input', 'block_id': 'evidence', 'label': {'type': 'plain_text', 'text': '증빙'},
                       'element': {'type': 'static_select', 'action_id': 'evidence', 'options': ev,
                                   'initial_option': ev[0]}})
        if payee:   # 예전에 보낸 계좌로 채웠으면 확인 안내(11-01)
            blocks.append(ea.payee_saved_note())
        blocks += ea.transfer_input_blocks(payee)
    return blocks
