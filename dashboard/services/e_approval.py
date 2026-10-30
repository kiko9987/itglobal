"""전자결재 엔진 (2026-10-30) — 플렉스 워크플로우 대체, 1차 양식 = 자금집행.

결재선(지출형): 요청 → ① 승인(고광일) → ② 이체 완료(경영지원 명단) → 끝.
  - 승인자가 직접 요청하면 ① 은 자동 통과.
  - 반려(사유 필수)·요청 취소(요청자·승인자)는 끝나기 전 언제든.
상태: pending(승인 대기) → approved(이체 대기) → done(이체 완료)
      pending/approved → rejected(반려) / cancelled(취소)

여기는 슬랙 I/O 없는 장부·상태 전환·카드 모양만. 슬랙 처리는 blueprints/slack_eapproval.py.
설계: docs/e_approval_design.md
"""
from __future__ import annotations

import json
import logging
import os
import re
from datetime import datetime
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

FORM_FUND = 'fund'
FORM_LABELS = {FORM_FUND: '자금집행'}

PENDING, APPROVED, DONE, REJECTED, CANCELLED = 'pending', 'approved', 'done', 'rejected', 'cancelled'
OPEN_STATUSES = (PENDING, APPROVED)
STATUS_LABELS = {
    PENDING: ('🟡', '승인 대기'),
    APPROVED: ('🔵', '이체 대기'),
    DONE: ('🟢', '이체 완료'),
    REJECTED: ('🔴', '반려'),
    CANCELLED: ('⚪', '요청 취소'),
}

EVIDENCE_TYPES = [('전자계산서', '전자계산서'), ('수기계산서', '수기계산서'), ('기타영수증', '기타영수증')]
URGENCY_TYPES = [('일반', '일반'), ('긴급', '긴급')]

# 슬랙 식별자
A_APPROVE, A_REJECT, A_TRANSFER, A_CANCEL = 'eap_approve', 'eap_reject', 'eap_transfer', 'eap_cancel'
A_NEW = 'eap_new'                     # 채널 메뉴 [💸 자금집행 요청]
A_NEW_FROM_PROJECT = 'eap_new_project'  # 공사확정 카드 [💸 자금집행 요청]
A_PROJECT = 'eap_project'             # 새 요청 모달 프로젝트 external_select
CB_NEW, CB_REJECT = 'eap_new_submit', 'eap_reject_submit'

_DEFAULT_APPROVERS = ('U04UL2ZLJAX',)  # 고광일


def _env_ids(name: str) -> tuple:
    return tuple(x.strip() for x in os.getenv(name, '').split(',') if x.strip())


def approver_ids() -> tuple:
    """① 승인할 수 있는 사람."""
    return _env_ids('EAPPROVAL_APPROVER_IDS') or _DEFAULT_APPROVERS


def executor_ids() -> tuple:
    """② 이체 완료를 누를 수 있는 사람 = 정산 승인 명단(SB·KiKO·JW·YG, 경영지원 보조 추가 시 거기에)."""
    from dashboard.services.settlement_checkers import checker_ids
    return checker_ids()


def executor_dm_ids() -> tuple:
    """이체 요청 DM 받을 사람 — 기본은 명단 첫 줄(대표 경영지원). 나머지는 카드에서 대신 누를 수 있다."""
    ids = _env_ids('EAPPROVAL_EXECUTOR_DM_IDS')
    if ids:
        return ids
    ex = executor_ids()
    return ex[:1]


def channel_id() -> str:
    return os.getenv('SLACK_EAPPROVAL_CHANNEL', '').strip()


# ── 장부 ──────────────────────────────────────────────────────────────
_DDL = [
    '''CREATE TABLE IF NOT EXISTS approvals (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        form TEXT NOT NULL,
        status TEXT NOT NULL,
        requester_id TEXT,
        requester_name TEXT,
        project_code TEXT,
        cost_category TEXT,
        content TEXT,
        evidence TEXT,
        grn TEXT,
        partner TEXT,
        payee_name TEXT,
        bank TEXT,
        account TEXT,
        amount INTEGER NOT NULL DEFAULT 0,
        urgency TEXT,
        memo TEXT,
        origin_channel TEXT,
        origin_ts TEXT,
        card_channel TEXT,
        card_ts TEXT,
        dm_refs TEXT NOT NULL DEFAULT '[]',
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )''',
    'CREATE INDEX IF NOT EXISTS idx_approvals_status ON approvals(status)',
    'CREATE INDEX IF NOT EXISTS idx_approvals_origin ON approvals(origin_channel, origin_ts)',
    'CREATE INDEX IF NOT EXISTS idx_approvals_project ON approvals(project_code)',
    '''CREATE TABLE IF NOT EXISTS approval_steps (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        approval_id INTEGER NOT NULL,
        step TEXT NOT NULL,
        actor_id TEXT,
        actor_name TEXT,
        result TEXT NOT NULL,
        reason TEXT,
        at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )''',
    'CREATE INDEX IF NOT EXISTS idx_approval_steps_aid ON approval_steps(approval_id)',
    '''CREATE TABLE IF NOT EXISTS payee_accounts (
        partner_key TEXT PRIMARY KEY,
        partner TEXT,
        payee_name TEXT,
        bank TEXT,
        account TEXT,
        last_used TIMESTAMP
    )''',
]
_table_ready = False


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


def _ensure_tables(conn) -> None:
    global _table_ready
    if _table_ready:
        return
    for ddl in _DDL:
        conn.execute(ddl)
    conn.commit()
    _table_ready = True


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _row(conn, aid: int) -> Optional[dict]:
    r = conn.execute('SELECT * FROM approvals WHERE id=?', (aid,)).fetchone()
    return dict(r) if r else None


def get(aid: int) -> Optional[dict]:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        return _row(conn, aid)


def steps(aid: int) -> List[dict]:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        return [dict(r) for r in conn.execute(
            'SELECT * FROM approval_steps WHERE approval_id=? ORDER BY id', (aid,)).fetchall()]


def find_by_origin(channel: str, ts: str) -> Optional[dict]:
    """계산서 글에서 시작한 결재(가장 최근 것)."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute(
            'SELECT * FROM approvals WHERE origin_channel=? AND origin_ts=? ORDER BY id DESC LIMIT 1',
            (channel, ts)).fetchone()
        return dict(r) if r else None


def partner_key(name: str) -> str:
    """거래처 이름 비교용 — (주)·㈜·주식회사·공백·기호 제거, 소문자."""
    s = re.sub(r'\(주\)|㈜|주식회사|\(유\)|유한회사', '', name or '')
    return re.sub(r'[\s\.\,\-\(\)·]', '', s).lower()


def find_unlinked_match(grn: str, amount: int, partner: str) -> Optional[dict]:
    """계산서 글과 G/R·금액·거래처가 같은데 아직 계산서 글에 안 묶인 자금집행(반려·취소 제외).

    딱 한 건일 때만 돌려준다(여러 건이면 사람이 고른다)."""
    key = partner_key(partner)
    if not key or not amount:
        return None
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        rows = conn.execute(
            '''SELECT * FROM approvals WHERE form=? AND origin_ts IS NULL AND amount=? AND grn=?
               AND status IN (?,?,?) ORDER BY id DESC LIMIT 20''',
            (FORM_FUND, int(amount), grn, PENDING, APPROVED, DONE)).fetchall()
    hits = [dict(r) for r in rows if partner_key(r['partner']) == key]
    return hits[0] if len(hits) == 1 else None


def link_origin(aid: int, channel: str, ts: str) -> None:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute('UPDATE approvals SET origin_channel=?, origin_ts=?, updated_at=? WHERE id=? AND origin_ts IS NULL',
                     (channel, ts, _now(), aid))
        conn.commit()


def create(form: str, requester_id: str, requester_name: str, fields: Dict) -> dict:
    """결재 생성. 승인자가 직접 요청하면 ① 자동 통과(바로 이체 대기)."""
    auto = requester_id in approver_ids()
    status = APPROVED if auto else PENDING
    now = _now()
    cols = ('project_code', 'cost_category', 'content', 'evidence', 'grn', 'partner', 'payee_name',
            'bank', 'account', 'amount', 'urgency', 'memo', 'origin_channel', 'origin_ts')
    vals = [fields.get(c) for c in cols]
    vals[cols.index('amount')] = int(fields.get('amount') or 0)
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        cur = conn.execute(
            f'''INSERT INTO approvals (form, status, requester_id, requester_name, {", ".join(cols)},
                created_at, updated_at) VALUES (?,?,?,?,{",".join("?" * len(cols))},?,?)''',
            (form, status, requester_id, requester_name, *vals, now, now))
        aid = cur.lastrowid
        conn.execute('INSERT INTO approval_steps (approval_id, step, actor_id, actor_name, result, at) VALUES (?,?,?,?,?,?)',
                     (aid, 'request', requester_id, requester_name, 'requested', now))
        if auto:
            conn.execute('INSERT INTO approval_steps (approval_id, step, actor_id, actor_name, result, reason, at) '
                         'VALUES (?,?,?,?,?,?,?)',
                         (aid, 'approve', requester_id, requester_name, 'auto', '요청자가 승인자', now))
        conn.commit()
        return _row(conn, aid)


_TRANSITIONS = {
    # 동작: (허용 이전 상태, 다음 상태, 단계 이름, 결과)
    'approve': ((PENDING,), APPROVED, 'approve', 'approved'),
    'reject': (OPEN_STATUSES, REJECTED, 'approve', 'rejected'),
    'transfer': ((APPROVED,), DONE, 'transfer', 'done'),
    'cancel': (OPEN_STATUSES, CANCELLED, 'cancel', 'cancelled'),
}


def can_act(action: str, row: dict, user_id: str) -> bool:
    """권한 — 승인·반려=승인자(이체 대기 중 반려는 이체 담당도), 이체=이체 담당, 취소=요청자·승인자."""
    if action == 'approve':
        return user_id in approver_ids()
    if action == 'reject':
        return user_id in approver_ids() or (row.get('status') == APPROVED and user_id in executor_ids())
    if action == 'transfer':
        return user_id in executor_ids()
    if action == 'cancel':
        return user_id == row.get('requester_id') or user_id in approver_ids()
    return False


def transition(aid: int, action: str, actor_id: str, actor_name: str, reason: str = '') -> Optional[dict]:
    """상태 전환(두 사람이 동시에 눌러도 한 번만 — 이전 상태 조건부 UPDATE). 실패면 None."""
    allowed, nxt, step, result = _TRANSITIONS[action]
    now = _now()
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        cur = conn.execute(
            f'UPDATE approvals SET status=?, updated_at=? WHERE id=? AND status IN ({",".join("?" * len(allowed))})',
            (nxt, now, aid, *allowed))
        if cur.rowcount != 1:
            conn.rollback()
            return None
        conn.execute('INSERT INTO approval_steps (approval_id, step, actor_id, actor_name, result, reason, at) '
                     'VALUES (?,?,?,?,?,?,?)', (aid, step, actor_id, actor_name, result, reason or None, now))
        conn.commit()
        return _row(conn, aid)


def set_card(aid: int, channel: str, ts: str) -> None:
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute('UPDATE approvals SET card_channel=?, card_ts=? WHERE id=?', (channel, ts, aid))
        conn.commit()


def add_dm_ref(aid: int, channel: str, ts: str, role: str, user_id: str) -> None:
    """결재 DM 위치 기록 — 상태가 바뀌면 이 DM 들도 같이 고친다."""
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute('SELECT dm_refs FROM approvals WHERE id=?', (aid,)).fetchone()
        refs = json.loads((r['dm_refs'] if r else None) or '[]')
        refs.append({'ch': channel, 'ts': ts, 'role': role, 'user': user_id})
        conn.execute('UPDATE approvals SET dm_refs=? WHERE id=?', (json.dumps(refs), aid))
        conn.commit()


def dm_refs(row: dict) -> List[dict]:
    return json.loads(row.get('dm_refs') or '[]')


# ── 거래처 계좌 기억 ─────────────────────────────────────────────────
def lookup_payee(partner: str) -> Optional[dict]:
    key = partner_key(partner)
    if not key:
        return None
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        r = conn.execute('SELECT * FROM payee_accounts WHERE partner_key=?', (key,)).fetchone()
        return dict(r) if r else None


def remember_payee(partner: str, payee_name: str, bank: str, account: str) -> None:
    key = partner_key(partner)
    if not (key and payee_name and bank and account):
        return
    with _db()._get_connection() as conn:
        _ensure_tables(conn)
        conn.execute(
            '''INSERT INTO payee_accounts (partner_key, partner, payee_name, bank, account, last_used)
               VALUES (?,?,?,?,?,?)
               ON CONFLICT(partner_key) DO UPDATE SET partner=excluded.partner, payee_name=excluded.payee_name,
                 bank=excluded.bank, account=excluded.account, last_used=excluded.last_used''',
            (key, partner, payee_name, bank, account, _now()))
        conn.commit()


# ── 입력 검사 ─────────────────────────────────────────────────────────
def parse_amount(s) -> int:
    """'533,500원' → 533500. 숫자가 없으면 0."""
    digits = re.sub(r'[^\d]', '', str(s or ''))
    return int(digits) if digits else 0


def normalize_account(s: str) -> str:
    """계좌번호 — 숫자와 하이픈만."""
    return re.sub(r'[^\d\-]', '', (s or '').strip())


def grn_from_code(code: str) -> str:
    c = (code or '').strip()[:1].upper()
    return c if c in ('G', 'R') else ''


def validate_transfer(payee_name: str, bank: str, account: str) -> Dict[str, str]:
    """이체 정보 오류 {block_id: 문구}."""
    errs = {}
    if not (payee_name or '').strip():
        errs['payee_name'] = '예금주를 넣어 주세요.'
    if not (bank or '').strip():
        errs['bank'] = '은행을 넣어 주세요.'
    acct = normalize_account(account)
    if len(re.sub(r'\D', '', acct)) < 8:
        errs['account'] = '계좌번호를 숫자로 넣어 주세요.'
    return errs


# ── 카드 모양 ─────────────────────────────────────────────────────────
def _fmt_when(s: str) -> str:
    return str(s or '')[5:16].replace('-', '.')


def title(row: dict) -> str:
    return f"{FORM_LABELS.get(row.get('form'), '결재')} #{row['id']}"


def step_line(row: dict, step_rows: List[dict], approver_name: str = '고광일') -> str:
    """'결재: ✅ 고광일 승인 10.30 14:20 → ⏳ 이체' 한 줄."""
    by = {s['step']: s for s in step_rows}
    st = row.get('status')
    a = by.get('approve')
    if a and a['result'] == 'auto':
        first = '✅ 승인 (요청자가 승인자)'
    elif a and a['result'] == 'approved':
        first = f"✅ {a['actor_name']} 승인 {_fmt_when(a['at'])}"
    elif a and a['result'] == 'rejected':
        first = f"❌ {a['actor_name']} 반려 {_fmt_when(a['at'])}"
    else:
        first = f'⏳ {approver_name} 승인' if st == PENDING else f'⏸ {approver_name} 승인'
    t = by.get('transfer')
    if t:
        second = f"✅ {t['actor_name']} 이체 {_fmt_when(t['at'])}"
    elif st == APPROVED:
        second = '⏳ 이체'
    else:
        second = '⏸ 이체'
    return f'결재: {first} → {second}'


def card_text(row: dict, step_rows: List[dict], project_line: str, quote_amount: int = 0) -> str:
    emoji, label = STATUS_LABELS.get(row.get('status'), ('', row.get('status')))
    from dashboard.services.purchase_invoice_assign import cost_label
    lines = [f"{emoji} *{title(row)}* · {label}"
             + (' · 🚨 긴급' if row.get('urgency') == '긴급' else ''),
             f'*{project_line}*']
    content = (row.get('content') or '').strip()
    lines.append(cost_label(row.get('cost_category')) + (f' — {content}' if content else ''))
    lines.append(f"거래처: {row.get('partner') or '-'} ({row.get('grn') or '-'} · {row.get('evidence') or '-'})")
    lines.append(f"이체 계좌: {row.get('bank') or '-'} {row.get('account') or '-'} ({row.get('payee_name') or '-'})")
    lines.append(f"금액: *{int(row.get('amount') or 0):,}원* (VAT 포함)")
    if quote_amount:
        lines.append(f'고객 견적(VAT 별도): {int(quote_amount):,}원')
    memo = (row.get('memo') or '').strip()
    if memo:
        lines.append(f'비고: {memo}')
    lines.append(f"요청: {row.get('requester_name') or '-'} {_fmt_when(row.get('created_at'))}")
    lines.append(step_line(row, step_rows))
    rej = next((s for s in reversed(step_rows) if s['result'] in ('rejected', 'cancelled')), None)
    if rej and rej.get('reason'):
        lines.append(f"{'반려' if rej['result'] == 'rejected' else '취소'} 사유: {rej['reason']}")
    elif rej:
        lines.append(f"{rej['actor_name']} 취소 {_fmt_when(rej['at'])}")
    return '\n'.join(lines)


def _btn(text: str, action: str, aid: int, style: Optional[str] = None, confirm: Optional[dict] = None) -> dict:
    b = {'type': 'button', 'action_id': action, 'value': str(aid),
         'text': {'type': 'plain_text', 'text': text, 'emoji': True}}
    if style:
        b['style'] = style
    if confirm:
        b['confirm'] = confirm
    return b


def _confirm(title_: str, text: str, ok: str) -> dict:
    return {'title': {'type': 'plain_text', 'text': title_},
            'text': {'type': 'plain_text', 'text': text},
            'confirm': {'type': 'plain_text', 'text': ok},
            'deny': {'type': 'plain_text', 'text': '아니요'}}


def buttons(row: dict, role: str = 'all') -> List[dict]:
    """상태·받는 사람 역할에 맞는 버튼. role: all(채널 카드) / approver / executor / requester."""
    st, aid = row.get('status'), row['id']
    amt = f"{int(row.get('amount') or 0):,}원"
    out = []
    if st == PENDING and role in ('all', 'approver'):
        out += [_btn('승인', A_APPROVE, aid, 'primary'), _btn('반려', A_REJECT, aid, 'danger')]
    if st == APPROVED and role in ('all', 'executor'):
        out.append(_btn('💸 이체 완료', A_TRANSFER, aid, 'primary', _confirm(
            '이체 완료', f"{row.get('bank')} {row.get('account')} ({row.get('payee_name')})\n{amt} 이체를 마쳤나요?",
            '이체 완료')))
    if st in OPEN_STATUSES and role in ('all', 'requester'):
        out.append(_btn('요청 취소', A_CANCEL, aid, None, _confirm(
            '요청 취소', f'{title(row)} 요청을 취소할까요?', '요청 취소')))
    return out


def card_blocks(row: dict, step_rows: List[dict], project_line: str, quote_amount: int = 0,
                role: str = 'all') -> List[dict]:
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn',
                                           'text': card_text(row, step_rows, project_line, quote_amount)}}]
    btns = buttons(row, role)
    if btns:
        blocks.append({'type': 'actions', 'elements': btns})
    if row.get('status') in OPEN_STATUSES and role == 'all':
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn',
                       'text': '견적서·계산서·영수증은 이 카드 스레드에 올려 주세요.'}]})
    return blocks


def dm_intro(row: dict, role: str) -> str:
    """DM 맨 위 한 줄."""
    if role == 'approver':
        return f"{row.get('requester_name') or '-'}님이 {title(row)} 승인을 요청했어요."
    if role == 'executor':
        return f'{title(row)} 승인이 끝났어요. 이체하고 [이체 완료]를 눌러 주세요.'
    return f'{title(row)} 요청을 올렸어요. 진행되면 이 메시지가 바뀝니다.'


def menu_blocks() -> List[dict]:
    """#전자_결재 맨 위에 고정해 둘 메뉴."""
    return [
        {'type': 'section', 'text': {'type': 'mrkdwn', 'text':
            '*전자결재 요청*\n현장 계산서는 #영업_관리 계산서 글의 [📍 현장 지정]에서 바로 올릴 수 있어요. '
            '계산서 없이 먼저 보내야 하는 돈은 아래 버튼으로 요청하세요.'}},
        {'type': 'actions', 'elements': [{
            'type': 'button', 'action_id': A_NEW, 'style': 'primary',
            'text': {'type': 'plain_text', 'text': '💸 자금집행 요청', 'emoji': True}}]},
    ]


# ── 새 요청 모달 ──────────────────────────────────────────────────────
def _opts(pairs) -> List[dict]:
    return [{'text': {'type': 'plain_text', 'text': label}, 'value': v} for v, label in pairs]


def transfer_input_blocks(prefill: Optional[dict] = None, urgency: str = '일반') -> List[dict]:
    """예금주·은행·계좌·집행 시기 입력 (새 요청 모달·현장 지정 모달 공용)."""
    p = prefill or {}

    def _text(block_id, label, placeholder, value=''):
        el = {'type': 'plain_text_input', 'action_id': block_id,
              'placeholder': {'type': 'plain_text', 'text': placeholder}}
        if value:
            el['initial_value'] = value
        return {'type': 'input', 'block_id': block_id, 'label': {'type': 'plain_text', 'text': label},
                'element': el}

    urg_opts = _opts(URGENCY_TYPES)
    return [
        _text('payee_name', '예금주', '예: 윤정민', p.get('payee_name', '')),
        _text('bank', '은행', '예: 신한', p.get('bank', '')),
        _text('account', '계좌번호', '예: 110-511-825643', p.get('account', '')),
        {'type': 'input', 'block_id': 'urgency', 'label': {'type': 'plain_text', 'text': '집행 시기'},
         'element': {'type': 'radio_buttons', 'action_id': 'urgency', 'options': urg_opts,
                     'initial_option': next(o for o in urg_opts if o['value'] == urgency)}},
    ]


def new_request_view(meta: dict, fixed_project: Optional[str] = None, fixed_label: str = '') -> dict:
    """계산서 없이 시작하는 자금집행 모달. fixed_project 면 현장 고르기 없이 고정(공사확정 카드에서)."""
    from dashboard.services.purchase_invoice_assign import COST_CATEGORIES
    blocks: List[dict] = []
    if fixed_project:
        blocks.append({'type': 'section', 'text': {'type': 'mrkdwn', 'text': f'*{fixed_label or fixed_project}*'}})
    else:
        blocks.append({'type': 'input', 'block_id': 'project', 'label': {'type': 'plain_text', 'text': '현장'},
                       'element': {'type': 'external_select', 'action_id': A_PROJECT, 'min_query_length': 0,
                                   'placeholder': {'type': 'plain_text',
                                                   'text': '코드·상호·주소로 검색 (비우면 내 담당 현장)'}}})
    ev_opts = _opts(EVIDENCE_TYPES)
    blocks += [
        {'type': 'input', 'block_id': 'evidence', 'label': {'type': 'plain_text', 'text': '증빙'},
         'element': {'type': 'static_select', 'action_id': 'evidence', 'options': ev_opts,
                     'initial_option': ev_opts[0]}},
        {'type': 'input', 'block_id': 'cost', 'label': {'type': 'plain_text', 'text': '원가 분류'},
         'element': {'type': 'static_select', 'action_id': 'cost', 'options': _opts(COST_CATEGORIES),
                     'placeholder': {'type': 'plain_text', 'text': '원가 분류'}}},
        {'type': 'input', 'block_id': 'content', 'label': {'type': 'plain_text', 'text': '내용'},
         'element': {'type': 'plain_text_input', 'action_id': 'content',
                     'placeholder': {'type': 'plain_text', 'text': '예: FCU 교체 부품 구매'}}},
        {'type': 'input', 'block_id': 'partner', 'label': {'type': 'plain_text', 'text': '거래처(상호)'},
         'element': {'type': 'plain_text_input', 'action_id': 'partner',
                     'placeholder': {'type': 'plain_text', 'text': '예: 에스와이에이전시'}}},
        {'type': 'input', 'block_id': 'amount', 'label': {'type': 'plain_text', 'text': '금액 (VAT 포함, 원)'},
         'element': {'type': 'plain_text_input', 'action_id': 'amount',
                     'placeholder': {'type': 'plain_text', 'text': '예: 277200'}}},
        {'type': 'input', 'block_id': 'grn', 'optional': True,
         'label': {'type': 'plain_text', 'text': '사업자 구분'},
         'hint': {'type': 'plain_text', 'text': '비우면 공사코드(G/R)를 따릅니다. 프로젝트0·공통비는 골라 주세요.'},
         'element': {'type': 'static_select', 'action_id': 'grn',
                     'options': _opts([('G', 'G'), ('R', 'R'), ('N', 'N')]),
                     'placeholder': {'type': 'plain_text', 'text': 'G / R / N'}}},
        {'type': 'section', 'text': {'type': 'mrkdwn',
         'text': '*이체 계좌* — 예전에 같은 거래처로 요청한 적 있으면 비워 둬도 됩니다.'}},
    ]
    tb = transfer_input_blocks()
    for b in tb[:3]:
        b['optional'] = True
    blocks += tb
    blocks.append({'type': 'input', 'block_id': 'memo', 'optional': True,
                   'label': {'type': 'plain_text', 'text': '비고'},
                   'element': {'type': 'plain_text_input', 'action_id': 'memo', 'multiline': True}})
    return {
        'type': 'modal', 'callback_id': CB_NEW,
        'title': {'type': 'plain_text', 'text': '자금집행 요청'},
        'submit': {'type': 'plain_text', 'text': '요청'},
        'close': {'type': 'plain_text', 'text': '취소'},
        'private_metadata': json.dumps(meta),
        'blocks': blocks,
    }


def read_new_request(view: dict) -> Dict:
    """새 요청 모달 값 → 필드(검사 전)."""
    v = view['state']['values']

    def _sel(bid, aid=None):
        return (((v.get(bid) or {}).get(aid or bid) or {}).get('selected_option') or {}).get('value') or ''

    def _txt(bid):
        return (((v.get(bid) or {}).get(bid) or {}).get('value') or '').strip()

    meta = json.loads(view.get('private_metadata') or '{}')
    return {
        'project_code': meta.get('project') or _sel('project', A_PROJECT),
        'evidence': _sel('evidence'), 'cost_category': _sel('cost'), 'content': _txt('content'),
        'partner': _txt('partner'), 'amount': parse_amount(_txt('amount')), 'grn': _sel('grn'),
        'payee_name': _txt('payee_name'), 'bank': _txt('bank'), 'account': normalize_account(_txt('account')),
        'urgency': _sel('urgency') or '일반', 'memo': _txt('memo'),
    }


def complete_new_request(f: Dict) -> Dict[str, str]:
    """새 요청 값 보완(사업자 구분·기억한 계좌) + 오류 {block_id: 문구}. f 를 고친다."""
    errs = {}
    if not f.get('project_code'):
        errs['project'] = '현장을 골라 주세요.'
    if f.get('amount', 0) <= 0:
        errs['amount'] = '금액을 숫자로 넣어 주세요.'
    if not f.get('grn'):
        f['grn'] = grn_from_code(f.get('project_code', ''))
        if not f['grn'] and 'project' not in errs:
            errs['grn'] = '프로젝트0·공통비는 사업자 구분을 골라 주세요.'
    if not (f.get('payee_name') or f.get('bank') or f.get('account')):
        saved = lookup_payee(f.get('partner', ''))
        if saved:
            f.update(payee_name=saved['payee_name'], bank=saved['bank'], account=saved['account'])
        else:
            errs['account'] = '처음 요청하는 거래처예요. 예금주·은행·계좌를 넣어 주세요.'
    if 'account' not in errs:
        errs.update(validate_transfer(f.get('payee_name'), f.get('bank'), f.get('account')))
    return errs
