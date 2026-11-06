# -*- coding: utf-8 -*-
"""재직·경력증명서 신청 (2026-11-06) — 플렉스 「문서·증명서」 대체.

흐름: 전자결재 메뉴 [📜 증명서 신청] → 창(종류·용도·제출처·주민번호 포함·받는 방법·부수·필요일·메모) →
신청자 DM 카드 + 경영지원 DM 카드 [📎 발급 완료][반려] → 경영지원이 서류를 직접 만들어(당분간 수동, 11-06 KiKO)
PDF 를 올리면 신청자 DM 으로 파일이 간다. 채널 카드는 없다(용도·주민번호 포함 여부가 동료에게 보이지 않게).
플렉스 때 설정: 신청 → 경영지원(SB) 승인 → 발급, 선택지(용도·주민번호 포함·그룹 입사일·서명 방식).
휴직·사직은 메뉴에 없다 — 경영지원에 직접 말한다(11-06 KiKO).
"""
import json
from contextlib import contextmanager
from datetime import date, datetime
from typing import Dict, List, Optional

TYPES = ['재직증명서', '경력증명서']
PURPOSES = ['금융기관 제출', '관공서 제출', '회사·학교 제출', '기타']
SSN_OPTS = ['생년월일만 (뒷자리 빼고)', '주민등록번호 전체']
RECEIVE = ['PDF 파일 (DM으로 받기)', '직인 찍은 원본 (사무실에서 받기)']

PENDING, ISSUED, REJECTED, CANCELLED = 'pending', 'issued', 'rejected', 'cancelled'
STATUS = {PENDING: ('🟡', '발급 대기'), ISSUED: ('🟢', '발급 완료'), REJECTED: ('🔴', '반려'),
          CANCELLED: ('⚪', '취소')}

A_OPEN, CB_REQ = 'cert_open', 'cert_submit'
A_ISSUE, CB_ISSUE = 'cert_issue', 'cert_issue_submit'
A_REJECT, CB_REJECT = 'cert_reject', 'cert_reject_submit'
A_CANCEL = 'cert_cancel'
FILE_BLOCK = 'cert_files'

_DDL = ['''CREATE TABLE IF NOT EXISTS cert_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    requester_id TEXT NOT NULL,
    requester_name TEXT,
    ctype TEXT NOT NULL,
    purpose TEXT,
    dest TEXT,
    ssn TEXT,
    receive TEXT,
    copies INTEGER DEFAULT 1,
    need_by TEXT,
    memo TEXT,
    status TEXT NOT NULL,
    handled_by TEXT,
    handled_note TEXT,
    files TEXT,
    dm_refs TEXT,
    created_at TEXT NOT NULL,
    handled_at TEXT
)''']
_table_ready = False


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


@contextmanager
def _conn():
    global _table_ready
    with _db()._get_connection() as conn:
        if not _table_ready:
            for ddl in _DDL:
                conn.execute(ddl)
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def handler_ids() -> tuple:
    """발급 DM 받는 사람 — 전자결재 이체 DM 과 같은 사람(시험 모드도 같이 따른다)."""
    from dashboard.services import e_approval as ea
    return ea.executor_dm_ids()


def can_handle(uid: str) -> bool:
    from dashboard.services import e_approval as ea
    return bool(uid) and (uid in ea.executor_ids() or uid in handler_ids())


# ── 창 ────────────────────────────────────────────────────────────────
def _plain(t: str) -> dict:
    return {'type': 'plain_text', 'text': t, 'emoji': True}


def _sel(block_id: str, label: str, opts: List[str], initial: str = '', optional: bool = False) -> dict:
    el = {'type': 'static_select', 'action_id': block_id, 'placeholder': _plain('고르기'),
          'options': [{'text': _plain(o), 'value': o} for o in opts]}
    if initial in opts:
        el['initial_option'] = {'text': _plain(initial), 'value': initial}
    return {'type': 'input', 'block_id': block_id, 'label': _plain(label), 'element': el, 'optional': optional}


def _txt(block_id: str, label: str, optional: bool = True, placeholder: str = '', multiline: bool = False,
         initial: str = '', max_len: int = 200) -> dict:
    el = {'type': 'plain_text_input', 'action_id': block_id, 'multiline': multiline, 'max_length': max_len}
    if placeholder:
        el['placeholder'] = _plain(placeholder)
    if initial:
        el['initial_value'] = initial
    return {'type': 'input', 'block_id': block_id, 'label': _plain(label), 'element': el, 'optional': optional}


def request_view(name: str) -> dict:
    return {
        'type': 'modal', 'callback_id': CB_REQ, 'title': _plain('증명서 신청'),
        'submit': _plain('신청'), 'close': _plain('취소'), 'notify_on_close': True,
        'blocks': [
            {'type': 'section', 'text': {'type': 'mrkdwn', 'text':
                f'*{name}* 님 증명서를 경영지원에 요청해요. 신청 내용은 본인과 경영지원만 봐요.'}},
            _sel('c_type', '증명서 종류', TYPES, '재직증명서'),
            _sel('c_purpose', '용도', PURPOSES),
            _txt('c_dest', '제출처', placeholder='예: ○○은행 대출, ○○구청'),
            _sel('c_ssn', '주민등록번호 표기', SSN_OPTS, SSN_OPTS[0]),
            _sel('c_receive', '받는 방법', RECEIVE, RECEIVE[0]),
            _txt('c_copies', '부수 (원본일 때)', placeholder='1', max_len=2),
            {'type': 'input', 'block_id': 'c_need', 'optional': True, 'label': _plain('필요한 날'),
             'element': {'type': 'datepicker', 'action_id': 'c_need'}},
            _txt('c_memo', '메모', placeholder='예: 영문 필요, 입사일을 그룹 입사일로', multiline=True, max_len=500),
        ],
    }


def read_request(view: dict) -> Dict[str, str]:
    v = view['state']['values']

    def sel(k):
        return ((v.get(k, {}).get(k) or {}).get('selected_option') or {}).get('value') or ''

    def txt(k):
        return ((v.get(k, {}).get(k) or {}).get('value') or '').strip()
    return {'ctype': sel('c_type'), 'purpose': sel('c_purpose'), 'dest': txt('c_dest'), 'ssn': sel('c_ssn'),
            'receive': sel('c_receive'), 'copies': txt('c_copies'),
            'need_by': (v.get('c_need', {}).get('c_need') or {}).get('selected_date') or '', 'memo': txt('c_memo')}


def validate(vals: Dict[str, str]) -> Dict[str, str]:
    errs = {}
    if vals['copies'] and (not vals['copies'].isdigit() or not 1 <= int(vals['copies']) <= 10):
        errs['c_copies'] = '1~10 사이 숫자로 적어 주세요.'
    if vals['need_by'] and vals['need_by'] < date.today().isoformat():
        errs['c_need'] = '지난 날짜예요.'
    return errs


def issue_view(row: dict) -> dict:
    original = row.get('receive', '').startswith('직인')
    return {
        'type': 'modal', 'callback_id': CB_ISSUE, 'private_metadata': json.dumps({'id': row['id']}),
        'title': _plain('증명서 발급'), 'submit': _plain('발급 완료'), 'close': _plain('취소'),
        'blocks': [
            {'type': 'section', 'text': {'type': 'mrkdwn', 'text': summary_line(row)}},
            {'type': 'input', 'block_id': FILE_BLOCK, 'optional': original, 'label': _plain('증명서 파일 (PDF)'),
             'element': {'type': 'file_input', 'action_id': FILE_BLOCK, 'max_files': 3},
             'hint': _plain('올리면 신청자 DM 으로 바로 가요.' + (' 원본만 주는 경우엔 비워 두세요.' if original else ''))},
            _txt('i_note', '신청자에게 남길 말', placeholder='예: 원본은 경영지원 책상에 있어요'),
        ],
    }


def reject_view(row: dict) -> dict:
    return {
        'type': 'modal', 'callback_id': CB_REJECT, 'private_metadata': json.dumps({'id': row['id']}),
        'title': _plain('증명서 반려'), 'submit': _plain('반려'), 'close': _plain('취소'),
        'blocks': [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': summary_line(row)}},
                   _txt('r_reason', '사유', optional=False, multiline=True, max_len=500)],
    }


def read_files(view: dict) -> List[dict]:
    el = (((view.get('state') or {}).get('values') or {}).get(FILE_BLOCK) or {}).get(FILE_BLOCK) or {}
    return [{'id': f.get('id', ''), 'name': f.get('name') or f.get('title') or 'file',
             'url': f.get('url_private_download') or f.get('url_private') or ''}
            for f in (el.get('files') or []) if f.get('id')]


# ── 장부 ──────────────────────────────────────────────────────────────
def create(uid: str, name: str, vals: Dict[str, str]) -> dict:
    copies = int(vals['copies']) if vals.get('copies', '').isdigit() else 1
    with _conn() as conn:
        cur = conn.execute(
            'INSERT INTO cert_requests (requester_id, requester_name, ctype, purpose, dest, ssn, receive, copies, '
            'need_by, memo, status, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)',
            (uid, name, vals['ctype'], vals['purpose'], vals['dest'] or None, vals['ssn'], vals['receive'], copies,
             vals['need_by'] or None, vals['memo'] or None, PENDING, _now()))
        return dict(conn.execute('SELECT * FROM cert_requests WHERE id=?', (cur.lastrowid,)).fetchone())


def get(rid: int) -> Optional[dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM cert_requests WHERE id=?', (rid,)).fetchone()
        return dict(r) if r else None


def finish(rid: int, status: str, actor: str, note: str = '', files: Optional[List[str]] = None) -> Optional[dict]:
    """대기 중일 때만 바꾼다(두 번 눌러도 한 번)."""
    with _conn() as conn:
        cur = conn.execute('UPDATE cert_requests SET status=?, handled_by=?, handled_note=?, files=?, handled_at=? '
                           'WHERE id=? AND status=?',
                           (status, actor, note or None, json.dumps(files or [], ensure_ascii=False), _now(),
                            rid, PENDING))
        if cur.rowcount != 1:
            return None
        return dict(conn.execute('SELECT * FROM cert_requests WHERE id=?', (rid,)).fetchone())


def add_ref(rid: int, ch: str, ts: str, role: str) -> None:
    with _conn() as conn:
        r = conn.execute('SELECT dm_refs FROM cert_requests WHERE id=?', (rid,)).fetchone()
        refs = json.loads((r['dm_refs'] if r else None) or '[]')
        refs.append({'ch': ch, 'ts': ts, 'role': role})
        conn.execute('UPDATE cert_requests SET dm_refs=? WHERE id=?', (json.dumps(refs), rid))


def refs(row: dict) -> List[dict]:
    return json.loads(row.get('dm_refs') or '[]')


def mine(uid: str, limit: int = 10) -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute(
            'SELECT * FROM cert_requests WHERE requester_id=? ORDER BY id DESC LIMIT ?', (uid, limit)).fetchall()]


# ── 카드 ──────────────────────────────────────────────────────────────
def summary_line(row: dict) -> str:
    bits = [row['purpose'] or '']
    if row.get('dest'):
        bits.append(row['dest'])
    return f"📜 *{row['ctype']}* — {row['requester_name']} ({' · '.join(b for b in bits if b)})"


def card_text(row: dict) -> str:
    """결재 카드와 같은 틀(11-06) — 머리 · 무엇 · 구분선 · '*항목* : 값' · 구분선 · 처리."""
    from dashboard.services.e_approval import CARD_SEP
    emoji, label = STATUS.get(row['status'], ('', row['status']))
    rows = [('용도', f"{row['purpose']}" + (f" ({row['dest']})" if row.get('dest') else '')),
            ('주민등록번호', row['ssn']),
            ('받는 방법', f"{row['receive']}" + (f" · {row['copies']}부" if (row.get('receive') or '').startswith('직인') else ''))]
    if row.get('need_by'):
        d = date.fromisoformat(row['need_by'])
        rows.append(('필요한 날', f'{d.month}월 {d.day}일'))
    if row.get('memo'):
        rows.append(('메모', row['memo']))
    lines = [f"{emoji} *증명서 신청* · {label}", f"*{row['ctype']}* · {row['requester_name']}", CARD_SEP]
    lines += [f'*{k}* : {v}' for k, v in rows]
    if row['status'] == ISSUED:
        lines += [CARD_SEP, f"✅ 발급 완료 ({row.get('handled_by') or ''})"
                  + (f": {row['handled_note']}" if row.get('handled_note') else '')]
    elif row['status'] == REJECTED:
        lines += [CARD_SEP, f"❌ 반려 사유 ({row.get('handled_by') or ''}): {row.get('handled_note') or ''}"]
    return '\n'.join(f'>{ln}' for ln in '\n'.join(lines).split('\n'))


def _btn(text: str, action: str, rid: int, style: str = '') -> dict:
    b = {'type': 'button', 'text': _plain(text), 'action_id': action, 'value': str(rid)}
    if style:
        b['style'] = style
    return b


def card_blocks(row: dict, role: str) -> List[dict]:
    """role: handler(경영지원) / requester."""
    from dashboard.services.e_approval import spaced
    els = None
    if row['status'] == PENDING:
        if role == 'handler':
            els = [_btn('📎 발급 완료', A_ISSUE, row['id'], 'primary'), _btn('반려', A_REJECT, row['id'])]
        else:
            els = [_btn('신청 취소', A_CANCEL, row['id'])]
    return spaced(card_text(row)[:2900], els)
