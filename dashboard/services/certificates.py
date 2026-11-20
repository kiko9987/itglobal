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
_MIGRATE = ['ALTER TABLE cert_requests ADD COLUMN company TEXT', 'ALTER TABLE cert_issues ADD COLUMN company TEXT']
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
            for sql in _MIGRATE:                    # 두 법인 선택(11-20) — 옛 표에 회사 칸
                try:
                    conn.execute(sql)
                except Exception:
                    pass
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
            _sel('c_company', '발급 회사', [company_label(k) for k in COMPANY_ORDER]),
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
    return {'ctype': sel('c_type'), 'company': company_key(sel('c_company')),
            'purpose': sel('c_purpose'), 'dest': txt('c_dest'), 'ssn': sel('c_ssn'),
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
            'INSERT INTO cert_requests (requester_id, requester_name, ctype, company, purpose, dest, ssn, receive, '
            'copies, need_by, memo, status, created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)',
            (uid, name, vals['ctype'], vals.get('company') or None, vals['purpose'], vals['dest'] or None, vals['ssn'],
             vals['receive'], copies, vals['need_by'] or None, vals['memo'] or None, PENDING, _now()))
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
    rows = ([('발급 회사', company_label(row['company']))] if row.get('company') in COMPANIES else []) + [
            ('용도', f"{row['purpose']}" + (f" ({row['dest']})" if row.get('dest') else '')),
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
            # 11-20: [📄 바로 발급] = 인사카드로 직인 PDF 를 만들어 보냄. 서식이 다르거나 영문 등은 [📎 파일 올려 발급]
            els = [_btn('📄 바로 발급', A_AUTO, row['id'], 'primary'), _btn('📎 파일 올려 발급', A_ISSUE, row['id']),
                   _btn('반려', A_REJECT, row['id'])]
        else:
            els = [_btn('신청 취소', A_CANCEL, row['id'])]
    return spaced(card_text(row)[:2900], els)


# ── 자동 발급 (2026-11-20, 플렉스 「문서·증명서」처럼 직인 찍힌 PDF) ─────────────────────────
# 칸은 플렉스 발급본과 같게(기본 정보·재직 정보·회사 정보·발급 용도).
# 두 법인(11-20 부장님): 직원 소속이 (주)아이티글로벌·(주)아이티글로벌그룹으로 나뉘어 있지만 같은 회사로 보고
# 재직증명서는 원하는 쪽으로 발급해 준다(나중에 그룹으로 다 전적). 신청 창·그룹웨어에서 발급 회사를 고른다.
# 회사 정보는 플렉스 자료(아이티글로벌=회사 정보, 그룹=플렉스 발급본 40건). 바뀌면 .env CERT_<ITG|GROUP>_<NAME|CEO|BIZ_NO|ADDRESS|TEL>.
# 직인은 instance/cert/seal_<itg|group>.png(깃 밖). 그룹은 플렉스 설정의 직인(seal.png). 그림이 없는 회사는 바로 발급을 막는다.
A_AUTO = 'cert_auto'
_DDL.append('''CREATE TABLE IF NOT EXISTS cert_issues (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    person_id INTEGER NOT NULL,
    name TEXT,
    ctype TEXT NOT NULL,
    purpose TEXT,
    ssn_included INTEGER NOT NULL DEFAULT 0,
    group_date INTEGER NOT NULL DEFAULT 0,
    request_id INTEGER,
    issued_by TEXT,
    hr_doc_id INTEGER,
    created_at TEXT NOT NULL
)''')

COMPANIES = {
    'itg': {'name': '(주)아이티글로벌', 'ceo': '박정우', 'biz_no': '884-81-01425',
            'address': '경기도 김포시 걸포로 134-84(걸포동) 1층', 'tel': '1899-1377'},
    'group': {'name': '(주)아이티글로벌그룹', 'ceo': '김영숙', 'biz_no': '533-88-02845',
              'address': '경기도 김포시 걸포로 134-84(걸포동) 4층', 'tel': '1899-1377'},
}
COMPANY_ORDER = ['itg', 'group']


def company(key: str) -> Dict[str, str]:
    import os
    if key not in COMPANIES:
        raise ValueError('발급 회사를 골라 주세요.')
    return {k: (os.getenv(f'CERT_{key.upper()}_{k.upper()}') or '').strip() or v for k, v in COMPANIES[key].items()}


def company_label(key: str) -> str:
    return company(key)['name'] if key in COMPANIES else ''


def company_key(label: str) -> str:
    return next((k for k in COMPANY_ORDER if company_label(k) == label), '')


def seal_path(key: str) -> str:
    import os
    from pathlib import Path
    d = (os.getenv('CERT_SEAL_DIR') or '').strip() or str(Path(__file__).resolve().parents[2] / 'instance' / 'cert')
    p = os.path.join(d, f'seal_{key}.png')
    if key == 'group' and not os.path.isfile(p):
        p = os.path.join(d, 'seal.png')               # 플렉스에서 받은 그룹 직인(11-20 처음 이름)
    return p


def _paths() -> Dict[str, str]:
    import os
    from pathlib import Path
    root = Path(__file__).resolve().parents[2]
    fonts = (os.getenv('CERT_FONT_DIR') or '').strip() or r'C:\Windows\Fonts'
    return {'regular': os.path.join(fonts, 'malgun.ttf'), 'bold': os.path.join(fonts, 'malgunbd.ttf')}


def _kdate(s: Optional[str]) -> str:
    if not s:
        return ''
    try:
        d = date.fromisoformat(str(s)[:10])
    except ValueError:
        return str(s)
    return f'{d.year}년 {d.month}월 {d.day}일'


def cert_rows(person: dict, ctype: str, purpose: str, ssn_full: str = '', group_date: bool = False,
              company_key: str = 'group') -> List[tuple]:
    """[(구역 제목, [(칸, 값)])] — PDF 와 시험이 같이 쓴다."""
    start = person.get('group_hire_date') if group_date and person.get('group_hire_date') else person.get('hire_date')
    end = person.get('resign_date') if person.get('status') == '퇴사' else ''
    period = f"{_kdate(start)} ~ {_kdate(end) if end else '현재'}" if start else ''
    basic = [('성명', person.get('name') or ''), ('생년월일', _kdate(person.get('birth_date')))]
    if ssn_full:
        basic.append(('주민등록번호', ssn_full))
    basic.append(('주소', person.get('address') or ''))
    employed = ctype == '재직증명서'
    work = [('재직 기간' if employed else '근무 기간', period),
            ('소속', person.get('department') or ''),
            ('직위', person.get('job_rank') or person.get('job_title') or ''),
            ('직무', person.get('job_role') or '')]
    c = company(company_key)
    comp = [('회사명', c['name']), ('사업자등록번호', c['biz_no']), ('소재지', c['address']), ('대표 전화', c['tel'])]
    return [('기본 정보', basic), ('재직 정보' if employed else '경력 정보', work), ('회사 정보', comp),
            ('발급 용도', [('용도', purpose or '')])]


def render_pdf(person: dict, ctype: str, purpose: str, doc_no: str, ssn_full: str = '', group_date: bool = False,
               today: Optional[date] = None, company_key: str = 'group') -> bytes:
    from dashboard.services import pdf_mini as pm
    today = today or date.today()
    p = _paths()
    d = pm.Doc(title=f"{ctype} - {person.get('name') or ''}")
    d.font('R', pm.TTF.load(p['regular']))
    d.font('B', pm.TTF.load(p['bold']))
    W = pm.A4[0]
    L, R, cx = 56, W - 56, W / 2
    ink, sub, rule, fill = (25, 25, 25), (105, 105, 105), (200, 200, 200), (244, 245, 247)
    d.text(L, 40, f'발급번호  {doc_no}', 8.5, 'R', sub)
    d.text(cx, 76, ctype, 24, 'B', ink, align='center', spacing=8)
    y, label_w, row_h = 138, 112, 25
    for title, rows in cert_rows(person, ctype, purpose, ssn_full, group_date, company_key):
        d.text(L, y, title, 10.5, 'B', ink)
        y += 19
        d.line(L, y, R, y, 1.0, (70, 70, 70))
        for label, val in rows:
            d.rect(L, y + 0.4, label_w, row_h - 0.8, fill)
            d.text(L + 12, y + 7.5, label, 9.5, 'R', sub)
            val, size = (val or '-'), 9.5
            while d.text_width(val, size, 'R') > R - L - label_w - 24 and size > 7:
                size -= 0.5
            d.text(L + label_w + 12, y + 7.5 + (9.5 - size) / 2, val, size, 'R', ink)
            y += row_h
            d.line(L, y, R, y, 0.5, rule)
        y += 22
    d.text(cx, max(y + 18, 632), '위와 같이 재직 중임을 증명합니다.' if ctype == '재직증명서' else '위의 경력을 증명합니다.',
           12, 'R', ink, align='center')
    d.text(cx, max(y + 48, 664), _kdate(today.isoformat()), 11, 'R', ink, align='center')
    c = company(company_key)
    sign, mark = f"{c['name']}   대표이사  {c['ceo']}", '      (인)'
    sy = max(y + 96, 724)
    x0 = cx - (d.text_width(sign, 14, 'B') + d.text_width(mark, 12, 'R')) / 2
    x1 = x0 + d.text(x0, sy, sign, 14, 'B', ink)
    mark_w = d.text(x1, sy + 2, mark, 12, 'R', sub)
    in_cx = x1 + mark_w - d.text_width('(인)', 12, 'R') / 2
    d.image(open(seal_path(company_key), 'rb').read(), in_cx - 29, sy - 21, 58, 58)
    return d.output()


def ready(company_key: str) -> str:
    """자동 발급 준비 — 문제 문구(없으면 '')."""
    import os
    p = _paths()
    for k, label in (('regular', '글꼴(맑은 고딕)'), ('bold', '글꼴(맑은 고딕 굵게)')):
        if not os.path.isfile(p[k]):
            return f'{label} 파일이 없어 바로 발급을 할 수 없어요. [📎 파일 올려 발급]을 써 주세요.'
    if not os.path.isfile(seal_path(company_key)):
        return f'{company_label(company_key)} 직인 그림이 아직 없어 바로 발급을 할 수 없어요. [📎 파일 올려 발급]을 써 주세요.'
    return ''


def issue_pdf(person_id: int, ctype: str, purpose: str, issued_by: str, ssn: bool = False,
              group_date: bool = False, request_id: Optional[int] = None, today: Optional[date] = None,
              company_key: str = '') -> dict:
    """직인 PDF 를 만들고 발급 기록·인사카드 서류(발급 증명서)에 남김 → {'pdf','file_name','doc_no','issue'}.
    issued_by = 발급한 사람(메일 또는 이름) — 인사카드 수정 기록에 남는다."""
    from dashboard.services import hr
    if ctype not in TYPES:
        raise ValueError('증명서 종류가 올바르지 않습니다.')
    if company_key not in COMPANIES:
        raise ValueError('발급 회사를 골라 주세요.')
    why = ready(company_key)
    if why:
        raise RuntimeError(why)
    person = hr.get_person(person_id)
    if not person:
        raise LookupError('없는 구성원입니다.')
    if ctype == '재직증명서' and person.get('status') == '퇴사':
        raise ValueError('퇴사한 분은 재직증명서를 발급할 수 없어요. 경력증명서로 발급해 주세요.')
    ssn_full = ''
    if ssn:
        if not person.get('ssn_enc'):
            raise ValueError('인사카드에 주민등록번호가 없어요. 생년월일만으로 발급하거나 인사카드에 먼저 넣어 주세요.')
        ssn_full = hr.decrypt_ssn(person['ssn_enc'])
    today = today or date.today()
    by = (issued_by or 'system').strip()
    with _conn() as conn:
        rid = conn.execute('INSERT INTO cert_issues (person_id, name, ctype, company, purpose, ssn_included, group_date, '
                           'request_id, issued_by, created_at) VALUES (?,?,?,?,?,?,?,?,?,?)',
                           (person_id, person['name'], ctype, company_key, purpose or None, int(bool(ssn)),
                            int(bool(group_date)), request_id, by, _now())).lastrowid
    doc_no = f'{today.year}-{rid:04d}'
    pdf = render_pdf(person, ctype, purpose, doc_no, ssn_full, group_date, today, company_key)
    short = '아이티글로벌그룹' if company_key == 'group' else '아이티글로벌'
    file_name = f"{ctype}_{person['name']}_{short}_{today.strftime('%Y%m%d')}.pdf"
    saved = hr.save_document(person_id, '발급 증명서', file_name, pdf, by, trusted=True)
    with _conn() as conn:
        conn.execute('UPDATE cert_issues SET hr_doc_id=? WHERE id=?', (saved.get('id'), rid))
        row = dict(conn.execute('SELECT * FROM cert_issues WHERE id=?', (rid,)).fetchone())
    hr.log_access(person_id, by, 'cert_issue', f'{ctype} {doc_no} {company_label(company_key)}' + (' · 주민번호 포함' if ssn else ''))
    return {'pdf': pdf, 'file_name': file_name, 'doc_no': doc_no, 'issue': row}


def issues(person_id: int, limit: int = 20) -> List[dict]:
    with _conn() as conn:
        return [dict(r) for r in conn.execute('SELECT * FROM cert_issues WHERE person_id=? ORDER BY id DESC LIMIT ?',
                                              (person_id, limit)).fetchall()]


def request_opts(row: dict) -> dict:
    """슬랙 신청 → 자동 발급 옵션: 용도(제출처 포함)·주민번호 전체 여부."""
    return {'purpose': purpose_text(row.get('purpose') or '', row.get('dest') or ''),
            'ssn': (row.get('ssn') or '').startswith('주민등록번호 전체'),
            'company': row.get('company') or ''}


def purpose_text(purpose: str, dest: str = '') -> str:
    """증명서 '용도' 칸 — 용도 (제출처). '기타'에 적은 내용이 있으면 그 내용만(11-20 부장님: '기타 (…)'는 어색)."""
    purpose, dest = (purpose or '').strip(), (dest or '').strip()
    if purpose == '기타' and dest:
        return dest
    if dest:
        return f'{purpose} ({dest})' if purpose else dest
    return purpose
