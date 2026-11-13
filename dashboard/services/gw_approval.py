"""그룹웨어 결재함 (2026-11-10, KiKO 결정 C안).

역할 나눔: 슬랙 = 알림·작성·빠른 처리, 그룹웨어 = 결재함(목록·검색·엑셀)·문서 상세·처리 단추·플렉스 보관함.
- 장부는 슬랙 전자결재와 같은 표(approvals·approval_steps, services/e_approval.py) 하나.
- 처리 단추는 슬랙 버튼과 같은 함수(blueprints/slack_eapproval.apply_action)를 부른다 →
  권한·상태 전환(동시 클릭 한 번만)·슬랙 카드 갱신·다음 사람 DM 이 슬랙에서 누른 것과 똑같다.
- 사람 구분은 인사카드의 슬랙 ID(로그인 메일 → 인사카드 → slack_id).
- 플렉스 보관함: 플렉스 「내 문서함」 내보내기(instance/flex_export/flex_workflow_export_*.json)를
  flex_approval_archive 표로 옮겨 읽기 전용으로 보여 준다(12/7 플렉스 해지 대비). 카드번호는 뒤 4자리만.
"""
from __future__ import annotations

import glob
import io
import json
import os
import re
from contextlib import contextmanager
from typing import Dict, List, Optional

from dashboard.services import e_approval as ea
from dashboard.services import hr
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

# ── 슬랙 신원 확인 ────────────────────────────────────────────────────
# 2026-11-13 보안 점검: 웹 결재는 인사카드의 slack_id 를 그대로 처리자로 썼다. 인사카드는 인사 관리자가
# 아무 값으로나 고칠 수 있어, 결재권 없는 인사 관리자가 자기 카드에 결재자 ID 를 넣으면 웹에서 승인·최종 승인
# (그리고 본인 몫인 이체 완료까지) 혼자 할 수 있었다. 이제 로그인 메일로 슬랙에 물어본 그 사람의 ID 와
# 인사카드 ID 가 같을 때만 인정한다. 슬랙 쪽 신원은 인사 관리자가 바꿀 수 없다.
_SID_CACHE: Dict[str, tuple] = {}
_SID_OK_TTL = 6 * 3600
_SID_FAIL_TTL = 300


def _lookup_slack_id(email: str) -> Optional[str]:
    """로그인 메일 → 슬랙 user ID. 슬랙에 없는 메일이면 '', 조회 자체를 못 하면 None."""
    from slack_sdk import WebClient
    # 메일 조회 권한(users:read.email)은 메인 봇에 있음 — 전자결재 봇엔 없음(2026-11-13 확인)
    for env in ('SLACK_BOT_TOKEN', 'SLACK_PROJECT_BOT_TOKEN', 'SLACK_EAPPROVAL_BOT_TOKEN'):
        tok = os.getenv(env, '').strip()
        if not tok:
            continue
        try:
            r = WebClient(token=tok, timeout=5).users_lookupByEmail(email=email)
            return ((r.get('user') or {}).get('id') or '').strip()
        except Exception as e:
            err = str(getattr(getattr(e, 'response', None), 'data', {}) or e)
            if 'users_not_found' in err:
                return ''
            logger.warning(f'[GW결재] 슬랙 신원 조회 실패({env}): {err[:120]}')
    return None


def verified_slack_id(email: str, card_sid: str) -> str:
    """인사카드 slack_id 를 슬랙이 아는 그 사람의 ID 와 같을 때만 돌려준다. 다르거나 확인 못 하면 ''."""
    email = (email or '').strip().lower()
    card_sid = (card_sid or '').strip()
    if not email or not card_sid:
        return ''
    import time
    now = time.time()
    hit = _SID_CACHE.get(email)
    if hit and hit[1] > now:
        real = hit[0]
    else:
        found = _lookup_slack_id(email)
        real = found or ''
        _SID_CACHE[email] = (real, now + (_SID_OK_TTL if found is not None else _SID_FAIL_TTL))
    if real and real == card_sid:
        return card_sid
    if real:
        logger.warning(f'[GW결재] 인사카드 슬랙 ID 가 본인과 다름 — 결재 신원 인정 안 함 ({email})')
    return ''


# ── 보는 사람 ─────────────────────────────────────────────────────────
def viewer(email: str) -> dict:
    p = hr.find_by_email(email) if email else None
    sid = verified_slack_id(email, (p or {}).get('slack_id') or '')
    approver = bool(sid) and sid in ea.approver_ids()
    final = bool(sid) and sid in ea.final_approver_ids()
    executor = bool(sid) and (sid in ea.executor_ids() or sid in ea.n_executor_ids())
    see_all = approver or final or executor or hr.is_hr_admin(email)
    return {'email': email, 'slack_id': sid, 'name': (p or {}).get('name') or '', 'see_all': see_all,
            'linked': bool(sid), 'archive_since': _archive_since(p)}


def _archive_since(p: Optional[dict]) -> str:
    """플렉스 옛 문서는 작성자 이름으로만 묶여 있다 — 생년월일이 다른 동명이인 카드가 있으면 이 사람 입사일
    이후 문서만 (2026-11-13 보안 점검 2차 #6). 같은 사람 재입사(생년월일 같음)는 예전 문서도 본인 것.
    입사일을 모르면 아무 문서도 안 보이게 '9999'. 동명이인이 없으면 ''."""
    if not p or not (p.get('name') or '').strip():
        return ''
    me_birth = (p.get('birth_date') or '').strip()
    try:
        others = [x for x in hr.same_name_people(p['name']) if x.get('id') != p.get('id')]
    except Exception:
        return '9999'
    if not any(not me_birth or (x.get('birth_date') or '').strip() != me_birth for x in others):
        return ''
    return (p.get('hire_date') or '').strip()[:10] or '9999'


def _waiting_on(row: dict, sid: str) -> bool:
    from dashboard.services.gw_home import waiting_on
    return bool(sid) and waiting_on(ea, row, sid)


def _can_view(row: dict, v: dict) -> bool:
    if v['see_all']:
        return True
    sid = v['slack_id']
    return bool(sid) and (row.get('requester_id') == sid or _waiting_on(row, sid))


# ── 목록 ──────────────────────────────────────────────────────────────
STATUS_GROUPS = {
    'open': ea.OPEN_STATUSES,
    'done': (ea.DONE,),
    'rejected': (ea.REJECTED,),
    'cancelled': (ea.CANCELLED,),
}


def _all_rows() -> List[dict]:
    with ea._db()._get_connection() as conn:
        ea._ensure_tables(conn)
        return [dict(r) for r in conn.execute('SELECT * FROM approvals ORDER BY id DESC').fetchall()]


def _summary(row: dict, v: dict) -> dict:
    emoji, label = ea.status_label(row)
    return {
        'id': row['id'], 'title': ea.title(row), 'form': row.get('form'),
        'form_label': ea.FORM_LABELS.get(row.get('form'), row.get('form') or ''),
        'status': row.get('status'), 'status_label': label, 'status_emoji': emoji,
        'requester': row.get('requester_name') or '', 'project_code': row.get('project_code') or '',
        'what': ea.what_of(row), 'amount': int(row.get('amount') or 0),
        'urgent': (row.get('urgency') or '') == '긴급', 'grn': row.get('grn') or '',
        'created_at': (row.get('created_at') or '')[:16], 'updated_at': (row.get('updated_at') or '')[:16],
        'my_turn': _waiting_on(row, v['slack_id']),
    }


def _match(row: dict, q: str) -> bool:
    if not q:
        return True
    q = q.lower().replace(' ', '')
    hay = ' '.join(str(x or '') for x in (
        ea.title(row), row.get('partner'), row.get('content'), row.get('project_code'), row.get('requester_name'),
        row.get('memo'), row.get('payee_name'), ea.what_of(row), row.get('amount'))).lower().replace(' ', '')
    return q in hay


def list_docs(v: dict, box: str = 'todo', q: str = '', form: str = '', status: str = '',
              date_from: str = '', date_to: str = '') -> dict:
    rows = _all_rows()
    sid = v['slack_id']
    todo = [r for r in rows if _waiting_on(r, sid)]
    mine = [r for r in rows if sid and r.get('requester_id') == sid]
    boxes = {'todo': todo, 'mine': mine}
    if v['see_all']:
        boxes['all'] = rows
    if box not in boxes:
        box = 'todo'
    pick = boxes[box]
    if form:
        pick = [r for r in pick if r.get('form') == form]
    if status in STATUS_GROUPS:
        pick = [r for r in pick if r.get('status') in STATUS_GROUPS[status]]
    if date_from:
        pick = [r for r in pick if (r.get('created_at') or '')[:10] >= date_from]
    if date_to:
        pick = [r for r in pick if (r.get('created_at') or '')[:10] <= date_to]
    pick = [r for r in pick if _match(r, q)]
    forms_seen = sorted({r.get('form') for r in rows if r.get('form')}, key=lambda f: ea.FORM_LABELS.get(f, f))
    return {
        'box': box, 'counts': {k: len(b) for k, b in boxes.items()},
        'items': [_summary(r, v) for r in pick],
        'forms': [{'key': f, 'label': ea.FORM_LABELS.get(f, f)} for f in forms_seen],
        'see_all': v['see_all'], 'linked': v['linked'],
    }


# ── 상세 ──────────────────────────────────────────────────────────────
_COLUMN_LABELS = [
    ('project_code', '현장 코드'), ('cost_category', '비용 구분'), ('content', '내용'), ('evidence', '증빙'),
    ('grn', '통장'), ('partner', '거래처'), ('payee_name', '예금주'), ('bank', '은행'), ('account', '계좌번호'),
    ('amount', '금액'), ('urgency', '긴급도'), ('memo', '비고 (특이사항)'),
]
_SUPPLY_EXTRA = [('link', '상품 링크'), ('qty', '수량'), ('unit_price', '단가'), ('reason', '사유')]
_SENSITIVE = re.compile(r'계좌|카드번호|CVC|유효기간')


def _fields(row: dict, v: dict) -> List[dict]:
    """칸 이름·값 목록. 계좌·카드 칸은 결재선 사람(see_all)과 요청자 본인만."""
    show_secret = v['see_all'] or (v['slack_id'] and row.get('requester_id') == v['slack_id'])
    out = []
    form = ea.forms.get(row.get('form'))
    if form:
        vals = ea.extra(row).get('f') or {}
        for f in form.fields:
            if not ea.forms.is_shown(f, vals):
                continue
            val = vals.get(f.key)
            if isinstance(val, list):
                val = ', '.join(str(x) for x in val)
            if val in (None, ''):
                continue
            if f.kind == 'money':
                try:
                    val = f'{int(str(val).replace(",", "")):,}원'
                except ValueError:
                    pass
            secret = bool(_SENSITIVE.search(f.label))
            out.append({'label': f.label, 'value': '(경영지원·결재자만)' if secret and not show_secret else str(val),
                        'wide': f.kind == 'multiline', 'section': f.section if f.section not in ('', '-') else ''})
        return out
    x = ea.extra(row)
    for key, label in _COLUMN_LABELS:
        val = row.get(key)
        if key == 'amount':
            val = f'{int(val or 0):,}원'
        elif key == 'grn' and val:
            val = ea.GRN_NAMES.get(val, val)
        if val in (None, ''):
            continue
        secret = key in ('account', 'bank', 'payee_name') and row.get('form') == ea.forms.FORM_EXPENSE
        secret = secret or key == 'account'
        out.append({'label': label, 'value': '(경영지원·결재자만)' if secret and not show_secret else str(val),
                    'wide': key in ('content', 'memo'), 'section': ''})
    if row.get('form') == ea.FORM_SUPPLY:
        for key, label in _SUPPLY_EXTRA:
            if x.get(key):
                out.append({'label': label, 'value': str(x[key]), 'wide': key in ('link', 'reason'), 'section': ''})
    return out


_STEP_NAMES = {'request': '요청', 'approve': '승인', 'final': '최종 승인', 'transfer': '처리', 'cancel': '취소', 'revise': '수정 요청',
               'resubmit': '다시 올림'}
_RESULT_NAMES = {'requested': '요청', 'approved': '승인', 'rejected': '반려', 'done': '완료', 'cancelled': '취소', 'revise': '수정 요청',
                 'resubmitted': '다시 올림', 'auto': '자동 통과'}

ACTIONS = ('approve', 'final', 'transfer', 'reject', 'revise', 'cancel')


def actions_for(row: dict, v: dict) -> List[dict]:
    sid = v['slack_id']
    if not sid:
        return []
    out = []
    w = ea.step2_word(row)
    labels = {'approve': '승인', 'final': '최종 승인', 'transfer': f'{w} 완료', 'reject': '반려',
              'revise': '수정 요청', 'cancel': '요청 취소'}
    for a in ACTIONS:
        allowed = ea._TRANSITIONS[a][0]
        if row.get('status') not in allowed or not ea.can_act(a, row, sid):
            continue
        if a == 'approve' and row.get('requester_id') == sid:
            continue
        out.append({'key': a, 'label': labels[a], 'needs_reason': a in ('reject', 'revise'),
                    'primary': a in ('approve', 'final', 'transfer')})
    return out


def detail(aid: int, v: dict) -> Optional[dict]:
    row = ea.get(aid)
    if not row or not _can_view(row, v):
        return None
    from dashboard.services.gw_home import _card_link
    steps = []
    for s in ea.steps(aid):
        name = _STEP_NAMES.get(s['step'], s['step'])
        if s['result'] == 'rejected':          # 승인 칸에서 반려 → '반려'만
            name = '반려'
        elif s['result'] == 'auto':            # 요청자가 승인자 등 → '승인 (자동)'
            name = f'{name} (자동)'
        elif s['step'] == 'transfer':          # 이체 완료·구매 완료·처리 완료
            name = f'{ea.step2_word(row)} 완료'
        steps.append({'step': name, 'result': name, 'actor': s.get('actor_name') or '',
                      'reason': s.get('reason') or '', 'at': (s.get('at') or '')[:16]})
    return dict(_summary(row, v), fields=_fields(row, v), steps=steps, actions=actions_for(row, v),
                slack_link=_card_link(row), requester_id_is_me=row.get('requester_id') == v['slack_id'])


def act(aid: int, action: str, v: dict, reason: str = '') -> Optional[str]:
    """처리 — 실패면 보여 줄 문구, 성공이면 None. 슬랙 버튼과 같은 함수."""
    if action not in ACTIONS:
        return '알 수 없는 처리입니다.'
    if not v['slack_id']:
        return '인사카드에 슬랙 계정이 연결되지 않아 처리할 수 없어요. 경영지원에 말씀해 주세요.'
    if action in ('reject', 'revise') and not (reason or '').strip():
        return '사유를 적어 주세요.' if action == 'reject' else '무엇을 고쳐야 하는지 적어 주세요.'
    from dashboard.blueprints import slack_eapproval as se
    return se.apply_action(action, aid, v['slack_id'], (reason or '').strip()[:300])


def export_xlsx(v: dict, **filters) -> bytes:
    import openpyxl
    from openpyxl.styles import Font
    data = list_docs(v, **filters)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '결재함'
    ws.append(['번호', '양식', '제목', '상태', '요청자', '현장 코드', '무엇', '금액', '긴급', '올린 때', '마지막 변경'])
    for c in ws[1]:
        c.font = Font(bold=True)
    from dashboard.utils.spreadsheet_safe import xlsx_append_text   # 요청자가 쓴 글이 수식으로 안 되게 (2026-11-13)
    for it in data['items']:
        xlsx_append_text(ws, [it['id'], it['form_label'], it['title'], it['status_label'], it['requester'],
                              it['project_code'], it['what'], it['amount'], '긴급' if it['urgent'] else '',
                              it['created_at'], it['updated_at']])
    for col, w in zip('ABCDEFGHIJK', (7, 14, 26, 12, 10, 12, 30, 14, 6, 17, 17)):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── 플렉스 보관함 ─────────────────────────────────────────────────────
_ARCHIVE_DDL = '''CREATE TABLE IF NOT EXISTS flex_approval_archive (
    doc_no TEXT PRIMARY KEY,
    title TEXT, form TEXT, status TEXT, writer TEXT,
    created_at TEXT, updated_at TEXT, attachments INTEGER NOT NULL DEFAULT 0,
    memo TEXT, fields TEXT NOT NULL DEFAULT '{}',
    amount INTEGER NOT NULL DEFAULT 0, project_code TEXT, partner TEXT,
    imported_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
)'''
_archive_ready = False

FLEX_STATUS = {'DONE': '완료', 'CANCELED': '취소', 'DECLINED': '반려', 'IN_PROGRESS': '진행 중(플렉스)'}


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


@contextmanager
def _aconn():
    global _archive_ready
    with _db()._get_connection() as conn:
        if not _archive_ready:
            conn.execute(_ARCHIVE_DDL)
            conn.execute('CREATE INDEX IF NOT EXISTS idx_flex_arch_created ON flex_approval_archive(created_at)')
            conn.commit()
            _archive_ready = True
        yield conn
        conn.commit()


def flex_form(title: str) -> str:
    """플렉스 문서 제목 → 양식 이름(법인카드 지출결의는 달·사람이 제목에 붙어 있음)."""
    t = (title or '').strip()
    if '법인카드' in t:
        return '법인카드 지출결의'
    t = re.sub(r'^[GR](?=(계약서|카드결제|보증보험))', '', t)
    t = re.sub(r'^\(매출\)\s*[GR]?', '(매출)', t)
    return t or '기타'


def _mask_card(v: str) -> str:
    d = re.sub(r'\D', '', v or '')
    return f'****-{d[-4:]}' if len(d) > 4 else (v or '')


def _num(v) -> int:
    try:
        return int(re.sub(r'[^\d\-]', '', str(v or '')) or 0)
    except ValueError:
        return 0


def latest_export(folder: Optional[str] = None) -> Optional[str]:
    folder = folder or os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                                    'instance', 'flex_export')
    files = sorted(glob.glob(os.path.join(folder, 'flex_workflow_export_*.json')))
    return files[-1] if files else None


def import_flex_export(path: str, apply: bool = True) -> dict:
    with open(path, encoding='utf-8') as fp:
        data = json.load(fp)
    docs = data.get('docs') or []
    rows = []
    for d in docs:
        f = dict(d.get('f') or {})
        for k in list(f):
            if '카드번호' in k:
                f[k] = _mask_card(f[k])
        amount = next((_num(f[k]) for k in ('총 금액(VAT포함)', '청구 금액', '금액(VAT포함)', '원 매출(VAT포함)')
                       if f.get(k)), 0)
        rows.append((str(d.get('c') or ''), d.get('t') or '', flex_form(d.get('t')), FLEX_STATUS.get(d.get('s'), d.get('s') or ''),
                     d.get('w') or '', (d.get('at') or '').replace('T', ' ')[:16], (d.get('up') or '').replace('T', ' ')[:16],
                     int(d.get('att') or 0), d.get('memo') or '', json.dumps(f, ensure_ascii=False), amount,
                     (f.get('프로젝트 코드') or f.get('프로젝트코드') or '').strip(), (f.get('상호명') or '').strip()))
    rows = [r for r in rows if r[0]]
    if apply:
        with _aconn() as conn:
            conn.executemany('INSERT OR REPLACE INTO flex_approval_archive (doc_no, title, form, status, writer, created_at, '
                             'updated_at, attachments, memo, fields, amount, project_code, partner) '
                             'VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)', rows)
    from collections import Counter
    return {'count': len(rows), 'exported_at': data.get('exported_at'),
            'by_status': dict(Counter(r[3] for r in rows)), 'by_form': dict(Counter(r[2] for r in rows).most_common(12))}


def archive_list(v: dict, q: str = '', form: str = '', date_from: str = '', date_to: str = '',
                 page: int = 1, size: int = 50) -> dict:
    where, args = [], []
    if not v['see_all']:
        if not v['name']:
            return {'items': [], 'total': 0, 'forms': [], 'page': 1, 'size': size}
        where.append('writer = ?')
        args.append(v['name'])
        if v.get('archive_since'):
            where.append('created_at >= ?')
            args.append(v['archive_since'])
    if form:
        where.append('form = ?')
        args.append(form)
    if date_from:
        where.append('created_at >= ?')
        args.append(date_from)
    if date_to:
        where.append('created_at <= ?')
        args.append(date_to + ' 99')
    if q:
        where.append("(title || ' ' || writer || ' ' || IFNULL(partner,'') || ' ' || IFNULL(project_code,'') || ' ' || "
                     "IFNULL(memo,'') || ' ' || fields || ' ' || doc_no) LIKE ?")
        args.append(f'%{q.strip()}%')
    w = ('WHERE ' + ' AND '.join(where)) if where else ''
    page = max(1, int(page or 1))
    with _aconn() as conn:
        total = conn.execute(f'SELECT COUNT(*) FROM flex_approval_archive {w}', args).fetchone()[0]
        rows = conn.execute(f'SELECT doc_no, title, form, status, writer, created_at, amount, project_code, partner, attachments '
                            f'FROM flex_approval_archive {w} ORDER BY created_at DESC, doc_no DESC LIMIT ? OFFSET ?',
                            (*args, size, (page - 1) * size)).fetchall()
        fw, fa = '', []
        if not v['see_all']:
            fw, fa = 'WHERE writer = ?', [v['name']]
            if v.get('archive_since'):
                fw, fa = fw + ' AND created_at >= ?', fa + [v['archive_since']]
        forms_ = [r[0] for r in conn.execute(f'SELECT form, COUNT(*) c FROM flex_approval_archive {fw} GROUP BY form '
                                              f'ORDER BY c DESC', fa).fetchall()]
    keys = ('doc_no', 'title', 'form', 'status', 'writer', 'created_at', 'amount', 'project_code', 'partner', 'attachments')
    return {'items': [dict(zip(keys, r)) for r in rows], 'total': total, 'forms': forms_, 'page': page, 'size': size}


def archive_detail(doc_no: str, v: dict) -> Optional[dict]:
    with _aconn() as conn:
        r = conn.execute('SELECT doc_no, title, form, status, writer, created_at, updated_at, attachments, memo, fields '
                         'FROM flex_approval_archive WHERE doc_no = ?', (doc_no,)).fetchone()
    if not r:
        return None
    keys = ('doc_no', 'title', 'form', 'status', 'writer', 'created_at', 'updated_at', 'attachments', 'memo', 'fields')
    d = dict(zip(keys, r))
    # 이름이 빈 사람(인사카드 없음)이 작성자 빈 문서를 보는 일 막기 — 목록과 같은 규칙 (2026-11-13)
    if not v['see_all'] and (not v['name'] or d['writer'] != v['name']
                             or (v.get('archive_since') and (d['created_at'] or '') < v['archive_since'])):
        return None
    try:
        fields = json.loads(d['fields'] or '{}')
    except ValueError:
        fields = {}
    d['fields'] = [{'label': k, 'value': str(val)} for k, val in fields.items() if str(val or '').strip()]
    return d
