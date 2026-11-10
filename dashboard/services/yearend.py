"""연말정산 자료 모으기 (2026-11-10) — 계산·신고는 세무사, 그룹웨어는 모으고·넘기고·나눠 준다.

두 방식 모두 받는다(KiKO 11-10 "둘 다 대비"):
  ① 홈택스 「간소화자료 일괄제공」 — 회사가 직원 명단을 홈택스에 등록 → 직원이 홈택스에서 동의 →
     국세청이 회사·세무사에 바로 넘김. 그룹웨어는 명단(이름·주민번호) 뽑기와 "동의했어요" 확인·알림.
  ② 간소화 PDF 직접 업로드 — 예전 방식.
공통: 부양가족 변경 확인, 간소화에 안 나오는 추가 증빙 업로드, 제출 → 경영지원 현황판(보완 요청·완료·
     미제출자 슬랙 알림·전체 자료 묶음 내려받기) → 세무사 결과(원천징수영수증)를 각자 서류함에.
서류는 인사 서류함(hr_documents)에 꼬리표 ye<귀속연도> 를 붙여 둔다. 열람은 본인 + 인사카드 관리자.
"""
from __future__ import annotations

import csv
import io
import os
import zipfile
from contextlib import contextmanager
from datetime import date, datetime
from typing import Dict, List, Optional

from dashboard.services import hr
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

PREP, OPEN, CLOSED = 'prep', 'open', 'closed'
SEASON_LABEL = {PREP: '준비 중', OPEN: '진행 중', CLOSED: '마감'}
DRAFT, SUBMITTED, RETURNED, DONE = 'draft', 'submitted', 'returned', 'done'
STATUS_LABEL = {DRAFT: '작성 중', SUBMITTED: '제출', RETURNED: '보완 요청', DONE: '완료'}
METHOD_BULK, METHOD_PDF = 'bulk', 'pdf'
KIND_TYPES = {'simple': '연말정산 간소화자료', 'extra': '연말정산 추가증빙', 'result': '원천징수영수증'}

_DDL = [
    '''CREATE TABLE IF NOT EXISTS ye_seasons (
        year INTEGER PRIMARY KEY,
        status TEXT NOT NULL DEFAULT 'prep',
        due_date TEXT,
        notice TEXT,
        updated_by TEXT,
        updated_at TIMESTAMP
    )''',
    '''CREATE TABLE IF NOT EXISTS ye_submissions (
        year INTEGER NOT NULL,
        person_id INTEGER NOT NULL,
        method TEXT,
        bulk_consent INTEGER NOT NULL DEFAULT 0,
        family_consent INTEGER NOT NULL DEFAULT 0,
        family_status TEXT,
        family_note TEXT,
        note TEXT,
        status TEXT NOT NULL DEFAULT 'draft',
        returned_reason TEXT,
        submitted_at TEXT,
        done_at TEXT,
        updated_at TEXT,
        PRIMARY KEY (year, person_id)
    )''',
]
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


def current_year(today: Optional[date] = None) -> int:
    """지금 다루는 귀속연도 — 1~3월은 작년 귀속(연말정산 기간), 그 밖엔 올해 귀속(준비)."""
    today = today or date.today()
    return today.year - 1 if today.month <= 3 else today.year


def tag(year: int) -> str:
    return f'ye{year}'


# ── 시즌(경영지원이 열고 닫음) ─────────────────────────────────────────
def season(year: int) -> dict:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM ye_seasons WHERE year = ?', (year,)).fetchone()
    s = dict(r) if r else {'year': year, 'status': PREP, 'due_date': None, 'notice': None}
    s['status_label'] = SEASON_LABEL.get(s['status'], s['status'])
    return s


def set_season(year: int, by: str, status: Optional[str] = None, due_date: Optional[str] = None,
               notice: Optional[str] = None) -> dict:
    if not hr.is_hr_admin(by):
        raise PermissionError('연말정산 설정은 경영지원·대표만 할 수 있습니다.')
    cur = season(year)
    st = status or cur['status']
    if st not in SEASON_LABEL:
        raise ValueError('상태 값이 올바르지 않습니다.')
    due = cur.get('due_date') if due_date is None else (due_date or None)
    if due:
        date.fromisoformat(due)
    nt = cur.get('notice') if notice is None else (notice.strip() or None)
    with _conn() as conn:
        conn.execute('INSERT INTO ye_seasons (year, status, due_date, notice, updated_by, updated_at) VALUES (?,?,?,?,?,?) '
                     'ON CONFLICT(year) DO UPDATE SET status=excluded.status, due_date=excluded.due_date, '
                     'notice=excluded.notice, updated_by=excluded.updated_by, updated_at=excluded.updated_at',
                     (year, st, due, nt, by.lower(), _now()))
    return season(year)


# ── 직원 제출 ─────────────────────────────────────────────────────────
def _sub(year: int, person_id: int) -> dict:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM ye_submissions WHERE year = ? AND person_id = ?', (year, person_id)).fetchone()
    if r:
        d = dict(r)
    else:
        d = {'year': year, 'person_id': person_id, 'method': None, 'bulk_consent': 0, 'family_consent': 0,
             'family_status': None, 'family_note': None, 'note': None, 'status': DRAFT, 'returned_reason': None,
             'submitted_at': None, 'done_at': None}
    for k in ('bulk_consent', 'family_consent'):
        d[k] = bool(d.get(k))
    d['status_label'] = STATUS_LABEL.get(d['status'], d['status'])
    return d


def _save_sub(year: int, person_id: int, **vals) -> None:
    cur = _sub(year, person_id)
    cur.update(vals)
    with _conn() as conn:
        conn.execute('INSERT INTO ye_submissions (year, person_id, method, bulk_consent, family_consent, family_status, '
                     'family_note, note, status, returned_reason, submitted_at, done_at, updated_at) '
                     'VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(year, person_id) DO UPDATE SET '
                     'method=excluded.method, bulk_consent=excluded.bulk_consent, family_consent=excluded.family_consent, '
                     'family_status=excluded.family_status, family_note=excluded.family_note, note=excluded.note, '
                     'status=excluded.status, returned_reason=excluded.returned_reason, submitted_at=excluded.submitted_at, '
                     'done_at=excluded.done_at, updated_at=excluded.updated_at',
                     (year, person_id, cur['method'], int(bool(cur['bulk_consent'])), int(bool(cur['family_consent'])),
                      cur['family_status'], cur['family_note'], cur['note'], cur['status'], cur['returned_reason'],
                      cur['submitted_at'], cur['done_at'], _now()))


def my_view(person: dict, year: Optional[int] = None) -> dict:
    year = year or current_year()
    docs = hr.documents(person['id'], tag=tag(year))
    return {
        'year': year, 'season': season(year), 'submission': _sub(year, person['id']),
        'families': hr.families(person['id']),
        'simple_docs': [d for d in docs if d['doc_type'] == KIND_TYPES['simple']],
        'extra_docs': [d for d in docs if d['doc_type'] == KIND_TYPES['extra']],
        'result_docs': [d for d in docs if d['doc_type'] == KIND_TYPES['result']],
        'name': person['name'],
    }


def _editable(year: int, person_id: int) -> dict:
    s = season(year)
    if s['status'] != OPEN:
        raise PermissionError('지금은 연말정산 자료를 받는 기간이 아닙니다.')
    sub = _sub(year, person_id)
    if sub['status'] not in (DRAFT, RETURNED):
        raise PermissionError('이미 제출했습니다. 고치려면 「수정하기」를 눌러 주세요.')
    return sub


def save_answers(person_id: int, year: int, data: Dict) -> dict:
    _editable(year, person_id)
    vals = {}
    if 'method' in data:
        if data['method'] not in (METHOD_BULK, METHOD_PDF, None, ''):
            raise ValueError('자료 제출 방식이 올바르지 않습니다.')
        vals['method'] = data['method'] or None
    for k in ('bulk_consent', 'family_consent'):
        if k in data:
            vals[k] = bool(data[k])
    if 'family_status' in data:
        if data['family_status'] not in ('none', 'changed', None, ''):
            raise ValueError('부양가족 변경 여부가 올바르지 않습니다.')
        vals['family_status'] = data['family_status'] or None
    for k in ('family_note', 'note'):
        if k in data:
            vals[k] = (str(data[k] or '').strip()[:1000]) or None
    _save_sub(year, person_id, **vals)
    return _sub(year, person_id)


def upload(person_id: int, year: int, kind: str, file_name: str, data: bytes, by: str) -> dict:
    """본인 자료 업로드(간소화 PDF·추가 증빙). 기간 안·제출 전에만."""
    if kind not in ('simple', 'extra'):
        raise ValueError('자료 종류가 올바르지 않습니다.')
    _editable(year, person_id)
    return hr.save_document(person_id, KIND_TYPES[kind], file_name, data, by, tag=tag(year), trusted=True)


def delete_upload(person_id: int, year: int, doc_id: int, by: str) -> None:
    _editable(year, person_id)
    d = hr.document_file(doc_id)
    if not d or d['person_id'] != person_id or d.get('tag') != tag(year) or d['doc_type'] == KIND_TYPES['result']:
        raise PermissionError('지울 수 없는 자료입니다.')
    hr.delete_document(doc_id, by, trusted=True)


def problems(view: dict) -> List[str]:
    """제출 전 빠진 것(화면 안내·제출 막기 공용)."""
    sub = view['submission']
    out = []
    if sub['method'] == METHOD_BULK:
        if not sub['bulk_consent']:
            out.append('홈택스에서 회사 자료 제공에 동의한 뒤 「동의했어요」를 체크해 주세요.')
    elif sub['method'] == METHOD_PDF:
        if not view['simple_docs']:
            out.append('홈택스 간소화 자료 PDF를 업로드해 주세요.')
    else:
        out.append('간소화 자료를 어떻게 낼지 골라 주세요.')
    if sub['family_status'] not in ('none', 'changed'):
        out.append('부양가족 변경이 있는지 골라 주세요.')
    elif sub['family_status'] == 'changed' and not (sub.get('family_note') or '').strip():
        out.append('바뀐 부양가족 내용을 적어 주세요.')
    return out


def submit(person_id: int, year: int, person: dict) -> dict:
    _editable(year, person_id)
    v = my_view(person, year)
    p = problems(v)
    if p:
        raise ValueError(p[0])
    _save_sub(year, person_id, status=SUBMITTED, submitted_at=_now(), returned_reason=None)
    hr.log_access(person_id, f'person:{person_id}', 'ye_submit', str(year))
    return _sub(year, person_id)


def reopen(person_id: int, year: int) -> dict:
    """본인이 제출을 되돌려 고치기 — 기간 안이고 경영지원이 완료 처리하기 전까지."""
    if season(year)['status'] != OPEN:
        raise PermissionError('지금은 연말정산 자료를 받는 기간이 아닙니다.')
    sub = _sub(year, person_id)
    if sub['status'] != SUBMITTED:
        raise PermissionError('제출한 상태에서만 수정할 수 있습니다.')
    _save_sub(year, person_id, status=DRAFT)
    return _sub(year, person_id)


# ── 경영지원 현황판 ────────────────────────────────────────────────────
def admin_rows(year: int) -> List[dict]:
    rows = []
    for p in hr.list_people(include_resigned=False):
        sub = _sub(year, p['id'])
        docs = hr.documents(p['id'], tag=tag(year))
        rows.append({
            'person_id': p['id'], 'name': p['name'], 'department': p.get('department') or '',
            'access': p.get('access'), 'job_role': p.get('job_role') or '',
            'status': sub['status'], 'status_label': sub['status_label'], 'method': sub['method'],
            'bulk_consent': sub['bulk_consent'], 'family_consent': sub['family_consent'],
            'family_status': sub['family_status'], 'family_note': sub['family_note'], 'note': sub['note'],
            'returned_reason': sub['returned_reason'], 'submitted_at': sub['submitted_at'],
            'simple': sum(1 for d in docs if d['doc_type'] == KIND_TYPES['simple']),
            'extra': sum(1 for d in docs if d['doc_type'] == KIND_TYPES['extra']),
            'result': sum(1 for d in docs if d['doc_type'] == KIND_TYPES['result']),
        })
    return rows


def set_status(person_id: int, year: int, status: str, by: str, reason: str = '') -> dict:
    if not hr.is_hr_admin(by):
        raise PermissionError('경영지원·대표만 할 수 있습니다.')
    if status == RETURNED:
        if not reason.strip():
            raise ValueError('보완 요청 사유를 적어 주세요.')
        _save_sub(year, person_id, status=RETURNED, returned_reason=reason.strip()[:500])
    elif status == DONE:
        _save_sub(year, person_id, status=DONE, done_at=_now())
    elif status == SUBMITTED:      # 완료를 되돌릴 때
        _save_sub(year, person_id, status=SUBMITTED, done_at=None)
    else:
        raise ValueError('상태 값이 올바르지 않습니다.')
    hr.log_access(person_id, by, f'ye_{status}', reason.strip()[:200])
    return _sub(year, person_id)


def roster_xlsx(year: int, by: str) -> bytes:
    """홈택스 일괄제공 신청용 명단(재직자 이름·주민번호 전체). 받으면 기록이 남는다."""
    if not hr.is_hr_admin(by):
        raise PermissionError('경영지원·대표만 받을 수 있습니다.')
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f'{year}년 귀속 명단'
    ws.append(['성명', '주민등록번호', '입사일', '제출 방식'])
    for p in hr.list_people(include_resigned=False):
        full = hr.get_person(p['id'])
        ssn = ''
        if full.get('ssn_enc'):
            try:
                ssn = hr.decrypt_ssn(full['ssn_enc'])
            except Exception:
                ssn = '(읽기 실패)'
        m = _sub(year, p['id'])['method']
        ws.append([p['name'], ssn, p.get('hire_date') or '',
                   {'bulk': '일괄제공', 'pdf': 'PDF 업로드'}.get(m or '', '')])
    for col, w in zip('ABCD', (12, 18, 12, 12)):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    hr.log_access(None, by, 'ye_roster', f'{year} 재직자 명단(주민번호 포함)')
    return buf.getvalue()


def bundle_zip(year: int, by: str) -> bytes:
    """세무사 전달용 — 사람별 폴더에 간소화·추가 증빙 + 요약표(CSV). 받으면 기록이 남는다."""
    if not hr.is_hr_admin(by):
        raise PermissionError('경영지원·대표만 받을 수 있습니다.')
    buf = io.BytesIO()
    summary = io.StringIO()
    w = csv.writer(summary)
    w.writerow(['성명', '부서', '상태', '제출 방식', '홈택스 동의', '부양가족 동의', '부양가족 변경', '변경 내용', '메모',
                '간소화 PDF', '추가 증빙', '제출 시각'])
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for r in admin_rows(year):
            w.writerow([r['name'], r['department'], r['status_label'],
                        {'bulk': '일괄제공', 'pdf': 'PDF 업로드'}.get(r['method'] or '', ''),
                        '예' if r['bulk_consent'] else '', '예' if r['family_consent'] else '',
                        {'none': '없음', 'changed': '있음'}.get(r['family_status'] or '', ''),
                        r['family_note'] or '', r['note'] or '', r['simple'], r['extra'], r['submitted_at'] or ''])
            for d in hr.documents(r['person_id'], tag=tag(year)):
                if d['doc_type'] == KIND_TYPES['result']:
                    continue
                f = hr.document_file(d['id'])
                if f and os.path.exists(f['path']):
                    kind = '간소화' if d['doc_type'] == KIND_TYPES['simple'] else '추가증빙'
                    z.write(f['path'], f"{r['name']}_{r['person_id']}/{kind}_{d['id']}_{d['file_name']}")
        z.writestr(f'{year}년귀속_연말정산_요약.csv', '﻿' + summary.getvalue())
    hr.log_access(None, by, 'ye_bundle', f'{year} 전체 자료 묶음')
    return buf.getvalue()


def upload_result(person_id: int, year: int, file_name: str, data: bytes, by: str) -> dict:
    """세무사가 준 근로소득 원천징수영수증 — 그 사람 서류함에(꼬리표 붙여)."""
    if not hr.is_hr_admin(by):
        raise PermissionError('경영지원·대표만 할 수 있습니다.')
    return hr.save_document(person_id, KIND_TYPES['result'], file_name, data, by, tag=tag(year))


# ── 슬랙 알림(휴가 봇 DM) ─────────────────────────────────────────────
def _dm(slack_id: str, text: str) -> bool:
    try:
        from dashboard.blueprints.slack_leave import leave_client
        client = leave_client()
        if not client or not slack_id:
            return False
        client.chat_postMessage(channel=slack_id, text=text)
        return True
    except Exception as e:
        logger.warning(f'[연말정산] 슬랙 DM 실패({slack_id}): {e}')
        return False


def _due_text(year: int) -> str:
    due = season(year).get('due_date')
    if not due:
        return ''
    d = date.fromisoformat(due)
    return f' 마감은 *{d.month}월 {d.day}일*이에요.'


def notify_pending(year: int, by: str) -> dict:
    """아직 제출 안 했거나 보완 요청을 받은 사람에게 DM."""
    if not hr.is_hr_admin(by):
        raise PermissionError('경영지원·대표만 할 수 있습니다.')
    if season(year)['status'] != OPEN:
        raise PermissionError('연말정산이 진행 중일 때만 알림을 보낼 수 있습니다.')
    sent, skipped = [], []
    for r in admin_rows(year):
        if r['status'] in (SUBMITTED, DONE):
            continue
        p = hr.get_person(r['person_id'])
        if r['status'] == RETURNED:
            text = (f'*{year}년 귀속 연말정산* 자료 보완을 요청드려요.\n사유: {r["returned_reason"] or "-"}\n'
                    f'그룹웨어 → 인사 → 연말정산에서 고친 뒤 다시 제출해 주세요.{_due_text(year)}')
        else:
            text = (f'*{year}년 귀속 연말정산* 자료를 아직 제출하지 않으셨어요.\n'
                    f'그룹웨어 → 인사 → 연말정산에서 간소화 자료 제출 방식을 고르고 제출해 주세요.{_due_text(year)}')
        (sent if _dm(p.get('slack_id') or '', text) else skipped).append(r['name'])
    hr.log_access(None, by, 'ye_notify', f'{year} 보냄 {len(sent)} · 못 보냄 {len(skipped)}')
    return {'sent': sent, 'skipped': skipped}


def notify_result(person_id: int, year: int) -> bool:
    p = hr.get_person(person_id)
    return _dm((p or {}).get('slack_id') or '',
               f'*{year}년 귀속 근로소득 원천징수영수증*이 그룹웨어 서류함에 올라왔어요.\n'
               '그룹웨어 → 인사 → 연말정산(또는 내 정보 → 서류)에서 볼 수 있어요.')
