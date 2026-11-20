# -*- coding: utf-8 -*-
"""본인 정보 변경 요청 (2026-11-20, 플렉스 「정보 변경 승인」 대체) — 직원이 내 정보에서 요청 → 경영지원이 반영·반려.

- 연락처(전화·개인 메일·주소), 비상연락망(가족 이름·관계·연락처), 서류(등본·통장 사본 등) 세 가지.
- 지금까지는 인사카드를 경영지원만 고칠 수 있었다(직원은 말로 부탁). 요청은 대기 → [반영]하면 인사카드에 들어가고
  [반려]하면 사유가 본인에게 간다. 서류는 반영 전까지 임시 폴더(hr_docs/_pending)에만 있다.
- 같은 종류(연락처·비상연락망) 요청은 하나씩 — 확인 중이면 취소하고 다시 낸다. 서류는 여러 장 가능.
"""
import json
import os
import re
from contextlib import contextmanager
from datetime import datetime
from typing import Dict, List, Optional

PENDING, APPROVED, REJECTED, CANCELLED = 'pending', 'approved', 'rejected', 'cancelled'
INFO, FAMILY, DOC = 'info', 'family', 'doc'
KIND_LABEL = {INFO: '연락처', FAMILY: '비상연락망', DOC: '서류'}
INFO_FIELDS = [('phone', '전화'), ('personal_email', '개인 메일'), ('address', '주소')]
SELF_DOC_TYPES = ['주민등록등본', '가족관계증명서', '급여계좌사본', '신분증사본', '건강보험득실확인서', '이력서',
                  '학력증명서', '경력증명서', '장애인증명서', '국가유공자증명서', '기타']
MAX_FAMILY = 10

_DDL = ['''CREATE TABLE IF NOT EXISTS hr_change_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    person_id INTEGER NOT NULL,
    name TEXT,
    kind TEXT NOT NULL,
    payload TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'pending',
    requested_by TEXT,
    created_at TEXT NOT NULL,
    handled_by TEXT,
    handled_at TEXT,
    reason TEXT,
    staged TEXT
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


def _row(r) -> Optional[dict]:
    if not r:
        return None
    d = dict(r)
    d['payload'] = json.loads(d.get('payload') or '{}')
    d['kind_label'] = KIND_LABEL.get(d['kind'], d['kind'])
    d.pop('staged', None)
    return d


def _digits(v: str) -> str:
    return re.sub(r'\D', '', v or '')


def _check_phone(v: str, what: str = '전화') -> str:
    v = (v or '').strip()
    if v and not 9 <= len(_digits(v)) <= 11:
        raise ValueError(f'{what} 번호를 확인해 주세요(숫자 9~11자리).')
    return v


def _pending_of(person_id: int, kind: str) -> Optional[dict]:
    with _conn() as conn:
        return _row(conn.execute('SELECT * FROM hr_change_requests WHERE person_id=? AND kind=? AND status=?',
                                 (person_id, kind, PENDING)).fetchone())


def _insert(person: dict, kind: str, payload: dict, by: str, staged: str = None) -> dict:
    with _conn() as conn:
        rid = conn.execute('INSERT INTO hr_change_requests (person_id, name, kind, payload, status, requested_by, '
                           'created_at, staged) VALUES (?,?,?,?,?,?,?,?)',
                           (person['id'], person['name'], kind, json.dumps(payload, ensure_ascii=False), PENDING,
                            by, _now(), staged)).lastrowid
        return _row(conn.execute('SELECT * FROM hr_change_requests WHERE id=?', (rid,)).fetchone())


def request_info(person: dict, values: Dict, by: str) -> dict:
    """연락처 바꾸기 요청 — 바뀐 칸만 담는다."""
    if _pending_of(person['id'], INFO):
        raise ValueError('확인 중인 연락처 요청이 있어요. 취소하고 다시 내 주세요.')
    new, old = {}, {}
    for k, label in INFO_FIELDS:
        if k not in values:
            continue
        v = (values.get(k) or '').strip()
        if k == 'phone':
            v = _check_phone(v)
        if k == 'personal_email':
            v = v.lower()
            if v and ('@' not in v or len(v) > 120):
                raise ValueError('개인 메일 형식을 확인해 주세요.')
        if k == 'address' and len(v) > 200:
            raise ValueError('주소가 너무 길어요(200자까지).')
        if (person.get(k) or '') != v and not (k == 'phone' and _digits(person.get(k)) == _digits(v) and v):
            new[k], old[k] = v, person.get(k) or ''
    if not new:
        raise ValueError('바뀐 내용이 없어요.')
    return _insert(person, INFO, {'new': new, 'old': old}, by)


def request_family(person: dict, rows: List[Dict], by: str) -> dict:
    """비상연락망 바꾸기 요청 — 원하는 목록 전체(기존 줄은 id 유지, 새 줄은 id 없음, 빠진 줄은 지움)."""
    from dashboard.services import hr
    if _pending_of(person['id'], FAMILY):
        raise ValueError('확인 중인 비상연락망 요청이 있어요. 취소하고 다시 내 주세요.')
    cur = {f['id']: f for f in hr.families(person['id'])}
    clean = []
    for r in rows or []:
        name, rel = (r.get('name') or '').strip()[:30], (r.get('relation') or '').strip()[:20]
        phone = _check_phone(r.get('phone'), '연락처')
        if not (name or rel or phone):
            continue
        if not name or not rel:
            raise ValueError('이름과 관계를 모두 적어 주세요.')
        fid = r.get('id')
        if fid is not None:
            fid = int(fid)
            if fid not in cur:
                raise ValueError('없는 줄이 섞여 있어요. 화면을 새로 고쳐 주세요.')
        clean.append({'id': fid, 'name': name, 'relation': rel, 'phone': phone})
    if len(clean) > MAX_FAMILY:
        raise ValueError(f'비상연락망은 {MAX_FAMILY}명까지예요.')
    old = [{'id': f['id'], 'name': f.get('name') or '', 'relation': f.get('relation') or '', 'phone': f.get('phone') or ''}
           for f in cur.values()]
    if [(r['id'], r['name'], r['relation'], _digits(r['phone'])) for r in clean] == \
            [(r['id'], r['name'], r['relation'], _digits(r['phone'])) for r in old]:
        raise ValueError('바뀐 내용이 없어요.')
    return _insert(person, FAMILY, {'rows': clean, 'old': old}, by)


def _pending_dir() -> str:
    from dashboard.services import hr
    d = os.path.join(hr.docs_dir(), '_pending')
    os.makedirs(d, exist_ok=True)
    return d


def request_doc(person: dict, doc_type: str, file_name: str, data: bytes, by: str) -> dict:
    from dashboard.services import hr
    if doc_type not in SELF_DOC_TYPES:
        raise ValueError('서류 종류를 골라 주세요.')
    if not data:
        raise ValueError('빈 파일입니다.')
    if len(data) > hr.DOC_MAX_BYTES:
        raise ValueError('파일이 너무 커요(8MB까지).')
    if not (file_name or '').lower().endswith(hr.DOC_EXT):
        raise ValueError('올릴 수 없는 파일 형식이에요.')
    safe = hr._safe_name(file_name)
    row = _insert(person, DOC, {'doc_type': doc_type, 'file_name': safe, 'size': len(data)}, by)
    path = os.path.join(_pending_dir(), f"{row['id']}_{safe}")
    with open(path, 'wb') as fp:
        fp.write(data)
    with _conn() as conn:
        conn.execute('UPDATE hr_change_requests SET staged=? WHERE id=?', (path, row['id']))
    return row


def get(rid: int) -> Optional[dict]:
    with _conn() as conn:
        return _row(conn.execute('SELECT * FROM hr_change_requests WHERE id=?', (rid,)).fetchone())


def staged_path(rid: int) -> Optional[str]:
    with _conn() as conn:
        r = conn.execute('SELECT staged FROM hr_change_requests WHERE id=? AND status=?', (rid, PENDING)).fetchone()
    return r['staged'] if r and r['staged'] and os.path.isfile(r['staged']) else None


def pending(person_id: Optional[int] = None) -> List[dict]:
    sql, args = 'SELECT * FROM hr_change_requests WHERE status=?', [PENDING]
    if person_id is not None:
        sql += ' AND person_id=?'
        args.append(person_id)
    with _conn() as conn:
        return [_row(r) for r in conn.execute(sql + ' ORDER BY id', args).fetchall()]


def mine(person_id: int, limit: int = 10) -> List[dict]:
    """본인 화면 — 확인 중 + 최근 처리된 것(취소 제외)."""
    with _conn() as conn:
        return [_row(r) for r in conn.execute(
            'SELECT * FROM hr_change_requests WHERE person_id=? AND status != ? ORDER BY id DESC LIMIT ?',
            (person_id, CANCELLED, limit)).fetchall()]


def _drop_staged(rid: int) -> None:
    with _conn() as conn:
        r = conn.execute('SELECT staged FROM hr_change_requests WHERE id=?', (rid,)).fetchone()
        conn.execute('UPDATE hr_change_requests SET staged=NULL WHERE id=?', (rid,))
    if r and r['staged']:
        try:
            os.remove(r['staged'])
        except OSError:
            pass


def _close(rid: int, status: str, by: str, reason: str = '') -> Optional[dict]:
    with _conn() as conn:
        cur = conn.execute('UPDATE hr_change_requests SET status=?, handled_by=?, handled_at=?, reason=? '
                           'WHERE id=? AND status=?', (status, by, _now(), reason or None, rid, PENDING))
        if cur.rowcount != 1:
            return None
        return _row(conn.execute('SELECT * FROM hr_change_requests WHERE id=?', (rid,)).fetchone())


def cancel(rid: int, person_id: int, by: str) -> dict:
    row = get(rid)
    if not row or row['person_id'] != person_id:
        raise LookupError('없는 요청이에요.')
    done = _close(rid, CANCELLED, by)
    if not done:
        raise ValueError('이미 처리된 요청이에요.')
    _drop_staged(rid)
    return done


def reject(rid: int, admin: str, reason: str) -> dict:
    from dashboard.services import hr
    if not hr.is_hr_admin(admin):
        raise PermissionError('경영지원·대표만 처리할 수 있습니다.')
    if not (reason or '').strip():
        raise ValueError('반려 사유를 적어 주세요.')
    done = _close(rid, REJECTED, admin, reason.strip()[:300])
    if not done:
        raise ValueError('이미 처리된 요청이에요.')
    _drop_staged(rid)
    return done


def approve(rid: int, admin: str) -> dict:
    """반영 — 인사카드에 넣고 닫는다. 넣다가 실패하면 대기 그대로."""
    from dashboard.services import hr
    if not hr.is_hr_admin(admin):
        raise PermissionError('경영지원·대표만 처리할 수 있습니다.')
    row = get(rid)
    if not row or row['status'] != PENDING:
        raise ValueError('이미 처리된 요청이에요.')
    pid, pl = row['person_id'], row['payload']
    if row['kind'] == INFO:
        hr.update_person(pid, pl['new'], admin)
    elif row['kind'] == FAMILY:
        keep = {r['id'] for r in pl['rows'] if r.get('id')}
        with hr._conn() as conn:
            for f in conn.execute('SELECT id FROM hr_families WHERE person_id=?', (pid,)).fetchall():
                if f['id'] not in keep:
                    conn.execute('DELETE FROM hr_families WHERE id=? AND person_id=?', (f['id'], pid))
            for r in pl['rows']:
                if r.get('id'):
                    conn.execute('UPDATE hr_families SET name=?, relation=?, phone=? WHERE id=? AND person_id=?',
                                 (r['name'], r['relation'], r['phone'] or None, r['id'], pid))
                else:
                    conn.execute("INSERT INTO hr_families (person_id, name, relation, phone, source) VALUES (?,?,?,?,'self')",
                                 (pid, r['name'], r['relation'], r['phone'] or None))
        hr.log_access(pid, admin, 'edit', 'families')
    elif row['kind'] == DOC:
        path = staged_path(rid)
        if not path:
            raise LookupError('올린 파일을 찾지 못했어요. 반려하고 다시 올려 달라고 해 주세요.')
        with open(path, 'rb') as fp:
            data = fp.read()
        hr.save_document(pid, pl['doc_type'], pl['file_name'], data, row.get('requested_by') or admin, trusted=True)
    done = _close(rid, APPROVED, admin)
    if not done:
        raise ValueError('이미 처리된 요청이에요.')
    if row['kind'] == DOC:
        _drop_staged(rid)
    return done


def summary(row: dict) -> str:
    """한 줄 요약 — 슬랙 알림·홈 할 일."""
    pl = row['payload']
    if row['kind'] == INFO:
        return ', '.join(dict(INFO_FIELDS)[k] for k in pl.get('new', {}))
    if row['kind'] == FAMILY:
        return f"{len(pl.get('rows', []))}명으로"
    return f"{pl.get('doc_type', '')} · {pl.get('file_name', '')}"
