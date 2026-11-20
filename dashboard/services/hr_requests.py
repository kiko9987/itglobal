# -*- coding: utf-8 -*-
"""내 정보 직접 고치기 (2026-11-20) — 연락처·비상연락망·서류는 본인이 바로 고친다.

처음엔 '변경 요청 → 경영지원 반영'으로 만들었다가 같은 날 부장님: "본인 정보는 본인이 해야 맞다, 다만 전부는 아니고".
그래서 본인이 바로 고칠 수 있는 것만 정했다:
  - 연락처(전화·개인 메일·주소), 비상연락망(가족 이름·관계·연락처), 서류 올리기(등본·통장 사본 등)
  - 이름·주민번호·입사일·부서·직위 같은 인사 정보, 서류 지우기는 계속 경영지원만.
고치면 바로 인사카드에 들어가고, 경영지원에게 알림 DM, 인사카드 수정 기록(본인 수정)에 남는다.
이 표(hr_change_requests)는 무엇을 언제 바꿨는지 남기는 기록이다(status='applied').
"""
import json
import re
from contextlib import contextmanager
from datetime import datetime
from typing import Dict, List, Optional

APPLIED = 'applied'
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
    status TEXT NOT NULL DEFAULT 'applied',
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


def _record(person: dict, kind: str, payload: dict, by: str) -> dict:
    with _conn() as conn:
        rid = conn.execute('INSERT INTO hr_change_requests (person_id, name, kind, payload, status, requested_by, '
                           'created_at, handled_by, handled_at) VALUES (?,?,?,?,?,?,?,?,?)',
                           (person['id'], person['name'], kind, json.dumps(payload, ensure_ascii=False), APPLIED,
                            by, _now(), '본인', _now())).lastrowid
        return _row(conn.execute('SELECT * FROM hr_change_requests WHERE id=?', (rid,)).fetchone())


def update_info(person: dict, values: Dict, by: str) -> dict:
    """연락처 바로 고치기 — 바뀐 칸만."""
    from dashboard.services import hr
    new, old = {}, {}
    for k, _label in INFO_FIELDS:
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
        if (person.get(k) or '') != v:
            new[k], old[k] = v, person.get(k) or ''
    if not new:
        raise ValueError('바뀐 내용이 없어요.')
    sets = ', '.join(f'{k} = ?' for k in new)
    with hr._conn() as conn:
        conn.execute(f'UPDATE hr_people SET {sets}, updated_at = ?, updated_by = ? WHERE id = ?',
                     (*[v or None for v in new.values()], _now(), (by or '').lower(), person['id']))
    hr.log_access(person['id'], by, 'self_edit', ','.join(sorted(new)))
    return _record(person, INFO, {'new': new, 'old': old}, by)


def update_family(person: dict, rows: List[Dict], by: str, by_admin: bool = False) -> dict:
    """비상연락망 바로 고치기 — 원하는 목록 전체(기존 줄은 id 유지, 새 줄은 id 없음, 빠진 줄은 지움)."""
    from dashboard.services import hr
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
    pid = person['id']
    keep = {r['id'] for r in clean if r['id']}
    with hr._conn() as conn:
        for fid in cur:
            if fid not in keep:
                conn.execute('DELETE FROM hr_families WHERE id=? AND person_id=?', (fid, pid))
        for r in clean:
            if r['id']:
                conn.execute('UPDATE hr_families SET name=?, relation=?, phone=? WHERE id=? AND person_id=?',
                             (r['name'], r['relation'], r['phone'] or None, r['id'], pid))
            else:
                conn.execute("INSERT INTO hr_families (person_id, name, relation, phone, source) VALUES (?,?,?,?,'self')",
                             (pid, r['name'], r['relation'], r['phone'] or None))
    hr.log_access(pid, by, 'edit' if by_admin else 'self_edit', 'families')
    if by_admin:                       # 경영지원이 인사카드에서 고친 것(11-20) — 수정 기록만, 본인 수정 기록 표엔 안 남김
        return {'person_id': pid, 'name': person['name'], 'kind': FAMILY, 'kind_label': KIND_LABEL[FAMILY],
                'payload': {'rows': clean, 'old': old}}
    return _record(person, FAMILY, {'rows': clean, 'old': old}, by)


def upload_doc(person: dict, doc_type: str, file_name: str, data: bytes, by: str) -> dict:
    """서류 바로 올리기 — 인사카드 서류에 들어감(지우기는 경영지원만)."""
    from dashboard.services import hr
    if doc_type not in SELF_DOC_TYPES:
        raise ValueError('서류 종류를 골라 주세요.')
    if not (file_name or '').lower().endswith(hr.DOC_EXT):
        raise ValueError('올릴 수 없는 파일 형식이에요.')
    saved = hr.save_document(person['id'], doc_type, file_name, data, by, trusted=True)
    return _record(person, DOC, {'doc_type': doc_type, 'file_name': hr._safe_name(file_name), 'size': len(data),
                                 'doc_id': saved.get('id')}, by)


def history(person_id: int, limit: int = 10) -> List[dict]:
    with _conn() as conn:
        return [_row(r) for r in conn.execute(
            'SELECT * FROM hr_change_requests WHERE person_id=? AND status=? ORDER BY id DESC LIMIT ?',
            (person_id, APPLIED, limit)).fetchall()]


def summary(row: dict) -> str:
    """한 줄 요약 — 경영지원 알림."""
    pl = row['payload']
    if row['kind'] == INFO:
        return ', '.join(dict(INFO_FIELDS)[k] for k in pl.get('new', {}))
    if row['kind'] == FAMILY:
        return f"{len(pl.get('rows', []))}명"
    return f"{pl.get('doc_type', '')} · {pl.get('file_name', '')}"
