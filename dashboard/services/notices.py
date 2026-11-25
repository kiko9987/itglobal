"""그룹웨어 공지사항·개발 노트(2026-11-25 부장님 '그룹웨어인데 공지사항이 없네').

- 공지사항(kind='notice'): 카테고리마다 보는 사람을 정한다(전 직원 / 사무실 / 현장·미화 / 부서 고르기).
  기본 카테고리 = 전사 공지(전 직원) · 사무실 공지(사무실) · 현장 공지(현장·미화). 관리자가 더하고 고친다.
- 개발 노트(kind='devnote'): 그룹웨어·관리 사이트·슬랙 봇이 바뀐 내용. 사무실 직원이 본다(슬랙 「관리 시스템 업데이트 안내」 대신).
- 쓰기·고치기·지우기·카테고리 관리 = 그룹웨어 관리자(admin·super_admin)만. 관리자는 모든 글을 본다.
- 읽음 확인: 글을 열면 읽음으로 남고, 관리자는 대상 중 누가 읽었고 누가 안 읽었는지 본다.
- 슬랙 알림은 글마다 고른다(카테고리에 채널이 있어야). 기본 꺼짐.
- 현장·미화 = 인사카드 접근 방식 field(슬랙 1회용 링크로 들어오는 사람). 사무실 = 그 밖의 재직자.
- 첨부: instance/notice_files/<글 번호>/ (git 제외). 지운 글은 숨김(deleted_at)만 — 파일은 남긴다.
"""
from __future__ import annotations

import json
import os
import re
from contextlib import contextmanager
from datetime import datetime
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)
_LOG = '[공지]'

NOTICE, DEVNOTE = 'notice', 'devnote'
KINDS = (NOTICE, DEVNOTE)
KIND_LABEL = {NOTICE: '공지사항', DEVNOTE: '개발 노트'}
AUD_ALL, AUD_OFFICE, AUD_FIELD, AUD_DEPT = 'all', 'office', 'field', 'dept'
AUDIENCES = {AUD_ALL: '전 직원', AUD_OFFICE: '사무실 직원', AUD_FIELD: '현장·미화', AUD_DEPT: '부서 고르기'}
DEFAULT_CATEGORIES = (('전사 공지', AUD_ALL), ('사무실 공지', AUD_OFFICE), ('현장 공지', AUD_FIELD))

TITLE_MAX, BODY_MAX = 120, 20000
FILE_EXT = ('.pdf', '.jpg', '.jpeg', '.png', '.gif', '.webp', '.heic', '.doc', '.docx', '.hwp', '.hwpx',
            '.xls', '.xlsx', '.ppt', '.pptx', '.zip', '.txt')
FILE_MAX_BYTES = 15 * 1024 * 1024
FILES_PER_NOTICE = 10


class NoticeError(ValueError):
    """사용자에게 그대로 보여 줄 입력 오류."""


DDL = [
    '''CREATE TABLE IF NOT EXISTS notice_categories (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        audience TEXT NOT NULL DEFAULT 'all',
        depts TEXT NOT NULL DEFAULT '[]',
        slack_channel TEXT,
        sort INTEGER NOT NULL DEFAULT 0,
        deleted_at TEXT
    )''',
    '''CREATE TABLE IF NOT EXISTS notices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kind TEXT NOT NULL DEFAULT 'notice',
        category_id INTEGER,
        title TEXT NOT NULL,
        body TEXT NOT NULL DEFAULT '',
        pinned INTEGER NOT NULL DEFAULT 0,
        author_name TEXT,
        author_email TEXT,
        created_at TEXT NOT NULL,
        updated_at TEXT,
        updated_by TEXT,
        deleted_at TEXT,
        slack_channel TEXT,
        slack_ts TEXT,
        flex_id TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_notices_kind ON notices(kind, deleted_at)',
    '''CREATE TABLE IF NOT EXISTS notice_files (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        notice_id INTEGER NOT NULL,
        file_name TEXT NOT NULL,
        stored_name TEXT NOT NULL,
        size INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL,
        deleted_at TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_notice_files_notice ON notice_files(notice_id)',
    '''CREATE TABLE IF NOT EXISTS notice_reads (
        notice_id INTEGER NOT NULL,
        reader_key TEXT NOT NULL,
        reader_name TEXT,
        read_at TEXT NOT NULL,
        PRIMARY KEY (notice_id, reader_key)
    )''',
]
_table_ready = False


def _db():
    from dashboard.utils.user_database import get_user_database
    return get_user_database()


@contextmanager
def _conn():
    """사용자 DB 연결 — 처음 한 번 표·기본 카테고리를 만들고, 오류 없이 끝나면 저장한다.
    ⚠️ 연결 안에서 다시 _conn() 을 열지 않는다(DB 잠금)."""
    global _table_ready
    with _db()._get_connection() as conn:
        try:
            if not _table_ready:
                for q in DDL:
                    conn.execute(q)
                if not conn.execute('SELECT 1 FROM notice_categories LIMIT 1').fetchone():
                    for i, (name, aud) in enumerate(DEFAULT_CATEGORIES):
                        conn.execute('INSERT INTO notice_categories (name, audience, sort) VALUES (?, ?, ?)', (name, aud, i))
                conn.commit()
                _table_ready = True
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


# ── 보는 사람 ─────────────────────────────────────────────────────────
def viewer(person: Optional[dict], email: str = '', admin: bool = False, field: bool = False) -> dict:
    """화면에서 넘겨주는 보는 사람 — {key, name, dept, field, admin}. 인사카드가 있으면 그 번호로 읽음을 남긴다."""
    from dashboard.services import hr
    p = person or {}
    is_field = bool(field) or p.get('access') == hr.ACCESS_FIELD
    key = f"p{p['id']}" if p.get('id') else (f'e{email.lower()}' if email else '')
    return {'key': key, 'name': p.get('name') or (email.split('@')[0] if email else ''),
            'dept': p.get('department') or '', 'field': is_field, 'admin': bool(admin) and not field}


def _dept_match(dept: str, depts: List[str]) -> bool:
    return any(d and (d == dept or d in dept) for d in depts)


def can_see(row: dict, cat: Optional[dict], v: dict) -> bool:
    if v.get('admin'):
        return True
    if row['kind'] == DEVNOTE:
        return not v.get('field')
    if not cat:
        return False
    aud = cat['audience']
    if aud == AUD_ALL:
        return True
    if aud == AUD_OFFICE:
        return not v.get('field')
    if aud == AUD_FIELD:
        return bool(v.get('field'))
    return _dept_match(v.get('dept') or '', cat.get('depts') or [])


# ── 카테고리 ──────────────────────────────────────────────────────────
def _cat_row(r) -> dict:
    c = dict(r)
    try:
        c['depts'] = json.loads(c.get('depts') or '[]')
    except ValueError:
        c['depts'] = []
    c['audience_label'] = AUDIENCES.get(c['audience'], c['audience'])
    if c['audience'] == AUD_DEPT and c['depts']:
        c['audience_label'] = ', '.join(c['depts'])
    return c


def categories(include_deleted: bool = False) -> List[dict]:
    with _conn() as conn:
        rows = conn.execute('SELECT * FROM notice_categories ' + ('' if include_deleted else 'WHERE deleted_at IS NULL ')
                            + 'ORDER BY sort, id').fetchall()
    return [_cat_row(r) for r in rows]


def _cat_map() -> Dict[int, dict]:
    return {c['id']: c for c in categories(include_deleted=True)}


def visible_categories(v: dict) -> List[dict]:
    """보는 사람이 볼 수 있는 카테고리(거르기 칩)."""
    return [c for c in categories() if can_see({'kind': NOTICE}, c, v)]


def _clean_channel(ch: str) -> str:
    ch = (ch or '').strip().lstrip('#')
    if ch and not re.fullmatch(r'[A-Z0-9]{8,12}', ch):
        raise NoticeError('슬랙 채널은 채널 ID(C로 시작하는 영문·숫자)로 적어 주세요. 채널 이름 옆 ⋮ → 채널 세부정보 맨 아래에 있어요.')
    return ch


def save_category(data: dict, admin: bool, cid: Optional[int] = None) -> dict:
    if not admin:
        raise PermissionError('카테고리는 관리자만 고칠 수 있어요.')
    name = (data.get('name') or '').strip()[:20]
    aud = data.get('audience') or AUD_ALL
    if not name:
        raise NoticeError('카테고리 이름을 적어 주세요.')
    if aud not in AUDIENCES:
        raise NoticeError('보는 사람을 골라 주세요.')
    depts = [str(d).strip()[:30] for d in (data.get('depts') or []) if str(d).strip()] if aud == AUD_DEPT else []
    if aud == AUD_DEPT and not depts:
        raise NoticeError('보여 줄 부서를 하나 이상 골라 주세요.')
    ch = _clean_channel(data.get('slack_channel') or '')
    with _conn() as conn:
        dup = conn.execute('SELECT id FROM notice_categories WHERE name=? AND deleted_at IS NULL AND id<>?',
                           (name, cid or 0)).fetchone()
        if dup:
            raise NoticeError(f'"{name}" 카테고리가 이미 있어요.')
        if cid:
            cur = conn.execute('UPDATE notice_categories SET name=?, audience=?, depts=?, slack_channel=? '
                               'WHERE id=? AND deleted_at IS NULL',
                               (name, aud, json.dumps(depts, ensure_ascii=False), ch or None, cid))
            if not cur.rowcount:
                raise LookupError('없는 카테고리예요. 새로 고쳐 주세요.')
        else:
            n = conn.execute('SELECT COALESCE(MAX(sort), 0) + 1 FROM notice_categories').fetchone()[0]
            cid = conn.execute('INSERT INTO notice_categories (name, audience, depts, slack_channel, sort) VALUES (?, ?, ?, ?, ?)',
                               (name, aud, json.dumps(depts, ensure_ascii=False), ch or None, n)).lastrowid
    return next(c for c in categories() if c['id'] == cid)


def delete_category(cid: int, admin: bool) -> None:
    if not admin:
        raise PermissionError('카테고리는 관리자만 고칠 수 있어요.')
    with _conn() as conn:
        n = conn.execute('SELECT COUNT(*) FROM notices WHERE category_id=? AND deleted_at IS NULL', (cid,)).fetchone()[0]
        if n:
            raise NoticeError(f'이 카테고리에 공지가 {n}건 있어요. 공지를 다른 카테고리로 옮긴 뒤 지워 주세요.')
        if conn.execute('SELECT COUNT(*) FROM notice_categories WHERE deleted_at IS NULL').fetchone()[0] <= 1:
            raise NoticeError('카테고리는 하나 이상 있어야 해요.')
        conn.execute('UPDATE notice_categories SET deleted_at=? WHERE id=?', (_now(), cid))


# ── 글 ────────────────────────────────────────────────────────────────
def _files(conn, nids: List[int]) -> Dict[int, List[dict]]:
    out: Dict[int, List[dict]] = {}
    if not nids:
        return out
    q = ','.join('?' * len(nids))
    for r in conn.execute(f'SELECT id, notice_id, file_name, size FROM notice_files WHERE notice_id IN ({q}) '
                          f'AND deleted_at IS NULL ORDER BY id', nids):
        out.setdefault(r['notice_id'], []).append({'id': r['id'], 'name': r['file_name'], 'size': r['size']})
    return out


def _public(row: dict, cat: Optional[dict], files: List[dict], read: bool, with_body: bool) -> dict:
    d = {'id': row['id'], 'kind': row['kind'], 'title': row['title'], 'pinned': bool(row['pinned']),
         'author': row['author_name'] or '', 'created_at': row['created_at'], 'updated_at': row['updated_at'],
         'category_id': row['category_id'], 'category': (cat or {}).get('name') or '',
         'audience': (cat or {}).get('audience_label') or ('사무실 직원' if row['kind'] == DEVNOTE else ''),
         'files': files, 'read': read, 'slack_sent': bool(row['slack_ts'])}
    if with_body:
        d['body'] = row['body']
    else:
        d['preview'] = re.sub(r'\s+', ' ', row['body'] or '')[:90]
    return d


def list_notices(kind: str, v: dict, category_id: Optional[int] = None, limit: int = 0, q: str = '') -> List[dict]:
    """보는 사람이 볼 수 있는 글 — 고정 먼저, 그다음 최신순."""
    if kind not in KINDS:
        raise NoticeError('없는 게시판이에요.')
    cats = _cat_map()
    with _conn() as conn:
        rows = [dict(r) for r in conn.execute(
            'SELECT * FROM notices WHERE kind=? AND deleted_at IS NULL ORDER BY pinned DESC, created_at DESC, id DESC',
            (kind,))]
        rows = [r for r in rows if can_see(r, cats.get(r['category_id']), v)
                and (not category_id or r['category_id'] == category_id)]
        if q:
            ql = q.strip().lower()
            rows = [r for r in rows if ql in (r['title'] or '').lower() or ql in (r['body'] or '').lower()]
        if limit:
            rows = rows[:limit]
        ids = [r['id'] for r in rows]
        read = set()
        if ids and v.get('key'):
            qm = ','.join('?' * len(ids))
            read = {r[0] for r in conn.execute(f'SELECT notice_id FROM notice_reads WHERE reader_key=? AND notice_id IN ({qm})',
                                               (v['key'], *ids))}
        files = _files(conn, ids)
    return [_public(r, cats.get(r['category_id']), files.get(r['id'], []), r['id'] in read, False) for r in rows]


def unread_count(kind: str, v: dict) -> int:
    return sum(1 for n in list_notices(kind, v) if not n['read'])


def _get_row(nid: int) -> Optional[dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM notices WHERE id=? AND deleted_at IS NULL', (nid,)).fetchone()
    return dict(r) if r else None


def get(nid: int, v: dict, mark_read: bool = True) -> dict:
    row = _get_row(nid)
    cats = _cat_map()
    if not row:
        raise LookupError('없는 글이에요. 지워졌을 수 있어요.')
    cat = cats.get(row['category_id'])
    if not can_see(row, cat, v):
        raise PermissionError('볼 수 없는 글이에요.')
    with _conn() as conn:
        if mark_read and v.get('key'):
            conn.execute('INSERT OR IGNORE INTO notice_reads (notice_id, reader_key, reader_name, read_at) VALUES (?, ?, ?, ?)',
                         (nid, v['key'], v.get('name') or '', _now()))
        files = _files(conn, [nid]).get(nid, [])
    out = _public(row, cat, files, True, True)
    if v.get('admin'):
        out['reads'] = read_status(nid)
    return out


def _check(kind: str, data: dict, cats: Dict[int, dict]) -> dict:
    title = (data.get('title') or '').strip()
    body = (data.get('body') or '').replace('\r\n', '\n').strip()
    if not title:
        raise NoticeError('제목을 적어 주세요.')
    if len(title) > TITLE_MAX:
        raise NoticeError(f'제목은 {TITLE_MAX}자까지 적을 수 있어요.')
    if len(body) > BODY_MAX:
        raise NoticeError('본문이 너무 길어요. 나눠서 올리거나 파일로 붙여 주세요.')
    cid = None
    if kind == NOTICE:
        try:
            cid = int(data.get('category_id') or 0)
        except (TypeError, ValueError):
            cid = 0
        if not cid or cid not in cats or cats[cid].get('deleted_at'):
            raise NoticeError('카테고리를 골라 주세요.')
    return {'title': title, 'body': body, 'category_id': cid, 'pinned': 1 if data.get('pinned') else 0}


def save(kind: str, data: dict, author: dict, admin: bool, nid: Optional[int] = None) -> dict:
    """새 글(nid 없음)·고치기. author = {name, email}. 슬랙 알림은 따로(notify)."""
    if not admin:
        raise PermissionError('공지는 관리자만 쓸 수 있어요.')
    if kind not in KINDS:
        raise NoticeError('없는 게시판이에요.')
    c = _check(kind, data, _cat_map())
    now = _now()
    with _conn() as conn:
        if nid:
            cur = conn.execute('UPDATE notices SET title=?, body=?, category_id=?, pinned=?, updated_at=?, updated_by=? '
                               'WHERE id=? AND kind=? AND deleted_at IS NULL',
                               (c['title'], c['body'], c['category_id'], c['pinned'], now, author.get('name') or '', nid, kind))
            if not cur.rowcount:
                raise LookupError('없는 글이에요. 새로 고쳐 주세요.')
        else:
            nid = conn.execute('INSERT INTO notices (kind, category_id, title, body, pinned, author_name, author_email, created_at) '
                               'VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
                               (kind, c['category_id'], c['title'], c['body'], c['pinned'], author.get('name') or '',
                                (author.get('email') or '').lower(), now)).lastrowid
            # 쓴 사람은 읽은 것으로
            if author.get('key'):
                conn.execute('INSERT OR IGNORE INTO notice_reads (notice_id, reader_key, reader_name, read_at) VALUES (?, ?, ?, ?)',
                             (nid, author['key'], author.get('name') or '', now))
    logger.info(f"{_LOG} {KIND_LABEL[kind]} {'고침' if data.get('_edit') else '저장'} #{nid} {c['title'][:30]} by {author.get('name')}")
    return {'id': nid}


def delete(nid: int, admin: bool, by: str = '') -> None:
    if not admin:
        raise PermissionError('공지는 관리자만 지울 수 있어요.')
    with _conn() as conn:
        cur = conn.execute('UPDATE notices SET deleted_at=?, updated_by=? WHERE id=? AND deleted_at IS NULL', (_now(), by, nid))
    if not cur.rowcount:
        raise LookupError('없는 글이에요. 새로 고쳐 주세요.')
    logger.info(f'{_LOG} #{nid} 지움 by {by}')


# ── 첨부 ──────────────────────────────────────────────────────────────
def files_dir() -> str:
    d = os.getenv('GW_NOTICE_FILES_DIR', '').strip()
    if not d:
        from pathlib import Path
        d = str(Path(__file__).resolve().parents[2] / 'instance' / 'notice_files')
    return d


def _safe_name(name: str) -> str:
    name = os.path.basename((name or '').replace('\\', '/')).strip() or '파일'
    name = re.sub(r'[\x00-\x1f<>:"/\\|?*]', '_', name)
    return name[-120:]


def add_file(nid: int, file_name: str, data: bytes, admin: bool) -> dict:
    if not admin:
        raise PermissionError('첨부는 관리자만 올릴 수 있어요.')
    if not _get_row(nid):
        raise LookupError('없는 글이에요.')
    if not data:
        raise NoticeError('빈 파일이에요.')
    if len(data) > FILE_MAX_BYTES:
        raise NoticeError('파일이 너무 커요(15MB까지 올릴 수 있어요).')
    if not (file_name or '').lower().endswith(FILE_EXT):
        raise NoticeError('올릴 수 없는 파일 형식이에요(문서·사진·압축 파일만).')
    folder = os.path.join(files_dir(), str(nid))
    os.makedirs(folder, exist_ok=True)
    safe = _safe_name(file_name)
    with _conn() as conn:
        if conn.execute('SELECT COUNT(*) FROM notice_files WHERE notice_id=? AND deleted_at IS NULL', (nid,)).fetchone()[0] \
                >= FILES_PER_NOTICE:
            raise NoticeError(f'첨부는 한 글에 {FILES_PER_NOTICE}개까지예요.')
        fid = conn.execute('INSERT INTO notice_files (notice_id, file_name, stored_name, size, created_at) VALUES (?, ?, ?, ?, ?)',
                           (nid, safe, 'pending', len(data), _now())).lastrowid
        stored = f'{fid}_{safe}'
        with open(os.path.join(folder, stored), 'wb') as fp:
            fp.write(data)
        conn.execute('UPDATE notice_files SET stored_name=? WHERE id=?', (stored, fid))
    return {'id': fid, 'name': safe, 'size': len(data)}


def remove_file(fid: int, admin: bool) -> None:
    if not admin:
        raise PermissionError('첨부는 관리자만 지울 수 있어요.')
    with _conn() as conn:
        cur = conn.execute('UPDATE notice_files SET deleted_at=? WHERE id=? AND deleted_at IS NULL', (_now(), fid))
    if not cur.rowcount:
        raise LookupError('없는 파일이에요.')


def file_for(fid: int, v: dict) -> dict:
    """내려받을 파일 {path, name} — 그 글을 볼 수 있는 사람만."""
    with _conn() as conn:
        f = conn.execute('SELECT * FROM notice_files WHERE id=? AND deleted_at IS NULL', (fid,)).fetchone()
    if not f:
        raise LookupError('없는 파일이에요.')
    row = _get_row(f['notice_id'])
    if not row or not can_see(row, _cat_map().get(row['category_id']), v):
        raise PermissionError('볼 수 없는 파일이에요.')
    return {'path': os.path.join(files_dir(), str(f['notice_id']), f['stored_name']), 'name': f['file_name']}


# ── 읽음 확인 ─────────────────────────────────────────────────────────
def audience_people(row: dict, cat: Optional[dict]) -> List[dict]:
    """이 글을 볼 대상(재직자) [{key, name, dept}] — 관리자 포함 여부와 상관없이 대상 규칙대로."""
    from dashboard.services import hr
    out = []
    for p in hr.list_people():
        if p.get('status') == hr.STATUS_RESIGNED:
            continue
        v = viewer(p)
        if can_see(row, cat, dict(v, admin=False)):
            out.append({'key': v['key'], 'name': p.get('name') or '', 'dept': p.get('department') or ''})
    return out


def read_status(nid: int) -> dict:
    row = _get_row(nid)
    if not row:
        raise LookupError('없는 글이에요.')
    people = audience_people(row, _cat_map().get(row['category_id']))
    with _conn() as conn:
        reads = {r['reader_key']: r['read_at'] for r in conn.execute(
            'SELECT reader_key, read_at FROM notice_reads WHERE notice_id=?', (nid,))}
    done = [dict(p, read_at=reads[p['key']]) for p in people if p['key'] in reads]
    left = [p for p in people if p['key'] not in reads]
    done.sort(key=lambda p: p['read_at'])
    left.sort(key=lambda p: (p['dept'], p['name']))
    return {'total': len(people), 'read': done, 'unread': left}


# ── 슬랙 알림 ─────────────────────────────────────────────────────────
def _slack_client():
    from slack_sdk import WebClient
    tok = os.getenv('NOTICE_SLACK_BOT_TOKEN') or os.getenv('SLACK_LEAVE_BOT_TOKEN') or ''
    return WebClient(token=tok) if tok else None


def channel_for(row: dict) -> str:
    if row['kind'] == DEVNOTE:
        return (os.getenv('DEVNOTE_SLACK_CHANNEL') or '').strip()
    return ((_cat_map().get(row['category_id']) or {}).get('slack_channel') or '').strip()


def notify(nid: int, admin: bool, client=None) -> dict:
    """슬랙 채널에 제목·링크를 올린다. 채널이 없거나 실패하면 NoticeError(글은 이미 저장됨)."""
    if not admin:
        raise PermissionError('관리자만 알릴 수 있어요.')
    row = _get_row(nid)
    if not row:
        raise LookupError('없는 글이에요.')
    ch = channel_for(row)
    if not ch:
        raise NoticeError('이 카테고리에 슬랙 채널이 정해져 있지 않아 알리지 못했어요. 카테고리 관리에서 채널을 넣어 주세요.'
                          if row['kind'] == NOTICE else '개발 노트 슬랙 채널이 정해져 있지 않아 알리지 못했어요.')
    client = client or _slack_client()
    if not client:
        raise NoticeError('슬랙 봇 설정이 없어 알리지 못했어요.')
    from dashboard.services import gw_login
    cat = (_cat_map().get(row['category_id']) or {}).get('name') or KIND_LABEL[row['kind']]
    path = 'devnotes' if row['kind'] == DEVNOTE else 'notices'
    url = f'{gw_login.public_url()}/{path}/{nid}'
    text = f"📢 [{cat}] {row['title']}\n<{url}|그룹웨어에서 보기>"
    try:
        r = client.chat_postMessage(channel=ch, text=text, unfurl_links=False)
    except Exception as e:
        logger.error(f'{_LOG} #{nid} 슬랙 알림 실패({ch}): {e}')
        raise NoticeError('슬랙으로 알리지 못했어요. 봇이 그 채널에 들어가 있는지 확인해 주세요. 글은 저장됐어요.')
    with _conn() as conn:
        conn.execute('UPDATE notices SET slack_channel=?, slack_ts=? WHERE id=?', (ch, r.get('ts'), nid))
    return {'channel': ch, 'ts': r.get('ts')}
