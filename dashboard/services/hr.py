"""인사 장부 — 구성원·인사카드·발령 이력 (그룹웨어 1단계, 2026-11-10).

플렉스 해지 후 인사 정보를 볼 곳. 사용자 DB(users.db)에 별도 표로 둔다.

열람 범위(KiKO 11-10):
  - 본인: 자기 인사카드만
  - 결재자(JW·YG): 전 직원 휴가 현황(휴가 화면) — 인사카드는 안 봄
  - 경영지원·대표(HR 관리자): 인사카드 전체 + 수정
주민등록번호: 저장하되 암호화(개인정보보호법 제24조의2). 화면은 뒷자리 가림, HR 관리자만 [전체 보기],
  보기 기록은 남기지 않음(11-10 KiKO — 내부 그룹웨어). 수정·서류 올리기·지우기만 hr_access_log. 암호화 키는 .env GW_HR_KEY(Fernet). 키가 없으면 저장 안 함.
연봉 등 급여 정보는 옮기지 않는다(요청 범위 밖).
"""
from __future__ import annotations

import json
import os
import re
from contextlib import contextmanager
from datetime import date, datetime
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

STATUS_ACTIVE, STATUS_LEAVE, STATUS_RESIGNED = '재직', '휴직', '퇴사'
STATUSES = (STATUS_ACTIVE, STATUS_LEAVE, STATUS_RESIGNED)
ACCESS_STAFF, ACCESS_FIELD = 'staff', 'field'   # field = 현장·미화(개인 메일 로그인, 제한 권한)

# 화면·수정에 쓰는 칸(주민번호 제외). (키, 이름)
FIELDS = [
    ('name', '이름'), ('status', '상태'),
    ('hire_date', '입사일'), ('group_hire_date', '그룹 입사일'),
    ('resign_date', '퇴사일'), ('last_work_date', '마지막 근무일'), ('resign_reason', '퇴사 사유'),
    ('department', '부서'), ('job_role', '직무'), ('job_rank', '직급'), ('job_title', '직책'),
    ('contract_type', '계약 형태'),
    ('work_email', '회사 메일'), ('personal_email', '개인 메일'), ('phone', '전화'),
    ('gender', '성별'), ('birth_date', '생년월일'), ('address', '주소'),
    ('login_email', '그룹웨어 로그인 메일'), ('access', '로그인 방식'),
    ('slack_id', '슬랙 ID'), ('memo', '메모'),
]
EDITABLE = [k for k, _ in FIELDS]
_DATE_FIELDS = ('hire_date', 'group_hire_date', 'resign_date', 'last_work_date', 'birth_date')

_DDL = [
    '''CREATE TABLE IF NOT EXISTS hr_people (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        status TEXT NOT NULL DEFAULT '재직',
        hire_date TEXT, group_hire_date TEXT, resign_date TEXT, last_work_date TEXT, resign_reason TEXT,
        department TEXT, job_role TEXT, job_rank TEXT, job_title TEXT, contract_type TEXT,
        work_email TEXT, personal_email TEXT, phone TEXT, gender TEXT, birth_date TEXT, address TEXT,
        ssn_enc TEXT,
        login_email TEXT, access TEXT NOT NULL DEFAULT 'staff',
        slack_id TEXT, flex_id TEXT UNIQUE, memo TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_by TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_hr_people_login ON hr_people(login_email)',
    'CREATE INDEX IF NOT EXISTS idx_hr_people_work ON hr_people(work_email)',
    '''CREATE TABLE IF NOT EXISTS hr_appointments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        person_id INTEGER NOT NULL,
        apply_date TEXT,
        department TEXT, job_role TEXT, job_rank TEXT, job_title TEXT,
        memo TEXT, source TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )''',
    'CREATE INDEX IF NOT EXISTS idx_hr_appt_person ON hr_appointments(person_id)',
    '''CREATE TABLE IF NOT EXISTS hr_access_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        person_id INTEGER,
        viewer_email TEXT NOT NULL,
        action TEXT NOT NULL,
        detail TEXT,
        at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
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
            _migrate(conn)
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _row(r) -> dict:
    return {k: r[k] for k in r.keys()} if r is not None else None


# ── 역할 ──────────────────────────────────────────────────────────────
# 기본값은 코드에, 바꾼 값은 gw_settings(그룹웨어 관리 화면)에. 이메일 소문자.
# 경영지원(KiKO·SB) + 대표(박용구·박정우, KiKO 11-10 지정). 플렉스상 대표이사 김영숙은 넣지 않음.
_DEFAULT_HR_ADMINS = ('kiko@itg-aircon.com', 'sb@itg-aircon.com', 'yg@itg-aircon.com', 'jw@itg-aircon.com')
# 휴가 현황만 보는 사람(인사카드 X). 지금 결재자 JW·YG 는 대표라 위에 이미 들어 있음 — 결재자가 바뀌면 여기.
_DEFAULT_LEAVE_VIEWERS = ('jw@itg-aircon.com', 'yg@itg-aircon.com')


def _setting_list(key: str, default) -> List[str]:
    try:
        from dashboard.services import groupware as gw
        raw = gw.get_setting(key)
    except Exception as e:
        logger.warning(f'[HR] 역할 설정 읽기 실패({key}) → 기본값: {e}')
        raw = None
    if raw is None:
        return [e.lower() for e in default]
    try:
        return [str(e).strip().lower() for e in json.loads(raw) if str(e).strip()]
    except Exception:
        return [e.lower() for e in default]


def hr_admins() -> List[str]:
    return _setting_list('hr_admins', _DEFAULT_HR_ADMINS)


def leave_viewers() -> List[str]:
    return _setting_list('leave_viewers', _DEFAULT_LEAVE_VIEWERS)


def is_hr_admin(email: str) -> bool:
    return bool(email) and email.strip().lower() in hr_admins()


def can_view_all_leave(email: str) -> bool:
    return is_hr_admin(email) or (bool(email) and email.strip().lower() in leave_viewers())


# ── 주민등록번호 암호화 ────────────────────────────────────────────────
class SsnKeyMissing(RuntimeError):
    pass


def _fernet():
    key = os.getenv('GW_HR_KEY', '').strip()
    if not key:
        raise SsnKeyMissing('GW_HR_KEY 가 없어 주민등록번호를 저장·조회할 수 없습니다.')
    from cryptography.fernet import Fernet
    return Fernet(key.encode())


def normalize_ssn(raw: str) -> str:
    """숫자 13자리 → 'YYMMDD-NNNNNNN'. 형식이 아니면 ValueError."""
    digits = re.sub(r'\D', '', raw or '')
    if len(digits) != 13:
        raise ValueError('주민등록번호는 숫자 13자리여야 합니다.')
    return f'{digits[:6]}-{digits[6:]}'


def encrypt_ssn(raw: str) -> str:
    return _fernet().encrypt(normalize_ssn(raw).encode()).decode()


def decrypt_ssn(token: str) -> str:
    return _fernet().decrypt(token.encode()).decode()


def mask_ssn(full: str) -> str:
    """'900101-1234567' → '900101-1******'."""
    if not full or len(full) < 8:
        return ''
    return full[:8] + '*' * 6


# ── 조회 ──────────────────────────────────────────────────────────────
_LIST_COLS = 'id, name, status, department, job_role, job_rank, job_title, hire_date, resign_date, access, login_email, work_email'


def list_people(include_resigned: bool = False) -> List[dict]:
    sql = f'SELECT {_LIST_COLS} FROM hr_people'
    if not include_resigned:
        sql += " WHERE status != '퇴사'"
    sql += " ORDER BY CASE status WHEN '재직' THEN 0 WHEN '휴직' THEN 1 ELSE 2 END, hire_date, name"
    with _conn() as conn:
        return [_row(r) for r in conn.execute(sql).fetchall()]


def same_name_people(name: str) -> List[dict]:
    """이름이 같은 인사카드(퇴사 포함) — id·생년월일·근무 기간만. 플렉스 옛 문서를 사람·근무 기간으로 묶는 데 씀
    (2026-11-13, 11-14 퇴사일 추가)."""
    with _conn() as conn:
        return [dict(r) for r in conn.execute('SELECT id, name, birth_date, hire_date, resign_date, last_work_date '
                                              'FROM hr_people WHERE name = ?',
                                              ((name or '').strip(),)).fetchall()]


def get_person(person_id: int) -> Optional[dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM hr_people WHERE id = ?', (person_id,)).fetchone()
        return _row(r)


def find_by_email(email: str) -> Optional[dict]:
    """로그인 메일 → 인사카드(로그인 메일 우선, 없으면 회사 메일). 퇴사자는 제외."""
    e = (email or '').strip().lower()
    if not e:
        return None
    with _conn() as conn:
        for col in ('login_email', 'work_email', 'personal_email'):
            r = conn.execute(f"SELECT * FROM hr_people WHERE lower({col}) = ? AND status != '퇴사' "
                             'ORDER BY id LIMIT 1', (e,)).fetchone()
            if r is not None:
                return _row(r)
    return None


def find_by_slack(slack_id: str) -> Optional[dict]:
    """슬랙 ID → 재직 중인 인사카드(퇴사자 제외)."""
    if not slack_id:
        return None
    with _conn() as conn:
        r = conn.execute("SELECT * FROM hr_people WHERE slack_id = ? AND status != '퇴사' ORDER BY id LIMIT 1",
                         (slack_id,)).fetchone()
    return _row(r)


def identity_email(p: dict) -> str:
    """그룹웨어 세션에 쓸 이메일 — 로그인 메일 > 개인 메일 > 자리표시(현장 직원 중 메일 없는 사람)."""
    return (p.get('login_email') or p.get('personal_email') or f'field-{p["id"]}@itg-aircon.invalid').lower()


def appointments(person_id: int) -> List[dict]:
    with _conn() as conn:
        rows = conn.execute('SELECT * FROM hr_appointments WHERE person_id = ? '
                            'ORDER BY apply_date DESC, id DESC', (person_id,)).fetchall()
    return [_row(r) for r in rows]


def card(person_id: int, viewer_email: str) -> Optional[dict]:
    """화면에 보낼 인사카드. 주민번호는 가린 값만, 원본 암호문은 절대 내보내지 않는다."""
    p = get_person(person_id)
    if not p:
        return None
    enc = p.pop('ssn_enc', None)
    p['ssn_masked'] = ''
    p['has_ssn'] = bool(enc)
    if enc:
        try:
            p['ssn_masked'] = mask_ssn(decrypt_ssn(enc))
        except SsnKeyMissing:
            p['ssn_masked'] = '(키 없음)'
        except Exception:
            p['ssn_masked'] = '(읽기 실패)'
    p['appointments'] = appointments(person_id)
    p['can_reveal_ssn'] = is_hr_admin(viewer_email)
    p['can_edit'] = is_hr_admin(viewer_email)
    return p


def reveal_ssn(person_id: int, viewer_email: str) -> str:
    """HR 관리자만. 열람 기록은 남기지 않는다(내부 그룹웨어 — 11-10 KiKO 결정)."""
    if not is_hr_admin(viewer_email):
        raise PermissionError('주민등록번호 전체 보기는 경영지원·대표만 할 수 있습니다.')
    p = get_person(person_id)
    if not p or not p.get('ssn_enc'):
        raise LookupError('저장된 주민등록번호가 없습니다.')
    return decrypt_ssn(p['ssn_enc'])


def log_access(person_id: Optional[int], viewer_email: str, action: str, detail: str = '') -> None:
    with _conn() as conn:
        conn.execute('INSERT INTO hr_access_log (person_id, viewer_email, action, detail, at) VALUES (?, ?, ?, ?, ?)',
                     (person_id, (viewer_email or '').lower(), action, detail, _now()))


# 보기만 한 기록(11-10 이전에 쌓인 것)은 화면 '수정 기록'에서 뺀다 — 이제는 남기지도 않음
VIEW_ACTIONS = ('ssn_view', 'family_ssn_view', 'doc_download')


def access_log(person_id: int, limit: int = 30) -> List[dict]:
    """수정 기록(수정·추가·서류 업로드/삭제). 보기 기록은 제외."""
    ph = ','.join('?' * len(VIEW_ACTIONS))
    with _conn() as conn:
        rows = conn.execute(f'SELECT viewer_email, action, detail, at FROM hr_access_log WHERE person_id = ? '
                            f'AND action NOT IN ({ph}) ORDER BY id DESC LIMIT ?',
                            (person_id, *VIEW_ACTIONS, limit)).fetchall()
    return [_row(r) for r in rows]


# ── 수정 ──────────────────────────────────────────────────────────────
def _clean(values: Dict) -> Dict:
    out = {}
    for k in EDITABLE:
        if k not in values:
            continue
        v = values[k]
        v = '' if v is None else str(v).strip()
        if k in _DATE_FIELDS and v:
            try:
                v = date.fromisoformat(v[:10]).isoformat()
            except ValueError:
                raise ValueError(f'{dict(FIELDS)[k]} 날짜 형식이 아닙니다: {v}')
        if k == 'status' and v not in STATUSES:
            raise ValueError(f'상태는 {"·".join(STATUSES)} 중 하나여야 합니다.')
        if k == 'access' and v not in (ACCESS_STAFF, ACCESS_FIELD):
            raise ValueError('로그인 방식 값이 올바르지 않습니다.')
        if k in ('work_email', 'personal_email', 'login_email'):
            v = v.lower()
            if v and '@' not in v:
                raise ValueError(f'{dict(FIELDS)[k]} 형식이 아닙니다.')
        if k == 'name' and not v:
            raise ValueError('이름은 비울 수 없습니다.')
        if k in ('name', 'status', 'access'):
            out[k] = v
        else:
            out[k] = v or None      # 빈 칸은 NULL
    return out


def update_person(person_id: int, values: Dict, by: str) -> dict:
    """HR 관리자 수정. 바뀐 칸만 쓰고, 무엇을 바꿨는지 접근 기록에 남긴다(값은 남기지 않음)."""
    if not is_hr_admin(by):
        raise PermissionError('인사카드 수정은 경영지원·대표만 할 수 있습니다.')
    cur = get_person(person_id)
    if not cur:
        raise LookupError('없는 구성원입니다.')
    clean = _clean(values)
    changed = {k: v for k, v in clean.items() if (cur.get(k) or None) != (v or None)}
    ssn_raw = (values.get('ssn') or '').strip()
    if ssn_raw:
        changed['ssn_enc'] = encrypt_ssn(ssn_raw)
    if not changed:
        return card(person_id, by)
    sets = ', '.join(f'{k} = ?' for k in changed)
    with _conn() as conn:
        conn.execute(f'UPDATE hr_people SET {sets}, updated_at = ?, updated_by = ? WHERE id = ?',
                     (*changed.values(), _now(), by.lower(), person_id))
    log_access(person_id, by, 'edit', ','.join(sorted(k for k in changed)))
    return card(person_id, by)


def create_person(values: Dict, by: str) -> dict:
    if not is_hr_admin(by):
        raise PermissionError('구성원 추가는 경영지원·대표만 할 수 있습니다.')
    clean = _clean({'status': STATUS_ACTIVE, 'access': ACCESS_STAFF, **values})
    if not clean.get('name'):
        raise ValueError('이름은 비울 수 없습니다.')
    ssn_raw = (values.get('ssn') or '').strip()
    if ssn_raw:
        clean['ssn_enc'] = encrypt_ssn(ssn_raw)
    cols = ', '.join(clean)
    with _conn() as conn:
        cur = conn.execute(f'INSERT INTO hr_people ({cols}, updated_by) VALUES ({", ".join("?" * len(clean))}, ?)',
                           (*clean.values(), by.lower()))
        pid = cur.lastrowid
    log_access(pid, by, 'create')
    return card(pid, by)


# ── 이관(플렉스) ───────────────────────────────────────────────────────
_FLEX_FOLLOW = ('status', 'resign_date', 'last_work_date', 'resign_reason')


def upsert_from_flex(rec: Dict, appts: List[Dict], by: str = 'flex-import') -> str:
    """flex_id 기준으로 넣거나 고친다. 사람이 이미 고친 칸은 덮어쓰지 않는다(updated_by 가 import 가 아니면 비어 있는 칸만).
    반환: 'insert' | 'update' | 'skip'."""
    fid = rec['flex_id']
    clean = _clean({k: v for k, v in rec.items() if k in EDITABLE})
    ssn_raw = rec.get('ssn') or ''
    with _conn() as conn:
        r = conn.execute('SELECT * FROM hr_people WHERE flex_id = ?', (fid,)).fetchone()
        if r is None:
            data = dict(clean, flex_id=fid)
            if ssn_raw:
                try:
                    data['ssn_enc'] = encrypt_ssn(ssn_raw)
                except ValueError:
                    logger.warning(f'[HR] 주민번호 형식 아님 — 건너뜀(flex_id={fid})')
            cols = ', '.join(data)
            cur = conn.execute(f'INSERT INTO hr_people ({cols}, updated_by) VALUES ({", ".join("?" * len(data))}, ?)',
                               (*data.values(), by))
            pid, result = cur.lastrowid, 'insert'
        else:
            cur_row = _row(r)
            pid = cur_row['id']
            human_edited = (cur_row.get('updated_by') or '') != by
            changed = {}
            for k, v in clean.items():
                old = cur_row.get(k)
                # 사람이 고친 카드는 퇴사 관련 칸만 플렉스 값을 받는다(11-10: 일부러 비운 직책에 플렉스 '매니저'가
                # 다시 채워지지 않게). 플렉스에서 퇴사 처리된 건 놓치면 안 되니 그 칸들은 비어 있을 때 채움.
                if human_edited:
                    if k == 'status':
                        if v == STATUS_ACTIVE:          # 플렉스가 재직이면 사람이 정한 상태 유지
                            continue
                    elif k not in _FLEX_FOLLOW or old not in (None, ''):
                        continue
                if (old or None) != (v or None):
                    changed[k] = v
            if ssn_raw and not cur_row.get('ssn_enc'):
                try:
                    changed['ssn_enc'] = encrypt_ssn(ssn_raw)
                except ValueError:
                    pass
            if changed:
                sets = ', '.join(f'{k} = ?' for k in changed)
                conn.execute(f'UPDATE hr_people SET {sets}, updated_at = ? WHERE id = ?',
                             (*changed.values(), _now(), pid))
            result = 'update' if changed else 'skip'
        # 발령 이력: 플렉스 출처는 통째로 다시 씀(사람이 넣은 건 source 가 달라 보존)
        conn.execute("DELETE FROM hr_appointments WHERE person_id = ? AND source = 'flex'", (pid,))
        for a in appts:
            conn.execute('INSERT INTO hr_appointments (person_id, apply_date, department, job_role, job_rank, job_title, memo, source) '
                         "VALUES (?, ?, ?, ?, ?, ?, ?, 'flex')",
                         (pid, a.get('apply_date'), a.get('department'), a.get('job_role'), a.get('job_rank'),
                          a.get('job_title'), a.get('memo')))
    return result


# ── 경력·학력·가족·서류 (11-10 추가 — 플렉스 구성원 상세에 있던 것) ───────────
# 열람: 본인 + 인사카드 관리자. 가족 주민번호는 본인 주민번호처럼 암호화(가린 값만 화면에).
# 서류 파일: instance/hr_docs/<인사카드 번호>/ (git 제외). 올리기·지우기만 hr_access_log 에 남긴다(보기·내려받기는 안 남김, 11-10).
_DDL += [
    '''CREATE TABLE IF NOT EXISTS hr_careers (
        id INTEGER PRIMARY KEY AUTOINCREMENT, person_id INTEGER NOT NULL,
        company TEXT, department TEXT, job_role TEXT, job_rank TEXT, job_title TEXT, contract_type TEXT,
        begin_date TEXT, end_date TEXT, source TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_hr_careers_person ON hr_careers(person_id)',
    '''CREATE TABLE IF NOT EXISTS hr_educations (
        id INTEGER PRIMARY KEY AUTOINCREMENT, person_id INTEGER NOT NULL,
        school TEXT, major TEXT, level TEXT, graduate_type TEXT, begin_date TEXT, end_date TEXT, source TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_hr_educations_person ON hr_educations(person_id)',
    '''CREATE TABLE IF NOT EXISTS hr_families (
        id INTEGER PRIMARY KEY AUTOINCREMENT, person_id INTEGER NOT NULL,
        name TEXT, relation TEXT, gender TEXT, phone TEXT, ssn_enc TEXT,
        deduction_relation TEXT, default_deduction INTEGER NOT NULL DEFAULT 0,
        child_deduction INTEGER NOT NULL DEFAULT 0, health_dependent INTEGER NOT NULL DEFAULT 0,
        handicap TEXT, source TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_hr_families_person ON hr_families(person_id)',
    '''CREATE TABLE IF NOT EXISTS hr_documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT, person_id INTEGER NOT NULL,
        doc_type TEXT, file_name TEXT NOT NULL, stored_name TEXT NOT NULL, size INTEGER,
        flex_file_key TEXT UNIQUE, uploaded_by TEXT,
        created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP, tag TEXT
    )''',
    'CREATE INDEX IF NOT EXISTS idx_hr_documents_person ON hr_documents(person_id)',
]

DOC_TYPES = ['주민등록등본', '가족관계증명서', '건강보험득실확인서', '근로계약서', '연봉계약서', '이력서', '학력증명서',
             '경력증명서', '원천징수영수증', '신분증사본', '급여계좌사본', '장애인증명서', '국가유공자증명서',
             '연말정산 간소화자료', '연말정산 추가증빙', '기타']
DOC_EXT = ('.pdf', '.jpg', '.jpeg', '.png', '.heic', '.gif', '.webp', '.doc', '.docx', '.hwp', '.hwpx',
           '.xls', '.xlsx', '.zip')
DOC_MAX_BYTES = 8 * 1024 * 1024
_CONTRACT_KO = {'REGULAR': '정규직', 'CONTRACT': '계약직', 'DISPATCH': '파견', 'TEMPORARY': '계약직',
                'PART_TIME': '시간제', 'INTERN': '인턴'}
_EDU_LEVEL_KO = {'HIGH_SCHOOL': '고등학교', 'BACHELOR_COLLEGE': '전문대학', 'BACHELOR_UNIVERSITY': '대학교',
                 'MASTER': '대학원(석사)', 'DOCTOR': '대학원(박사)', 'MIDDLE_SCHOOL': '중학교'}
_GRAD_KO = {'GRADUATED': '졸업', 'ATTENDING': '재학', 'DROPPED_OUT': '중퇴', 'EXPECTED': '졸업 예정',
            'COMPLETED': '수료', 'ON_LEAVE': '휴학'}
_GENDER_KO = {'MALE': '남', 'FEMALE': '여'}


def docs_dir() -> str:
    d = os.getenv('GW_HR_DOCS_DIR', '').strip()
    if not d:
        from pathlib import Path
        d = str(Path(__file__).resolve().parents[2] / 'instance' / 'hr_docs')
    return d


def _rows(sql: str, args=()) -> List[dict]:
    with _conn() as conn:
        return [_row(r) for r in conn.execute(sql, args).fetchall()]


def careers(person_id: int) -> List[dict]:
    return _rows('SELECT * FROM hr_careers WHERE person_id = ? ORDER BY begin_date DESC, id', (person_id,))


def educations(person_id: int) -> List[dict]:
    return _rows('SELECT * FROM hr_educations WHERE person_id = ? ORDER BY begin_date DESC, id', (person_id,))


def families(person_id: int) -> List[dict]:
    """가족 — 주민번호는 가린 값만(ssn_masked), 암호문은 내보내지 않는다."""
    out = []
    for f in _rows('SELECT * FROM hr_families WHERE person_id = ? ORDER BY id', (person_id,)):
        enc = f.pop('ssn_enc', None)
        f['has_ssn'] = bool(enc)
        f['ssn_masked'] = ''
        if enc:
            try:
                f['ssn_masked'] = mask_ssn(decrypt_ssn(enc))
            except Exception:
                f['ssn_masked'] = '(읽기 실패)'
        for k in ('default_deduction', 'child_deduction', 'health_dependent'):
            f[k] = bool(f.get(k))
        out.append(f)
    return out


def documents(person_id: int, tag: Optional[str] = None) -> List[dict]:
    """서류 목록. tag 를 주면 그 꼬리표(예: 연말정산 ye2026)만."""
    _ensure_doc_tag()
    sql = 'SELECT id, person_id, doc_type, file_name, size, uploaded_by, created_at, tag FROM hr_documents WHERE person_id = ?'
    args = [person_id]
    if tag is not None:
        sql += ' AND tag = ?'
        args.append(tag)
    return _rows(sql + ' ORDER BY id', tuple(args))


def _migrate(conn) -> None:
    """이미 있는 운영 표에 나중에 생긴 칸 추가 — hr_documents.tag(11-10 연말정산용 꼬리표)."""
    import sqlite3
    try:
        conn.execute('ALTER TABLE hr_documents ADD COLUMN tag TEXT')
    except sqlite3.OperationalError:
        pass      # 이미 있음


def _ensure_doc_tag() -> None:
    """(예전 이름 유지) 칸 추가는 표를 처음 열 때 _migrate 가 한다."""
    with _conn():
        pass


def private_sections(person_id: int) -> dict:
    """경력·학력·가족·서류 — 본인·인사카드 관리자에게만 보낸다(권한 확인은 부르는 쪽)."""
    return {'careers': careers(person_id), 'educations': educations(person_id),
            'families': families(person_id), 'documents': documents(person_id), 'doc_types': DOC_TYPES}


def document_file(doc_id: int) -> Optional[dict]:
    with _conn() as conn:
        r = conn.execute('SELECT * FROM hr_documents WHERE id = ?', (doc_id,)).fetchone()
    d = _row(r)
    if d:
        d['path'] = os.path.join(docs_dir(), str(d['person_id']), d['stored_name'])
    return d


def _safe_name(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', '_', (name or '').strip()) or 'file'
    return name[-120:]


def save_document(person_id: int, doc_type: str, file_name: str, data: bytes, by: str,
                  flex_file_key: Optional[str] = None, tag: Optional[str] = None, trusted: bool = False) -> dict:
    """서류 파일 저장 + 장부 한 줄. 인사카드 관리자(또는 이관 스크립트 by='flex-import')만.
    trusted=True 는 부르는 쪽이 이미 권한을 확인한 경우(연말정산 본인 업로드 — by 에 본인 표시)."""
    _ensure_doc_tag()
    if not trusted and by != 'flex-import' and not is_hr_admin(by):
        raise PermissionError('서류 업로드는 경영지원·대표만 할 수 있습니다.')
    if not get_person(person_id):
        raise LookupError('없는 구성원입니다.')
    if not data:
        raise ValueError('빈 파일입니다.')
    if len(data) > DOC_MAX_BYTES:
        raise ValueError('파일이 너무 큽니다(8MB까지 업로드할 수 있습니다).')
    if not (file_name or '').lower().endswith(DOC_EXT):
        raise ValueError('업로드할 수 없는 파일 형식입니다.')
    doc_type = doc_type if doc_type in DOC_TYPES else '기타'
    folder = os.path.join(docs_dir(), str(person_id))
    os.makedirs(folder, exist_ok=True)
    with _conn() as conn:
        if flex_file_key:
            hit = conn.execute('SELECT id FROM hr_documents WHERE flex_file_key = ?', (flex_file_key,)).fetchone()
            if hit:
                return {'id': hit[0], 'skipped': True}
        cur = conn.execute('INSERT INTO hr_documents (person_id, doc_type, file_name, stored_name, size, flex_file_key, '
                           'uploaded_by, created_at, tag) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)',
                           (person_id, doc_type, _safe_name(file_name), 'pending', len(data), flex_file_key,
                            by.lower(), _now(), tag))
        doc_id = cur.lastrowid
        stored = f'{doc_id}_{_safe_name(file_name)}'
        with open(os.path.join(folder, stored), 'wb') as fp:
            fp.write(data)
        conn.execute('UPDATE hr_documents SET stored_name = ? WHERE id = ?', (stored, doc_id))
    if by != 'flex-import':
        log_access(person_id, by, 'doc_upload', doc_type)
    return {'id': doc_id, 'skipped': False}


def delete_document(doc_id: int, by: str, trusted: bool = False) -> None:
    """trusted=True 는 부르는 쪽이 이미 권한을 확인한 경우(연말정산 본인 자료 — 제출 전)."""
    if not trusted and not is_hr_admin(by):
        raise PermissionError('서류 삭제는 경영지원·대표만 할 수 있습니다.')
    d = document_file(doc_id)
    if not d:
        raise LookupError('없는 서류입니다.')
    with _conn() as conn:
        conn.execute('DELETE FROM hr_documents WHERE id = ?', (doc_id,))
    try:
        os.remove(d['path'])
    except OSError:
        pass
    log_access(d['person_id'], by, 'doc_delete', f'{d.get("doc_type") or ""} {d.get("file_name")}')


def reveal_family_ssn(family_id: int, viewer_email: str) -> str:
    if not is_hr_admin(viewer_email):
        raise PermissionError('주민등록번호 전체 보기는 경영지원·대표만 할 수 있습니다.')
    with _conn() as conn:
        r = conn.execute('SELECT person_id, name, ssn_enc FROM hr_families WHERE id = ?', (family_id,)).fetchone()
    if not r or not r[2]:
        raise LookupError('저장된 주민등록번호가 없습니다.')
    return decrypt_ssn(r[2])


def import_flex_details(person_id: int, careers_: List[Dict], educations_: List[Dict], families_: List[Dict]) -> dict:
    """플렉스 경력·학력·가족을 다시 쓴다(출처 flex 만 지우고 넣음 — 화면에서 사람이 넣은 줄은 보존)."""
    with _conn() as conn:
        for t in ('hr_careers', 'hr_educations', 'hr_families'):
            conn.execute(f"DELETE FROM {t} WHERE person_id = ? AND source = 'flex'", (person_id,))
        for c in careers_:
            conn.execute('INSERT INTO hr_careers (person_id, company, department, job_role, job_rank, job_title, '
                         "contract_type, begin_date, end_date, source) VALUES (?,?,?,?,?,?,?,?,?,'flex')",
                         (person_id, c.get('name'), c.get('departmentName'), c.get('jobRoleName'), c.get('jobRankName'),
                          c.get('jobTitleName'), _CONTRACT_KO.get(c.get('contractType') or '', c.get('contractType')),
                          (c.get('beginDate') or '')[:10] or None, (c.get('endDate') or '')[:10] or None))
        for e in educations_:
            conn.execute('INSERT INTO hr_educations (person_id, school, major, level, graduate_type, begin_date, end_date, '
                         "source) VALUES (?,?,?,?,?,?,?,'flex')",
                         (person_id, e.get('name'), e.get('majorName'),
                          _EDU_LEVEL_KO.get(e.get('level') or '', e.get('level')),
                          _GRAD_KO.get(e.get('graduateType') or '', e.get('graduateType')),
                          (e.get('beginDate') or '')[:10] or None, (e.get('endDate') or '')[:10] or None))
        for f in families_:
            enc = None
            if f.get('ssn'):
                try:
                    enc = encrypt_ssn(f['ssn'])
                except ValueError:
                    enc = None
            conn.execute('INSERT INTO hr_families (person_id, name, relation, gender, phone, ssn_enc, deduction_relation, '
                         'default_deduction, child_deduction, health_dependent, handicap, source) '
                         "VALUES (?,?,?,?,?,?,?,?,?,?,?,'flex')",
                         (person_id, f.get('name'), (f.get('relation') or {}).get('label'),
                          _GENDER_KO.get(f.get('gender') or ''), f.get('phoneNumber') or None, enc,
                          (f.get('deductionRelation') or {}).get('label'), int(bool(f.get('isDefaultDeduction'))),
                          int(bool(f.get('isChildrenDeduction'))), int(bool(f.get('isDependentHealthInsurance'))),
                          (f.get('handicap') or {}).get('label')))
    return {'careers': len(careers_), 'educations': len(educations_), 'families': len(families_)}
