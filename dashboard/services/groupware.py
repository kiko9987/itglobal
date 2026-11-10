"""그룹웨어 새 화면(왼쪽 줄 메뉴) — 시범 명단과 메뉴 구성.

정식 오픈 전까지는 시범 명단에 있는 사람만 새 틀을 본다. 나머지는 기존 PM 화면 그대로.
- 시범 명단·전체 켜기는 사용자 DB의 별도 표(gw_pilots·gw_settings)에 둔다 → 재시작 없이 바로 반영.
- 시범 사용자 본인도 「예전 화면으로」를 누르면 세션 동안 기존 틀로 돌아간다(session['gw_off']).
- 기존 users 표는 건드리지 않는다.
"""
from __future__ import annotations

import time
from contextlib import contextmanager
from datetime import datetime

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

# 처음 표를 만들 때 넣는 시범 사용자(첫 시범은 KiKO 한 사람)
_SEED_PILOTS = ('kiko@itg-aircon.com',)

_DDL = [
    '''CREATE TABLE IF NOT EXISTS gw_pilots (
        email TEXT PRIMARY KEY,
        added_by TEXT,
        added_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )''',
    '''CREATE TABLE IF NOT EXISTS gw_settings (
        key TEXT PRIMARY KEY,
        value TEXT,
        updated_by TEXT,
        updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
    )''',
]
_table_ready = False

# 화면마다 DB를 읽지 않게 잠깐 기억(명단을 바꾸면 바로 지움)
_CACHE_SECONDS = 30
_cache = {'at': 0.0, 'pilots': frozenset(), 'open_all': False}


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
            seeded = conn.execute("SELECT value FROM gw_settings WHERE key = 'seeded'").fetchone()
            if not seeded:
                for email in _SEED_PILOTS:
                    conn.execute('INSERT OR IGNORE INTO gw_pilots (email, added_by) VALUES (?, ?)',
                                 (email, 'seed'))
                conn.execute("INSERT OR REPLACE INTO gw_settings (key, value, updated_by) VALUES ('seeded', '1', 'seed')")
            conn.commit()
            _table_ready = True
        yield conn
        conn.commit()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _invalidate() -> None:
    _cache['at'] = 0.0


def _load() -> dict:
    if time.time() - _cache['at'] < _CACHE_SECONDS:
        return _cache
    try:
        with _conn() as conn:
            pilots = frozenset(r[0].lower() for r in conn.execute('SELECT email FROM gw_pilots'))
            row = conn.execute("SELECT value FROM gw_settings WHERE key = 'open_all'").fetchone()
        _cache.update(at=time.time(), pilots=pilots, open_all=bool(row and row[0] == '1'))
    except Exception as e:  # DB 문제로 화면이 안 뜨는 일은 없게 — 기존 틀로
        logger.warning(f'[GW] 시범 명단 읽기 실패 → 기존 화면: {e}')
        _cache.update(at=time.time(), pilots=frozenset(), open_all=False)
    return _cache


# ── 명단 ──────────────────────────────────────────────────────────────
def is_open_all() -> bool:
    return _load()['open_all']


def is_pilot(email: str) -> bool:
    return bool(email) and email.lower() in _load()['pilots']


def can_use(email: str) -> bool:
    """새 틀을 쓸 수 있는 사람인가(전체 켜기 또는 시범 명단)."""
    return bool(email) and (is_open_all() or is_pilot(email))


def wants_new_layout(email: str, session_off: bool) -> bool:
    """이번 화면을 새 틀로 그릴지 — 쓸 수 있고, 본인이 「예전 화면」으로 돌리지 않았을 때."""
    return can_use(email) and not session_off


def list_pilots() -> list[dict]:
    with _conn() as conn:
        rows = conn.execute('SELECT email, added_by, added_at FROM gw_pilots ORDER BY added_at').fetchall()
    return [{'email': r[0], 'added_by': r[1], 'added_at': r[2]} for r in rows]


def add_pilot(email: str, by: str) -> None:
    email = (email or '').strip().lower()
    if not email or '@' not in email:
        raise ValueError('이메일 형식이 아닙니다.')
    with _conn() as conn:
        conn.execute('INSERT OR IGNORE INTO gw_pilots (email, added_by, added_at) VALUES (?, ?, ?)',
                     (email, by, _now()))
    _invalidate()


def remove_pilot(email: str) -> None:
    with _conn() as conn:
        conn.execute('DELETE FROM gw_pilots WHERE email = ?', ((email or '').strip().lower(),))
    _invalidate()


def set_open_all(on: bool, by: str) -> None:
    set_setting('open_all', '1' if on else '0', by)


def get_setting(key: str):
    """그룹웨어 설정 값(문자열) — 없으면 None."""
    with _conn() as conn:
        row = conn.execute('SELECT value FROM gw_settings WHERE key = ?', (key,)).fetchone()
    return row[0] if row else None


def set_setting(key: str, value: str, by: str) -> None:
    with _conn() as conn:
        conn.execute('INSERT OR REPLACE INTO gw_settings (key, value, updated_by, updated_at) VALUES (?, ?, ?, ?)',
                     (key, value, by, _now()))
    _invalidate()


# ── 메뉴 ──────────────────────────────────────────────────────────────
# 왼쪽 줄 = 영역, 위쪽 줄 = 그 영역의 하위 메뉴. 주소는 기존 그대로.
# admin=True 는 관리자(admin·super_admin)에게만 보인다(기존 상단 버튼과 같은 조건).
AREAS = [
    {
        'key': 'home', 'label': '홈', 'icon': 'fas fa-home',
        'items': [{'href': '/home', 'label': '홈', 'icon': 'fas fa-home', 'field': True}],
    },
    {
        'key': 'sales', 'label': '영업', 'icon': 'fas fa-briefcase',
        'items': [
            {'href': '/projects', 'label': '프로젝트', 'icon': 'fas fa-folder-open'},
            {'href': '/leads', 'label': '고객 리드', 'icon': 'fas fa-user-tie'},
            {'href': '/stats', 'label': '매출통계', 'icon': 'fas fa-chart-bar'},
            {'href': '/case-posts', 'label': '시공사례', 'icon': 'fas fa-camera', 'admin': True},
        ],
    },
    {
        'key': 'support', 'label': '경영지원', 'icon': 'fas fa-coins',
        'items': [
            {'href': '/settlement/collected', 'label': '월별 수금', 'icon': 'fas fa-file-invoice-dollar'},
        ],
    },
    {
        'key': 'leave', 'label': '휴가', 'icon': 'fas fa-umbrella-beach',
        'items': [
            {'href': '/leave', 'label': '내 휴가', 'icon': 'fas fa-calendar-check', 'field': True},
            {'href': '/leave/team', 'label': '전사 현황', 'icon': 'fas fa-calendar-alt', 'leave_all': True},
        ],
    },
    {
        'key': 'hr', 'label': '인사', 'icon': 'fas fa-id-card',
        'items': [
            {'href': '/hr/me', 'label': '내 정보', 'icon': 'fas fa-user', 'field': True},
            {'href': '/hr/people', 'label': '구성원', 'icon': 'fas fa-users', 'hr_admin': True},
        ],
    },
    {
        'key': 'admin', 'label': '관리', 'icon': 'fas fa-cog', 'bottom': True,
        'items': [
            {'href': '/gw/admin', 'label': '그룹웨어 설정', 'icon': 'fas fa-sliders-h', 'admin': True},
        ],
    },
]


def _matches(path: str, href: str) -> bool:
    return path == href or path.startswith(href.rstrip('/') + '/')


def _visible(it: dict, is_admin: bool, is_hr_admin: bool, field: bool = False, leave_all: bool = False) -> bool:
    if field:                       # 현장·미화 세션: field=True 메뉴만
        return bool(it.get('field'))
    if it.get('admin') and not is_admin:
        return False
    if it.get('hr_admin') and not is_hr_admin:
        return False
    if it.get('leave_all') and not (leave_all or is_hr_admin):
        return False
    return True


def build_nav(path: str, permission_level: str, is_hr_admin: bool = False, field: bool = False,
              leave_all: bool = False) -> dict:
    """현재 주소에 맞춰 왼쪽 줄(영역)과 위쪽 줄(하위 메뉴)을 만든다.
    bottom=True 영역(관리)은 왼쪽 줄 아래쪽에 따로 둔다. field=True 는 현장·미화 세션."""
    is_admin = permission_level in ('admin', 'super_admin')
    areas, active_items = [], []
    for area in AREAS:
        items = [dict(it) for it in area['items'] if _visible(it, is_admin, is_hr_admin, field, leave_all)]
        if not items:
            continue
        # 가장 길게 맞는 메뉴 하나만 켠다(/leave/team 에서 /leave 까지 켜지지 않게)
        best = max((it['href'] for it in items if _matches(path, it['href'])), key=len, default=None)
        for it in items:
            it['active'] = it['href'] == best
        hit = best is not None
        areas.append({'key': area['key'], 'label': area['label'], 'icon': area['icon'],
                      'href': items[0]['href'], 'active': hit, 'bottom': bool(area.get('bottom'))})
        if hit:
            active_items = items
    return {'areas': areas, 'items': active_items}
