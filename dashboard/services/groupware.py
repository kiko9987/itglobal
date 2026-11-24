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


# 개발 중 화면(차량, 11-21 부장님 — 스케줄 근무는 11-25 시공·미화 빼고 모두에게 엶) — 이 명단만 들어가고, 다른 사람은 '아직 개발 중' 안내.
# 메뉴는 그대로 보인다. 모두에게 열 때는 비운다(()).
DEV_PREVIEW_USERS: tuple = ('kiko@itg-aircon.com',)


# 방문 일정(영업 > 방문 일정)만 관리자(admin·super_admin)에게 먼저 연다(11-23 부장님) — 고치기·확정은 그대로 일정 담당만.
# 차량 배정은 위 명단 그대로. 모두에게 열 때는 False 로(그러면 일정 담당 + 오늘 온라인 당번 메뉴로 돌아감).
VISIT_PLAN_ADMIN_ONLY: bool = True
ADMIN_LEVELS = ('admin', 'super_admin')


def dev_preview_ok(email: str) -> bool:
    return not DEV_PREVIEW_USERS or (email or '').strip().lower() in DEV_PREVIEW_USERS


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
        # 현장·미화 세션엔 안 보임 — /home 은 현장이면 /leave 로 넘겨 '홈'과 '휴가'가 같은 화면을 열었음(11-16 최종 점검)
        'items': [{'href': '/home', 'label': '홈', 'icon': 'fas fa-home'}],
    },
    {
        # 결재(11-10) — 슬랙 전자결재와 같은 장부. 작성은 슬랙, 결재함·처리 단추·플렉스 보관함은 여기
        'key': 'approval', 'label': '결재', 'icon': 'fas fa-file-signature',
        'items': [
            {'href': '/approval', 'label': '결재함', 'icon': 'fas fa-inbox'},
            {'href': '/approval/archive', 'label': '플렉스 보관함', 'icon': 'fas fa-box-archive'},
        ],
    },
    {
        'key': 'sales', 'label': '영업', 'icon': 'fas fa-briefcase',
        'items': [
            # 순서(11-18 부장님): 자주 쓰는 프로젝트 → 일을 따오는 입구(리드·입찰) → 끝난 공사 글(시공사례) → 보고(통계)
            {'href': '/projects', 'label': '프로젝트', 'icon': 'fas fa-folder-open'},
            {'href': '/leads', 'label': '고객 리드', 'icon': 'fas fa-user-tie'},
            {'href': '/visits/plan', 'label': '방문 일정', 'icon': 'fas fa-route', 'planner': True,   # 캔버스2·/일정확정 대신(11-22)
             'also': ('/vehicles/assign',)},   # 차량 배정은 방문 일정 탭으로만 들어감(11-24) — 그 화면에서도 이 메뉴가 켜짐
            {'href': '/bids', 'label': '입찰 공고', 'icon': 'fas fa-gavel'},   # 나라장터(11-18)
            {'href': '/case-posts', 'label': '시공 사례', 'icon': 'fas fa-camera', 'admin': True},
            {'href': '/stats', 'label': '매출 통계', 'icon': 'fas fa-chart-bar'},
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
            {'href': '/leave/team', 'label': '전사 휴가 관리', 'icon': 'fas fa-calendar-alt', 'leave_all': True},
            {'href': '/leave/saturday', 'label': '스케줄 근무', 'icon': 'fas fa-calendar-week', 'sat_off': True},   # 기술영업 격주휴무(11-20)
        ],
    },
    {
        'key': 'hr', 'label': '인사', 'icon': 'fas fa-id-card',
        'items': [
            {'href': '/hr/me', 'label': '내 정보', 'icon': 'fas fa-user', 'field': True},
            {'href': '/hr/people', 'label': '구성원', 'icon': 'fas fa-users', 'hr_admin': True},
            {'href': '/yearend', 'label': '연말정산', 'icon': 'fas fa-receipt', 'field': True},
            {'href': '/yearend/admin', 'label': '연말정산 현황', 'icon': 'fas fa-clipboard-check', 'hr_admin': True},
        ],
    },
    {
        # 차량(11-20) — 운행 기록은 위치라 경영지원만. 차량 배정은 11-24 부장님: 이 영역에서 빼고 방문 일정 탭으로만.
        # 휴가·인사 뒤에 둔다 — 현장 세션은 첫 영역이 로고 자리라 휴가가 먼저여야 한다.
        'key': 'vehicle', 'label': '차량', 'icon': 'fas fa-car',
        'items': [
            {'href': '/vehicles/now', 'label': '실시간 위치', 'icon': 'fas fa-location-dot'},   # 사무실 전 직원(11-24). 현장·미화는 뺌(11-25 부장님)
            {'href': '/vehicles', 'label': '운행일지', 'icon': 'fas fa-route', 'hr_admin': True},   # U+커넥트 운행(11-22 이름)
            {'href': '/vehicles/manage', 'label': '차량 관리', 'icon': 'fas fa-screwdriver-wrench', 'hr_admin': True},  # 11-22
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


def _uplus_on() -> bool:
    """U+ 메뉴를 보일지 — 키가 있거나, 키를 잘못 적었을 때(화면 머리에 '설정 오류'가 보이게, 11-24 조용한 실패 점검)."""
    try:
        from dashboard.services import uplus_api as up
        return up.enabled('status') or bool(up.config_error('status'))
    except Exception:
        return False


def _visible(it: dict, is_admin: bool, is_hr_admin: bool, field: bool = False, leave_all: bool = False,
             sat_off: bool = False, planner: bool = False) -> bool:
    if it.get('uplus') and not _uplus_on():   # U+ 키가 들어오기 전엔 빈 화면이라 숨김(11-24)
        return False
    if field:                       # 현장·미화 세션: field=True 메뉴만
        return bool(it.get('field'))
    if it.get('admin') and not is_admin:
        return False
    if it.get('hr_admin') and not is_hr_admin:
        return False
    if it.get('leave_all') and not (leave_all or is_hr_admin):
        return False
    if it.get('sat_off') and not sat_off:
        return False
    if it.get('planner') and not planner:      # 방문 일정 — 일정 담당(JW·대표님·부장님)만
        return False
    return True


def build_nav(path: str, permission_level: str, is_hr_admin: bool = False, field: bool = False,
              leave_all: bool = False, sat_off: bool = False, planner: bool = False) -> dict:
    """현재 주소에 맞춰 왼쪽 줄(영역)과 위쪽 줄(하위 메뉴)을 만든다.
    bottom=True 영역(관리)은 왼쪽 줄 아래쪽에 따로 둔다. field=True 는 현장·미화 세션."""
    is_admin = permission_level in ('admin', 'super_admin')
    areas, active_items = [], []
    owner = next((it['href'] for area in AREAS for it in area['items']
                  if any(_matches(path, x) for x in it.get('also', ()))
                  and _visible(it, is_admin, is_hr_admin, field, leave_all, sat_off, planner)), None)
    for area in AREAS:
        items = [dict(it) for it in area['items'] if _visible(it, is_admin, is_hr_admin, field, leave_all, sat_off, planner)]
        if not items:
            continue
        # 가장 길게 맞는 메뉴 하나만 켠다(/leave/team 에서 /leave 까지 켜지지 않게)
        best = max((it['href'] for it in items if _matches(path, it['href'])), key=len, default=None)
        if owner:               # 메뉴에 없는 화면은 맡은 메뉴만 켠다(차량 배정 → 방문 일정, /vehicles 운행일지가 켜지지 않게)
            best = owner if owner in [it['href'] for it in items] else None
        for it in items:
            it['active'] = it['href'] == best
        hit = best is not None
        # children = 그 영역의 하위 메뉴(왼쪽 줄을 펼쳤을 때 바로 고를 수 있게, 당근 비즈니스처럼 — 11-10)
        areas.append({'key': area['key'], 'label': area['label'], 'icon': area['icon'],
                      'href': items[0]['href'], 'active': hit, 'bottom': bool(area.get('bottom')),
                      'children': items})
        if hit:
            active_items = items
    return {'areas': areas, 'items': active_items}
