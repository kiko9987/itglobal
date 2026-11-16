# -*- coding: utf-8 -*-
"""휴가 기능 켜기 전 점검 — 읽기 전용 (2026-11-16 상태 머신 점검 묶음 G).

LEAVE_ENABLED 를 켜는 날 한꺼번에 터질 수 있는 것을 사람별로 보여 준다. DB 는 읽기 전용(mode=ro)으로만 연다.
  ① 휴가 명단 ↔ 인사카드: 퇴사·입사일·슬랙 ID 가 어긋난 사람
  ② accrue_from(이 날 뒤 연차만 자동 발생)이 빈 사람 — 비어 있으면 입사일부터 몰려 생긴다
  ③ 오늘 켜면 자동 발생이 넣을 연차(사람별) — 3줄 넘게 몰리는 사람은 자동 발생이 보류하고 관리자에게 알린다
  ④ 플렉스 잔여를 읽은 날(AS_OF) 뒤의 플렉스 연차 기록 — 잔여에 이미 빠졌는지 경영지원이 확인(두 번 잡힘 방지).
     보관 기록(gw_leave_history)은 가져온 날까지만 있으니, 옮기는 날 플렉스에서 다시 받아 확인할 것.

사용:  .venv/Scripts/python.exe scripts/leave_cutover_check.py [--db 경로] [--today YYYY-MM-DD]
create_app() 은 부르지 않는다(스케줄러 이중 기동 방지). 아무것도 쓰지 않는다.
"""
import argparse
import sqlite3
import sys
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

DEFAULT_DB = ROOT / 'instance' / 'users.db'
HOLD_OVER = 3          # dashboard.services.leave.ACCRUE_HOLD_OVER 와 같게


def _ro(path: Path) -> sqlite3.Connection:
    conn = sqlite3.connect(f'file:{path.as_posix()}?mode=ro', uri=True)
    conn.row_factory = sqlite3.Row
    return conn


def _table(conn, name: str) -> bool:
    return bool(conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name=?", (name,)).fetchone())


def check(conn: sqlite3.Connection, today: date, as_of: date) -> dict:
    """점검 결과 dict — 테스트가 그대로 확인한다."""
    from dashboard.services.leave import due_grants
    staff = [dict(r) for r in conn.execute('SELECT * FROM leave_staff ORDER BY name')]
    hr = {}
    if _table(conn, 'hr_people'):
        for r in conn.execute("SELECT name, status, hire_date, slack_id FROM hr_people "
                              "WHERE slack_id IS NOT NULL AND slack_id != '' ORDER BY id"):
            r = dict(r)
            cur = hr.get(r['slack_id'])
            if cur is None or (cur['status'] == '퇴사' and r['status'] != '퇴사'):
                hr[r['slack_id']] = r
    have = {}
    for r in conn.execute('SELECT name, source FROM leave_grants'):
        have.setdefault(r['name'], set()).add(r['source'])
    out = {'mismatch': [], 'no_accrue_from': [], 'accrue': [], 'flex_after': []}
    for s in staff:
        p = hr.get(s.get('slack_id') or '')
        if not p:
            out['mismatch'].append((s['name'], '인사카드에 같은 슬랙 ID 없음'))
        else:
            if p['status'] == '퇴사' and s.get('active'):
                out['mismatch'].append((s['name'], '인사카드 퇴사인데 휴가 명단 재직'))
            if p.get('hire_date') and p['hire_date'] != s.get('hire_date'):
                out['mismatch'].append((s['name'], f"입사일 다름: 휴가 명단 {s.get('hire_date')} / 인사카드 {p['hire_date']}"))
        if not s.get('active') or (p and p['status'] == '퇴사'):
            continue
        if not s.get('accrue_from'):
            out['no_accrue_from'].append(s['name'])
        person = dict(s, hire_date=(p or {}).get('hire_date') or s.get('hire_date'))
        due = [g for g in due_grants(person, today) if g['source'] not in have.get(s['name'], set())]
        if due:
            out['accrue'].append((s['name'], len(due), sum(g['days'] for g in due), len(due) > HOLD_OVER))
    if _table(conn, 'gw_leave_history'):
        for r in conn.execute("SELECT name, start, \"end\", days, unit, status FROM gw_leave_history "
                              "WHERE type_label='연차' AND status IN ('승인','등록','승인 대기') AND start > ? "
                              'ORDER BY name, start', (as_of.isoformat(),)):
            out['flex_after'].append(dict(r))
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--db', default=str(DEFAULT_DB))
    ap.add_argument('--today', default='')
    ap.add_argument('--as-of', default='2026-11-10', help='플렉스 잔여를 읽은 날(scripts/leave_import_flex.AS_OF)')
    a = ap.parse_args()
    today = date.fromisoformat(a.today) if a.today else date.today()
    res = check(_ro(Path(a.db)), today, date.fromisoformat(a.as_of))
    print(f'휴가 기능 켜기 전 점검 — 기준일 {today}, 플렉스 잔여 읽은 날 {a.as_of}\n')
    print('① 휴가 명단 ↔ 인사카드')
    for n, why in res['mismatch'] or [('-', '어긋난 사람 없음')]:
        print(f'   {n}: {why}')
    print('② accrue_from 이 빈 사람(입사일부터 자동 발생)')
    print('   ' + (', '.join(res['no_accrue_from']) or '없음'))
    print('③ 오늘 켜면 자동 발생이 넣을 연차')
    for n, cnt, days, held in res['accrue'] or [('-', 0, 0, False)]:
        if n == '-':
            print('   없음')
            break
        print(f"   {n}: {cnt}줄 {days:g}일" + ('  ← 한꺼번에 몰림: 자동 발생이 보류하고 관리자에게 알림' if held else ''))
    print(f'④ {a.as_of} 뒤 플렉스 연차 기록(잔여에서 이미 빠졌는지 확인)')
    for r in res['flex_after'] or [{}]:
        if not r:
            print('   없음 (보관 기록은 가져온 날까지만 — 옮기는 날 플렉스에서 다시 확인)')
            break
        print(f"   {r['name']} {r['start']}~{r['end'] or r['start']} {r['days'] or ''}일 {r['unit'] or ''} ({r['status']})")
    return 0


if __name__ == '__main__':
    sys.exit(main())
