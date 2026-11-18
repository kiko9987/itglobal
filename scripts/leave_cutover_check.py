# -*- coding: utf-8 -*-
"""휴가 기능 켜기 전 점검 — 읽기 전용 (2026-11-16 상태 머신 점검 묶음 G).

LEAVE_ENABLED 를 켜는 날 한꺼번에 터질 수 있는 것을 사람별로 보여 준다. DB 는 읽기 전용(mode=ro)으로만 연다.
  ① 휴가 명단 ↔ 인사카드: 퇴사·입사일·슬랙 ID 가 어긋난 사람
  ② accrue_from(이 날 뒤 연차만 자동 발생)이 빈 사람 — 비어 있으면 입사일부터 몰려 생긴다
  ③ 오늘 켜면 자동 발생이 넣을 연차(사람별) — 3줄 넘게 몰리는 사람은 자동 발생이 보류하고 관리자에게 알린다
  ④ 플렉스 잔여를 읽은 날(AS_OF) 뒤에 '신청된' 플렉스 연차 — 그 잔여엔 아직 안 빠진 것(11-18 R3: 플렉스는 휴가 날짜가 아니라
     신청이 잡힐 때 뺀다. 예전엔 휴가 날짜로 걸러 이미 빠진 걸 또 보여 줬음). 취소는 시각이 없어 여기 안 나옴 → 잔여 대조로.
     보관 기록(gw_leave_history)은 hr_import_flex.py 가 갱신한다 — 전환 날엔 그걸 먼저 돌릴 것(R4).
  ⑤ 오늘 뒤 날짜로 플렉스에서 확정된 휴가 — 장부로 옮길 대상(leave_import_flex.py ③). 옮겼는지 표시.
     승인 대기·취소 승인 대기는 전환 전에 플렉스에서 끝내야 한다.
  ⑥ 장부의 신청 수(상태별) — 시험 신청이 남아 있으면 정식 전환 전에 정리.

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
    out = {'mismatch': [], 'no_accrue_from': [], 'accrue': [], 'flex_after': [], 'flex_future': [], 'requests': {}}
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
    moved = set()
    if _table(conn, 'leave_requests'):
        for r in conn.execute("SELECT reason FROM leave_requests WHERE reason LIKE 'flex:%' "
                              "AND status IN ('wait_1','wait_2','approved')"):
            moved.add(r['reason'][5:])
        for r in conn.execute('SELECT status, COUNT(*) AS n FROM leave_requests GROUP BY status'):
            out['requests'][r['status']] = r['n']
    if _table(conn, 'gw_leave_history'):
        # 신청 시각(requested_at 'YYYY-MM-DD HH:MM')이 잔여를 읽은 날 다음 날 0시 이후 = 그 잔여에 아직 안 빠짐
        after = (as_of.toordinal() + 1)
        after_s = date.fromordinal(after).isoformat()
        for r in conn.execute("SELECT name, start, \"end\", days, unit, status, requested_at FROM gw_leave_history "
                              "WHERE type_label='연차' AND status IN ('승인','등록','승인 대기') AND requested_at >= ? "
                              'ORDER BY name, start', (after_s,)):
            out['flex_after'].append(dict(r))
        for r in conn.execute("SELECT flex_event_id, name, type_label, start, \"end\", days, unit, status FROM gw_leave_history "
                              "WHERE status IN ('승인','등록','승인 대기','취소 승인 대기') AND start > ? "
                              'ORDER BY start, name', (today.isoformat(),)):
            r = dict(r)
            r['moved'] = r['flex_event_id'] in moved
            out['flex_future'].append(r)
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
    print(f'④ {a.as_of} 뒤에 신청된 플렉스 연차(그 잔여엔 아직 안 빠짐 — 잔여를 다시 읽을 것)')
    for r in res['flex_after'] or [{}]:
        if not r:
            print('   없음 (보관 기록은 hr_import_flex.py 로 받은 날까지만)')
            break
        print(f"   {r['name']} {r['start']}~{r['end'] or r['start']} {r['days'] or ''}일 {r['unit'] or ''} "
              f"({r['status']}, 신청 {r['requested_at']})")
    print(f'⑤ {today} 뒤 날짜로 플렉스에서 잡힌 휴가(장부로 옮길 대상 — leave_import_flex.py ③)')
    for r in res['flex_future'] or [{}]:
        if not r:
            print('   없음')
            break
        mark = '✓ 장부에 옮김' if r['moved'] else ('✋ 플렉스에서 먼저 끝낼 것' if '대기' in r['status'] else '아직 안 옮김')
        print(f"   {r['name']} {r['start']}~{r['end'] or r['start']} {r['type_label']} {r['days'] or ''}일 "
              f"({r['status']}) — {mark}")
    print('⑥ 장부 신청 수(상태별)')
    print('   ' + (', '.join(f'{k} {v}건' for k, v in sorted(res['requests'].items())) or '0건'))
    return 0


if __name__ == '__main__':
    sys.exit(main())
