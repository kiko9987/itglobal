# -*- coding: utf-8 -*-
"""플렉스 휴가 현황 → 휴가 장부 (2026-11-05, 전환 준비 11-18).

플렉스엔 내려받기가 없어 휴가 관리 → 휴가 보유 현황·구성원 관리 화면에서 읽은 값을 아래 SNAPSHOT 에 적고 돌린다.
옮기는 날 다시 읽어 SNAPSHOT 을 고치고 한 번 더 돌리면 된다(같은 사람은 덮어쓴다).

  - 직원: 이름·슬랙 ID·입사일. 슬랙이 없는 직원은 slack_id 를 비워 둔다(경영지원이 대신 신청).
  - 연차: 플렉스 남은 연차를 'flex:이름' 부여 한 줄로. 쓸 수 있는 기간 = 다음 입사 기념일 전날까지.
  - accrue_from = 기준일 다음 날 — 그 뒤 생기는 연차만 자동 발생이 넣는다(그 전 몫은 플렉스 잔여에 들어 있음).
  - 이미 쓴(잡힌) 연차가 있는 'flex:' 줄은 days 만 고친다.

전환 준비(11-17 EGM 검토 → 11-18 반영):
  - --as-of: 플렉스를 다시 받은 날(전환 날). 기본은 아래 AS_OF.
  - 잔여 대조(R9): 플렉스 요약 엑셀(플렉스_요약_YYYYMMDD.xlsx '구성원' 시트 '연차 잔여(YY-MM)')과 SNAPSHOT 을 비교만 한다.
    다르면 고칠 목록을 보여 주고 --apply 를 멈춘다(손 입력 실수 방지). 엑셀 값이 맞으면 SNAPSHOT 을 고쳐 다시 돌린다.
  - 플렉스로 이미 승인된 앞날 휴가 옮기기(E7·R7): 플렉스 보관 기록(gw_leave_history — hr_import_flex.py 가 먼저 갱신)에서
    --as-of 뒤 날짜의 '연차' 승인 건을 장부에 '승인 완료' 신청으로 넣는다(DM·공지 없음). 해지 뒤에도 취소·겹침 검사가 되게.
      · 플렉스는 신청이 잡힌 순간 그 휴가 날짜의 달 몫에서 잔여를 뺀다(11-18 김태현·권태숙 월별 현황으로 확인).
        그래서 휴가 날짜가 --as-of 달 이내면 그 일수는 이미 SNAPSHOT 잔여에서 빠져 있다 → 'flex:' 줄을 같은 일수만큼
        올리고 거기서 뗀다(남은 연차는 그대로). 다음 달 이후 날짜면 아직 안 빠졌으므로 올리지 않고 평소처럼 뗀다.
      · 승인 대기·취소 승인 대기 건, 반나절이 섞인 여러 날, 연차 아닌 종류(경조사·군소집 등)는 옮기지 않고 목록만 보여 준다.
      · 다시 돌려도 같은 건은 다시 넣지 않는다(사유 칸 'flex:<플렉스 ID>').

사용:  .venv/Scripts/python.exe scripts/leave_import_flex.py [--as-of YYYY-MM-DD] [--summary 요약.xlsx]           # 미리 보기
       .venv/Scripts/python.exe scripts/leave_import_flex.py --as-of YYYY-MM-DD --apply                       # 장부에 쓰기
       (--allow-diff: 잔여 대조가 달라도 쓰기 — 엑셀이 틀린 걸 확인했을 때만)
create_app() 은 부르지 않는다(스케줄러 이중 기동 방지).
"""
import argparse
import glob
import os
import sys
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Tuple

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

AS_OF = date(2026, 11, 10)   # 플렉스 자료를 읽은 날(11-05 화면 → 11-10 다시 받은 자료로 갱신). 전환 날엔 --as-of 로
FLEX_DIR = Path(r'C:\Users\SECOM\Desktop\ITG-Project\flex_export_20261106')

# (이름, 슬랙 ID, 입사일, 플렉스 남은 연차) — 2026-11-05 휴가 보유 현황 → 2026-11-10 다시 받은 자료로 대조(플렉스_요약_20261110.xlsx)
SNAPSHOT = [
    ('강정권', 'U0BHH8UD99P', '2026-04-01', 4),
    ('고광일', 'U04UL2ZLJAX', '2025-01-01', 14),      # 플렉스 15.5 → KiKO 가 누락분 반영해 14 로 고침(11-05)
    ('권진현', 'U0CETGE1480', '2026-11-02', 0),
    ('권태숙', 'U0CFWU7U14G', '2023-07-24', 12),       # 미화 — 개인 메일로 슬랙 가입(11-05)
    ('권태훈', 'U0BJ9M096EL', '2024-04-01', 5),
    ('김민욱', 'U0CFWGS61HU', '2026-11-09', 0),        # 11-09 입사(11-10 플렉스에서 확인). 첫해는 매달 1일씩 자동 발생
    ('김종연', 'U0CF4E3BWQH', '2022-05-23', 7),        # 시공 — 개인 메일로 슬랙 가입(11-05). 11-18 플렉스 8(11-30 연차 취소)
    ('김태현', 'U0CFWU7KW8Y', '2022-05-23', 9),        # 시공 — 개인 메일로 슬랙 가입(11-05). 11-18 플렉스 7(11-20·12-24 새 승인)
    ('박민우', 'U0BJ9M1MRJ4', '2022-02-03', 8),
    ('박민재', 'U0BH93XCS1H', '2024-10-02', 14),
    ('박용구', 'U050XNG5VBR', '2007-01-01', 24),
    ('박정우', 'U051ENK8AAY', '2019-08-07', 18),
    ('빈승정', 'U0BHDBCEUHL', '2024-08-12', 13),
    ('이상효', 'U0CDTS9UEH3', '2026-11-02', 0),
    ('이종원', 'U0CE4T9MFHB', '2026-11-03', 0),
    ('조성헌', 'U0BJ9M0QFK2', '2025-04-22', 0),        # 플렉스 빈칸 = 다 씀(11-05 KiKO 확인)
    ('주영민', 'U0BHK2UET44', '2024-04-08', 6),
    ('황샛별', 'U0BHC2JV7U5', '2023-04-04', 1),
]
# 김영숙 대표이사(플렉스 잔여 26일)는 휴가 장부에 넣지 않음 — KiKO 11-10 결정.

UNIT_TYPE = {'DAY': 'annual', 'HALF_DAY_AM': 'half_am', 'HALF_DAY_PM': 'half_pm'}
DONE = ('승인', '등록')                       # 플렉스 상태 중 '확정된 휴가'
WAITING = ('승인 대기', '취소 승인 대기')      # 전환 전에 플렉스에서 정리해야 할 것
FLEX_ACTOR = ('flex', '플렉스 이관')


def next_anniversary(hire: date, after: date) -> date:
    from dashboard.services.leave import add_years
    n = 1
    while add_years(hire, n) <= after:
        n += 1
    return add_years(hire, n)


# ── 잔여 대조 (요약 엑셀 ↔ SNAPSHOT) ────────────────────────────────────
def latest_summary(flex_dir: Path = FLEX_DIR) -> Optional[str]:
    files = sorted(glob.glob(os.path.join(str(flex_dir), '플렉스_요약_*.xlsx')))
    return files[-1] if files else None


def latest_promotion(flex_dir: Optional[Path] = None) -> Optional[str]:
    files = sorted(glob.glob(str((flex_dir or FLEX_DIR) / '원본' / 'flex_promotion_*.json')))
    return files[-1] if files else None


def summary_left(path: str) -> Tuple[str, Dict[str, Optional[float]]]:
    """요약 엑셀 '구성원' 시트의 재직자 연차 잔여 → (칸 이름, {이름: 일수 또는 None(빈칸)})."""
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True)['구성원']
    rows = ws.iter_rows(values_only=True)
    head = list(next(rows))
    col = next((i for i, h in enumerate(head) if str(h or '').startswith('연차 잔여')), None)
    if col is None:
        raise SystemExit(f"{path}: '연차 잔여' 칸이 없습니다")
    i_name, i_state = head.index('이름'), head.index('상태')
    out: Dict[str, Optional[float]] = {}
    for r in rows:
        if not r or not r[i_name] or '퇴사' in str(r[i_state] or ''):
            continue
        v = r[col]
        out[str(r[i_name])] = None if v in (None, '') else float(v)
    return head[col], out


def compare_left(snapshot, flex_left: Dict[str, Optional[float]]) -> List[Tuple[str, float, Optional[float]]]:
    """SNAPSHOT 과 엑셀 잔여가 다른 사람 [(이름, SNAPSHOT, 엑셀)] — 엑셀 빈칸은 0 으로 본다(조성헌 등 다 쓴 사람)."""
    diff = []
    for name, _uid, _hire, left in snapshot:
        if name not in flex_left:
            diff.append((name, float(left), None))
            continue
        x = flex_left[name] or 0.0
        if abs(x - float(left)) > 1e-9:
            diff.append((name, float(left), x))
    return diff


# ── 플렉스로 이미 승인된 앞날 휴가 ─────────────────────────────────────
def covered(start: str, as_of: date) -> bool:
    """휴가 날짜가 as_of 달 이내 → 플렉스 잔여(받은 달 월말 값)에서 이미 빠짐."""
    return start[:7] <= as_of.strftime('%Y-%m')


def flex_future(conn, as_of: date) -> Dict[str, List[dict]]:
    """gw_leave_history 의 as_of 뒤 날짜 휴가 → {'move': 옮길 연차·반차, 'skip': 손으로 볼 것(이유 포함)}."""
    out = {'move': [], 'skip': []}
    if not conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='gw_leave_history'").fetchone():
        return out
    for r in conn.execute('SELECT flex_event_id, name, type_label, start, "end", days, unit, status FROM gw_leave_history '
                          'WHERE start > ? ORDER BY start, name', (as_of.isoformat(),)):
        r = dict(r)
        if r['status'] in WAITING:
            out['skip'].append(dict(r, why='플렉스에서 승인·취소를 먼저 끝내 주세요'))
        elif r['status'] not in DONE:
            continue                                   # 취소·반려 — 옮길 것 없음
        elif r['type_label'] != '연차':
            out['skip'].append(dict(r, why='연차 아닌 휴가 — 필요하면 장부에서 손으로'))
        elif r['unit'] not in UNIT_TYPE:
            out['skip'].append(dict(r, why=f"단위 {r['unit']} — 반나절 섞인 여러 날은 손으로"))
        else:
            out['move'].append(r)
    return out


def _flex_gid(conn, name: str, on: str = '') -> Optional[int]:
    """그 사람의 'flex:' 줄 id — on(휴가 마지막 날)까지 살아 있을 때만(같은 달에 기념일이 끼면 None → 손으로)."""
    row = conn.execute('SELECT id, valid_to FROM leave_grants WHERE source=?', (f'flex:{name}',)).fetchone()
    if not row or (on and row[1] < on):
        return None
    return row[0]


def moved_covered_days(conn, name: str, as_of: date) -> float:
    """이미 옮긴 플렉스 휴가 중 플렉스 잔여에서 빠져 있던(= flex 줄을 올려 준) 일수."""
    rows = conn.execute("SELECT start, days FROM leave_requests WHERE name=? AND reason LIKE 'flex:%' "
                        "AND status IN ('wait_1','wait_2','approved')", (name,)).fetchall()
    return sum(r[1] for r in rows if covered(r[0], as_of))


def move_one(conn, r: dict, slack_id: str, as_of: date, now: str) -> str:
    """한 건 옮기기 → 'moved' | 'exists' | 'no_balance'. 커밋은 부르는 쪽."""
    from dashboard.services import leave as lv
    tag = f"flex:{r['flex_event_id']}"
    if conn.execute('SELECT 1 FROM leave_requests WHERE reason=?', (tag,)).fetchone():
        return 'exists'
    ltype = UNIT_TYPE[r['unit']]
    end = r['start'] if ltype in lv.HALF_TYPES else (r['end'] or r['start'])
    days = float(r['days'] or lv.count_days(ltype, date.fromisoformat(r['start']), date.fromisoformat(end)))
    cur = conn.execute(
        'INSERT INTO leave_requests (name, slack_id, submitter_id, submitter_name, ltype, sub, start, "end", days, '
        'reason, status, created_at, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)',
        (r['name'], slack_id or None, FLEX_ACTOR[0], FLEX_ACTOR[1], ltype, None, r['start'], end, days, tag,
         lv.APPROVED, now, now))
    rid = cur.lastrowid
    if covered(r['start'], as_of):
        gid = _flex_gid(conn, r['name'], end)
        if gid is None:
            conn.execute('DELETE FROM leave_requests WHERE id=?', (rid,))
            return 'no_balance'
        conn.execute('UPDATE leave_grants SET days = days + ? WHERE id=?', (days, gid))
        conn.execute('INSERT INTO leave_uses (request_id, grant_id, days) VALUES (?,?,?)', (rid, gid, days))
    elif not lv._allocate(conn, rid, r['name'], date.fromisoformat(r['start']), days, date.fromisoformat(end)):
        conn.execute('DELETE FROM leave_requests WHERE id=?', (rid,))
        return 'no_balance'
    for step, actor, result, why in (('request', FLEX_ACTOR, 'requested', None),
                                      ('approve_1', FLEX_ACTOR, 'auto', '플렉스에서 승인'),
                                      ('approve_2', FLEX_ACTOR, 'auto', '플렉스에서 승인')):
        conn.execute('INSERT INTO leave_steps (request_id, step, actor_id, actor_name, result, reason, at) '
                     'VALUES (?,?,?,?,?,?,?)', (rid, step, actor[0], actor[1], result, why, now))
    return 'moved'


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--apply', action='store_true')
    ap.add_argument('--as-of', default=AS_OF.isoformat(), help='플렉스를 다시 받은 날(전환 날)')
    ap.add_argument('--summary', default='', help='플렉스 요약 엑셀(기본: 가장 최근 플렉스_요약_*.xlsx)')
    ap.add_argument('--allow-diff', action='store_true', help='잔여 대조가 달라도 쓰기(엑셀이 틀린 걸 확인했을 때만)')
    ap.add_argument('--promotion', default='', help='플렉스 촉진 기록 JSON(기본: 원본 폴더의 가장 최근 flex_promotion_*.json)')
    a = ap.parse_args(argv)
    as_of = date.fromisoformat(a.as_of)
    apply = a.apply

    from dashboard.services import leave as lv

    # ① 잔여 대조
    path = a.summary or latest_summary()
    if path and os.path.exists(path):
        col, flex_left = summary_left(path)
        diff = compare_left(SNAPSHOT, flex_left)
        print(f'① 잔여 대조: {os.path.basename(path)} [{col}] ↔ SNAPSHOT')
        for name, snap, x in diff:
            print(f'   {name}: SNAPSHOT {snap:g} / 엑셀 {"(없음)" if x is None else f"{x:g}"}  ← SNAPSHOT 을 고쳐 주세요')
        if not diff:
            print('   모두 같음')
        if diff and apply and not a.allow_diff:
            print('\n잔여가 달라 쓰지 않았습니다. SNAPSHOT 을 고치거나, 엑셀이 틀린 걸 확인했으면 --allow-diff.')
            return 1
    else:
        print('① 잔여 대조: 요약 엑셀이 없어 건너뜀(--summary 로 지정)')

    # ② 잔여 줄
    print(f'\n② 잔여 줄 (기준일 {as_of})')
    accrue_from = (as_of + timedelta(days=1)).isoformat()
    for name, uid, hire_s, left in SNAPSHOT:
        hire = date.fromisoformat(hire_s)
        until = next_anniversary(hire, as_of) - timedelta(days=1)
        print(f'{name:4} 입사 {hire_s}  슬랙 {uid or "(없음)":13}  남은 연차 {left:>4}일  ~{until}')
        if not apply:
            continue
        lv.upsert_staff(name, hire_s, uid, accrue_from=accrue_from, note='플렉스 이관')
        with lv._conn() as conn:
            # 기준일까지 자동으로 생긴 월차·기념일 연차는 플렉스 잔여에 이미 들어 있다 — 안 쓴 줄은 지워 두 번 안 잡히게(11-05 점검)
            gone = conn.execute(
                "DELETE FROM leave_grants WHERE name=? AND (source LIKE 'monthly:%' OR source LIKE 'annual:%') "
                'AND valid_from<=? AND id NOT IN (SELECT grant_id FROM leave_uses)', (name, as_of.isoformat())).rowcount
            if gone:
                print(f'      자동 발생 {gone}줄 지움(플렉스 잔여에 포함)')
            used = conn.execute("SELECT COALESCE(SUM(u.days),0) FROM leave_uses u JOIN leave_grants g ON g.id=u.grant_id "
                                'WHERE g.source=?', (f'flex:{name}',)).fetchone()[0]
            extra = moved_covered_days(conn, name, as_of)        # 이미 옮긴 플렉스 휴가 몫(잔여에서 빠져 있던 것)
        total = float(left) + extra
        if used and total < used:
            print(f'      ⚠️ 슬랙으로 이미 {used:g}일 썼는데 플렉스 잔여가 {left}일 — 확인 필요(그대로 넣으면 음수)')
        if not lv.add_grant(name, total, as_of, until, f'flex:{name}', f'플렉스 잔여 ({as_of})'):
            # 잔여 0 이 된 사람도 고친다(예전 값이 남으면 연차가 부풀려짐)
            with lv._conn() as conn:
                conn.execute('UPDATE leave_grants SET days=?, valid_from=?, valid_to=?, note=? WHERE source=?',
                             (total, as_of.isoformat(), until.isoformat(), f'플렉스 잔여 ({as_of})', f'flex:{name}'))

    # ③ 플렉스로 이미 승인된 앞날 휴가
    print(f'\n③ 플렉스로 승인된 {as_of} 뒤 휴가 → 장부 \'승인 완료\'로 (DM·공지 없음)')
    slack = {n: u for n, u, _h, _l in SNAPSHOT}
    with lv._conn() as conn:
        fut = flex_future(conn, as_of)
        now = lv._now()
        for r in fut['move']:
            how = '플렉스 잔여에서 이미 빠짐 → flex 줄을 올리고 거기서 뗌' if covered(r['start'], as_of) else '다음 달 이후 → 평소처럼 뗌'
            line = f"   {r['name']} {r['start']}~{r['end'] or r['start']} {r['days']:g}일 {r['unit']} — {how}"
            if r['name'] not in slack:
                print(line + '  ⚠️ 휴가 명단에 없음 — 건너뜀')
                continue
            if not apply:
                print(line)
                continue
            res = move_one(conn, r, slack[r['name']], as_of, now)
            print(line + {'moved': '  ✓ 옮김', 'exists': '  (이미 옮김)', 'no_balance': '  ⚠️ 뗄 연차가 없어 못 옮김'}[res])
        conn.commit()
        if not fut['move']:
            print('   없음')
        for r in fut['skip']:
            print(f"   ✋ {r['name']} {r['start']} {r['type_label']} {r['days'] or ''}일 ({r['status']}) — {r['why']}")
        # 장부엔 옮겼는데 플렉스에서 그 뒤 취소된 것
        if conn.execute("SELECT 1 FROM sqlite_master WHERE type='table' AND name='gw_leave_history'").fetchone():
            for r in conn.execute("SELECT q.id, q.name, q.start, h.status FROM leave_requests q JOIN gw_leave_history h "
                                  "ON q.reason = 'flex:' || h.flex_event_id WHERE q.status IN ('wait_1','wait_2','approved') "
                                  "AND h.status NOT IN ('승인','등록')").fetchall():
                print(f'   ⚠️ 장부 #{r[0]} {r[1]} {r[2]} — 플렉스에선 지금 {r[3]}. 장부에서 취소해 주세요')
    # ④ 연차 촉진 기록(11-18) — 플렉스 촉진 문서(flex_promotion_YYYYMMDD.json, scripts/flex/flex_promotion_backup.js)
    promo = a.promotion or latest_promotion()
    print(f"\n④ 연차 촉진 기록 ← {os.path.basename(promo) if promo else '(파일 없음 — 건너뜀)'}")
    if promo and os.path.exists(promo):
        import json
        from dashboard.services import leave_promo as lp
        rows = lp.import_flex(promo, apply=apply)
        for r in rows:
            if r['grant_to'] < as_of.isoformat():
                continue                     # 끝난 기간은 기록만(목록에선 생략)
            plan = ', '.join(x['date'] for x in json.loads(r['plan_dates'])) or '-'
            print(f"   {r['name']} ~{r['grant_to']} 촉구 {r['noticed_at'][:10]} {r['days']:g}일 · {r['note']} · "
                  f"계획 {plan} · 지정 기한 {r['designate_due']}")
        print(f'   전체 {len(rows)}줄(끝난 기간 포함)' + (' 넣음(이미 있는 문서는 그대로)' if apply else ''))

    print('\n' + ('장부에 썼습니다.' if apply else '미리 보기입니다. --apply 로 쓰기.'))
    return 0


if __name__ == '__main__':
    sys.exit(main())
