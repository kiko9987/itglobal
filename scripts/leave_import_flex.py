# -*- coding: utf-8 -*-
"""플렉스 휴가 현황 → 휴가 장부 (2026-11-05).

플렉스엔 내려받기가 없어 휴가 관리 → 휴가 보유 현황·구성원 관리 화면에서 읽은 값을 아래 SNAPSHOT 에 적고 돌린다.
옮기는 날 다시 읽어 SNAPSHOT 을 고치고 한 번 더 돌리면 된다(같은 사람은 덮어쓴다).

  - 직원: 이름·슬랙 ID·입사일. 슬랙이 없는 직원은 slack_id 를 비워 둔다(경영지원이 대신 신청).
  - 연차: 플렉스 남은 연차를 'flex:이름' 부여 한 줄로. 쓸 수 있는 기간 = 다음 입사 기념일 전날까지.
  - accrue_from = 기준일 다음 날 — 그 뒤 생기는 연차만 자동 발생이 넣는다(그 전 몫은 플렉스 잔여에 들어 있음).
  - 이미 쓴(잡힌) 연차가 있는 'flex:' 줄은 days 만 고친다.

사용:  .venv/Scripts/python.exe scripts/leave_import_flex.py           # 미리 보기(쓰지 않음)
       .venv/Scripts/python.exe scripts/leave_import_flex.py --apply   # 장부에 쓰기
create_app() 은 부르지 않는다(스케줄러 이중 기동 방지).
"""
import sys
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

AS_OF = date(2026, 11, 5)   # 플렉스 화면을 읽은 날

# (이름, 슬랙 ID, 입사일, 플렉스 남은 연차) — 2026-11-05 휴가 보유 현황·구성원 관리
SNAPSHOT = [
    ('강정권', 'U0BHH8UD99P', '2026-04-01', 4),
    ('고광일', 'U04UL2ZLJAX', '2025-01-01', 14),      # 플렉스 15.5 → KiKO 가 누락분 반영해 14 로 고침(11-05)
    ('권진현', 'U0CETGE1480', '2026-11-02', 0),
    ('권태숙', 'U0CFWU7U14G', '2023-07-24', 12),       # 미화 — 개인 메일로 슬랙 가입(11-05)
    ('권태훈', 'U0BJ9M096EL', '2024-04-01', 5),
    ('김종연', 'U0CF4E3BWQH', '2022-05-23', 7),        # 시공 — 개인 메일로 슬랙 가입(11-05)
    ('김태현', 'U0CFWU7KW8Y', '2022-05-23', 10),       # 시공 — 개인 메일로 슬랙 가입(11-05)
    ('박민우', 'U0BJ9M1MRJ4', '2022-02-03', 8),
    ('박민재', 'U0BH93XCS1H', '2024-10-02', 14),
    ('박용구', 'U050XNG5VBR', '2007-01-01', 24),
    ('박정우', 'U051ENK8AAY', '2019-08-07', 18),
    ('빈승정', 'U0BHDBCEUHL', '2024-08-12', 13),
    ('이상효', 'U0CDTS9UEH3', '2026-11-02', 0),
    ('이종원', 'U0CE4T9MFHB', '2026-11-03', 0),
    ('조성헌', 'U0BJ9M0QFK2', '2025-04-22', 0),        # 플렉스 빈칸 — 확인 필요
    ('주영민', 'U0BHK2UET44', '2024-04-08', 6),
    ('황샛별', 'U0BHC2JV7U5', '2023-04-04', 1),
]


def next_anniversary(hire: date, after: date) -> date:
    from dashboard.services.leave import add_years
    n = 1
    while add_years(hire, n) <= after:
        n += 1
    return add_years(hire, n)


def main(apply: bool) -> None:
    from dashboard.services import leave as lv
    accrue_from = (AS_OF + timedelta(days=1)).isoformat()
    for name, uid, hire_s, left in SNAPSHOT:
        hire = date.fromisoformat(hire_s)
        until = next_anniversary(hire, AS_OF) - timedelta(days=1)
        print(f'{name:4} 입사 {hire_s}  슬랙 {uid or "(없음)":13}  남은 연차 {left:>4}일  ~{until}')
        if not apply:
            continue
        lv.upsert_staff(name, hire_s, uid, accrue_from=accrue_from, note='플렉스 이관')
        with lv._conn() as conn:
            # 기준일까지 자동으로 생긴 월차·기념일 연차는 플렉스 잔여에 이미 들어 있다 — 안 쓴 줄은 지워 두 번 안 잡히게(11-05 점검)
            gone = conn.execute(
                "DELETE FROM leave_grants WHERE name=? AND (source LIKE 'monthly:%' OR source LIKE 'annual:%') "
                'AND valid_from<=? AND id NOT IN (SELECT grant_id FROM leave_uses)', (name, AS_OF.isoformat())).rowcount
            if gone:
                print(f'      자동 발생 {gone}줄 지움(플렉스 잔여에 포함)')
            used = conn.execute("SELECT COALESCE(SUM(u.days),0) FROM leave_uses u JOIN leave_grants g ON g.id=u.grant_id "
                                'WHERE g.source=?', (f'flex:{name}',)).fetchone()[0]
        if used and left < used:
            print(f'      ⚠️ 슬랙으로 이미 {used:g}일 썼는데 플렉스 잔여가 {left}일 — 확인 필요(그대로 넣으면 음수)')
        if not lv.add_grant(name, left, AS_OF, until, f'flex:{name}', f'플렉스 잔여 ({AS_OF})'):
            # 잔여 0 이 된 사람도 고친다(예전 값이 남으면 연차가 부풀려짐)
            with lv._conn() as conn:
                conn.execute('UPDATE leave_grants SET days=?, valid_from=?, valid_to=?, note=? WHERE source=?',
                             (float(left), AS_OF.isoformat(), until.isoformat(), f'플렉스 잔여 ({AS_OF})',
                              f'flex:{name}'))
    print('\n' + ('장부에 썼습니다.' if apply else '미리 보기입니다. --apply 로 쓰기.'))


if __name__ == '__main__':
    main('--apply' in sys.argv)
