"""이카운트 매출전표 장부 기준선 넣기 (한 번만, 2026-10-30).

장부(users.db ecount_sales_entries)의 출발점 = '이카운트에 이미 들어 있는 공사별 매출'.
이카운트는 조회 API 가 없어 시트 값으로 채우되, 이카운트 화면과 대조해 다른 공사는 --adjust 로 맞춘다.

  확정일 ≤ --cutoff 인 공사 → 시트 목표값(공급가=총액 1, 부가세=총액 2−총액 1)으로 기준선.
  확정일 > --cutoff 인 공사 → 넣지 않음 (첫 목록에 '공사 확정'으로 뜸. 이미 넣었으면 [넣음]).
  --skip 코드      → 이카운트에 매출전표가 없는 공사 (첫 목록에 뜸).
  --adjust 코드=공급가,부가세 → 이카운트 실제 합계가 시트와 다른 공사.

cutoff 다음 날 이후 금액·부가세가 바뀐 공사가 있으면 기준선이 틀어지므로 멈춘다(--allow-changed 로 확인 후 진행).
기본은 미리보기. --apply 를 줘야 장부에 쓴다. 이미 장부에 있는 공사는 건너뛰므로 다시 돌려도 안전.

예) .venv/Scripts/python.exe scripts/ecount_sales_seed.py --cutoff 2026-10-28 \
      --skip R4180-YM --adjust G4152-JW=11242000,1124200 --apply
"""
import argparse
import os
import sqlite3
import sys
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from dotenv import load_dotenv  # noqa: E402

load_dotenv(os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), '.env'))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--cutoff', required=True, help='이 날짜(포함)까지 확정된 공사를 기준선으로, YYYY-MM-DD')
    ap.add_argument('--skip', nargs='*', default=[])
    ap.add_argument('--adjust', nargs='*', default=[], help='코드=공급가,부가세')
    ap.add_argument('--allow-changed', nargs='*', default=[], help='cutoff 뒤 바뀌었지만 이카운트에 반영된 공사')
    ap.add_argument('--apply', action='store_true')
    a = ap.parse_args()

    from dashboard.services import ecount_sales as ecs
    from dashboard.services.project_service import get_project_records
    from dashboard.utils.user_database import get_user_database

    adjust = {}
    for s in a.adjust:
        code, nums = s.split('=', 1)
        sup, vat = nums.split(',')
        adjust[code.strip()] = (int(sup), int(vat))

    recs = get_project_records(force_refresh=True)
    if len(recs) < 1000:
        sys.exit(f'시트 레코드가 {len(recs)}건뿐 — 조회 실패로 보고 멈춤')

    rows, later = [], []
    for r in recs:
        code = str(r.get('프로젝트 코드') or '').strip()
        conf = str(r.get('공사 확정') or '').strip()[:10]
        if not code or not conf or code in a.skip:
            continue
        if conf > a.cutoff:
            later.append((code, conf))
            continue
        sup, vat = adjust.get(code, ecs.target_of(r))
        rows.append((code, sup, vat, 'adjust(이카운트 대조)' if code in adjust else ''))

    # cutoff 다음 날 0시(한국) 이후 금액·부가세·취소 변경 — 기준선이 '그 뒤 값'을 먹으면 차액이 사라진다
    since = (datetime.fromisoformat(a.cutoff) + timedelta(days=1) - timedelta(hours=9)).strftime('%Y-%m-%d %H:%M:%S')
    seeded = {c for c, *_ in rows}
    db = get_user_database()
    with db._get_connection() as conn:
        changed = conn.execute(
            "SELECT project_code, timestamp, field_name, old_value, new_value FROM audit_logs "
            "WHERE timestamp >= ? AND (field_name IN ('총액 1','부가세') OR action IN ('CANCEL_PROJECT','RESUME_PROJECT')) "
            "ORDER BY timestamp", (since,)).fetchall()
    bad = [c for c in changed if c[0] in seeded and c[0] not in a.allow_changed and c[0] not in adjust]

    print(f'기준선 {len(rows)}건 (조정 {len(adjust)}, 제외 {len(a.skip)}), cutoff 뒤 확정 {len(later)}건 → 첫 목록에 뜸')
    for code, conf in sorted(later, key=lambda x: x[1]):
        print(f'  확정 {conf} {code}')
    for c in changed:
        mark = '!!' if c in bad else '  '
        print(f'{mark} cutoff 뒤 변경 {c[0]} {c[1]} {c[2]}: {c[3]} -> {c[4]}')
    if bad:
        sys.exit('cutoff 뒤 바뀐 공사(!!)가 기준선에 들어갑니다. 이카운트 반영 여부 확인 후 --adjust 또는 --allow-changed')
    if not a.apply:
        print('미리보기 — 장부에 쓰려면 --apply')
        return
    try:
        n = ecs.seed_baseline(rows, by=f'seed:{a.cutoff}')
    except sqlite3.Error as exc:
        sys.exit(f'장부 쓰기 실패: {exc}')
    print(f'장부에 기준선 {n}건 넣음 (이미 있던 공사 {len(rows) - n}건 건너뜀)')


if __name__ == '__main__':
    main()
