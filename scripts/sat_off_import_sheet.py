"""격주휴무 신청서(엑셀) 한 달치를 그룹웨어 토요일 휴무표로 옮긴다(2026-11-20, 부장님 — 12월분부터 그룹웨어로).

드라이브 '공유 드라이브/기술영업 휴무자료/<연도>일정/<연도>년 격주휴무 신청서.xlsx' 의 월 시트(1~12)에서
토요일 칸 바로 아래 칸에 적힌 이름을 읽는다. 이름이 붙어 있거나(빈승정 이상효) 줄이 섞여 있어도
인사카드의 기술영업 명단과 맞는 이름만 넣는다. 명단에 없는 글자는 따로 보여 준다.

옮긴 달은 '작성 중'으로 둔다 — 빈승정 대리가 그룹웨어에서 확인하고 결재를 올리면 된다.
이미 그룹웨어에 고른 사람이 있는 달은 건드리지 않는다(덮어쓰기 없음).

사용:
  python scripts/sat_off_import_sheet.py 2026-12            # 미리보기(쓰지 않음)
  python scripts/sat_off_import_sheet.py 2026-12 --apply    # 넣기(넣기 전에 instance/users.db 백업)
create_app() 을 부르지 않는다(스케줄러 이중 기동 방지).
"""
import argparse
import re
import shutil
import sys
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

DEFAULT_SHEET = 'G:/공유 드라이브/공유 드라이브/기술영업 휴무자료/{y}일정/{y}년 격주휴무 신청서.xlsx'
SOURCE = '격주휴무 신청서 이관'


def read_month(path: str, month: str, roster):
    import openpyxl
    y, m = (int(x) for x in month.split('-'))
    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
    if str(m) not in wb.sheetnames:
        raise SystemExit(f'시트 {m} 이 없어요: {wb.sheetnames}')
    ws = wb[str(m)]
    out, leftover = {}, {}
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if not (isinstance(v, (datetime, date)) and v.weekday() == 5):
                continue
            d = v.date() if isinstance(v, datetime) else v
            if (d.year, d.month) != (y, m):
                continue
            raw = re.sub(r'\s', '', str(ws.cell(row=c.row + 1, column=c.column).value or ''))
            names = [n for n in roster if n in raw]
            rest = raw
            for n in names:
                rest = rest.replace(n, '')
            rest = re.sub(r'[/,·.\-]', '', rest)
            out[d.isoformat()] = sorted(names)
            if rest:
                leftover[d.isoformat()] = rest
    return out, leftover


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('month')
    ap.add_argument('--sheet', default='')
    ap.add_argument('--db', default=str(ROOT / 'instance' / 'users.db'))
    ap.add_argument('--apply', action='store_true')
    a = ap.parse_args()

    from dashboard.services import leave as lv
    from dashboard.services import sat_off as so
    from dashboard.utils.user_database import UserDatabase

    month = so.parse_month(a.month)
    if not month:
        raise SystemExit('달은 YYYY-MM 으로 적어 주세요.')
    path = a.db
    if a.apply:
        bak = f"{a.db}.bak-satoff-{month}-{datetime.now():%Y%m%d%H%M%S}"
        shutil.copy2(a.db, bak)
        print(f'백업: {bak}')
    else:                                   # 미리보기는 사본으로(표 만들기조차 운영 DB 에 하지 않게)
        import tempfile
        path = str(Path(tempfile.mkdtemp()) / 'users_preview.db')
        shutil.copy2(a.db, path)
    db = UserDatabase(db_path=path)
    lv._db = lambda: db

    roster = {m['name']: m for m in so.members()}
    sheet = a.sheet or DEFAULT_SHEET.format(y=month[:4])
    days, leftover = read_month(sheet, month, list(roster))
    for d in sorted(days):
        print(f"{d} (토): {', '.join(days[d]) or '-'}" + (f'   ※ 명단에 없는 글자: {leftover[d]}' if d in leftover else ''))

    have = {k: v for k, v in so.picks(month).items() if v}
    if have:
        print(f'그룹웨어에 이미 고른 사람이 있어 넣지 않아요: {have}')
        return
    if so.month_row(month)['status'] != so.OPEN:
        print('이미 결재 중이거나 확정된 달이라 넣지 않아요.')
        return
    if not a.apply:
        print('미리보기예요. 넣으려면 --apply 를 붙여 주세요.')
        return
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    n = 0
    with lv._conn() as conn:
        for d, names in days.items():
            if so.holiday_name(date.fromisoformat(d)):
                continue
            for name in names:
                conn.execute('INSERT OR IGNORE INTO sat_off_days (day, name, slack_id, added_by, added_at) '
                             'VALUES (?, ?, ?, ?, ?)', (d, name, roster[name]['slack_id'], SOURCE, now))
                n += 1
        conn.execute('INSERT INTO sat_off_log (at, actor, action, month, detail) VALUES (?, ?, ?, ?, ?)',
                     (now, SOURCE, 'import', month, f'엑셀 격주휴무 신청서에서 {n}건 옮김'))
    print(f'{n}건 넣었어요. 상태는 작성 중이에요.')


if __name__ == '__main__':
    main()
