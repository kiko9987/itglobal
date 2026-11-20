# -*- coding: utf-8 -*-
"""회계연도 부여 미리 보기(2026-11-18) — 휴가 명단으로 fiscal_grants 결과를 표로. 읽기만 한다.

  .venv/Scripts/python.exe scripts/leave_fiscal_preview.py [--today 2026-12-15]

운영 users.db 를 읽기 전용(mode=ro)으로 열고 사본에서 계산한다. create_app() 은 부르지 않는다(스케줄러 이중 기동 방지).
"""
import argparse
import shutil
import sqlite3
import sys
import tempfile
from datetime import date
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--today', default='2026-12-15')
    ap.add_argument('--db', default=str(ROOT / 'instance' / 'users.db'))
    a = ap.parse_args(argv)
    today = date.fromisoformat(a.today)
    tmp = Path(tempfile.mkdtemp()) / 'users.db'
    src = sqlite3.connect(f'file:{a.db}?mode=ro', uri=True)
    dst = sqlite3.connect(str(tmp))
    src.backup(dst)
    src.close()
    dst.close()
    try:
        from dashboard.services import leave as lv
        from dashboard.utils.user_database import UserDatabase
        d = UserDatabase(db_path=str(tmp))
        lv._db = lambda: d
        print(f'기준일 {today} · 회계연도 시작 {lv.FISCAL_FROM}\n')
        print(f"{'이름':5} {'입사일':10}  {'2027 비례분':>8}  {'2028 부여':>7}  내용")
        total = 0.0
        for p in lv.staff():
            g = {x['valid_from'].year: x for x in lv.fiscal_grants(p, date(2028, 1, 1))}
            now = {x['valid_from'].year for x in lv.fiscal_grants(p, today)}
            y27, y28 = g.get(2027), g.get(2028)
            total += y27['days'] if y27 else 0
            mark = '' if 2027 in now else ' (아직 안 생김)'
            print(f"{p['name']:5} {p['hire_date']:10}  {(y27 or {}).get('days', 0):>8g}  {(y28 or {}).get('days', 0):>7g}  "
                  f"{(y27 or {}).get('note', '-')}{mark}")
        print(f'\n2027 비례분 합: {total:g}일')
    finally:
        shutil.rmtree(tmp.parent, ignore_errors=True)
    return 0


if __name__ == '__main__':
    sys.exit(main())
