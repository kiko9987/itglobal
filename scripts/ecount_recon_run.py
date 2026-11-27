"""이카운트 밤 대조 손 실행(2026-11-28) — 스케줄러 없이 한 번 돌려 보기. create_app 을 쓰지 않는다.

  # 1) 이카운트 읽기만(비밀번호가 들어가는 단계 — 부장님이 직접 실행). 결과 JSON·로그는 instance/ecount_recon/
  node tools/ecount_reader/read_ecount.mjs --out instance/ecount_recon

  # 2) 읽은 결과로 시험 대조 — DB 를 바꾸지 않고 무엇이 바뀔지만 보여 줌
  .venv/Scripts/python.exe scripts/ecount_recon_run.py --from-file instance/ecount_recon/ecount_read_YYYYMMDD_HHMMSS.json

  # 3) 실제 반영(장부 맞춤 + ③ 대조 기록) — 장부는 자동 백업 뒤 조정 줄만 덧붙임
  .venv/Scripts/python.exe scripts/ecount_recon_run.py --from-file <json> --apply

  --from-file 없이 --apply 만 주면 이카운트 읽기부터 한 번에 함(밤 작업과 같음).
  --dm-preview 는 아침 DM 글을 보내지 않고 화면에만 찍음.
"""
import argparse
import os
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
os.chdir(ROOT)
sys.stdout.reconfigure(encoding='utf-8')

from dotenv import load_dotenv  # noqa: E402

load_dotenv(ROOT / '.env')

from dashboard.services import ecount_recon as er  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('--from-file', help='읽기 도구 결과 JSON(없으면 이카운트를 새로 읽음)')
    ap.add_argument('--apply', action='store_true', help='DB 에 반영(없으면 시험만)')
    ap.add_argument('--dm-preview', action='store_true', help='아침 DM 글 미리 보기(보내지 않음)')
    a = ap.parse_args()
    if not a.from_file and not a.apply:
        print('이카운트를 새로 읽으려면 --apply 를 함께 주세요(읽기는 node 도구로 따로 해도 돼요).')
        return 2
    res = er.run_nightly(read_path=a.from_file, apply=a.apply)
    if not res['ok']:
        print('[실패]', res['error'])
        return 1
    s = res['summary']
    print(f"{'[반영]' if a.apply else '[시험 — DB 그대로]'} 이카운트 매출 공사 {s['eco_codes']}개 · 매입 공사 {s['buy_codes']}개 · 공사 현황 {s['sheet_rows']}줄")
    print(f"장부 맞춤: {s['ledger_fixed'] if a.apply else len(res['fixes'])}건" + (' (한도 넘어 보류)' if s['ledger_held'] else ''))
    for f in res['fixes'][:60]:
        print(f"   {f['code']}: 장부 {f['cur'][0]:,}/{f['cur'][1]:,} → 이카운트 {f['eco'][0]:,}/{f['eco'][1]:,}")
    for side, found in (('매출', res['sales_found']), ('매입', res['buy_found']), ('시트 원가', res['cost_found'])):
        print(f'{side} 차이 {len(found)}건:', dict(Counter(i["category"] for i in found.values())))
    for c, i in sorted(res['sales_found'].items()):
        eco = f"{i['eco_supply']:,}/{i['eco_vat']:,}"
        sheet = '-' if i['sheet_supply'] is None else f"{i['sheet_supply']:,}/{i['sheet_vat']:,}"
        print(f"   {i['category']:8} {c}{'(이카운트 ' + i['ecount_code'] + ')' if i['ecount_code'] != c else ''} 이카운트 {eco} · 공사 현황 {sheet}")
    print(f"②에 '이카운트에 이미 있음' 경고가 붙을 공사: {len(res['mismatch'])}건")
    if a.apply:
        for side in ('sales', 'buy', 'cost'):
            r = s[side]
            print(f"{side}: 새로 {len(r.get('new') or [])} · 다시 처리 전 {len(r.get('reopened') or [])} · 맞음(완료) {len(r.get('resolved') or [])}")
    if a.dm_preview and a.apply:
        from dashboard.services import ecount_audit as ea
        text, blocks = er.build_dm(er.last_run(), er.upload_counts(), ea.open_items('sales'), ea.open_items('buy'))
        print('\n[아침 DM 미리 보기]', text)
        for b in blocks:
            print(b.get('text', {}).get('text') or b.get('elements', [{}])[0].get('text'))
    return 0


if __name__ == '__main__':
    sys.exit(main())
