"""Flex 자금집행 기록 → 전자결재 거래처 계좌 기억(payee_accounts) 넣기 (2026-11-01, 한 번만).

입력: instance/flex_export/flex_workflow_export_20261101.json (Flex 워크플로우 '내 문서함' 내보내기)
- 완료(DONE)된 자금집행요청* 문서만. 거래처마다 가장 최근 계좌 하나.
- 상호가 거래처 이름이 아닌 것('.', '개인', '개인용달' 등)·반환/환불 건은 뺀다(엉뚱한 계좌가 채워지지 않게).
- last_used = 그 문서 작성일(최근 순 목록이 실제 거래 순서가 되게). 이미 있는 거래처는 건드리지 않는다.
create_app() 을 부르지 않는다(스케줄러 이중 기동 방지).

  .venv/Scripts/python.exe scripts/eapproval_seed_payees_from_flex.py           # 미리보기(CSV 저장)
  .venv/Scripts/python.exe scripts/eapproval_seed_payees_from_flex.py --apply   # 실제로 넣기
"""
import argparse
import csv
import json
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from dotenv import load_dotenv  # noqa: E402

load_dotenv(os.path.join(ROOT, '.env'))

SRC = os.path.join(ROOT, 'instance', 'flex_export', 'flex_workflow_export_20261101.json')
PREVIEW = os.path.join(ROOT, 'instance', 'flex_export', 'payee_seed_preview.csv')
# 거래처 이름이 아니라 '종류'를 적은 상호 — 건마다 다른 사람 계좌라 기억하면 엉뚱하게 채워진다
_GENERIC = re.compile(r'^(개인|개인통장|개인용달|개인화물|개별용달|용달|용달화물|화물|퀵|퀵서비스|사다리|사다리차|크레인|'
                      r'기타|없음|미정|현금|카드|인건비|일당|일용직|기사)$')
_REFUND = re.compile(r'반환|환불|과입금|과수금')


def build(docs):
    from dashboard.services.e_approval import partner_key, normalize_account
    best, counts, accts, skipped = {}, {}, {}, []
    for d in docs:
        if not d['t'].startswith('자금집행요청') or d['s'] != 'DONE':
            continue
        f = d['f']
        partner = str(f.get('상호명') or '').strip()
        key = partner_key(partner)
        acct = normalize_account(str(f.get('계좌번호') or ''))
        holder, bank = str(f.get('예금주명') or '').strip(), str(f.get('은행명') or '').strip()
        if len(key) < 2 or _GENERIC.match(key) or _REFUND.search(str(f.get('내용') or '')):
            skipped.append(partner or '(빈칸)')
            continue
        if len(re.sub(r'\D', '', acct)) < 8 or not (holder and bank):
            skipped.append(partner)
            continue
        counts[key] = counts.get(key, 0) + 1
        accts.setdefault(key, set()).add(re.sub(r'\D', '', acct))
        if key not in best or d['at'] > best[key]['at']:
            best[key] = {'partner': partner, 'payee_name': holder, 'bank': bank, 'account': acct, 'at': d['at']}
    rows = []
    for k, b in best.items():
        n, na = counts[k], len(accts[k])
        if n >= 3 and na / n >= 0.75:   # 거의 매번 다른 계좌 = 종류 이름(예: 사다리차 4건 4계좌)
            skipped.append(f"{b['partner']}({n}건 {na}계좌)")
            continue
        rows.append({**b, 'key': k, 'n': n, 'accts': na})
    rows.sort(key=lambda r: (-r['n'], r['partner']))
    return rows, skipped


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--apply', action='store_true')
    args = ap.parse_args()
    with open(SRC, encoding='utf-8') as fh:
        docs = json.load(fh)['docs']
    rows, skipped = build(docs)
    with open(PREVIEW, 'w', encoding='utf-8-sig', newline='') as fh:
        w = csv.writer(fh)
        w.writerow(['상호', '예금주', '은행', '계좌', '최근 거래일', '거래 건수', '쓴 계좌 수'])
        for r in rows:
            w.writerow([r['partner'], r['payee_name'], r['bank'], r['account'], r['at'][:10], r['n'], r['accts']])
    print(f'거래처 {len(rows)}곳 (계좌 여러 개 {sum(1 for r in rows if r["accts"] > 1)}곳) / 뺀 건 {len(skipped)}')
    print(f'미리보기: {PREVIEW}')
    print('뺀 상호(중복 제외):', ', '.join(sorted(set(skipped))))
    if not args.apply:
        print('미리보기만 했습니다. 넣으려면 --apply')
        return
    from dashboard.services import e_approval as ea
    added = kept = 0
    with ea._db()._get_connection() as conn:
        ea._ensure_tables(conn)
        for r in rows:
            cur = conn.execute(
                'INSERT OR IGNORE INTO payee_accounts (partner_key, partner, payee_name, bank, account, last_used) '
                'VALUES (?,?,?,?,?,?)',
                (r['key'], r['partner'], r['payee_name'], r['bank'], r['account'],
                 r['at'][:19].replace('T', ' ')))
            if cur.rowcount:
                added += 1
            else:
                kept += 1
        conn.commit()
    print(f'넣음 {added}곳 / 이미 있어서 그대로 {kept}곳')


if __name__ == '__main__':
    main()
