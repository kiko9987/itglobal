# -*- coding: utf-8 -*-
"""플렉스 결재 보관함 갱신용 내보내기 만들기 (2026-11-20 해지 전 점검).

보관함(flex_approval_archive)은 11-01 「내 문서함」 내보내기(1,305건)로 채워져 있어 그 뒤 문서·상태가 빠진다.
11-06 전체 받기(flex_data_*.json, 상세 1,336건)에 다시 받기(flex_refresh_*.json, SINCE 뒤 바뀐 문서 상세)를 덮어
같은 모양(flex_workflow_export_*.json)으로 만든다 → scripts/gw_import_flex_approvals.py --file 로 넣는다(같은 번호는 덮어씀).

  .venv/Scripts/python.exe scripts/flex/flex_archive_refresh.py                 # 미리 보기(쓰지 않음)
  .venv/Scripts/python.exe scripts/flex/flex_archive_refresh.py --write         # instance/flex_export/ 에 파일 쓰기
  .venv/Scripts/python.exe scripts/gw_import_flex_approvals.py --file instance/flex_export/flex_workflow_export_YYYYMMDD.json --apply

선택 칸(SELECT) 값은 플렉스 원본이 '["일반"]' 처럼 목록 글자라 '일반' 으로 푼다(11-01 내보내기와 같게).
메모 = 본문 HTML 을 글자로. 카드번호는 보관함에 넣을 때 뒤 4자리만 남는다(gw_approval.import_flex_export).
"""
import argparse
import glob
import html
import json
import os
import re
import sys
from datetime import datetime, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
FLEX_RAW = os.path.join(os.path.dirname(ROOT), 'flex_export_20261106', '원본')
OUT_DIR = os.path.join(ROOT, 'instance', 'flex_export')


def _text(h: str) -> str:
    h = re.sub(r'<br\s*/?>|</p>|</div>|</li>', '\n', h or '', flags=re.I)
    h = html.unescape(re.sub(r'<[^>]+>', '', h)).replace('\xa0', ' ')
    return re.sub(r'\n{3,}', '\n\n', h).strip()


def _value(v) -> str:
    v = '' if v is None else str(v)
    if v.startswith('['):
        try:
            lst = json.loads(v)
            if isinstance(lst, list):
                return ', '.join(str(z) for z in lst)
        except ValueError:
            pass
    return v


def compact(d: dict) -> dict:
    f = {}
    for i in sorted(d.get('inputs') or [], key=lambda i: (i.get('inputField') or {}).get('displayOrder') or 0):
        name = (i.get('inputField') or {}).get('name')
        if name:
            f[name] = _value(i.get('value'))
    return {'c': d.get('code'), 't': d.get('title') or '', 's': d.get('status') or '', 'w': (d.get('writer') or {}).get('name') or '',
            'at': d.get('writtenAt') or d.get('createdAt') or '', 'up': d.get('updatedAt') or '',
            'f': f, 'att': len(d.get('attachments') or []), 'memo': _text(d.get('content'))}


def main(argv=None) -> int:
    sys.stdout.reconfigure(encoding='utf-8')
    ap = argparse.ArgumentParser()
    ap.add_argument('--full', default=os.path.join(FLEX_RAW, 'flex_data_20261106.json'))
    ap.add_argument('--refresh', help='다시 받기 파일(기본: 원본 폴더의 가장 최근 flex_refresh_*.json)')
    ap.add_argument('--write', action='store_true')
    a = ap.parse_args(argv)
    refresh = a.refresh or sorted(glob.glob(os.path.join(FLEX_RAW, 'flex_refresh_*.json')))[-1]
    docs = {}
    for path in (a.full, refresh):                       # 뒤 파일(다시 받기)이 앞 것을 덮는다
        for x in json.load(open(path, encoding='utf-8')).get('approvalDocs') or []:
            d = x.get('document', x)
            if d.get('code'):
                docs[d['code']] = compact(d)
    by = {}
    for d in docs.values():
        by[d['s']] = by.get(d['s'], 0) + 1
    print(f'전체 {os.path.basename(a.full)} + 다시 받기 {os.path.basename(refresh)} → 문서 {len(docs)}건 {by}')
    for d in docs.values():
        if d['s'] == 'IN_PROGRESS':
            print(f"  아직 진행 중(플렉스에서 정리 필요): {d['c']} {d['t']} {d['w']} {d['at'][:10]}")
    if not a.write:
        print('→ 미리 보기만(--write 로 파일 쓰기)')
        return 0
    stamp = re.search(r'(\d{8})', os.path.basename(refresh)).group(1)
    out = os.path.join(OUT_DIR, f'flex_workflow_export_{stamp}.json')
    os.makedirs(OUT_DIR, exist_ok=True)
    json.dump({'exported_at': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ'),
               'source': f'flex.team 전체 받기 + 다시 받기({stamp})', 'count': len(docs),
               'docs': sorted(docs.values(), key=lambda d: d['at'], reverse=True)},
              open(out, 'w', encoding='utf-8'), ensure_ascii=False)
    print('→ 썼습니다:', out)
    return 0


if __name__ == '__main__':
    sys.exit(main())
