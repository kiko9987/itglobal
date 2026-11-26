"""차량 유류비·하이패스(통행료) — SB 법인카드 월 지출결의서에서 자동으로(2026-11-26 부장님 '자동으로 읽어 와').

주유 카드·하이패스 카드는 차마다 따로라 결재 내용에 차 번호가 안 남는다 → 결재에서 모으는 차량 관리 비용
(vehicle_manage.sync_costs)엔 유류·통행료가 거의 비어 있었다(1~10월 유류 9만 6천 원 vs 실제 약 3,500만 원).
SB 가 카드 청구서가 나오면(다음 달 20일 무렵) 손으로 채우는 파일을 읽기만 한다 — SB 일은 늘지 않음.

파일: 공유 드라이브 경영지원관리Ⅱ / AP 지급계정 / 3. 법인카드 / 'YYYY년 법인카드 월 지출결의서(SB).xlsx'
- 시트 = 'YYMM'(2601~). 칸: B 종류(유류비·하이패스 — 첫 줄에만), C NO, D 카드번호, E 차량번호, F 차량종류, G 사용액.
- 구획은 B 칸에 다음 글(…합계·신용카드·개인지급·종류 등)이 나오면 끝.
- 차는 차량번호 끝 4자리로 찾는다('81루 4904' 같은 오타도). 사람 이름 줄은 PERSON_PLATE 로만(박민우=MW 스파크).
- 지금 차 목록에 없는 줄도 1년 통계엔 필요해서 버리지 않는다(11-26 부장님): 판 차(SOLD — GV80)는 그 번호로,
  차를 모르는 카드(퇴사자 강성환 등)는 UNASSIGNED('미지정') 로 넣고 내용에 누구 카드인지 적는다.
  차량 관리 첫 화면 표 아래 '지난 차·차량 미지정' 줄로 보이고 올해 차량비 합계에도 들어감.
- 차량 관리엔 한 달에 한 줄(그 달 마지막 날, '10월 유류비 (월 합계)'). SB 가 금액을 고치면 다음 읽기 때 같이 고침.
- 서버(LocalSystem)는 G: 가 안 보여 서비스 계정 Drive API 로 받는다(자금일보와 같은 방식).
"""
from __future__ import annotations

import calendar
import os
import re
import tempfile
from datetime import date
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

DRIVE_NAME = '경영지원관리Ⅱ'
FOLDER_PATH = ('AP 지급계정', '3. 법인카드')
FILE_RE = re.compile(r'^(\d{4})년 법인카드 월 지출결의서.*\.xlsx$')
SECTIONS = {'유류비': '유류', '하이패스': '통행료'}
# 차량번호 칸에 사람 이름이 적힌 카드 — MW 스파크(개인 차, 회사가 유류비 지원, 11-22). 11월 시트부터는 53우1783 로 적힘
PERSON_PLATE = {'박민우': '53우1783', 'MW': '53우1783'}
# 지금 목록에 없는 차·카드(1년 통계용) — 화면 이름도 여기서
SOLD = {'254누2593': 'GV80 2593 (26.10.02 매각)'}
UNASSIGNED = '미지정'
EXTRA_LABELS = dict(SOLD, **{UNASSIGNED: '차량 미지정 (퇴사자 카드 등)'})
# 같은 카드의 다른 이름 — 하이패스 B704931 이 1~2월 'SH', 그 뒤 오른쪽 표에 '강성환'(2610 시트)
PERSON_ALIAS = {'SH': '강성환'}
SOURCE = 'card'


def _month_of(sheet: str) -> Optional[date]:
    m = re.fullmatch(r'(\d{2})(\d{2})', str(sheet).strip())
    if not m:
        return None
    y, mon = 2000 + int(m.group(1)), int(m.group(2))
    if not 1 <= mon <= 12:
        return None
    return date(y, mon, calendar.monthrange(y, mon)[1])


def parse_workbook(path: str, fleet: List[Dict]) -> Dict[str, List[Dict]]:
    """{'rows': [{month_end, kind, card, plate, amount, label}], 'unmatched': [{month, kind, card, label, amount}]}"""
    import openpyxl
    from dashboard.services import vehicle_manage as vm
    plates = {c['plate'] for c in fleet}
    out = {'rows': [], 'unmatched': []}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            end = _month_of(ws.title)
            if not end:
                continue
            kind = None
            for row in ws.iter_rows(min_row=1, max_col=7, values_only=True):
                row = list(row) + [None] * (7 - len(row))
                b = str(row[1] or '').strip()
                if b:
                    kind = SECTIONS.get(b)        # 새 구획 시작이거나(유류비·하이패스) 끝(합계·신용카드·종류…)
                    if not kind:
                        continue
                if not kind:
                    continue
                card, who, model, amt = (str(row[3] or '').strip(), str(row[4] or '').strip(),
                                         str(row[5] or '').strip(), row[6])
                if not isinstance(amt, (int, float)) or amt <= 0:
                    continue
                label = ' '.join(x for x in (who, model) if x)
                found = vm.find_plates(who, fleet) if who else []
                plate = found[0] if len(found) == 1 else PERSON_PLATE.get(who)
                person = ''
                if plate not in plates:
                    digits = re.sub(r'\D', '', who)
                    sold = next((p for p in SOLD if digits and p.endswith(digits[-4:])), None)
                    if sold:
                        plate = sold
                    else:
                        plate, person = UNASSIGNED, PERSON_ALIAS.get(who, who) or '(빈칸)'
                    out['unmatched'].append({'month': end.strftime('%Y-%m'), 'kind': kind, 'card': card,
                                             'label': label or '(빈칸)', 'amount': int(amt), 'plate': plate})
                out['rows'].append({'month_end': end.isoformat(), 'kind': kind, 'card': card, 'plate': plate,
                                    'amount': int(round(amt)), 'label': label, 'person': person})
    finally:
        wb.close()
    return out


def save_rows(rows: List[Dict]) -> Dict[str, int]:
    """차량 관리 비용에 넣기 — 같은 달·종류·카드·차는 한 줄(source_key). 금액이 바뀌면 고침."""
    from dashboard.services import vehicle_manage as vm
    added = changed = 0
    now = vm._now()
    with vm._conn() as conn:
        for r in rows:
            m = int(r['month_end'][5:7])
            name = '유류비' if r['kind'] == '유류' else '하이패스 통행료'
            item = f'{m}월 {name} (월 합계)' + (f" · {r['person']} 카드" if r.get('person') else '')
            vendor = f"법인카드 {r['card']}"[:40]
            key = f"card:{r['month_end'][:7]}:{r['kind']}:{r['card']}:{r['plate']}" + (f":{r['person']}" if r.get('person') else '')
            cur = conn.execute('SELECT id, amount FROM vehicle_costs WHERE source_key = ?', (key,)).fetchone()
            if cur is None:
                conn.execute('INSERT INTO vehicle_costs (plate, day, kind, item, amount, shared, vendor, source, source_key, '
                             'created_by, created_at) VALUES (?,?,?,?,?,1,?,?,?,?,?)',
                             (r['plate'], r['month_end'], r['kind'], item, r['amount'], vendor, SOURCE, key, 'sync', now))
                added += 1
            elif int(cur['amount']) != r['amount']:
                conn.execute('UPDATE vehicle_costs SET amount = ?, item = ? WHERE id = ?', (r['amount'], item, cur['id']))
                changed += 1
    return {'added': added, 'changed': changed}


def _folder(svc) -> str:
    drives = svc.drives().list(pageSize=100).execute().get('drives', [])
    drive = next((d for d in drives if d.get('name') == DRIVE_NAME), None)
    if not drive:
        raise RuntimeError(f"서비스 계정이 공유 드라이브 '{DRIVE_NAME}' 를 못 봄")
    parent = drive['id']
    for name in FOLDER_PATH:
        q = (f"name = '{name}' and '{parent}' in parents and mimeType = 'application/vnd.google-apps.folder' "
             "and trashed = false")
        found = svc.files().list(q=q, fields='files(id)', corpora='drive', driveId=drive['id'], supportsAllDrives=True,
                                 includeItemsFromAllDrives=True).execute().get('files', [])
        if not found:
            raise RuntimeError(f"공유 드라이브에 '{'/'.join(FOLDER_PATH)}' 폴더가 없음")
        parent = found[0]['id']
    return parent


def _year_files(svc, folder: str, years: List[int]) -> List[dict]:
    r = svc.files().list(q=f"'{folder}' in parents and trashed = false", pageSize=200,
                         fields='files(id, name, modifiedTime)', supportsAllDrives=True,
                         includeItemsFromAllDrives=True).execute()
    out = []
    for f in r.get('files', []):
        m = FILE_RE.match(f['name'])
        if m and int(m.group(1)) in years:
            out.append(f)
    return out


def run(today: Optional[date] = None, path: Optional[str] = None) -> Dict:
    """올해(1~2월엔 지난해 파일도 — 12월 청구가 1월에 채워짐) SB 파일을 읽어 차량 관리 비용에 넣는다.
    path 를 주면 드라이브 대신 그 파일(시험·손으로 채우기)."""
    from dashboard.services import vehicle_manage as vm
    today = today or date.today()
    fleet = vm._fleet()
    paths: List[str] = []
    tmp = None
    if path:
        paths = [path]
    else:
        from dashboard.services import fund_daily as fd
        svc = fd._drive()
        years = [today.year] + ([today.year - 1] if today.month <= 2 else [])
        files = _year_files(svc, _folder(svc), years)
        if not files:
            return {'status': 'no_file', 'years': years}
        tmp = tempfile.mkdtemp(prefix='card_costs_')
        for f in files:
            p = os.path.join(tmp, f"{f['id']}.xlsx")
            fd._download(svc, f['id'], p)
            paths.append(p)
    res = {'status': 'ok', 'added': 0, 'changed': 0, 'unmatched': []}
    try:
        for p in paths:
            parsed = parse_workbook(p, fleet)
            saved = save_rows(parsed['rows'])
            res['added'] += saved['added']
            res['changed'] += saved['changed']
            res['unmatched'] += parsed['unmatched']
    finally:
        if tmp:
            import shutil
            shutil.rmtree(tmp, ignore_errors=True)
    if res['added'] or res['changed']:
        logger.info(f"[차량관리] 법인카드 지출결의서에서 유류·통행료 새로 {res['added']}줄, 고침 {res['changed']}줄")
    return res
