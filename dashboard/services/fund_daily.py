"""자금일보 자동 초안(2026-11-26 부장님 — SB 가 매일 손으로 채우던 자금일보 엑셀을 서버가 미리 채워 둔다).

흐름
- 매일 07:00(SB 08:00 출근 전) 전 영업일 시트를 그 달 자금일보 파일(드라이브 경영지원관리Ⅱ/자금관련/자금일보/'YY.MM 자금일보.xlsx')에
  새로 만든다 — 앞 시트를 본떠서(서식·합친 칸·조건부 서식), 위쪽 계좌 표의 전일잔액 수식은 앞 시트를 가리키게.
- 상세 내역(19~127행)은 입출금 장부(services/bank_ledger, 은행 문자)에서 채운다:
  · 입금: 공사 코드가 붙은 줄(입금 지정 확인) → 계정 = 그 공사 유입 구분, 적요 = '{코드} 매출수금완료'(미수금 0) / '매출수금 중'
  · 출금: 전자결재 이체 완료와 금액으로 짝지음 → 계정(자재비·도급비 = 외상매입), 적요 '{공사 코드 또는 0 이니셜} {내용}', 상대방 = 거래처
  · 그 밖: 지난 자금일보에서 같은 상대방에 쓴 계정·적요를 제안. 못 찾으면 비워 두고 칸에 메모('확인 필요')
- 그날 시트가 이미 있으면(SB 가 만들었으면) 건드리지 않는다. 새 달 파일이 없으면 만들지 않고 알린다(SB 가 만든 뒤 다음 날).
- 문자가 안 오는 계좌(우리·수협·국민·P 기업·G 하나·R 하나 9011)는 SB 가 그 시트에 직접 넣는다.

설정(.env)
- FUND_DAILY_MODE  off(기본) / copy(시험 — 사본 폴더에만) / live(SB 파일에 직접)
- FUND_DAILY_FOLDER_ID  자금일보 드라이브 폴더 ID(없으면 공유 드라이브 경영지원관리Ⅱ/자금관련/자금일보 를 찾음)
  ※ 서버(Flask)는 LocalSystem 이라 G: 가 안 보여 서비스 계정 Drive API 로 받고 올린다(외부 백업과 같은 방식).
- FUND_DAILY_COPY_DIR  copy 모드에서 쓸 폴더(기본 바탕화면\\ITG-Project\\자금일보_자동초안_시험)
"""
from __future__ import annotations

import os
import re
import shutil
from collections import Counter, defaultdict
from copy import copy
from datetime import date, datetime, timedelta
from typing import Dict, List, Optional, Tuple

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)
_LOG = '[자금일보]'

FIRST_ROW, LAST_ROW = 19, 127          # 상세 내역 줄
ACCT_ROWS = range(6, 14)              # 위쪽 계좌 표(전일잔액 G = 앞 시트 L)
IN_COLS = ('B', 'C', 'D', 'G', 'H')   # 구분 · 계정 · 적요 · 입금자명 · 금액
OUT_COLS = ('I', 'J', 'K', 'N', 'O')
SUPPLIER_COSTS = ('자재비', '도급비', '외주비', '장비비', '장비 임차', '운반비')   # 자금일보 계정 '외상매입'
PROJECT_RE = re.compile(r'^[A-Z]\d{3,5}-[A-Z]{1,3}$')
CHECK_NOTE = '그룹웨어: 계정·적요를 찾지 못했어요. 확인해 주세요.'


def mode() -> str:
    m = (os.getenv('FUND_DAILY_MODE') or 'off').strip().lower()
    return m if m in ('off', 'copy', 'live') else 'off'


def fund_dir() -> str:
    return os.getenv('FUND_DAILY_DIR') or r'G:\공유 드라이브\경영지원관리Ⅱ\자금관련\자금일보'


def copy_dir() -> str:
    return os.getenv('FUND_DAILY_COPY_DIR') or r'C:\Users\SECOM\Desktop\ITG-Project\자금일보_자동초안_시험'


def month_file(d: date) -> str:
    return f'{d:%y.%m} 자금일보.xlsx'


# ── 지난 자금일보에서 배우기 ───────────────────────────────────────────
def norm(s) -> str:
    s = str(s or '')
    s = re.sub(r'\(주\)|㈜|주식회사|\(유\)|유한회사', '', s)
    s = re.sub(r'[\s()\[\]·.,:/-]', '', s)
    return s.lower()


_hint_cache: Dict[str, object] = {}


def learn_hints(folder: Optional[str] = None) -> Dict[str, Dict[str, Counter]]:
    """{'in': {상대방: Counter((계정, 적요))}, 'out': {...}} — 그 폴더의 모든 월 파일(읽기만). 하루 한 번 다시 읽음."""
    folder = folder or fund_dir()
    key = f'{folder}|{date.today().isoformat()}'
    if _hint_cache.get('key') == key:
        return _hint_cache['val']           # type: ignore[return-value]
    import openpyxl
    hints: Dict[str, Dict[str, Counter]] = {'in': defaultdict(Counter), 'out': defaultdict(Counter)}
    try:
        names = sorted(f for f in os.listdir(folder) if f.endswith('자금일보.xlsx') and not f.startswith('~$'))
    except OSError as e:
        logger.warning(f'{_LOG} 지난 자금일보 폴더를 못 읽음: {e}')
        names = []
    for f in names:
        try:
            wb = openpyxl.load_workbook(os.path.join(folder, f), data_only=True, read_only=True)
        except Exception as e:
            logger.warning(f'{_LOG} {f} 읽기 실패: {e}')
            continue
        for ws in wb.worksheets:
            for r in ws.iter_rows(min_row=FIRST_ROW, max_row=LAST_ROW, min_col=2, max_col=15, values_only=True):
                b, c, d, _, _, g, h, i, j, k, _, _, n, o = r
                if h and b and g and c:
                    hints['in'][norm(g)][(str(c), str(d or ''))] += 1
                if o and i and n and j:
                    hints['out'][norm(n)][(str(j), str(k or ''))] += 1
        wb.close()
    _hint_cache.update(key=key, val=hints)
    return hints


def _hint(hints: Dict[str, Counter], name: str) -> Tuple[str, str]:
    """가장 많이 쓴 (계정, 적요). 적요가 공사마다 다른 상대방이면(공사 코드가 든 적요) 계정만."""
    for cand in [name] + [p for p in re.split(r'[:/]', name or '') if p]:
        c = hints.get(norm(cand))
        if c:
            (title, memo), _n = c.most_common(1)[0]
            titles = Counter()
            for (t, _m), k in c.items():
                titles[t] += k
            title = titles.most_common(1)[0][0]
            same_memo = sum(k for (t, m), k in c.items() if m == memo)
            if re.search(r'[A-Z]\d{3,5}-', memo) or same_memo < 2:
                memo = ''
            return title, memo
    return '', ''


# ── 상세 내역 만들기 ──────────────────────────────────────────────────
def _short(name: str) -> str:
    return re.sub(r'^\s*(\(주\)|㈜)\s*', '', name or '').strip()


def deposit_row(tx: dict, project: Optional[dict], hints) -> dict:
    """입금 한 줄 {B,C,D,G,H, check}. project = 공사 캐시 행(유입 구분·미수금)."""
    codes = [c for c in (tx.get('project_code') or '').split(',') if c]
    title, memo, check = '', '', False
    if codes:
        title = str((project or {}).get('유입 구분') or '').strip()
        parts = []
        for c in codes:
            unpaid = (project or {}).get('미수금') if c == codes[0] else None
            try:
                done = unpaid is not None and float(unpaid) <= 0
            except (TypeError, ValueError):
                done = False
            parts.append(f'{c} 매출수금완료' if done else f'{c} 매출수금 중')
        memo = ' / '.join(parts)
        if not title:
            title = _hint(hints, tx.get('counterparty') or '')[0]
    else:
        title, memo = _hint(hints, tx.get('counterparty') or '')
        check = not title
    return {'B': tx['code'], 'C': title, 'D': memo, 'G': _short(tx.get('counterparty') or ''), 'H': tx['amount'],
            'check': check or not title}


def outflow_row(tx: dict, approval: Optional[dict], hints, initial_of) -> dict:
    """출금 한 줄 {I,J,K,N,O, check}. approval = 짝지어진 전자결재(이체 완료)."""
    if approval:
        cost = approval.get('cost_category') or ''
        pc = (approval.get('project_code') or '').strip()
        head = pc if PROJECT_RE.match(pc) else f"0 {initial_of(approval.get('requester_name') or '')}".strip()
        partner = approval.get('partner') or approval.get('payee_name') or tx.get('counterparty') or ''
        title = '외상매입' if cost in SUPPLIER_COSTS else (_hint(hints, partner)[0] or cost)
        return {'I': tx['code'], 'J': title, 'K': f"{head} {approval.get('content') or ''}".strip(), 'N': partner,
                'O': tx['amount'], 'check': not title}
    title, memo = _hint(hints, tx.get('counterparty') or '')
    return {'I': tx['code'], 'J': title, 'K': memo, 'N': tx.get('counterparty') or '', 'O': tx['amount'],
            'check': not title}


# 자금일보 '구분' → 사업자(전자결재 grn = 이체할 통장 G·R·N 과 맞추기)
CODE_ENTITY = {'1011': 'G', '1012': 'G', '1013': 'G', '1014': 'G', '9011': 'R', '9012': 'R', '4011': 'P', '5011': 'N',
               '농협': 'N', '하도급': 'R'}


def match_approvals(outs: List[dict], approvals: List[dict]) -> Dict[int, dict]:
    """출금 ↔ 전자결재 이체 완료 — 같은 금액, 이체 시각이 문자 시각 ±1일, 이체 통장(grn)이 있으면 같은 사업자.
    한 결재는 한 번만. {출금 index: 결재}."""
    used, out = set(), {}
    for i, tx in enumerate(outs):
        ent = CODE_ENTITY.get(str(tx.get('code') or ''), '')
        try:
            t = datetime.strptime(tx['at'][:16], '%Y-%m-%d %H:%M')
        except (KeyError, ValueError):
            continue
        best = None
        for a in approvals:
            if a['id'] in used or int(a.get('amount') or 0) != int(tx['amount']):
                continue
            grn = str(a.get('grn') or '')[:1].upper()
            if grn and ent and grn != ent:
                continue
            try:
                at = datetime.strptime(str(a['done_at'])[:16], '%Y-%m-%d %H:%M')
            except ValueError:
                continue
            gap = abs((at - t).total_seconds())
            if gap <= 86400 and (best is None or gap < best[0]):
                best = (gap, a)
        if best:
            used.add(best[1]['id'])
            out[i] = best[1]
    return out


def transfers_done(start: str, end: str) -> List[dict]:
    """그 기간 이체 완료된 전자결재(자금집행 등)."""
    from dashboard.utils.user_database import get_user_database
    with get_user_database()._get_connection() as c:
        return [dict(r) for r in c.execute(
            "SELECT a.id, a.form, a.amount, a.project_code, a.cost_category, a.content, a.partner, a.payee_name, "
            "a.requester_name, a.grn, s.at AS done_at FROM approvals a JOIN approval_steps s ON s.approval_id=a.id "
            "WHERE s.step='transfer' AND s.result='done' AND s.at >= ? AND s.at < ?", (start, end))]


def build_rows(txs: List[dict], hints=None, projects: Optional[Dict[str, dict]] = None,
               approvals: Optional[List[dict]] = None, initial_of=None) -> Tuple[List[dict], List[dict]]:
    """거래 [{code, direction, amount, counterparty, project_code, at}] → (입금 줄, 출금 줄). 시각순."""
    hints = hints if hints is not None else {'in': {}, 'out': {}}
    projects = projects or {}
    initial_of = initial_of or (lambda n: '')
    ins = [t for t in sorted(txs, key=lambda x: x.get('at') or '') if t['direction'] == 'in']
    outs = [t for t in sorted(txs, key=lambda x: x.get('at') or '') if t['direction'] == 'out']
    pairs = match_approvals(outs, approvals or [])
    # 농협(N통장) 입금은 공사 입금으로 지정된 것만 — N통장은 대표님 개인 용도도 섞여 있음(11-26 부장님)
    in_rows = [deposit_row(t, projects.get((t.get('project_code') or '').split(',')[0]), hints['in']) for t in ins
               if not (t.get('code') == '농협' and not t.get('project_code'))]
    out_rows = [outflow_row(t, pairs.get(i), hints['out'], initial_of) for i, t in enumerate(outs)
                if not _skip_outflow(t, pairs.get(i))]
    return in_rows, out_rows


def _skip_outflow(tx: dict, approval: Optional[dict]) -> bool:
    """자금일보에 안 넣는 출금 — 농협(N통장) 출금은 전자결재 N 출금(grn N)과 짝지어질 때만 넣는다.
    11-26 부장님 '농협은 전자결재에 N출금이 있을 때만' — 나머지는 대표님(YG) 개인 용도(11/25 1,000,000원 ×5).
    장부·입출금 화면에는 그대로 남는다."""
    if tx.get('code') != '농협':
        return False
    return not (approval and str(approval.get('grn') or '')[:1].upper() == 'N')


# ── 시트 쓰기 ─────────────────────────────────────────────────────────
def add_day_sheet(path_in: str, path_out: str, day: date, in_rows: List[dict], out_rows: List[dict],
                  replace: bool = False, suffix: str = '') -> str:
    """그 달 파일에 그날 시트 추가 → path_out 저장. 결과 'made' / 'exists' / 'no_template'.
    replace=True 는 시험(사본)에서만 — 같은 이름 시트를 지우고 다시 만든다.
    suffix — 시험(사본)에서 SB 가 만든 같은 날 시트와 나란히 두려고 '26 자동' 처럼 이름을 다르게."""
    import openpyxl
    from openpyxl.comments import Comment
    wb = openpyxl.load_workbook(path_in)
    name = str(day.day) + suffix
    if name in wb.sheetnames:
        if not replace:
            return 'exists'
        del wb[name]
    prevs = [ws for ws in wb.worksheets if ws.title.isdigit() and int(ws.title) < day.day]
    if not prevs:
        return 'no_template'
    prev = max(prevs, key=lambda w: int(w.title))
    ws = wb.copy_worksheet(prev)
    ws.title = name
    # 시트 순서: 앞 시트 바로 뒤
    wb.move_sheet(ws, offset=wb.sheetnames.index(prev.title) + 1 - wb.sheetnames.index(name))
    # copy_worksheet 가 옮기지 않는 조건부 서식·열 너비 보정
    for cf in prev.conditional_formatting:
        for rule in cf.rules:
            ws.conditional_formatting.add(str(cf.sqref), rule)
    for k, dim in prev.column_dimensions.items():
        ws.column_dimensions[k].width = dim.width
    ws.sheet_view.tabSelected = False
    for r in ACCT_ROWS:
        ws[f'G{r}'] = f"=SUM('{prev.title}'!L{r}:M{r})"
    # 상세 내역 비우기(값만 — 서식·합친 칸은 그대로)
    for row in ws.iter_rows(min_row=FIRST_ROW, max_row=LAST_ROW, min_col=2, max_col=15):
        for c in row:
            if type(c).__name__ == 'MergedCell':
                continue
            if c.value is not None and not (isinstance(c.value, str) and c.value.startswith('=')):
                c.value = None
            c.comment = None
    cap = LAST_ROW - FIRST_ROW + 1
    for cols, rows in ((IN_COLS, in_rows), (OUT_COLS, out_rows)):
        for i, row in enumerate(rows[:cap]):
            r = FIRST_ROW + i
            for col in cols:
                v = row.get(col)
                if v not in (None, ''):
                    ws[f'{col}{r}'] = int(v) if col in ('H', 'O') else (int(v) if str(v).isdigit() else v)
            if row.get('check'):
                ws[f'{cols[1]}{r}'].comment = Comment(CHECK_NOTE, '그룹웨어')
    # SB 처럼 쓰는 줄만 보이고 나머지는 숨김(앞 시트의 숨긴 줄이 그대로 따라와 채운 줄이 가려졌음) — 빈 줄 2개는 보이게(손 입력용)
    used = max(len(in_rows), len(out_rows))
    for r in range(FIRST_ROW, LAST_ROW + 1):
        ws.row_dimensions[r].hidden = r >= FIRST_ROW + used + 2
    if len(in_rows) > cap or len(out_rows) > cap:
        logger.error(f'{_LOG} {day} 상세 내역이 {cap}줄을 넘어 일부를 못 넣음 — 입금 {len(in_rows)} 출금 {len(out_rows)}')
    wb.active = wb.sheetnames.index(name)
    wb.save(path_out)
    return 'made'


# ── 매일 ─────────────────────────────────────────────────────────────
def prev_business_day(today: date) -> date:
    from dashboard.services import company_holidays as ch
    d = today - timedelta(days=1)
    while d.weekday() >= 5 or ch.is_holiday(d):
        d -= timedelta(days=1)
    return d


def ledger_txs(start: str, end: str) -> List[dict]:
    """장부의 입금·출금(잔액 맞추기 빼고) — [start, end) 기간."""
    from dashboard.services import bank_ledger as bl
    rows = bl.list_tx(start, (date.fromisoformat(end) - timedelta(days=1)).isoformat())
    return [{'code': r['acct_code'] or '', 'direction': r['direction'], 'amount': r['amount'],
             'counterparty': r['counterparty'] or '', 'project_code': r['project_code'] or '', 'at': r['at']}
            for r in rows if r['direction'] in ('in', 'out') and r['acct_code']]


# ── 드라이브(서버는 LocalSystem 이라 G: 가 안 보임 → 서비스 계정으로 Drive API, 외부 백업과 같은 방식) ─────
DRIVE_NAME = '경영지원관리Ⅱ'
FOLDER_PATH = ('자금관련', '자금일보')


def _drive():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    from pathlib import Path
    cred = (os.getenv('GOOGLE_CREDENTIALS_FILE') or os.getenv('GOOGLE_APPLICATION_CREDENTIALS') or 'credentials.json').strip()
    if not os.path.isabs(cred):
        cred = str(Path(__file__).resolve().parents[2] / cred)
    creds = Credentials.from_service_account_file(cred, scopes=['https://www.googleapis.com/auth/drive'])
    return build('drive', 'v3', credentials=creds, cache_discovery=False)


def _folder_id(svc) -> str:
    fid = (os.getenv('FUND_DAILY_FOLDER_ID') or '').strip()
    if fid:
        return fid
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


def _files(svc, folder_id: str) -> Dict[str, dict]:
    """{이름: {id, modifiedTime}} — 자금일보 월 파일만."""
    out, token = {}, None
    while True:
        r = svc.files().list(q=f"'{folder_id}' in parents and trashed = false", pageSize=200, pageToken=token,
                             fields='nextPageToken, files(id, name, modifiedTime)', supportsAllDrives=True,
                             includeItemsFromAllDrives=True).execute()
        for f in r.get('files', []):
            if f['name'].endswith('자금일보.xlsx'):
                out[f['name']] = f
        token = r.get('nextPageToken')
        if not token:
            return out


def _download(svc, file_id: str, path: str) -> None:
    import io
    from googleapiclient.http import MediaIoBaseDownload
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, svc.files().get_media(fileId=file_id, supportsAllDrives=True))
    done = False
    while not done:
        _st, done = dl.next_chunk()
    with open(path, 'wb') as f:
        f.write(buf.getvalue())


def _upload(svc, file_id: str, path: str) -> None:
    from googleapiclient.http import MediaFileUpload
    media = MediaFileUpload(path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    svc.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()


def work_dir() -> str:
    from pathlib import Path
    d = os.getenv('FUND_DAILY_WORK_DIR') or str(Path(__file__).resolve().parents[2] / 'instance' / 'fund_daily')
    os.makedirs(d, exist_ok=True)
    return d


def run(today: Optional[date] = None, force_mode: str = '', svc=None) -> dict:
    """07:00 — 전 영업일 시트 만들기. 반환 {'status', 'day', ...}.
    copy: 그 달 파일을 받아 'N 자동' 시트를 넣고 사본 폴더에 저장(드라이브엔 안 올림).
    live: 받기 전 사본을 instance/fund_daily/backup 에 남기고, 고친 파일을 같은 파일의 새 버전으로 올림.
          받은 뒤 그 사이 누가 고쳤으면(수정 시각이 바뀜) 올리지 않는다."""
    m = force_mode or mode()
    if m == 'off':
        return {'status': 'off'}
    today = today or date.today()
    day = prev_business_day(today)
    svc = svc or _drive()
    folder = _folder_id(svc)
    files = _files(svc, folder)
    meta = files.get(month_file(day))
    if not meta:
        logger.warning(f'{_LOG} {month_file(day)} 이 없어 {day} 시트를 못 만듦 — 새 달 파일은 SB 가 처음 만든 뒤부터')
        return {'status': 'no_month_file', 'day': day.isoformat()}
    wd = work_dir()
    hint_dir = os.path.join(wd, 'hints')
    os.makedirs(hint_dir, exist_ok=True)
    for name, f in files.items():              # 지난 자금일보(계정·적요 배우기) — 바뀐 파일만 다시 받음
        local = os.path.join(hint_dir, name)
        stamp = local + '.mtime'
        if not (os.path.exists(local) and os.path.exists(stamp) and open(stamp).read() == f['modifiedTime']):
            _download(svc, f['id'], local)
            open(stamp, 'w').write(f['modifiedTime'])
    src = os.path.join(hint_dir, month_file(day))
    if m == 'copy':
        os.makedirs(copy_dir(), exist_ok=True)
        path = os.path.join(copy_dir(), month_file(day).replace('.xlsx', ' (자동 초안).xlsx'))
    else:
        os.makedirs(os.path.join(wd, 'backup'), exist_ok=True)
        shutil.copy2(src, os.path.join(wd, 'backup', f'{datetime.now():%Y%m%d-%H%M} {month_file(day)}'))
        path = os.path.join(wd, 'upload.xlsx')
    shutil.copy2(src, path)
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    prev_days = sorted(int(n) for n in wb.sheetnames if n.isdigit() and int(n) < day.day)
    wb.close()
    start = day.replace(day=prev_days[-1]) + timedelta(days=1) if prev_days else day.replace(day=1)
    end = day + timedelta(days=1)
    txs = ledger_txs(start.isoformat(), end.isoformat())
    from dashboard.services.payment_sync import _project_cache_row
    from dashboard.services.e_approval import initial
    projects = {}
    for t in txs:
        c = (t['project_code'] or '').split(',')[0]
        if c and c not in projects:
            projects[c] = _project_cache_row(c) or {}
    ins, outs = build_rows(txs, learn_hints(hint_dir), projects,
                           transfers_done((start - timedelta(days=1)).isoformat(), (end + timedelta(days=1)).isoformat()),
                           initial)
    res = add_day_sheet(path, path, day, ins, outs, suffix=' 자동' if m == 'copy' else '')
    if res == 'made' and m == 'live':
        now_meta = svc.files().get(fileId=meta['id'], fields='modifiedTime', supportsAllDrives=True).execute()
        if now_meta.get('modifiedTime') != meta['modifiedTime']:
            logger.error(f'{_LOG} {month_file(day)} 이 받는 사이 바뀌어 올리지 않음 — 다시 돌리면 됨', extra={'force_alert': True})
            return {'status': 'changed_meanwhile', 'day': day.isoformat()}
        _upload(svc, meta['id'], path)
    logger.info(f'{_LOG} {m} {day} 시트 {res} — 입금 {len(ins)} 출금 {len(outs)} (확인 필요 '
                f'{sum(1 for r in ins + outs if r.get("check"))})')
    return {'status': res, 'day': day.isoformat(), 'mode': m, 'in': len(ins), 'out': len(outs), 'path': path}
