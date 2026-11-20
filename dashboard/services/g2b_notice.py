"""입찰 공고문 자동 읽기 (2026-11-20 부장님 3단계).

추산초처럼 신용등급 통과 기준·제조사 확약서·현장설명 필수 같은 조건은 API 칸에 없고 첨부 공고문에만 있다.
→ 나라장터 첨부 중 '공고문'을 받아 kordoc(doc_viewer)으로 글로 바꾸고,
   ① '입찰참가자격'·'현장설명' 단락 원문 ② 눈여겨볼 조건 문구(확약서·직접생산·실적·신용등급…)를 저장해 상세에 보여 준다.

- 자동: 모으기 뒤 입찰 가능한 공고(마감 전) 중 아직 안 읽은 것만 몇 건씩. 그 밖은 상세의 [공고문 읽기] 버튼.
- 파일·변환 실패는 줄에 사유만 남기고(하루 뒤 다시 시도) 목록은 그대로 보인다.
- 문구 찾기는 '있다'만 알려 준다 — 판정은 사람이(원문 줄을 함께 보여 줌). 신용등급 'OO 이상'만 우리 등급과 비교.
"""
import json
import re
import threading
from datetime import datetime, timedelta
from typing import Dict, List, Optional

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_MAX_BYTES = 20 * 1024 * 1024
_SECTION_CHARS = 3500          # 단락 하나 상한(글자)
_LINE_CHARS = 400
_run_lock = threading.Lock()

DDL = '''CREATE TABLE IF NOT EXISTS g2b_notice (
    bid_no TEXT PRIMARY KEY,
    status TEXT NOT NULL DEFAULT '',      -- ok / fail / none(공고문 파일 없음)
    message TEXT NOT NULL DEFAULT '',
    file_name TEXT NOT NULL DEFAULT '',
    sections TEXT NOT NULL DEFAULT '[]',
    flags TEXT NOT NULL DEFAULT '[]',
    read_at TEXT NOT NULL DEFAULT ''
)'''


# ── 1) 첨부 중 공고문 고르기 ──
_NOT_NOTICE = ('내역', '시방', '도면', '규격서', '설계', '산출', '특수조건', '청렴', '서약', '양식', '서식', '현황표')
_EXT_RANK = {'hwpx': 0, 'hwp': 1, 'pdf': 2}   # 한글 원본이 표·줄바꿈이 더 깨끗하게 나옴


def pick_notice_file(files: List[dict]) -> Optional[dict]:
    cands = []
    for i, f in enumerate(files or []):
        name = str(f.get('name') or '')
        ext = name.rsplit('.', 1)[-1].lower() if '.' in name else ''
        if ext not in _EXT_RANK or not f.get('url'):
            continue
        stem = name.rsplit('.', 1)[0]
        if '공고' not in stem or any(w in stem for w in _NOT_NOTICE):
            continue
        cands.append((_EXT_RANK[ext], i, f))
    return min(cands, key=lambda x: (x[0], x[1]))[2] if cands else None


# ── 2) 글 다듬기·단락 자르기 ──
def _clean(line: str) -> str:
    s = re.sub(r'<br\s*/?>', ' ', line, flags=re.I)
    s = re.sub(r'<[^>]+>', '', s)
    s = s.replace('**', '').replace('\\*', '*').replace('\\_', '_')
    s = re.sub(r'!\[[^\]]*\]\([^)]*\)', '', s)
    if s.strip().startswith('|'):
        if re.match(r'^\|?\s*:?-{3,}', s.strip()):
            return ''
        s = ' · '.join(c.strip() for c in s.strip().strip('|').split('|') if c.strip())
    s = re.sub(r'\s+', ' ', s).strip()
    if re.match(r'^[\s*\-=_~─━═·.]{5,}$', s):
        return ''
    return s


# 큰 제목: '3. 입찰참가자격', '3 입찰참가자격', 'Ⅲ. …', '제3조 …' 처럼 짧은 줄(3.1. 같은 하위 번호는 아님)
_HEAD_RE = re.compile(r'^(?:(\d{1,2})\s*[.)]?|([ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ])\s*\.?)\s+(\S.{0,28})$')


def _heads(lines: List[str]) -> List[tuple]:
    out = []
    for i, ln in enumerate(lines):
        m = _HEAD_RE.match(ln)
        if m and not re.search(r'[다요]\.?$', m.group(3)):   # 문장(…합니다.)은 제목 아님
            out.append((i, m.group(3).strip()))
    return out


_WANT = (('참가자격', '입찰참가자격'), ('현장설명', '현장설명'))


def sections_of(text: str) -> List[dict]:
    lines = [c for c in (_clean(x) for x in text.splitlines()) if c]
    heads = _heads(lines)
    out = []
    for key, title in _WANT:
        for n, (i, h) in enumerate(heads):
            # 앞쪽 목차·근거 법령 줄('Ⅲ. …입찰참가자격등록규정', '…등록증 등록정보 확인')은 건너뜀
            if key not in h.replace(' ', '') or re.search(r'규정|등록증|기준$', h):
                continue
            if n + 1 < len(heads) and heads[n + 1][0] == i + 1:
                continue   # 바로 다음 줄이 또 제목 = 목차
            end = heads[n + 1][0] if n + 1 < len(heads) else min(len(lines), i + 40)
            body, size = [], 0
            for ln in lines[i + 1:end]:
                ln = ln[:_LINE_CHARS]
                if size + len(ln) > _SECTION_CHARS:
                    body.append('… (이하 공고문 참조)')
                    break
                body.append(ln)
                size += len(ln)
            if body:
                out.append({'title': title, 'lines': body})
            break
    return out


# ── 3) 눈여겨볼 조건 문구 ──
#   level: warn(참가·낙찰에 걸릴 수 있음) / info(알아 둘 것)
_RULES = [
    ('credit', 'warn', '신용평가등급 기준', re.compile(r'신용\s*(평가\s*)?등급')),
    ('confirm', 'warn', '제조사·대리점 확약서', re.compile(r'(제조|대리점|공급|A/?S|기술지원).{0,30}확약서|확약서.{0,20}(제조|대리점)')),
    ('direct', 'warn', '직접생산 확인', re.compile(r'직접\s*생산')),
    ('perf', 'warn', '실적 조건', re.compile(r'(이행|납품|시공|공사|수행|계약)\s*실적')),
    ('sme', 'info', '중소기업·소기업 제한', re.compile(r'(중소기업자?\s*(간\s*경쟁|만|로\s*제한|에\s*한)|소기업.{0,12}(제한|만\s*참)|소상공인.{0,12}(제한|만\s*참))')),
    ('brief_must', 'warn', '현장설명 참석 필수', re.compile(r'현장\s*설명.{0,40}(참석하지\s*않은|참석한\s*자|참석\s*(필수|의무)|미참석)')),
    ('brief_none', 'info', '현장설명 없음', re.compile(r'현장\s*설명을?\s*실시하지\s*않')),
    ('region', 'warn', '소재지 제한 문구', re.compile(r'(주된\s*영업소|본점|본사|법인\s*등기부상\s*소재지).{0,50}(소재|두고|있는\s*자|있는\s*업체)')),
    ('ks', 'info', 'KS 인증 제품', re.compile(r'KS\s*[A-Z]?\s*\d{3,4}|KS\s*(제품\s*)?인증')),
    ('energy', 'info', '에너지효율 1등급', re.compile(r'효율\s*1\s*등급')),
    ('nojoint', 'info', '공동도급 불가', re.compile(r'공동\s*도급.{0,20}(허용되지\s*않|불가|허용하지\s*않)')),
]
_GRADE_RE = re.compile(r'(AAA|AA[+\-0°]?|A[+\-0°]?|BBB[+\-0°]?|BB[+\-0°]?|B[+\-0°]?|CCC[+\-0°]?)\s*(?:등급\s*)?이상')


def _grade_norm(g: str) -> str:
    g = g.replace('°', '0')
    return g + '0' if g in ('AA', 'A', 'BBB', 'BB', 'B') else g


def credit_check(line: str, ours: str) -> Optional[bool]:
    """'BB- 이상' 같은 기준이 줄에 있으면 우리 등급이 넘는지(True/False), 판단 못 하면 None."""
    from dashboard.services.g2b_bids import CREDIT_GRADES
    m = _GRADE_RE.search(line or '')
    if not m or not ours or ours not in CREDIT_GRADES:
        return None
    need = _grade_norm(m.group(1))
    if need not in CREDIT_GRADES:
        return None
    return CREDIT_GRADES.index(ours) <= CREDIT_GRADES.index(need)


# 같은 조건 문구가 여러 줄이면 실제 기준이 적힌 줄을 먼저(일반 안내·근거 법령 줄보다)
_PREFER = {
    'credit': _GRADE_RE,
    'perf': re.compile(r'실적이\s*있는|실적.{0,15}(이상|보유|있는\s*(자|업체))'),
}
# 제재·위반 사유 목록 줄은 조건이 아님(부정당업자 사유 '직접생산기준을 위반하여 납품하는 행위' 등)
_SKIP = {
    'direct': re.compile(r'위반|하청생산'),
    'perf': re.compile(r'적용기준일|실적\s*증명서?\s*(발급|제출\s*안내)'),
}


def flags_of(text: str) -> List[dict]:
    lines = [c for c in (_clean(x) for x in text.splitlines()) if c]
    out = []
    for key, level, label, rx in _RULES:
        hits = [ln for ln in lines if rx.search(ln) and not (key in _SKIP and _SKIP[key].search(ln))]
        if not hits:
            continue
        pref = _PREFER.get(key)
        best = next((ln for ln in hits if pref and pref.search(ln)), hits[0])
        out.append({'key': key, 'level': level, 'label': label, 'line': best[:320]})
    if any(f['key'] == 'brief_must' for f in out):
        out = [f for f in out if f['key'] != 'brief_none']
    return out


def analyse(text: str) -> dict:
    return {'sections': sections_of(text), 'flags': flags_of(text)}


# ── 4) 받기·저장 ──
def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def _save(bid_no: str, status: str, message: str = '', file_name: str = '', res: Optional[dict] = None) -> dict:
    from dashboard.services.g2b_bids import _conn
    res = res or {}
    row = {'bid_no': bid_no, 'status': status, 'message': message, 'file_name': file_name,
           'sections': json.dumps(res.get('sections') or [], ensure_ascii=False),
           'flags': json.dumps(res.get('flags') or [], ensure_ascii=False), 'read_at': _now()}
    with _conn() as conn:
        conn.execute('INSERT OR REPLACE INTO g2b_notice (bid_no, status, message, file_name, sections, flags, read_at) '
                     'VALUES (:bid_no, :status, :message, :file_name, :sections, :flags, :read_at)', row)
    return view(row)


def _download(url: str) -> bytes:
    import requests
    r = requests.get(url, timeout=(10, 60), stream=True, headers={'User-Agent': 'Mozilla/5.0'})
    try:
        if r.status_code != 200:
            raise RuntimeError(f'파일을 받지 못했어요(응답 {r.status_code}).')
        if 'text/html' in (r.headers.get('Content-Type') or ''):
            raise RuntimeError('나라장터가 파일 대신 안내 화면을 돌려줬어요.')
        buf = bytearray()
        for chunk in r.iter_content(64 * 1024):
            buf += chunk
            if len(buf) > _MAX_BYTES:
                raise RuntimeError('공고문 파일이 너무 커요(20MB 넘음).')
        return bytes(buf)
    finally:
        r.close()


def read_notice(bid_no: str) -> dict:
    """공고 하나의 공고문을 받아 읽고 저장. 결과(view)를 돌려준다."""
    from dashboard.services import doc_viewer
    from dashboard.services.g2b_bids import _conn, _safe_err, sort_files
    with _conn() as conn:
        r = conn.execute('SELECT files FROM g2b_bids WHERE bid_no=?', (bid_no,)).fetchone()
    if not r:
        raise LookupError('공고를 찾지 못했어요.')
    f = pick_notice_file(sort_files(json.loads(r[0] or '[]')))
    if not f:
        return _save(bid_no, 'none', '첨부 중 공고문(한글·PDF)을 찾지 못했어요.')
    ext = f['name'].rsplit('.', 1)[-1].lower()
    try:
        data = _download(f['url'])
        text = doc_viewer.convert_to_markdown(data, ext)
    except doc_viewer.DocConvertError as e:
        return _save(bid_no, 'fail', f'공고문을 글로 바꾸지 못했어요: {e}', f['name'])
    except Exception as e:
        logger.warning(f'[G2B] 공고문 읽기 실패 {bid_no}: {_safe_err(e)}')
        msg = str(e) if isinstance(e, RuntimeError) else '공고문을 받지 못했어요. 잠시 뒤 다시 해 주세요.'
        return _save(bid_no, 'fail', msg, f['name'])
    res = analyse(text)
    if not res['sections'] and not res['flags']:
        return _save(bid_no, 'fail', '공고문은 받았지만 참가자격 단락을 찾지 못했어요. 첨부파일을 직접 확인해 주세요.', f['name'])
    return _save(bid_no, 'ok', '', f['name'], res)


def run_pending(limit: int = 40) -> int:   # 한 건 1~2초(11-20 실측 31건 37초)
    """입찰 가능한 공고(마감 전) 중 아직 안 읽은 것·실패 뒤 하루 지난 것을 몇 건씩."""
    from dashboard.services.g2b_bids import ELIGIBLE, _conn, _safe_err
    if not _run_lock.acquire(blocking=False):
        return 0
    try:
        now_s = datetime.now().strftime('%Y-%m-%d %H:%M')
        retry = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d %H:%M:%S')
        with _conn() as conn:
            rows = conn.execute(
                "SELECT b.bid_no FROM g2b_bids b LEFT JOIN g2b_notice n ON n.bid_no = b.bid_no "
                "WHERE b.category = ? AND b.close_at >= ? AND b.files NOT IN ('', '[]') "
                "AND (n.bid_no IS NULL OR (n.status = 'fail' AND n.read_at < ?)) "
                'ORDER BY b.close_at ASC LIMIT ?', (ELIGIBLE, now_s, retry, limit)).fetchall()
        done = 0
        for (bid_no,) in rows:
            try:
                read_notice(bid_no)
                done += 1
            except Exception as e:
                logger.warning(f'[G2B] 공고문 자동 읽기 실패 {bid_no}: {_safe_err(e)}')
        if done:
            logger.info(f'[G2B] 공고문 읽기 {done}건')
        return done
    finally:
        _run_lock.release()


def view(row: Optional[dict]) -> Optional[dict]:
    if not row:
        return None
    from dashboard.services.g2b_bids import get_profile
    flags = json.loads(row.get('flags') or '[]')
    grade = get_profile().get('credit_grade') or ''
    for f in flags:
        if f.get('key') == 'credit':
            f['ok'] = credit_check(f.get('line') or '', grade)
    return {'status': row.get('status') or '', 'message': row.get('message') or '',
            'file_name': row.get('file_name') or '', 'read_at': (row.get('read_at') or '')[:16],
            'sections': json.loads(row.get('sections') or '[]'), 'flags': flags}


def views_for(bid_nos: List[str]) -> Dict[str, dict]:
    from dashboard.services.g2b_bids import _conn
    if not bid_nos:
        return {}
    out = {}
    with _conn() as conn:
        for i in range(0, len(bid_nos), 500):
            part = bid_nos[i:i + 500]
            q = 'SELECT * FROM g2b_notice WHERE bid_no IN (%s)' % ','.join('?' * len(part))
            for r in conn.execute(q, part).fetchall():
                out[r['bid_no']] = view(dict(r))
    return out
