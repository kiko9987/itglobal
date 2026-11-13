# -*- coding: utf-8 -*-
"""월별 수금 완료 정산 자료 (에어테이블 '수금 완료 💲 - YY년M월' 보기 대체, 2026-11-08).

경영지원(SB)이 매달 에어테이블 보기를 내려받아 정산에 쓰던 것을 PM 에서 바로 보고 엑셀로 받게 한다.
- 'N월' = 전월 26일 ~ 당월 25일 (정산 25일 기준). 에어테이블 25년1월(74건)·26년9월(106건) 보기와
  시트 수금 날짜를 대조해 확정 — 8/25 수금분은 9월이 아니라 8월.
  명절이 긴 달은 기준일이 바뀌어 달별로 기간을 바꿔 저장할 수 있음(2026-11-09, resolve_range).
- 조건: 수금 확인 체크 + 수금 날짜가 기간 안.
- 묶음: 담당자 → 사업자 (에어테이블 보기와 같음). 칸 순서도 그 보기와 같고 맨 앞에 담당자만 더함
  (에어테이블은 묶음 칸이라 내려받은 파일에 담당자가 빠졌음).
"""
from __future__ import annotations

import io
import re
from datetime import date, datetime, timedelta
from typing import Iterable, Optional

COLUMNS = [
    '담당자', '프로젝트 코드', '사업자', '유입 구분', '사업자명', '현장 주소',
    '총액 1', '부가세', '총액 2', '계약금', '중도금', '잔금', '미수금',
    '계산서', '계약금 계산서', '중도금 계산서', '잔금 계산서', '수금 날짜', '수금 확인',
    '도급 구분', '시공자', '공사 내용', '수금 관련 특이사항',
    '제품대', '도급비', '자재비', '기타비', '순익', '마진율',
]
MONEY = ['총액 1', '총액 2', '계약금', '중도금', '잔금', '미수금', '제품대', '도급비', '자재비', '기타비', '순익']
CHECKS = ['부가세', '수금 확인']
SUM_FIELDS = list(MONEY)  # 에어테이블 보기처럼 계약금·중도금·잔금까지 합계
CUTOFF_DAY = 25


def period_range(year: int, month: int) -> tuple[date, date]:
    """'year년 month월' 정산 기간 = 전월 26일 ~ 당월 25일."""
    py, pm = (year - 1, 12) if month == 1 else (year, month - 1)
    return date(py, pm, CUTOFF_DAY + 1), date(year, month, CUTOFF_DAY)


def period_of(d: date) -> tuple[int, int]:
    """날짜가 속한 정산 월. 26일부터는 다음 달."""
    if d.day > CUTOFF_DAY:
        return (d.year + 1, 1) if d.month == 12 else (d.year, d.month + 1)
    return d.year, d.month


def _key(year: int, month: int) -> str:
    return f'{year}-{month:02d}'


def _shift(year: int, month: int, k: int) -> tuple[int, int]:
    n = year * 12 + (month - 1) + k
    return n // 12, n % 12 + 1


def _saved(saved: Optional[dict], year: int, month: int) -> Optional[tuple[date, date]]:
    """저장된 기간 'YYYY-MM-DD|YYYY-MM-DD' → (시작, 끝). 없거나 깨졌으면 None."""
    v = (saved or {}).get(_key(year, month))
    try:
        s, e = str(v).split('|')
        return date.fromisoformat(s), date.fromisoformat(e)
    except Exception:
        return None


def resolve_range(year: int, month: int, saved: Optional[dict] = None) -> tuple[date, date, bool]:
    """정산 기간 (시작, 끝, 직접 정했는지).

    명절이 긴 달은 경영지원이 기간을 바꿔 정함(2026-11-09 SB 요청) → 그 달에 저장된 기간을 씀.
    저장이 없으면 전월 26일 ~ 당월 25일인데, 앞 달 끝을 바꿔 뒀으면 시작일은 그 다음 날부터
    (앞 달을 10/23 에 끊었으면 이번 달은 10/24 부터 → 빠지거나 두 번 들어가는 날이 없게).
    """
    own = _saved(saved, year, month)
    if own:
        return own[0], own[1], True
    s, e = period_range(year, month)
    prev = _saved(saved, *_shift(year, month, -1))
    if prev and prev[1] < e:
        s = prev[1] + timedelta(days=1)
    return s, e, False


MAX_RANGE_DAYS = 62


def check_range(start: date, end: date) -> Optional[str]:
    """저장 전에 기간이 말이 되는지. 문제 있으면 안내 문구."""
    if start > end:
        return '시작일이 끝날보다 늦습니다.'
    if (end - start).days + 1 > MAX_RANGE_DAYS:
        return f'기간은 {MAX_RANGE_DAYS}일을 넘길 수 없습니다.'
    return None


def _md(d: date) -> str:
    return f'{d.month}/{d.day}'


def range_notes(year: int, month: int, saved: Optional[dict] = None) -> list[str]:
    """앞뒤 달 기간과 겹치거나 비는 날 안내 (그 날 수금분이 두 번 들어가거나 빠짐)."""
    s, e, _ = resolve_range(year, month, saved)
    notes = []
    py, pm = _shift(year, month, -1)
    ps, pe, _ = resolve_range(py, pm, saved)
    if pe >= s:
        notes.append(f'{pm}월({_md(ps)}~{_md(pe)})과 {_md(s)}~{_md(min(pe, e))}가 겹칩니다. '
                     f'그 사이 수금분은 두 달에 모두 들어갑니다.')
    elif (s - pe).days > 1:
        notes.append(f'{pm}월이 {_md(pe)}에 끝나서 {_md(pe + timedelta(days=1))}~{_md(s - timedelta(days=1))} '
                     f'수금분은 어느 달에도 들어가지 않습니다.')
    ny, nm = _shift(year, month, 1)
    nxt = _saved(saved, ny, nm)
    if nxt:
        ns, ne = nxt
        if ns <= e:
            notes.append(f'{nm}월({_md(ns)}~{_md(ne)})과 {_md(max(ns, s))}~{_md(e)}가 겹칩니다. '
                         f'그 사이 수금분은 두 달에 모두 들어갑니다.')
        elif (ns - e).days > 1:
            notes.append(f'{nm}월이 {_md(ns)}부터라서 {_md(e + timedelta(days=1))}~{_md(ns - timedelta(days=1))} '
                         f'수금분은 어느 달에도 들어가지 않습니다.')
    return notes


def period_label(year: int, month: int, saved: Optional[dict] = None) -> str:
    s, e, custom = resolve_range(year, month, saved)
    return f"{year % 100}년 {month}월 ({_md(s)}~{_md(e)}){' · 기간 바꿈' if custom else ''}"


def parse_period(s: str, today: Optional[date] = None) -> tuple[int, int]:
    """'2026-10' → (2026, 10). 비었거나 틀리면 오늘이 속한 정산 월."""
    m = re.fullmatch(r'\s*(\d{4})-(\d{1,2})\s*', str(s or ''))
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(1)), int(m.group(2))
    return period_of(today or date.today())


def _date(v) -> Optional[date]:
    if v is None:
        return None
    if isinstance(v, datetime):   # pandas Timestamp 포함 (datetime 하위형)
        try:
            return date(v.year, v.month, v.day)
        except Exception:   # NaT (year 가 nan)
            return None
    if isinstance(v, date):
        return v
    m = re.match(r'\s*(\d{4})[-./]\s*(\d{1,2})[-./]\s*(\d{1,2})', str(v))
    if not m:
        return None
    try:
        return date(*map(int, m.groups()))
    except ValueError:
        return None


def _checked(v) -> bool:
    if isinstance(v, bool):
        return v
    return str(v or '').strip().upper() in ('TRUE', '1', 'Y', 'O', '✓', '✔')


def _money(v) -> int:
    s = re.sub(r'[^\d.\-]', '', str(v if v is not None else ''))
    if s in ('', '-', '.', '-.'):
        return 0
    try:
        return int(round(float(s)))
    except ValueError:
        return 0


def _text(v) -> str:
    if v is None:
        return ''
    if hasattr(v, 'strftime'):
        try:
            return v.strftime('%Y-%m-%d')
        except Exception:
            return ''
    s = str(v).strip()
    return '' if s.lower() in ('nan', 'nat', 'none') else s


def collect(records: Iterable[dict], year: int, month: int, saved: Optional[dict] = None) -> dict:
    """시트 행들 → 정산 기간의 수금 완료 행 + 담당자→사업자 묶음 합계.

    saved = 달별로 바꿔 둔 기간 {'YYYY-MM': 'YYYY-MM-DD|YYYY-MM-DD'} (없으면 26~25일 규칙).
    """
    start, end, custom = resolve_range(year, month, saved)
    d_start, d_end = period_range(year, month)
    rows = []
    for r in records:
        code = _text(r.get('프로젝트 코드'))
        if not code or not _checked(r.get('수금 확인')):
            continue
        d = _date(r.get('수금 날짜'))
        if not d or not (start <= d <= end):
            continue
        row = {}
        for c in COLUMNS:
            v = r.get(c)
            if c in MONEY:
                row[c] = _money(v)
            elif c in CHECKS:
                row[c] = _checked(v)
            elif c == '수금 날짜':
                row[c] = d.isoformat()
            else:
                row[c] = _text(v)
        row['담당자'] = row['담당자'] or '(담당자 없음)'
        rows.append(row)
    rows.sort(key=lambda x: (x['담당자'], x['사업자'], x['수금 날짜'], x['프로젝트 코드']))

    def _sum(rs):
        return {f: sum(x[f] for x in rs) for f in SUM_FIELDS} | {'건수': len(rs)}

    groups = []
    for owner in sorted({x['담당자'] for x in rows}):
        o_rows = [x for x in rows if x['담당자'] == owner]
        subs = []
        for biz in sorted({x['사업자'] for x in o_rows}):
            b_rows = [x for x in o_rows if x['사업자'] == biz]
            subs.append({'사업자': biz or '(사업자 없음)', 'rows': b_rows, 'sum': _sum(b_rows)})
        groups.append({'담당자': owner, 'groups': subs, 'sum': _sum(o_rows)})
    return {
        'year': year, 'month': month,
        'start': start.isoformat(), 'end': end.isoformat(),
        'custom': custom, 'label': period_label(year, month, saved),
        'default_start': d_start.isoformat(), 'default_end': d_end.isoformat(),
        'notes': range_notes(year, month, saved),
        'columns': COLUMNS, 'money': MONEY,
        'rows': rows, 'groups': groups, 'total': _sum(rows),
    }


def recent_periods(today: Optional[date] = None, n: int = 24, saved: Optional[dict] = None) -> list[dict]:
    """선택 목록용 최근 정산 월 (현재 정산 월부터 거꾸로)."""
    y, m = period_of(today or date.today())
    out = []
    for _ in range(n):
        out.append({'value': _key(y, m), 'label': period_label(y, m, saved)})
        y, m = _shift(y, m, -1)
    return out


def to_xlsx(report: dict) -> bytes:
    """엑셀: '수금 완료' 시트(행 그대로, 에어테이블 내려받기와 같은 칸) + '담당자별 합계' 시트."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = '수금 완료'
    ws.append(COLUMNS)
    from dashboard.utils.spreadsheet_safe import xlsx_append_text   # 시트 글이 수식으로 안 되게 (2026-11-13)
    for row in report['rows']:
        xlsx_append_text(ws, [('✓' if row[c] else '') if c in CHECKS else row[c] for c in COLUMNS])
    bold = Font(bold=True)
    head_fill = PatternFill('solid', fgColor='E9ECEF')
    for cell in ws[1]:
        cell.font = bold
        cell.fill = head_fill
        cell.alignment = Alignment(horizontal='center')
    for i, c in enumerate(COLUMNS, start=1):
        if c in MONEY:
            for (cell,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                cell.number_format = '#,##0'
    widths = {'담당자': 9, '프로젝트 코드': 13, '사업자': 8, '유입 구분': 9, '사업자명': 22, '현장 주소': 34,
              '공사 내용': 34, '수금 관련 특이사항': 30, '수금 날짜': 11}
    from openpyxl.utils import get_column_letter
    for i, c in enumerate(COLUMNS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 12)
    ws.freeze_panes = 'C2'

    ws2 = wb.create_sheet('담당자별 합계')
    head = ['담당자', '사업자', '건수'] + SUM_FIELDS
    ws2.append(head)
    for cell in ws2[1]:
        cell.font = bold
        cell.fill = head_fill
    sub_fill = PatternFill('solid', fgColor='F8F9FA')
    for g in report['groups']:
        for s in g['groups']:
            ws2.append([g['담당자'], s['사업자'], s['sum']['건수']] + [s['sum'][f] for f in SUM_FIELDS])
        ws2.append([f"{g['담당자']} 소계", '', g['sum']['건수']] + [g['sum'][f] for f in SUM_FIELDS])
        for cell in ws2[ws2.max_row]:
            cell.font = bold
            cell.fill = sub_fill
    t = report['total']
    ws2.append(['전체 합계', '', t['건수']] + [t[f] for f in SUM_FIELDS])
    for cell in ws2[ws2.max_row]:
        cell.font = bold
    for row in ws2.iter_rows(min_row=2, min_col=4, max_col=len(head)):
        for cell in row:
            cell.number_format = '#,##0'
    for i, w in enumerate([14, 10, 7] + [14] * len(SUM_FIELDS), start=1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def xlsx_filename(report: dict) -> str:
    s, e = date.fromisoformat(report['start']), date.fromisoformat(report['end'])
    return (f"수금완료_{report['year'] % 100}년{report['month']}월"
            f"({s.month:02d}.{s.day:02d}-{e.month:02d}.{e.day:02d}).xlsx")
