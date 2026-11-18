"""휴가 엑셀 내려받기(2026-11-18, 플렉스 휴가 관리 맞추기) — 보유 현황·사용 내역·월별 연차 사용.

전 직원 휴가를 볼 수 있는 사람(hr.can_view_all_leave)만. 화면과 같은 계산(leave_view·leave)을 그대로 쓴다.
직원이 쓴 사유가 엑셀 수식으로 실행되지 않게 xlsx_append_text 로 넣는다(2026-11-13 보안 점검).
"""
from __future__ import annotations

import io
from datetime import date, timedelta
from typing import List, Optional

from dashboard.services import hr
from dashboard.services import leave as lv
from dashboard.services import leave_view

KINDS = {'balance': '휴가 보유 현황', 'usage': '휴가 사용 내역', 'monthly': '월별 연차 사용'}


def _book(title: str, head: List[str], rows: List[list], widths: List[int], number_cols=()) -> bytes:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from dashboard.utils.spreadsheet_safe import xlsx_append_text

    wb = Workbook()
    ws = wb.active
    ws.title = title[:31]
    ws.append(head)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', fgColor='E9ECEF')
        cell.alignment = Alignment(horizontal='center')
    for r in rows:
        xlsx_append_text(ws, r)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for c in number_cols:
        for (cell,) in ws.iter_rows(min_row=2, min_col=c, max_col=c):
            cell.number_format = '0.0#'
    ws.freeze_panes = 'B2'
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def balance_rows(today: date) -> List[list]:
    out = []
    for p in hr.list_people(include_resigned=False):
        full = hr.get_person(p['id'])
        s = leave_view.summary(full, today)
        soon = min(s['grants'], key=lambda g: g['valid_to']) if s['grants'] else None
        name = leave_view.leave_name(full)
        custom = ''
        if name:
            parts = [f"{c['label']} {lv.fmt_days(c['left'])}" for c in lv.custom_balances(name, today) if c['left'] > 0]
            custom = ', '.join(parts)
        out.append([p['name'], p.get('department') or '', p.get('hire_date') or '',
                    s['balance'] if s['linked'] else None,
                    (f"{lv.fmt_days(soon['days'])} · {soon['valid_to']}까지" if soon else ''),
                    s['annual_used'], s['sick_used'], s['pending'], custom])
    return out


def balance_xlsx(today: Optional[date] = None) -> bytes:
    today = today or date.today()
    head = ['이름', '부서', '입사일', f'남은 연차({today.isoformat()} 기준)', '가장 먼저 소멸',
            f'{today.year}년 쓴 연차', f'{today.year}년 병가', '승인 대기(건)', '맞춤 휴가 남은 일수']
    return _book('보유 현황', head, balance_rows(today), [10, 12, 12, 18, 22, 14, 12, 12, 28], number_cols=(4, 6, 7))


def usage_rows(year: int) -> List[list]:
    out = []
    y = str(year)
    for p in hr.list_people(include_resigned=True):
        full = hr.get_person(p['id'])
        for e in leave_view.entries_for(full):
            if e['start'][:4] != y and (e.get('end') or e['start'])[:4] != y:
                continue
            out.append([p['name'], e['type'], e['start'], e.get('end') or e['start'], e.get('half') or '',
                        e['days'], e['status'], e.get('reason') or '', e.get('requested_at') or '',
                        '플렉스' if e['source'] == 'flex' else '그룹웨어·슬랙'])
    out.sort(key=lambda r: (r[2], r[0]))
    return out


def usage_xlsx(year: int) -> bytes:
    head = ['이름', '종류', '시작일', '마지막 날', '반차', '일수', '상태', '사유', '신청 시각', '기록']
    return _book(f'{year} 사용 내역', head, usage_rows(year), [10, 18, 12, 12, 7, 7, 9, 36, 17, 13], number_cols=(6,))


def _split_by_month(e: dict, year: int) -> dict:
    """한 휴가의 일수를 달마다 — 한 달 안이면 그대로, 달을 넘으면 그 달에 든 근무일만큼."""
    s, t = lv._d(e['start']), lv._d(e.get('end') or e['start'])
    if not s or not t:
        return {}
    if (s.year, s.month) == (t.year, t.month):
        return {s.month: e['days']} if s.year == year else {}
    out = {}
    for m in range(1, 13):
        a = max(s, date(year, m, 1))
        b = min(t, date(year + (m == 12), m % 12 + 1, 1) - timedelta(days=1))
        if a <= b:
            n = lv.count_days(lv.ANNUAL, a, b)
            if n:
                out[m] = n
    return out


def monthly_rows(year: int) -> List[list]:
    out = []
    for p in hr.list_people(include_resigned=True):
        full = hr.get_person(p['id'])
        per = [0.0] * 12
        for e in leave_view.entries_for(full):
            if e['status'] != leave_view.ST_OK or not e['annual']:
                continue
            for m, d in _split_by_month(e, year).items():
                per[m - 1] += d
        if not any(per) and p.get('status') == hr.STATUS_RESIGNED:
            continue
        out.append([p['name'], p.get('department') or ''] + [round(x, 2) for x in per] + [round(sum(per), 2)])
    return out


def monthly_xlsx(year: int) -> bytes:
    head = ['이름', '부서'] + [f'{m}월' for m in range(1, 13)] + ['합계']
    return _book(f'{year} 월별 연차', head, monthly_rows(year), [10, 12] + [7] * 12 + [8],
                 number_cols=tuple(range(3, 16)))


def build(kind: str, year: int, today: Optional[date] = None):
    """(파일 내용, 파일 이름) — kind 가 틀리면 ValueError."""
    if kind == 'balance':
        t = today or date.today()
        return balance_xlsx(t), f'휴가_보유현황_{t.isoformat()}.xlsx'
    if kind == 'usage':
        return usage_xlsx(year), f'휴가_사용내역_{year}.xlsx'
    if kind == 'monthly':
        return monthly_xlsx(year), f'월별_연차사용_{year}.xlsx'
    raise ValueError(kind)
