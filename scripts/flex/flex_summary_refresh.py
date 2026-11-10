"""플렉스 요약 엑셀 다시 만들기(2026-11-10) — 가장 최근 원본(flex_refresh_YYYYMMDD.json) 기준.
11/6 전체 내보내기의 휴가 상세(결재선·댓글)·결재 문서·첨부 목록은 이어 붙인다. 파일 옮기기·첨부 풀기는 안 함.
받기: scripts/flex/flex_refresh_browser.js → 이 스크립트 → scripts/hr_import_flex.py. 날짜는 아래 REFRESH 를 바꾼다."""
import json
import os
from datetime import datetime, timezone, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

OUT = r'C:\Users\SECOM\Desktop\ITG-Project\flex_export_20261106'
RAW = os.path.join(OUT, '원본')
KST = timezone(timedelta(hours=9))
old = json.load(open(os.path.join(RAW, 'flex_data_20261106.json'), encoding='utf-8'))
REFRESH = '20261110'   # 다시 받은 날
d = json.load(open(os.path.join(RAW, f'flex_refresh_{REFRESH}.json'), encoding='utf-8'))
d['timeOffDetail'] = old.get('timeOffDetail') or {}
_docs = {x['document']['documentKey']: x for x in old['approvalDocs']}
for x in d['approvalDocs']:
    _docs[x['document']['documentKey']] = x
d['approvalDocs'] = list(_docs.values())
for _u, _o in d['users'].items():           # 오늘은 2024~2027년만 받음 → 그 전 연도는 11/6 자료로
    _old_years = ((old['users'].get(_u) or {}).get('annualStatusByYear') or {})
    for _y, _v in _old_years.items():
        _o.setdefault('annualStatusByYear', {}).setdefault(_y, _v)
if not (d['company'].get('timeOffForms') or {}).get('timeOffForms'):
    d['company']['timeOffForms'] = old['company']['timeOffForms']

users = d['users']
uname = {u: o['name'] for u, o in users.items()}
policy = {'107705': '연차'}
for f in d['company']['timeOffForms'].get('timeOffForms', []):
    cf = f.get('customerTimeOffForm') or {}
    if cf.get('timeOffPolicyId'):
        policy[str(cf['timeOffPolicyId'])] = (cf.get('displayInfo') or {}).get('name', '')
STATUS_KO = {'IN_EMPLOY': '재직', 'RESIGNATION': '퇴사', 'IN_APPRENTICESHIP': '수습', 'LEAVE_OF_ABSENCE': '휴직',
             'RESIGNATION_SCHEDULED': '퇴사 예정', 'LEAVE_OF_ABSENCE_SCHEDULED': '휴직 예정',
             'COMPANY_JOIN_SCHEDULED': '입사 예정'}
USE_KO = {'APPROVAL_COMPLETED': '승인', 'APPROVAL_WAITING': '승인 대기', 'APPROVAL_DECLINED': '반려',
          'TIME_OFF_CANCELED': '취소', 'TIME_OFF_REGISTERED': '등록', 'CANCEL_APPROVAL_WAITING': '취소 승인 대기',
          'CANCEL_APPROVAL_COMPLETED': '취소 승인'}
DOC_KO = {'DONE': '완료', 'IN_PROGRESS': '진행 중', 'DECLINED': '반려', 'CANCELED': '취소'}


def ts(v):
    if v in (None, ''):
        return ''
    if isinstance(v, (int, float)):
        return datetime.fromtimestamp(v / 1000, KST).strftime('%Y-%m-%d %H:%M')
    try:
        return datetime.fromisoformat(str(v).replace('Z', '+00:00')).astimezone(KST).strftime('%Y-%m-%d %H:%M')
    except Exception:
        return str(v)


def first(lst, key='name'):
    return ', '.join(x.get(key, '') for x in (lst or []) if x.get(key))


def birth(ssn):
    if not ssn or len(ssn) < 8:
        return ''
    yy, g = ssn[:2], ssn[7]
    cent = '19' if g in '1256' else '20'
    return f'{cent}{yy}-{ssn[2:4]}-{ssn[4:6]}'


wb = openpyxl.Workbook()


def sheet(title, header, rows, widths=None):
    ws = wb.create_sheet(title)
    ws.append(header)
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='4B5563')
        c.alignment = Alignment(vertical='center')
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    for i, w in enumerate(widths or [], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    return ws


wb.remove(wb.active)

# 구성원
rows = []
for u, o in sorted(users.items(), key=lambda x: (('RESIGNATION' in x[1]['search']['tagInfo']['userStatuses']), x[1]['name'])):
    s = o['search']
    b, p, e = s['basicInfo'], s.get('privateInfo') or {}, s.get('employeeInfo') or {}
    res = (o.get('resignations') or [{}])
    res = res[0] if res else {}
    left = ''
    st26 = ((o.get('annualStatusByYear') or {}).get('2026') or {}).get('userAnnualTimeOffDetailStatus') or {}
    for m in st26.get('yearMonthlyStatus', []):
        if m['yearMonth'] == '2026-11':
            left = m['remainingTime']['timeOffDays']
    rows.append([b['name'], '/'.join(STATUS_KO.get(x, x) for x in s['tagInfo']['userStatuses']),
                 b.get('companyJoinDate', ''), b.get('companyGroupJoinDate', ''),
                 res.get('resignationDate', ''), res.get('lastWorkingDate', ''),
                 '/'.join(x for x in (res.get('reason'), res.get('subReason')) if x),
                 first(e.get('departments')), first(e.get('jobRoles')), first(e.get('jobRanks')), first(e.get('jobTitles')),
                 b.get('email', ''), p.get('personalEmail', ''), p.get('phoneNumber') or b.get('phoneNumber', ''),
                 {'MALE': '남', 'FEMALE': '여'}.get(p.get('gender'), ''), birth(p.get('ssn', '')),
                 first(((s.get('contractInfo') or {}).get('employmentContracts')), 'contractType'),
                 left, u])
sheet('구성원', ['이름', '상태', '입사일', '그룹 입사일', '퇴사일', '마지막 근무일', '퇴사 사유', '부서', '직무', '직급', '직책',
               '회사 메일', '개인 메일', '전화', '성별', '생년월일', '계약 형태', '연차 잔여(26-11)', '플렉스 ID'], rows,
      [9, 10, 11, 11, 11, 11, 18, 10, 10, 10, 8, 26, 26, 14, 5, 11, 10, 9, 12])

# 인사 발령
rows = []
for u, o in users.items():
    for a in o.get('personnelAppointments') or []:
        ei = a.get('employeeInfo') or {}
        rows.append([o['name'], a.get('applyDate', ''),
                     ', '.join((x.get('department') or {}).get('name', '') for x in ei.get('positions', [])),
                     ', '.join((x.get('jobRole') or x).get('name', '') for x in ei.get('jobRoles', [])),
                     ', '.join((x.get('jobRank') or x).get('name', '') for x in ei.get('jobRanks', [])),
                     ', '.join((x.get('jobTitle') or x).get('name', '') for x in ei.get('jobTitles', [])),
                     a.get('memo', '') or ''])
rows.sort(key=lambda r: (r[0], r[1]))
sheet('인사발령', ['이름', '적용일', '부서', '직무', '직급', '직책', '메모'], rows, [9, 11, 14, 12, 12, 10, 30])

# 휴가 사용 내역
rows = []
det = d.get('timeOffDetail') or {}
for u, o in users.items():
    for x in ((o.get('timeOffUsesAll') or {}).get('timeOffUses') or []):
        eid = x['userTimeOffRegisterEventId']
        dd = det.get(eid) or {}
        ev = ((dd.get('use') or {}).get('timeOffRegisterEvent') or x)
        task = ((dd.get('content') or {}).get('task') or {})
        appr = []
        for ln in ((task.get('approvalProcess') or {}).get('lines') or []):
            for ac in ln.get('actors', []):
                appr.append(f"{(ac.get('resolvedTarget') or {}).get('displayName', '')}({ac.get('status', '')})")
        unit = ','.join(sorted({b.get('timeOffRegisterUnit', '') for b in ev.get('userTimeOffRegisterEventBlocks', [])}))
        comments = ' / '.join(f"{(c.get('writer') or {}).get('name', '')}: {c.get('content')}"
                              for c in task.get('comments', []) if c.get('content'))
        rows.append([o['name'], policy.get(str(x.get('timeOffPolicyId')), x.get('timeOffPolicyId')),
                     x['timeOffRegisterDateFrom'], x['timeOffRegisterDateTo'],
                     (x.get('useTime') or {}).get('timeOffDays'), unit,
                     USE_KO.get(x.get('timeOffUseStatus'), x.get('timeOffUseStatus')),
                     ev.get('memo', '') or task.get('content', ''), ' → '.join(appr), comments,
                     ts(x.get('timeOffRegisteredAt')), eid])
rows.sort(key=lambda r: (r[2], r[0]), reverse=True)
sheet('휴가 사용 내역', ['이름', '종류', '시작', '끝', '일수', '단위', '상태', '사유(메모)', '승인', '댓글', '신청 시각', '플렉스 ID'], rows,
      [9, 14, 11, 11, 6, 8, 10, 30, 36, 30, 16, 10])

# 연차 월별 현황
rows = []
for u, o in users.items():
    for y, v in sorted((o.get('annualStatusByYear') or {}).items()):
        st = (v or {}).get('userAnnualTimeOffDetailStatus') or {}
        for m in st.get('yearMonthlyStatus', []):
            g = lambda k: m[k]['totalAmountWithDays']['timeOffDays']
            vals = [g('assigns'), g('uses'), g('expirations'), g('adjusts')]
            rem = m['remainingTime']['timeOffDays']
            if any(vals) or rem:
                rows.append([o['name'], m['yearMonth'], *vals, rem])
rows.sort(key=lambda r: (r[0], r[1]))
sheet('연차 월별', ['이름', '연월', '부여', '사용', '소멸', '조정', '잔여'], rows, [9, 9, 7, 7, 7, 7, 7])

# 결재 문서
rows = []
for x in d['approvalDocs']:
    doc, pr = x['document'], x.get('approvalProcess') or {}
    appr = []
    for ln in pr.get('lines') or []:
        for ac in ln.get('actors', []):
            appr.append(f"{(ac.get('resolvedTarget') or {}).get('displayName', '')}({ac.get('status', '')})")
    inputs = '\n'.join(f"{(i.get('inputField') or {}).get('name', '')}: {i.get('value', '')}" for i in doc.get('inputs') or [])
    comments = '\n'.join(f"{(c.get('writer') or {}).get('name', '')}: {c.get('content')}"
                         for c in doc.get('comments') or [] if c.get('content') and not c.get('writtenBySystem'))
    att = '\n'.join(a['file']['fileName'] for a in doc.get('attachments') or [])
    rows.append([doc.get('code'), doc.get('title'), DOC_KO.get(doc.get('status'), doc.get('status')),
                 (doc.get('writer') or {}).get('name', ''), ts(doc.get('writtenAt')), ts(pr.get('terminatedAt')),
                 ' → '.join(appr), inputs[:32000], comments[:32000], att, doc.get('documentKey')])
rows.sort(key=lambda r: str(r[0]), reverse=True)
ws = sheet('결재 문서', ['번호', '제목', '상태', '작성자', '작성', '종결', '결재선', '입력 내용', '댓글', '첨부', '문서 키'], rows,
           [11, 26, 8, 8, 16, 16, 30, 50, 30, 30, 12])
for r in ws.iter_rows(min_row=2):
    for c in r:
        c.alignment = Alignment(wrap_text=True, vertical='top')


# 첨부 파일 목록은 11/6 요약에서 그대로
_ow = openpyxl.load_workbook(os.path.join(OUT, '플렉스_요약_20261106.xlsx'), read_only=True)
_rows = list(_ow['첨부 파일 목록'].iter_rows(values_only=True))
sheet('첨부 파일 목록', list(_rows[0]), [list(r) for r in _rows[1:]], [14, 90, 10, 12])

xp = os.path.join(OUT, f'플렉스_요약_{REFRESH}.xlsx')
wb.save(xp)
print('엑셀', xp)
for w in wb.worksheets:
    print(' ', w.title, w.max_row - 1)
