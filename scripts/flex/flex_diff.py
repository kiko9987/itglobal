"""플렉스 11/6 전체 내보내기 ↔ 오늘 다시 받은 자료 비교 — 값은 찍지 않고 '무엇이 바뀌었는지'만.
사용: python -I flex_diff.py <옛 json> <새 json>
"""
import json
import sys
from collections import Counter

old = json.load(open(sys.argv[1], encoding='utf-8'))
new = json.load(open(sys.argv[2], encoding='utf-8'))
SENSITIVE = {'ssn', 'phoneNumber', 'addressFull', 'addressDetails', 'address1', 'address2', 'address3', 'amount',
             'birthDate', 'email', 'originEmail', 'addressZipCode', 'addressCoordinateLatitude',
             'addressCoordinateLongitude'}


def flat(o, prefix=''):
    out = {}
    if isinstance(o, dict):
        for k, v in o.items():
            out.update(flat(v, f'{prefix}.{k}' if prefix else k))
    elif isinstance(o, list):
        out[prefix] = json.dumps(o, ensure_ascii=False, sort_keys=True)[:2000]
    else:
        out[prefix] = o
    return out


ou, nu = old['users'], new['users']
print(f'구성원: 11/6 {len(ou)}명 → 오늘 {len(nu)}명')
for u in nu.keys() - ou.keys():
    s = nu[u]['search']['basicInfo']
    print(f'  + 새로 생김: {nu[u]["name"]} (입사 {s.get("companyJoinDate")})')
for u in ou.keys() - nu.keys():
    print(f'  - 없어짐: {ou[u]["name"]}')

SECTIONS = ['employee', 'personalBundle', 'employeeBundle', 'employmentContracts', 'salaryContracts',
            'personnelAppointments', 'resignations', 'leaveOfAbsences', 'workExperiences', 'educationExperiences',
            'dependentFamilies', 'careers', 'licenses', 'memos', 'documentLegacies', 'documentFiles',
            'certificateHistories', 'annualPolicy', 'timeOffContext']
IGNORE_SUFFIX = ('companyWorkingPeriod.days', 'companyWorkingPeriod.totalDays', 'companyWorkingPeriod.months',
                 'companyWorkingPeriod.totalMonths', 'companyWorkingPeriod.years', 'workingDayInfo.workingDays',
                 'workingDayInfo.workingMonths', 'url', 'Url', 'signedUrl', 'expiresAt', 'expiredAt')
print('\n[사람별 바뀐 항목] (값은 안 찍음, 근속일수·파일 링크 만료 같은 자동 변화 제외)')
for u in sorted(nu.keys() & ou.keys(), key=lambda x: nu[x]['name']):
    changes = []
    for sec in SECTIONS + ['search']:
        a, b = flat(ou[u].get(sec)), flat(nu[u].get(sec))
        for k in sorted(set(a) | set(b)):
            if k.endswith(IGNORE_SUFFIX) or 'Durations' in k or 'signed' in k.lower() or k.endswith('.url'):
                continue
            if a.get(k) != b.get(k):
                leaf = k.split('.')[-1]
                changes.append(f'{sec}.{leaf}' + ('(민감)' if leaf in SENSITIVE else ''))
    if changes:
        c = Counter(changes)
        print(f'  {nu[u]["name"]}: ' + ', '.join(f'{k}×{v}' if v > 1 else k for k, v in c.most_common(12)))

print('\n[휴가 사용 내역]')
def uses(d):
    m = {}
    for u, v in d['users'].items():
        for x in ((v.get('timeOffUsesAll') or {}).get('timeOffUses') or []):
            m[str(x.get('userTimeOffRegisterEventId'))] = (d['users'][u]['name'], x)
    return m
uo, un = uses(old), uses(new)
print(f'  11/6 {len(uo)}건 → 오늘 {len(un)}건')
for k in sorted(un.keys() - uo.keys()):
    name, x = un[k]
    print(f'  + {name} {x.get("timeOffRegisterDateFrom")}~{x.get("timeOffRegisterDateTo")} 상태={x.get("timeOffUseStatus")} 취소={x.get("canceled")}')
for k in sorted(uo.keys() & un.keys()):
    a, b = uo[k][1], un[k][1]
    diff = [f for f in ('timeOffUseStatus', 'canceled', 'timeOffRegisterDateFrom', 'timeOffRegisterDateTo', 'useTime')
            if a.get(f) != b.get(f)] + (['approval'] if json.dumps(a.get('approvalStatus'), sort_keys=True) != json.dumps(b.get('approvalStatus'), sort_keys=True) else [])
    if diff:
        print(f'  ~ {un[k][0]} {b.get("timeOffRegisterDateFrom")} 바뀜: {",".join(diff)} → {b.get("timeOffUseStatus")}')
for k in sorted(uo.keys() - un.keys()):
    print(f'  - 없어짐: {uo[k][0]} {uo[k][1].get("timeOffRegisterDateFrom")}')

print('\n[연차 현황(올해) 바뀐 사람]')
for u in sorted(nu.keys(), key=lambda x: nu[x]['name']):
    a = json.dumps(((ou.get(u) or {}).get('annualStatusByYear') or {}).get('2026'), sort_keys=True)
    b = json.dumps((nu[u].get('annualStatusByYear') or {}).get('2026'), sort_keys=True)
    if a != b:
        print(f'  {nu[u]["name"]}')

print('\n[결재 문서]')
keys_old = {d.get('document', {}).get('documentKey') or d.get('documentKey') for d in old.get('approvalDocs') or []}
lst = new.get('approvalList') or []
print(f'  목록 {len(lst)}건 받음 · 11/6 상세 {len(keys_old)}건')
newer = [x for x in lst if x['key'] not in keys_old]
print(f'  11/6 이후 새 문서 {len(newer)}건: ' + ', '.join(f'{x["title"]}({x["status"]})' for x in newer[:20]))
print(f'\n오류 {len(new.get("errors") or [])}건')
for e in (new.get('errors') or [])[:10]:
    print('  ', e.get('s'), e.get('u'))
