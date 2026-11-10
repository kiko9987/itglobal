# -*- coding: utf-8 -*-
"""플렉스 구성원·인사발령 → 인사 장부(hr_people·hr_appointments) (2026-11-10).

읽는 곳: flex_export_20261106/ 의 가장 최근 요약 엑셀(플렉스_요약_YYYYMMDD.xlsx: 구성원·인사발령·휴가 사용 내역)
        + 가장 최근 원본 JSON(원본/flex_refresh_*.json, 없으면 flex_data_*.json: 주민번호·주소·휴가 건↔직원).
  2026-11-10 다시 받은 자료(김민욱 입사·휴가 2건)부터 자동으로 최신을 고른다. --summary/--raw 로 직접 지정 가능.
  - 연봉 등 급여 정보는 옮기지 않는다.
  - 같은 사람(플렉스 ID)은 다시 돌려도 한 줄. 사람이 화면에서 고친 칸은 덮어쓰지 않는다.
  - 회사 메일(@itg-aircon.com)이 있으면 그 메일로 로그인(staff), 없으면 현장·미화(field) — 로그인 메일은 경영지원이 나중에 넣음.
  - 슬랙 ID 는 휴가 장부(leave_staff)의 같은 이름에서 가져온다.
  - 주민번호는 .env GW_HR_KEY 로 암호화해 저장. 키가 없으면 --apply 를 멈춘다.
  - 화면 출력엔 주민번호·전화·주소를 찍지 않는다.

사용:  .venv/Scripts/python.exe scripts/hr_import_flex.py                    # 미리 보기(쓰지 않음)
       .venv/Scripts/python.exe scripts/hr_import_flex.py --apply            # 운영 사용자 DB 에 쓰기
       .venv/Scripts/python.exe scripts/hr_import_flex.py --db <경로> --apply  # 다른 DB(시험)에 쓰기
create_app() 은 부르지 않는다(스케줄러 이중 기동 방지).
"""
import argparse
import json
import os
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

DEFAULT_FLEX_DIR = Path(r'C:/Users/SECOM/Desktop/ITG-Project/flex_export_20261106')
COMPANY_DOMAIN = '@itg-aircon.com'


def _d(v):
    """엑셀 날짜 칸 → 'YYYY-MM-DD' 또는 None."""
    if v in (None, ''):
        return None
    if isinstance(v, (datetime, date)):
        return v.strftime('%Y-%m-%d')
    s = str(v).strip()
    if not s or s == '-':
        return None
    try:
        return date.fromisoformat(s[:10].replace('.', '-').replace('/', '-')).isoformat()
    except ValueError:
        return None


def _s(v):
    s = '' if v is None else str(v).strip()
    return s if s and s != '-' else None


_CONTRACT = {'REGULAR': '정규직', 'CONTRACT': '계약직', 'TEMPORARY': '계약직', 'PART_TIME': '시간제',
             'INTERN': '인턴', 'DISPATCH': '파견'}


def _contract(v):
    s = _s(v)
    return _CONTRACT.get(s.upper(), s) if s else None


def _status(v):
    s = str(v or '')
    if '퇴사' in s:
        return '퇴사'
    if '휴직' in s:
        return '휴직'
    return '재직'


def _sheet(wb, title):
    ws = wb[title]
    rows = ws.iter_rows(values_only=True)
    hdr = [str(h).strip() if h is not None else '' for h in next(rows)]
    return [dict(zip(hdr, r)) for r in rows if any(c not in (None, '') for c in r)]


def _latest(flex_dir: Path, patterns) -> Path:
    """이름 끝 날짜(YYYYMMDD)가 가장 늦은 파일. patterns 는 앞에 있을수록 우선(같은 날짜면)."""
    best = None
    for rank, pat in enumerate(patterns):
        for f in flex_dir.glob(pat):
            key = (f.stem[-8:], -rank)
            if best is None or key > best[0]:
                best = (key, f)
    if not best:
        raise FileNotFoundError(f'{flex_dir} 에 {patterns} 파일이 없습니다.')
    return best[1]


def summary_path(flex_dir: Path) -> Path:
    return _latest(flex_dir, ['플렉스_요약_*.xlsx'])


def raw_path(flex_dir: Path) -> Path:
    return _latest(flex_dir / '원본', ['flex_refresh_*.json', 'flex_data_*.json'])


def _private(flex_json, fid):
    """원본 JSON 에서 주민번호·주소만."""
    u = (flex_json.get('users') or {}).get(fid) or {}
    p = ((u.get('personalBundle') or {}).get('personal') or {})
    ssn = (p.get('ssn') or '').strip()
    addr = ' '.join(x for x in [(p.get('addressFull') or '').strip(), (p.get('addressDetails') or '').strip()] if x)
    return ssn, addr or None


def build_records(flex_dir: Path, leave_slack: dict, summary: Path = None, raw: Path = None):
    import openpyxl
    wb = openpyxl.load_workbook(summary or summary_path(flex_dir), read_only=True, data_only=True)
    members = _sheet(wb, '구성원')
    appts_rows = _sheet(wb, '인사발령')
    flex_json = json.load(open(raw or raw_path(flex_dir), encoding='utf-8'))

    # 같은 이름이 둘 이상(예: 재직 박민재·퇴사 박민재 = 다른 사람)이면 발령 적용일이 근무 기간 안에 드는 사람에게
    by_name = {}
    for m in members:
        by_name.setdefault(_s(m.get('이름')), []).append(m)

    def _owner(name, apply_date):
        cands = by_name.get(name) or []
        if len(cands) == 1:
            return _s(cands[0].get('플렉스 ID'))
        hits = [c for c in cands
                if (_d(c.get('입사일')) or '0000') <= (apply_date or '9999') <= (_d(c.get('퇴사일')) or '9999')]
        return _s(hits[0].get('플렉스 ID')) if len(hits) == 1 else None

    appts = {}
    skipped_appts = 0
    for a in appts_rows:
        fid = _owner(_s(a.get('이름')), _d(a.get('적용일')))
        if not fid:
            skipped_appts += 1
            continue
        appts.setdefault(fid, []).append({
            'apply_date': _d(a.get('적용일')), 'department': _s(a.get('부서')), 'job_role': _s(a.get('직무')),
            'job_rank': _s(a.get('직급')), 'job_title': _s(a.get('직책')), 'memo': _s(a.get('메모')),
        })

    records = []
    for m in members:
        fid = _s(m.get('플렉스 ID'))
        name = _s(m.get('이름'))
        if not fid or not name:
            continue
        work = (_s(m.get('회사 메일')) or '').lower() or None
        personal = (_s(m.get('개인 메일')) or '').lower() or None
        # 플렉스 '회사 메일' 칸에 개인 메일(네이버 등)을 넣어 둔 현장 직원 → 개인 메일로 옮김
        if work and not work.endswith(COMPANY_DOMAIN):
            personal = personal or work
            work = None
        status = _status(m.get('상태'))
        active = status != '퇴사'
        staff = bool(work and work.endswith(COMPANY_DOMAIN))
        ssn, addr = _private(flex_json, fid)
        records.append({
            'flex_id': fid, 'name': name, 'status': status,
            'hire_date': _d(m.get('입사일')), 'group_hire_date': _d(m.get('그룹 입사일')),
            'resign_date': _d(m.get('퇴사일')), 'last_work_date': _d(m.get('마지막 근무일')),
            'resign_reason': _s(m.get('퇴사 사유')),
            'department': _s(m.get('부서')), 'job_role': _s(m.get('직무')), 'job_rank': _s(m.get('직급')),
            'job_title': _s(m.get('직책')), 'contract_type': _contract(m.get('계약 형태')),
            'work_email': work if (work and '@' in work) else None,
            'personal_email': personal if (personal and '@' in personal) else None,
            'phone': _s(m.get('전화')), 'gender': _s(m.get('성별')), 'birth_date': _d(m.get('생년월일')),
            'address': addr,
            # 퇴사자는 로그인·슬랙 연결 없음(같은 이름 재직자와 섞이지 않게)
            'login_email': work if (active and staff) else None,
            'access': 'field' if (active and not staff) else 'staff',
            'slack_id': leave_slack.get(name) if active else None,
            'ssn': ssn,
            '_appts': appts.get(fid, []),
        })
    return records, skipped_appts


def build_leave_history(flex_dir: Path, summary: Path = None, raw: Path = None):
    """플렉스 '휴가 사용 내역' 시트 → 보관용 줄. 건 번호로 원본 JSON 에서 직원(플렉스 ID)을 찾는다(동명이인 구분)."""
    import openpyxl
    wb = openpyxl.load_workbook(summary or summary_path(flex_dir), read_only=True, data_only=True)
    hist = _sheet(wb, '휴가 사용 내역')
    flex_json = json.load(open(raw or raw_path(flex_dir), encoding='utf-8'))
    owner = {}
    for uid, v in (flex_json.get('users') or {}).items():
        for x in ((v.get('timeOffUsesAll') or {}).get('timeOffUses') or []):
            if x.get('userTimeOffRegisterEventId') is not None:
                owner[str(x['userTimeOffRegisterEventId'])] = uid
    rows, unmatched = [], 0
    for h in hist:
        eid = _s(h.get('플렉스 ID'))
        if not eid:
            continue
        uid = owner.get(eid)
        if not uid:
            unmatched += 1
        start = _d(h.get('시작'))
        if not start:
            continue
        days = h.get('일수')
        rows.append({
            'flex_event_id': eid, 'person_flex_id': uid, 'name': _s(h.get('이름')) or '',
            'type_label': _s(h.get('종류')), 'start': start, 'end': _d(h.get('끝')) or start,
            'days': float(days) if isinstance(days, (int, float)) else None, 'unit': _s(h.get('단위')),
            'status': _s(h.get('상태')), 'reason': _s(h.get('사유(메모)')), 'approval': _s(h.get('승인')),
            'comment': _s(h.get('댓글')), 'requested_at': _s(h.get('신청 시각')),
        })
    return rows, unmatched


def build_details_and_docs(flex_dir: Path, raw: Path):
    """원본 JSON 의 경력·학력·가족 + 직원서류 파일(11/6 내보내기 첨부파일/직원서류/<이름>_<플렉스ID>/).
    반환: ({플렉스ID: (경력, 학력, 가족)}, [서류], 받아 둔 파일이 없는 서류 수)."""
    import os
    flex_json = json.load(open(raw, encoding='utf-8'))
    types = {t['userDocumentIdHash']: t['name'] for t in (flex_json['company'].get('userDocumentTypes') or [])}
    base = flex_dir / '첨부파일' / '직원서류'
    folders = {}
    if base.exists():
        for sub in os.listdir(base):
            if '_' in sub:
                folders[sub.rsplit('_', 1)[1]] = base / sub
    details, docs, missing = {}, [], 0
    for fid, u in (flex_json.get('users') or {}).items():
        cs = (u.get('workExperiences') or {}).get('workExperienceResponses') or []
        es = (u.get('educationExperiences') or {}).get('educationExperienceResponses') or []
        fs = (u.get('dependentFamilies') or {}).get('responses') or []
        if cs or es or fs:
            details[fid] = (cs, es, fs)
        folder = folders.get(fid)
        local = os.listdir(folder) if folder else []
        for f in u.get('documentFiles') or []:
            name = f.get('fileName') or ''
            hit = next((x for x in local if x == name or x.endswith('_' + name) or x.endswith(name)), None)
            if not hit:
                missing += 1
                continue
            docs.append({'flex_id': fid, 'type': types.get(f.get('userDocumentIdHash'), '기타'), 'file_name': name,
                         'file_key': f.get('fileKey'), 'path': str(folder / hit)})
    return details, docs, missing


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--apply', action='store_true')
    ap.add_argument('--db', help='쓸 사용자 DB 경로(기본: 운영 instance/users.db)')
    ap.add_argument('--flex-dir', default=str(DEFAULT_FLEX_DIR))
    ap.add_argument('--summary', help='요약 엑셀(기본: 가장 최근 플렉스_요약_*.xlsx)')
    ap.add_argument('--raw', help='원본 JSON(기본: 가장 최근 flex_refresh_*/flex_data_*)')
    args = ap.parse_args()

    try:
        from dotenv import load_dotenv
        load_dotenv(ROOT / '.env')
    except Exception:
        pass

    from dashboard.services import groupware as gw
    from dashboard.services import hr
    if args.db:
        from dashboard.utils.user_database import UserDatabase
        d = UserDatabase(db_path=args.db)
        hr._db = lambda: d
        gw._db = lambda: d
        hr._table_ready = False
        gw._table_ready = False

    # 휴가 장부의 이름 → 슬랙 ID
    leave_slack = {}
    try:
        with hr._db()._get_connection() as conn:
            for r in conn.execute('SELECT name, slack_id FROM leave_staff'):
                if r[1]:
                    leave_slack[r[0]] = r[1]
    except Exception as e:
        print(f'(휴가 장부 없음 — 슬랙 ID 비움: {e})')

    fdir = Path(args.flex_dir)
    summary = Path(args.summary) if args.summary else summary_path(fdir)
    raw = Path(args.raw) if args.raw else raw_path(fdir)
    print(f'읽는 파일: {summary.name} + {raw.name}')
    records, skipped_appts = build_records(fdir, leave_slack, summary, raw)
    key_ok = bool(os.getenv('GW_HR_KEY', '').strip())
    by_status = Counter(r['status'] for r in records)
    print(f'구성원 {len(records)}명 — ' + ' · '.join(f'{k} {v}' for k, v in by_status.items()))
    print(f'주민번호 있음 {sum(1 for r in records if r["ssn"])}명 · 암호화 키 {"있음" if key_ok else "없음"}')
    print(f'발령 이력 {sum(len(r["_appts"]) for r in records)}건 (이름이 겹치거나 없는 행 {skipped_appts}건 건너뜀)')
    print()
    print(f'{"이름":<6} {"상태":<3} {"부서":<10} {"직급":<6} {"접근":<6} 로그인 메일 / 슬랙 / 주민번호 / 발령')
    for r in records:
        print(f'{r["name"]:<6} {r["status"]:<3} {(r["department"] or "-"):<10} {(r["job_rank"] or "-"):<6} '
              f'{r["access"]:<6} {r["login_email"] or "(없음)"} / {"O" if r["slack_id"] else "-"} / '
              f'{"O" if r["ssn"] else "-"} / {len(r["_appts"])}')

    history, unmatched = build_leave_history(fdir, summary, raw)
    from collections import Counter as _C
    print(f'\n휴가 기록 {len(history)}건 (직원 못 찾음 {unmatched}건) — '
          + ' · '.join(f'{k} {v}' for k, v in _C(h['status'] for h in history).most_common()))

    details, docs, missing_docs = build_details_and_docs(fdir, raw)
    print(f'경력 {sum(len(x[0]) for x in details.values())}건 · 학력 {sum(len(x[1]) for x in details.values())}건 · '
          f'가족 {sum(len(x[2]) for x in details.values())}명(주민번호 {sum(1 for x in details.values() for f in x[2] if f.get("ssn"))}명) · '
          f'서류 {len(docs)}개(받아 둔 파일 없음 {missing_docs}개)')

    if not args.apply:
        print('\n미리 보기만 했습니다. 쓰려면 --apply')
        return
    if any(r['ssn'] for r in records) and not key_ok:
        sys.exit('\nGW_HR_KEY 가 없어 멈춤 — 주민번호를 평문으로 저장하지 않습니다. .env 에 키를 넣고 다시 돌리세요.')
    result = Counter()
    for r in records:
        appts = r.pop('_appts')
        result[hr.upsert_from_flex(r, appts)] += 1
    print(f'\n저장: 새로 {result["insert"]} · 고침 {result["update"]} · 그대로 {result["skip"]}')

    # 휴가 기록 보관(인사카드 번호 연결)
    from dashboard.services import leave_view
    if args.db:
        leave_view._db = lambda: d
        leave_view._table_ready = False
    with hr._db()._get_connection() as conn:
        pid_by_flex = {r[0]: r[1] for r in conn.execute('SELECT flex_id, id FROM hr_people WHERE flex_id IS NOT NULL')}
    for h in history:
        h['person_id'] = pid_by_flex.get(h.pop('person_flex_id'))
    n = leave_view.upsert_flex_history(history)
    print(f'휴가 기록 보관: {n}건 (인사카드 연결 {sum(1 for h in history if h["person_id"])}건)')

    # 경력·학력·가족(출처 flex 줄만 다시 씀) + 서류 파일(같은 플렉스 파일은 한 번만)
    for fid, (cs, es, fs) in details.items():
        pid = pid_by_flex.get(fid)
        if pid:
            hr.import_flex_details(pid, cs, es, fs)
    saved = skipped = 0
    for doc in docs:
        pid = pid_by_flex.get(doc['flex_id'])
        if not pid:
            continue
        r = hr.save_document(pid, doc['type'], doc['file_name'], open(doc['path'], 'rb').read(), 'flex-import',
                             flex_file_key=doc['file_key'])
        skipped += r['skipped']
        saved += not r['skipped']
    print(f'경력·학력·가족 반영 {len(details)}명 · 서류 새로 {saved}개 · 이미 있음 {skipped}개 → {hr.docs_dir()}')


if __name__ == '__main__':
    main()
