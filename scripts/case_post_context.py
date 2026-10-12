"""시공사례 아침 초안 작성 도우미 — Claude 예약 작업(case-post-morning-draft) 전용 (2026-10-12).

PM 페이지(/case-posts)에서 접수된 건을 읽고, 작성한 초안을 돌려놓는다.
⚠️ create_app() 호출 금지(운영 스케줄러 이중 기동). 시트는 읽기만 한다.

사용:
  python scripts/case_post_context.py list                 # 접수(queued) 건 id 목록
  python scripts/case_post_context.py show <job_id>        # 접수 정보 + 파일 경로 + 공사현황·리드 칸
  python scripts/case_post_context.py save <job_id> <draft.json 경로>   # 초안 저장 + 상태 '초안 완료'
  python scripts/case_post_context.py fail <job_id> "사유"  # 작성 못한 사유 기록(상태는 접수 유지)

draft.json 형식:
{
  "homepage": {"title": str, "meta": str, "body": str,
               "photo_order": [{"file": "photo_01.jpg", "caption": str}, ...]},
  "blog": {"title": str, "body": str, "hashtags": "#a #b"},
  "mask_photos": [{"file": "photo_03.jpg", "reason": "간판 상호 보임"}],
  "check": ["..."],
  "after": ["..."]
}
"""
import contextlib
import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(os.environ.get('CASE_POST_ROOT') or r'C:\Users\SECOM\Documents\시공사례_게시글')
INBOX = ROOT / '접수'
_JOB_ID_RE = re.compile(r'^\d{8}-\d{6}_[A-Za-z0-9-]{1,30}$')

PROJECT_FIELDS = ['프로젝트 코드', '현장 주소', '공사 구분', '기계 분류', '브랜드', '공사 시작', '공사 종료',
                  '공사 내용', '유입 구분', 'Lead No']
LEAD_FIELDS = ['상담 시간', '플랫폼', '방문 주소', '문의 내용', '상담 내용', '키워드']


def _read(path: Path):
    return json.loads(path.read_text(encoding='utf-8'))


def _write(path: Path, data) -> None:
    tmp = path.with_suffix(path.suffix + '.tmp')
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(tmp, path)


def _job_dir(job_id: str) -> Path:
    if not _JOB_ID_RE.match(job_id):
        sys.exit(f'잘못된 job_id: {job_id}')
    d = INBOX / job_id
    if not (d / 'job.json').is_file():
        sys.exit(f'접수 건 없음: {job_id}')
    return d


def cmd_list():
    ids = []
    if INBOX.is_dir():
        for d in sorted(INBOX.iterdir()):
            try:
                if d.is_dir() and _read(d / 'job.json').get('status') == 'queued':
                    ids.append(d.name)
            except FileNotFoundError:
                continue
    print(json.dumps(ids, ensure_ascii=False))


def _sheet_context(code: str, lead_no_hint: str):
    # 서비스 import·시트 로딩 로그가 stdout(JSON 출력)에 섞이지 않게 stderr 로 돌린다
    with contextlib.redirect_stdout(sys.stderr):
        return _sheet_context_inner(code, lead_no_hint)


def _sheet_context_inner(code: str, lead_no_hint: str):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parents[1] / '.env')
    from dashboard.services.project_service import get_project_records
    from dashboard.services.lead_service import get_lead_by_no
    project = next((r for r in (get_project_records() or [])
                    if (r.get('프로젝트 코드') or '').strip() == code), None)
    p = {k: (project or {}).get(k) for k in PROJECT_FIELDS} if project else None
    lead_no = str((project or {}).get('Lead No') or lead_no_hint or '').strip()
    lead = get_lead_by_no(lead_no) if lead_no else None
    l = {k: lead.get(k) for k in LEAD_FIELDS} if lead else None
    return p, lead_no, l


def cmd_show(job_id: str):
    d = _job_dir(job_id)
    job = _read(d / 'job.json')
    project, lead_no, lead = _sheet_context(job['code'], job.get('lead_no', ''))
    out = {
        'job_id': job_id,
        'status': job.get('status'),
        'code': job['code'],
        'memo': job.get('memo', ''),
        'order_text': job.get('order_text', ''),  # 카톡 단체방 작업지시 글(붙여넣기) — 작업지시서 본문
        'files': [{'kind': f['kind'], 'name': f['name'], 'path': str(d / 'files' / f['name'])}
                  for f in job.get('files', [])],
        'project': project,
        'lead_no': lead_no,
        'lead': lead,
        'draft_output': str(d / 'draft.json'),
    }
    print(json.dumps(out, ensure_ascii=False, indent=2, default=str))


def _validate(draft: dict, file_names: set):
    errs = []
    hp, bl = draft.get('homepage') or {}, draft.get('blog') or {}
    for sec, obj, keys in (('homepage', hp, ('title', 'meta', 'body')), ('blog', bl, ('title', 'body', 'hashtags'))):
        for k in keys:
            if not isinstance(obj.get(k), str) or not obj.get(k).strip():
                errs.append(f'{sec}.{k} 비어 있음')
    for p in hp.get('photo_order') or []:
        if p.get('file') not in file_names:
            errs.append(f'photo_order 에 없는 파일: {p.get("file")}')
    for m in draft.get('mask_photos') or []:
        if m.get('file') not in file_names:
            errs.append(f'mask_photos 에 없는 파일: {m.get("file")}')
    for k in ('check', 'after'):
        if not isinstance(draft.get(k, []), list):
            errs.append(f'{k} 는 목록이어야 함')
    return errs


def _to_markdown(job: dict, draft: dict) -> str:
    hp, bl = draft['homepage'], draft['blog']
    lines = [f"# {job['code']} 시공사례 초안 ({job['id']})", '', '[홈페이지 시공사례]',
             f"제목: {hp['title']}", f"검색 설명문: {hp['meta']}", '본문:', hp['body'], '', '사진 순서:']
    lines += [f"{i}) {p['file']} — {p.get('caption', '')}" for i, p in enumerate(hp.get('photo_order') or [], 1)]
    lines += ['', '[네이버 블로그]', f"제목: {bl['title']}", '본문:', bl['body'], '', f"해시태그: {bl['hashtags']}",
              '', '[가려야 할 사진]']
    lines += [f"- {m['file']}: {m['reason']}" for m in draft.get('mask_photos') or []] or ['- 없음']
    lines += ['', '[확인 필요]'] + ([f'- {c}' for c in draft.get('check') or []] or ['- 없음'])
    lines += ['', '[올린 뒤 할 일]'] + ([f'- {a}' for a in draft.get('after') or []] or ['- 없음'])
    return '\n'.join(lines) + '\n'


def cmd_save(job_id: str, draft_path: str):
    d = _job_dir(job_id)
    job = _read(d / 'job.json')
    if job.get('status') != 'queued':
        sys.exit(f'접수 상태가 아님({job.get("status")}) — 저장 안 함')
    draft = _read(Path(draft_path))
    errs = _validate(draft, {f['name'] for f in job.get('files', [])})
    if errs:
        sys.exit('초안 형식 오류:\n- ' + '\n- '.join(errs))
    now = datetime.now().strftime('%Y-%m-%d %H:%M')
    draft['generated_at'] = now
    if (d / 'draft.json').exists():  # '다시 작성' 이전 초안 보존
        os.replace(d / 'draft.json', d / 'draft.prev.json')
    _write(d / 'draft.json', draft)
    md_name = f"{job['id'][:8]}_{job['code']}.md"
    (ROOT / md_name).write_text(_to_markdown(job, draft), encoding='utf-8')
    job = _read(d / 'job.json')  # 그사이 페이지에서 바뀐 값 반영
    if job.get('status') != 'queued':
        sys.exit(f'저장 중 상태가 바뀜({job.get("status")}) — draft.json 은 저장됨, 상태는 그대로 둠')
    job.update(status='drafted', drafted_at=now, draft_error='')
    _write(d / 'job.json', job)
    print(f'저장 완료: {d / "draft.json"} / {ROOT / md_name}')


def cmd_fail(job_id: str, reason: str):
    d = _job_dir(job_id)
    job = _read(d / 'job.json')
    job['draft_error'] = f"{datetime.now().strftime('%m-%d %H:%M')} {reason}"[:500]
    _write(d / 'job.json', job)
    print('기록함')


if __name__ == '__main__':
    a = sys.argv[1:]
    if a[:1] == ['list']:
        cmd_list()
    elif a[:1] == ['show'] and len(a) == 2:
        cmd_show(a[1])
    elif a[:1] == ['save'] and len(a) == 3:
        cmd_save(a[1], a[2])
    elif a[:1] == ['fail'] and len(a) == 3:
        cmd_fail(a[1], a[2])
    else:
        sys.exit(__doc__)
