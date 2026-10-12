"""시공사례 게시글 접수 블루프린트 — 완료 사진·작업지시서 업로드 → 아침 자동 초안 (2026-10-12).

흐름:
  1. 페이지(/case-posts)에서 프로젝트 선택 + 사진·작업지시서·메모 업로드 → [접수] (status=queued)
  2. 평일 07:20(+분산 지연) Claude 예약 작업(case-post-morning-draft)이 queued 건을 construction-case-post
     스킬로 작성 → draft.json 저장 + status=drafted  (scripts/case_post_context.py 경유)
  3. 페이지에서 초안 확인·수정·복사 → 아임웹·블로그에 직접 게시 → [게시함] (status=posted)

저장: 이 PC 로컬 폴더(CASE_POST_ROOT/접수/{job_id}/) = Claude 가 읽는 원본.
      + 프로젝트 Drive 폴더의 '시공사례' 하위 폴더에 백업(폴더 경로 없으면 백업만 skip).
자동 발행 없음. 시트는 읽기만 한다.

⚠️ 상태변경 엔드포인트는 '/api/' 접두사 필수(보안 미들웨어 세션 인증 분기).
   본문·메모는 FormData 로 받는다 — JSON 입력검증(문자열 패턴·5000자)에 글 본문이 걸리지 않게.
권한: 지금은 관리자만(admin_required).
"""
import json
import os
import re
import threading
from datetime import datetime
from pathlib import Path

from flask import Blueprint, render_template, request, session, jsonify, send_from_directory, abort

from ..auth import admin_required
from ..utils.logging_config import get_logger

logger = get_logger(__name__)

case_posts_bp = Blueprint('case_posts', __name__)

# Flask 서비스는 LocalSystem 으로 돌아 Path.home() 이 systemprofile → 절대경로 기본값.
CASE_POST_ROOT = Path(os.environ.get('CASE_POST_ROOT')
                      or r'C:\Users\SECOM\Documents\시공사례_게시글')
INBOX = CASE_POST_ROOT / '접수'
DRIVE_SUBFOLDER = '시공사례'

STATUS_LABELS = {
    'uploading': '사진 보관',
    'queued': '접수',
    'drafted': '초안 완료',
    'posted': '게시함',
    'cancelled': '취소',
}
ALLOWED_EXT = {'.jpg', '.jpeg', '.png', '.webp', '.heic', '.pdf'}
_JOB_ID_RE = re.compile(r'^\d{8}-\d{6}_[A-Za-z0-9-]{1,30}$')
_FILE_RE = re.compile(r'^(photo|order)_\d{2,3}\.[a-z]{3,4}$')
_CODE_RE = re.compile(r'^[A-Za-z0-9-]{2,30}$')

_job_lock = threading.Lock()


# ── 로컬 job 저장소 ────────────────────────────────────────────────

def _job_dir(job_id: str) -> Path:
    if not _JOB_ID_RE.match(job_id or ''):
        abort(404)
    d = INBOX / job_id
    if not (d / 'job.json').is_file():
        abort(404)
    return d


def _read_json(path: Path):
    try:
        return json.loads(path.read_text(encoding='utf-8'))
    except FileNotFoundError:
        return None


def _write_json(path: Path, data) -> None:
    """원자적 저장 (임시파일 → replace). 예약 작업 스크립트와 같은 방식."""
    tmp = path.with_suffix(path.suffix + '.tmp')
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    os.replace(tmp, path)


def _update_job(job_id: str, **fields) -> dict:
    d = _job_dir(job_id)
    with _job_lock:
        job = _read_json(d / 'job.json') or {}
        job.update(fields)
        _write_json(d / 'job.json', job)
    return job


def _current_email() -> str:
    return ((session.get('user') or {}).get('email') or '').strip()


def _now() -> str:
    return datetime.now().strftime('%Y-%m-%d %H:%M')


def _job_summary(job: dict) -> dict:
    files = job.get('files') or []
    return {
        'id': job.get('id'),
        'code': job.get('code'),
        'site': job.get('site') or '',
        'status': job.get('status'),
        'status_label': STATUS_LABELS.get(job.get('status'), job.get('status')),
        'created_at': job.get('created_at'),
        'drafted_at': job.get('drafted_at'),
        'posted_at': job.get('posted_at'),
        'photo_count': sum(1 for f in files if f.get('kind') == 'photo'),
        'order_count': sum(1 for f in files if f.get('kind') == 'order'),
    }


# ── 프로젝트 조회 (읽기 전용) ─────────────────────────────────────

def _project_brief(r: dict) -> dict:
    return {
        'code': (r.get('프로젝트 코드') or '').strip(),
        'site': str(r.get('현장 주소') or '').strip(),
        'biz': str(r.get('사업자명') or '').strip(),
        'kind': str(r.get('공사 구분') or '').strip(),
        'start': str(r.get('공사 시작') or '').strip(),
        'end': str(r.get('공사 종료') or '').strip(),
        'lead_no': str(r.get('Lead No') or '').strip(),
    }


def _find_project(code: str):
    from ..services.project_service import get_project_records
    for r in (get_project_records() or []):
        if (r.get('프로젝트 코드') or '').strip() == code:
            return r
    return None


# ── Drive 백업 ────────────────────────────────────────────────────

def _drive_backup(job_id: str, code: str, filename: str, content: bytes, mimetype: str) -> str:
    """프로젝트 Drive 폴더/시공사례/ 에 백업. 성공 시 '', 실패 시 사유 문자열."""
    try:
        from ..services.business_license_handler import _project_folder_id_for
        from ..utils.google_drive import find_or_create_folder, upload_file
        job = _read_json(INBOX / job_id / 'job.json') or {}
        folder_id = job.get('drive_folder_id')
        if not folder_id:
            project_folder = _project_folder_id_for(code)
            if not project_folder:
                return '프로젝트 폴더 경로가 시트에 없어 드라이브 백업을 건너뜀'
            folder_id = (find_or_create_folder(DRIVE_SUBFOLDER, project_folder) or {}).get('id')
            if not folder_id:
                return '드라이브 시공사례 폴더 생성 실패'
            _update_job(job_id, drive_folder_id=folder_id)
        res = upload_file(folder_id, f'{job_id[:8]}_{filename}', content, mimetype)
        return '' if res else '드라이브 업로드 실패'
    except Exception as exc:
        logger.warning(f'[CASE_POST] 드라이브 백업 실패 ({job_id}/{filename}): {exc}')
        return '드라이브 업로드 오류'


# ── 페이지 ────────────────────────────────────────────────────────

@case_posts_bp.route('/case-posts')
@admin_required
def case_posts_page():
    return render_template('case_posts.html', user=session.get('user', {}))


# ── API ───────────────────────────────────────────────────────────

@case_posts_bp.route('/api/case-posts/projects', methods=['GET'])
@admin_required
def api_search_projects():
    """코드·현장 주소·사업자명 일부로 검색. 빈 검색어면 최근 공사 15건."""
    from ..services.project_service import get_project_records
    q = (request.args.get('q') or '').strip().lower()
    rows = [_project_brief(r) for r in (get_project_records() or [])]
    rows = [r for r in rows if r['code']]
    if q:
        rows = [r for r in rows
                if q in r['code'].lower() or q in r['site'].lower() or q in r['biz'].lower()]
    else:
        today = datetime.now().strftime('%Y-%m-%d')
        rows = [r for r in rows if r['start'] and r['start'] <= today]
        rows.sort(key=lambda r: (r['end'] or r['start']), reverse=True)
    return jsonify({'success': True, 'items': rows[-15:][::-1] if q else rows[:15]})


@case_posts_bp.route('/api/case-posts/jobs', methods=['GET'])
@admin_required
def api_list_jobs():
    jobs = []
    if INBOX.is_dir():
        for d in INBOX.iterdir():
            job = _read_json(d / 'job.json') if d.is_dir() else None
            if job:
                jobs.append(_job_summary(job))
    jobs.sort(key=lambda j: j.get('id') or '', reverse=True)
    return jsonify({'success': True, 'items': jobs})


@case_posts_bp.route('/api/case-posts/jobs', methods=['POST'])
@admin_required
def api_create_job():
    """접수 시작 — 폴더·job.json 생성 (status=uploading, 예약 작업은 아직 안 가져감)."""
    code = (request.form.get('code') or '').strip()
    memo = (request.form.get('memo') or '').strip()[:2000]
    if not _CODE_RE.match(code):
        return jsonify({'success': False, 'error': '프로젝트 코드를 선택해 주세요.'}), 400
    project = _find_project(code)
    if not project:
        return jsonify({'success': False, 'error': f'공사현황에서 {code} 를 찾지 못했습니다.'}), 404
    brief = _project_brief(project)
    job_id = f"{datetime.now().strftime('%Y%m%d-%H%M%S')}_{code}"
    d = INBOX / job_id
    (d / 'files').mkdir(parents=True, exist_ok=False)
    job = {
        'id': job_id,
        'code': code,
        'site': brief['site'],
        'lead_no': brief['lead_no'],
        'memo': memo,
        'status': 'uploading',
        'created_at': _now(),
        'created_by': _current_email(),
        'files': [],
        'drive_folder_id': '',
        'drive_errors': [],
    }
    _write_json(d / 'job.json', job)
    logger.info(f'[CASE_POST] 접수 시작 {job_id} by {job["created_by"]}')
    return jsonify({'success': True, 'job': _job_summary(job)})


@case_posts_bp.route('/api/case-posts/by-project/<code>', methods=['GET'])
@admin_required
def api_job_by_project(code):
    """PM 아코디언 '시공 완료 사진' 칸 — 그 프로젝트의 가장 최근 접수(취소 제외) + 파일 목록."""
    code = (code or '').strip()
    if not _CODE_RE.match(code):
        return jsonify({'success': True, 'job': None})
    latest = None
    if INBOX.is_dir():
        for d in INBOX.glob(f'*_{code}'):
            job = _read_json(d / 'job.json') if d.is_dir() else None
            if job and job.get('code') == code and job.get('status') != 'cancelled':
                if latest is None or job['id'] > latest['id']:
                    latest = job
    if not latest:
        return jsonify({'success': True, 'job': None})
    return jsonify({'success': True, 'job': {**_job_summary(latest), 'files': latest.get('files', []),
                                             'drive_errors': latest.get('drive_errors', [])}})


@case_posts_bp.route('/api/case-posts/jobs/<job_id>/files', methods=['POST'])
@admin_required
def api_upload_file(job_id):
    """파일 한 개씩 업로드 (요청 크기 제한 8MB 회피 — 프론트에서 사진 축소 후 한 장씩)."""
    d = _job_dir(job_id)
    kind = (request.form.get('kind') or '').strip()
    f = request.files.get('file')
    if kind not in ('photo', 'order') or not f:
        return jsonify({'success': False, 'error': '파일 종류가 올바르지 않습니다.'}), 400
    ext = Path(f.filename or '').suffix.lower() or '.jpg'
    if ext == '.jpeg':
        ext = '.jpg'
    if ext not in ALLOWED_EXT:
        return jsonify({'success': False, 'error': f'{ext} 파일은 올릴 수 없습니다(사진·PDF만).'}), 400
    content = f.read()
    if not content:
        return jsonify({'success': False, 'error': '빈 파일입니다.'}), 400

    with _job_lock:
        job = _read_json(d / 'job.json') or {}
        if job.get('status') not in ('uploading', 'queued'):
            return jsonify({'success': False, 'error': '이미 초안이 작성된 건에는 파일을 더할 수 없습니다.'}), 409
        n = sum(1 for x in job.get('files', []) if x.get('kind') == kind) + 1
        name = f'{kind}_{n:02d}{ext}'
        (d / 'files' / name).write_bytes(content)
        job.setdefault('files', []).append({
            'name': name, 'kind': kind, 'original': (f.filename or '')[:200], 'size': len(content),
        })
        _write_json(d / 'job.json', job)

    drive_err = _drive_backup(job_id, job['code'], name, content, f.mimetype or 'application/octet-stream')
    if drive_err:
        with _job_lock:
            job = _read_json(d / 'job.json') or {}
            errs = job.get('drive_errors') or []
            if drive_err not in errs:
                errs.append(drive_err)
            job['drive_errors'] = errs
            _write_json(d / 'job.json', job)
    return jsonify({'success': True, 'name': name, 'drive_error': drive_err})


@case_posts_bp.route('/api/case-posts/jobs/<job_id>/submit', methods=['POST'])
@admin_required
def api_submit_job(job_id):
    d = _job_dir(job_id)
    job = _read_json(d / 'job.json') or {}
    if not any(x.get('kind') == 'photo' for x in job.get('files', [])):
        return jsonify({'success': False, 'error': '사진을 한 장 이상 올려 주세요.'}), 400
    if job.get('status') != 'uploading':
        return jsonify({'success': True, 'job': _job_summary(job)})
    job = _update_job(job_id, status='queued', queued_at=_now())
    logger.info(f'[CASE_POST] 접수 완료 {job_id} (사진 {_job_summary(job)["photo_count"]}장)')
    return jsonify({'success': True, 'job': _job_summary(job)})


@case_posts_bp.route('/api/case-posts/jobs/<job_id>', methods=['GET'])
@admin_required
def api_job_detail(job_id):
    d = _job_dir(job_id)
    job = _read_json(d / 'job.json') or {}
    draft = _read_json(d / 'draft.json')
    return jsonify({
        'success': True,
        'job': {**_job_summary(job), 'memo': job.get('memo', ''), 'lead_no': job.get('lead_no', ''),
                'files': job.get('files', []), 'drive_errors': job.get('drive_errors', []),
                'draft_error': job.get('draft_error', '')},
        'draft': draft,
    })


@case_posts_bp.route('/api/case-posts/jobs/<job_id>/files/<name>', methods=['GET'])
@admin_required
def api_job_file(job_id, name):
    d = _job_dir(job_id)
    if not _FILE_RE.match(name or ''):
        abort(404)
    return send_from_directory(d / 'files', name, max_age=3600)


@case_posts_bp.route('/api/case-posts/jobs/<job_id>/draft', methods=['POST'])
@admin_required
def api_save_draft(job_id):
    """초안 수정 저장. 처음 수정 시 Claude 원본을 draft.orig.json 으로 보존."""
    d = _job_dir(job_id)
    raw = request.form.get('draft') or ''
    try:
        new = json.loads(raw)
        assert isinstance(new, dict)
    except Exception:
        return jsonify({'success': False, 'error': '초안 형식이 올바르지 않습니다.'}), 400
    with _job_lock:
        cur = _read_json(d / 'draft.json')
        if cur is None:
            return jsonify({'success': False, 'error': '아직 초안이 없습니다.'}), 409
        if not (d / 'draft.orig.json').exists():
            _write_json(d / 'draft.orig.json', cur)
        for sec in ('homepage', 'blog'):
            if isinstance(new.get(sec), dict):
                cur.setdefault(sec, {})
                for k, v in new[sec].items():
                    if k in ('title', 'meta', 'body', 'hashtags') and isinstance(v, str):
                        cur[sec][k] = v[:20000]
        cur['edited_at'] = _now()
        _write_json(d / 'draft.json', cur)
    return jsonify({'success': True})


@case_posts_bp.route('/api/case-posts/jobs/<job_id>/status', methods=['POST'])
@admin_required
def api_set_status(job_id):
    """게시함 / 게시 취소(초안 완료로) / 접수 취소 / 다시 작성(queued)."""
    d = _job_dir(job_id)
    action = (request.form.get('action') or '').strip()
    job = _read_json(d / 'job.json') or {}
    st = job.get('status')
    if action == 'posted' and st == 'drafted':
        job = _update_job(job_id, status='posted', posted_at=_now())
    elif action == 'unposted' and st == 'posted':
        job = _update_job(job_id, status='drafted', posted_at=None)
    elif action == 'cancel' and st in ('uploading', 'queued', 'drafted'):
        job = _update_job(job_id, status='cancelled', cancelled_at=_now())
    elif action == 'requeue' and st in ('drafted', 'cancelled'):
        job = _update_job(job_id, status='queued', queued_at=_now(), draft_error='')
    else:
        return jsonify({'success': False, 'error': '지금 상태에서는 할 수 없는 작업입니다.'}), 409
    return jsonify({'success': True, 'job': _job_summary(job)})
