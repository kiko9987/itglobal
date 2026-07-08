"""사업자등록증 스레드 첨부 → Google Drive 자동 저장.

흐름:
1. 슬랙 #공사_확정 채널 카드 스레드에 파일 첨부 이벤트 도착
2. project_thread:{channel}|{thread_ts} 매핑으로 프로젝트 코드 조회
3. 프로젝트 시트에서 '견적서 및 계약서 폴더 경로' 값(폴더 ID) 조회
4. 그 폴더 안에 '사업자등록증/' 하위 폴더 생성 (없으면)
5. 파일 다운로드 (Slack API) → Drive 업로드
   - 기존 '사업자등록증.{ext}' 있으면 '사업자등록증_{N}.{ext}' 로 rename 후 새 파일을
     '사업자등록증.{ext}' 로 저장. 즉 최신이 항상 '사업자등록증.{ext}'.
6. 완료 후 스레드에 성공 답글 (매니저 안내)

계산서 요청 시 verify_license_exists()로 사업자등록증.{ext} 존재 여부 검증.

Shared Drive 사용 중이라 모든 Drive API 호출에 supportsAllDrives=True 필수.
"""

from __future__ import annotations

import io
import logging
import os
import re
import time
from typing import Optional, Tuple

import urllib.request

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload

logger = logging.getLogger(__name__)

LICENSE_FOLDER_NAME = '사업자등록증'
LICENSE_BASENAME = '사업자등록증'
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive']

_drive_service = None


def _get_drive():
    global _drive_service
    if _drive_service is None:
        creds_path = os.getenv('GOOGLE_APPLICATION_CREDENTIALS', 'credentials.json')
        credentials = service_account.Credentials.from_service_account_file(
            creds_path, scopes=DRIVE_SCOPES,
        )
        _drive_service = build('drive', 'v3', credentials=credentials, cache_discovery=False)
    return _drive_service


def _project_folder_id_for(code: str) -> Optional[str]:
    """프로젝트 코드 → 시트의 '견적서 및 계약서 폴더 경로' 값. 폴더 ID 형식이면 반환."""
    from dashboard.services.project_service import get_project_records
    records = get_project_records() or []
    for r in records:
        if (r.get('프로젝트 코드') or '').strip() == code:
            raw = str(r.get('견적서 및 계약서 폴더 경로') or '').strip()
            if raw and re.match(r'^[a-zA-Z0-9_-]{20,}$', raw):
                return raw
            return None
    return None


def _find_license_subfolder(drive, parent_id: str) -> Optional[str]:
    """부모 폴더 안 '사업자등록증' 하위 폴더 ID (없으면 None)."""
    q = (
        f"'{parent_id}' in parents and trashed=false "
        f"and mimeType='application/vnd.google-apps.folder' "
        f"and name='{LICENSE_FOLDER_NAME}'"
    )
    resp = drive.files().list(
        q=q,
        fields='files(id,name)',
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    files = resp.get('files', [])
    return files[0]['id'] if files else None


def _create_license_subfolder(drive, parent_id: str) -> str:
    """부모 폴더 안에 '사업자등록증' 하위 폴더 생성 → id 반환."""
    body = {
        'name': LICENSE_FOLDER_NAME,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id],
    }
    resp = drive.files().create(body=body, fields='id', supportsAllDrives=True).execute()
    return resp['id']


def _get_or_create_license_subfolder(drive, parent_id: str) -> str:
    fid = _find_license_subfolder(drive, parent_id)
    if fid:
        return fid
    logger.info(f'[LICENSE] 사업자등록증 폴더 생성: parent={parent_id}')
    return _create_license_subfolder(drive, parent_id)


def _list_folder_files(drive, folder_id: str) -> list:
    """폴더 내 non-folder 파일 목록."""
    q = (
        f"'{folder_id}' in parents and trashed=false "
        f"and mimeType!='application/vnd.google-apps.folder'"
    )
    resp = drive.files().list(
        q=q,
        fields='files(id,name)',
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        pageSize=100,
    ).execute()
    return resp.get('files', [])


def _canonical_name(code: str, ext: str) -> str:
    """{프로젝트코드} 사업자등록증.{ext} 형식.
    (2026-07-09 사용자 확정: 파일이 폴더 밖으로 나가도 어느 프로젝트인지 즉시 인지)
    """
    return f'{code} {LICENSE_BASENAME}.{ext}'


def _next_copy_index(files: list, code: str, ext: str) -> int:
    """이미 존재하는 '{code} 사업자등록증_{N}.{ext}' 중 최댓값+1. 없으면 1."""
    prefix = f'{code} {LICENSE_BASENAME}'
    pat = re.compile(rf'^{re.escape(prefix)}_(\d+)\.{re.escape(ext)}$')
    max_n = 0
    for f in files:
        m = pat.match(f['name'])
        if m:
            try:
                max_n = max(max_n, int(m.group(1)))
            except Exception:
                pass
    return max_n + 1


def _download_slack_file(url: str, slack_bot_token: str) -> bytes:
    """Slack file URL → bytes. private URL이라 봇 토큰 Authorization 필요."""
    req = urllib.request.Request(
        url,
        headers={'Authorization': f'Bearer {slack_bot_token}'},
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read()


def _guess_ext(filename: str, mimetype: str) -> str:
    """확장자 추출 (소문자)."""
    if filename and '.' in filename:
        ext = filename.rsplit('.', 1)[1].lower()
        if ext:
            return ext
    # mimetype fallback
    mime_ext = {
        'image/jpeg': 'jpg',
        'image/png': 'png',
        'image/heic': 'heic',
        'application/pdf': 'pdf',
    }
    return mime_ext.get((mimetype or '').lower(), 'bin')


def save_business_license(code: str, file_bytes: bytes, filename: str, mimetype: str) -> dict:
    """사업자등록증 파일을 프로젝트 폴더에 저장.

    Returns:
        {'ok': bool, 'reason': str, 'file_name': str, 'file_id': str}
    """
    parent = _project_folder_id_for(code)
    if not parent:
        return {'ok': False, 'reason': 'no_project_folder'}

    drive = _get_drive()
    license_folder = _get_or_create_license_subfolder(drive, parent)

    ext = _guess_ext(filename, mimetype)
    canonical = _canonical_name(code, ext)

    # 저장 파일명 규칙 (2026-07-09 사용자 확정):
    #   원본이 없으면 '{code} 사업자등록증.{ext}' (canonical)
    #   원본이 이미 있으면 '{code} 사업자등록증_{N}.{ext}' (N=1,2,3...). 원본은 유지.
    # 계산서 요청 검증은 항상 canonical 존재 여부만 확인.
    existing = _list_folder_files(drive, license_folder)
    canonical_exists = any(f['name'] == canonical for f in existing)
    if canonical_exists:
        next_n = _next_copy_index(existing, code, ext)
        save_name = f'{code} {LICENSE_BASENAME}_{next_n}.{ext}'
    else:
        save_name = canonical

    # 새 파일 업로드
    media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype=mimetype or 'application/octet-stream')
    body = {
        'name': save_name,
        'parents': [license_folder],
    }
    up = drive.files().create(
        body=body,
        media_body=media,
        fields='id,name',
        supportsAllDrives=True,
    ).execute()
    logger.info(f'[LICENSE] 저장 완료: {up["name"]} (project={code}, id={up["id"]})')
    return {
        'ok': True,
        'reason': 'saved',
        'file_name': up['name'],
        'file_id': up['id'],
    }


def verify_license_exists(code: str) -> bool:
    """프로젝트에 '사업자등록증.{ext}' 파일이 하나라도 있으면 True.

    계산서 요청 검증에 사용. 파일 이름의 확장자 부분은 상관없지만 basename이 정확히
    '사업자등록증'이어야 함(백업본 '사업자등록증_1' 등은 인정 안 함).
    """
    parent = _project_folder_id_for(code)
    if not parent:
        return False
    drive = _get_drive()
    fid = _find_license_subfolder(drive, parent)
    if not fid:
        return False
    # 2026-07-09 규칙: canonical = '{code} 사업자등록증.{ext}'.
    # 확장자 무관 basename 매치 (같은 프로젝트에 pdf/png 등 여러 확장자 원본 허용).
    expected_base = f'{code} {LICENSE_BASENAME}'
    for f in _list_folder_files(drive, fid):
        name = f['name']
        if '.' not in name:
            continue
        base = name.rsplit('.', 1)[0]
        if base == expected_base:
            return True
    return False


def resolve_project_from_thread(channel: str, thread_ts: str) -> Optional[str]:
    """channel|thread_ts → 프로젝트 코드 (없으면 None)."""
    if not channel or not thread_ts:
        return None
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        v = rc.get(f'project_thread:{channel}|{thread_ts}')
        if isinstance(v, bytes):
            v = v.decode('utf-8')
        return v or None
    except Exception as exc:
        logger.warning(f'[LICENSE] 스레드→프로젝트 조회 실패: {exc}')
        return None


def handle_thread_file_share(event: dict, slack_bot_token: str) -> Optional[dict]:
    """슬랙 message.file_share 이벤트 진입점.

    event 안의 files 배열을 순회하며 각 이미지/PDF를 사업자등록증으로 저장.
    스레드가 프로젝트 카드가 아니거나 파일이 지원 안 되는 확장자면 skip.

    Returns:
        {'code': str, 'saved': [file_names], 'skipped': [reasons]} or None if not relevant.
    """
    channel = event.get('channel')
    thread_ts = event.get('thread_ts')
    if not thread_ts:
        return None  # 스레드 답글 아님

    code = resolve_project_from_thread(channel, thread_ts)
    if not code:
        return None  # 프로젝트 카드 스레드 아님

    files = event.get('files') or []
    if not files:
        return None

    saved = []
    skipped = []
    for f in files:
        name = f.get('name') or ''
        mimetype = f.get('mimetype') or ''
        url = f.get('url_private_download') or f.get('url_private')
        if not url:
            skipped.append(f'{name}: url 없음')
            continue

        # 지원 확장자: 이미지 or PDF 만
        ext = _guess_ext(name, mimetype).lower()
        if ext not in {'jpg', 'jpeg', 'png', 'heic', 'gif', 'pdf', 'webp'}:
            skipped.append(f'{name}: 지원 안 되는 확장자({ext})')
            continue

        try:
            content = _download_slack_file(url, slack_bot_token)
            res = save_business_license(code, content, name, mimetype)
            if res.get('ok'):
                saved.append(res['file_name'])
            else:
                skipped.append(f'{name}: 저장 실패({res.get("reason")})')
        except Exception as exc:
            logger.error(f'[LICENSE] 처리 예외 ({name}): {exc}', exc_info=True)
            skipped.append(f'{name}: 예외 {type(exc).__name__}')

    # 하나라도 저장 성공했으면 원본 공사 확정 카드의 사업자등록증 배지를 ✅로 갱신.
    if saved:
        try:
            from dashboard.services.project_slack_notifier import refresh_project_card_license
            refresh_project_card_license(code)
        except Exception as exc:
            logger.warning(f'[LICENSE] 원본 카드 갱신 실패 ({code}): {exc}')

    return {'code': code, 'saved': saved, 'skipped': skipped, 'thread_ts': thread_ts, 'channel': channel}
