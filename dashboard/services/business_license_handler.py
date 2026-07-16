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


# 최대 파일 크기 (사업자등록증은 사진·PDF 이므로 넉넉하게 50MB)
_MAX_LICENSE_BYTES = 50 * 1024 * 1024  # 50MB


def _download_slack_file(url: str, slack_bot_token: str) -> bytes:
    """Slack file URL → bytes. private URL이라 봇 토큰 Authorization 필요.

    - 2026-07-09 네트워크 timeout 재시도 (최대 3회 지수 백오프)
    - 2026-07-10 파일 크기 50MB 상한 (대용량 다운로드로 메모리·타임아웃 방지)
    """
    import http.client as _http_client
    import socket
    import time as _time

    last_exc = None
    for attempt in range(3):
        try:
            req = urllib.request.Request(
                url,
                headers={'Authorization': f'Bearer {slack_bot_token}'},
            )
            with urllib.request.urlopen(req, timeout=45) as r:
                # Content-Length 로 사전 차단
                try:
                    length = int(r.headers.get('Content-Length', '0'))
                except (TypeError, ValueError):
                    length = 0
                if length and length > _MAX_LICENSE_BYTES:
                    raise ValueError(
                        f'파일 크기 초과: {length / 1024 / 1024:.1f}MB > 50MB'
                    )
                # 스트림 읽기 상한 (헤더 누락 대비 실제 읽는 바이트도 체크)
                data = r.read(_MAX_LICENSE_BYTES + 1)
                if len(data) > _MAX_LICENSE_BYTES:
                    raise ValueError(
                        f'파일 크기 초과: >{_MAX_LICENSE_BYTES / 1024 / 1024:.0f}MB (실 스트림)'
                    )
                return data
        except ValueError:
            # 크기 초과는 재시도 무의미 — 즉시 propagate
            raise
        except (TimeoutError, socket.timeout, _http_client.IncompleteRead, ConnectionError) as exc:
            last_exc = exc
            wait = 1.0 * (attempt + 1)
            logger.warning(
                f'[LICENSE] Slack 파일 다운로드 네트워크 에러 ({type(exc).__name__}) '
                f'재시도 {attempt+1}/3 — {wait}s 후'
            )
            _time.sleep(wait)
    raise last_exc if last_exc else RuntimeError('Slack 파일 다운로드 실패')


# ─────────────────────────────────────────────────────────────
# 매직 넘버 (파일 시그니처) 검증 — 확장자 위장 방어 (2026-07-10)
# 사업자등록증에 허용되는 실제 파일 종류만 승인
# ─────────────────────────────────────────────────────────────
_MAGIC_SIGNATURES = [
    (b'\xff\xd8\xff', 'jpg'),                    # JPEG (JFIF/EXIF/etc)
    (b'\x89PNG\r\n\x1a\n', 'png'),               # PNG
    (b'%PDF-', 'pdf'),                           # PDF
    (b'GIF87a', 'gif'),                          # GIF87a
    (b'GIF89a', 'gif'),                          # GIF89a
    (b'RIFF', 'webp'),                           # WebP (RIFF + 'WEBP' at offset 8)
]


def _validate_magic_bytes(data: bytes, expected_ext: str) -> bool:
    """파일 첫 몇 바이트로 실제 종류 확인. 확장자 위장 방어.

    - HEIC 는 offset 4~11 에 'ftyp' 시그니처 있어 별도 처리
    - WebP 는 offset 8~11 에 'WEBP' 있어야 함
    - 그 외는 prefix 매칭
    """
    if not data or len(data) < 8:
        return False
    ext = (expected_ext or '').lower()
    # HEIC
    if ext == 'heic':
        return len(data) >= 12 and data[4:8] == b'ftyp' and b'heic' in data[8:16]
    # WebP
    if ext == 'webp':
        return data[:4] == b'RIFF' and len(data) >= 12 and data[8:12] == b'WEBP'
    # PDF/JPEG/PNG/GIF prefix 매칭
    for sig, sig_ext in _MAGIC_SIGNATURES:
        if data.startswith(sig) and sig_ext == ext:
            return True
    # 특수: JPG variants (all start with FFD8FF)
    if ext == 'jpg' and data[:3] == b'\xff\xd8\xff':
        return True
    return False


def _sanitize_filename_for_log(name: str) -> str:
    """로그·에러 메시지용 파일명 sanitize.

    Google Drive API 는 자체 필터링하지만 로그에 raw filename 노출 시
    특수문자·경로 traversal 조작 가능. 로그 표시용만 정리.
    """
    if not name:
        return '(no-name)'
    # 제어 문자 + 경로 구분자 제거, 길이 100 제한
    import re as _re
    cleaned = _re.sub(r'[\x00-\x1f\x7f\\/]', '_', str(name))[:100]
    return cleaned or '(empty)'


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

    ext = _guess_ext(filename, mimetype)

    # 매직 넘버 검증 — 확장자 위장 방어 (2026-07-10)
    # 사업자등록증은 이미지·PDF 만 허용. 실제 파일 종류가 확장자와 일치하지 않으면 거부.
    if not _validate_magic_bytes(file_bytes, ext):
        logger.warning(
            f'[LICENSE] 매직 넘버 불일치 거부: filename={_sanitize_filename_for_log(filename)} '
            f'ext={ext} first_bytes={file_bytes[:8].hex() if file_bytes else "empty"}'
        )
        return {'ok': False, 'reason': 'invalid_file_signature'}

    drive = _get_drive()
    license_folder = _get_or_create_license_subfolder(drive, parent)

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


_MIMETYPE_BY_EXT = {
    'pdf': 'application/pdf',
    'png': 'image/png',
    'jpg': 'image/jpeg',
    'jpeg': 'image/jpeg',
    'heic': 'image/heic',
    'gif': 'image/gif',
    'webp': 'image/webp',
}


def fetch_license_canonical(code: str) -> Optional[dict]:
    """canonical 사업자등록증 파일 다운로드. 세금계산서 카드 스레드 첨부용.

    Returns: {'file_name': str, 'content': bytes, 'mimetype': str} or None.
    canonical 여러 확장자면 pdf > png > jpg 우선.
    """
    parent = _project_folder_id_for(code)
    if not parent:
        return None
    drive = _get_drive()
    fid = _find_license_subfolder(drive, parent)
    if not fid:
        return None
    expected_base = f'{code} {LICENSE_BASENAME}'
    ext_priority = {'pdf': 0, 'png': 1, 'jpg': 2, 'jpeg': 2, 'webp': 3, 'gif': 4, 'heic': 5}
    candidates = []
    for f in _list_folder_files(drive, fid):
        name = f['name']
        if '.' not in name:
            continue
        base, ext = name.rsplit('.', 1)
        if base != expected_base:
            continue
        candidates.append((ext_priority.get(ext.lower(), 99), name, f['id'], ext.lower()))
    if not candidates:
        return None
    candidates.sort()
    _, name, file_id, ext = candidates[0]
    try:
        request = drive.files().get_media(fileId=file_id, supportsAllDrives=True)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        content = buf.getvalue()
    except Exception as exc:
        logger.warning(f'[LICENSE/FETCH] Drive 다운로드 실패 ({code}, {name}): {exc}')
        return None
    return {
        'file_name': name,
        'content': content,
        'mimetype': _MIMETYPE_BY_EXT.get(ext, 'application/octet-stream'),
    }


def resolve_project_from_thread(channel: str, thread_ts: str, slack_bot_token: str = '') -> Optional[str]:
    """channel|thread_ts → 프로젝트 코드 (없으면 None).

    1차: Redis 매핑 (카드 발송 시점 저장). 최근 카드만 있음.
    2차 (fallback, 2026-07-13): Slack API 로 스레드 root 메시지 조회 후
    텍스트에서 프로젝트 코드 정규식 추출. 오래된 카드도 매칭 가능.
    """
    if not channel or not thread_ts:
        return None
    # 1차: Redis
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
        v = rc.get(f'project_thread:{channel}|{thread_ts}')
        if isinstance(v, bytes):
            v = v.decode('utf-8')
        if v:
            return v
    except Exception as exc:
        logger.warning(f'[LICENSE] 스레드→프로젝트 Redis 조회 실패: {exc}')

    # 2차 fallback: Slack API 로 root 메시지 텍스트에서 프로젝트 코드 추출
    if not slack_bot_token:
        return None
    try:
        import re
        from slack_sdk import WebClient
        client = WebClient(token=slack_bot_token)
        resp = client.conversations_replies(channel=channel, ts=thread_ts, limit=1)
        msgs = resp.get('messages') or []
        if not msgs:
            return None
        root_text = msgs[0].get('text', '') or ''
        # 2026-07-16 사고: 세금계산서 발행 완료 카드도 헤더에 프로젝트 코드를 포함
        # (`[세금계산서 발행 완료] G3560-YG`) → 정규식 매칭만 하면 오라우팅.
        # 반드시 `[공사 확정]` 헤더가 있는 스레드만 사업자등록증 저장 대상.
        if '[공사 확정]' not in root_text:
            logger.info(
                f'[LICENSE] 스레드 root 에 [공사 확정] 헤더 없음 → skip '
                f'(channel={channel} ts={thread_ts})'
            )
            return None
        # 공사확정 카드 헤더 패턴: G1234-SH / R1234-MW / P1234-YG 등
        # Slack fallback text 는 백틱이 벗겨진 상태로 옴 → 백틱 옵셔널.
        m = re.search(r'`?([GRNP]\d{3,5}-[A-Z]{1,3})`?', root_text)
        if m:
            code = m.group(1)
            logger.info(f'[LICENSE] 스레드→프로젝트 slack fallback 매칭: {code}')
            # Redis 에 다시 저장 (다음 첨부에서 slack API 재조회 방지, 30일 TTL)
            try:
                from dashboard.utils.redis_client import get_redis_client
                rc = get_redis_client().redis
                rc.setex(f'project_thread:{channel}|{thread_ts}', 86400 * 30, code)
            except Exception:
                pass
            return code
    except Exception as exc:
        logger.warning(f'[LICENSE] 스레드→프로젝트 slack fallback 실패: {exc}')
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

    code = resolve_project_from_thread(channel, thread_ts, slack_bot_token)
    if not code:
        return None  # 프로젝트 카드 스레드 아님

    files = event.get('files') or []
    if not files:
        return None

    saved = []
    skipped = []
    ocr_content = None  # OCR 대상 (첫 번째로 성공 저장된 파일 content)
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
                # OCR 은 첫 파일만 (여러 개 첨부해도 하나만 처리 — API cost 절감)
                if ocr_content is None:
                    ocr_content = content
            else:
                skipped.append(f'{name}: 저장 실패({res.get("reason")})')
        except Exception as exc:
            logger.error(f'[LICENSE] 처리 예외 ({name}): {exc}', exc_info=True)
            skipped.append(f'{name}: 예외 {type(exc).__name__}')

    # 하나라도 저장 성공했으면 원본 공사 확정 카드의 사업자등록증 배지를 ✅로 갱신.
    # 오래된 카드는 Redis 매핑이 없어 skip 되던 문제 → thread_ts (= 카드 ts) 를
    # fallback 인자로 전달 (2026-07-13).
    if saved:
        try:
            from dashboard.services.project_slack_notifier import refresh_project_card_license
            refresh_project_card_license(
                code,
                fallback_channel=channel,
                fallback_message_ts=thread_ts,
            )
        except Exception as exc:
            logger.warning(f'[LICENSE] 원본 카드 갱신 실패 ({code}): {exc}')

    # OCR — 첫 저장 파일에서 법인명·상호 추출 (Google Vision API, 2026-07-13).
    # 실패 시 조용히 skip. 성공 시 시트 사업자명 자동 반영 시도 (안전장치 있음).
    business_name = ''
    biz_update_status = ''  # '' | 'saved' | 'match' | 'mismatch' | 'error'
    biz_update_existing = ''
    if ocr_content:
        try:
            from dashboard.services.business_license_ocr import ocr_business_license
            business_name = ocr_business_license(ocr_content) or ''
        except Exception as exc:
            logger.warning(f'[LICENSE/OCR] 예외: {exc}')
        if business_name:
            try:
                biz_update_status, biz_update_existing = _maybe_update_business_name(code, business_name)
            except Exception as exc:
                logger.warning(f'[LICENSE/OCR] 사업자명 자동 반영 실패: {exc}')
                biz_update_status = 'error'

    return {
        'code': code, 'saved': saved, 'skipped': skipped,
        'thread_ts': thread_ts, 'channel': channel,
        'business_name': business_name,
        'biz_update_status': biz_update_status,
        'biz_update_existing': biz_update_existing,
    }


def _maybe_update_business_name(code: str, ocr_name: str) -> tuple:
    """OCR 로 추출한 법인명·상호를 시트 사업자명 필드에 자동 반영.

    정책 (오탐 리스크 최소화):
      - 비어있으면 자동 저장 → 'saved'
      - 기존값 있고 OCR 결과와 일치 → 'match' (안내 불필요)
      - 기존값 있고 다름 → 'mismatch' (덮어쓰지 않음, 매니저 확인 유도)

    반환: (status, existing_value)
    """
    if not code or not ocr_name:
        return '', ''
    # 캐시된 프로젝트 데이터에서 현재 사업자명 조회
    from dashboard.services.project_service import get_project_records, get_sheets_manager
    records = get_project_records() or []
    target_row = None
    existing = ''
    for i, r in enumerate(records, start=2):
        if (r.get('프로젝트 코드') or '').strip() == code:
            existing = (r.get('사업자명') or '').strip()
            target_row = i  # 시트 row (header 제외 + 1-indexed)
            break
    if target_row is None:
        logger.warning(f'[LICENSE/OCR] 프로젝트 못 찾음: {code}')
        return '', ''

    # 정규화 비교 (공백 무시)
    def _norm(s: str) -> str:
        import re
        return re.sub(r'\s+', '', s).strip()

    # '-' 도 미기입으로 간주 (2026-07-13 사용자 정책)
    if existing and existing != '-':
        if _norm(existing) == _norm(ocr_name):
            return 'match', existing
        return 'mismatch', existing

    # 비어있거나 '-' 이면 시트 update
    import os
    sheet_id = os.getenv('GOOGLE_SHEET_ID', '').strip()
    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '').strip()
    if not (sheet_id and sheet_name):
        logger.warning('[LICENSE/OCR] GOOGLE_SHEET_ID/NAME 미설정')
        return 'error', ''
    manager = get_sheets_manager()
    result = manager.batch_update_fields(
        sheet_id, sheet_name, target_row,
        {'사업자명': ocr_name},
        project_code=code,
    )
    if not result.get('success'):
        logger.warning(f'[LICENSE/OCR] 시트 업데이트 실패: {result}')
        return 'error', ''

    # 캐시 무효화 → 다음 조회부터 최신값 반영
    try:
        from dashboard.services.project_service import invalidate_project_cache
        invalidate_project_cache(code, trigger_refresh=False)
    except Exception:
        try:
            from dashboard.utils.smart_cache_manager import get_cache_manager
            get_cache_manager().invalidate_pattern('projects_list')
        except Exception as exc:
            logger.debug(f'[LICENSE/OCR] 캐시 무효화 실패 (무시): {exc}')

    logger.info(f'[LICENSE/OCR] 사업자명 자동 저장: {code} → {ocr_name!r}')
    return 'saved', ''
