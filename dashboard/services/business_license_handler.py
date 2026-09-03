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
import json
import logging
import os
import re
import threading
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

# 스레드별 Drive 클라이언트 (프로세스 싱글톤 금지).
# googleapiclient(httplib2)는 thread-safe 하지 않아 여러 스레드가 한 인스턴스를 공유하면
# heap corruption 위험 (2026-07-09 GoogleSheetsManager 동일 사고). Waitress 웹 요청 스레드와
# 슬랙 Bolt 핸들러 스레드가 각자 인스턴스를 갖도록 threading.local 사용.
_drive_local = threading.local()


def _get_drive():
    svc = getattr(_drive_local, 'service', None)
    if svc is None:
        creds_path = os.getenv('GOOGLE_APPLICATION_CREDENTIALS', 'credentials.json')
        credentials = service_account.Credentials.from_service_account_file(
            creds_path, scopes=DRIVE_SCOPES,
        )
        svc = build('drive', 'v3', credentials=credentials, cache_discovery=False)
        _drive_local.service = svc
    return svc


def _project_folder_id_for(code: str, fresh: bool = False) -> Optional[str]:
    """프로젝트 코드 → 시트의 '견적서 및 계약서 폴더 경로' 값을 진짜 폴더 ID로 정규화.

    2026-08-18: 폴더칸에 폴더가 아니라 그 안의 '파일' 링크를 붙여넣은 오입력(파일→상위폴더)과
    URL 형태 붙여넣기를 resolve_folder_id 로 저장 시점에 자동 교정. 등록 검증(1c37308)을
    우회한 옛 데이터까지 여기서 교정된다. 유효한 폴더 ID 를 못 얻으면 None
    (→ 호출부에서 'no_project_folder' 명확 안내).

    fresh=True: 캐시 대신 시트 강제 재조회. PM에서 폴더 경로를 방금 넣고 바로 재첨부하면
    캐시 지연으로 '폴더 없음' 오판정되던 문제 방지 (2026-08-26 R4035-JW 계기, 저장 실패 시 재시도용).
    """
    from dashboard.services.project_service import get_project_records
    records = get_project_records(force_refresh=fresh) or []
    raw = None
    for r in records:
        if (r.get('프로젝트 코드') or '').strip() == code:
            raw = str(r.get('견적서 및 계약서 폴더 경로') or '').strip()
            break
    if not raw:
        return None
    try:
        from dashboard.utils.google_drive import resolve_folder_id
        res = resolve_folder_id(raw)
        reason = res.get('reason')
        # 로컬경로 등 비-ID, 상위폴더조차 없는 파일 → 사용 불가 (명확 안내 유도)
        if reason in ('not_id_format', 'file_no_parent', 'empty'):
            return None
        # ok_folder·file_to_parent(교정됨)·no_service·lookup_failed → 최선값 사용
        val = str(res.get('value') or '').strip()
        if re.fullmatch(r'[a-zA-Z0-9_-]{20,}', val):
            return val
        return None
    except Exception as exc:
        logger.warning(f'[LICENSE] 폴더 경로 정규화 실패 ({code}): {exc}')
        # 폴백: 기존 규칙 (bare 폴더 ID 만 인정)
        if re.match(r'^[a-zA-Z0-9_-]{20,}$', raw):
            return raw
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
    if ext == 'jpeg':
        ext = 'jpg'  # .jpeg 확장자도 JPEG 로 인정 (리사이즈 파일 등)
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
            return 'jpg' if ext == 'jpeg' else ext  # .jpeg → jpg 정규화
    # mimetype fallback
    mime_ext = {
        'image/jpeg': 'jpg',
        'image/png': 'png',
        'image/heic': 'heic',
        'application/pdf': 'pdf',
    }
    return mime_ext.get((mimetype or '').lower(), 'bin')


def save_business_license(code: str, file_bytes: bytes, filename: str, mimetype: str,
                          parent_override: Optional[str] = None) -> dict:
    """사업자등록증 파일을 프로젝트 폴더에 저장.

    parent_override: 폴더 ID를 직접 지정(자동 재시도용). write-behind 시트 반영 지연을
        우회해 방금 등록된 폴더에 바로 저장하기 위함.

    Returns:
        {'ok': bool, 'reason': str, 'file_name': str, 'file_id': str}
    """
    parent = str(parent_override or '').strip() or _project_folder_id_for(code)
    if not parent:
        # 캐시 지연 대비: 폴더 경로를 방금 PM에 넣었을 수 있으니 시트 fresh 재조회 후 재판정
        parent = _project_folder_id_for(code, fresh=True)
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

    # 저장 파일명 규칙 (2026-08-18 개정 — '최신이 canonical'):
    #   새로 올린 파일이 항상 '{code} 사업자등록증.{ext}' (canonical) 가 된다.
    #   기존 canonical(확장자 무관)이 있으면 '{code} 사업자등록증_{N}.{ext}' 백업으로 밀어낸다.
    #   → 잘못 첨부 후 정정(삭제→새 첨부) 시 '최신=올바른' 파일이 canonical =
    #     계산서 요청 때 fetch_license_canonical 이 최신본을 첨부. (R3916-TH 계기)
    #   구 규칙('first-wins', 2026-07-09)은 정정해도 잘못된 첫 파일이 canonical 로 남아
    #   계산서에 계속 첨부되던 문제 → 반전.
    existing = _list_folder_files(drive, license_folder)
    expected_base = f'{code} {LICENSE_BASENAME}'
    canonical_cleared = True  # 새 파일을 canonical 로 저장 가능한가 (같은 확장자 canonical 이 안 남았는가)
    for f in list(existing):
        nm = f.get('name', '')
        if '.' not in nm:
            continue
        b, e = nm.rsplit('.', 1)
        if b != expected_base:
            continue  # 이미 백업(_N)이거나 무관 파일 — 유지
        e_low = e.lower()
        n = _next_copy_index(existing, code, e_low)
        backup_name = f'{code} {LICENSE_BASENAME}_{n}.{e}'
        try:
            drive.files().update(
                fileId=f['id'], body={'name': backup_name},
                fields='id', supportsAllDrives=True,
            ).execute()
            logger.info(f'[LICENSE] 기존본 백업 전환: {nm} → {backup_name} (project={code})')
        except Exception as exc:
            logger.warning(f'[LICENSE] 기존본 백업 rename 실패 ({code}, {nm}): {exc}')
            if e_low == ext:
                canonical_cleared = False  # 같은 확장자 canonical 이 남음 → 이름충돌·dup 방지

    if canonical_cleared:
        save_name = canonical
    else:
        # 폴백: 같은 확장자 canonical 을 못 밀어냄(Drive 오류) → dup 방지 위해 백업명 저장.
        # 파일 유실 방지 우선(최신이 canonical 은 못 되지만 새 파일은 보존). rename 실패는 드묾.
        n = _next_copy_index(_list_folder_files(drive, license_folder), code, ext)
        save_name = f'{code} {LICENSE_BASENAME}_{n}.{ext}'
        logger.warning(f'[LICENSE] canonical 확보 실패 → 새 파일 백업 저장 ({code}): {save_name}')

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
    invalidate_license_state(code)  # 상태 캐시 무효화 → PM 뱃지 즉시 최신 반영 (슬랙·PM 공통)
    # 사업자명 재사용 인덱스 갱신 — 이후 같은 사업자명의 다른 공사가 이 등록증을 재사용(복사).
    try:
        _index_license_source(_project_norm_biz(code), code, up['id'], ext)
    except Exception as exc:
        logger.debug(f'[LICENSE] 재사용 인덱스 갱신 실패 ({code}): {exc}')
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


def get_license_view_url(code: str) -> Optional[str]:
    """canonical 사업자등록증 파일의 Drive 열람 URL (없으면 None).

    PM '열람' 버튼용 — 새 탭에서 사용자의 구글 세션으로 연다(회사 공유드라이브 접근 권한 전제).
    canonical 여러 확장자면 fetch_license_canonical 과 동일 우선순위(pdf>png>jpg).
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
        candidates.append((ext_priority.get(ext.lower(), 99), f['id']))
    if not candidates:
        return None
    candidates.sort()
    return f'https://drive.google.com/file/d/{candidates[0][1]}/view'


_LICENSE_STATE_TTL = 30 * 24 * 3600  # 등록증 상태 캐시 (30일=사실상 영구).
# 모든 쓰기 경로(save_business_license·trash_license_canonical)가 invalidate 하므로 길게 잡아도 정확.
# 백그라운드 워머(warm_license_states)가 시작·주기적으로 미리 채워 '첫 조회'도 즉시.
# (30일 TTL 은 앱 밖 수동 Drive 조작 같은 예외 상황의 자동 자가치유 안전망.)


def _license_state_key(code: str) -> str:
    return f'license_state:{code}'


def invalidate_license_state(code: str) -> None:
    """등록증 상태 캐시 무효화 (업로드/변경 시). save_business_license 성공 시 자동 호출."""
    try:
        from dashboard.utils.redis_client import get_redis_client
        get_redis_client().redis.delete(_license_state_key((code or '').strip()))
    except Exception:
        pass


# ── 사업자명(상호) → 등록증 재사용 인덱스 (2026-09-02) ──
# 반복 거래처: 같은 사업자명의 기존 프로젝트 등록증을 새 프로젝트로 '복사'해 재사용한다.
# (거래처 blanket 예외 폐지 → 모든 공사에 등록증 필요. 사업자명 매칭되면 재사용, 아니면 차단.)
# 이유: 인테리어 업체 경유라도 고객과 직접 금전 거래가 많아 프로젝트마다 각자 등록증이 필요.
_BIZ_INDEX_KEY = 'license_biz_index'  # Redis hash: norm_biz -> {"code","file_id","ext"}
_EXT_PRIORITY = {'pdf': 0, 'png': 1, 'jpg': 2, 'jpeg': 2, 'webp': 3, 'gif': 4, 'heic': 5}


# 유입 플레이스홀더 — '사업자명' 필드에 상호 대신 유입구분/기본값이 들어간 경우.
# 이 값들은 회사 식별자가 아니므로 등록증 재사용/인덱싱/거래처 매칭의 키가 되면 안 된다.
# (예: '온라인' 274건이 한 키로 묶여 무관한 고객의 등록증이 계산서에 오첨부되는 사고 방지. 2026-09-02)
_PLACEHOLDER_BIZ = {
    '온라인', '거래처', '전화', '소개', '기타', '카톡', '카카오', '방문',
    '미정', '개인', '없음', 'none', 'null', 'na', 'n/a',
}


def _norm_biz(name: str) -> str:
    """사업자명(상호) 정규화 — 상호→이메일 캐시와 동일 규칙(법인표기 유지, 정밀 매칭).

    유입 플레이스홀더(온라인/거래처/- 등)와 1자 이하는 '' 반환 → 등록증 재사용·인덱싱
    대상에서 제외(무관한 고객 간 등록증 오첨부 방지). own(자기 폴더 파일)은 폴더 기반이라 무영향.
    """
    try:
        from dashboard.services.partner_status_sync import _norm_name
        nb = _norm_name(name)
    except Exception:
        nb = re.sub(r'\s+', '', str(name or '')).strip().lower()
    if nb in _PLACEHOLDER_BIZ or len(nb) <= 1:
        return ''
    return nb


def _project_norm_biz(code: str) -> str:
    try:
        from dashboard.services.project_service import get_project_records
        for r in (get_project_records() or []):
            if (r.get('프로젝트 코드') or '').strip() == code:
                return _norm_biz(r.get('사업자명') or '')
    except Exception:
        pass
    return ''


def _index_license_source(norm_biz: str, code: str, file_id: str, ext: str) -> None:
    """이 프로젝트를 해당 사업자명의 등록증 '원본 소스'로 인덱싱 (재사용 복사 출처)."""
    if not norm_biz or not file_id:
        return
    try:
        from dashboard.utils.redis_client import get_redis_client
        get_redis_client().redis.hset(
            _BIZ_INDEX_KEY, norm_biz,
            json.dumps({'code': code, 'file_id': file_id, 'ext': ext}, ensure_ascii=False))
    except Exception:
        pass


def _get_license_source(norm_biz: str) -> Optional[dict]:
    if not norm_biz:
        return None
    try:
        from dashboard.utils.redis_client import get_redis_client
        v = get_redis_client().redis.hget(_BIZ_INDEX_KEY, norm_biz)
        if v:
            return json.loads(v.decode() if isinstance(v, bytes) else v)
    except Exception:
        pass
    return None


def _own_canonical(code: str, drive=None):
    """프로젝트 자기 폴더의 canonical 등록증 → (file_id, ext) or None."""
    parent = _project_folder_id_for(code)
    if not parent:
        return None
    drive = drive or _get_drive()
    fid = _find_license_subfolder(drive, parent)
    if not fid:
        return None
    expected_base = f'{code} {LICENSE_BASENAME}'
    cands = []
    for f in _list_folder_files(drive, fid):
        nm = f['name']
        if '.' not in nm:
            continue
        base, ext = nm.rsplit('.', 1)
        if base != expected_base:
            continue
        cands.append((_EXT_PRIORITY.get(ext.lower(), 99), f['id'], ext.lower()))
    if not cands:
        return None
    cands.sort()
    return (cands[0][1], cands[0][2])


def _copy_license_to_project(code: str, source: dict):
    """source(다른 프로젝트의 등록증)를 code 프로젝트 폴더로 **복사** → (file_id, ext) or None.
    반복 거래처 재사용: 각 공사가 물리적으로 등록증을 보유하게 한다 (원본은 유지)."""
    src_id = (source or {}).get('file_id')
    ext = ((source or {}).get('ext') or 'pdf').lower()
    if not src_id:
        return None
    parent = _project_folder_id_for(code)
    if not parent:
        return None
    drive = _get_drive()
    sub = _get_or_create_license_subfolder(drive, parent)
    new_name = _canonical_name(code, ext)
    try:
        r = drive.files().copy(
            fileId=src_id, body={'name': new_name, 'parents': [sub]},
            fields='id,name', supportsAllDrives=True,
        ).execute()
    except Exception as exc:
        logger.warning(f'[LICENSE/REUSE] 복사 실패 ({code} ← {source.get("code")}): {exc}')
        return None
    logger.info(f'[LICENSE/REUSE] 등록증 재사용 복사: {source.get("code")} → {code} ({new_name})')
    return (r['id'], ext)


def get_license_state(code: str, use_cache: bool = True) -> dict:
    """등록증 상태 → {'exists','view_url','source'}. source: 'own'|'reuse'|None (표시/게이트용, **복사 안 함**).

    - own: 자기 폴더에 등록증 있음.
    - reuse: 자기 폴더엔 없지만 **같은 사업자명**의 기존 등록증이 있어 재사용 가능(열람은 원본 링크).
      실제 물리 복사는 계산서 요청 시 ensure_license() 가 수행.
    - None: 없음 → 계산서 요청 차단(업로드 필요).
    Redis 캐시(30일). 업로드·삭제·복사 시 invalidate.
    """
    code = (code or '').strip()
    if not code:
        return {'exists': False, 'view_url': None, 'source': None}
    if use_cache:
        try:
            from dashboard.utils.redis_client import get_redis_client
            raw = get_redis_client().redis.get(_license_state_key(code))
            if raw:
                return json.loads(raw.decode() if isinstance(raw, bytes) else raw)
        except Exception:
            pass

    state = {'exists': False, 'view_url': None, 'source': None}
    nb = _project_norm_biz(code)
    oc = _own_canonical(code)
    if oc:
        fid, ext = oc
        state = {'exists': True, 'view_url': f'https://drive.google.com/file/d/{fid}/view',
                 'source': 'own', 'file_id': fid, 'ext': ext}
        _index_license_source(nb, code, fid, ext)  # 인덱스 최신화
    else:
        src = _get_license_source(nb)
        if src and src.get('code') != code and src.get('file_id'):
            state = {
                'exists': True,
                'view_url': f'https://drive.google.com/file/d/{src.get("file_id")}/view',
                'source': 'reuse',
            }
        else:
            # 등록증 파일이 없어도 **거래처 탭에 상호가 있으면 계산서 발행 가능**
            # (홈택스 발행 이력 있는 거래처. 사용자 통찰 2026-09-02). 파일 없음 → view_url 없음.
            try:
                from dashboard.services.partner_status_sync import is_partner_known
                if is_partner_known(nb):
                    state = {'exists': True, 'view_url': None, 'source': 'partner'}
            except Exception:
                pass
    try:
        from dashboard.utils.redis_client import get_redis_client
        get_redis_client().redis.setex(_license_state_key(code), _LICENSE_STATE_TTL, json.dumps(state))
    except Exception:
        pass
    return state


def ensure_license(code: str):
    """계산서 첨부 직전 호출 — 자기 폴더에 등록증을 **보장**(없으면 사업자명 매칭 소스에서 복사).

    Returns (file_id, ext) or None. 복사 시 상태 캐시 무효화(다음 조회 own 반영).
    """
    code = (code or '').strip()
    if not code:
        return None
    oc = _own_canonical(code)
    if oc:
        return oc
    src = _get_license_source(_project_norm_biz(code))
    if src and src.get('code') != code:
        copied = _copy_license_to_project(code, src)
        if copied:
            invalidate_license_state(code)
            return copied
    return None


def trash_license_canonical(code: str) -> dict:
    """canonical 사업자등록증을 Drive **휴지통으로 이동**(복구 가능, 영구삭제 아님).

    PM 등록증 [삭제] 버튼용. 오분류·복구 이력이 있어 영구삭제 대신 trashed=True.
    백업본(_N)은 건드리지 않고 현재 canonical 만 이동 → 상태는 '없음'이 되어 재업로드 유도.
    Returns {'ok': bool, 'reason': str, 'file_name'?}.
    """
    code = (code or '').strip()
    if not code:
        return {'ok': False, 'reason': 'no_code'}
    parent = _project_folder_id_for(code)
    if not parent:
        return {'ok': False, 'reason': 'no_project_folder'}
    drive = _get_drive()
    fid = _find_license_subfolder(drive, parent)
    if not fid:
        return {'ok': False, 'reason': 'no_license'}
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
        candidates.append((ext_priority.get(ext.lower(), 99), name, f['id']))
    if not candidates:
        return {'ok': False, 'reason': 'no_license'}
    candidates.sort()
    _, name, file_id = candidates[0]
    try:
        drive.files().update(
            fileId=file_id, body={'trashed': True}, supportsAllDrives=True,
        ).execute()
    except Exception as exc:
        logger.warning(f'[LICENSE] 휴지통 이동 실패 ({code}, {name}): {exc}')
        return {'ok': False, 'reason': 'drive_error'}
    logger.info(f'[LICENSE] 휴지통 이동(삭제): {name} (project={code}, id={file_id})')
    invalidate_license_state(code)
    return {'ok': True, 'reason': 'trashed', 'file_name': name}


def warm_license_states(throttle: float = 0.15, limit: Optional[int] = None) -> int:
    """등록증 상태 캐시 + 사업자명 재사용 인덱스 백그라운드 구축 (2-pass, 2026-09-02).

    Pass A: 비취소 프로젝트 own-check → 사업자명→등록증 인덱스 구축 + own 상태 캐시.
      (이미 'own' 캐시된 건은 캐시의 file_id/ext 로 인덱싱 → Drive 재호출 없이 재실행 저렴.)
    Pass B: own 없는 프로젝트 → **인덱스 완성 후** 사업자명 매칭되면 '재사용', 아니면 '없음' 캐시.
      2-pass 라 '매칭 소스보다 먼저 처리돼 없음으로 굳는' 순서 문제 없음.
      (물리 복사는 계산서 요청 시 ensure_license 가 수행 — 워머는 상태/인덱스만.)
    반환: own-check 처리 건수.
    """
    try:
        from dashboard.services.project_service import get_project_records
        records = get_project_records() or []
    except Exception as exc:
        logger.warning(f'[LICENSE/WARM] 프로젝트 목록 로드 실패: {exc}')
        return 0

    def _code_num(r):
        m = re.search(r'[GPR](\d+)', str(r.get('프로젝트 코드') or ''))
        return int(m.group(1)) if m else 0
    records = sorted(records, key=_code_num, reverse=True)
    active = [
        r for r in records
        if (r.get('프로젝트 코드') or '').strip()
        and '공사취소' not in str(r.get('수금 관련 특이사항') or '').replace(' ', '')
    ]

    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    except Exception:
        rc = None

    def _cache(code, st):
        if rc is not None:
            try:
                rc.setex(_license_state_key(code), _LICENSE_STATE_TTL, json.dumps(st))
            except Exception:
                pass

    drive = _get_drive()
    needing = []
    checked = 0
    # ── Pass A: 인덱스 구축 + own 캐시 ──
    for r in active:
        code = (r.get('프로젝트 코드') or '').strip()
        nb = _norm_biz(r.get('사업자명') or '')
        # 이미 own 캐시면 캐시의 file_id/ext 로 인덱싱 (Drive 재호출 회피)
        cached = None
        if rc is not None:
            try:
                raw = rc.get(_license_state_key(code))
                cached = json.loads(raw.decode() if isinstance(raw, bytes) else raw) if raw else None
            except Exception:
                cached = None
        # 신포맷 own 캐시(file_id 보유)면 캐시로 인덱싱 → Drive 재호출 회피.
        # 그 외(구포맷·reuse·없음·미캐시)는 반드시 own-check (구포맷 own 을 needing 오분류 방지).
        if cached and cached.get('source') == 'own' and cached.get('file_id'):
            _index_license_source(nb, code, cached.get('file_id'), cached.get('ext') or 'pdf')
            continue
        try:
            oc = _own_canonical(code, drive)
        except Exception as exc:
            logger.debug(f'[LICENSE/WARM] {code} own-check 실패: {exc}')
            continue
        if oc:
            fid, ext = oc
            _index_license_source(nb, code, fid, ext)
            _cache(code, {'exists': True, 'view_url': f'https://drive.google.com/file/d/{fid}/view',
                          'source': 'own', 'file_id': fid, 'ext': ext})
        else:
            needing.append((code, nb))
        checked += 1
        if limit and checked >= limit:
            break
        if throttle:
            time.sleep(throttle)

    # ── Pass B: 재사용 참조 / 거래처 탭 발행가능 / 없음 캐시 (인덱스 완성 후) ──
    try:
        from dashboard.services.partner_status_sync import is_partner_known
    except Exception:
        is_partner_known = None
    reused = 0
    partner = 0
    for code, nb in needing:
        src = _get_license_source(nb)
        if src and src.get('code') != code and src.get('file_id'):
            _cache(code, {'exists': True,
                          'view_url': f'https://drive.google.com/file/d/{src.get("file_id")}/view',
                          'source': 'reuse'})
            reused += 1
        elif is_partner_known and is_partner_known(nb):
            # 등록증 파일 없어도 거래처 탭에 있으면 발행 가능
            _cache(code, {'exists': True, 'view_url': None, 'source': 'partner'})
            partner += 1
        else:
            _cache(code, {'exists': False, 'view_url': None, 'source': None})
    logger.info(f'[LICENSE/WARM] 상태·인덱스 워밍: own-check {checked}건 / 재사용 {reused} / 거래처발행가능 {partner} / 대상 {len(active)}')
    return checked


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


def _reason_message(reason: str) -> str:
    """save_business_license 실패 reason → 매니저용 명확·조치가능 안내 (2026-08-18)."""
    r = str(reason or '')
    if r == 'no_project_folder':
        return ('⚠️ 견적서·계약서 폴더가 없거나 잘못 입력돼 저장하지 못했습니다. '
                'PM 사이트에서 이 프로젝트의 "견적서 및 계약서 폴더 경로"에 올바른 *폴더* 링크를 '
                '입력한 뒤 파일을 다시 첨부해 주세요.')
    if r == 'invalid_file_signature':
        return ('이미지·PDF 형식이 아니거나 파일이 손상돼 저장하지 못했습니다. '
                '사업자등록증 원본(사진·PDF)이 맞는지 확인해 주세요.')
    return f'저장 실패({r})'


def _exception_message(exc: Exception) -> str:
    """저장 중 예외 → 매니저용 명확 안내. parentNotAFolder(폴더칸에 파일 링크) 특별 처리."""
    detail = str(exc)
    if 'parentNotAFolder' in detail or 'not a folder' in detail.lower():
        return ('⚠️ "견적서 및 계약서 폴더 경로"에 폴더가 아니라 파일 링크가 입력돼 있어 '
                '저장하지 못했습니다. PM 사이트에서 해당 칸을 *폴더* 링크로 고친 뒤 다시 첨부해 주세요.')
    return (f'저장 중 오류가 발생했습니다 ({type(exc).__name__}). '
            '잠시 후 다시 시도하거나, 계속 실패하면 폴더 경로를 확인해 주세요.')


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
    pending_files = []  # 폴더 미등록으로 실패 → 폴더 경로 등록 시 자동 재시도용 (2026-08-28)
    ocr_result = None       # 저장된 사업자등록증의 OCR 분석 결과 (사업자명 반영·경고 재사용)
    not_license_any = False  # 사업자등록증 아닌 파일이 하나라도 있었나 (세금계산서·정산문서 등)
    is_card_any = False
    ambiguous_any = False    # 정산문서/카드는 아닌데 등록증 제목도 못 잡음 = 흐릿한 등록증 의심(안내)
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
            # 문서 종류 판별 — 사업자등록증/고유번호증만 저장. 세금계산서·지출품의서·정산 PDF 등
            #   (같은 스레드에 자주 올라옴)이 canonical 을 덮어쓰던 사고 방지 (2026-09-01 R3883-SJ).
            try:
                from dashboard.services.business_license_ocr import analyze_business_license
                _an = analyze_business_license(content)
            except Exception as exc:
                logger.warning(f'[LICENSE] OCR 판별 예외 ({name}): {exc}')
                _an = {}
            if _an.get('is_card'):
                is_card_any = True
            if not _an.get('is_license'):
                not_license_any = True
                # 명백한 정산문서(세금계산서 등)·카드가 아니고 텍스트는 있는데 제목만 못 잡았으면
                #   흐릿한 등록증일 수 있음 → ambiguous(가벼운 안내). 명백한 문서/카드는 조용히 skip
                #   (공사확정 스레드엔 정산문서가 자주 올라오므로 소음 방지). skipped 에 넣지 않음.
                if _an.get('has_text') and not _an.get('doc_negative') and not _an.get('is_card'):
                    ambiguous_any = True
                logger.info(
                    f'[LICENSE] 비-등록증 저장 skip ({code}): {_sanitize_filename_for_log(name)} '
                    f'(neg={_an.get("doc_negative")} card={_an.get("is_card")} text={_an.get("has_text")})')
                continue

            res = save_business_license(code, content, name, mimetype)
            if res.get('ok'):
                saved.append(res['file_name'])
                if ocr_result is None:
                    ocr_result = _an  # 첫 저장 등록증의 분석 결과 재사용 (재 OCR 안 함)
            else:
                skipped.append(f'{name}: {_reason_message(res.get("reason"))}')
                if res.get('reason') == 'no_project_folder':
                    pending_files.append({'name': name, 'url': url, 'mime': mimetype})
        except Exception as exc:
            logger.error(f'[LICENSE] 처리 예외 ({name}): {exc}', exc_info=True)
            skipped.append(f'{name}: {_exception_message(exc)}')

    # 폴더 미등록으로 실패한 첨부는 기록 → PM에서 폴더 경로 등록되면 자동 재저장
    # (retry_pending_license, notify_project_field_changes 에서 호출). 성공분 있으면
    # 이미 폴더 유효한 상태라 기록 안 함.
    if pending_files and not saved:
        try:
            from dashboard.utils.redis_client import get_redis_client
            _rc = get_redis_client().redis
            _rc.set(f'license_pending:{code}', json.dumps({
                'channel': channel, 'thread_ts': thread_ts,
                'files': pending_files, 'created': time.time(),
            }), ex=14 * 86400)
            logger.info(
                f'[LICENSE] {code} 폴더 미등록 실패 {len(pending_files)}장 → pending 기록'
                f'(폴더 경로 등록 시 자동 재시도)')
        except Exception as exc:
            logger.warning(f'[LICENSE] pending 기록 실패 ({code}): {exc}')

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

    # 저장된 사업자등록증의 OCR 결과(ocr_result)로 법인명·상호 시트 자동 반영.
    #   루프에서 문서 종류 판별 시 이미 OCR 했으므로 재호출 안 함 (API cost 절감).
    business_name = ''
    biz_update_status = ''  # '' | 'saved' | 'match' | 'mismatch' | 'error'
    biz_update_existing = ''
    not_license = not_license_any  # 사업자등록증 아닌 파일이 하나라도 있었나 (안내용)
    is_card = is_card_any          # 카드 이미지 감지
    if ocr_result:
        business_name = (ocr_result.get('name') or '')
        if business_name:
            try:
                biz_update_status, biz_update_existing = _maybe_update_business_name(code, business_name)
            except Exception as exc:
                logger.warning(f'[LICENSE/OCR] 사업자명 자동 반영 실패: {exc}')
                biz_update_status = 'error'
            # OCR 사업자명이 시트에 새로 저장됐으면 원본 공사 확정 카드도 재렌더 (2026-08-10).
            # 위 배지 refresh(saved 블록)는 OCR 저장 전에 실행돼 사업자명='-' 상태로 카드가
            # 그려짐 → 저장 후 한 번 더 갱신해 사업자명을 카드에 반영. 시트 실제값 기준
            # (force_refresh) 으로 latest_data 주입해 캐시 지연 회피.
            if biz_update_status == 'saved':
                try:
                    from dashboard.services.project_service import get_project_records
                    from dashboard.services.project_slack_notifier import (
                        refresh_project_card_license as _refresh_card,
                    )
                    _recs = get_project_records(force_refresh=True) or []
                    _proj = next(
                        (r for r in _recs
                         if (r.get('프로젝트 코드') or '').strip() == code), None
                    )
                    if _proj:
                        _refresh_card(
                            code, latest_data=_proj,
                            fallback_channel=channel, fallback_message_ts=thread_ts,
                        )
                except Exception as exc:
                    logger.warning(
                        f'[LICENSE/OCR] 사업자명 반영 후 카드 갱신 실패 ({code}): {exc}'
                    )

    return {
        'code': code, 'saved': saved, 'skipped': skipped,
        'thread_ts': thread_ts, 'channel': channel,
        'business_name': business_name,
        'biz_update_status': biz_update_status,
        'biz_update_existing': biz_update_existing,
        'not_license': not_license,
        'is_card': is_card,
        'ambiguous': ambiguous_any,
    }


def retry_pending_license(code: str, folder_hint: str = '') -> dict:
    """폴더 미등록으로 실패했던 사업자등록증 첨부를, 폴더 경로가 나중에 등록되면 자동 재저장.

    notify_project_field_changes(견적서 폴더 경로 변경 시)에서 호출. write-behind 시트 반영
    지연을 피하려 folder_hint(방금 등록된 폴더값)를 받아 그 폴더에 바로 저장. 성공 시
    배지 갱신 + '저장 완료' 스레드 댓글 + pending 키 삭제. (2026-08-28)
    """
    result = {'code': code, 'retried': 0, 'saved': []}
    try:
        from dashboard.utils.redis_client import get_redis_client
        rc = get_redis_client().redis
    except Exception:
        return result
    key = f'license_pending:{code}'
    raw = rc.get(key)
    if not raw:
        return result
    try:
        pend = json.loads(raw.decode() if isinstance(raw, bytes) else raw)
    except Exception:
        rc.delete(key)
        return result

    # 저장할 폴더 ID 결정: hint(방금 등록값) 우선 → 없으면 시트 fresh 재조회
    parent = ''
    hint = str(folder_hint or '').strip()
    if hint:
        try:
            from dashboard.utils.google_drive import resolve_folder_id
            res = resolve_folder_id(hint)
            if res.get('reason') not in ('not_id_format', 'file_no_parent', 'empty'):
                v = str(res.get('value') or '').strip()
                if re.fullmatch(r'[a-zA-Z0-9_-]{20,}', v):
                    parent = v
        except Exception as exc:
            logger.warning(f'[LICENSE/재시도] folder_hint 정규화 실패 ({code}): {exc}')
    if not parent:
        parent = _project_folder_id_for(code, fresh=True) or ''
    if not parent:
        logger.info(f'[LICENSE/재시도] {code} 유효 폴더 아직 없음 — pending 유지')
        return result

    token = os.getenv('SLACK_PROJECT_BOT_TOKEN', '').strip()
    channel = pend.get('channel')
    thread_ts = pend.get('thread_ts')
    files = pend.get('files') or []
    saved = []
    for f in files:
        name = f.get('name') or ''
        url = f.get('url') or ''
        mime = f.get('mime') or ''
        if not url:
            continue
        try:
            content = _download_slack_file(url, token)
            r = save_business_license(code, content, name, mime, parent_override=parent)
            if r.get('ok'):
                saved.append(r['file_name'])
            else:
                logger.warning(f'[LICENSE/재시도] 저장 실패 ({code}, {name}): {r.get("reason")}')
        except Exception as exc:
            logger.warning(f'[LICENSE/재시도] 처리 예외 ({code}, {name}): {exc}')
    result['retried'] = len(files)
    result['saved'] = saved
    if not saved:
        return result  # pending 유지 (다음 기회 재시도)

    # 배지 갱신
    try:
        from dashboard.services.project_slack_notifier import refresh_project_card_license
        refresh_project_card_license(code, fallback_channel=channel, fallback_message_ts=thread_ts)
    except Exception as exc:
        logger.warning(f'[LICENSE/재시도] 배지 갱신 실패 ({code}): {exc}')
    # '저장 완료' 스레드 댓글
    if token and channel and thread_ts:
        try:
            from slack_sdk import WebClient
            WebClient(token=token).chat_postMessage(
                channel=channel, thread_ts=thread_ts,
                text=(':white_check_mark: 폴더 경로 등록 확인 — 사업자등록증 자동 저장 완료\n'
                      + '\n'.join(f'  • {fn}' for fn in saved)),
                unfurl_links=False,
            )
        except Exception as exc:
            logger.warning(f'[LICENSE/재시도] 완료 댓글 실패 ({code}): {exc}')
    rc.delete(key)
    logger.info(f'[LICENSE/재시도] {code} 폴더 등록 후 {len(saved)}장 자동 저장 완료')
    return result


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
