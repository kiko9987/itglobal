"""Google Drive API 헬퍼 — lead별 폴더 생성 + 사진 파일 업로드.

- 서비스 계정(credentials.json) 인증
- 루트 폴더 ID 환경변수 GOOGLE_DRIVE_VISIT_FOLDER_ID
"""

import io
import os
import time
from typing import Optional, List, Dict

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseUpload

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

# 재시도 대상 HTTP 상태 — Drive API 는 quota/일시 장애에서 이 코드들을 반환.
# 429 rate limit (bulk 업로드 시 자주), 5xx 는 서버 측 일시 오류.
_DRIVE_RETRIABLE_STATUS = frozenset({429, 500, 502, 503, 504})

_SCOPES = [
    'https://www.googleapis.com/auth/drive',
]

# Thread-safety: googleapiclient 는 not thread-safe (2026-07-09 heap corruption
# 사고 참조: GoogleSheetsManager 프로세스 싱글톤 → 크래시). 스레드별 격리 필수.
# threading.local() 로 각 스레드가 자기 service 인스턴스 갖도록.
import threading as _threading
_drive_service_tls = _threading.local()


def _get_drive_service():
    """스레드별 Google Drive API client (lazy 초기화).

    2026-07-10: 이전 전역 싱글톤이었으나 Waitress request 스레드와 방문 사진
    업로드 daemon 스레드가 공유하면 heap corruption 위험 → threading.local 로 격리.
    """
    svc = getattr(_drive_service_tls, 'service', None)
    if svc is not None:
        return svc
    cred_file = os.getenv('GOOGLE_CREDENTIALS_FILE', 'credentials.json').strip()
    if not os.path.isabs(cred_file):
        cred_file = os.path.join(os.getcwd(), cred_file)
    if not os.path.exists(cred_file):
        logger.error(f"[DRIVE] credentials 파일 없음: {cred_file}")
        return None
    try:
        creds = Credentials.from_service_account_file(cred_file, scopes=_SCOPES)
        svc = build('drive', 'v3', credentials=creds, cache_discovery=False)
        _drive_service_tls.service = svc
        return svc
    except Exception as exc:
        logger.error(f"[DRIVE] 인증 실패: {exc}", exc_info=True)
        return None


def find_or_create_folder(name: str, parent_id: str) -> Optional[dict]:
    """parent_id 안에서 name과 일치하는 폴더를 찾거나 새로 생성. {'id', 'webViewLink'} 반환.

    2026-07-09 네트워크 timeout 재시도 (최대 3회 지수 백오프).
    """
    import http.client as _http_client
    import socket
    import time as _time

    service = _get_drive_service()
    if not service or not parent_id:
        return None

    safe_name = name.replace("'", "\\'")
    query = (
        f"name = '{safe_name}' and "
        f"'{parent_id}' in parents and "
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )
    meta = {
        'name': name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id],
    }

    last_exc = None
    for attempt in range(3):
        try:
            resp = service.files().list(
                q=query, fields='files(id, name, webViewLink)',
                supportsAllDrives=True, includeItemsFromAllDrives=True,
            ).execute()
            files = resp.get('files', [])
            if files:
                return files[0]

            created = service.files().create(
                body=meta, fields='id, name, webViewLink',
                supportsAllDrives=True,
            ).execute()
            logger.info(f"[DRIVE] 폴더 생성: {name} ({created.get('id')})")
            return created
        except (TimeoutError, socket.timeout, _http_client.IncompleteRead, ConnectionError) as exc:
            last_exc = exc
            wait = 0.5 * (attempt + 1)
            logger.warning(
                f"[DRIVE] 네트워크 에러 ({type(exc).__name__}) 재시도 {attempt+1}/3 "
                f"— {wait}s 후 ({name})"
            )
            _time.sleep(wait)
        except Exception as exc:
            logger.error(f"[DRIVE] 폴더 생성/조회 실패 ({name}): {exc}", exc_info=True)
            return None
    logger.error(f"[DRIVE] 폴더 생성/조회 3회 재시도 모두 실패 ({name}): {last_exc}")
    return None


def search_folder_by_name(name: str, limit: int = 10) -> List[Dict]:
    """이름으로 폴더 검색 (parent 무관, 공유 드라이브 포함).

    Windows 로컬 경로에서 추출한 폴더 이름으로 Drive에서 매칭 폴더 찾기용.
    각 결과에 parents 필드 포함 → 동명 폴더 disambiguation에 사용.

    Args:
        name: 정확 일치 검색할 폴더 이름
        limit: 최대 결과 수 (기본 10)

    Returns:
        List of {'id': str, 'name': str, 'parents': List[str]} dicts. 실패 시 빈 리스트.
    """
    service = _get_drive_service()
    if not service or not name:
        return []
    try:
        safe_name = name.replace("'", "\\'")
        query = (
            f"name = '{safe_name}' and "
            f"mimeType = 'application/vnd.google-apps.folder' and "
            f"trashed = false"
        )
        resp = service.files().list(
            q=query,
            fields='files(id, name, parents)',
            pageSize=limit,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            corpora='allDrives',  # 공유 드라이브까지 검색 범위 확장 (2026-07-07 발견: 이거 없으면 옛 폴더 못 찾음)
        ).execute()
        return resp.get('files', []) or []
    except Exception as exc:
        logger.warning(f"[DRIVE] 폴더 이름 검색 실패 ({name}): {exc}")
        return []


def get_folder_name(folder_id: str) -> Optional[str]:
    """폴더 ID로 이름 조회. 없으면 None."""
    service = _get_drive_service()
    if not service or not folder_id:
        return None
    try:
        resp = service.files().get(
            fileId=folder_id, fields='name',
            supportsAllDrives=True,
        ).execute()
        return resp.get('name') or None
    except Exception as exc:
        logger.debug(f"[DRIVE] 폴더 이름 조회 실패 ({folder_id}): {exc}")
        return None


def rename_folder(folder_id: str, new_name: str) -> bool:
    """폴더 이름 변경. 성공 시 True."""
    service = _get_drive_service()
    if not service or not folder_id:
        return False
    try:
        service.files().update(
            fileId=folder_id, body={'name': new_name},
            fields='id, name', supportsAllDrives=True,
        ).execute()
        logger.info(f"[DRIVE] 폴더 이름 변경: {folder_id} → {new_name}")
        return True
    except Exception as exc:
        logger.error(f"[DRIVE] 폴더 이름 변경 실패 ({folder_id} → {new_name}): {exc}",
                     exc_info=True)
        return False


def upload_file(folder_id: str, filename: str, content: bytes,
                mimetype: str = 'application/octet-stream',
                max_retries: int = 3) -> Optional[dict]:
    """folder_id 안에 파일 업로드. {'id', 'name', 'webViewLink'} 반환.

    HTTP 429/5xx 는 지수 백오프로 최대 max_retries 회 재시도.
    현장 사진 50~60장 bulk 업로드 시 Drive quota 로 429 발생 가능성 대응.
    """
    service = _get_drive_service()
    if not service:
        return None
    meta = {'name': filename, 'parents': [folder_id]}
    last_exc: Optional[Exception] = None
    for attempt in range(max_retries):
        try:
            # MediaIoBaseUpload 는 재시도마다 새로 만들어야 함 (스트림 소진 방지).
            media = MediaIoBaseUpload(
                io.BytesIO(content), mimetype=mimetype, resumable=False,
            )
            result = service.files().create(
                body=meta, media_body=media,
                fields='id, name, webViewLink',
                supportsAllDrives=True,
            ).execute()
            logger.info(
                f"[DRIVE] 파일 업로드: {filename} → {folder_id} ({result.get('id')})"
            )
            return result
        except HttpError as exc:
            status = getattr(getattr(exc, 'resp', None), 'status', 0)
            try:
                status = int(status)
            except (TypeError, ValueError):
                status = 0
            if status in _DRIVE_RETRIABLE_STATUS and attempt < max_retries - 1:
                delay = 2 ** attempt + 0.5  # 1.5, 2.5, 4.5s
                logger.warning(
                    f"[DRIVE] 업로드 재시도 예정 ({filename}) — HTTP {status} "
                    f"[{attempt + 1}/{max_retries}] {delay}s sleep"
                )
                time.sleep(delay)
                last_exc = exc
                continue
            logger.error(
                f"[DRIVE] 파일 업로드 실패 ({filename}) HTTP {status}: {exc}",
                exc_info=True,
            )
            return None
        except Exception as exc:
            logger.error(f"[DRIVE] 파일 업로드 실패 ({filename}): {exc}", exc_info=True)
            return None
    logger.error(
        f"[DRIVE] 파일 업로드 최종 실패 ({filename}) — {max_retries}회 재시도 후 포기: "
        f"{last_exc}"
    )
    return None
