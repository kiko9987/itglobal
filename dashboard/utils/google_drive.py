"""Google Drive API 헬퍼 — lead별 폴더 생성 + 사진 파일 업로드.

- 서비스 계정(credentials.json) 인증
- 루트 폴더 ID 환경변수 GOOGLE_DRIVE_VISIT_FOLDER_ID
"""

import io
import os
from typing import Optional

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

_SCOPES = [
    'https://www.googleapis.com/auth/drive',
]

_drive_service = None


def _get_drive_service():
    """Google Drive API client (lazy 초기화)."""
    global _drive_service
    if _drive_service is not None:
        return _drive_service
    cred_file = os.getenv('GOOGLE_CREDENTIALS_FILE', 'credentials.json').strip()
    if not os.path.isabs(cred_file):
        cred_file = os.path.join(os.getcwd(), cred_file)
    if not os.path.exists(cred_file):
        logger.error(f"[DRIVE] credentials 파일 없음: {cred_file}")
        return None
    try:
        creds = Credentials.from_service_account_file(cred_file, scopes=_SCOPES)
        _drive_service = build('drive', 'v3', credentials=creds, cache_discovery=False)
        return _drive_service
    except Exception as exc:
        logger.error(f"[DRIVE] 인증 실패: {exc}", exc_info=True)
        return None


def find_or_create_folder(name: str, parent_id: str) -> Optional[dict]:
    """parent_id 안에서 name과 일치하는 폴더를 찾거나 새로 생성. {'id', 'webViewLink'} 반환."""
    service = _get_drive_service()
    if not service or not parent_id:
        return None
    try:
        # 동일 이름 폴더 검색 (이름 단일 인자에 작은따옴표 escape)
        safe_name = name.replace("'", "\\'")
        query = (
            f"name = '{safe_name}' and "
            f"'{parent_id}' in parents and "
            f"mimeType = 'application/vnd.google-apps.folder' and "
            f"trashed = false"
        )
        resp = service.files().list(
            q=query, fields='files(id, name, webViewLink)',
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        files = resp.get('files', [])
        if files:
            return files[0]

        # 새 폴더 생성
        meta = {
            'name': name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id],
        }
        created = service.files().create(
            body=meta, fields='id, name, webViewLink',
            supportsAllDrives=True,
        ).execute()
        logger.info(f"[DRIVE] 폴더 생성: {name} ({created.get('id')})")
        return created
    except Exception as exc:
        logger.error(f"[DRIVE] 폴더 생성/조회 실패 ({name}): {exc}", exc_info=True)
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
                mimetype: str = 'application/octet-stream') -> Optional[dict]:
    """folder_id 안에 파일 업로드. {'id', 'name', 'webViewLink'} 반환."""
    service = _get_drive_service()
    if not service:
        return None
    try:
        media = MediaIoBaseUpload(
            io.BytesIO(content), mimetype=mimetype, resumable=False,
        )
        meta = {'name': filename, 'parents': [folder_id]}
        result = service.files().create(
            body=meta, media_body=media,
            fields='id, name, webViewLink',
            supportsAllDrives=True,
        ).execute()
        logger.info(
            f"[DRIVE] 파일 업로드: {filename} → {folder_id} ({result.get('id')})"
        )
        return result
    except Exception as exc:
        logger.error(f"[DRIVE] 파일 업로드 실패 ({filename}): {exc}", exc_info=True)
        return None
