"""
Google Calendar OAuth & service helpers
"""

from __future__ import annotations

import json
import logging
import os
from dataclasses import dataclass
from typing import Optional

from flask import current_app
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logger = logging.getLogger(__name__)

CALENDAR_SCOPES = ['https://www.googleapis.com/auth/calendar.events']


class CalendarConfigurationError(Exception):
    """Raised when calendar integration is not properly configured."""


class CalendarCredentialsError(Exception):
    """Raised when calendar credentials are missing or invalid."""


def _project_root() -> str:
    return os.path.dirname(os.path.dirname(os.path.dirname(__file__)))


def _resolve_path(path: str) -> str:
    if os.path.isabs(path):
        return path
    return os.path.abspath(os.path.join(_project_root(), path))


def get_client_secret_file() -> str:
    client_secret = os.getenv('GOOGLE_CALENDAR_CLIENT_SECRET_FILE')
    if not client_secret:
        raise CalendarConfigurationError('환경 변수 GOOGLE_CALENDAR_CLIENT_SECRET_FILE 이 설정되지 않았습니다.')

    resolved = _resolve_path(client_secret)
    if not os.path.exists(resolved):
        raise CalendarConfigurationError(f'Google Calendar OAuth client 파일을 찾을 수 없습니다: {resolved}')
    return resolved


def get_token_file() -> str:
    token_file = os.getenv('GOOGLE_CALENDAR_TOKEN_FILE')
    if token_file:
        resolved = _resolve_path(token_file)
    else:
        try:
            instance_path = current_app.instance_path  # type: ignore[attr-defined]
        except RuntimeError:
            instance_path = os.path.join(_project_root(), 'instance')
        if not os.path.exists(instance_path):
            os.makedirs(instance_path, exist_ok=True)
        resolved = os.path.join(instance_path, 'google_calendar_token.json')

    os.makedirs(os.path.dirname(resolved), exist_ok=True)
    return resolved


def create_flow(redirect_uri: str, state: Optional[str] = None) -> Flow:
    """Create an OAuth flow for Google Calendar."""
    client_secret = get_client_secret_file()
    flow = Flow.from_client_secrets_file(
        client_secret,
        scopes=CALENDAR_SCOPES,
        state=state
    )
    flow.redirect_uri = redirect_uri
    return flow


def load_credentials() -> Optional[Credentials]:
    token_file = get_token_file()
    if not os.path.exists(token_file):
        return None
    try:
        credentials = Credentials.from_authorized_user_file(token_file, CALENDAR_SCOPES)
        return credentials
    except (ValueError, json.JSONDecodeError) as exc:
        # 2026-07-09 Calendar 도입 유예 중 — 반복 노이즈 방지 debug 로 낮춤.
        logger.debug(f"[CALENDAR] 토큰 파일 파싱 실패: {exc}")
        return None


def save_credentials(credentials: Credentials) -> None:
    token_file = get_token_file()
    with open(token_file, 'w', encoding='utf-8') as token_handle:
        token_handle.write(credentials.to_json())
    logger.info(f"[CALENDAR] 토큰 저장 완료: {token_file}")


def has_credentials() -> bool:
    credentials = load_credentials()
    if not credentials:
        return False
    if credentials.expired and not credentials.refresh_token:
        return False
    return True


def get_calendar_service():
    """Google Calendar API service 반환.

    Thread-safety: 매번 새 build() 호출로 새 service 인스턴스 반환.
    googleapiclient 는 not thread-safe (GoogleSheetsManager 사고 참조,
    2026-07-09 heap corruption) 이라 캐싱하지 말고 호출자별로 새로 받아야 함.
    캐싱 유혹 시 반드시 threading.local() 로.
    """
    credentials = load_credentials()
    if not credentials:
        raise CalendarCredentialsError('Google Calendar 토큰이 없습니다. 먼저 OAuth 연동을 진행해주세요.')

    if credentials.expired and credentials.refresh_token:
        try:
            credentials.refresh(Request())
            save_credentials(credentials)
            logger.info("[CALENDAR] 토큰을 성공적으로 갱신했습니다.")
        except Exception as exc:
            logger.error(f"[CALENDAR] 토큰 갱신 실패: {exc}")
            raise CalendarCredentialsError('Google Calendar 토큰 갱신에 실패했습니다.') from exc

    if credentials.expired and not credentials.refresh_token:
        raise CalendarCredentialsError('Google Calendar 토큰이 만료되었고 refresh token이 없습니다.')

    try:
        service = build('calendar', 'v3', credentials=credentials, cache_discovery=False)
        return service
    except HttpError as exc:
        logger.error(f"[CALENDAR] Google Calendar 서비스 초기화 실패: {exc}")
        raise CalendarCredentialsError('Google Calendar 서비스 초기화에 실패했습니다.') from exc

