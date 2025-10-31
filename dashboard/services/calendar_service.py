"""
Google Calendar integration service for project schedules.
"""

from __future__ import annotations

import logging
import os
from datetime import datetime, timedelta
from typing import Optional

from googleapiclient.errors import HttpError

from dashboard.utils.google_calendar import (
    CalendarConfigurationError,
    CalendarCredentialsError,
    get_calendar_service,
    has_credentials,
    get_client_secret_file,
)
from dashboard.utils.user_database import get_calendar_event_repository

logger = logging.getLogger(__name__)


def _calendar_id() -> str:
    calendar_id = os.getenv('GOOGLE_CALENDAR_ID')
    if not calendar_id:
        raise CalendarConfigurationError('환경 변수 GOOGLE_CALENDAR_ID 가 설정되지 않았습니다.')
    return calendar_id


def is_calendar_enabled() -> bool:
    try:
        get_client_secret_file()
    except CalendarConfigurationError:
        return False
    calendar_id = os.getenv('GOOGLE_CALENDAR_ID')
    return bool(calendar_id)


def _parse_date(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        return datetime.strptime(value, '%Y-%m-%d')
    except ValueError:
        return None


def _build_event_summary(project: dict) -> str:
    parts = [
        project.get('프로젝트 코드'),
        project.get('시공자'),
        project.get('공사 내용') or project.get('현장 주소') or project.get('거래처') or project.get('사업자')
    ]
    return ' · '.join(filter(None, parts))


def _build_event_description(project: dict) -> str:
    lines = [
        f"프로젝트 코드: {project.get('프로젝트 코드')}",
        f"공사 내용: {project.get('공사 내용') or '-'}",
        f"사업자: {project.get('사업자') or '-'}",
        f"담당자: {project.get('담당자') or '-'}",
        f"현장 주소: {project.get('현장 주소') or '-'}",
        f"고객명: {project.get('현장 담당자') or '-'}",
        f"연락처: {project.get('담당자 연락처') or '-'}",
    ]
    return '\n'.join(lines)


def _build_event_body(project: dict) -> Optional[dict]:
    start_dt = _parse_date(project.get('공사 시작'))
    if not start_dt:
        return None

    end_dt = _parse_date(project.get('공사 종료')) or start_dt
    if end_dt < start_dt:
        end_dt = start_dt

    end_for_all_day = end_dt + timedelta(days=1)  # Google Calendar all-day events are exclusive at end date.

    event = {
        'summary': _build_event_summary(project),
        'location': project.get('현장 주소', ''),
        'description': _build_event_description(project),
        'start': {
            'date': start_dt.strftime('%Y-%m-%d')
        },
        'end': {
            'date': end_for_all_day.strftime('%Y-%m-%d')
        },
        'reminders': {
            'useDefault': True
        }
    }
    return event


def create_project_calendar_event(project: dict) -> Optional[str]:
    """
    Create a new calendar event for the given project.
    Returns the Google Calendar event ID on success.
    """
    if not is_calendar_enabled():
        logger.debug("[CALENDAR] 비활성화 상태 - 이벤트 생성 건너뜀")
        return None

    event_body = _build_event_body(project)
    if not event_body:
        logger.debug(f"[CALENDAR] 공사 시작일이 없어 이벤트를 생성하지 않습니다: {project.get('프로젝트 코드')}")
        return None

    try:
        service = get_calendar_service()
    except CalendarCredentialsError as exc:
        logger.warning(f"[CALENDAR] 토큰이 없어 이벤트 생성 불가: {exc}")
        return None

    calendar_id = _calendar_id()

    try:
        created_event = service.events().insert(calendarId=calendar_id, body=event_body).execute()
        event_id = created_event.get('id')
        if event_id:
            repo = get_calendar_event_repository()
            repo.upsert_event(project.get('프로젝트 코드'), calendar_id, event_id)
            logger.info(f"[CALENDAR] 이벤트 생성 완료: {project.get('프로젝트 코드')} → {event_id}")
        return event_id
    except HttpError as exc:
        logger.error(f"[CALENDAR] 이벤트 생성 실패({project.get('프로젝트 코드')}): {exc}")
        return None


def update_project_calendar_event(project: dict, old_project_code: Optional[str] = None) -> None:
    """
    Update an existing calendar event if it exists.
    """
    if not is_calendar_enabled():
        return

    if not has_credentials():
        logger.debug("[CALENDAR] 토큰이 없어 업데이트 건너뜀")
        return

    repo = get_calendar_event_repository()
    project_code = project.get('프로젝트 코드')
    entry = repo.get_event(project_code)

    if not entry and old_project_code:
        entry = repo.get_event(old_project_code)
        if entry:
            repo.update_project_code(old_project_code, project_code)

    if not entry:
        logger.debug(f"[CALENDAR] 등록된 이벤트가 없어 업데이트를 건너뜁니다: {project_code}")
        return

    event_body = _build_event_body(project)
    calendar_id = entry['calendar_id']
    event_id = entry['event_id']

    try:
        service = get_calendar_service()
    except CalendarCredentialsError as exc:
        logger.warning(f"[CALENDAR] 토큰이 없어 이벤트 업데이트 불가: {exc}")
        return

    try:
        if event_body:
            service.events().patch(calendarId=calendar_id, eventId=event_id, body=event_body).execute()
            repo.upsert_event(project_code, calendar_id, event_id)
            logger.info(f"[CALENDAR] 이벤트 업데이트 완료: {project_code}")
        else:
            # 공사 시작일이 사라진 경우 이벤트 삭제
            delete_project_calendar_event(project_code)
    except HttpError as exc:
        if exc.resp.status == 404:
            logger.warning(f"[CALENDAR] 이벤트를 찾을 수 없어 매핑을 삭제합니다: {project_code}")
            repo.delete_event(project_code)
        else:
            logger.error(f"[CALENDAR] 이벤트 업데이트 실패({project_code}): {exc}")


def delete_project_calendar_event(project_code: str) -> None:
    """
    Delete a calendar event associated with the given project.
    """
    if not is_calendar_enabled():
        return

    repo = get_calendar_event_repository()
    entry = repo.get_event(project_code)
    if not entry:
        logger.debug(f"[CALENDAR] 삭제할 이벤트가 없습니다: {project_code}")
        return

    try:
        service = get_calendar_service()
    except CalendarCredentialsError:
        logger.warning("[CALENDAR] 토큰이 없어 이벤트 삭제를 진행하지 못했습니다.")
        return

    try:
        service.events().delete(calendarId=entry['calendar_id'], eventId=entry['event_id']).execute()
        logger.info(f"[CALENDAR] 이벤트 삭제 완료: {project_code}")
    except HttpError as exc:
        if exc.resp.status != 404:
            logger.error(f"[CALENDAR] 이벤트 삭제 실패({project_code}): {exc}")
    finally:
        repo.delete_event(project_code)
