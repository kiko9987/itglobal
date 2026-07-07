"""
Google Calendar 역방향 동기화 스케줄러
캘린더에서 삭제된 이벤트를 주기적으로 감지하여 Google Sheets에 반영합니다.
"""

import time
import threading
import logging
from datetime import datetime, timedelta, timezone
from typing import List, Dict, Optional
from googleapiclient.errors import HttpError
import atexit

from dashboard.utils.google_calendar import get_calendar_service, has_credentials
from dashboard.utils.user_database import get_calendar_event_repository, get_user_database
from dashboard.utils.google_sheets import GoogleSheetsManager
from dashboard.services.calendar_service import is_calendar_enabled
from dashboard.services.project_service import invalidate_project_cache

logger = logging.getLogger(__name__)


class CalendarSyncScheduler:
    """
    Google Calendar 역방향 동기화 스케줄러
    주기적으로 DB의 캘린더 이벤트가 Google Calendar에 존재하는지 확인하고,
    삭제된 이벤트를 감지하면 Google Sheets의 날짜 필드를 비웁니다.
    """

    def __init__(self, check_interval_seconds: int = 300):
        """
        Args:
            check_interval_seconds: 체크 주기 (초). 기본값 300초 (5분)
        """
        self.check_interval = check_interval_seconds
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._enabled = False

    def start(self):
        """스케줄러 시작"""
        if not is_calendar_enabled():
            logger.info("[CALENDAR_SYNC] 캘린더가 비활성화되어 동기화 스케줄러를 시작하지 않습니다.")
            return

        if not has_credentials():
            logger.info("[CALENDAR_SYNC] 캘린더 인증 정보가 없어 동기화 스케줄러를 시작하지 않습니다.")
            return

        if self._thread and self._thread.is_alive():
            logger.warning("[CALENDAR_SYNC] 스케줄러가 이미 실행 중입니다.")
            return

        self._enabled = True
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._scheduler_loop, daemon=True, name="CalendarSyncScheduler")
        self._thread.start()
        logger.info(f"[CALENDAR_SYNC] 동기화 스케줄러 시작 (체크 주기: {self.check_interval}초)")

        # 프로그램 종료 시 스케줄러 정리
        atexit.register(self.stop)

    def stop(self):
        """스케줄러 중지"""
        if not self._enabled:
            return

        logger.info("[CALENDAR_SYNC] 스케줄러 중지 요청")
        self._stop_event.set()

        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=5)
            logger.info("[CALENDAR_SYNC] 스케줄러 종료 완료")

        self._enabled = False

    def _scheduler_loop(self):
        """스케줄러 메인 루프"""
        logger.info("[CALENDAR_SYNC] 스케줄러 루프 시작")

        while not self._stop_event.is_set():
            try:
                self._sync_deleted_events()
            except Exception as e:
                logger.error(f"[CALENDAR_SYNC] 동기화 작업 중 예외 발생: {e}", exc_info=True)

            # 인터럽트 가능한 대기
            self._stop_event.wait(timeout=self.check_interval)

        logger.info("[CALENDAR_SYNC] 스케줄러 루프 종료")

    def run_sync_once(self):
        """
        수동으로 한 번 동기화를 실행하는 공개 메서드

        웹 UI에서 '캘린더 동기화' 버튼을 클릭할 때 호출됩니다.
        """
        logger.info("[CALENDAR_SYNC] 수동 동기화 요청 - run_sync_once() 호출")
        self._sync_deleted_events()

    def _sync_deleted_events(self):
        """
        DB의 모든 캘린더 이벤트를 확인하여:
        1. 삭제된 이벤트를 감지하고 처리
        2. 날짜가 변경된 이벤트를 감지하고 시트로 반영 (시트 master)
        """
        try:
            # Calendar API 서비스 초기화
            service = get_calendar_service()
        except Exception as e:
            logger.warning(f"[CALENDAR_SYNC] Calendar 서비스 초기화 실패: {e}")
            return

        # DB에서 모든 캘린더 이벤트 매핑 조회
        repo = get_calendar_event_repository()
        all_events = self._get_all_calendar_events(repo)

        if not all_events:
            logger.debug("[CALENDAR_SYNC] DB에 등록된 캘린더 이벤트가 없습니다.")
            return

        logger.debug(f"[CALENDAR_SYNC] {len(all_events)}개 이벤트 확인 시작")

        # 성능 최적화: 전체 프로젝트 데이터를 한 번만 로드하고 dict로 변환
        from dashboard.services.project_service import get_project_records
        all_projects = get_project_records()
        # KeyError 방지: 프로젝트 코드가 없거나 빈 값인 row는 건너뜀
        # 방어적 가드: all_projects에 dict가 아닌 항목이 섞여 들어오는 경우 무시 + 로그
        project_map = {}
        invalid_types = []
        for p in all_projects:
            if not isinstance(p, dict):
                invalid_types.append(type(p).__name__)
                continue
            code = p.get('프로젝트 코드')
            if code:
                project_map[code] = p

        if invalid_types:
            logger.warning(
                f"[CALENDAR_SYNC] all_projects에 비정상 항목 발견: "
                f"타입={invalid_types[:5]}{'...' if len(invalid_types) > 5 else ''} "
                f"(총 {len(invalid_types)}개, 전체 {len(all_projects)}개 중)"
            )

        logger.debug(f"[CALENDAR_SYNC] 시트 데이터 로드 완료: {len(project_map)}개 프로젝트 (dict 변환)")

        deleted_count = 0
        updated_count = 0

        for event_mapping in all_events:
            project_code = event_mapping['project_code']
            calendar_id = event_mapping['calendar_id']
            event_id = event_mapping['event_id']
            db_updated_at = event_mapping.get('updated_at')

            # Google Calendar에서 이벤트 조회
            calendar_event = self._get_calendar_event(service, calendar_id, event_id)

            if calendar_event is None:
                # 삭제된 이벤트 처리
                logger.info(f"[CALENDAR_SYNC] ❌ 삭제된 이벤트 감지: {project_code} (event_id: {event_id})")

                # Google Sheets에서 날짜 필드 비우기
                success = self._clear_project_dates(project_code)

                if success:
                    # DB에서 이벤트 매핑 삭제
                    repo.delete_event(project_code)
                    deleted_count += 1
                    logger.info(f"[CALENDAR_SYNC] ✅ 프로젝트 {project_code}의 날짜 필드를 비우고 DB 매핑 삭제 완료")
                else:
                    logger.error(f"[CALENDAR_SYNC] ⚠️ 프로젝트 {project_code}의 날짜 필드 비우기 실패")
            else:
                # 날짜 변경 감지 및 양방향 동기화 (최근 변경 기준)
                date_updated = self._sync_calendar_dates_bidirectional(
                    project_code, calendar_event, project_map, db_updated_at, repo, calendar_id
                )
                if date_updated:
                    updated_count += 1

        if deleted_count > 0 or updated_count > 0:
            logger.info(f"[CALENDAR_SYNC] 동기화 완료: {deleted_count}개 삭제, {updated_count}개 날짜 업데이트")
        else:
            logger.debug("[CALENDAR_SYNC] 변경 사항 없음")

    def _get_all_calendar_events(self, repo) -> List[Dict]:
        """
        DB에서 모든 캘린더 이벤트 매핑 조회
        """
        try:
            db = get_user_database()
            with db._get_connection() as conn:
                cursor = conn.execute('''
                    SELECT project_code, calendar_id, event_id, created_at, updated_at
                    FROM project_calendar_events
                    ORDER BY updated_at DESC
                ''')
                rows = cursor.fetchall()
                return [dict(row) for row in rows]
        except Exception as e:
            logger.error(f"[CALENDAR_SYNC] DB 조회 실패: {e}")
            return []

    def _get_calendar_event(self, service, calendar_id: str, event_id: str) -> Optional[dict]:
        """
        Google Calendar에서 이벤트 조회

        Returns:
            dict: 이벤트 정보 (존재하는 경우)
            None: 이벤트 삭제됨 또는 접근 불가
        """
        try:
            event = service.events().get(calendarId=calendar_id, eventId=event_id).execute()
            return event
        except HttpError as e:
            if e.resp.status == 404:
                # 404: 이벤트가 삭제됨
                return None
            else:
                # 기타 에러 (네트워크, 권한 등) - 빈 dict 반환 (삭제로 처리하지 않음)
                logger.warning(f"[CALENDAR_SYNC] 이벤트 확인 중 오류 (event_id: {event_id}): {e}")
                return {}  # 빈 dict: 에러 상황, 삭제로 간주하지 않음
        except Exception as e:
            logger.error(f"[CALENDAR_SYNC] 이벤트 확인 중 예외 (event_id: {event_id}): {e}")
            return {}  # 빈 dict: 예외 상황, 삭제로 간주하지 않음

    def _clear_project_dates(self, project_code: str) -> bool:
        """
        Google Sheets에서 프로젝트의 날짜 필드를 비웁니다.

        Thread-Safe: GoogleSheetsManager.batch_update_cells를 사용하여 락 보호 적용

        Args:
            project_code: 프로젝트 코드

        Returns:
            성공 여부
        """
        try:
            sheets_manager = GoogleSheetsManager()

            # 환경 변수에서 시트 정보 가져오기
            import os
            sheet_id = os.getenv('GOOGLE_SHEET_ID')
            sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황의 사본')

            if not sheet_id:
                logger.error("[CALENDAR_SYNC] GOOGLE_SHEET_ID 환경 변수가 설정되지 않았습니다.")
                return False

            # 프로젝트 코드로 행 찾기 (find_row_by_project_code는 이미 락 보호됨)
            row_index = sheets_manager.find_row_by_project_code(sheet_id, project_code, f"{sheet_name}!A:A")

            if row_index is None:
                logger.warning(f"[CALENDAR_SYNC] 프로젝트를 찾을 수 없습니다: {project_code}")
                return False

            # 날짜 필드 비우기 (공사 시작: J열, 공사 종료: K열)
            # 2026-05-22 컬럼 시프트 반영: 옛 I/J → J/K
            # batch_update_cells 사용하여 락 보호 적용
            updates = [
                {'range': f"{sheet_name}!J{row_index}", 'values': [['']]},  # 공사 시작
                {'range': f"{sheet_name}!K{row_index}", 'values': [['']]}   # 공사 종료
            ]

            sheets_manager.batch_update_cells(sheet_id, updates)

            logger.info(f"[CALENDAR_SYNC] 프로젝트 {project_code} (행 {row_index})의 날짜 필드를 비웠습니다.")

            # 캐시 무효화: 서버 측 데이터가 즉시 갱신되도록 함
            try:
                invalidate_project_cache(project_code)
                logger.info(f"[CALENDAR_SYNC] 프로젝트 {project_code} 캐시 무효화 완료")
            except Exception as cache_err:
                logger.warning(f"[CALENDAR_SYNC] 캐시 무효화 실패 (무시): {cache_err}")

            return True

        except Exception as e:
            logger.error(f"[CALENDAR_SYNC] 프로젝트 {project_code} 날짜 필드 비우기 실패: {e}", exc_info=True)
            return False

    def _update_sheet_dates(self, project_code: str, start_date: str, end_date: str) -> bool:
        """
        Google Sheets의 프로젝트 날짜 필드를 업데이트합니다.

        Thread-Safe: GoogleSheetsManager.batch_update_cells를 사용하여 락 보호 적용

        Args:
            project_code: 프로젝트 코드
            start_date: 공사 시작일 (YYYY-MM-DD)
            end_date: 공사 종료일 (YYYY-MM-DD)

        Returns:
            성공 여부
        """
        try:
            sheets_manager = GoogleSheetsManager()

            # 환경 변수에서 시트 정보 가져오기
            import os
            sheet_id = os.getenv('GOOGLE_SHEET_ID')
            sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황의 사본')

            if not sheet_id:
                logger.error("[CALENDAR_SYNC] GOOGLE_SHEET_ID 환경 변수가 설정되지 않았습니다.")
                return False

            # 프로젝트 코드로 행 찾기 (find_row_by_project_code는 이미 락 보호됨)
            row_index = sheets_manager.find_row_by_project_code(sheet_id, project_code, f"{sheet_name}!A:A")

            if row_index is None:
                logger.warning(f"[CALENDAR_SYNC] 프로젝트를 찾을 수 없습니다: {project_code}")
                return False

            # 날짜 필드 업데이트 (공사 시작: J열, 공사 종료: K열)
            # 2026-05-22 컬럼 시프트 반영: 옛 I/J → J/K
            # batch_update_cells 사용하여 락 보호 적용
            updates = [
                {'range': f"{sheet_name}!J{row_index}", 'values': [[start_date]]},  # 공사 시작
                {'range': f"{sheet_name}!K{row_index}", 'values': [[end_date]]}     # 공사 종료
            ]

            sheets_manager.batch_update_cells(sheet_id, updates)

            logger.info(
                f"[CALENDAR_SYNC] 프로젝트 {project_code} (행 {row_index}) 날짜 업데이트: "
                f"{start_date} ~ {end_date}"
            )

            # 캐시 무효화: 서버 측 데이터가 즉시 갱신되도록 함
            try:
                invalidate_project_cache(project_code)
                logger.info(f"[CALENDAR_SYNC] 프로젝트 {project_code} 캐시 무효화 완료")
            except Exception as cache_err:
                logger.warning(f"[CALENDAR_SYNC] 캐시 무효화 실패 (무시): {cache_err}")

            return True

        except Exception as e:
            logger.error(f"[CALENDAR_SYNC] 프로젝트 {project_code} 날짜 업데이트 실패: {e}", exc_info=True)
            return False

    def _sync_calendar_dates_bidirectional(
        self, project_code: str, calendar_event: dict, project_map: dict,
        db_updated_at: str, repo, calendar_id: str
    ) -> bool:
        """
        캘린더 이벤트의 날짜와 시트의 날짜를 비교하고,
        불일치 시 최근 변경된 쪽을 기준으로 양방향 동기화

        Args:
            project_code: 프로젝트 코드
            calendar_event: 캘린더 이벤트 정보
            project_map: 프로젝트 코드를 키로 하는 프로젝트 dict
            db_updated_at: DB에 저장된 마지막 동기화 시간
            repo: Calendar event repository
            calendar_id: Google Calendar ID

        Returns:
            True: 날짜 불일치가 감지되어 동기화 수행함
            False: 날짜가 일치하거나 업데이트 불필요
        """
        try:
            # 캘린더 이벤트에서 날짜 추출
            cal_start = calendar_event.get('start', {}).get('date')
            cal_end_raw = calendar_event.get('end', {}).get('date')
            cal_updated = calendar_event.get('updated')  # RFC3339 형식

            if not cal_start:
                logger.debug(f"[CALENDAR_SYNC] 캘린더 이벤트에 날짜 정보 없음: {project_code}")
                return False

            # Google Calendar all-day 이벤트는 end가 exclusive (다음날)이므로 1일 빼기
            cal_end = None
            if cal_end_raw:
                try:
                    end_dt = datetime.strptime(cal_end_raw, '%Y-%m-%d') - timedelta(days=1)
                    cal_end = end_dt.strftime('%Y-%m-%d')
                except ValueError:
                    logger.warning(f"[CALENDAR_SYNC] 캘린더 종료일 파싱 실패: {cal_end_raw}")

            # Google Sheets에서 프로젝트 데이터 조회 (O(1) dict 조회)
            sheet_project = project_map.get(project_code)

            if not sheet_project:
                logger.warning(f"[CALENDAR_SYNC] 시트에서 프로젝트를 찾을 수 없음: {project_code}")
                return False

            # None 안전 처리: None이면 빈 문자열로 fallback
            sheet_start = (sheet_project.get('공사 시작') or '').strip()
            sheet_end = (sheet_project.get('공사 종료') or '').strip()

            # 날짜 비교
            dates_match = (cal_start == sheet_start) and (cal_end == sheet_end or (not cal_end and not sheet_end))

            if dates_match:
                # 일치: 아무 작업 안 함
                return False

            # 불일치 감지: 어느 쪽이 최신인지 판단
            logger.warning(
                f"⚠️ [CALENDAR_SYNC] 날짜 불일치 감지: {project_code}\n"
                f"  - 시트: {sheet_start} ~ {sheet_end}\n"
                f"  - 캘린더: {cal_start} ~ {cal_end}"
            )

            # 시간 비교: 캘린더의 updated vs DB의 updated_at
            calendar_is_newer = False
            if cal_updated and db_updated_at:
                try:
                    # RFC3339 → datetime 변환 (aware datetime, UTC)
                    # Google Calendar updated: '2024-01-15T12:34:56.789Z'
                    cal_time_str = cal_updated.replace('Z', '+00:00')
                    cal_time = datetime.fromisoformat(cal_time_str)

                    # DB 시간 파싱 (naive → aware 변환, UTC 가정)
                    # DB updated_at: '2025-10-26 04:20:00' 또는 '2025-10-26T04:20:00Z'
                    db_time_str = db_updated_at.replace('Z', '+00:00')  # Z 포맷 대비
                    db_time = datetime.fromisoformat(db_time_str)
                    if db_time.tzinfo is None:
                        # naive datetime을 UTC aware로 변환
                        db_time = db_time.replace(tzinfo=timezone.utc)

                    calendar_is_newer = cal_time > db_time
                    logger.debug(
                        f"[CALENDAR_SYNC] 시간 비교: 캘린더={cal_updated}, DB={db_updated_at}, "
                        f"캘린더가 최신={calendar_is_newer}"
                    )
                except Exception as e:
                    logger.warning(f"[CALENDAR_SYNC] 시간 비교 실패, 기본값(시트 우선) 사용: {e}")
                    calendar_is_newer = False

            if calendar_is_newer:
                # 캘린더가 최신 → 시트 업데이트
                logger.debug(f"  → 캘린더가 최신: 시트를 캘린더 날짜로 업데이트")
                success = self._update_sheet_dates(project_code, cal_start, cal_end or cal_start)
                if success:
                    # DB 업데이트 시간 갱신 (기존 calendar_id와 event_id 유지)
                    event_id = calendar_event.get('id')
                    if event_id:
                        repo.upsert_event(project_code, calendar_id, event_id)
                    logger.debug(f"✅ [CALENDAR_SYNC] 시트를 캘린더 날짜로 동기화 완료: {project_code}")
                    return True
                else:
                    logger.error(f"[CALENDAR_SYNC] 시트 업데이트 실패: {project_code}")
                    return False
            else:
                # 시트가 최신 또는 시간 비교 불가 → 캘린더 업데이트
                logger.debug(f"  → 시트가 최신: 캘린더를 시트 날짜로 업데이트")
                from dashboard.services.calendar_service import update_project_calendar_event
                update_project_calendar_event(sheet_project)
                logger.debug(f"✅ [CALENDAR_SYNC] 캘린더를 시트 날짜로 동기화 완료: {project_code}")
                return True

        except Exception as e:
            logger.error(f"[CALENDAR_SYNC] 양방향 동기화 실패 ({project_code}): {e}", exc_info=True)
            return False

# 전역 스케줄러 인스턴스
_scheduler_instance: Optional[CalendarSyncScheduler] = None


def get_calendar_sync_scheduler() -> CalendarSyncScheduler:
    """전역 스케줄러 인스턴스 반환"""
    global _scheduler_instance
    if _scheduler_instance is None:
        _scheduler_instance = CalendarSyncScheduler()
    return _scheduler_instance


def start_calendar_sync_scheduler():
    """스케줄러 시작 (Flask 앱 초기화 시 호출)"""
    scheduler = get_calendar_sync_scheduler()
    scheduler.start()


def stop_calendar_sync_scheduler():
    """스케줄러 중지"""
    global _scheduler_instance
    if _scheduler_instance:
        _scheduler_instance.stop()
