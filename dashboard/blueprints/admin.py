"""
관리자 블루프린트
캐시 관리, 시스템 모니터링, 관리자 전용 기능들
"""

import logging
from datetime import datetime, timedelta
from flask import Blueprint, render_template, jsonify, request, make_response
from dashboard.auth import admin_required, login_required
from dashboard.utils.logging_config import get_logger
from dashboard.utils.error_helpers import generate_error_id

logger = get_logger(__name__)

admin_bp = Blueprint('admin', __name__, url_prefix='/admin')


def _no_cache(resp):
    """관리자 페이지 응답에 no-cache 헤더 강제. 2026-07-10 추가.
    브라우저(Chrome/Edge) 가 관리자 HTML 을 강력 캐시해 페이지 배포 후에도
    Ctrl+Shift+R 로도 반영 안 되는 사고 관측 → 관리자 페이지는 매번 최신을
    받도록 강제. 관리자 트래픽은 소수라 캐시 이득보다 정합성이 훨씬 중요.
    """
    resp.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    resp.headers['Pragma'] = 'no-cache'
    resp.headers['Expires'] = '0'
    return resp


@admin_bp.route('/cache-monitor')
@admin_required
def cache_monitor_page():
    """캐시 모니터링 페이지 (관리자 전용)"""
    return _no_cache(make_response(render_template('cache_monitor.html')))

@admin_bp.route('/docs')
@admin_required
def docs_management_page():
    """API 문서 관리 페이지 (관리자 전용)"""
    return render_template('docs_management.html')

# API 엔드포인트들
@admin_bp.route('/api/cache-stats')
@login_required
def get_cache_stats():
    """캐시 상태 모니터링 API"""
    try:
        from dashboard.utils.smart_cache_manager import get_smart_cache, cache_stats

        # 기본 캐시 통계
        basic_stats = cache_stats()

        # 스마트 캐시 통계
        smart_cache = get_smart_cache()
        smart_stats = smart_cache.get_cache_info()

        return jsonify({
            'basic_cache': basic_stats,
            'smart_cache': smart_stats,
            'timestamp': datetime.now().isoformat(),
            'status': 'healthy'
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캐시 통계 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캐시 통계를 조회할 수 없습니다',
            'error_id': error_id
        }), 500

@admin_bp.route('/api/cache-clear', methods=['POST'])
@admin_required
def clear_cache():
    """캐시 전체 삭제 API (관리자 전용)"""
    try:
        from dashboard.utils.smart_cache_manager import get_smart_cache, cache_clear
        from dashboard.utils.user_database import get_audit_repository
        from flask import session

        # 기본 캐시 삭제
        cache_clear()

        # 스마트 캐시 삭제
        smart_cache = get_smart_cache()
        smart_cache.clear()

        # 감사 로그 기록
        try:
            audit_repo = get_audit_repository()
            user_email = session.get('user', {}).get('email', 'unknown')
            audit_repo.log_action(
                user_email=user_email,
                action='CACHE_CLEAR',
                details='전체 캐시 삭제',
                field_name='cache',
                old_value='전체 캐시',
                new_value='삭제됨',
                ip_address=request.remote_addr
            )
        except Exception as log_error:
            logger.warning(f"감사 로그 기록 실패: {log_error}")

        return jsonify({
            'success': True,
            'message': '캐시가 성공적으로 삭제되었습니다.',
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캐시 삭제 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캐시를 삭제할 수 없습니다',
            'error_id': error_id
        }), 500

@admin_bp.route('/api/cache/status')
@login_required
def get_cache_status():
    """캐시 상태 조회 (모든 사용자)"""
    try:
        from dashboard.utils.smart_cache_manager import get_smart_cache
        from dashboard.utils.background_prefetch import get_background_prefetch

        cache = get_smart_cache()
        prefetch = get_background_prefetch()

        # 캐시 기본 정보
        cache_info = cache.get_cache_info()

        # 프리패치 상태 (간소화된 정보만)
        prefetch_stats = prefetch.get_stats()

        # 마지막 성공적인 프리패치 시간
        last_success = None
        has_recent_errors = False

        if 'current_sheet_data' in prefetch_stats.get('jobs', {}):
            job_info = prefetch_stats['jobs']['current_sheet_data']
            last_success = job_info.get('last_prefetch')
            has_recent_errors = job_info.get('error_count', 0) > 0

        # 캐시 건강 상태 판단
        cache_health = 'healthy'
        if not prefetch_stats.get('running', False):
            cache_health = 'warning'  # 프리패치가 실행되지 않음
        elif has_recent_errors:
            cache_health = 'error'    # 최근 오류 발생
        elif last_success:
            # 마지막 성공이 10분 이상 전이면 경고
            try:
                # ISO 형식 날짜 파싱 (YYYY-MM-DDTHH:MM:SS.ffffff)
                last_time = datetime.fromisoformat(last_success.replace('Z', '+00:00'))
                if datetime.now() - last_time > timedelta(minutes=10):
                    cache_health = 'warning'
            except Exception as e:
                logger.warning(f"처리 중 오류 무시: {e}")
                pass

        return jsonify({
            'cache_info': cache_info,
            'prefetch_running': prefetch_stats.get('running', False),
            'last_successful_update': last_success,
            'cache_health': cache_health,
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캐시 상태 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캐시 상태를 조회할 수 없습니다',
            'error_id': error_id
        }), 500

@admin_bp.route('/api/cache/refresh', methods=['POST'])
@login_required
def manual_cache_refresh():
    """수동 캐시 새로고침 (폴백 전략)"""
    try:
        from dashboard.services.project_service import invalidate_project_cache, load_data
        from flask import session

        # 기존 캐시 무효화
        invalidate_project_cache()

        # 새 데이터 강제 로드
        fresh_data = load_data()

        logger.info(f"[MANUAL_REFRESH] 사용자 요청으로 캐시 수동 새로고침 완료")

        data_count = 0
        if fresh_data is not None:
            if hasattr(fresh_data, '__len__'):
                data_count = len(fresh_data)
            elif hasattr(fresh_data, 'shape'):  # pandas DataFrame
                data_count = fresh_data.shape[0]

        # 감사 로그 기록
        try:
            from dashboard.utils.user_database import get_audit_repository
            audit_repo = get_audit_repository()
            user_email = session.get('user', {}).get('email', 'unknown')
            audit_repo.log_action(
                user_email=user_email,
                action='CACHE_REFRESH',
                details=f'수동 캐시 새로고침 (데이터 {data_count}건)',
                field_name='cache',
                old_value='이전 캐시',
                new_value='새로고침됨',
                ip_address=request.remote_addr
            )
        except Exception as log_error:
            logger.warning(f"감사 로그 기록 실패: {log_error}")

        return jsonify({
            'success': True,
            'message': f'캐시가 성공적으로 새로고침되었습니다. ({data_count}건의 데이터 로드)',
            'data_count': data_count,
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캐시 새로고침 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캐시를 새로고침할 수 없습니다',
            'error_id': error_id
        }), 500

@admin_bp.route('/api/calendar/sync-all', methods=['POST'])
@admin_required
def sync_all_projects_to_calendar():
    """
    기존 프로젝트를 캘린더에 일괄 등록 (관리자 전용)

    - 공사 시작일이 있는 모든 프로젝트 조회
    - 아직 캘린더에 등록되지 않은 프로젝트만 필터링
    - 백그라운드에서 순차 등록
    - 진행 상황을 로그로 출력

    Query Parameters:
        limit (int, optional): 등록할 최대 개수 (기본값: 제한 없음)
            예: ?limit=10 (최근 10개만 등록)

    Returns:
        JSON: 등록 시작 확인 메시지 (비동기 실행)
    """
    try:
        from dashboard.services.project_service import get_project_records
        from dashboard.services.calendar_service import create_project_calendar_event, is_calendar_enabled
        from dashboard.utils.user_database import get_calendar_event_repository
        from flask import session
        import threading

        # Query parameter로 limit 받기
        limit = request.args.get('limit', type=int)
        if limit is not None and limit <= 0:
            return jsonify({
                'success': False,
                'error': 'limit 값은 1 이상이어야 합니다.'
            }), 400

        # 캘린더 활성화 확인
        if not is_calendar_enabled():
            return jsonify({
                'success': False,
                'error': 'Google Calendar 연동이 설정되지 않았습니다.'
            }), 400

        def background_sync():
            """백그라운드 동기화 작업"""
            try:
                logger.info("[CALENDAR_SYNC_ALL] 일괄 동기화 시작")

                # 모든 프로젝트 조회
                all_projects = get_project_records()
                logger.info(f"[CALENDAR_SYNC_ALL] 전체 프로젝트: {len(all_projects)}개")

                # 공사 시작일이 있는 프로젝트만 필터링
                projects_with_date = [
                    p for p in all_projects
                    if p.get('공사 시작') and str(p.get('공사 시작')).strip()
                ]
                logger.info(f"[CALENDAR_SYNC_ALL] 공사 시작일 있음: {len(projects_with_date)}개")

                # 이미 등록된 프로젝트 확인
                repo = get_calendar_event_repository()
                registered_codes = set()
                for project in projects_with_date:
                    code = project.get('프로젝트 코드')
                    if code and repo.get_event(code):
                        registered_codes.add(code)

                logger.info(f"[CALENDAR_SYNC_ALL] 이미 등록됨: {len(registered_codes)}개")

                # 등록되지 않은 프로젝트만 추출
                projects_to_sync = [
                    p for p in projects_with_date
                    if p.get('프로젝트 코드') not in registered_codes
                ]

                # 역순으로 정렬 (최근 프로젝트부터)
                projects_to_sync = list(reversed(projects_to_sync))

                # limit 적용 (최근 프로젝트부터)
                total_candidates = len(projects_to_sync)
                if limit is not None and limit < total_candidates:
                    projects_to_sync = projects_to_sync[:limit]
                    logger.info(f"[CALENDAR_SYNC_ALL] 등록 대상: {total_candidates}개 중 최근 {limit}개만 등록")
                else:
                    logger.info(f"[CALENDAR_SYNC_ALL] 등록 대상: {len(projects_to_sync)}개")

                # 순차 등록
                success_count = 0
                error_count = 0

                for i, project in enumerate(projects_to_sync, 1):
                    project_code = project.get('프로젝트 코드')
                    try:
                        event_id = create_project_calendar_event(project)
                        if event_id:
                            success_count += 1
                            logger.info(f"[CALENDAR_SYNC_ALL] [{i}/{len(projects_to_sync)}] 등록 성공: {project_code}")
                        else:
                            error_count += 1
                            logger.warning(f"[CALENDAR_SYNC_ALL] [{i}/{len(projects_to_sync)}] 등록 실패: {project_code} (이벤트 ID 없음)")
                    except Exception as e:
                        error_count += 1
                        logger.error(f"[CALENDAR_SYNC_ALL] [{i}/{len(projects_to_sync)}] 등록 오류: {project_code} - {e}")

                logger.info(
                    f"[CALENDAR_SYNC_ALL] 일괄 동기화 완료 - "
                    f"성공: {success_count}개, 실패: {error_count}개, "
                    f"이미 등록: {len(registered_codes)}개"
                )

            except Exception as e:
                logger.error(f"[CALENDAR_SYNC_ALL] 백그라운드 동기화 오류: {e}", exc_info=True)

        # 백그라운드 스레드로 실행
        sync_thread = threading.Thread(target=background_sync, daemon=True)
        sync_thread.start()

        # 감사 로그 기록
        try:
            from dashboard.utils.user_database import get_audit_repository
            audit_repo = get_audit_repository()
            user_email = session.get('user', {}).get('email', 'unknown')
            audit_repo.log_action(
                user_email=user_email,
                action='CALENDAR_SYNC_ALL',
                details='기존 프로젝트 캘린더 일괄 등록 시작',
                field_name='calendar',
                old_value='미등록',
                new_value='일괄 등록 시작',
                ip_address=request.remote_addr
            )
        except Exception as log_error:
            logger.warning(f"감사 로그 기록 실패: {log_error}")

        message = '캘린더 일괄 동기화를 시작했습니다. 서버 로그에서 진행 상황을 확인하세요.'
        if limit is not None:
            message = f'캘린더 일괄 동기화를 시작했습니다 (최대 {limit}개). 서버 로그에서 진행 상황을 확인하세요.'

        return jsonify({
            'success': True,
            'message': message,
            'limit': limit,
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 캘린더 일괄 동기화 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '캘린더 일괄 동기화를 시작할 수 없습니다',
            'error_id': error_id
        }), 500


# ─────────────────────────────────────────────────────────────
# 프로젝트 시퀀스 관리 (2026-07-10)
#   매니저가 시트에서 프로젝트를 수동 삭제하면 DB 시퀀스가 시트 max 보다 커져
#   다음 신규 등록 시 gap 이 생긴다. 자동 하향 조정은 위험(정말 gap 을 두려는
#   케이스도 있음)해서 관리자가 명시적으로 리셋할 수 있는 API 만 제공.
# ─────────────────────────────────────────────────────────────

@admin_bp.route('/api/sequence/status', methods=['GET'])
@admin_required
def sequence_status():
    """DB 시퀀스 vs 시트 max 조회. gap 이 얼마인지 즉시 파악."""
    try:
        import re
        from dashboard.utils.user_database import get_user_database
        from dashboard.services.project_service import load_data

        pattern = request.args.get('pattern', 'GLOBAL').strip()
        udb = get_user_database()
        seq_info = udb.get_sequence_status(pattern) or {}
        db_current = int(seq_info.get('current_number', 0))

        # 2026-07-10 fix — load_data() 로 시트 코드 조회.
        #   기존엔 _get_payment_service() (payment_sync 싱글톤) 로 직접 시트 API 호출했는데,
        #   Flask 요청 스레드에서 실행되면 다른 스레드 googleapiclient 와 충돌해 SSL 에러
        #   (BAD_RECORD_TYPE) 발생. memory: googleapiclient thread-safety 위반.
        #   load_data() 는 threading.local 로 관리되는 GoogleSheetsManager 를 사용해 안전.
        sheet_max = 0
        top_codes: list[str] = []
        df = load_data()
        if df is not None and '프로젝트 코드' in df.columns:
            codes = df['프로젝트 코드'].astype(str).tolist()
            nums = []
            for c in codes:
                m = re.match(r'^[GRN](\d{4})-', c)
                if m:
                    nums.append(int(m.group(1)))
            if nums:
                sheet_max = max(nums)
                top_codes = [c for c in codes if re.match(r'^[GRN]\d{4}-', c)]
                top_codes = sorted(top_codes, key=lambda x: int(re.match(r'^[GRN](\d{4})-', x).group(1)), reverse=True)[:5]

        gap = db_current - sheet_max
        return jsonify({
            'success': True,
            'pattern': pattern,
            'db_current_number': db_current,
            'sheet_max_number': sheet_max,
            'gap': gap,
            'next_issued_if_no_reset': db_current + 1,
            'next_issued_if_reset': sheet_max + 1,
            'sheet_top_5_codes': top_codes,
            'last_updated': seq_info.get('last_updated', ''),
            'recommendation': (
                f'⚠ 시트에 없는 번호 {gap}개 예약됨. 리셋 권장.' if gap > 0
                else '✓ 정합 상태 (gap 없음).'
            ),
        })
    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 시퀀스 상태 조회 오류: {e}", exc_info=True)
        return jsonify({'success': False, 'error': '조회 실패', 'error_id': error_id}), 500


@admin_bp.route('/api/sequence/reset', methods=['POST'])
@admin_required
def sequence_reset():
    """DB 시퀀스를 시트 max 로 리셋. 다음 등록부터 시트 max+1 발급.

    Body (선택):
        pattern: 시퀀스 이름 (기본 GLOBAL)
        target: 명시적 목표 값 (미지정 시 시트 max)
    """
    try:
        import re
        from dashboard.utils.user_database import get_user_database
        from dashboard.services.project_service import load_data

        payload = request.get_json(silent=True) or {}
        pattern = str(payload.get('pattern', 'GLOBAL')).strip()
        explicit_target = payload.get('target')

        udb = get_user_database()
        before = udb.get_sequence_status(pattern) or {}
        db_before = int(before.get('current_number', 0))

        # 목표 값 결정
        if explicit_target is not None:
            try:
                target = int(explicit_target)
            except (TypeError, ValueError):
                return jsonify({'success': False, 'error': 'target 은 정수여야 합니다.'}), 400
        else:
            # 2026-07-10 fix — load_data() 로 안전하게 조회 (status endpoint 동일 이유).
            df = load_data()
            if df is None or '프로젝트 코드' not in df.columns:
                return jsonify({'success': False, 'error': '시트 데이터 로드 실패'}), 500
            nums = []
            for c in df['프로젝트 코드'].astype(str):
                m = re.match(r'^[GRN](\d{4})-', c)
                if m:
                    nums.append(int(m.group(1)))
            if not nums:
                return jsonify({'success': False, 'error': '시트에서 프로젝트 코드를 찾지 못했습니다.'}), 500
            target = max(nums)

        if target >= db_before:
            return jsonify({
                'success': False,
                'error': (
                    f'target ({target}) 이 현재 DB 시퀀스 ({db_before}) 보다 크거나 같습니다. '
                    f'하향 조정용 API 라 안전상 거부합니다. 강제로 올리려면 다른 경로 사용.'
                ),
            }), 400

        # SQLite UPDATE
        with udb._get_connection() as conn:
            conn.execute('BEGIN IMMEDIATE')
            try:
                conn.execute(
                    'UPDATE project_code_sequences '
                    'SET current_number = ?, last_updated = CURRENT_TIMESTAMP '
                    'WHERE pattern = ?',
                    (target, pattern),
                )
                conn.commit()
            except Exception:
                conn.rollback()
                raise

        after = udb.get_sequence_status(pattern) or {}
        logger.warning(
            f'[SEQUENCE/RESET] {pattern}: {db_before} → {target} (by admin)'
        )
        return jsonify({
            'success': True,
            'pattern': pattern,
            'before': db_before,
            'after': int(after.get('current_number', target)),
            'next_issued': int(after.get('current_number', target)) + 1,
        })
    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 시퀀스 리셋 오류: {e}", exc_info=True)
        return jsonify({'success': False, 'error': '리셋 실패', 'error_id': error_id}), 500