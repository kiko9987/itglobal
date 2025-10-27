"""
데이터 관리 블루프린트
데이터 새로고침, 캐시 관리, 데이터 로드 상태 관리 등 데이터 관리 기능들
"""

import logging
import threading
import time
from datetime import datetime
from flask import Blueprint, jsonify, request, session
from dashboard.auth import login_required, admin_required
from dashboard.utils.logging_config import get_logger
from dashboard.utils.smart_cache_manager import smart_get, CacheStrategy

logger = get_logger(__name__)

data_management_bp = Blueprint('data_management', __name__, url_prefix='/api')

# 비동기 새로고침 중복 실행 방지
_refresh_lock = threading.Lock()
_refresh_in_progress = False

@data_management_bp.route('/refresh-data')
@login_required
def refresh_data():
    """데이터 새로고침 API (강제 새로고침)"""
    try:
        from dashboard.services.project_service import load_data, get_last_update
        from dashboard.utils.smart_cache_manager import get_smart_cache

        # 캐시 초기화
        get_smart_cache().clear()

        # 구글시트에서 강제 새로고침
        df = load_data(force_refresh=True)
        if df is None:
            return jsonify({'error': '데이터 새로고침 실패'}), 500

        last_ts = get_last_update()

        # SocketIO 실시간 알림 (선택적)
        try:
            from dashboard import socketio
            socketio.emit('data_updated', {
                'message': '데이터가 업데이트되었습니다',
                'timestamp': last_ts.isoformat() if last_ts else None,
                'record_count': len(df),
                'action': 'refresh'
            })
        except ImportError:
            logger.debug("SocketIO를 사용할 수 없음")

        return jsonify({
            'success': True,
            'message': '데이터 새로고침 완료',
            'timestamp': last_ts.isoformat() if last_ts else None,
            'formatted_time': last_ts.strftime('%Y-%m-%d %H:%M:%S') if last_ts else None,
            'record_count': len(df)
        })

    except Exception as e:
        logger.error(f"데이터 새로고침 오류: {str(e)}")
        return jsonify({'success': False, 'error': '데이터 새로고침 중 오류가 발생했습니다.'}), 500

@data_management_bp.route('/quick-refresh')
@login_required
def quick_refresh_data():
    """빠른 데이터 새로고침 API (캐시 우선 사용)"""
    try:
        from dashboard.services.project_service import load_data, get_last_update

        # 캐시 우선 사용하여 빠른 로드
        df = load_data(force_refresh=False)
        if df is None:
            return jsonify({'error': '데이터 로드 실패'}), 500

        # 캐시 사용 여부 확인
        cache_key = "current_sheet_data"
        is_cached = smart_get(cache_key, CacheStrategy.CRITICAL_DATA) is not None
        last_ts = get_last_update()

        # SocketIO 실시간 알림 (선택적)
        try:
            from dashboard import socketio
            socketio.emit('data_updated', {
                'message': '빠른 새로고침 완료' + (' (캐시 사용)' if is_cached else ' (직접 로드)'),
                'timestamp': last_ts.isoformat() if last_ts else None,
                'record_count': len(df),
                'from_cache': is_cached,
                'action': 'quick_refresh'
            })
        except ImportError:
            logger.debug("SocketIO를 사용할 수 없음")

        return jsonify({
            'success': True,
            'message': '빠른 새로고침 완료' + (' (캐시 사용)' if is_cached else ' (직접 로드)'),
            'timestamp': last_ts.isoformat() if last_ts else None,
            'formatted_time': last_ts.strftime('%Y-%m-%d %H:%M:%S') if last_ts else None,
            'record_count': len(df),
            'from_cache': is_cached
        })

    except Exception as e:
        logger.error(f"빠른 새로고침 오류: {str(e)}")
        return jsonify({'success': False, 'error': '빠른 새로고침 중 오류가 발생했습니다.'}), 500

@data_management_bp.route('/cache/refresh', methods=['POST'])
@login_required
def manual_cache_refresh():
    """수동 캐시 새로고침 (폴백 전략)"""
    try:
        from dashboard.services.project_service import load_data, invalidate_project_cache

        # 기존 캐시 무효화
        invalidate_project_cache()

        # 새 데이터 강제 로드
        fresh_data = load_data()

        logger.info(f"[MANUAL_REFRESH] 사용자 요청으로 캐시 수동 새로고침 완료")

        # 데이터 개수 계산
        data_count = 0
        if fresh_data is not None:
            if hasattr(fresh_data, '__len__'):
                data_count = len(fresh_data)
            elif hasattr(fresh_data, 'shape'):  # pandas DataFrame
                data_count = fresh_data.shape[0]

        return jsonify({
            'success': True,
            'message': '캐시가 성공적으로 새로고침되었습니다.',
            'record_count': data_count,
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        logger.error(f"수동 캐시 새로고침 오류: {str(e)}")
        return jsonify({
            'success': False,
            'error': '캐시 새로고침 중 오류가 발생했습니다.'
        }), 500

@data_management_bp.route('/data-status')
@login_required
def get_data_status():
    """데이터 상태 조회 (타임스탬프, 레코드 수 등)"""
    try:
        from dashboard.services.project_service import get_last_update

        # 현재 데이터 상태 확인
        df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
        last_ts = get_last_update()

        # 캐시 상태 확인
        cache_key = "current_sheet_data"
        is_cached = smart_get(cache_key, CacheStrategy.CRITICAL_DATA) is not None

        return jsonify({
            'success': True,
            'data_loaded': df is not None,
            'record_count': len(df) if df is not None else 0,
            'last_update': last_ts.isoformat() if last_ts else None,
            'formatted_time': last_ts.strftime('%Y-%m-%d %H:%M:%S') if last_ts else None,
            'cache_status': 'cached' if is_cached else 'not_cached',
            'status_timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        logger.error(f"데이터 상태 조회 오류: {str(e)}")
        return jsonify({
            'success': False,
            'error': '데이터 상태 조회 중 오류가 발생했습니다.'
        }), 500

@data_management_bp.route('/refresh-async', methods=['POST'])
@login_required
def refresh_data_async():
    """비동기 데이터 새로고침 (백그라운드 처리)"""
    global _refresh_in_progress

    try:
        # 중복 실행 방지 체크
        with _refresh_lock:
            if _refresh_in_progress:
                return jsonify({
                    'success': False,
                    'message': '이미 백그라운드 새로고침이 진행 중입니다.',
                    'timestamp': datetime.now().isoformat()
                }), 409  # Conflict

            _refresh_in_progress = True

        from dashboard.services.project_service import load_data

        def refresh_task():
            """백그라운드에서 실행되는 새로고침 작업"""
            global _refresh_in_progress
            try:
                time.sleep(0.5)  # 구글 시트 반영 대기
                df = load_data(force_refresh=True)  # 강제 새로고침
                logger.info("백그라운드 데이터 새로고침 완료")

                # SocketIO 실시간 알림 (선택적)
                try:
                    from dashboard import socketio
                    socketio.emit('data_updated', {
                        'message': '백그라운드 데이터 새로고침 완료',
                        'timestamp': datetime.now().isoformat(),
                        'record_count': len(df) if df is not None else 0,
                        'action': 'async_refresh'
                    })
                except ImportError:
                    logger.debug("SocketIO를 사용할 수 없음")

            except Exception as e:
                logger.warning(f"백그라운드 데이터 새로고침 실패: {e}")
            finally:
                # 완료 후 플래그 해제
                with _refresh_lock:
                    _refresh_in_progress = False
                logger.debug("백그라운드 새로고침 플래그 해제")

        # 백그라운드 스레드로 실행
        thread = threading.Thread(target=refresh_task, daemon=True)
        thread.start()

        return jsonify({
            'success': True,
            'message': '백그라운드 새로고침이 시작되었습니다.',
            'timestamp': datetime.now().isoformat()
        })

    except Exception as e:
        # 오류 발생 시 플래그 해제
        with _refresh_lock:
            _refresh_in_progress = False
        logger.error(f"비동기 새로고침 오류: {str(e)}")
        return jsonify({
            'success': False,
            'error': '비동기 새로고침 시작 중 오류가 발생했습니다.'
        }), 500

def parse_amount(amount_str):
    """금액 문자열을 숫자로 변환"""
    if not amount_str:
        return 0

    # 문자열에서 숫자와 소수점, 음수 부호만 추출
    import re
    cleaned = re.sub(r'[^\d.-]', '', str(amount_str))

    try:
        return float(cleaned) if cleaned else 0
    except ValueError:
        return 0