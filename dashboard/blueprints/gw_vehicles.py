"""그룹웨어 「차량 운행」(2026-11-20) — U+커넥트 운행일지를 올려 차량별 운행과 추정 운전자를 본다.
규칙은 services/vehicle_log.py. 위치 기록이라 인사 관리자(경영지원)만 — 현장 세션은 groupware 문지기가 막는다.
상태 바꾸는 API 는 /api/ 로 시작해야 보안 미들웨어를 통과한다.
"""
import re
import threading
from datetime import date

from flask import Blueprint, jsonify, redirect, render_template, request

from dashboard.blueprints.hr import _hr_admin, _me, gw_page
from dashboard.services import vehicle_log as vl
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

gw_vehicles_bp = Blueprint('gw_vehicles', __name__)
_MONTH_RE = re.compile(r'^\d{4}-(0[1-9]|1[0-2])$')
_resolving = threading.Lock()


def _deny():
    return jsonify({'success': False, 'message': '경영지원만 볼 수 있어요.'}), 403


def _month() -> str:
    m = (request.args.get('month') or '').strip()
    return m if _MONTH_RE.match(m) else date.today().strftime('%Y-%m')


def _leads() -> list:
    from dashboard.services.lead_service import load_leads_data
    df = load_leads_data()
    return [] if df is None else df.to_dict('records')


def _resolve_later(start: str, end: str) -> None:
    """방문 주소의 동 찾기(행안부) — 1분쯤 걸려 올리기 응답과 떼어 뒤에서."""
    def run():
        if not _resolving.acquire(blocking=False):
            return
        try:
            n = vl.resolve_dongs(vl.unresolved(_leads(), start, end))
            logger.info(f'[차량운행] 방문 주소 동 찾기 {n}건 ({start}~{end})')
        except Exception as e:
            logger.error(f'[차량운행] 동 찾기 실패: {e}', exc_info=True)
        finally:
            _resolving.release()
    threading.Thread(target=run, name='vehicle-dong', daemon=True).start()


@gw_vehicles_bp.route('/vehicles')
@gw_page
def page_vehicles():
    if not _hr_admin():
        return redirect('/home')
    return render_template('gw/vehicles.html')


@gw_vehicles_bp.route('/api/vehicles', methods=['GET'])
@gw_page
def api_summary():
    if not _hr_admin():
        return _deny()
    try:
        data = vl.month_summary(_month(), _leads())
    except Exception as e:
        logger.error(f'[차량운행] 요약 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'}), 500
    return jsonify({'success': True, **data, 'months': vl.months(), 'last_upload': vl.last_upload(),
                    'resolving': _resolving.locked()})


@gw_vehicles_bp.route('/api/vehicles/car', methods=['GET'])
@gw_page
def api_car():
    if not _hr_admin():
        return _deny()
    plate = (request.args.get('plate') or '').strip()[:20]
    if not plate:
        return jsonify({'success': False, 'message': '차량을 골라 주세요.'}), 400
    try:
        data = vl.vehicle_days(_month(), plate, _leads())
    except Exception as e:
        logger.error(f'[차량운행] 차량 상세 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'}), 500
    return jsonify({'success': True, **data})


@gw_vehicles_bp.route('/api/vehicles/upload', methods=['POST'])
@gw_page
def api_upload():
    if not _hr_admin():
        return _deny()
    f = request.files.get('file')
    if not f:
        return jsonify({'success': False, 'message': '파일을 골라 주세요.'}), 400
    try:
        res = vl.import_upload(f.filename or '', f.read(), _me())
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error(f'[차량운행] 올리기 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '올리지 못했어요. 파일을 확인하고 다시 해 주세요.'}), 500
    logger.info(f'[차량운행] 운행일지 올림 by {_me()}: 차량 {res["vehicles"]}대, 새 운행 {res["added"]}건')
    rng = vl.trip_range()
    if rng:
        _resolve_later(*rng)
    return jsonify({'success': True, **res})
