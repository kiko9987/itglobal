"""그룹웨어 「차량 운행」·「차량 배정」(2026-11-20) — U+커넥트 운행일지를 올려 차량별 운행과 운전자를 보고,
다음 날 차마다 탈 사람을 정한다(배정 규칙은 services/vehicle_assign.py).
운행 규칙은 services/vehicle_log.py. 운행 기록은 위치라 인사 관리자(경영지원)만 — 현장 세션은 groupware 문지기가 막는다.
배정은 직원 모두(현장 포함) 볼 수 있고 고치기는 배정 담당자(설정 vehicle_assigners, 기본 SJ)·인사 관리자.
상태 바꾸는 API 는 /api/ 로 시작해야 보안 미들웨어를 통과한다.
"""
import re
import threading
from datetime import date

from flask import Blueprint, jsonify, redirect, render_template, request

from dashboard.blueprints.hr import _hr_admin, _is_field, _me, gw_page
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
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


# ── 차량 배정(11-20) — 보기는 직원 모두(현장 포함), 고치기는 배정하는 사람(SJ)·인사 관리자 ──────────
def _can_assign() -> bool:
    return not _is_field() and va.can_assign(_me())


def _day() -> str:
    d = (request.args.get('day') or '').strip()
    return d if va.valid_day(d) else va.next_workday()


@gw_vehicles_bp.route('/vehicles/assign')
@gw_page
def page_assign():
    return render_template('gw/vehicle_assign.html')


@gw_vehicles_bp.route('/api/vehicles/assign', methods=['GET'])
@gw_page
def api_assign_get():
    day = _day()
    try:        # 일정확정 결과(리드 시트 방문 예정일·영업 담당자) — 못 읽어도 배정은 할 수 있게
        leads = _leads()
        visitors, teams = va.visitors(day, leads), va.teams(day, leads)
    except Exception as e:
        logger.warning(f'[차량배정] 방문 일정 읽기 실패: {e}')
        visitors, teams = [], []
    try:
        view = va.day_view(day, teams)
        people = va.staff()
    except Exception as e:
        logger.error(f'[차량배정] 불러오기 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'}), 500
    who = {p['code']: p for p in people}
    for c in view['cars']:
        c['fixed_name'] = va.display(c['fixed_user'], who) if c['fixed_user'] else ''
    if view.get('confirmed_by'):
        p = hr.find_by_email(view['confirmed_by'])
        view['confirmed_name'] = (p or {}).get('name') or view['confirmed_by']
    return jsonify({'success': True, **view, 'day_label': va.day_label(day),
                    'staff': [{'code': p['code'], 'name': p['name']} for p in people], 'visitors': visitors, 'teams': teams,
                    'can_edit': _can_assign(), 'can_see_log': _hr_admin()})


@gw_vehicles_bp.route('/api/vehicles/assign', methods=['POST'])
@gw_page
def api_assign_save():
    if not _can_assign():
        return jsonify({'success': False, 'message': '차량 배정은 배정 담당자만 고칠 수 있어요.'}), 403
    body = request.get_json(silent=True) or {}
    day = (body.get('day') or '').strip()
    try:
        res = va.save_day(day, body.get('rows') or [], _me(), confirm=bool(body.get('confirm')),
                          notify=body.get('notify', True) is not False)
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error(f'[차량배정] 저장 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '저장하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    logger.info(f"[차량배정] {day} 저장 by {_me()}: 바뀐 차 {len(res['changed'])}대, DM {res['notified']}건"
                f"{' (확정)' if body.get('confirm') else ''}")
    return jsonify({'success': True, **res})


@gw_vehicles_bp.route('/api/vehicles/fleet/<plate>', methods=['POST'])
@gw_page
def api_fleet_update(plate):
    if not _can_assign():
        return jsonify({'success': False, 'message': '차량 설정은 배정 담당자만 고칠 수 있어요.'}), 403
    try:
        car = va.update_car(plate, request.get_json(silent=True) or {}, _me())
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.info(f'[차량배정] 차량 설정 {plate} by {_me()}')
    return jsonify({'success': True, 'car': car})


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
