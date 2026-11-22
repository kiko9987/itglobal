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

from dashboard.blueprints.hr import _hr_admin, _is_field, _me, _my_person, dev_preview_gate, gw_page
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import vehicle_driver as vd
from dashboard.services import vehicle_log as vl
from dashboard.services import vehicle_manage as vm
from dashboard.services import visit_plan as vp
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

gw_vehicles_bp = Blueprint('gw_vehicles', __name__)
gw_vehicles_bp.before_request(dev_preview_gate)   # 개발 중 — 부장님만(11-21)
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
    leads: list = []
    try:        # 일정확정 결과(리드 시트 방문 예정일·영업 담당자) — 못 읽어도 배정은 할 수 있게
        leads = _leads()
        visitors, teams = va.visitors(day, leads), va.teams(day, leads)
        try:    # 그룹웨어에서 확정한 날은 그 팀 그대로(시트 캐시가 늦게 바뀌어도, 11-22)
            teams = vp.confirmed_teams(day) or teams
        except Exception as e:
            logger.warning(f'[차량배정] 방문 일정 팀 읽기 실패: {e}')
        # 갈 곳을 동까지 보이게 — 아직 못 찾은 방문 주소는 뒤에서 찾고(행안부), 화면이 잠시 뒤 다시 받아 간다
        resolving = bool(vl.unresolved(leads, day, day))
        if resolving and not _resolving.locked():
            _resolve_later(day, day)
    except Exception as e:
        logger.warning(f'[차량배정] 방문 일정 읽기 실패: {e}')
        visitors, teams, resolving = [], [], False
    try:        # 그날 방문 일정 칸(캔버스1과 같은 묶음, 시트에서 바로)
        schedule = va.day_schedule(day, leads)
    except Exception as e:
        logger.warning(f'[차량배정] 방문 일정 칸 실패: {e}', exc_info=True)
        schedule = []
    try:        # 그날 쉬는 사람(그룹웨어 휴가·토요일 휴무표)
        off = va.day_off(day)
    except Exception as e:
        logger.warning(f'[차량배정] 휴가 읽기 실패: {e}')
        off = {'off': [], 'half': []}
    try:        # 그날 온라인 문의 당번(/일정확정 때 캔버스2 '온라인' 섹션에서 저장)
        online = va.online_duty(day)
    except Exception as e:
        logger.warning(f'[차량배정] 온라인 당번 읽기 실패: {e}')
        online = []
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
                    'resolving': resolving, 'schedule': schedule, 'off': off['off'], 'half': off['half'], 'online': online,
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


# ── ① 방문 일정(11-22 부장님) — 슬랙 캔버스2 + /일정확정 대신 그룹웨어에서 짜고 확정. 규칙은 services/visit_plan.py
#    보기는 사무실 직원(현장 세션은 문지기가 막음), 고치기·확정은 일정 담당(JW·대표님·부장님)
def _can_plan() -> bool:
    return not _is_field() and vp.can_plan(_me())


def _fresh_leads() -> list:
    from dashboard.services.lead_service import load_leads_data
    df = load_leads_data(force_refresh=True)
    return [] if df is None or df.empty else df.to_dict('records')


@gw_vehicles_bp.route('/api/vehicles/plan', methods=['GET'])
@gw_page
def api_plan_get():
    day = _day()
    try:
        data = vp.load(day, _leads(), _me() if not _is_field() else '')
    except Exception as e:
        logger.error(f'[방문일정] 불러오기 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'}), 500
    data['can_edit'] = bool(data.pop('can_plan')) and data['editable'] and not _is_field()
    return jsonify({'success': True, **data})


def _plan_body():
    body = request.get_json(silent=True) or {}
    return (body.get('day') or '').strip(), body.get('plan') or {}


@gw_vehicles_bp.route('/api/vehicles/plan', methods=['POST'])
@gw_page
def api_plan_save():
    if not _can_plan():
        return jsonify({'success': False, 'message': '방문 일정은 일정 담당자만 고칠 수 있어요.'}), 403
    day, plan = _plan_body()
    try:
        vp.save(day, plan, _me(), _leads())
        data = vp.load(day, _leads(), _me())
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error(f'[방문일정] 저장 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '저장하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    data['can_edit'] = bool(data.pop('can_plan')) and data['editable']
    return jsonify({'success': True, **data})


@gw_vehicles_bp.route('/api/vehicles/plan/confirm', methods=['POST'])
@gw_page
def api_plan_confirm():
    if not _can_plan():
        return jsonify({'success': False, 'message': '방문 일정은 일정 담당자만 확정할 수 있어요.'}), 403
    day, plan = _plan_body()
    try:
        st = vp.start_confirm(day, plan, _me(), _leads(), _fresh_leads)
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error(f'[방문일정] 확정 시작 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '확정하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    logger.info(f'[방문일정] {day} 확정 시작 by {_me()}')
    return jsonify({'success': True, 'run': st})


@gw_vehicles_bp.route('/api/vehicles/plan/state', methods=['GET'])
@gw_page
def api_plan_state():
    try:
        return jsonify({'success': True, 'run': vp.run_state(_day())})
    except Exception as e:
        logger.error(f'[방문일정] 상태 읽기 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '상태를 읽지 못했어요.'}), 500


# ── [운행 종료](11-22 부장님) — 운행을 마친 사람이 바로 운전자를 적는다(저녁 DM 과 같은 자리에 저장) ──
def _my_code() -> str:
    p = _my_person() or {}
    return vd.code_of_name(p.get('name') or '')


@gw_vehicles_bp.route('/api/vehicles/trip', methods=['GET'])
@gw_page
def api_trip_mine():
    day = date.today().isoformat()
    try:
        code = _my_code()
        trips = vd.my_trips(code, day) if code else []
    except Exception as e:
        logger.error(f'[운전자] 내 차 읽기 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'}), 500
    return jsonify({'success': True, 'day': day, 'day_label': va.day_label(day), 'trips': trips})


@gw_vehicles_bp.route('/api/vehicles/trip/end', methods=['POST'])
@gw_page
def api_trip_end():
    body = request.get_json(silent=True) or {}
    code = _my_code()
    if not code:
        return jsonify({'success': False, 'message': '차량 배정 명단에 없는 분이에요.'}), 403
    day = (body.get('day') or '').strip()
    plate = (body.get('plate') or '').strip()[:20]
    try:
        vd.end_trip(day, plate, code, body.get('am') or '', body.get('pm') or '')
        trips = vd.my_trips(code, day)
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error(f'[운전자] 운행 종료 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '저장하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    return jsonify({'success': True, 'trips': trips})


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


# ── 차량 관리(11-22) — 보험·검사 만기, 정비·비용(결재에서 자동), 계기판 km. 경영지원만 ─────────────
@gw_vehicles_bp.route('/vehicles/manage')
@gw_page
def page_manage():
    if not _hr_admin():
        return redirect('/home')
    return render_template('gw/vehicle_manage.html')


@gw_vehicles_bp.route('/api/vehicles/manage', methods=['GET'])
@gw_page
def api_manage():
    if not _hr_admin():
        return _deny()
    try:
        vm.sync_costs()              # 결재에서 새로 생긴 차량 비용(10분에 한 번)
    except Exception as e:
        logger.warning(f'[차량관리] 결재 비용 가져오기 실패: {e}', exc_info=True)
    try:
        return jsonify({'success': True, **vm.overview(), 'kinds': list(vm.KINDS)})
    except Exception as e:
        logger.error(f'[차량관리] 불러오기 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'}), 500


@gw_vehicles_bp.route('/api/vehicles/manage/car', methods=['GET'])
@gw_page
def api_manage_car():
    if not _hr_admin():
        return _deny()
    plate = (request.args.get('plate') or '').strip()[:20]
    if not plate:
        return jsonify({'success': False, 'message': '차량을 골라 주세요.'}), 400
    return jsonify({'success': True, 'plate': plate, 'info': vm.info(plate), 'costs': vm.costs(plate=plate),
                    'odometer': vm.odometer(plate)})


@gw_vehicles_bp.route('/api/vehicles/manage/info/<plate>', methods=['POST'])
@gw_page
def api_manage_info(plate):
    if not _hr_admin():
        return _deny()
    try:
        info = vm.update_info(plate, request.get_json(silent=True) or {}, _me())
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.info(f'[차량관리] 차량 정보 {plate} by {_me()}')
    return jsonify({'success': True, 'info': info})


@gw_vehicles_bp.route('/api/vehicles/manage/costs', methods=['POST'])
@gw_page
def api_manage_cost_add():
    if not _hr_admin():
        return _deny()
    b = request.get_json(silent=True) or {}
    try:
        res = vm.add_cost(b.get('plate') or '', b.get('day') or '', b.get('kind') or '', b.get('item') or '',
                          b.get('amount'), b.get('vendor') or '', _me())
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.info(f"[차량관리] 비용 직접 추가 {b.get('plate')} by {_me()}")
    return jsonify({'success': True, **res})


@gw_vehicles_bp.route('/api/vehicles/manage/costs/<int:cost_id>/delete', methods=['POST'])
@gw_page
def api_manage_cost_delete(cost_id):
    if not _hr_admin():
        return _deny()
    try:
        vm.delete_cost(cost_id)
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.info(f'[차량관리] 비용 지움 #{cost_id} by {_me()}')
    return jsonify({'success': True})


@gw_vehicles_bp.route('/api/vehicles/manage/odometer', methods=['POST'])
@gw_page
def api_manage_odometer():
    if not _hr_admin():
        return _deny()
    b = request.get_json(silent=True) or {}
    try:
        res = vm.add_odometer(b.get('plate') or '', b.get('day') or '', b.get('km'), _me(), b.get('note') or '')
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.info(f"[차량관리] 계기판 km {b.get('plate')} by {_me()}")
    return jsonify({'success': True, **res})
