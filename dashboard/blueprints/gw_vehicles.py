"""그룹웨어 「차량 운행」·「차량 배정」(2026-11-20) — U+커넥트 운행일지를 올려 차량별 운행과 운전자를 보고,
다음 날 차마다 탈 사람을 정한다(배정 규칙은 services/vehicle_assign.py).
운행 규칙은 services/vehicle_log.py. 운행 기록은 위치라 인사 관리자(경영지원)만 — 현장 세션은 groupware 문지기가 막는다.
방문 일정·차량 배정은 시공·미화를 뺀 직원 모두 보고, 고치기·확정은 관리자만(11-25 부장님).
상태 바꾸는 API 는 /api/ 로 시작해야 보안 미들웨어를 통과한다.
"""
import re
import threading
from datetime import date, timedelta

from flask import Blueprint, jsonify, redirect, render_template, request, session

from dashboard.blueprints.hr import _hr_admin, _is_admin, _is_field, _me, _my_person, dev_preview_gate, gw_page
from dashboard.services import hr
from dashboard.services import vehicle_assign as va
from dashboard.services import vehicle_driver as vd
from dashboard.services import vehicle_log as vl
from dashboard.services import vehicle_manage as vm
from dashboard.services import visit_plan as vp
from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

gw_vehicles_bp = Blueprint('gw_vehicles', __name__)
# 방문 일정·차량 배정(+ 배정 화면의 차량 설정) — 시공·미화를 뺀 모두에게 연다(11-25 부장님). 고치기는 각 API 가 관리자만.
_SCHED_PATHS = ('/visits/plan', '/api/vehicles/plan', '/vehicles/assign', '/api/vehicles/assign', '/api/vehicles/fleet/')


def _can_view() -> bool:
    """방문 일정·차량 배정을 볼 수 있나 — 현장 세션이 아니고 시공·미화 부서가 아닌 재직자(관리자는 늘)."""
    return not _is_field() and (_is_admin() or vp.can_view(_me()))


_NOW_PATHS = ('/vehicles/now', '/api/vehicles/now')     # 실시간 위치 — 전 직원(11-24 부장님: 온라인 당번도 봐야 함)


def _gate():
    """개발 중 — 부장님만(11-21). 방문 일정·차량 배정은 시공·미화 뺀 모두(11-25),
    실시간 위치는 전 직원(현장 포함, 11-24)."""
    if request.path.startswith(_SCHED_PATHS):
        if not session.get('user') or _can_view():
            return None         # 로그인 확인은 각 화면(gw_page)이 한다
        if request.path.startswith('/api/'):
            return jsonify({'success': False, 'message': '볼 수 없는 화면이에요.'}), 403
        return redirect('/home')
    if request.path.startswith(_NOW_PATHS):
        return None
    if _hr_open(request.path):
        return None
    return dev_preview_gate()


# 운행일지·차량 관리 — 경영지원(인사 관리자)에게 연다(11-24 부장님 "메뉴는 다 오픈"). 화면·API 안에서 _hr_admin 을 다시 본다.
_HR_EXACT = ('/vehicles', '/api/vehicles', '/api/vehicles/car', '/api/vehicles/upload', '/api/vehicles/logbook.xlsx',
             '/api/vehicles/insights', '/api/vehicles/route')
_HR_PREFIX = ('/vehicles/manage', '/api/vehicles/manage')


def _hr_open(path: str) -> bool:
    return path in _HR_EXACT or path.startswith(_HR_PREFIX)


def _kakao_js() -> str:
    """카카오 지도 JavaScript 키(카카오 개발자 › 내 애플리케이션 › 앱 키). 없으면 지도 대신 목록만."""
    import os
    return (os.getenv('KAKAO_JS_KEY') or '').strip()


gw_vehicles_bp.before_request(_gate)
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


class _LeadsUnavailable(Exception):
    """리드 시트를 못 읽음 — 방문 일정은 빈 리드로 열거나 저장하면 안 된다."""


_LEADS_DOWN_MSG = '리드 시트를 읽지 못했어요. 잠시 뒤 다시 해 주세요.'


def _plan_leads() -> list:
    """방문 일정(불러오기·저장·확정·당일 배정)용 리드. 못 읽으면 멈춘다(11-24 조용한 실패 점검 1번:
    빈 리드로 저장하면 짜 둔 코스의 현장이 모두 빠진 채 덮어써졌음)."""
    from dashboard.services.lead_service import load_leads_data
    df = load_leads_data()
    if df is None or df.empty:
        logger.error('[방문일정] 리드 시트를 읽지 못해 방문 일정 작업을 멈춤')
        raise _LeadsUnavailable()
    return df.to_dict('records')


_LEAD_DAY = re.compile(r'^(\d{4})[-./](\d{1,2})[-./](\d{1,2})')


def _day_leads(leads: list, day: str) -> list:
    """그날에 걸친 방문 리드만(11-24 부장님: 날짜를 바꿀 때 오래 걸림 — 아래 함수마다 리드 5천 건 날짜를 다시 읽던 것).
    대부분(지난 날·다른 날 한 날짜)은 앞 날짜만 보고 바로 거르고, 기간('~')이거나 모양이 다른 것만 두 날짜 읽기
    (일정 목록·방문 일정 = 시작~끝, 운행일지 = vl._visit_days)로 본다 — 어느 쪽에도 빠지지 않게."""
    from dashboard.services import visit_assignment_sync as vas
    d = date.fromisoformat(day)
    out = []
    for lead in leads:
        vd = lead.get('방문 예정일')
        if not vd:
            continue
        s = str(vd).strip().lstrip("'")
        m = _LEAD_DAY.match(s)
        if m and '~' not in s:
            try:
                if date(int(m.group(1)), int(m.group(2)), int(m.group(3))) == d:
                    out.append(lead)
            except ValueError:
                pass
            continue
        st = vas._parse_visit_date_start(vd)
        if st and st > d:
            continue
        en = vas._parse_visit_date_end(vd) if st else None
        if (st and st <= d <= (en if en and en >= st else st)) or day in vl._visit_days(vd):
            out.append(lead)
    return out


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
    missing: list = []
    try:        # 확정한 날인데 운전자 답이 없는 차(운행 종료 감지·아침 다시 묻기 뒤에도 남은 것, 11-24)
        from dashboard.services import vehicle_trip_watch as tw
        m = data.get('month') or _month()
        y, mo = int(m[:4]), int(m[5:7])
        last = (date(y + (mo == 12), mo % 12 + 1, 1) - timedelta(days=1)).isoformat()
        missing = tw.missing(f'{m}-01', last)
    except Exception as e:
        logger.warning(f'[차량운행] 운전자 미확인 읽기 실패: {e}')
    return jsonify({'success': True, **data, 'months': vl.months(), 'last_upload': vl.last_upload(),
                    'resolving': _resolving.locked(), 'missing_drivers': missing})


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


# ── 차량 배정(11-20) — 보기는 시공·미화 뺀 모두, 고치기는 관리자만(11-25) ──────────
def _can_assign() -> bool:
    return not _is_field() and va.can_assign(_me())


def _day() -> str:
    d = (request.args.get('day') or '').strip()
    return d if va.valid_day(d) else va.next_workday()


def _vtabs() -> dict:
    """방문 일정 ↔ 차량 배정 탭(11-24 부장님). 두 화면은 같은 사람에게 열려 있어(11-25) 늘 둘 다."""
    ok = _can_view()
    return {'plan': ok, 'cars': ok}


@gw_vehicles_bp.route('/vehicles/assign')
@gw_page
def page_assign():
    return render_template('gw/vehicle_assign.html', vtabs=_vtabs(), vtab_on='cars')


@gw_vehicles_bp.route('/api/vehicles/assign', methods=['GET'])
@gw_page
def api_assign_get():
    """배정 화면 한 날짜. 곁들이는 정보(리드 시트·방문 일정 칸·휴가·온라인 당번·방문 일정 확정)는 못 읽어도
    배정은 할 수 있게 열되, 못 읽은 칸 이름을 partial 로 싣는다 — 빈 목록으로 바꿔 success 만 보내면 화면이
    '이날 방문 일정이 없어요'·휴무 표시 없음으로 보여 휴가자를 차에 넣고 확정할 수 있었다(11-24 조용한 실패 점검 18번).
    partial: 'leads'(리드 시트) · 'visits'(방문 사람·코스) · 'teams'(확정한 방문 일정 팀) · 'schedule'(방문 일정 칸)
             · 'off'(휴가·토요 휴무) · 'online'(온라인 당번) · 'plan'(방문 일정 확정 정보)"""
    day = _day()
    partial: list = []
    leads: list = []
    visitors, teams, resolving = [], [], False
    # 일정확정 결과(리드 시트 방문 예정일·영업 담당자). 리드 시트는 수천 줄이라 빈 결과 = 못 읽음
    #   (load_leads_data 가 시트 오류·빈 시트를 None 으로 돌려주고 _leads 가 그걸 [] 로 바꾼다)
    try:
        all_leads = _leads()
    except Exception as e:
        logger.error(f'[차량배정] {day} 리드 시트 읽기 실패 — 방문·코스 표시 없이 엶: {e}', exc_info=True)
        all_leads = None
    else:
        if not all_leads:
            logger.error(f'[차량배정] {day} 리드 시트를 읽지 못함(빈 결과) — 방문·코스 표시 없이 엶')
    if not all_leads:
        partial.append('leads')
    else:
        try:
            leads = _day_leads(all_leads, day)      # 그날 리드만 한 번 골라 아래 모두에 넘긴다(모두 그날만 봄)
            visitors, teams = va.visitors(day, leads), va.teams(day, leads)
        except Exception as e:
            logger.error(f'[차량배정] {day} 방문 일정 읽기 실패: {e}', exc_info=True)
            visitors, teams = [], []
            partial.append('visits')
        try:    # 그룹웨어에서 확정한 날은 그 팀 그대로(시트 캐시가 늦게 바뀌어도, 11-22)
            teams = vp.confirmed_teams(day, leads) or teams     # 슬랙으로 정한 오늘·지난 날은 시트 코스(11-23)
        except Exception as e:
            logger.error(f'[차량배정] {day} 방문 일정 팀 읽기 실패: {e}', exc_info=True)
            partial.append('teams')
        try:    # 갈 곳을 동까지 보이게 — 아직 못 찾은 방문 주소는 뒤에서 찾고(행안부), 화면이 잠시 뒤 다시 받아 간다
            resolving = bool(vl.unresolved(leads, day, day))
            if resolving and not _resolving.locked():
                _resolve_later(day, day)
        except Exception as e:      # 동 찾기는 곁가지(못 찾으면 시·구까지만 보임)
            logger.warning(f'[차량배정] {day} 방문 주소 동 찾기 준비 실패: {e}')
            resolving = False
    try:        # 그날 방문 일정 칸(캔버스1과 같은 묶음, 시트에서 바로)
        schedule = va.day_schedule(day, leads)
    except Exception as e:
        logger.error(f'[차량배정] {day} 방문 일정 칸 실패: {e}', exc_info=True)
        schedule = []
        partial.append('schedule')
    try:        # 그날 쉬는 사람(그룹웨어 휴가·토요일 휴무표) — 못 읽으면 휴가자를 차에 넣고 확정할 수 있어 ERROR
        off = va.day_off(day)
    except Exception as e:
        logger.error(f'[차량배정] {day} 휴가 읽기 실패 — 휴무 표시 없이 엶: {e}', exc_info=True)
        off = {'off': [], 'half': []}
        partial.append('off')
    try:        # 그날 온라인 문의 당번(/일정확정 때 캔버스2 '온라인' 섹션에서 저장)
        online = va.online_duty(day)
    except Exception as e:
        logger.error(f'[차량배정] {day} 온라인 당번 읽기 실패: {e}', exc_info=True)
        online = []
        partial.append('online')
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
    try:        # 그룹웨어 방문 일정 확정 한 줄(11-22) — 화면 위에 '○ 방문 일정 확정 · 코스 n개'
        info = vp.confirmed_info(day)
        view['plan'] = {**info, 'can_open': _can_view(), 'cars_saved_at': va.last_saved_at(day)} if info else None
        view['work'] = vp.confirmed_work(day)      # 사내 업무(차 필요 없음, 11-23)
        view['courses'] = vp.confirmed_courses(day, leads)     # 위 칸을 코스로(11-23) — 확정한 날 + 슬랙으로 정한 오늘·지난 날
    except Exception as e:
        logger.error(f'[차량배정] {day} 방문 일정 확정 정보 실패: {e}', exc_info=True)
        view['plan'] = None
        view['work'] = []
        view['courses'] = None
        partial.append('plan')
    return jsonify({'success': True, **view, 'day_label': va.day_label(day), 'partial': partial,
                    'staff': [{'code': p['code'], 'name': p['name']} for p in people], 'visitors': visitors, 'teams': teams,
                    'resolving': resolving, 'schedule': schedule, 'off': off['off'], 'half': off['half'], 'online': online,
                    'can_edit': _can_assign(), 'can_see_log': _hr_admin(), 'dm_off': not va.CAR_DM_ON})


@gw_vehicles_bp.route('/api/vehicles/assign', methods=['POST'])
@gw_page
def api_assign_save():
    if not _can_assign():
        return jsonify({'success': False, 'message': '차량 배정은 관리자만 수정할 수 있어요.'}), 403
    body = request.get_json(silent=True) or {}
    day = (body.get('day') or '').strip()
    # base = 화면이 불러올 때 받은 그날 판. 안 보낸 옛 화면(재시작 직후 열려 있던 것)은 예전처럼 저장(11-24 조용한 실패 점검 17번)
    base = body.get('base')
    try:
        res = va.save_day(day, body.get('rows') or [], _me(), confirm=bool(body.get('confirm')),
                          notify=body.get('notify', True) is not False, base=None if base is None else str(base))
    except va.StaleSave as e:
        return jsonify({'success': False, 'conflict': True, 'message': str(e)}), 409
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error(f'[차량배정] 저장 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '저장하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    logger.info(f"[차량배정] {day} 저장 by {_me()}: 바뀐 차 {len(res['changed'])}대, DM {res['notified']}건"
                f"{' (확정)' if body.get('confirm') else ''}")
    return jsonify({'success': True, **res})


@gw_vehicles_bp.route('/api/vehicles/assign/resend', methods=['POST'])
@gw_page
def api_assign_resend():
    """못 받은 사람에게 배정 DM 다시 보내기(11-24 조용한 실패 점검 추가분)."""
    if not _can_assign():
        return jsonify({'success': False, 'message': '차량 배정은 관리자만 수정할 수 있어요.'}), 403
    body = request.get_json(silent=True) or {}
    day = (body.get('day') or '').strip()
    codes = [str(c)[:20] for c in (body.get('codes') or []) if c][:50]
    try:
        res = va.resend(day, codes)
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error(f'[차량배정] 다시 보내기 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '보내지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    logger.info(f"[차량배정] {day} DM 다시 보냄 by {_me()}: {res['notified']}명")
    return jsonify({'success': True, **res})


# ── 방문 일정(11-22 부장님, 영업 > 방문 일정 /visits/plan) — 슬랙 캔버스2 + /일정확정 대신 그룹웨어에서 짜고 확정.
#    규칙은 services/visit_plan.py. 보기는 시공·미화 뺀 모두, 고치기·확정·당일 배정은 관리자만(11-25), 현장 세션은 문지기가 막음
def _can_plan() -> bool:
    return not _is_field() and vp.can_plan(_me())


@gw_vehicles_bp.route('/visits/plan')
@gw_page
def page_visit_plan():
    return render_template('gw/visit_plan.html', vtabs=_vtabs(), vtab_on='plan')


def _fresh_leads() -> list:
    from dashboard.services.lead_service import load_leads_data
    df = load_leads_data(force_refresh=True)
    return [] if df is None or df.empty else df.to_dict('records')


@gw_vehicles_bp.route('/api/vehicles/plan', methods=['GET'])
@gw_page
def api_plan_get():
    day = _day()
    try:
        data = vp.load(day, _plan_leads(), _me() if not _is_field() else '')
    except _LeadsUnavailable:
        return jsonify({'success': False, 'message': _LEADS_DOWN_MSG}), 503
    except Exception as e:
        logger.error(f'[방문일정] 불러오기 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'}), 500
    data['can_edit'] = bool(data.pop('can_plan')) and data['editable'] and not _is_field()
    data['to_cars'] = _can_assign()     # 확정이 끝나면 차량 배정 탭으로 바로 넘어갈 사람(고칠 수 있는 사람, 11-24)
    data['can_urgent'] = not _is_field() and vp.can_urgent(_me(), day)
    return jsonify({'success': True, **data})


@gw_vehicles_bp.route('/api/vehicles/plan/urgent', methods=['POST'])
@gw_page
def api_plan_urgent():
    """당일 배정(11-23) — 오늘 생긴 현장을 근처 매니저에게. 관리자만(11-25)."""
    if _is_field():
        return jsonify({'success': False, 'message': '사무실에서만 할 수 있어요.'}), 403
    body = request.get_json(silent=True) or {}
    day = (body.get('day') or '').strip()
    try:
        # 확인은 캐시된 리드로(빠르게) — 시트 기록·DM 은 뒤에서(11-23: 한 번에 11초쯤 걸렸음)
        res = vp.urgent_assign(day, str(body.get('lead_no') or '').strip(), body.get('members') or [], _me(), _plan_leads())
    except _LeadsUnavailable:
        return jsonify({'success': False, 'message': _LEADS_DOWN_MSG}), 503
    except PermissionError as e:
        return jsonify({'success': False, 'message': str(e)}), 403
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error(f'[방문일정] 당일 배정 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '당일 배정하지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500
    return jsonify({'success': True, **res})


def _plan_body():
    body = request.get_json(silent=True) or {}
    return (body.get('day') or '').strip(), body.get('plan') or {}


@gw_vehicles_bp.route('/api/vehicles/plan', methods=['POST'])
@gw_page
def api_plan_save():
    if not _can_plan():
        return jsonify({'success': False, 'message': '방문 일정은 관리자만 수정할 수 있어요.'}), 403
    day, plan = _plan_body()
    try:
        leads = _plan_leads()
        vp.save(day, plan, _me(), leads)
        data = vp.load(day, leads, _me())
    except _LeadsUnavailable:
        return jsonify({'success': False, 'message': _LEADS_DOWN_MSG}), 503
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
        return jsonify({'success': False, 'message': '방문 일정은 관리자만 확정할 수 있어요.'}), 403
    day, plan = _plan_body()
    try:
        st = vp.start_confirm(day, plan, _me(), _plan_leads(), _fresh_leads)
    except _LeadsUnavailable:
        return jsonify({'success': False, 'message': _LEADS_DOWN_MSG}), 503
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
        return jsonify({'success': False, 'message': '차량 설정은 관리자만 수정할 수 있어요.'}), 403
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


# ── U+커넥트 기능 옮기기(11-24) — 운행기록부·운전 습관·가동률·경로(경영지원), 실시간 위치(전 직원) ──────────
def _month_span() -> tuple:
    m = _month()
    y, mo = int(m[:4]), int(m[5:7])
    last = date(y + (mo == 12), mo % 12 + 1, 1) - timedelta(days=1)
    return m, f'{m}-01', last.isoformat()


@gw_vehicles_bp.route('/api/vehicles/logbook.xlsx', methods=['GET'])
@gw_page
def api_logbook():
    """국세청 업무용승용차 운행기록부 — 차량마다 한 장(plate 를 주면 그 차만)."""
    if not _hr_admin():
        return _deny()
    from flask import Response
    from urllib.parse import quote
    m, start, end = _month_span()
    plate = (request.args.get('plate') or '').strip()[:20]
    with vl._conn() as conn:
        plates = [r[0] for r in conn.execute('SELECT DISTINCT plate FROM vehicle_trips WHERE day BETWEEN ? AND ? ORDER BY plate',
                                            (start, end)).fetchall()]
    if plate:
        plates = [plate]
    if not plates:
        return jsonify({'success': False, 'message': f'{m} 운행 기록이 없어요.'}), 404
    data = vl.logbook_xlsx(plates, start, end, corp='아이티글로벌')
    name = f"운행기록부_{m}{'_' + plate if plate else ''}.xlsx"
    logger.info(f'[차량운행] 운행기록부 내려받기 {name} by {_me()}')
    return Response(data, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f"attachment; filename*=UTF-8''{quote(name)}"})


@gw_vehicles_bp.route('/api/vehicles/insights', methods=['GET'])
@gw_page
def api_insights():
    """운전 습관(사람별·차량별)·가동률 — 이 달."""
    if not _hr_admin():
        return _deny()
    from dashboard.services import uplus_api as up
    from dashboard.services import vehicle_api_sync as vas
    m, start, end = _month_span()
    try:
        return jsonify({'success': True, 'month': m, 'habits': vas.habits(start, end), 'util': vas.utilization(start, end),
                        'api': up.enabled('eco')})
    except Exception as e:
        logger.error(f'[차량운행] 운전 습관·가동률 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'}), 500


@gw_vehicles_bp.route('/api/vehicles/route', methods=['GET'])
@gw_page
def api_route():
    """운행 한 건의 경로 좌표(카카오 지도 선)."""
    if not _hr_admin():
        return _deny()
    from dashboard.services import vehicle_api_sync as vas
    plate = (request.args.get('plate') or '').strip()[:20]
    day = (request.args.get('day') or '').strip()
    if not plate or not va.valid_day(day):
        return jsonify({'success': False, 'message': '차량과 날짜를 골라 주세요.'}), 400
    hm = re.compile(r'^\d{2}:\d{2}$')
    start = request.args.get('start') or ''
    end = request.args.get('end') or ''
    try:
        pts = vas.route(plate, day, start if hm.match(start) else '', end if hm.match(end) else '')
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.warning(f'[차량운행] 경로 실패 {plate} {day}: {e}')
        return jsonify({'success': False, 'message': 'U+ 에서 경로를 받지 못했어요. 잠시 뒤 다시 해 주세요.'}), 502
    return jsonify({'success': True, 'points': pts, 'kakao_js': _kakao_js()})


@gw_vehicles_bp.route('/vehicles/now')
@gw_page
def page_now():
    return render_template('gw/vehicle_now.html', kakao_js=_kakao_js())


@gw_vehicles_bp.route('/api/vehicles/now', methods=['GET'])
@gw_page
def api_now():
    from dashboard.services import vehicle_api_sync as vas
    try:
        return jsonify({'success': True, **vas.now_view(), 'kakao_js': _kakao_js()})
    except Exception as e:
        logger.error(f'[차량위치] 불러오기 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '불러오지 못했어요. 잠시 뒤 새로 고쳐 주세요.'}), 500


@gw_vehicles_bp.route('/api/vehicles/now/near', methods=['GET'])
@gw_page
def api_now_near():
    from dashboard.services import vehicle_api_sync as vas
    addr = (request.args.get('addr') or '').strip()[:120]
    if len(addr) < 2:
        return jsonify({'success': False, 'message': '주소를 적어 주세요.'}), 400
    try:
        return jsonify({'success': True, **vas.nearest(addr)})
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    except Exception as e:
        logger.error(f'[차량위치] 가까운 차 찾기 실패: {e}', exc_info=True)
        return jsonify({'success': False, 'message': '찾지 못했어요. 잠시 뒤 다시 해 주세요.'}), 500


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
                    'odometer': vm.odometer(plate), 'accidents': vm.accidents(plate)})


@gw_vehicles_bp.route('/api/vehicles/manage/accidents', methods=['POST'])
@gw_page
def api_manage_accident_save():
    """사고 기록 넣기(id 가 있으면 고치기)."""
    if not _hr_admin():
        return _deny()
    b = request.get_json(silent=True) or {}
    try:
        acc_id = int(b.get('id') or 0) or None
        res = vm.save_accident(b, _me(), acc_id)
    except (TypeError, ValueError) as e:
        return jsonify({'success': False, 'message': str(e) if str(e) else '값을 확인해 주세요.'}), 400
    logger.info(f"[차량관리] 사고 기록 {'고침' if acc_id else '넣음'} {b.get('plate')} by {_me()}")
    return jsonify({'success': True, **res})


@gw_vehicles_bp.route('/api/vehicles/manage/accidents/<int:acc_id>/delete', methods=['POST'])
@gw_page
def api_manage_accident_delete(acc_id):
    if not _hr_admin():
        return _deny()
    try:
        vm.delete_accident(acc_id)
    except ValueError as e:
        return jsonify({'success': False, 'message': str(e)}), 400
    logger.info(f'[차량관리] 사고 기록 지움 #{acc_id} by {_me()}')
    return jsonify({'success': True})


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
