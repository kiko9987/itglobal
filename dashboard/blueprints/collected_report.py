# -*- coding: utf-8 -*-
"""월별 수금 완료 정산 화면·엑셀 (에어테이블 '수금 완료 💲 - YY년M월' 보기 대체, 2026-11-08).

- GET  /settlement/collected                    화면 (정산 월 선택·담당자→사업자 묶음 합계)
- GET  /api/settlement/collected?period=YYYY-MM JSON
- GET  /api/settlement/collected.xlsx?period=…  엑셀 내려받기
- POST /api/settlement/collected/range          그 달 정산 기간 바꾸기·되돌리기 (2026-11-09)
데이터는 PM 공사 목록과 같은 캐시(공사현황 시트). refresh=true 면 시트에서 새로 읽음.
달별로 바꾼 기간은 Redis 해시 settlement:collected:ranges 에 {'YYYY-MM': 'YYYY-MM-DD|YYYY-MM-DD'}.
"""
import io
from datetime import date

from flask import Blueprint, jsonify, redirect, render_template, request, send_file, session

from ..auth import login_required
from ..services import collected_report as cr
from ..utils.logging_config import get_logger
from ..utils.error_helpers import generate_error_id

logger = get_logger(__name__)

collected_report_bp = Blueprint('collected_report', __name__)

RANGES_KEY = 'settlement:collected:ranges'


def _records(refresh: bool) -> list:
    from .projects import _load_projects_data
    df = _load_projects_data(refresh)
    if df is None or len(df) == 0:
        return []
    return df.fillna('').to_dict('records')


def _redis():
    from ..utils.redis_client import get_redis_client
    return get_redis_client().redis


def _saved_ranges() -> dict:
    """바꿔 둔 기간. Redis 가 안 되면 빈 값 → 26~25일 규칙으로 보여줌."""
    try:
        return _redis().hgetall(RANGES_KEY) or {}
    except Exception as e:
        logger.warning(f'정산 기간 설정 읽기 실패(기본 26~25일로 표시): {e}')
        return {}


RANGE_EDIT_DENIED = '정산 기간은 경영지원과 대표님만 바꿀 수 있어요.'


def _can_edit_range() -> bool:
    """정산 기간 저장·되돌리기는 정산 승인 명단(경영지원 SB·KiKO + 대표 JW·YG, settlement_checkers)만 — 부장님 결정 11-17."""
    from ..services.settlement_checkers import checker_emails
    u = session.get('user') or {}
    return isinstance(u, dict) and (u.get('email') or '').strip().lower() in checker_emails()


def _report():
    year, month = cr.parse_period(request.args.get('period', ''))
    refresh = request.args.get('refresh', 'false').lower() == 'true'
    return cr.collect(_records(refresh), year, month, _saved_ranges())


@collected_report_bp.before_request
def _admin_only():
    """재무(월별 수금)는 관리자만(admin·super_admin) — 11-25 부장님 '저 카테고리는 어드민한테만'.
    예전엔 로그인한 누구나 열 수 있었다(담당자별 수금 금액이 보임). 로그인 확인은 각 경로의 login_required 가 한다."""
    u = session.get('user') or {}
    if not isinstance(u, dict) or not u or u.get('permission_level') in ('admin', 'super_admin'):
        return None
    if request.path.startswith('/api/'):
        return jsonify({'success': False, 'error': '관리자만 볼 수 있는 화면이에요.'}), 403
    return redirect('/home')


@collected_report_bp.route('/settlement/collected')
@login_required
def collected_page():
    return render_template('collected_report.html', periods=cr.recent_periods(saved=_saved_ranges()),
                           can_edit_range=_can_edit_range())


@collected_report_bp.route('/api/settlement/collected')
@login_required
def collected_json():
    try:
        return jsonify({'success': True, 'data': _report()})
    except Exception as e:
        eid = generate_error_id()
        logger.error(f'[{eid}] 수금 완료 정산 자료 오류: {e}', exc_info=True)
        return jsonify({'success': False, 'error': '자료를 불러오지 못했습니다.', 'error_id': eid}), 500


@collected_report_bp.route('/api/settlement/collected/range', methods=['POST'])
@login_required
def collected_range():
    """그 달 정산 기간 저장. reset=true 면 기본(26~25일)으로 되돌림. 경영지원·대표만."""
    if not _can_edit_range():
        return jsonify({'success': False, 'error': RANGE_EDIT_DENIED}), 403
    body = request.get_json(silent=True) or {}
    period = str(body.get('period') or '')
    try:
        y, m = (int(x) for x in period.split('-'))
        if not 1 <= m <= 12:
            raise ValueError
    except ValueError:
        return jsonify({'success': False, 'error': '정산 월이 올바르지 않습니다.'}), 400
    key = f'{y}-{m:02d}'
    u = session.get('user') or {}
    user = (u.get('name') or u.get('email')) if isinstance(u, dict) else u
    try:
        if body.get('reset'):
            _redis().hdel(RANGES_KEY, key)
            logger.info(f'정산 기간 되돌림: {key} (기본 26~25일) by {user}')
            return jsonify({'success': True})
        try:
            s, e = date.fromisoformat(str(body.get('start'))), date.fromisoformat(str(body.get('end')))
        except ValueError:
            return jsonify({'success': False, 'error': '날짜가 올바르지 않습니다.'}), 400
        err = cr.check_range(s, e)
        if err:
            return jsonify({'success': False, 'error': err}), 400
        ds, de = cr.period_range(y, m)
        if (s, e) == (ds, de):
            _redis().hdel(RANGES_KEY, key)   # 기본과 같으면 저장하지 않음
        else:
            _redis().hset(RANGES_KEY, key, f'{s.isoformat()}|{e.isoformat()}')
        logger.info(f'정산 기간 저장: {key} = {s} ~ {e} by {user}')
        return jsonify({'success': True})
    except Exception as ex:
        eid = generate_error_id()
        logger.error(f'[{eid}] 정산 기간 저장 오류: {ex}', exc_info=True)
        return jsonify({'success': False, 'error': '기간을 저장하지 못했습니다.', 'error_id': eid}), 500


@collected_report_bp.route('/api/settlement/collected.xlsx')
@login_required
def collected_xlsx():
    try:
        rep = _report()
        return send_file(io.BytesIO(cr.to_xlsx(rep)), as_attachment=True,
                         download_name=cr.xlsx_filename(rep),
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        eid = generate_error_id()
        logger.error(f'[{eid}] 수금 완료 엑셀 오류: {e}', exc_info=True)
        return jsonify({'success': False, 'error': '엑셀을 만들지 못했습니다.', 'error_id': eid}), 500
