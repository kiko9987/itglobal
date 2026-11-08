# -*- coding: utf-8 -*-
"""월별 수금 완료 정산 화면·엑셀 (에어테이블 '수금 완료 💲 - YY년M월' 보기 대체, 2026-11-08).

- GET /settlement/collected                    화면 (정산 월 선택·담당자→사업자 묶음 합계)
- GET /api/settlement/collected?period=YYYY-MM JSON
- GET /api/settlement/collected.xlsx?period=…  엑셀 내려받기
데이터는 PM 공사 목록과 같은 캐시(공사현황 시트). refresh=true 면 시트에서 새로 읽음.
"""
import io

from flask import Blueprint, jsonify, render_template, request, send_file

from ..auth import login_required
from ..services import collected_report as cr
from ..utils.logging_config import get_logger
from ..utils.error_helpers import generate_error_id

logger = get_logger(__name__)

collected_report_bp = Blueprint('collected_report', __name__)


def _records(refresh: bool) -> list:
    from .projects import _load_projects_data
    df = _load_projects_data(refresh)
    if df is None or len(df) == 0:
        return []
    return df.fillna('').to_dict('records')


def _report():
    year, month = cr.parse_period(request.args.get('period', ''))
    refresh = request.args.get('refresh', 'false').lower() == 'true'
    return cr.collect(_records(refresh), year, month)


@collected_report_bp.route('/settlement/collected')
@login_required
def collected_page():
    return render_template('collected_report.html', periods=cr.recent_periods())


@collected_report_bp.route('/api/settlement/collected')
@login_required
def collected_json():
    try:
        return jsonify({'success': True, 'data': _report()})
    except Exception as e:
        eid = generate_error_id()
        logger.error(f'[{eid}] 수금 완료 정산 자료 오류: {e}', exc_info=True)
        return jsonify({'success': False, 'error': '자료를 불러오지 못했습니다.', 'error_id': eid}), 500


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
