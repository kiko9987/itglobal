"""
분석 및 리포트 블루프린트
요약 통계, 지역별/브랜드별 분석, 미수금 분석 등 데이터 분석 기능들
"""

import logging
from datetime import datetime
from flask import Blueprint, jsonify
from dashboard.auth import login_required
from dashboard.utils.logging_config import get_logger
from dashboard.utils.smart_cache_manager import smart_get, CacheStrategy
from dashboard.utils.error_helpers import generate_error_id

logger = get_logger(__name__)

analytics_bp = Blueprint('analytics', __name__, url_prefix='/api')

@analytics_bp.route('/summary')
@login_required
def get_summary():
    """요약 통계 API"""
    try:
        from dashboard.services.project_service import load_data, get_last_update
        from dashboard.utils.data_analyzer import DataAnalyzer
        from dashboard.utils.converters import convert_numpy_int64

        df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
        if df is None:
            df = load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500

        analyzer = DataAnalyzer(df)
        summary = analyzer.get_summary_stats()

        # numpy int64를 Python int로 변환
        summary = convert_numpy_int64(summary)

        # 추가 정보
        last_ts = get_last_update()
        summary['last_update'] = last_ts.isoformat() if last_ts else None
        summary['total_records'] = int(len(df))  # 명시적으로 int로 변환

        return jsonify(summary)

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 요약 통계 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '요약 통계를 생성할 수 없습니다.',
            'error_id': error_id
        }), 500

@analytics_bp.route('/regional-analysis')
@login_required
def get_regional_analysis():
    """지역별 분석 API"""
    try:
        from dashboard.services.project_service import load_data
        from dashboard.utils.data_analyzer import DataAnalyzer

        df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
        if df is None:
            df = load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500

        analyzer = DataAnalyzer(df)
        regional = analyzer.get_regional_analysis()

        result = regional.to_dict('records') if not regional.empty else []
        return jsonify(result)

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 지역별 분석 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '지역별 분석을 수행할 수 없습니다.',
            'error_id': error_id
        }), 500

@analytics_bp.route('/outstanding-analysis')
@login_required
def get_outstanding_analysis():
    """미수금 분석 API"""
    try:
        from dashboard.services.project_service import load_data
        from dashboard.utils.data_analyzer import DataAnalyzer

        df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
        if df is None:
            df = load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500

        analyzer = DataAnalyzer(df)
        outstanding = analyzer.get_outstanding_analysis()

        # DataFrame을 dict로 변환
        result = {}
        for key, value in outstanding.items():
            if hasattr(value, 'to_dict'):
                result[key] = value.to_dict('records')
            else:
                result[key] = value

        return jsonify(result)

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 미수금 분석 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '미수금 분석을 수행할 수 없습니다.',
            'error_id': error_id
        }), 500

@analytics_bp.route('/brand-analysis')
@login_required
def get_brand_analysis():
    """브랜드별 분석 API"""
    try:
        from dashboard.services.project_service import load_data
        from dashboard.utils.data_analyzer import DataAnalyzer

        df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
        if df is None:
            df = load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500

        analyzer = DataAnalyzer(df)
        brands = analyzer.get_brand_analysis()

        result = brands.to_dict('records') if not brands.empty else []
        return jsonify(result)

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 브랜드별 분석 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '브랜드별 분석을 수행할 수 없습니다.',
            'error_id': error_id
        }), 500

@analytics_bp.route('/monthly-sales')
@login_required
def get_monthly_sales():
    """월별 매출 API"""
    try:
        from dashboard.services.project_service import load_data
        from dashboard.utils.data_analyzer import DataAnalyzer

        df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
        if df is None:
            df = load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500

        analyzer = DataAnalyzer(df)
        monthly = analyzer.get_monthly_sales()

        result = monthly.to_dict('records') if not monthly.empty else []
        return jsonify(result)

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 월별 매출 API 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '월별 매출 분석을 수행할 수 없습니다.',
            'error_id': error_id
        }), 500