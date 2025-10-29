"""
고객 리드 관리 블루프린트
리드 생성, 조회, 수정, 상태 변경 API 제공
"""

import logging
from flask import Blueprint, render_template, redirect, url_for, request, session, jsonify

from ..auth import login_required, get_user_role
from ..services.lead_service import (
    get_lead_records,
    get_lead_by_no,
    create_lead,
    update_lead,
    update_lead_status,
    invalidate_leads_cache,
    search_leads_by_customer
)
from ..utils.logging_config import get_logger
from ..api.responses import APIResponse, APIErrorCode

logger = get_logger(__name__)

# 블루프린트 생성
leads_bp = Blueprint('leads', __name__, url_prefix='/leads')


@leads_bp.route('/')
@login_required
def index():
    """
    리드 관리 메인 페이지
    """
    try:
        user_email = session.get('user_email', 'Unknown')
        user_role = get_user_role()

        return render_template(
            'leads.html',
            user_email=user_email,
            user_role=user_role
        )

    except Exception as exc:
        logger.error(f"[LEADS] 페이지 로드 실패: {exc}", exc_info=True)
        return redirect('/projects')


@leads_bp.route('/api/list', methods=['GET'])
@login_required
def api_list_leads():
    """
    리드 목록 조회 API

    Returns:
        {
            "success": true,
            "leads": [...],
            "count": 123
        }
    """
    try:
        # 강제 새로고침 여부
        force_refresh = request.args.get('force_refresh', 'false').lower() == 'true'

        if force_refresh:
            invalidate_leads_cache()

        # 리드 목록 조회
        leads = get_lead_records()

        return APIResponse.success(
            data={
                'leads': leads,
                'count': len(leads)
            }
        )

    except Exception as exc:
        logger.error(f"[API] 리드 목록 조회 실패: {exc}", exc_info=True)
        return APIResponse.error(
            message='리드 목록 조회 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR,
            status_code=500
        )


@leads_bp.route('/api/create', methods=['POST'])
@login_required
def api_create_lead():
    """
    새 리드 생성 API

    Request Body:
        {
            "거래처": "글로벌",
            "연락처": "010-1234-5678",
            "고객명": "홍길동",
            "방문 주소": "서울시 강남구...",
            "상담 내용": "주방 리모델링 상담",
            "방문 예정일": "2025-01-15",
            "담당자": "김철수",
            "비고": ""
        }

    Returns:
        {
            "success": true,
            "lead_no": "L0042",
            "message": "리드 L0042가 생성되었습니다"
        }
    """
    try:
        data = request.get_json()

        if not data:
            return APIResponse.error(
                message='요청 데이터가 없습니다',
                error_code=APIErrorCode.BAD_REQUEST,
                status_code=400
            )

        # 리드 생성
        result = create_lead(data)

        if not result['success']:
            return APIResponse.error(
                message=result['message'],
                error_code=APIErrorCode.VALIDATION_FAILED,
                status_code=400
            )

        return APIResponse.created(
            data={'lead_no': result['lead_no']},
            message=result['message']
        )

    except Exception as exc:
        logger.error(f"[API] 리드 생성 실패: {exc}", exc_info=True)
        return APIResponse.error(
            message='리드 생성 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR,
            status_code=500
        )


@leads_bp.route('/api/update/<lead_no>', methods=['PUT'])
@login_required
def api_update_lead(lead_no):
    """
    리드 정보 업데이트 API

    Args:
        lead_no: 리드 번호 (예: L0042)

    Request Body:
        {
            "방문 주소": "서울시 서초구...",
            "상담 내용": "추가 상담 내용",
            "비고": "재연락 필요"
        }

    Returns:
        {
            "success": true,
            "message": "리드 L0042가 업데이트되었습니다"
        }
    """
    try:
        data = request.get_json()

        if not data:
            return APIResponse.error(
                message='업데이트할 데이터가 없습니다',
                error_code=APIErrorCode.BAD_REQUEST,
                status_code=400
            )

        # 리드 업데이트
        result = update_lead(lead_no, data)

        if not result['success']:
            return APIResponse.error(
                message=result['message'],
                error_code=APIErrorCode.NOT_FOUND,
                status_code=404
            )

        return APIResponse.updated(
            message=result['message']
        )

    except Exception as exc:
        logger.error(f"[API] 리드 업데이트 실패: {exc}", exc_info=True)
        return APIResponse.error(
            message='리드 업데이트 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR,
            status_code=500
        )


@leads_bp.route('/api/status/<lead_no>', methods=['PATCH'])
@login_required
def api_update_lead_status(lead_no):
    """
    리드 상태 변경 API (빠른 업데이트)

    Args:
        lead_no: 리드 번호

    Request Body:
        {
            "status": "견적 제출"
        }

    Returns:
        {
            "success": true,
            "message": "상태가 업데이트되었습니다"
        }
    """
    try:
        data = request.get_json()

        if not data or 'status' not in data:
            return APIResponse.error(
                message='상태 값이 필요합니다',
                error_code=APIErrorCode.MISSING_PARAMETER,
                status_code=400
            )

        new_status = data['status']

        # 유효한 상태 값 검증
        valid_statuses = ['유선 상담', '방문 예약', '견적 제출', '방문 취소', '공사 확정', '공사 드랍']
        if new_status not in valid_statuses:
            return APIResponse.error(
                message=f'유효하지 않은 상태값입니다: {new_status}',
                error_code=APIErrorCode.INVALID_PARAMETER,
                status_code=400
            )

        # 상태 업데이트
        result = update_lead_status(lead_no, new_status)

        if not result['success']:
            return APIResponse.error(
                message=result['message'],
                error_code=APIErrorCode.NOT_FOUND,
                status_code=404
            )

        return APIResponse.updated(
            message=result['message']
        )

    except Exception as exc:
        logger.error(f"[API] 리드 상태 업데이트 실패: {exc}", exc_info=True)
        return APIResponse.error(
            message='상태 업데이트 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR,
            status_code=500
        )


@leads_bp.route('/api/search', methods=['GET'])
@login_required
def api_search_leads():
    """
    고객명/연락처로 리드 검색 API (프로젝트 등록 시 자동완성용)

    Query Params:
        - customer_name: 고객명 (필수)
        - phone: 연락처 (선택)

    Returns:
        {
            "success": true,
            "leads": [...],
            "count": 3
        }
    """
    try:
        customer_name = request.args.get('customer_name', '').strip()

        if not customer_name:
            return APIResponse.error(
                message='고객명이 필요합니다',
                error_code=APIErrorCode.MISSING_PARAMETER,
                status_code=400
            )

        phone = request.args.get('phone', '').strip()

        # 리드 검색
        results = search_leads_by_customer(customer_name, phone if phone else None)

        return APIResponse.success(
            data={
                'leads': results,
                'count': len(results)
            }
        )

    except Exception as exc:
        logger.error(f"[API] 리드 검색 실패: {exc}", exc_info=True)
        return APIResponse.error(
            message='리드 검색 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR,
            status_code=500
        )


@leads_bp.route('/api/<lead_no>', methods=['GET'])
@login_required
def api_get_lead(lead_no):
    """
    단일 리드 조회 API

    Args:
        lead_no: 리드 번호

    Returns:
        {
            "success": true,
            "lead": {...}
        }
    """
    try:
        lead = get_lead_by_no(lead_no)

        if not lead:
            return APIResponse.not_found(
                resource=f'리드 {lead_no}'
            )

        return APIResponse.success(
            data={'lead': lead}
        )

    except Exception as exc:
        logger.error(f"[API] 리드 조회 실패: {exc}", exc_info=True)
        return APIResponse.error(
            message='리드 조회 중 오류가 발생했습니다',
            error_code=APIErrorCode.INTERNAL_ERROR,
            status_code=500
        )
