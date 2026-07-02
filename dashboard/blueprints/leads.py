"""
고객 리드 관리 블루프린트
리드 생성, 조회, 수정, 상태 변경 API 제공
"""

import logging
from flask import Blueprint, render_template, redirect, url_for, request, session, jsonify

from ..auth import login_required, get_user_role, editor_required
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
        user_email = session.get('user', {}).get('email', 'Unknown')
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
@editor_required
def api_create_lead():
    """
    새 리드 생성 API (온라인 문의 현황 15열 구조)

    Request Body:
        {
            "상담 시간": "2026.05.26 10:30",  # 비우면 현재 시각 자동
            "플랫폼": "홈페이지",              # 홈페이지/전화 등
            "상태": "유선 상담",               # 미지정 시 '유선 상담'
            "방문 예정일": "2026-06-01",
            "고객 연락처": "010-1234-5678",
            "이메일": "...",
            "고객명": "홍길동",
            "방문 주소": "서울시 강남구...",
            "상담 내용": "...",
            "키워드": "...",
            "온라인 상담자": "...",
            "영업 담당자": "조성헌",
            "마지막 연락일": "",
            "피드백": ""
        }

    옛 키 별칭(거래처→플랫폼, 담당자→영업 담당자, 연락처→고객 연락처, 비고→피드백)도 허용.

    Returns:
        {
            "success": true,
            "lead_no": "L-00042",
            "message": "리드 L-00042가 생성되었습니다"
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
@editor_required
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
@editor_required
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

        # 유효한 상태 값 검증 (온라인 문의 현황 시트 기준)
        valid_statuses = [
            '상담 대기', '유선 상담', '부재중',
            '방문 예약', '방문 대기', '방문 완료', '방문 취소',
            '견적 제출', '문의 드랍',
            '공사 확정', '공사 취소', '공사 드랍',
        ]
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


@leads_bp.route('/api/search-for-project', methods=['GET'])
@login_required
def api_search_leads_for_project():
    """새 프로젝트 등록 모달의 '리드 불러오기' 자동완성용.

    필터:
      - 기본: 로그인 사용자의 담당 리드만 (영업 담당자 = user_alias_map[email])
        · 리드 No 정확 매칭(score 100)은 담당자 무관 항상 노출 (트러스트: 정확한 ID 안다는 뜻)
        · ?all=1 으로 담당자 필터 우회 (관리자용)
      - 빈 검색어: 상태 = '방문 예약' / '견적 제출' 최근 리드만 (현재 파이프라인)
      - 검색어 있음: 상태 무관 전체 리드에서 매칭 (몇 달 전 문의 재개 케이스 지원)
      - 이미 프로젝트로 등록된 lead_no 는 항상 제외

    Query Params:
      - q: 검색어 (lead_no / 이름 / 연락처 / 주소, 빈값이면 최근 30건)
      - limit: 최대 반환 개수 (기본 30, 최대 50)
      - all: '1' 이면 담당자 필터 우회 (전체 팀 리드 검색)

    Returns:
      {
        "success": true,
        "leads": [
          {
            "lead_no": "L-03095", "name": "홍길동", "phone": "010-...",
            "email": "...", "address": "...", "platform": "홈페이지",
            "inquiry": "..." (인입 원본), "consultation": "..." (매니저 상담)
          },
          ...
        ]
      }
    """
    import re as _re
    try:
        q = (request.args.get('q') or '').strip().lower()
        q_digits = _re.sub(r'\D', '', q)
        show_all = (request.args.get('all') or '').strip() == '1'
        try:
            limit = int(request.args.get('limit', 30))
        except Exception:
            limit = 30
        limit = max(1, min(limit, 50))

        # 로그인 사용자의 영업 담당자 alias 조회 (user_alias_map)
        user_alias = ''
        if not show_all:
            try:
                from ..services.project_service import get_project_config
                cfg = get_project_config() or {}
                alias_map = cfg.get('user_alias_map') or {}
                user_email = (session.get('user') or {}).get('email', '')
                user_alias = str(alias_map.get(user_email, '') or '').strip()
            except Exception as exc:
                logger.warning(f"[API] user_alias_map 조회 실패: {exc}")

        # 이미 프로젝트에 등록된 lead_no 목록 (제외용)
        registered_leads = set()
        try:
            from ..services.project_service import get_project_records
            projects = get_project_records() or []
            for p in projects:
                ln = str(p.get('Lead No') or '').strip()
                if ln and ln.startswith('L-'):
                    registered_leads.add(ln)
        except Exception as exc:
            logger.warning(f"[API] 프로젝트 lead_no 조회 실패: {exc}")

        # 리드 목록 필터 + 매칭
        all_leads = get_lead_records() or []
        # 빈 query 기본 노출은 현재 파이프라인만, 검색어 있으면 상태 무관 전체
        PIPELINE_STATUSES = {'방문 예약', '견적 제출'}
        matched = []
        for lead in all_leads:
            lead_no = str(lead.get('리드 No') or '').strip()
            if not lead_no.startswith('L-'):
                continue
            if lead_no in registered_leads:
                continue
            status = str(lead.get('상태') or '').strip()

            # 빈 query 는 현재 파이프라인만 표시 (스크롤 폭탄 방지)
            if not q and status not in PIPELINE_STATUSES:
                continue

            name = str(lead.get('고객명') or '').strip()
            phone = str(lead.get('고객 연락처') or '').strip()
            phone_digits = _re.sub(r'\D', '', phone)
            address = str(lead.get('방문 주소') or '').strip()

            # 점수 매칭 — query 있으면 상태 무관 전체 검색 (몇 달 전 리드 재개 지원)
            # 현재 파이프라인 상태(방문 예약/견적 제출)에 소폭 가점을 줘서 상단 노출
            score = 0
            pipeline_bonus = 5 if status in PIPELINE_STATUSES else 0
            is_exact_lead_no = False
            if not q:
                score = 1
            elif q.upper() in lead_no.upper():
                score = 100 + pipeline_bonus
                is_exact_lead_no = True
            elif q in name.lower():
                score = 90 + pipeline_bonus
            elif q_digits and q_digits in phone_digits:
                score = 80 + pipeline_bonus
            elif q in address.lower():
                score = 50 + pipeline_bonus

            if score <= 0:
                continue

            # 담당자 필터 — 리드 No 정확 매칭이면 담당자 무관 노출 (트러스트: 사용자가 ID 안다는 뜻)
            if user_alias and not is_exact_lead_no:
                lead_owner = str(lead.get('영업 담당자') or '').strip()
                if lead_owner != user_alias:
                    continue

            try:
                sort_key = int(lead_no.split('-')[1])
            except Exception:
                sort_key = 0
            matched.append({
                'score': score,
                'sort_key': sort_key,
                'lead_no': lead_no,
                'name': name,
                'phone': phone,
                'email': str(lead.get('이메일') or '').strip(),
                'address': address,
                'platform': str(lead.get('플랫폼') or '').strip(),
                'status': status,
                'inquiry': str(lead.get('문의 내용') or lead.get('상담 내용') or '').strip(),
                'consultation': str(lead.get('상담 내용') or lead.get('피드백') or '').strip(),
                'folder_id_sheet': str(lead.get('폴더 ID') or '').strip(),  # 시트 P열 값
            })

        # 정렬 — 점수 내림, 최신 lead_no 우선
        matched.sort(key=lambda x: (-x['score'], -x['sort_key']))
        result = [
            {k: v for k, v in m.items() if k not in ('score', 'sort_key')}
            for m in matched[:limit]
        ]

        # 방문 사진 폴더 ID — 시트 P열 우선, 없으면 Redis fallback
        # (사진 업로드된 리드만 값 채워짐)
        try:
            from ..utils.redis_client import get_redis_client
            rc = get_redis_client().redis
        except Exception:
            rc = None
        for item in result:
            folder_id = item.pop('folder_id_sheet', '') or ''
            if not folder_id and rc is not None:
                try:
                    cached = rc.get(f"visit_folder:{item['lead_no']}")
                    if cached:
                        folder_id = cached.decode('utf-8') if isinstance(cached, bytes) else cached
                except Exception:
                    pass
            if folder_id:
                item['folder_id'] = folder_id
                item['folder_link'] = f"https://drive.google.com/drive/folders/{folder_id}"

        return APIResponse.success(data={'leads': result, 'count': len(result)})

    except Exception as exc:
        logger.error(f"[API] 프로젝트용 리드 검색 실패: {exc}", exc_info=True)
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
