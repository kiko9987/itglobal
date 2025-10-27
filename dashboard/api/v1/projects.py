"""
프로젝트 관리 API v1 엔드포인트
표준화된 응답 포맷과 버저닝 시스템 적용
"""

from flask import Blueprint, request, g
from datetime import datetime
from typing import Dict, List, Optional

from ..responses import APIResponse, NotFoundException, ValidationException, ConflictException
from ..validation_simple import (
    validate_json, validate_query_params, get_pagination_params,
    PaginationSchema, CreateProjectSchema, PatchProjectSchema
)
from ..versioning import versioned_route, version_compatibility_check

# v1 프로젝트 API Blueprint 생성
projects_v1_bp = Blueprint('projects_v1', __name__, url_prefix='/api/v1/projects')


@projects_v1_bp.route('', methods=['GET'])
@versioned_route('v1', '/projects')
@validate_query_params(PaginationSchema)
def list_projects():
    """
    프로젝트 목록 조회 (페이지네이션 지원)
    ---
    tags:
      - Projects v1
    summary: 프로젝트 목록 조회
    description: 페이지네이션과 필터링을 지원하는 프로젝트 목록 조회
    parameters:
      - name: page
        in: query
        type: integer
        minimum: 1
        default: 1
        description: 페이지 번호
      - name: limit
        in: query
        type: integer
        minimum: 1
        maximum: 100
        default: 20
        description: 페이지당 항목 수
      - name: sort
        in: query
        type: string
        default: created_at
        description: 정렬 필드
      - name: order
        in: query
        type: string
        enum: [asc, desc]
        default: desc
        description: 정렬 순서
      - name: status
        in: query
        type: string
        enum: [active, completed, cancelled, on_hold]
        description: 상태별 필터링
      - name: search
        in: query
        type: string
        description: 프로젝트명 검색
    responses:
      200:
        description: 프로젝트 목록 조회 성공
        schema:
          $ref: '#/definitions/StandardResponse'
    """
    from ..responses import PaginationHelper

    # 검증된 페이지네이션 파라미터 가져오기
    pagination_params = g.validated_data

    # 추가 필터 파라미터
    status_filter = request.args.get('status')
    search_query = request.args.get('search')

    # 모의 프로젝트 데이터 생성 (실제로는 데이터베이스에서 조회)
    mock_projects = generate_mock_projects(200)

    # 필터 적용
    filtered_projects = apply_filters(mock_projects, status_filter, search_query)

    # 페이지네이션 적용
    page = pagination_params['page']
    limit = pagination_params['limit']
    start_idx = (page - 1) * limit
    end_idx = start_idx + limit

    paginated_projects = filtered_projects[start_idx:end_idx]

    return PaginationHelper.paginated_response(
        items=paginated_projects,
        page=page,
        limit=limit,
        total=len(filtered_projects),
        message=f"{len(paginated_projects)}개의 프로젝트를 조회했습니다"
    )


@projects_v1_bp.route('', methods=['POST'])
@versioned_route('v1', '/projects')
@validate_json(CreateProjectSchema)
def create_project():
    """
    새 프로젝트 생성
    ---
    tags:
      - Projects v1
    summary: 프로젝트 생성
    description: 새로운 프로젝트를 생성합니다
    parameters:
      - name: body
        in: body
        required: true
        schema:
          type: object
          properties:
            name:
              type: string
              description: 프로젝트명 (필수)
              example: "웹사이트 리뉴얼"
            description:
              type: string
              description: 프로젝트 설명
              example: "회사 홈페이지 전면 리뉴얼 프로젝트"
            start_date:
              type: string
              format: date
              description: 시작일 (필수)
              example: "2025-01-01"
            end_date:
              type: string
              format: date
              description: 종료일
              example: "2025-06-30"
            budget:
              type: number
              description: 예산
              example: 50000000
          required:
            - name
            - start_date
    responses:
      201:
        description: 프로젝트 생성 성공
      400:
        description: 입력 데이터 유효성 검사 실패
      409:
        description: 중복된 프로젝트명
    """
    validated_data = g.validated_data

    # 중복 프로젝트명 체크 (모의)
    if check_duplicate_project_name(validated_data['name']):
        raise ConflictException(
            resource="Project",
            message=f"프로젝트명 '{validated_data['name']}'이 이미 존재합니다"
        )

    # 프로젝트 생성 (모의)
    new_project = create_new_project(validated_data)

    return APIResponse.created(
        data=new_project,
        message=f"프로젝트 '{new_project['name']}'이 성공적으로 생성되었습니다"
    )


@projects_v1_bp.route('/<project_code>', methods=['GET'])
@versioned_route('v1', '/projects/<project_code>')
def get_project(project_code: str):
    """
    특정 프로젝트 조회
    ---
    tags:
      - Projects v1
    summary: 프로젝트 상세 조회
    description: 프로젝트 코드로 특정 프로젝트의 상세 정보를 조회합니다
    parameters:
      - name: project_code
        in: path
        required: true
        type: string
        description: 프로젝트 코드
        example: "PRJ001"
    responses:
      200:
        description: 프로젝트 조회 성공
      404:
        description: 프로젝트를 찾을 수 없음
    """
    # 프로젝트 조회 (모의)
    project = find_project_by_code(project_code)

    if not project:
        raise NotFoundException("프로젝트", project_code)

    return APIResponse.success(
        data=project,
        message=f"프로젝트 '{project_code}' 정보를 조회했습니다"
    )


@projects_v1_bp.route('/<project_code>', methods=['PATCH'])
@versioned_route('v1', '/projects/<project_code>')
@validate_json(PatchProjectSchema)
def update_project(project_code: str):
    """
    프로젝트 부분 업데이트
    ---
    tags:
      - Projects v1
    summary: 프로젝트 업데이트
    description: 프로젝트의 특정 필드만 업데이트합니다
    parameters:
      - name: project_code
        in: path
        required: true
        type: string
        description: 프로젝트 코드
      - name: body
        in: body
        required: true
        schema:
          type: object
          properties:
            name:
              type: string
              description: 프로젝트명
            description:
              type: string
              description: 프로젝트 설명
            status:
              type: string
              enum: [active, completed, cancelled, on_hold]
              description: 프로젝트 상태
            end_date:
              type: string
              format: date
              description: 종료일
            budget:
              type: number
              description: 예산
            progress:
              type: integer
              minimum: 0
              maximum: 100
              description: 진행률
    responses:
      200:
        description: 프로젝트 업데이트 성공
      404:
        description: 프로젝트를 찾을 수 없음
      400:
        description: 입력 데이터 유효성 검사 실패
    """
    validated_data = g.validated_data

    # 프로젝트 존재 확인
    project = find_project_by_code(project_code)
    if not project:
        raise NotFoundException("프로젝트", project_code)

    # 프로젝트 업데이트 (모의)
    updated_project = update_project_data(project, validated_data)

    # 변경된 필드 추적
    changed_fields = list(validated_data.keys())

    return APIResponse.success(
        data=updated_project,
        message=f"프로젝트 '{project_code}'이 업데이트되었습니다",
        changed_fields=changed_fields
    )


@projects_v1_bp.route('/<project_code>/status', methods=['PUT'])
@versioned_route('v1', '/projects/<project_code>/status')
def change_project_status(project_code: str):
    """
    프로젝트 상태 변경
    ---
    tags:
      - Projects v1
    summary: 프로젝트 상태 변경
    description: 프로젝트의 상태를 변경합니다
    parameters:
      - name: project_code
        in: path
        required: true
        type: string
        description: 프로젝트 코드
      - name: body
        in: body
        required: true
        schema:
          type: object
          properties:
            status:
              type: string
              enum: [active, completed, cancelled, on_hold]
              description: 새로운 상태
            reason:
              type: string
              description: 상태 변경 사유
          required:
            - status
    responses:
      200:
        description: 상태 변경 성공
      404:
        description: 프로젝트를 찾을 수 없음
      400:
        description: 잘못된 상태 변경 요청
    """
    data = request.get_json() or {}
    new_status = data.get('status')
    reason = data.get('reason', '')

    if not new_status:
        return APIResponse.validation_error(
            details={'status': ['상태 필드는 필수입니다']},
            message="상태 변경 요청이 유효하지 않습니다"
        )

    # 프로젝트 존재 확인
    project = find_project_by_code(project_code)
    if not project:
        raise NotFoundException("프로젝트", project_code)

    # 상태 변경 가능 여부 확인
    if not is_valid_status_transition(project['status'], new_status):
        return APIResponse.error(
            "INVALID_STATUS_TRANSITION",
            f"'{project['status']}'에서 '{new_status}'로 상태 변경이 불가능합니다",
            details={
                "current_status": project['status'],
                "requested_status": new_status,
                "valid_transitions": get_valid_status_transitions(project['status'])
            },
            status_code=400
        )

    # 상태 변경 실행 (모의)
    updated_project = change_status(project, new_status, reason)

    return APIResponse.success(
        data=updated_project,
        message=f"프로젝트 '{project_code}'의 상태가 '{new_status}'로 변경되었습니다"
    )


@projects_v1_bp.route('/summary', methods=['GET'])
@versioned_route('v1', '/projects/summary')
def get_projects_summary():
    """
    프로젝트 요약 통계
    ---
    tags:
      - Projects v1
    summary: 프로젝트 요약 통계
    description: 전체 프로젝트의 요약 통계 정보를 조회합니다
    responses:
      200:
        description: 요약 통계 조회 성공
    """
    # 요약 통계 계산 (모의)
    summary = calculate_project_summary()

    return APIResponse.success(
        data=summary,
        message="프로젝트 요약 통계를 조회했습니다"
    )


# 유틸리티 함수들 (실제로는 서비스 레이어에 위치)

def generate_mock_projects(count: int) -> List[Dict]:
    """모의 프로젝트 데이터 생성"""
    projects = []
    statuses = ['active', 'completed', 'cancelled', 'on_hold']

    for i in range(1, count + 1):
        project = {
            "id": f"PRJ{i:03d}",
            "name": f"프로젝트 {i}",
            "description": f"프로젝트 {i}에 대한 설명입니다",
            "status": statuses[i % len(statuses)],
            "start_date": "2025-01-01",
            "end_date": "2025-12-31" if i % 3 != 0 else None,
            "budget": i * 1000000,
            "progress": min(i * 5, 100),
            "created_at": datetime.utcnow().isoformat() + "Z",
            "updated_at": datetime.utcnow().isoformat() + "Z",
            "created_by": "admin@itglobal.com"
        }
        projects.append(project)

    return projects


def apply_filters(projects: List[Dict], status_filter: str, search_query: str) -> List[Dict]:
    """프로젝트 필터링 적용"""
    filtered = projects

    if status_filter:
        filtered = [p for p in filtered if p['status'] == status_filter]

    if search_query:
        search_lower = search_query.lower()
        filtered = [p for p in filtered if search_lower in p['name'].lower()]

    return filtered


def check_duplicate_project_name(name: str) -> bool:
    """프로젝트명 중복 확인 (모의)"""
    # 실제로는 데이터베이스 조회
    return name.lower() in ['기존 프로젝트', 'duplicate project']


def create_new_project(data: Dict) -> Dict:
    """새 프로젝트 생성 (모의)"""
    import uuid

    new_project = {
        "id": f"PRJ{len(generate_mock_projects(50)) + 1:03d}",
        "name": data['name'],
        "description": data.get('description', ''),
        "status": "active",
        "start_date": data['start_date'].isoformat(),
        "end_date": data.get('end_date').isoformat() if data.get('end_date') else None,
        "budget": float(data.get('budget', 0)),
        "progress": 0,
        "created_at": datetime.utcnow().isoformat() + "Z",
        "updated_at": datetime.utcnow().isoformat() + "Z",
        "created_by": "current_user@itglobal.com"
    }

    return new_project


def find_project_by_code(project_code: str) -> Optional[Dict]:
    """프로젝트 코드로 프로젝트 찾기 (모의)"""
    projects = generate_mock_projects(50)
    return next((p for p in projects if p['id'] == project_code), None)


def update_project_data(project: Dict, update_data: Dict) -> Dict:
    """프로젝트 데이터 업데이트 (모의)"""
    updated = project.copy()
    updated.update(update_data)
    updated['updated_at'] = datetime.utcnow().isoformat() + "Z"

    # 날짜 필드 처리
    for date_field in ['start_date', 'end_date']:
        if date_field in update_data and update_data[date_field]:
            if hasattr(update_data[date_field], 'isoformat'):
                updated[date_field] = update_data[date_field].isoformat()

    return updated


def is_valid_status_transition(current_status: str, new_status: str) -> bool:
    """상태 전환 유효성 확인"""
    valid_transitions = {
        'active': ['completed', 'cancelled', 'on_hold'],
        'on_hold': ['active', 'cancelled'],
        'completed': [],  # 완료된 프로젝트는 상태 변경 불가
        'cancelled': []   # 취소된 프로젝트는 상태 변경 불가
    }

    return new_status in valid_transitions.get(current_status, [])


def get_valid_status_transitions(current_status: str) -> List[str]:
    """현재 상태에서 가능한 상태 전환 목록"""
    valid_transitions = {
        'active': ['completed', 'cancelled', 'on_hold'],
        'on_hold': ['active', 'cancelled'],
        'completed': [],
        'cancelled': []
    }

    return valid_transitions.get(current_status, [])


def change_status(project: Dict, new_status: str, reason: str) -> Dict:
    """프로젝트 상태 변경 (모의)"""
    updated = project.copy()
    updated['status'] = new_status
    updated['updated_at'] = datetime.utcnow().isoformat() + "Z"

    # 상태 변경 이력 추가 (실제로는 별도 테이블에 저장)
    if 'status_history' not in updated:
        updated['status_history'] = []

    updated['status_history'].append({
        "from_status": project['status'],
        "to_status": new_status,
        "reason": reason,
        "changed_at": datetime.utcnow().isoformat() + "Z",
        "changed_by": "current_user@itglobal.com"
    })

    return updated


def calculate_project_summary() -> Dict:
    """프로젝트 요약 통계 계산 (모의)"""
    projects = generate_mock_projects(100)

    summary = {
        "total_projects": len(projects),
        "by_status": {
            "active": len([p for p in projects if p['status'] == 'active']),
            "completed": len([p for p in projects if p['status'] == 'completed']),
            "cancelled": len([p for p in projects if p['status'] == 'cancelled']),
            "on_hold": len([p for p in projects if p['status'] == 'on_hold'])
        },
        "total_budget": sum(p['budget'] for p in projects),
        "average_progress": sum(p['progress'] for p in projects) / len(projects) if projects else 0,
        "upcoming_deadlines": 5,  # 임박한 마감일 프로젝트 수
        "overdue_projects": 2     # 지연된 프로젝트 수
    }

    return summary