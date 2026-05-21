"""
시공자(Constructor) 관리 블루프린트
- 활성 시공자 조회: 모든 로그인 사용자 (프로젝트 드롭다운용)
- 전체 CRUD: 관리자 전용
"""

from flask import Blueprint, jsonify, request

from dashboard.auth import admin_required, login_required
from dashboard.utils.logging_config import get_logger
from dashboard.utils.user_database import get_constructor_repository, ConstructorRepository

logger = get_logger(__name__)

constructors_bp = Blueprint('constructors', __name__, url_prefix='/api')


# -----------------------------------------------------------------------------
# 공개 (로그인 사용자) 엔드포인트
# -----------------------------------------------------------------------------

@constructors_bp.route('/constructors/active', methods=['GET'])
@login_required
def get_active_constructors():
    """활성 시공자 목록 - 카테고리별로 그룹화 (드롭다운용)"""
    try:
        repo = get_constructor_repository()
        grouped = repo.get_grouped(active_only=True)
        return jsonify({'success': True, 'data': grouped})
    except Exception as e:
        logger.error(f"[CONSTRUCTOR] 활성 시공자 조회 실패: {e}", exc_info=True)
        return jsonify({'success': False, 'message': '시공자 목록을 불러오지 못했습니다.'}), 500


@constructors_bp.route('/constructors/all-public', methods=['GET'])
@login_required
def get_all_constructors_public():
    """모든 시공자 (활성+비활성) - 카테고리별로 그룹화 (편집 모달용)

    기존 프로젝트에 저장된 비활성 시공자를 옵션으로 표시하기 위한 공용 엔드포인트.
    민감 정보(timestamps 등) 제외, 이름/카테고리/활성여부만 노출.
    """
    try:
        repo = get_constructor_repository()
        grouped = repo.get_grouped(active_only=False)
        # 민감 정보 제외, 필요 필드만
        simplified = {}
        for cat, items in grouped.items():
            simplified[cat] = [
                {'name': c['name'], 'is_active': c['is_active']}
                for c in items
            ]
        return jsonify({'success': True, 'data': simplified})
    except Exception as e:
        logger.error(f"[CONSTRUCTOR] 전체 공용 시공자 조회 실패: {e}", exc_info=True)
        return jsonify({'success': False, 'message': '시공자 목록을 불러오지 못했습니다.'}), 500


# -----------------------------------------------------------------------------
# 관리자 전용 엔드포인트
# -----------------------------------------------------------------------------

@constructors_bp.route('/admin/constructors', methods=['GET'])
@admin_required
def list_constructors():
    """전체 시공자 목록 (활성+비활성)"""
    try:
        repo = get_constructor_repository()
        all_constructors = repo.get_all(active_only=False)
        return jsonify({
            'success': True,
            'data': all_constructors,
            'total': len(all_constructors),
        })
    except Exception as e:
        logger.error(f"[CONSTRUCTOR] 전체 시공자 조회 실패: {e}", exc_info=True)
        return jsonify({'success': False, 'message': '시공자 목록을 불러오지 못했습니다.'}), 500


@constructors_bp.route('/admin/constructors', methods=['POST'])
@admin_required
def create_constructor():
    """시공자 추가

    Request body:
        {
            "name": "홍길동",
            "category": "메인" | "서브" | "내부",
            "is_active": true (선택, 기본 true)
        }
    """
    payload = request.get_json(silent=True) or {}
    name = payload.get('name')
    category = payload.get('category')
    is_active = payload.get('is_active', True)

    if not name or not category:
        return jsonify({'success': False, 'message': '이름과 카테고리는 필수입니다.'}), 400

    if category not in ConstructorRepository.VALID_CATEGORIES:
        return jsonify({
            'success': False,
            'message': f"잘못된 카테고리: {category} (허용: {', '.join(ConstructorRepository.VALID_CATEGORIES)})"
        }), 400

    try:
        repo = get_constructor_repository()
        ok, msg, data = repo.create(name=name, category=category, is_active=bool(is_active))
        if not ok:
            return jsonify({'success': False, 'message': msg}), 400
        return jsonify({'success': True, 'message': msg, 'data': data}), 201
    except Exception as e:
        logger.error(f"[CONSTRUCTOR] 추가 실패: {e}", exc_info=True)
        return jsonify({'success': False, 'message': '시공자 추가 중 오류가 발생했습니다.'}), 500


@constructors_bp.route('/admin/constructors/<int:constructor_id>', methods=['PATCH'])
@admin_required
def update_constructor(constructor_id: int):
    """시공자 수정 (부분 업데이트)

    Request body (모두 선택):
        {
            "name": "...",
            "category": "메인" | "서브" | "내부",
            "is_active": true | false,
            "display_order": 0
        }
    """
    payload = request.get_json(silent=True) or {}

    # 허용 필드만 추출
    name = payload.get('name')
    category = payload.get('category')
    is_active = payload.get('is_active')
    display_order = payload.get('display_order')

    if category is not None and category not in ConstructorRepository.VALID_CATEGORIES:
        return jsonify({
            'success': False,
            'message': f"잘못된 카테고리: {category}"
        }), 400

    try:
        repo = get_constructor_repository()
        ok, msg, data = repo.update(
            constructor_id=constructor_id,
            name=name,
            category=category,
            is_active=bool(is_active) if is_active is not None else None,
            display_order=int(display_order) if display_order is not None else None,
        )
        if not ok:
            status = 404 if '찾을 수 없' in msg else 400
            return jsonify({'success': False, 'message': msg}), status
        return jsonify({'success': True, 'message': msg, 'data': data})
    except Exception as e:
        logger.error(f"[CONSTRUCTOR] 수정 실패 (id={constructor_id}): {e}", exc_info=True)
        return jsonify({'success': False, 'message': '시공자 수정 중 오류가 발생했습니다.'}), 500


@constructors_bp.route('/admin/constructors/<int:constructor_id>', methods=['DELETE'])
@admin_required
def delete_constructor(constructor_id: int):
    """시공자 영구 삭제 (보통 is_active=false 권장)"""
    try:
        repo = get_constructor_repository()
        ok, msg = repo.delete(constructor_id)
        if not ok:
            return jsonify({'success': False, 'message': msg}), 404
        return jsonify({'success': True, 'message': msg})
    except Exception as e:
        logger.error(f"[CONSTRUCTOR] 삭제 실패 (id={constructor_id}): {e}", exc_info=True)
        return jsonify({'success': False, 'message': '시공자 삭제 중 오류가 발생했습니다.'}), 500
