"""
메타데이터 관리 블루프린트
프로젝트 코드 미리보기, 드롭다운 옵션 등 UI 지원을 위한 메타데이터 제공 기능들
"""

import logging
import re
from datetime import datetime
from flask import Blueprint, jsonify, request
from dashboard.auth import login_required
from dashboard.utils.logging_config import get_logger
from dashboard.utils.smart_cache_manager import smart_get, CacheStrategy
from dashboard.utils.error_helpers import generate_error_id

logger = get_logger(__name__)

metadata_bp = Blueprint('metadata', __name__, url_prefix='/api')

@metadata_bp.route('/preview-project-code')
@login_required
def preview_project_code():
    """프로젝트 코드 미리보기 생성"""
    try:
        company = request.args.get('company', '').strip()
        owner = request.args.get('owner', '').strip()

        if not company or not owner:
            return jsonify({"ok": False, "error": "사업자와 담당자가 필요합니다"})

        # 현재 데이터 로드
        from dashboard.services.project_service import load_data
        df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
        if df is None:
            df = load_data()
        if df is None:
            return jsonify({"ok": False, "error": "데이터를 불러올 수 없습니다"})

        # 프로젝트 코드 생성 로직
        try:
            # 사업자명을 기반으로 회사 코드 생성 (구글 시트 규칙)
            company_code = ""
            if company:
                # 회사 코드 매핑
                company_mapping = {
                    '플렌트': 'P',
                    '글로벌': 'G',
                    '글로벌그룹': 'R'
                }
                company_code = company_mapping.get(company, 'G')  # 기본값 G

            # 담당자 이름을 기반으로 접미사 코드 생성
            # 우선순위: 1) 사용자 DB(Redis) 2) project_config 3) 폴백(이름 앞 2글자)
            owner_suffix = ""
            if owner:
                # 1. 먼저 사용자 DB에서 담당자 이름으로 이메일 찾기
                try:
                    from dashboard.utils.user_database import get_all_users
                    from dashboard.blueprints.users import _get_user_suffix

                    all_users = get_all_users()
                    user_email = None

                    # 이름으로 사용자 찾기
                    for user in all_users:
                        if user.get('name', '').strip() == owner.strip():
                            user_email = user.get('email')
                            break

                    # 사용자를 찾았으면 _get_user_suffix 함수로 접미사 조회
                    if user_email:
                        owner_suffix = _get_user_suffix(user_email, owner)
                        if owner_suffix:
                            logger.info(f"[PROJECT_CODE] 사용자 DB에서 접미사 찾음: {owner} → {owner_suffix}")

                except Exception as e:
                    logger.warning(f"[PROJECT_CODE] 사용자 DB 접미사 조회 실패: {e}")

                # 2. 사용자 DB에서 못 찾았으면 project_config 확인
                if not owner_suffix:
                    from dashboard.services.project_service import PROJECT_CONFIG
                    owner_suffix = PROJECT_CONFIG.get('owner_suffix_map', {}).get(owner, '')
                    if owner_suffix:
                        logger.info(f"[PROJECT_CODE] project_config에서 접미사 찾음: {owner} → {owner_suffix}")

                # 3. 둘 다 없으면 폴백 (이름 앞 2글자)
                if not owner_suffix:
                    owner_suffix = owner[:2].upper() if len(owner) >= 2 else owner.upper()
                    logger.info(f"[PROJECT_CODE] 폴백 접미사 생성: {owner} → {owner_suffix}")

            # 기존 프로젝트 코드와 중복 확인
            existing_codes = []
            if '프로젝트 코드' in df.columns:
                existing_codes = [str(code).strip() for code in df['프로젝트 코드'].dropna()
                                if str(code).strip() and str(code).strip() != 'nan']

            # 전체 프로젝트 코드에서 최대 순번 찾기 (사업자 무관)
            max_sequence = 0
            for code in existing_codes:
                if '-' in code:
                    try:
                        # G0001-IT, R0123-DN 등의 형식에서 숫자 부분 추출
                        # 알파벳 이후 하이픈 전까지의 숫자
                        parts = code.split('-')[0]
                        # 알파벳을 제외한 숫자만 추출
                        number_part = ''.join(filter(str.isdigit, parts))
                        if number_part:
                            max_sequence = max(max_sequence, int(number_part))
                    except Exception as e:
                        continue

            # 다음 순번 (4자리)
            sequence = max_sequence + 1

            # 프로젝트 코드 생성: G0001-IT 형식
            project_code = f"{company_code}{sequence:04d}-{owner_suffix}"

            logger.info(f"[PROJECT_CODE_PREVIEW] 생성된 코드: {project_code} (사업자: {company}, 담당자: {owner})")

            return jsonify({
                "ok": True,
                "project_code": project_code,
                "company": company,
                "owner": owner,
                "company_code": company_code,
                "owner_suffix": owner_suffix,
                "sequence": sequence
            })

        except Exception as code_error:
            logger.error(f"프로젝트 코드 생성 오류: {code_error}")
            # 기본 코드 생성 실패 시 간단한 대체 코드
            fallback_code = f"PROJ{datetime.now().strftime('%y%m%d%H%M')}"
            return jsonify({
                "ok": True,
                "project_code": fallback_code,
                "company": company,
                "owner": owner,
                "note": "기본 코드 생성 실패로 대체 코드 사용"
            })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 프로젝트 코드 미리보기 오류: {str(e)}", exc_info=True)
        return jsonify({
            "ok": False,
            "error": "프로젝트 코드 생성 중 오류가 발생했습니다.",
            "error_id": error_id
        })

@metadata_bp.route('/meta/options', methods=['GET'])
@login_required
def get_meta_options():
    """드롭다운용 옵션 API (사업자, 담당자 목록)

    메타데이터는 별도 캐시로 관리되며 10분 TTL을 가짐.
    프로젝트 생성/수정/삭제 시에도 메타데이터 캐시는 유지됨.
    """
    try:
        # 메타데이터 캐시 확인 (10분 TTL)
        metadata_cache_key = "metadata_options"
        cached_metadata = smart_get(metadata_cache_key, CacheStrategy.METADATA)

        if cached_metadata is not None:
            logger.debug("[META_OPTIONS] 캐시된 메타데이터 사용")
            return jsonify(cached_metadata)

        # 캐시 미스 시 데이터 로드 및 메타데이터 추출
        logger.info("[META_OPTIONS] 메타데이터 캐시 미스 - 새로 추출")
        from dashboard.services.project_service import load_data

        df = smart_get("current_sheet_data", CacheStrategy.CRITICAL_DATA)
        if df is None:
            df = load_data()
        if df is None:
            return jsonify({'error': '데이터를 불러올 수 없습니다.'}), 500

        # 사업자 목록 추출
        companies = []
        if "사업자" in df.columns:
            companies = sorted(set(x.strip() for x in df["사업자"].astype(str)
                                 if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a", "nan")))

        # 담당자 목록 추출
        owners = []
        if "담당자" in df.columns:
            owners = sorted(set(x.strip() for x in df["담당자"].astype(str)
                              if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a", "nan")))

        # 거래처 목록 추출 (추가 정보)
        clients = []
        if "거래처" in df.columns:
            clients = sorted(set(x.strip() for x in df["거래처"].astype(str)
                               if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a", "nan")))

        # 지역 목록 추출 (추가 정보)
        regions = []
        address_columns = ["현장 주소", "주소", "지역"]
        for col in address_columns:
            if col in df.columns:
                for addr in df[col].astype(str):
                    if addr and addr.strip() not in ("-", "없음", "N/A", "n/a", "nan"):
                        # 주소에서 시/도 추출
                        addr = addr.strip()
                        if addr:
                            # 정규식으로 시/도 추출
                            region_match = re.match(r'^([가-힣]+[시도])', addr)
                            if region_match:
                                regions.append(region_match.group(1))

        regions = sorted(set(regions))

        # 상태 목록 추출
        statuses = []
        status_columns = ["상태", "진행상태", "공사상태"]
        for col in status_columns:
            if col in df.columns:
                status_values = [x.strip() for x in df[col].astype(str)
                               if x.strip() and x.strip() not in ("-", "없음", "N/A", "n/a", "nan")]
                statuses.extend(status_values)

        statuses = sorted(set(statuses))

        # 통계 정보 추가
        total_projects = len(df)
        active_projects = 0

        # 활성 프로젝트 계산 (상태가 '완료'가 아닌 것들)
        for col in status_columns:
            if col in df.columns:
                completed_mask = df[col].astype(str).str.contains('완료|종료|마감', case=False, na=False)
                active_projects = len(df[~completed_mask])
                break

        # 메타데이터 응답 구성
        metadata_response = {
            'success': True,
            'options': {
                'companies': companies,
                'owners': owners,
                'clients': clients,
                'regions': regions,
                'statuses': statuses
            },
            'statistics': {
                'total_projects': total_projects,
                'active_projects': active_projects,
                'total_companies': len(companies),
                'total_owners': len(owners),
                'total_clients': len(clients)
            },
            'last_updated': datetime.now().isoformat()
        }

        # 메타데이터 캐시 저장 (10분 TTL)
        from dashboard.utils.smart_cache_manager import smart_set
        smart_set(metadata_cache_key, metadata_response, CacheStrategy.METADATA)

        logger.info(f"[META_OPTIONS] 메타데이터 추출 완료 - 사업자: {len(companies)}개, 담당자: {len(owners)}개, 거래처: {len(clients)}개")

        return jsonify(metadata_response)

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 메타 옵션 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '메타 옵션 조회 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500

@metadata_bp.route('/system-info')
@login_required
def get_system_info():
    """시스템 정보 조회 (추가 메타데이터)"""
    try:
        from dashboard.services.project_service import get_last_update

        # 마지막 업데이트 시간
        last_update = get_last_update()

        # 시스템 버전 정보 (환경변수나 설정에서 가져올 수 있음)
        import os
        system_info = {
            'app_name': 'ITG Dashboard',
            'version': '1.0.0',
            'last_data_update': last_update.isoformat() if last_update else None,
            'environment': os.getenv('FLASK_ENV', 'production'),
            'debug_mode': os.getenv('FLASK_DEBUG', 'False').lower() == 'true',
            'current_time': datetime.now().isoformat()
        }

        return jsonify({
            'success': True,
            'system_info': system_info
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 시스템 정보 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '시스템 정보 조회 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500