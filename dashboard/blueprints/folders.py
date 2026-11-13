"""
폴더 관리 블루프린트
Google Drive 폴더 통합, 프로젝트 폴더 관리, 폴더 ID 변환 등 파일 시스템 관리 기능들
"""

import os
import re
import logging
from datetime import datetime
from flask import Blueprint, jsonify, request, session
from dashboard.auth import login_required, admin_required
from dashboard.utils.logging_config import get_logger
from dashboard.utils.error_helpers import generate_error_id

logger = get_logger(__name__)

folders_bp = Blueprint('folders', __name__, url_prefix='/api')

@folders_bp.route('/projects/<project_code>/folder_id')
@login_required
def get_project_folder_id(project_code):
    """프로젝트의 Google Drive 폴더 ID 반환"""
    try:
        from dashboard.utils.smart_cache_manager import smart_get, smart_set
        from dashboard.services.project_service import get_project_by_code

        # 캐시에서 폴더 ID 확인
        folder_id_cache_key = f"folder_id_{project_code}"
        cached_folder_id = smart_get(folder_id_cache_key)

        if cached_folder_id:
            logger.debug(f"[FOLDER_ID] 캐시에서 폴더 ID 조회: {project_code} -> {cached_folder_id}")
            return jsonify({'success': True, 'folder_id': cached_folder_id})

        # 캐시에 없으면 현재 폴더 경로에서 ID 추출 시도
        project_data = get_project_by_code(project_code)
        if not project_data:
            return jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다'})

        folder_path = project_data.get('견적서 및 계약서 폴더 경로', '')
        folder_id = None
        # 폴더 ID 유효 형식 (Google Drive ID 표준: 20~100자 alphanumeric + _ + -)
        _folder_id_pattern = re.compile(r'^[a-zA-Z0-9_-]{20,100}$')

        # Case 1: Google Drive URL에서 폴더 ID 추출
        if folder_path and 'drive.google.com' in folder_path:
            try:
                if '/folders/' in folder_path:
                    folder_id = folder_path.split('/folders/')[-1].split('?')[0].split('/')[0]
                elif 'id=' in folder_path:
                    folder_id = folder_path.split('id=')[-1].split('&')[0]

                if folder_id and not _folder_id_pattern.match(folder_id):
                    folder_id = None  # 파싱 결과 이상값이면 skip

                if folder_id:
                    smart_set(folder_id_cache_key, folder_id, ttl=86400)
                    logger.info(f"[FOLDER_ID] URL에서 추출: {project_code} -> {folder_id}")
                    return jsonify({'success': True, 'folder_id': folder_id})

            except Exception as e:
                logger.warning(f"[FOLDER_ID] URL 파싱 실패: {folder_path}, 오류: {e}")

        # Case 2: 셀에 이미 순수 폴더 ID가 저장된 경우 (convert_folder_paths_to_ids 실행 후)
        # 셀 값이 폴더 ID 형식이면 그대로 사용 + 캐시 저장
        if folder_path and _folder_id_pattern.match(folder_path.strip()):
            folder_id = folder_path.strip()
            smart_set(folder_id_cache_key, folder_id, ttl=86400)
            logger.info(f"[FOLDER_ID] 셀에 저장된 순수 ID 인식: {project_code} -> {folder_id}")
            return jsonify({'success': True, 'folder_id': folder_id})

        # 폴더 ID를 찾을 수 없는 경우
        logger.warning(f"[FOLDER_ID] 폴더 ID를 찾을 수 없음: {project_code}, 경로: {folder_path}")
        return jsonify({
            'success': False,
            'error': '폴더 ID를 찾을 수 없습니다',
            'folder_path': folder_path
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 폴더 ID 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '폴더 ID 조회 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500

# 폴더 ID 유효성 검사: 20~100자 alphanumeric + _ + - (Google Drive ID 표준)
_FOLDER_ID_PATTERN = re.compile(r'^[a-zA-Z0-9_-]{20,100}$')


def _extract_url_folder_id(folder_path: str):
    """Google Drive URL에서 folder_id 추출. 유효 형식만 반환, 아니면 None."""
    folder_id = None
    try:
        if '/folders/' in folder_path:
            folder_id = folder_path.split('/folders/')[-1].split('?')[0].split('/')[0]
        elif 'id=' in folder_path:
            folder_id = folder_path.split('id=')[-1].split('&')[0]
    except Exception:
        return None
    if folder_id and _FOLDER_ID_PATTERN.match(folder_id):
        return folder_id
    return None


def _extract_leaf_and_parent_from_windows_path(folder_path: str):
    """Windows 경로에서 마지막 폴더 이름과 부모 폴더 이름 추출.

    Returns:
        (leaf_name: str | None, parent_name: str | None)
    """
    # 2026-07-08 None/빈 값 방어 (test_core_flows에서 발견된 실제 버그)
    if not folder_path:
        return None, None
    normalized = folder_path.replace('/', '\\')
    segments = [s.strip() for s in normalized.split('\\') if s.strip()]
    if not segments:
        return None, None
    leaf = segments[-1]
    parent = segments[-2] if len(segments) >= 2 else None
    return leaf, parent


def _pick_folder_id_with_disambiguation(matches, parent_name):
    """Drive API 검색 결과 여러 개일 때 부모 폴더 이름으로 좁힘.

    Args:
        matches: [{'id', 'name', 'parents': [id...]}, ...]
        parent_name: 사용자 시트 경로에서 추출한 부모 폴더 이름

    Returns:
        folder_id (str) 정확히 1개로 좁혀지면. 아니면 None.
    """
    if len(matches) == 1:
        return matches[0]['id']
    if not parent_name:
        return None

    from dashboard.utils.google_drive import get_folder_name
    candidates = []
    for m in matches:
        parents = m.get('parents', []) or []
        for pid in parents:
            pname = get_folder_name(pid)
            if pname and pname.strip() == parent_name.strip():
                candidates.append(m)
                break
    if len(candidates) == 1:
        return candidates[0]['id']
    return None


def _background_folder_resolve(sheet_id, sheet_name, admin_email, ip_address):
    """백그라운드 스레드에서 실행. 시트 셀 URL/경로 → 폴더 ID 변환 + write.

    세 케이스:
      - Google Drive URL → 정규식 추출
      - 순수 폴더 ID (이미 변환된 상태) → 그대로 유지 (재실행 대비)
      - Windows 로컬 경로 → Drive API로 마지막 폴더 이름 검색 (부모로 disambiguate)

    실패한 항목은 시트 write 대상에 안 들어가서 원본 유지.
    """
    import time
    from dashboard.utils.smart_cache_manager import smart_delete, smart_set
    from dashboard.services.project_service import load_data
    from dashboard.utils.google_sheets import GoogleSheetsManager
    from dashboard.utils.google_drive import search_folder_by_name

    try:
        logger.info("[FOLDER_CONVERT_BG] 백그라운드 실행 시작")

        # 최신 데이터 로드 (캐시 무시)
        smart_delete("current_sheet_data")
        smart_delete("project_data_cache")
        df = load_data(force_refresh=True)
        if df is None:
            logger.error("[FOLDER_CONVERT_BG] 데이터 로드 실패")
            return

        total = len(df)
        stats = {
            'url_matched': 0,           # Case 1: Drive URL → ID 추출
            'already_id': 0,            # Case 2: 이미 순수 ID
            'name_unique': 0,           # Case 3a: 이름 검색 결과 1건
            'name_disambiguated': 0,    # Case 3b: 이름 여러 개 → 부모로 좁힘 성공
            'name_ambiguous_skip': 0,   # Case 3c: 여러 개 남음 → skip
            'name_not_found': 0,        # Case 3d: 이름 검색 결과 0건
            'invalid_or_empty': 0,      # 빈 값 · 형식 이상
            'drive_api_error': 0,       # API 예외
        }
        sheet_updates = []  # {'range', 'values'}
        # 폴더 경로 컬럼 레터를 필드명으로 파생 (하드코딩 금지 — 컬럼 시프트 자동 정정)
        #   (예전 기본값 'AO' 는 열이 늘며 틀린 칸이 됨 → 기본값 없이 헤더 지도에서만, 2026-10-22)
        _folder_col = GoogleSheetsManager().get_field_to_letter()['견적서 및 계약서 폴더 경로']

        for idx, project in df.iterrows():
            project_code = project.get('프로젝트 코드', '')
            folder_path = str(project.get('견적서 및 계약서 폴더 경로', '') or '').strip()

            if not project_code:
                continue

            folder_id = None

            if not folder_path or folder_path == '-':
                stats['invalid_or_empty'] += 1
            elif 'drive.google.com' in folder_path:
                folder_id = _extract_url_folder_id(folder_path)
                if folder_id:
                    stats['url_matched'] += 1
                else:
                    stats['invalid_or_empty'] += 1
            elif _FOLDER_ID_PATTERN.match(folder_path):
                # 이미 변환된 상태 — 재확인만
                folder_id = folder_path
                stats['already_id'] += 1
            elif folder_path.startswith(('G:', 'g:', 'C:', 'c:', 'D:', 'd:', 'F:', 'f:')) or '\\' in folder_path:
                leaf, parent = _extract_leaf_and_parent_from_windows_path(folder_path)
                if leaf:
                    try:
                        matches = search_folder_by_name(leaf, limit=10)
                        if not matches:
                            stats['name_not_found'] += 1
                        elif len(matches) == 1:
                            folder_id = matches[0]['id']
                            stats['name_unique'] += 1
                        else:
                            picked = _pick_folder_id_with_disambiguation(matches, parent)
                            if picked:
                                folder_id = picked
                                stats['name_disambiguated'] += 1
                            else:
                                stats['name_ambiguous_skip'] += 1
                                logger.info(
                                    f"[FOLDER_CONVERT_BG] {project_code} 동명 {len(matches)}개, "
                                    f"부모 매칭 후에도 실패 → skip (leaf={leaf!r}, parent={parent!r})"
                                )
                        time.sleep(0.05)  # Drive API rate limit 여유
                    except Exception as exc:
                        stats['drive_api_error'] += 1
                        logger.warning(f"[FOLDER_CONVERT_BG] {project_code} Drive API 오류: {exc}")
                else:
                    stats['invalid_or_empty'] += 1
            else:
                stats['invalid_or_empty'] += 1

            if folder_id and _FOLDER_ID_PATTERN.match(folder_id):
                folder_id_cache_key = f"folder_id_{project_code}"
                smart_set(folder_id_cache_key, folder_id, ttl=86400)
                # already_id 케이스는 시트가 이미 순수 ID라 write 불필요 (셀 값 이미 같음)
                # 원본 문자열과 다르면(즉 URL이나 Windows 경로 변환 결과) write 필요
                if folder_id != folder_path:
                    sheet_row = idx + 2
                    sheet_updates.append({
                        'range': f'{sheet_name}!{_folder_col}{sheet_row}',  # 폴더 경로 (필드명 파생)
                        'values': [[folder_id]]
                    })

            # 진행률 로그 (200개마다)
            processed = idx + 1
            if processed % 200 == 0 or processed == total:
                matched = stats['url_matched'] + stats['name_unique'] + stats['name_disambiguated'] + stats['already_id']
                logger.info(
                    f"[FOLDER_CONVERT_BG] 진행 {processed}/{total} — "
                    f"매칭 {matched}, skip {stats['name_ambiguous_skip']+stats['name_not_found']}, "
                    f"write 예정 {len(sheet_updates)}"
                )

        # 시트 batch write (500개 chunk)
        sheet_write_count = 0
        if sheet_updates:
            manager = GoogleSheetsManager()
            for i in range(0, len(sheet_updates), 500):
                chunk = sheet_updates[i:i + 500]
                try:
                    batch_result = manager.batch_update_cells(sheet_id, chunk)
                    written = batch_result.get('totalUpdatedCells', 0) if isinstance(batch_result, dict) else len(chunk)
                    sheet_write_count += written
                    logger.info(f"[FOLDER_CONVERT_BG] 시트 write chunk {i // 500 + 1}: {written}건")
                except Exception as exc:
                    logger.error(f"[FOLDER_CONVERT_BG] 시트 write chunk 실패: {exc}", exc_info=True)

            smart_delete("current_sheet_data")
            smart_delete("project_data_cache")

        # 감사 로그
        try:
            from dashboard.utils.user_database import get_audit_repository
            audit_repo = get_audit_repository()
            audit_repo.log_action(
                user_email=admin_email,
                action='FOLDER_PATH_CONVERSION',
                details=f'백그라운드 변환 완료: {stats}, 시트 셀 write {sheet_write_count}건',
                field_name='folder_paths',
                old_value='URL/경로/ID 혼재',
                new_value=f'{sheet_write_count}개 셀을 폴더 ID로 갱신',
                ip_address=ip_address,
            )
        except Exception as exc:
            logger.warning(f"[FOLDER_CONVERT_BG] 감사 로그 실패: {exc}")

        logger.info(
            f"[FOLDER_CONVERT_BG] 완료: total={total}, stats={stats}, "
            f"시트 write={sheet_write_count}"
        )

    except Exception as exc:
        logger.error(f"[FOLDER_CONVERT_BG] 백그라운드 실행 오류: {exc}", exc_info=True)


@folders_bp.route('/folder/convert-paths-to-ids', methods=['POST'])
@admin_required
def convert_folder_paths_to_ids():
    """폴더 경로 → Drive ID 일괄 변환 (관리자 전용, 백그라운드 실행).

    지원 케이스:
      1. Google Drive URL → 정규식 추출
      2. 이미 순수 폴더 ID → 그대로 유지 (재실행 안전)
      3. Windows 로컬 경로 (G:\\… 등) → Drive API 이름 검색 (부모로 disambiguate)

    3186건 규모라 15~25분 예상. 즉시 202 응답 후 스레드로 실행.
    진행률은 service_stdout.log의 [FOLDER_CONVERT_BG] 태그로 확인.
    """
    import threading

    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        return jsonify({
            'success': False,
            'error': 'Google Sheets ID가 설정되지 않았습니다.'
        }), 500

    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황')
    admin_email = session.get('user', {}).get('email', 'unknown')
    ip_address = request.remote_addr

    logger.info(f"=== 폴더 경로 변환 백그라운드 실행 요청: {admin_email} ===")

    thread = threading.Thread(
        target=_background_folder_resolve,
        args=(sheet_id, sheet_name, admin_email, ip_address),
        daemon=True,
        name='FolderPathResolveBG',
    )
    thread.start()

    return jsonify({
        'success': True,
        'message': '폴더 경로 변환이 백그라운드에서 시작됐습니다. 서버 로그의 [FOLDER_CONVERT_BG] 태그로 진행률 확인.',
        'background': True,
    }), 202

@folders_bp.route('/folder/name/<project_code>')
@login_required
def get_folder_name(project_code):
    """프로젝트 폴더명 가져오기"""
    try:
        logger.info(f"폴더명 API 호출됨 - 프로젝트 코드: {project_code}")

        user_data = session.get('user', {})
        user_email = user_data.get('email', '')
        logger.info(f"사용자 이메일: {user_email}")

        if not user_email:
            logger.error("로그인되지 않은 사용자의 폴더명 API 접근")
            return jsonify({'success': False, 'error': '로그인 필요'}), 401

        # 프로젝트 데이터 조회
        from dashboard.services.project_service import load_data
        df = load_data()
        if df is None:
            return jsonify({'success': False, 'error': '데이터를 불러올 수 없습니다'}), 500

        # DataFrame을 딕셔너리 리스트로 변환
        df = df.fillna('')  # NaN 값을 빈 문자열로 변환
        projects_data = df.to_dict('records')

        # 프로젝트 찾기
        project = None
        for proj in projects_data:
            if proj.get('프로젝트 코드', '') == project_code:
                project = proj
                break

        if not project:
            logger.warning(f"프로젝트 코드 {project_code}를 찾을 수 없습니다")
            return jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다'})

        # 폴더명 추출 (여러 가능한 컬럼명 확인)
        folder_name_keys = ['견적서 및 계약서 폴더명', '폴더명', 'folder_name']
        folder_name = ''

        for key in folder_name_keys:
            if project.get(key):
                folder_name = project.get(key)
                break

        # 폴더명이 없으면 프로젝트 코드로 대체
        if not folder_name:
            folder_name = project_code
            logger.info(f"폴더명이 없어 프로젝트 코드로 대체: {project_code}")

        logger.info(f"폴더명 조회 성공: {project_code} -> {folder_name}")

        return jsonify({
            'success': True,
            'folder_name': folder_name,
            'project_code': project_code
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 폴더명 조회 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '폴더명 조회 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500

def _resolve_drive_folder_id(value):
    """셀 값 → Drive 폴더 ID. 순수 폴더 ID 또는 https://drive.google.com/ 링크만 인정, 그 외는 None."""
    v = str(value or '').strip()
    if not v:
        return None
    if _FOLDER_ID_PATTERN.match(v):
        return v
    if v.startswith(('https://drive.google.com/', 'http://drive.google.com/')):
        return _extract_url_folder_id(v)
    return None


@folders_bp.route('/folder/open/<project_code>')
@login_required
def open_project_folder(project_code):
    """프로젝트 문서 폴더를 여는 데 필요한 정보를 돌려준다. 서버는 어떤 프로그램도 실행하지 않는다.

    2026-11-13 보안 점검: 예전에는 셀 값('견적서 및 계약서 폴더 경로')을 cmd `start`·PowerShell 문자열에
    그대로 넣어 서버에서 실행했다 → 편집자가 셀에 명령을 섞으면 서버(LocalSystem)에서 실행될 수 있었다.
    게다가 서버는 LocalSystem 세션이라 직원 화면에 창을 띄우지도 못했다(itgfolder.js 주석).
    그래서 서버 실행을 없애고, 검증된 Drive 폴더 ID와 웹 주소만 돌려준다. 실제 열기는 브라우저가 한다.
    다시 subprocess/os.startfile 로 되돌리지 말 것.
    """
    try:
        user_email = (session.get('user') or {}).get('email', '')
        if not user_email:
            return jsonify({'success': False, 'error': '로그인 필요'}), 401

        from dashboard.services.project_service import load_data
        df = load_data()
        if df is None:
            return jsonify({'success': False, 'error': '데이터를 불러올 수 없습니다'}), 500

        df = df.fillna('')
        project = next(
            (p for p in df.to_dict('records') if p.get('프로젝트 코드', '') == project_code),
            None,
        )
        if not project:
            return jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다'})

        folder_path = str(project.get('견적서 및 계약서 폴더 경로', '') or '').strip()
        if not folder_path or folder_path == '-':
            return jsonify({'success': False, 'error': '폴더 경로가 설정되지 않았습니다'})

        folder_id = _resolve_drive_folder_id(folder_path)
        if folder_id:
            return jsonify({
                'success': True,
                'folder_type': 'google_drive',
                'folder_id': folder_id,
                'url': f'https://drive.google.com/drive/folders/{folder_id}',
            })

        # Drive 폴더가 아닌 값(옛 PC 경로 등)은 서버가 열 수 없다 → 화면이 경로를 복사해 준다.
        return jsonify({
            'success': False,
            'folder_type': 'local_path',
            'folder_path': folder_path,
            'error': 'Drive 폴더가 아니라 서버에서 열 수 없습니다',
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 폴더 열기 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '폴더 열기 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500
