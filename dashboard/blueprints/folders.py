"""
폴더 관리 블루프린트
Google Drive 폴더 통합, 프로젝트 폴더 관리, 폴더 ID 변환 등 파일 시스템 관리 기능들
"""

import os
import re
import logging
import subprocess
from datetime import datetime
from flask import Blueprint, jsonify, request, session
from dashboard.auth import login_required, admin_required
from dashboard.utils.logging_config import get_logger
from dashboard.utils.error_helpers import generate_error_id

logger = get_logger(__name__)

folders_bp = Blueprint('folders', __name__, url_prefix='/api')

def find_folder_path_by_drive_id(folder_id):
    """Google Drive 폴더 ID를 로컬 경로로 변환.

    우선순위:
    1) GOOGLE_DRIVE_WINDOWS_BASE_PATH 하위에서 Drive API로 조회한 폴더명 매칭
       (Slack 방문사진 파이프라인 등 신규 API 생성 폴더)
    2) .shortcut-targets-by-id/{id} (Drive Desktop 바로가기)
    """
    if not folder_id:
        return None

    # 1) BASE_PATH + Drive API 폴더명 조회
    base_path = os.getenv('GOOGLE_DRIVE_WINDOWS_BASE_PATH', '').strip()
    if base_path:
        try:
            from dashboard.utils.google_drive import get_folder_name
            folder_name = get_folder_name(folder_id)
            if folder_name:
                candidate = os.path.join(base_path, folder_name)
                if os.path.exists(candidate):
                    logger.info(f"[FOLDER_PATH] BASE_PATH 매칭: {candidate}")
                    return candidate
        except Exception as e:
            logger.debug(f"[FOLDER_PATH] Drive API 조회 실패: {e}")

    # 2) Google Drive for Desktop 바로가기 경로
    for drive_letter in ['G:', 'F:', 'E:', 'D:', 'C:']:
        local_path = f"{drive_letter}\\.shortcut-targets-by-id\\{folder_id}"
        if os.path.exists(local_path):
            logger.info(f"[FOLDER_PATH] 바로가기 매칭: {local_path}")
            return local_path

    logger.warning(f"[FOLDER_PATH] 폴더 ID {folder_id}의 로컬 경로를 찾을 수 없습니다")
    return None

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
                        'range': f'{sheet_name}!AL{sheet_row}',
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

    sheet_name = os.getenv('GOOGLE_SHEET_NAME', '공사 현황의 사본')
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

@folders_bp.route('/folder/open/<project_code>')
@login_required
def open_project_folder(project_code):
    """프로젝트 폴더를 윈도우 탐색기로 열기"""
    try:
        user_data = session.get('user', {})
        user_email = user_data.get('email', '')
        if not user_email:
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
            return jsonify({'success': False, 'error': '프로젝트를 찾을 수 없습니다'})

        # 폴더 경로 추출
        folder_path = project.get('견적서 및 계약서 폴더 경로', '')

        if not folder_path:
            return jsonify({'success': False, 'error': '폴더 경로가 설정되지 않았습니다'})

        # Google Drive 폴더 ID만 있는 경우 로컬 경로로 변환
        # 폴더 ID는 보통 20자 이상의 영숫자와 -_ 조합 (예: 1ZsnooHtIPe_4UxKq020Gk8TseLtEMdws)
        import re
        folder_id_pattern = r'^[a-zA-Z0-9_-]{20,}$'
        if re.match(folder_id_pattern, folder_path.strip()):
            # 폴더 ID로 판단되면 로컬 경로 찾기
            logger.info(f"[FOLDER_OPEN] 폴더 ID 감지: {project_code} -> {folder_path}")
            local_path = find_folder_path_by_drive_id(folder_path.strip())

            if local_path:
                logger.info(f"[FOLDER_OPEN] 로컬 경로 발견: {local_path}")
                folder_path = local_path
            else:
                logger.warning(f"[FOLDER_OPEN] 로컬 경로를 찾을 수 없어 웹 URL로 대체")
                folder_path = f'https://drive.google.com/drive/folders/{folder_path.strip()}'

        # Google Drive 링크인 경우 브라우저로 열기
        if 'drive.google.com' in folder_path:
            try:
                # Windows에서 기본 브라우저로 열기
                subprocess.run(['start', folder_path], shell=True, check=True)
                logger.info(f"[FOLDER_OPEN] Google Drive 폴더 열기 성공: {project_code}")

                # 감사 로그 기록
                try:
                    from dashboard.utils.user_database import get_audit_repository
                    audit_repo = get_audit_repository()
                    audit_repo.log_action(
                        user_email=user_email,
                        action='FOLDER_OPEN',
                        details=f'프로젝트 {project_code} 폴더 열기 (Google Drive)',
                        field_name='folder_path',
                        old_value='',
                        new_value=folder_path[:100],  # 경로가 길 수 있으므로 100자로 제한
                        ip_address=request.remote_addr
                    )
                except Exception as log_error:
                    logger.warning(f"감사 로그 기록 실패: {log_error}")

                return jsonify({
                    'success': True,
                    'message': 'Google Drive 폴더가 브라우저에서 열렸습니다',
                    'folder_type': 'google_drive'
                })

            except subprocess.CalledProcessError as e:
                logger.error(f"브라우저로 폴더 열기 실패: {e}")
                return jsonify({
                    'success': False,
                    'error': '브라우저로 폴더를 열 수 없습니다',
                    'folder_url': folder_path  # 사용자가 수동으로 열 수 있도록 URL 제공
                })

        # 로컬 경로인 경우 탐색기로 열기
        elif os.path.exists(folder_path):
            opened = False
            error_msg = None

            # 방법 1: PowerShell로 탐색기 열고 창 활성화
            try:
                logger.info(f"[FOLDER_OPEN] PowerShell로 탐색기 시작: {folder_path}")

                # PowerShell 스크립트 - 탐색기 열고 창을 강제로 포그라운드로 가져오기
                ps_script = f'''
                $path = "{folder_path}"
                $shell = New-Object -ComObject "Shell.Application"
                $shell.Open($path)
                Start-Sleep -Milliseconds 500

                # 탐색기 창 찾아서 활성화
                Add-Type @"
                    using System;
                    using System.Runtime.InteropServices;
                    public class WinAPI {{
                        [DllImport("user32.dll")]
                        [return: MarshalAs(UnmanagedType.Bool)]
                        public static extern bool SetForegroundWindow(IntPtr hWnd);

                        [DllImport("user32.dll")]
                        [return: MarshalAs(UnmanagedType.Bool)]
                        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

                        [DllImport("user32.dll")]
                        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
                    }}
"@

                # 잠시 대기 후 탐색기 창 찾기
                Start-Sleep -Milliseconds 300
                $explorerWindows = Get-Process | Where-Object {{$_.ProcessName -eq "explorer"}} | Select-Object -First 1
                if ($explorerWindows) {{
                    $hwnd = $explorerWindows.MainWindowHandle
                    if ($hwnd -ne [IntPtr]::Zero) {{
                        [WinAPI]::ShowWindow($hwnd, 9)  # SW_RESTORE
                        [WinAPI]::SetForegroundWindow($hwnd)
                    }}
                }}
                '''

                # PowerShell 실행
                subprocess.run(['powershell', '-WindowStyle', 'Hidden', '-Command', ps_script],
                              capture_output=True,
                              creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0)

                opened = True
                logger.info(f"[FOLDER_OPEN] PowerShell 성공")
            except Exception as e:
                logger.warning(f"[FOLDER_OPEN] PowerShell 실패: {e}")
                error_msg = str(e)

            # 방법 2: os.startfile() fallback
            if not opened:
                try:
                    logger.info(f"[FOLDER_OPEN] os.startfile() 시도")
                    os.startfile(folder_path)
                    opened = True
                    logger.info(f"[FOLDER_OPEN] os.startfile() 성공")
                except Exception as e:
                    logger.warning(f"[FOLDER_OPEN] os.startfile() 실패: {e}")
                    error_msg = str(e)

            # 방법 3: shell=True fallback
            if not opened:
                try:
                    logger.info(f"[FOLDER_OPEN] shell=True 시도")
                    subprocess.run(f'explorer "{folder_path}"', shell=True, check=True)
                    opened = True
                    logger.info(f"[FOLDER_OPEN] shell=True 성공")
                except Exception as e:
                    logger.error(f"[FOLDER_OPEN] shell=True 실패: {e}")
                    error_msg = str(e)

            if opened:
                # 감사 로그 기록
                try:
                    from dashboard.utils.user_database import get_audit_repository
                    audit_repo = get_audit_repository()
                    audit_repo.log_action(
                        user_email=user_email,
                        action='FOLDER_OPEN',
                        details=f'프로젝트 {project_code} 폴더 열기 (로컬)',
                        field_name='folder_path',
                        old_value='',
                        new_value=folder_path[:100],
                        ip_address=request.remote_addr
                    )
                except Exception as log_error:
                    logger.warning(f"감사 로그 기록 실패: {log_error}")

                return jsonify({
                    'success': True,
                    'message': '폴더가 탐색기에서 열렸습니다',
                    'folder_type': 'local'
                })
            else:
                logger.error(f"[FOLDER_OPEN] 모든 방법 실패: {error_msg}")
                return jsonify({'success': False, 'error': f'폴더를 열 수 없습니다: {error_msg}'})

        else:
            return jsonify({
                'success': False,
                'error': '폴더 경로가 올바르지 않거나 접근할 수 없습니다',
                'folder_path': folder_path
            })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 폴더 열기 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '폴더 열기 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500