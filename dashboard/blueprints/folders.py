"""
폴더 관리 블루프린트
Google Drive 폴더 통합, 프로젝트 폴더 관리, 폴더 ID 변환 등 파일 시스템 관리 기능들
"""

import os
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
    """Google Drive 폴더 ID를 로컬 경로로 변환 (간단 버전)"""
    if not folder_id:
        return None

    # Google Drive for Desktop의 .shortcut-targets-by-id 경로 확인
    # 우선순위: G: > F: > E: > D: > C:
    for drive_letter in ['G:', 'F:', 'E:', 'D:', 'C:']:
        local_path = f"{drive_letter}\\.shortcut-targets-by-id\\{folder_id}"
        if os.path.exists(local_path):
            logger.info(f"[FOLDER_PATH] 로컬 경로 발견: {local_path}")
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

        # Google Drive URL에서 폴더 ID 추출
        if folder_path and 'drive.google.com' in folder_path:
            try:
                # URL에서 폴더 ID 추출 로직
                if '/folders/' in folder_path:
                    folder_id = folder_path.split('/folders/')[-1].split('?')[0].split('/')[0]
                elif 'id=' in folder_path:
                    folder_id = folder_path.split('id=')[-1].split('&')[0]

                if folder_id:
                    # 캐시에 저장 (24시간)
                    smart_set(folder_id_cache_key, folder_id, ttl=86400)
                    logger.info(f"[FOLDER_ID] 폴더 ID 추출 성공: {project_code} -> {folder_id}")

                    return jsonify({'success': True, 'folder_id': folder_id})

            except Exception as e:
                logger.warning(f"[FOLDER_ID] URL 파싱 실패: {folder_path}, 오류: {e}")

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

@folders_bp.route('/folder/convert-paths-to-ids', methods=['POST'])
@admin_required
def convert_folder_paths_to_ids():
    """기존 프로젝트들의 폴더 경로를 Google Drive ID로 일괄 변환 (관리자 전용)"""
    try:
        from dashboard.utils.smart_cache_manager import smart_delete, smart_set
        from dashboard.services.project_service import load_data
        from dashboard.utils.google_sheets import get_sheets_manager

        logger.info("=== 폴더 경로 → ID 일괄 변환 시작 ===")

        # Google Sheets에서 직접 최신 데이터 읽기 (캐시 무시)
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            return jsonify({
                'success': False,
                'error': 'Google Sheets ID가 설정되지 않았습니다.'
            }), 500

        # 모든 캐시 삭제하여 완전히 새로운 데이터 보장
        smart_delete("current_sheet_data")
        smart_delete("project_data_cache")
        logger.info("🗑️ 모든 캐시 삭제 완료")

        # 최신 데이터 로드
        df = load_data(force_refresh=True)
        if df is None:
            return jsonify({
                'success': False,
                'error': '데이터를 불러올 수 없습니다.'
            }), 500

        # 폴더 경로 → ID 변환 작업
        converted_count = 0
        total_projects = len(df)
        results = []

        for idx, project in df.iterrows():
            project_code = project.get('프로젝트 코드', '')
            folder_path = project.get('견적서 및 계약서 폴더 경로', '')

            if not project_code:
                continue

            # Google Drive URL에서 폴더 ID 추출
            folder_id = None
            if folder_path and 'drive.google.com' in folder_path:
                try:
                    if '/folders/' in folder_path:
                        folder_id = folder_path.split('/folders/')[-1].split('?')[0].split('/')[0]
                    elif 'id=' in folder_path:
                        folder_id = folder_path.split('id=')[-1].split('&')[0]

                    if folder_id:
                        # 캐시에 저장
                        folder_id_cache_key = f"folder_id_{project_code}"
                        smart_set(folder_id_cache_key, folder_id, ttl=86400)
                        converted_count += 1

                        results.append({
                            'project_code': project_code,
                            'folder_id': folder_id,
                            'original_path': folder_path[:50] + '...' if len(folder_path) > 50 else folder_path
                        })

                except Exception as e:
                    logger.warning(f"프로젝트 {project_code} 변환 실패: {e}")
                    results.append({
                        'project_code': project_code,
                        'error': str(e),
                        'original_path': folder_path[:50] + '...' if len(folder_path) > 50 else folder_path
                    })

        # 감사 로그 기록
        try:
            from dashboard.utils.user_database import get_audit_repository
            audit_repo = get_audit_repository()
            admin_email = session.get('user', {}).get('email', 'unknown')
            audit_repo.log_action(
                user_email=admin_email,
                action='FOLDER_PATH_CONVERSION',
                details=f'폴더 경로 일괄 변환: {converted_count}/{total_projects}개 성공',
                field_name='folder_paths',
                old_value='기존 경로들',
                new_value=f'{converted_count}개 ID로 변환',
                ip_address=request.remote_addr
            )
        except Exception as log_error:
            logger.warning(f"감사 로그 기록 실패: {log_error}")

        logger.info(f"=== 폴더 경로 → ID 변환 완료: {converted_count}/{total_projects} ===")

        return jsonify({
            'success': True,
            'message': f'{converted_count}개 프로젝트의 폴더 ID 변환 완료',
            'converted_count': converted_count,
            'total_projects': total_projects,
            'results': results[:20]  # 처음 20개만 반환
        })

    except Exception as e:
        error_id = generate_error_id()
        logger.error(f"[{error_id}] 폴더 경로 변환 오류: {str(e)}", exc_info=True)
        return jsonify({
            'success': False,
            'error': '폴더 경로 변환 중 오류가 발생했습니다.',
            'error_id': error_id
        }), 500

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