"""
인라인 편집을 위한 간단한 API 엔드포인트
"""

from flask import request, jsonify
import os
from dashboard.utils.google_sheets import GoogleSheetsManager
from datetime import datetime
import logging

logger = logging.getLogger(__name__)

def register_inline_update_routes(app, socketio, load_data):
    """인라인 업데이트 라우트 등록"""
    
    @app.route('/api/inline-update', methods=['POST'])
    def inline_update():
        """간단한 인라인 업데이트 API"""
        try:
            data = request.get_json()
            logger.info(f"인라인 업데이트 요청: {data}")
            
            project_code = data.get('projectCode')
            if not project_code:
                return jsonify({'ok': False, 'error': '프로젝트 코드가 필요합니다.'}), 400
            
            sheet_id = os.getenv('GOOGLE_SHEET_ID')
            if not sheet_id:
                return jsonify({'ok': False, 'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500
            
            manager = GoogleSheetsManager()
            
            # 프로젝트 코드로 행 찾기 - 직접 구현
            logger.info(f"프로젝트 코드 {project_code}의 행 번호를 찾는 중...")
            
            # A열에서 프로젝트 코드 검색
            search_result = manager.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range='공사 현황!A:A'
            ).execute()
            
            values = search_result.get('values', [])
            row_number = None
            
            for i, row in enumerate(values):
                if row and len(row) > 0 and row[0] == project_code:
                    row_number = i + 1  # 1부터 시작
                    break
            
            if not row_number:
                return jsonify({'ok': False, 'error': '프로젝트를 찾을 수 없습니다.'}), 404
            
            logger.info(f"프로젝트 {project_code}을 {row_number}행에서 발견")
            
            # 업데이트할 셀들
            updates = []
            
            # 필드별로 해당 열에 업데이트
            if data.get('company'):
                updates.append({
                    'range': f'공사 현황!B{row_number}',
                    'values': [[data['company']]]
                })
            
            if data.get('siteManager'):
                updates.append({
                    'range': f'공사 현황!N{row_number}',  # N열: 현장 담당자
                    'values': [[data['siteManager']]]
                })
            
            if data.get('contractType'):
                updates.append({
                    'range': f'공사 현황!L{row_number}',  # L열: 도급 구분
                    'values': [[data['contractType']]]
                })
            
            if data.get('managerPhone'):
                updates.append({
                    'range': f'공사 현황!O{row_number}',  # O열: 담당자 연락처
                    'values': [[data['managerPhone']]]
                })
            
            if data.get('constructor'):
                updates.append({
                    'range': f'공사 현황!M{row_number}',  # M열: 시공자
                    'values': [[data['constructor']]]
                })
            
            if data.get('managerEmail'):
                updates.append({
                    'range': f'공사 현황!P{row_number}',  # P열: 담당자 이메일
                    'values': [[data['managerEmail']]]
                })
            
            if data.get('address'):
                updates.append({
                    'range': f'공사 현황!E{row_number}',  # E열: 현장 주소
                    'values': [[data['address']]]
                })
            
            # 배치 업데이트 실행
            if updates:
                batch_update_body = {
                    'valueInputOption': 'USER_ENTERED',
                    'data': updates
                }
                
                logger.info(f"{len(updates)}개 셀 업데이트 시작...")
                
                batch_result = manager.service.spreadsheets().values().batchUpdate(
                    spreadsheetId=sheet_id,
                    body=batch_update_body
                ).execute()
                
                updated_cells = batch_result.get('totalUpdatedCells', 0)
                logger.info(f"업데이트 완료: {updated_cells}개 셀")
            
            # 로컬 데이터 새로고침
            load_data()
            
            # 실시간 알림
            if socketio:
                socketio.emit('data_updated', {
                    'message': f"프로젝트가 수정되었습니다: {project_code}",
                    'timestamp': datetime.now().isoformat(),
                    'action': 'inline_update',
                    'project_code': project_code
                })
            
            return jsonify({
                'ok': True,
                'message': '성공적으로 업데이트되었습니다.',
                'project_code': project_code,
                'updated_cells': updated_cells if updates else 0
            })
            
        except Exception as e:
            logger.error(f"인라인 업데이트 오류: {str(e)}", exc_info=True)
            return jsonify({'ok': False, 'error': str(e)}), 500
    
    return inline_update