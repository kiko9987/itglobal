# 인라인 편집 API 엔드포인트 추가 코드
from flask import jsonify, request
from dashboard.utils.google_sheets import GoogleSheetsManager
import os
from datetime import datetime

def register_inline_update_route(app, socketio, load_data):
    @app.route('/api/update-project-inline', methods=['POST'])
    def update_project_inline():
        """프로젝트 인라인 편집 API - 구글 시트 직접 업데이트"""
        try:
            data = request.get_json()
            project_code = data.get('프로젝트 코드')
            
            if not project_code:
                return jsonify({'ok': False, 'error': '프로젝트 코드가 필요합니다.'}), 400
            
            sheet_id = os.getenv('GOOGLE_SHEET_ID')
            if not sheet_id:
                return jsonify({'ok': False, 'error': 'GOOGLE_SHEET_ID가 설정되지 않았습니다.'}), 500
            
            manager = GoogleSheetsManager()
            
            # 프로젝트가 있는 행 찾기
            row_number = manager.find_row_by_project_code(sheet_id, project_code)
            
            if not row_number:
                return jsonify({'ok': False, 'error': '프로젝트를 찾을 수 없습니다.'}), 404
            
            # 현재 행의 데이터를 가져오기
            current_row_range = f'공사 현황!A{row_number}:AM{row_number}'
            result = manager.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=current_row_range
            ).execute()
            
            current_values = result.get('values', [[]])[0] if result.get('values') else []
            
            # 현재 값을 리스트로 확장 (39개 컬럼)
            while len(current_values) < 39:
                current_values.append('')
            
            # 컬럼 매핑 가져오기
            column_mapping = manager.get_column_mapping()
            
            # 업데이트할 필드만 변경
            for field_name, new_value in data.items():
                if field_name == '프로젝트 코드':
                    continue
                
                # 필드명에 해당하는 컬럼 인덱스 찾기
                column_index = None
                for col_letter, col_name in column_mapping.items():
                    if col_name == field_name:
                        if len(col_letter) == 1:
                            column_index = ord(col_letter) - ord('A')
                        else:
                            column_index = (ord(col_letter[0]) - ord('A') + 1) * 26 + (ord(col_letter[1]) - ord('A'))
                        break
                
                if column_index is not None and column_index < len(current_values):
                    if new_value == '-' or new_value == '':
                        current_values[column_index] = ''
                    else:
                        current_values[column_index] = str(new_value)
            
            # 구글 시트 업데이트
            update_result = manager.update_row(sheet_id, row_number, current_values)
            
            # 로컬 데이터 새로고침
            load_data()
            
            # 실시간 업데이트 알림
            socketio.emit('data_updated', {
                'message': f"프로젝트가 수정되었습니다: {project_code}",
                'timestamp': datetime.now().isoformat(),
                'action': 'inline_update',
                'project_code': project_code,
                'updated_fields': list(data.keys())
            })
            
            return jsonify({
                'ok': True,
                'message': '성공적으로 업데이트되었습니다.',
                'project_code': project_code
            })
            
        except Exception as e:
            print(f"인라인 업데이트 오류: {str(e)}")
            return jsonify({'ok': False, 'error': str(e)}), 500
    
    return update_project_inline