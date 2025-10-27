#!/usr/bin/env python3
"""
구글 시트 쓰기 권한 테스트
"""

import os
import sys
from dotenv import load_dotenv

# 프로젝트 루트 추가
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

load_dotenv()

from dashboard.utils.google_sheets import GoogleSheetsManager

def test_write_permission():
    """구글 시트 쓰기 권한 테스트"""
    try:
        sheet_id = os.getenv('GOOGLE_SHEET_ID')
        if not sheet_id:
            print("[ERROR] GOOGLE_SHEET_ID가 설정되지 않았습니다.")
            return False
        
        manager = GoogleSheetsManager()
        
        # 첫 번째 프로젝트 찾기
        df = manager.get_sheet_data(sheet_id)
        if df.empty:
            print("[ERROR] 데이터를 가져올 수 없습니다.")
            return False
        
        first_project_code = df.iloc[0]['프로젝트 코드']
        print(f"테스트 프로젝트: {first_project_code}")
        
        # 행 번호 찾기
        row_number = manager.find_row_by_project_code(sheet_id, first_project_code)
        if not row_number:
            print("[ERROR] 프로젝트 행을 찾을 수 없습니다.")
            return False
        
        print(f"행 번호: {row_number}")
        
        # 현재 E열(현장 주소) 값 읽기
        current_range = f'공사 현황!E{row_number}'
        result = manager.service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=current_range
        ).execute()
        
        current_value = result.get('values', [['']])[0][0] if result.get('values') else ''
        print(f"현재 현장주소: {current_value}")
        
        # 테스트 값으로 업데이트 시도
        test_value = f"{current_value} [TEST UPDATE]"
        print(f"테스트 업데이트 시도: {test_value}")
        
        update_result = manager.service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=current_range,
            valueInputOption='USER_ENTERED',
            body={'values': [[test_value]]}
        ).execute()
        
        print(f"[SUCCESS] 업데이트 성공: {update_result.get('updatedCells', 0)}개 셀")
        
        # 원래 값으로 복원
        restore_result = manager.service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=current_range,
            valueInputOption='USER_ENTERED',
            body={'values': [[current_value]]}
        ).execute()
        
        print("[SUCCESS] 원래 값으로 복원 완료")
        return True
        
    except Exception as e:
        error_msg = str(e)
        print(f"[ERROR] 쓰기 테스트 실패: {error_msg}")
        
        if "403" in error_msg:
            print("[PERMISSION] 권한 문제: 서비스 계정에 구글 시트 편집 권한을 부여하세요.")
            print("   1. 구글 시트 열기")
            print("   2. 우측 상단 '공유' 버튼 클릭")  
            print("   3. 서비스 계정 이메일 주소 찾기")
            print("   4. 권한을 '편집자'로 변경")
        elif "404" in error_msg:
            print("[NOT_FOUND] 시트를 찾을 수 없음: 시트 ID와 시트명을 확인하세요.")
        
        return False

if __name__ == "__main__":
    print("구글 시트 쓰기 권한 테스트")
    print("=" * 50)
    
    success = test_write_permission()
    
    print("=" * 50)
    if success:
        print("[SUCCESS] 구글 시트 쓰기 권한이 정상적으로 작동합니다!")
    else:
        print("[ERROR] 구글 시트 쓰기 권한에 문제가 있습니다.")