import os
import pandas as pd
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from googleapiclient.discovery import build
from datetime import datetime
import logging

# 로깅 설정
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class GoogleSheetsManager:
    """구글 시트 연동 관리 클래스"""
    
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    
    def __init__(self, credentials_file='credentials.json'):
        """
        구글 시트 매니저 초기화
        
        Args:
            credentials_file: 구글 API 서비스 계정 자격증명 파일 경로
        """
        self.credentials_file = credentials_file
        self.service = None
        self._authenticate()
    
    def _authenticate(self):
        """구글 API 인증 처리 (서비스 계정 방식)"""
        if not os.path.exists(self.credentials_file):
            raise FileNotFoundError(
                f"구글 API 자격증명 파일이 없습니다: {self.credentials_file}\n"
                "Google Cloud Console에서 서비스 계정 JSON 키를 다운로드하여 credentials.json으로 저장하세요."
            )
        
        try:
            # 서비스 계정 자격증명 로드
            creds = ServiceAccountCredentials.from_service_account_file(
                self.credentials_file, scopes=self.SCOPES
            )
            
            # 서비스 객체 생성
            self.service = build('sheets', 'v4', credentials=creds)
            logger.info("구글 시트 API 인증 완료 (서비스 계정)")
            
        except Exception as e:
            logger.error(f"구글 API 인증 실패: {str(e)}")
            raise
    
    def get_sheet_data(self, sheet_id, range_name='공사 현황!A:AM'):
        """
        구글 시트에서 데이터 가져오기 (에러 처리 강화)
        
        Args:
            sheet_id: 구글 시트 ID
            range_name: 데이터 범위
            
        Returns:
            pandas.DataFrame: 시트 데이터
        """
        try:
            if not sheet_id:
                raise ValueError("시트 ID가 제공되지 않았습니다.")
            
            # 시트 데이터 가져오기 (함수 계산 결과 포함)
            result = self.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name,
                valueRenderOption='FORMATTED_VALUE',  # 함수 계산 결과를 포맷된 값으로
                dateTimeRenderOption='FORMATTED_STRING'  # 날짜 포맷된 문자열로
            ).execute()
            
            values = result.get('values', [])
            
            if not values:
                logger.warning("시트에 데이터가 없습니다.")
                return pd.DataFrame()
            
            if len(values) < 2:
                logger.warning("시트에 헤더만 있고 데이터가 없습니다.")
                return pd.DataFrame(columns=values[0] if values else [])
            
            # DataFrame 생성
            df = pd.DataFrame(values[1:], columns=values[0])
            
            # 데이터 전처리
            df = self._preprocess_data(df)
            
            logger.info(f"구글 시트에서 {len(df)}행의 데이터를 가져왔습니다.")
            return df
            
        except Exception as e:
            error_msg = f"구글 시트 데이터 가져오기 오류: {str(e)}"
            logger.error(error_msg)
            
            # 구체적인 에러 유형별 처리
            if "404" in str(e):
                logger.error("시트를 찾을 수 없습니다. 시트 ID와 권한을 확인하세요.")
            elif "403" in str(e):
                logger.error("시트 접근 권한이 없습니다. 서비스 계정에 뷰어 권한을 부여하세요.")
            elif "401" in str(e):
                logger.error("인증 실패. 서비스 계정 키를 확인하세요.")
            
            raise
    
    def _preprocess_data(self, df):
        """데이터 전처리"""
        # 빈 행 제거
        df = df.dropna(how='all')
        
        # 프로젝트 코드가 있는 행만 필터링 (실제 데이터만)
        if '프로젝트 코드' in df.columns:
            original_count = len(df)
            df = df[df['프로젝트 코드'].notna() & (df['프로젝트 코드'].astype(str).str.strip() != '')]
            filtered_count = len(df)
            logger.info(f"프로젝트 코드 필터링: {original_count}행 → {filtered_count}행")
        else:
            logger.warning("프로젝트 코드 컬럼을 찾을 수 없습니다.")
        
        # 날짜 컬럼 처리
        date_columns = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')
        
        # 숫자 컬럼 처리 (함수 계산 결과 포함)
        numeric_columns = ['총액 1', '총액 2', '총액2', '계약금', '중도금', '잔금', 
                          '미수금', '미수금W', '제품대', '도급비', '자재비', '기타비', '순익', '마진율']
        for col in numeric_columns:
            if col in df.columns:
                # 쉼표, 원화기호, 공백 제거 후 숫자 변환 (구글 시트 포맷된 값 처리)
                df[col] = df[col].astype(str).str.replace(',', '').str.replace('￦', '').str.replace('₩', '').str.replace('-', '').str.strip()
                # 빈 문자열이나 'nan'을 NaN으로 처리
                df[col] = df[col].replace(['', 'nan', 'None'], pd.NA)
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # 불린 컬럼 처리
        boolean_columns = ['부가세', '수금 확인']
        for col in boolean_columns:
            if col in df.columns:
                df[col] = df[col].map({'TRUE': True, 'FALSE': False}).fillna(False)
        
        return df
    
    def get_sheet_metadata(self, sheet_id):
        """시트 메타데이터 가져오기"""
        try:
            result = self.service.spreadsheets().get(spreadsheetId=sheet_id).execute()
            return {
                'title': result.get('properties', {}).get('title', ''),
                'sheets': [sheet.get('properties', {}).get('title', '') 
                          for sheet in result.get('sheets', [])],
                'last_updated': datetime.now().isoformat()
            }
        except Exception as e:
            logger.error(f"시트 메타데이터 가져오기 오류: {str(e)}")
            return {}
    
    def validate_connection(self, sheet_id):
        """구글 시트 연결 테스트"""
        try:
            metadata = self.get_sheet_metadata(sheet_id)
            if metadata:
                logger.info(f"구글 시트 연결 성공: {metadata['title']}")
                return True
            else:
                logger.error("구글 시트 연결 실패")
                return False
        except Exception as e:
            logger.error(f"구글 시트 연결 테스트 오류: {str(e)}")
            return False

def test_google_sheets_connection():
    """구글 시트 연결 테스트 함수"""
    from dotenv import load_dotenv
    load_dotenv()
    
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if not sheet_id:
        print("GOOGLE_SHEET_ID가 .env 파일에 설정되지 않았습니다.")
        return
    
    try:
        manager = GoogleSheetsManager()
        if manager.validate_connection(sheet_id):
            print("✅ 구글 시트 연결 성공!")
            
            # 샘플 데이터 가져오기
            df = manager.get_sheet_data(sheet_id)
            print(f"📊 데이터 크기: {df.shape}")
            print(f"📋 컬럼 수: {len(df.columns)}")
        else:
            print("❌ 구글 시트 연결 실패")
    except Exception as e:
        print(f"❌ 오류 발생: {str(e)}")

    def append_row(self, sheet_id, values, range_name='공사 현황!A:AM'):
        """
        구글 시트에 새 행 추가
        
        Args:
            sheet_id: 구글 시트 ID
            values: 추가할 데이터 리스트
            range_name: 데이터 범위
            
        Returns:
            dict: 추가 결과
        """
        try:
            body = {
                'values': [values]
            }
            
            result = self.service.spreadsheets().values().append(
                spreadsheetId=sheet_id,
                range=range_name,
                valueInputOption='USER_ENTERED',
                insertDataOption='INSERT_ROWS',
                body=body
            ).execute()
            
            logger.info(f"새 행 추가 성공: {result.get('updates', {}).get('updatedRows', 0)}행")
            return result
            
        except Exception as e:
            logger.error(f"행 추가 오류: {str(e)}")
            raise
    
    def update_row(self, sheet_id, row_number, values, range_name='공사 현황!A{row}:AM{row}'):
        """
        구글 시트의 특정 행 업데이트
        
        Args:
            sheet_id: 구글 시트 ID
            row_number: 행 번호 (1부터 시작)
            values: 업데이트할 데이터 리스트
            range_name: 데이터 범위 템플릿
            
        Returns:
            dict: 업데이트 결과
        """
        try:
            # 범위 설정
            actual_range = range_name.format(row=row_number)
            
            body = {
                'values': [values]
            }
            
            result = self.service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=actual_range,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            
            logger.info(f"행 업데이트 성공: {row_number}행, {result.get('updatedCells', 0)}셀")
            return result
            
        except Exception as e:
            logger.error(f"행 업데이트 오류: {str(e)}")
            raise
    
    def find_row_by_project_code(self, sheet_id, project_code, range_name='공사 현황!A:A'):
        """
        프로젝트 코드로 행 번호 찾기
        
        Args:
            sheet_id: 구글 시트 ID
            project_code: 찾을 프로젝트 코드
            range_name: 검색할 범위
            
        Returns:
            int: 행 번호 (없으면 None)
        """
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            
            for i, row in enumerate(values):
                if row and len(row) > 0 and row[0] == project_code:
                    return i + 1  # 1부터 시작하는 행 번호
            
            return None
            
        except Exception as e:
            logger.error(f"행 찾기 오류: {str(e)}")
            return None
    
    def get_next_project_code(self, sheet_id, region_code='IT'):
        """
        다음 프로젝트 코드 생성
        
        Args:
            sheet_id: 구글 시트 ID
            region_code: 지역 코드 (예: IT, YG, JW 등)
            
        Returns:
            str: 새 프로젝트 코드
        """
        try:
            # 기존 데이터에서 해당 지역의 최대 번호 찾기
            df = self.get_sheet_data(sheet_id)
            
            if df.empty or '프로젝트 코드' not in df.columns:
                return f"G0001-{region_code}"
            
            # 해당 지역 코드가 포함된 프로젝트 코드 찾기
            region_projects = df[df['프로젝트 코드'].str.contains(f'-{region_code}', na=False)]
            
            if region_projects.empty:
                return f"G0001-{region_code}"
            
            # 번호 추출 및 최대값 찾기
            max_num = 0
            for code in region_projects['프로젝트 코드']:
                try:
                    # G0001-IT 형태에서 숫자 부분 추출
                    num_part = code.split('-')[0][1:]  # G 제거 후 숫자 부분
                    num = int(num_part)
                    max_num = max(max_num, num)
                except (ValueError, IndexError):
                    continue
            
            next_num = max_num + 1
            return f"G{next_num:04d}-{region_code}"
            
        except Exception as e:
            logger.error(f"프로젝트 코드 생성 오류: {str(e)}")
            return f"G0001-{region_code}"
    
    def get_column_mapping(self):
        """컬럼 매핑 정보 반환"""
        return {
            'A': '프로젝트 코드',
            'B': '사업자', 
            'C': '담당자',
            'D': '거래처',
            'E': '현장 주소',
            'F': '공사 구분',
            'G': '기계 분류',
            'H': '브랜드',
            'I': '공사 시작',
            'J': '공사 종료',
            'K': '공사 내용',
            'L': '도급 구분',
            'M': '시공자',
            'N': '현장 담당자',
            'O': '담당자 연락처',
            'P': '담당자 이메일',
            'Q': '총액 1',
            'R': '부가세',
            'S': '총액 2',
            'T': '계약금',
            'U': '중도금',
            'V': '잔금',
            'W': '미수금',
            'X': '계산서',
            'Y': '수금 날짜',
            'Z': '수금 확인',
            'AA': '제품대',
            'AB': '도급비',
            'AC': '자재비',
            'AD': '기타비',
            'AE': '순익',
            'AF': '마진율',
            'AG': '비고',
            'AH': '계약금 입금자명',
            'AI': '중도금 입금자명',
            'AJ': '잔금 입금자명',
            'AK': '견적서 및 계약서 폴더 경로',
            'AL': '공사 확정',
            'AM': 'Airtable Record ID'
        }

if __name__ == "__main__":
    test_google_sheets_connection()