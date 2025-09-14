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
    """구글 시트 연동 관리 클래스 (Thread-Safe 싱글톤)"""
    
    SCOPES = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive.readonly'
    ]
    _instance = None
    _lock = None
    
    def __new__(cls, credentials_file='credentials.json'):
        """Thread-Safe 싱글톤 구현"""
        if cls._lock is None:
            import threading
            cls._lock = threading.Lock()
        
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(GoogleSheetsManager, cls).__new__(cls)
                cls._instance._initialized = False
            return cls._instance
    
    def __init__(self, credentials_file='credentials.json'):
        """
        구글 시트 매니저 초기화 (한 번만 실행)
        
        Args:
            credentials_file: 구글 API 서비스 계정 자격증명 파일 경로
        """
        if self._initialized:
            return
            
        self.credentials_file = credentials_file
        self.service = None
        self._authenticate()
        self._initialized = True
    
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
    
    def get_sheet_data(self, sheet_id, range_name='공사 현황의 사본!A:AM'):
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
            
            # 공사 종료 날짜 원시 데이터 재확인 (사용자 입력 데이터 확인)
            if '공사 종료' in df.columns:
                logger.info("=== 최근 10개 프로젝트의 공사 종료 원시 데이터 재확인 ===")
                recent_projects = df.tail(10)
                for i, (idx, row) in enumerate(recent_projects.iterrows()):
                    project_code = row.get('프로젝트 코드', 'N/A')
                    start_date = row.get('공사 시작', 'N/A')
                    end_date = row.get('공사 종료', 'N/A') 
                    logger.info(f"  프로젝트 {project_code}: 시작='{start_date}' 종료='{end_date}' (종료타입:{type(end_date)})")
            
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
        
        # 날짜 컬럼 처리 (빈 값은 빈 문자열로 유지)
        date_columns = ['공사 시작', '공사 종료', '수금 날짜', '공사 확정']
        for col in date_columns:
            if col in df.columns:
                # 빈 값이 아닌 경우에만 날짜 변환 시도
                df[col] = df[col].apply(lambda x: 
                    pd.to_datetime(x, errors='coerce') if x and str(x).strip() and str(x).strip().lower() != 'none' 
                    else pd.NaT
                )
        
        # 숫자 컬럼 처리 (함수 계산 결과 포함) - 성능 최적화
        numeric_columns = ['총액 1', '총액 2', '총액2', '계약금', '중도금', '잔금', 
                          '미수금', '미수금W', '제품대', '도급비', '자재비', '기타비', '순익', '마진율']
        
        # 존재하는 숫자 컬럼만 필터링하여 처리 속도 향상
        existing_numeric_columns = [col for col in numeric_columns if col in df.columns]
        
        if existing_numeric_columns:
            # 벡터화된 일괄 처리로 성능 대폭 향상
            for col in existing_numeric_columns:
                df[col] = (df[col].astype(str, copy=False)
                          .str.replace(r'[,￦₩\-]', '', regex=True)
                          .str.strip()
                          .replace(['', 'nan', 'None', 'NaN'], pd.NA))
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

    def find_next_empty_row(self, sheet_id, range_name='공사 현황의 사본!A:A'):
        """
        다음 빈 행 번호 찾기 (수식이 미리 설정된 행에 데이터 추가용)
        
        Returns:
            int: 다음 빈 행 번호 (1부터 시작)
        """
        try:
            result = self.service.spreadsheets().values().get(
                spreadsheetId=sheet_id,
                range=range_name,
                valueRenderOption='FORMATTED_VALUE'
            ).execute()
            
            values = result.get('values', [])
            
            # 프로젝트 코드(A열)가 비어있는 첫 번째 행 찾기
            for i, row in enumerate(values):
                if i == 0:  # 헤더 행 스킵
                    continue
                    
                if not row or len(row) == 0 or not row[0] or not row[0].strip():
                    return i + 1  # 행 번호는 1부터 시작
            
            # 빈 행을 찾지 못한 경우 마지막 다음 행 반환
            return len(values) + 1
            
        except Exception as e:
            logger.error(f"빈 행 찾기 오류: {str(e)}")
            return None

    def append_row(self, sheet_id, values, range_name='공사 현황의 사본!A:AM'):
        """
        구글 시트의 다음 빈 행에 데이터 추가 (수식이 미리 설정된 행에 덮어쓰기)
        
        Args:
            sheet_id: 구글 시트 ID
            values: 추가할 데이터 리스트
            range_name: 데이터 범위 (사용하지 않음, 호환성용)
            
        Returns:
            dict: 추가 결과
        """
        try:
            # 다음 빈 행 번호 찾기
            next_row = self.find_next_empty_row(sheet_id)
            if not next_row:
                raise Exception("다음 빈 행을 찾을 수 없습니다")
            
            # 특정 행에 데이터 업데이트 (수식이 있는 빈 행에 덮어쓰기)
            actual_range = f'공사 현황의 사본!A{next_row}:AM{next_row}'
            body = {
                'values': [values]
            }
            
            result = self.service.spreadsheets().values().update(
                spreadsheetId=sheet_id,
                range=actual_range,
                valueInputOption='USER_ENTERED',
                body=body
            ).execute()
            
            logger.info(f"빈 행({next_row})에 데이터 추가 성공: {result.get('updatedCells', 0)}셀 업데이트")
            return result
            
        except Exception as e:
            logger.error(f"빈 행 데이터 추가 오류: {str(e)}")
            raise
    
    def update_row(self, sheet_id, row_number, values, range_name='공사 현황의 사본!A{row}:AM{row}'):
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
    
    def batch_update_cells(self, sheet_id, updates):
        """
        구글 시트의 여러 셀을 일괄 업데이트
        
        Args:
            sheet_id: 구글 시트 ID
            updates: 업데이트할 데이터 리스트 [{'range': 'A1', 'values': [['data']]}]
            
        Returns:
            bool: 성공 여부
        """
        try:
            # batchUpdate API를 사용하여 여러 셀 일괄 업데이트
            data = []
            for update in updates:
                data.append({
                    'range': update['range'],
                    'values': update['values']
                })
            
            body = {
                'valueInputOption': 'USER_ENTERED',
                'data': data
            }
            
            result = self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=sheet_id,
                body=body
            ).execute()
            
            updated_cells = result.get('totalUpdatedCells', 0)
            logger.info(f"배치 업데이트 성공: {updated_cells}개 셀 업데이트")
            return True
            
        except Exception as e:
            logger.error(f"배치 업데이트 오류: {str(e)}")
            return False
    
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
            'AG': '수금 관련 특이사항',
            'AH': '계약금 입금자명',
            'AI': '중도금 입금자명',
            'AJ': '잔금 입금자명',
            'AK': '견적서 및 계약서 폴더 경로',
            'AL': '공사 확정',
            'AM': 'Airtable Record ID'
        }

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

if __name__ == "__main__":
    test_google_sheets_connection()