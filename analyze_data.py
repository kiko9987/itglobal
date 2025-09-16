import sys
import os
sys.path.append('dashboard')

from utils.google_sheets import GoogleSheetsManager
from dotenv import load_dotenv

load_dotenv('.env')

def analyze_current_data():
    """현재 시스템의 데이터 구조 분석"""
    try:
        print("=== Google Sheets 데이터 구조 분석 ===\n")

        manager = GoogleSheetsManager()
        sheet_id = os.getenv('GOOGLE_SHEET_ID')

        # 헤더만 가져오기
        df_header = manager.get_sheet_data(sheet_id, '공사 현황의 사본!A1:AM1')

        print("Google Sheets 컬럼 구조:")
        print("-" * 50)
        for i, col in enumerate(df_header.columns):
            print(f"{i+1:2d}. {col}")

        print(f"\n총 {len(df_header.columns)}개 컬럼")

        # 샘플 데이터 가져오기 (처음 3개 행)
        df_sample = manager.get_sheet_data(sheet_id, '공사 현황의 사본!A1:AM4')

        print("\n=== 샘플 데이터 (처음 3개 프로젝트) ===")
        print("-" * 50)

        if len(df_sample) > 1:
            # 주요 컬럼들만 출력
            key_columns = ['프로젝트 코드', '현장명', '사업자', '담당자', '수금 관련 특이사항', '총액1', '총액2']

            for i, (_, row) in enumerate(df_sample.iterrows()):
                if i == 0:  # 헤더 스킵
                    continue
                print(f"\n프로젝트 {i}:")
                for col in key_columns:
                    if col in row:
                        print(f"  {col}: {row[col]}")

        # 취소된 프로젝트 찾기
        print("\n=== 취소된 프로젝트 분석 ===")
        print("-" * 50)

        df_all = manager.get_sheet_data(sheet_id, '공사 현황의 사본!A1:AM100')  # 100개만 분석

        if '수금 관련 특이사항' in df_all.columns:
            cancelled = df_all[df_all['수금 관련 특이사항'] == '공사취소']
            print(f"취소된 프로젝트: {len(cancelled)}개")

            if len(cancelled) > 0:
                print("취소된 프로젝트 목록:")
                for _, row in cancelled.head(5).iterrows():  # 최대 5개
                    code = row.get('프로젝트 코드', 'N/A')
                    name = row.get('현장명', 'N/A')
                    print(f"  - {code}: {name}")

        # 필수 필드 분석
        print("\n=== 필수 필드 분석 ===")
        print("-" * 50)

        required_fields = ['프로젝트 코드', '현장 주소', '현장명', '사업자', '담당자']
        missing_data = {}

        for field in required_fields:
            if field in df_all.columns:
                empty_count = df_all[field].isna().sum() + (df_all[field] == '').sum()
                missing_data[field] = empty_count
                print(f"{field}: {empty_count}개 누락")

        return df_header.columns.tolist()

    except Exception as e:
        print(f"오류 발생: {e}")
        return []

if __name__ == "__main__":
    columns = analyze_current_data()