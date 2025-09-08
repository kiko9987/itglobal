import pandas as pd
from dashboard.utils.google_sheets import GoogleSheetsManager
import os
from dotenv import load_dotenv

load_dotenv()

manager = GoogleSheetsManager()
sheet_id = os.getenv('GOOGLE_SHEET_ID')
df = manager.get_sheet_data(sheet_id)

print(f'Total rows loaded: {len(df)}')

# 프로젝트 코드 컬럼 확인
if '프로젝트 코드' in df.columns:
    # 2806을 포함하는 모든 행 찾기
    filtered = df[df['프로젝트 코드'].astype(str).str.contains('2806', na=False)]
    print(f'Rows containing "2806": {len(filtered)}')
    
    if not filtered.empty:
        print("\n2806 관련 프로젝트:")
        for idx, row in filtered.iterrows():
            print(f"  - {row['프로젝트 코드']}: {row.get('현장 주소', 'N/A')}")
    
    # 최근 프로젝트 코드들 확인
    print("\n최근 프로젝트 코드 (뒤에서 10개):")
    recent_codes = df['프로젝트 코드'].dropna().tail(10)
    for code in recent_codes:
        print(f"  - {code}")
else:
    print("프로젝트 코드 컬럼이 없습니다!")
    print(f"Available columns: {list(df.columns[:10])}...")