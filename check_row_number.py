import pandas as pd
from dashboard.utils.google_sheets import GoogleSheetsManager
import os
from dotenv import load_dotenv

load_dotenv()

# 범위를 늘려서 데이터 가져오기
manager = GoogleSheetsManager()
sheet_id = os.getenv('GOOGLE_SHEET_ID')

# 원래 범위로 가져오기
os.environ['SHEET_RANGE'] = '공사 현황!A1:AM10000'
df1 = manager.get_sheet_data(sheet_id)
print(f"Original range (A1:AM10000): {len(df1)} rows")

# 범위를 늘려서 가져오기
os.environ['SHEET_RANGE'] = '공사 현황!A1:AM20000'
df2 = manager.get_sheet_data(sheet_id)
print(f"Extended range (A1:AM20000): {len(df2)} rows")

# 2806 찾기
if '프로젝트 코드' in df2.columns:
    g2806 = df2[df2['프로젝트 코드'].astype(str).str.contains('2806', na=False)]
    if not g2806.empty:
        for idx, row in g2806.iterrows():
            print(f"\nG2806 found at row {idx + 2}:")  # +2 because of 0-index and header
            print(f"  프로젝트 코드: {row['프로젝트 코드']}")
            print(f"  현장 주소: {row.get('현장 주소', 'N/A')}")
    
    # 마지막 몇 개 프로젝트 확인
    print("\n마지막 5개 프로젝트:")
    last_projects = df2['프로젝트 코드'].dropna().tail(5)
    for i, code in enumerate(last_projects):
        print(f"  Row {len(df2) - 5 + i + 2}: {code}")