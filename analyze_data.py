import pandas as pd
from dashboard.utils.google_sheets import GoogleSheetsManager
import os
from dotenv import load_dotenv

load_dotenv()

manager = GoogleSheetsManager()
sheet_id = os.getenv('GOOGLE_SHEET_ID')
df = manager.get_sheet_data(sheet_id)

print(f"Total rows: {len(df)}")

if '프로젝트 코드' in df.columns:
    # 프로젝트 코드가 있는 행만
    with_code = df[df['프로젝트 코드'].notna() & (df['프로젝트 코드'] != '')]
    print(f"Rows with project code: {len(with_code)}")
    
    # G로 시작하는 프로젝트 코드만
    g_projects = df[df['프로젝트 코드'].astype(str).str.startswith('G', na=False)]
    print(f"Projects starting with 'G': {len(g_projects)}")
    
    # 마지막 G 프로젝트들
    last_g_projects = g_projects['프로젝트 코드'].tail(10)
    print("\n마지막 10개 G 프로젝트:")
    for code in last_g_projects:
        print(f"  - {code}")
    
    # 2804, 2805, 2806 찾기
    for num in ['2804', '2805', '2806']:
        found = df[df['프로젝트 코드'].astype(str).str.contains(num, na=False)]
        if not found.empty:
            print(f"\n{num} 프로젝트:")
            for idx, row in found.iterrows():
                print(f"  - Row {idx+2}: {row['프로젝트 코드']}")
else:
    print("프로젝트 코드 컬럼이 없습니다!")