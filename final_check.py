import pandas as pd
from dashboard.utils.google_sheets import GoogleSheetsManager
from dashboard.app import load_data, current_data
import os
from dotenv import load_dotenv
import requests

load_dotenv()

print("=== 1. Google Sheets 직접 확인 ===")
manager = GoogleSheetsManager()
sheet_id = os.getenv('GOOGLE_SHEET_ID')
df_sheets = manager.get_sheet_data(sheet_id)
print(f"Sheets 데이터: {len(df_sheets)}행")

if '프로젝트 코드' in df_sheets.columns:
    g2806 = df_sheets[df_sheets['프로젝트 코드'].astype(str).str.contains('2806', na=False)]
    print(f"2806 in Sheets: {list(g2806['프로젝트 코드'])}")

print("\n=== 2. App 내부 데이터 확인 ===")
df_app = current_data if current_data is not None else load_data()
print(f"App 데이터: {len(df_app)}행")

if '프로젝트 코드' in df_app.columns:
    g2806_app = df_app[df_app['프로젝트 코드'].astype(str).str.contains('2806', na=False)]
    print(f"2806 in App: {list(g2806_app['프로젝트 코드'])}")
    
    # fillna 후
    df_filled = df_app.fillna('')
    records = df_filled.to_dict('records')
    found_2806 = [r for r in records if '2806' in str(r.get('프로젝트 코드', ''))]
    print(f"2806 after to_dict: {len(found_2806)} records")

print("\n=== 3. API 응답 확인 ===")
response = requests.get('http://localhost:5000/api/projects/list')
api_data = response.json()
print(f"API 응답: {len(api_data)}개")

api_2806 = [p for p in api_data if '2806' in str(p.get('프로젝트 코드', ''))]
print(f"2806 in API: {len(api_2806)} found")

# API 데이터의 마지막 G 프로젝트 확인
g_projects = [p for p in api_data if str(p.get('프로젝트 코드', '')).startswith('G')]
if g_projects:
    last_5 = g_projects[-5:]
    print("\nAPI의 마지막 5개 G 프로젝트:")
    for p in last_5:
        print(f"  - {p.get('프로젝트 코드')}")

print("\n=== 4. 문제 진단 ===")
print(f"Sheets에는 있음: {len(g2806) > 0}")
print(f"App 내부에 있음: {len(g2806_app) > 0}")
print(f"to_dict 후에도 있음: {len(found_2806) > 0}")
print(f"API 응답에 있음: {len(api_2806) > 0}")

if len(found_2806) > 0 and len(api_2806) == 0:
    print("\n문제: to_dict는 정상이지만 API 응답에서 누락됨")
    print("Flask API 라우트에서 문제가 발생하는 것으로 보임")