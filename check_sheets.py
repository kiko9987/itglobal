#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import sys
from dotenv import load_dotenv

# 환경 변수 로드
load_dotenv()

# Google Sheets API
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Credentials
creds_file = os.getenv('GOOGLE_CREDENTIALS_FILE') or 'credentials.json'
sheet_id = os.getenv('GOOGLE_SHEET_ID')

print(f"Credentials: {creds_file}")
print(f"Sheet ID: {sheet_id}")
print()

try:
    # 인증
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    credentials = service_account.Credentials.from_service_account_file(
        creds_file, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=credentials)

    # 시트 메타데이터 가져오기
    metadata = service.spreadsheets().get(spreadsheetId=sheet_id).execute()

    print("=== 사용 가능한 시트(탭) 목록 ===")
    for sheet in metadata.get('sheets', []):
        title = sheet['properties']['title']
        sheet_id_num = sheet['properties']['sheetId']
        print(f"  - {title} (ID: {sheet_id_num})")

except Exception as e:
    print(f"ERROR: {e}")
    sys.exit(1)
