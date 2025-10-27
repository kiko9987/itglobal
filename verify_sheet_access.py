#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Verify sheet access by testing the exact same way Flask server does
"""
import os
import sys
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

load_dotenv()

creds_file = os.getenv('GOOGLE_CREDENTIALS_FILE') or 'credentials.json'
sheet_id = os.getenv('GOOGLE_SHEET_ID')

print(f"Credentials: {creds_file}")
print(f"Sheet ID: {sheet_id}")
print()

try:
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    credentials = service_account.Credentials.from_service_account_file(
        creds_file, scopes=SCOPES)
    service = build('sheets', 'v4', credentials=credentials)

    # Try to read exactly the range Flask is trying to read
    test_range = "공사 현황의 사본!A1:AM5000"
    print(f"Attempting to read: {test_range}")
    print(f"Range repr: {repr(test_range)}")
    print()

    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=test_range,
            valueRenderOption='FORMATTED_VALUE',
            dateTimeRenderOption='FORMATTED_STRING'
        ).execute()

        values = result.get('values', [])
        print(f"SUCCESS! Read {len(values)} rows")
        if values:
            print(f"First row columns: {len(values[0])}")
            print(f"First few columns: {values[0][:5]}")

    except HttpError as e:
        print(f"HTTP ERROR: {e.resp.status} - {e._get_reason()}")
        print(f"Error details: {e}")
        print()

        # Try alternate range format
        print("Trying alternate range format...")
        alt_range = "공사 현황의 사본!A1:AM"
        result2 = service.spreadsheets().values().get(
            spreadsheetId=sheet_id,
            range=alt_range,
            valueRenderOption='FORMATTED_VALUE',
            dateTimeRenderOption='FORMATTED_STRING'
        ).execute()

        values2 = result2.get('values', [])
        print(f"Alternate format SUCCESS! Read {len(values2)} rows")

except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
