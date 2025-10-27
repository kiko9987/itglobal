#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
GoogleSheetsManager 싱글톤을 강제로 리셋하고 재초기화
"""
from dotenv import load_dotenv
load_dotenv()

# GoogleSheetsManager 싱글톤 리셋
from dashboard.utils.google_sheets import GoogleSheetsManager

print("Resetting GoogleSheetsManager singleton...")
GoogleSheetsManager._instance = None
GoogleSheetsManager._lock = None
print("Singleton reset complete")

# 새 인스턴스 생성 테스트
print("\nCreating new instance...")
manager = GoogleSheetsManager()
print("New instance created successfully")

# 테스트 읽기
import os
sheet_id = os.getenv('GOOGLE_SHEET_ID')
sheet_range = "공사 현황의 사본!A1:AM5000"

print(f"\nTesting with sheet_id: {sheet_id}")
print(f"Testing with range: {sheet_range}")

try:
    df = manager.get_sheet_data(sheet_id, sheet_range)
    print(f"SUCCESS: Read {len(df)} rows")
except Exception as e:
    print(f"FAILED: {e}")
