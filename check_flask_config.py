#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
from dotenv import load_dotenv

load_dotenv()

print("=== Flask Configuration ===")
print(f"GOOGLE_SHEET_ID: {os.getenv('GOOGLE_SHEET_ID')}")
print(f"GOOGLE_CREDENTIALS_FILE: {os.getenv('GOOGLE_CREDENTIALS_FILE')}")
print(f"SHEET_RANGE: {os.getenv('SHEET_RANGE')}")
print()

# Also check project_config.json
import json
try:
    with open('dashboard/project_config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
    print("=== project_config.json ===")
    print(f"sheet_id: {config.get('sheet_id')}")
    print(f"sheet_range: {config.get('sheet_range')}")
except Exception as e:
    print(f"Error reading project_config.json: {e}")
