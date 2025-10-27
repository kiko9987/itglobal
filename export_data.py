import os
import sys
import json
from dotenv import load_dotenv

# .env 파일 로드
load_dotenv()

# Add the dashboard directory to the Python path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'dashboard'))

# 환경변수 설정
os.environ['PROJECT_ROOT'] = os.path.dirname(__file__)

try:
    from dashboard.services.project_service import get_project_records

    print("직접 Google Sheets에서 데이터를 가져옵니다...")

    # 데이터 로드
    projects = get_project_records(force_refresh=True)

    print(f"총 {len(projects)}개의 프로젝트 데이터를 가져왔습니다.")

    # 2805와 2806을 포함하는 모든 레코드 찾기
    for search_term in ['2805', '2806']:
        found_records = []
        for i, record in enumerate(projects):
            for key, value in record.items():
                if search_term in str(value):
                    found_records.append({
                        'index': i,
                        'key': key,
                        'value': value,
                        'project_code': record.get('프로젝트 코드', 'N/A')
                    })

        print(f"\n=== {search_term} 검색 결과 ===")
        if found_records:
            for r in found_records:
                print(f"Index {r['index']}: {r['key']} = {r['value']}")
                print(f"  Project Code: {r['project_code']}")
        else:
            print(f"{search_term} not found")

    # G로 시작하는 모든 프로젝트 코드를 정렬해서 확인
    g_codes = []
    for i, record in enumerate(projects):
        code = record.get('프로젝트 코드', '')
        if str(code).startswith('G') and len(str(code)) > 1:
            g_codes.append((str(code), i))

    # 정렬
    g_codes.sort()

    print(f"\n=== G 프로젝트 코드 (정렬됨) ===")
    print(f"Total G codes: {len(g_codes)}")

    # 2800 이후만 출력
    after_2800 = [(code, idx) for code, idx in g_codes if '280' in code or '281' in code]
    print("\n2800번대 프로젝트들:")
    for code, idx in after_2800:
        print(f"  Index {idx}: {code}")

except Exception as e:
    print(f"오류 발생: {e}")
    print("환경변수와 경로를 확인해주세요.")