import requests
import json

# 세션을 사용해서 쿠키 유지
session = requests.Session()

# 먼저 메인 페이지에 접속해서 세션 생성
session.get('http://localhost:5000/')

# API 호출
response = session.get('http://localhost:5000/api/projects/list')

# 응답 확인
print(f"Status code: {response.status_code}")
print(f"Response text (first 200 chars): {response.text[:200]}")

if response.status_code != 200:
    print("API call failed, trying different approach...")
    # 직접 JSON 파일에서 데이터 로드
    try:
        with open('data/projects.json', 'r', encoding='utf-8') as f:
            data = json.load(f)
        print("Loaded data from local JSON file")
    except FileNotFoundError:
        print("Local JSON file not found, exiting...")
        exit(1)
else:
    data = response.json()

print(f"Total records: {len(data)}")

# 2805와 2806을 포함하는 모든 레코드 찾기
for search_term in ['2805', '2806']:
    found_records = []
    for i, record in enumerate(data):
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
for i, record in enumerate(data):
    code = record.get('프로젝트 코드', '')
    if str(code).startswith('G') and len(str(code)) > 1:
        g_codes.append((str(code), i))

# 정렬
g_codes.sort()

print("\n=== G 프로젝트 코드 (정렬됨) ===")
print(f"Total G codes: {len(g_codes)}")

# 2800 이후만 출력
after_2800 = [(code, idx) for code, idx in g_codes if '280' in code or '281' in code]
print("\n2800번대 프로젝트들:")
for code, idx in after_2800:
    print(f"  Index {idx}: {code}")