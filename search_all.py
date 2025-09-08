import requests
import json

response = requests.get('http://localhost:5000/api/projects/list')
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