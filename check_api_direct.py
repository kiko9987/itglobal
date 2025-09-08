import requests
import json

# API 직접 호출
response = requests.get('http://localhost:5000/api/projects/list')
data = response.json()

print(f"Total from API: {len(data)}")

# 프로젝트 코드가 있는 것만
with_codes = [p for p in data if p.get('프로젝트 코드') and str(p.get('프로젝트 코드')).strip()]
print(f"With project codes: {len(with_codes)}")

# G로 시작하는 것만
g_projects = [p for p in data if str(p.get('프로젝트 코드', '')).startswith('G')]
print(f"G projects: {len(g_projects)}")

# 2806 찾기
for p in data:
    code = str(p.get('프로젝트 코드', ''))
    if '2806' in code:
        print(f"\nFound 2806!")
        print(f"  Code: {code}")
        print(f"  Address: {p.get('현장 주소', 'N/A')}")
        # 모든 키 출력
        print("  All keys:", list(p.keys())[:10])
        break
else:
    print("\n2806 not found in API response")
    
# 마지막 G 프로젝트들
print("\nLast 5 G projects:")
last_g = [p for p in g_projects if p.get('프로젝트 코드')][-5:]
for p in last_g:
    print(f"  {p.get('프로젝트 코드')}: {p.get('현장 주소', 'N/A')[:30]}")