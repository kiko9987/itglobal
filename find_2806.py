import requests
import json

# API 호출
response = requests.get('http://localhost:5000/api/projects/list')
projects = response.json()

print(f"Total projects from API: {len(projects)}")

# 모든 값에서 G2806 찾기
found = False
for i, project in enumerate(projects):
    for key, value in project.items():
        if 'G2806' in str(value) or '2806' in str(value):
            found = True
            print(f"\nFound at index {i}:")
            print(f"  Key: {key}")
            print(f"  Value: {value}")
            # 주요 필드들 출력
            for k in project.keys():
                if any(x in str(k).lower() for x in ['addr', 'code', 'proj', 'id']):
                    print(f"  {k}: {project[k]}")
            break
    if found:
        break

if not found:
    print("\n2806을 찾을 수 없습니다.")