import requests

# API에서 첫 번째 프로젝트의 컬럼명 확인
response = requests.get('http://localhost:5002/api/projects/list')
data = response.json()

if data:
    first_project = data[0]
    
    print("=== 프로젝트 데이터 컬럼명 ===")
    for key in sorted(first_project.keys()):
        if '총액' in key or '미수금' in key or key in ['S', 'W']:
            value = first_project[key]
            print(f"{key}: {value} (type: {type(value)})")
    
    print("\n=== 모든 컬럼명 (처음 20개) ===")
    for i, key in enumerate(sorted(first_project.keys())[:20]):
        print(f"{i+1:2d}. {key}")
else:
    print("데이터가 없습니다.")