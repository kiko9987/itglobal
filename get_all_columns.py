import requests

# API에서 모든 컬럼 확인
response = requests.get('http://localhost:5002/api/projects/list')
data = response.json()

if data and len(data) > 0:
    first_project = data[0]
    
    print("=== 모든 컬럼 목록 ===")
    columns = list(first_project.keys())
    
    for i, col in enumerate(columns, 1):
        sample_value = first_project[col]
        # 값이 너무 길면 자르기
        if isinstance(sample_value, str) and len(sample_value) > 50:
            sample_value = sample_value[:50] + "..."
        print(f"{i:2d}. {col}: {sample_value}")
    
    print(f"\n총 {len(columns)}개 컬럼")
else:
    print("데이터가 없습니다.")