import requests

# API에서 담당자 목록 확인
response = requests.get('http://localhost:5002/api/projects/list')
data = response.json()

if data:
    # 담당자 컬럼 찾기
    manager_columns = []
    first_project = data[0]
    
    for key in first_project.keys():
        if '담당자' in key or '영업' in key:
            manager_columns.append(key)
    
    print("담당자 관련 컬럼:", manager_columns)
    
    # 담당자 목록 추출
    if manager_columns:
        main_manager_col = manager_columns[0]  # 첫 번째 담당자 컬럼 사용
        
        managers = set()
        for project in data:
            manager = project.get(main_manager_col, '')
            if manager and str(manager).strip():
                managers.add(str(manager).strip())
        
        print(f"\n'{main_manager_col}' 컬럼의 담당자 목록:")
        for manager in sorted(managers):
            print(f"  - {manager}")
        
        print(f"\n총 {len(managers)}명의 담당자")
    else:
        print("담당자 컬럼을 찾을 수 없습니다.")
        print("사용 가능한 컬럼 (처음 10개):")
        for i, key in enumerate(list(first_project.keys())[:10]):
            print(f"  {i+1}. {key}")
else:
    print("데이터가 없습니다.")