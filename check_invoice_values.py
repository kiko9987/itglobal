import requests
import json
from collections import Counter

# API 호출해서 계산서 필드 값들 확인
response = requests.get('http://localhost:5002/api/projects/list')
data = response.json()

print(f"총 프로젝트 수: {len(data)}")

# 계산서 필드의 모든 값들 수집
invoice_values = []
for project in data:
    invoice_value = project.get('계산서', '')
    if invoice_value and str(invoice_value).strip():
        invoice_values.append(str(invoice_value).strip())

print(f"\n계산서 필드에 값이 있는 프로젝트: {len(invoice_values)}개")

# 값들의 빈도수 확인
value_counts = Counter(invoice_values)
print(f"\n계산서 필드의 모든 값들:")
for value, count in value_counts.most_common():
    print(f"  '{value}': {count}개")

# 샘플 데이터 몇 개 확인
print(f"\n샘플 프로젝트들의 계산서 값:")
for i, project in enumerate(data[:10]):
    invoice_val = project.get('계산서', '')
    project_code = project.get('프로젝트 코드', 'N/A')
    print(f"  {project_code}: '{invoice_val}'")