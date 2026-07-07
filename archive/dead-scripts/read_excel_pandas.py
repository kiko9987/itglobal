import pandas as pd
import sys

# UTF-8 출력 설정
sys.stdout.reconfigure(encoding='utf-8')

try:
    # 엑셀 파일 읽기
    df = pd.read_excel('data/아이티 공사 현황 (2).xlsx', sheet_name='공사 현황')
    
    print("=== 데이터 기본 정보 ===")
    print(f"데이터 크기: {df.shape[0]}행 x {df.shape[1]}열")
    print(f"총 {len(df.columns)}개 컬럼")
    
    print("\n=== 컬럼명 ===")
    for i, col in enumerate(df.columns):
        print(f"{i+1:2d}. {col}")
    
    print("\n=== 처음 3행 샘플 데이터 ===")
    # 주요 컬럼만 선택해서 보기
    main_cols = ['프로젝트 코드', '업체명', '담당구', '판로처', '공사 주소', '공사 내용', 
                 '설치일', '공사 시작', '공사 완료', '고객 연락처', '업체 연락처명', '금액']
    
    # 실제 존재하는 컬럼만 선택
    available_cols = [col for col in main_cols if col in df.columns]
    
    if available_cols:
        print(df[available_cols].head(3).to_string(index=False))
    else:
        # 처음 5개 컬럼 표시
        print(df.iloc[:3, :5].to_string(index=False))
    
    print(f"\n=== 데이터 타입별 분포 ===")
    print(df.dtypes.value_counts())
    
    # 빈 값 체크
    print(f"\n=== 빈 값(결측치) 현황 ===")
    null_counts = df.isnull().sum()
    print(f"전체 빈 값 비율: {(null_counts.sum() / (df.shape[0] * df.shape[1]) * 100):.1f}%")
    
    # 주요 필드별 빈 값
    for col in df.columns[:10]:  # 처음 10개 컬럼만
        null_count = df[col].isnull().sum()
        if null_count > 0:
            print(f"  {col}: {null_count}개 ({null_count/len(df)*100:.1f}%)")

except Exception as e:
    print(f"오류 발생: {e}")
    import traceback
    traceback.print_exc()