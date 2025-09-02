import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional
import logging

logger = logging.getLogger(__name__)

class DataAnalyzer:
    """냉난방기 설치 공사 데이터 분석 클래스"""
    
    def __init__(self, df: pd.DataFrame):
        """
        데이터 분석기 초기화
        
        Args:
            df: 공사 데이터 DataFrame
        """
        self.df = df
        self.current_date = datetime.now()
    
    def get_summary_stats(self) -> Dict[str, Any]:
        """전체 요약 통계"""
        try:
            total_projects = len(self.df)
            completed_projects = len(self.df[self.df['공사 종료'].notna()])
            in_progress_projects = len(self.df[
                (self.df['공사 시작'].notna()) & 
                (self.df['공사 종료'].isna())
            ])
            
            # 금액 관련 통계
            total_amount = self.df['총액 2'].sum()
            total_received = self.df['계약금'].fillna(0).sum() + \
                           self.df['중도금'].fillna(0).sum() + \
                           self.df['잔금'].fillna(0).sum()
            total_outstanding = self.df['미수금'].sum()
            
            # 평균 프로젝트 금액
            avg_project_amount = self.df['총액 2'].mean()
            
            return {
                'total_projects': total_projects,
                'completed_projects': completed_projects,
                'in_progress_projects': in_progress_projects,
                'pending_projects': total_projects - completed_projects - in_progress_projects,
                'total_amount': total_amount,
                'total_received': total_received,
                'total_outstanding': total_outstanding,
                'avg_project_amount': avg_project_amount,
                'collection_rate': (total_received / total_amount * 100) if total_amount > 0 else 0
            }
        except Exception as e:
            logger.error(f"요약 통계 계산 오류: {str(e)}")
            return {}
    
    def get_monthly_sales(self, year: Optional[int] = None) -> pd.DataFrame:
        """월별 매출 현황"""
        try:
            if year is None:
                year = self.current_date.year
            
            # 해당 연도 데이터 필터링
            df_year = self.df[
                self.df['공사 시작'].dt.year == year
            ].copy() if '공사 시작' in self.df.columns else self.df.copy()
            
            if df_year.empty or '공사 시작' not in df_year.columns:
                return pd.DataFrame()
            
            # 월별 그룹화
            df_year['월'] = df_year['공사 시작'].dt.month
            monthly_stats = df_year.groupby('월').agg({
                '총액 2': ['sum', 'mean', 'count'],
                '미수금': 'sum',
                '순익': 'sum'
            }).round(0)
            
            monthly_stats.columns = ['총매출', '평균금액', '건수', '미수금', '순익']
            monthly_stats = monthly_stats.reset_index()
            
            return monthly_stats
        except Exception as e:
            logger.error(f"월별 매출 계산 오류: {str(e)}")
            return pd.DataFrame()
    
    def get_regional_analysis(self) -> pd.DataFrame:
        """지역별 분석"""
        try:
            if '담당자' not in self.df.columns:
                return pd.DataFrame()
            
            regional_stats = self.df.groupby('담당자').agg({
                '총액 2': ['sum', 'mean', 'count'],
                '미수금': 'sum',
                '순익': 'sum',
                '마진율': 'mean'
            }).round(2)
            
            regional_stats.columns = ['총매출', '평균금액', '건수', '미수금', '총순익', '평균마진율']
            regional_stats = regional_stats.sort_values('총매출', ascending=False)
            regional_stats = regional_stats.reset_index()
            
            return regional_stats
        except Exception as e:
            logger.error(f"지역별 분석 오류: {str(e)}")
            return pd.DataFrame()
    
    def get_brand_analysis(self) -> pd.DataFrame:
        """브랜드별 분석"""
        try:
            if '브랜드' not in self.df.columns:
                return pd.DataFrame()
            
            brand_stats = self.df.groupby('브랜드').agg({
                '총액 2': ['sum', 'count'],
                '순익': 'sum',
                '마진율': 'mean'
            }).round(2)
            
            brand_stats.columns = ['총매출', '건수', '총순익', '평균마진율']
            brand_stats = brand_stats.sort_values('총매출', ascending=False)
            brand_stats = brand_stats.reset_index()
            
            return brand_stats
        except Exception as e:
            logger.error(f"브랜드별 분석 오류: {str(e)}")
            return pd.DataFrame()
    
    def get_outstanding_analysis(self) -> Dict[str, Any]:
        """미수금 분석"""
        try:
            # 미수금이 있는 프로젝트만 필터링
            outstanding_df = self.df[self.df['미수금'] > 0].copy()
            
            if outstanding_df.empty:
                return {'total_cases': 0, 'total_amount': 0, 'details': pd.DataFrame()}
            
            # 미수금 기간별 분류
            current_date = self.current_date
            outstanding_df['미수금_기간'] = outstanding_df['공사 종료'].apply(
                lambda x: self._classify_outstanding_period(x, current_date) 
                if pd.notna(x) else '진행중'
            )
            
            # 미수금 요약
            period_summary = outstanding_df.groupby('미수금_기간').agg({
                '미수금': ['sum', 'count'],
                '담당자': lambda x: list(set(x))
            }).round(0)
            
            period_summary.columns = ['미수금총액', '건수', '담당자목록']
            
            # 상위 미수금 프로젝트
            top_outstanding = outstanding_df.nlargest(10, '미수금')[
                ['프로젝트 코드', '현장 주소', '담당자', '미수금', '공사 종료', '고객 연락처']
            ].copy()
            
            return {
                'total_cases': len(outstanding_df),
                'total_amount': outstanding_df['미수금'].sum(),
                'period_summary': period_summary.reset_index(),
                'top_outstanding': top_outstanding,
                'by_person': outstanding_df.groupby('담당자')['미수금'].sum().sort_values(ascending=False).reset_index()
            }
        except Exception as e:
            logger.error(f"미수금 분석 오류: {str(e)}")
            return {'total_cases': 0, 'total_amount': 0, 'details': pd.DataFrame()}
    
    def _classify_outstanding_period(self, completion_date: datetime, current_date: datetime) -> str:
        """미수금 기간 분류"""
        if pd.isna(completion_date):
            return '진행중'
        
        days_diff = (current_date - completion_date).days
        
        if days_diff <= 30:
            return '30일 이내'
        elif days_diff <= 60:
            return '31-60일'
        elif days_diff <= 90:
            return '61-90일'
        else:
            return '90일 초과'
    
    def check_missing_data(self) -> Dict[str, Any]:
        """빈 칸 검증 및 누락 데이터 체크"""
        try:
            # 중요 필드 정의
            critical_fields = [
                '프로젝트 코드', '사업자', '담당자', '거래처', 
                '현장 주소', '공사 내용', '공사 시작', '총액 2',
                '현장 담당자', '담당자 연락처'
            ]
            
            # 존재하는 중요 필드만 선택
            available_critical_fields = [f for f in critical_fields if f in self.df.columns]
            
            missing_analysis = {}
            
            # 각 필드별 누락 현황
            for field in available_critical_fields:
                missing_count = self.df[field].isna().sum()
                missing_percentage = (missing_count / len(self.df)) * 100
                
                missing_analysis[field] = {
                    'missing_count': missing_count,
                    'missing_percentage': missing_percentage
                }
            
            # 영업사원별 누락 현황
            person_missing = {}
            if '담당자' in self.df.columns:
                for person in self.df['담당자'].dropna().unique():
                    person_df = self.df[self.df['담당자'] == person]
                    person_missing_count = 0
                    person_critical_missing = []
                    
                    for field in available_critical_fields:
                        field_missing = person_df[field].isna().sum()
                        if field_missing > 0:
                            person_missing_count += field_missing
                            person_critical_missing.append({
                                'field': field,
                                'missing_count': field_missing,
                                'projects': person_df[person_df[field].isna()]['프로젝트 코드'].tolist() 
                                          if '프로젝트 코드' in self.df.columns else []
                            })
                    
                    person_missing[person] = {
                        'total_missing': person_missing_count,
                        'critical_missing': person_critical_missing
                    }
            
            # 가장 누락이 많은 프로젝트들
            self.df['missing_count'] = self.df[available_critical_fields].isna().sum(axis=1)
            top_missing_projects = self.df.nlargest(20, 'missing_count')[
                ['프로젝트 코드', '담당자', '현장 주소', 'missing_count'] + 
                [f for f in available_critical_fields[:5] if f not in ['프로젝트 코드', '담당자', '현장 주소']]
            ].copy() if 'missing_count' in self.df.columns else pd.DataFrame()
            
            # numpy 타입을 Python 기본 타입으로 변환하는 함수
            def convert_numpy_types(obj):
                if hasattr(obj, 'item'):  # numpy scalar
                    return obj.item()
                elif hasattr(obj, 'tolist'):  # numpy array
                    return obj.tolist()
                elif isinstance(obj, dict):
                    return {k: convert_numpy_types(v) for k, v in obj.items()}
                elif isinstance(obj, list):
                    return [convert_numpy_types(item) for item in obj]
                else:
                    return obj
            
            result = {
                'field_analysis': convert_numpy_types(missing_analysis),
                'person_analysis': convert_numpy_types(person_missing),
                'top_missing_projects': top_missing_projects.to_dict('records') if not top_missing_projects.empty else [],
                'total_critical_fields': int(len(available_critical_fields)),
                'overall_missing_rate': float((
                    self.df[available_critical_fields].isna().sum().sum() / 
                    (len(self.df) * len(available_critical_fields)) * 100
                ) if available_critical_fields else 0)
            }
            
            return result
            
        except Exception as e:
            logger.error(f"누락 데이터 체크 오류: {str(e)}")
            return {}
    
    def get_completion_timeline(self) -> pd.DataFrame:
        """공사 완료 timeline 분석"""
        try:
            if '공사 시작' not in self.df.columns or '공사 종료' not in self.df.columns:
                return pd.DataFrame()
            
            # 완료된 프로젝트만
            completed_df = self.df[
                self.df['공사 시작'].notna() & 
                self.df['공사 종료'].notna()
            ].copy()
            
            if completed_df.empty:
                return pd.DataFrame()
            
            # 소요 기간 계산
            completed_df['소요일수'] = (
                completed_df['공사 종료'] - completed_df['공사 시작']
            ).dt.days
            
            # 월별 완료 건수
            completed_df['완료월'] = completed_df['공사 종료'].dt.to_period('M')
            timeline = completed_df.groupby('완료월').agg({
                '총액 2': 'sum',
                '소요일수': ['mean', 'count'],
                '순익': 'sum'
            }).round(1)
            
            timeline.columns = ['총매출', '평균소요일', '완료건수', '총순익']
            return timeline.reset_index()
            
        except Exception as e:
            logger.error(f"완료 timeline 분석 오류: {str(e)}")
            return pd.DataFrame()

def test_data_analyzer():
    """데이터 분석기 테스트"""
    # 샘플 데이터로 테스트
    from dashboard.utils.google_sheets import GoogleSheetsManager
    import os
    from dotenv import load_dotenv
    
    load_dotenv()
    
    try:
        # 로컬 엑셀 파일로 테스트
        df = pd.read_excel('data/아이티 공사 현황.xlsx', sheet_name='공사 현황')
        analyzer = DataAnalyzer(df)
        
        print("=== 전체 요약 통계 ===")
        summary = analyzer.get_summary_stats()
        for key, value in summary.items():
            print(f"{key}: {value}")
        
        print("\\n=== 지역별 분석 ===")
        regional = analyzer.get_regional_analysis()
        print(regional.head())
        
        print("\\n=== 미수금 분석 ===")
        outstanding = analyzer.get_outstanding_analysis()
        print(f"총 미수금 건수: {outstanding['total_cases']}")
        print(f"총 미수금 금액: {outstanding['total_amount']:,}원")
        
        print("\\n=== 누락 데이터 분석 ===")
        missing = analyzer.check_missing_data()
        print(f"전체 누락률: {missing['overall_missing_rate']:.1f}%")
        
    except Exception as e:
        print(f"테스트 오류: {str(e)}")

if __name__ == "__main__":
    test_data_analyzer()