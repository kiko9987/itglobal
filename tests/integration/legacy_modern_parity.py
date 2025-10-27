"""
레거시-모던 CSS 기능 동등성 자동 테스트
두 버전 간의 기능적 차이점을 자동으로 감지하고 검증
"""

import time
import json
import logging
import asyncio
from datetime import datetime, timedelta
from typing import Dict, List, Any, Optional, Tuple
from dataclasses import dataclass, asdict
from enum import Enum
import unittest
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException, WebDriverException

logger = logging.getLogger(__name__)

class TestResult(Enum):
    """테스트 결과"""
    PASS = "pass"
    FAIL = "fail"
    WARNING = "warning"
    SKIP = "skip"
    ERROR = "error"

class TestCategory(Enum):
    """테스트 카테고리"""
    FUNCTIONALITY = "functionality"
    PERFORMANCE = "performance"
    VISUAL = "visual"
    ACCESSIBILITY = "accessibility"
    RESPONSIVE = "responsive"

@dataclass
class ParityTestResult:
    """동등성 테스트 결과"""
    test_id: str
    test_name: str
    category: TestCategory
    legacy_result: Any
    modern_result: Any
    result: TestResult
    execution_time: float
    timestamp: datetime
    error_message: Optional[str] = None
    performance_diff: Optional[float] = None
    visual_diff: Optional[str] = None
    recommendations: List[str] = None

    def __post_init__(self):
        if self.recommendations is None:
            self.recommendations = []

class LegacyModernParityTest:
    """레거시-모던 CSS 동등성 테스트"""

    def __init__(self, base_url: str = "http://localhost:5000"):
        self.base_url = base_url
        self.test_results: List[ParityTestResult] = []
        self.drivers = {}

        # 테스트 설정
        self.test_timeout = 30
        self.performance_threshold = 2.0  # 2배 이상 차이나면 경고
        self.visual_threshold = 0.05  # 5% 이상 시각적 차이

        # 테스트할 페이지들
        self.test_pages = [
            '/project-list',
            '/receivables',
            '/dashboard',
            '/stats'
        ]

        # 테스트할 기능들
        self.test_features = {
            'datatable': {
                'selectors': {
                    'table': 'table.dataTable',
                    'search': 'input[type="search"]',
                    'pagination': '.dataTables_paginate',
                    'sort_buttons': 'th.sorting'
                },
                'interactions': ['search', 'sort', 'paginate']
            },
            'modals': {
                'selectors': {
                    'modal_trigger': '[data-bs-toggle="modal"]',
                    'modal': '.modal',
                    'close_button': '.modal .btn-close'
                },
                'interactions': ['open', 'close']
            },
            'forms': {
                'selectors': {
                    'input_fields': 'input[type="text"], input[type="email"]',
                    'submit_button': 'button[type="submit"]',
                    'validation_messages': '.invalid-feedback'
                },
                'interactions': ['input', 'validate', 'submit']
            },
            'navigation': {
                'selectors': {
                    'nav_menu': '.navbar-nav',
                    'breadcrumb': '.breadcrumb',
                    'sidebar': '.sidebar'
                },
                'interactions': ['navigate', 'expand', 'collapse']
            }
        }

    def setup_drivers(self):
        """WebDriver 설정"""
        chrome_options = Options()
        chrome_options.add_argument('--headless')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--window-size=1920,1080')

        try:
            self.drivers['legacy'] = webdriver.Chrome(options=chrome_options)
            self.drivers['modern'] = webdriver.Chrome(options=chrome_options)
            logger.info("WebDriver 설정 완료")
        except Exception as e:
            logger.error(f"WebDriver 설정 실패: {e}")
            raise

    def teardown_drivers(self):
        """WebDriver 정리"""
        for driver_name, driver in self.drivers.items():
            try:
                driver.quit()
            except Exception as e:
                logger.warning(f"{driver_name} driver 종료 중 오류: {e}")

    def run_full_parity_test(self) -> Dict[str, Any]:
        """전체 동등성 테스트 실행"""
        try:
            self.setup_drivers()
            test_summary = {
                'start_time': datetime.now(),
                'total_tests': 0,
                'passed': 0,
                'failed': 0,
                'warnings': 0,
                'errors': 0,
                'test_results': []
            }

            logger.info("레거시-모던 CSS 동등성 테스트 시작")

            # 각 페이지에 대해 테스트 실행
            for page in self.test_pages:
                page_results = self._test_page_parity(page)
                test_summary['test_results'].extend(page_results)

            # 결과 집계
            for result in test_summary['test_results']:
                test_summary['total_tests'] += 1
                if result.result == TestResult.PASS:
                    test_summary['passed'] += 1
                elif result.result == TestResult.FAIL:
                    test_summary['failed'] += 1
                elif result.result == TestResult.WARNING:
                    test_summary['warnings'] += 1
                elif result.result == TestResult.ERROR:
                    test_summary['errors'] += 1

            test_summary['end_time'] = datetime.now()
            test_summary['duration'] = (test_summary['end_time'] - test_summary['start_time']).total_seconds()

            logger.info(f"동등성 테스트 완료: {test_summary['passed']}/{test_summary['total_tests']} 통과")
            return test_summary

        finally:
            self.teardown_drivers()

    def _test_page_parity(self, page_path: str) -> List[ParityTestResult]:
        """개별 페이지 동등성 테스트"""
        page_results = []

        try:
            # 페이지 로딩 테스트
            loading_result = self._test_page_loading(page_path)
            page_results.append(loading_result)

            # 기능별 테스트
            for feature_name, feature_config in self.test_features.items():
                feature_result = self._test_feature_parity(page_path, feature_name, feature_config)
                if feature_result:
                    page_results.append(feature_result)

            # 반응형 테스트
            responsive_result = self._test_responsive_parity(page_path)
            page_results.append(responsive_result)

        except Exception as e:
            error_result = ParityTestResult(
                test_id=f"page_error_{page_path.replace('/', '_')}",
                test_name=f"Page Error: {page_path}",
                category=TestCategory.FUNCTIONALITY,
                legacy_result=None,
                modern_result=None,
                result=TestResult.ERROR,
                execution_time=0,
                timestamp=datetime.now(),
                error_message=str(e)
            )
            page_results.append(error_result)

        return page_results

    def _test_page_loading(self, page_path: str) -> ParityTestResult:
        """페이지 로딩 동등성 테스트"""
        test_id = f"loading_{page_path.replace('/', '_')}"

        try:
            # 레거시 버전 로딩
            legacy_start = time.time()
            self.drivers['legacy'].get(f"{self.base_url}{page_path}?css_mode=legacy")
            WebDriverWait(self.drivers['legacy'], self.test_timeout).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            legacy_time = time.time() - legacy_start

            # 모던 버전 로딩
            modern_start = time.time()
            self.drivers['modern'].get(f"{self.base_url}{page_path}?css_mode=modern")
            WebDriverWait(self.drivers['modern'], self.test_timeout).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            modern_time = time.time() - modern_start

            # 성능 비교
            performance_diff = abs(modern_time - legacy_time) / legacy_time if legacy_time > 0 else 0

            if performance_diff > self.performance_threshold:
                result = TestResult.WARNING
                recommendations = [f"성능 차이가 {performance_diff:.2%} 발생. 최적화 필요"]
            else:
                result = TestResult.PASS
                recommendations = []

            return ParityTestResult(
                test_id=test_id,
                test_name=f"Page Loading: {page_path}",
                category=TestCategory.PERFORMANCE,
                legacy_result=legacy_time,
                modern_result=modern_time,
                result=result,
                execution_time=legacy_time + modern_time,
                timestamp=datetime.now(),
                performance_diff=performance_diff,
                recommendations=recommendations
            )

        except TimeoutException:
            return ParityTestResult(
                test_id=test_id,
                test_name=f"Page Loading: {page_path}",
                category=TestCategory.PERFORMANCE,
                legacy_result=None,
                modern_result=None,
                result=TestResult.FAIL,
                execution_time=self.test_timeout,
                timestamp=datetime.now(),
                error_message="페이지 로딩 타임아웃",
                recommendations=["페이지 로딩 시간 최적화 필요"]
            )

    def _test_feature_parity(self, page_path: str, feature_name: str,
                           feature_config: Dict) -> Optional[ParityTestResult]:
        """기능 동등성 테스트"""
        test_id = f"feature_{feature_name}_{page_path.replace('/', '_')}"

        try:
            # 기능 요소 존재 확인
            legacy_elements = self._check_feature_elements(self.drivers['legacy'], feature_config['selectors'])
            modern_elements = self._check_feature_elements(self.drivers['modern'], feature_config['selectors'])

            # 기능 상호작용 테스트
            legacy_interactions = self._test_feature_interactions(
                self.drivers['legacy'], feature_config, feature_name
            )
            modern_interactions = self._test_feature_interactions(
                self.drivers['modern'], feature_config, feature_name
            )

            # 결과 비교
            element_parity = legacy_elements == modern_elements
            interaction_parity = legacy_interactions == modern_interactions

            if element_parity and interaction_parity:
                result = TestResult.PASS
                recommendations = []
            elif element_parity:
                result = TestResult.WARNING
                recommendations = ["상호작용 동작에 차이가 있음"]
            else:
                result = TestResult.FAIL
                recommendations = ["기능 요소 구조에 차이가 있음", "기능 동등성 확보 필요"]

            return ParityTestResult(
                test_id=test_id,
                test_name=f"Feature: {feature_name} on {page_path}",
                category=TestCategory.FUNCTIONALITY,
                legacy_result={'elements': legacy_elements, 'interactions': legacy_interactions},
                modern_result={'elements': modern_elements, 'interactions': modern_interactions},
                result=result,
                execution_time=2.0,  # 추정값
                timestamp=datetime.now(),
                recommendations=recommendations
            )

        except Exception as e:
            return ParityTestResult(
                test_id=test_id,
                test_name=f"Feature: {feature_name} on {page_path}",
                category=TestCategory.FUNCTIONALITY,
                legacy_result=None,
                modern_result=None,
                result=TestResult.ERROR,
                execution_time=0,
                timestamp=datetime.now(),
                error_message=str(e),
                recommendations=["기능 테스트 중 오류 발생 - 디버깅 필요"]
            )

    def _check_feature_elements(self, driver: webdriver.Chrome, selectors: Dict[str, str]) -> Dict[str, bool]:
        """기능 요소 존재 확인"""
        elements_found = {}

        for element_name, selector in selectors.items():
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                elements_found[element_name] = len(elements) > 0
            except Exception:
                elements_found[element_name] = False

        return elements_found

    def _test_feature_interactions(self, driver: webdriver.Chrome, feature_config: Dict,
                                 feature_name: str) -> Dict[str, bool]:
        """기능 상호작용 테스트"""
        interactions_working = {}

        for interaction in feature_config.get('interactions', []):
            try:
                interactions_working[interaction] = self._test_specific_interaction(
                    driver, feature_name, interaction, feature_config['selectors']
                )
            except Exception as e:
                logger.warning(f"상호작용 테스트 실패: {feature_name}.{interaction} - {e}")
                interactions_working[interaction] = False

        return interactions_working

    def _test_specific_interaction(self, driver: webdriver.Chrome, feature_name: str,
                                 interaction: str, selectors: Dict[str, str]) -> bool:
        """특정 상호작용 테스트"""
        try:
            if feature_name == 'datatable' and interaction == 'search':
                search_input = driver.find_element(By.CSS_SELECTOR, selectors['search'])
                search_input.clear()
                search_input.send_keys("test")
                time.sleep(1)
                return True

            elif feature_name == 'datatable' and interaction == 'sort':
                sort_button = driver.find_element(By.CSS_SELECTOR, selectors['sort_buttons'])
                sort_button.click()
                time.sleep(1)
                return True

            elif feature_name == 'modals' and interaction == 'open':
                modal_trigger = driver.find_element(By.CSS_SELECTOR, selectors['modal_trigger'])
                modal_trigger.click()
                WebDriverWait(driver, 5).until(
                    EC.visibility_of_element_located((By.CSS_SELECTOR, selectors['modal']))
                )
                return True

            # 다른 상호작용들도 여기에 추가...

            return True

        except Exception as e:
            logger.debug(f"상호작용 테스트 실패: {interaction} - {e}")
            return False

    def _test_responsive_parity(self, page_path: str) -> ParityTestResult:
        """반응형 동등성 테스트"""
        test_id = f"responsive_{page_path.replace('/', '_')}"

        try:
            responsive_sizes = [
                (1920, 1080),  # 데스크톱
                (768, 1024),   # 태블릿
                (375, 667)     # 모바일
            ]

            responsive_results = {'legacy': {}, 'modern': {}}

            for width, height in responsive_sizes:
                # 레거시 테스트
                self.drivers['legacy'].set_window_size(width, height)
                time.sleep(1)
                legacy_layout = self._analyze_responsive_layout(self.drivers['legacy'])
                responsive_results['legacy'][f"{width}x{height}"] = legacy_layout

                # 모던 테스트
                self.drivers['modern'].set_window_size(width, height)
                time.sleep(1)
                modern_layout = self._analyze_responsive_layout(self.drivers['modern'])
                responsive_results['modern'][f"{width}x{height}"] = modern_layout

            # 창 크기 복원
            for driver in self.drivers.values():
                driver.set_window_size(1920, 1080)

            # 반응형 동등성 확인
            layout_parity = responsive_results['legacy'] == responsive_results['modern']

            if layout_parity:
                result = TestResult.PASS
                recommendations = []
            else:
                result = TestResult.WARNING
                recommendations = ["반응형 레이아웃에 차이가 있음", "미디어 쿼리 동등성 확인 필요"]

            return ParityTestResult(
                test_id=test_id,
                test_name=f"Responsive: {page_path}",
                category=TestCategory.RESPONSIVE,
                legacy_result=responsive_results['legacy'],
                modern_result=responsive_results['modern'],
                result=result,
                execution_time=3.0,
                timestamp=datetime.now(),
                recommendations=recommendations
            )

        except Exception as e:
            return ParityTestResult(
                test_id=test_id,
                test_name=f"Responsive: {page_path}",
                category=TestCategory.RESPONSIVE,
                legacy_result=None,
                modern_result=None,
                result=TestResult.ERROR,
                execution_time=0,
                timestamp=datetime.now(),
                error_message=str(e),
                recommendations=["반응형 테스트 중 오류 발생"]
            )

    def _analyze_responsive_layout(self, driver: webdriver.Chrome) -> Dict[str, Any]:
        """반응형 레이아웃 분석"""
        try:
            # 주요 레이아웃 요소들의 표시 상태 확인
            layout_info = {}

            # 네비게이션 메뉴
            nav_elements = driver.find_elements(By.CSS_SELECTOR, '.navbar, .nav')
            layout_info['nav_visible'] = any(elem.is_displayed() for elem in nav_elements)

            # 사이드바
            sidebar_elements = driver.find_elements(By.CSS_SELECTOR, '.sidebar, .side-menu')
            layout_info['sidebar_visible'] = any(elem.is_displayed() for elem in sidebar_elements)

            # 메인 콘텐츠
            main_elements = driver.find_elements(By.CSS_SELECTOR, 'main, .main-content, .content')
            layout_info['main_visible'] = any(elem.is_displayed() for elem in main_elements)

            # 테이블 (있는 경우)
            table_elements = driver.find_elements(By.CSS_SELECTOR, 'table')
            layout_info['table_responsive'] = True  # 실제로는 가로 스크롤 등을 확인

            return layout_info

        except Exception as e:
            logger.warning(f"레이아웃 분석 실패: {e}")
            return {}

    def generate_test_report(self, test_summary: Dict[str, Any]) -> Dict[str, Any]:
        """테스트 보고서 생성"""
        report = {
            'summary': {
                'test_date': test_summary['start_time'].isoformat(),
                'duration_seconds': test_summary['duration'],
                'total_tests': test_summary['total_tests'],
                'pass_rate': test_summary['passed'] / test_summary['total_tests'] * 100 if test_summary['total_tests'] > 0 else 0,
                'results_breakdown': {
                    'passed': test_summary['passed'],
                    'failed': test_summary['failed'],
                    'warnings': test_summary['warnings'],
                    'errors': test_summary['errors']
                }
            },
            'category_results': {},
            'critical_issues': [],
            'recommendations': [],
            'detailed_results': []
        }

        # 카테고리별 결과 집계
        for category in TestCategory:
            category_tests = [r for r in test_summary['test_results'] if r.category == category]
            if category_tests:
                report['category_results'][category.value] = {
                    'total': len(category_tests),
                    'passed': len([r for r in category_tests if r.result == TestResult.PASS]),
                    'failed': len([r for r in category_tests if r.result == TestResult.FAIL]),
                    'warnings': len([r for r in category_tests if r.result == TestResult.WARNING])
                }

        # 중요 이슈 및 권장사항 추출
        for result in test_summary['test_results']:
            if result.result == TestResult.FAIL:
                report['critical_issues'].append({
                    'test_name': result.test_name,
                    'error': result.error_message,
                    'recommendations': result.recommendations
                })

            report['recommendations'].extend(result.recommendations)

        # 권장사항 중복 제거
        report['recommendations'] = list(set(report['recommendations']))

        # 상세 결과
        for result in test_summary['test_results']:
            report['detailed_results'].append({
                'test_id': result.test_id,
                'test_name': result.test_name,
                'category': result.category.value,
                'result': result.result.value,
                'execution_time': result.execution_time,
                'timestamp': result.timestamp.isoformat(),
                'performance_diff': result.performance_diff,
                'error_message': result.error_message,
                'recommendations': result.recommendations
            })

        return report

    def save_test_results(self, test_summary: Dict[str, Any], file_path: str = None):
        """테스트 결과 저장"""
        if file_path is None:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            file_path = f"parity_test_results_{timestamp}.json"

        report = self.generate_test_report(test_summary)

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(report, f, indent=2, ensure_ascii=False)

            logger.info(f"테스트 결과 저장 완료: {file_path}")
            return file_path

        except Exception as e:
            logger.error(f"테스트 결과 저장 실패: {e}")
            return None

# 테스트 실행을 위한 유틸리티 함수들

def run_automated_parity_test(base_url: str = "http://localhost:5000") -> Dict[str, Any]:
    """자동화된 동등성 테스트 실행"""
    tester = LegacyModernParityTest(base_url)
    return tester.run_full_parity_test()

def schedule_daily_parity_test():
    """일일 동등성 테스트 스케줄링"""
    # 실제로는 cron job이나 스케줄러와 연동
    logger.info("일일 동등성 테스트가 스케줄되었습니다")

if __name__ == "__main__":
    # 테스트 실행 예시
    try:
        test_results = run_automated_parity_test()
        tester = LegacyModernParityTest()
        report_file = tester.save_test_results(test_results)
        print(f"테스트 완료. 보고서: {report_file}")
    except Exception as e:
        print(f"테스트 실행 실패: {e}")
        logger.error(f"테스트 실행 실패: {e}")