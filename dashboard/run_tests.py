#!/usr/bin/env python3
"""
테스트 실행 스크립트
"""

import subprocess
import sys
import os

def run_tests():
    """모든 테스트 실행"""
    print("테스트 시작...")
    
    # 현재 디렉토리를 dashboard로 변경
    os.chdir(os.path.dirname(__file__))
    
    try:
        # pytest 실행
        result = subprocess.run([
            sys.executable, '-m', 'pytest',
            'tests/',
            '-v',
            '--tb=short',
            '--cov=utils',
            '--cov-report=term-missing',
            '--cov-report=html:htmlcov'
        ], check=False)
        
        if result.returncode == 0:
            print("SUCCESS: 모든 테스트 통과!")
            print("COVERAGE: 커버리지 보고서 htmlcov/index.html")
            return True
        else:
            print("FAILED: 테스트 실패")
            return False
            
    except FileNotFoundError:
        print("ERROR: pytest가 설치되어 있지 않습니다.")
        print("pip install pytest pytest-cov 를 실행하세요.")
        return False
    except Exception as e:
        print(f"ERROR: 테스트 실행 오류: {str(e)}")
        return False

def run_security_scan():
    """보안 스캔 실행"""
    print("SECURITY: 보안 스캔 시작...")
    
    try:
        # bandit로 보안 취약점 스캔
        result = subprocess.run([
            sys.executable, '-m', 'bandit',
            '-r', 'utils/',
            '-f', 'json',
            '-o', 'security_report.json'
        ], check=False, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("SUCCESS: 보안 스캔 완료 - 문제 없음")
        else:
            print("WARNING: 보안 취약점 발견")
            print("REPORT: security_report.json")
            
    except FileNotFoundError:
        print("WARNING: bandit가 설치되어 있지 않습니다.")
        print("pip install bandit 를 실행하세요.")

def run_code_quality_check():
    """코드 품질 검사"""
    print("QUALITY: 코드 품질 검사...")
    
    try:
        # flake8으로 코딩 스타일 검사
        result = subprocess.run([
            sys.executable, '-m', 'flake8',
            'utils/',
            '--max-line-length=100',
            '--ignore=E203,W503'
        ], check=False)
        
        if result.returncode == 0:
            print("SUCCESS: 코딩 스타일 검사 통과")
        else:
            print("WARNING: 코딩 스타일 개선 필요")
            
    except FileNotFoundError:
        print("WARNING: flake8이 설치되어 있지 않습니다.")

if __name__ == "__main__":
    success = True
    
    # 테스트 실행
    if not run_tests():
        success = False
    
    print("\n" + "="*50)
    
    # 보안 스캔
    run_security_scan()
    
    print("\n" + "="*50)
    
    # 코드 품질 검사
    run_code_quality_check()
    
    print("\n" + "="*50)
    
    if success:
        print("SUCCESS: 모든 검사 완료!")
        sys.exit(0)
    else:
        print("FAILED: 일부 검사 실패")
        sys.exit(1)