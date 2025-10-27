#!/usr/bin/env python3
"""
개발 서버 시작 스크립트 - 포트 5001에서 실행
"""
import os
import sys
import io

# Windows 콘솔 UTF-8 출력 설정
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# 환경 변수 설정 (.env 파일의 PORT를 사용하거나 기본값 5000)
if 'PORT' not in os.environ:
    os.environ['PORT'] = '5000'
os.environ['FLASK_ENV'] = 'development'
os.environ['FLASK_DEBUG'] = '1'

# 현재 디렉토리를 Python 경로에 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

# 대시보드 앱 실행
if __name__ == '__main__':
    print(f"개발 서버 시작 - 포트: {os.environ['PORT']}")
    from dashboard import app
    # app.py에서 직접 가져오기 (UTF-8 인코딩 명시)
    exec(open(os.path.join(current_dir, 'dashboard', 'app.py'), encoding='utf-8').read())