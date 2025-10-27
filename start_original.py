#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys

# UTF-8 인코딩 설정
os.environ['PYTHONIOENCODING'] = 'utf-8'

# 경로 설정
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)
sys.path.insert(0, os.path.join(current_dir, 'dashboard'))

# 환경 변수 설정
os.environ['PROJECT_ROOT'] = current_dir
os.environ['PYTHONPATH'] = current_dir
os.environ['FLASK_ENV'] = 'development'
os.environ['PORT'] = '5004'

print("=== IT Global Dashboard (원본) 시작 ===")

try:
    # 원본 dashboard 앱 import
    from dashboard.app import app, socketio

    print("대시보드 앱과 SocketIO 로드 완료!")
    print("서버 주소: http://localhost:5004")
    print("SocketIO 지원: 활성화")

    # SocketIO와 함께 실행 (어제와 동일한 방식)
    socketio.run(app, debug=True, host='0.0.0.0', port=5004, use_reloader=False)

except Exception as e:
    print(f"원본 앱 실행 오류: {e}")

    # 간단한 대체 서버
    from flask import Flask
    simple_app = Flask(__name__)

    @simple_app.route('/')
    def index():
        return f'''
        <h1>IT Global Dashboard</h1>
        <p>원본 앱 로드 실패</p>
        <p>오류: {str(e)}</p>
        <p><a href="http://localhost:5003">대체 서버로 이동</a></p>
        '''

    simple_app.run(host='0.0.0.0', port=5004, debug=True)