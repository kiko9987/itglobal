#!/usr/bin/env python3
import os
import sys

# 현재 디렉토리를 Python path에 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)
sys.path.insert(0, os.path.join(current_dir, 'dashboard'))

# 환경 변수 설정
os.environ['PROJECT_ROOT'] = current_dir
os.environ['PYTHONPATH'] = current_dir
os.environ['FLASK_ENV'] = 'development'
os.environ['FLASK_DEBUG'] = '1'

try:
    print("Starting simple Flask server...")
    from dashboard import app

    if hasattr(app, 'app'):
        flask_app = app.app
    else:
        flask_app = app

    print("Flask app loaded successfully")
    flask_app.run(host='0.0.0.0', port=5000, debug=True, use_reloader=False)

except Exception as e:
    print(f"Error starting main app: {e}")
    print("Starting basic Flask server instead...")

    from flask import Flask
    basic_app = Flask(__name__)

    @basic_app.route('/')
    def hello():
        return '''
        <h1>IT Global Dashboard</h1>
        <p>Server is running but main app failed to load.</p>
        <p>Error: ''' + str(e) + '''</p>
        '''

    basic_app.run(host='0.0.0.0', port=5000, debug=True)