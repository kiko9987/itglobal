#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os
import sys

# UTF-8 인코딩 설정
os.environ['PYTHONIOENCODING'] = 'utf-8'

# 현재 디렉토리를 Python path에 추가
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)
sys.path.insert(0, os.path.join(current_dir, 'dashboard'))

# 환경 변수 설정
os.environ['PROJECT_ROOT'] = current_dir
os.environ['PYTHONPATH'] = current_dir
os.environ['FLASK_ENV'] = 'development'
os.environ['PORT'] = '5000'

print("=== ITG Dashboard Starting ===")
print("Loading dashboard application...")

try:
    # 기본 Flask 앱으로 시작
    from flask import Flask, render_template_string

    app = Flask(__name__)
    app.config['SECRET_KEY'] = 'dev-secret-key'

    @app.route('/')
    def index():
        return render_template_string("""
        <!DOCTYPE html>
        <html>
        <head>
            <title>IT Global Dashboard</title>
            <meta charset="utf-8">
            <style>
                body { font-family: Arial, sans-serif; margin: 40px; }
                .container { max-width: 800px; margin: 0 auto; }
                .btn { padding: 10px 20px; margin: 5px; background: #007bff; color: white; text-decoration: none; border-radius: 4px; }
                .btn:hover { background: #0056b3; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>IT Global Dashboard</h1>
                <p>Server is running successfully!</p>
                <p>Available endpoints:</p>
                <ul>
                    <li><a href="/projects" class="btn">Projects</a></li>
                    <li><a href="/api/projects/list" class="btn">API Test</a></li>
                    <li><a href="/login" class="btn">Login</a></li>
                </ul>
                <p><small>Server running on port 5000</small></p>
            </div>
        </body>
        </html>
        """)

    @app.route('/projects')
    def projects():
        return "<h1>Projects Page</h1><p>Project management coming soon...</p>"

    @app.route('/api/projects/list')
    def api_test():
        return {"status": "ok", "message": "API is working", "projects": []}

    @app.route('/login')
    def login():
        return "<h1>Login Page</h1><p>Authentication coming soon...</p>"

    print("Basic Flask server ready!")
    print("Access the dashboard at: http://localhost:5000")

    app.run(host='0.0.0.0', port=5000, debug=True, use_reloader=False)

except Exception as e:
    print(f"Error: {e}")
    print("Please check the application setup.")