#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import locale

# UTF-8 인코딩 강제 설정
try:
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.detach())
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.detach())
except:
    pass

# 환경 변수로 인코딩 설정
os.environ['PYTHONIOENCODING'] = 'utf-8'
os.environ['LANG'] = 'en_US.UTF-8'

# 경로 설정
current_dir = os.path.dirname(os.path.abspath(__file__))
dashboard_dir = os.path.join(current_dir, 'dashboard')

sys.path.insert(0, current_dir)
sys.path.insert(0, dashboard_dir)

# 환경 변수 설정
os.environ['PROJECT_ROOT'] = current_dir
os.environ['PYTHONPATH'] = current_dir
os.environ['FLASK_ENV'] = 'development'
os.environ['PORT'] = '5002'

print("=== IT Global Dashboard 시작 ===")

try:
    # 대시보드 앱 import 및 실행
    os.chdir(current_dir)

    # dashboard.app 모듈 import
    from dashboard.app import app

    print("대시보드 앱 로드 완료!")
    print("서버 주소: http://localhost:5002")
    print("로그인 후 프로젝트 관리 페이지로 이동 가능합니다.")

    # Flask 앱 실행 (socketio 대신 기본 Flask 사용)
    app.run(host='0.0.0.0', port=5002, debug=True, use_reloader=False)

except ImportError as e:
    print(f"모듈 import 오류: {e}")
    print("기본 대시보드로 실행합니다...")

    from flask import Flask, render_template_string, redirect, url_for, session

    app = Flask(__name__)
    app.secret_key = 'dev-secret-key'

    @app.route('/')
    def index():
        if 'user' in session:
            return redirect('/projects')
        return render_template_string("""
        <!DOCTYPE html>
        <html>
        <head>
            <title>IT Global Dashboard</title>
            <meta charset="utf-8">
            <style>
                body { font-family: Arial, sans-serif; margin: 40px; background: #f5f5f5; }
                .container { max-width: 600px; margin: 0 auto; background: white; padding: 40px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                .btn { padding: 12px 24px; margin: 10px 0; background: #007bff; color: white; text-decoration: none; border-radius: 4px; display: inline-block; border: none; cursor: pointer; }
                .btn:hover { background: #0056b3; }
                .login-form { margin-top: 20px; }
                input { padding: 10px; margin: 5px 0; width: 100%; box-sizing: border-box; border: 1px solid #ddd; border-radius: 4px; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>IT Global Dashboard</h1>
                <p>프로젝트 관리 시스템에 오신 것을 환영합니다.</p>

                <div class="login-form">
                    <h3>로그인</h3>
                    <form method="POST" action="/login">
                        <input type="email" name="email" placeholder="이메일 주소" required>
                        <input type="password" name="password" placeholder="비밀번호" required>
                        <button type="submit" class="btn">로그인</button>
                    </form>
                </div>

                <p><small>Demo: admin@test.com / admin</small></p>
            </div>
        </body>
        </html>
        """)

    @app.route('/login', methods=['POST'])
    def login():
        email = request.form.get('email', '')
        password = request.form.get('password', '')

        # 간단한 인증 (실제로는 데이터베이스 확인)
        if email == 'admin@test.com' and password == 'admin':
            session['user'] = email
            return redirect('/projects')

        return redirect('/')

    @app.route('/projects')
    def projects():
        if 'user' not in session:
            return redirect('/')

        return render_template_string("""
        <!DOCTYPE html>
        <html>
        <head>
            <title>프로젝트 관리 - IT Global</title>
            <meta charset="utf-8">
            <style>
                body { font-family: Arial, sans-serif; margin: 0; background: #f5f5f5; }
                .header { background: #007bff; color: white; padding: 20px; }
                .container { max-width: 1200px; margin: 20px auto; padding: 20px; }
                .btn { padding: 10px 20px; margin: 5px; background: #007bff; color: white; text-decoration: none; border-radius: 4px; }
                .btn.secondary { background: #6c757d; }
                .project-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 20px; margin-top: 20px; }
                .project-card { background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
            </style>
        </head>
        <body>
            <div class="header">
                <h1>프로젝트 관리</h1>
                <p>사용자: {{ session.user }}</p>
                <a href="/logout" class="btn secondary">로그아웃</a>
            </div>

            <div class="container">
                <div style="margin-bottom: 20px;">
                    <a href="/projects/new" class="btn">새 프로젝트</a>
                    <a href="/api/projects/list" class="btn secondary">API 테스트</a>
                </div>

                <div class="project-grid">
                    <div class="project-card">
                        <h3>프로젝트 A</h3>
                        <p>상태: 진행중</p>
                        <p>금액: 1,000,000원</p>
                    </div>
                    <div class="project-card">
                        <h3>프로젝트 B</h3>
                        <p>상태: 완료</p>
                        <p>금액: 2,500,000원</p>
                    </div>
                </div>
            </div>
        </body>
        </html>
        """)

    @app.route('/logout')
    def logout():
        session.pop('user', None)
        return redirect('/')

    @app.route('/api/projects/list')
    def api_projects():
        return {
            "status": "success",
            "projects": [
                {"id": 1, "name": "프로젝트 A", "status": "진행중", "amount": 1000000},
                {"id": 2, "name": "프로젝트 B", "status": "완료", "amount": 2500000}
            ]
        }

    from flask import request
    app.run(host='0.0.0.0', port=5002, debug=True, use_reloader=False)

except Exception as e:
    print(f"오류 발생: {e}")
    print("기본 서버로 실행합니다...")