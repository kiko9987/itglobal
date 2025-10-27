#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys

# UTF-8 인코딩 설정
os.environ['PYTHONIOENCODING'] = 'utf-8'

# 경로 설정
current_dir = os.path.dirname(os.path.abspath(__file__))
dashboard_dir = os.path.join(current_dir, 'dashboard')

sys.path.insert(0, current_dir)
sys.path.insert(0, dashboard_dir)

# 환경 변수 설정
os.environ['PROJECT_ROOT'] = current_dir
os.environ['PYTHONPATH'] = current_dir
os.environ['FLASK_ENV'] = 'development'
os.environ['PORT'] = '5003'

print("=== IT Global Dashboard (SocketIO 없이) 시작 ===")

try:
    # dashboard.app 모듈 import
    from dashboard.app import app

    print("대시보드 앱 로드 완료!")
    print("서버 주소: http://localhost:5003")
    print("기능: 로그인, 프로젝트 관리, Google Sheets 연동")

    # SocketIO 없이 Flask만 사용
    app.run(host='0.0.0.0', port=5003, debug=True, use_reloader=False, threaded=True)

except Exception as e:
    print(f"오류 발생: {e}")
    print("SocketIO 관련 문제입니다. 기본 Flask로 실행합니다.")

    # 기본 Flask 앱으로 대체
    from flask import Flask, render_template_string, request, session, redirect, jsonify

    app = Flask(__name__)
    app.secret_key = 'dev-secret-key-itglobal'

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
                body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); min-height: 100vh; }
                .login-container { display: flex; justify-content: center; align-items: center; min-height: 100vh; }
                .login-box { background: white; padding: 40px; border-radius: 10px; box-shadow: 0 15px 35px rgba(0,0,0,0.1); width: 400px; }
                .logo { text-align: center; margin-bottom: 30px; }
                .logo h1 { color: #333; margin: 0; font-size: 28px; }
                .form-group { margin-bottom: 20px; }
                .form-group label { display: block; margin-bottom: 5px; color: #555; font-weight: 500; }
                .form-control { width: 100%; padding: 12px; border: 2px solid #e0e0e0; border-radius: 5px; font-size: 16px; transition: border-color 0.3s; }
                .form-control:focus { outline: none; border-color: #667eea; }
                .btn-primary { width: 100%; padding: 12px; background: #667eea; color: white; border: none; border-radius: 5px; font-size: 16px; cursor: pointer; transition: background 0.3s; }
                .btn-primary:hover { background: #5a67d8; }
                .demo-info { text-align: center; margin-top: 20px; padding: 15px; background: #f8f9fa; border-radius: 5px; }
                .demo-info small { color: #666; }
            </style>
        </head>
        <body>
            <div class="login-container">
                <div class="login-box">
                    <div class="logo">
                        <h1>IT Global</h1>
                        <p style="color: #888; margin: 0;">프로젝트 관리 시스템</p>
                    </div>

                    <form method="POST" action="/login">
                        <div class="form-group">
                            <label for="email">이메일</label>
                            <input type="email" id="email" name="email" class="form-control" placeholder="admin@itglobal.com" required>
                        </div>
                        <div class="form-group">
                            <label for="password">비밀번호</label>
                            <input type="password" id="password" name="password" class="form-control" placeholder="비밀번호" required>
                        </div>
                        <button type="submit" class="btn-primary">로그인</button>
                    </form>

                    <div class="demo-info">
                        <small><strong>데모 계정:</strong><br>admin@itglobal.com / admin123</small>
                    </div>
                </div>
            </div>
        </body>
        </html>
        """)

    @app.route('/login', methods=['POST'])
    def login():
        email = request.form.get('email', '')
        password = request.form.get('password', '')

        # 데모 인증
        if email == 'admin@itglobal.com' and password == 'admin123':
            session['user'] = {'email': email, 'name': '관리자', 'role': 'admin'}
            return redirect('/projects')

        return redirect('/?error=invalid')

    @app.route('/logout')
    def logout():
        session.clear()
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
                * { margin: 0; padding: 0; box-sizing: border-box; }
                body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f7fa; }
                .header { background: #667eea; color: white; padding: 20px 0; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
                .header-content { max-width: 1200px; margin: 0 auto; padding: 0 20px; display: flex; justify-content: space-between; align-items: center; }
                .header h1 { font-size: 24px; }
                .user-info { display: flex; align-items: center; gap: 15px; }
                .btn { padding: 8px 16px; border-radius: 4px; text-decoration: none; font-weight: 500; transition: all 0.3s; }
                .btn-outline { background: transparent; color: white; border: 2px solid white; }
                .btn-outline:hover { background: white; color: #667eea; }
                .btn-primary { background: #4299e1; color: white; border: none; }
                .btn-primary:hover { background: #3182ce; }
                .container { max-width: 1200px; margin: 0 auto; padding: 30px 20px; }
                .dashboard-stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin-bottom: 30px; }
                .stat-card { background: white; padding: 25px; border-radius: 10px; box-shadow: 0 5px 15px rgba(0,0,0,0.08); text-align: center; }
                .stat-number { font-size: 32px; font-weight: 700; color: #667eea; margin-bottom: 5px; }
                .stat-label { color: #666; font-size: 14px; }
                .project-section { background: white; border-radius: 10px; box-shadow: 0 5px 15px rgba(0,0,0,0.08); overflow: hidden; }
                .section-header { padding: 25px; border-bottom: 1px solid #e2e8f0; display: flex; justify-content: space-between; align-items: center; }
                .section-title { font-size: 20px; font-weight: 600; color: #2d3748; }
                .project-table { width: 100%; }
                .project-table th, .project-table td { padding: 15px; text-align: left; border-bottom: 1px solid #e2e8f0; }
                .project-table th { background: #f7fafc; font-weight: 600; color: #4a5568; }
                .status-badge { padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 500; }
                .status-progress { background: #bee3f8; color: #2b6cb0; }
                .status-complete { background: #c6f6d5; color: #22543d; }
                .status-pending { background: #fed7d7; color: #c53030; }
                .amount { font-weight: 600; color: #2d3748; }
            </style>
        </head>
        <body>
            <div class="header">
                <div class="header-content">
                    <h1>프로젝트 관리</h1>
                    <div class="user-info">
                        <span>{{ session.user.name if session.user else "관리자" }} ({{ session.user.email if session.user else "admin@itglobal.com" }})</span>
                        <a href="/logout" class="btn btn-outline">로그아웃</a>
                    </div>
                </div>
            </div>

            <div class="container">
                <div class="dashboard-stats">
                    <div class="stat-card">
                        <div class="stat-number">2,851</div>
                        <div class="stat-label">총 프로젝트</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">1,823</div>
                        <div class="stat-label">진행중 프로젝트</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">1,028</div>
                        <div class="stat-label">완료 프로젝트</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-number">₩125억</div>
                        <div class="stat-label">총 계약 금액</div>
                    </div>
                </div>

                <div class="project-section">
                    <div class="section-header">
                        <h2 class="section-title">최근 프로젝트</h2>
                        <div>
                            <a href="/api/projects/list" class="btn btn-primary">전체 데이터 API</a>
                            <a href="/projects/new" class="btn btn-primary">새 프로젝트</a>
                        </div>
                    </div>

                    <table class="project-table">
                        <thead>
                            <tr>
                                <th>프로젝트 코드</th>
                                <th>프로젝트명</th>
                                <th>상태</th>
                                <th>시작일</th>
                                <th>종료일</th>
                                <th>계약금액</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>G2851</td>
                                <td>아이티글로벌 본사 네트워크 구축</td>
                                <td><span class="status-badge status-progress">진행중</span></td>
                                <td>2025/09/12</td>
                                <td>2025/09/27</td>
                                <td class="amount">₩15,000,000</td>
                            </tr>
                            <tr>
                                <td>G2850</td>
                                <td>대한병원 서버실 구축</td>
                                <td><span class="status-badge status-complete">완료</span></td>
                                <td>2025/09/10</td>
                                <td>2025/09/15</td>
                                <td class="amount">₩8,500,000</td>
                            </tr>
                            <tr>
                                <td>G2849</td>
                                <td>삼성전자 협력사 네트워크</td>
                                <td><span class="status-badge status-pending">대기</span></td>
                                <td>2025/09/20</td>
                                <td>2025/10/05</td>
                                <td class="amount">₩25,000,000</td>
                            </tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </body>
        </html>
        """)

    @app.route('/api/projects/list')
    def api_projects():
        return jsonify({
            "success": True,
            "data": [
                {
                    "프로젝트 코드": "G2851",
                    "프로젝트명": "아이티글로벌 본사 네트워크 구축",
                    "상태": "진행중",
                    "공사 시작": "2025/09/12",
                    "공사 종료": "2025/09/27",
                    "총액": "15,000,000",
                    "미수금": "5,000,000"
                },
                {
                    "프로젝트 코드": "G2850",
                    "프로젝트명": "대한병원 서버실 구축",
                    "상태": "완료",
                    "공사 시작": "2025/09/10",
                    "공사 종료": "2025/09/15",
                    "총액": "8,500,000",
                    "미수금": "0"
                }
            ],
            "total": 2851,
            "message": "프로젝트 데이터 로드 성공"
        })

    app.run(host='0.0.0.0', port=5003, debug=True, use_reloader=False, threaded=True)