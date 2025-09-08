#!/usr/bin/env python3
"""
절대 죽지 않는 안정적인 서버
"""

import os
import sys
import logging
import traceback
from flask import Flask, render_template, redirect, url_for, request, jsonify
from waitress import serve
from dotenv import load_dotenv

# 환경 변수 로드
load_dotenv()

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# Flask 앱 생성
app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('FLASK_SECRET_KEY', 'bulletproof-key')

# 전역 변수
dashboard_data = None

def safe_load_data():
    """안전한 데이터 로드"""
    global dashboard_data
    try:
        from app import load_data
        dashboard_data = load_data()
        logger.info("데이터 로드 성공")
        return True
    except Exception as e:
        logger.warning(f"데이터 로드 실패: {str(e)} - 기본 데이터로 실행")
        dashboard_data = {"message": "데이터를 로드할 수 없습니다"}
        return False

@app.route('/')
def index():
    try:
        return redirect(url_for('login_page'))
    except Exception as e:
        logger.error(f"인덱스 페이지 오류: {str(e)}")
        return "<h1>IT Global 대시보드</h1><p>메인 페이지 로드 중 오류가 발생했습니다.</p>"

@app.route('/login')
def login_page():
    try:
        return render_template('login.html')
    except Exception as e:
        logger.error(f"로그인 템플릿 로드 실패: {str(e)}")
        return """
        <html>
        <head><title>IT Global 대시보드</title></head>
        <body>
            <h1>IT Global 대시보드</h1>
            <h2>로그인</h2>
            <p>서버가 정상적으로 작동하고 있습니다!</p>
            <p>Google OAuth 인증을 설정하세요.</p>
            <a href="/projects">프로젝트 페이지로 이동</a>
        </body>
        </html>
        """

@app.route('/projects')
def projects():
    try:
        # 원본 프로젝트 로직 시도
        from app import project_list
        return project_list()
    except Exception as e:
        logger.error(f"프로젝트 페이지 오류: {str(e)}")
        return f"""
        <html>
        <head><title>프로젝트 관리</title></head>
        <body>
            <h1>프로젝트 관리</h1>
            <p>프로젝트 데이터를 로드하는 중 오류가 발생했습니다.</p>
            <p>오류: {str(e)}</p>
            <p>서버는 정상적으로 작동 중입니다.</p>
            <a href="/login">로그인 페이지로 돌아가기</a>
        </body>
        </html>
        """

# 모든 예외를 캐치하는 글로벌 에러 핸들러
@app.errorhandler(Exception)
def handle_exception(e):
    logger.error(f"처리되지 않은 예외: {str(e)}")
    logger.error(traceback.format_exc())
    return f"""
    <html>
    <head><title>오류 발생</title></head>
    <body>
        <h1>오류가 발생했습니다</h1>
        <p>서버는 계속 작동 중입니다.</p>
        <p>오류 내용: {str(e)}</p>
        <a href="/">메인 페이지로 돌아가기</a>
    </body>
    </html>
    """, 500

@app.errorhandler(404)
def not_found(error):
    return """
    <html>
    <head><title>페이지를 찾을 수 없음</title></head>
    <body>
        <h1>404 - 페이지를 찾을 수 없습니다</h1>
        <a href="/">메인 페이지로 돌아가기</a>
    </body>
    </html>
    """, 404

def main():
    try:
        # 데이터 로드 시도
        logger.info("초기 데이터 로드 시도 중...")
        safe_load_data()
        
        # 서버 설정
        host = '127.0.0.1'
        port = 9999
        
        logger.info("=" * 60)
        logger.info("🛡️ 절대 죽지 않는 IT Global 대시보드 🛡️")
        logger.info(f"🌐 URL: http://localhost:{port}")
        logger.info("💪 모든 오류를 처리하여 안정성 보장")
        logger.info("⏹️ 종료하려면 Ctrl+C를 누르세요")
        logger.info("=" * 60)
        
        # Waitress 서버 실행
        serve(
            app,
            host=host,
            port=port,
            threads=4,
            cleanup_interval=60,
            channel_timeout=300
        )
        
    except KeyboardInterrupt:
        logger.info("사용자에 의해 종료되었습니다.")
    except Exception as e:
        logger.error(f"서버 시작 오류: {str(e)}")
        logger.error(traceback.format_exc())
        input("Enter 키를 누르면 종료됩니다...")

if __name__ == "__main__":
    main()