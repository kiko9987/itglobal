#!/usr/bin/env python3
"""
프로덕션용 안정적인 서버 실행
"""

import os
import sys
import logging
from flask import Flask, render_template, redirect, url_for
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

# 간단한 Flask 앱 생성
app = Flask(__name__)
app.config['SECRET_KEY'] = os.getenv('FLASK_SECRET_KEY', 'simple-key-for-testing')

@app.route('/')
def index():
    return redirect(url_for('login_page'))

@app.route('/login')
def login_page():
    try:
        return render_template('login.html')
    except Exception as e:
        logger.error(f"로그인 페이지 오류: {str(e)}")
        return f"<h1>로그인 페이지</h1><p>임시 로그인 페이지입니다.</p><p>오류: {str(e)}</p>"

@app.route('/projects')
def projects():
    return "<h1>프로젝트 페이지</h1><p>프로젝트 목록이 여기에 표시됩니다.</p>"

@app.errorhandler(404)
def not_found(error):
    return "<h1>404</h1><p>페이지를 찾을 수 없습니다</p>", 404

@app.errorhandler(500)  
def internal_error(error):
    logger.error(f"서버 내부 오류: {str(error)}")
    return "<h1>500</h1><p>서버 내부 오류가 발생했습니다</p>", 500

@app.errorhandler(Exception)
def handle_exception(e):
    logger.error(f"처리되지 않은 예외: {str(e)}", exc_info=True)
    return "<h1>오류</h1><p>예상치 못한 오류가 발생했습니다</p>", 500

def main():
    try:
        host = '127.0.0.1'
        port = int(os.getenv('PORT', 5000))
        
        logger.info("=" * 50)
        logger.info("프로덕션 서버 시작")
        logger.info(f"URL: http://localhost:{port}")
        logger.info("종료하려면 Ctrl+C를 누르세요")
        logger.info("=" * 50)
        
        # Waitress 서버로 실행 (개발 모드 없음)
        serve(
            app,
            host=host,
            port=port,
            threads=4,              # 멀티스레드
            cleanup_interval=30,    # 30초마다 정리
            channel_timeout=120     # 2분 타임아웃
        )
        
    except KeyboardInterrupt:
        logger.info("사용자에 의해 종료되었습니다.")
    except Exception as e:
        logger.error(f"서버 시작 오류: {str(e)}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()