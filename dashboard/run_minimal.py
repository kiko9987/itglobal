#!/usr/bin/env python3
"""
최소한의 기능으로 안정적인 서버 실행
"""

import os
import sys
import logging
from flask import Flask, render_template, redirect, url_for, request, session
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
        return f"로그인 페이지를 로드할 수 없습니다: {str(e)}", 500

@app.route('/projects')
def projects():
    return "프로젝트 페이지 (임시)"

@app.errorhandler(404)
def not_found(error):
    return "페이지를 찾을 수 없습니다", 404

@app.errorhandler(500)  
def internal_error(error):
    logger.error(f"서버 내부 오류: {str(error)}")
    return "서버 내부 오류가 발생했습니다", 500

@app.errorhandler(Exception)
def handle_exception(e):
    logger.error(f"처리되지 않은 예외: {str(e)}", exc_info=True)
    return "예상치 못한 오류가 발생했습니다", 500

def main():
    try:
        host = '127.0.0.1'
        port = int(os.getenv('PORT', 5000))
        
        logger.info(f"최소 기능 서버 시작: http://localhost:{port}")
        logger.info("종료하려면 Ctrl+C를 누르세요")
        
        app.run(
            host=host,
            port=port,
            debug=False,
            use_reloader=False,
            threaded=True
        )
        
    except KeyboardInterrupt:
        logger.info("사용자에 의해 종료되었습니다.")
    except Exception as e:
        logger.error(f"서버 시작 오류: {str(e)}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()