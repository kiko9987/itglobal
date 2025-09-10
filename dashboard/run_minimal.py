#!/usr/bin/env python3
"""
최소한의 기능으로 안정적인 서버 실행
"""

import os
import sys
import logging
from flask import render_template, redirect, url_for, request, session

# 프로젝트 루트 경로 추가
sys.path.append(os.path.dirname(__file__))
from utils.server_config import setup_basic_app, get_server_config

logger = logging.getLogger(__name__)

# 공통 설정으로 앱 생성
app = setup_basic_app("Minimal Dashboard")

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

def main():
    try:
        config = get_server_config()
        
        logger.info(f"최소 기능 서버 시작: http://localhost:{config['port']}")
        logger.info("종료하려면 Ctrl+C를 누르세요")
        
        app.run(
            host=config['host'],
            port=config['port'],
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