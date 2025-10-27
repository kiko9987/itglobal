"""
공통 서버 설정 모듈
"""

import os
import logging
from flask import Flask
from dotenv import load_dotenv

def setup_basic_app(app_name="Dashboard"):
    """기본 Flask 앱 설정"""
    load_dotenv()
    
    # 로깅 설정
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler()]
    )
    
    app = Flask(__name__)
    app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key-here')
    
    # 공통 에러 핸들러
    @app.errorhandler(404)
    def not_found(error):
        return "페이지를 찾을 수 없습니다", 404

    @app.errorhandler(500)  
    def internal_error(error):
        logging.error(f"서버 내부 오류: {str(error)}")
        return "서버 내부 오류가 발생했습니다", 500

    @app.errorhandler(Exception)
    def handle_exception(e):
        logging.error(f"처리되지 않은 예외: {str(e)}", exc_info=True)
        return "예상치 못한 오류가 발생했습니다", 500
    
    return app

def get_server_config():
    """서버 설정 반환"""
    return {
        'host': '127.0.0.1',
        'port': int(os.getenv('PORT', 5000)),
        'debug': os.getenv('DEBUG', 'False').lower() == 'true'
    }