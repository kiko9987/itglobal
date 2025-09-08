#!/usr/bin/env python3
"""
안정적인 서버 실행 스크립트
"""

import os
import sys
import logging
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

def main():
    try:
        # app 모듈 임포트
        from app import app, socketio, load_data
        
        # Flask 앱에 에러 핸들러 추가
        @app.errorhandler(Exception)
        def handle_exception(e):
            logger.error(f"처리되지 않은 예외: {str(e)}", exc_info=True)
            return "서버 내부 오류가 발생했습니다.", 500
        
        # 초기 데이터 로드
        logger.info("초기 데이터 로드 중...")
        load_data()
        
        # 서버 설정
        host = '127.0.0.1'  # localhost만 사용
        port = int(os.getenv('PORT', 5000))
        
        logger.info(f"서버 시작: http://localhost:{port}")
        logger.info("종료하려면 Ctrl+C를 누르세요")
        
        # 서버 시작 (안정적인 설정)
        app.run(
            host=host,
            port=port,
            debug=False,           # 디버그 모드 완전 비활성화
            use_reloader=False,    # 자동 재시작 비활성화
            threaded=True          # 멀티스레드 지원
        )
        
    except KeyboardInterrupt:
        logger.info("사용자에 의해 종료되었습니다.")
    except Exception as e:
        logger.error(f"서버 시작 오류: {str(e)}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()