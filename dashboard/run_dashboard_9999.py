#!/usr/bin/env python3
"""
원본 대시보드를 포트 9999로 실행
"""

import os
import sys
import logging
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

def main():
    try:
        # 원본 app 모듈 임포트
        from app import app, load_data
        
        # 초기 데이터 로드
        logger.info("초기 데이터 로드 중...")
        try:
            load_data()
            logger.info("데이터 로드 완료")
        except Exception as e:
            logger.warning(f"데이터 로드 실패: {str(e)} - 서버는 계속 시작합니다")
        
        # 서버 설정
        host = '127.0.0.1'
        port = 9999
        
        logger.info("=" * 60)
        logger.info("🎉 IT Global 대시보드 서버 시작! 🎉")
        logger.info(f"🌐 URL: http://localhost:{port}")
        logger.info("🔐 Google OAuth 인증 지원")
        logger.info("⏹️  종료하려면 Ctrl+C를 누르세요")
        logger.info("=" * 60)
        
        # Waitress 프로덕션 서버로 실행
        serve(
            app,
            host=host,
            port=port,
            threads=6,
            cleanup_interval=30,    
            channel_timeout=120,
            connection_limit=1000
        )
        
    except KeyboardInterrupt:
        logger.info("사용자에 의해 종료되었습니다.")
    except Exception as e:
        logger.error(f"서버 시작 오류: {str(e)}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()