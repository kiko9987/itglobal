#!/usr/bin/env python
"""
dashboard/app.py - 레거시 호환성을 위한 간단한 팩토리 래퍼

이 파일은 기존 `python -m dashboard.app` 방식의 레거시 호환성을 제공합니다.
실제 애플리케이션 로직은 __init__.py의 create_app 팩토리에서 처리됩니다.

새로운 방식: run_server.py 사용 (권장)
레거시 방식: python -m dashboard.app (호환성용)
"""

import os
import sys
from dotenv import load_dotenv

# 환경 변수 로드
load_dotenv()

# 팩토리 패턴을 사용한 앱 생성
from dashboard import create_app

def main():
    """메인 실행 함수"""
    # 환경 설정
    config_name = os.getenv('FLASK_ENV', 'development')
    port = int(os.getenv('PORT', 5000))
    debug = os.getenv('FLASK_DEBUG', 'False').lower() in ('true', '1')

    # Flask 앱과 SocketIO 인스턴스 생성
    app, socketio = create_app(config_name)

    # 앱 인스턴스를 전역으로 노출 (레거시 호환성)
    globals()['app'] = app
    globals()['socketio'] = socketio

    # __name__ == '__main__' 블록에서 사용할 수 있도록 반환
    return app, socketio, port, debug

# 팩토리 실행 및 전역 변수 설정
app, socketio, port, debug = main()

if __name__ == '__main__':
    """
    레거시 호환성을 위한 직접 실행 지원
    권장 방법: run_server.py 사용
    """
    from dashboard.utils.logging_config import get_logger

    logger = get_logger(__name__)
    logger.info(f"레거시 모드로 대시보드 서버 시작: http://localhost:{port}")
    logger.info("권장 사항: run_server.py를 사용하세요")

    # Flask 디버그 모드 명시적 설정
    app.debug = debug

    # SocketIO 사용 가능 여부에 따라 서버 시작
    if socketio:
        socketio.run(app, debug=debug, host='0.0.0.0', port=port, use_reloader=True)
    else:
        app.run(debug=debug, host='0.0.0.0', port=port, use_reloader=True)