#!/usr/bin/env python3
"""
냉난방기 설치 공사 대시보드 시작 스크립트
"""

import os
import sys
import logging
from pathlib import Path

# 프로젝트 루트 경로 설정
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

# 환경 변수 로드
from dotenv import load_dotenv
load_dotenv()

# 로깅 설정
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('dashboard.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

def check_requirements():
    """필요한 설정 및 파일 체크"""
    checks = []
    
    # .env 파일 체크
    env_file = PROJECT_ROOT / '.env'
    if env_file.exists():
        checks.append(("[OK]", ".env 파일 존재"))
    else:
        checks.append(("[FAIL]", ".env 파일 없음 - .env.example을 참고하여 생성하세요"))
    
    # 구글 시트 ID 체크
    sheet_id = os.getenv('GOOGLE_SHEET_ID')
    if sheet_id:
        checks.append(("[OK]", f"구글 시트 ID 설정됨: {sheet_id[:20]}..."))
    else:
        checks.append(("[FAIL]", "GOOGLE_SHEET_ID 환경변수가 설정되지 않았습니다"))
    
    # 구글 자격증명 파일 체크
    creds_file = PROJECT_ROOT / 'credentials.json'
    if creds_file.exists():
        checks.append(("[OK]", "Google API 자격증명 파일 존재"))
    else:
        checks.append(("[FAIL]", "credentials.json 파일 없음 - Google Cloud Console에서 다운로드하세요"))
    
    # 데이터 폴더 체크
    data_folder = PROJECT_ROOT / 'data'
    if data_folder.exists():
        checks.append(("[OK]", "데이터 폴더 존재"))
        excel_files = list(data_folder.glob('*.xlsx'))
        if excel_files:
            checks.append(("[OK]", f"엑셀 파일 발견: {len(excel_files)}개"))
        else:
            checks.append(("[WARN]", "데이터 폴더에 엑셀 파일 없음"))
    else:
        checks.append(("[FAIL]", "data 폴더 없음"))
    
    # 결과 출력
    print("=" * 60)
    print("시스템 요구사항 체크")
    print("=" * 60)
    
    for status, message in checks:
        print(f"{status} {message}")
    
    print("=" * 60)
    
    # 에러가 있는지 체크
    errors = [check for check in checks if check[0] == "[FAIL]"]
    if errors:
        print("설정을 완료한 후 다시 실행해주세요.")
        return False
    
    print("모든 요구사항이 충족되었습니다!")
    return True

def main():
    """메인 함수"""
    print("냉난방기 설치 공사 대시보드")
    print("=" * 60)
    
    # 요구사항 체크
    if not check_requirements():
        sys.exit(1)
    
    print("\n대시보드를 시작합니다...")
    
    try:
        # 대시보드 앱 실행
        from dashboard.app import app, socketio
        
        # 초기 데이터 로드
        logger.info("초기 데이터 로드 중...")
        # boot 함수가 있다면 호출하여 초기 데이터를 로드
        try:
            from dashboard.app import boot
            boot()
        except ImportError:
            logger.info("boot 함수가 없습니다. 서버만 시작합니다.")
        
        # 서버 설정
        host = os.getenv('HOST', '0.0.0.0')
        port = int(os.getenv('PORT', 5000))
        debug = os.getenv('DEBUG', 'True').lower() == 'true'
        
        print(f"대시보드 URL: http://localhost:{port}")
        print("종료하려면 Ctrl+C를 누르세요")
        print("=" * 60)
        
        # Flask-SocketIO 서버 시작
        socketio.run(
            app, 
            host=host, 
            port=port, 
            debug=debug,
            use_reloader=False  # 알림 시스템과의 충돌 방지
        )
        
    except KeyboardInterrupt:
        print("\n사용자에 의해 종료되었습니다.")
    except Exception as e:
        logger.error(f"서버 시작 오류: {str(e)}")
        print(f"서버 시작 실패: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()