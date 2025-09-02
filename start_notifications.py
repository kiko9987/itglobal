#!/usr/bin/env python3
"""
알림 시스템 시작 스크립트
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
        logging.FileHandler('notifications.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

def main():
    """메인 함수"""
    print("📢 냉난방기 설치 공사 알림 시스템")
    print("=" * 60)
    
    # 설정 체크
    required_settings = [
        ('GOOGLE_SHEET_ID', '구글 시트 ID'),
        ('EMAIL_USERNAME', '이메일 계정'),
        ('EMAIL_PASSWORD', '이메일 패스워드')
    ]
    
    missing_settings = []
    for setting, description in required_settings:
        if not os.getenv(setting):
            missing_settings.append(f"❌ {setting} ({description})")
        else:
            print(f"✅ {description} 설정됨")
    
    if missing_settings:
        print("\n⚠️  다음 설정이 누락되었습니다:")
        for setting in missing_settings:
            print(f"   {setting}")
        print("\n.env 파일을 확인하고 필요한 설정을 추가해주세요.")
        sys.exit(1)
    
    # 선택적 설정 체크
    optional_settings = [
        ('SLACK_WEBHOOK_URL', '슬랙 웹훅'),
        ('ADMIN_EMAILS', '관리자 이메일'),
        ('SALES_EMAILS', '영업사원 이메일')
    ]
    
    print("\n선택적 설정:")
    for setting, description in optional_settings:
        if os.getenv(setting):
            print(f"✅ {description} 설정됨")
        else:
            print(f"⚠️  {description} 설정 안됨 (선택사항)")
    
    print("=" * 60)
    print("🚀 알림 시스템을 시작합니다...")
    print("📅 스케줄: 매일 오전 9시, 오후 6시 실행")
    print("🛑 종료하려면 Ctrl+C를 누르세요")
    print("=" * 60)
    
    try:
        # 알림 스케줄러 실행
        from dashboard.utils.notification_system import run_notification_scheduler
        run_notification_scheduler()
        
    except KeyboardInterrupt:
        print("\n🛑 사용자에 의해 종료되었습니다.")
    except Exception as e:
        logger.error(f"알림 시스템 오류: {str(e)}")
        print(f"❌ 알림 시스템 실패: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()