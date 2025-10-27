"""
실서버 배포 시 한 번만 실행하는 마이그레이션
기존 퇴사자 4명을 users 테이블에 추가하고 퇴사 처리

Usage:
    python migrations/add_resigned_users.py
"""
import sys
import os

# 프로젝트 루트를 Python 경로에 추가
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dashboard.utils.user_database import get_user_database


def migrate_resigned_users():
    """
    기존 퇴사자 4명을 users 테이블에 추가하고 퇴사 처리
    이 4명은 테이블에 담당했던 공사 히스토리가 남아있음
    """
    resigned_users = [
        {'name': '심장원', 'email': 'resigned1@itg-aircon.com'},
        {'name': '아이티', 'email': 'resigned2@itg-aircon.com'},
        {'name': '이근혁', 'email': 'resigned3@itg-aircon.com'},
        {'name': '김단이', 'email': 'resigned4@itg-aircon.com'}
    ]

    user_db = get_user_database()

    print("=" * 60)
    print("기존 퇴사자 마이그레이션 시작")
    print("=" * 60)

    success_count = 0
    error_count = 0

    for user in resigned_users:
        name = user['name']
        email = user['email']

        try:
            # 더미 사용자 생성 (비밀번호는 랜덤, 로그인 불가)
            success, message = user_db.create_user(
                name=name,
                email=email,
                password='resigned_dummy_password_no_login',
                permission_level='viewer'
            )

            if success:
                print(f"[CREATE] {name} ({email}): 사용자 생성됨")
            elif '이미 존재' in message:
                print(f"[SKIP] {name} ({email}): 이미 존재하는 사용자")
            else:
                print(f"[ERROR] {name} 생성 실패: {message}")
                error_count += 1
                continue

            # 퇴사 처리
            success2, message2 = user_db.update_user_status(email, '퇴사')
            if success2:
                print(f"[RESIGNED] {name} ({email}): 퇴사 처리 완료")
                success_count += 1
            else:
                print(f"[ERROR] {name} 퇴사 처리 실패: {message2}")
                error_count += 1

        except Exception as e:
            print(f"[ERROR] {name} 처리 중 오류: {e}")
            error_count += 1

    print("=" * 60)
    print(f"마이그레이션 완료: 성공 {success_count}명, 실패 {error_count}명")
    print("=" * 60)

    # 최종 퇴사자 목록 확인
    resigned_managers = user_db.get_resigned_managers()
    print(f"\n현재 퇴사자 목록 ({len(resigned_managers)}명):")
    for manager in resigned_managers:
        print(f"  - {manager}")

    return success_count, error_count


if __name__ == '__main__':
    try:
        success, error = migrate_resigned_users()
        sys.exit(0 if error == 0 else 1)
    except Exception as e:
        print(f"\n[FATAL ERROR] 마이그레이션 실패: {e}")
        sys.exit(1)
