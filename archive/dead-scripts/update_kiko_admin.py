import sys
sys.path.insert(0, 'dashboard')
from dashboard.utils.user_database import get_user_database

# 통합 DB 사용 (instance/users.db)
user_db = get_user_database()

# 권한 업데이트 (세션 무효화 포함)
success, message = user_db.update_user_permission('kiko@itg-aircon.com', 'super_admin')
print(message)

if success:
    # 확인
    user = user_db.get_user_by_email('kiko@itg-aircon.com')
    if user:
        print(f'Email: {user["email"]}, Role: {user["permission_level"]}, Active: {user["is_active"]}')
