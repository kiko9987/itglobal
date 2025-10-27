import sys
sys.path.insert(0, 'dashboard')
from dashboard.utils.user_database import get_user_database

# 통합 DB 사용 (instance/users.db)
user_db = get_user_database()
user = user_db.get_user_by_email('kiko@itg-aircon.com')

if user:
    print(f'Email: {user["email"]}, Role: {user["permission_level"]}, Active: {user["is_active"]}')
else:
    print('User not found')
