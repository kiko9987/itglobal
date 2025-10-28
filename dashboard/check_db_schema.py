import sqlite3

print("=== Claude Project DB Schema ===")
try:
    conn = sqlite3.connect('C:/Users/kiko9/Desktop/Claude Project/dashboard/instance/users.db')
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = cursor.fetchall()
    print(f"Tables: {[t[0] for t in tables]}")

    for table in tables:
        print(f"\n{table[0]} columns:")
        cursor.execute(f"PRAGMA table_info({table[0]})")
        for col in cursor.fetchall():
            print(f"  {col[1]} ({col[2]})")
    conn.close()
except Exception as e:
    print(f"Error: {e}")

print("\n\n=== Current itglobal DB Schema ===")
try:
    conn = sqlite3.connect('C:/Users/kiko9/Desktop/itglobal/dashboard/users.db')
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
    tables = cursor.fetchall()
    print(f"Tables: {[t[0] for t in tables]}")

    for table in tables:
        print(f"\n{table[0]} columns:")
        cursor.execute(f"PRAGMA table_info({table[0]})")
        for col in cursor.fetchall():
            print(f"  {col[1]} ({col[2]})")

        # 데이터 개수도 확인
        cursor.execute(f"SELECT COUNT(*) FROM {table[0]}")
        count = cursor.fetchone()[0]
        print(f"  Row count: {count}")

        # user 테이블이면 실제 데이터 출력
        if table[0] == 'user':
            cursor.execute(f"SELECT * FROM {table[0]}")
            for row in cursor.fetchall():
                print(f"  {row}")

    conn.close()
except Exception as e:
    print(f"Error: {e}")
