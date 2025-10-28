import sqlite3

print("=== Claude Project Users ===")
conn = sqlite3.connect('C:/Users/kiko9/Desktop/Claude Project/dashboard/instance/users.db')
cursor = conn.cursor()
cursor.execute('SELECT email, role FROM users')
for row in cursor.fetchall():
    print(f'{row[0]}: {row[1]}')
conn.close()

print("\n=== Current itglobal Users ===")
conn = sqlite3.connect('C:/Users/kiko9/Desktop/itglobal/dashboard/users.db')
cursor = conn.cursor()
cursor.execute('SELECT email, role FROM users')
for row in cursor.fetchall():
    print(f'{row[0]}: {row[1]}')
conn.close()
