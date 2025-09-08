import socket
import threading
import time

def test_server():
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        sock.bind(('127.0.0.1', 9999))
        sock.listen(1)
        print('테스트 서버가 포트 9999에서 리스닝 시작')
        
        conn, addr = sock.accept()
        print(f'연결됨: {addr}')
        response = 'HTTP/1.1 200 OK\r\nContent-Type: text/html\r\n\r\n<h1>Test Success!</h1>'
        conn.send(response.encode('utf-8'))
        conn.close()
    except Exception as e:
        print(f'서버 오류: {e}')
    finally:
        sock.close()

# 서버 시작
server_thread = threading.Thread(target=test_server)
server_thread.daemon = True
server_thread.start()

time.sleep(1)

# 연결 테스트
try:
    client_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    client_sock.connect(('127.0.0.1', 9999))
    response = client_sock.recv(1024)
    print('연결 성공! 응답 받음')
    client_sock.close()
    print('기본 소켓 연결은 정상 작동합니다!')
except Exception as e:
    print('연결 실패:', e)

time.sleep(1)