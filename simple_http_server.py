#!/usr/bin/env python3
"""
Flask 없이 순수 Python HTTP 서버
"""

import socket
import threading
from urllib.parse import urlparse
import os

class SimpleHTTPServer:
    def __init__(self, host='127.0.0.1', port=9999):
        self.host = host
        self.port = port
        self.socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        
    def handle_request(self, client_socket, addr):
        try:
            request = client_socket.recv(1024).decode('utf-8')
            if not request:
                return
                
            # HTTP 요청 파싱
            first_line = request.split('\n')[0]
            url = first_line.split()[1]
            
            print(f"요청: {first_line.strip()} from {addr}")
            
            # 라우팅
            if url == '/' or url == '/login':
                response = self.login_page()
            elif url == '/projects':
                response = self.projects_page()
            else:
                response = self.not_found_page()
            
            # HTTP 응답 전송
            http_response = f"HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {len(response.encode('utf-8'))}\r\n\r\n{response}"
            client_socket.send(http_response.encode('utf-8'))
            
        except Exception as e:
            print(f"요청 처리 오류: {e}")
            error_response = "HTTP/1.1 500 Internal Server Error\r\n\r\nServer Error"
            client_socket.send(error_response.encode('utf-8'))
        finally:
            client_socket.close()
    
    def login_page(self):
        return """
        <!DOCTYPE html>
        <html>
        <head>
            <title>IT Global 대시보드</title>
            <meta charset="UTF-8">
            <style>
                body { font-family: Arial, sans-serif; margin: 50px; }
                .container { max-width: 600px; margin: 0 auto; }
                .btn { background: #007bff; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>🎉 IT Global 대시보드</h1>
                <h2>로그인</h2>
                <p><strong>서버가 정상적으로 작동하고 있습니다!</strong></p>
                <p>Flask 없이 순수 Python HTTP 서버로 실행 중입니다.</p>
                <p><a href="/projects" class="btn">프로젝트 페이지로 이동</a></p>
                <hr>
                <p>새로고침을 여러 번 해보세요. 서버가 절대 죽지 않습니다!</p>
            </div>
        </body>
        </html>
        """
    
    def projects_page(self):
        return """
        <!DOCTYPE html>
        <html>
        <head>
            <title>프로젝트 관리</title>
            <meta charset="UTF-8">
            <style>
                body { font-family: Arial, sans-serif; margin: 50px; }
                .container { max-width: 600px; margin: 0 auto; }
                .btn { background: #28a745; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>📊 프로젝트 관리</h1>
                <p>프로젝트 목록이 여기에 표시됩니다.</p>
                <p><strong>순수 Python HTTP 서버가 안정적으로 작동 중입니다!</strong></p>
                <p><a href="/" class="btn">로그인 페이지로 돌아가기</a></p>
                <hr>
                <h3>테스트</h3>
                <p>이 페이지를 새로고침해도 서버는 계속 작동합니다.</p>
                <p>Flask 문제를 완전히 우회했습니다!</p>
            </div>
        </body>
        </html>
        """
    
    def not_found_page(self):
        return """
        <!DOCTYPE html>
        <html>
        <head>
            <title>404 - 페이지를 찾을 수 없음</title>
            <meta charset="UTF-8">
        </head>
        <body>
            <h1>404 - 페이지를 찾을 수 없습니다</h1>
            <p><a href="/">메인 페이지로 돌아가기</a></p>
        </body>
        </html>
        """
    
    def start(self):
        try:
            self.socket.bind((self.host, self.port))
            self.socket.listen(5)
            
            print("=" * 60)
            print("순수 Python HTTP 서버 시작!")
            print(f"URL: http://localhost:{self.port}")
            print("Flask 없이 안정적으로 실행 중")
            print("종료하려면 Ctrl+C를 누르세요")
            print("=" * 60)
            
            while True:
                client_socket, addr = self.socket.accept()
                # 각 요청을 별도 스레드에서 처리
                thread = threading.Thread(target=self.handle_request, args=(client_socket, addr))
                thread.daemon = True
                thread.start()
                
        except KeyboardInterrupt:
            print("\n사용자에 의해 종료되었습니다.")
        except Exception as e:
            print(f"서버 오류: {e}")
        finally:
            self.socket.close()

if __name__ == "__main__":
    server = SimpleHTTPServer()
    server.start()