#!/usr/bin/env python3
"""
통합 개발 환경 스타트 스크립트
- 포트 자동 감지 및 할당
- Flask + Vite 개발서버 동시 실행
- 브라우저 자동 열기
"""

import os
import sys
import time
import socket
import subprocess
import threading
import webbrowser
from pathlib import Path

# UTF-8 출력 설정 (Windows 콘솔 호환)
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

class DevServer:
    def __init__(self):
        self.flask_port = 5000
        self.vite_port = 5173
        self.processes = []

    def is_port_available(self, port):
        """포트 사용 가능 여부 확인"""
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            try:
                sock.bind(('localhost', port))
                return True
            except OSError:
                return False

    def find_available_port(self, start_port=5000, max_attempts=50):
        """사용 가능한 포트 찾기"""
        for port in range(start_port, start_port + max_attempts):
            if self.is_port_available(port):
                return port
        raise Exception(f"사용 가능한 포트를 찾을 수 없습니다 ({start_port}-{start_port + max_attempts})")

    def kill_existing_servers(self):
        """기존 서버들 종료"""
        print("🔄 기존 서버들 확인 중...")

        # Python 서버들 확인
        try:
            result = subprocess.run(['netstat', '-ano'], capture_output=True, text=True)
            lines = result.stdout.split('\n')

            ports_to_check = [5000, 5001, 5002, 5003, 5173, 5174, 5175]
            for line in lines:
                if 'LISTENING' in line:
                    for port in ports_to_check:
                        if f':{port} ' in line:
                            print(f"⚠️  포트 {port}가 사용 중입니다")
        except:
            pass

    def start_flask_server(self):
        """Flask 개발서버 시작"""
        print(f"🐍 Flask 서버 시작 중... (포트: {self.flask_port})")

        env = os.environ.copy()
        env['PORT'] = str(self.flask_port)
        env['FLASK_ENV'] = 'development'
        env['FLASK_DEBUG'] = '1'

        process = subprocess.Popen(
            [sys.executable, 'run_server.py'],
            cwd=Path(__file__).parent,
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            bufsize=1,
            universal_newlines=True
        )

        self.processes.append(('Flask', process))

        # Flask 서버 로그 출력
        def flask_logger():
            for line in iter(process.stdout.readline, ''):
                if line.strip():
                    print(f"[Flask] {line.strip()}")

        threading.Thread(target=flask_logger, daemon=True).start()
        return process

    def start_vite_server(self):
        """Vite 개발서버 시작"""
        print(f"🔥 Vite 서버 시작 중... (포트: {self.vite_port})")

        dashboard_dir = Path(__file__).parent / 'dashboard'

        # Vite 설정에서 Flask 포트 업데이트
        self.update_vite_config()

        process = subprocess.Popen(
            ['npm', 'run', 'dev'],
            cwd=dashboard_dir,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding='utf-8',
            bufsize=1,
            universal_newlines=True,
            shell=True  # Windows에서 npm.cmd 찾기 위해
        )

        self.processes.append(('Vite', process))

        # Vite 서버 로그 출력
        def vite_logger():
            for line in iter(process.stdout.readline, ''):
                if line.strip():
                    print(f"[Vite] {line.strip()}")
                    # Vite 서버가 시작되면 포트 감지
                    if 'Local:' in line and 'localhost:' in line:
                        import re
                        match = re.search(r'localhost:(\d+)', line)
                        if match:
                            self.vite_port = int(match.group(1))
                            print(f"✅ Vite 서버 포트 감지: {self.vite_port}")

        threading.Thread(target=vite_logger, daemon=True).start()
        return process

    def update_vite_config(self):
        """Vite 설정에서 Flask 포트 업데이트"""
        vite_config_path = Path(__file__).parent / 'dashboard' / 'vite.config.js'

        try:
            with open(vite_config_path, 'r', encoding='utf-8') as f:
                content = f.read()

            # 프록시 설정의 Flask 포트 업데이트
            import re
            content = re.sub(
                r'http://localhost:\d+',
                f'http://localhost:{self.flask_port}',
                content
            )

            with open(vite_config_path, 'w', encoding='utf-8') as f:
                f.write(content)

            print(f"✅ Vite 설정 업데이트: Flask 포트 {self.flask_port}")

        except Exception as e:
            print(f"⚠️  Vite 설정 업데이트 실패: {e}")

    def wait_for_servers(self):
        """서버들이 시작될 때까지 대기"""
        print("⏳ 서버 시작 대기 중...")

        # Flask 서버 대기
        for attempt in range(30):
            if self.is_port_available(self.flask_port):
                time.sleep(1)
            else:
                print(f"✅ Flask 서버 준비 완료 (포트: {self.flask_port})")
                break
        else:
            print("❌ Flask 서버 시작 실패")
            return False

        # Vite 서버 대기 (추가 시간)
        time.sleep(3)
        print(f"✅ Vite 서버 준비 완료 (포트: {self.vite_port})")
        return True

    def open_browser(self):
        """브라우저 열기"""
        url = f"http://localhost:{self.vite_port}/projects"
        print(f"🌐 브라우저 열기: {url}")

        try:
            # Chrome으로 열기 시도
            chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            if os.path.exists(chrome_path):
                subprocess.Popen([
                    chrome_path,
                    '--new-window',
                    '--disable-web-security',
                    '--disable-features=VizDisplayCompositor',
                    '--incognito',
                    url
                ])
            else:
                webbrowser.open(url)
        except Exception as e:
            print(f"⚠️  브라우저 열기 실패: {e}")
            print(f"수동으로 접속해주세요: {url}")

    def start(self):
        """개발 환경 시작"""
        print("🚀 IT Global 개발 환경 시작")
        print("=" * 50)

        try:
            # 포트 확인 및 할당
            self.kill_existing_servers()

            # Flask 포트 할당
            if not self.is_port_available(self.flask_port):
                self.flask_port = self.find_available_port(5000)
                print(f"🔄 Flask 포트 변경: {self.flask_port}")

            # Vite 포트 할당
            if not self.is_port_available(self.vite_port):
                self.vite_port = self.find_available_port(5173)
                print(f"🔄 Vite 포트 변경: {self.vite_port}")

            # 서버들 시작
            flask_process = self.start_flask_server()
            time.sleep(2)  # Flask 서버가 먼저 시작되도록

            vite_process = self.start_vite_server()

            # 서버 준비 대기
            if not self.wait_for_servers():
                return

            # 브라우저 열기
            self.open_browser()

            print("\n" + "=" * 50)
            print("🎉 개발 환경 준비 완료!")
            print(f"📱 Frontend (Vite): http://localhost:{self.vite_port}/projects")
            print(f"⚙️  Backend (Flask): http://localhost:{self.flask_port}")
            print("🔥 핫리로드 활성화 - 파일 수정 시 자동 반영")
            print("\n💡 개발 팁:")
            print("   - JavaScript/CSS 파일 수정 → 즉시 반영")
            print("   - Python 파일 수정 → Flask 서버 재시작 필요")
            print("   - Ctrl+C로 모든 서버 종료")
            print("=" * 50)

            # 메인 스레드에서 대기
            try:
                while True:
                    time.sleep(1)
                    # 프로세스 상태 확인
                    for name, process in self.processes:
                        if process.poll() is not None:
                            print(f"❌ {name} 서버가 종료되었습니다")
                            return
            except KeyboardInterrupt:
                print("\n🛑 서버 종료 중...")

        except Exception as e:
            print(f"❌ 오류 발생: {e}")

        finally:
            self.cleanup()

    def cleanup(self):
        """정리 작업"""
        print("🧹 정리 작업 중...")
        for name, process in self.processes:
            try:
                process.terminate()
                process.wait(timeout=5)
                print(f"✅ {name} 서버 종료")
            except:
                try:
                    process.kill()
                    print(f"🔪 {name} 서버 강제 종료")
                except:
                    pass

if __name__ == "__main__":
    server = DevServer()
    server.start()