#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
개발 서버 관리 도구 - 중복 실행 방지 및 안전한 서버 관리
"""

import os
import sys
import psutil
import time
import signal
import subprocess
from pathlib import Path

class DevServerManager:
    def __init__(self, project_root=None):
        self.project_root = project_root or os.getcwd()
        self.pid_file = os.path.join(self.project_root, '.dev_server.pid')
        self.default_port = 5000

    def is_port_in_use(self, port):
        """포트 사용 여부 확인"""
        for conn in psutil.net_connections():
            if conn.laddr.port == port and conn.status == psutil.CONN_LISTEN:
                return True
        return False

    def find_server_processes(self):
        """현재 실행 중인 서버 프로세스 찾기"""
        server_processes = []
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                cmdline = ' '.join(proc.info['cmdline'] or [])
                if ('python' in proc.info['name'].lower() and
                    ('run_server.py' in cmdline or 'start_servers.py' in cmdline)):
                    server_processes.append(proc)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return server_processes

    def kill_existing_servers(self):
        """기존 서버 프로세스 모두 종료"""
        processes = self.find_server_processes()
        killed_count = 0

        for proc in processes:
            try:
                print(f"종료 중: PID {proc.pid} - {' '.join(proc.cmdline())}")
                proc.terminate()
                proc.wait(timeout=5)
                killed_count += 1
            except (psutil.NoSuchProcess, psutil.TimeoutExpired):
                try:
                    proc.kill()
                    killed_count += 1
                except psutil.NoSuchProcess:
                    pass
            except Exception as e:
                print(f"프로세스 {proc.pid} 종료 실패: {e}")

        # PID 파일 정리
        if os.path.exists(self.pid_file):
            os.remove(self.pid_file)

        return killed_count

    def save_pid(self, pid):
        """현재 서버 PID 저장"""
        with open(self.pid_file, 'w') as f:
            f.write(str(pid))

    def get_saved_pid(self):
        """저장된 PID 가져오기"""
        if os.path.exists(self.pid_file):
            try:
                with open(self.pid_file, 'r') as f:
                    return int(f.read().strip())
            except (ValueError, FileNotFoundError):
                pass
        return None

    def is_server_running(self):
        """서버가 실행 중인지 확인"""
        pid = self.get_saved_pid()
        if pid:
            try:
                proc = psutil.Process(pid)
                cmdline = ' '.join(proc.cmdline())
                if 'run_server.py' in cmdline or 'start_servers.py' in cmdline:
                    return True
            except psutil.NoSuchProcess:
                pass
        return False

    def start_clean_server(self, port=None, debug=True):
        """깨끗한 단일 서버 시작"""
        port = port or self.default_port

        print("[CLEAN] 기존 서버 프로세스 정리 중...")
        killed = self.kill_existing_servers()
        if killed > 0:
            print(f"[OK] {killed}개 프로세스 종료됨")
            time.sleep(2)  # 포트 해제 대기

        # 포트 확인
        if self.is_port_in_use(port):
            print(f"[ERROR] 포트 {port}이 여전히 사용 중입니다.")
            return False

        print(f"[START] 새 서버 시작 중 (포트 {port})...")

        # 환경 변수 설정
        env = os.environ.copy()
        env['PORT'] = str(port)
        env['FLASK_DEBUG'] = 'True' if debug else 'False'

        # 서버 시작 (백그라운드 아님)
        cmd = [sys.executable, 'run_server.py']

        try:
            # subprocess로 시작하여 PID 추적 가능
            proc = subprocess.Popen(
                cmd,
                cwd=self.project_root,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
                universal_newlines=True
            )

            # PID 저장
            self.save_pid(proc.pid)

            print(f"✅ 서버 시작됨 (PID: {proc.pid})")
            print(f"🌐 http://localhost:{port}")
            print("📝 로그 확인을 위해 잠시 대기 중...")

            # 초기 로그 확인 (5초)
            start_time = time.time()
            while time.time() - start_time < 5:
                if proc.poll() is not None:
                    print("❌ 서버가 예기치 않게 종료됨")
                    return False

                try:
                    line = proc.stdout.readline()
                    if line:
                        print(f"[SERVER] {line.strip()}")
                        if "Running on" in line:
                            print("✅ 서버 정상 시작 확인됨")
                            return True
                except:
                    pass

                time.sleep(0.1)

            return True

        except Exception as e:
            print(f"❌ 서버 시작 실패: {e}")
            return False

    def stop_server(self):
        """현재 서버 정지"""
        pid = self.get_saved_pid()
        if pid:
            try:
                proc = psutil.Process(pid)
                proc.terminate()
                proc.wait(timeout=5)
                print(f"✅ 서버 정지됨 (PID: {pid})")
            except (psutil.NoSuchProcess, psutil.TimeoutExpired):
                print("⚠️ 서버가 이미 종료되었거나 강제 종료됨")

        if os.path.exists(self.pid_file):
            os.remove(self.pid_file)

    def status(self):
        """서버 상태 확인"""
        print("🔍 서버 상태 확인 중...")

        # 저장된 PID 확인
        pid = self.get_saved_pid()
        if pid:
            try:
                proc = psutil.Process(pid)
                print(f"✅ 메인 서버 실행 중 (PID: {pid})")
                print(f"   포트: {self.default_port}")
                print(f"   명령: {' '.join(proc.cmdline())}")
            except psutil.NoSuchProcess:
                print("❌ 저장된 PID의 프로세스가 존재하지 않음")
        else:
            print("❌ 저장된 서버 PID 없음")

        # 모든 서버 프로세스 확인
        processes = self.find_server_processes()
        if processes:
            print(f"\n⚠️ 발견된 서버 프로세스: {len(processes)}개")
            for proc in processes:
                try:
                    print(f"   PID {proc.pid}: {' '.join(proc.cmdline())}")
                except:
                    print(f"   PID {proc.pid}: <명령 확인 불가>")
        else:
            print("\n✅ 추가 서버 프로세스 없음")

        # 포트 사용 현황
        used_ports = []
        for port in [5000, 5001, 5002, 5010, 5020, 8000, 8080, 8888]:
            if self.is_port_in_use(port):
                used_ports.append(port)

        if used_ports:
            print(f"\n🔌 사용 중인 포트: {', '.join(map(str, used_ports))}")
        else:
            print("\n✅ 주요 포트 모두 사용 가능")

def main():
    manager = DevServerManager()

    if len(sys.argv) < 2:
        print("사용법:")
        print("  python dev_server_manager.py start [port]  # 서버 시작")
        print("  python dev_server_manager.py stop          # 서버 정지")
        print("  python dev_server_manager.py restart [port] # 서버 재시작")
        print("  python dev_server_manager.py status        # 상태 확인")
        print("  python dev_server_manager.py clean         # 모든 서버 정리")
        return

    command = sys.argv[1].lower()

    if command == 'start':
        port = int(sys.argv[2]) if len(sys.argv) > 2 else 5000
        success = manager.start_clean_server(port)
        if success:
            print("\n🎉 서버가 성공적으로 시작되었습니다!")
            print("🛑 서버를 정지하려면: python dev_server_manager.py stop")
        else:
            print("\n❌ 서버 시작에 실패했습니다.")

    elif command == 'stop':
        manager.stop_server()

    elif command == 'restart':
        port = int(sys.argv[2]) if len(sys.argv) > 2 else 5000
        print("🔄 서버 재시작 중...")
        manager.stop_server()
        time.sleep(2)
        manager.start_clean_server(port)

    elif command == 'status':
        manager.status()

    elif command == 'clean':
        print("🧹 모든 서버 프로세스 정리 중...")
        killed = manager.kill_existing_servers()
        print(f"✅ {killed}개 프로세스가 정리되었습니다.")

    else:
        print(f"❌ 알 수 없는 명령: {command}")

if __name__ == "__main__":
    main()