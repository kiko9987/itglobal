#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Flask 서버 안전 종료 스크립트
터미널이 죽지 않도록 안전하게 프로세스를 종료합니다.
"""

import subprocess
import time
import sys

def run_command(cmd):
    """명령어를 실행하고 결과를 반환"""
    try:
        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            text=True,
            timeout=5
        )
        return result.stdout, result.returncode
    except subprocess.TimeoutExpired:
        print(f"[!] Command timeout: {cmd}")
        return "", 1
    except Exception as e:
        print(f"[ERROR] Error running command: {e}")
        return "", 1

def find_process_on_port(port=5000):
    """특정 포트를 사용하는 프로세스 찾기"""
    cmd = f'netstat -ano | findstr :{port}'
    output, _ = run_command(cmd)

    pids = set()
    for line in output.split('\n'):
        if 'LISTENING' in line:
            parts = line.split()
            if parts:
                pid = parts[-1]
                pids.add(pid)

    return list(pids)

def stop_process(pid, force=False):
    """프로세스 종료"""
    if force:
        cmd = f'taskkill /PID {pid} /F /T'
    else:
        cmd = f'taskkill /PID {pid} /T'

    output, returncode = run_command(cmd)
    return returncode == 0

def main():
    print("=" * 50)
    print("Flask Server Safe Shutdown")
    print("=" * 50)
    print()

    # 5000 포트 사용 프로세스 찾기
    print("[*] Checking for server on port 5000...")
    pids = find_process_on_port(5000)

    if not pids:
        print("[OK] No process found on port 5000")
    else:
        print(f"[*] Found {len(pids)} process(es): {', '.join(pids)}")

        for pid in pids:
            print(f"\n[!] Stopping process {pid}...")

            # 먼저 정상 종료 시도
            if stop_process(pid, force=False):
                print(f"[OK] Process {pid} stopped gracefully")
            else:
                print(f"[!] Graceful stop failed, forcing...")
                time.sleep(1)

                # 강제 종료
                if stop_process(pid, force=True):
                    print(f"[OK] Process {pid} force stopped")
                else:
                    print(f"[ERROR] Failed to stop process {pid}")

    # Python 프로세스 확인
    print("\n[*] Checking for Python processes...")
    output, _ = run_command('tasklist | findstr python.exe')

    if output.strip():
        print("[*] Found Python processes:")
        print(output)

        response = input("\n[?] Stop all Python processes? (y/n): ").lower()
        if response == 'y':
            run_command('taskkill /IM python.exe /F /T')
            print("[OK] All Python processes stopped")
    else:
        print("[OK] No Python processes found")

    print("\n" + "=" * 50)
    print("Done!")
    print("=" * 50)

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n[!] Interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"\n[ERROR] Error: {e}")
        sys.exit(1)
