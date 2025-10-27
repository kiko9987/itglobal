#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
서버 프로세스 정리 및 재시작 도구
"""

import os
import sys
import psutil
import time

def kill_python_servers():
    """Python 서버 프로세스 모두 종료"""
    killed_count = 0
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            if ('python' in proc.info['name'].lower() and
                proc.info['cmdline'] and
                any('run_server.py' in arg or 'start_servers.py' in arg
                    for arg in proc.info['cmdline'])):
                print(f"종료: PID {proc.pid}")
                proc.terminate()
                proc.wait(timeout=3)
                killed_count += 1
        except (psutil.NoSuchProcess, psutil.TimeoutExpired, psutil.AccessDenied):
            pass
    return killed_count

def start_clean_server():
    """깨끗한 서버 시작"""
    print("[1/3] 기존 서버 프로세스 정리...")
    killed = kill_python_servers()
    print(f"     {killed}개 프로세스 종료됨")

    if killed > 0:
        print("[2/3] 포트 해제 대기...")
        time.sleep(3)

    print("[3/3] 새 서버 시작...")
    os.environ['PORT'] = '5000'
    os.environ['FLASK_DEBUG'] = 'True'

    # 새 서버 시작
    os.system('python run_server.py')

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == 'clean':
        print("서버 프로세스 정리 중...")
        killed = kill_python_servers()
        print(f"완료: {killed}개 프로세스 정리됨")
    else:
        start_clean_server()