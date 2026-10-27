# -*- coding: utf-8 -*-
"""서버 재시작 승인 요청 — 사장님 슬랙 DM 에 [지금 재시작] [취소] 카드를 보낸다.

사용:  .venv/Scripts/python.exe scripts/request_restart.py "문서 변환 봇 조각 파일 반영"
       (대기 중인 요청이 있으면 이유만 덧붙이고 같은 카드를 갱신)
       .venv/Scripts/python.exe scripts/request_restart.py --status   # 현재 요청 상태

승인되면 SYSTEM 워치독(watchdog.ps1)이 1분 안에 재시작하고 카드에 결과를 적는다.
create_app() 은 부르지 않는다(스케줄러 이중 기동 방지) — 슬랙 WebClient 만 쓴다.
"""
import json
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def _load_env():
    env_path = ROOT / '.env'
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if not line or line.startswith('#') or '=' not in line:
            continue
        key, _, val = line.partition('=')
        val = val.split(' #')[0].strip().strip('"').strip("'")
        os.environ.setdefault(key.strip(), val)


def main(argv):
    _load_env()
    from dashboard.services import restart_request as rr

    if argv[:1] == ['--status']:
        print(json.dumps(rr.load_request(), ensure_ascii=False, indent=1))
        return 0
    reason = ' '.join(argv).strip()
    if not reason:
        print('사용: request_restart.py "재시작 이유"')
        return 2
    from slack_sdk import WebClient
    client = WebClient(token=os.environ['SLACK_BOT_TOKEN'])
    req = rr.request_restart(reason, client)
    print(f"요청 {req['id']} ({req['status']}) — 커밋 {len(req['commits'])}개, "
          f"빌드 필요 {'예' if req['build_needed'] else '아니오'}")
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
