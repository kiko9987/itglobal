# -*- coding: utf-8 -*-
"""서버 재시작 승인 요청 (2026-10-28).

재시작은 관리자 PowerShell 에서 restart_server.ps1 을 돌려야 해서 매번 사장님 손이 갔다
(10-26 하루 4번). 이제는:
  ① Claude 등이 `python scripts/request_restart.py "이유"` 로 요청 → 사장님 DM 에 카드
     [지금 재시작] [취소]
  ② 사장님이 [지금 재시작] → 요청 파일을 'approved' 로
  ③ SYSTEM 권한으로 1분마다 도는 watchdog.ps1 이 approved 를 보고 재시작 → 카드에 결과
Claude 는 요청까지만 하고, 실행은 사장님 승인 + 워치독만 한다.

요청 상태는 워치독(PowerShell)도 읽어야 해서 Redis 가 아니라 파일
dashboard/logs/restart_request.json 하나에 둔다(요청은 한 번에 하나, 대기 중이면 덧붙임).
"""
import json
import os
import re
import subprocess
import time
import uuid
from pathlib import Path
from typing import Optional

PROJECT_ROOT = Path(__file__).resolve().parents[2]
REQUEST_PATH = PROJECT_ROOT / 'dashboard' / 'logs' / 'restart_request.json'
_LOG_DIR = PROJECT_ROOT / 'dashboard' / 'logs'
_SRC_DIR = PROJECT_ROOT / 'dashboard' / 'src'
_BUILD_MANIFEST = PROJECT_ROOT / 'dashboard' / 'static' / 'dist' / '.vite' / 'manifest.json'

PENDING_TTL_SEC = 30 * 60      # 답이 없는 요청은 30분 뒤 만료
ACTIVE_STATUSES = ('pending', 'approved', 'running')
APPROVE_ACTION = 'server_restart_approve'
CANCEL_ACTION = 'server_restart_cancel'


def approver_id() -> str:
    """승인할 수 있는 사람 = SLACK_ADMIN_CHANNEL(사장님 슬랙 사용자 ID)."""
    return os.getenv('SLACK_ADMIN_CHANNEL', '').strip()


# ─────────────────────────────────────────────────────────────
# 요청 파일
# ─────────────────────────────────────────────────────────────
def load_request(path: Path = None) -> Optional[dict]:
    path = path or REQUEST_PATH
    if not path.exists():
        return None
    try:
        # 워치독(PowerShell 5.1)은 BOM 붙은 UTF-8 로 쓴다
        return json.loads(path.read_text(encoding='utf-8-sig'))
    except (ValueError, OSError):
        return None


def save_request(req: dict, path: Path = None) -> None:
    path = path or REQUEST_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix('.tmp')
    tmp.write_text(json.dumps(req, ensure_ascii=False, indent=1), encoding='utf-8')
    os.replace(tmp, path)


def active_request(path: Path = None, now: float = None) -> Optional[dict]:
    """대기·승인·실행 중인 요청. 만료된 대기 요청은 None."""
    req = load_request(path)
    if not req or req.get('status') not in ACTIVE_STATUSES:
        return None
    now = now or time.time()
    if req['status'] == 'pending' and now - req.get('created_at', 0) > PENDING_TTL_SEC:
        return None
    return req


# ─────────────────────────────────────────────────────────────
# 반영 대기 내용 (마지막 기동 이후 커밋, 화면 빌드 필요 여부)
# ─────────────────────────────────────────────────────────────
_SERVING_RE = re.compile(r'^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\.\d+ .*Serving on')


def last_server_start() -> Optional[str]:
    """서버 로그의 마지막 'Serving on' 시각(YYYY-MM-DD HH:MM:SS). 회전된 로그까지 본다."""
    latest = None
    for log in sorted(_LOG_DIR.glob('itglobal_dashboard.log*')):
        try:
            with open(log, encoding='utf-8', errors='replace') as f:
                for line in f:
                    m = _SERVING_RE.match(line)
                    if m and (latest is None or m.group(1) > latest):
                        latest = m.group(1)
        except OSError:
            continue
    return latest


def commits_since(since: Optional[str]) -> list:
    """마지막 기동 이후 main 커밋 '짧은해시 제목' 목록(최신이 위)."""
    if not since:
        return []
    try:
        out = subprocess.run(
            ['git', '-C', str(PROJECT_ROOT), 'log', f'--since={since}', '--format=%h %s'],
            capture_output=True, timeout=15,
        ).stdout.decode('utf-8', 'replace')
    except (OSError, subprocess.SubprocessError):
        return []
    return [line.strip() for line in out.splitlines() if line.strip()]


def frontend_build_needed() -> bool:
    """화면 소스(dashboard/src)가 마지막 빌드보다 새로우면 True — 재시작만으론 반영 안 됨."""
    if not _SRC_DIR.exists():
        return False
    try:
        built = _BUILD_MANIFEST.stat().st_mtime
    except OSError:
        return True
    newest = max((p.stat().st_mtime for p in _SRC_DIR.rglob('*') if p.is_file()), default=0)
    return newest > built


# ─────────────────────────────────────────────────────────────
# 카드
# ─────────────────────────────────────────────────────────────
def _summary_lines(req: dict) -> list:
    lines = [f"• {r}" for r in req.get('reasons', [])]
    commits = req.get('commits') or []
    if commits:
        lines.append(f"\n반영 대기 커밋 {len(commits)}개 (마지막 기동 {req.get('server_started') or '알 수 없음'} 이후)")
        lines += [f"  `{c.split(' ', 1)[0]}` {c.split(' ', 1)[1] if ' ' in c else ''}" for c in commits[:10]]
        if len(commits) > 10:
            lines.append(f"  … 외 {len(commits) - 10}개")
    if req.get('build_needed'):
        lines.append('\n:warning: 화면 소스가 마지막 빌드보다 새롭습니다. 재시작만으로는 화면 변경이 반영되지 않아요(빌드 먼저).')
    return lines


def build_blocks(req: dict) -> list:
    status = req.get('status')
    head = {
        'pending': ':arrows_counterclockwise: *서버 재시작 요청*',
        'approved': ':hourglass_flowing_sand: *재시작 승인됨* — 1분 안에 재시작합니다',
        'running': ':hourglass_flowing_sand: *재시작 중…*',
        'done': ':white_check_mark: *재시작 완료*',
        'failed': ':x: *재시작 실패*',
        'cancelled': ':no_entry_sign: *재시작 취소됨*',
        'expired': ':zzz: *재시작 요청 만료*',
    }.get(status, '*서버 재시작 요청*')
    blocks = [{'type': 'section', 'text': {'type': 'mrkdwn', 'text': head + '\n' + '\n'.join(_summary_lines(req))}}]
    if req.get('result'):
        blocks.append({'type': 'context', 'elements': [{'type': 'mrkdwn', 'text': req['result']}]})
    if status == 'pending':
        blocks.append({'type': 'actions', 'elements': [
            {'type': 'button', 'action_id': APPROVE_ACTION, 'value': req['id'], 'style': 'primary',
             'text': {'type': 'plain_text', 'text': '지금 재시작'}},
            {'type': 'button', 'action_id': CANCEL_ACTION, 'value': req['id'],
             'text': {'type': 'plain_text', 'text': '취소'}},
        ]})
    return blocks


def _card_text(req: dict) -> str:
    return f"서버 재시작 요청 ({req.get('status')})"


def update_card(client, req: dict) -> None:
    if req.get('channel') and req.get('ts'):
        client.chat_update(channel=req['channel'], ts=req['ts'],
                           text=_card_text(req), blocks=build_blocks(req))


# ─────────────────────────────────────────────────────────────
# 요청 · 승인 · 취소
# ─────────────────────────────────────────────────────────────
def request_restart(reason: str, client, path: Path = None, now: float = None) -> dict:
    """재시작 요청. 대기 중인 요청이 있으면 이유를 덧붙이고 같은 카드를 갱신한다."""
    now = now or time.time()
    started = last_server_start()
    req = active_request(path, now)
    if req and req['status'] == 'pending':
        if reason and reason not in req['reasons']:
            req['reasons'].append(reason)
        req['commits'] = commits_since(started)
        req['build_needed'] = frontend_build_needed()
        save_request(req, path)
        update_card(client, req)
        return req
    if req:
        raise RuntimeError(f"이미 {req['status']} 상태인 재시작이 있습니다 — 끝난 뒤 다시 요청하세요.")

    channel = approver_id()
    if not channel:
        raise RuntimeError('SLACK_ADMIN_CHANNEL 미설정 — 승인받을 사람이 없습니다.')
    req = {
        'id': uuid.uuid4().hex[:12], 'status': 'pending', 'created_at': now,
        'reasons': [reason] if reason else [], 'server_started': started,
        'commits': commits_since(started), 'build_needed': frontend_build_needed(),
    }
    resp = client.chat_postMessage(channel=channel, text=_card_text(req), blocks=build_blocks(req))
    req['channel'] = resp['channel']
    req['ts'] = resp['ts']
    save_request(req, path)
    return req


def decide(request_id: str, user_id: str, approve: bool, client, path: Path = None,
           now: float = None) -> str:
    """버튼 처리. 돌려주는 값은 누른 사람에게 보여 줄 안내(빈 문자열이면 정상 처리)."""
    if not user_id or user_id != approver_id():
        return '재시작 승인은 대표님만 할 수 있어요.'
    now = now or time.time()
    req = load_request(path)
    if not req or req.get('id') != request_id:
        return '이미 처리됐거나 바뀐 요청입니다.'
    if req.get('status') != 'pending':
        return f"이미 '{req.get('status')}' 상태인 요청입니다."
    if now - req.get('created_at', 0) > PENDING_TTL_SEC:
        req['status'] = 'expired'
        save_request(req, path)
        update_card(client, req)
        return '30분이 지나 만료된 요청입니다. 필요하면 다시 요청해 주세요.'
    req['status'] = 'approved' if approve else 'cancelled'
    req['decided_at'] = now
    req['decided_by'] = user_id
    if approve:
        req['approved_at'] = now   # 워치독은 승인 후 10분이 지나면 실행하지 않는다
    save_request(req, path)
    update_card(client, req)
    return ''
