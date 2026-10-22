# -*- coding: utf-8 -*-
"""슬랙 문서 보기 — HWP·PDF 등을 kordoc 으로 글로 바꿔 스레드에 답글 (2026-10-22).

관공서 공고·과업지시서는 HWP 로 오는데 슬랙 모바일은 HWP 미리보기가 안 돼
현장 매니저가 열어 볼 방법이 없었다.
- 리드·방문 카드 스레드 등 채널에 HWP·HWPX 를 올리면 자동으로 내용을 답글.
- 봇 DM(또는 전용 채널 SLACK_DOC_VIEW_CHANNEL)에 올리면 PDF·엑셀 등 모든 형식 변환.

- 변환: tools/kordoc 의 kordoc CLI(Node) 를 subprocess 로 호출. 설치는 그 폴더에서 `npm ci`.
- 외부 통신 차단(KORDOC_OFFLINE=1)·접근 폴더를 임시 폴더로 제한(KORDOC_ROOT)·OCR 끔.
- 표는 슬랙이 그리지 못하므로 줄마다 '• 칸 · 칸 · 칸' 으로 풀어 쓴다.
- 긴 문서는 앞부분만 메시지로, 전체는 .txt 파일로 같은 스레드에 첨부.
"""
import json
import logging
import os
import re
import shutil
import subprocess
import tempfile
import threading
import urllib.request
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)

SUPPORTED_EXTS = {'hwp', 'hwpx', 'pdf', 'docx', 'xlsx', 'xls', 'pptx'}
_MAX_BYTES = 30 * 1024 * 1024        # 슬랙 다운로드·변환 상한 30MB
_CONVERT_TIMEOUT = 120               # kordoc 한 건 상한(초)
_MESSAGE_LIMIT = 3500                # 스레드 본문 메시지 글자 수 — 넘으면 .txt 첨부
_PROJECT_ROOT = Path(__file__).resolve().parents[2]
_KORDOC_CLI = _PROJECT_ROOT / 'tools' / 'kordoc' / 'node_modules' / 'kordoc' / 'dist' / 'cli.js'

# 변환은 Node 프로세스 하나에 수백 MB 를 쓸 수 있어 동시 2건으로 묶는다.
_convert_slots = threading.BoundedSemaphore(2)


def file_ext(name: str) -> str:
    return name.rsplit('.', 1)[-1].lower() if '.' in (name or '') else ''


def is_supported(name: str) -> bool:
    return file_ext(name) in SUPPORTED_EXTS


def _node_path() -> Optional[str]:
    """서비스(LocalSystem) PATH 에 node 가 없을 수 있어 기본 설치 경로도 본다."""
    explicit = os.getenv('KORDOC_NODE', '').strip()
    if explicit:
        return explicit
    found = shutil.which('node')
    if found:
        return found
    default = Path(r'C:\Program Files\nodejs\node.exe')
    return str(default) if default.exists() else None


class DocConvertError(Exception):
    """매니저에게 그대로 보여 줄 수 있는 한국어 사유."""


def convert_to_markdown(data: bytes, ext: str) -> str:
    """파일 바이트 → kordoc 마크다운. 실패는 DocConvertError(한국어 사유)."""
    node = _node_path()
    if not node or not _KORDOC_CLI.exists():
        logger.error(f'[DOCVIEW] kordoc 런타임 없음 node={node} cli={_KORDOC_CLI.exists()}')
        raise DocConvertError('변환 도구가 서버에 설치되어 있지 않습니다. 관리자에게 알려 주세요.')

    with _convert_slots, tempfile.TemporaryDirectory(prefix='docview_') as tmp:
        src = Path(tmp) / f'input.{ext}'
        out = Path(tmp) / 'output.md'
        src.write_bytes(data)
        env = dict(os.environ, KORDOC_OFFLINE='1', KORDOC_ROOT=tmp)
        try:
            proc = subprocess.run(
                [node, str(_KORDOC_CLI), str(src), '-o', str(out), '--table-format', 'gfm'],
                cwd=tmp, env=env, capture_output=True, timeout=_CONVERT_TIMEOUT,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
        except subprocess.TimeoutExpired:
            raise DocConvertError(f'변환이 {_CONVERT_TIMEOUT}초 안에 끝나지 않았습니다. 파일이 너무 크거나 복잡합니다.')

        if proc.returncode != 0 or not out.exists():
            reason = _kordoc_error(proc.stdout.decode('utf-8', 'replace'))
            logger.warning(f'[DOCVIEW] kordoc 실패 rc={proc.returncode} reason={reason!r} '
                           f'stderr={proc.stderr.decode("utf-8", "replace")[-300:]!r}')
            raise DocConvertError(reason or '문서를 읽지 못했습니다.')
        return out.read_text(encoding='utf-8')


def _kordoc_error(stdout: str) -> str:
    """kordoc 은 실패 시 stdout 에 {"success": false, "error": "…"} JSON 을 낸다."""
    start = stdout.find('{')
    if start < 0:
        return ''
    try:
        return str(json.loads(stdout[start:]).get('error') or '')
    except (ValueError, AttributeError):
        return ''


# ─────────────────────────────────────────────────────────────
# 마크다운 → 슬랙 글
# ─────────────────────────────────────────────────────────────
_TAG_RE = re.compile(r'<[^>]+>')
_IMAGE_RE = re.compile(r'^!\[[^\]]*\]\([^)]*\)$')
_SEPARATOR_RE = re.compile(r'^\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)*\|?$')


def _clean_inline(text: str) -> str:
    text = re.sub(r'<br\s*/?>', ' ', text, flags=re.I)
    text = _TAG_RE.sub('', text)
    text = text.replace('\\*', '*').replace('\\_', '_')
    text = re.sub(r'\s+', ' ', text).strip()
    # 슬랙 mrkdwn 제어문자 escape
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def _split_row(line: str) -> list:
    inner = line.strip()
    if inner.startswith('|'):
        inner = inner[1:]
    if inner.endswith('|'):
        inner = inner[:-1]
    return [_clean_inline(c) for c in inner.split('|')]


def markdown_to_slack(md: str) -> str:
    """kordoc 마크다운을 슬랙에서 읽기 좋은 글로.

    - 제목(#) → 굵게, 표 → 행마다 '• 칸 · 칸'(빈 칸 생략), 3칸 이상 표의 머리행은 굵게.
    - 이미지 줄·':' 만 남은 줄(PDF 머리글 상자의 잔재)·HTML 태그·빈 줄 반복 제거.
    """
    lines = md.replace('\r\n', '\n').split('\n')
    out = []
    i = 0
    while i < len(lines):
        raw = lines[i].rstrip()
        stripped = raw.strip()

        if stripped.startswith('|'):
            block = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                block.append(lines[i].strip())
                i += 1
            out.extend(_table_to_lines(block))
            out.append('')
            continue

        i += 1
        if not stripped or stripped == ':' or _IMAGE_RE.match(stripped):
            out.append('')
            continue
        if stripped.startswith('<!--'):
            continue
        heading = re.match(r'^#{1,6}\s+(.*)$', stripped)
        if heading:
            text = _clean_inline(heading.group(1))
            if text:
                out.append(f'*{text}*')
            continue
        text = _clean_inline(stripped)
        if text:
            out.append(text)

    # 빈 줄 반복 정리
    result = re.sub(r'\n{3,}', '\n\n', '\n'.join(out)).strip()
    return result


def _table_to_lines(block: list) -> list:
    rows = [_split_row(line) for line in block if not _SEPARATOR_RE.match(line)]
    if not rows:
        return []
    lines = []
    width = max(len(r) for r in rows)
    for idx, cells in enumerate(rows):
        filled = [c for c in cells if c]
        if not filled:
            continue
        if idx == 0 and width >= 3:
            lines.append('*' + ' · '.join(filled) + '*')
        else:
            lines.append('• ' + ' · '.join(filled))
    return lines


def split_for_message(text: str, limit: int = _MESSAGE_LIMIT) -> tuple:
    """(메시지에 넣을 앞부분, 잘렸는지). 문단 경계에서 자른다."""
    if len(text) <= limit:
        return text, False
    cut = text.rfind('\n', 0, limit)
    if cut < limit // 2:
        cut = limit
    return text[:cut].rstrip(), True


# ─────────────────────────────────────────────────────────────
# 슬랙 연결
# ─────────────────────────────────────────────────────────────
def _download(url: str, token: str) -> bytes:
    req = urllib.request.Request(url, headers={'Authorization': f'Bearer {token}'})
    with urllib.request.urlopen(req, timeout=45) as r:
        data = r.read(_MAX_BYTES + 1)
    if len(data) > _MAX_BYTES:
        raise DocConvertError(f'파일이 {_MAX_BYTES // 1024 // 1024}MB 보다 커서 변환하지 않았습니다.')
    # 슬랙이 권한 없는 요청에 로그인 HTML 을 200 으로 돌려주는 경우 방어
    if data[:15].lower().startswith((b'<!doctype html', b'<html')):
        raise DocConvertError('슬랙에서 파일을 받지 못했습니다(권한 문제). 관리자에게 알려 주세요.')
    return data


def _claim(file_id: str, ts: str) -> bool:
    """같은 파일·메시지 중복 처리 방지(슬랙 재전송·이벤트 중복). Redis 장애 시엔 그냥 진행."""
    try:
        from dashboard.utils.redis_client import get_redis_client
        return bool(get_redis_client().redis.set(f'docview:{file_id}:{ts}', '1', nx=True, ex=86400))
    except Exception as exc:
        logger.warning(f'[DOCVIEW] dedup Redis 실패(진행): {exc}')
        return True


def _react(client, channel: str, ts: str, add: str = '', remove: str = '') -> None:
    for fn, name in ((client.reactions_remove, remove), (client.reactions_add, add)):
        if not name:
            continue
        try:
            fn(channel=channel, timestamp=ts, name=name)
        except Exception:
            pass


HWP_EXTS = {'hwp', 'hwpx'}


def doc_view_mode(event: dict) -> Optional[str]:
    """문서 보기 대상이면 'all'(모든 형식) / 'hwp'(HWP·HWPX만), 아니면 None.

    - DM·전용 채널(SLACK_DOC_VIEW_CHANNEL) 최상위 글: 'all' — 일부러 보내 온 파일.
    - 그 밖의 채널(리드·방문 카드 스레드 등): HWP·HWPX 가 있을 때만 'hwp'. PDF 는 슬랙이
      미리보기를 해 주고 우리 견적서 PDF 가 스레드에 자주 올라와, 변환하면 답글만 쌓인다.
    """
    if event.get('bot_id') or event.get('subtype') not in ('file_share', None, ''):
        return None
    files = event.get('files') or []
    if not files:
        return None
    if event.get('channel_type') == 'im':
        return 'all'
    thread_ts = event.get('thread_ts')
    top_level = not thread_ts or thread_ts == event.get('ts')
    target = os.getenv('SLACK_DOC_VIEW_CHANNEL', '').strip()
    if target and event.get('channel') == target and top_level:
        return 'all'
    if any(file_ext(f.get('name') or '') in HWP_EXTS for f in files):
        return 'hwp'
    return None


DM_GUIDE = (
    ':page_facing_up: HWP·PDF·워드·엑셀·PPT 파일을 여기로 보내 주시면 내용을 글로 바꿔 답해 드려요.\n'
    '리드·방문 카드 스레드에 올린 HWP 파일은 따로 보내지 않아도 자동으로 바꿔 드립니다.'
)


def start_doc_view(event: dict, client, token: str) -> Optional[str]:
    """대상이면 백그라운드로 변환을 시작하고 mode 를, 아니면 None 을 돌려준다.

    슬랙 핸들러(process_before_response)가 바로 반환하도록 변환은 스레드로 넘긴다.
    같은 채널에 봇이 둘 있어도 _claim(파일·메시지당 1회)으로 답글은 한 번만 달린다.
    """
    mode = doc_view_mode(event)
    if not mode:
        return None

    def _bg():
        try:
            handle_doc_view(event, client, token, mode)
        except Exception as exc:
            logger.error(f'[DOCVIEW] 처리 실패: {exc}', exc_info=True)
    threading.Thread(target=_bg, daemon=True).start()
    return mode


def handle_doc_view(event: dict, client, token: str, mode: str = 'all') -> None:
    """파일마다 변환해 그 메시지 스레드에 답글. 백그라운드 스레드에서 호출.

    답글은 원 메시지가 스레드 답글이면 그 스레드에, 최상위 글이면 그 글의 스레드에 단다.
    """
    channel = event.get('channel')
    ts = event.get('ts')
    reply_ts = event.get('thread_ts') or ts
    files = event.get('files') or []
    allowed = SUPPORTED_EXTS if mode == 'all' else HWP_EXTS
    targets = [f for f in files if file_ext(f.get('name') or '') in allowed]
    if not targets:
        if mode != 'all':
            return
        names = ', '.join(f.get('name') or '?' for f in files)
        client.chat_postMessage(
            channel=channel, thread_ts=reply_ts,
            text=f':information_source: {names} — 글로 바꿀 수 있는 파일은 HWP·HWPX·PDF·워드·엑셀·PPT 입니다.',
        )
        return

    for f in targets:
        name = f.get('name') or '문서'
        if not _claim(f.get('id') or name, ts):
            logger.info(f'[DOCVIEW] 이미 처리 중/완료 — skip {name}')
            continue
        _react(client, channel, ts, add='hourglass_flowing_sand')
        try:
            url = f.get('url_private_download') or f.get('url_private')
            if not url:
                raise DocConvertError('파일 주소를 받지 못했습니다.')
            md = convert_to_markdown(_download(url, token), file_ext(name))
            body = markdown_to_slack(md)
            if not body:
                raise DocConvertError('글자를 찾지 못했습니다. 스캔본(사진으로 된 PDF)은 아직 읽지 못합니다.')
            _post_result(client, channel, reply_ts, name, body)
            _react(client, channel, ts, add='white_check_mark', remove='hourglass_flowing_sand')
            logger.info(f'[DOCVIEW] 완료 {name} chars={len(body)}')
        except DocConvertError as exc:
            client.chat_postMessage(channel=channel, thread_ts=reply_ts, text=f':x: *{name}* — {exc}')
            _react(client, channel, ts, add='x', remove='hourglass_flowing_sand')
        except Exception as exc:
            logger.error(f'[DOCVIEW] 예외 {name}: {exc}', exc_info=True)
            client.chat_postMessage(
                channel=channel, thread_ts=reply_ts,
                text=f':x: *{name}* — 변환 중 오류가 났습니다. 관리자에게 알려 주세요.',
            )
            _react(client, channel, ts, add='x', remove='hourglass_flowing_sand')


def _post_result(client, channel: str, ts: str, name: str, body: str) -> None:
    head, truncated = split_for_message(body)
    title = f':page_facing_up: *{_clean_inline(name)}* 내용'
    if truncated:
        title += ' (앞부분 · 전체는 아래 파일)'
    client.chat_postMessage(
        channel=channel, thread_ts=ts, text=f'{title}\n\n{head}',
        unfurl_links=False, unfurl_media=False,
    )
    if truncated:
        stem = name.rsplit('.', 1)[0]
        plain = body.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&')
        client.files_upload_v2(
            channel=channel, thread_ts=ts,
            filename=f'{stem}_내용.txt', title=f'{stem} 전체 내용',
            content=plain,
        )
