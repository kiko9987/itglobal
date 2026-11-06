# -*- coding: utf-8 -*-
"""세콤매니저(근태) REST 전송 받기 — 같은 PC(세콤 PC)의 세콤매니저 → 이 서버.

세콤매니저 「파일 > ERP 설정 > REST 서버」의 URL 을 http://127.0.0.1:5000/secom 으로 두면
세콤이 출입·출퇴근 기록을 보낸다(원래는 에스원 S1ESP 클라우드로 가는 기능).
  · v1/secom/manager/alarm        — 태그 한 번마다 {at, sabun, type(0출입/1출근/2외출/3복귀/4퇴근)}
  · v1/secom/manager/workhistory  — 하루 요약 {date, sabun, ws_at(출근), wc_at(퇴근)}
「접속 확인」 때 세콤이 무엇을 묻는지 아직 모름(2026-11-06) → 1단계는 받은 요청을 전부 그대로
파일로 남기고 성공으로 답한다. 모양을 본 뒤 저장·슬랙 알림을 붙인다.

보안: 로그인 없는 기계 요청이라 같은 PC 안에서 온 것만 받는다 — 접속 주소가 127.0.0.1 이고
프록시(Caddy) 머리말이 없을 때만. 바깥에서 홈페이지 주소로 들어오면 Caddy 를 거쳐
X-Forwarded-For 가 붙으므로 거부된다.
"""

import json
import os
import time

from flask import Blueprint, current_app, jsonify, request

from dashboard.utils.logging_config import get_logger

logger = get_logger(__name__)

secom_bp = Blueprint('secom_inbound', __name__, url_prefix='/secom')

_LOOPBACK = ('127.0.0.1', '::1', 'localhost')
_PROXY_HEADERS = ('X-Forwarded-For', 'X-Forwarded-Host', 'X-Real-IP', 'Via', 'Forwarded')


def is_local_request() -> bool:
    if (request.remote_addr or '') not in _LOOPBACK:
        return False
    return not any(request.headers.get(h) for h in _PROXY_HEADERS)


def capture_dir() -> str:
    d = os.path.join(current_app.instance_path, 'secom')
    os.makedirs(d, exist_ok=True)
    return d


def _capture(sub: str) -> None:
    """받은 요청 한 건을 그대로 한 줄(JSON)로 — 날짜별 파일."""
    rec = {
        'ts': time.strftime('%Y-%m-%d %H:%M:%S'),
        'method': request.method,
        'path': sub,
        'query': request.query_string.decode('utf-8', 'replace'),
        'headers': {k: v for k, v in request.headers.items()},
        'body': request.get_data(as_text=True)[:200000],
    }
    path = os.path.join(capture_dir(), time.strftime('capture_%Y%m%d.jsonl'))
    with open(path, 'a', encoding='utf-8') as f:
        f.write(json.dumps(rec, ensure_ascii=False) + '\n')


@secom_bp.route('/', defaults={'sub': ''}, methods=['GET', 'POST', 'PUT', 'PATCH', 'DELETE'])
@secom_bp.route('/<path:sub>', methods=['GET', 'POST', 'PUT', 'PATCH', 'DELETE'])
def secom_any(sub: str):
    if not is_local_request():
        logger.warning(f'[SECOM] 바깥 요청 거부 (ip={request.remote_addr}, path=/{sub})')
        return jsonify({'error': 'forbidden'}), 403
    try:
        _capture(sub)
    except Exception as exc:
        logger.error(f'[SECOM] 요청 기록 실패 (/{sub}): {exc}', exc_info=True)
    logger.info(f'[SECOM] {request.method} /{sub} ({request.content_length or 0}B)')
    # 무엇을 기대하는지 모르는 단계 — 흔한 성공 모양을 두루 담아 답한다
    return jsonify({'result': True, 'success': True, 'code': 0, 'status': 'OK', 'message': 'OK'}), 200
