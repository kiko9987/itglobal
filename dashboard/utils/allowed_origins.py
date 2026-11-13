# -*- coding: utf-8 -*-
"""이 사이트로 인정하는 출처(Origin) 목록 — 실시간 소켓 CORS 와 /api 출처 확인이 같이 쓴다.

2026-11-13: 소켓 CORS 를 '*' 에서 라이브러리 기본값(요청 Host 기준)으로 바꿨더니, 서버가 Caddy 뒤에 있어
https://pm.itg-aircon.com 을 자기 주소로 못 알아보고 실시간 연결을 전부 거절했다. → 공개 주소를 명시한다.
바꾸려면 .env 의 ALLOWED_ORIGINS (쉼표로 여러 개).
"""
import os

_DEFAULT = (
    'https://pm.itg-aircon.com',
    'http://localhost:5000',
    'http://127.0.0.1:5000',
)


def allowed_origins() -> list:
    raw = os.getenv('ALLOWED_ORIGINS', '').strip()
    vals = [v.strip().rstrip('/') for v in raw.split(',') if v.strip()] if raw else list(_DEFAULT)
    return vals


def allowed_hosts() -> set:
    """출처에서 'scheme://' 를 뗀 host[:port] 집합 (소문자)."""
    out = set()
    for o in allowed_origins():
        out.add(o.split('://', 1)[-1].lower())
    return out
