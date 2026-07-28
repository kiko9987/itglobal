"""국세청 사업자등록 상태조회 클라이언트 (공공데이터포털 odcloud).

거래처 사업자등록번호의 계속/휴업/폐업 상태를 배치 조회해 거래처 탭에 기록.
"우리 데이터가 마지막 발행 시점 스냅샷이라 상대방 등록 변경(특히 폐업/신설)을
못 따라간다"는 문제의 안전망 — 폐업 번호로 세금계산서 발행하는 사고 방지.

API: POST https://api.odcloud.kr/api/nts-businessman/v1/status
  query : serviceKey (공공데이터포털 인증키), returnType=JSON
  body  : {"b_no": ["1234567890", ...]}   # 최대 100개, 하이픈 제외 10자리
  resp  : {"data": [{"b_no","b_stt","b_stt_cd","tax_type","end_dt", ...}]}
    b_stt: '계속사업자' | '휴업자' | '폐업자' | '' (국세청 미등록)
    end_dt: 폐업일 (YYYYMMDD, 폐업자만)

주소·상호·업태·종목 변경은 이 API로 조회 불가 (상태·과세유형만). 참조 메모리:
nts-businessman-api. 키 env: NTS_SERVICE_KEY (Decoding 키 권장).
"""
from __future__ import annotations

import logging
import os
import re
import time
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

_ENDPOINT = 'https://api.odcloud.kr/api/nts-businessman/v1/status'
_BATCH = 100  # API 1회 최대 100건

# b_stt_cd → 사람이 읽는 상태 (b_stt 가 비어오는 케이스 방어용 매핑)
_STT_BY_CD = {'01': '계속사업자', '02': '휴업자', '03': '폐업자'}


def normalize_bno(v: str) -> str:
    """등록번호 → 하이픈 제거 10자리 숫자. 유효하지 않으면 빈 문자열."""
    digits = re.sub(r'\D', '', str(v or ''))
    return digits if len(digits) == 10 else ''


def check_business_status(
    b_nos: List[str], *,
    service_key: Optional[str] = None,
    timeout: int = 15, retry: int = 2,
) -> Dict[str, dict]:
    """등록번호 리스트 → {정규화번호: {b_stt, b_stt_cd, tax_type, end_dt, ...}}.

    키 미설정·전부 실패 시 빈 dict (호출자에서 graceful skip). 조회 안 된 번호는
    결과에 없음 (국세청 미등록 = data 에 b_stt='' 로 오거나 누락).
    """
    key = (service_key or os.getenv('NTS_SERVICE_KEY', '')).strip()
    if not key:
        logger.warning('[NTS] NTS_SERVICE_KEY 미설정 — 상태조회 skip')
        return {}
    try:
        import requests
    except Exception as exc:
        logger.error(f'[NTS] requests 모듈 없음: {exc}')
        return {}

    # 정규화 + dedup (10자리만)
    norm: List[str] = []
    seen = set()
    for v in b_nos:
        n = normalize_bno(v)
        if n and n not in seen:
            seen.add(n)
            norm.append(n)

    out: Dict[str, dict] = {}
    total_batches = (len(norm) + _BATCH - 1) // _BATCH
    for bi in range(0, len(norm), _BATCH):
        chunk = norm[bi:bi + _BATCH]
        idx = bi // _BATCH + 1
        for attempt in range(retry + 1):
            try:
                r = requests.post(
                    _ENDPOINT,
                    params={'serviceKey': key, 'returnType': 'JSON'},
                    json={'b_no': chunk},
                    timeout=timeout,
                )
                if r.status_code != 200:
                    logger.warning(
                        f'[NTS] HTTP {r.status_code} (batch {idx}/{total_batches}): '
                        f'{r.text[:200]}'
                    )
                    if attempt < retry:
                        time.sleep(1.5)
                        continue
                    break
                data = (r.json() or {}).get('data') or []
                for d in data:
                    bno = normalize_bno(d.get('b_no'))
                    if not bno:
                        continue
                    # b_stt 빈값이면 코드로 보정
                    if not (d.get('b_stt') or '').strip():
                        cd = (d.get('b_stt_cd') or '').strip()
                        if cd in _STT_BY_CD:
                            d['b_stt'] = _STT_BY_CD[cd]
                    out[bno] = d
                logger.info(f'[NTS] batch {idx}/{total_batches} 조회 완료 ({len(data)}건)')
                break
            except Exception as exc:
                logger.warning(f'[NTS] 요청 예외 (batch {idx}, try {attempt}): {exc}')
                if attempt < retry:
                    time.sleep(1.5)
                    continue
    logger.info(f'[NTS] 총 {len(out)}/{len(norm)}건 상태 확보')
    return out
