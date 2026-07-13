"""사업자등록증 OCR — Google Cloud Vision API 로 법인명·상호 추출.

용도: 사업자등록증 이미지/PDF 첨부 후 자동으로 법인명(단체명) 또는 상호를 뽑아
프로젝트 시트 사업자명 필드에 pre-fill.

무료 티어 첫 1,000 units/월. 계산서 요청 대비 매우 넉넉.
실패 시 조용히 skip — 매니저가 계산서 모달에서 수동 입력 가능.
"""
from __future__ import annotations

import logging
import os
import re
from typing import Optional

logger = logging.getLogger(__name__)

# 사업자등록증 라벨 패턴 — 법인 또는 개인 사업자용.
#   법인용: '법 인 명 (단 체 명) : (주)크리스아이티'
#   개인용: '상      호 : 홍길동상회'
# OCR 결과는 공백·개행이 편차 있으므로 유연 매치.
_NAME_PATTERNS = [
    # 법인명(단체명) : XXX  ← 우선순위 높음
    re.compile(r'법\s*인\s*명[^:\n]*[:：]\s*([^\n]+)', re.MULTILINE),
    # 상호 : XXX  (개인 사업자)
    re.compile(r'(?<!단체)상\s*호\s*[:：]\s*([^\n]+)', re.MULTILINE),
]

# 추출된 값 뒤에 붙은 다른 라벨/노이즈 제거 패턴
_TRAILING_LABEL_RE = re.compile(
    r'\s*(?:대\s*표\s*자|성\s*명|주\s*민|등\s*록\s*번\s*호|법\s*인\s*등\s*록|'
    r'개\s*업\s*연\s*월\s*일|사\s*업\s*장|생\s*년\s*월\s*일|주\s*소).*$'
)


def _clean_name(raw: str) -> str:
    """추출된 원문에서 뒤따르는 다른 라벨·괄호 노이즈 제거."""
    name = raw.strip()
    # 뒤 라벨 자름 (예: '(주)크리스아이티 대표자 홍길동' → '(주)크리스아이티')
    name = _TRAILING_LABEL_RE.sub('', name).strip()
    # 남은 문장부호 정리
    name = re.sub(r'\s{2,}', ' ', name).strip(' .,:;')
    return name


def extract_business_name_from_text(text: str) -> str:
    """OCR 결과 텍스트에서 법인명/상호 추출. 실패 시 빈 문자열."""
    if not text:
        return ''
    for pat in _NAME_PATTERNS:
        m = pat.search(text)
        if not m:
            continue
        name = _clean_name(m.group(1))
        if len(name) >= 2 and len(name) <= 40:
            return name
    return ''


def ocr_business_license(image_bytes: bytes) -> str:
    """이미지/PDF 바이트 → Vision API OCR → 법인명·상호 추출.

    반환: 추출된 이름 or 빈 문자열.
    """
    if not image_bytes:
        return ''
    try:
        from google.cloud import vision
    except ImportError:
        logger.warning('[LICENSE/OCR] google-cloud-vision 미설치 — skip')
        return ''

    try:
        client = vision.ImageAnnotatorClient()
        image = vision.Image(content=image_bytes)
        # document_text_detection: 문서 이미지에 최적화 (양식·표 잘 인식)
        response = client.document_text_detection(image=image, image_context={'language_hints': ['ko']})
        if response.error and response.error.message:
            logger.warning(f'[LICENSE/OCR] Vision API 에러: {response.error.message}')
            return ''
        full_text = response.full_text_annotation.text if response.full_text_annotation else ''
        name = extract_business_name_from_text(full_text)
        if name:
            logger.info(f'[LICENSE/OCR] 법인명·상호 추출 성공: {name!r}')
        else:
            logger.info('[LICENSE/OCR] 법인명·상호 매치 실패 (OCR 텍스트 확인 필요)')
        return name
    except Exception as exc:
        logger.warning(f'[LICENSE/OCR] Vision API 호출 실패: {type(exc).__name__}: {exc}')
        return ''
