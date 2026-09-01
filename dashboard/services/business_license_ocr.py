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
    # OCR 라벨 깨짐 fallback (2026-08-10): 개인사업자 '상호'의 '상'이 앞 줄 노이즈에
    # 붙어('상성개' 등) 값 줄이 '호 : 신세계 유통' 만 남는 케이스(신세계유통·더밸런스짐·
    # 꽁지네김밥 관측). 값이 한글/영문으로 시작할 때만 매치(등록번호 등 숫자값 제외).
    # 최하 우선순위 — 정상 패턴 실패 시에만.
    re.compile(r'^\s*호\s*[:：]\s*([가-힣A-Za-z][^\n]*)', re.MULTILINE),
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
    # 꼬리 영문 괄호 제거 (2026-08-10): 법인명이 '한글상호 (English Co., Ltd.)' 형태일 때
    # 영문 병기는 잘라 한글 공식 상호만. 괄호 안에 한글 없고 ASCII 문자 포함 + 닫는 괄호는
    # 선택(OCR 절단 대비). '(주)'·'(지점)' 등 한글 괄호는 보존.
    name = re.sub(r'\s*\([^가-힣)]*[A-Za-z][^가-힣)]*\)?\s*$', '', name).strip()
    # 남은 문장부호 정리
    name = re.sub(r'\s{2,}', ' ', name).strip(' .,:;')
    return name


# 사업자등록번호 (10자리, 하이픈/공백 편차 허용)
_BNO_RE = re.compile(r'(\d{3})\D?(\d{2})\D?(\d{5})')


def _bno_checksum_ok(b: str) -> bool:
    """국세청 사업자등록번호 체크섬 검증 — OCR 오인식(숫자 오독) 방어."""
    if len(b) != 10 or not b.isdigit():
        return False
    w = [1, 3, 7, 1, 3, 7, 1, 3, 5]
    s = sum(int(b[i]) * w[i] for i in range(9))
    s += (int(b[8]) * 5) // 10
    return (10 - (s % 10)) % 10 == int(b[9])


def extract_business_no_from_text(text: str) -> str:
    """OCR 텍스트에서 **체크섬 유효한** 사업자번호 10자리 추출. 없으면 빈 문자열.

    법인등록번호(13자리)·전화·날짜 등 다른 숫자열은 체크섬으로 걸러짐.
    """
    for m in _BNO_RE.finditer(text or ''):
        b = m.group(1) + m.group(2) + m.group(3)
        if _bno_checksum_ok(b):
            return b
    return ''


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


def _ocr_full_text(client, image_bytes: bytes) -> str:
    """Vision OCR full text 추출 — PDF/이미지 자동 분기.

    이미지(JPEG/PNG 등): document_text_detection (이미지 API).
    PDF: batch_annotate_files (파일 API, 인라인 최대 5장). 이미지 API 에 PDF 를 넣으면
      'Bad image data' 로 거부되므로 반드시 파일 API 로 분기 (2026-08-10 G3977-JK).
    """
    from google.cloud import vision
    _head = image_bytes[:8].lstrip()
    is_pdf = image_bytes[:5] == b'%PDF-' or _head[:5] == b'%PDF-'
    if is_pdf:
        input_config = vision.InputConfig(content=image_bytes, mime_type='application/pdf')
        feature = vision.Feature(type_=vision.Feature.Type.DOCUMENT_TEXT_DETECTION)
        # 사업자등록증은 1장 — 첫 페이지만 (inline 최대 5장)
        req = vision.AnnotateFileRequest(
            input_config=input_config, features=[feature], pages=[1],
        )
        resp = client.batch_annotate_files(requests=[req])
        parts = []
        for file_resp in resp.responses:
            for page_resp in file_resp.responses:
                if getattr(page_resp, 'error', None) and page_resp.error.message:
                    logger.warning(f'[LICENSE/OCR] Vision PDF 에러: {page_resp.error.message}')
                    continue
                fta = getattr(page_resp, 'full_text_annotation', None)
                if fta and fta.text:
                    parts.append(fta.text)
        return '\n'.join(parts)
    # 이미지 경로
    image = vision.Image(content=image_bytes)
    response = client.document_text_detection(image=image, image_context={'language_hints': ['ko']})
    if response.error and response.error.message:
        logger.warning(f'[LICENSE/OCR] Vision API 에러: {response.error.message}')
        return ''
    return response.full_text_annotation.text if response.full_text_annotation else ''


# 신용/체크카드 이미지 오첨부 감지 키워드 (2026-08-10 G3879-JSH — 사업자등록증 대신
# 카드 사진 첨부). OCR 텍스트에 하나라도 있으면 카드 이미지로 추정 (안내 구체화용).
_CARD_KEYWORDS = (
    '신용카드', '체크카드', '카드번호', '후불교통', '유효기간',
    '삼성카드', '신한카드', '현대카드', '국민카드', '롯데카드',
    '우리카드', '하나카드', '비씨카드', '농협카드',
    'american express', 'visa', 'mastercard', 'master card',
    'cardmember', 'valid thru',
)


def analyze_business_license(image_bytes: bytes) -> dict:
    """사업자등록증 OCR 종합 분석 — 이름·사업자번호·카드여부·텍스트유무.

    반환: {'name': str, 'bno': str, 'is_card': bool, 'has_text': bool}
      - name/bno 둘 다 없음 = 사업자등록증이 아닐 가능성(오첨부·판독불가) → 호출부 경고.
      - is_card: 카드 이미지 키워드 감지 → 안내 문구 구체화.
    Vision 은 1회만 호출 (ocr_business_license 도 이걸 재사용).
    """
    result = {'name': '', 'bno': '', 'is_card': False, 'has_text': False,
              'is_license': False, 'doc_negative': False}
    if not image_bytes:
        return result
    try:
        from google.cloud import vision
    except ImportError:
        logger.warning('[LICENSE/OCR] google-cloud-vision 미설치 — skip')
        return result
    try:
        client = vision.ImageAnnotatorClient()
        # PDF/이미지 자동 분기 (PDF 는 파일 API, 2026-08-10 'Bad image data' fix)
        full_text = _ocr_full_text(client, image_bytes)
    except Exception as exc:
        logger.warning(f'[LICENSE/OCR] Vision API 호출 실패: {type(exc).__name__}: {exc}')
        return result

    result['has_text'] = bool((full_text or '').strip())
    _low = (full_text or '').lower()
    result['is_card'] = any(k in _low for k in _CARD_KEYWORDS)

    # 문서 종류 판별 — '사업자등록증'/'고유번호증' 제목이 있고 세금계산서/정산문서가 아닐 때만
    #   진짜 사업자등록증으로 인정. 세금계산서·지출품의서는 사업자번호·상호가 있어도 여기서
    #   걸러 canonical 오저장(덮어쓰기) 방지. (2026-09-01 R3883-SJ 계기)
    _norm = re.sub(r'\s+', '', full_text or '')
    result['doc_negative'] = any(n in _norm for n in (
        '세금계산서', '계산서', '지출품의서', '품의서', '견적서', '거래명세',
        '지불의뢰', '입금표', '거래처원장'))
    result['is_license'] = (
        any(p in _norm for p in ('사업자등록증', '고유번호증'))
        and not result['doc_negative'])

    # ① 사업자번호(숫자, OCR 90%+체크섬) → 거래처 탭 역조회로 정답 상호.
    #    한글 상호 오인식(예: 미덕원→이억원) 회피. 번호 미매칭·신규는 ②로.
    bno = extract_business_no_from_text(full_text)
    result['bno'] = bno
    name = ''
    if bno:
        try:
            from dashboard.services.partner_status_sync import get_partner_name_by_bno
            looked = get_partner_name_by_bno(bno)
            if looked:
                # 거래처 탭 상호도 꼬리 영문 괄호 정리 (2026-08-10) — '한글(English…' 통일
                name = _clean_name(looked) or looked
                logger.info(f'[LICENSE/OCR] 사업자번호 {bno} → 거래처 탭 상호: {name!r}')
        except Exception as exc:
            logger.debug(f'[LICENSE/OCR] 번호 역조회 실패 (상호 OCR 로 fallback): {exc}')
    # ② 상호 정규식 추출 (거래처 탭에 없는 신규 사업자 or 번호 매칭 실패)
    if not name:
        name = extract_business_name_from_text(full_text)
        if name:
            logger.info(f'[LICENSE/OCR] 법인명·상호 추출 성공: {name!r}')
    if not name:
        logger.info(
            f'[LICENSE/OCR] 법인명·상호 매치 실패 '
            f'(bno={bno or "없음"}, card={result["is_card"]}, {len(full_text or "")}자)'
        )
    result['name'] = name
    return result


def ocr_business_license(image_bytes: bytes) -> str:
    """이미지/PDF 바이트 → 법인명·상호 추출 (name 만, 하위호환 wrapper)."""
    return analyze_business_license(image_bytes).get('name', '')
