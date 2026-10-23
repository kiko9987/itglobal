"""발행된 세금계산서 OCR — 공급받는자(등록번호·상호·성명·사업장·업태·종목·이메일) 읽기 (2026-10-23).

계산서 스레드에 경영지원이 올린 세금계산서에는 홈택스에 있는 거래처 정보가 다 들어 있다. 이것을
거래처 탭에 넣어 두면, 홈택스엔 있지만 거래처 탭(예전 내보내기)엔 없던 거래처를 다음 계산서부터
알아본다 (G4158 서울아산가정의학과의원 — '부담자 등록증을 찾지 못했습니다' 경고 계기).

글자만 이어 읽으면 표 칸 순서가 섞이고 줄바꿈된 상호가 끊긴다('서울아산가정의학과의' / '원').
그래서 Vision 단어 위치로 읽는다: 홈택스 계산서는 왼쪽이 공급자, 오른쪽이 공급받는자이고
라벨(상호·성명·사업장·업태·종목·이메일) 오른쪽 같은 줄(칸 안 두 줄 포함)이 값이다.
"""
from __future__ import annotations

import logging
import re
from statistics import median
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)

_BNO_RE = re.compile(r'(\d{3})\D?(\d{2})\D?(\d{5})')
_EMAIL_RE = re.compile(r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+')
_LABELS = ('상호', '성명', '사업장', '업태', '종목', '이메일')
# 값에 섞이면 안 되는 표 라벨 조각
_LABEL_TOKENS = {'상호', '성명', '사업장', '업태', '종목', '이메일', '등록', '번호', '등번', '종사업장',
                 '공급받는자', '공급받는', '자', '공', '공급자'}


def ocr_words(content: bytes) -> List[dict]:
    """Vision 문서 OCR → 단어 목록 [{text, x0, y0, x1, y1}] (좌표는 쪽 크기 기준 0~1). PDF는 첫 장."""
    from google.cloud import vision
    client = vision.ImageAnnotatorClient()
    head = content[:8].lstrip()
    pages = []
    if content[:5] == b'%PDF-' or head[:5] == b'%PDF-':
        req = vision.AnnotateFileRequest(
            input_config=vision.InputConfig(content=content, mime_type='application/pdf'),
            features=[vision.Feature(type_=vision.Feature.Type.DOCUMENT_TEXT_DETECTION)], pages=[1])
        for fr in client.batch_annotate_files(requests=[req]).responses:
            for pr in fr.responses:
                if pr.full_text_annotation:
                    pages += list(pr.full_text_annotation.pages)
    else:
        resp = client.document_text_detection(image=vision.Image(content=content),
                                               image_context={'language_hints': ['ko']})
        if resp.error and resp.error.message:
            logger.warning(f'[INVOICE/OCR] Vision 에러: {resp.error.message}')
            return []
        if resp.full_text_annotation:
            pages = list(resp.full_text_annotation.pages)
    if not pages:
        return []
    page = pages[0]
    w, h = float(page.width or 1), float(page.height or 1)
    out = []
    for block in page.blocks:
        for para in block.paragraphs:
            for word in para.words:
                text = ''.join(s.text for s in word.symbols)
                bb = word.bounding_box
                if bb.normalized_vertices:
                    xs = [v.x for v in bb.normalized_vertices]
                    ys = [v.y for v in bb.normalized_vertices]
                else:
                    xs = [v.x / w for v in bb.vertices]
                    ys = [v.y / h for v in bb.vertices]
                if text.strip():
                    out.append({'text': text, 'x0': min(xs), 'x1': max(xs), 'y0': min(ys), 'y1': max(ys)})
    return out


def _bno_ok(b: str) -> bool:
    from dashboard.services.business_license_ocr import _bno_checksum_ok
    return _bno_checksum_ok(b)


def _yc(w: dict) -> float:
    return (w['y0'] + w['y1']) / 2


def _join(words: List[dict]) -> str:
    """칸 안 단어들 → 글. 같은 줄은 간격 보고 띄우고, 줄바꿈은 한글-한글·숫자 앞 쉼표면 붙인다."""
    if not words:
        return ''
    hs = [w['y1'] - w['y0'] for w in words]
    tol = max(median(hs) * 0.5, 1e-4)
    lines: List[List[dict]] = []
    for w in sorted(words, key=_yc):
        if lines and abs(_yc(w) - _yc(lines[-1][-1])) <= tol:
            lines[-1].append(w)
        else:
            lines.append([w])
    texts = []
    for ln in lines:
        ln.sort(key=lambda w: w['x0'])
        lh = median([w['y1'] - w['y0'] for w in ln])
        s = ln[0]['text']
        for a, b in zip(ln, ln[1:]):
            # 띄어쓰기 간격은 글자 높이의 0.19~0.47, 붙은 조각('걸포'+'로', '134-84'+',')은 0.14 이하
            gap = b['x0'] - a['x1']
            glue = '' if (gap < lh * 0.17 or re.match(r'^[,.)\]%(]', b['text']) or re.search(r'[(\[]$', a['text'])) else ' '
            s += glue + b['text']
        texts.append(s)
    out = texts[0]
    for t in texts[1:]:
        if re.search(r'[가-힣]$', out) and re.match(r'^[가-힣]', t):
            out += t                       # 칸이 좁아 줄바꿈된 낱말 ('…가정의학과의' + '원')
        elif out.endswith(',') and re.match(r'^\d', t):
            out += t                       # '505,506,' + '509호'
        else:
            out += ' ' + t
    return re.sub(r'\s{2,}', ' ', out).strip()


def parse_invoice_buyer(words: List[dict]) -> Dict[str, str]:
    """단어 위치 → 공급받는자 정보 {bno, name, ceo, addr, biz_type, biz_item, email}. 못 읽은 칸은 ''.

    라벨이 공급자(왼쪽)·공급받는자(오른쪽) 두 번 나오면 오른쪽 것이 공급받는자. 값은 라벨 오른쪽,
    같은 줄 다음 라벨(상호→성명, 업태→종목) 전까지, 위아래로는 줄 간격 절반 안(칸 안 두 줄).
    """
    res = {'bno': '', 'name': '', 'ceo': '', 'addr': '', 'biz_type': '', 'biz_item': '', 'email': '',
           'supplier_bno': ''}
    if not words:
        return res
    # 표 머리 영역: '작성일자' 줄 위
    top_limit = min([_yc(w) for w in words if w['text'].replace(' ', '') in ('작성일자', '작성일')] or [1.0])
    head = [w for w in words if _yc(w) < top_limit]
    lab = {k: sorted([w for w in head if w['text'] == k], key=lambda w: w['x0']) for k in _LABELS}
    if len(lab['상호']) < 2:
        return res                          # 공급자·공급받는자 상호 라벨 둘 다 보여야 홈택스 계산서 양식
    buyer_x = lab['상호'][-1]['x0'] - 0.03  # 공급받는자 영역 왼쪽 경계 (등록번호 라벨 '등록'은 상호 라벨 줄과 같은 x)
    supplier_right = lab['상호'][0]['x1']
    if buyer_x <= supplier_right:
        return res
    b = {}
    for k, ws in lab.items():
        right = [w for w in ws if w['x0'] >= buyer_x]
        if right:
            b[k] = min(right, key=_yc) if k == '이메일' else right[-1] if k != '사업장' else right[0]
    rows = sorted(_yc(b[k]) for k in ('상호', '사업장', '업태', '이메일') if k in b)
    gaps = [y2 - y1 for y1, y2 in zip(rows, rows[1:]) if y2 > y1]
    row_h = median(gaps) if gaps else 0.05
    band = row_h * 0.55

    def value(k, stop=None):
        L = b.get(k)
        if not L:
            return ''
        x_end = stop['x0'] - 0.002 if stop else 1.01
        vs = [w for w in head if w['x0'] >= L['x1'] - 0.002 and w['x1'] <= x_end + 0.002
              and abs(_yc(w) - _yc(L)) <= band and w is not L and w['text'] not in _LABEL_TOKENS]
        return _join(vs)

    res['name'] = value('상호', b.get('성명') if b.get('성명') and _yc(b['성명']) - _yc(b['상호']) < band else None)
    res['ceo'] = value('성명')
    res['addr'] = value('사업장')
    res['biz_type'] = value('업태', b.get('종목') if b.get('종목') and abs(_yc(b['종목']) - _yc(b['업태'])) < band else None)
    res['biz_item'] = value('종목')
    # 등록번호: 공급받는자 영역의 체크섬 맞는 번호 (공급자 번호는 왼쪽 — 같은 번호면 잘못 읽은 것)
    def first_bno(ws):
        for w in sorted(ws, key=_yc):
            m = _BNO_RE.search(w['text'])
            if m and _bno_ok(''.join(m.groups())):
                return f'{m.group(1)}-{m.group(2)}-{m.group(3)}'
        return ''
    res['bno'] = first_bno([w for w in head if w['x0'] >= buyer_x])
    res['supplier_bno'] = first_bno([w for w in head if w['x1'] < buyer_x])
    for w in sorted([w for w in head if w['x0'] >= buyer_x], key=_yc):
        m = _EMAIL_RE.search(w['text'])
        if m:
            res['email'] = m.group(0).lower()
            break
    # 이메일 칸 값이 라벨 줄 값으로도 잡혔으면 정리
    for k in ('addr', 'biz_type', 'biz_item', 'ceo', 'name'):
        res[k] = _EMAIL_RE.sub('', res[k]).strip(' ,')
    # 상호 꼬리 영문 괄호('뉴타입베이크클럽 (New-type Bake club)')는 등록증 OCR과 같은 규칙으로 뗌
    from dashboard.services.business_license_ocr import _clean_name
    res['name'] = _clean_name(res['name']) or res['name']
    return res


def read_invoice_buyer(content: bytes) -> Optional[Dict[str, str]]:
    """세금계산서 파일 → 공급받는자 정보. 홈택스 계산서 양식이 아니거나 등록번호·상호를 못 읽으면 None."""
    try:
        words = ocr_words(content)
    except Exception as exc:
        logger.warning(f'[INVOICE/OCR] Vision 호출 실패: {type(exc).__name__}: {exc}')
        return None
    text = ''.join(w['text'] for w in words)
    if not any(k in text for k in ('공급받는자', '공급가액', '승인번호')):
        return None
    r = parse_invoice_buyer(words)
    if not (r['bno'] and r['name']) or r['bno'] == r.get('supplier_bno'):
        return None
    return r
