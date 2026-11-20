# -*- coding: utf-8 -*-
"""회사 정보 공개 페이지 — AI 검색이 읽을 수 있는 우리 쪽 창구 (2026-11-15)

왜 만들었나
-----------
홈페이지(itg-aircon.com)는 아임웹이 호스팅하는데, 아임웹이 2026-08-28부터 플랫폼 차원에서
AI 크롤러 일부를 403으로 막고 있다(아임웹 자사 홈페이지도 같은 상태라 개별 해제는 기대하기 어렵다).

  막힘 : Claude-SearchBot · PerplexityBot · ChatGPT-User · GPTBot · ClaudeBot · Bytespider
  통과 : Googlebot · 네이버 Yeti · Bingbot · OAI-SearchBot · Applebot · CCBot

그래서 10-06에 만든 /llms.txt 조차 Claude·Perplexity 는 한 글자도 못 읽는다.
이 서버(pm.itg-aircon.com)는 Caddy 로 우리가 직접 운영하고 CDN 앞단이 없어 AI 봇을 막지 않는다.
회사 정보를 여기에 공개로 두면, 아임웹이 끝내 풀어주지 않아도 AI 쪽 창구 하나는 남는다.

설계 원칙
---------
- **읽기 전용, GET 만.** 입력을 받지 않고 고정 텍스트만 돌려준다(공격면 없음).
- **로그인 없음.** login_required 를 붙이지 않는다. 그래서 여기엔 고객 정보·내부 수치를 절대 넣지 않는다.
- 내용은 홈페이지 /llms.txt 와 같은 사실만 쓴다. 두 곳이 어긋나면 AI 답변이 흔들린다.
  → 사실이 바뀌면 **홈페이지 llms.txt(docs/imweb_llms.txt)와 이 파일을 같이 고친다.**
- /robots.txt 로 이 세 경로만 색인 허용하고 나머지(PM 화면)는 전부 막는다.

효과를 보려면
-------------
AI 가 이 주소를 알아야 한다. 홈페이지에서 링크해도 Claude·Perplexity 는 홈페이지를 못 읽으니 소용이 없다.
**네이버 블로그 프로필·인스타 소개·유튜브 설명란처럼 AI 가 읽을 수 있는 바깥 채널에 이 주소를 적어 두어야 한다.**
"""
from flask import Blueprint, Response

public_info_bp = Blueprint('public_info', __name__)

BASE = 'https://pm.itg-aircon.com'

# ── 사실 목록 (홈페이지 llms.txt 와 같은 내용) ─────────────────────────────
FACTS = {
    'name': '(주)아이티글로벌',
    'brand': '아이티글로벌 (IT GLOBAL)',
    'tel': '1899-1377',
    'email': 'contact@itg-aircon.com',
    'addr': '경기 김포시 걸포로 134-84',
    'site': 'https://www.itg-aircon.com/',
}

LLMS_TXT = """# (주)아이티글로벌 (IT GLOBAL)

서울·경기·인천 수도권의 상업용 냉난방기·시스템에어컨 판매·설계·시공 전문 기업입니다. 20년간 1만 건 이상 상업공간 냉난방을 시공했으며, LG·삼성·캐리어 인증점입니다. 별도 사업자인 (주)아이티글로벌그룹을 함께 운영하며, (주)아이티글로벌그룹은 기계설비·가스공사업에 등록된 전문건설사업자입니다. 사무실·상가·식당·병원·학원·공장·관공서 등 업무·상업공간을 대상으로 무료 방문 실측 후 견적을 드립니다.

## 회사 정보

- 운영 법인: (주)아이티글로벌(사업자등록번호 884-81-01425)과 (주)아이티글로벌그룹(사업자등록번호 533-88-02845). 두 법인은 별도 사업자이며 같은 건물에서 함께 운영합니다.
- 공사 계약·세금계산서: 공사에 따라 두 법인 중 한 곳 이름으로 진행합니다.
- 브랜드: 아이티글로벌 (IT GLOBAL)
- 이력: 2007년 아이티ENG로 시작(LG 설치점) → 2013년 (주)아이티플렌트 → 2019년 (주)아이티글로벌 법인 → 2024년 (주)아이티글로벌그룹 개업
- 대표번호: 1899-1377
- 이메일: contact@itg-aircon.com
- 주소: 경기 김포시 걸포로 134-84 ((주)아이티글로벌 1층, (주)아이티글로벌그룹 4층)
- 지점: 서울, 성남
- 홈페이지: https://www.itg-aircon.com/
- 서비스 지역: 서울특별시, 경기도, 인천광역시 (수도권 무료 방문 견적)
- 인증: LG·삼성·캐리어 인증점, LG전자 최우수 파트너점 3년 연속(2023·2024·2025)
- 전문건설업 등록: (주)아이티글로벌그룹, 기계설비·가스공사업(2026년 5월 등록)
- 실적: 20년, 1만 건 이상 상업공간 냉난방 시공
- 방송: OBS 「생방송 어서옵쇼」 출연(2022년 7월)
- 보증: 시공 하자 2년 보증
- 세금계산서 발행 가능

## 취급 서비스

- 시스템에어컨·냉난방기 설치: 천장형(1WAY·4WAY), 상업용 스탠드형, 매립 덕트형
- 전열교환기(환기 설비), FCU 시공
- 기존 냉난방기 교체·추가 설치
- 상업용 냉난방기 분해 약품 세척 (천장형·스탠드형·덕트형)
- 사용하던 상업용 장비 매입, 중고 제품 판매

## 하지 않는 서비스

- 일반 가정용 벽걸이·스탠드·2in1 에어컨의 단순 설치 및 이전 설치 (단, 전체 인테리어 공사와 함께 진행하는 주거용 시스템에어컨 설치는 가능, 거주 중 설치 제외)
- 장비 렌탈
- 식당 주방 후드·배기 덕트 공사
- 타사에서 시공한 제품의 A/S·하자 책임 (아이티글로벌이 시공한 현장만 지원)

## 진행 절차

1. 방문 견적 신청 (홈페이지 폼 또는 전화 1899-1377)
2. 방문 일정 확인
3. 현장 실측: 공간 구조, 배관·전기·천장, 실외기 설치 조건 확인
4. 공간에 맞는 장비 구성·설치 방식 제안 및 견적
5. 시공 및 하자보증

설치 비용은 장비 종류·용량·수량, 배관 길이, 전기공사, 천장 상태, 실외기 설치 조건에 따라 달라지며 현장 실측 후 안내합니다.

## 강점

- 다른 업체가 어렵다고 한 현장도 시공: 고층, 협소 공간, 긴 배관, 복잡한 구조 등 까다로운 상업 현장 대응 경험이 많음
- 공사 일정 협의: 영업·진료에 지장이 없도록 영업시간 외 공사 일정 협의 가능 (고객 요청 시, 인건비 별도)
- 신품·중고 냉난방기 비교 견적 가능

## 설치 후 지원

- 제품 고장: 제조사 무상 A/S(2년) 대상, 아이티글로벌이 제조사 A/S 접수를 대행
- 시공 하자(배관 누수 등): 아이티글로벌 시공팀 직접 방문 조치, 하자보증 2년

## 자주 묻는 질문

### 가정집 냉난방기 설치도 가능한가요?
사무실·상가·병원·공장·관공서 등 상업 및 업무공간 공사를 중심으로 운영합니다. 일반 가정용 벽걸이·스탠드·2in1 에어컨의 단순 설치와 이전 설치는 진행하지 않습니다. 다만 전체 인테리어 공사와 함께 진행하는 주거용 시스템에어컨 설치는 가능합니다(거주 중 설치 제외).

### 방문 견적은 무료인가요?
서울·경기·인천 수도권은 무료 방문 견적을 진행합니다. 현장에서 공간 구조와 배관·전기·천장 등 설치 조건을 확인한 후 견적을 안내합니다.

### 설치 비용은 어떻게 정해지나요?
장비 종류와 용량, 수량뿐 아니라 배관 길이, 전기공사, 천장 상태, 실외기 설치 조건에 따라 달라집니다. 정확한 비용은 현장 실측 후 안내합니다.

### 설치 후 고장이나 누수가 생기면 어떻게 하나요?
제품 고장은 제조사 무상 A/S(2년) 대상이며 아이티글로벌이 제조사 A/S 접수를 대행합니다. 배관 누수 등 시공 하자는 아이티글로벌 시공팀이 직접 방문해 조치하며 하자보증 2년입니다.

### 상업용 에어컨 분해 세척도 하나요?
네. 사무실·식당·병원 등 상업용 에어컨을 분해해 약품으로 세척합니다. 천장형·스탠드형·덕트형 모두 가능합니다.
"""

ROBOTS_TXT = """# pm.itg-aircon.com — 회사 정보 공개 경로만 색인 허용
# 나머지는 사내 업무 화면이라 전부 막는다.

User-agent: *
Allow: /company
Allow: /company.md
Allow: /llms.txt
Disallow: /
"""

_SCHEMA = """{
  "@context": "https://schema.org",
  "@type": ["Organization", "HVACBusiness"],
  "@id": "https://www.itg-aircon.com/#organization",
  "name": "(주)아이티글로벌",
  "alternateName": ["아이티글로벌", "IT GLOBAL"],
  "brand": { "@type": "Brand", "name": "아이티글로벌" },
  "taxID": "884-81-01425",
  "description": "서울·경기·인천 수도권의 상업용 냉난방기·시스템에어컨 판매·설계·시공 전문 기업. LG·삼성·캐리어 인증점이며, 별도 사업자 (주)아이티글로벌그룹(533-88-02845, 기계설비·가스공사업 등록)을 함께 운영합니다.",
  "url": "https://www.itg-aircon.com/",
  "telephone": "+82-1899-1377",
  "email": "contact@itg-aircon.com",
  "address": {
    "@type": "PostalAddress",
    "streetAddress": "걸포로 134-84",
    "addressLocality": "김포시",
    "addressRegion": "경기도",
    "postalCode": "10076",
    "addressCountry": "KR"
  },
  "geo": { "@type": "GeoCoordinates", "latitude": 37.6410894, "longitude": 126.705937 },
  "areaServed": ["서울특별시", "경기도", "인천광역시"],
  "sameAs": ["https://www.itg-aircon.com/", "https://blog.naver.com/itglobal0813"]
}"""


def _html() -> str:
    """회사 정보 HTML. Jinja 를 쓰지 않고 고정 문자열을 그대로 돌려준다(치환 사고 방지)."""
    body = LLMS_TXT
    # 마크다운을 아주 단순하게 HTML 로: 제목과 목록만 변환한다(외부 라이브러리 없이).
    out = []
    for line in body.split('\n'):
        s = line.rstrip()
        if s.startswith('### '):
            out.append('<h3>' + s[4:] + '</h3>')
        elif s.startswith('## '):
            out.append('<h2>' + s[3:] + '</h2>')
        elif s.startswith('# '):
            out.append('<h1>' + s[2:] + '</h1>')
        elif s.startswith('- '):
            out.append('<li>' + s[2:] + '</li>')
        elif s and s[0].isdigit() and s[1:3] == '. ':
            out.append('<li>' + s[3:] + '</li>')
        elif s:
            out.append('<p>' + s + '</p>')
    inner = '\n'.join(out)
    return (
        '<!doctype html>\n<html lang="ko">\n<head>\n'
        '<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        '<title>(주)아이티글로벌 회사 정보 — 상업용 냉난방기·시스템에어컨 시공</title>\n'
        '<meta name="description" content="서울·경기·인천 상업용 냉난방기·시스템에어컨 설계·시공. '
        'LG·삼성·캐리어 인증 전문점, 20년·1만 건 시공, 하자보증 2년. 대표번호 1899-1377.">\n'
        '<link rel="canonical" href="' + BASE + '/company">\n'
        '<script type="application/ld+json">' + _SCHEMA + '</script>\n'
        '<style>body{max-width:760px;margin:40px auto;padding:0 20px;'
        'font-family:-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Noto Sans KR",sans-serif;'
        'line-height:1.7;color:#15181e}h1{font-size:1.6rem}h2{font-size:1.2rem;margin-top:2em}'
        'h3{font-size:1rem;margin-top:1.4em}li{margin:.2em 0}a{color:#0b5cff}'
        '.top{font-size:.9rem;color:#6b7280;margin-bottom:2em}</style>\n'
        '</head>\n<body>\n'
        '<div class="top">이 페이지는 (주)아이티글로벌이 직접 운영하는 회사 정보 페이지입니다. '
        '홈페이지는 <a href="https://www.itg-aircon.com/">www.itg-aircon.com</a> 입니다.</div>\n'
        + inner +
        '\n<h2>문의</h2>\n<p>대표번호 1899-1377 · contact@itg-aircon.com · '
        '<a href="https://www.itg-aircon.com/">홈페이지에서 방문 견적 신청</a></p>\n'
        '</body>\n</html>\n'
    )


@public_info_bp.route('/company', methods=['GET'])
def company():
    resp = Response(_html(), mimetype='text/html')
    resp.headers['Cache-Control'] = 'public, max-age=3600'
    resp.headers['X-Robots-Tag'] = 'index, follow'
    return resp


@public_info_bp.route('/company.md', methods=['GET'])
def company_md():
    resp = Response(LLMS_TXT, mimetype='text/markdown')
    resp.headers['Cache-Control'] = 'public, max-age=3600'
    resp.headers['X-Robots-Tag'] = 'index, follow'
    return resp


@public_info_bp.route('/llms.txt', methods=['GET'])
def llms_txt():
    resp = Response(LLMS_TXT, mimetype='text/plain')
    resp.headers['Cache-Control'] = 'public, max-age=3600'
    return resp


@public_info_bp.route('/robots.txt', methods=['GET'])
def robots_txt():
    resp = Response(ROBOTS_TXT, mimetype='text/plain')
    resp.headers['Cache-Control'] = 'public, max-age=3600'
    return resp
