# -*- coding: utf-8 -*-
"""유선 상담 재통화 분류(A~E) + 리마인드 A/B 수집·렌더 회귀 테스트 (2026-09-15).

규칙 기반 5분류: A=곧 재통화 / B=추후 팔로업 / C=먼 미래 / D=조건부 / E=단순 종결.
메시지엔 A·B만 카테고리 구분 노출.
"""
import sys
sys.path.insert(0, '.')

from datetime import date

from dashboard.services import absent_remind as ar
from dashboard.services.absent_remind import classify_consult_memo as clf


# ── 분류기 ──────────────────────────────────────────────
def test_classify_A_soon_callback():
    assert clf('바쁘셔서 추후 다시 연락주기로함.') == 'A'
    assert clf('미팅 시간 조율 때문에 다시 통화 예정') == 'A'
    assert clf('스탠드 견적 문의 / 내일 연락 다시 줄 예정') == 'A'


def test_classify_B_followup():
    assert clf('신규설치할지 내부 협의후 다시 연락준다함.') == 'B'
    assert clf('건물 계약 후 연락달라고 요청함') == 'B'
    assert clf('임대 계약 후 전화 다시 받기로 함.') == 'B'   # 명사+다시+동사 패턴


def test_classify_C_far_future():
    assert clf('1등급 지원금 희망. 내년에 설치나고 다시 연락준다 함.') == 'C'
    assert clf('지원사업 나오면 다시 연락주신다고 합니다.') == 'C'


def test_classify_D_conditional():
    assert clf('설치 필요시 다시 연락달라함.') == 'D'
    assert clf('제품 설치가 필요할때 다시 문의 달라고 상담') == 'D'


def test_classify_E_closed():
    assert clf('단순 제품 단가문의') == 'E'
    assert clf('2019년에 설치한 천장형 제품 1대. 매입 불가 상담') == 'E'
    assert clf('다른업체 선정') == 'E'
    assert clf('') == 'E'


# ── 수집 (A/B만, 오늘·E 제외) ────────────────────────────
def _lead(dt, status, memo='', consultant='c', sales=''):
    return {'리드 No': f'L-{abs(hash((dt, memo))) % 90000 + 1000}',
            '상담 시간': dt.strftime('%Y.%m.%d. 10:00'), '상태': status,
            '상담 내용': memo, '온라인 상담자': consultant, '영업 담당자': sales,
            '고객명': 't', '고객 연락처': '010-0000-0000', '플랫폼': '전화'}


def test_collect_callback_ab(monkeypatch):
    today = date.today()
    prev = ar._previous_business_day(today)
    leads = [
        _lead(prev, '유선 상담', '바빠서 다시 통화 예정'),        # A
        _lead(prev, '유선 상담', '내부 협의후 다시 연락준다함'),  # B
        _lead(prev, '유선 상담', '내년에 다시 연락준다 함'),      # C → 제외
        _lead(prev, '유선 상담', '필요시 다시 연락달라함'),        # D → 제외
        _lead(prev, '유선 상담', '단순 단가 문의'),               # E → 제외
        _lead(today, '유선 상담', '바빠서 다시 통화 예정'),        # 오늘 → 제외
    ]
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'get_lead_records', lambda: leads)
    _u, _r, _q, callback = ar.collect_absent_leads()
    assert len(callback['A']) == 1
    assert len(callback['B']) == 1
    # C·D·E·오늘 은 어디에도 안 들어감
    assert '단순' not in str(callback)


# ── 렌더 ────────────────────────────────────────────────
def test_build_text_shows_ab_sections():
    prev = ar._previous_business_day(date.today())
    cb = {
        'A': [_lead(prev, '유선 상담', '바빠서 다시 통화 예정')],
        'B': [_lead(prev, '유선 상담', '계약 후 다시 연락주기로 함')],
    }
    text, total = ar.build_remind_text([], {}, client=None, callback=cb)
    assert '재통화 예정 (1건)' in text
    assert '추후 팔로업 (1건)' in text
    assert total == 2
