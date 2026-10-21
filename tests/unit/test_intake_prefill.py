# -*- coding: utf-8 -*-
"""수금 지정 창 미리 채움 — 나눠 내는 공사 입금자명('301호임대인'·부담자 이름)으로 프로젝트·낸 사람 (2026-10-21)."""
import json
from unittest.mock import patch

from dashboard.services.payer_split import guess_intake_target

ADDR = '인천 부평구 부영로189번길 51 오남프라자 '
G4144 = {'프로젝트 코드': 'G4144-JW', '현장 주소': ADDR + '301호', '사업자명': '-',
         '계약금': 8000000, '중도금': 0, '잔금': 0,
         '계약금 부담자': '임대인 · 임기환 · 10,000,000 · 현금',
         '잔금 부담자': '임차인 · 이동엽(클라이맥스국어학원) · 3,252,000 · 현금'}
G4158 = {'프로젝트 코드': 'G4158-JW', '현장 주소': ADDR + '507호', '계약금': 6424880, '잔금': 6160880,
         '계약금 부담자': '임대인 · 이현주 · 8,031,100 · 계산서',
         '잔금 부담자': '임차인 · 민선홍(서울아산가정의학) · 7,701,100 · 계산서'}
G4159 = {'프로젝트 코드': 'G4159-JW', '현장 주소': ADDR + '508호', '계약금': 1441880, '잔금': 1177880,
         '계약금 부담자': '임대인 · 이현주 · 1,802,350 · 계산서',
         '잔금 부담자': '임차인 · 김지민(다정약국) · 1,472,350 · 계산서'}
G4155 = {'프로젝트 코드': 'G4155-JW', '현장 주소': ADDR + '501~504호', '계약금': 0, '잔금': 0,
         '계약금 부담자': '임대인 · 최영희 · 19,791,200 · 계산서',
         '잔금 부담자': '임차인 · 기종원(위너스골프) · 18,471,200 · 계산서'}
OTHER_301 = {'프로젝트 코드': 'R4000-XX', '현장 주소': '서울 강남구 어딘가 301호', '계약금': 0,
             '계약금 부담자': '-', '잔금 부담자': '-'}   # 나눠 내는 공사 아님 → 후보 아님
RECS = [G4144, G4158, G4159, G4155, OTHER_301]


def test_room_and_role_in_depositor():
    assert guess_intake_target('301호임대인', 2000000, RECS) == ('G4144-JW', '계약금')
    assert guess_intake_target('301호 임차인', 3252000, RECS) == ('G4144-JW', '잔금')
    assert guess_intake_target('504호위너스', 18471200, RECS) == ('G4155-JW', '잔금')   # '501~504호' 주소


def test_room_without_role_uses_exact_remaining():
    assert guess_intake_target('301호', 2000000, RECS) == ('G4144-JW', '계약금')     # 남은 돈 정확히 일치
    assert guess_intake_target('301호', 1000000, RECS) is None                     # 두 사람 모두 가능 → 모름


def test_name_match_across_rooms_needs_exact_remaining():
    # 이현주는 여러 호실 임대인 → 남은 돈과 금액이 같은 호실 하나일 때만
    assert guess_intake_target('이현주약국에어컨', 360470, RECS) == ('G4159-JW', '계약금')
    assert guess_intake_target('이현주', 300000, RECS) is None
    assert guess_intake_target('다정약국.김지민', 294470, RECS) == ('G4159-JW', '잔금')


def test_no_guess_when_unclear():
    assert guess_intake_target('홍길동', 1000000, RECS) is None
    assert guess_intake_target('301호임대인', 5000000, RECS) is None    # 남은 돈(200만)보다 훨씬 큼
    assert guess_intake_target('', 2000000, RECS) is None
    assert guess_intake_target('301호임대인', 0, RECS) is None


def test_open_modal_prefills_project_stage_and_notice():
    from dashboard.blueprints import slack_bot as sb
    sms = '[Web발신]\n농협, 10/21 09:05\n352-****-1682-33\n입금 2,000,000원\n301호임대인'
    with patch('dashboard.services.project_service.get_project_records', return_value=RECS):
        opt, stage, details, guessed = sb._guess_intake_prefill(sms)
        assert guessed and opt['value'] == 'G4144-JW' and stage == '계약금'
        assert opt['text']['text'].startswith('G4144-JW | - | 인천 부평구')
        view = sb._build_payment_intake_view('i1', 'C1', '1.0', sms, selected_project=opt,
                                             project_details=details, stage_value=stage, guessed=guessed)
    txt = json.dumps(view, ensure_ascii=False)
    assert '미리 골랐습니다' in txt and '누가 낸 돈인가요' in txt
    stage_block = next(b for b in view['blocks'] if b.get('block_id') == 'stage')
    assert stage_block['element']['initial_option']['value'] == '계약금'
    assert view['blocks'][2]['element']['initial_option']['value'] == 'G4144-JW'


def test_open_modal_no_prefill_for_ordinary_deposit():
    from dashboard.blueprints import slack_bot as sb
    with patch('dashboard.services.project_service.get_project_records', return_value=RECS):
        assert sb._guess_intake_prefill('입금 1,000,000원\n주식회사아무개') == (None, None, None, False)
