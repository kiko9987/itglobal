# -*- coding: utf-8 -*-
"""채널톡 봇 입력 창 → 리드·카드·스레드 (2026-11-19 L-04606 계기)."""
import sys

sys.path.insert(0, '.')

import pytest

from dashboard.services import channeltalk_form as cf

FORM = {'type': 'custom', 'submittedAt': 1795011170216, 'inputs': [
    {'label': '이름 or 상호', 'bindingKey': 'user.profile.name', 'value': '세하건설'},
    {'label': '연락처', 'bindingKey': 'user.profile.mobileNumber', 'value': '+821075342739'},
    {'label': '설치 장소 (주소)', 'bindingKey': None, 'value': '화성시 동탄'},
    {'label': '문의 내용', 'bindingKey': None, 'value': '1층 상업공간 40평, 2,3층주택 각 30평 신축건물의 냉난방기 설치'},
]}

CARD = ('⠀\n접수번호: `L-04606`\n:채널톡: *새 문의 접수 알림 - 온라인 (채널톡)*\n'
        '--------------------------------------------\n문의시간 : 2026.11.18. 23:11\n'
        '이름 / 상호 : 리본 369\n연락처 : -\n문의 내용 : 📝 견적 문의\n'
        '--------------------------------------------\n\n<https://x|*💬 채팅 열기*>\n⠀')


def test_parse_form():
    v = cf.parse_form(FORM)
    assert v['name'] == '세하건설' and v['phone'] == '+821075342739'
    assert v['address'] == '화성시 동탄' and '신축' in v['inquiry']
    # 빈 창(고객이 안 냄)은 아무것도 없다 — L-04604
    empty = {'type': 'custom', 'inputs': [{'label': '연락처', 'value': None}]}
    assert cf.parse_form(empty) == {}


def test_lead_updates_fill_blanks_and_keep_inflow_marker():
    cur = {'고객명': '리본 369', '고객 연락처': '-', '방문 주소': '-',
           '문의 내용': '[유입:네이버/검색]\n📝 견적 문의'}
    up = cf.lead_updates(cur, cf.parse_form(FORM))
    assert up['고객 연락처'] == '010-7534-2739'
    assert up['고객명'] == '세하건설'            # 채널톡 익명 이름만 바꾼다
    assert up['방문 주소']
    assert up['문의 내용'].startswith('[유입:네이버/검색]')   # 첫 줄 유입 표시 보존
    assert '■ 고객 입력 창' in up['문의 내용'] and '신축' in up['문의 내용']


def test_lead_updates_do_not_overwrite_real_values():
    cur = {'고객명': '김대표', '고객 연락처': '010-1111-2222', '방문 주소': '김포 걸포로 1',
           '문의 내용': '■ 고객 입력 창\n설치 장소: 화성시 동탄\n문의: 1층 상업공간 40평, 2,3층주택 각 30평 신축건물의 냉난방기 설치'}
    up = cf.lead_updates(cur, cf.parse_form(FORM))
    assert up == {}                                 # 이미 있는 값·같은 묶음은 안 건드린다


def test_patch_card_text():
    v = cf.parse_form(FORM)
    t = cf.patch_card_text(CARD, v, '010-7534-2739')
    assert '이름 / 상호 : 세하건설' in t and '연락처 : 010-7534-2739' in t
    assert '설치 장소 : 화성시 동탄' in t and '문의 내용 : 1층 상업공간' in t
    assert '📝 견적 문의' not in t
    assert '채팅 열기' in t and '접수번호: `L-04606`' in t      # 나머지 줄은 그대로
    assert cf.patch_card_text(t, v, '010-7534-2739') == t        # 두 번 해도 같다


def test_thread_note():
    n = cf.thread_note(cf.parse_form(FORM), '010-7534-2739')
    assert '고객이 입력 창에 남긴 내용' in n
    for s in ('세하건설', '010-7534-2739', '화성시 동탄', '신축'):
        assert s in n


def test_apply_once_per_form(monkeypatch):
    calls = {'update': [], 'note': 0}

    class _R:
        d = {}

        def set(self, k, v, nx=False, ex=None):
            if nx and k in self.d:
                return False
            self.d[k] = v
            return True

        def get(self, k):
            return 'L-04606' if k.startswith('channeltalk_chat_lead:') else self.d.get(k)

    r = _R()
    monkeypatch.setattr(cf, '_rc', lambda: r)
    monkeypatch.setattr(cf, 'submitted_forms', lambda cid: [('m1', FORM)])
    import dashboard.services.lead_service as ls
    monkeypatch.setattr(ls, 'get_lead_by_no', lambda no: {'고객명': '리본 369', '고객 연락처': '-'})
    monkeypatch.setattr(ls, 'update_lead', lambda no, up: calls['update'].append((no, up)) or {'success': True})
    monkeypatch.setenv('SLACK_CHANNELTALK_CHANNEL', '')        # 슬랙 부분은 건너뜀
    assert cf.apply_submitted_forms('chat1') == 1
    assert calls['update'] and calls['update'][0][0] == 'L-04606'
    assert cf.apply_submitted_forms('chat1') == 0               # 같은 창은 한 번만


def test_webhook_wires_form_check():
    import inspect
    from dashboard.blueprints import channeltalk as ch
    src = inspect.getsource(ch)
    assert 'schedule_check(' in src
