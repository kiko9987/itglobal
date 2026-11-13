# -*- coding: utf-8 -*-
"""2026-11-13 보안 점검 2차 — 한 줄 수정 회귀 테스트.

① SMS 이름 칸(입금자·출금 상대) 슬랙 표기 이스케이프 — 반환 후보·완료 카드, 입금 경고 DM
② 리드 카드 제목의 '[유입:…]' 값 이스케이프
③ 사업자명(sync_biz_cell) 수식 방지 — 계산서 OCR 부담자 이름
④ 전화·방문 워크플로 동기화 병합 쓰기 수식 방지
"""
import sys
from datetime import datetime

sys.path.insert(0, '.')

import dashboard.blueprints.slack_bot as sb
from dashboard.services import lead_sync as ls
from dashboard.services import payer_split as ps
from dashboard.services import refund_match as rm

EVIL = '<!channel>'


# ── ① SMS 이름 칸 ──────────────────────────────────────────

def _c(partner):
    return {'code': 'R4163-TH', 'stage': '계약금', 'amount': 200000, 'stage_value': 200000,
            'paid': 2700000, 'total2': 2500000, 'unpaid_before': 0, 'unpaid_after': 0,
            'reason': 'x', 'kind': 'overpay', 'name_typo': True,
            'dep_date': '2026-10-03', 'dep_bank': '하나', 'dep_partner': partner}


def _out(partner):
    return {'id': 'o1', 'text': '출금200,000원', 'amount': 200000, 'partner': partner,
            'bank': '하나', 'acct_code': 'R', 'ts': int(datetime(2026, 10, 8, 14, 30).timestamp())}


def test_refund_cards_escape_partner_names():
    body = rm.build_candidate_blocks('o1', _out(EVIL), _c(EVIL))[0]['text']['text']
    done = rm.build_done_blocks(_out(EVIL), _c(EVIL), 200000, 0, 'MY')[0]['text']['text']
    for t in (body, done):
        assert EVIL not in t and '&lt;!channel&gt;' in t
    assert '이름 확인 : 입금 &lt;!channel&gt; / 출금 &lt;!channel&gt;' in body


def test_refund_cards_plain_names_unchanged():
    body = rm.build_candidate_blocks('o1', _out('김보라(오라라(O'), _c('김보라(오라라(O'))[0]['text']['text']
    assert '>입금 : 10/03  하나  200,000원  김보라(오라라(O' in body


def test_intake_warn_dm_escapes_depositor(monkeypatch):
    monkeypatch.setattr(sb, '_resolve_manager_initial', lambda u: 'MJ')
    t = sb._build_intake_warn_dm([('G4139-MJ', '잔금', 650000, [])], 'U', True,
                                 deposit={'partner': '<https://evil.example|확인>', 'date_md': '09/30'})
    assert '<https://evil.example' not in t
    assert '입금 : 650,000원 · &lt;https://evil.example|확인&gt; · 09/30' in t


# ── ② 리드 카드 제목 ───────────────────────────────────────

def test_lead_title_escapes_typed_inflow_marker():
    blocks, _ = ls.build_inquiry_blocks({'고객명': 'x', '문의 내용': '[유입:<!channel>]\n견적 문의'},
                                        'L-00001', '홈페이지')
    title = blocks[0]['text']['text'].split('\n')[1]
    assert '<!channel>' not in title and '&lt;!channel&gt;' in title


# ── ③ 사업자명 수식 방지 ───────────────────────────────────

class _BizMgr:
    def __init__(self, biz, ver='3'):
        self.cells = {'F10': biz, 'AT10': ver}
        self.writes = []

    def get_cell_value(self, sid, sn, a):
        return self.cells.get(a, '')

    def update_cell_value(self, sid, sn, a, v):
        self.writes.append((a, v))
        # 시트가 앞 ' 를 떼고 보여 주는 동작 흉내
        self.cells[a] = v[1:] if v.startswith("'") else v
        return True


def test_sync_biz_cell_quotes_formula_name():
    m = _BizMgr('-')
    plan = {'계약금': {'name': '=HYPERLINK("https://x","a")', 'stage': '계약금'}}
    out = ps.sync_biz_cell(m, 's', 'n', 10, {'사업자명': 'F', '_version': 'AT'}, plan)
    assert m.writes[0] == ('F10', "'" + ps.split_biz_value(plan))
    assert out['사업자명'] == ps.split_biz_value(plan)          # 돌려주는 값은 원래 글
    # 다시 돌려도 같은 값으로 보여 또 쓰지 않음(무한 갱신 없음)
    assert ps.sync_biz_cell(m, 's', 'n', 10, {'사업자명': 'F', '_version': 'AT'}, plan) == {}


# ── ④ 워크플로 병합 쓰기 ───────────────────────────────────

class _Exec:
    def __init__(self, sink, body):
        sink.append(body)

    def execute(self):
        return {}


class _Values:
    def __init__(self):
        self.bodies = []

    def batchUpdate(self, spreadsheetId, body):
        return _Exec(self.bodies, body)


class _Mgr:
    def __init__(self):
        self.v = _Values()
        sheets = type('S', (), {'values': lambda _s: self.v})()
        self.service = type('Svc', (), {'spreadsheets': lambda _s: sheets})()


def test_workflow_merge_quotes_formula_values():
    m = _Mgr()
    ls._update_existing_lead_from_workflow(
        m, {'sheet_id': 's', 'sheet_name': '리드'}, 12,
        {'고객명': '=IMAGE("https://x")', '방문 주소': '서울 강남구 테헤란로 1', '고객 연락처': '010-1234-5678'})
    vals = {d['range'].split('!')[1]: d['values'][0][0] for d in m.v.bodies[0]['data']}
    assert m.v.bodies[0]['valueInputOption'] == 'USER_ENTERED'
    assert vals['H12'] == "'=IMAGE(\"https://x\")"
    assert vals['I12'] == '서울 강남구 테헤란로 1' and vals['F12'] == '010-1234-5678'
