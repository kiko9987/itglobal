# -*- coding: utf-8 -*-
"""입출금 장부(11-26) — 은행 문자 → 장부(잔액 포함), 중복 한 줄, 입금 지정 적요, 손 입력·잔액 맞추기."""
import sys
from datetime import datetime

sys.path.insert(0, '.')

import pytest

from dashboard.services import bank_ledger as bl

IBK = "[Web발신]\n2026/11/26 09:12\n입금 2,018,800원\n잔액 310,869,202원\n국민카드\n452***38801011\n기업"
HANA_OUT = "[Web발신]\n하나,11/26, 10:05\n255******31304\n출금897,303원\n일산전기조명\n잔액240,963,932원"
NH_IN = "[Web발신]\n농협 입금7,450,600원\n11/26 09:15 352-****-1682-33 이재철(푸른 잔액82,517,408원"
PERSONAL = "[Web발신]\n2026/11/26 09:12\n입금 50,000원\n잔액 1,000원\n홍길동\n123***45678901\n기업"
NOW = datetime(2026, 11, 26, 12, 0)


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(bl, '_db', lambda: d)
    monkeypatch.setattr(bl, '_table_ready', False)
    yield d


def _acct(code):
    return next(a for a in bl.accounts() if a['code'] == code)


def test_parse_when_formats():
    assert bl.parse_when('2026/08/10 14:49', NOW) == '2026-08-10 14:49'
    assert bl.parse_when('하나,08/07, 15:33', NOW) == '2026-08-07 15:33'
    assert bl.parse_when('11/26 09:15 352-', NOW) == '2026-11-26 09:15'
    assert bl.parse_when('12/31 23:50', datetime(2027, 1, 2)) == '2026-12-31 23:50'    # 1월에 받은 작년 12월 문자
    assert bl.parse_when('글자만', NOW) == '2026-11-26 12:00'


def test_sms_to_ledger_with_balance(db):
    bl.record_from_text(IBK, 'sms:jw')
    bl.record_from_text(HANA_OUT, 'sms:jw')
    bl.record_from_text(NH_IN, 'sms:jw')
    rows = {r['acct_code']: r for r in bl.list_tx('2026-11-26', '2026-11-26')}
    assert rows['1011']['direction'] == 'in' and rows['1011']['amount'] == 2_018_800 and rows['1011']['balance'] == 310_869_202
    assert rows['1011']['counterparty'] == '국민카드' and rows['1011']['at'] == '2026-11-26 09:12'
    assert rows['9012']['direction'] == 'out' and rows['9012']['amount'] == 897_303 and rows['9012']['balance'] == 240_963_932
    assert rows['농협']['amount'] == 7_450_600 and rows['농협']['balance'] == 82_517_408
    assert all('잔액' not in (r['raw'] or '') for r in rows.values())        # 잔액은 숫자 칸에만, 원문엔 없음


def test_same_sms_from_three_phones_is_one_row(db):
    for dev in ('sms:jw', 'sms:yg', 'channel:U1'):
        bl.record_from_text(IBK.replace('[Web발신]', '[Web발신] '), dev)
    assert len(bl.list_tx('2026-11-26', '2026-11-26')) == 1


def test_personal_account_not_recorded(db):
    bl.record_from_text(PERSONAL, 'sms:jw')
    assert bl.list_tx('2026-11-26', '2026-11-26') == []


def test_designation_sets_memo_and_split(db):
    from dashboard.services.sms_intake import dedup_hash
    bl.record_from_text(IBK, 'sms:jw')
    uid = dedup_hash('x', IBK)
    bl.designate(uid, 'G4160-JW', '중도금')
    r = bl.list_tx('2026-11-26', '2026-11-26')[0]
    assert r['memo'] == 'G4160-JW 매출수금 중도금' and r['project_code'] == 'G4160-JW' and r['title'] == '매출수금'
    bl.designate(uid, 'G4160-JW', '중도금')                                   # 다시 눌러도 한 번
    bl.designate(uid, 'R4323-TH', '잔금', 500_000)
    r = bl.list_tx('2026-11-26', '2026-11-26')[0]
    assert r['memo'] == 'G4160-JW 매출수금 중도금 / R4323-TH 매출수금 잔금 500,000'
    assert r['project_code'] == 'G4160-JW,R4323-TH'
    assert bl.list_tx('2026-11-26', '2026-11-26', unassigned=True) == []


def test_manual_entry_balance_and_summary(db):
    woori = _acct('1013')['id']
    bl.add_manual({'account_id': woori, 'at': '2026-11-26 08:00', 'direction': 'bal', 'balance': '63,520,584'}, 'SB', True)
    bl.add_manual({'account_id': woori, 'at': '2026-11-26 11:00', 'direction': 'out', 'amount': '1,000',
                   'balance': '63,519,584', 'counterparty': '경기남부도로', 'title': '여비교통비'}, 'SB', True)
    s = {a['code']: a for a in bl.summary('2026-11-26')}
    assert s['1013']['balance'] == 63_519_584 and s['1013']['day_out'] == 1_000
    assert s['1011']['balance'] is None
    with pytest.raises(PermissionError):
        bl.add_manual({'account_id': woori, 'at': '2026-11-26 08:00', 'direction': 'in', 'amount': 1}, 'x', False)
    with pytest.raises(bl.LedgerError):
        bl.add_manual({'account_id': woori, 'at': '2026-11-26 08:00', 'direction': 'in', 'amount': ''}, 'SB', True)
    with pytest.raises(bl.LedgerError):
        bl.add_manual({'account_id': woori, 'at': '어제', 'direction': 'in', 'amount': 5}, 'SB', True)


def test_sms_rows_keep_bank_values(db):
    bl.record_from_text(IBK, 'sms:jw')
    r = bl.list_tx('2026-11-26', '2026-11-26')[0]
    bl.update(r['id'], {'amount': 1, 'memo': 'G4160-JW 매출수금 중', 'title': '거래처'}, 'SB', True)
    r = bl.list_tx('2026-11-26', '2026-11-26')[0]
    assert r['amount'] == 2_018_800 and r['memo'] == 'G4160-JW 매출수금 중' and r['title'] == '거래처'
    with pytest.raises(bl.LedgerError):
        bl.delete(r['id'], 'SB', True)                                       # 은행 문자 줄은 못 지움
    m = bl.add_manual({'account_id': r['account_id'], 'at': '2026-11-26 13:00', 'direction': 'in', 'amount': 5}, 'SB', True)
    bl.delete(m, 'SB', True)
    assert len(bl.list_tx('2026-11-26', '2026-11-26')) == 1


def test_report_balances_only_for_no_sms_accounts(db):
    """SB 자금일보 잔액 → 문자 안 오는 자금일보 계좌만(11-26). 같은 날 다시 돌면 고치고, 그대로면 줄을 안 늘림."""
    bals = {'1011': 306_638_537, '1012': 91_420_704, '1013': 63_520_584, '9011': 0, '9012': 247_053_932}
    bl.record_sms(HANA_OUT, bl.OUT, 'h1', 'sms', NOW)                       # 9012 는 은행 문자 잔액이 이미 있음
    assert bl.sync_report_balances('2026-11-25', bals) == 4                 # 1012·1013·9011 + 1011(문자 잔액 없어 시작점)
    s = {a['code']: a for a in bl.summary('2026-11-25')}
    assert s['1012']['balance'] == 91_420_704 and s['9011']['balance'] == 0 and s['1011']['balance'] == 306_638_537
    assert s['9012']['balance'] is None                                     # 문자(11/26) 계좌는 자금일보로 안 덮음
    bl.record_sms(IBK, bl.IN, 'i1', 'sms', NOW)                             # 1011 문자가 오면 그 뒤론 문자 잔액
    assert bl.sync_report_balances('2026-11-26', {'1011': 1}) == 0
    assert bl.sync_report_balances('2026-11-25', bals) == 0                 # 다시 돌려도 그대로
    assert bl.sync_report_balances('2026-11-25', dict(bals, **{'1013': 63_500_000})) == 1   # SB 가 고친 값
    assert bl.summary('2026-11-25')[2]['balance'] == 63_500_000
    assert bl.sync_report_balances('2026-11-26', dict(bals, **{'1013': 63_500_000})) == 0   # 다음 날 같으면 줄 안 늘림
    assert bl.sync_report_balances('2026-11-27', dict(bals, **{'1012': 91_000_000, '1013': 63_500_000})) == 1
    rows = bl.list_tx('2026-11-25', '2026-11-27', _acct('1012')['id'])
    assert [r['balance'] for r in rows] == [91_000_000, 91_420_704] and rows[0]['source'] == bl.REPORT
    with pytest.raises(bl.LedgerError):
        bl.delete(rows[0]['id'], 'SB', True)                                # 자금일보에서 온 줄은 못 지움



def test_link_approvals_fills_empty_only(db):
    """출금 문자 ↔ 전자결재 이체 완료 — 짝이면 결재 번호, 빈 계정·적요·공사 코드만 결재 내용으로(사람이 쓴 건 그대로)."""
    bl.record_sms(HANA_OUT, bl.OUT, 'h1', 'sms', NOW)                       # 9012 출금 897,303원 11/26 10:05
    appr = [{'id': 84, 'amount': 897_303, 'project_code': 'R4294-JW', 'cost_category': '자재비', 'content': '조명 자재',
             'partner': '일산전기조명', 'requester_name': '', 'grn': 'R', 'done_at': '2026-11-26 10:30:00'},
            {'id': 85, 'amount': 897_303, 'project_code': '', 'cost_category': '자재비', 'content': '다른 통장',
             'partner': '', 'requester_name': '', 'grn': 'G', 'done_at': '2026-11-26 10:06:00'}]
    assert bl.link_approvals('2026-11-26', '2026-11-27', appr) == 1
    r = bl.list_tx('2026-11-26', '2026-11-26', _acct('9012')['id'])[0]
    assert r['approval_id'] == 84                                           # G 통장 결재(85)는 R 통장 출금과 안 맞음
    assert (r['title'], r['memo'], r['project_code']) == ('외상매입', 'R4294-JW 조명 자재', 'R4294-JW')
    assert bl.link_approvals('2026-11-26', '2026-11-27', appr) == 0          # 이미 짝
    assert bl.linked_approval_ids() == {84} and bl.sms_start().startswith('2026-11-26')
