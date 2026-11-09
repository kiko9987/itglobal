"""매입 계산서 현장 지정 (전자결재 1단계, 2026-10-30) — 글 해석·프로젝트 목록·장부·블록."""
import json

import pytest

from dashboard.services import purchase_invoice_assign as pinv
from dashboard.services.pin_remind import is_invoice_request


@pytest.fixture(autouse=True)
def _initials(monkeypatch):
    """이니셜 표는 고정값으로(실제 표는 Redis 를 읽어 테스트가 느려진다)."""
    from dashboard.services import e_approval as ea
    monkeypatch.setattr(ea, '_suffix_map', lambda: {'박정우': 'JW', '황샛별': 'SB', '강성환': 'SH'})


# ── 글 해석 ─────────────────────────────────────────────────────────
@pytest.mark.parametrize('text,expect', [
    ('10/29 G 533,500원 강원포장 회신 바랍니다.', ('10/29', 'G', 533500, '강원포장')),
    ('10/29 G 306,900원 유유산업 회신 바랍니다.', ('10/29', 'G', 306900, '유유산업')),
    ('9/2 r 277200원 에스와이에이전시 회신 부탁드립니다', ('09/02', 'R', 277200, '에스와이에이전시')),
    ('10/30 N 150,000운 사다리차 김기사', ('10/30', 'N', 150000, '사다리차 김기사')),  # '원' 오타
    ('10/30 G 1,100,000 ㈜대성크레인\n추가 설명 줄', ('10/30', 'G', 1100000, '㈜대성크레인')),
])
def test_parse_invoice_post(text, expect):
    p = pinv.parse_invoice_post(text)
    assert (p['date_md'], p['grn'], p['amount'], p['partner']) == expect
    assert is_invoice_request(text)  # 자동 고정 판정과 같은 글을 잡는다


@pytest.mark.parametrize('text', [
    '계약서 진행 부탁드립니다',
    '09/10 G 12 회의',          # 금액 아님
    '입금 452***38801011 100,000원',
    '',
])
def test_parse_invoice_post_rejects(text):
    assert pinv.parse_invoice_post(text) is None


def test_parse_code_in_post():
    # 글에 공사 코드를 같이 쓰면 거래처에 섞여 들어갔다(11-09 #6 '서울크레인 G4160-JW')
    p = pinv.parse_invoice_post('11/09 G 1,100,000원 서울크레인 G4160-JW 회신 바랍니다.')
    assert p['partner'] == '서울크레인' and p['code'] == 'G4160-JW'
    assert pinv.split_code('r4003-th 대성크레인') == ('대성크레인', 'R4003-TH')
    assert pinv.split_code('강원포장') == ('강원포장', '') and 'code' not in pinv.parse_invoice_post('10/29 G 533,500원 강원포장')
    # 버튼 값에 코드 'c' — 예전 버튼(거래처에 코드 섞임)도 떼어 읽는다
    v = json.loads(pinv.button_value('C1', '1.0', p))
    assert v['c'] == 'G4160-JW' and pinv.parsed_from_value(v)['code'] == 'G4160-JW'
    old = pinv.parsed_from_value({'d': '11/09', 'g': 'G', 'a': 1100000, 'p': '서울크레인 G4160-JW'})
    assert old['partner'] == '서울크레인' and old['code'] == 'G4160-JW'
    view = pinv.modal_view(p, {})
    assert 'G4160-JW' in json.dumps(view['blocks'], ensure_ascii=False)


def test_code_mismatch_warns_once(monkeypatch):
    from dashboard.blueprints import slack_bot as sb

    class _R:
        def __init__(self): self.d = {}
        def get(self, k): return self.d.get(k)
        def set(self, k, v, ex=None): self.d[k] = v

    class _C:
        redis = _R()
    import dashboard.utils.redis_client as rc
    monkeypatch.setattr(rc, 'get_redis_client', lambda: _C)
    view = {'id': 'V1', 'private_metadata': json.dumps({'d': '11/09', 'g': 'G', 'a': 1, 'p': '서울크레인', 'c': 'G4160-JW'})}
    assert sb._pinv_code_mismatch_once(view, 'G4160-JW') == ''          # 같으면 통과
    assert 'G4160-JW' in sb._pinv_code_mismatch_once(view, 'G4155-JW')  # 처음 다르면 막음
    assert sb._pinv_code_mismatch_once(view, 'G4155-JW') == ''          # 한 번 더 누르면 통과


def test_post_summary():
    p = pinv.parse_invoice_post('10/29 G 533,500원 강원포장 회신 바랍니다.')
    assert pinv.post_summary(p) == '10/29 G 533,500원 강원포장'


# ── 프로젝트 목록 ───────────────────────────────────────────────────
RECS = [
    {'프로젝트 코드': 'G4049-JW', '사업자명': '비즈모아', '현장 주소': '광명 양지로 21', '담당자': '박정우'},
    {'프로젝트 코드': 'G4100-JW', '사업자명': '강원상사', '현장 주소': '수원 팔달구 1', '담당자': '박정우'},
    {'프로젝트 코드': 'R4120-YM', '사업자명': '서울시립대', '현장 주소': '서울 동대문구', '담당자': '주영민',
     '거래처명': '강원인테리어'},
]


def _vals(opts):
    return [o['value'] for o in opts]


def test_options_empty_query_mine_recent_first_with_specials_on_top():
    opts = pinv.build_project_options(RECS, '', manager_name='박정우')
    assert _vals(opts) == [pinv.STOCK_CODE, pinv.COMMON_CODE, 'G4100-JW', 'G4049-JW']


def test_options_empty_query_without_manager_only_specials():
    assert _vals(pinv.build_project_options(RECS, '', manager_name=None)) == [pinv.STOCK_CODE, pinv.COMMON_CODE]


def test_options_specials_filter_for_new_fund_request():
    """자금집행 새 요청은 공통비만 — 프로젝트0(회사 재고)은 비품 신청으로."""
    only_common = (pinv.COMMON_CODE,)
    assert _vals(pinv.build_project_options(RECS, '', manager_name=None, specials=only_common)) == [pinv.COMMON_CODE]
    assert _vals(pinv.build_project_options(RECS, '재고', specials=only_common)) == []


def test_no_site_option_only_for_corpdocs():
    """법인서류 요청은 「현장 없음」(공통비 아님) — 자금집행 등 기본 목록엔 안 나온다(11-05)."""
    only = (pinv.NO_SITE_CODE,)
    opts = pinv.build_project_options(RECS, '', manager_name=None, specials=only)
    assert _vals(opts) == [pinv.NO_SITE_CODE] and opts[0]['text']['text'] == pinv.NO_SITE_LABEL
    assert _vals(pinv.build_project_options(RECS, '입찰', specials=only)) == [pinv.NO_SITE_CODE]
    assert pinv.NO_SITE_CODE not in _vals(pinv.build_project_options(RECS, '', manager_name=None))
    assert pinv.project_display(pinv.NO_SITE_CODE, None) == '현장 없음'


def test_options_search_matches_code_biz_addr_partner():
    assert _vals(pinv.build_project_options(RECS, '강원')) == ['R4120-YM', 'G4100-JW']  # 거래처명·상호
    assert _vals(pinv.build_project_options(RECS, '4049')) == ['G4049-JW']
    assert _vals(pinv.build_project_options(RECS, '재고')) == [pinv.STOCK_CODE]


def test_options_respect_limit_and_label_length():
    many = [{'프로젝트 코드': f'G{1000 + i}-JW', '사업자명': '가' * 40, '현장 주소': '나' * 80,
             '담당자': '박정우'} for i in range(50)]
    opts = pinv.build_project_options(many, '', manager_name='박정우', limit=30)
    assert len(opts) == 30
    assert all(len(o['text']['text']) <= 75 for o in opts)


def test_project_display():
    assert pinv.project_display('G4049-JW', RECS[0]) == 'G4049-JW · 광명 양지로 21'
    assert pinv.project_display(pinv.STOCK_CODE, None) == '프로젝트 0 · 회사 재고'
    assert pinv.project_display('G9999-XX', None) == 'G9999-XX'


# ── 장부 ────────────────────────────────────────────────────────────
@pytest.fixture
def tmp_db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    db = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(pinv, '_db', lambda: db)
    monkeypatch.setattr(pinv, '_table_ready', False)
    return db


PARSED = {'date_md': '10/29', 'grn': 'G', 'amount': 533500, 'partner': '강원포장'}


def test_save_then_reassign_keeps_history(tmp_db):
    assert pinv.get_assignment('C1', '111.1') is None
    row = pinv.save_assignment('C1', '111.1', PARSED, 'G4049-JW', '기타비:운반비', '', 'U1', '박정우')
    assert row['project_code'] == 'G4049-JW' and row['amount'] == 533500
    assert json.loads(row['history']) == []
    pinv.set_reply_ts(row['id'], '222.2')

    row2 = pinv.save_assignment('C1', '111.1', PARSED, 'G4100-JW', '기타비:크레인', '크레인 반입', 'U2', '강성환')
    assert row2['id'] == row['id']
    assert row2['project_code'] == 'G4100-JW' and row2['assigned_name'] == '강성환'
    assert row2['reply_ts'] == '222.2'  # 다시 지정해도 답글 위치 유지 → 같은 답글을 고친다
    hist = json.loads(row2['history'])
    assert len(hist) == 1 and hist[0]['project_code'] == 'G4049-JW'


def test_assignments_are_per_message(tmp_db):
    pinv.save_assignment('C1', '1.0', PARSED, 'G4049-JW', '자재비', '', 'U1', 'a')
    pinv.save_assignment('C1', '2.0', PARSED, pinv.COMMON_CODE, '기타비:기타', '', 'U1', 'a')
    assert pinv.get_assignment('C1', '1.0')['project_code'] == 'G4049-JW'
    assert pinv.get_assignment('C1', '2.0')['project_code'] == pinv.COMMON_CODE


# ── 슬랙 블록 ───────────────────────────────────────────────────────
def test_prompt_button_value_roundtrip():
    blocks = pinv.prompt_blocks('C1', '111.1', PARSED)
    btn = next(b for b in blocks if b['type'] == 'actions')['elements'][0]
    assert blocks[-1]['elements'][0]['text'] == '⠀'          # 다른 카드와 같은 아래 여백(11-06)
    assert btn['action_id'] == pinv.ACTION_OPEN
    v = json.loads(btn['value'])
    assert (v['ch'], v['ts']) == ('C1', '111.1')
    assert pinv.parsed_from_value(v) == PARSED


def test_result_text_and_reassign_button(tmp_db):
    row = pinv.save_assignment('C1', '111.1', PARSED, 'G4049-JW', '기타비:운반비', '1층 반입', 'U1', '박정우')
    txt = pinv.result_text(row, 'G4049-JW · 광명 양지로 21')
    assert 'G4049-JW · 광명 양지로 21' in txt and '기타비 · 운반비' in txt and '지정 JW' in txt
    assert '1층 반입' in txt and '다시 지정됨' not in txt
    row = pinv.save_assignment('C1', '111.1', PARSED, 'G4100-JW', '자재비', '', 'U2', '강성환')
    assert '다시 지정됨' in pinv.result_text(row, 'G4100-JW')
    btn = next(b for b in pinv.result_blocks(row, 'G4100-JW', 'C1', '111.1') if b['type'] == 'actions')['elements'][0]
    assert btn['action_id'] == pinv.ACTION_OPEN
    assert pinv.parsed_from_value(json.loads(btn['value'])) == PARSED


def test_modal_view_prefills_current():
    cur = {'project_code': 'G4049-JW', 'cost_category': '기타비:크레인', 'memo': '크레인'}
    opt = {'text': {'type': 'plain_text', 'text': 'G4049-JW | 비즈모아'}, 'value': 'G4049-JW'}
    view = pinv.modal_view(PARSED, {'ch': 'C1'}, cur, opt)
    assert view['callback_id'] == pinv.CALLBACK_SUBMIT
    by_id = {b.get('block_id'): b for b in view['blocks']}
    assert by_id['project']['element']['initial_option'] == opt
    assert by_id['project']['element']['action_id'] == pinv.ACTION_PROJECT
    assert by_id['cost']['element']['initial_option']['value'] == '기타비:크레인'
    assert by_id['memo']['element']['initial_value'] == '크레인'
    assert by_id['memo']['optional'] is True


def test_modal_view_fresh_has_no_initials():
    view = pinv.modal_view(PARSED, {'ch': 'C1'})
    by_id = {b.get('block_id'): b for b in view['blocks']}
    assert 'initial_option' not in by_id['project']['element']
    assert 'initial_option' not in by_id['cost']['element']
    assert 'initial_value' not in by_id['memo']['element']


# ── 슬랙 처리 (가짜 클라이언트) ─────────────────────────────────────
class _FakeClient:
    def __init__(self):
        self.calls = []

    def __getattr__(self, name):
        def _call(**kw):
            self.calls.append((name, kw))
            if name == 'chat_postMessage':
                return {'ts': '999.9'}
            if name == 'views_open':
                return {'view': {'id': 'V1'}}
            return {}
        return _call

    def names(self):
        return [n for n, _ in self.calls]


@pytest.fixture
def sb(monkeypatch, tmp_db):
    from dashboard.blueprints import slack_bot as sb
    from dashboard.services import e_approval as ea
    monkeypatch.setattr(ea, '_db', lambda: tmp_db)
    monkeypatch.setattr(ea, '_table_ready', False)
    monkeypatch.setattr(ea, '_suffix_map', lambda: {'박정우': 'JW', '황샛별': 'SB'})
    monkeypatch.setenv('SLACK_INVOICE_CHANNEL_ID', 'CINV')
    monkeypatch.setattr(sb, '_resolve_intake_manager_name', lambda uid: '박정우')
    monkeypatch.setattr(sb, '_pinv_project_rec', lambda code: RECS[0] if code == 'G4049-JW' else None)

    class _R:
        def __init__(self):
            self.d = {}

        def get(self, k):
            return self.d.get(k)

        def set(self, k, v, ex=None):
            self.d[k] = v

    fake_redis = type('RC', (), {'redis': _R()})()
    from dashboard.utils import redis_client
    monkeypatch.setattr(redis_client, 'get_redis_client', lambda: fake_redis)
    sb._fake_redis = fake_redis.redis
    return sb


POST = '10/29 G 533,500원 강원포장 회신 바랍니다.'


def test_prompt_off_by_default(sb, monkeypatch):
    monkeypatch.delenv('PURCHASE_INVOICE_ASSIGN_ENABLED', raising=False)
    c = _FakeClient()
    sb._maybe_post_pinv_prompt(c, 'CINV', '111.1', POST)
    assert c.calls == []


def test_prompt_posted_only_for_invoice_post_in_sales_channel(sb, monkeypatch):
    monkeypatch.setenv('PURCHASE_INVOICE_ASSIGN_ENABLED', 'true')
    c = _FakeClient()
    sb._maybe_post_pinv_prompt(c, 'COTHER', '111.1', POST)           # 다른 채널
    sb._maybe_post_pinv_prompt(c, 'CINV', '111.2', '입금 452***38801011 100,000원')  # 입금 문자
    assert c.calls == []
    sb._maybe_post_pinv_prompt(c, 'CINV', '111.1', POST)
    (name, kw), = c.calls
    assert name == 'chat_postMessage' and kw['thread_ts'] == '111.1'
    assert json.loads(next(b for b in kw['blocks'] if b['type'] == 'actions')['elements'][0]['value'])['a'] == 533500
    assert sb._fake_redis.get('pinv_prompt:CINV:111.1')            # 버튼 위치 기억(자금집행 창에서 지정 시 정리용)


# ── 11-01 자금집행 창에서 계산서 고르기 ──────────────────────────────
class _PinsClient(_FakeClient):
    def __init__(self, items):
        super().__init__()
        self.items = items

    def pins_list(self, **kw):
        return {'items': [{'message': m} for m in self.items]}


def test_open_invoice_posts_skips_assigned_deposit_and_requested(sb, monkeypatch):
    from dashboard.blueprints import slack_eapproval as se
    from dashboard.services import e_approval as ea
    pins = [{'ts': '100.1', 'text': POST},
            {'ts': '100.2', 'text': '10/30 R 99,000원 대성크레인 회신 바랍니다.'},
            {'ts': '100.3', 'text': '입금 452***38801011 100,000원'},
            {'ts': '100.4', 'text': '10/31 G 55,000원 사다리차 회신 바랍니다.'},
            {'ts': '100.5', 'text': '계약서 진행 부탁드립니다'}]
    monkeypatch.setattr(se, '_invoice_client', lambda: _PinsClient(pins))
    pinv.save_assignment('CINV', '100.2', pinv.parse_invoice_post(pins[1]['text']), 'G4049-JW', '자재비', '',
                         'U1', '박정우')
    ea.create(ea.FORM_FUND, 'U1', '박정우', {'amount': 55000, 'origin_channel': 'CINV', 'origin_ts': '100.4'})
    items = se.open_invoice_posts(force=True)
    assert [i['ts'] for i in items] == ['100.1']
    opts = se.invoice_options('')
    assert opts[0]['value'] == ea.INVOICE_NONE and opts[1]['text']['text'] == '10/29 G 533,500원 강원포장'
    assert se.invoice_options('크레인') == [ea.invoice_option(None)]


def _inv_view(meta_inv, **vals):
    from dashboard.services import e_approval as ea
    base = {'evidence': '전자계산서', 'cost': '기타비:운반비', 'content': '', 'payee_name': '김강원',
            'bank': '신한', 'account': '110-511-825643'}
    base.update(vals)
    sel = {'evidence', 'cost', 'urgency'}
    v = {k: {k: ({'selected_option': {'value': x}} if k in sel else {'value': x})} for k, x in base.items()}
    v['project'] = {ea.A_PROJECT: {'selected_option': {'value': 'G4049-JW'}}}
    return {'private_metadata': json.dumps({'project': '', 'inv': meta_inv}), 'state': {'values': v}}


def test_new_request_with_invoice_uses_post_values_and_blocks_duplicate(sb):
    from dashboard.services import e_approval as ea
    inv = {'ch': 'CINV', 'ts': '111.1', 'parsed': pinv.parse_invoice_post(POST)}
    f = ea.read_new_request(_inv_view(inv))
    assert ea.complete_new_request(f) == {}
    assert (f['partner'], f['amount'], f['grn'], f['origin_ts']) == ('강원포장', 533500, 'G', '111.1')
    ea.create(ea.FORM_FUND, 'U1', '박정우', f)
    f2 = ea.read_new_request(_inv_view(inv))
    assert 'invoice' in ea.complete_new_request(f2)                 # 같은 계산서로 두 번 못 올림
    v = ea.new_request_view({'inv': inv})
    ids = [b.get('block_id') for b in v['blocks']]
    assert 'partner' not in ids and 'amount' not in ids and ids[0] == 'invoice'
    assert v['blocks'][0]['element']['initial_option']['text']['text'] == '10/29 G 533,500원 강원포장'
    assert 'partner' in [b.get('block_id') for b in ea.new_request_view({})['blocks']]


def test_new_request_with_invoice_also_finishes_site_assignment(sb, monkeypatch):
    from dashboard.blueprints import slack_eapproval as se
    from dashboard.services import e_approval as ea
    eap, inv_c = _FakeClient(), _FakeClient()
    monkeypatch.setenv('SLACK_EAPPROVAL_CHANNEL', 'CEAP')
    monkeypatch.setattr(se, 'eap_client', lambda: eap)
    monkeypatch.setattr(se, '_invoice_client', lambda: inv_c)
    monkeypatch.setattr(se, '_name', lambda uid: '박정우')
    monkeypatch.setattr(se, '_project_rec', lambda code: None)
    sb._fake_redis.set('pinv_prompt:CINV:111.1', '500.5')
    inv = {'ch': 'CINV', 'ts': '111.1', 'parsed': pinv.parse_invoice_post(POST)}
    f = ea.read_new_request(_inv_view(inv))
    ea.complete_new_request(f)
    row = se.process_new_submit('U1', f)
    a = pinv.get_assignment('CINV', '111.1')
    assert a['project_code'] == 'G4049-JW' and a['cost_category'] == '기타비:운반비'
    posted = [kw for n, kw in inv_c.calls if n == 'chat_postMessage']
    assert posted and posted[0]['thread_ts'] == '111.1' and ea.title(row) in posted[0]['text']
    assert any(n == 'chat_update' and kw['ts'] == '500.5' for n, kw in inv_c.calls)   # 버튼 메시지 정리
    assert any(n == 'pins_remove' and kw['timestamp'] == '111.1' for n, kw in inv_c.calls)


def _submit_body(code='G4049-JW', cost='기타비:운반비', memo=None, click_ts='500.5'):
    meta = {'ch': 'CINV', 'ts': '111.1', 'click_ts': click_ts,
            'd': '10/29', 'g': 'G', 'a': 533500, 'p': '강원포장'}
    view = {'private_metadata': json.dumps(meta), 'state': {'values': {
        'project': {pinv.ACTION_PROJECT: {'selected_option': {'value': code}}},
        'cost': {'cost': {'selected_option': {'value': cost}}},
        'memo': {'memo': {'value': memo}},
    }}}
    return {'user': {'id': 'U1'}}, view


def test_submit_first_time_replies_cleans_button_and_unpins(sb):
    c = _FakeClient()
    sb._process_pinv_submit(c, *_submit_body())
    assert c.names() == ['chat_postMessage', 'chat_update', 'pins_remove']
    post = c.calls[0][1]
    assert post['thread_ts'] == '111.1' and 'G4049-JW · 광명 양지로 21' in post['text']
    assert c.calls[1][1]['ts'] == '500.5'          # 버튼 메시지 정리
    assert c.calls[2][1]['timestamp'] == '111.1'   # 원 글 고정 해제
    row = pinv.get_assignment('CINV', '111.1')
    assert row['reply_ts'] == '999.9' and row['assigned_name'] == '박정우'


def test_submit_reassign_edits_same_reply(sb):
    sb._process_pinv_submit(_FakeClient(), *_submit_body())
    c2 = _FakeClient()
    sb._process_pinv_submit(c2, *_submit_body(code=pinv.STOCK_CODE, cost='자재비', click_ts='999.9'))
    assert c2.names() == ['chat_update', 'pins_remove']
    upd = c2.calls[0][1]
    assert upd['ts'] == '999.9' and '프로젝트 0 · 회사 재고' in upd['text'] and '다시 지정됨' in upd['text']


def test_open_modal_prefills_existing(sb):
    pinv.save_assignment('CINV', '111.1', PARSED, 'G4049-JW', '기타비:크레인', '', 'U1', '박정우')
    c = _FakeClient()
    body = {'trigger_id': 'T', 'message': {'ts': '999.9'},
            'actions': [{'value': pinv.button_value('CINV', '111.1', PARSED)}]}
    sb._open_pinv_modal(c, body)
    assert c.names() == ['views_open', 'views_update']
    view = c.calls[1][1]['view']
    by_id = {b.get('block_id'): b for b in view['blocks']}
    assert by_id['project']['element']['initial_option']['value'] == 'G4049-JW'
    assert json.loads(view['private_metadata'])['click_ts'] == '999.9'


# ── 2단계: 「이체 필요」 → 자금집행 ──────────────────────────────────
META = {'ch': 'CINV', 'ts': '111.1', 'click_ts': '500.5', 'd': '10/29', 'g': 'G', 'a': 533500, 'p': '강원포장'}


def _ids(view):
    return [b.get('block_id') for b in view['blocks']]


def test_modal_toggle_shows_transfer_fields_with_remembered_account(sb):
    from dashboard.services import e_approval as ea
    view = sb._pinv_modal_for(dict(META), transfer=False)
    assert 'transfer_toggle' in _ids(view) and 'payee_name' not in _ids(view)
    ea.remember_payee('강원포장', '김강원', '신한', '110-511-825643')
    view = sb._pinv_modal_for(dict(META), transfer=True)
    by_id = {b.get('block_id'): b for b in view['blocks']}
    assert by_id['transfer_toggle']['elements'][0]['initial_options']
    assert by_id['account']['element']['initial_value'] == '110-511-825643'
    assert by_id['evidence']['element']['initial_option']['value'] == '전자계산서'


def test_modal_shows_existing_approval_instead_of_toggle(sb):
    from dashboard.services import e_approval as ea
    row = ea.create(ea.FORM_FUND, 'U1', '박정우', {'amount': 533500, 'grn': 'G', 'partner': '강원포장',
                                                   'origin_channel': 'CINV', 'origin_ts': '111.1'})
    view = sb._pinv_modal_for(dict(META), transfer=False)
    assert 'transfer_toggle' not in _ids(view)
    assert f'자금집행 #{row["id"]}' in json.dumps(view, ensure_ascii=False)
    ea.transition(row['id'], 'reject', 'U04UL2ZLJAX', '고광일', 'x')
    assert 'transfer_toggle' in _ids(sb._pinv_modal_for(dict(META), transfer=False))  # 반려되면 다시 올릴 수 있다


def test_modal_links_earlier_request_with_same_invoice(sb):
    from dashboard.services import e_approval as ea
    row = ea.create(ea.FORM_FUND, 'U1', '박정우', {'project_code': 'G4049-JW', 'cost_category': '자재비',
                                                   'amount': 533500, 'grn': 'G', 'partner': '㈜강원포장'})
    meta = dict(META)
    view = sb._pinv_modal_for(meta, transfer=False)
    assert meta['link_aid'] == row['id'] and 'transfer_toggle' not in _ids(view)
    by_id = {b.get('block_id'): b for b in view['blocks']}
    assert by_id['project']['element']['initial_option']['value'] == 'G4049-JW'
    assert by_id['cost']['element']['initial_option']['value'] == '자재비'
    body, v = _submit_body()
    v['private_metadata'] = json.dumps(meta)
    c = _FakeClient()
    sb._process_pinv_submit(c, body, v)
    assert ea.get(row['id'])['origin_ts'] == '111.1'
    assert f'자금집행 #{row["id"]}' in c.calls[0][1]['text'] and '연결' in c.calls[0][1]['text']


def _transfer_values(view, payee='김강원', bank='신한', account='110-511-825643'):
    vals = view['state']['values']
    vals['transfer_toggle'] = {pinv.ACTION_TRANSFER_TOGGLE: {'selected_options': [{'value': 'yes'}]}}
    vals['evidence'] = {'evidence': {'selected_option': {'value': '수기계산서'}}}
    vals['payee_name'] = {'payee_name': {'value': payee}}
    vals['bank'] = {'bank': {'value': bank}}
    vals['account'] = {'account': {'value': account}}
    vals['urgency'] = {'urgency': {'selected_option': {'value': '긴급'}}}
    return view


def test_validate_transfer_fields_and_bot_config(sb, monkeypatch):
    _b, v = _submit_body()
    assert sb._validate_pinv_submit(v) == {}                      # 이체 안 함 → 검사 없음
    v = _transfer_values(v, payee='', account='12')
    assert set(sb._validate_pinv_submit(v)) == {'payee_name', 'account'}
    v = _transfer_values(_submit_body()[1])
    monkeypatch.delenv('SLACK_EAPPROVAL_BOT_TOKEN', raising=False)
    assert 'account' in sb._validate_pinv_submit(v)               # 봇 미설정이면 막음
    monkeypatch.setenv('SLACK_EAPPROVAL_BOT_TOKEN', 'x')
    monkeypatch.setenv('SLACK_EAPPROVAL_CHANNEL', 'CEAP')
    assert sb._validate_pinv_submit(v) == {}


def test_submit_with_transfer_starts_approval(sb, monkeypatch):
    from dashboard.blueprints import slack_eapproval as se
    got = {}

    def fake_start(uid, name, ch, ts, parsed, code, cost, memo, ev, payee, bank, acct, urg):
        got.update(uid=uid, ch=ch, ts=ts, code=code, ev=ev, payee=payee, acct=acct, urg=urg,
                   amount=parsed['amount'])
        return {'id': 7, 'form': 'fund', 'status': 'pending'}
    monkeypatch.setattr(se, 'start_from_invoice', fake_start)
    body, v = _submit_body()
    c = _FakeClient()
    sb._process_pinv_submit(c, body, _transfer_values(v))
    assert got == {'uid': 'U1', 'ch': 'CINV', 'ts': '111.1', 'code': 'G4049-JW', 'ev': '수기계산서',
                   'payee': '김강원', 'acct': '110-511-825643', 'urg': '긴급', 'amount': 533500}
    assert '자금집행 #7 결재 올림' in c.calls[0][1]['text']


def test_submit_with_transfer_failure_still_assigns(sb, monkeypatch):
    from dashboard.blueprints import slack_eapproval as se

    def boom(*a, **k):
        raise RuntimeError('slack down')
    monkeypatch.setattr(se, 'start_from_invoice', boom)
    body, v = _submit_body()
    c = _FakeClient()
    sb._process_pinv_submit(c, body, _transfer_values(v))
    assert pinv.get_assignment('CINV', '111.1')['project_code'] == 'G4049-JW'
    assert '결재는 올라가지 않았어요' in c.calls[0][1]['text']
