"""보안 점검 3차(2026-11-14) — 공사·A/S·정산·계산서 카드와 문의·상담 창에 들어가는 시트·리드·OCR 값의
슬랙 표기(& < >) 이스케이프. 고객 글이 '<!channel>'·위장 링크로 동작하지 않아야 한다."""
import pytest

EVIL = '<!channel> <https://x.example|견적서 확인> &'
SAFE = '&lt;!channel&gt;'


def _clean(text: str) -> None:
    assert SAFE in text
    assert '<!channel>' not in text and '<https://x.example' not in text


def _project(**over):
    d = {'프로젝트 코드': 'G9999-JW', '유입 구분': EVIL, '거래처명': '(주)' + EVIL, '사업자명': EVIL,
         '현장 주소': EVIL, '발주처 담당자': EVIL, '발주처 연락처': EVIL, '발주처 이메일': EVIL,
         '공사 내용': EVIL, '도급 구분': EVIL, '시공자': EVIL, '총액 1': 1000000, '부가세': True,
         '공사 시작': EVIL, '공사 종료': '2026-11-20'}
    d.update(over)
    return d


def test_project_card_escapes_values_but_keeps_own_links():
    from dashboard.services import project_slack_notifier as psn
    text = psn._build_message(_project(**{'견적서 및 계약서 폴더 경로': 'a' * 25}), 'G9999-JW',
                              thread_permalink='https://w.slack.com/archives/C1/p1')
    _clean(text)
    for label in ('사업자명', '현장 주소', '발주처 담당자', '공사 내용', '시공자', '공사 시작'):
        line = next(ln for ln in text.split('\n') if label in ln)
        assert SAFE in line, label
    # 코드가 만든 링크는 그대로
    assert f'<itgfolder://{"a" * 25}|탐색기에서 열기>' in text
    assert '<https://w.slack.com/archives/C1/p1|(첨부하기)>' in text
    # 로컬 경로 글은 이스케이프
    text2 = psn._build_message(_project(**{'견적서 및 계약서 폴더 경로': 'C:\\공사\\' + EVIL}), 'G9999-JW')
    assert 'C:\\공사\\&lt;!channel&gt;' in text2


def test_project_card_button_value_stays_raw():
    # 계산서 요청 창 미리 채우기는 원래 값 — 창에서 이스케이프된 글자가 보이면 안 된다
    import json
    from dashboard.services import project_slack_notifier as psn
    v = json.loads(psn._build_invoice_button_value(_project(), 'G9999-JW'))
    assert v['biz'] == EVIL and v['addr'] == EVIL


def test_settlement_dm_site_lines():
    from dashboard.services import project_slack_notifier as psn
    lines = psn._dm_site_lines(_project())
    assert lines and all(SAFE in ln for ln in lines)


def test_info_edit_dm_escapes_changes(monkeypatch):
    from dashboard.services import project_slack_notifier as psn
    import dashboard.services.settlement_checkers as sc
    sent = []

    class _WC:
        def __init__(self, token):
            pass

        def conversations_open(self, users):
            return {'channel': {'id': 'D1'}}

        def chat_postMessage(self, **kw):
            sent.append(kw['text'])
    import slack_sdk
    monkeypatch.setattr(slack_sdk, 'WebClient', _WC)
    monkeypatch.setattr(sc, 'info_notice_ids', lambda email: ['U1'])
    monkeypatch.setenv('SLACK_BOT_TOKEN', 'xoxb-test')
    monkeypatch.setattr(psn, '_project_card_permalink', lambda code: 'https://w.slack.com/p1')
    ok = psn.notify_info_edit_to_settlement(
        'G9999-JW', [{'field_name': '현장 주소', 'old_value': '광명', 'new_value': EVIL}],
        editor_email='mgr@itg-aircon.com', latest_data=_project())
    assert ok and sent
    _clean(sent[0])
    assert '<https://w.slack.com/p1|공사 확정 카드 보기>' in sent[0]


def test_as_card_text_escapes_both_layouts():
    from dashboard.blueprints import slack_bot as sb
    data = {'No': 'AS-1', '현장주소': EVIL, '공사내용': EVIL, '요청 내용': EVIL, '요청자': EVIL,
            '조치 내용': EVIL}
    _clean(sb._build_as_card_text(dict(data), view_state='accepted'))           # 코드 이전 공사
    proj = {'inflow': EVIL, 'biz': EVIL, 'partner': '(주)' + EVIL, 'address': EVIL, 'client_manager': EVIL,
            'client_phone': EVIL, 'client_email': EVIL, 'work_content': EVIL, 'contract_type': EVIL,
            'contractor': EVIL, 'amount': EVIL, 'work_start': EVIL, 'work_end': EVIL}
    text = sb._build_as_card_text(dict(data, **{'프로젝트 코드': 'G9999-JW'}), view_state='completed', proj=proj)
    _clean(text)
    assert '*[A/S 조치 완료]*' in text   # 굵은 글씨·코드 표기는 그대로


def test_site_head_lines():
    from dashboard.blueprints import slack_bot as sb
    lines = sb._site_head_lines(_project())
    assert lines and all(SAFE in ln for ln in lines)


def test_invoice_request_card(monkeypatch):
    from dashboard.blueprints import slack_bot as sb
    posted = []

    class _C:
        def conversations_join(self, **kw):
            return {}

        def chat_postMessage(self, **kw):
            posted.append(kw)
            return {'ok': False}

    class _App:
        client = _C()
    monkeypatch.setenv('SLACK_INVOICE_REQUEST_CHANNEL_ID', 'CINV')
    monkeypatch.setattr(sb, '_invoice_slack_app', _App())
    monkeypatch.setattr(sb, '_partner_status_warn', lambda biz: '')
    r = sb.post_invoice_request('G9999-JW', EVIL, EVIL, '1000000', 'sep', EVIL, EVIL, 'JW', dedup_check=False)
    assert r == {'ok': False, 'reason': 'post_failed'}
    text = posted[0]['blocks'][0]['text']['text']
    _clean(text)
    _clean(posted[0]['text'])
    for label in ('사업자명', '현장 주소', '이메일', '요청사항'):
        assert SAFE in next(ln for ln in text.split('\n') if label in ln), label


def test_consult_info_blocks():
    from dashboard.blueprints import slack_bot as sb
    lead = {'고객명': EVIL, '고객 연락처': '010-1234-5678', '이메일': EVIL, '문의 내용': EVIL,
            '상담 내용': '지난번 ' + EVIL, '키워드': EVIL, '상담 시간': '2026-11-14 09:00'}
    blocks = sb._build_consult_info_blocks(lead, 'L-00001')
    texts = [b['text']['text'] for b in blocks if b.get('type') == 'section']
    assert len(texts) == 2
    for t in texts:
        _clean(t)


def test_fund_move_candidate_name():
    # 입금 카드의 '법인 간 이체' 후보 줄 — 이름·은행은 앞선 입금 문자 원문에서 온다
    from dashboard.blueprints import sms_inbound as si
    cand = {'code': 'G9999-JW', 'stage': '잔금', 'bank': EVIL, 'to': 'global', 'amount': 1000, 'partner': EVIL}
    blocks = si._build_intake_blocks('i1', '[Web발신] 입금 1,000원', {'bank': '하나'},
                                     transfer={'cands': [cand], 'evidence': False})
    t = next(b['text']['text'] for b in blocks if b.get('type') == 'section' and '법인 간 이체' in b['text']['text'])
    _clean(t)
