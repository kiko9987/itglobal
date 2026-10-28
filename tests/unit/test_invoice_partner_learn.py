# -*- coding: utf-8 -*-
"""발행된 세금계산서 → 거래처 탭 (2026-10-23 G4158 서울아산가정의학과의원: 홈택스엔 있는데 거래처 탭엔
없어 '부담자 등록증을 찾지 못했습니다' 경고). 계산서 OCR 단어 위치로 공급받는자 읽기 · 거래처 탭 넣기 ·
부담자 사업자번호로 거래처 정보 찾기 · 자리표시 사업자명 정리."""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.services.invoice_ocr as io_
import dashboard.services.partner_status_sync as ps


def W(text, x0, x1, y0, y1=None):
    return {'text': text, 'x0': x0, 'x1': x1, 'y0': y0, 'y1': (y1 if y1 is not None else y0 + 0.01)}


# 홈택스 계산서 흉내 — 왼쪽 공급자, 오른쪽 공급받는자. 상호·사업장은 칸 안에서 두 줄로 줄바꿈.
INVOICE = [
    W('승인번호', 0.05, 0.12, 0.02),
    W('등록', 0.05, 0.08, 0.10), W('884-81-01425', 0.12, 0.25, 0.10),
    W('등록', 0.55, 0.58, 0.10), W('831-03-01097', 0.62, 0.75, 0.10),
    W('상호', 0.05, 0.08, 0.15), W('(주)아이티글로벌', 0.12, 0.25, 0.15), W('성명', 0.30, 0.33, 0.15),
    W('박정우', 0.36, 0.42, 0.15),
    W('상호', 0.55, 0.58, 0.15), W('서울아산가정의학과의', 0.62, 0.78, 0.140), W('원', 0.62, 0.64, 0.153),
    W('성명', 0.80, 0.83, 0.15), W('민선홍', 0.86, 0.92, 0.15),
    W('사업장', 0.05, 0.09, 0.20), W('경기도', 0.12, 0.16, 0.20), W('김포시', 0.17, 0.21, 0.20),
    W('사업장', 0.55, 0.59, 0.20), W('인천광역시', 0.62, 0.68, 0.193), W('부평구', 0.69, 0.73, 0.193),
    W('508호(산곡', 0.74, 0.80, 0.193), W('동', 0.62, 0.635, 0.205), W(',', 0.635, 0.64, 0.205),
    W('오남프라자)', 0.65, 0.72, 0.205),
    W('업태', 0.05, 0.08, 0.25), W('도매업', 0.12, 0.17, 0.25), W('종목', 0.25, 0.28, 0.25), W('배관', 0.32, 0.36, 0.25),
    W('업태', 0.55, 0.58, 0.25), W('보건업', 0.62, 0.67, 0.25), W('종목', 0.75, 0.78, 0.25),
    W('가정의학과', 0.82, 0.90, 0.25),
    W('이메일', 0.05, 0.09, 0.30), W('tax@itg-aircon.com', 0.12, 0.30, 0.30),
    W('이메일', 0.55, 0.59, 0.30), W('clemensss@naver.com', 0.62, 0.80, 0.30),
    W('이메일', 0.55, 0.59, 0.33),
    W('작성일자', 0.05, 0.11, 0.40), W('공급가액', 0.20, 0.27, 0.40), W('7,001,000', 0.20, 0.28, 0.43),
    W('공급받는자', 0.50, 0.53, 0.20),
]


def test_parse_buyer_from_layout():
    r = io_.parse_invoice_buyer(INVOICE)
    assert r['bno'] == '831-03-01097' and r['supplier_bno'] == '884-81-01425'
    assert r['name'] == '서울아산가정의학과의원'                # 줄바꿈된 상호 이어 붙임
    assert r['ceo'] == '민선홍'
    assert r['addr'] == '인천광역시 부평구 508호(산곡동, 오남프라자)'
    assert r['biz_type'] == '보건업' and r['biz_item'] == '가정의학과'
    assert r['email'] == 'clemensss@naver.com'


def test_read_invoice_buyer_guards(monkeypatch):
    monkeypatch.setattr(io_, 'ocr_words', lambda content: INVOICE)
    assert io_.read_invoice_buyer(b'x')['name'] == '서울아산가정의학과의원'
    # 등록증·사진 등 계산서 양식이 아니면 None
    monkeypatch.setattr(io_, 'ocr_words', lambda content: [W('사업자등록증', 0.1, 0.3, 0.1), W('상호', 0.1, 0.2, 0.2)])
    assert io_.read_invoice_buyer(b'x') is None
    # 공급받는자 번호가 공급자 번호와 같으면(잘못 읽음) None
    same = [dict(w, text='884-81-01425') if w['text'] == '831-03-01097' else w for w in INVOICE]
    monkeypatch.setattr(io_, 'ocr_words', lambda content: same)
    assert io_.read_invoice_buyer(b'x') is None


# ── 거래처 탭 넣기 ──
class _Exec:
    def __init__(self, v):
        self.v = v

    def execute(self):
        return self.v


class _Sheet:
    def __init__(self, rows):
        self.rows, self.updates, self.batches = rows, [], []

    def values(self):
        s = self

        class V:
            def get(self, **k):
                return _Exec({'values': s.rows})

            def update(self, **k):
                s.updates.append(k)
                return _Exec({})

            def batchUpdate(self, **k):
                s.batches.append(k)
                return _Exec({})
        return V()


class _Lock:
    def __init__(self):
        self.held = False

    def set(self, k, v, nx=False, ex=None):
        if self.held:
            return False
        self.held = True
        return True

    def delete(self, k):
        self.held = False


@pytest.fixture
def tab(monkeypatch):
    holder = {}

    def _use(rows):
        sh = _Sheet(rows)
        svc = type('S', (), {'spreadsheets': lambda self: sh})()
        monkeypatch.setattr(ps, '_sheet', lambda: (type('M', (), {'service': svc})(), 'SID'))
        holder['lock'] = _Lock()
        import dashboard.utils.redis_client as rcm
        monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': holder['lock']})())
        holder['rebuilt'] = 0
        monkeypatch.setattr(ps, 'rebuild_partner_caches', lambda: holder.__setitem__('rebuilt', holder['rebuilt'] + 1))
        return sh
    return _use, holder


BUYER = {'bno': '831-03-01097', 'name': '서울아산가정의학과의원', 'ceo': '민선홍',
         'addr': '인천광역시 부평구 부영로189번길 51', 'biz_type': '보건업', 'biz_item': '가정의학과',
         'email': 'clemensss@naver.com'}


def test_upsert_adds_new_partner_row(tab):
    use, h = tab
    sh = use([['등록번호', '상호명'], ['', '당근'], ['105-86-00001', '주식회사 기존상사']])
    assert ps.upsert_partner_from_invoice(BUYER) == 'added'
    up = sh.updates[0]
    assert up['range'] == '거래처!A4:G4' and up['valueInputOption'] == 'RAW'
    assert up['body']['values'][0][:3] == ['831-03-01097', '서울아산가정의학과의원', '민선홍']
    assert h['rebuilt'] == 1 and not h['lock'].held                    # 캐시 다시 만들고 잠금 풂


def test_upsert_fills_only_empty_cells(tab):
    use, h = tab
    sh = use([['등록번호', '상호명'], ['8310301097', '서울아산가정의학과', '', '', '보건업']])
    assert ps.upsert_partner_from_invoice(BUYER) == 'filled'
    ranges = {d['range']: d['values'][0][0] for d in sh.batches[0]['body']['data']}
    assert '거래처!B2' not in ranges                                   # 이미 적힌 상호는 그대로
    assert ranges['거래처!C2'] == '민선홍' and ranges['거래처!G2'] == 'clemensss@naver.com'
    assert '거래처!E2' not in ranges


def test_upsert_same_and_skip(tab):
    use, h = tab
    full = ['831-03-01097', '서울아산가정의학과의원', '민선홍', '인천', '보건업', '가정의학과', 'a@b.com']
    sh = use([['등록번호'], full])
    assert ps.upsert_partner_from_invoice(BUYER) == 'same' and not sh.updates and not sh.batches
    assert h['rebuilt'] == 0
    assert ps.upsert_partner_from_invoice(dict(BUYER, bno='831-03-01098')) == 'skip'   # 체크섬 틀림
    assert ps.upsert_partner_from_invoice(dict(BUYER, name='')) == 'skip'


def test_partner_info_by_bno(monkeypatch):
    monkeypatch.setattr(ps, 'get_partner_name_by_bno', lambda b: '서울아산가정의학과의원' if b == '831-03-01097' else None)
    monkeypatch.setattr(ps, 'get_partner_info', lambda n: {'name': n, 'bno': '831-03-01097'})
    assert ps.get_partner_info_by_bno('831-03-01097')['name'] == '서울아산가정의학과의원'
    assert ps.get_partner_info_by_bno('000-00-00000') is None


# ── 슬랙: 부담자 칸 번호로 거래처 정보 · 발행 계산서에서 배우기 ──
def test_payer_partner_info_falls_back_to_bno(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    monkeypatch.setattr(ps, 'get_partner_info', lambda n: None)
    monkeypatch.setattr(ps, 'get_partner_info_by_bno', lambda b: {'name': '서울아산가정의학과의원'} if b else None)
    payer = {'name': '민선홍(서울아산가정의학)', 'biz_no': '831-03-01097'}
    assert sb._payer_partner_info('민선홍(서울아산가정의학)', payer)['name'] == '서울아산가정의학과의원'
    assert sb._payer_partner_info('민선홍(서울아산가정의학)', {'name': 'x'}) is None


def test_learn_from_issued_invoice(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    import dashboard.services.business_license_handler as blh
    monkeypatch.setattr(blh, '_download_slack_file', lambda url, token: b'img')
    monkeypatch.setattr(io_, 'read_invoice', lambda content: dict(BUYER, supplier_bno='884-81-01425'))
    got, filled, moves = [], [], []
    monkeypatch.setattr(ps, 'upsert_partner_from_invoice', lambda info: got.append(info['bno']) or 'added')
    monkeypatch.setattr(sb, '_fill_payer_from_invoice', lambda *a, **k: filled.append((a, k)))
    monkeypatch.setattr(sb, '_close_fund_moves_by_invoice',
                        lambda client, ch, ts, meta, inv: moves.append((ch, ts, inv['supplier_bno'])) or 0)
    event = {'channel': 'CINV', 'thread_ts': '9.0',
             'files': [{'mimetype': 'image/png', 'url_private': 'u'}, {'mimetype': 'text/plain', 'url_private': 'v'}]}
    meta = {'code': 'G4158-JW', 'stages': '잔금', 'biz': '민선홍(서울아산가정의학)'}
    assert sb._learn_from_issued_invoice(None, event, meta) == 1
    assert got == ['831-03-01097'] and moves == [('CINV', '9.0', '884-81-01425')]
    (args, kw), = filled
    assert args[:3] == ('G4158-JW', '잔금', '민선홍(서울아산가정의학)')
    assert kw == {'bno': '831-03-01097', 'buyer_name': '서울아산가정의학과의원', 'buyer_ceo': '민선홍'}
    # 공급받는자를 못 읽어도(주민번호 발행) 자금 이동 정리는 함, 거래처 탭은 안 건드림
    monkeypatch.setattr(io_, 'read_invoice', lambda content: dict(BUYER, bno='', supplier_bno='533-88-02845'))
    got.clear(), filled.clear(), moves.clear()
    assert sb._learn_from_issued_invoice(None, event, meta) == 0 and not got and not filled and len(moves) == 1
    # 계산서 양식이 아니면 아무것도 안 함
    monkeypatch.setattr(io_, 'read_invoice', lambda content: None)
    moves.clear()
    assert sb._learn_from_issued_invoice(None, event, meta) == 0 and not got and not moves


def test_close_fund_moves_by_invoice_posts_note(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    import dashboard.services.fund_transfer as ft
    req = {'code': 'G3982-JW', 'stage': '잔금', 'amount': 2310000}
    calls = []
    monkeypatch.setattr(ft, 'close_by_invoice', lambda code, stages, sup: calls.append((code, stages, sup)) or [req])
    posted = []
    c = type('C', (), {'chat_postMessage': lambda self, **k: posted.append(k['text'])})()
    n = sb._close_fund_moves_by_invoice(c, 'CINV', '9.0', {'code': 'G3982-JW', 'stages': '잔금'},
                                        {'supplier_bno': '533-88-02845'})
    assert n == 1 and calls == [('G3982-JW', '잔금', 'R')]
    assert '글로벌그룹 계좌로 들어온 돈이고 계산서도 글로벌그룹 이름으로' in posted[0]
    # 우리 법인 번호가 아니면(못 읽음) 아무것도 안 함
    assert sb._close_fund_moves_by_invoice(c, 'CINV', '9.0', {'code': 'G3982-JW', 'stages': '잔금'},
                                           {'supplier_bno': ''}) == 0 and len(calls) == 1


# ── 자리표시 사업자명 ──
def test_placeholder_biz():
    import dashboard.services.business_license_handler as blh
    for v in ('', '-', '온라인', '거래처', '당근', '숨고', '홈페이지', '중고판매', ' 큐플레이스 '):
        assert blh.is_placeholder_biz(v), v
    for v in ('한', '데이즈헤어', '(주)엠아이텍'):
        assert not blh.is_placeholder_biz(v), v
    assert blh._norm_biz('당근') == ''                                 # 등록증 재사용·거래처 판정 키 안 됨


def test_ocr_fills_over_placeholder_biz(monkeypatch):
    """옛 공사에 남은 '온라인'이 등록증 상호 자동 기재를 막던 문제 — 자리표시는 빈칸처럼 채운다."""
    import dashboard.services.business_license_handler as blh
    import dashboard.services.project_service as psv
    import dashboard.services.partner_names as pn
    import dashboard.utils.redis_client as rcm
    recs = [{'프로젝트 코드': 'R1-X', '사업자명': '온라인'}, {'프로젝트 코드': 'R2-X', '사업자명': '한'}]
    monkeypatch.setattr(psv, 'get_project_records', lambda *a, **k: recs)
    writes = []
    mgr = type('M', (), {'batch_update_fields': lambda self, *a, **k: writes.append(a) or {'success': True}})()
    monkeypatch.setattr(psv, 'get_sheets_manager', lambda: mgr)
    monkeypatch.setattr(psv, 'invalidate_project_cache', lambda *a, **k: None)
    monkeypatch.setattr(rcm, 'get_redis_client', lambda: type('C', (), {'redis': type('R', (), {'setex': lambda *a: None})()})())
    monkeypatch.setattr(pn, 'apply_place_on_biz_change', lambda *a, **k: 0)
    monkeypatch.setenv('GOOGLE_SHEET_ID', 'SID')
    monkeypatch.setenv('GOOGLE_SHEET_NAME', '공사 현황')
    assert blh._maybe_update_business_name('R1-X', '데이즈헤어') == ('saved', '')
    assert writes and writes[0][3] == {'사업자명': '데이즈헤어'}
    assert blh._maybe_update_business_name('R2-X', '다른상호') == ('mismatch', '한')   # 한 글자 실제 상호는 지킴



# ── 부담자 이름 = 실제 발행 상호(대표자) (2026-10-28) ─────────────────

def test_official_payer_name():
    from dashboard.services.payer_split import official_payer_name as on
    assert on('위너스 골프존', '기중현', '기종원(위너스골프)') == '위너스 골프존(기중현)'
    assert on('제이케이디 영재학원', '김현준', '주식회사 제이케이디학원') == '제이케이디 영재학원(김현준)'
    assert on('주식회사 제이케이디학원', '김현준(단독 대표)', '') == '주식회사 제이케이디학원(김현준)'
    assert on('이현주', '이현주', '이현주') == '이현주'                       # 개인 사업자 = 이름만
    assert on('다정약국', '', '김지민(다정약국)') == '다정약국(김지민)'        # 대표자 모르면 지금 이름의 사람 부분
    assert on('서울아산가정의학과의원', '', '서울아산가정의학과의원') == '서울아산가정의학과의원'
    assert on('', '홍길동', '홍길동상사') == '홍길동상사'


def _split_rec(name, bno):
    return {'프로젝트 코드': 'G4155-JW', '계약금': '0', '잔금': '0',
            '계약금 부담자': '임대인 · 최영희 · 132-38-00668 · 19,791,200 · 계산서', '중도금 부담자': '-',
            '잔금 부담자': f"임차인 · {name}{' · ' + bno if bno else ''} · 18,471,200 · 계산서"}


def test_fill_payer_renames_to_issued_buyer(monkeypatch):
    import dashboard.blueprints.slack_bot as sb
    import dashboard.services.project_service as psvc
    import dashboard.services.payer_split as psp
    saved = []
    monkeypatch.setattr(psvc, 'get_project_records', lambda force_refresh=False: [_split_rec('기종원(위너스골프)', '443-05-03193')])
    monkeypatch.setattr(ps, 'get_partner_info', lambda n: {'ceo': '기중현'} if n == '위너스 골프존' else None)
    monkeypatch.setattr(psp, 'save_payers', lambda code, plan: saved.append(plan) or {})
    sb._fill_payer_from_invoice('G4155-JW', '잔금', '기종원(위너스골프)', '', bno='443-05-03193',
                                buyer_name='위너스 골프존', buyer_ceo='기현')    # OCR 대표 '기현'보다 거래처 탭 '기중현'
    (plan,), = [saved]
    assert plan['잔금']['name'] == '위너스 골프존(기중현)' and plan['계약금']['name'] == '최영희'


def test_fill_payer_keeps_name_when_issued_to_other_business(monkeypatch):
    """칸 번호와 발행 번호가 다르면(수정발행·다른 사업자) 이름을 바꾸지 않음 — G4148 같은 건은 사람이 정정."""
    import dashboard.blueprints.slack_bot as sb
    import dashboard.services.project_service as psvc
    import dashboard.services.payer_split as psp
    saved = []
    monkeypatch.setattr(psvc, 'get_project_records',
                        lambda force_refresh=False: [_split_rec('제이케이디 영재학원', '405-81-13558')])
    monkeypatch.setattr(psp, 'save_payers', lambda code, plan: saved.append(plan) or {})
    sb._fill_payer_from_invoice('G4155-JW', '잔금', '제이케이디 영재학원', '', bno='310-96-05190',
                                buyer_name='제이케이디 영재학원', buyer_ceo='김현준')
    assert saved == []


def test_normalize_payer_names_by_bno(monkeypatch):
    from dashboard.services.payer_split import normalize_payer_names
    monkeypatch.setattr(ps, 'get_partner_name_by_bno', lambda b: '다정약국' if b == '799-29-00421' else None)
    monkeypatch.setattr(ps, 'get_partner_info', lambda n: {'ceo': '김지민'} if n == '다정약국' else None)
    plan = {'잔금': {'name': '김지민', 'biz_no': '799-29-00421'}, '계약금': {'name': '임기환', 'biz_no': ''}}
    out = normalize_payer_names(plan)
    assert out['잔금']['name'] == '다정약국(김지민)' and out['계약금']['name'] == '임기환'
