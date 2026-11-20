"""나라장터 입찰 공고(2026-11-18) — 지역·면허 묶음·까다로운 조건 분류, 모으기(정정공고), 아침 DM, 화면 API."""
import json
from datetime import datetime, timedelta

import pytest

from dashboard.services import g2b_bids as g

MECH, REFRIG, ELEC = '기계설비ㆍ가스공사업', '냉매회수업', '전기공사업'


@pytest.fixture
def db(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'users.db'))
    monkeypatch.setattr(g, '_db', lambda: d)
    monkeypatch.setattr(g, '_table_ready', False)
    monkeypatch.setattr(g, '_profile_cache', {'at': 0.0, 'val': None})   # 회사 정보는 30초 기억 — 테스트마다 새로
    monkeypatch.setenv('G2B_SERVICE_KEY', 'test-key')
    monkeypatch.delenv('G2B_LICENSES', raising=False)
    monkeypatch.delenv('G2B_LICENSES_PLANNED', raising=False)
    return d


def _future(days=3):
    return (datetime.now() + timedelta(days=days)).strftime('%Y-%m-%d %H:%M:%S')


def _item(no='R26BK0001', ord_='000', title='OO초 냉난방개선 부대공사', **kw):
    it = {'bidNtceNo': no, 'bidNtceOrd': ord_, 'bidNtceNm': title, 'ntceInsttNm': '경기도김포교육지원청',
          'dminsttNm': '경기도김포교육지원청', 'cntrctCnclsMthdNm': '수의계약',
          'sucsfbidMthdNm': '소액수의견적-소액수의견적(2인 이상 견적 제출)', 'presmptPrce': '58208182',
          'bidNtceDt': '2026-11-17 10:00:00', 'bidClseDt': _future(), 'opengDt': _future(),
          'bidNtceDtlUrl': f'https://www.g2b.go.kr/link/PNPE027_01/single/?bidPbancNo={no}&bidPbancOrd={ord_}',
          'ntceSpecDocUrl1': 'https://www.g2b.go.kr/x1', 'ntceSpecFileNm1': '입찰공고문.hwp',
          'arsltCmptYn': 'N', 'pqEvalYn': '', 'dsgntCmptYn': 'N'}
    it.update(kw)
    return it


# ── 분류 ──────────────────────────────────────────────────────────────
def test_region_needs_gyeonggi_whole_or_gimpo():
    assert g.region_ok(['경기도'])
    assert g.region_ok(['경기도 고양시', '경기도 김포시', '경기도 파주시'])
    assert g.region_ok(['경기도', '서울특별시', '인천광역시'])
    assert not g.region_ok(['경기도 군포시', '경기도 수원시'])   # 시 단위 제한 — 김포 업체는 못 들어감
    assert not g.region_ok(['인천광역시'])
    assert not g.region_ok([])


def test_license_group_is_and_inside_or_between():
    # 같은 묶음 = 모두 있어야
    ok, miss = g.license_status([[MECH, REFRIG]], [MECH], [REFRIG])
    assert not ok and miss == [REFRIG]
    # 묶음끼리는 하나만 맞으면
    assert g.license_status([[ELEC], [MECH]], [MECH])[0]
    # 가운뎃점·업종코드 달라도 같은 면허
    assert g.license_status([['기계설비·가스공사업/6202']], [MECH])[0]
    # 제한 없음
    assert g.license_status([], [MECH]) == (True, [])
    # 준비 중 면허로도 안 되면 모자란 목록 없음
    assert g.license_status([[ELEC]], [MECH], [REFRIG]) == (False, [])


@pytest.mark.parametrize('kw,reason', [
    ({'arsltCmptYn': 'Y'}, '실적제한'),
    ({'pqEvalYn': 'Y'}, '사전심사(PQ)'),
    ({'cntrctCnclsMthdNm': '지명경쟁'}, '지명경쟁'),
    ({'sucsfbidMthdNm': '적격심사제-추정가격이 10억원 미만 4억원 이상', 'presmptPrce': '450000000'}, '4억 이상 적격심사'),
    ({'sucsfbidMthdNm': '적격심사제-추정가격이 4억원 미만 2억원 이상', 'presmptPrce': '340000000'}, ''),
    ({'sucsfbidMthdNm': '최저가낙찰제', 'presmptPrce': '454545455'}, ''),   # 최저가는 금액 커도 남김
])
def test_hard_reason(kw, reason):
    assert g.hard_reason(_item(**kw)) == reason


def test_classify_categories():
    held, planned = [MECH], [REFRIG]
    assert g.classify(_item(), ['경기도'], [[MECH]], held, planned)['category'] == g.ELIGIBLE
    c = g.classify(_item(), ['경기도 김포시'], [[MECH, REFRIG]], held, planned)
    assert c['category'] == g.LICENSE_ONLY and c['missing'] == [REFRIG] and c['reason'] == '냉매회수업 면허 없음'
    assert g.classify(_item(), [], [[MECH]], held, planned)['category'] == g.ELIGIBLE        # 전국 공고 = 가능(11-18)
    c = g.classify(_item(bidNtceNm='파주송화초 냉난방기 전기공사'), [], [], held, planned)
    assert c['category'] == g.OTHER and c['reason'] == '면허 안 맞음(전기공사업)'           # 면허 제한 없어도 전기공사는 안 됨
    assert g.classify(_item(bidNtceNm='대학본부1층, 3층 냉난방기 공사'), [], [], held, planned)['category'] == g.ELIGIBLE
    c = g.classify(_item(), ['경기도 수원시'], [[MECH]], held, planned)
    assert c['category'] == g.OTHER and c['reason'] == '김포 업체 참가 불가(지역 제한)'
    assert g.classify(_item(arsltCmptYn='Y'), ['경기도'], [[MECH]], held, planned)['reason'] == '실적제한 있음'
    c = g.classify(_item(), ['경기도'], [[ELEC]], held, planned)
    assert c['category'] == g.OTHER and c['reason'] == '면허 안 맞음(전기공사업)'
    # 사유가 여럿이면 모두 — 지역도 안 되고 냉매회수업도 없으면 '면허만 없음'이 아님
    c = g.classify(_item(), ['인천광역시'], [[MECH, REFRIG]], held, planned)
    assert c['category'] == g.OTHER and c['reason'] == '김포 업체 참가 불가(지역 제한) · 냉매회수업 면허 없음'


# ── 모으기 ────────────────────────────────────────────────────────────
def _fake_api(monkeypatch, items, limits, nuri=None, apt=None):
    nuri = nuri or {}
    apt = apt or []
    calls = {'limits': 0}

    def search(kw, start, end, op=g.OP_LIST):
        assert (end - start) < timedelta(days=7)   # 7일 넘으면 나라장터 07 오류
        if kw != g.KEYWORDS[0]:
            return []
        want = '물품' if op == g.OP_LIST_THNG else '공사'   # 테스트 공고는 'T_kind'로 종류 지정(없으면 공사)
        return [{k: v for k, v in i.items() if k != 'T_kind'} for i in items if i.get('T_kind', '공사') == want]

    def fetch_limits(no, ord_, source='나라장터'):
        calls['limits'] += 1
        return limits[no]
    monkeypatch.setattr(g, 'search', search)
    monkeypatch.setattr(g, 'fetch_limits', fetch_limits)
    monkeypatch.setattr(g, 'nuri_list', lambda op, start, end: [dict(i) for i in nuri.get(op, [])])
    monkeypatch.setattr(g, 'apt_list', lambda start, end: [dict(i) for i in apt])
    return calls


def test_collect_saves_and_skips_known(db, monkeypatch):
    items = [_item('A1'), _item('B2', title='수원 냉난방', ), _item('C3', title='김포 냉난방 철거')]
    limits = {'A1': (['경기도'], [[MECH]]), 'B2': (['경기도 수원시'], [[MECH]]),
              'C3': (['경기도 김포시'], [[MECH, REFRIG]])}
    calls = _fake_api(monkeypatch, items, limits)
    r = g.collect(days=7)
    assert r['new'] == 3 and calls['limits'] == 3
    data = g.list_bids(g.ELIGIBLE)
    assert [x['bid_no'] for x in data['items']] == ['A1']
    assert data['counts'] == {g.ELIGIBLE: 1, g.BLOCKED: 2, g.LICENSE_ONLY: 1}
    blocked = g.list_bids(g.BLOCKED)['items']
    assert [x['bid_no'] for x in blocked] == ['C3', 'B2']          # 면허만 없는 공고가 먼저
    assert blocked[0]['reason'] == '냉매회수업 면허 없음' and not blocked[1]['region_ok']
    assert data['items'][0]['files'][0]['name'] == '입찰공고문.hwp'
    assert data['last_collect']
    # 다시 모아도 이미 있는 건 지역·면허를 다시 부르지 않는다
    g.collect(days=2)
    assert calls['limits'] == 3


def test_correction_notice_updates_but_keeps_mark(db, monkeypatch):
    calls = _fake_api(monkeypatch, [_item('A1')], {'A1': (['경기도'], [[MECH]])})
    g.collect()
    g.set_mark('A1', 'review', '고광일')
    _fake_api(monkeypatch, [_item('A1', ord_='001', title='정정 냉난방')], {'A1': (['경기도'], [[MECH]])})
    r = g.collect()
    assert r['changed'] == 1
    it = g.list_bids()['items'][0]
    assert it['title'] == '정정 냉난방' and it['bid_ord'] == '001' and it['mark'] == 'review'
    assert calls['limits'] == 1


def test_reclassify_after_new_license(db, monkeypatch):
    _fake_api(monkeypatch, [_item('C3')], {'C3': (['경기도 김포시'], [[MECH, REFRIG]])})
    g.collect()
    assert g.list_bids(g.BLOCKED)['counts'][g.LICENSE_ONLY] == 1
    monkeypatch.setenv('G2B_LICENSES', f'{MECH},{REFRIG}')
    assert g.reclassify() == 1
    assert [x['bid_no'] for x in g.list_bids(g.ELIGIBLE)['items']] == ['C3']


def test_closed_hidden_unless_asked(db, monkeypatch):
    past = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d %H:%M:%S')
    _fake_api(monkeypatch, [_item('OLD', bidClseDt=past)], {'OLD': (['경기도'], [[MECH]])})
    g.collect()
    assert g.list_bids()['items'] == []
    assert g.list_bids(include_closed=True)['items'][0]['closed']


def test_missing_close_falls_back_to_opening(db, monkeypatch):
    _fake_api(monkeypatch, [_item('N1', bidClseDt='')], {'N1': (['경기도'], [[MECH]])})
    g.collect()
    assert g.list_bids()['items'][0]['bid_no'] == 'N1'


def test_mark_validation(db, monkeypatch):
    _fake_api(monkeypatch, [_item('A1')], {'A1': (['경기도'], [[MECH]])})
    g.collect()
    with pytest.raises(ValueError):
        g.set_mark('A1', 'bogus', 'x')
    with pytest.raises(LookupError):
        g.set_mark('NOPE', 'join', 'x')
    assert g.set_mark('A1', 'join', '고광일')['mark_label'] == '참여'
    assert g.set_mark('A1', '', '고광일')['mark_by'] == ''


def test_detail_from_raw(db, monkeypatch):
    _fake_api(monkeypatch, [_item('A1', bidBeginDt='2026-11-16 10:00:00', bidQlfctRgstDt='2026-11-17 18:00',
                                  sucsfbidLwltRate='89.745', prearngPrceDcsnMthdNm='복수예가', totPrdprcNum='15',
                                  drwtPrdprcNum='4', VAT='34049272', cmmnSpldmdMethdNm='(없음)공동수급불허',
                                  ntceInsttOfclNm='신수견', ntceInsttOfclTelNo='031-980-1137',
                                  ntceInsttOfclEmailAdrs='a@korea.kr', rgnDutyJntcontrctYn='N')],
              {'A1': (['경기도'], [[MECH]])})
    g.collect()
    d = g.list_bids()['items'][0]['detail']
    assert d['begin_at'] == '2026-11-16 10:00' and d['qual_at'] == '2026-11-17 18:00'
    assert d['lower_rate'] == '89.745%' and d['price_count'] == '15개 중 4개 추첨' and d['vat'] == 34049272
    assert d['contact_name'] == '신수견' and d['contact_email'] == 'a@korea.kr' and d['joint_region'] == '아니오'


def test_files_sorted_by_leading_number():
    names = ['03_입찰조서.hwp', '04_내역서.xlsx', '01_입찰공고서.pdf', '★시방서.hwp', '02_과업지시서.pdf', '10_도면.pdf']
    got = [f['name'] for f in g.sort_files([{'name': n, 'url': 'u'} for n in names])]
    assert got == ['01_입찰공고서.pdf', '02_과업지시서.pdf', '03_입찰조서.hwp', '04_내역서.xlsx', '10_도면.pdf', '★시방서.hwp']


def test_single_price_has_no_draw_count():
    d = g._detail(json.dumps({'prearngPrceDcsnMthdNm': '단일예가', 'totPrdprcNum': '1', 'drwtPrdprcNum': '0'}))
    assert d['price_method'] == '단일예가' and d['price_count'] == ''


def test_old_rows_get_detail_on_next_collect_without_limit_calls(db, monkeypatch):
    calls = _fake_api(monkeypatch, [_item('A1')], {'A1': (['경기도'], [[MECH]])})
    g.collect()
    with g._conn() as conn:   # 처음 버전(상세 칸 없음)으로 저장된 줄 흉내
        conn.execute("UPDATE g2b_bids SET raw='' WHERE bid_no='A1'")
    assert g.list_bids()['items'][0]['detail'] == {}
    _fake_api(monkeypatch, [_item('A1', sucsfbidLwltRate='87.745')], {'A1': (['경기도'], [[MECH]])})
    g.collect()
    assert g.list_bids()['items'][0]['detail']['lower_rate'] == '87.745%'


def test_raw_column_added_to_old_table(tmp_path, monkeypatch):
    import sqlite3
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'u.db'))
    with d._get_connection() as conn:   # 11-18 첫 버전 표(raw 없음)
        conn.execute(g._DDL[0])
        conn.commit()
    monkeypatch.setattr(g, '_db', lambda: d)
    monkeypatch.setattr(g, '_table_ready', False)
    with g._conn() as conn:
        cols = {r[1] for r in conn.execute('PRAGMA table_info(g2b_bids)')}
    assert 'raw' in cols


# ── 아침 DM ───────────────────────────────────────────────────────────
def test_digest_only_when_new_eligible_and_once(db, monkeypatch):
    sent = []
    monkeypatch.setattr(g, '_post_slack', lambda ch, text: sent.append((ch, text)) or True)
    _fake_api(monkeypatch, [_item('C3')], {'C3': (['경기도 김포시'], [[MECH, REFRIG]])})
    g.collect()
    r = g.send_morning_digest(force=True)
    assert r['sent'] is False and sent == []            # 면허만 없는 공고만 있으면 안 보냄

    _fake_api(monkeypatch, [_item('A1', title='정평중 냉난방개선 부대공사')], {'A1': (['경기도'], [[MECH]])})
    g.collect()
    r = g.send_morning_digest(force=True)
    assert r == {'sent': True, 'eligible': 1, 'license_only': 1}
    ch, text = sent[0]
    assert ch == 'U04UL2ZLJAX'
    assert '정평중 냉난방개선 부대공사' in text and '나라장터에서 보기' in text and '|공고문 HWP>' in text
    assert '면허만 없어서 못 들어가는 새 공고도 1건' in text and '/bids|그룹웨어에서 전체 보기' in text
    # 같은 공고로 다시 보내지 않는다
    assert g.send_morning_digest(force=True)['sent'] is False
    assert len(sent) == 1


def test_briefing_in_detail_and_digest(db, monkeypatch):
    sent = []
    monkeypatch.setattr(g, '_post_slack', lambda ch, text: sent.append(text) or True)
    soon_ = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d 14:00:00')
    _fake_api(monkeypatch, [_item('A1', dcmtgOprtnDt=soon_, dcmtgOprtnPlce='본관 2층 대회의실'),
                            _item('A2', title='냉난방 설명회 없음')],
              {'A1': (['경기도'], [[MECH]]), 'A2': (['경기도'], [[MECH]])})
    g.collect()
    d = {x['bid_no']: x['detail'] for x in g.list_bids()['items']}
    assert d['A1']['briefing_at'] == soon_[:16] and d['A1']['briefing_place'] == '본관 2층 대회의실'
    assert d['A2']['briefing_at'] == ''
    g.send_morning_digest(force=True)
    assert sent[0].count('현장설명회') == 1 and '본관 2층 대회의실' in sent[0] and '이미 지났어요' not in sent[0]


def test_briefing_line_says_when_passed():
    past = (datetime.now() - timedelta(hours=2)).strftime('%Y-%m-%d %H:%M:%S')
    line = g._briefing_line({'raw': json.dumps({'dcmtgOprtnDt': past})})
    assert '이미 지났어요' in line


def test_digest_skips_weekend(db):
    from datetime import date
    assert g.send_morning_digest(today=date(2026, 11, 21)) == {'skipped': '영업일 아님'}   # 토요일


def test_won_format():
    assert g._won(116415455) == '1억 1,641만 원'
    assert g._won(58208182) == '5,820만 원'
    assert g._won(400000000) == '4억 원'
    assert g._won(1) == '비공개(공고문 확인)'   # 발주처가 1원으로 넣은 금액 비공개(11-19 동원대학)


def test_home_summary_excludes_skip(db, monkeypatch):
    _fake_api(monkeypatch, [_item('A1'), _item('A2')], {'A1': (['경기도'], [[MECH]]), 'A2': (['경기도'], [[MECH]])})
    g.collect()
    assert g.home_summary()[0] == 2
    g.set_mark('A2', 'skip', 'x')
    n, sig = g.home_summary()
    assert n == 1 and sig == 'A1'


def test_error_text_hides_service_key(db, monkeypatch):
    monkeypatch.setenv('G2B_SERVICE_KEY', 'abc+/=123secret')
    e = RuntimeError('Max retries exceeded with url: /x?serviceKey=abc%2B%2F%3D123secret&type=json (timed out)')
    s = g._safe_err(e)
    assert 'secret' not in s and 'serviceKey=***' in s
    assert 'secret' not in g._safe_err(RuntimeError('key abc+/=123secret'))


def test_disabled_without_key(db, monkeypatch):
    monkeypatch.delenv('G2B_SERVICE_KEY')
    assert g.collect() == {'skipped': 'G2B_SERVICE_KEY 없음'}


# ── 메뉴 ──────────────────────────────────────────────────────────────
def test_menu_has_bids_for_office_not_field():
    from dashboard.services import groupware as gw
    nav = gw.build_nav('/bids', 'user')
    sales = next(a for a in nav['areas'] if a['key'] == 'sales')
    assert any(c['href'] == '/bids' and c['active'] for c in sales['children'])
    field = gw.build_nav('/bids', 'user', field=True)
    assert not any(c['href'] == '/bids' for a in field['areas'] for c in a['children'])


def test_new_badge_by_posted_date(db, monkeypatch):
    fresh = (datetime.now() - timedelta(hours=3)).strftime('%Y-%m-%d %H:%M:%S')
    old = (datetime.now() - timedelta(days=5)).strftime('%Y-%m-%d %H:%M:%S')
    _fake_api(monkeypatch, [_item('F1', bidNtceDt=fresh), _item('O1', bidNtceDt=old)],
              {'F1': (['경기도'], [[MECH]]), 'O1': (['경기도'], [[MECH]])})
    g.collect()   # 둘 다 방금 모았지만 New 는 게시 24시간 안인 것만
    assert {x['bid_no']: x['is_new'] for x in g.list_bids()['items']} == {'F1': True, 'O1': False}


def test_bootstrap_reclassifies_existing_rows(db, monkeypatch):
    _fake_api(monkeypatch, [_item('C3')], {'C3': (['경기도 김포시'], [[MECH, REFRIG]])})
    g.collect()
    with g._conn() as conn:   # 옛 사유 문구로 저장된 줄 흉내
        conn.execute("UPDATE g2b_bids SET reason='면허 없음: 냉매회수업' WHERE bid_no='C3'")
    monkeypatch.setattr(g, 'collect', lambda **k: pytest.fail('표가 차 있으면 다시 모으지 않음'))
    g.maybe_bootstrap()
    assert g.list_bids(g.BLOCKED)['items'][0]['reason'] == '냉매회수업 면허 없음'


def test_reclassify_updates_old_pq_wording(db, monkeypatch):
    _fake_api(monkeypatch, [_item('P1', pqEvalYn='Y')], {'P1': (['경기도'], [[MECH]])})
    g.collect()
    with g._conn() as conn:   # 11-18 앞선 표기로 저장된 줄
        conn.execute("UPDATE g2b_bids SET hard_reason='PQ(사전심사)', reason='PQ(사전심사) 있음' WHERE bid_no='P1'")
    g.reclassify()
    with g._conn() as conn:
        r = conn.execute("SELECT hard_reason, reason FROM g2b_bids WHERE bid_no='P1'").fetchone()
    assert r[0] == '사전심사(PQ)' and '사전심사(PQ) 있음' in r[1]


def test_keywords_catch_dotted_titles():
    # '냉·난방', '냉.난방' 은 '냉난방'으로는 안 잡힘 → 검색어에 '난방'이 있어야
    assert '난방' in g.KEYWORDS and '냉난방' not in g.KEYWORDS
    for t in ('냉·난방기 교체', '냉.난방기 교체 공사', '냉난방 개선', '냉 난방'):
        assert any(k in t for k in g.KEYWORDS)
    assert {'공기조화', '냉각탑'} <= set(g.KEYWORDS)


def test_unrelated_heating_titles_are_not_eligible():
    held, planned = [MECH], [REFRIG]
    c = g.classify(_item(title='파주지사 지역난방설비동 채광창 교체공사'), ['경기도'], [[MECH]], held, planned)
    assert c['category'] == g.OTHER and c['reason'] == '냉난방 공사 아님(제목)'
    assert g.is_relevant('한국지역난방공사 사옥 냉난방기 교체')      # 지역난방 말고 냉난방도 있으면 관련
    assert g.is_relevant('냉.난방기 교체') and g.is_relevant('의정관 냉각탑 교체')
    assert not g.is_relevant('2027년 재난방송온라인시스템 유지관리')


# ── 물품(구매·설치) 11-19 ─────────────────────────────────────────────
def _thng(no='T1', title='[긴급] 추산초 냉난방 구매 및 설치', **kw):
    it = _item(no, title=title, cntrctCnclsMthdNm='일반경쟁', presmptPrce='800106363', asignBdgtAmt='880117000',
               sucsfbidMthdNm='적격심사제-추정가격이 10억원 미만 고시금액 이상인 물품의 평가기준(물품구매)',
               mnfctYn='N', prdctClsfcLmtYn='Y', dlvrDaynum='80', dlvryCndtnNm='현장설치도',
               purchsObjPrdctList='[1^4010178702^냉난방기],[1^4010180601^전기히트펌프],[1^4010178701^히트펌프용실내기]',
               T_kind='물품')
    it.update(kw)
    return it


def test_thng_install_collected_and_eligible_even_over_400m(db, monkeypatch):
    _fake_api(monkeypatch, [_thng('T1'), _thng('T2', title='OO중 냉난방기 구매'), _item('A1')],
              {'T1': ([], [[MECH]]), 'T2': ([], [[MECH]]), 'A1': (['경기도'], [[MECH]])})
    r = g.collect()
    assert r['new'] == 2                                   # 설치 없는 단순 구매(T2)는 저장도 안 함
    items = {x['bid_no']: x for x in g.list_bids()['items']}
    t = items['T1']
    assert t['bid_kind'] == '물품' and items['A1']['bid_kind'] == '공사'
    assert t['category'] == g.ELIGIBLE                     # 8억 물품 적격심사여도 공사용 '4억' 기준은 안 씀
    assert t['detail']['products'] == ['냉난방기(4010178702)', '전기히트펌프(4010180601)', '히트펌프용실내기(4010178701)']
    assert t['detail']['budget'] == 880117000 and t['detail']['deliver_days'] == '80'


def test_thng_maker_only_is_blocked():
    held, planned = [MECH], [REFRIG]
    c = g.classify({**_thng(mnfctYn='Y'), '_kind': '물품'}, [], [[MECH]], held, planned)
    assert c['category'] == g.OTHER and '제조업체만 참가' in c['reason']
    assert g.classify({**_item(sucsfbidMthdNm='적격심사제', presmptPrce='450000000'), '_kind': '공사'},
                      ['경기도'], [[MECH]], held, planned)['reason'] == '4억 이상 적격심사'


def test_digest_lists_thng_without_kind_tag(db, monkeypatch):
    sent = []
    monkeypatch.setattr(g, '_post_slack', lambda ch, text: sent.append(text) or True)
    _fake_api(monkeypatch, [_thng('T1')], {'T1': ([], [[MECH]])})
    g.collect()
    g.send_morning_digest(force=True)
    assert '*[긴급] 추산초 냉난방 구매 및 설치*' in sent[0] and '[물품]' not in sent[0]   # 딱지 없이(11-19 부장님)


def test_old_rows_default_to_construction(tmp_path, monkeypatch):
    from dashboard.utils.user_database import UserDatabase
    d = UserDatabase(db_path=str(tmp_path / 'u.db'))
    with d._get_connection() as conn:
        conn.execute(g._DDL[0])
        conn.execute("INSERT INTO g2b_bids (bid_no, first_seen, updated_at) VALUES ('X1', 'a', 'b')")
        conn.commit()
    monkeypatch.setattr(g, '_db', lambda: d)
    monkeypatch.setattr(g, '_table_ready', False)
    with g._conn() as conn:
        assert conn.execute("SELECT kind FROM g2b_bids WHERE bid_no='X1'").fetchone()[0] == '공사'


def test_area_of_for_area_filter():
    assert g.area_of('', '경기도교육청 경기도포천교육지원청') == '수도권'
    assert g.area_of('', '서울교통공사') == '수도권'
    assert g.area_of('경기도 광주시', '') == '수도권'                     # '광주'보다 '경기'가 먼저
    assert g.area_of('', '전남광주통합특별시교육청순천만생태문화교육원') == '호남'
    assert g.area_of('', '대전광역시교육청', '조달청 대전지방조달청') == '충청'
    assert g.area_of('', '경상남도교육청 경상남도남해교육지원청') == '영남'
    assert g.area_of('', '계명대학교', '계명대학교') == '기타'
    assert g.area_of('경기도 용인시 처인구', '부산광역시교육청') == '수도권'   # 공사현장이 먼저



# ── 누리장터 민간입찰 11-20 ───────────────────────────────────────────
def _nuri(no='N1', title='지구촌교회(수지채플)시스템에어컨 교체공사 입찰', **kw):
    it = {'bidNtceNo': no, 'bidNtceOrd': '000', 'ntceNm': title, 'nticeDt': '2026-11-19 18:08:26',
          'ntceInsttNm': '지구촌교회', 'cntrctMthdNm': '일반경쟁', 'sucsfbidMthdNm': '최저가낙찰제',
          'bidQlfctNm': '공고서 참조', 'rgnLmtDivNm': '공고서 참조', 'ofclNm': '이승호', 'ofclTelNo': '031-000-0000',
          'bidClseDt': _future(), 'opengDt': _future(), 'refAmt': '7150000', 'asignBdgtAmt': '0',
          'sptDscrptDt': '2026-11-24 10:00:00', 'sptDscrptPlce': '본당 1층',
          'ntceSpecDocUrl1': 'https://www.g2b.go.kr/x', 'ntceSpecDocNm1': '입찰공고서.hwp',
          'cnstwkDtlList': '[1^시스템에어컨 교체공사^경기도 용인시 수지구 ...^2026-12-31]'}
    it.update(kw)
    return it


def test_nuri_collected_normalized_and_filtered(db, monkeypatch):
    nuri = {g.NURI_KINDS[0][1]: [_nuri('N1'), _nuri('N2', ntceNm='본관 도장공사')],
            g.NURI_KINDS[1][1]: [_nuri('N3', ntceNm='항온항습기 구매')]}
    _fake_api(monkeypatch, [], {'N1': ([], [])}, nuri)
    r = g.collect()
    assert r['new'] == 1                                   # 도장공사(무관)·항온항습기 구매(설치 없음)는 안 넣음
    it = g.list_bids()['items'][0]
    assert it['source'] == '누리장터' and it['title'].startswith('지구촌교회') and it['category'] == g.ELIGIBLE
    assert it['price'] == 7150000 and it['contact'] == '이승호 031-000-0000' and it['area'] == '수도권'
    assert it['detail']['briefing_at'] == '2026-11-24 10:00' and it['files'][0]['name'] == '입찰공고서.hwp'
    assert it['detail_url'] == ''                           # 나라장터 상세 주소가 빈 화면이라 저장 안 함


def test_nuri_digest_tag(db, monkeypatch):
    sent = []
    monkeypatch.setattr(g, '_post_slack', lambda ch, text: sent.append(text) or True)
    _fake_api(monkeypatch, [], {'N1': ([], [])}, {g.NURI_KINDS[0][1]: [_nuri('N1')]})
    g.collect()
    g.send_morning_digest(force=True)
    assert '지구촌교회 (누리장터)' in sent[0]



# ── 공동주택(K-apt) 11-20 ─────────────────────────────────────────────
def _apt(num='20261117000000001', title='휘트니스센터 천장형 시스템 냉난방기 설치 공사 업체 선정', **kw):
    it = {'bidNum': num, 'aptCode': 'A1', 'bidArea': 41, 'bidTitle': title, 'bidContent': '첨부 참조',
          'bidRegDate': '2026-11-19', 'bidDeadline': _future(), 'bidKaptname': '목동올리브힐', 'bidState': '1',
          'codeAuth': '01', 'codeKind': '01', 'codeClassifyType1': '02', 'codeClassifyType2': '02',
          'codeSucWay': '03', 'bidFieldDesDate': '2026-11-23 14:00:00', 'bidFieldDesLoc': '관리사무소',
          'bidMngCertYn': 'Y', 'bidCreditConfYn': 'N'}
    it.update(kw)
    return it


def test_apt_collected_with_kapt_link_and_filters(db, monkeypatch):
    apt = [_apt('1'), _apt('2', bidTitle='난방배관 설비보호제 납품 및 투입업체 선정'),
           _apt('3', bidTitle='난방배관 교체공사 감리업자 선정'), _apt('4', bidTitle='주택관리업체 선정'),
           _apt('5', bidTitle='전기실 에어컨 설치', codeKind='03')]
    _fake_api(monkeypatch, [], {}, apt=apt)
    r = g.collect()
    assert r['new'] == 2                                    # 설비보호제·감리·무관은 안 넣음
    items = {x['bid_no']: x for x in g.list_bids(g.ELIGIBLE)['items'] + g.list_bids(g.BLOCKED)['items']}
    a = items['KAPT-1']
    assert a['source'] == '공동주택' and a['category'] == g.ELIGIBLE and a['area'] == '수도권'
    assert a['detail_url'] == 'https://www.k-apt.go.kr/bid/bidDetail.do?bidNum=1'
    assert a['contract_method'] == '일반경쟁' and a['detail']['briefing_at'] == '2026-11-23 14:00'
    assert a['detail']['apt_perf_cert'] == '예' and a['detail']['apt_auth'] == 'K-apt'
    assert items['KAPT-5']['category'] == g.OTHER and '지명경쟁' in items['KAPT-5']['reason']


def test_apt_digest_tag(db, monkeypatch):
    sent = []
    monkeypatch.setattr(g, '_post_slack', lambda ch, text: sent.append(text) or True)
    _fake_api(monkeypatch, [], {}, apt=[_apt('1')])
    g.collect()
    g.send_morning_digest(force=True)
    assert '목동올리브힐 (공동주택)' in sent[0] and 'k-apt.go.kr/bid/bidDetail.do?bidNum=1|K-apt에서 보기' in sent[0]


def test_apt_area_12_is_jeonnam_gwangju():
    it = g.apt_normalize(_apt('9', bidArea=12))
    assert it['cnstrtsiteRgnNm'] == '전남광주통합특별시' and g.area_of(it['cnstrtsiteRgnNm']) == '호남'



# ── 우리 회사 정보·적격심사 85점 11-20 ───────────────────────────────
def test_qual_need_table2_by_credit_grade():
    need = lambda gr, b=0: g.qual_need('T2', gr, b, False, 80.495)['need_pct']   # noqa: E731
    assert need('A-') == 80.5 and need('BBB-') == 80.85 and need('B+') == 81.5 and need('없음') == 85.5
    assert need('B+', 1) == 81.0                                      # 신인도 +1 이면 0.5%p 내려감
    assert g.qual_need('T3', 'B+', 0, False, 84.245)['need_pct'] == 84.5
    assert g.qual_need('T3', '없음', 0, True, 84.245)['need_pct'] == 84.25   # 창업 7년 이내 = 신용 만점
    assert g.qual_need('T1', 'B+') == {} and g.qual_need('T2', '') == {}


def test_qual_table_only_local_goods():
    m = '적격심사제-추정가격이 10억원 미만 고시금액 이상인 물품의 평가기준(물품구매)'
    local = '지방자치단체 입찰시 낙찰자 결정기준'
    assert g.qual_table({'_kind': '물품', 'sucsfbidMthdNm': m, 'sucsfbidMthdAppStd': local}) == 'T2'
    assert g.qual_table({'_kind': '공사', 'sucsfbidMthdNm': m, 'sucsfbidMthdAppStd': local}) == ''
    assert g.qual_table({'_kind': '물품', 'sucsfbidMthdNm': m, 'sucsfbidMthdAppStd': '조달청'}) == ''


def test_profile_save_validates_and_reclassifies(db, monkeypatch):
    _fake_api(monkeypatch, [_item('C3')], {'C3': (['경기도 김포시'], [[MECH, REFRIG]])})
    g.collect()
    assert g.list_bids(g.ELIGIBLE)['items'] == []
    with pytest.raises(ValueError):
        g.save_profile({'credit_grade': 'ZZZ'}, 'x')
    p = g.save_profile({'credit_grade': 'B+', 'licenses': f'{MECH}, {REFRIG}', 'planned': '',
                        'products': '4010178702 냉난방기' + chr(10) + '아무말' + chr(10) + '4010180601 전기히트펌프', 'bonus': '5'}, '고광일')
    assert p['credit_grade'] == 'B+' and p['licenses'] == [MECH, REFRIG] and p['planned'] == []
    assert [x['code'] for x in p['products']] == ['4010178702', '4010180601'] and p['bonus'] == 2.0
    assert [x['bid_no'] for x in g.list_bids(g.ELIGIBLE)['items']] == ['C3']   # 면허 바꾸자 바로 다시 분류


def test_ours_products_and_qual_on_view(db, monkeypatch):
    t = _thng('T1', sucsfbidLwltRate='80.495', sucsfbidMthdAppStd='지방자치단체 입찰시 낙찰자 결정기준')
    _fake_api(monkeypatch, [t], {'T1': ([], [[MECH]])})
    g.collect()
    g.save_profile({'credit_grade': 'B+', 'licenses': MECH, 'products': '4010178702 냉난방기'}, 'x')
    ours = g.list_bids()['items'][0]['ours']
    assert ours['products_known'] and ours['products_missing'] == ['4010180601', '4010178701']
    assert ours['qual']['need_pct'] == 81.5 and ours['qual']['gap_pct'] == 1.0 and ours['grade'] == 'B+'
