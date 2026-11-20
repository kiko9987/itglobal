"""입찰 공고문 자동 읽기(2026-11-20) — 공고문 고르기, 단락·조건 문구 찾기, 신용등급 비교, 읽기·자동 실행."""
import pytest

from dashboard.services import g2b_bids as g
from dashboard.services import g2b_notice as n
from test_g2b_bids import MECH, _fake_api, _item, db  # noqa: F401  (db 픽스처 재사용)

NOTICE_MD = """
| Ⅰ. 물품 구매 입찰공고 |
Ⅲ. 국가종합전자조달시스템 입찰참가자격등록규정
Ⅳ. 지방자치단체 입찰시 낙찰자 결정기준

2. 입찰 및 계약방식

 2.4. **공동도급 및 하도급이 허용되지 않습니다.**

3. 입찰참가자격

3.1. 다음 사항을 모두 갖춘 업체이어야 합니다.

  3.1.2. 「산업표준화법」에 따른 **KS제품인증서(표준번호: KSC9306)**를 소지한 자

  3.1.3. **공급업체는 당해 물품 제조사 또는 대리점의 정품인증서 및 물품공급 및 기술지원[A/S] 확약서를 제출하여야 합니다.**

  3.1.5. 신용평가등급이 BB- 이상인 업체이어야 합니다.

4. 현장설명

 **4.1. 이 사업은 별도의 현장설명을 실시하지 않습니다.**

5. 입찰참가신청 및 입찰보증금

5.2.2. 입찰보증금 납부를 확약하는 내용이 포함된 입찰서를 제출
| 10 | [공사]<br>하도급지킴이<br>이용 확약서 | 하도급지킴이를 이용 |
"""


def test_pick_notice_file_prefers_hwpx_and_skips_specs():
    files = [{'name': '03_공사시방서.hwp', 'url': 'u3'}, {'name': '★ 공고문[추산초].pdf', 'url': 'u1'},
             {'name': '★ 공고문[추산초].hwpx', 'url': 'u2'}, {'name': '내역서(공고).xlsx', 'url': 'u4'}]
    assert n.pick_notice_file(files)['url'] == 'u2'
    assert n.pick_notice_file([{'name': '01_입찰공고서.pdf', 'url': 'a'}])['url'] == 'a'
    assert n.pick_notice_file([{'name': '시방서.hwp', 'url': 'a'}, {'name': '도면.pdf', 'url': 'b'}]) is None


def test_sections_skip_table_of_contents():
    secs = {s['title']: s['lines'] for s in n.sections_of(NOTICE_MD)}
    assert set(secs) == {'입찰참가자격', '현장설명'}
    assert secs['입찰참가자격'][0].startswith('3.1. ')
    assert any('확약서' in x for x in secs['입찰참가자격'])
    assert '**' not in ''.join(secs['입찰참가자격'])          # 굵게 표시 기호는 뺌
    assert secs['현장설명'] == ['4.1. 이 사업은 별도의 현장설명을 실시하지 않습니다.']


def test_flags_found_and_noise_skipped():
    keys = {f['key']: f for f in n.flags_of(NOTICE_MD)}
    assert set(keys) == {'confirm', 'ks', 'credit', 'brief_none', 'nojoint'}
    assert '제조사' in keys['confirm']['line']                    # 입찰보증금 '확약'·하도급지킴이 확약서가 아님
    assert keys['confirm']['level'] == 'warn' and keys['brief_none']['level'] == 'info'


def test_brief_must_hides_brief_none():
    md = '현장설명을 실시하지 않습니다.\n현장설명회에 참석하지 않은 자는 입찰에 참가할 수 없습니다.'
    assert [f['key'] for f in n.flags_of(md)] == ['brief_must']


def test_credit_check_against_our_grade():
    line = '신용평가등급이 BB- 이상인 업체'
    assert n.credit_check(line, 'B+') is False
    assert n.credit_check(line, 'BB-') is True
    assert n.credit_check('신용평가등급 BBB° 이상', 'BBB0') is True
    assert n.credit_check('신용평가등급 확인서 제출', 'B+') is None   # 기준 등급이 없으면 판단 안 함
    assert n.credit_check(line, '') is None


def _seed(monkeypatch, files=None):
    it = _item('N1', ntceSpecFileNm1='★ 공고문.hwpx') if files is None else _item('N1', **files)
    _fake_api(monkeypatch, [it], {'N1': (['경기도'], [[MECH]])})
    g.collect(days=7)


def test_read_notice_saves_and_shows_in_list(db, monkeypatch):
    _seed(monkeypatch)
    from dashboard.services import doc_viewer
    got = {}
    monkeypatch.setattr(n, '_download', lambda url: got.setdefault('url', url) and b'data')
    monkeypatch.setattr(doc_viewer, 'convert_to_markdown', lambda data, ext: got.setdefault('ext', ext) and NOTICE_MD)
    g.save_profile({'credit_grade': 'B+', 'licenses': MECH}, 't')
    v = n.read_notice('N1')
    assert got == {'url': 'https://www.g2b.go.kr/x1', 'ext': 'hwpx'}
    assert v['status'] == 'ok' and v['file_name'] == '★ 공고문.hwpx'
    item = g.list_bids()['items'][0]
    credit = [f for f in item['notice']['flags'] if f['key'] == 'credit'][0]
    assert credit['ok'] is False                                  # B+ < BB- 이상
    assert [s['title'] for s in item['notice']['sections']] == ['입찰참가자격', '현장설명']


def test_read_notice_failures_are_kept_with_reason(db, monkeypatch):
    _seed(monkeypatch, {'ntceSpecFileNm1': '시방서.hwp'})
    assert n.read_notice('N1')['status'] == 'none'
    with g._conn() as conn:
        conn.execute("""UPDATE g2b_bids SET files='[{"name": "공고문.pdf", "url": "https://www.g2b.go.kr/x1"}]'""")

    def boom(url):
        raise RuntimeError('파일을 받지 못했어요(응답 500).')
    monkeypatch.setattr(n, '_download', boom)
    v = n.read_notice('N1')
    assert v['status'] == 'fail' and '500' in v['message']
    with pytest.raises(LookupError):
        n.read_notice('없음')


def test_run_pending_reads_eligible_once(db, monkeypatch):
    _seed(monkeypatch)
    calls = []
    monkeypatch.setattr(n, 'read_notice', lambda no: calls.append(no) or n._save(no, 'ok'))
    assert n.run_pending() == 1 and calls == ['N1']
    assert n.run_pending() == 0                                   # 읽은 건 다시 안 읽음
    assert g.list_bids()['items'][0]['notice']['status'] == 'ok'

