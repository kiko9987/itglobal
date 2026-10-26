# -*- coding: utf-8 -*-
"""슬랙 문서 보기(doc_viewer) 회귀 테스트 (2026-10-22).

- 마크다운 → 슬랙 글: 표는 '• 칸 · 칸', 머리글 상자 잔재(':')·이미지 줄 제거.
- 대상 판별: 전용 채널 최상위 첨부·DM 첨부만, 스레드 답글·봇·다른 채널은 제외.
- 처리 흐름: 변환 성공/실패/미지원 형식 안내, 긴 글은 .txt 첨부, 중복 이벤트 1회만.
- 실제 kordoc 변환: tools/kordoc 런타임이 있을 때만(최소 DOCX 를 만들어 변환).
"""
import io
import sys
import zipfile

import pytest

sys.path.insert(0, '.')

from dashboard.services import doc_viewer as dv


ESTIMATE_MD = """:

수신

### 견적명 : 엘지 멀티-V 천장형 냉난방기 탈착 공사 건

| 번 호<br>No. | 품 명<br>Description | 금 액<br>Amount Price | 비 고<br>Remark |
| --- | --- | --- | --- |
| 1 | 멀티-V 실내기 | ₩ 100,000 | 철거 |
| 2 | 멀티-V 공사비 | ₩ 250,000 |  |

| 합계/Total | ₩ 700,000 |
| --- | --- |

![image](images/x/image_001.png)
"""


def test_markdown_to_slack_flattens_tables_and_drops_noise():
    out = dv.markdown_to_slack(ESTIMATE_MD)
    assert '*견적명 : 엘지 멀티-V 천장형 냉난방기 탈착 공사 건*' in out
    assert '*번 호 No. · 품 명 Description · 금 액 Amount Price · 비 고 Remark*' in out
    assert '• 1 · 멀티-V 실내기 · ₩ 100,000 · 철거' in out
    assert '• 2 · 멀티-V 공사비 · ₩ 250,000' in out        # 빈 칸 생략
    assert '• 합계/Total · ₩ 700,000' in out               # 2칸 표는 머리행도 일반 줄
    assert '|' not in out and '---' not in out
    assert 'image' not in out
    assert not any(line.strip() == ':' for line in out.split('\n'))


def test_markdown_to_slack_drops_decorative_separator_lines():
    md = ('제한경쟁 / 적격심사\n\n' + '\\*' * 84 + '\n\n이 입찰 설명서는\n\n' + '-' * 150
          + '\n\n- 다 음 -\n\n═══════\n\n1. 물품구매')
    out = dv.markdown_to_slack(md)
    assert out.split('\n\n') == ['제한경쟁 / 적격심사', '이 입찰 설명서는', '- 다 음 -', '1. 물품구매']


def test_markdown_to_slack_escapes_slack_control_chars():
    out = dv.markdown_to_slack('A <b>굵게</b> & B < C')
    assert out == 'A 굵게 &amp; B &lt; C'


def test_split_for_message_cuts_on_line_boundary():
    text = '\n'.join(f'줄{i:04d}' for i in range(2000))
    head, truncated = dv.split_for_message(text, limit=100)
    assert truncated and len(head) <= 100 and head.endswith(head.split('\n')[-1])
    assert dv.split_for_message('짧음', limit=100) == ('짧음', False)


def _event(**kw):
    base = {'type': 'message', 'subtype': 'file_share', 'channel': 'CDOC', 'channel_type': 'channel',
            'ts': '100.1', 'user': 'U1', 'files': [{'id': 'F1', 'name': '공고.hwp', 'url_private': 'https://x'}]}
    base.update(kw)
    return base


PDF_FILES = [{'id': 'F9', 'name': '견적서.pdf', 'url_private': 'https://x'}]


def test_doc_view_mode(monkeypatch):
    monkeypatch.delenv('SLACK_DOC_VIEW_CHANNEL', raising=False)
    # DM 은 모든 형식
    assert dv.doc_view_mode(_event(channel='DXX', channel_type='im', files=PDF_FILES)) == 'all'
    # 일반 채널·카드 스레드는 HWP 가 있을 때만
    assert dv.doc_view_mode(_event(channel='CLEAD', thread_ts='99.0')) == 'hwp'
    assert dv.doc_view_mode(_event(channel='CLEAD')) == 'hwp'
    assert dv.doc_view_mode(_event(channel='CLEAD', thread_ts='99.0', files=PDF_FILES)) is None
    # 봇·빈 첨부·수정 이벤트는 제외
    assert dv.doc_view_mode(_event(bot_id='B1')) is None
    assert dv.doc_view_mode(_event(files=[])) is None
    assert dv.doc_view_mode(_event(subtype='message_changed')) is None
    # 전용 채널 최상위 글은 모든 형식, 그 채널의 스레드 답글은 일반 규칙
    monkeypatch.setenv('SLACK_DOC_VIEW_CHANNEL', 'CDOC')
    assert dv.doc_view_mode(_event(files=PDF_FILES)) == 'all'
    assert dv.doc_view_mode(_event(thread_ts='99.0', files=PDF_FILES)) is None


def test_start_doc_view_runs_in_background(monkeypatch):
    calls = []
    monkeypatch.setattr(dv, 'handle_doc_view', lambda e, c, t, m: calls.append(m))
    assert dv.start_doc_view(_event(channel='CLEAD', thread_ts='99.0'), None, 'xoxb') == 'hwp'
    assert dv.start_doc_view(_event(channel='CLEAD', files=PDF_FILES), None, 'xoxb') is None
    import time
    for _ in range(50):
        if calls:
            break
        time.sleep(0.01)
    assert calls == ['hwp']


class FakeClient:
    def __init__(self):
        self.posts, self.uploads, self.reactions = [], [], []

    def chat_postMessage(self, **kw):
        self.posts.append(kw)

    def files_upload_v2(self, **kw):
        self.uploads.append(kw)

    def reactions_add(self, **kw):
        self.reactions.append(('+', kw['name']))

    def reactions_remove(self, **kw):
        self.reactions.append(('-', kw['name']))


@pytest.fixture
def no_redis(monkeypatch):
    seen = set()

    def claim(file_id, ts):
        key = (file_id, ts)
        if key in seen:
            return False
        seen.add(key)
        return True
    monkeypatch.setattr(dv, '_claim', claim)
    monkeypatch.setattr(dv, '_download', lambda url, token: b'bytes')


def test_handle_doc_view_posts_converted_text(monkeypatch, no_redis):
    monkeypatch.setattr(dv, 'convert_to_markdown', lambda data, ext: ESTIMATE_MD)
    c = FakeClient()
    dv.handle_doc_view(_event(), c, 'xoxb')
    assert len(c.posts) == 1 and c.posts[0]['thread_ts'] == '100.1'
    assert '공고.hwp' in c.posts[0]['text'] and '멀티-V 실내기' in c.posts[0]['text']
    assert c.uploads == []
    assert ('+', 'white_check_mark') in c.reactions

    # 같은 이벤트가 다시 와도(슬랙 재전송) 두 번 답하지 않는다
    dv.handle_doc_view(_event(), c, 'xoxb')
    assert len(c.posts) == 1


def test_handle_doc_view_medium_text_is_one_message_of_blocks(monkeypatch, no_redis):
    # 슬랙은 4,000자 넘는 text 를 메시지 여러 개로 쪼갠다(과업지시서 6,145자가 답글 2개로 달림)
    md = '\n\n'.join(f'제{i}조 내용입니다 ' * 5 for i in range(80))       # 약 6.5천 자
    monkeypatch.setattr(dv, 'convert_to_markdown', lambda data, ext: md)
    c = FakeClient()
    dv.handle_doc_view(_event(), c, 'xoxb')
    assert len(c.posts) == 1 and c.uploads == []
    blocks = c.posts[0]['blocks']
    assert '공고.hwp' in blocks[0]['text']['text']
    assert all(len(b['text']['text']) <= 3000 for b in blocks)
    joined = '\n\n'.join(b['text']['text'] for b in blocks[1:])
    assert joined == dv.markdown_to_slack(md)                    # 칸을 나눠도 빠지는 글 없음


def test_handle_doc_view_short_text_is_plain_message(monkeypatch, no_redis):
    monkeypatch.setattr(dv, 'convert_to_markdown', lambda data, ext: '짧은 공고 ' * 100)
    c = FakeClient()
    dv.handle_doc_view(_event(), c, 'xoxb')
    assert len(c.posts) == 1 and 'blocks' not in c.posts[0]


def test_handle_doc_view_long_spec_is_whole_in_one_message(monkeypatch, no_redis):
    # L-04373 공사시방서(25,801자)가 앞부분 3천 자 + .txt 로 가 '잘린 것 같다' → 전부 한 메시지에
    md = '\n\n'.join(f'제{i}조 내용입니다 ' * 5 for i in range(330))       # 약 2.7만 자
    monkeypatch.setattr(dv, 'convert_to_markdown', lambda data, ext: md)
    c = FakeClient()
    dv.handle_doc_view(_event(), c, 'xoxb')
    assert len(c.posts) == 1 and c.uploads == []
    blocks = c.posts[0]['blocks']
    assert len(blocks) <= 50
    assert '제329조' in blocks[-1]['text']['text']


def test_handle_doc_view_falls_back_to_file_when_slack_rejects_length(monkeypatch, no_redis):
    md = '\n\n'.join(f'제{i}조 내용입니다 ' * 5 for i in range(330))
    monkeypatch.setattr(dv, 'convert_to_markdown', lambda data, ext: md)

    class Rejecting(FakeClient):
        def chat_postMessage(self, **kw):
            if 'blocks' in kw:
                raise RuntimeError("The server responded with: {'ok': False, 'error': 'msg_blocks_too_long'}")
            super().chat_postMessage(**kw)
    c = Rejecting()
    dv.handle_doc_view(_event(), c, 'xoxb')
    assert len(c.uploads) == 1 and '제329조' in c.uploads[0]['content']
    assert ('+', 'white_check_mark') in c.reactions


def test_handle_doc_view_huge_text_is_one_file_with_clear_preview_end(monkeypatch, no_redis):
    huge_md = '\n\n'.join(f'제{i}조 내용입니다 ' * 5 for i in range(3000))  # 약 19만 자 > 5만 자
    monkeypatch.setattr(dv, 'convert_to_markdown', lambda data, ext: huge_md)
    c = FakeClient()
    dv.handle_doc_view(_event(), c, 'xoxb')
    assert c.posts == [] and len(c.uploads) == 1             # 답글은 파일 하나뿐
    up = c.uploads[0]
    assert up['filename'] == '공고_내용.txt' and up['thread_ts'] == '100.1'
    assert '제2999조' in up['content']
    comment = up['initial_comment']
    assert comment.rstrip().endswith('첨부 파일에 있어요._')     # 말없이 끊기지 않게
    assert '약 19만 자' in comment and len(comment) < 3300


def test_handle_doc_view_reports_failure_in_korean(monkeypatch, no_redis):
    def boom(data, ext):
        raise dv.DocConvertError('지원하지 않는 파일 형식입니다.')
    monkeypatch.setattr(dv, 'convert_to_markdown', boom)
    c = FakeClient()
    dv.handle_doc_view(_event(), c, 'xoxb')
    assert '지원하지 않는 파일 형식입니다.' in c.posts[0]['text']
    assert ('+', 'x') in c.reactions


def test_handle_doc_view_empty_text_means_scan(monkeypatch, no_redis):
    monkeypatch.setattr(dv, 'convert_to_markdown', lambda data, ext: '![image](a.png)\n\n:')
    c = FakeClient()
    dv.handle_doc_view(_event(), c, 'xoxb')
    assert '스캔본' in c.posts[0]['text']


def test_handle_doc_view_unsupported_file_gets_guide(no_redis):
    c = FakeClient()
    dv.handle_doc_view(_event(files=[{'id': 'F2', 'name': '사진.jpg'}]), c, 'xoxb')
    assert 'HWP·HWPX·PDF' in c.posts[0]['text']
    assert c.reactions == []


def test_handle_doc_view_hwp_mode_converts_only_hwp_into_card_thread(monkeypatch, no_redis):
    seen = []
    monkeypatch.setattr(dv, 'convert_to_markdown', lambda data, ext: seen.append(ext) or '공고 본문')
    c = FakeClient()
    files = [{'id': 'F1', 'name': '공고.hwp', 'url_private': 'u'}, PDF_FILES[0],
             {'id': 'F3', 'name': '사진.jpg', 'url_private': 'u'}]
    dv.handle_doc_view(_event(thread_ts='99.0', files=files), c, 'xoxb', 'hwp')
    assert seen == ['hwp']                                  # PDF·사진은 건드리지 않음
    assert len(c.posts) == 1 and c.posts[0]['thread_ts'] == '99.0'   # 카드 스레드에 답글


def test_handle_doc_view_hwp_mode_silent_without_hwp(no_redis):
    c = FakeClient()
    dv.handle_doc_view(_event(files=PDF_FILES), c, 'xoxb', 'hwp')
    assert c.posts == [] and c.reactions == []


def _minimal_docx(text: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w') as z:
        z.writestr('[Content_Types].xml',
                   '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '</Types>')
        z.writestr('_rels/.rels',
                   '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   '</Relationships>')
        z.writestr('word/document.xml',
                   '<?xml version="1.0" encoding="UTF-8"?>'
                   '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                   f'<w:body><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:body></w:document>')
    return buf.getvalue()


@pytest.mark.skipif(not dv._KORDOC_CLI.exists() or not dv._node_path(),
                    reason='tools/kordoc 런타임 미설치')
def test_convert_to_markdown_real_kordoc():
    md = dv.convert_to_markdown(_minimal_docx('과업지시서 현장설명회 일정'), 'docx')
    assert '과업지시서 현장설명회 일정' in md

    with pytest.raises(dv.DocConvertError) as exc:
        dv.convert_to_markdown(b'garbage', 'hwp')
    assert '지원하지 않는' in str(exc.value)
