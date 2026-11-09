"""수기세금계산서 드라이브 자동 저장(2026-11-09) — 기존 폴더 규칙 'G 26.09.10.png'."""
from datetime import date

from dashboard.services import handwritten_invoice as hw


def test_invoice_date():
    assert hw.invoice_date('11/09', date(2026, 11, 10)) == date(2026, 11, 9)
    assert hw.invoice_date('12/30', date(2027, 1, 3)) == date(2026, 12, 30)   # 1월에 올린 12월 계산서
    assert hw.invoice_date('', date(2026, 11, 10)) == date(2026, 11, 10)
    assert hw.invoice_date('2/30', date(2026, 3, 1)) == date(2026, 3, 1)      # 없는 날짜


def test_names_follow_folder_rule():
    assert hw.base_name('G', date(2026, 9, 10)) == 'G 26.09.10'
    ex = {'G 26.11.09.jpg'}
    assert hw.unique_name(set(), 'G 26.11.09', '.jpg') == 'G 26.11.09.jpg'
    assert hw.unique_name(ex, 'G 26.11.09', '.jpg', '서울크레인') == 'G 26.11.09 서울크레인.jpg'
    ex.add('G 26.11.09 서울크레인.jpg')
    assert hw.unique_name(ex, 'G 26.11.09', '.jpg', '서울크레인') == 'G 26.11.09 서울크레인 (2).jpg'
    assert hw.unique_name(ex, 'G 26.11.09', '.png') == 'G 26.11.09.png'


class _Req:
    def __init__(self, v): self.v = v
    def execute(self, **k): return self.v


class _Files:
    def __init__(self, d): self.d = d

    def list(self, q='', **k):
        if "mimeType='application/vnd.google-apps.folder'" in q:
            return _Req({'files': [{'id': 'Y25', 'name': 'G 2025년'}, {'id': 'Y26', 'name': '2026년'}]})
        return _Req({'files': [{'name': n} for n in self.d.existing]})

    def create(self, body=None, media_body=None, **k):
        self.d.created.append((body.get('parents'), body['name']))
        return _Req({'id': 'NEW'})


class _Drive:
    def __init__(self):
        self.existing, self.created = ['G 26.11.09.jpg'], []
    def files(self): return _Files(self)


def test_save(monkeypatch):
    import dashboard.services.business_license_handler as bl
    d = _Drive()
    monkeypatch.setattr(bl, '_get_drive', lambda: d)
    files = [{'name': 'IMG_1.jpeg', 'content': b'1'}, {'name': 'memo.txt', 'content': b'x'},
             {'name': 'scan.pdf', 'content': b'2'}]
    got = hw.save('G', date(2026, 11, 9), '서울크레인', files)
    # 같은 날 파일이 이미 있으면 거래처 이름을 붙이고, 사진·PDF 만, 2026년 폴더
    assert got == ['G 26.11.09 서울크레인.jpg', 'G 26.11.09.pdf']
    assert all(p == ['Y26'] for p, _ in d.created)
    assert hw.save('N', date(2026, 11, 9), 'x', files) == []                     # N 은 폴더 없음


def test_save_handwritten_on_transfer(monkeypatch):
    from dashboard.blueprints import slack_eapproval as se
    row = {'id': 1, 'grn': 'R', 'partner': '용달화물', 'evidence': '수기계산서', 'created_at': '2026-11-09 10:00:00',
           'extra': '{"files": [{"name": "a.jpg", "url": "https://f/a"}]}', 'card_channel': 'C', 'card_ts': '1.0'}
    logs, got = [], {}
    monkeypatch.setattr(se, 'invoice_post_files', lambda ch, ts: [])
    monkeypatch.setattr(se, '_download', lambda url: b'IMG')
    monkeypatch.setattr(se, '_thread_log', lambda r, t, broadcast=False: logs.append(t))
    monkeypatch.setattr(hw, 'save', lambda grn, d, p, files: got.update(grn=grn, d=d, n=len(files)) or ['R 26.11.09.jpg'])
    assert se.save_handwritten(row) == ['R 26.11.09.jpg']
    assert got == {'grn': 'R', 'd': date(2026, 11, 9), 'n': 1}
    assert '2026년 › R 26.11.09.jpg' in logs[0] and '엑셀 장부' in logs[0]
