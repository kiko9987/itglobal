"""셀 메모 읽기 — 노트·시작 위치만 받는지(2026-11-15).

예전엔 fields 없이 includeGridData=True 로 5열×4천 행의 값·서식까지 통째로 받아 3분마다 4.3초,
큰 응답을 푸는 동안 서버 전체가 느려졌다.
"""
from dashboard.utils.google_sheets import GoogleSheetsManager


class _Req:
    def __init__(self, result):
        self.result = result

    def execute(self):
        return self.result


class _Svc:
    def __init__(self, result):
        self.result = result
        self.kwargs = None

    def spreadsheets(self):
        return self

    def get(self, **kw):
        self.kwargs = kw
        return _Req(self.result)


class _Mgr(GoogleSheetsManager):
    service = None   # 테스트에서 가짜 서비스를 넣으려고 클래스 속성으로 덮음

    def __init__(self, svc):   # 실제 인증 안 함
        self.service = svc

    def _execute_with_retry(self, request_factory, operation_name=''):
        return request_factory().execute()


def test_requests_only_notes_and_positions():
    svc = _Svc({'sheets': []})
    _Mgr(svc).get_cell_notes('sid', '공사 현황', ['V', 'W', 'Z'])
    assert svc.kwargs['includeGridData'] is True
    assert svc.kwargs['fields'] == 'sheets(data(startRow,startColumn,rowData(values(note))))'
    assert svc.kwargs['ranges'] == ['공사 현황!V:Z']


def test_parses_masked_response_with_positions():
    # fields 로 줄인 응답: 노트 없는 칸은 빈 {} 로 자리만 남고, startRow 0 은 생략된다
    masked = {'sheets': [{'data': [{'startColumn': 21, 'rowData': [
        {},                                                  # 1행(헤더)
        {'values': [{'note': '입금 3/1'}, {}, {}, {}, {'note': '계산서 메모'}]},   # 2행: V, Z
        {'values': [{}, {'note': '중도금 메모'}]},             # 3행: W
    ]}]}]}
    out = _Mgr(_Svc(masked)).get_cell_notes('sid', '공사 현황', ['V', 'W', 'Z'])
    assert out == {2: {'V': '입금 3/1', 'Z': '계산서 메모'}, 3: {'W': '중도금 메모'}}
