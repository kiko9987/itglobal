# -*- coding: utf-8 -*-
"""나눠 내는 공사 부담자 등록증 찾기·계산서 스레드 등록증 첨부 (2026-10-22 G4159).

- 임차인 → 임대인 순으로 올려 임차인 등록증이 백업(_1)으로 밀리면 계산서 요청 때 못 찾던 문제:
  find_payer_license 가 등록증 폴더 전체를 OCR 상호로 대조.
- 계산서 스레드에 올린 등록증이 '발행 완료'로 처리되던 문제: _save_license_from_invoice_thread.
- 계산서 스레드 등록증은 프로젝트 대표 등록증을 바꾸지 않음: save_business_license(keep_canonical).
"""
import sys

sys.path.insert(0, '.')

import pytest

import dashboard.blueprints.slack_bot as sb
import dashboard.services.business_license_handler as blh
import dashboard.services.business_license_ocr as ocr

# 오남프라자 등록증 폴더 (여러 호실 공사가 한 폴더를 같이 씀)
FILES = [
    {'id': 'F59', 'name': 'G4159-JW 사업자등록증.jpg', 'createdTime': '2026-10-21T23:49:20Z'},
    {'id': 'F59_1', 'name': 'G4159-JW 사업자등록증_1.jpg', 'createdTime': '2026-10-21T23:48:25Z'},
    {'id': 'F57', 'name': 'G4157-JW 사업자등록증.jpg', 'createdTime': '2026-10-16T06:05:20Z'},
    {'id': 'F57_1', 'name': 'G4157-JW 사업자등록증_1.pdf', 'createdTime': '2026-10-15T05:48:41Z'},
    {'id': 'F46', 'name': 'G4146-JW 사업자등록증.jpg', 'createdTime': '2026-10-15T04:40:20Z'},
]
OCR = {'F59': '명:이현주', 'F59_1': '다정약국', 'F57': '명:이현주', 'F57_1': '부평센트럴이비인후과의원', 'F46': ''}


@pytest.fixture
def folder(monkeypatch):
    monkeypatch.setattr(blh, '_get_drive', lambda: object())
    monkeypatch.setattr(blh, '_project_folder_id_for', lambda code, fresh=False: 'P')
    monkeypatch.setattr(blh, '_find_license_subfolder', lambda drive, parent: 'S')
    monkeypatch.setattr(blh, '_list_folder_files', lambda drive, fid: [dict(f) for f in FILES])
    monkeypatch.setattr(blh, '_file_biz', lambda fid: OCR.get(fid, ''))
    monkeypatch.setattr(blh, '_download_drive_file', lambda drive, fid: fid.encode())
    monkeypatch.setattr(blh, '_get_license_source', lambda key: None)


def test_payer_license_from_backup(folder):
    # 다정약국 등록증은 _1 로 밀려 있어도 상호로 찾음 (오늘 G4159 경고 원인)
    lic = blh.find_payer_license('G4159-JW', {'name': '김지민(다정약국)'})
    assert lic['content'] == b'F59_1' and lic['file_name'] == 'G4159-JW 사업자등록증_1.jpg'
    assert lic['mimetype'] == 'image/jpeg'


def test_payer_license_same_folder_other_unit(folder):
    # G4160 은 자기 파일이 없음 → 같은 폴더의 이현주 등록증 중 최근 것
    assert blh.find_payer_license('G4160-JW', {'name': '이현주'})['content'] == b'F59'
    # 자기 공사 이름 파일이 있으면 그게 먼저
    assert blh.find_payer_license('G4157-JW', {'name': '이현주'})['content'] == b'F57'
    lic = blh.find_payer_license('G4157-JW', {'name': '부평센트럴이비인후과의원'})
    assert lic['content'] == b'F57_1' and lic['mimetype'] == 'application/pdf'


def test_payer_license_needs_ocr_name(folder):
    # 상호를 못 읽은 파일(F46)은 쓰지 않음, 등록증 없는 부담자는 None
    assert blh.find_payer_license('G4146-JW', {'name': '이희승'}) is None
    assert blh.find_payer_license('G4160-JW', {'name': '은상범(바디홀릭 필라테스)'}) is None
    assert blh.find_payer_license('G4160-JW', {'name': ''}) is None


def test_payer_license_index_fallback(folder, monkeypatch):
    # 폴더에 없으면 상호 색인(다른 공사 폴더) — 이름 키 정확 일치
    idx = {'제이케이디학원': {'code': 'G3999-XX', 'file_id': 'FX', 'ext': 'pdf'}}
    monkeypatch.setattr(blh, '_get_license_source', lambda key: idx.get(key))
    lic = blh.find_payer_license('G4146-JW', {'name': '주식회사 제이케이디학원'})
    assert lic['content'] == b'FX' and lic['file_name'] == 'G3999-XX 사업자등록증.pdf'


def test_split_payer_label(monkeypatch):
    import dashboard.services.project_service as psv
    rec = {'프로젝트 코드': 'G4159-JW', '사업자명': '다정약국',
           '계약금 부담자': '임대인 · 이현주 · kamello@hanmail.net · 1,802,350 · 계산서',
           '잔금 부담자': '임차인 · 김지민(다정약국) · 1,472,350 · 계산서'}
    monkeypatch.setattr(psv, 'get_project_records', lambda: [rec])
    assert blh.split_payer_label('G4159-JW', '명 : 이현주') == '임대인 이현주'
    assert blh.split_payer_label('G4159-JW', '다정약국') == '임차인 김지민(다정약국)'
    assert blh.split_payer_label('G4159-JW', '전혀다른상호') == ''
    assert blh.split_payer_label('R4101-JSH', '다정약국') == ''


# ── save_business_license(keep_canonical) ──
class _Exec:
    def __init__(self, v):
        self.v = v

    def execute(self, **k):
        return self.v


class _FakeDrive:
    def __init__(self):
        self.log = []

    def files(self):
        drive = self

        class _F:
            def update(self, fileId, body, **k):
                drive.log.append(('rename', fileId, body['name']))
                return _Exec({'id': fileId})

            def create(self, body, media_body, **k):
                drive.log.append(('create', body['name']))
                return _Exec({'id': 'NEW', 'name': body['name']})
        return _F()


@pytest.fixture
def save_env(monkeypatch):
    drive = _FakeDrive()
    state = {'existing': [], 'indexed': [], 'biz': {}}
    monkeypatch.setattr(blh, '_validate_magic_bytes', lambda data, ext: True)
    monkeypatch.setattr(blh, '_project_folder_id_for', lambda code, fresh=False: 'P')
    monkeypatch.setattr(blh, '_get_drive', lambda: drive)
    monkeypatch.setattr(blh, '_get_or_create_license_subfolder', lambda d, p: 'S')
    monkeypatch.setattr(blh, '_list_folder_files', lambda d, fid: [dict(f) for f in state['existing']])
    monkeypatch.setattr(blh, 'invalidate_license_state', lambda code: None)
    monkeypatch.setattr(blh, '_set_file_biz', lambda fid, nb: state['biz'].__setitem__(fid, nb))
    monkeypatch.setattr(blh, '_file_biz', lambda fid: state['biz'].get(fid, ''))
    monkeypatch.setattr(blh, '_project_norm_biz', lambda code: '다정약국')
    monkeypatch.setattr(blh, '_index_license_source', lambda nb, code, fid, ext: state['indexed'].append(fid))
    return drive, state


def test_keep_canonical_saves_backup(save_env):
    drive, state = save_env
    state['existing'] = [{'id': 'C', 'name': 'G4159-JW 사업자등록증.jpg', 'md5Checksum': 'aaa'},
                         {'id': 'B1', 'name': 'G4159-JW 사업자등록증_1.jpg', 'md5Checksum': 'bbb'}]
    r = blh.save_business_license('G4159-JW', b'new', 'x.jpg', 'image/jpeg', ocr_biz='다정약국',
                                  keep_canonical=True)
    assert r['ok'] and r['backup'] and r['file_name'] == 'G4159-JW 사업자등록증_2.jpg'
    assert drive.log == [('create', 'G4159-JW 사업자등록증_2.jpg')]      # 대표 등록증은 그대로
    assert state['indexed'] == ['NEW']                                   # 상호를 읽었으니 색인


def test_keep_canonical_backup_without_ocr_not_indexed(save_env):
    drive, state = save_env
    state['existing'] = [{'id': 'C', 'name': 'G4159-JW 사업자등록증.jpg', 'md5Checksum': 'aaa'}]
    r = blh.save_business_license('G4159-JW', b'new', 'x.jpg', 'image/jpeg', keep_canonical=True)
    assert r['ok'] and r['backup'] and state['indexed'] == []


def test_keep_canonical_without_canonical_becomes_canonical(save_env):
    drive, state = save_env
    r = blh.save_business_license('G4159-JW', b'new', 'x.jpg', 'image/jpeg', keep_canonical=True)
    assert r['ok'] and not r['backup'] and r['file_name'] == 'G4159-JW 사업자등록증.jpg'


def test_default_save_still_pushes_canonical(save_env):
    drive, state = save_env
    state['existing'] = [{'id': 'C', 'name': 'G4159-JW 사업자등록증.jpg', 'md5Checksum': 'aaa'}]
    r = blh.save_business_license('G4159-JW', b'new', 'x.jpg', 'image/jpeg')
    assert r['file_name'] == 'G4159-JW 사업자등록증.jpg'
    assert drive.log == [('rename', 'C', 'G4159-JW 사업자등록증_1.jpg'), ('create', 'G4159-JW 사업자등록증.jpg')]


def test_keep_canonical_dedup_any_file(save_env):
    import hashlib
    drive, state = save_env
    md5 = hashlib.md5(b'same').hexdigest()
    state['existing'] = [{'id': 'C', 'name': 'G4159-JW 사업자등록증.jpg', 'md5Checksum': 'aaa'},
                         {'id': 'O', 'name': 'G4157-JW 사업자등록증.jpg', 'md5Checksum': md5}]
    r = blh.save_business_license('G4159-JW', b'same', 'x.jpg', 'image/jpeg', keep_canonical=True)
    assert r['dedup'] and r['file_id'] == 'O' and drive.log == [] and state['indexed'] == []


# ── 계산서 스레드 첨부 ──
class _Client:
    token = 'xoxb-test'

    def __init__(self):
        self.posts = []

    def chat_postMessage(self, **k):
        self.posts.append(k)
        return {'ok': True}


def _file(name='image.png', mt='image/png'):
    return {'name': name, 'mimetype': mt, 'url_private_download': f'https://files/{name}'}


@pytest.fixture
def inv_env(monkeypatch):
    saved = []
    results = {}
    monkeypatch.setattr(blh, '_download_slack_file', lambda url, tok: url.encode())
    monkeypatch.setattr(ocr, 'analyze_business_license', lambda content: results[content.decode()])
    monkeypatch.setattr(blh, 'save_business_license',
                        lambda code, content, name, mt, ocr_biz='', keep_canonical=False, **k:
                        saved.append((code, name, ocr_biz, keep_canonical)) or
                        {'ok': True, 'reason': 'saved', 'file_name': 'G4159-JW 사업자등록증_2.png', 'file_id': 'N'})
    monkeypatch.setattr(blh, 'split_payer_label', lambda code, name: '임차인 김지민(다정약국)' if name == '다정약국' else '')
    return saved, results


LIC = {'is_license': True, 'invoice_like': False, 'name': '다정약국'}
INV = {'is_license': False, 'invoice_like': True, 'name': '(주)아이티글로벌'}
META = {'code': 'G4159-JW', 'biz': '김지민(다정약국)'}


def test_invoice_thread_license_saved_not_completed(inv_env):
    saved, results = inv_env
    results['https://files/lic.png'] = LIC
    cli = _Client()
    ev = {'files': [_file('lic.png')]}
    assert sb._save_license_from_invoice_thread(cli, 'C1', '1.2', ev, META) is True
    assert saved == [('G4159-JW', 'lic.png', '다정약국', True)]
    text = cli.posts[0]['text']
    assert cli.posts[0]['thread_ts'] == '1.2'
    assert '사업자등록증으로 확인해 `G4159-JW` 등록증 폴더에 저장했습니다 (상호: 다정약국)' in text
    assert '*임차인 김지민(다정약국)* 부담자의 계산서를 요청하면 다음부터 자동으로 붙습니다' in text
    assert '발행한 계산서를 이 스레드에 올리면' in text


@pytest.mark.parametrize('kinds', [
    [INV],                                                        # 발행한 계산서
    [LIC, INV],                                                   # 섞임 → 기존 동작
    [{'is_license': True, 'invoice_like': True, 'name': ''}],     # 홈택스 화면에 '사업자등록증' 메뉴 글자
])
def test_invoice_thread_non_license_keeps_old_flow(inv_env, kinds):
    saved, results = inv_env
    files = []
    for i, an in enumerate(kinds):
        results[f'https://files/f{i}.png'] = an
        files.append(_file(f'f{i}.png'))
    cli = _Client()
    assert sb._save_license_from_invoice_thread(cli, 'C1', '1.2', {'files': files}, META) is False
    assert saved == [] and cli.posts == []


def test_invoice_thread_ocr_failure_keeps_old_flow(inv_env, monkeypatch):
    saved, _ = inv_env

    def _boom(content):
        raise RuntimeError('vision down')
    monkeypatch.setattr(ocr, 'analyze_business_license', _boom)
    cli = _Client()
    assert sb._save_license_from_invoice_thread(cli, 'C1', '1.2', {'files': [_file()]}, META) is False
    assert saved == [] and cli.posts == []


def test_invoice_thread_ignores_non_image(inv_env):
    cli = _Client()
    ev = {'files': [_file('a.xlsx', 'application/vnd.ms-excel')]}
    assert sb._save_license_from_invoice_thread(cli, 'C1', '1.2', ev, META) is False


def test_invoice_thread_license_save_fail_still_not_completed(inv_env, monkeypatch):
    saved, results = inv_env
    results['https://files/lic.png'] = LIC
    monkeypatch.setattr(blh, 'save_business_license',
                        lambda *a, **k: {'ok': False, 'reason': 'no_project_folder'})
    cli = _Client()
    assert sb._save_license_from_invoice_thread(cli, 'C1', '1.2', {'files': [_file('lic.png')]}, META) is True
    text = cli.posts[0]['text']
    assert '견적서·계약서 폴더' in text and '자동으로 붙습니다' not in text
