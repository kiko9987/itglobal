# -*- coding: utf-8 -*-
"""폴더 열기(/api/folder/open) — 서버는 아무것도 실행하지 않고 검증된 Drive 주소만 돌려준다.

2026-11-13 보안 점검: 셀 값이 cmd `start`·PowerShell 문자열로 서버에서 실행되던 문제 회귀 방지.
"""
import pathlib
import subprocess
from unittest.mock import patch

import pandas as pd
import pytest
from flask import Flask

from dashboard.blueprints import folders as fo

ID = '1ZsnooHtIPe_4UxKq020Gk8TseLtEMdws'


@pytest.mark.parametrize('value,expected', [
    (ID, ID),
    (f'  {ID}  ', ID),
    (f'https://drive.google.com/drive/folders/{ID}', ID),
    (f'https://drive.google.com/drive/folders/{ID}?usp=sharing', ID),
    (f'https://drive.google.com/open?id={ID}', ID),
    ('drive.google.com&calc', None),                       # 예전 cmd 주입 형태
    (f'drive.google.com/{ID}&calc', None),                 # https:// 로 시작하지 않으면 링크로 안 봄
    (f'https://drive.google.com/drive/folders/{ID}&calc', None),
    ('https://evil.example/drive.google.com/folders/' + ID, None),
    (r'C:\Windows\$(calc)\..', None),                      # 예전 PowerShell 주입 형태
    (r'G:\공유 드라이브\현장\2026', None),
    ('', None),
    ('-', None),
    (None, None),
])
def test_resolve_drive_folder_id(value, expected):
    assert fo._resolve_drive_folder_id(value) == expected


def test_module_has_no_process_launch():
    src = pathlib.Path(fo.__file__).read_text(encoding='utf-8')
    code = '\n'.join(l for l in src.splitlines() if not l.strip().startswith(('#', '"""')) and '되돌리지 말 것' not in l)
    for bad in ('import subprocess', 'subprocess.run', 'os.startfile', 'shell=True', 'powershell'):
        assert bad not in code, bad


@pytest.fixture()
def client():
    app = Flask(__name__)
    app.secret_key = 'test'
    app.register_blueprint(fo.folders_bp)
    with patch('dashboard.auth._validate_session', return_value=None):
        c = app.test_client()
        with c.session_transaction() as s:
            s['user'] = {'email': 'tester@example.invalid', 'permission_level': 'editor'}
        yield c


def _rows(value):
    return pd.DataFrame([{'프로젝트 코드': 'T-1', '견적서 및 계약서 폴더 경로': value}])


@pytest.mark.parametrize('value', [
    'drive.google.com&calc',
    f'https://drive.google.com/drive/folders/{ID}&calc',
    r'C:\Windows\$(calc)\..',
    r'C:\Windows',
])
def test_open_never_launches_process(client, value):
    with patch('dashboard.services.project_service.load_data', return_value=_rows(value)), \
         patch.object(subprocess, 'run') as run, patch.object(subprocess, 'Popen') as popen:
        r = client.get('/api/folder/open/T-1')
    data = r.get_json()
    assert data['success'] is False
    assert data['folder_type'] == 'local_path'
    assert 'url' not in data
    run.assert_not_called()
    popen.assert_not_called()


def test_open_drive_url_returns_canonical_url(client):
    raw = f'https://drive.google.com/drive/folders/{ID}?usp=sharing'
    with patch('dashboard.services.project_service.load_data', return_value=_rows(raw)), \
         patch.object(subprocess, 'run') as run:
        r = client.get('/api/folder/open/T-1')
    data = r.get_json()
    assert data == {
        'success': True,
        'folder_type': 'google_drive',
        'folder_id': ID,
        'url': f'https://drive.google.com/drive/folders/{ID}',
    }
    run.assert_not_called()


def test_open_missing_project_and_empty_path(client):
    with patch('dashboard.services.project_service.load_data', return_value=_rows('-')):
        assert client.get('/api/folder/open/T-1').get_json()['success'] is False
        assert client.get('/api/folder/open/NOPE').get_json()['error'] == '프로젝트를 찾을 수 없습니다'
