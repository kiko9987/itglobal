"""폴더 등록 후 다시 저장(retry_pending_license)도 등록증 OCR 상호로 사업자명을 채우는지 (2026-11-09 R4251-TH)."""
import json

import dashboard.services.business_license_handler as blh


class _FakeRedis:
    def __init__(self, data):
        self.data = dict(data)

    def get(self, k):
        return self.data.get(k)

    def delete(self, k):
        self.data.pop(k, None)


def _setup(monkeypatch, ocr_name='로보온(주)', save_ok=True):
    pend = {'channel': 'C1', 'thread_ts': '1.0',
            'files': [{'name': 'image.png', 'url': 'https://x/image.png', 'mime': 'image/png'}]}
    fake = _FakeRedis({'license_pending:R1-X': json.dumps(pend)})

    class _RC:
        redis = fake
    import dashboard.utils.redis_client as rcmod
    monkeypatch.setattr(rcmod, 'get_redis_client', lambda: _RC())
    monkeypatch.setenv('SLACK_PROJECT_BOT_TOKEN', '')  # 댓글은 안 보냄
    import dashboard.utils.google_drive as gd
    monkeypatch.setattr(gd, 'resolve_folder_id', lambda v: {'value': v, 'reason': 'ok_folder'})
    monkeypatch.setattr(blh, '_download_slack_file', lambda url, tok: b'img')
    import dashboard.services.business_license_ocr as ocr
    monkeypatch.setattr(ocr, 'analyze_business_license', lambda b: {'is_license': True, 'name': ocr_name})
    saves = []

    def _save(code, content, name, mime, parent_override=None, ocr_biz=''):
        saves.append((code, parent_override, ocr_biz))
        return {'ok': save_ok, 'file_name': 'R1-X 사업자등록증.png', 'reason': 'x'}
    monkeypatch.setattr(blh, 'save_business_license', _save)
    import dashboard.services.project_slack_notifier as psn
    monkeypatch.setattr(psn, 'refresh_project_card_license', lambda *a, **k: None)
    monkeypatch.setattr(blh, 'split_payer_label', lambda code, nm: '')
    applied = []

    def _apply(code, nm, ch, ts):
        applied.append((code, nm, ch, ts))
        return 'saved', ''
    monkeypatch.setattr(blh, '_apply_ocr_business_name', _apply)
    return fake, saves, applied


def test_retry_fills_business_name(monkeypatch):
    fake, saves, applied = _setup(monkeypatch)
    r = blh.retry_pending_license('R1-X', folder_hint='1wTtyQQ-gZvVm04yCuE8pkSIL3r7CC-v_')
    assert r['saved'] == ['R1-X 사업자등록증.png']
    assert saves == [('R1-X', '1wTtyQQ-gZvVm04yCuE8pkSIL3r7CC-v_', '로보온(주)')]
    assert applied == [('R1-X', '로보온(주)', 'C1', '1.0')]
    assert r['biz_update_status'] == 'saved'
    assert 'license_pending:R1-X' not in fake.data


def test_retry_no_ocr_name_skips_business_name(monkeypatch):
    fake, saves, applied = _setup(monkeypatch, ocr_name='')
    r = blh.retry_pending_license('R1-X', folder_hint='1wTtyQQ-gZvVm04yCuE8pkSIL3r7CC-v_')
    assert r['saved'] and applied == []


def test_retry_save_fail_keeps_pending_and_no_business_name(monkeypatch):
    fake, saves, applied = _setup(monkeypatch, save_ok=False)
    r = blh.retry_pending_license('R1-X', folder_hint='1wTtyQQ-gZvVm04yCuE8pkSIL3r7CC-v_')
    assert r['saved'] == [] and applied == []
    assert 'license_pending:R1-X' in fake.data
