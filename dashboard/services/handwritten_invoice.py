"""수기세금계산서 드라이브 자동 저장 (2026-11-09).

자금집행 증빙이 「수기계산서」인 결재가 [이체 완료]되면 계산서 사진을 공유 드라이브
`★아이티글로벌 거래처(자재,공구)AS\\★운반비 수기세금계산서\\{연도}년` 에 넣는다(플렉스 안내문 3번 대신).

기존 규칙(경영지원 SB 가 손으로 넣던 것, 2026년 폴더 11장 확인):
  - 연도 폴더: '2026년' (옛 폴더는 'G 2025년' 처럼 앞에 G 가 붙어 있어 '{연도}년' 으로 끝나는 것을 찾는다)
  - 파일 이름: '{G|R} YY.MM.DD.확장자' — 날짜는 계산서 날짜(계산서 글 '11/09 G …' 의 날짜)
  - 같은 날 두 장이면 뒤에 설명을 붙였다('25.07.22 사다리차.jpg') → 거래처 이름, 그래도 겹치면 (2)
  - 사업자(G/R)만 — N 은 계산서가 없어 폴더도 없다
같은 폴더의 '{연도}년 수기세금계산서.xlsx' 장부(사업자번호·상호·대표자·금액)는 경영지원이 직접 쓴다
(손글씨 계산서라 번호를 잘못 읽을 수 있어 자동으로 쓰지 않음).
"""
from __future__ import annotations

import io
import logging
import os
import re
from datetime import date, datetime
from typing import List, Optional

logger = logging.getLogger(__name__)

_DEFAULT_ROOT = '1yAHz1fFqwuabJzudXzI0UGO8z9TKW1SO'   # ★운반비 수기세금계산서
_EXT_OK = ('.jpg', '.jpeg', '.png', '.pdf', '.heic', '.webp', '.gif')
_MIME = {'.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.pdf': 'application/pdf',
         '.heic': 'image/heic', '.webp': 'image/webp', '.gif': 'image/gif'}
EVIDENCE = '수기계산서'


def root_folder_id() -> str:
    return os.getenv('HANDWRITTEN_INVOICE_FOLDER_ID', '').strip() or _DEFAULT_ROOT


def invoice_date(date_md: str, base: date) -> date:
    """'11/09' + 기준일(결재·계산서 글 날짜) → 날짜. 1월에 12월 계산서를 올린 경우는 작년."""
    m = re.match(r'\s*(\d{1,2})/(\d{1,2})', date_md or '')
    if not m:
        return base
    mo, d = int(m.group(1)), int(m.group(2))
    y = base.year - (1 if mo > base.month + 1 else 0)
    try:
        return date(y, mo, d)
    except ValueError:
        return base


def base_name(grn: str, d: date) -> str:
    return f'{grn} {d:%y.%m.%d}'


def unique_name(existing: set, base: str, ext: str, partner: str = '') -> str:
    """'G 26.11.09.jpg' → 겹치면 'G 26.11.09 서울크레인.jpg' → 'G 26.11.09 서울크레인 (2).jpg'."""
    low = {e.lower() for e in existing}
    cand = f'{base}{ext}'
    if cand.lower() not in low:
        return cand
    p = re.sub(r'[\\/:*?"<>|]', '', partner or '').strip()[:20]
    stem = f'{base} {p}' if p else base
    cand = f'{stem}{ext}'
    n = 2
    while cand.lower() in low:
        cand = f'{stem} ({n}){ext}'
        n += 1
    return cand


def _ext(name: str) -> str:
    e = os.path.splitext(name or '')[1].lower()
    return e if e in _EXT_OK else ''


def _year_folder(drive, year: int) -> str:
    """'{연도}년' 으로 끝나는 하위 폴더(옛 'G 2025년' 포함). 없으면 '{연도}년' 을 만든다."""
    root = root_folder_id()
    items = drive.files().list(
        q=f"'{root}' in parents and trashed=false and mimeType='application/vnd.google-apps.folder'",
        fields='files(id,name)', pageSize=100, supportsAllDrives=True, includeItemsFromAllDrives=True,
    ).execute().get('files', [])
    exact = [f for f in items if f['name'].strip() == f'{year}년']
    loose = [f for f in items if f['name'].strip().endswith(f'{year}년')]
    hit = (exact or loose or [None])[0]
    if hit:
        return hit['id']
    logger.info(f'[수기계산서] {year}년 폴더 생성')
    return drive.files().create(body={'name': f'{year}년', 'mimeType': 'application/vnd.google-apps.folder',
                                      'parents': [root]}, fields='id', supportsAllDrives=True).execute()['id']


def save(grn: str, inv_date: date, partner: str, files: List[dict]) -> List[str]:
    """files=[{'name', 'content'}] → 드라이브에 넣은 파일 이름들. 사진·PDF 만, G/R 만."""
    grn = (grn or '').upper()[:1]
    if grn not in ('G', 'R'):
        return []
    files = [f for f in files if _ext(f.get('name', '')) and f.get('content')]
    if not files:
        return []
    from googleapiclient.http import MediaIoBaseUpload
    from dashboard.services.business_license_handler import _get_drive
    drive = _get_drive()
    folder = _year_folder(drive, inv_date.year)
    existing = {f['name'] for f in drive.files().list(
        q=f"'{folder}' in parents and trashed=false", fields='files(name)', pageSize=1000,
        supportsAllDrives=True, includeItemsFromAllDrives=True).execute().get('files', [])}
    saved = []
    for f in files:
        ext = _ext(f['name'])
        ext_out = '.jpg' if ext == '.jpeg' else ext   # 폴더의 기존 파일은 .jpg
        name = unique_name(existing, base_name(grn, inv_date), ext_out, partner)
        media = MediaIoBaseUpload(io.BytesIO(f['content']), mimetype=_MIME.get(ext, 'application/octet-stream'))
        drive.files().create(body={'name': name, 'parents': [folder]}, media_body=media,
                             fields='id', supportsAllDrives=True).execute()
        existing.add(name)
        saved.append(name)
    logger.info(f'[수기계산서] 저장 {inv_date.year}년: {saved}')
    return saved


def base_date_of(row: dict) -> date:
    """기준일 — 계산서 글 시각(origin_ts) → 결재 올린 날."""
    ts = row.get('origin_ts')
    try:
        if ts:
            return datetime.fromtimestamp(float(ts)).date()
    except (TypeError, ValueError):
        pass
    try:
        return datetime.strptime(str(row.get('created_at') or '')[:10], '%Y-%m-%d').date()
    except ValueError:
        return date.today()
