"""외부 백업 — 매일 백업을 암호로 잠가 구글 공유 드라이브(경영지원관리Ⅱ)에 올림 (2026-10-30).

backup/ 은 이 PC 의 C: 에만 있어 PC 고장·랜섬웨어·도난이면 백업도 같이 잃는다(10-29 점검).
03:15 일일 백업(scripts/backup_daily.py)이 끝나면 그날 users.db·Redis dump.rdb·비밀값 파일을 tar.gz 로
묶고 BACKUP_ENCRYPT_PASSWORD 로 잠가(scrypt → AES-256-GCM) 서비스 계정으로 바로 올린다.
서버(Flask)는 LocalSystem 이라 G: 드라이브가 안 보여 Drive API 로 올림 — 서비스 계정은 그 공유 드라이브
콘텐츠 관리자.

- 비밀값(.env 등)이 들어 있어 암호 없이는 열 수 없게 잠근다. 암호는 .env 와 사장님 보관본 두 곳뿐 —
  PC 를 잃으면 .env 도 잃으니 보관본이 없으면 복원 불가.
- 보관: 매일분 90일 + 매달 1일분 400일. 오래된 건 휴지통으로(영구 삭제 아님, 30일 안 되살릴 수 있음).
- 복원: python scripts/offsite_backup.py restore <받은 .itgbak 파일> [풀 폴더]
"""
import io
import json
import os
import re
import tarfile
import time
from datetime import date, datetime, timedelta
from pathlib import Path

MAGIC = b'ITGBAK1\n'
_SALT_LEN, _NONCE_LEN = 16, 12
_SCRYPT = dict(n=2 ** 15, r=8, p=1)
DRIVE_NAME = '경영지원관리Ⅱ'
FOLDER_NAME = 'PM시스템 백업 (삭제 금지)'
DAILY_KEEP_DAYS = 90
MONTHLY_KEEP_DAYS = 400
MIN_KEEP = 7                     # 이보다 적게 남기는 정리는 안 함 (날짜가 틀어진 경우 대비)
STATUS_FILE = 'offsite_last.json'
_NAME_RE = re.compile(r'^itg-backup-(\d{8})\.itgbak$')


def password() -> str:
    return (os.getenv('BACKUP_ENCRYPT_PASSWORD') or '').strip()


def _key(pw: str, salt: bytes) -> bytes:
    from cryptography.hazmat.primitives.kdf.scrypt import Scrypt
    return Scrypt(salt=salt, length=32, **_SCRYPT).derive(pw.encode('utf-8'))


def encrypt(data: bytes, pw: str) -> bytes:
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    if not pw:
        raise ValueError('암호 없음')
    salt, nonce = os.urandom(_SALT_LEN), os.urandom(_NONCE_LEN)
    head = MAGIC + salt + nonce
    return head + AESGCM(_key(pw, salt)).encrypt(nonce, data, head)


def decrypt(blob: bytes, pw: str) -> bytes:
    """암호가 틀리거나 파일이 손상되면 cryptography.exceptions.InvalidTag."""
    from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    if not blob.startswith(MAGIC):
        raise ValueError('ITG 백업 파일이 아님')
    i = len(MAGIC)
    salt, nonce = blob[i:i + _SALT_LEN], blob[i + _SALT_LEN:i + _SALT_LEN + _NONCE_LEN]
    head = blob[:i + _SALT_LEN + _NONCE_LEN]
    return AESGCM(_key(pw, salt)).decrypt(nonce, blob[len(head):], head)


def pack(backup_root, stamp: str) -> bytes:
    """그날(stamp=YYYYMMDD) 백업 파일을 tar.gz 로. 하나라도 없으면 FileNotFoundError — 반쪽 백업은 안 올림."""
    root = Path(backup_root)
    items = [(root / 'users_db' / f'{stamp}.db', 'users.db'),
             (root / 'redis' / f'{stamp}.rdb', 'dump.rdb')]
    sec = root / 'secrets' / stamp
    items += [(p, f'secrets/{p.name}') for p in sorted(sec.iterdir()) if p.is_file()] if sec.is_dir() else []
    missing = [str(p) for p, _ in items if not p.is_file()]
    if missing or not sec.is_dir():
        raise FileNotFoundError(', '.join(missing) or str(sec))
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode='w:gz') as tar:
        for p, arc in items:
            tar.add(str(p), arcname=arc)
    return buf.getvalue()


def file_name(stamp: str) -> str:
    return f'itg-backup-{stamp}.itgbak'


def to_trash(files, today: date) -> list:
    """정리할 파일 [{'id','name'}] — 매일분 90일·매달 1일분 400일 넘은 것. 남는 게 MIN_KEEP 미만이면 정리 안 함."""
    keep, drop = [], []
    for f in files:
        m = _NAME_RE.match(f.get('name') or '')
        if not m:
            continue                 # 사람이 넣은 다른 파일은 손대지 않음
        try:
            d = datetime.strptime(m.group(1), '%Y%m%d').date()
        except ValueError:
            continue
        age = (today - d).days
        (keep if age <= DAILY_KEEP_DAYS or (d.day == 1 and age <= MONTHLY_KEEP_DAYS) else drop).append(f)
    return drop if len(keep) >= MIN_KEEP else []


# backup_daily.py 와 같은 프로세스(서버와 따로)에서 돌아 dashboard 패키지를 import 하지 않는다 —
# 그러면 로깅 설정이 서버와 같은 로그 파일에 핸들러를 하나 더 열어 회전이 꼬인다. 그래서 scripts/ 에 둠.
def _service():
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
    cred = os.getenv('GOOGLE_CREDENTIALS_FILE', 'credentials.json').strip()
    if not os.path.isabs(cred):
        cred = str(Path(__file__).resolve().parents[1] / cred)
    creds = Credentials.from_service_account_file(cred, scopes=['https://www.googleapis.com/auth/drive'])
    return build('drive', 'v3', credentials=creds, cache_discovery=False)


def _retriable(exc) -> bool:
    try:
        status = int(getattr(getattr(exc, 'resp', None), 'status', 0) or 0)
    except (TypeError, ValueError):
        status = 0
    if status in (429, 500, 502, 503, 504):
        return True
    return status == 403 and 'ateLimitExceeded' in str(exc)     # (user)RateLimitExceeded


def _call(req, tries: int = 4):
    for i in range(tries):
        try:
            return req.execute()
        except Exception as exc:
            if i == tries - 1 or not (_retriable(exc) or isinstance(exc, (TimeoutError, ConnectionError))):
                raise
            time.sleep(2 ** i + 0.5)


def _folder_id(svc) -> str:
    fid = (os.getenv('OFFSITE_BACKUP_FOLDER_ID') or '').strip()
    if fid:
        return fid
    drives = _call(svc.drives().list(pageSize=100)).get('drives', [])
    drive = next((d for d in drives if d.get('name') == DRIVE_NAME), None)
    if not drive:
        raise RuntimeError(f"서비스 계정이 공유 드라이브 '{DRIVE_NAME}' 를 못 봄 — 구성원에서 빠졌는지 확인")
    q = (f"name = '{FOLDER_NAME}' and '{drive['id']}' in parents and "
         "mimeType = 'application/vnd.google-apps.folder' and trashed = false")
    found = _call(svc.files().list(q=q, fields='files(id)', corpora='drive', driveId=drive['id'],
                                   supportsAllDrives=True, includeItemsFromAllDrives=True)).get('files', [])
    if found:
        return found[0]['id']
    meta = {'name': FOLDER_NAME, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [drive['id']]}
    return _call(svc.files().create(body=meta, fields='id', supportsAllDrives=True))['id']


def _list(svc, folder_id: str) -> list:
    out, token = [], None
    while True:
        resp = _call(svc.files().list(q=f"'{folder_id}' in parents and trashed = false",
                                      fields='nextPageToken, files(id, name, size)', pageSize=1000,
                                      pageToken=token, supportsAllDrives=True, includeItemsFromAllDrives=True))
        out += resp.get('files', [])
        token = resp.get('nextPageToken')
        if not token:
            return out


def _trash(svc, file_id: str):
    _call(svc.files().update(fileId=file_id, body={'trashed': True}, supportsAllDrives=True))


def upload(backup_root, stamp: str = None, svc=None, today: date = None) -> str:
    """그날 백업을 잠가 올리고 오래된 것 정리. 결과 문구('OK: …' / 'skip: …'), 실패는 예외."""
    from googleapiclient.http import MediaIoBaseUpload
    pw = password()
    if not pw:
        return 'skip: BACKUP_ENCRYPT_PASSWORD 없음'
    today = today or date.today()
    stamp = stamp or today.strftime('%Y%m%d')
    blob = encrypt(pack(backup_root, stamp), pw)
    svc = svc or _service()
    fid = _folder_id(svc)
    name = file_name(stamp)
    before = _list(svc, fid)
    media = MediaIoBaseUpload(io.BytesIO(blob), mimetype='application/octet-stream', resumable=True)
    new = _call(svc.files().create(body={'name': name, 'parents': [fid]}, media_body=media,
                                   fields='id, size', supportsAllDrives=True))
    if int(new.get('size') or 0) != len(blob):
        raise RuntimeError(f"올라간 크기가 다름 ({new.get('size')} ≠ {len(blob)})")
    for f in before:                 # 같은 날 다시 돌린 경우 — 새 파일이 올라간 뒤에 옛 파일을 치움
        if f.get('name') == name:
            _trash(svc, f['id'])
    old = to_trash([f for f in before if f.get('name') != name] + [{'id': new['id'], 'name': name}], today)
    for f in old:
        _trash(svc, f['id'])
    Path(backup_root, STATUS_FILE).write_text(json.dumps(
        {'at': time.time(), 'name': name, 'id': new['id'], 'bytes': len(blob)}), encoding='utf-8')
    return f'OK: {name} ({len(blob) // 1024} KB → {DRIVE_NAME}/{FOLDER_NAME}), 정리 {len(old)}건'


# ── 서류 백업(2026-11-20, 플렉스 해지 전 점검) ─────────────────────────────────────────────
# 인사 서류(instance/hr_docs — 등본·계약서 등, 플렉스에서 옮긴 것 포함)와 플렉스 결재 보관 내보내기(instance/flex_export)는
# users.db 밖 파일이라 매일 묶음에 없었다. 매일 같은 시각에 보되 바뀐 날만 올린다(파일 이름·크기·고친 시각 지문).
DOCS_KEEP = 30                   # 서류 묶음은 최근 30벌(바뀐 날만 생기므로 날짜가 아니라 개수로)
DOCS_STATUS = 'offsite_docs_last.json'
_DOCS_RE = re.compile(r'^itg-docs-(\d{8})\.itgbak$')
VOLUME_BYTES = 300 * 1024 * 1024  # 큰 폴더 한 번 올리기 — 한 조각 300MB(메모리에 잠근 사본까지 두 벌)


def docs_sources(project_root) -> list:
    root = Path(project_root)
    hr = (os.getenv('GW_HR_DOCS_DIR') or '').strip() or str(root / 'instance' / 'hr_docs')
    return [(Path(hr), 'hr_docs'), (root / 'instance' / 'flex_export', 'flex_export'),
            (root / 'instance' / 'cert', 'cert')]      # 증명서 직인 그림(11-20)


def walk(src, arc: str) -> list:
    """폴더 안 파일 [(경로, 묶음 안 이름)] — 이름 순. 폴더가 없으면 빈 목록."""
    src = Path(src)
    if not src.is_dir():
        return []
    return [(p, f"{arc}/{p.relative_to(src).as_posix()}") for p in sorted(src.rglob('*')) if p.is_file()]


def fingerprint(files) -> str:
    import hashlib
    h = hashlib.sha256()
    for p, arc in files:
        st = p.stat()
        h.update(f'{arc}\0{st.st_size}\0{int(st.st_mtime)}\n'.encode('utf-8'))
    return h.hexdigest()


def pack_files(files, gz: bool = True) -> bytes:
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode='w:gz' if gz else 'w') as tar:
        for p, arc in files:
            tar.add(str(p), arcname=arc)
    return buf.getvalue()


def _put(svc, folder_id: str, name: str, blob: bytes) -> str:
    from googleapiclient.http import MediaIoBaseUpload
    media = MediaIoBaseUpload(io.BytesIO(blob), mimetype='application/octet-stream', resumable=True,
                              chunksize=32 * 1024 * 1024)
    new = _call(svc.files().create(body={'name': name, 'parents': [folder_id]}, media_body=media,
                                   fields='id, size', supportsAllDrives=True))
    if int(new.get('size') or 0) != len(blob):
        raise RuntimeError(f"올라간 크기가 다름 ({name}: {new.get('size')} ≠ {len(blob)})")
    return new['id']


def upload_docs(project_root, backup_root, svc=None, today: date = None) -> str:
    """인사 서류·플렉스 보관 내보내기를 잠가 올림 — 지난번과 같으면 건너뜀. 결과 문구, 실패는 예외."""
    pw = password()
    if not pw:
        return 'skip: BACKUP_ENCRYPT_PASSWORD 없음'
    files = [f for src, arc in docs_sources(project_root) for f in walk(src, arc)]
    if not files:
        return 'skip: 서류 없음'
    fp = fingerprint(files)
    status = Path(backup_root, DOCS_STATUS)
    try:
        if json.loads(status.read_text(encoding='utf-8')).get('fp') == fp:
            return f'skip: 바뀐 서류 없음 ({len(files)}개)'
    except Exception:
        pass
    today = today or date.today()
    name = f"itg-docs-{today.strftime('%Y%m%d')}.itgbak"
    blob = encrypt(pack_files(files), pw)
    svc = svc or _service()
    fid = _folder_id(svc)
    before = _list(svc, fid)
    new_id = _put(svc, fid, name, blob)
    for f in before:
        if f.get('name') == name:
            _trash(svc, f['id'])
    ours = sorted([f for f in before if _DOCS_RE.match(f.get('name') or '') and f.get('name') != name],
                  key=lambda f: f['name'], reverse=True)
    old = ours[DOCS_KEEP - 1:]                              # 새 것 + 앞선 29벌만 남김
    for f in old:
        _trash(svc, f['id'])
    Path(backup_root).mkdir(parents=True, exist_ok=True)
    status.write_text(json.dumps({'at': time.time(), 'name': name, 'id': new_id, 'bytes': len(blob),
                                  'files': len(files), 'fp': fp}), encoding='utf-8')
    return f'OK: {name} (서류 {len(files)}개, {len(blob) // 1024} KB), 정리 {len(old)}건'


def volumes(files, limit: int = VOLUME_BYTES) -> list:
    """파일을 순서대로 limit 안쪽 조각으로 나눔. limit 보다 큰 파일은 혼자 한 조각."""
    out, cur, size = [], [], 0
    for p, arc in files:
        s = p.stat().st_size
        if cur and size + s > limit:
            out.append(cur)
            cur, size = [], 0
        cur.append((p, arc))
        size += s
    if cur:
        out.append(cur)
    return out


def upload_archive(src_dir, label: str, svc=None, limit: int = VOLUME_BYTES, log=print) -> list:
    """큰 폴더를 한 번 잠가 올림(플렉스 해지 전 자료 등) — {label}-NNofMM.itgbak 조각들.
    같은 이름·같은 크기 조각이 이미 있으면 건너뜀(끊겼다 다시 돌려도 이어서). 조각마다 tar(압축 없음, 대부분 pdf·zip)."""
    pw = password()
    if not pw:
        raise RuntimeError('BACKUP_ENCRYPT_PASSWORD 없음')
    src = Path(src_dir)
    vols = volumes(walk(src, src.name), limit)
    svc = svc or _service()
    fid = _folder_id(svc)
    have = {f.get('name'): int(f.get('size') or 0) for f in _list(svc, fid)}
    done = []
    for i, vol in enumerate(vols, 1):
        name = f'{label}-{i:02d}of{len(vols):02d}.itgbak'
        blob = encrypt(pack_files(vol, gz=False), pw)
        if have.get(name) == len(blob):
            log(f'  {name} 이미 있음 — 건너뜀')
            done.append((name, len(blob), len(vol), 'skip'))
            continue
        _put(svc, fid, name, blob)
        log(f'  {name} 올림 ({len(vol)}개, {len(blob) // (1024 * 1024)} MB)')
        done.append((name, len(blob), len(vol), 'ok'))
    return done


def last_success_age_hours(backup_root, now: float = None):
    """마지막 외부 백업 성공 후 지난 시간(시간). 기록 없으면 None."""
    try:
        at = json.loads(Path(backup_root, STATUS_FILE).read_text(encoding='utf-8'))['at']
    except Exception:
        return None
    return ((now or time.time()) - float(at)) / 3600


def restore(src: str, out_dir: str, pw: str) -> list:
    """받은 .itgbak 을 풀어 out_dir 에 꺼냄 — 매일분은 users.db·dump.rdb·secrets/, 서류는 hr_docs/·flex_export/,
    큰 폴더 조각은 그 폴더 이름 아래(조각을 모두 같은 곳에 풀면 원래 폴더). 운영 파일은 건드리지 않음."""
    data = decrypt(Path(src).read_bytes(), pw)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    with tarfile.open(fileobj=io.BytesIO(data), mode='r:*') as tar:   # 매일분·서류=gz, 큰 폴더 조각=tar
        tar.extractall(out, filter='data')
        return tar.getnames()


_RESTORE_HELP = """다음 순서로 되돌립니다 (서버를 멈춘 상태에서):
  1) secrets/ 의 .env·credentials.json → 프로젝트 폴더로 (instance 토큰은 instance/ 로)
  2) users.db → instance/users.db
  3) Redis: 컨테이너를 멈추고 dump.rdb 를 /data/dump.rdb 로 docker cp 한 뒤 다시 시작
     (appendonly 를 켜 둔 컨테이너는 appendonlydir 를 비워야 dump.rdb 를 읽음)
  4) 서버 시작 후 /api/health 확인"""


if __name__ == '__main__':
    import getpass
    import sys
    root = Path(__file__).resolve().parents[1]
    try:
        from dotenv import load_dotenv
        load_dotenv(root / '.env')
    except ImportError:
        pass
    args = sys.argv[1:]
    if args[:1] == ['upload']:              # 수동 올리기: upload [YYYYMMDD]
        print(upload(root / 'backup', args[1] if len(args) > 1 else None))
    elif args[:1] == ['upload-docs']:       # 인사 서류·플렉스 보관 내보내기(바뀐 날만)
        print(upload_docs(root, root / 'backup'))
    elif args[:1] == ['upload-archive'] and len(args) >= 3:   # 큰 폴더 한 번: upload-archive <폴더> <이름표>
        res = upload_archive(args[1], args[2])
        print(f"끝: 조각 {len(res)}개, {sum(r[1] for r in res) // (1024 * 1024)} MB, 파일 {sum(r[2] for r in res)}개")
    elif args[:1] == ['restore'] and len(args) >= 2:
        pw = password() or getpass.getpass('백업 암호: ')
        dest = args[2] if len(args) > 2 else str(Path(args[1]).with_suffix(''))
        for n in restore(args[1], dest, pw):
            print('  ', n)
        print(f'풀린 곳: {dest}\n{_RESTORE_HELP}')
    else:
        print('사용법: python scripts/offsite_backup.py upload [YYYYMMDD]\n'
              '        python scripts/offsite_backup.py upload-docs\n'
              '        python scripts/offsite_backup.py upload-archive <폴더> <이름표>\n'
              '        python scripts/offsite_backup.py restore <파일.itgbak> [풀 폴더]')
        sys.exit(2)
