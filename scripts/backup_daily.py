"""매일 새벽 자동 백업.

- instance/users.db (SQLite, 정식 백업 API 스냅샷) → backup/users_db/{YYYYMMDD}.db
- Redis 전체 데이터 dump → backup/redis/{YYYYMMDD}.rdb (또는 SAVE 후 rdb 복사)
- 시크릿 파일 (.env, credentials.json, token.json) → backup/secrets/{YYYYMMDD}/*
- 30일 이상된 로그 파일 (dashboard/logs/*.log*) 자동 삭제
- 최근 30일치 백업 유지, 그 이전은 자동 삭제
- 그날 백업을 암호로 잠가 구글 공유 드라이브(경영지원관리Ⅱ)에 복사 — offsite_backup.py

APScheduler 로 매일 새벽 03:15 실행 (Flask 부팅 시 등록).
수동 실행: python scripts/backup_daily.py
"""
import os
import shutil
import sys
from datetime import datetime, timedelta
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
BACKUP_ROOT = PROJECT_ROOT / 'backup'
USERS_DB_SRC = PROJECT_ROOT / 'instance' / 'users.db'
# Redis: Docker 컨테이너(redis:alpine) 안의 /data/dump.rdb 를 docker cp 로 가져옴.
# 컨테이너 이름 기본값 'redis' (docker-compose 서비스명).
REDIS_CONTAINER = os.getenv('REDIS_BACKUP_CONTAINER', 'redis')
# 실제 파일 시스템에 Redis dump.rdb가 마운트돼 있으면 그 경로도 시도 (레거시)
REDIS_RDB_FALLBACK = Path('C:/Program Files/Redis/dump.rdb')
SECRETS_FILES = [
    PROJECT_ROOT / '.env',
    PROJECT_ROOT / 'credentials.json',
    PROJECT_ROOT / 'instance' / 'google_calendar_token.json',
]
LOG_DIR = PROJECT_ROOT / 'dashboard' / 'logs'
RETENTION_DAYS = 30
LOG_RETENTION_DAYS = 30


def redis_cli_cmd(*args: str) -> list:
    """컨테이너 안 redis-cli 명령 — 비밀번호(REDIS_PASSWORD)가 있으면 REDISCLI_AUTH 로 전달 (2026-10-29 Redis 비밀번호 적용).

    비밀번호 없이 BGSAVE 하면 'NOAUTH' 를 돌려주고도 종료 코드가 0 이라, 옛 dump.rdb 를 복사하고 OK 로 남는다."""
    pw = (os.getenv('REDIS_PASSWORD') or '').strip()
    auth = ['-e', f'REDISCLI_AUTH={pw}'] if pw else []
    return ['docker', 'exec', *auth, REDIS_CONTAINER, 'redis-cli', *args]


def bgsave_started(out: str) -> bool:
    """BGSAVE 응답이 실제로 저장을 시작했는지 (진행 중인 저장이 있어도 OK)."""
    o = (out or '').lower()
    return 'background saving started' in o or 'already in progress' in o or 'scheduled' in o


def _ensure_dir(p: Path) -> None:
    p.mkdir(parents=True, exist_ok=True)


def _timestamp() -> str:
    return datetime.now().strftime('%Y%m%d')


def _prune(dir_path: Path, keep_days: int) -> int:
    """RETENTION_DAYS 이전 파일 삭제. 삭제 개수 반환."""
    if not dir_path.exists():
        return 0
    cutoff = datetime.now() - timedelta(days=keep_days)
    removed = 0
    for f in dir_path.iterdir():
        if not f.is_file():
            continue
        try:
            mtime = datetime.fromtimestamp(f.stat().st_mtime)
            if mtime < cutoff:
                f.unlink()
                removed += 1
        except Exception:
            pass
    return removed


def sqlite_snapshot(src: Path, dest: Path) -> None:
    """SQLite 정식 백업(online backup API)으로 한 시점 스냅샷 → dest (2026-10-30 복원 연습).

    users.db 는 WAL 방식이라 최근 변경이 users.db-wal 에 머물 수 있는데, 예전 shutil.copy2 는 본 파일만
    복사해 그 변경이 빠지거나 쓰기 도중이면 어긋난 파일이 될 수 있었다. 결과는 WAL 없이 파일 하나로 열리게
    journal_mode=DELETE 로 바꾸고 무결성 확인 뒤에 제자리로 옮김(실패하면 옛 백업 그대로)."""
    import sqlite3
    tmp = dest.with_name(dest.name + '.tmp')
    tmp.unlink(missing_ok=True)
    s = sqlite3.connect(str(src), timeout=30)
    d = sqlite3.connect(str(tmp))
    try:
        s.backup(d)
        d.execute('PRAGMA journal_mode=DELETE')
        ok = d.execute('PRAGMA integrity_check').fetchone()[0]
    finally:
        d.close()
        s.close()
    if ok != 'ok':
        tmp.unlink(missing_ok=True)
        raise RuntimeError(f'백업본 무결성 실패: {ok}')
    os.replace(tmp, dest)


def backup_users_db() -> str:
    if not USERS_DB_SRC.exists():
        return f'skip: {USERS_DB_SRC} 없음'
    dest_dir = BACKUP_ROOT / 'users_db'
    _ensure_dir(dest_dir)
    dest = dest_dir / f'{_timestamp()}.db'
    try:
        sqlite_snapshot(USERS_DB_SRC, dest)
    except Exception as exc:
        return f'실패: {type(exc).__name__}: {exc}'
    pruned = _prune(dest_dir, RETENTION_DAYS)
    return f'OK: {dest.name} ({dest.stat().st_size // 1024} KB), 정리 {pruned}건'


def backup_redis() -> str:
    """Redis BGSAVE 트리거 후 dump.rdb 복사.

    우선순위:
    1) Docker 컨테이너(REDIS_CONTAINER)에서 `redis-cli BGSAVE` + `docker cp` 로 가져오기
    2) 로컬 파일 시스템에 dump.rdb 있으면 (Windows 네이티브 Redis) 그 경로 복사
    """
    import subprocess
    import time as _time
    dest_dir = BACKUP_ROOT / 'redis'
    _ensure_dir(dest_dir)
    dest = dest_dir / f'{_timestamp()}.rdb'

    # 1) Docker 시도
    try:
        exists = subprocess.run(
            ['docker', 'ps', '--filter', f'name={REDIS_CONTAINER}', '--format', '{{.Names}}'],
            capture_output=True, text=True, timeout=10,
        )
        if REDIS_CONTAINER in (exists.stdout or ''):
            bg = subprocess.run(
                redis_cli_cmd('BGSAVE'),
                capture_output=True, text=True, timeout=15,
            )
            if bg.returncode == 0 and bgsave_started(bg.stdout):
                _time.sleep(3)  # BGSAVE 완료 대기 (dump.rdb 크기 작으면 즉시)
                cp = subprocess.run(
                    ['docker', 'cp', f'{REDIS_CONTAINER}:/data/dump.rdb', str(dest)],
                    capture_output=True, text=True, timeout=60,
                )
                if cp.returncode == 0 and dest.exists():
                    pruned = _prune(dest_dir, RETENTION_DAYS)
                    return f'OK: {dest.name} ({dest.stat().st_size // 1024} KB, docker), 정리 {pruned}건'
                return f'docker cp 실패: {cp.stderr[:200]}'
            return f'docker BGSAVE 실패: {(bg.stderr or bg.stdout or "")[:200]}'
    except FileNotFoundError:
        pass  # docker 명령 없음 → fallback
    except subprocess.TimeoutExpired:
        return 'docker 명령 timeout'
    except Exception as exc:
        return f'docker 시도 예외: {exc}'

    # 2) Fallback: 로컬 파일
    if REDIS_RDB_FALLBACK.exists():
        try:
            shutil.copy2(REDIS_RDB_FALLBACK, dest)
            pruned = _prune(dest_dir, RETENTION_DAYS)
            return f'OK: {dest.name} ({dest.stat().st_size // 1024} KB, local), 정리 {pruned}건'
        except PermissionError:
            return f'권한 오류: {REDIS_RDB_FALLBACK} 읽기 실패'

    return f'skip: Docker 컨테이너 없음 + 로컬 dump.rdb 없음'


def backup_secrets() -> str:
    """시크릿 파일 3개 (.env, credentials.json, token.json) 백업."""
    dest_dir = BACKUP_ROOT / 'secrets' / _timestamp()
    _ensure_dir(dest_dir)
    copied = []
    missing = []
    for src in SECRETS_FILES:
        if not src.exists():
            missing.append(src.name)
            continue
        try:
            shutil.copy2(src, dest_dir / src.name)
            copied.append(src.name)
        except Exception as exc:
            return f'복사 실패 ({src.name}): {exc}'
    # 30일 이상된 secrets 스냅샷 폴더 삭제
    parent = BACKUP_ROOT / 'secrets'
    pruned = 0
    if parent.exists():
        cutoff = datetime.now() - timedelta(days=RETENTION_DAYS)
        for d in parent.iterdir():
            if not d.is_dir():
                continue
            try:
                mtime = datetime.fromtimestamp(d.stat().st_mtime)
                if mtime < cutoff:
                    shutil.rmtree(d, ignore_errors=True)
                    pruned += 1
            except Exception:
                pass
    parts = [f'복사 {len(copied)}건 ({", ".join(copied)})']
    if missing:
        parts.append(f'없음: {", ".join(missing)}')
    parts.append(f'정리 {pruned}개 폴더')
    return 'OK: ' + ', '.join(parts)


def cleanup_old_logs() -> str:
    """dashboard/logs/ 에서 30일 이상된 로그 파일 삭제.

    - .log, .log.1 ~ .log.N (rotation 백업)
    - service_stdout-YYYYMMDDTHHMMSS.NNN.log 등 stale 서비스 로그
    """
    if not LOG_DIR.exists():
        return f'skip: {LOG_DIR} 없음'
    cutoff = datetime.now() - timedelta(days=LOG_RETENTION_DAYS)
    removed = 0
    total_bytes = 0
    for f in LOG_DIR.iterdir():
        if not f.is_file():
            continue
        name = f.name.lower()
        # .log 로 끝나거나 .log.N (rotation)
        if not (name.endswith('.log') or '.log.' in name):
            continue
        try:
            mtime = datetime.fromtimestamp(f.stat().st_mtime)
            if mtime < cutoff:
                size = f.stat().st_size
                f.unlink()
                removed += 1
                total_bytes += size
        except Exception:
            pass
    mb = total_bytes / (1024 * 1024)
    return f'OK: {removed}개 파일 삭제 ({mb:.1f} MB 확보)'


def backup_offsite() -> str:
    """그날 백업을 잠가 구글 공유 드라이브(경영지원관리Ⅱ)에 올림 — scripts/offsite_backup.py (2026-10-30)."""
    try:
        if str(Path(__file__).resolve().parent) not in sys.path:
            sys.path.insert(0, str(Path(__file__).resolve().parent))
        import offsite_backup
        return offsite_backup.upload(BACKUP_ROOT)
    except Exception as exc:
        return f'실패: {type(exc).__name__}: {str(exc)[:300]}'


def run_all() -> dict:
    print(f'[BACKUP] 시작 @ {datetime.now().isoformat(timespec="seconds")}')
    result = {
        'users_db': backup_users_db(),
        'redis': backup_redis(),
        'secrets': backup_secrets(),
        'log_cleanup': cleanup_old_logs(),
    }
    # 로컬 백업이 하나라도 실패하면 반쪽 묶음을 올리지 않음 (pack 도 파일 없으면 거부)
    result['offsite'] = (backup_offsite() if not failed_items(result)
                         else 'skip: 로컬 백업 실패로 건너뜀')
    for k, v in result.items():
        print(f'  {k}: {v}')
    return result


def failed_items(result: dict) -> list:
    """꼭 남아야 하는 백업(사용자 DB·Redis·비밀값) 중 'OK' 가 아닌 항목 — 하나라도 있으면 실패로 끝냄.

    예전엔 Redis 백업이 'skip'·'docker cp 실패'여도 종료 코드 0 이라 '[BACKUP] 완료'로만 남고 알림이
    없었음 (2026-10-29 점검). 로그 정리(log_cleanup)는 없어도 되는 일이라 제외.
    외부 백업(offsite)은 실제 실패만 — 암호 미설정·로컬 실패로 건너뜀(skip)은 제외 (2026-10-30)."""
    return [f'{k}: {v}' for k, v in result.items()
            if (k in ('users_db', 'redis', 'secrets') and not str(v).startswith('OK'))
            or (k == 'offsite' and str(v).startswith('실패'))]


if __name__ == '__main__':
    try:    # 수동 실행 때도 REDIS_PASSWORD 등 .env 값 사용 (스케줄 실행은 서버 환경을 물려받음)
        from dotenv import load_dotenv
        load_dotenv(PROJECT_ROOT / '.env')
    except ImportError:
        pass
    _bad = failed_items(run_all())
    if _bad:
        print('[BACKUP] 일부 실패:\n  ' + '\n  '.join(_bad), file=sys.stderr)
    sys.exit(1 if _bad else 0)
