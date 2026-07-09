"""매일 새벽 자동 백업.

- instance/users.db (SQLite) → backup/users_db/{YYYYMMDD}.db
- Redis 전체 데이터 dump → backup/redis/{YYYYMMDD}.rdb (또는 SAVE 후 rdb 복사)
- 시크릿 파일 (.env, credentials.json, token.json) → backup/secrets/{YYYYMMDD}/*
- 30일 이상된 로그 파일 (dashboard/logs/*.log*) 자동 삭제
- 최근 30일치 백업 유지, 그 이전은 자동 삭제

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


def backup_users_db() -> str:
    if not USERS_DB_SRC.exists():
        return f'skip: {USERS_DB_SRC} 없음'
    dest_dir = BACKUP_ROOT / 'users_db'
    _ensure_dir(dest_dir)
    dest = dest_dir / f'{_timestamp()}.db'
    shutil.copy2(USERS_DB_SRC, dest)
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
                ['docker', 'exec', REDIS_CONTAINER, 'redis-cli', 'BGSAVE'],
                capture_output=True, text=True, timeout=15,
            )
            if bg.returncode == 0:
                _time.sleep(3)  # BGSAVE 완료 대기 (dump.rdb 크기 작으면 즉시)
                cp = subprocess.run(
                    ['docker', 'cp', f'{REDIS_CONTAINER}:/data/dump.rdb', str(dest)],
                    capture_output=True, text=True, timeout=60,
                )
                if cp.returncode == 0 and dest.exists():
                    pruned = _prune(dest_dir, RETENTION_DAYS)
                    return f'OK: {dest.name} ({dest.stat().st_size // 1024} KB, docker), 정리 {pruned}건'
                return f'docker cp 실패: {cp.stderr[:200]}'
            return f'docker BGSAVE 실패: {bg.stderr[:200]}'
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


def run_all() -> dict:
    print(f'[BACKUP] 시작 @ {datetime.now().isoformat(timespec="seconds")}')
    result = {
        'users_db': backup_users_db(),
        'redis': backup_redis(),
        'secrets': backup_secrets(),
        'log_cleanup': cleanup_old_logs(),
    }
    for k, v in result.items():
        print(f'  {k}: {v}')
    return result


if __name__ == '__main__':
    sys.exit(0 if run_all() else 1)
