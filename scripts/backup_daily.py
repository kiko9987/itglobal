"""매일 새벽 자동 백업.

- instance/users.db (SQLite) → backup/users_db/{YYYYMMDD}.db
- Redis 전체 데이터 dump → backup/redis/{YYYYMMDD}.rdb (또는 SAVE 후 rdb 복사)
- 최근 30일치 유지, 그 이전은 자동 삭제

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
REDIS_RDB_SRC = Path('C:/Program Files/Redis/dump.rdb')  # 표준 Windows Redis 설치 경로
RETENTION_DAYS = 30


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
    """Redis BGSAVE 트리거 후 dump.rdb 복사."""
    try:
        import redis
        rc = redis.Redis(host='localhost', port=6379)
        rc.bgsave()
    except Exception as exc:
        return f'BGSAVE 실패: {exc}'
    # BGSAVE 는 백그라운드라 완료 대기 (최대 30초)
    import time
    for _ in range(30):
        try:
            last_save = rc.lastsave()
            time.sleep(1)
            if rc.lastsave() != last_save:
                break
        except Exception:
            break
    if not REDIS_RDB_SRC.exists():
        return f'skip: {REDIS_RDB_SRC} 없음 (Redis 경로 확인 필요)'
    dest_dir = BACKUP_ROOT / 'redis'
    _ensure_dir(dest_dir)
    dest = dest_dir / f'{_timestamp()}.rdb'
    try:
        shutil.copy2(REDIS_RDB_SRC, dest)
    except PermissionError:
        return f'권한 오류: {REDIS_RDB_SRC} 읽기 실패 (Redis 서비스 권한 확인)'
    pruned = _prune(dest_dir, RETENTION_DAYS)
    return f'OK: {dest.name} ({dest.stat().st_size // 1024} KB), 정리 {pruned}건'


def run_all() -> dict:
    print(f'[BACKUP] 시작 @ {datetime.now().isoformat(timespec="seconds")}')
    result = {
        'users_db': backup_users_db(),
        'redis': backup_redis(),
    }
    for k, v in result.items():
        print(f'  {k}: {v}')
    return result


if __name__ == '__main__':
    sys.exit(0 if run_all() else 1)
