"""
사용자 관리 SQLite 데이터베이스
파일 기반 JSON에서 SQLite로 마이그레이션하여 동시성과 보안 강화
"""

import sqlite3
import hashlib
import secrets
import json
import os
import threading
import shutil
import logging
from datetime import datetime, timedelta
from contextlib import contextmanager
from typing import Optional, Dict, List, Tuple

logger = logging.getLogger(__name__)

# 보안 설정 상수
PBKDF2_ITERATIONS = 100000  # PBKDF2 반복 횟수 (OWASP 권장: 100,000+)
PBKDF2_HASH_NAME = 'sha256'  # 해시 알고리즘
SALT_LENGTH = 16  # Salt 길이 (바이트)


class UserDatabase:
    """SQLite 기반 사용자 관리 시스템"""

    def __init__(self, db_path: str = None):
        if db_path:
            self.db_path = db_path
        else:
            # instance 폴더를 기본 경로로 사용 (운영 환경 안전성)
            project_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
            instance_dir = os.path.join(project_root, 'instance')

            # instance 폴더가 없으면 생성
            if not os.path.exists(instance_dir):
                os.makedirs(instance_dir)
                logger.info(f"[USER_DB] instance 폴더 생성: {instance_dir}")

            self.db_path = os.path.join(instance_dir, 'users.db')

            # 기존 DB 마이그레이션 처리
            self._migrate_existing_db()

        self.lock = threading.RLock()  # 스레드 안전성 보장
        logger.info(f"[USER_DB] SQLite 데이터베이스 경로: {self.db_path}")
        self._init_database()

    def _init_database(self):
        """데이터베이스 초기화 및 테이블 생성"""
        with self._get_connection() as conn:
            conn.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT UNIQUE NOT NULL,
                    name TEXT NOT NULL,
                    password_hash TEXT,
                    permission_level TEXT NOT NULL DEFAULT 'viewer',
                    is_active BOOLEAN NOT NULL DEFAULT 1,
                    is_resigned BOOLEAN NOT NULL DEFAULT 0,
                    auth_method TEXT NOT NULL DEFAULT 'password',
                    picture TEXT DEFAULT '',
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    last_login TIMESTAMP,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            # 기존 테이블에 is_resigned 컬럼 추가 (마이그레이션)
            try:
                conn.execute('ALTER TABLE users ADD COLUMN is_resigned BOOLEAN NOT NULL DEFAULT 0')
                logger.info("[USER_DB] is_resigned 컬럼이 추가되었습니다.")
            except sqlite3.OperationalError:
                # 컬럼이 이미 존재하는 경우 무시
                pass

            # 세션 무효화 타임스탬프 컬럼 추가 (마이그레이션)
            try:
                conn.execute('ALTER TABLE users ADD COLUMN session_invalidate_after TIMESTAMP')
                logger.info("[USER_DB] session_invalidate_after 컬럼이 추가되었습니다.")
            except sqlite3.OperationalError:
                # 컬럼이 이미 존재하는 경우 무시
                pass

            # 인덱스 생성 (성능 향상)
            conn.execute('CREATE INDEX IF NOT EXISTS idx_users_email ON users(email)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_users_active ON users(is_active)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_users_permission ON users(permission_level)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_users_resigned ON users(is_resigned)')

            # 사용자 세션 테이블
            conn.execute('''
                CREATE TABLE IF NOT EXISTS user_sessions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id INTEGER NOT NULL,
                    session_token TEXT UNIQUE NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    expires_at TIMESTAMP NOT NULL,
                    is_active BOOLEAN NOT NULL DEFAULT 1,
                    user_agent TEXT,
                    ip_address TEXT,
                    FOREIGN KEY (user_id) REFERENCES users (id) ON DELETE CASCADE
                )
            ''')

            conn.execute('CREATE INDEX IF NOT EXISTS idx_sessions_token ON user_sessions(session_token)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_sessions_active ON user_sessions(is_active)')

            # 감사 로그 테이블 (통합)
            conn.execute('''
                CREATE TABLE IF NOT EXISTS audit_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_email TEXT,
                    action TEXT NOT NULL,
                    details TEXT,
                    project_code TEXT,
                    field_name TEXT,
                    old_value TEXT,
                    new_value TEXT,
                    ip_address TEXT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_email) REFERENCES users (email) ON DELETE SET NULL
                )
            ''')

            conn.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_user_email ON audit_logs(user_email)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_timestamp ON audit_logs(timestamp)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_project_code ON audit_logs(project_code)')

            # 프로젝트 코드 시퀀스 테이블 (다중 프로세스 환경에서 중복 방지)
            conn.execute('''
                CREATE TABLE IF NOT EXISTS project_code_sequences (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    pattern TEXT UNIQUE NOT NULL,
                    current_number INTEGER NOT NULL DEFAULT 0,
                    last_updated TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(pattern)
                )
            ''')

            conn.execute('CREATE INDEX IF NOT EXISTS idx_project_code_sequences_pattern ON project_code_sequences(pattern)')

            conn.execute('''
                CREATE TABLE IF NOT EXISTS project_calendar_events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_code TEXT UNIQUE NOT NULL,
                    calendar_id TEXT NOT NULL,
                    event_id TEXT NOT NULL,
                    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
                )
            ''')

            conn.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_calendar_events_code ON project_calendar_events(project_code)')

            conn.commit()

    def _migrate_existing_db(self):
        """기존 dashboard/users.db를 instance 폴더로 이동"""
        try:
            # 기존 DB 경로 (이전 위치)
            old_db_path = os.path.join(os.path.dirname(__file__), '..', 'users.db')

            if os.path.exists(old_db_path) and not os.path.exists(self.db_path):
                logger.info(f"[USER_DB] 기존 DB 파일 발견, instance 폴더로 이동 시작...")
                logger.info(f"[USER_DB] 이동: {old_db_path} -> {self.db_path}")

                # 파일 이동
                shutil.move(old_db_path, self.db_path)
                logger.info(f"[USER_DB] 기존 DB 파일을 instance 폴더로 성공적으로 이동했습니다.")

        except Exception as e:
            logger.error(f"[USER_DB] DB 파일 이동 실패: {e}")
            # 실패해도 계속 진행 (새 DB 생성)

    @contextmanager
    def _get_connection(self):
        """스레드 안전한 데이터베이스 연결"""
        with self.lock:
            conn = sqlite3.connect(self.db_path, timeout=30.0)
            conn.row_factory = sqlite3.Row  # 딕셔너리 형태로 결과 반환
            conn.execute('PRAGMA foreign_keys = ON')
            try:
                yield conn
            finally:
                conn.close()

    def _hash_password(self, password: str) -> str:
        """비밀번호 해싱 (기존 방식과 호환)"""
        salt = secrets.token_hex(SALT_LENGTH)
        password_hash = hashlib.pbkdf2_hmac(PBKDF2_HASH_NAME, password.encode('utf-8'), salt.encode('utf-8'), PBKDF2_ITERATIONS)
        return f"{salt}:{password_hash.hex()}"

    def _verify_password(self, stored_password: str, provided_password: str) -> bool:
        """비밀번호 검증 (기존 방식과 호환)"""
        try:
            salt, password_hash = stored_password.split(':')
            provided_hash = hashlib.pbkdf2_hmac(PBKDF2_HASH_NAME, provided_password.encode('utf-8'), salt.encode('utf-8'), PBKDF2_ITERATIONS)
            return password_hash == provided_hash.hex()
        except (ValueError, AttributeError) as e:
            # 잘못된 비밀번호 형식 또는 None 값
            logger.warning(f"비밀번호 검증 실패: {e}")
            return False

    def authenticate_user(self, email: str, password: str) -> Optional[Dict]:
        """사용자 인증 (기존 인터페이스 호환)"""
        with self._get_connection() as conn:
            cursor = conn.execute('''
                SELECT * FROM users
                WHERE email = ? AND is_active = 1
            ''', (email.lower(),))

            user = cursor.fetchone()
            if not user:
                return None

            if not user['password_hash']:  # OAuth 사용자
                return None

            if not self._verify_password(user['password_hash'], password):
                return None

            # 로그인 시간 업데이트
            conn.execute('''
                UPDATE users
                SET last_login = CURRENT_TIMESTAMP, updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
            ''', (user['id'],))
            conn.commit()

            return self._user_to_dict(user)

    def authenticate_google_user(self, google_user_info: Dict) -> Optional[Dict]:
        """구글 OAuth 사용자 인증 및 자동 등록 (기존 인터페이스 호환)"""
        email = google_user_info['email'].lower()
        name = google_user_info['name']
        picture = google_user_info.get('picture', '')

        with self._get_connection() as conn:
            # 기존 사용자 확인
            cursor = conn.execute('SELECT * FROM users WHERE email = ?', (email,))
            user = cursor.fetchone()

            if not user:
                # 첫 번째 OAuth 사용자인지 확인
                cursor = conn.execute("SELECT COUNT(*) as count FROM users WHERE auth_method = 'google_oauth'")
                is_first_oauth_user = cursor.fetchone()['count'] == 0

                permission_level = 'admin' if is_first_oauth_user else 'viewer'

                # 새 사용자 생성
                cursor = conn.execute('''
                    INSERT INTO users (email, name, password_hash, permission_level, is_active,
                                     auth_method, picture, created_at, last_login, updated_at)
                    VALUES (?, ?, '', ?, 1, 'google_oauth', ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                ''', (email, name, permission_level, picture))

                user_id = cursor.lastrowid
                conn.commit()

                # 로그 출력
                if is_first_oauth_user:
                    print(f"[PARTY] 첫 번째 Google 사용자 등록: {name} ({email}) - 관리자 권한 자동 부여!")
                else:
                    print(f"[SUCCESS] 새 Google 사용자 등록: {name} ({email}) - viewer 권한")

                # 새로 생성된 사용자 정보 조회
                cursor = conn.execute('SELECT * FROM users WHERE id = ?', (user_id,))
                user = cursor.fetchone()

            else:
                # 기존 사용자 정보 업데이트
                if not user['is_active']:
                    return None

                conn.execute('''
                    UPDATE users
                    SET name = ?, picture = ?, auth_method = 'google_oauth',
                        last_login = CURRENT_TIMESTAMP, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                ''', (name, picture, user['id']))
                conn.commit()

                # 업데이트된 사용자 정보 다시 조회
                cursor = conn.execute('SELECT * FROM users WHERE id = ?', (user['id'],))
                user = cursor.fetchone()

            return self._user_to_dict(user)

    def get_user_by_email(self, email: str) -> Optional[Dict]:
        """이메일로 사용자 정보 가져오기"""
        with self._get_connection() as conn:
            cursor = conn.execute('SELECT * FROM users WHERE email = ?', (email.lower(),))
            user = cursor.fetchone()
            return self._user_to_dict(user) if user else None

    def get_all_users(self) -> List[Dict]:
        """모든 사용자 정보 가져오기 (관리자용)"""
        with self._get_connection() as conn:
            cursor = conn.execute('''
                SELECT * FROM users
                ORDER BY created_at DESC
            ''')
            users = cursor.fetchall()
            logger.info(f"[USER_DB] get_all_users() 호출: DB에서 {len(users)}명 조회됨")
            result = [self._user_to_dict(user) for user in users]
            logger.info(f"[USER_DB] get_all_users() 반환: {len(result)}명의 딕셔너리 변환 완료")
            for user in result:
                logger.info(f"[USER_DB]   - {user['name']} ({user['email']}) is_resigned={user['is_resigned']}")
            return result

    def get_resigned_managers(self) -> List[Dict]:
        """퇴사 처리된 담당자 정보 목록 반환 (이름, 이메일)"""
        with self._get_connection() as conn:
            cursor = conn.execute('''
                SELECT name, email FROM users
                WHERE is_resigned = 1
                ORDER BY name
            ''')
            return [{'name': row['name'], 'email': row['email']} for row in cursor.fetchall()]

    def create_user(self, name: str, email: str, password: str, permission_level: str = 'viewer') -> Tuple[bool, str]:
        """새 사용자 생성"""
        with self._get_connection() as conn:
            # 중복 확인
            cursor = conn.execute('SELECT COUNT(*) as count FROM users WHERE email = ?', (email.lower(),))
            if cursor.fetchone()['count'] > 0:
                return False, "이미 존재하는 사용자입니다."

            # 사용자 생성
            password_hash = self._hash_password(password)
            conn.execute('''
                INSERT INTO users (email, name, password_hash, permission_level, is_active,
                                 auth_method, created_at, updated_at)
                VALUES (?, ?, ?, ?, 1, 'password', CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
            ''', (email.lower(), name, password_hash, permission_level))
            conn.commit()

            return True, "사용자가 생성되었습니다."

    def update_user_permission(self, email: str, permission_level: str) -> Tuple[bool, str]:
        """사용자 권한 업데이트 (모든 세션 무효화)"""
        with self._get_connection() as conn:
            cursor = conn.execute('''
                UPDATE users
                SET permission_level = ?,
                    updated_at = CURRENT_TIMESTAMP,
                    session_invalidate_after = CURRENT_TIMESTAMP
                WHERE email = ?
            ''', (permission_level, email.lower()))

            if cursor.rowcount == 0:
                return False, "사용자를 찾을 수 없습니다."

            conn.commit()
            return True, "권한이 업데이트되었습니다."

    def toggle_user_status(self, email: str, is_active: bool) -> Tuple[bool, str]:
        """사용자 활성화/비활성화"""
        with self._get_connection() as conn:
            cursor = conn.execute('''
                UPDATE users
                SET is_active = ?, updated_at = CURRENT_TIMESTAMP
                WHERE email = ?
            ''', (is_active, email.lower()))

            if cursor.rowcount == 0:
                return False, "사용자를 찾을 수 없습니다."

            conn.commit()
            status = "활성화" if is_active else "비활성화"
            return True, f"사용자가 {status}되었습니다."

    def update_user_status(self, email: str, status: str) -> Tuple[bool, str]:
        """사용자 상태 업데이트 (활성/비활성/퇴사) - 모든 세션 무효화"""
        status_map = {
            '활성': (True, False),
            'active': (True, False),
            '비활성': (False, False),
            'inactive': (False, False),
            '퇴사': (False, True),
            'resigned': (False, True)
        }

        if status not in status_map:
            return False, f"유효하지 않은 상태입니다: {status}"

        is_active, is_resigned = status_map[status]

        with self._get_connection() as conn:
            cursor = conn.execute('''
                UPDATE users
                SET is_active = ?, is_resigned = ?,
                    updated_at = CURRENT_TIMESTAMP,
                    session_invalidate_after = CURRENT_TIMESTAMP
                WHERE email = ?
            ''', (is_active, is_resigned, email.lower()))

            if cursor.rowcount == 0:
                return False, "사용자를 찾을 수 없습니다."

            conn.commit()
            status_text = {
                (True, False): "활성화",
                (False, False): "비활성화",
                (False, True): "퇴사 처리"
            }.get((is_active, is_resigned), "상태 변경")

            return True, f"사용자가 {status_text}되었습니다."

    def delete_user(self, email: str) -> Tuple[bool, str]:
        """사용자 삭제"""
        with self._get_connection() as conn:
            cursor = conn.execute('DELETE FROM users WHERE email = ?', (email.lower(),))

            if cursor.rowcount == 0:
                return False, "사용자를 찾을 수 없습니다."

            conn.commit()
            return True, "사용자가 삭제되었습니다."

    def _user_to_dict(self, user_row) -> Dict:
        """SQLite Row를 딕셔너리로 변환 (기존 인터페이스 호환)"""
        if not user_row:
            return None

        # sqlite3.Row는 .get() 메서드가 없으므로 try-except 사용
        try:
            is_resigned = bool(user_row['is_resigned'])
        except (KeyError, IndexError):
            is_resigned = False

        return {
            'id': user_row['id'],
            'name': user_row['name'],
            'email': user_row['email'],
            'permission_level': user_row['permission_level'],
            'is_active': bool(user_row['is_active']),
            'is_resigned': is_resigned,
            'auth_method': user_row['auth_method'],
            'picture': user_row['picture'] or '',
            'created_at': user_row['created_at'],
            'last_login': user_row['last_login'],
            'updated_at': user_row['updated_at']
        }

    def migrate_from_json(self, json_file_path: str) -> Tuple[bool, str]:
        """기존 users.json에서 SQLite로 데이터 마이그레이션"""
        try:
            if not os.path.exists(json_file_path):
                return True, "JSON 파일이 없어 마이그레이션을 건너뜁니다."

            with open(json_file_path, 'r', encoding='utf-8') as f:
                json_users = json.load(f)

            migrated_count = 0
            with self._get_connection() as conn:
                for email, user_data in json_users.items():
                    # 이미 존재하는지 확인
                    cursor = conn.execute('SELECT COUNT(*) as count FROM users WHERE email = ?', (email.lower(),))
                    if cursor.fetchone()['count'] > 0:
                        continue

                    # 데이터 마이그레이션
                    conn.execute('''
                        INSERT INTO users (email, name, password_hash, permission_level, is_active,
                                         auth_method, picture, created_at, last_login, updated_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    ''', (
                        email.lower(),
                        user_data.get('name', ''),
                        user_data.get('password_hash', ''),
                        user_data.get('permission_level', 'viewer'),
                        user_data.get('is_active', True),
                        user_data.get('auth_method', 'password'),
                        user_data.get('picture', ''),
                        user_data.get('created_at'),
                        user_data.get('last_login')
                    ))
                    migrated_count += 1

                conn.commit()

            return True, f"{migrated_count}명의 사용자를 성공적으로 마이그레이션했습니다."

        except Exception as e:
            return False, f"마이그레이션 실패: {str(e)}"

    def backup_to_json(self, backup_file_path: str) -> Tuple[bool, str]:
        """SQLite 데이터를 JSON으로 백업"""
        try:
            users = self.get_all_users()
            backup_data = {}

            for user in users:
                # password_hash는 백업하지 않음 (보안)
                user_backup = user.copy()
                if 'password_hash' in user_backup:
                    del user_backup['password_hash']
                backup_data[user['email']] = user_backup

            with open(backup_file_path, 'w', encoding='utf-8') as f:
                json.dump(backup_data, f, ensure_ascii=False, indent=2)

            return True, f"{len(users)}명의 사용자를 백업했습니다."

        except Exception as e:
            return False, f"백업 실패: {str(e)}"

    def get_next_project_number(self, pattern: str, current_max_from_sheet: int = 0) -> int:
        """
        프로젝트 코드 패턴에 대한 다음 시퀀스 번호 반환 (다중 프로세스 안전)

        Args:
            pattern: 프로젝트 코드 패턴 (예: "GLOBAL")
            current_max_from_sheet: 현재 스프레드시트의 최대 프로젝트 번호

        Returns:
            int: 다음 사용 가능한 번호 (1부터 시작)

        Note:
            - SQLite의 트랜잭션 격리를 활용하여 동시성 제어
            - Gunicorn workers=4 환경에서도 안전하게 작동
            - threading.RLock보다 더 강력한 프로세스 간 동기화
            - 매번 스프레드시트 max와 비교하여 시퀀스가 뒤처지지 않도록 자동 조정
            - 수동으로 스프레드시트에 프로젝트가 추가되어도 자동 감지
        """
        with self._get_connection() as conn:
            # BEGIN IMMEDIATE로 즉시 쓰기 락 획득 (다른 프로세스 대기)
            conn.execute('BEGIN IMMEDIATE')

            try:
                # 시퀀스가 있는지 확인
                cursor = conn.execute('''
                    SELECT current_number FROM project_code_sequences
                    WHERE pattern = ?
                ''', (pattern,))

                existing = cursor.fetchone()

                if existing is None:
                    # 시퀀스가 없으면 스프레드시트 max값으로 초기화
                    logger.info(f"[SEQUENCE] 시퀀스 최초 생성: pattern={pattern}, max_from_sheet={current_max_from_sheet}")
                    conn.execute('''
                        INSERT INTO project_code_sequences (pattern, current_number)
                        VALUES (?, ?)
                    ''', (pattern, current_max_from_sheet))
                else:
                    # 시퀀스가 이미 존재 - 스프레드시트 max와 비교하여 필요시 조정
                    current_sequence = existing['current_number']

                    if current_sequence < current_max_from_sheet:
                        logger.warning(
                            f"[SEQUENCE] 시퀀스({current_sequence})가 스프레드시트 max({current_max_from_sheet})보다 작음. "
                            f"자동으로 {current_max_from_sheet}로 조정합니다."
                        )
                        conn.execute('''
                            UPDATE project_code_sequences
                            SET current_number = ?, last_updated = CURRENT_TIMESTAMP
                            WHERE pattern = ?
                        ''', (current_max_from_sheet, pattern))

                # 번호 증가 (ATOMIC)
                conn.execute('''
                    UPDATE project_code_sequences
                    SET current_number = current_number + 1,
                        last_updated = CURRENT_TIMESTAMP
                    WHERE pattern = ?
                ''', (pattern,))

                # 증가된 번호 조회
                cursor = conn.execute('''
                    SELECT current_number FROM project_code_sequences
                    WHERE pattern = ?
                ''', (pattern,))

                result = cursor.fetchone()
                next_number = result['current_number'] if result else 1

                conn.commit()
                logger.debug(f"[SEQUENCE] pattern={pattern}, next_number={next_number}")
                return next_number

            except Exception as e:
                conn.rollback()
                logger.error(f"[SEQUENCE] 시퀀스 생성 실패: {e}")
                raise


# 싱글톤 패턴으로 데이터베이스 인스턴스 관리
_user_db_instance = None

def get_user_database(db_path: str = None) -> UserDatabase:
    """사용자 데이터베이스 싱글톤 인스턴스 반환"""
    global _user_db_instance
    if _user_db_instance is None:
        _user_db_instance = UserDatabase(db_path=db_path)
    return _user_db_instance


class AuditLogRepository:
    """감사 로그 저장소 (통합)"""

    def __init__(self, user_db: UserDatabase):
        self.user_db = user_db

    def log_action(self, user_email: str, action: str, details: str = None,
                  project_code: str = None, field_name: str = None,
                  old_value: str = None, new_value: str = None,
                  ip_address: str = None) -> bool:
        """액션 로깅"""
        try:
            with self.user_db._get_connection() as conn:
                conn.execute('''
                    INSERT INTO audit_logs (
                        user_email, action, details, project_code,
                        field_name, old_value, new_value, ip_address
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (user_email, action, details, project_code,
                      field_name, old_value, new_value, ip_address))
                conn.commit()
                return True
        except Exception as e:
            logger.error(f"감사 로그 저장 실패: {str(e)}")
            return False

    def log_actions_batch(self, actions: List[Dict]) -> Tuple[bool, int]:
        """
        여러 액션을 한번의 트랜잭션으로 로깅 (성능 최적화)

        Args:
            actions: 로그 데이터 리스트
                [
                    {
                        'user_email': str,
                        'action': str,
                        'details': str (optional),
                        'project_code': str (optional),
                        'field_name': str (optional),
                        'old_value': str (optional),
                        'new_value': str (optional),
                        'ip_address': str (optional)
                    },
                    ...
                ]

        Returns:
            (success: bool, count: int) - 성공 여부와 저장된 로그 수
        """
        if not actions:
            return True, 0

        try:
            with self.user_db._get_connection() as conn:
                for action_data in actions:
                    conn.execute('''
                        INSERT INTO audit_logs (
                            user_email, action, details, project_code,
                            field_name, old_value, new_value, ip_address
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        action_data.get('user_email'),
                        action_data.get('action'),
                        action_data.get('details'),
                        action_data.get('project_code'),
                        action_data.get('field_name'),
                        action_data.get('old_value'),
                        action_data.get('new_value'),
                        action_data.get('ip_address')
                    ))

                conn.commit()  # 한번만 커밋
                return True, len(actions)
        except Exception as e:
            logger.error(f"배치 감사 로그 저장 실패: {str(e)}")
            return False, 0

    def get_recent_logs(self, limit: int = 100, user_email: str = None) -> List[Dict]:
        """최근 로그 조회"""
        try:
            with self.user_db._get_connection() as conn:
                if user_email:
                    cursor = conn.execute('''
                        SELECT * FROM audit_logs
                        WHERE user_email = ?
                        ORDER BY timestamp DESC LIMIT ?
                    ''', (user_email, limit))
                else:
                    cursor = conn.execute('''
                        SELECT * FROM audit_logs
                        ORDER BY timestamp DESC LIMIT ?
                    ''', (limit,))

                return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            logger.error(f"감사 로그 조회 실패: {str(e)}")
            return []


# 싱글톤 감사 로그 저장소 인스턴스
_audit_log_repository = None

def get_audit_repository() -> AuditLogRepository:
    """감사 로그 저장소 싱글톤 인스턴스 반환"""
    global _audit_log_repository
    if _audit_log_repository is None:
        user_db = get_user_database()
        _audit_log_repository = AuditLogRepository(user_db)
    return _audit_log_repository


class CalendarEventRepository:
    """Google Calendar 이벤트 매핑 저장소"""

    def __init__(self, user_db: UserDatabase):
        self.user_db = user_db

    def upsert_event(self, project_code: str, calendar_id: str, event_id: str) -> None:
        with self.user_db._get_connection() as conn:
            conn.execute('''
                INSERT INTO project_calendar_events (project_code, calendar_id, event_id, created_at, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
                ON CONFLICT(project_code) DO UPDATE SET
                    calendar_id = excluded.calendar_id,
                    event_id = excluded.event_id,
                    updated_at = CURRENT_TIMESTAMP
            ''', (project_code, calendar_id, event_id))
            conn.commit()

    def get_event(self, project_code: str) -> Optional[Dict]:
        with self.user_db._get_connection() as conn:
            cursor = conn.execute('''
                SELECT project_code, calendar_id, event_id, created_at, updated_at
                FROM project_calendar_events
                WHERE project_code = ?
            ''', (project_code,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def delete_event(self, project_code: str) -> None:
        with self.user_db._get_connection() as conn:
            conn.execute('DELETE FROM project_calendar_events WHERE project_code = ?', (project_code,))
            conn.commit()

    def update_project_code(self, old_code: str, new_code: str) -> None:
        with self.user_db._get_connection() as conn:
            conn.execute('''
                UPDATE project_calendar_events
                SET project_code = ?, updated_at = CURRENT_TIMESTAMP
                WHERE project_code = ?
            ''', (new_code, old_code))
            conn.commit()


_calendar_event_repository: Optional[CalendarEventRepository] = None


def get_calendar_event_repository() -> CalendarEventRepository:
    """Calendar 이벤트 저장소 싱글톤 인스턴스"""
    global _calendar_event_repository
    if _calendar_event_repository is None:
        user_db = get_user_database()
        _calendar_event_repository = CalendarEventRepository(user_db)
    return _calendar_event_repository
