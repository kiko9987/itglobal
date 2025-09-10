"""
SQLite 데이터베이스 관리 시스템
파일 기반 사용자 관리를 데이터베이스로 마이그레이션
"""

import sqlite3
import os
import json
import threading
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from contextlib import contextmanager
import logging

logger = logging.getLogger(__name__)

class DatabaseManager:
    """SQLite 데이터베이스 매니저"""
    
    def __init__(self, db_path: str = None):
        if db_path is None:
            db_path = os.path.join(os.path.dirname(__file__), '..', 'data', 'dashboard.db')
        
        self.db_path = db_path
        self._lock = threading.RLock()
        self._ensure_database_directory()
        self._init_database()
    
    def _ensure_database_directory(self):
        """데이터베이스 디렉토리 생성"""
        db_dir = os.path.dirname(self.db_path)
        if not os.path.exists(db_dir):
            os.makedirs(db_dir)
    
    def _init_database(self):
        """데이터베이스 초기화 및 테이블 생성"""
        with self.get_connection() as conn:
            # 사용자 테이블
            conn.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT UNIQUE NOT NULL,
                    name TEXT NOT NULL,
                    password_hash TEXT,
                    permission_level TEXT DEFAULT 'viewer',
                    is_active BOOLEAN DEFAULT 1,
                    auth_method TEXT DEFAULT 'local',
                    picture TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    last_login TIMESTAMP,
                    UNIQUE(email)
                )
            ''')
            
            # 세션 테이블 (JWT 리프레시 토큰 저장용)
            conn.execute('''
                CREATE TABLE IF NOT EXISTS user_sessions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_email TEXT NOT NULL,
                    refresh_token_id TEXT UNIQUE NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    expires_at TIMESTAMP NOT NULL,
                    is_active BOOLEAN DEFAULT 1,
                    FOREIGN KEY (user_email) REFERENCES users (email) ON DELETE CASCADE
                )
            ''')
            
            # 감사 로그 테이블
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
            
            # 프로젝트 메타데이터 캐시 테이블
            conn.execute('''
                CREATE TABLE IF NOT EXISTS project_cache (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_code TEXT UNIQUE NOT NULL,
                    data JSON NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # 시스템 설정 테이블
            conn.execute('''
                CREATE TABLE IF NOT EXISTS system_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL,
                    description TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            # 인덱스 생성
            conn.execute('CREATE INDEX IF NOT EXISTS idx_users_email ON users(email)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_user_email ON audit_logs(user_email)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_audit_logs_timestamp ON audit_logs(timestamp)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_sessions_user_email ON user_sessions(user_email)')
            conn.execute('CREATE INDEX IF NOT EXISTS idx_sessions_expires_at ON user_sessions(expires_at)')
            
            conn.commit()
            logger.info("데이터베이스 초기화 완료")
    
    @contextmanager
    def get_connection(self):
        """Thread-safe 데이터베이스 연결"""
        with self._lock:
            conn = sqlite3.connect(
                self.db_path, 
                timeout=30.0,
                check_same_thread=False
            )
            conn.row_factory = sqlite3.Row  # 딕셔너리 스타일 접근
            conn.execute('PRAGMA foreign_keys = ON')  # 외래 키 제약 조건 활성화
            try:
                yield conn
            except Exception as e:
                conn.rollback()
                logger.error(f"데이터베이스 오류: {str(e)}")
                raise
            finally:
                conn.close()
    
    def migrate_from_json_users(self, json_file_path: str) -> int:
        """JSON 파일에서 사용자 데이터 마이그레이션"""
        if not os.path.exists(json_file_path):
            logger.info("마이그레이션할 JSON 파일이 없습니다.")
            return 0
        
        try:
            with open(json_file_path, 'r', encoding='utf-8') as f:
                users_data = json.load(f)
            
            migrated_count = 0
            with self.get_connection() as conn:
                for email, user_data in users_data.items():
                    try:
                        conn.execute('''
                            INSERT OR IGNORE INTO users (
                                email, name, password_hash, permission_level, 
                                is_active, auth_method, picture, created_at, last_login
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            email,
                            user_data.get('name', ''),
                            user_data.get('password_hash', ''),
                            user_data.get('permission_level', 'viewer'),
                            user_data.get('is_active', True),
                            user_data.get('auth_method', 'local'),
                            user_data.get('picture', ''),
                            user_data.get('created_at'),
                            user_data.get('last_login')
                        ))
                        migrated_count += 1
                    except Exception as e:
                        logger.error(f"사용자 {email} 마이그레이션 실패: {str(e)}")
                
                conn.commit()
            
            logger.info(f"{migrated_count}명의 사용자를 마이그레이션했습니다.")
            return migrated_count
            
        except Exception as e:
            logger.error(f"JSON 마이그레이션 실패: {str(e)}")
            return 0

class UserRepository:
    """사용자 데이터 저장소"""
    
    def __init__(self, db_manager: DatabaseManager):
        self.db = db_manager
    
    def create_user(self, email: str, name: str, password_hash: str = '', 
                   permission_level: str = 'viewer', auth_method: str = 'local',
                   picture: str = '') -> bool:
        """사용자 생성"""
        try:
            with self.db.get_connection() as conn:
                conn.execute('''
                    INSERT INTO users (
                        email, name, password_hash, permission_level, 
                        auth_method, picture
                    ) VALUES (?, ?, ?, ?, ?, ?)
                ''', (email, name, password_hash, permission_level, auth_method, picture))
                conn.commit()
                return True
        except sqlite3.IntegrityError:
            logger.warning(f"사용자 {email} 이미 존재")
            return False
        except Exception as e:
            logger.error(f"사용자 생성 실패: {str(e)}")
            return False
    
    def get_user_by_email(self, email: str) -> Optional[Dict[str, Any]]:
        """이메일로 사용자 조회"""
        try:
            with self.db.get_connection() as conn:
                cursor = conn.execute(
                    'SELECT * FROM users WHERE email = ? AND is_active = 1', 
                    (email,)
                )
                row = cursor.fetchone()
                return dict(row) if row else None
        except Exception as e:
            logger.error(f"사용자 조회 실패: {str(e)}")
            return None
    
    def update_user(self, email: str, **kwargs) -> bool:
        """사용자 정보 업데이트"""
        if not kwargs:
            return False
        
        # 허용된 필드만 업데이트
        allowed_fields = {
            'name', 'password_hash', 'permission_level', 
            'is_active', 'picture', 'last_login'
        }
        
        update_fields = {k: v for k, v in kwargs.items() if k in allowed_fields}
        if not update_fields:
            return False
        
        try:
            with self.db.get_connection() as conn:
                set_clause = ', '.join(f'{field} = ?' for field in update_fields.keys())
                query = f'UPDATE users SET {set_clause} WHERE email = ?'
                
                conn.execute(query, list(update_fields.values()) + [email])
                conn.commit()
                return conn.total_changes > 0
        except Exception as e:
            logger.error(f"사용자 업데이트 실패: {str(e)}")
            return False
    
    def get_all_users(self) -> List[Dict[str, Any]]:
        """모든 사용자 조회"""
        try:
            with self.db.get_connection() as conn:
                cursor = conn.execute('''
                    SELECT id, email, name, permission_level, is_active, 
                           auth_method, picture, created_at, last_login 
                    FROM users ORDER BY created_at DESC
                ''')
                return [dict(row) for row in cursor.fetchall()]
        except Exception as e:
            logger.error(f"사용자 목록 조회 실패: {str(e)}")
            return []
    
    def delete_user(self, email: str) -> bool:
        """사용자 삭제"""
        try:
            with self.db.get_connection() as conn:
                conn.execute('DELETE FROM users WHERE email = ?', (email,))
                conn.commit()
                return conn.total_changes > 0
        except Exception as e:
            logger.error(f"사용자 삭제 실패: {str(e)}")
            return False

class AuditLogRepository:
    """감사 로그 저장소"""
    
    def __init__(self, db_manager: DatabaseManager):
        self.db = db_manager
    
    def log_action(self, user_email: str, action: str, details: str = None,
                  project_code: str = None, field_name: str = None,
                  old_value: str = None, new_value: str = None,
                  ip_address: str = None) -> bool:
        """액션 로깅"""
        try:
            with self.db.get_connection() as conn:
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
    
    def get_recent_logs(self, limit: int = 100, user_email: str = None) -> List[Dict[str, Any]]:
        """최근 로그 조회"""
        try:
            with self.db.get_connection() as conn:
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

# 전역 데이터베이스 인스턴스
_db_manager = None
_user_repository = None
_audit_repository = None

def init_database(db_path: str = None):
    """데이터베이스 초기화"""
    global _db_manager, _user_repository, _audit_repository
    
    _db_manager = DatabaseManager(db_path)
    _user_repository = UserRepository(_db_manager)
    _audit_repository = AuditLogRepository(_db_manager)
    
    logger.info("데이터베이스 시스템 초기화 완료")

def get_user_repository() -> UserRepository:
    """사용자 저장소 반환"""
    if _user_repository is None:
        raise RuntimeError("Database not initialized")
    return _user_repository

def get_audit_repository() -> AuditLogRepository:
    """감사 로그 저장소 반환"""
    if _audit_repository is None:
        raise RuntimeError("Database not initialized")
    return _audit_repository

def get_db_manager() -> DatabaseManager:
    """데이터베이스 매니저 반환"""
    if _db_manager is None:
        raise RuntimeError("Database not initialized")
    return _db_manager