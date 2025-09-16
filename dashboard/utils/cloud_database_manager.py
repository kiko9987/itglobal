"""
GCP 클라우드 데이터베이스 관리자
- Cloud SQL (PostgreSQL) 통합
- Cloud Memorystore (Redis) 연동
- 멀티 인스턴스 데이터 일관성 보장
"""

import os
import logging
import sqlalchemy
from sqlalchemy import create_engine, text
from sqlalchemy.orm import sessionmaker
from sqlalchemy.pool import NullPool
import redis
from google.cloud.sql.connector import Connector
from typing import Dict, Any, Optional, List
import json
from datetime import datetime, timedelta

logger = logging.getLogger(__name__)

class CloudDatabaseManager:
    """GCP 클라우드 데이터베이스 관리자"""

    def __init__(self):
        self.db_engine = None
        self.redis_client = None
        self.Session = None
        self.connector = None
        self.init_connections()

    def init_connections(self):
        """데이터베이스 연결 초기화"""
        try:
            # Cloud SQL 연결
            self._init_cloud_sql()

            # Cloud Memorystore (Redis) 연결
            self._init_cloud_redis()

            # 기본 테이블 생성
            self._create_tables()

            logger.info("클라우드 데이터베이스 연결 성공")

        except Exception as e:
            logger.error(f"데이터베이스 연결 실패: {e}")
            # 로컬 폴백
            self._init_local_fallback()

    def _init_cloud_sql(self):
        """Cloud SQL 연결 초기화"""
        # 환경변수에서 설정 읽기
        project_id = os.getenv('GOOGLE_CLOUD_PROJECT')
        region = os.getenv('CLOUD_SQL_REGION', 'asia-northeast3')
        instance_name = os.getenv('CLOUD_SQL_INSTANCE', 'itglobal-main')
        database_name = os.getenv('CLOUD_SQL_DATABASE', 'itglobal_db')
        db_user = os.getenv('CLOUD_SQL_USER', 'postgres')
        db_password = os.getenv('CLOUD_SQL_PASSWORD')

        if not all([project_id, db_password]):
            raise ValueError("Cloud SQL 환경변수가 설정되지 않았습니다")

        # Cloud SQL Connector 사용
        self.connector = Connector()

        def getconn():
            conn = self.connector.connect(
                f"{project_id}:{region}:{instance_name}",
                "pg8000",
                user=db_user,
                password=db_password,
                db=database_name,
            )
            return conn

        # SQLAlchemy 엔진 생성
        self.db_engine = create_engine(
            "postgresql+pg8000://",
            creator=getconn,
            poolclass=NullPool,
        )

        self.Session = sessionmaker(bind=self.db_engine)

    def _init_cloud_redis(self):
        """Cloud Memorystore (Redis) 연결 초기화"""
        redis_host = os.getenv('CLOUD_REDIS_HOST')
        redis_port = int(os.getenv('CLOUD_REDIS_PORT', 6379))
        redis_auth = os.getenv('CLOUD_REDIS_AUTH')

        if redis_host:
            self.redis_client = redis.Redis(
                host=redis_host,
                port=redis_port,
                password=redis_auth,
                decode_responses=True,
                socket_connect_timeout=5,
                socket_timeout=5,
                retry_on_timeout=True
            )

            # 연결 테스트
            self.redis_client.ping()
            logger.info(f"Cloud Redis 연결 성공: {redis_host}:{redis_port}")
        else:
            # 로컬 Redis 폴백
            self.redis_client = redis.Redis(host='localhost', port=6379, decode_responses=True)

    def _init_local_fallback(self):
        """로컬 환경 폴백"""
        logger.warning("클라우드 연결 실패 - 로컬 환경으로 폴백")

        # SQLite 폴백
        db_path = os.path.join(os.path.dirname(__file__), '..', 'database.db')
        self.db_engine = create_engine(f'sqlite:///{db_path}')
        self.Session = sessionmaker(bind=self.db_engine)

        # 로컬 Redis 폴백
        try:
            self.redis_client = redis.Redis(host='localhost', port=6379, decode_responses=True)
            self.redis_client.ping()
        except:
            logger.warning("Redis 연결 실패 - 메모리 캐시로 폴백")
            self.redis_client = None

    def _create_tables(self):
        """필수 테이블 생성"""
        tables_sql = """
        -- 사용자 테이블
        CREATE TABLE IF NOT EXISTS users (
            id VARCHAR(255) PRIMARY KEY,
            email VARCHAR(255) UNIQUE NOT NULL,
            name VARCHAR(255) NOT NULL,
            role VARCHAR(50) NOT NULL DEFAULT 'viewer',
            permissions JSON,
            regions JSON,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- 세션 테이블 (멀티 인스턴스 세션 공유)
        CREATE TABLE IF NOT EXISTS user_sessions (
            session_id VARCHAR(255) PRIMARY KEY,
            user_id VARCHAR(255) NOT NULL,
            session_data JSON NOT NULL,
            expires_at TIMESTAMP NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        );

        -- 감사 로그 테이블
        CREATE TABLE IF NOT EXISTS audit_logs (
            id SERIAL PRIMARY KEY,
            user_id VARCHAR(255),
            action VARCHAR(100) NOT NULL,
            resource_type VARCHAR(100),
            resource_id VARCHAR(255),
            details JSON,
            ip_address VARCHAR(45),
            user_agent TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- 시스템 설정 테이블
        CREATE TABLE IF NOT EXISTS system_settings (
            key VARCHAR(255) PRIMARY KEY,
            value JSON NOT NULL,
            description TEXT,
            updated_by VARCHAR(255),
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- 프로젝트 캐시 테이블 (Google Sheets 캐시)
        CREATE TABLE IF NOT EXISTS project_cache (
            cache_key VARCHAR(255) PRIMARY KEY,
            cache_data JSON NOT NULL,
            expires_at TIMESTAMP NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        -- 인덱스 생성
        CREATE INDEX IF NOT EXISTS idx_sessions_expires ON user_sessions(expires_at);
        CREATE INDEX IF NOT EXISTS idx_audit_user_time ON audit_logs(user_id, created_at);
        CREATE INDEX IF NOT EXISTS idx_cache_expires ON project_cache(expires_at);
        """

        with self.db_engine.connect() as conn:
            # PostgreSQL의 경우 개별 문장으로 실행
            statements = [stmt.strip() for stmt in tables_sql.split(';') if stmt.strip()]
            for statement in statements:
                try:
                    conn.execute(text(statement))
                    conn.commit()
                except Exception as e:
                    if "already exists" not in str(e).lower():
                        logger.error(f"테이블 생성 실패: {statement[:50]}... - {e}")

    def get_user(self, user_id: str) -> Optional[Dict[str, Any]]:
        """사용자 정보 조회"""
        try:
            with self.Session() as session:
                result = session.execute(
                    text("SELECT * FROM users WHERE id = :user_id"),
                    {"user_id": user_id}
                ).fetchone()

                if result:
                    return {
                        'id': result.id,
                        'email': result.email,
                        'name': result.name,
                        'role': result.role,
                        'permissions': json.loads(result.permissions) if result.permissions else [],
                        'regions': json.loads(result.regions) if result.regions else []
                    }
                return None
        except Exception as e:
            logger.error(f"사용자 조회 실패: {e}")
            return None

    def save_user(self, user_data: Dict[str, Any]) -> bool:
        """사용자 정보 저장"""
        try:
            with self.Session() as session:
                session.execute(text("""
                    INSERT INTO users (id, email, name, role, permissions, regions, updated_at)
                    VALUES (:id, :email, :name, :role, :permissions, :regions, CURRENT_TIMESTAMP)
                    ON CONFLICT (id) DO UPDATE SET
                        email = EXCLUDED.email,
                        name = EXCLUDED.name,
                        role = EXCLUDED.role,
                        permissions = EXCLUDED.permissions,
                        regions = EXCLUDED.regions,
                        updated_at = CURRENT_TIMESTAMP
                """), {
                    'id': user_data['id'],
                    'email': user_data['email'],
                    'name': user_data['name'],
                    'role': user_data.get('role', 'viewer'),
                    'permissions': json.dumps(user_data.get('permissions', [])),
                    'regions': json.dumps(user_data.get('regions', []))
                })
                session.commit()
                return True
        except Exception as e:
            logger.error(f"사용자 저장 실패: {e}")
            return False

    def get_session(self, session_id: str) -> Optional[Dict[str, Any]]:
        """세션 데이터 조회"""
        try:
            with self.Session() as session:
                result = session.execute(
                    text("SELECT session_data FROM user_sessions WHERE session_id = :session_id AND expires_at > CURRENT_TIMESTAMP"),
                    {"session_id": session_id}
                ).fetchone()

                if result:
                    return json.loads(result.session_data)
                return None
        except Exception as e:
            logger.error(f"세션 조회 실패: {e}")
            return None

    def save_session(self, session_id: str, user_id: str, session_data: Dict[str, Any], expires_at: datetime) -> bool:
        """세션 데이터 저장"""
        try:
            with self.Session() as session:
                session.execute(text("""
                    INSERT INTO user_sessions (session_id, user_id, session_data, expires_at)
                    VALUES (:session_id, :user_id, :session_data, :expires_at)
                    ON CONFLICT (session_id) DO UPDATE SET
                        user_id = EXCLUDED.user_id,
                        session_data = EXCLUDED.session_data,
                        expires_at = EXCLUDED.expires_at
                """), {
                    'session_id': session_id,
                    'user_id': user_id,
                    'session_data': json.dumps(session_data),
                    'expires_at': expires_at
                })
                session.commit()
                return True
        except Exception as e:
            logger.error(f"세션 저장 실패: {e}")
            return False

    def log_audit(self, user_id: str, action: str, details: Dict[str, Any],
                  resource_type: str = None, resource_id: str = None,
                  ip_address: str = None, user_agent: str = None) -> bool:
        """감사 로그 기록"""
        try:
            with self.Session() as session:
                session.execute(text("""
                    INSERT INTO audit_logs (user_id, action, resource_type, resource_id,
                                          details, ip_address, user_agent)
                    VALUES (:user_id, :action, :resource_type, :resource_id,
                           :details, :ip_address, :user_agent)
                """), {
                    'user_id': user_id,
                    'action': action,
                    'resource_type': resource_type,
                    'resource_id': resource_id,
                    'details': json.dumps(details),
                    'ip_address': ip_address,
                    'user_agent': user_agent
                })
                session.commit()
                return True
        except Exception as e:
            logger.error(f"감사 로그 저장 실패: {e}")
            return False

    def cache_project_data(self, cache_key: str, data: Any, ttl_seconds: int = 300) -> bool:
        """프로젝트 데이터 캐싱"""
        try:
            expires_at = datetime.now() + timedelta(seconds=ttl_seconds)

            # Redis 우선 캐싱
            if self.redis_client:
                try:
                    self.redis_client.setex(
                        f"project:{cache_key}",
                        ttl_seconds,
                        json.dumps(data, default=str)
                    )
                except Exception as redis_error:
                    logger.warning(f"Redis 캐싱 실패: {redis_error}")

            # 데이터베이스 캐싱 (폴백)
            with self.Session() as session:
                session.execute(text("""
                    INSERT INTO project_cache (cache_key, cache_data, expires_at)
                    VALUES (:cache_key, :cache_data, :expires_at)
                    ON CONFLICT (cache_key) DO UPDATE SET
                        cache_data = EXCLUDED.cache_data,
                        expires_at = EXCLUDED.expires_at
                """), {
                    'cache_key': cache_key,
                    'cache_data': json.dumps(data, default=str),
                    'expires_at': expires_at
                })
                session.commit()

            return True
        except Exception as e:
            logger.error(f"데이터 캐싱 실패: {e}")
            return False

    def get_cached_project_data(self, cache_key: str) -> Optional[Any]:
        """캐시된 프로젝트 데이터 조회"""
        try:
            # Redis 우선 조회
            if self.redis_client:
                try:
                    cached = self.redis_client.get(f"project:{cache_key}")
                    if cached:
                        return json.loads(cached)
                except Exception as redis_error:
                    logger.warning(f"Redis 조회 실패: {redis_error}")

            # 데이터베이스 폴백
            with self.Session() as session:
                result = session.execute(
                    text("SELECT cache_data FROM project_cache WHERE cache_key = :cache_key AND expires_at > CURRENT_TIMESTAMP"),
                    {"cache_key": cache_key}
                ).fetchone()

                if result:
                    return json.loads(result.cache_data)

            return None
        except Exception as e:
            logger.error(f"캐시 조회 실패: {e}")
            return None

    def cleanup_expired_data(self) -> int:
        """만료된 데이터 정리"""
        cleaned_count = 0
        try:
            with self.Session() as session:
                # 만료된 세션 정리
                result = session.execute(
                    text("DELETE FROM user_sessions WHERE expires_at < CURRENT_TIMESTAMP")
                )
                cleaned_count += result.rowcount

                # 만료된 캐시 정리
                result = session.execute(
                    text("DELETE FROM project_cache WHERE expires_at < CURRENT_TIMESTAMP")
                )
                cleaned_count += result.rowcount

                session.commit()

            logger.info(f"만료된 데이터 정리 완료: {cleaned_count}개")
            return cleaned_count
        except Exception as e:
            logger.error(f"데이터 정리 실패: {e}")
            return 0

    def get_system_health(self) -> Dict[str, Any]:
        """시스템 상태 확인"""
        health = {
            'database': False,
            'redis': False,
            'timestamp': datetime.now().isoformat()
        }

        # 데이터베이스 상태 확인
        try:
            with self.db_engine.connect() as conn:
                conn.execute(text("SELECT 1"))
                health['database'] = True
        except Exception as e:
            logger.error(f"데이터베이스 상태 확인 실패: {e}")

        # Redis 상태 확인
        try:
            if self.redis_client:
                self.redis_client.ping()
                health['redis'] = True
        except Exception as e:
            logger.error(f"Redis 상태 확인 실패: {e}")

        return health

    def close_connections(self):
        """연결 종료"""
        try:
            if self.connector:
                self.connector.close()
            if self.redis_client:
                self.redis_client.close()
        except Exception as e:
            logger.error(f"연결 종료 실패: {e}")


# 전역 인스턴스
cloud_db_manager = None

def get_cloud_db_manager() -> CloudDatabaseManager:
    """클라우드 데이터베이스 매니저 인스턴스 반환"""
    global cloud_db_manager
    if cloud_db_manager is None:
        cloud_db_manager = CloudDatabaseManager()
    return cloud_db_manager

def init_cloud_database():
    """클라우드 데이터베이스 초기화"""
    global cloud_db_manager
    cloud_db_manager = CloudDatabaseManager()
    return cloud_db_manager