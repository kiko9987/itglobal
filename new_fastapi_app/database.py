"""
데이터베이스 설정 및 연결 관리
"""

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, Session
from sqlalchemy.pool import StaticPool
from contextlib import contextmanager
import logging

from config import settings
from models import Base

# 로깅 설정
logger = logging.getLogger(__name__)

# SQLAlchemy 엔진 생성
engine = create_engine(
    settings.database_url,
    connect_args={"check_same_thread": False},  # SQLite용
    poolclass=StaticPool,
    echo=settings.debug  # 디버그 모드에서 SQL 쿼리 로깅
)

# 세션 팩토리 생성
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

def create_tables():
    """데이터베이스 테이블 생성"""
    try:
        Base.metadata.create_all(bind=engine)
        logger.info("✅ 데이터베이스 테이블 생성 완료")
    except Exception as e:
        logger.error(f"❌ 데이터베이스 테이블 생성 실패: {e}")
        raise

def get_db() -> Session:
    """데이터베이스 세션 의존성"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

@contextmanager
def get_db_context():
    """컨텍스트 매니저로 데이터베이스 세션 사용"""
    db = SessionLocal()
    try:
        yield db
        db.commit()
    except Exception as e:
        db.rollback()
        logger.error(f"데이터베이스 트랜잭션 롤백: {e}")
        raise
    finally:
        db.close()

def init_database():
    """데이터베이스 초기화"""
    try:
        # 테이블 생성
        create_tables()

        # 기본 관리자 계정 생성
        from services.user_service import create_default_admin_users
        create_default_admin_users()

        logger.info("✅ 데이터베이스 초기화 완료")

    except Exception as e:
        logger.error(f"❌ 데이터베이스 초기화 실패: {e}")
        raise

# 데이터베이스 헬스체크
def check_database_health() -> dict:
    """데이터베이스 연결 상태 확인"""
    try:
        with get_db_context() as db:
            # 간단한 쿼리 실행
            result = db.execute("SELECT 1").fetchone()
            return {
                "status": "healthy",
                "message": "데이터베이스 연결 정상"
            }
    except Exception as e:
        logger.error(f"데이터베이스 헬스체크 실패: {e}")
        return {
            "status": "unhealthy",
            "message": f"데이터베이스 연결 실패: {str(e)}"
        }