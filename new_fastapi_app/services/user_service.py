"""
사용자 관리 서비스
"""

from sqlalchemy.orm import Session
from typing import Optional, List
from datetime import datetime
import logging

from models import User
from config import settings

logger = logging.getLogger(__name__)

class UserService:
    """사용자 관리 서비스"""

    @staticmethod
    def get_user_by_email(db: Session, email: str) -> Optional[User]:
        """이메일로 사용자 조회"""
        try:
            return db.query(User).filter(User.email == email).first()
        except Exception as e:
            logger.error(f"사용자 조회 실패 (email: {email}): {e}")
            return None

    @staticmethod
    def get_user_by_id(db: Session, user_id: int) -> Optional[User]:
        """ID로 사용자 조회"""
        try:
            return db.query(User).filter(User.id == user_id).first()
        except Exception as e:
            logger.error(f"사용자 조회 실패 (id: {user_id}): {e}")
            return None

    @staticmethod
    def create_user(db: Session, email: str, name: str, role: str = "viewer") -> User:
        """새 사용자 생성"""
        try:
            user = User(
                email=email,
                name=name,
                role=role,
                is_active=True
            )

            db.add(user)
            db.commit()
            db.refresh(user)

            logger.info(f"새 사용자 생성: {email} ({role})")
            return user

        except Exception as e:
            db.rollback()
            logger.error(f"사용자 생성 실패: {e}")
            raise

    @staticmethod
    def create_default_users(db: Session) -> List[User]:
        """기본 사용자들 생성"""
        created_users = []

        try:
            # 관리자 사용자들 생성
            for admin_email in settings.admin_emails:
                existing_user = UserService.get_user_by_email(db, admin_email)
                if not existing_user:
                    # 이메일에서 이름 추출 (@ 앞부분)
                    name = admin_email.split('@')[0]
                    user = UserService.create_user(db, admin_email, name, "admin")
                    created_users.append(user)
                    logger.info(f"기본 관리자 생성: {admin_email}")
                else:
                    # 기존 사용자의 역할을 관리자로 업데이트
                    if existing_user.role != "admin":
                        existing_user.role = "admin"
                        db.commit()
                        logger.info(f"기존 사용자를 관리자로 승격: {admin_email}")

            # 편집자 사용자들 생성
            for editor_email in settings.editor_emails:
                existing_user = UserService.get_user_by_email(db, editor_email)
                if not existing_user:
                    name = editor_email.split('@')[0]
                    user = UserService.create_user(db, editor_email, name, "editor")
                    created_users.append(user)
                    logger.info(f"기본 편집자 생성: {editor_email}")
                else:
                    # 기존 사용자의 역할을 편집자로 업데이트 (관리자가 아닌 경우)
                    if existing_user.role not in ["admin", "editor"]:
                        existing_user.role = "editor"
                        db.commit()
                        logger.info(f"기존 사용자를 편집자로 승격: {editor_email}")

            return created_users

        except Exception as e:
            logger.error(f"기본 사용자 생성 실패: {e}")
            raise

    @staticmethod
    def get_or_create_user_from_email(db: Session, email: str, name: str = None) -> User:
        """이메일로 사용자 조회하거나 새로 생성"""
        try:
            # 기존 사용자 확인
            user = UserService.get_user_by_email(db, email)
            if user:
                return user

            # 새 사용자 생성
            if not name:
                name = email.split('@')[0]

            # 역할 결정
            role = "viewer"  # 기본값
            if email in settings.admin_emails:
                role = "admin"
            elif email in settings.editor_emails:
                role = "editor"

            return UserService.create_user(db, email, name, role)

        except Exception as e:
            logger.error(f"사용자 조회/생성 실패: {e}")
            raise

    @staticmethod
    def update_user_last_login(db: Session, user_id: int) -> bool:
        """사용자 마지막 로그인 시간 업데이트"""
        try:
            user = UserService.get_user_by_id(db, user_id)
            if user:
                user.last_login = datetime.utcnow()
                db.commit()
                return True
            return False

        except Exception as e:
            logger.error(f"마지막 로그인 시간 업데이트 실패: {e}")
            return False

    @staticmethod
    def deactivate_user(db: Session, user_id: int) -> bool:
        """사용자 비활성화"""
        try:
            user = UserService.get_user_by_id(db, user_id)
            if user:
                user.is_active = False
                db.commit()
                logger.info(f"사용자 비활성화: {user.email}")
                return True
            return False

        except Exception as e:
            logger.error(f"사용자 비활성화 실패: {e}")
            return False

    @staticmethod
    def get_all_users(db: Session, include_inactive: bool = False) -> List[User]:
        """모든 사용자 조회"""
        try:
            query = db.query(User)
            if not include_inactive:
                query = query.filter(User.is_active == True)

            return query.order_by(User.created_at.desc()).all()

        except Exception as e:
            logger.error(f"사용자 목록 조회 실패: {e}")
            return []