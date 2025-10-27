"""
인증 관련 서비스
"""

from fastapi import Depends, HTTPException, status, Request
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from jose import jwt, JWTError
from sqlalchemy.orm import Session
from typing import Optional
import logging

from database import get_db
from models import User
from config import settings

logger = logging.getLogger(__name__)

# JWT 토큰 스키마
security = HTTPBearer(auto_error=False)

class AuthService:
    """인증 서비스"""

    @staticmethod
    def create_access_token(data: dict) -> str:
        """액세스 토큰 생성"""
        try:
            payload = data.copy()
            encoded_jwt = jwt.encode(payload, settings.secret_key, algorithm="HS256")
            return encoded_jwt
        except Exception as e:
            logger.error(f"토큰 생성 실패: {e}")
            raise

    @staticmethod
    def verify_token(token: str) -> Optional[dict]:
        """토큰 검증"""
        try:
            payload = jwt.decode(token, settings.secret_key, algorithms=["HS256"])
            return payload
        except JWTError as e:
            logger.warning(f"토큰 검증 실패: {e}")
            return None

    @staticmethod
    def get_user_by_email(db: Session, email: str) -> Optional[User]:
        """이메일로 사용자 조회"""
        return db.query(User).filter(User.email == email).first()

    @staticmethod
    def create_or_update_user(db: Session, user_data: dict) -> User:
        """사용자 생성 또는 업데이트 (OAuth용)"""
        try:
            email = user_data.get("email")
            existing_user = db.query(User).filter(User.email == email).first()

            if existing_user:
                # 기존 사용자 정보 업데이트
                existing_user.name = user_data.get("name", existing_user.name)
                existing_user.google_id = user_data.get("google_id", existing_user.google_id)
                existing_user.is_active = True
                db.commit()
                db.refresh(existing_user)
                return existing_user
            else:
                # 새 사용자 생성
                role = "admin" if email in settings.admin_emails else "user"
                new_user = User(
                    email=email,
                    name=user_data.get("name"),
                    google_id=user_data.get("google_id"),
                    role=role,
                    is_active=True
                )
                db.add(new_user)
                db.commit()
                db.refresh(new_user)
                return new_user

        except Exception as e:
            db.rollback()
            logger.error(f"사용자 생성/업데이트 실패: {e}")
            raise

# 의존성 함수들
async def get_current_user(
    request: Request,
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(security),
    db: Session = Depends(get_db)
) -> User:
    """현재 사용자 가져오기"""

    # 개발 모드에서는 기본 사용자 사용 (임시)
    if settings.debug and not credentials:
        # 첫 번째 관리자를 기본 사용자로 사용
        admin_email = settings.admin_emails[0] if settings.admin_emails else "admin@example.com"
        user = db.query(User).filter(User.email == admin_email).first()

        if not user:
            # 기본 관리자 생성
            user = User(
                email=admin_email,
                name="관리자",
                role="admin",
                is_active=True
            )
            db.add(user)
            db.commit()
            db.refresh(user)

        return user

    if not credentials:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="인증이 필요합니다",
            headers={"WWW-Authenticate": "Bearer"},
        )

    # 토큰 검증
    payload = AuthService.verify_token(credentials.credentials)
    if not payload:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="유효하지 않은 토큰입니다",
            headers={"WWW-Authenticate": "Bearer"},
        )

    # 사용자 조회
    email = payload.get("email")
    if not email:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="토큰에 사용자 정보가 없습니다",
            headers={"WWW-Authenticate": "Bearer"},
        )

    user = db.query(User).filter(User.email == email).first()
    if not user or not user.is_active:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="사용자를 찾을 수 없거나 비활성화되었습니다",
            headers={"WWW-Authenticate": "Bearer"},
        )

    return user

async def get_admin_user(current_user: User = Depends(get_current_user)) -> User:
    """관리자 권한 확인"""
    if not current_user.is_admin:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="관리자 권한이 필요합니다"
        )
    return current_user

async def get_editor_user(current_user: User = Depends(get_current_user)) -> User:
    """편집자 권한 확인"""
    if not current_user.can_edit:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="편집 권한이 필요합니다"
        )
    return current_user

# 선택적 인증 (로그인하지 않아도 접근 가능)
async def get_current_user_optional(
    request: Request,
    credentials: Optional[HTTPAuthorizationCredentials] = Depends(security),
    db: Session = Depends(get_db)
) -> Optional[User]:
    """현재 사용자 가져오기 (선택적)"""
    try:
        return await get_current_user(request, credentials, db)
    except HTTPException:
        return None