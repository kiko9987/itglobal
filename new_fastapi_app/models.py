"""
데이터베이스 모델 정의
"""

from sqlalchemy import Column, Integer, String, Text, Boolean, DateTime, ForeignKey, Numeric
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from datetime import datetime
from typing import Optional

Base = declarative_base()

class Project(Base):
    """프로젝트 모델 - Google Sheets의 39개 컬럼을 모두 포함"""
    __tablename__ = "projects"

    # Primary Key
    id = Column(Integer, primary_key=True, index=True)

    # 기본 정보 (1-16)
    project_code = Column(String(50), unique=True, index=True)  # 프로젝트 코드
    site_name = Column(String(200))  # 현장명
    business = Column(String(100))  # 사업자
    contractor = Column(String(100))  # 시공자
    site_address = Column(Text)  # 현장 주소
    site_phone = Column(String(50))  # 현장 전화
    construction_classification = Column(String(100))  # 공사 분류
    manager = Column(String(100))  # 담당자
    manager_phone = Column(String(50))  # 담당자 전화
    manager_address = Column(Text)  # 담당자 주소
    manager_email = Column(String(200))  # 담당자 이메일
    manager_region = Column(String(100))  # 담당자 지역
    created_date = Column(String(50))  # 등록일
    site_manager = Column(String(100))  # 현장 담당자
    customer_name = Column(String(100))  # 고객명
    customer_email = Column(String(200))  # 고객 이메일

    # 금액 정보 (17-23)
    total_amount_1 = Column(Numeric(15, 2))  # 총액1
    tax_amount = Column(Numeric(15, 2))  # 부가세
    total_amount_2 = Column(Numeric(15, 2))  # 총액2
    contract_deposit = Column(Numeric(15, 2))  # 계약금
    mid_payment = Column(Numeric(15, 2))  # 중도금
    final_payment = Column(Numeric(15, 2))  # 잔금
    outstanding_amount = Column(Numeric(15, 2))  # 미수금

    # 일정 정보 (24-32)
    construction_confirmed = Column(String(50))  # 공사확정
    construction_date = Column(String(50))  # 공사 날짜
    construction_confirmed_yn = Column(String(10))  # 공사 확정 여부
    installation_brand = Column(String(100))  # 설치 브랜드
    delivery_date = Column(String(50))  # 납품일
    delivery_note = Column(Text)  # 납품 비고
    installation_note = Column(Text)  # 설치 특이사항
    completion_status = Column(String(50))  # 완료 상태

    # 수금 관련 (33-39)
    collection_notes = Column(Text)  # 수금 관련 특이사항
    collection_auto_account = Column(String(200))  # 수금 자동 계좌번호
    mid_payment_date = Column(String(50))  # 중도금 입금일
    final_payment_date = Column(String(50))  # 잔금 입금일
    total_profit_loss_info = Column(Text)  # 총 손익정보 편집 내용
    construction_confirmed_final = Column(String(50))  # 공사 확정 (최종)
    airtable_record_id = Column(String(100))  # Airtable Record ID

    # 추가 메타데이터
    is_cancelled = Column(Boolean, default=False, index=True)  # 취소 여부 (파생 필드)
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    def __repr__(self):
        return f"<Project(code={self.project_code}, name={self.site_name})>"

    @property
    def status(self) -> str:
        """프로젝트 상태 반환"""
        if self.is_cancelled or self.collection_notes == "공사취소":
            return "cancelled"
        elif self.construction_confirmed:
            return "confirmed"
        else:
            return "pending"

    @property
    def status_display(self) -> str:
        """상태 표시용 텍스트"""
        status_map = {
            "cancelled": "취소됨",
            "confirmed": "확정됨",
            "pending": "대기중"
        }
        return status_map.get(self.status, "알 수 없음")


class User(Base):
    """사용자 모델"""
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    email = Column(String(200), unique=True, index=True)
    name = Column(String(100))
    role = Column(String(50), default="user")  # admin, editor, user
    is_active = Column(Boolean, default=True)
    hashed_password = Column(String(255))  # OAuth 사용 시 비어있을 수 있음
    google_id = Column(String(100), unique=True, nullable=True)  # Google OAuth ID

    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    # 관계
    audit_logs = relationship("AuditLog", back_populates="user")
    project_locks = relationship("ProjectLock", back_populates="user")

    def __repr__(self):
        return f"<User(email={self.email}, name={self.name})>"

    @property
    def is_admin(self) -> bool:
        """관리자 여부"""
        return self.role == "admin"

    @property
    def can_edit(self) -> bool:
        """편집 권한 여부"""
        return self.role in ["admin", "editor"]


class AuditLog(Base):
    """감사 로그 모델"""
    __tablename__ = "audit_logs"

    id = Column(Integer, primary_key=True, index=True)
    user_id = Column(Integer, ForeignKey("users.id"))
    project_code = Column(String(50), index=True)
    action = Column(String(100))  # CREATE, UPDATE, DELETE, CANCEL, RESUME
    field_name = Column(String(100))
    old_value = Column(Text)
    new_value = Column(Text)
    description = Column(Text)
    ip_address = Column(String(45))
    user_agent = Column(Text)

    created_at = Column(DateTime, default=datetime.utcnow, index=True)

    # 관계
    user = relationship("User", back_populates="audit_logs")

    def __repr__(self):
        return f"<AuditLog(project={self.project_code}, action={self.action})>"


class ProjectLock(Base):
    """프로젝트 잠금 모델"""
    __tablename__ = "project_locks"

    id = Column(Integer, primary_key=True, index=True)
    project_code = Column(String(50), index=True)
    user_id = Column(Integer, ForeignKey("users.id"))
    locked_fields = Column(Text)  # JSON 배열
    expires_at = Column(DateTime)

    created_at = Column(DateTime, default=datetime.utcnow)

    # 관계
    user = relationship("User", back_populates="project_locks")

    def __repr__(self):
        return f"<ProjectLock(project={self.project_code}, user={self.user_id})>"


class SyncLog(Base):
    """Google Sheets 동기화 로그"""
    __tablename__ = "sync_logs"

    id = Column(Integer, primary_key=True, index=True)
    direction = Column(String(20))  # "to_sheets", "from_sheets"
    status = Column(String(20))  # "success", "error", "partial"
    records_processed = Column(Integer, default=0)
    records_successful = Column(Integer, default=0)
    records_failed = Column(Integer, default=0)
    error_message = Column(Text)
    processing_time = Column(Numeric(10, 3))  # 처리 시간 (초)

    created_at = Column(DateTime, default=datetime.utcnow, index=True)

    def __repr__(self):
        return f"<SyncLog(direction={self.direction}, status={self.status})>"