"""
프로젝트 관련 Pydantic 스키마
"""

from pydantic import BaseModel, validator
from typing import Optional, List
from decimal import Decimal
from datetime import datetime

class ProjectBase(BaseModel):
    """프로젝트 기본 스키마"""
    project_code: Optional[str] = None
    site_name: Optional[str] = None
    business: Optional[str] = None
    contractor: Optional[str] = None
    site_address: Optional[str] = None
    site_phone: Optional[str] = None
    construction_classification: Optional[str] = None
    manager: Optional[str] = None
    manager_phone: Optional[str] = None
    manager_address: Optional[str] = None
    manager_email: Optional[str] = None
    manager_region: Optional[str] = None
    created_date: Optional[str] = None
    site_manager: Optional[str] = None
    customer_name: Optional[str] = None
    customer_email: Optional[str] = None

    # 금액 정보
    total_amount_1: Optional[Decimal] = None
    tax_amount: Optional[Decimal] = None
    total_amount_2: Optional[Decimal] = None
    contract_deposit: Optional[Decimal] = None
    mid_payment: Optional[Decimal] = None
    final_payment: Optional[Decimal] = None
    outstanding_amount: Optional[Decimal] = None

    # 일정 정보
    construction_confirmed: Optional[str] = None
    construction_date: Optional[str] = None
    construction_confirmed_yn: Optional[str] = None
    installation_brand: Optional[str] = None
    delivery_date: Optional[str] = None
    delivery_note: Optional[str] = None
    installation_note: Optional[str] = None
    completion_status: Optional[str] = None

    # 수금 관련
    collection_notes: Optional[str] = None
    collection_auto_account: Optional[str] = None
    mid_payment_date: Optional[str] = None
    final_payment_date: Optional[str] = None
    total_profit_loss_info: Optional[str] = None
    construction_confirmed_final: Optional[str] = None
    airtable_record_id: Optional[str] = None

class ProjectCreate(ProjectBase):
    """프로젝트 생성 스키마"""
    project_code: str
    site_name: str
    site_address: str

    @validator('project_code')
    def validate_project_code(cls, v):
        if not v or len(v.strip()) == 0:
            raise ValueError('프로젝트 코드는 필수입니다')
        return v.strip()

    @validator('site_name')
    def validate_site_name(cls, v):
        if not v or len(v.strip()) == 0:
            raise ValueError('현장명은 필수입니다')
        return v.strip()

    @validator('site_address')
    def validate_site_address(cls, v):
        if not v or len(v.strip()) == 0:
            raise ValueError('현장 주소는 필수입니다')
        return v.strip()

class ProjectUpdate(ProjectBase):
    """프로젝트 수정 스키마 (모든 필드 선택적)"""
    pass

class ProjectResponse(ProjectBase):
    """프로젝트 응답 스키마"""
    id: int
    is_cancelled: bool
    created_at: datetime
    updated_at: datetime

    # 파생 속성
    status: str
    status_display: str

    class Config:
        from_attributes = True

class ProjectListResponse(BaseModel):
    """프로젝트 목록 응답 스키마"""
    projects: List[ProjectResponse]
    total: int
    page: int
    page_size: int
    total_pages: int

class ProjectStatistics(BaseModel):
    """프로젝트 통계 스키마"""
    total_projects: int
    active_projects: int
    cancelled_projects: int
    total_revenue: float
    active_percentage: float

class ProjectFilter(BaseModel):
    """프로젝트 필터 스키마"""
    search: Optional[str] = None
    status: Optional[str] = None  # "active", "cancelled", "all"
    manager: Optional[str] = None
    business: Optional[str] = None
    page: int = 1
    page_size: int = 20

    @validator('page')
    def validate_page(cls, v):
        if v < 1:
            raise ValueError('페이지 번호는 1 이상이어야 합니다')
        return v

    @validator('page_size')
    def validate_page_size(cls, v):
        if v < 1 or v > 100:
            raise ValueError('페이지 크기는 1-100 사이여야 합니다')
        return v

    @validator('status')
    def validate_status(cls, v):
        if v and v not in ['active', 'cancelled', 'all']:
            raise ValueError('상태는 active, cancelled, all 중 하나여야 합니다')
        return v