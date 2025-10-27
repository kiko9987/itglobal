"""
관리 기능 라우터
"""

from fastapi import APIRouter, Depends, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session

from database import get_db, check_database_health
from models import User
from services.auth_service import get_admin_user
from services.audit_service import get_audit_logs

router = APIRouter()
templates = Jinja2Templates(directory="templates")

@router.get("/", response_class=HTMLResponse)
async def admin_dashboard(
    request: Request,
    current_user: User = Depends(get_admin_user)
):
    """관리자 대시보드"""
    return templates.TemplateResponse(
        "admin/dashboard.html",
        {
            "request": request,
            "user": current_user,
            "page_title": "관리자 대시보드"
        }
    )

@router.get("/health", response_class=HTMLResponse)
async def health_check_page(
    request: Request,
    current_user: User = Depends(get_admin_user)
):
    """시스템 상태 확인 페이지"""
    db_health = check_database_health()

    return templates.TemplateResponse(
        "admin/health.html",
        {
            "request": request,
            "user": current_user,
            "db_health": db_health,
            "page_title": "시스템 상태"
        }
    )

@router.get("/audit-logs", response_class=HTMLResponse)
async def audit_logs_page(
    request: Request,
    current_user: User = Depends(get_admin_user),
    db: Session = Depends(get_db)
):
    """감사 로그 페이지"""
    # 최근 100개 로그 조회
    logs = get_audit_logs(db, limit=100)

    return templates.TemplateResponse(
        "admin/audit_logs.html",
        {
            "request": request,
            "user": current_user,
            "logs": logs,
            "page_title": "감사 로그"
        }
    )