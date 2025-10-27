"""
인증 관련 라우터
"""

from fastapi import APIRouter, Depends, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session

from database import get_db
from services.auth_service import get_current_user_optional

router = APIRouter()
templates = Jinja2Templates(directory="templates")

@router.get("/login", response_class=HTMLResponse)
async def login_page(
    request: Request,
    message: str = None,
    current_user = Depends(get_current_user_optional)
):
    """로그인 페이지"""
    # 이미 로그인된 경우 프로젝트 페이지로 리다이렉트
    if current_user:
        return RedirectResponse(url="/projects", status_code=302)

    return templates.TemplateResponse(
        "auth/login.html",
        {
            "request": request,
            "message": message,
            "page_title": "로그인"
        }
    )

@router.get("/logout")
async def logout():
    """로그아웃"""
    # 쿠키 삭제를 위한 응답
    response = RedirectResponse(url="/auth/login?message=logout_success", status_code=302)
    response.delete_cookie("access_token")
    return response

# 개발용 로그인 (운영에서는 제거)
@router.post("/dev-login")
async def dev_login(
    request: Request,
    db: Session = Depends(get_db)
):
    """개발용 로그인 (OAuth 구현 전 임시)"""
    from services.auth_service import AuthService
    from config import settings

    if not settings.debug:
        raise HTTPException(status_code=404, detail="Not found")

    # 첫 번째 관리자로 로그인
    admin_email = settings.admin_emails[0] if settings.admin_emails else "admin@example.com"

    # 토큰 생성
    token = AuthService.create_access_token({"email": admin_email})

    # 프로젝트 페이지로 리다이렉트하며 토큰을 쿠키에 설정
    response = RedirectResponse(url="/projects", status_code=302)
    response.set_cookie(
        key="access_token",
        value=f"Bearer {token}",
        httponly=True,
        max_age=settings.access_token_expire_minutes * 60
    )

    return response