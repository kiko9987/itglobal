"""
ITG 대시보드 FastAPI 애플리케이션
냉난방기 설치 공사 관리 시스템
"""

from fastapi import FastAPI, Request, Depends, HTTPException, status
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, RedirectResponse
from contextlib import asynccontextmanager
import os
from pathlib import Path

# 로컬 모듈
from database import engine, create_tables
from routers import projects, auth, admin
from services.auth_service import get_current_user
from config import settings

# 애플리케이션 수명주기 관리
@asynccontextmanager
async def lifespan(app: FastAPI):
    """애플리케이션 시작/종료 시 실행되는 함수"""
    # 시작 시
    print("🚀 ITG 대시보드 FastAPI 애플리케이션 시작")

    # 데이터베이스 테이블 생성
    create_tables()
    print("✅ 데이터베이스 테이블 초기화 완료")

    yield

    # 종료 시
    print("🔄 ITG 대시보드 애플리케이션 종료")

# FastAPI 애플리케이션 생성
app = FastAPI(
    title="ITG 대시보드",
    description="냉난방기 설치 공사 관리 시스템",
    version="2.0.0",
    lifespan=lifespan
)

# 정적 파일 제공
app.mount("/static", StaticFiles(directory="static"), name="static")

# 템플릿 설정
templates = Jinja2Templates(directory="templates")

# 라우터 등록
app.include_router(auth.router, prefix="/auth", tags=["인증"])
app.include_router(projects.router, prefix="/projects", tags=["프로젝트"])
app.include_router(admin.router, prefix="/admin", tags=["관리"])

# 홈페이지 (프로젝트 목록으로 리다이렉트)
@app.get("/", response_class=RedirectResponse)
async def root():
    """홈페이지 - 프로젝트 목록으로 리다이렉트"""
    return RedirectResponse(url="/projects", status_code=302)

# 대시보드 메인 페이지
@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(
    request: Request,
    current_user: dict = Depends(get_current_user)
):
    """대시보드 메인 페이지"""
    return templates.TemplateResponse(
        "dashboard.html",
        {
            "request": request,
            "user": current_user,
            "page_title": "대시보드"
        }
    )

# 헬스체크 엔드포인트
@app.get("/health")
async def health_check():
    """헬스체크 엔드포인트"""
    return {
        "status": "healthy",
        "app": "ITG Dashboard",
        "version": "2.0.0"
    }

# 에러 핸들러
@app.exception_handler(404)
async def not_found_handler(request: Request, exc: HTTPException):
    """404 에러 핸들러"""
    return templates.TemplateResponse(
        "error.html",
        {
            "request": request,
            "error_code": 404,
            "error_message": "페이지를 찾을 수 없습니다."
        },
        status_code=404
    )

@app.exception_handler(500)
async def internal_error_handler(request: Request, exc: HTTPException):
    """500 에러 핸들러"""
    return templates.TemplateResponse(
        "error.html",
        {
            "request": request,
            "error_code": 500,
            "error_message": "서버 내부 오류가 발생했습니다."
        },
        status_code=500
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        log_level="info"
    )