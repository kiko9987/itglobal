"""
프로젝트 관련 라우터
"""

from fastapi import APIRouter, Depends, HTTPException, Request, Form, Query
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from typing import Optional
import math

from database import get_db
from models import User
from services.auth_service import get_current_user, get_editor_user
from services.project_service import ProjectService
from schemas.project_schemas import (
    ProjectCreate, ProjectUpdate, ProjectResponse,
    ProjectListResponse, ProjectFilter, ProjectStatistics
)

router = APIRouter()
templates = Jinja2Templates(directory="templates")

# HTML 페이지 라우트
@router.get("/", response_class=HTMLResponse)
async def projects_page(
    request: Request,
    search: Optional[str] = Query(None),
    status: Optional[str] = Query(None),
    manager: Optional[str] = Query(None),
    page: int = Query(1, ge=1),
    page_size: int = Query(20, ge=1, le=100),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """프로젝트 목록 페이지"""
    try:
        # 프로젝트 목록 조회
        skip = (page - 1) * page_size
        projects, total = ProjectService.get_projects(
            db=db,
            skip=skip,
            limit=page_size,
            search=search,
            status=status,
            manager=manager
        )

        # 페이징 정보
        total_pages = math.ceil(total / page_size)

        # 통계 정보
        stats = ProjectService.get_project_statistics(db)

        return templates.TemplateResponse(
            "projects/list.html",
            {
                "request": request,
                "projects": projects,
                "total": total,
                "page": page,
                "page_size": page_size,
                "total_pages": total_pages,
                "search": search or "",
                "status": status or "",
                "manager": manager or "",
                "stats": stats,
                "user": current_user,
                "page_title": "프로젝트 관리"
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"프로젝트 목록 조회 실패: {str(e)}")

@router.get("/{project_code}", response_class=HTMLResponse)
async def project_detail_page(
    project_code: str,
    request: Request,
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """프로젝트 상세 페이지"""
    try:
        project = ProjectService.get_project_by_code(db, project_code)
        if not project:
            raise HTTPException(status_code=404, detail="프로젝트를 찾을 수 없습니다")

        return templates.TemplateResponse(
            "projects/detail.html",
            {
                "request": request,
                "project": project,
                "user": current_user,
                "page_title": f"프로젝트 상세 - {project.project_code}"
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"프로젝트 상세 조회 실패: {str(e)}")

@router.get("/{project_code}/edit", response_class=HTMLResponse)
async def project_edit_page(
    project_code: str,
    request: Request,
    current_user: User = Depends(get_editor_user),
    db: Session = Depends(get_db)
):
    """프로젝트 편집 페이지"""
    try:
        project = ProjectService.get_project_by_code(db, project_code)
        if not project:
            raise HTTPException(status_code=404, detail="프로젝트를 찾을 수 없습니다")

        if project.is_cancelled:
            return RedirectResponse(
                url=f"/projects/{project_code}?error=취소된 프로젝트는 편집할 수 없습니다",
                status_code=302
            )

        return templates.TemplateResponse(
            "projects/edit.html",
            {
                "request": request,
                "project": project,
                "user": current_user,
                "page_title": f"프로젝트 편집 - {project.project_code}"
            }
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"프로젝트 편집 페이지 로드 실패: {str(e)}")

# 프로젝트 액션 라우트
@router.post("/{project_code}/cancel")
async def cancel_project(
    project_code: str,
    current_user: User = Depends(get_editor_user),
    db: Session = Depends(get_db)
):
    """프로젝트 취소"""
    try:
        project = ProjectService.cancel_project(db, project_code, current_user.id)
        if not project:
            return RedirectResponse(
                url="/projects?error=프로젝트를 찾을 수 없습니다",
                status_code=302
            )

        return RedirectResponse(
            url=f"/projects?success=프로젝트 {project_code}가 취소되었습니다",
            status_code=302
        )

    except ValueError as e:
        return RedirectResponse(
            url=f"/projects?error={str(e)}",
            status_code=302
        )
    except Exception as e:
        return RedirectResponse(
            url=f"/projects?error=프로젝트 취소 중 오류가 발생했습니다",
            status_code=302
        )

@router.post("/{project_code}/resume")
async def resume_project(
    project_code: str,
    current_user: User = Depends(get_editor_user),
    db: Session = Depends(get_db)
):
    """프로젝트 재개"""
    try:
        project = ProjectService.resume_project(db, project_code, current_user.id)
        if not project:
            return RedirectResponse(
                url="/projects?error=프로젝트를 찾을 수 없습니다",
                status_code=302
            )

        return RedirectResponse(
            url=f"/projects?success=프로젝트 {project_code}가 재개되었습니다",
            status_code=302
        )

    except ValueError as e:
        return RedirectResponse(
            url=f"/projects?error={str(e)}",
            status_code=302
        )
    except Exception as e:
        return RedirectResponse(
            url=f"/projects?error=프로젝트 재개 중 오류가 발생했습니다",
            status_code=302
        )

@router.post("/{project_code}/update")
async def update_project(
    project_code: str,
    request: Request,
    current_user: User = Depends(get_editor_user),
    db: Session = Depends(get_db)
):
    """프로젝트 업데이트 (폼 데이터)"""
    try:
        # 폼 데이터 파싱
        form_data = await request.form()

        # ProjectUpdate 스키마에 맞게 데이터 변환
        update_data = {}

        # 필드 매핑 (폼 필드명 → 모델 필드명)
        field_mapping = {
            "site_name": "site_name",
            "site_address": "site_address",
            "business": "business",
            "manager": "manager",
            "manager_email": "manager_email",
            "total_amount_1": "total_amount_1",
            "total_amount_2": "total_amount_2",
            "contract_deposit": "contract_deposit",
            "mid_payment": "mid_payment",
            "final_payment": "final_payment",
            "collection_notes": "collection_notes",
            # 필요한 다른 필드들 추가
        }

        for form_field, model_field in field_mapping.items():
            if form_field in form_data:
                value = form_data[form_field].strip() if form_data[form_field] else None
                if value:  # 빈 값이 아닌 경우만
                    update_data[model_field] = value

        # 업데이트 실행
        if update_data:
            project_update = ProjectUpdate(**update_data)
            project = ProjectService.update_project(db, project_code, project_update, current_user.id)

            if not project:
                return RedirectResponse(
                    url="/projects?error=프로젝트를 찾을 수 없습니다",
                    status_code=302
                )

            return RedirectResponse(
                url=f"/projects/{project_code}?success=프로젝트가 업데이트되었습니다",
                status_code=302
            )
        else:
            return RedirectResponse(
                url=f"/projects/{project_code}?info=변경사항이 없습니다",
                status_code=302
            )

    except ValueError as e:
        return RedirectResponse(
            url=f"/projects/{project_code}/edit?error={str(e)}",
            status_code=302
        )
    except Exception as e:
        return RedirectResponse(
            url=f"/projects/{project_code}/edit?error=프로젝트 업데이트 중 오류가 발생했습니다",
            status_code=302
        )

# API 엔드포인트
@router.get("/api/list", response_model=ProjectListResponse)
async def get_projects_api(
    filter_params: ProjectFilter = Depends(),
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """프로젝트 목록 API"""
    try:
        skip = (filter_params.page - 1) * filter_params.page_size
        projects, total = ProjectService.get_projects(
            db=db,
            skip=skip,
            limit=filter_params.page_size,
            search=filter_params.search,
            status=filter_params.status,
            manager=filter_params.manager
        )

        total_pages = math.ceil(total / filter_params.page_size)

        return ProjectListResponse(
            projects=projects,
            total=total,
            page=filter_params.page,
            page_size=filter_params.page_size,
            total_pages=total_pages
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"프로젝트 목록 조회 실패: {str(e)}")

@router.get("/api/stats", response_model=ProjectStatistics)
async def get_project_statistics_api(
    current_user: User = Depends(get_current_user),
    db: Session = Depends(get_db)
):
    """프로젝트 통계 API"""
    try:
        return ProjectService.get_project_statistics(db)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"프로젝트 통계 조회 실패: {str(e)}")