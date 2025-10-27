"""
프로젝트 관련 비즈니스 로직
"""

from sqlalchemy.orm import Session
from sqlalchemy import and_, or_, desc
from typing import List, Optional, Dict, Any
from datetime import datetime
import logging

from models import Project, AuditLog
from schemas.project_schemas import ProjectCreate, ProjectUpdate, ProjectResponse
from services.audit_service import log_action

logger = logging.getLogger(__name__)

class ProjectService:
    """프로젝트 관련 서비스"""

    @staticmethod
    def get_projects(
        db: Session,
        skip: int = 0,
        limit: int = 20,
        search: Optional[str] = None,
        status: Optional[str] = None,
        manager: Optional[str] = None
    ) -> tuple[List[Project], int]:
        """프로젝트 목록 조회 (페이징, 필터링)"""
        try:
            query = db.query(Project)

            # 필터링
            if search:
                search_filter = or_(
                    Project.project_code.ilike(f"%{search}%"),
                    Project.site_name.ilike(f"%{search}%"),
                    Project.site_address.ilike(f"%{search}%"),
                    Project.business.ilike(f"%{search}%")
                )
                query = query.filter(search_filter)

            if status:
                if status == "cancelled":
                    query = query.filter(Project.is_cancelled == True)
                elif status == "active":
                    query = query.filter(Project.is_cancelled == False)

            if manager:
                query = query.filter(Project.manager.ilike(f"%{manager}%"))

            # 총 개수
            total = query.count()

            # 정렬 및 페이징
            projects = query.order_by(desc(Project.created_at)).offset(skip).limit(limit).all()

            return projects, total

        except Exception as e:
            logger.error(f"프로젝트 목록 조회 실패: {e}")
            raise

    @staticmethod
    def get_project_by_code(db: Session, project_code: str) -> Optional[Project]:
        """프로젝트 코드로 프로젝트 조회"""
        try:
            return db.query(Project).filter(Project.project_code == project_code).first()
        except Exception as e:
            logger.error(f"프로젝트 조회 실패 ({project_code}): {e}")
            raise

    @staticmethod
    def create_project(db: Session, project_data: ProjectCreate, user_id: int) -> Project:
        """새 프로젝트 생성"""
        try:
            # 프로젝트 코드 중복 확인
            existing = db.query(Project).filter(Project.project_code == project_data.project_code).first()
            if existing:
                raise ValueError(f"프로젝트 코드가 이미 존재합니다: {project_data.project_code}")

            # 새 프로젝트 생성
            project = Project(**project_data.dict())
            db.add(project)
            db.commit()
            db.refresh(project)

            # 감사 로그 기록
            log_action(
                db=db,
                user_id=user_id,
                project_code=project.project_code,
                action="CREATE_PROJECT",
                description=f"새 프로젝트 생성: {project.site_name}"
            )

            logger.info(f"프로젝트 생성 완료: {project.project_code}")
            return project

        except Exception as e:
            db.rollback()
            logger.error(f"프로젝트 생성 실패: {e}")
            raise

    @staticmethod
    def update_project(
        db: Session,
        project_code: str,
        project_data: ProjectUpdate,
        user_id: int
    ) -> Optional[Project]:
        """프로젝트 업데이트"""
        try:
            project = db.query(Project).filter(Project.project_code == project_code).first()
            if not project:
                return None

            # 취소된 프로젝트는 수정 불가
            if project.is_cancelled:
                raise ValueError("취소된 프로젝트는 수정할 수 없습니다")

            # 변경사항 추적
            changes = []
            update_data = project_data.dict(exclude_unset=True)

            for field, new_value in update_data.items():
                if hasattr(project, field):
                    old_value = getattr(project, field)
                    if old_value != new_value:
                        changes.append({
                            "field": field,
                            "old_value": str(old_value) if old_value else "",
                            "new_value": str(new_value) if new_value else ""
                        })
                        setattr(project, field, new_value)

            # 업데이트 시간 갱신
            project.updated_at = datetime.utcnow()

            if changes:
                db.commit()
                db.refresh(project)

                # 감사 로그 기록
                for change in changes:
                    log_action(
                        db=db,
                        user_id=user_id,
                        project_code=project.project_code,
                        action="UPDATE_PROJECT",
                        field_name=change["field"],
                        old_value=change["old_value"],
                        new_value=change["new_value"],
                        description=f"프로젝트 필드 수정: {change['field']}"
                    )

                logger.info(f"프로젝트 업데이트 완료: {project.project_code} ({len(changes)}개 필드)")

            return project

        except Exception as e:
            db.rollback()
            logger.error(f"프로젝트 업데이트 실패 ({project_code}): {e}")
            raise

    @staticmethod
    def cancel_project(db: Session, project_code: str, user_id: int) -> Optional[Project]:
        """프로젝트 취소"""
        try:
            project = db.query(Project).filter(Project.project_code == project_code).first()
            if not project:
                return None

            if project.is_cancelled:
                raise ValueError("이미 취소된 프로젝트입니다")

            # 취소 처리
            old_status = project.collection_notes or ""
            project.is_cancelled = True
            project.collection_notes = "공사취소"
            project.updated_at = datetime.utcnow()

            db.commit()
            db.refresh(project)

            # 감사 로그 기록
            log_action(
                db=db,
                user_id=user_id,
                project_code=project.project_code,
                action="CANCEL_PROJECT",
                field_name="수금 관련 특이사항",
                old_value=old_status,
                new_value="공사취소",
                description=f"프로젝트 취소: {project.site_name}"
            )

            logger.info(f"프로젝트 취소 완료: {project.project_code}")
            return project

        except Exception as e:
            db.rollback()
            logger.error(f"프로젝트 취소 실패 ({project_code}): {e}")
            raise

    @staticmethod
    def resume_project(db: Session, project_code: str, user_id: int) -> Optional[Project]:
        """프로젝트 재개"""
        try:
            project = db.query(Project).filter(Project.project_code == project_code).first()
            if not project:
                return None

            if not project.is_cancelled:
                raise ValueError("취소되지 않은 프로젝트입니다")

            # 재개 처리
            project.is_cancelled = False
            project.collection_notes = "-"  # 또는 빈 값
            project.updated_at = datetime.utcnow()

            db.commit()
            db.refresh(project)

            # 감사 로그 기록
            log_action(
                db=db,
                user_id=user_id,
                project_code=project.project_code,
                action="RESUME_PROJECT",
                field_name="수금 관련 특이사항",
                old_value="공사취소",
                new_value="-",
                description=f"프로젝트 재개: {project.site_name}"
            )

            logger.info(f"프로젝트 재개 완료: {project.project_code}")
            return project

        except Exception as e:
            db.rollback()
            logger.error(f"프로젝트 재개 실패 ({project_code}): {e}")
            raise

    @staticmethod
    def get_project_statistics(db: Session) -> Dict[str, Any]:
        """프로젝트 통계 정보"""
        try:
            total_projects = db.query(Project).count()
            active_projects = db.query(Project).filter(Project.is_cancelled == False).count()
            cancelled_projects = db.query(Project).filter(Project.is_cancelled == True).count()

            # 총 금액 계산
            total_amount = db.query(Project.total_amount_2).filter(
                and_(Project.total_amount_2.isnot(None), Project.is_cancelled == False)
            ).all()

            total_revenue = sum(amount[0] for amount in total_amount if amount[0])

            return {
                "total_projects": total_projects,
                "active_projects": active_projects,
                "cancelled_projects": cancelled_projects,
                "total_revenue": float(total_revenue) if total_revenue else 0.0,
                "active_percentage": round((active_projects / total_projects * 100), 1) if total_projects > 0 else 0
            }

        except Exception as e:
            logger.error(f"프로젝트 통계 조회 실패: {e}")
            return {
                "total_projects": 0,
                "active_projects": 0,
                "cancelled_projects": 0,
                "total_revenue": 0.0,
                "active_percentage": 0.0
            }