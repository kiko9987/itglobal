"""
감사 로그 서비스
"""

from sqlalchemy.orm import Session
from typing import Optional
from datetime import datetime
import logging

from models import AuditLog

logger = logging.getLogger(__name__)

def log_action(
    db: Session,
    user_id: int,
    project_code: str,
    action: str,
    description: str,
    field_name: Optional[str] = None,
    old_value: Optional[str] = None,
    new_value: Optional[str] = None,
    ip_address: Optional[str] = None,
    user_agent: Optional[str] = None
) -> AuditLog:
    """감사 로그 기록"""
    try:
        audit_log = AuditLog(
            user_id=user_id,
            project_code=project_code,
            action=action,
            field_name=field_name,
            old_value=old_value,
            new_value=new_value,
            description=description,
            ip_address=ip_address,
            user_agent=user_agent
        )

        db.add(audit_log)
        db.commit()
        db.refresh(audit_log)

        logger.info(f"감사 로그 기록: {action} - {project_code} by {user_id}")
        return audit_log

    except Exception as e:
        db.rollback()
        logger.error(f"감사 로그 기록 실패: {e}")
        raise

def get_audit_logs(
    db: Session,
    project_code: Optional[str] = None,
    user_id: Optional[int] = None,
    action: Optional[str] = None,
    limit: int = 100,
    offset: int = 0
):
    """감사 로그 조회"""
    try:
        query = db.query(AuditLog)

        if project_code:
            query = query.filter(AuditLog.project_code == project_code)

        if user_id:
            query = query.filter(AuditLog.user_id == user_id)

        if action:
            query = query.filter(AuditLog.action == action)

        return query.order_by(AuditLog.created_at.desc()).offset(offset).limit(limit).all()

    except Exception as e:
        logger.error(f"감사 로그 조회 실패: {e}")
        raise