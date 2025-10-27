"""
Google Sheets 동기화 서비스
"""

from sqlalchemy.orm import Session
from typing import List, Dict, Any, Optional
from datetime import datetime
import logging
import json

from models import Project, SyncLog
from services.project_service import ProjectService
from schemas.project_schemas import ProjectCreate

logger = logging.getLogger(__name__)

class SheetsSyncService:
    """Google Sheets 동기화 서비스"""

    # Google Sheets 컬럼 매핑 (0-based index)
    COLUMN_MAPPING = {
        0: "project_code",          # 시공번호
        1: "creation_date",         # 시공일자
        2: "site_name",             # 현장명
        3: "site_address",          # 현장주소
        4: "business",              # 공사업종
        5: "manager",               # 현장담당자
        6: "manager_email",         # 담당자 이메일
        7: "total_amount_1",        # 총공사금액1
        8: "total_amount_2",        # 총공사금액2
        9: "contract_deposit",      # 계약금
        10: "mid_payment",          # 중도금
        11: "final_payment",        # 잔금
        12: "collection_status_1",  # 수금현황1
        13: "collection_status_2",  # 수금현황2
        14: "collection_status_3",  # 수금현황3
        15: "collection_status_4",  # 수금현황4
        16: "collection_status_5",  # 수금현황5
        17: "collection_status_6",  # 수금현황6
        18: "collection_notes",     # 수금비고
        # 추가 컬럼들은 필요에 따라 매핑
    }

    @staticmethod
    def parse_sheet_row(row: List[str]) -> Dict[str, Any]:
        """시트 행 데이터를 프로젝트 데이터로 변환"""
        try:
            project_data = {}

            for col_index, field_name in SheetsSyncService.COLUMN_MAPPING.items():
                if col_index < len(row):
                    value = row[col_index].strip() if row[col_index] else None

                    # 특별한 변환이 필요한 필드들
                    if field_name == "creation_date" and value:
                        try:
                            # 날짜 형식 변환 (YYYY-MM-DD 형식으로 변환)
                            project_data[field_name] = datetime.strptime(value, "%Y-%m-%d").date()
                        except ValueError:
                            logger.warning(f"날짜 형식 오류: {value}")
                            project_data[field_name] = None

                    elif field_name in ["total_amount_1", "total_amount_2", "contract_deposit", "mid_payment", "final_payment"]:
                        # 금액 필드 처리
                        if value:
                            try:
                                # 쉼표 제거 후 숫자로 변환
                                project_data[field_name] = int(value.replace(",", "").replace("원", ""))
                            except ValueError:
                                logger.warning(f"금액 형식 오류: {value}")
                                project_data[field_name] = 0
                        else:
                            project_data[field_name] = 0

                    else:
                        project_data[field_name] = value

            # 취소 상태 확인 (수금비고에 "공사취소"가 있는 경우)
            collection_notes = project_data.get("collection_notes", "")
            project_data["is_cancelled"] = "공사취소" in collection_notes if collection_notes else False

            return project_data

        except Exception as e:
            logger.error(f"시트 행 파싱 실패: {e}")
            return {}

    @staticmethod
    def sync_from_sheets(db: Session, sheet_data: List[List[str]], sync_user_id: int) -> Dict[str, int]:
        """Google Sheets 데이터를 데이터베이스로 동기화"""
        try:
            stats = {
                "total_rows": len(sheet_data),
                "created": 0,
                "updated": 0,
                "errors": 0,
                "skipped": 0
            }

            for row_index, row in enumerate(sheet_data):
                try:
                    # 빈 행 건너뛰기
                    if not row or not any(row):
                        stats["skipped"] += 1
                        continue

                    # 헤더 행 건너뛰기
                    if row_index == 0 and row[0] in ["시공번호", "project_code"]:
                        stats["skipped"] += 1
                        continue

                    # 행 데이터 파싱
                    project_data = SheetsSyncService.parse_sheet_row(row)

                    if not project_data.get("project_code"):
                        logger.warning(f"프로젝트 코드 없음 (행 {row_index + 1})")
                        stats["errors"] += 1
                        continue

                    project_code = project_data["project_code"]

                    # 기존 프로젝트 확인
                    existing_project = ProjectService.get_project_by_code(db, project_code)

                    if existing_project:
                        # 기존 프로젝트 업데이트
                        updated = False

                        for field, value in project_data.items():
                            if field != "project_code" and hasattr(existing_project, field):
                                current_value = getattr(existing_project, field)
                                if current_value != value:
                                    setattr(existing_project, field, value)
                                    updated = True

                        if updated:
                            existing_project.updated_at = datetime.utcnow()
                            db.commit()
                            stats["updated"] += 1
                            logger.info(f"프로젝트 업데이트: {project_code}")
                        else:
                            stats["skipped"] += 1

                    else:
                        # 새 프로젝트 생성
                        try:
                            project_create = ProjectCreate(**project_data)
                            new_project = ProjectService.create_project(db, project_create, sync_user_id)
                            stats["created"] += 1
                            logger.info(f"새 프로젝트 생성: {project_code}")

                        except Exception as e:
                            logger.error(f"프로젝트 생성 실패 ({project_code}): {e}")
                            stats["errors"] += 1

                except Exception as e:
                    logger.error(f"행 처리 실패 (행 {row_index + 1}): {e}")
                    stats["errors"] += 1

            # 동기화 로그 기록
            SheetsSyncService.log_sync_result(db, sync_user_id, stats)

            logger.info(f"동기화 완료: {stats}")
            return stats

        except Exception as e:
            logger.error(f"시트 동기화 실패: {e}")
            raise

    @staticmethod
    def export_to_sheets_format(db: Session) -> List[List[str]]:
        """데이터베이스 데이터를 Google Sheets 형식으로 내보내기"""
        try:
            projects = db.query(Project).order_by(Project.created_at.desc()).all()

            # 헤더 행
            headers = [
                "시공번호", "시공일자", "현장명", "현장주소", "공사업종", "현장담당자", "담당자이메일",
                "총공사금액1", "총공사금액2", "계약금", "중도금", "잔금",
                "수금현황1", "수금현황2", "수금현황3", "수금현황4", "수금현황5", "수금현황6",
                "수금비고"
            ]

            result = [headers]

            for project in projects:
                row = [
                    project.project_code or "",
                    project.creation_date.strftime("%Y-%m-%d") if project.creation_date else "",
                    project.site_name or "",
                    project.site_address or "",
                    project.business or "",
                    project.manager or "",
                    project.manager_email or "",
                    str(project.total_amount_1) if project.total_amount_1 else "0",
                    str(project.total_amount_2) if project.total_amount_2 else "0",
                    str(project.contract_deposit) if project.contract_deposit else "0",
                    str(project.mid_payment) if project.mid_payment else "0",
                    str(project.final_payment) if project.final_payment else "0",
                    project.collection_status_1 or "",
                    project.collection_status_2 or "",
                    project.collection_status_3 or "",
                    project.collection_status_4 or "",
                    project.collection_status_5 or "",
                    project.collection_status_6 or "",
                    project.collection_notes or ""
                ]
                result.append(row)

            return result

        except Exception as e:
            logger.error(f"시트 형식 내보내기 실패: {e}")
            raise

    @staticmethod
    def log_sync_result(db: Session, user_id: int, stats: Dict[str, int]) -> SyncLog:
        """동기화 결과 로그 기록"""
        try:
            sync_log = SyncLog(
                user_id=user_id,
                sync_type="import",
                total_records=stats["total_rows"],
                created_records=stats["created"],
                updated_records=stats["updated"],
                error_records=stats["errors"],
                details=json.dumps(stats, ensure_ascii=False)
            )

            db.add(sync_log)
            db.commit()
            db.refresh(sync_log)

            return sync_log

        except Exception as e:
            logger.error(f"동기화 로그 기록 실패: {e}")
            return None

    @staticmethod
    def get_sync_history(db: Session, limit: int = 50) -> List[SyncLog]:
        """동기화 히스토리 조회"""
        try:
            return db.query(SyncLog).order_by(SyncLog.created_at.desc()).limit(limit).all()
        except Exception as e:
            logger.error(f"동기화 히스토리 조회 실패: {e}")
            return []

    @staticmethod
    def validate_sheet_data(sheet_data: List[List[str]]) -> Dict[str, Any]:
        """시트 데이터 유효성 검사"""
        try:
            validation_result = {
                "is_valid": True,
                "errors": [],
                "warnings": [],
                "total_rows": len(sheet_data),
                "valid_rows": 0,
                "invalid_rows": 0
            }

            for row_index, row in enumerate(sheet_data):
                row_errors = []

                # 빈 행 확인
                if not row or not any(row):
                    continue

                # 헤더 행 건너뛰기
                if row_index == 0:
                    continue

                # 프로젝트 코드 필수 확인
                if not row[0] or not row[0].strip():
                    row_errors.append("프로젝트 코드가 비어있습니다")

                # 날짜 형식 확인
                if len(row) > 1 and row[1]:
                    try:
                        datetime.strptime(row[1], "%Y-%m-%d")
                    except ValueError:
                        row_errors.append("날짜 형식이 올바르지 않습니다 (YYYY-MM-DD 형식 필요)")

                # 금액 필드 확인
                for col_index in [7, 8, 9, 10, 11]:  # 금액 관련 컬럼들
                    if len(row) > col_index and row[col_index]:
                        try:
                            int(row[col_index].replace(",", "").replace("원", ""))
                        except ValueError:
                            row_errors.append(f"컬럼 {col_index + 1}: 금액 형식이 올바르지 않습니다")

                if row_errors:
                    validation_result["invalid_rows"] += 1
                    validation_result["errors"].append({
                        "row": row_index + 1,
                        "errors": row_errors
                    })
                else:
                    validation_result["valid_rows"] += 1

            if validation_result["invalid_rows"] > 0:
                validation_result["is_valid"] = False

            return validation_result

        except Exception as e:
            logger.error(f"시트 데이터 유효성 검사 실패: {e}")
            return {
                "is_valid": False,
                "errors": [f"유효성 검사 중 오류 발생: {str(e)}"],
                "warnings": [],
                "total_rows": 0,
                "valid_rows": 0,
                "invalid_rows": 0
            }