"""
애플리케이션 설정
"""

from pydantic_settings import BaseSettings
from typing import List, Dict
import os
from pathlib import Path

class Settings(BaseSettings):
    """애플리케이션 설정 클래스"""

    # 기본 설정
    app_name: str = "ITG 대시보드"
    app_version: str = "2.0.0"
    debug: bool = True

    # 데이터베이스 설정
    database_url: str = "sqlite:///./itg_dashboard.db"

    # 보안 설정
    secret_key: str = "your-secret-key-here-change-in-production"
    access_token_expire_minutes: int = 480  # 8시간

    # Google Sheets 설정
    google_credentials_file: str = "credentials.json"
    google_sheet_id: str = ""
    google_sheet_range: str = "공사 현황의 사본!A1:AM5000"

    # 관리자 설정
    admin_emails: List[str] = [
        "jw@itg-aircon.com",
        "yg@itg-aircon.com",
        "sb@itg-aircon.com",
        "kiko@itg-aircon.com"
    ]

    # 사용자 별칭 매핑
    user_alias_map: Dict[str, str] = {
        "jw@itg-aircon.com": "박정우",
        "yg@itg-aircon.com": "박용구",
        "sb@itg-aircon.com": "황샛별",
        "kiko@itg-aircon.com": "고광일"
    }

    # 회사 코드 매핑
    company_prefix_map: Dict[str, str] = {
        "글로벌": "G",
        "글로벌그룹": "R",
        "플렌트": "P",
        "아이티": "I"
    }

    # 담당자 접미사 매핑
    owner_suffix_map: Dict[str, str] = {
        "박정우": "JW",
        "강성환": "SH",
        "박용구": "YG",
        "박민우": "MW",
        "이근혁": "GH",
        "아이티": "IT",
        "김단이": "DN",
        "권태훈": "TH",
        "주영민": "YM",
        "빈승정": "SJ",
        "심장원": "SJW",
        "박민재": "MJ",
        "조성헌": "JSH",
        "황해승": "HS",
        "강민석": "MS",
        "황샛별": "SB"
    }

    # 필수 필드
    required_fields: List[str] = [
        "프로젝트 코드",
        "현장 주소"
    ]

    # 페이징 설정
    default_page_size: int = 20
    max_page_size: int = 100

    class Config:
        env_file = ".env"
        env_file_encoding = "utf-8"

    def __init__(self, **kwargs):
        super().__init__(**kwargs)

        # 환경변수에서 Google Sheets 설정 로드
        self.google_sheet_id = os.getenv("GOOGLE_SHEET_ID", "")

        # 데이터베이스 디렉토리 생성
        db_dir = Path(self.database_url.replace("sqlite:///", "")).parent
        db_dir.mkdir(exist_ok=True)

# 전역 설정 인스턴스
settings = Settings()