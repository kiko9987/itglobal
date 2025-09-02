import logging
from typing import Dict, List, Optional
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

class GoogleSheetsManager:
    SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

    def __init__(self, sheet_id: str, range_a1: str, service_account_json: str):
        self.sheet_id = sheet_id
        self.range_a1 = range_a1
        self.service_account_json = service_account_json
        self._service = self._build_service()
        self._headers: List[str] = []
        self._last_df: Optional[pd.DataFrame] = None

    def _build_service(self):
        if not self.service_account_json:
            logger.warning("Google API 키가 설정되지 않았습니다. 더미 모드로 실행됩니다.")
            return None
        try:
            creds = Credentials.from_service_account_file(
                self.service_account_json, scopes=self.SCOPES
            )
            return build("sheets", "v4", credentials=creds, cache_discovery=False)
        except Exception as e:
            logger.warning(f"Google Sheets 연결 실패: {e}. 더미 모드로 실행됩니다.")
            return None

    def fetch_dataframe(self) -> pd.DataFrame:
        if not self._service:
            # 더미 데이터 반환
            dummy_data = {
                '프로젝트 코드': ['G0001-YG', 'P0002-JW', 'R0003-SH'],
                '현장 주소': ['서울시 강남구', '부산시 해운대구', '대구시 수성구'],
                '사업자': ['글로벌', '플렌트', '글로벌그룹'],
                '담당자': ['박용구', '박정우', '강성환'],
                '등록일': ['2025-01-15', '2025-01-16', '2025-01-17']
            }
            df = pd.DataFrame(dummy_data)
            self._headers = list(df.columns)
            self._last_df = df
            return df
            
        values = self._service.spreadsheets().values().get(
            spreadsheetId=self.sheet_id, range=self.range_a1
        ).execute().get("values", [])

        if not values:
            return pd.DataFrame()

        self._headers = values[0]
        rows = values[1:]
        df = pd.DataFrame(rows, columns=self._headers).fillna("").replace({None: ""})
        self._last_df = df
        return df

    def append_row(self, record: Dict[str, str]) -> None:
        if not self._service:
            logger.warning("더미 모드에서는 데이터 추가가 지원되지 않습니다.")
            return
            
        if not self._headers:
            self.fetch_dataframe()
        row = [record.get(h, "") for h in self._headers]
        self._service.spreadsheets().values().append(
            spreadsheetId=self.sheet_id,
            range=self.range_a1.split("!")[0],
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body={"values": [row]},
        ).execute()
