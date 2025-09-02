import os
from typing import List, Dict, Any
import pandas as pd
from datetime import datetime

def _parse_date(s: str):
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            continue
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None

class DataAnalyzer:
    def __init__(self, df: pd.DataFrame):
        self.df = df.copy()
        self.required_fields = [x.strip() for x in os.getenv("REQUIRED_FIELDS", "").split(",") if x.strip()]
        self.confirm_col = os.getenv("CONFIRM_COLUMN", "공사 확정")
        self.deposit_date_col = os.getenv("DEPOSIT_DATE_COLUMN", "계약금 입금일")
        self.created_col = os.getenv("CREATED_DATE_COLUMN", "등록일")
        self.overdue_days = int(os.getenv("OVERDUE_CONFIRM_DAYS", "2"))
        self.enable_overdue = os.getenv("ENABLE_OVERDUE_RULE", "").lower() in ("1","true","y","yes")

    def missing_fields(self) -> List[Dict[str, Any]]:
        if not self.required_fields:
            return []
        out: List[Dict[str, Any]] = []
        for _, row in self.df.iterrows():
            miss = []
            for col in self.required_fields:
                if col not in self.df.columns or str(row.get(col, "")).strip() == "":
                    miss.append(col)
            if miss:
                d = row.to_dict()
                d["_missing_fields"] = miss
                out.append(d)
        return out
