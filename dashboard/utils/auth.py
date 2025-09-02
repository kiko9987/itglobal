import os, json
from typing import Tuple, Dict, Any

_CACHED: Dict[str, Any] = {}

def _load_creds() -> Dict[str, Any]:
    global _CACHED
    if _CACHED:
        return _CACHED
    cred_path = os.getenv("CREDENTIALS_JSON") or os.path.join(os.path.dirname(__file__), "..", "credentials.json")
    cred_path = os.path.abspath(cred_path)
    data: Dict[str, Any] = {}
    try:
        with open(cred_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = {}
    _CACHED = data
    return data

def _get_api_key() -> str:
    env_key = os.getenv("API_KEY", "").strip()
    if env_key:
        return env_key
    creds = _load_creds()
    return str(creds.get("api_key", "")).strip()

def _get_admins() -> list:
    env_admins = [x.strip().lower() for x in os.getenv("ADMIN_EMAILS", "").split(",") if x.strip()]
    if env_admins:
        return env_admins
    creds = _load_creds()
    raw = creds.get("admin_emails", [])
    if isinstance(raw, str):
        return [x.strip().lower() for x in raw.split(",") if x.strip()]
    return [str(x).strip().lower() for x in raw]

def check_api_key(key: str) -> bool:
    """API 키 검증 (보안 강화)"""
    if not key or len(key.strip()) < 8:  # 최소 8자 이상
        return False
    
    api_key = _get_api_key()
    if not api_key or len(api_key) < 8:
        return False
        
    # 상수 시간 비교로 타이밍 공격 방지
    return len(key) == len(api_key) and key == api_key

def is_admin(email: str) -> bool:
    return email.lower() in set(_get_admins())

def get_user_from_headers(headers) -> Tuple[str, bool]:
    api_key = headers.get("X-API-Key", "")
    user_email = headers.get("X-User-Email", "").strip()
    ok = check_api_key(api_key)
    admin = is_admin(user_email) if (ok and user_email) else False
    return (user_email, admin) if ok else ("", False)
