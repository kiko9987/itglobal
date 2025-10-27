"""
Google OAuth 구현
@itg-aircon.com 도메인 제한 포함
"""

import os
import json
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from flask import current_app
import logging

logger = logging.getLogger(__name__)

class GoogleOAuthManager:
    def __init__(self):
        self.client_config = self._get_oauth_config()
        # Config에서 allowed_domain 가져오기 (하드코딩 제거)
        self.allowed_domain = self._get_config_value('GOOGLE_OAUTH_ALLOWED_DOMAIN', 'itg-aircon.com')

    def _get_config_value(self, key, default=None):
        """Flask config나 환경 변수에서 값 가져오기"""
        try:
            # Flask app context가 있으면 config에서 가져오기
            return current_app.config.get(key, default)
        except RuntimeError:
            # App context가 없으면 환경 변수 또는 기본값
            return os.getenv(key, default)

    def _get_config_list(self, key, default=None):
        """
        리스트 값을 가져오기 (환경 변수면 콤마로 split)

        Args:
            key: 설정 키
            default: 기본값 (리스트 또는 None)

        Returns:
            list: 설정 값 (문자열이면 콤마로 split하여 리스트로 변환)

        Note:
            - Flask config에서는 리스트로 저장되어 있음
            - 환경 변수에서는 "url1,url2,url3" 형태의 문자열
            - 타입 안전성 보장
        """
        value = self._get_config_value(key, default)

        # 문자열이면 콤마로 split하여 리스트로 변환
        if isinstance(value, str):
            return [item.strip() for item in value.split(',') if item.strip()]

        # 이미 리스트면 그대로 반환
        if isinstance(value, list):
            return value

        # None이거나 다른 타입이면 기본값 반환
        return default if default is not None else []

    def _get_oauth_config(self):
        """OAuth 클라이언트 설정 로드"""
        # 환경 변수에서 OAuth 설정을 가져오거나 파일에서 로드
        client_id = os.getenv('GOOGLE_OAUTH_CLIENT_ID')
        client_secret = os.getenv('GOOGLE_OAUTH_CLIENT_SECRET')

        if client_id and client_secret:
            # Config에서 redirect_uris 가져오기 (리스트 타입 안전 보장)
            redirect_uris = self._get_config_list(
                'GOOGLE_OAUTH_REDIRECT_URIS',
                ["http://localhost:5000/auth/callback", "http://localhost:5001/auth/callback", "http://localhost:5004/auth/callback"]
            )

            return {
                "web": {
                    "client_id": client_id,
                    "client_secret": client_secret,
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "redirect_uris": redirect_uris
                }
            }

        # 파일에서 로드 시도
        try:
            oauth_file = os.path.join(os.path.dirname(__file__), 'google_oauth_credentials.json')
            with open(oauth_file, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            logger.warning("Google OAuth 설정을 찾을 수 없습니다. 환경 변수나 google_oauth_credentials.json 파일을 확인하세요.")
            return None

    def create_flow(self, redirect_uri=None):
        """OAuth flow 생성"""
        if not self.client_config:
            raise ValueError("Google OAuth 설정이 없습니다.")

        flow = Flow.from_client_config(
            self.client_config,
            scopes=[
                'https://www.googleapis.com/auth/userinfo.email',
                'https://www.googleapis.com/auth/userinfo.profile',
                'openid',
                'https://www.googleapis.com/auth/calendar.events'  # 캘린더 API 스코프 추가
            ]
        )

        # Config에서 기본 redirect_uri 가져오기 (하드코딩 제거)
        default_redirect_uri = self._get_config_value(
            'GOOGLE_OAUTH_DEFAULT_REDIRECT_URI',
            'http://localhost:5004/auth/callback'
        )
        flow.redirect_uri = redirect_uri or default_redirect_uri
        return flow

    def _has_valid_refresh_token(self) -> bool:
        """토큰 파일에 유효한 refresh_token이 있는지 확인하고, 만료 시 갱신"""
        try:
            # 토큰 파일 경로 (instance/google_calendar_token.json)
            token_path = os.path.join(
                os.path.dirname(os.path.dirname(__file__)),
                'instance',
                'google_calendar_token.json'
            )

            if not os.path.exists(token_path):
                return False

            with open(token_path, 'r') as f:
                token_data = json.load(f)

            refresh_token = token_data.get('refresh_token')
            if not refresh_token:
                return False

            # 실제로 토큰이 유효한지 검증 (revoke 여부 확인)
            creds = Credentials.from_authorized_user_info(token_data)

            # 토큰이 만료되었고 refresh_token이 있으면 갱신 시도
            if creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                    # 갱신 성공 시 토큰 파일 업데이트
                    from dashboard.utils.google_calendar import save_credentials as save_calendar_credentials
                    save_calendar_credentials(creds)
                    logger.info("[OAuth] 토큰 갱신 성공 (로그인 URL 생성 중)")
                except Exception as e:
                    # 갱신 실패 = revoke됨 or 유효하지 않음
                    logger.warning(f"[OAuth] 토큰 갱신 실패 (재인증 필요): {e}")
                    return False

            return True
        except Exception as e:
            logger.debug(f"토큰 파일 체크 실패: {e}")
            return False
    
    def get_authorization_url(self, redirect_uri=None):
        """구글 로그인 URL 생성"""
        try:
            flow = self.create_flow(redirect_uri)

            # 토큰 파일에 refresh_token이 있는지 체크 (UX 개선)
            has_refresh_token = self._has_valid_refresh_token()

            # refresh_token이 없으면 select_account + consent
            # 있으면 select_account만 (권한 승인 화면 스킵)
            prompt = 'select_account' if has_refresh_token else 'select_account consent'

            authorization_url, state = flow.authorization_url(
                access_type='offline',
                include_granted_scopes='true',
                prompt=prompt
            )
            return authorization_url, state
        except Exception as e:
            logger.error(f"OAuth URL 생성 오류: {str(e)}")
            return None, None
    
    def get_user_info(self, authorization_code, state, redirect_uri=None):
        """인증 코드로 사용자 정보 가져오기 (credentials도 함께 반환)"""
        try:
            flow = self.create_flow(redirect_uri)
            flow.fetch_token(code=authorization_code)

            # 사용자 정보 요청
            import requests
            credentials = flow.credentials
            userinfo_url = 'https://www.googleapis.com/oauth2/v2/userinfo'

            response = requests.get(userinfo_url, headers={
                'Authorization': f'Bearer {credentials.token}'
            })

            if response.status_code == 200:
                user_info = response.json()
                validated_user = self._validate_user(user_info)
                if validated_user:
                    # 사용자 정보와 credentials를 함께 반환
                    return validated_user, credentials
                return None, None
            else:
                logger.error(f"사용자 정보 가져오기 실패: {response.status_code}")
                return None, None

        except Exception as e:
            logger.error(f"OAuth 사용자 정보 가져오기 오류: {str(e)}")
            return None, None
    
    def _validate_user(self, user_info):
        """사용자 도메인 검증 및 정보 정리"""
        email = user_info.get('email', '').lower()
        
        if not email.endswith(f'@{self.allowed_domain}'):
            logger.warning(f"허용되지 않은 도메인: {email}")
            return None
        
        return {
            'email': email,
            'name': user_info.get('name', ''),
            'picture': user_info.get('picture', ''),
            'verified_email': user_info.get('verified_email', False)
        }

# 전역 OAuth 매니저 인스턴스 (Lazy Loading)
_google_oauth = None


def get_google_oauth():
    """
    OAuth 매니저 인스턴스 가져오기 (Lazy Loading)

    Returns:
        GoogleOAuthManager: OAuth 매니저 인스턴스

    Note:
        - 모듈 import 시점이 아닌 첫 호출 시 초기화
        - Flask app context가 있으면 app.config에서 설정 읽기
        - 운영 환경에서 설정 변경 가능
    """
    global _google_oauth
    if _google_oauth is None:
        _google_oauth = GoogleOAuthManager()
    return _google_oauth


def is_oauth_configured():
    """OAuth가 설정되어 있는지 확인"""
    oauth = get_google_oauth()
    return oauth.client_config is not None