"""
Google OAuth 구현
@itg-aircon.com 도메인 제한 포함
"""

import os
import json
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import logging

logger = logging.getLogger(__name__)

class GoogleOAuthManager:
    def __init__(self):
        self.client_config = self._get_oauth_config()
        self.allowed_domain = "itg-aircon.com"
        
    def _get_oauth_config(self):
        """OAuth 클라이언트 설정 로드"""
        # 환경 변수에서 OAuth 설정을 가져오거나 파일에서 로드
        client_id = os.getenv('GOOGLE_OAUTH_CLIENT_ID')
        client_secret = os.getenv('GOOGLE_OAUTH_CLIENT_SECRET')
        
        if client_id and client_secret:
            return {
                "web": {
                    "client_id": client_id,
                    "client_secret": client_secret,
                    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                    "token_uri": "https://oauth2.googleapis.com/token",
                    "redirect_uris": ["http://localhost:5000/auth/callback"]
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
                'openid'
            ]
        )
        
        flow.redirect_uri = redirect_uri or "http://localhost:5000/auth/callback"
        return flow
    
    def get_authorization_url(self, redirect_uri=None):
        """구글 로그인 URL 생성"""
        try:
            flow = self.create_flow(redirect_uri)
            authorization_url, state = flow.authorization_url(
                access_type='offline',
                include_granted_scopes='true',
                prompt='select_account'  # 계정 선택 강제
            )
            return authorization_url, state
        except Exception as e:
            logger.error(f"OAuth URL 생성 오류: {str(e)}")
            return None, None
    
    def get_user_info(self, authorization_code, state, redirect_uri=None):
        """인증 코드로 사용자 정보 가져오기"""
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
                return self._validate_user(user_info)
            else:
                logger.error(f"사용자 정보 가져오기 실패: {response.status_code}")
                return None
                
        except Exception as e:
            logger.error(f"OAuth 사용자 정보 가져오기 오류: {str(e)}")
            return None
    
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

# 전역 OAuth 매니저 인스턴스
google_oauth = GoogleOAuthManager()

def is_oauth_configured():
    """OAuth가 설정되어 있는지 확인"""
    return google_oauth.client_config is not None