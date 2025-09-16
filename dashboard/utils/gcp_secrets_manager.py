"""
GCP Secret Manager 통합
- 민감한 환경변수 안전 관리
- 런타임 시크릿 로딩
- 권한 기반 접근 제어
"""

import os
import logging
import json
from typing import Dict, Any, Optional
from google.cloud import secretmanager
from google.api_core import exceptions
import hashlib
import time

logger = logging.getLogger(__name__)

class GCPSecretsManager:
    """GCP Secret Manager 클라이언트"""

    def __init__(self):
        self.client = None
        self.project_id = None
        self.cache = {}
        self.cache_ttl = 300  # 5분 캐시
        self.init_client()

    def init_client(self):
        """Secret Manager 클라이언트 초기화"""
        try:
            self.project_id = os.getenv('GOOGLE_CLOUD_PROJECT')
            if not self.project_id:
                logger.warning("GOOGLE_CLOUD_PROJECT 환경변수가 설정되지 않음 - 로컬 환경으로 폴백")
                return

            self.client = secretmanager.SecretManagerServiceClient()
            logger.info(f"Secret Manager 초기화 완료: {self.project_id}")

        except Exception as e:
            logger.error(f"Secret Manager 초기화 실패: {e}")
            self.client = None

    def get_secret(self, secret_name: str, version: str = "latest") -> Optional[str]:
        """시크릿 값 조회"""
        if not self.client:
            return self._get_local_fallback(secret_name)

        # 캐시 확인
        cache_key = f"{secret_name}:{version}"
        if cache_key in self.cache:
            cached_data, timestamp = self.cache[cache_key]
            if time.time() - timestamp < self.cache_ttl:
                return cached_data

        try:
            name = f"projects/{self.project_id}/secrets/{secret_name}/versions/{version}"
            response = self.client.access_secret_version(request={"name": name})
            secret_value = response.payload.data.decode("UTF-8")

            # 캐시 저장
            self.cache[cache_key] = (secret_value, time.time())

            logger.debug(f"시크릿 조회 성공: {secret_name}")
            return secret_value

        except exceptions.NotFound:
            logger.error(f"시크릿을 찾을 수 없습니다: {secret_name}")
            return self._get_local_fallback(secret_name)
        except Exception as e:
            logger.error(f"시크릿 조회 실패: {secret_name} - {e}")
            return self._get_local_fallback(secret_name)

    def get_secret_json(self, secret_name: str, version: str = "latest") -> Optional[Dict[str, Any]]:
        """JSON 형태의 시크릿 조회"""
        secret_value = self.get_secret(secret_name, version)
        if secret_value:
            try:
                return json.loads(secret_value)
            except json.JSONDecodeError as e:
                logger.error(f"JSON 파싱 실패: {secret_name} - {e}")
        return None

    def create_secret(self, secret_name: str, secret_value: str,
                     labels: Dict[str, str] = None) -> bool:
        """새 시크릿 생성"""
        if not self.client:
            logger.warning("Secret Manager 비활성화 상태 - 시크릿 생성 불가")
            return False

        try:
            parent = f"projects/{self.project_id}"
            secret = {
                "replication": {"automatic": {}},
                "labels": labels or {}
            }

            # 시크릿 생성
            response = self.client.create_secret(
                request={"parent": parent, "secret_id": secret_name, "secret": secret}
            )

            # 초기 버전 추가
            self.client.add_secret_version(
                request={
                    "parent": response.name,
                    "payload": {"data": secret_value.encode("UTF-8")}
                }
            )

            logger.info(f"시크릿 생성 완료: {secret_name}")
            return True

        except exceptions.AlreadyExists:
            logger.warning(f"시크릿이 이미 존재합니다: {secret_name}")
            return self.update_secret(secret_name, secret_value)
        except Exception as e:
            logger.error(f"시크릿 생성 실패: {secret_name} - {e}")
            return False

    def update_secret(self, secret_name: str, secret_value: str) -> bool:
        """시크릿 값 업데이트"""
        if not self.client:
            return False

        try:
            parent = f"projects/{self.project_id}/secrets/{secret_name}"
            self.client.add_secret_version(
                request={
                    "parent": parent,
                    "payload": {"data": secret_value.encode("UTF-8")}
                }
            )

            # 캐시 무효화
            cache_keys_to_remove = [key for key in self.cache.keys() if key.startswith(f"{secret_name}:")]
            for key in cache_keys_to_remove:
                del self.cache[key]

            logger.info(f"시크릿 업데이트 완료: {secret_name}")
            return True

        except Exception as e:
            logger.error(f"시크릿 업데이트 실패: {secret_name} - {e}")
            return False

    def _get_local_fallback(self, secret_name: str) -> Optional[str]:
        """로컬 환경 폴백"""
        # 환경변수에서 조회
        env_value = os.getenv(secret_name.upper())
        if env_value:
            return env_value

        # .env 파일에서 조회 (개발 환경)
        env_file_path = os.path.join(os.path.dirname(__file__), '..', '.env')
        if os.path.exists(env_file_path):
            try:
                with open(env_file_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.strip() and not line.startswith('#'):
                            key, value = line.strip().split('=', 1)
                            if key.strip() == secret_name.upper():
                                return value.strip().strip('"\'')
            except Exception as e:
                logger.debug(f".env 파일 읽기 실패: {e}")

        return None

    def bulk_get_secrets(self, secret_names: list) -> Dict[str, str]:
        """여러 시크릿 일괄 조회"""
        results = {}
        for secret_name in secret_names:
            value = self.get_secret(secret_name)
            if value:
                results[secret_name] = value
        return results

    def get_database_config(self) -> Dict[str, str]:
        """데이터베이스 설정 조회"""
        secrets = self.bulk_get_secrets([
            'cloud_sql_instance',
            'cloud_sql_database',
            'cloud_sql_user',
            'cloud_sql_password',
            'cloud_redis_host',
            'cloud_redis_auth'
        ])

        # 기본값 설정
        config = {
            'CLOUD_SQL_INSTANCE': secrets.get('cloud_sql_instance', 'itglobal-main'),
            'CLOUD_SQL_DATABASE': secrets.get('cloud_sql_database', 'itglobal_db'),
            'CLOUD_SQL_USER': secrets.get('cloud_sql_user', 'postgres'),
            'CLOUD_SQL_PASSWORD': secrets.get('cloud_sql_password', ''),
            'CLOUD_REDIS_HOST': secrets.get('cloud_redis_host', ''),
            'CLOUD_REDIS_AUTH': secrets.get('cloud_redis_auth', '')
        }

        return config

    def get_google_credentials(self) -> Optional[Dict[str, Any]]:
        """Google API 인증 정보 조회"""
        # 서비스 계정 키 (JSON 형태)
        service_account_key = self.get_secret_json('google_service_account')
        if service_account_key:
            return service_account_key

        # 개별 필드로 저장된 경우
        credentials = self.bulk_get_secrets([
            'google_sheets_client_id',
            'google_sheets_client_secret',
            'google_oauth_client_id',
            'google_oauth_client_secret'
        ])

        if credentials:
            return {
                'sheets_client_id': credentials.get('google_sheets_client_id'),
                'sheets_client_secret': credentials.get('google_sheets_client_secret'),
                'oauth_client_id': credentials.get('google_oauth_client_id'),
                'oauth_client_secret': credentials.get('google_oauth_client_secret')
            }

        return None

    def get_application_secrets(self) -> Dict[str, str]:
        """애플리케이션 시크릿 조회"""
        secrets = self.bulk_get_secrets([
            'flask_secret_key',
            'jwt_secret_key',
            'csrf_secret_key',
            'encryption_key'
        ])

        # 시크릿이 없으면 임시 키 생성 (개발 환경용)
        if not secrets.get('flask_secret_key'):
            logger.warning("Flask secret key가 설정되지 않음 - 임시 키 사용")
            secrets['flask_secret_key'] = hashlib.sha256(
                f"temp_key_{self.project_id}_{int(time.time())}".encode()
            ).hexdigest()

        return secrets

    def setup_environment_variables(self):
        """환경변수 자동 설정"""
        # 데이터베이스 설정
        db_config = self.get_database_config()
        for key, value in db_config.items():
            if value:
                os.environ[key] = value

        # 애플리케이션 시크릿
        app_secrets = self.get_application_secrets()
        for key, value in app_secrets.items():
            if value:
                os.environ[key.upper()] = value

        # Google 인증 정보
        google_creds = self.get_google_credentials()
        if google_creds:
            if isinstance(google_creds, dict) and 'type' in google_creds:
                # 서비스 계정 키 파일 생성
                import tempfile
                with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
                    json.dump(google_creds, f)
                    os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = f.name
            else:
                # 개별 설정
                for key, value in google_creds.items():
                    if value:
                        os.environ[key.upper()] = value

        logger.info("환경변수 설정 완료")

    def health_check(self) -> Dict[str, Any]:
        """Secret Manager 상태 확인"""
        status = {
            'available': False,
            'project_id': self.project_id,
            'cache_size': len(self.cache),
            'error': None
        }

        if not self.client:
            status['error'] = 'Secret Manager 클라이언트 미초기화'
            return status

        try:
            # 테스트 시크릿 조회 시도
            test_secret = self.get_secret('health_check_test')
            status['available'] = True
            status['test_result'] = 'success' if test_secret else 'no_test_secret'
        except Exception as e:
            status['error'] = str(e)

        return status


# 전역 인스턴스
secrets_manager = None

def get_secrets_manager() -> GCPSecretsManager:
    """Secret Manager 인스턴스 반환"""
    global secrets_manager
    if secrets_manager is None:
        secrets_manager = GCPSecretsManager()
    return secrets_manager

def init_secrets():
    """시크릿 초기화 및 환경변수 설정"""
    global secrets_manager
    secrets_manager = GCPSecretsManager()
    secrets_manager.setup_environment_variables()
    return secrets_manager

def get_secret(name: str) -> Optional[str]:
    """편의 함수: 시크릿 조회"""
    return get_secrets_manager().get_secret(name)