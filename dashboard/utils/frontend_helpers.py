"""
프론트엔드 빌드 파일 관리 유틸리티
Vite로 빌드된 파일들의 해시를 읽어와 Flask 템플릿에서 사용
"""

import os
import json
import glob
from typing import Dict, Optional
import logging

logger = logging.getLogger(__name__)

class FrontendAssetManager:
    """프론트엔드 에셋 관리 클래스"""

    def __init__(self, dist_path: str = None):
        if dist_path is None:
            dist_path = os.path.join(os.path.dirname(__file__), '..', 'static', 'dist')

        self.dist_path = os.path.abspath(dist_path)
        self._asset_cache = {}
        self._manifest_cache = None
        self._manifest_mtime = None
        self._app = None  # Flask 앱 참조 (init_frontend_helpers에서 설정)
        # 캐시 버스팅용 타임스탬프 초기화
        import time
        self._cache_buster = str(int(time.time()))

    def get_asset_url(self, asset_name: str, cache_bust: bool = True) -> Optional[str]:
        """
        에셋 이름으로 해시가 포함된 실제 파일 URL 반환
        manifest 기반 단일 진입점 - manifest가 없을 때만 fallback 사용

        Args:
            asset_name: 'main.css', 'project-list.js' 등
            cache_bust: 캐시 버스팅 파라미터 추가 여부

        Returns:
            실제 파일 URL 또는 None
        """
        try:
            import time

            # Vite 개발 서버 사용 시 개발 서버 URL 반환
            if self._is_vite_dev_server_enabled():
                return self._get_vite_dev_url(asset_name)

            # 캐시에서 먼저 확인 (캐시 버스팅이 비활성화된 경우만)
            if not cache_bust and asset_name in self._asset_cache:
                return self._asset_cache[asset_name]

            # 1단계: manifest.json 기반 로딩 (우선순위)
            manifest = self._load_manifest()
            if manifest:
                url = self._get_url_from_manifest(manifest, asset_name)
                if url:
                    # 캐시 버스팅 파라미터 추가
                    if cache_bust:
                        url += f"?v={int(time.time())}"
                    self._asset_cache[asset_name] = url
                    return url

            # 2단계: manifest가 존재하지만 에셋을 찾지 못한 경우 - legacy 사용 안 함 (ver16 피드백)
            if manifest:  # manifest는 있지만 해당 에셋이 없는 경우
                logger.warning(f"모던 번들에서 에셋을 찾지 못함 (legacy 사용 안 함): {asset_name}")
                return None

            # 3단계: manifest 자체가 없는 경우에만 fallback 사용 (빌드 실패 상황)
            logger.debug(f"manifest 없음 - legacy fallback 사용: {asset_name}")
            return self._get_fallback_url(asset_name, cache_bust)

        except Exception as e:
            logger.warning(f"에셋 URL 조회 실패: {asset_name} - {e}")
            return self._get_fallback_url(asset_name, cache_bust)

    def _get_url_from_manifest(self, manifest: Dict, asset_name: str) -> Optional[str]:
        """manifest에서 에셋 URL 추출 (단일 진입점)"""
        # 직접 키로 찾기
        if asset_name in manifest:
            file_info = manifest[asset_name]
            if isinstance(file_info, dict) and 'file' in file_info:
                return f"/static/dist/{file_info['file']}"
            else:
                return f"/static/dist/{file_info}"

        # JS 파일의 경우 pages 디렉토리에서 찾기
        if asset_name.endswith('.js'):
            js_key = f"js/pages/{asset_name}"
            if js_key in manifest:
                file_info = manifest[js_key]
                if isinstance(file_info, dict) and 'file' in file_info:
                    return f"/static/dist/{file_info['file']}"
                else:
                    return f"/static/dist/{file_info}"

        # CSS 파일의 경우 css 디렉토리에서 찾기
        if asset_name.endswith('.css'):
            css_key = f"css/{asset_name}"
            if css_key in manifest:
                file_info = manifest[css_key]
                if isinstance(file_info, dict) and 'file' in file_info:
                    return f"/static/dist/{file_info['file']}"
                else:
                    return f"/static/dist/{file_info}"

            # CSS가 JS 엔트리 포인트의 'css' 배열에 있는 경우 찾기
            # 예: manifest["js/pages/leads.js"]["css"][0] = "css/leads.DhMKueFq.css"
            bundle_name = asset_name.replace('.css', '')
            js_entry_key = f"js/pages/{bundle_name}.js"
            if js_entry_key in manifest:
                entry_info = manifest[js_entry_key]
                if isinstance(entry_info, dict) and 'css' in entry_info:
                    css_files = entry_info['css']
                    if isinstance(css_files, list) and len(css_files) > 0:
                        # 첫 번째 CSS 파일 반환 (보통 하나만 있음)
                        css_file = css_files[0]
                        return f"/static/dist/{css_file}"

        # 패턴으로 찾기 (예: main.css -> css/main.css)
        for key, file_info in manifest.items():
            if key.endswith(asset_name):
                if isinstance(file_info, dict) and 'file' in file_info:
                    return f"/static/dist/{file_info['file']}"
                else:
                    return f"/static/dist/{file_info}"

        return None

    def _load_manifest(self) -> Optional[Dict]:
        """Vite manifest.json 파일 로드 (수정 시간 기반 캐시 무효화)"""
        # Vite는 manifest.json을 .vite 폴더에 생성함
        manifest_path = os.path.join(self.dist_path, '.vite', 'manifest.json')

        if not os.path.exists(manifest_path):
            self._manifest_cache = {}
            self._manifest_mtime = None
            return {}

        try:
            # 파일 수정 시간 확인
            current_mtime = os.path.getmtime(manifest_path)

            # 캐시가 있고 파일이 변경되지 않았다면 캐시 사용
            if (self._manifest_cache is not None and
                self._manifest_mtime is not None and
                current_mtime == self._manifest_mtime):
                return self._manifest_cache

            # 파일이 변경되었거나 첫 로드인 경우 새로 읽기
            with open(manifest_path, 'r', encoding='utf-8') as f:
                self._manifest_cache = json.load(f)
                self._manifest_mtime = current_mtime
                logger.info(f"Manifest 파일 갱신됨: {len(self._manifest_cache)}개 에셋 로드")
                return self._manifest_cache

        except Exception as e:
            logger.warning(f"manifest.json 로드 실패: {e}")
            self._manifest_cache = {}
            self._manifest_mtime = None
            return {}


    def _get_fallback_url(self, asset_name: str, cache_bust: bool = True) -> Optional[str]:
        """
        페이지별 폴백 URL 반환 - 확장 가능한 매핑 시스템

        전문가 피드백: 폴백 테이블을 페이지별로 확장하여 다른 페이지 대응
        """
        logger.warning(f"Manifest 없음, 페이지별 폴백 사용: {asset_name}")

        # 페이지별 자산 매핑 테이블 (확장 가능한 구조)
        page_asset_mappings = self._get_page_asset_mappings()

        # 1단계: 키워드 기반 간단한 매핑 (monitoring vs project)
        fallback_url = self._resolve_simplified_mapping(asset_name, page_asset_mappings)
        if fallback_url:
            return self._validate_asset_exists(fallback_url, asset_name)

        # 2단계: 기본 폴백 (project-list 기본값)
        return self._get_default_fallback(asset_name)

    def _get_page_asset_mappings(self) -> Dict:
        """간소화된 페이지별 자산 매핑 (ver16 분석: 실제 사용되는 매핑만 유지)"""
        return {
            # 실제 사용 가능한 legacy 파일만 지원 - monitoring은 템플릿에서 사용하지 않음
            'legacy_assets': {
                # 프로젝트 관련 (유일한 실제 사용 가능한 fallback)
                'project': {
                    'css': '/static/legacy/css/project-list.css',
                    'js': '/static/legacy/js/project-list.js'
                }
                # monitoring 매핑 제거: 템플릿이 asset_url 헬퍼를 사용하지 않고 직접 static 경로 참조
            }
        }

    def _resolve_simplified_mapping(self, asset_name: str, mappings: Dict) -> Optional[str]:
        """간소화된 매핑 해결 (ver16 분석: 프로젝트 관련 매핑만 지원)"""
        # 파일 타입 확인
        if asset_name.endswith('.css'):
            asset_type = 'css'
        elif asset_name.endswith('.js'):
            asset_type = 'js'
        else:
            return None

        # 프로젝트 관련 에셋만 fallback 제공 (monitoring은 템플릿이 직접 static 경로 사용)
        if 'project' in mappings['legacy_assets']:
            return mappings['legacy_assets']['project'].get(asset_type)

        return None

    def _get_default_fallback(self, asset_name: str) -> str:
        """기본 폴백 (최후 수단)"""
        if asset_name.endswith('.css'):
            return '/static/legacy/css/project-list.css'
        elif asset_name.endswith('.js'):
            return '/static/legacy/js/project-list.js'
        else:
            return f"/static/{asset_name}"


    def _validate_asset_exists(self, asset_url: str, original_asset_name: str) -> str:
        """자산 파일 존재 여부 검증 및 대체 제공"""
        try:
            # /static/ 경로를 실제 파일 시스템 경로로 변환
            if asset_url.startswith('/static/'):
                relative_path = asset_url[len('/static/'):]
                static_root = os.path.join(os.path.dirname(__file__), '..', 'static')
                file_path = os.path.join(static_root, relative_path)

                if os.path.exists(file_path):
                    return asset_url
                else:
                    logger.warning(f"자산 파일 누락: {file_path}, 기본 폴백 사용")
                    return self._get_default_fallback(original_asset_name)
        except Exception as e:
            logger.warning(f"자산 검증 중 오류: {e}")

        return asset_url  # 검증 실패 시 원본 URL 반환

    def get_page_specific_assets(self, page_name: str = None) -> Dict[str, str]:
        """간소화된 페이지별 자산 반환 (ver15 전문가 피드백 적용)"""
        if not page_name:
            try:
                from flask import request
                if request and request.endpoint:
                    page_name = request.endpoint
            except Exception as e:
                page_name = 'default'

        # 간소화된 매핑 로직
        page_lower = page_name.lower() if page_name else ''
        monitoring_keywords = ['monitor', 'dashboard', 'cache', 'system']

        # 모니터링 관련 페이지 감지
        if any(keyword in page_lower for keyword in monitoring_keywords):
            return {
                'css': '/static/legacy/css/monitoring-dashboard.css',
                'js': '/static/legacy/js/monitoring-dashboard.js'
            }

        # 기본값: 프로젝트 자산
        return {
            'css': '/static/legacy/css/project-list.css',
            'js': '/static/legacy/js/project-list.js'
        }

    def get_css_bundle(self, bundle_name: str = 'main') -> Optional[str]:
        """
        CSS 번들 URL 반환

        Args:
            bundle_name: CSS 번들 이름

        Returns:
            CSS 파일 URL
        """
        # 이미 .css 확장자가 있으면 제거 후 다시 추가 (중복 방지)
        if bundle_name.endswith('.css'):
            bundle_name = bundle_name[:-4]

        # 기본 로직 (모던 CSS 우선)
        url = self.get_asset_url(f"{bundle_name}.css")
        if not url:
            logger.error(f"CSS 번들을 찾을 수 없음: {bundle_name}")
        return url


    def get_js_bundle(self, bundle_name: str = 'main') -> Optional[str]:
        """JS 번들 URL 반환 (단일 진입점 사용)"""
        # 이미 .js 확장자가 있으면 제거 후 다시 추가 (중복 방지)
        if bundle_name.endswith('.js'):
            bundle_name = bundle_name[:-3]
        url = self.get_asset_url(f"{bundle_name}.js")
        if not url:
            logger.error(f"JS 번들을 찾을 수 없음: {bundle_name}")
        return url

    def get_all_css_bundles(self) -> Dict[str, str]:
        """모든 CSS 번들 반환"""
        bundles = {}
        css_dir = os.path.join(self.dist_path, 'css')

        if os.path.exists(css_dir):
            for file in os.listdir(css_dir):
                if file.endswith('.css'):
                    # 해시 제거하여 번들 이름 추출
                    bundle_name = file.split('.')[0]
                    bundles[bundle_name] = f"/static/dist/css/{file}"

        return bundles

    def get_all_js_bundles(self) -> Dict[str, str]:
        """모든 JS 번들 반환"""
        bundles = {}
        js_dir = os.path.join(self.dist_path, 'js')

        if os.path.exists(js_dir):
            for file in os.listdir(js_dir):
                if file.endswith('.js') and not file.endswith('.map'):
                    # 해시 제거하여 번들 이름 추출
                    bundle_name = file.split('.')[0]
                    bundles[bundle_name] = f"/static/dist/js/{file}"

        return bundles

    def get_service_worker_assets(self) -> Dict[str, list]:
        """서비스 워커가 캐싱할 자산 목록 반환"""
        assets = {
            'css_files': [],
            'js_files': [],
            'static_files': ['/']  # 기본 경로
        }

        # CSS 번들들 추가
        css_bundles = self.get_all_css_bundles()
        for name, url in css_bundles.items():
            if url:
                assets['css_files'].append(url)

        # JS 번들들 추가
        js_bundles = self.get_all_js_bundles()
        for name, url in js_bundles.items():
            if url:
                assets['js_files'].append(url)

        # 중복 제거
        assets['css_files'] = list(set(assets['css_files']))
        assets['js_files'] = list(set(assets['js_files']))

        return assets

    def _is_vite_dev_server_enabled(self) -> bool:
        """Vite 개발 서버 사용 여부 확인"""
        try:
            # 저장된 앱 참조 사용 (우선순위)
            if self._app:
                return self._app.config.get('VITE_DEV_SERVER_ENABLED', False)

            # 폴백: Flask 컨텍스트 사용
            from flask import current_app
            return current_app.config.get('VITE_DEV_SERVER_ENABLED', False)
        except Exception as e:
            return False

    def _get_vite_dev_url(self, asset_name: str) -> str:
        """Vite 개발 서버 URL 반환 (Vite root: 'src' 설정 반영)"""
        try:
            # 저장된 앱 참조 사용 (우선순위)
            if self._app:
                vite_host = self._app.config.get('VITE_DEV_SERVER_HOST', 'localhost')
                vite_port = self._app.config.get('VITE_DEV_SERVER_PORT', 5173)
            else:
                # 폴백: Flask 컨텍스트 사용
                from flask import current_app
                vite_host = current_app.config.get('VITE_DEV_SERVER_HOST', 'localhost')
                vite_port = current_app.config.get('VITE_DEV_SERVER_PORT', 5173)
        except Exception as e:
            vite_host = 'localhost'
            vite_port = 5173

        # Vite는 root: 'src'로 설정되어 있으므로 /src/ 경로를 포함하면 안 됨
        # 이미 전체 경로가 포함된 경우 (예: js/pages/project-list.js)
        if '/' in asset_name and not asset_name.startswith('http'):
            return f"http://{vite_host}:{vite_port}/{asset_name}"

        # CSS 파일
        if asset_name.endswith('.css'):
            return f"http://{vite_host}:{vite_port}/css/{asset_name}"

        # JS 파일 - 특정 파일별 매핑
        if asset_name.endswith('.js'):
            # authInterceptor는 utils에 있음
            if 'authInterceptor' in asset_name:
                return f"http://{vite_host}:{vite_port}/js/utils/{asset_name}"
            # 기본적으로 pages에서 찾기
            else:
                return f"http://{vite_host}:{vite_port}/js/pages/{asset_name}"

        # 기타 파일
        return f"http://{vite_host}:{vite_port}/{asset_name}"

    def clear_cache(self):
        """캐시 초기화"""
        self._asset_cache.clear()
        self._manifest_cache = None

    def force_reload_manifest(self):
        """manifest.json 강제 재로드"""
        self._manifest_cache = None
        self._asset_cache.clear()
        # 캐시 버스터도 새로 생성
        import time
        self._cache_buster = str(int(time.time()))
        return self._load_manifest()

    def is_modern_build_available(self) -> bool:
        """모던 빌드가 사용 가능한지 확인. 개발 모드에서는 Vite 개발 서버 확인. (요청 레벨 캐싱)"""
        # Flask request context에서 캐시 확인
        try:
            from flask import g
            if hasattr(g, '_modern_build_available_cache'):
                return g._modern_build_available_cache
        except Exception as e:
            logger.warning(f"처리 중 오류 무시: {e}")
            pass

        # Vite 개발 서버 감지 로직 import
        result = False
        try:
            from dashboard.app import is_vite_dev_server_running
            # 개발 모드에서는 Vite 개발 서버가 실행 중인지 확인
            if is_vite_dev_server_running():
                result = True
        except ImportError:
            pass

        # Vite 서버가 없으면 프로덕션 모드에서 빌드된 파일 확인
        if not result:
            result = (
                os.path.exists(os.path.join(self.dist_path, 'css')) and
                os.path.exists(os.path.join(self.dist_path, 'js'))
            )

        # 요청 레벨 캐시에 저장
        try:
            from flask import g
            g._modern_build_available_cache = result
        except Exception as e:
            logger.warning(f"처리 중 오류 무시: {e}")
            pass

        return result

# 전역 인스턴스
asset_manager = FrontendAssetManager()

def init_frontend_helpers(app):
    """Flask 앱에 프론트엔드 헬퍼 함수들 등록"""

    # Flask 앱 참조를 asset_manager에 저장
    asset_manager._app = app

    @app.template_global()
    def get_css_bundle(bundle_name='main'):
        """템플릿에서 CSS 번들 URL 가져오기"""
        return asset_manager.get_css_bundle(bundle_name)

    @app.template_global()
    def get_js_bundle(bundle_name='main'):
        """템플릿에서 JS 번들 URL 가져오기"""
        return asset_manager.get_js_bundle(bundle_name)

    @app.template_global()
    def is_modern_build():
        """모던 빌드 사용 가능 여부"""
        return asset_manager.is_modern_build_available()

    @app.template_global()
    def modern_build_available():
        """모던 빌드 사용 가능 여부 (함수 호출 형태)"""
        return asset_manager.is_modern_build_available()

    @app.template_global()
    def get_all_css_bundles():
        """모든 CSS 번들 목록"""
        return asset_manager.get_all_css_bundles()

    @app.template_global()
    def get_all_js_bundles():
        """모든 JS 번들 목록"""
        return asset_manager.get_all_js_bundles()

    @app.template_global()
    def cache_buster():
        """캐시 버스팅용 타임스탬프 반환"""
        import time
        # 기존 cache_buster가 있으면 사용, 없으면 현재 시간 사용
        if hasattr(asset_manager, '_cache_buster'):
            return asset_manager._cache_buster
        return str(int(time.time()))

    @app.context_processor
    def inject_frontend_assets():
        """템플릿 컨텍스트에 에셋 정보 주입"""
        return {
            'css_bundles': asset_manager.get_all_css_bundles(),
            'js_bundles': asset_manager.get_all_js_bundles(),
            'modern_build_available': asset_manager.is_modern_build_available,  # 함수로 호출 가능
            'get_css_bundle': asset_manager.get_css_bundle,  # 함수 직접 전달
            'get_js_bundle': asset_manager.get_js_bundle,    # 함수 직접 전달
            'get_vite_js_url': lambda: asset_manager.get_js_bundle('js/pages/project-list.js'),  # 호환성 함수
            'get_vite_css_url': lambda: asset_manager.get_css_bundle('css/main.css'),  # 호환성 함수
        }

    # 개발 모드에서 파일 변경 감지 및 조건부 캐시 클리어
    if app.debug:
        last_check_time = 0
        last_files_mtime = {}

        @app.before_request
        def clear_asset_cache():
            nonlocal last_check_time, last_files_mtime
            import time

            current_time = time.time()
            # 5초마다만 파일 변경 검사 (과도한 파일 시스템 체크 방지)
            if current_time - last_check_time < 5.0:
                return

            last_check_time = current_time
            files_changed = False

            # 검사할 파일들 목록
            check_files = [
                os.path.join(asset_manager.dist_path, '.vite', 'manifest.json'),
                os.path.join(asset_manager.dist_path, 'css'),
                os.path.join(asset_manager.dist_path, 'js')
            ]

            for file_path in check_files:
                if os.path.exists(file_path):
                    if os.path.isdir(file_path):
                        # 디렉토리의 경우 최신 파일 수정 시간 확인 (1단계만, 성능 최적화)
                        try:
                            latest_mtime = 0
                            # os.walk 대신 os.listdir로 1단계만 확인 (성능 향상)
                            for item in os.listdir(file_path):
                                item_path = os.path.join(file_path, item)
                                if os.path.isfile(item_path) and item.endswith(('.css', '.js', '.map')):
                                    mtime = os.path.getmtime(item_path)
                                    latest_mtime = max(latest_mtime, mtime)

                            if file_path not in last_files_mtime or last_files_mtime[file_path] != latest_mtime:
                                last_files_mtime[file_path] = latest_mtime
                                files_changed = True
                        except OSError:
                            pass
                    else:
                        # 개별 파일의 수정 시간 확인
                        try:
                            mtime = os.path.getmtime(file_path)
                            if file_path not in last_files_mtime or last_files_mtime[file_path] != mtime:
                                last_files_mtime[file_path] = mtime
                                files_changed = True
                        except OSError:
                            pass

            # 파일이 변경된 경우에만 캐시 클리어
            if files_changed:
                asset_manager.clear_cache()
                asset_manager.force_reload_manifest()
                # 캐시 버스팅을 위한 타임스탬프 추가
                asset_manager._cache_buster = str(int(current_time))
                logger.debug("에셋 파일 변경 감지됨 - 캐시 클리어 실행")


# 테스트용 함수들
def test_asset_manager():
    """에셋 매니저 테스트"""
    manager = FrontendAssetManager()

    print("=== Asset Manager 테스트 ===")
    print(f"모던 빌드 사용 가능: {manager.is_modern_build_available()}")
    print(f"메인 CSS: {manager.get_css_bundle('main')}")
    print(f"프로젝트 목록 JS: {manager.get_js_bundle('project-list')}")
    print(f"모든 CSS 번들: {manager.get_all_css_bundles()}")
    print(f"모든 JS 번들: {manager.get_all_js_bundles()}")

if __name__ == "__main__":
    test_asset_manager()