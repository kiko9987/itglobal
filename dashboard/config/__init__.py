"""
Configuration Factory
환경별 설정을 관리하는 중앙 팩토리
"""

import os
from .base import BaseConfig
from .development import DevelopmentConfig
from .production import ProductionConfig
from .testing import TestingConfig

# 환경별 설정 매핑
config_map = {
    'development': DevelopmentConfig,
    'production': ProductionConfig,
    'testing': TestingConfig,
    'default': DevelopmentConfig
}

def get_config(env_name=None):
    """
    환경에 맞는 설정 클래스 반환

    Args:
        env_name: 환경 이름 ('development', 'production', 'testing')
                 None인 경우 FLASK_ENV 환경변수 사용

    Returns:
        Config: 해당 환경의 설정 클래스
    """
    if env_name is None:
        env_name = os.environ.get('FLASK_ENV', 'development')

    return config_map.get(env_name, config_map['default'])

def create_config(env_name=None, overrides=None):
    """
    설정 인스턴스 생성

    Args:
        env_name: 환경 이름
        overrides: 추가 설정 딕셔너리

    Returns:
        Config: 설정 인스턴스
    """
    config_class = get_config(env_name)
    config = config_class()

    # 오버라이드 적용
    if overrides:
        for key, value in overrides.items():
            setattr(config, key, value)

    return config