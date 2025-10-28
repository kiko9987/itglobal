"""
DataFrame pickle 직렬화 테스트 (단위 테스트 - Redis 불필요)

pandas.DataFrame이 pickle로 올바르게 직렬화/역직렬화되는지 검증
"""

import pytest
import pickle
import pandas as pd
import numpy as np


def test_dataframe_pickle_serialization():
    """DataFrame pickle 직렬화/역직렬화 테스트"""
    # 테스트 DataFrame 생성
    df = pd.DataFrame({
        '프로젝트코드': ['TEST001', 'TEST002', 'TEST003'],
        '프로젝트명': ['테스트1', '테스트2', '테스트3'],
        '계약금액': [1000000, 2000000, 3000000],
        '진행율': [30.5, 50.0, 75.5],
        '완료여부': [False, False, True]
    })

    # Pickle 직렬화
    serialized = pickle.dumps(df, protocol=pickle.HIGHEST_PROTOCOL)
    assert isinstance(serialized, bytes)

    # Pickle 역직렬화
    deserialized = pickle.loads(serialized)
    assert isinstance(deserialized, pd.DataFrame)

    # 데이터 동일성 검증
    pd.testing.assert_frame_equal(df, deserialized)


def test_dataframe_with_nan_pickle():
    """NaN 값이 포함된 DataFrame pickle 테스트"""
    df = pd.DataFrame({
        'A': [1, 2, np.nan, 4],
        'B': ['a', None, 'c', 'd'],
        'C': [True, False, None, True]
    })

    serialized = pickle.dumps(df)
    deserialized = pickle.loads(serialized)

    pd.testing.assert_frame_equal(df, deserialized)


def test_dataframe_with_datetime_pickle():
    """datetime 컬럼이 포함된 DataFrame pickle 테스트"""
    df = pd.DataFrame({
        '날짜': pd.date_range('2025-01-01', periods=3),
        '값': [100, 200, 300]
    })

    serialized = pickle.dumps(df)
    deserialized = pickle.loads(serialized)

    pd.testing.assert_frame_equal(df, deserialized)


def test_empty_dataframe_pickle():
    """빈 DataFrame pickle 테스트"""
    df = pd.DataFrame()

    serialized = pickle.dumps(df)
    deserialized = pickle.loads(serialized)

    pd.testing.assert_frame_equal(df, deserialized)


def test_large_dataframe_pickle():
    """대용량 DataFrame pickle 테스트"""
    # 10,000행 DataFrame
    df = pd.DataFrame({
        f'col_{i}': np.random.rand(10000)
        for i in range(10)
    })

    serialized = pickle.dumps(df, protocol=pickle.HIGHEST_PROTOCOL)
    deserialized = pickle.loads(serialized)

    pd.testing.assert_frame_equal(df, deserialized)


def test_dataframe_dtypes_preserved():
    """DataFrame 데이터 타입 보존 테스트"""
    df = pd.DataFrame({
        'int_col': [1, 2, 3],
        'float_col': [1.1, 2.2, 3.3],
        'str_col': ['a', 'b', 'c'],
        'bool_col': [True, False, True],
        'datetime_col': pd.date_range('2025-01-01', periods=3)
    })

    serialized = pickle.dumps(df)
    deserialized = pickle.loads(serialized)

    # 데이터 타입 검증
    assert deserialized['int_col'].dtype == df['int_col'].dtype
    assert deserialized['float_col'].dtype == df['float_col'].dtype
    assert deserialized['str_col'].dtype == df['str_col'].dtype
    assert deserialized['bool_col'].dtype == df['bool_col'].dtype
    assert deserialized['datetime_col'].dtype == df['datetime_col'].dtype


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
