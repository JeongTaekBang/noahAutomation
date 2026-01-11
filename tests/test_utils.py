"""
utils 모듈 테스트
"""

import pandas as pd
import pytest

from po_generator.utils import (
    get_safe_value,
    format_currency,
)


class TestGetSafeValue:
    """get_safe_value 함수 테스트"""

    def test_existing_value(self):
        """존재하는 값"""
        data = pd.Series({'name': 'Test', 'qty': 10})
        assert get_safe_value(data, 'name') == 'Test'
        assert get_safe_value(data, 'qty') == 10

    def test_missing_key(self):
        """없는 키"""
        data = pd.Series({'name': 'Test'})
        assert get_safe_value(data, 'missing') == ''
        assert get_safe_value(data, 'missing', 'default') == 'default'

    def test_nan_value(self):
        """NaN 값"""
        data = pd.Series({'name': pd.NA, 'qty': float('nan')})
        assert get_safe_value(data, 'name') == ''
        assert get_safe_value(data, 'qty') == ''

    def test_nan_string_value(self):
        """문자열 'nan' (대소문자 무관)"""
        data = pd.Series({'lower': 'nan', 'upper': 'NaN', 'all_upper': 'NAN'})
        assert get_safe_value(data, 'lower') == ''
        assert get_safe_value(data, 'upper') == ''
        assert get_safe_value(data, 'all_upper') == ''

    def test_nan_like_string_not_filtered(self):
        """'nan'을 포함하지만 실제 값인 문자열은 유지"""
        data = pd.Series({
            'name1': 'Banana',
            'name2': 'Ferdinand',
            'name3': 'nano',
            'name4': 'ANNAN',
        })
        # 이 값들은 'nan' 문자열이 아니므로 그대로 반환되어야 함
        assert get_safe_value(data, 'name1') == 'Banana'
        assert get_safe_value(data, 'name2') == 'Ferdinand'
        assert get_safe_value(data, 'name3') == 'nano'
        assert get_safe_value(data, 'name4') == 'ANNAN'

    def test_none_value(self):
        """None 값"""
        data = pd.Series({'value': None})
        assert get_safe_value(data, 'value') == ''
        assert get_safe_value(data, 'value', 'fallback') == 'fallback'

    def test_zero_is_valid(self):
        """0은 유효한 값"""
        data = pd.Series({'qty': 0})
        # 0은 falsy지만 유효한 값으로 반환되어야 함
        result = get_safe_value(data, 'qty')
        assert result == 0


class TestFormatCurrency:
    """format_currency 함수 테스트"""

    def test_krw_format(self):
        """원화 포맷"""
        assert format_currency(1000000, 'KRW') == '₩1,000,000'
        assert format_currency(0, 'KRW') == '₩0'

    def test_usd_format(self):
        """달러 포맷"""
        assert format_currency(1234.56, 'USD') == '$1,234.56'
        assert format_currency(0, 'USD') == '$0.00'

    def test_large_number(self):
        """큰 숫자"""
        assert format_currency(999999999, 'KRW') == '₩999,999,999'
