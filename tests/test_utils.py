"""
utils 모듈 테스트
"""

from pathlib import Path
from unittest.mock import patch, MagicMock

import pandas as pd
import pytest

from po_generator.utils import (
    get_safe_value,
    format_currency,
    load_noah_po_lists,
    find_order_data,
    escape_excel_formula,
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

    def test_other_currency_defaults_to_usd_format(self):
        """KRW 외 통화는 USD 형식 사용"""
        assert format_currency(1234.56, 'EUR') == '$1,234.56'
        assert format_currency(1234.56, 'JPY') == '$1,234.56'


class TestLoadNoahPoLists:
    """load_noah_po_lists 함수 테스트"""

    def test_file_not_found_raises_error(self, tmp_path):
        """소스 파일이 없으면 FileNotFoundError"""
        with patch('po_generator.utils.NOAH_PO_LISTS_FILE', tmp_path / "없는파일.xlsx"):
            with pytest.raises(FileNotFoundError) as exc_info:
                load_noah_po_lists()
            assert "소스 파일을 찾을 수 없습니다" in str(exc_info.value)

    def test_loads_and_merges_sheets(self, tmp_path):
        """국내/해외 시트를 로드하고 합침"""
        # 테스트용 Excel 파일 생성
        test_file = tmp_path / "test_po_lists.xlsx"

        df_domestic = pd.DataFrame({
            'RCK Order no.': ['ND-0001', 'ND-0002'],
            'Customer name': ['고객A', '고객B'],
            'Item qty': [1, 2],
        })
        df_export = pd.DataFrame({
            'RCK Order no.': ['NE-0001'],
            'Customer name': ['Customer C'],
            'Item qty': [3],
            'Export only': ['Y'],  # 해외 시트에만 있는 컬럼
        })

        with pd.ExcelWriter(test_file) as writer:
            df_domestic.to_excel(writer, sheet_name='국내', index=False)
            df_export.to_excel(writer, sheet_name='해외', index=False)

        with patch('po_generator.utils.NOAH_PO_LISTS_FILE', test_file):
            result = load_noah_po_lists()

        # 3건 합쳐짐
        assert len(result) == 3
        # 시트 구분 컬럼 추가됨
        assert '_시트구분' in result.columns
        assert list(result['_시트구분']) == ['국내', '국내', '해외']
        # 모든 컬럼 통일됨
        assert 'Export only' in result.columns

    def test_adds_sheet_identifier_column(self, tmp_path):
        """_시트구분 컬럼이 올바르게 추가됨"""
        test_file = tmp_path / "test_po_lists.xlsx"

        df_domestic = pd.DataFrame({'RCK Order no.': ['ND-0001']})
        df_export = pd.DataFrame({'RCK Order no.': ['NE-0001']})

        with pd.ExcelWriter(test_file) as writer:
            df_domestic.to_excel(writer, sheet_name='국내', index=False)
            df_export.to_excel(writer, sheet_name='해외', index=False)

        with patch('po_generator.utils.NOAH_PO_LISTS_FILE', test_file):
            result = load_noah_po_lists()

        domestic_rows = result[result['_시트구분'] == '국내']
        export_rows = result[result['_시트구분'] == '해외']

        assert len(domestic_rows) == 1
        assert len(export_rows) == 1


class TestFindOrderData:
    """find_order_data 함수 테스트"""

    @pytest.fixture
    def sample_df(self):
        """테스트용 DataFrame"""
        return pd.DataFrame({
            'RCK Order no.': ['ND-0001', 'ND-0002', 'ND-0002', 'ND-0003'],
            'Customer name': ['고객A', '고객B', '고객B', '고객C'],
            'Item name': ['Item1', 'Item2-1', 'Item2-2', 'Item3'],
            'Item qty': [1, 2, 3, 4],
        })

    def test_single_item_returns_series(self, sample_df):
        """단일 아이템은 Series 반환"""
        result = find_order_data(sample_df, 'ND-0001')

        assert isinstance(result, pd.Series)
        assert result['Customer name'] == '고객A'
        assert result['Item qty'] == 1

    def test_multiple_items_returns_dataframe(self, sample_df):
        """다중 아이템은 DataFrame 반환"""
        result = find_order_data(sample_df, 'ND-0002')

        assert isinstance(result, pd.DataFrame)
        assert len(result) == 2
        assert list(result['Item name']) == ['Item2-1', 'Item2-2']

    def test_not_found_returns_none(self, sample_df):
        """없는 주문번호는 None 반환"""
        result = find_order_data(sample_df, 'ND-9999')

        assert result is None

    def test_empty_dataframe_returns_none(self):
        """빈 DataFrame에서 검색하면 None"""
        empty_df = pd.DataFrame({'RCK Order no.': []})
        result = find_order_data(empty_df, 'ND-0001')

        assert result is None

    def test_partial_match_not_found(self, sample_df):
        """부분 일치는 찾지 않음"""
        result = find_order_data(sample_df, 'ND-000')  # ND-0001의 일부

        assert result is None

    def test_case_sensitive_search(self, sample_df):
        """대소문자 구분"""
        result = find_order_data(sample_df, 'nd-0001')  # 소문자

        assert result is None


class TestEscapeExcelFormula:
    """escape_excel_formula 함수 테스트"""

    def test_equals_sign_escaped(self):
        """= 문자로 시작하면 이스케이프"""
        assert escape_excel_formula("=SUM(A1:A10)") == "'=SUM(A1:A10)"
        assert escape_excel_formula("=1+1") == "'=1+1"

    def test_plus_sign_escaped(self):
        """+ 문자로 시작하면 이스케이프"""
        assert escape_excel_formula("+82-10-1234-5678") == "'+82-10-1234-5678"
        assert escape_excel_formula("+cmd|calc") == "'+cmd|calc"

    def test_minus_sign_escaped(self):
        """- 문자로 시작하면 이스케이프"""
        assert escape_excel_formula("-1+1") == "'-1+1"
        assert escape_excel_formula("-@SUM(A1)") == "'-@SUM(A1)"

    def test_at_sign_escaped(self):
        """@ 문자로 시작하면 이스케이프"""
        assert escape_excel_formula("@SUM(A1:A10)") == "'@SUM(A1:A10)"
        assert escape_excel_formula("@user") == "'@user"

    def test_normal_string_not_escaped(self):
        """일반 문자열은 이스케이프하지 않음"""
        assert escape_excel_formula("Hello World") == "Hello World"
        assert escape_excel_formula("ND-0001") == "ND-0001"
        assert escape_excel_formula("고객명") == "고객명"

    def test_non_string_types_unchanged(self):
        """문자열이 아닌 타입은 그대로 반환"""
        assert escape_excel_formula(100) == 100
        assert escape_excel_formula(3.14) == 3.14
        assert escape_excel_formula(None) is None
        assert escape_excel_formula(['=SUM']) == ['=SUM']

    def test_empty_string_unchanged(self):
        """빈 문자열은 그대로 반환"""
        assert escape_excel_formula("") == ""

    def test_middle_formula_chars_not_escaped(self):
        """중간에 수식 문자가 있어도 이스케이프하지 않음"""
        assert escape_excel_formula("A=B+C") == "A=B+C"
        assert escape_excel_formula("test@example.com") == "test@example.com"
        assert escape_excel_formula("1+1=2") == "1+1=2"
