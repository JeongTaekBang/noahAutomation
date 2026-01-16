"""
validators 모듈 테스트
"""

import pandas as pd
import pytest

from po_generator.validators import (
    validate_order_data,
    validate_multiple_items,
    validate_required_fields,
    validate_ico_unit,
    validate_quantity,
    validate_delivery_date,
    ValidationResult,
)


class TestValidateRequiredFields:
    """필수 필드 검증 테스트"""

    def test_all_fields_present(self, valid_order_data: pd.Series):
        """모든 필수 필드가 있는 경우"""
        errors = validate_required_fields(valid_order_data)
        assert len(errors) == 0

    def test_missing_customer_name(self, invalid_order_data_missing_fields: pd.Series):
        """Customer name 누락"""
        errors = validate_required_fields(invalid_order_data_missing_fields)
        assert any('Customer name' in e for e in errors)

    def test_missing_model(self, invalid_order_data_missing_fields: pd.Series):
        """Model 누락"""
        errors = validate_required_fields(invalid_order_data_missing_fields)
        assert any('Model' in e for e in errors)


class TestValidateIcoUnit:
    """ICO Unit 검증 테스트"""

    def test_valid_ico_unit(self, valid_order_data: pd.Series):
        """유효한 ICO Unit"""
        errors = validate_ico_unit(valid_order_data)
        assert len(errors) == 0

    def test_zero_ico_unit(self, invalid_order_data_zero_ico: pd.Series):
        """ICO Unit이 0인 경우"""
        errors = validate_ico_unit(invalid_order_data_zero_ico)
        assert len(errors) == 1
        assert '0' in errors[0]

    def test_negative_ico_unit(self):
        """ICO Unit이 음수인 경우"""
        data = pd.Series({'ICO Unit': -1000})
        errors = validate_ico_unit(data)
        assert len(errors) == 1
        assert '음수' in errors[0]

    def test_invalid_ico_unit_string(self):
        """ICO Unit이 변환 불가능한 문자열인 경우"""
        data = pd.Series({'ICO Unit': 'invalid'})
        errors = validate_ico_unit(data)
        assert len(errors) == 1
        assert '올바르지 않습니다' in errors[0]

    def test_invalid_ico_unit_none(self):
        """ICO Unit이 None인 경우 (기본값 0 적용)"""
        data = pd.Series({'ICO Unit': None})
        errors = validate_ico_unit(data)
        # None은 get_safe_value에서 기본값 0으로 변환되므로 0 오류 발생
        assert len(errors) == 1
        assert '0' in errors[0]


class TestValidateQuantity:
    """수량 검증 테스트"""

    def test_valid_quantity(self, valid_order_data: pd.Series):
        """유효한 수량"""
        errors = validate_quantity(valid_order_data)
        assert len(errors) == 0

    def test_zero_quantity(self):
        """수량이 0인 경우"""
        data = pd.Series({'Item qty': 0})
        errors = validate_quantity(data)
        assert len(errors) == 1

    def test_negative_quantity(self):
        """수량이 음수인 경우"""
        data = pd.Series({'Item qty': -5})
        errors = validate_quantity(data)
        assert len(errors) == 1

    def test_invalid_quantity_string(self):
        """수량이 변환 불가능한 문자열인 경우"""
        data = pd.Series({'Item qty': 'many'})
        errors = validate_quantity(data)
        assert len(errors) == 1
        assert '올바르지 않습니다' in errors[0]

    def test_float_quantity_converted_to_int(self):
        """소수점 수량은 정수로 변환"""
        data = pd.Series({'Item qty': 2.7})
        errors = validate_quantity(data)
        # 2.7 -> 2 로 변환되어 유효함
        assert len(errors) == 0


class TestValidateDeliveryDate:
    """납기일 검증 테스트"""

    def test_valid_delivery_date(self, valid_order_data: pd.Series):
        """유효한 납기일 (30일 후)"""
        warnings, errors = validate_delivery_date(valid_order_data)
        assert len(errors) == 0
        assert len(warnings) == 0

    def test_past_delivery_date(self, invalid_order_data_past_delivery: pd.Series):
        """납기일이 과거인 경우"""
        warnings, errors = validate_delivery_date(invalid_order_data_past_delivery)
        assert len(errors) == 1
        assert '과거' in errors[0]

    def test_urgent_delivery_date(self, order_data_urgent_delivery: pd.Series):
        """납기일이 촉박한 경우 (7일 이내)"""
        warnings, errors = validate_delivery_date(order_data_urgent_delivery)
        assert len(errors) == 0
        assert len(warnings) == 1
        assert '촉박' in warnings[0]

    def test_missing_delivery_date(self):
        """납기일 미입력"""
        data = pd.Series({'Requested delivery date': None})
        warnings, errors = validate_delivery_date(data)
        assert len(warnings) == 1
        assert '입력되지 않았습니다' in warnings[0]

    def test_string_delivery_date_valid(self):
        """문자열 납기일 파싱 (유효한 형식)"""
        data = pd.Series({'Requested delivery date': '2099-12-31'})
        warnings, errors = validate_delivery_date(data)
        assert len(errors) == 0
        assert len(warnings) == 0

    def test_string_delivery_date_invalid_format(self):
        """문자열 납기일 파싱 실패 (잘못된 형식)"""
        data = pd.Series({'Requested delivery date': 'not-a-date'})
        warnings, errors = validate_delivery_date(data)
        assert len(warnings) == 1
        assert '형식을 확인하세요' in warnings[0]

    def test_string_delivery_date_korean_format(self):
        """한국어 형식 납기일"""
        data = pd.Series({'Requested delivery date': '2099년 12월 31일'})
        warnings, errors = validate_delivery_date(data)
        # pandas는 이 형식을 파싱하지 못할 수 있음
        # 파싱 실패 시 경고 발생
        assert len(errors) == 0  # 오류는 아님


class TestValidationResult:
    """ValidationResult 클래스 테스트"""

    def test_has_warnings_true(self):
        """경고가 있는 경우"""
        result = ValidationResult(warnings=['경고 메시지'], errors=[])
        assert result.has_warnings is True

    def test_has_warnings_false(self):
        """경고가 없는 경우"""
        result = ValidationResult(warnings=[], errors=[])
        assert result.has_warnings is False

    def test_has_errors_true(self):
        """오류가 있는 경우"""
        result = ValidationResult(warnings=[], errors=['오류 메시지'])
        assert result.has_errors is True

    def test_is_valid_with_only_warnings(self):
        """경고만 있고 오류가 없으면 유효함"""
        result = ValidationResult(warnings=['경고'], errors=[])
        assert result.is_valid is True
        assert result.has_warnings is True


class TestValidateOrderData:
    """전체 주문 데이터 검증 테스트"""

    def test_valid_order(self, valid_order_data: pd.Series):
        """유효한 주문"""
        result = validate_order_data(valid_order_data)
        assert isinstance(result, ValidationResult)
        assert result.is_valid
        assert not result.has_errors

    def test_multiple_errors(self, invalid_order_data_missing_fields: pd.Series):
        """여러 오류가 있는 경우"""
        result = validate_order_data(invalid_order_data_missing_fields)
        assert not result.is_valid
        assert result.has_errors
        assert len(result.errors) >= 2  # Customer name, Model 누락


class TestValidateMultipleItems:
    """다중 아이템 검증 테스트"""

    def test_valid_multiple_items(self, multiple_items_df: pd.DataFrame):
        """유효한 다중 아이템"""
        result = validate_multiple_items(multiple_items_df)
        assert result.is_valid

    def test_item_prefix_in_error_messages(self, multiple_items_df: pd.DataFrame):
        """오류 메시지에 아이템 번호 접두사"""
        # ICO Unit을 0으로 변경하여 오류 발생
        df = multiple_items_df.copy()
        df.loc[df.index[0], 'ICO Unit'] = 0

        result = validate_multiple_items(df)
        assert any('[아이템 1]' in e for e in result.errors)

    def test_item_prefix_in_warning_messages(self, multiple_items_df: pd.DataFrame):
        """경고 메시지에 아이템 번호 접두사"""
        # 납기일 제거하여 경고 발생
        df = multiple_items_df.copy()
        df.loc[df.index[1], 'Requested delivery date'] = None

        result = validate_multiple_items(df)
        assert any('[아이템 2]' in w for w in result.warnings)

    def test_multiple_items_multiple_errors(self, multiple_items_df: pd.DataFrame):
        """여러 아이템에서 여러 오류 발생"""
        df = multiple_items_df.copy()
        df.loc[df.index[0], 'ICO Unit'] = 0
        df.loc[df.index[1], 'Item qty'] = -1

        result = validate_multiple_items(df)
        assert any('[아이템 1]' in e for e in result.errors)
        assert any('[아이템 2]' in e for e in result.errors)
