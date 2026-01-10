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

    def test_item_prefix_in_messages(self, multiple_items_df: pd.DataFrame):
        """오류 메시지에 아이템 번호 접두사"""
        # ICO Unit을 0으로 변경하여 오류 발생
        df = multiple_items_df.copy()
        df.loc[df.index[0], 'ICO Unit'] = 0

        result = validate_multiple_items(df)
        assert any('[아이템 1]' in e for e in result.errors)
