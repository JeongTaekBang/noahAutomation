"""
데이터 검증 모듈
================

주문 데이터의 유효성을 검증합니다.
- 필수 필드 검증
- ICO Unit 검증
- 납기일 검증
"""

from __future__ import annotations

import logging
from datetime import datetime
from typing import NamedTuple

import pandas as pd

from po_generator.config import REQUIRED_FIELDS, MIN_LEAD_TIME_DAYS
from po_generator.utils import get_safe_value

logger = logging.getLogger(__name__)


class ValidationResult(NamedTuple):
    """검증 결과"""
    warnings: list[str]
    errors: list[str]

    @property
    def has_errors(self) -> bool:
        return len(self.errors) > 0

    @property
    def has_warnings(self) -> bool:
        return len(self.warnings) > 0

    @property
    def is_valid(self) -> bool:
        return not self.has_errors


def validate_required_fields(order_data: pd.Series) -> list[str]:
    """필수 필드 검증

    Args:
        order_data: 주문 데이터

    Returns:
        오류 메시지 목록
    """
    errors = []
    for field in REQUIRED_FIELDS:
        value = get_safe_value(order_data, field)
        if not value:
            errors.append(f"필수 필드 누락: {field}")
            logger.error(f"필수 필드 누락: {field}")
    return errors


def validate_ico_unit(order_data: pd.Series) -> list[str]:
    """ICO Unit 검증

    Args:
        order_data: 주문 데이터

    Returns:
        오류 메시지 목록
    """
    errors = []
    ico_unit = get_safe_value(order_data, 'ICO Unit', 0)

    try:
        ico_unit = float(ico_unit)
        if ico_unit == 0:
            errors.append("ICO Unit이 0입니다. 가격을 확인하세요.")
            logger.error("ICO Unit이 0입니다.")
        elif ico_unit < 0:
            errors.append(f"ICO Unit이 음수입니다: {ico_unit}")
            logger.error(f"ICO Unit이 음수입니다: {ico_unit}")
    except (ValueError, TypeError) as e:
        errors.append(f"ICO Unit 값이 올바르지 않습니다: {ico_unit}")
        logger.error(f"ICO Unit 변환 실패: {e}")

    return errors


def validate_quantity(order_data: pd.Series) -> list[str]:
    """수량 검증

    Args:
        order_data: 주문 데이터

    Returns:
        오류 메시지 목록
    """
    errors = []
    qty = get_safe_value(order_data, 'Item qty', 0)

    try:
        qty = int(float(qty))
        if qty <= 0:
            errors.append(f"수량이 올바르지 않습니다: {qty}")
            logger.error(f"수량이 올바르지 않습니다: {qty}")
    except (ValueError, TypeError) as e:
        errors.append(f"수량 값이 올바르지 않습니다: {qty}")
        logger.error(f"수량 변환 실패: {e}")

    return errors


def validate_delivery_date(order_data: pd.Series) -> tuple[list[str], list[str]]:
    """납기일 검증

    Args:
        order_data: 주문 데이터

    Returns:
        (경고 목록, 오류 목록) 튜플
    """
    warnings_list: list[str] = []
    errors_list: list[str] = []

    requested_date = get_safe_value(order_data, 'Requested delivery date')

    if not requested_date:
        warnings_list.append("납기일(Requested delivery date)이 입력되지 않았습니다.")
        logger.warning("납기일이 입력되지 않았습니다.")
        return warnings_list, errors_list

    try:
        if isinstance(requested_date, datetime):
            delivery_date = requested_date
        else:
            delivery_date = pd.to_datetime(requested_date)

        today = datetime.now()
        days_until_delivery = (delivery_date - today).days

        if days_until_delivery < 0:
            errors_list.append(
                f"납기일이 과거입니다: {delivery_date.strftime('%Y-%m-%d')}"
            )
            logger.error(f"납기일이 과거입니다: {delivery_date}")
        elif days_until_delivery < MIN_LEAD_TIME_DAYS:
            warnings_list.append(
                f"납기일이 촉박합니다: {days_until_delivery}일 후 "
                f"({delivery_date.strftime('%Y-%m-%d')})"
            )
            logger.warning(f"납기일이 촉박합니다: {days_until_delivery}일 후")
    except (ValueError, TypeError) as e:
        warnings_list.append(f"납기일 형식을 확인하세요: {requested_date}")
        logger.warning(f"납기일 파싱 실패: {e}")

    return warnings_list, errors_list


def validate_order_data(order_data: pd.Series) -> ValidationResult:
    """주문 데이터 전체 검증

    Args:
        order_data: 주문 데이터 Series

    Returns:
        ValidationResult(warnings, errors)
    """
    warnings_list: list[str] = []
    errors_list: list[str] = []

    # 1. 필수 필드 검증
    errors_list.extend(validate_required_fields(order_data))

    # 2. ICO Unit 검증
    errors_list.extend(validate_ico_unit(order_data))

    # 3. 수량 검증
    errors_list.extend(validate_quantity(order_data))

    # 4. 납기일 검증
    date_warnings, date_errors = validate_delivery_date(order_data)
    warnings_list.extend(date_warnings)
    errors_list.extend(date_errors)

    return ValidationResult(warnings=warnings_list, errors=errors_list)


def validate_multiple_items(items_df: pd.DataFrame) -> ValidationResult:
    """다중 아이템 검증

    Args:
        items_df: 다중 아이템 DataFrame

    Returns:
        ValidationResult(warnings, errors)
    """
    all_warnings: list[str] = []
    all_errors: list[str] = []

    for idx, (_, item) in enumerate(items_df.iterrows()):
        result = validate_order_data(item)

        for warn in result.warnings:
            all_warnings.append(f"[아이템 {idx + 1}] {warn}")

        for err in result.errors:
            all_errors.append(f"[아이템 {idx + 1}] {err}")

    return ValidationResult(warnings=all_warnings, errors=all_errors)
