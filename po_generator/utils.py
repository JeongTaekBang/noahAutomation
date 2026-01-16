"""
유틸리티 함수
=============

데이터 로드, 값 추출 등 공통 유틸리티 함수를 제공합니다.
"""

from __future__ import annotations

import logging
from typing import Any, Union

import pandas as pd

from po_generator.config import (
    NOAH_PO_LISTS_FILE,
    DOMESTIC_SHEET_INDEX,
    EXPORT_SHEET_INDEX,
    COLUMN_ALIASES,
)

logger = logging.getLogger(__name__)

# Excel 수식 인젝션 방지용 문자
FORMULA_ESCAPE_CHARS: tuple[str, ...] = ('=', '+', '-', '@')


def escape_excel_formula(value: Any) -> Any:
    """Excel 수식 인젝션 방지

    사용자 입력이 =, +, -, @로 시작하면 Excel에서 수식으로 해석될 수 있음.
    이를 방지하기 위해 앞에 작은따옴표(')를 추가.

    Args:
        value: 원본 값

    Returns:
        이스케이프된 값 (문자열인 경우만 처리)
    """
    if isinstance(value, str) and value and value[0] in FORMULA_ESCAPE_CHARS:
        return "'" + value
    return value


def get_safe_value(
    order_data: pd.Series,
    key: str,
    default: Any = ''
) -> Any:
    """안전하게 값 가져오기 (NaN 처리)

    Args:
        order_data: 주문 데이터 Series
        key: 가져올 컬럼 키
        default: 기본값 (값이 없거나 NaN인 경우)

    Returns:
        해당 키의 값 또는 기본값
    """
    value = order_data.get(key, default)

    # None 또는 pandas NaN/NA 체크
    if value is None or pd.isna(value):
        return default

    # 문자열 'nan' 체크 (대소문자 무관) - pandas가 가끔 이런 문자열을 반환
    if isinstance(value, str) and value.strip().lower() == 'nan':
        return default

    return value


def resolve_column(
    columns: Union[pd.Index, list[str]],
    key: str,
) -> str | None:
    """별칭에서 실제 컬럼명 찾기

    Args:
        columns: DataFrame의 컬럼 목록 (df.columns)
        key: 내부 키 (예: 'customer_name') 또는 실제 컬럼명

    Returns:
        실제 컬럼명 또는 None (찾지 못한 경우)
    """
    # 1. key가 이미 실제 컬럼명인 경우
    if key in columns:
        return key

    # 2. key가 내부 키인 경우, 별칭에서 찾기
    aliases = COLUMN_ALIASES.get(key)
    if aliases:
        for alias in aliases:
            if alias in columns:
                return alias

    # 3. 대소문자 무시 검색 (fallback)
    key_lower = key.lower()
    for col in columns:
        if col.lower() == key_lower:
            return col

    return None


def get_value(
    order_data: pd.Series,
    key: str,
    default: Any = '',
) -> Any:
    """내부 키로 값 가져오기 (별칭 지원)

    Args:
        order_data: 주문 데이터 Series
        key: 내부 키 (예: 'customer_name') 또는 실제 컬럼명
        default: 기본값 (값이 없거나 NaN인 경우)

    Returns:
        해당 키의 값 또는 기본값
    """
    # 실제 컬럼명 찾기
    actual_col = resolve_column(order_data.index, key)

    if actual_col is None:
        return default

    return get_safe_value(order_data, actual_col, default)


def load_noah_po_lists() -> pd.DataFrame:
    """NOAH_PO_Lists.xlsx에서 데이터 로드

    국내/해외 시트를 모두 로드하여 하나의 DataFrame으로 합칩니다.

    Returns:
        합쳐진 주문 데이터 DataFrame

    Raises:
        FileNotFoundError: 소스 파일이 없는 경우
    """
    if not NOAH_PO_LISTS_FILE.exists():
        raise FileNotFoundError(f"소스 파일을 찾을 수 없습니다: {NOAH_PO_LISTS_FILE}")

    logger.info(f"데이터 로드: {NOAH_PO_LISTS_FILE.name}")

    xl = pd.ExcelFile(NOAH_PO_LISTS_FILE)

    # 국내/해외 시트 모두 로드
    df_domestic = pd.read_excel(xl, sheet_name=DOMESTIC_SHEET_INDEX)
    df_export = pd.read_excel(xl, sheet_name=EXPORT_SHEET_INDEX)

    # 시트 구분 컬럼 추가
    df_domestic['_시트구분'] = '국내'
    df_export['_시트구분'] = '해외'

    # 컬럼 통일 (없는 컬럼은 NaN으로)
    all_columns = list(set(df_domestic.columns) | set(df_export.columns))
    for col in all_columns:
        if col not in df_domestic.columns:
            df_domestic[col] = pd.NA
        if col not in df_export.columns:
            df_export[col] = pd.NA

    df = pd.concat([df_domestic, df_export], ignore_index=True)

    logger.info(f"총 {len(df)}건의 주문 데이터 로드 완료")
    return df


def find_order_data(
    df: pd.DataFrame,
    order_no: str
) -> Union[pd.Series, pd.DataFrame, None]:
    """RCK Order No.로 주문 데이터 검색

    Args:
        df: 전체 주문 데이터
        order_no: 검색할 RCK Order No.

    Returns:
        - 단일 아이템: pd.Series
        - 다중 아이템: pd.DataFrame
        - 없음: None
    """
    # 컬럼 별칭으로 실제 컬럼명 찾기
    order_no_col = resolve_column(df.columns, 'order_no')
    if order_no_col is None:
        logger.error("주문번호 컬럼을 찾을 수 없습니다.")
        return None

    mask = df[order_no_col] == order_no
    match_count = mask.sum()

    if match_count == 0:
        logger.warning(f"주문번호 '{order_no}'를 찾을 수 없습니다.")
        return None

    if match_count > 1:
        logger.info(f"주문번호 '{order_no}': {match_count}개 아이템 발견 (다중 아이템)")
        return df[mask]  # DataFrame

    logger.info(f"주문번호 '{order_no}': 단일 아이템")
    return df[mask].iloc[0]  # Series


def format_currency(value: float, currency: str = 'KRW') -> str:
    """통화 포맷팅

    Args:
        value: 금액
        currency: 통화 코드 (KRW 또는 USD)

    Returns:
        포맷팅된 통화 문자열 (예: "₩1,000,000" 또는 "$1,000.00")
    """
    if currency == 'KRW':
        return f"₩{value:,.0f}"
    return f"${value:,.2f}"
