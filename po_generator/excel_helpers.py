"""
Excel 헬퍼 함수 모듈
====================

여러 모듈에서 공통으로 사용하는 Excel 관련 유틸리티 함수를 제공합니다.

- find_item_start_row_openpyxl: openpyxl 워크시트용
- find_item_start_row_xlwings: xlwings 워크시트용
"""

from __future__ import annotations

import logging
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet as OpenpyxlWorksheet
    import xlwings as xw

from po_generator.config import ITEM_START_ROW_FALLBACK

logger = logging.getLogger(__name__)


# === 헤더 라벨 프리셋 ===

# Purchase Order 헤더 라벨 (PO, 이력 추출용)
PO_HEADER_LABELS: tuple[str, ...] = (
    'No.',
    'Item Number',
    'Item\nNumber',
    'Item',
)

# 거래명세표 헤더 라벨
TS_HEADER_LABELS: tuple[str, ...] = (
    '월/일',
    '품명',
    'DESCRIPTION',
)

# Proforma Invoice 헤더 라벨
PI_HEADER_LABELS: tuple[str, ...] = (
    'No.',
    'Description',
    'DESCRIPTION',
)

# 기본 헤더 라벨 (모든 문서 유형에 공통)
DEFAULT_HEADER_LABELS: tuple[str, ...] = (
    'No.',
    'Item Number',
    'Item\nNumber',
    '품명',
    'Item',
)


def find_item_start_row_openpyxl(
    ws: OpenpyxlWorksheet,
    search_labels: tuple[str, ...] = DEFAULT_HEADER_LABELS,
    max_search_rows: int = 30,
    max_search_cols: int = 9,
    fallback_row: int = ITEM_START_ROW_FALLBACK,
) -> int:
    """템플릿에서 아이템 시작 행을 동적으로 찾기 (openpyxl 버전)

    헤더 레이블을 찾아서 그 다음 행이 아이템 시작 위치입니다.

    Args:
        ws: openpyxl Worksheet 객체
        search_labels: 검색할 헤더 레이블
        max_search_rows: 최대 검색 행 수
        max_search_cols: 최대 검색 열 수 (기본: 9, A-I)
        fallback_row: 헤더를 찾지 못했을 때 기본값

    Returns:
        아이템 시작 행 번호
    """
    for row in range(1, max_search_rows + 1):
        for col in range(1, max_search_cols + 1):
            cell_value = ws.cell(row=row, column=col).value
            if cell_value and any(
                label in str(cell_value) for label in search_labels
            ):
                logger.debug(
                    f"헤더 발견: Row {row}, 값='{cell_value}' -> 아이템 시작 Row {row + 1}"
                )
                return row + 1  # 레이블 다음 행이 데이터 시작

    logger.debug(f"헤더를 찾지 못함 -> 기본값 Row {fallback_row} 사용")
    return fallback_row


def find_item_start_row_xlwings(
    ws: xw.Sheet,
    search_labels: tuple[str, ...] = DEFAULT_HEADER_LABELS,
    max_search_rows: int = 30,
    columns: tuple[str, ...] = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'),
    fallback_row: int = ITEM_START_ROW_FALLBACK,
) -> int:
    """템플릿에서 아이템 시작 행을 동적으로 찾기 (xlwings 버전)

    헤더 레이블을 찾아서 그 다음 행이 아이템 시작 위치입니다.

    Args:
        ws: xlwings Sheet 객체
        search_labels: 검색할 헤더 레이블
        max_search_rows: 최대 검색 행 수
        columns: 검색할 열 문자 튜플
        fallback_row: 헤더를 찾지 못했을 때 기본값

    Returns:
        아이템 시작 행 번호
    """
    for row in range(1, max_search_rows + 1):
        for col in columns:
            cell_value = ws.range(f'{col}{row}').value
            if cell_value and any(
                label in str(cell_value) for label in search_labels
            ):
                logger.debug(
                    f"헤더 발견: Row {row}, 값='{cell_value}' -> 아이템 시작 Row {row + 1}"
                )
                return row + 1

    logger.debug(f"헤더를 찾지 못함 -> 기본값 Row {fallback_row} 사용")
    return fallback_row


# === 하위 호환성을 위한 별칭 ===
# 기존 코드에서 직접 사용하던 함수명 지원

def find_item_start_row(
    ws,
    search_labels: tuple[str, ...] = DEFAULT_HEADER_LABELS,
    max_search_rows: int = 30,
    fallback_row: int = ITEM_START_ROW_FALLBACK,
) -> int:
    """템플릿에서 아이템 시작 행을 동적으로 찾기 (자동 감지)

    워크시트 타입을 자동으로 감지하여 적절한 함수를 호출합니다.
    openpyxl.worksheet.worksheet.Worksheet인 경우 openpyxl 버전을,
    그 외의 경우 xlwings 버전을 사용합니다.

    Args:
        ws: openpyxl 또는 xlwings Worksheet 객체
        search_labels: 검색할 헤더 레이블
        max_search_rows: 최대 검색 행 수
        fallback_row: 헤더를 찾지 못했을 때 기본값

    Returns:
        아이템 시작 행 번호
    """
    # openpyxl인지 xlwings인지 자동 감지
    # openpyxl Worksheet는 cell() 메서드를 가지고 있음
    if hasattr(ws, 'cell') and callable(getattr(ws, 'cell', None)):
        return find_item_start_row_openpyxl(
            ws,
            search_labels=search_labels,
            max_search_rows=max_search_rows,
            fallback_row=fallback_row,
        )
    else:
        return find_item_start_row_xlwings(
            ws,
            search_labels=search_labels,
            max_search_rows=max_search_rows,
            fallback_row=fallback_row,
        )
