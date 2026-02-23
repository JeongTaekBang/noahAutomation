"""
거래명세표 생성 모듈 (xlwings 기반)
====================================

xlwings를 사용하여 템플릿 기반으로 거래명세표를 생성합니다.
이미지, 서식 등이 완벽하게 보존됩니다.

지원 문서 유형:
- DN: 납품 거래명세표 (DN_국내 데이터 사용)
- PMT: 선수금 거래명세표 (PMT_국내 데이터 사용)
"""

from __future__ import annotations

import logging
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.config import (
    TS_TEMPLATE_FILE,
    ITEM_START_ROW_FALLBACK,
    VAT_RATE_DOMESTIC,
)
from po_generator.utils import get_value
from po_generator.excel_helpers import (
    XlConstants,
    xlwings_app_context,
    prepare_template,
    cleanup_temp_file,
    find_item_start_row_xlwings,
    TS_HEADER_LABELS,
    batch_write_rows,
    delete_rows_range,
    find_text_in_column_batch,
)

logger = logging.getLogger(__name__)


# === 셀 매핑 (템플릿 기준) ===
CELL_DATE = 'B2'           # DATE : 날짜
CELL_CUSTOMER = 'B7'       # 고객명 귀하

# 레이블 검색 범위 (행 삽입 후에도 이 범위 내에 있음)
LABEL_SEARCH_START = 15
LABEL_SEARCH_END = 50

# 기본 행 위치 (동적 탐지 실패 시 폴백값)
# ITEM_START_ROW는 config의 ITEM_START_ROW_FALLBACK 사용
BASE_PO_ROW = 23           # PO No. 행 (폴백값)
BASE_TOTAL_ROW = 25        # 합계 행 (폴백값)


def create_ts_xlwings(
    template_path: Path,
    output_path: Path,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
    doc_type: str = 'DN',
) -> None:
    """xlwings로 거래명세표 생성

    Args:
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        order_data: 주문 데이터 (첫 번째 아이템 또는 단일 아이템)
        items_df: 다중 아이템인 경우 전체 아이템 DataFrame
        doc_type: 문서 유형 ('DN' 또는 'PMT')
    """
    # 템플릿 준비 (임시 폴더로 복사)
    temp_template, temp_output = prepare_template(template_path, "ts")

    # 출고일 가져오기 (없으면 오늘 날짜 사용)
    dispatch_date = get_value(order_data, 'dispatch_date', None)
    if dispatch_date is None or pd.isna(dispatch_date):
        dispatch_date = datetime.now()
    elif not isinstance(dispatch_date, datetime):
        try:
            dispatch_date = pd.to_datetime(dispatch_date)
        except (ValueError, TypeError):
            dispatch_date = datetime.now()
    dispatch_date_str = dispatch_date.strftime("%Y. %m. %d")

    try:
        # xlwings App 생명주기 관리
        with xlwings_app_context() as app:
            # 임시 템플릿 열기
            wb = app.books.open(str(temp_template))
            ws = wb.sheets[0]

            # 1. 헤더 정보 (출고일 사용)
            ws.range(CELL_DATE).value = f"DATE : {dispatch_date_str}"
            customer_name = get_value(order_data, 'customer_name', '')
            ws.range(CELL_CUSTOMER).value = f"{customer_name} 귀하"

            # DN/ADV 공통 처리 (remark만 다름)
            remark = '선수금' if doc_type == 'ADV' else ''
            _fill_ts_data(ws, order_data, items_df, dispatch_date, remark)

            # 임시 위치에 저장
            wb.save(str(temp_output))
            logger.info(f"거래명세표 생성 완료 (임시): {temp_output}")

    finally:
        # 임시 템플릿 삭제
        cleanup_temp_file(temp_template)

    # 최종 출력 경로로 이동
    shutil.move(str(temp_output), str(output_path))
    logger.info(f"거래명세표 저장 완료: {output_path}")


def _find_ts_subtotal_row(ws: xw.Sheet, start_row: int, max_search: int = 15) -> int:
    """소계(SUM) 수식이 있는 행 찾기 (배치 읽기 최적화)

    Args:
        ws: xlwings Sheet 객체
        start_row: 검색 시작 행
        max_search: 최대 검색 행 수

    Returns:
        소계 행 번호 (못 찾으면 start_row + 3)
    """
    # 배치 읽기로 15회 COM 호출 → 1회로 감소
    end_row = start_row + max_search - 1
    formulas = ws.range(f'E{start_row}:E{end_row}').formula

    # xlwings 범위 읽기는 tuple of tuples 반환: (('val1',), ('val2',), ...)
    # 단일 셀은 문자열 반환
    if isinstance(formulas, (list, tuple)) and formulas and isinstance(formulas[0], (list, tuple)):
        # 2D → 1D 평탄화 (각 행의 첫 번째 값만 추출)
        formulas = [f[0] if f else '' for f in formulas]
    elif not isinstance(formulas, (list, tuple)):
        formulas = [formulas]

    for idx, formula in enumerate(formulas):
        if formula and '=SUM' in str(formula):
            return start_row + idx

    return start_row + 3  # 기본값


def _restore_ts_item_borders(ws: xw.Sheet, item_start_row: int, num_items: int) -> None:
    """행 삭제 후 아이템 영역 테두리 복원

    Args:
        ws: xlwings Sheet 객체
        item_start_row: 아이템 시작 행
        num_items: 실제 아이템 수
    """
    # 마지막 아이템 행 (소계 바로 위)
    last_item_row = item_start_row + num_items - 1

    # 헤더 아래 행 (첫 번째 아이템 행 바로 위)의 아래 테두리
    header_bottom_row = item_start_row - 1
    ws.range(f'A{header_bottom_row}:H{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{header_bottom_row}:H{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    # 마지막 아이템 행의 아래 테두리
    ws.range(f'A{last_item_row}:H{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{last_item_row}:H{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    logger.debug(f"테두리 복원: Row {header_bottom_row} 하단, Row {last_item_row} 하단")


def _fill_ts_data(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None,
    dispatch_date: datetime,
    remark: str = '',
) -> None:
    """거래명세표 데이터 채우기 (DN/ADV 공통) - 배치 쓰기 최적화

    Args:
        ws: xlwings Sheet 객체
        order_data: 주문 데이터 (첫 번째 아이템)
        items_df: 다중 아이템인 경우 DataFrame
        dispatch_date: 출고일
        remark: 비고 텍스트 (예: '선수금')
    """
    # 아이템 준비
    if items_df is None:
        items_df = pd.DataFrame([order_data])
    num_items = len(items_df)

    # 아이템 시작 행 동적 탐지
    item_start_row = find_item_start_row_xlwings(
        ws,
        search_labels=TS_HEADER_LABELS,
        columns=('A', 'B', 'C', 'D'),
        fallback_row=ITEM_START_ROW_FALLBACK,
    )

    # 템플릿의 기존 아이템 행 수 계산 (소계 행 찾기)
    subtotal_row = _find_ts_subtotal_row(ws, item_start_row)
    template_item_count = subtotal_row - item_start_row
    logger.debug(f"템플릿 아이템 수: {template_item_count}, 실제 아이템 수: {num_items}")

    # 행 수 조정: 템플릿 예시보다 실제 아이템이 적으면 초과 행 삭제
    if num_items < template_item_count:
        rows_to_delete = template_item_count - num_items
        # 범위 삭제로 N회 COM 호출 → 1회로 감소
        delete_rows_range(ws, item_start_row + num_items, rows_to_delete)

        # 테두리 복원: 행 삭제로 사라진 테두리 다시 그리기
        _restore_ts_item_borders(ws, item_start_row, num_items)

    # 행 수 조정: 템플릿 예시보다 실제 아이템이 많으면 행 삽입
    elif num_items > template_item_count:
        rows_to_insert = num_items - template_item_count
        source_row = item_start_row
        for i in range(rows_to_insert):
            insert_row = item_start_row + template_item_count + i
            ws.range(f'{source_row}:{source_row}').api.Copy()
            ws.range(f'{insert_row}:{insert_row}').api.Insert(Shift=XlConstants.xlShiftDown)
        logger.debug(f"{rows_to_insert}개 행 삽입")

    # 기존 아이템 행 데이터 초기화 (서식은 유지) - 배치 초기화
    end_row = item_start_row + num_items - 1
    ws.range(f'A{item_start_row}:H{end_row}').value = None

    # 아이템 데이터 배치 쓰기 (N개 아이템 * 8열 COM 호출 → 1회로 감소)
    total_amount, total_tax = _fill_items_batch(ws, item_start_row, items_df, dispatch_date, remark)

    # 소계 행 수식 업데이트 (다중 아이템인 경우)
    subtotal_row = item_start_row + num_items
    if num_items > 1:
        last_item_row = item_start_row + num_items - 1
        ws.range(f'E{subtotal_row}').formula = f'=SUM(E{item_start_row}:E{last_item_row})'
        ws.range(f'G{subtotal_row}').formula = f'=SUM(G{item_start_row}:G{last_item_row})'
        ws.range(f'H{subtotal_row}').formula = f'=SUM(H{item_start_row}:H{last_item_row})'

    # PO No. 채우기 (레이블 위치를 찾아서 같은 행에 값 입력)
    customer_po = get_value(order_data, 'customer_po', '')
    po_row = find_text_in_column_batch(ws, 'A', 'PO No', LABEL_SEARCH_START, LABEL_SEARCH_END)
    if po_row is None:
        po_row = BASE_PO_ROW + (num_items - 1) if num_items > 1 else BASE_PO_ROW
    ws.range(f'B{po_row}').value = customer_po

    # 합계 채우기 (레이블 위치를 찾아서 같은 행에 값 입력)
    grand_total = total_amount + total_tax
    total_row = find_text_in_column_batch(ws, 'E', '합 계', LABEL_SEARCH_START, LABEL_SEARCH_END)
    if total_row is None:
        total_row = BASE_TOTAL_ROW + (num_items - 1) if num_items > 1 else BASE_TOTAL_ROW
    ws.range(f'G{total_row}').value = grand_total


def _fill_items_batch(
    ws: xw.Sheet,
    item_start_row: int,
    items_df: pd.DataFrame,
    dispatch_date: datetime,
    remark: str = '',
) -> tuple[int, int]:
    """아이템 데이터 배치 쓰기 (성능 최적화)

    N개 아이템 * 8열 = N*8회 COM 호출 → 1회로 감소

    Args:
        ws: xlwings Sheet 객체
        item_start_row: 아이템 시작 행
        items_df: 아이템 DataFrame
        dispatch_date: 출고일
        remark: 비고 텍스트

    Returns:
        (총 금액, 총 세액)
    """
    data_2d = []
    total_amount = 0
    total_tax = 0
    date_str = f"{dispatch_date.month}/{dispatch_date.day}"

    for item_idx, (_, item) in enumerate(items_df.iterrows()):
        # 수량
        raw_qty = get_value(item, 'item_qty', 1)
        try:
            qty = int(raw_qty) if pd.notna(raw_qty) else 1
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 수량 변환 실패 '{raw_qty}' -> 기본값 1 사용")
            qty = 1

        # 단가
        raw_price = get_value(item, 'sales_unit_price', 0)
        try:
            unit_price = int(float(raw_price)) if pd.notna(raw_price) else 0
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 단가 변환 실패 '{raw_price}' -> 기본값 0 사용")
            unit_price = 0

        # 금액/세액 계산
        amount = qty * unit_price
        tax = int(amount * VAT_RATE_DOMESTIC)
        total_amount += amount
        total_tax += tax

        # 행 데이터: A(월/일), B(품명), C(비고), D(규격), E(수량), F(단가), G(금액), H(세액)
        data_2d.append([
            date_str,
            get_value(item, 'item_name', ''),
            remark,
            "EA",
            qty,
            unit_price,
            amount,
            tax,
        ])

    # 한 번에 쓰기
    batch_write_rows(ws, f'A{item_start_row}', data_2d)

    return total_amount, total_tax


