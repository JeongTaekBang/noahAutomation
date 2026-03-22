"""
Order Confirmation 생성 모듈 (xlwings 기반)
============================================

Final Invoice와 동일한 레이아웃에 Dispatch date (H열), Shipping method (H11) 컬럼이 추가된 형태입니다.
Dispatch date는 SO_해외의 'EXW NOAH', Shipping method는 SO_해외의 'Shipping method' 컬럼 값을 사용합니다.

SO_해외 + Customer_해외 데이터를 사용합니다.
"""

from __future__ import annotations

import logging
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.utils import get_value
from po_generator.excel_helpers import (
    XlConstants,
    xlwings_app_context,
    prepare_template,
    cleanup_temp_file,
    delete_rows_range,
    find_text_in_column_batch,
)

logger = logging.getLogger(__name__)


def _to_text(value) -> str:
    """숫자를 문자열로 변환 (앞 0 보존, 뒤 .0 제거)"""
    if pd.isna(value) or value == '':
        return ''
    if isinstance(value, str):
        return value
    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)
    return str(value)


# === 셀 매핑 (Order Confirmation) ===
# Header (FI와 동일)
CELL_PO_NO = 'C7'
CELL_INVOICE_NO = 'H7'
CELL_PO_DATE = 'C8'
CELL_INVOICE_DATE = 'H8'
CELL_PAYMENT_TERMS = 'H9'
CELL_DELIVERY_TERMS = 'H10'
CELL_SHIPPING_METHOD = 'H11'
CELL_CUST_ADDR_1 = 'A13'
CELL_CUST_ADDR_2 = 'A14'
CELL_CUST_ADDR_3 = 'A15'
CELL_DELV_ADDR_1 = 'G13'
CELL_DELV_ADDR_2 = 'G14'
CELL_DELV_ADDR_3 = 'G15'

# 아이템 (헤더 Row 17, 데이터 Row 18~)
ITEM_START_ROW = 18

COL_ITEM_NAME = 'A'     # 품목명 (A:D 병합)
COL_QTY = 'E'           # 수량
COL_UNIT_PRICE = 'F'    # 단가
COL_CURRENCY = 'G'      # 통화
COL_DISPATCH = 'H'      # Dispatch date (OC 신규)
COL_AMOUNT = 'I'        # 금액


def create_oc_xlwings(
    template_path: Path,
    output_path: Path,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> None:
    """xlwings로 Order Confirmation 생성

    Args:
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        order_data: 주문 데이터 (첫 번째 아이템 또는 단일 아이템)
        items_df: 다중 아이템인 경우 전체 아이템 DataFrame
    """
    temp_template, temp_output = prepare_template(template_path, "oc")

    try:
        with xlwings_app_context() as app:
            wb = app.books.open(str(temp_template))
            ws = wb.sheets[0]

            _fill_header(ws, order_data)
            _fill_items(ws, order_data, items_df)

            wb.save(str(temp_output))
            logger.info(f"Order Confirmation 생성 완료 (임시): {temp_output}")

    finally:
        cleanup_temp_file(temp_template)

    shutil.move(str(temp_output), str(output_path))
    logger.info(f"Order Confirmation 저장 완료: {output_path}")


def _fill_header(ws: xw.Sheet, order_data: pd.Series) -> None:
    """헤더 정보 채우기 (SO_해외 + Customer_해외 기반)"""
    # Customer PO → SO_해외.Customer PO
    ws.range(CELL_PO_NO).value = get_value(order_data, 'customer_po', '')

    # Invoice No → SO_ID
    so_id = get_value(order_data, 'so_id', '')
    ws.range(CELL_INVOICE_NO).value = so_id

    # PO Date → SO_해외.PO receipt date
    po_date = get_value(order_data, 'po_receipt_date', '')
    if po_date and pd.notna(po_date):
        if isinstance(po_date, datetime):
            ws.range(CELL_PO_DATE).value = po_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_PO_DATE).value = str(po_date)

    # Invoice Date → 오늘 날짜 (OC 발행일)
    ws.range(CELL_INVOICE_DATE).value = datetime.now().strftime("%Y-%m-%d")

    # Payment Terms → Customer_해외.Payment terms
    payment_terms = get_value(order_data, 'payment_terms', '')
    if payment_terms:
        ws.range(CELL_PAYMENT_TERMS).value = payment_terms

    # Delivery Terms → SO_해외.Incoterms
    incoterms = get_value(order_data, 'incoterms', '')
    ws.range(CELL_DELIVERY_TERMS).value = incoterms

    # Shipping method → SO_해외.Shipping method
    shipping_method = get_value(order_data, 'shipping_method', '')
    if shipping_method:
        ws.range(CELL_SHIPPING_METHOD).value = shipping_method

    # Customer Address → Customer_해외.Bill to 1/2/3
    ws.range(CELL_CUST_ADDR_1).value = get_value(order_data, 'bill_to_1', '')
    ws.range(CELL_CUST_ADDR_2).value = get_value(order_data, 'bill_to_2', '')
    ws.range(CELL_CUST_ADDR_3).value = get_value(order_data, 'bill_to_3', '')

    # Delivery Address → SO_해외.납품 주소
    ws.range(CELL_DELV_ADDR_1).value = get_value(order_data, 'delivery_address', '')

    logger.debug(f"헤더 채우기 완료: SO_ID={so_id}")


def _restore_item_borders(ws: xw.Sheet, num_items: int) -> None:
    """행 삭제 후 아이템 영역 테두리 복원"""
    last_item_row = ITEM_START_ROW + num_items - 1

    header_bottom_row = ITEM_START_ROW - 1
    ws.range(f'A{header_bottom_row}:I{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{header_bottom_row}:I{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    ws.range(f'A{last_item_row}:I{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{last_item_row}:I{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    logger.debug(f"테두리 복원: Row {header_bottom_row} 하단, Row {last_item_row} 하단")


def _find_total_row(ws: xw.Sheet, start_row: int, max_search: int = 20) -> int:
    """'Total' 텍스트가 있는 행 찾기"""
    end_row = start_row + max_search - 1
    row = find_text_in_column_batch(ws, 'A', 'Total', start_row, end_row)
    return row if row is not None else start_row + 10


def _fill_items(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None,
) -> int:
    """아이템 데이터 채우기"""
    if items_df is None:
        items_df = pd.DataFrame([order_data])
    num_items = len(items_df)

    total_row = _find_total_row(ws, ITEM_START_ROW)
    template_item_count = total_row - ITEM_START_ROW
    logger.debug(f"템플릿 아이템 수: {template_item_count}, 실제 아이템 수: {num_items}")

    # 행 수 조정
    if num_items < template_item_count:
        rows_to_delete = template_item_count - num_items
        delete_rows_range(ws, ITEM_START_ROW + num_items, rows_to_delete)
        _restore_item_borders(ws, num_items)

    elif num_items > template_item_count:
        rows_to_insert = num_items - template_item_count

        original_last_row = ITEM_START_ROW + template_item_count - 1
        ws.range(f'A{original_last_row}:I{original_last_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlNone

        source_row = ITEM_START_ROW
        for i in range(rows_to_insert):
            insert_row = ITEM_START_ROW + template_item_count + i
            ws.range(f'{source_row}:{source_row}').api.Copy()
            ws.range(f'{insert_row}:{insert_row}').api.Insert(Shift=XlConstants.xlShiftDown)
        logger.debug(f"{rows_to_insert}개 행 삽입")

        _restore_item_borders(ws, num_items)

    _fill_items_batch(ws, items_df)
    _update_total_row(ws, num_items, order_data)

    return num_items - template_item_count if num_items > template_item_count else 0


def _update_total_row(ws: xw.Sheet, num_items: int, order_data: pd.Series) -> None:
    """Total 행의 수식과 Currency 업데이트"""
    total_row = ITEM_START_ROW + num_items
    last_item_row = total_row - 1

    ws.range(f'E{total_row}').formula = f"=SUM(E{ITEM_START_ROW}:E{last_item_row})"
    ws.range(f'F{total_row}').value = "EA"

    currency = get_value(order_data, 'currency', '')
    if currency:
        ws.range(f'{COL_CURRENCY}{total_row}').value = currency
        logger.debug(f"Currency 업데이트: {COL_CURRENCY}{total_row} = {currency}")

    sum_formula = f"=SUM(I{ITEM_START_ROW}:I{last_item_row})"
    ws.range(f'I{total_row}').formula = sum_formula
    logger.debug(f"Total 수식 업데이트: I{total_row} = {sum_formula}")


def _fill_items_batch(
    ws: xw.Sheet,
    items_df: pd.DataFrame,
) -> None:
    """아이템 데이터 배치 쓰기 (SO_해외 기반 + Dispatch date)"""
    num_items = len(items_df)
    end_row = ITEM_START_ROW + num_items - 1

    names = []
    qtys = []
    prices = []
    currencies = []
    dispatch_dates = []
    amounts = []

    for item_idx, (_, item) in enumerate(items_df.iterrows()):
        # 품목명: Model number + Item name (model number 있으면 앞에 붙임)
        raw_model = get_value(item, 'model', '')
        model = _to_text(raw_model)
        item_name = get_value(item, 'item_name', '')
        if model and item_name:
            full_name = f"{model} {item_name}"
        elif model:
            full_name = model
        else:
            full_name = str(item_name) if item_name else ''
        names.append(full_name)

        # 수량 → SO_해외.Item qty
        raw_qty = get_value(item, 'item_qty', 1)
        try:
            qty = int(raw_qty) if pd.notna(raw_qty) else 1
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 수량 변환 실패 '{raw_qty}' -> 기본값 1 사용")
            qty = 1
        qtys.append(qty)

        # 단가 → SO_해외.Sales Unit Price
        raw_price = get_value(item, 'sales_unit_price', 0)
        try:
            unit_price = float(raw_price) if pd.notna(raw_price) else 0
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 단가 변환 실패 '{raw_price}' -> 기본값 0 사용")
            unit_price = 0
        prices.append(unit_price)

        # 통화
        currency = get_value(item, 'currency', '')
        currencies.append(str(currency) if currency else '')

        # Dispatch date (SO_해외의 EXW NOAH)
        exw_noah = get_value(item, 'exw_noah', '')
        if exw_noah and pd.notna(exw_noah):
            if isinstance(exw_noah, datetime):
                dispatch_dates.append(exw_noah.strftime("%Y-%m-%d"))
            else:
                dispatch_dates.append(str(exw_noah))
        else:
            dispatch_dates.append('')

        # 금액
        amounts.append(qty * unit_price)

    # 열별 배치 쓰기 (6회 COM 호출)
    ws.range(f'{COL_ITEM_NAME}{ITEM_START_ROW}:{COL_ITEM_NAME}{end_row}').value = [[n] for n in names]
    ws.range(f'{COL_QTY}{ITEM_START_ROW}:{COL_QTY}{end_row}').value = [[q] for q in qtys]
    ws.range(f'{COL_UNIT_PRICE}{ITEM_START_ROW}:{COL_UNIT_PRICE}{end_row}').value = [[p] for p in prices]
    ws.range(f'{COL_CURRENCY}{ITEM_START_ROW}:{COL_CURRENCY}{end_row}').value = [[c] for c in currencies]
    ws.range(f'{COL_DISPATCH}{ITEM_START_ROW}:{COL_DISPATCH}{end_row}').value = [[d] for d in dispatch_dates]
    ws.range(f'{COL_AMOUNT}{ITEM_START_ROW}:{COL_AMOUNT}{end_row}').value = [[a] for a in amounts]

    ws.range(f'{ITEM_START_ROW}:{end_row}').rows.autofit()

    logger.debug(f"OC 아이템 배치 쓰기 완료: {num_items}개")
