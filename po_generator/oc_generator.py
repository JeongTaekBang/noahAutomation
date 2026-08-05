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

from po_generator.utils import get_value, to_text
from po_generator.excel_helpers import (
    ITEM_GRID_INNER_COLOR,
    XlConstants,
    xlwings_app_context,
    prepare_template,
    cleanup_temp_file,
    delete_rows_range,
    find_text_in_column_batch,
    insert_copied_rows,
    layout_address_rows,
    layout_item_rows,
)

logger = logging.getLogger(__name__)


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
ADDR_START_ROW = 13     # 주소 블록 첫 행 — 왼쪽 A13:E15(행별 병합), 오른쪽 G13:I15(세로 병합)

# 아이템 (헤더 Row 17, 데이터 Row 18~)
ITEM_START_ROW = 18

COL_ITEM_NAME = 'A'     # 품목명 (A:D 병합)
COL_QTY = 'E'           # 수량
COL_UNIT_PRICE = 'F'    # 단가
COL_CURRENCY = 'G'      # 통화
COL_DISPATCH = 'H'      # Dispatch date (OC 신규)
COL_AMOUNT = 'I'        # 금액

# 품목명 칸의 병합 범위(A:D)는 문서 5종 공통이라 excel_helpers가 소유한다
# (`ITEM_NAME_MERGED_COLS` — 여기서 재정의하지 말 것)


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
    bill_tos = [
        to_text(get_value(order_data, 'bill_to_1', '')),
        to_text(get_value(order_data, 'bill_to_2', '')),
        to_text(get_value(order_data, 'bill_to_3', '')),
    ]
    ws.range(CELL_CUST_ADDR_1).value = bill_tos[0]
    ws.range(CELL_CUST_ADDR_2).value = bill_tos[1]
    ws.range(CELL_CUST_ADDR_3).value = bill_tos[2]

    # Delivery Address → SO_해외.납품 주소
    delivery_addr = to_text(get_value(order_data, 'delivery_address', ''))
    ws.range(CELL_DELV_ADDR_1).value = delivery_addr

    # 주소 칸은 병합 셀이라 길면 경계에서 잘린다 — 줄바꿈 + 행 높이 확보
    layout_address_rows(ws, ADDR_START_ROW, bill_tos, delivery_addr)

    logger.debug(f"헤더 채우기 완료: SO_ID={so_id}")


def _restore_item_borders(ws: xw.Sheet, num_items: int) -> None:
    """행 조정 후 아이템 영역 테두리 정리

    프레임(헤더밴드 하단·마지막 행 하단)은 검정 thin으로 되그리고, 행 사이
    내부 가로선은 상단 규칙선과 같은 회색(`ITEM_GRID_INNER_COLOR`)으로 누른다
    — 검정 thin 격자는 PDF에서 0.96pt 실선이라 유독 무겁게 보인다(2026-08-05 보고).
    """
    last_item_row = ITEM_START_ROW + num_items - 1

    header_bottom_row = ITEM_START_ROW - 1
    ws.range(f'A{header_bottom_row}:I{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{header_bottom_row}:I{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    ws.range(f'A{last_item_row}:I{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{last_item_row}:I{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    if last_item_row > ITEM_START_ROW:
        ws.range(f'A{ITEM_START_ROW}:I{last_item_row}').api.Borders(
            XlConstants.xlInsideHorizontal
        ).Color = ITEM_GRID_INNER_COLOR

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
    """아이템 데이터 채우기 (부족분 삽입 → 값·행 높이 → 남는 행 삭제)

    표는 마지막 아이템에서 끝난다 — 남는 템플릿 행은 지운다. 예전에는 남는 높이만큼
    빈 행을 채워 한 페이지를 완성했는데, 아이템이 적은 문서(1아이템이 전체의 절반)가
    빈 격자 여러 줄을 달고 나가 정리했다 (2026-08-05 사용자 결정).
    """
    if items_df is None:
        items_df = pd.DataFrame([order_data])
    num_items = len(items_df)

    total_row = _find_total_row(ws, ITEM_START_ROW)
    template_count = total_row - ITEM_START_ROW
    logger.debug(f"템플릿 아이템 수: {template_count}, 실제 아이템 수: {num_items}")

    # 0. 템플릿 마지막 아이템 행의 하단 테두리를 미리 지운다 — 최종 행 수가 더 많으면
    #    표 중간에 선이 남는다. 마지막 행이 확정된 뒤 _restore_item_borders가 다시 그린다.
    last_tpl_row = ITEM_START_ROW + template_count - 1
    ws.range(f'A{last_tpl_row}:I{last_tpl_row}').api.Borders(
        XlConstants.xlEdgeBottom
    ).LineStyle = XlConstants.xlNone

    # 1. 부족한 행 삽입 (값 채우기 전에 — 배치 쓰기가 전 행을 덮도록)
    if num_items > template_count:
        insert_copied_rows(
            ws, ITEM_START_ROW + template_count, num_items - template_count,
            source_row=ITEM_START_ROW,
        )

    # 2. 값 채우기 → 병합 보장 + 행 높이 (품목명 칸은 A:D 병합이라 autofit이 먹지 않는다)
    names = _fill_items_batch(ws, items_df)
    layout_item_rows(ws, ITEM_START_ROW, names)

    # 3. 남는 템플릿 행 삭제 — 표가 마지막 아이템 바로 다음의 Total로 이어진다
    if num_items < template_count:
        delete_rows_range(ws, ITEM_START_ROW + num_items, template_count - num_items)

    _restore_item_borders(ws, num_items)
    _update_total_row(ws, num_items, order_data)

    return num_items - template_count


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
) -> list[str]:
    """아이템 데이터 배치 쓰기 (SO_해외 기반 + Dispatch date)

    Returns:
        각 행에 쓴 품목명 — 호출부가 행 높이를 재는 데 그대로 쓴다
    """
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
        model = to_text(raw_model)
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
            # int(float(...)): '2.0' 같은 숫자형 문자열도 안전하게 파싱 (int('2.0')은 ValueError)
            qty = int(float(raw_qty)) if pd.notna(raw_qty) else 1
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

    # 행 높이는 여기서 만지지 않는다 — 품목명 칸이 병합돼 있어 `rows.autofit()`이
    # 먹지 않기 때문(오히려 1줄로 줄여 잘라낸다). 호출부가 `layout_item_rows()`로 처리한다.
    logger.debug(f"OC 아이템 배치 쓰기 완료: {num_items}개")
    return names
