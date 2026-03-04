"""
Final Invoice 생성 모듈 (xlwings 기반)
=======================================

xlwings를 사용하여 템플릿 기반으로 대금 청구용 Invoice를 생성합니다.
이미지, 서식 등이 완벽하게 보존됩니다.

DN_해외 + Customer_해외 데이터를 사용합니다.
"""

from __future__ import annotations

import logging
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.config import FI_TEMPLATE_FILE
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
    """숫자를 문자열로 변환 (앞 0 보존, 뒤 .0 제거)

    Excel에서 숫자로 읽힌 값을 원래 텍스트 형태로 복원합니다.
    예: 12345.0 -> '12345', '0123' -> '0123'
    """
    if pd.isna(value) or value == '':
        return ''

    # 이미 문자열이면 그대로 반환
    if isinstance(value, str):
        return value

    # float인 경우 .0 제거
    if isinstance(value, float):
        # 정수로 변환 가능하면 정수로
        if value == int(value):
            return str(int(value))
        return str(value)

    return str(value)


# === 셀 매핑 (Final Invoice - 새 템플릿) ===
# Header
CELL_PO_NO = 'C7'               # Customer PO (C7:E7 병합)
CELL_INVOICE_NO = 'H7'          # Invoice No = DN_ID (H7:I7 병합)
CELL_PO_DATE = 'C8'             # Customer PO Date (C8:E8 병합)
CELL_INVOICE_DATE = 'H8'        # Invoice Date = 선적일 (H8:I8 병합)
CELL_PAYMENT_TERMS = 'H9'       # Payment Terms (H9:I9 병합)
CELL_DELIVERY_TERMS = 'H10'     # Delivery Terms = Incoterms (H10:I10 병합)
# Customer Address (A12~14)
CELL_CUST_ADDR_1 = 'A12'
CELL_CUST_ADDR_2 = 'A13'
CELL_CUST_ADDR_3 = 'A14'
# Delivery Address (G12~14)
CELL_DELV_ADDR_1 = 'G12'
CELL_DELV_ADDR_2 = 'G13'
CELL_DELV_ADDR_3 = 'G14'

# 아이템 (헤더 Row 16, 데이터 Row 17~)
ITEM_START_ROW = 17

COL_ITEM_NAME = 'A'     # 품목명 (A:D 병합)
COL_QTY = 'E'           # 수량
COL_UNIT_PRICE = 'F'    # 단가 (G→F)
COL_CURRENCY = 'G'      # 통화 (신규)
COL_AMOUNT = 'I'        # 금액 (I 유지, H:I는 헤더만 병합)


def create_fi_xlwings(
    template_path: Path,
    output_path: Path,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> None:
    """xlwings로 Final Invoice 생성

    Args:
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        order_data: 주문 데이터 (첫 번째 아이템 또는 단일 아이템)
        items_df: 다중 아이템인 경우 전체 아이템 DataFrame
    """
    # 템플릿 준비 (임시 폴더로 복사)
    temp_template, temp_output = prepare_template(template_path, "fi")

    try:
        # xlwings App 생명주기 관리
        with xlwings_app_context() as app:
            # 임시 템플릿 열기
            wb = app.books.open(str(temp_template))
            ws = wb.sheets[0]

            # 1. 헤더 정보 채우기
            _fill_header(ws, order_data)

            # 2. 아이템 데이터 채우기
            inserted_rows = _fill_items(ws, order_data, items_df)

            # 임시 위치에 저장
            wb.save(str(temp_output))
            logger.info(f"Final Invoice 생성 완료 (임시): {temp_output}")

    finally:
        # 임시 템플릿 삭제
        cleanup_temp_file(temp_template)

    # 최종 출력 경로로 이동
    shutil.move(str(temp_output), str(output_path))
    logger.info(f"Final Invoice 저장 완료: {output_path}")


def _fill_header(ws: xw.Sheet, order_data: pd.Series) -> None:
    """헤더 정보 채우기

    Args:
        ws: xlwings Sheet 객체
        order_data: 주문 데이터
    """
    # Customer PO → C7
    customer_po = get_value(order_data, 'customer_po', '')
    ws.range(CELL_PO_NO).value = customer_po

    # Invoice No (DN_ID) → H7
    dn_id = get_value(order_data, 'dn_id', '')
    ws.range(CELL_INVOICE_NO).value = dn_id

    # PO Date → C8
    po_date = get_value(order_data, 'po_receipt_date', '')
    if po_date and pd.notna(po_date):
        if isinstance(po_date, datetime):
            ws.range(CELL_PO_DATE).value = po_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_PO_DATE).value = str(po_date)

    # Invoice Date (선적일) → H8
    dispatch_date = get_value(order_data, 'dispatch_date', '')
    if dispatch_date and pd.notna(dispatch_date):
        if isinstance(dispatch_date, datetime):
            ws.range(CELL_INVOICE_DATE).value = dispatch_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_INVOICE_DATE).value = str(dispatch_date)
    else:
        ws.range(CELL_INVOICE_DATE).value = datetime.now().strftime("%Y-%m-%d")

    # Payment Terms → H9
    payment_terms = get_value(order_data, 'payment_terms', '')
    if payment_terms:
        ws.range(CELL_PAYMENT_TERMS).value = payment_terms

    # Delivery Terms (Incoterms) → H10 (SO_해외에서 JOIN된 값, 템플릿 기존값 덮어쓰기)
    incoterms = get_value(order_data, 'incoterms', '')
    ws.range(CELL_DELIVERY_TERMS).value = incoterms

    # Customer Address → A12/A13/A14
    bill_to_1 = get_value(order_data, 'bill_to_1', '')
    bill_to_2 = get_value(order_data, 'bill_to_2', '')
    bill_to_3 = get_value(order_data, 'bill_to_3', '')
    ws.range(CELL_CUST_ADDR_1).value = bill_to_1
    ws.range(CELL_CUST_ADDR_2).value = bill_to_2
    ws.range(CELL_CUST_ADDR_3).value = bill_to_3

    # Delivery Address → G12 (DN_해외의 Delivery Address)
    delivery_addr = get_value(order_data, 'delivery_address', '')
    ws.range(CELL_DELV_ADDR_1).value = delivery_addr

    logger.debug(f"헤더 채우기 완료: DN_ID={dn_id}, Customer={bill_to_1}")


def _restore_item_borders(ws: xw.Sheet, num_items: int) -> None:
    """행 삭제 후 아이템 영역 테두리 복원

    Args:
        ws: xlwings Sheet 객체
        num_items: 실제 아이템 수
    """
    # 마지막 아이템 행 (Total 바로 위)
    last_item_row = ITEM_START_ROW + num_items - 1

    # 헤더 아래 행 (첫 번째 아이템 행 바로 위 = Row 12)의 아래 테두리
    header_bottom_row = ITEM_START_ROW - 1
    ws.range(f'A{header_bottom_row}:I{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{header_bottom_row}:I{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    # 마지막 아이템 행의 아래 테두리
    ws.range(f'A{last_item_row}:I{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{last_item_row}:I{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    logger.debug(f"테두리 복원: Row {header_bottom_row} 하단, Row {last_item_row} 하단")


def _find_total_row(ws: xw.Sheet, start_row: int, max_search: int = 20) -> int:
    """'Total' 텍스트가 있는 행 찾기 (배치 읽기 최적화)

    Args:
        ws: xlwings Sheet 객체
        start_row: 검색 시작 행
        max_search: 최대 검색 행 수

    Returns:
        Total 행 번호 (못 찾으면 start_row + 10)
    """
    # 배치 읽기로 20회 COM 호출 → 1회로 감소
    end_row = start_row + max_search - 1
    row = find_text_in_column_batch(ws, 'A', 'Total', start_row, end_row)
    return row if row is not None else start_row + 10


def _fill_items(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None,
) -> int:
    """아이템 데이터 채우기 - 배치 쓰기 최적화

    Args:
        ws: xlwings Sheet 객체
        order_data: 주문 데이터 (첫 번째 아이템)
        items_df: 다중 아이템인 경우 DataFrame

    Returns:
        삽입된 행 수
    """
    # 아이템 준비
    if items_df is None:
        items_df = pd.DataFrame([order_data])
    num_items = len(items_df)

    # 템플릿의 기존 아이템 행 수 계산 (Total 행 찾기)
    total_row = _find_total_row(ws, ITEM_START_ROW)
    template_item_count = total_row - ITEM_START_ROW
    logger.debug(f"템플릿 아이템 수: {template_item_count}, 실제 아이템 수: {num_items}")

    # 행 수 조정: 템플릿 예시보다 실제 아이템이 적으면 초과 행 삭제
    if num_items < template_item_count:
        rows_to_delete = template_item_count - num_items
        # 범위 삭제로 N회 COM 호출 → 1회로 감소
        delete_rows_range(ws, ITEM_START_ROW + num_items, rows_to_delete)

        # 테두리 복원: 행 삭제로 사라진 테두리 다시 그리기
        _restore_item_borders(ws, num_items)

    # 행 수 조정: 템플릿 예시보다 실제 아이템이 많으면 행 삽입
    elif num_items > template_item_count:
        rows_to_insert = num_items - template_item_count

        # 삽입 전: 템플릿 원래 마지막 행의 하단 테두리 제거
        original_last_row = ITEM_START_ROW + template_item_count - 1
        ws.range(f'A{original_last_row}:I{original_last_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlNone

        source_row = ITEM_START_ROW
        for i in range(rows_to_insert):
            insert_row = ITEM_START_ROW + template_item_count + i
            ws.range(f'{source_row}:{source_row}').api.Copy()
            ws.range(f'{insert_row}:{insert_row}').api.Insert(Shift=XlConstants.xlShiftDown)
        logger.debug(f"{rows_to_insert}개 행 삽입")

        # 테두리 복원: 새 마지막 아이템 행에 하단 테두리 추가
        _restore_item_borders(ws, num_items)

    # 아이템 데이터 배치 쓰기
    _fill_items_batch(ws, items_df)

    # Total 행 수식 및 Currency 업데이트
    _update_total_row(ws, num_items, order_data)

    return num_items - template_item_count if num_items > template_item_count else 0


def _update_total_row(ws: xw.Sheet, num_items: int, order_data: pd.Series) -> None:
    """Total 행의 수식과 Currency 업데이트

    Args:
        ws: xlwings Sheet 객체
        num_items: 실제 아이템 수
        order_data: 주문 데이터 (Currency 정보용)
    """
    # Total 행 위치 계산 (아이템 마지막 행 + 1)
    total_row = ITEM_START_ROW + num_items
    last_item_row = total_row - 1

    # E열: Qty 합계 수식 업데이트
    qty_formula = f"=SUM(E{ITEM_START_ROW}:E{last_item_row})"
    ws.range(f'E{total_row}').formula = qty_formula

    # F열: 단위 "EA"
    ws.range(f'F{total_row}').value = "EA"

    # H열: Currency
    currency = get_value(order_data, 'currency', '')
    if currency:
        ws.range(f'H{total_row}').value = currency
        logger.debug(f"Currency 업데이트: H{total_row} = {currency}")

    # I열: Amount 합계 수식 업데이트
    sum_formula = f"=SUM(I{ITEM_START_ROW}:I{last_item_row})"
    ws.range(f'I{total_row}').formula = sum_formula
    logger.debug(f"Total 수식 업데이트: I{total_row} = {sum_formula}")


def _fill_items_batch(
    ws: xw.Sheet,
    items_df: pd.DataFrame,
) -> None:
    """아이템 데이터 배치 쓰기 (성능 최적화)

    FI는 열이 불연속적이므로(A, E, F, G, I) 열별로 배치 쓰기 수행

    Args:
        ws: xlwings Sheet 객체
        items_df: 아이템 DataFrame
    """
    num_items = len(items_df)
    end_row = ITEM_START_ROW + num_items - 1

    # 데이터 준비
    names = []
    qtys = []
    prices = []
    currencies = []
    amounts = []

    for item_idx, (_, item) in enumerate(items_df.iterrows()):
        # 품목명: Item 컬럼 사용 (DN_해외는 'Item' 컬럼)
        item_name = get_value(item, 'item_name', '')
        if not item_name:
            # DN_해외의 'Item' 컬럼 직접 참조
            item_name = item.get('Item', '') if 'Item' in item.index else ''
        names.append(str(item_name) if item_name else '')

        # 수량 (DN_해외는 'Qty' 컬럼)
        raw_qty = get_value(item, 'item_qty', '')
        if not raw_qty or (isinstance(raw_qty, str) and raw_qty == ''):
            raw_qty = item.get('Qty', 1) if 'Qty' in item.index else 1
        try:
            qty = int(raw_qty) if pd.notna(raw_qty) else 1
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 수량 변환 실패 '{raw_qty}' -> 기본값 1 사용")
            qty = 1
        qtys.append(qty)

        # 단가 (DN_해외는 'Unit Price' 컬럼)
        raw_price = get_value(item, 'unit_price', '')
        if not raw_price or (isinstance(raw_price, float) and raw_price == 0):
            raw_price = get_value(item, 'sales_unit_price', 0)
        try:
            unit_price = float(raw_price) if pd.notna(raw_price) else 0
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 단가 변환 실패 '{raw_price}' -> 기본값 0 사용")
            unit_price = 0
        prices.append(unit_price)

        # 통화 (SO_해외에서 JOIN된 Currency)
        currency = get_value(item, 'currency', '')
        currencies.append(str(currency) if currency else '')

        # 금액
        amounts.append(qty * unit_price)

    # 열별 배치 쓰기 (5회 COM 호출 - 아이템 수에 관계없이 고정)
    ws.range(f'{COL_ITEM_NAME}{ITEM_START_ROW}:{COL_ITEM_NAME}{end_row}').value = [[n] for n in names]
    ws.range(f'{COL_QTY}{ITEM_START_ROW}:{COL_QTY}{end_row}').value = [[q] for q in qtys]
    ws.range(f'{COL_UNIT_PRICE}{ITEM_START_ROW}:{COL_UNIT_PRICE}{end_row}').value = [[p] for p in prices]
    ws.range(f'{COL_CURRENCY}{ITEM_START_ROW}:{COL_CURRENCY}{end_row}').value = [[c] for c in currencies]
    ws.range(f'{COL_AMOUNT}{ITEM_START_ROW}:{COL_AMOUNT}{end_row}').value = [[a] for a in amounts]

    # 아이템 영역 행 높이 자동 조정
    ws.range(f'{ITEM_START_ROW}:{end_row}').rows.autofit()

    logger.debug(f"FI 아이템 배치 쓰기 완료: {num_items}개")
