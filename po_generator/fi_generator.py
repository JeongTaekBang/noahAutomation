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


# === 셀 매핑 (Final Invoice) ===
# Header
CELL_INVOICE_NO = 'G4'          # Invoice No
CELL_INVOICE_DATE = 'I4'        # Invoice Date (출고일)
CELL_BILL_TO_1 = 'A9'           # Bill to 주소 1줄
CELL_BILL_TO_2 = 'A10'          # Bill to 주소 2줄
CELL_BILL_TO_3 = 'A11'          # Bill to 주소 3줄
CELL_PAYMENT_TERMS = 'G8'       # Payment Terms 값 (G8:G9 병합)
CELL_DUE_DATE = 'I8'            # Due date 값 (I8:I9 병합)
CELL_PO_NO = 'G10'              # Customer PO No (G10:G11 병합)
CELL_PO_DATE = 'I10'            # Customer PO Date (I10:I11 병합)

# 아이템 시작 행 (헤더가 Row 13, 데이터는 Row 14부터)
ITEM_START_ROW = 14

# 아이템 열 (PI와 동일)
COL_ITEM_NAME = 'A'     # 품목명
COL_QTY = 'E'           # 수량
COL_UNIT_PRICE = 'G'    # 단가
COL_AMOUNT = 'I'        # 금액 (수량 * 단가)


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
    # Invoice No (DN_ID 사용)
    dn_id = get_value(order_data, 'dn_id', '')
    ws.range(CELL_INVOICE_NO).value = dn_id

    # Invoice Date (출고일)
    dispatch_date = get_value(order_data, 'dispatch_date', '')
    if dispatch_date and pd.notna(dispatch_date):
        if isinstance(dispatch_date, datetime):
            ws.range(CELL_INVOICE_DATE).value = dispatch_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_INVOICE_DATE).value = str(dispatch_date)
    else:
        # 출고일이 없으면 오늘 날짜
        ws.range(CELL_INVOICE_DATE).value = datetime.now().strftime("%Y-%m-%d")

    # Bill to (Customer_해외에서 JOIN된 데이터)
    bill_to_1 = get_value(order_data, 'bill_to_1', '')
    bill_to_2 = get_value(order_data, 'bill_to_2', '')
    bill_to_3 = get_value(order_data, 'bill_to_3', '')
    ws.range(CELL_BILL_TO_1).value = bill_to_1
    ws.range(CELL_BILL_TO_2).value = bill_to_2
    ws.range(CELL_BILL_TO_3).value = bill_to_3

    # Payment Terms (Customer_해외에서 가져온 기본값)
    payment_terms = get_value(order_data, 'payment_terms', '')
    if payment_terms:
        ws.range(CELL_PAYMENT_TERMS).value = payment_terms

    # Due date (비워둠 - 수동 입력)
    # ws.range(CELL_DUE_DATE).value = ''

    # Customer PO (SO_해외에서 JOIN된 Customer PO)
    customer_po = get_value(order_data, 'customer_po', '')
    ws.range(CELL_PO_NO).value = customer_po

    # PO Date (SO_해외에서 JOIN된 PO receipt date)
    po_date = get_value(order_data, 'po_receipt_date', '')
    if po_date and pd.notna(po_date):
        if isinstance(po_date, datetime):
            ws.range(CELL_PO_DATE).value = po_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_PO_DATE).value = str(po_date)

    logger.debug(f"헤더 채우기 완료: DN_ID={dn_id}, Bill to={bill_to_1}")


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

    # I열: Amount 합계 수식 업데이트
    sum_formula = f"=SUM(I{ITEM_START_ROW}:I{last_item_row})"
    ws.range(f'I{total_row}').formula = sum_formula
    logger.debug(f"Total 수식 업데이트: I{total_row} = {sum_formula}")

    # H열: Currency
    currency = get_value(order_data, 'currency', '')
    if currency:
        ws.range(f'H{total_row}').value = currency
        logger.debug(f"Currency 업데이트: H{total_row} = {currency}")


def _fill_items_batch(
    ws: xw.Sheet,
    items_df: pd.DataFrame,
) -> None:
    """아이템 데이터 배치 쓰기 (성능 최적화)

    FI는 열이 불연속적이므로(A, E, G, I) 열별로 배치 쓰기 수행

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

        # 금액
        amounts.append(qty * unit_price)

    # 열별 배치 쓰기 (4회 COM 호출 - 아이템 수에 관계없이 고정)
    ws.range(f'{COL_ITEM_NAME}{ITEM_START_ROW}:{COL_ITEM_NAME}{end_row}').value = [[n] for n in names]
    ws.range(f'{COL_QTY}{ITEM_START_ROW}:{COL_QTY}{end_row}').value = [[q] for q in qtys]
    ws.range(f'{COL_UNIT_PRICE}{ITEM_START_ROW}:{COL_UNIT_PRICE}{end_row}').value = [[p] for p in prices]
    ws.range(f'{COL_AMOUNT}{ITEM_START_ROW}:{COL_AMOUNT}{end_row}').value = [[a] for a in amounts]

    # 아이템 영역 행 높이 자동 조정
    ws.range(f'{ITEM_START_ROW}:{end_row}').rows.autofit()

    logger.debug(f"FI 아이템 배치 쓰기 완료: {num_items}개")
