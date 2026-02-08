"""
Proforma Invoice 생성 모듈 (xlwings 기반)
==========================================

xlwings를 사용하여 템플릿 기반으로 Proforma Invoice를 생성합니다.
이미지, 서식 등이 완벽하게 보존됩니다.

SO_해외 데이터를 사용합니다.
"""

from __future__ import annotations

import logging
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.config import PI_TEMPLATE_FILE
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


# === 셀 매핑 (Commercial Invoice 기준 - Proforma Invoice 동일) ===
# Header
CELL_CONSIGNED_TO = 'A9'        # 수취인 주소
CELL_CONSIGNED_COUNTRY = 'A10'  # 수취인 국가
CELL_CONSIGNED_TEL = 'C10'      # 수취인 전화번호
CELL_CONSIGNED_FAX = 'E10'      # 수취인 팩스번호
CELL_VESSEL = 'A12'             # 선박명/항공편
CELL_FROM = 'B13'               # 출발지
CELL_DESTINATION = 'B14'        # 도착 국가
CELL_DEPARTS = 'D15'            # 출발 예정일
CELL_INVOICE_NO = 'G4'          # Invoice No
CELL_LC_NO = 'G5'               # L/C No
CELL_INVOICE_DATE = 'I4'        # Invoice 발행일
CELL_LC_DATE = 'I5'             # L/C 발행일
CELL_HS_CODE = 'I11'            # HS CODE
CELL_PO_NO = 'G15'              # Customer PO No
CELL_PO_DATE = 'I15'            # Customer PO Date
CELL_CUSTOMER_PAGE2 = 'A53'     # 2페이지 헤더용 Customer name

# 아이템 시작 행
ITEM_START_ROW = 18

# 아이템 열
COL_ITEM_NAME = 'A'     # 품목명
COL_QTY = 'E'           # 수량
COL_UNIT_PRICE = 'G'    # 단가
COL_AMOUNT = 'I'        # 금액 (수량 * 단가)

# Incoterms / Currency 관련 셀
CELL_INCOTERMS = 'G17'          # Incoterms (단가 헤더 옆)
CELL_CURRENCY_TOTAL = 'H26'     # 합계 옆 통화 (템플릿 기준, 동적으로 조정됨)


def create_pi_xlwings(
    template_path: Path,
    output_path: Path,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> None:
    """xlwings로 Proforma Invoice 생성

    Args:
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        order_data: 주문 데이터 (첫 번째 아이템 또는 단일 아이템)
        items_df: 다중 아이템인 경우 전체 아이템 DataFrame
    """
    # 템플릿 준비 (임시 폴더로 복사)
    temp_template, temp_output = prepare_template(template_path, "pi")

    # 날짜
    today = datetime.now()
    today_str = today.strftime("%Y-%m-%d")

    try:
        # xlwings App 생명주기 관리
        with xlwings_app_context() as app:
            # 임시 템플릿 열기
            wb = app.books.open(str(temp_template))
            ws = wb.sheets[0]

            # 1. 헤더 정보 채우기
            _fill_header(ws, order_data, today_str)

            # 2. 아이템 데이터 채우기
            inserted_rows = _fill_items(ws, order_data, items_df)

            # 임시 위치에 저장
            wb.save(str(temp_output))
            logger.info(f"Proforma Invoice 생성 완료 (임시): {temp_output}")

    finally:
        # 임시 템플릿 삭제
        cleanup_temp_file(temp_template)

    # 최종 출력 경로로 이동
    shutil.move(str(temp_output), str(output_path))
    logger.info(f"Proforma Invoice 저장 완료: {output_path}")


def _fill_header(ws: xw.Sheet, order_data: pd.Series, today_str: str) -> None:
    """헤더 정보 채우기

    Args:
        ws: xlwings Sheet 객체
        order_data: 주문 데이터
        today_str: 오늘 날짜 문자열
    """
    # Invoice No (so_id 키 사용)
    so_id = get_value(order_data, 'so_id', '')
    ws.range(CELL_INVOICE_NO).value = so_id

    # Invoice Date
    ws.range(CELL_INVOICE_DATE).value = today_str

    # Customer 정보 (내부 키 사용)
    customer_name = get_value(order_data, 'customer_name', '')
    customer_address = get_value(order_data, 'customer_address', '')
    customer_country = get_value(order_data, 'customer_country', '')
    customer_tel = get_value(order_data, 'customer_tel', '')
    customer_fax = get_value(order_data, 'customer_fax', '')

    # Consigned to (고객명 + 주소)
    consigned_to = f"{customer_name}\n{customer_address}" if customer_address else customer_name
    ws.range(CELL_CONSIGNED_TO).value = consigned_to
    ws.range(CELL_CONSIGNED_COUNTRY).value = customer_country
    ws.range(CELL_CONSIGNED_TEL).value = customer_tel
    ws.range(CELL_CONSIGNED_FAX).value = customer_fax

    # 운송 정보
    ws.range(CELL_FROM).value = "INCHEON, KOREA"
    ws.range(CELL_DESTINATION).value = customer_country

    # Customer PO (내부 키 사용)
    customer_po = get_value(order_data, 'customer_po', '')
    po_date = get_value(order_data, 'po_receipt_date', '')
    ws.range(CELL_PO_NO).value = customer_po
    if po_date and pd.notna(po_date):
        if isinstance(po_date, datetime):
            ws.range(CELL_PO_DATE).value = po_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_PO_DATE).value = str(po_date)

    # Incoterms (G17)
    incoterms = get_value(order_data, 'incoterms', '')
    if incoterms:
        ws.range(CELL_INCOTERMS).value = incoterms

    # L/C 정보 (내부 키 사용)
    lc_no = get_value(order_data, 'lc_no', '')
    lc_date = get_value(order_data, 'lc_date', '')
    if lc_no:
        ws.range(CELL_LC_NO).value = lc_no
    if lc_date and pd.notna(lc_date):
        if isinstance(lc_date, datetime):
            ws.range(CELL_LC_DATE).value = lc_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_LC_DATE).value = str(lc_date)




def _restore_item_borders(ws: xw.Sheet, num_items: int) -> None:
    """행 삭제 후 아이템 영역 테두리 복원

    Args:
        ws: xlwings Sheet 객체
        num_items: 실제 아이템 수
    """
    # 마지막 아이템 행 (Total 바로 위)
    last_item_row = ITEM_START_ROW + num_items - 1

    # 헤더 아래 행 (첫 번째 아이템 행 바로 위 = Row 17)의 아래 테두리
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
        삽입된 행 수 (원래 1개 아이템 제외)
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
        # (이 테두리가 그대로 남아 중간에 선이 생기는 문제 방지)
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

    # 아이템 데이터 배치 쓰기 (N개 아이템 * 4열 COM 호출 → 1회로 감소)
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

    PI는 열이 불연속적이므로(A, E, G, I) 열별로 배치 쓰기 수행

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
        # 품목명: Model + Item name (Model은 텍스트로 변환하여 앞 0 보존)
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

        # 수량
        raw_qty = get_value(item, 'item_qty', 1)
        try:
            qty = int(raw_qty) if pd.notna(raw_qty) else 1
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 수량 변환 실패 '{raw_qty}' -> 기본값 1 사용")
            qty = 1
        qtys.append(qty)

        # 단가
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
    # xlwings는 1D 리스트를 수직으로 쓰려면 2D로 변환해야 함
    ws.range(f'{COL_ITEM_NAME}{ITEM_START_ROW}:{COL_ITEM_NAME}{end_row}').value = [[n] for n in names]
    ws.range(f'{COL_QTY}{ITEM_START_ROW}:{COL_QTY}{end_row}').value = [[q] for q in qtys]
    ws.range(f'{COL_UNIT_PRICE}{ITEM_START_ROW}:{COL_UNIT_PRICE}{end_row}').value = [[p] for p in prices]
    ws.range(f'{COL_AMOUNT}{ITEM_START_ROW}:{COL_AMOUNT}{end_row}').value = [[a] for a in amounts]

    # 아이템 영역 행 높이 자동 조정
    ws.range(f'{ITEM_START_ROW}:{end_row}').rows.autofit()

    logger.debug(f"PI 아이템 배치 쓰기 완료: {num_items}개")
