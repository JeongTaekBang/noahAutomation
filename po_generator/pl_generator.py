"""
Packing List 생성 모듈 (xlwings 기반)
======================================

xlwings를 사용하여 템플릿 기반으로 Packing List를 생성합니다.
이미지, 서식 등이 완벽하게 보존됩니다.

DN_해외 데이터를 사용합니다.
셀 레이아웃은 Commercial Invoice와 동일하나, 아이템 열이 다릅니다.
(단가/금액 대신 Weight/CBM)
"""

from __future__ import annotations

import logging
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.config import PL_TEMPLATE_FILE
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


# === 셀 매핑 (Packing List - CI와 헤더 동일) ===
# Header
CELL_CONSIGNED_TO = 'A9'
CELL_CONSIGNED_COUNTRY = 'A10'
CELL_CONSIGNED_TEL = 'C10'
CELL_CONSIGNED_FAX = 'E10'
CELL_VESSEL = 'A12'
CELL_FROM = 'B13'
CELL_DESTINATION = 'B14'
CELL_DEPARTS = 'D15'
CELL_INVOICE_NO = 'G4'
CELL_LC_NO = 'G5'
CELL_INVOICE_DATE = 'I4'
CELL_LC_DATE = 'I5'
CELL_HS_CODE = 'I11'
CELL_PO_NO = 'G15'
CELL_PO_DATE = 'I15'

# 아이템 시작 행 (Row 18 = 카테고리 라벨 유지)
ITEM_START_ROW = 19

# 아이템 열 (PL 고유: 단가/금액 대신 Weight/CBM)
COL_ITEM_NAME = 'A'
COL_QTY = 'E'
COL_NET_WEIGHT = 'F'    # Net Weight (KG/PC)
COL_GROSS_WEIGHT = 'H'  # Gross Weight (Kg)
COL_CBM = 'I'           # Measurement (CBM)

# Shipping Mark 영역 (템플릿 기준 고정 위치, 행 삽입/삭제 시 자동 이동)
CELL_SHIPPING_MARK_NAME = 'A34'
CELL_SHIPPING_MARK_BILLTO3 = 'A35'
CELL_SHIPPING_MARK_PO = 'C36'


def create_pl_xlwings(
    template_path: Path,
    output_path: Path,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> None:
    """xlwings로 Packing List 생성

    Args:
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        order_data: 주문 데이터 (첫 번째 아이템 또는 단일 아이템)
        items_df: 다중 아이템인 경우 전체 아이템 DataFrame
    """
    temp_template, temp_output = prepare_template(template_path, "pl")

    try:
        with xlwings_app_context() as app:
            wb = app.books.open(str(temp_template))
            ws = wb.sheets[0]

            # 1. 헤더 정보 채우기 (Shipping Mark 포함 — 행 조정 전에 고정 위치에 쓰기)
            _fill_header(ws, order_data)

            # 2. 아이템 데이터 채우기
            _fill_items(ws, order_data, items_df)

            wb.save(str(temp_output))
            logger.info(f"Packing List 생성 완료 (임시): {temp_output}")

    finally:
        cleanup_temp_file(temp_template)

    shutil.move(str(temp_output), str(output_path))
    logger.info(f"Packing List 저장 완료: {output_path}")


def _fill_header(ws: xw.Sheet, order_data: pd.Series) -> None:
    """헤더 정보 채우기

    PL은 Invoice No = dn_id, Date = dispatch_date (없으면 today)
    """
    # Invoice No (dn_id)
    dn_id = get_value(order_data, 'dn_id', '')
    ws.range(CELL_INVOICE_NO).value = dn_id

    # Date (dispatch_date, fallback today)
    dispatch_date = get_value(order_data, 'dispatch_date', '')
    if dispatch_date and pd.notna(dispatch_date):
        if isinstance(dispatch_date, datetime):
            ws.range(CELL_INVOICE_DATE).value = dispatch_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_INVOICE_DATE).value = str(dispatch_date)
    else:
        ws.range(CELL_INVOICE_DATE).value = datetime.now().strftime("%Y-%m-%d")

    # Consigned to (Delivery Address)
    customer_name = get_value(order_data, 'customer_name', '')
    customer_country = get_value(order_data, 'customer_country', '')
    customer_tel = get_value(order_data, 'customer_tel', '')
    customer_fax = get_value(order_data, 'customer_fax', '')

    delivery_address = get_value(order_data, 'delivery_address', '')
    ws.range(CELL_CONSIGNED_TO).value = delivery_address if delivery_address else customer_name
    ws.range(CELL_CONSIGNED_COUNTRY).value = customer_country
    ws.range(CELL_CONSIGNED_TEL).value = customer_tel
    ws.range(CELL_CONSIGNED_FAX).value = customer_fax

    # 운송 정보
    ws.range(CELL_FROM).value = "INCHEON, KOREA"
    ws.range(CELL_DESTINATION).value = customer_country

    # Customer PO
    customer_po = get_value(order_data, 'customer_po', '')
    po_date = get_value(order_data, 'po_receipt_date', '')
    ws.range(CELL_PO_NO).value = customer_po
    if po_date and pd.notna(po_date):
        if isinstance(po_date, datetime):
            ws.range(CELL_PO_DATE).value = po_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_PO_DATE).value = str(po_date)

    # Shipping Mark 영역
    bill_to_3 = get_value(order_data, 'bill_to_3', '')
    ws.range(CELL_SHIPPING_MARK_NAME).value = customer_name
    ws.range(CELL_SHIPPING_MARK_BILLTO3).value = bill_to_3
    ws.range(CELL_SHIPPING_MARK_PO).value = customer_po

    # L/C 정보
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
    """행 삭제 후 아이템 영역 테두리 복원"""
    last_item_row = ITEM_START_ROW + num_items - 1
    header_bottom_row = ITEM_START_ROW - 1

    ws.range(f'A{header_bottom_row}:I{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{header_bottom_row}:I{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    ws.range(f'A{last_item_row}:I{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{last_item_row}:I{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    logger.debug(f"테두리 복원: Row {header_bottom_row} 하단, Row {last_item_row} 하단")


def _find_total_row(ws: xw.Sheet, start_row: int, max_search: int = 20) -> int:
    """'Total' 텍스트가 있는 행 찾기 (배치 읽기 최적화)"""
    end_row = start_row + max_search - 1
    row = find_text_in_column_batch(ws, 'A', 'Total', start_row, end_row)
    return row if row is not None else start_row + 10


def _fill_items(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None,
) -> int:
    """아이템 데이터 채우기 - 배치 쓰기 최적화"""
    if items_df is None:
        items_df = pd.DataFrame([order_data])

    # Model number 기준 오름차순 정렬
    model_col = None
    for alias in ('Model number', 'Model', 'model'):
        if alias in items_df.columns:
            model_col = alias
            break
    if model_col:
        items_df = items_df.sort_values(by=model_col, ascending=True, na_position='last').reset_index(drop=True)

    num_items = len(items_df)

    total_row = _find_total_row(ws, ITEM_START_ROW)
    template_item_count = total_row - ITEM_START_ROW
    logger.debug(f"템플릿 아이템 수: {template_item_count}, 실제 아이템 수: {num_items}")

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
    _update_total_row(ws, num_items)

    return num_items - template_item_count if num_items > template_item_count else 0


def _update_total_row(ws: xw.Sheet, num_items: int) -> None:
    """Total 행의 수식 업데이트

    PL Total: E=SUM qty, G="KGS", H=SUM gross weight, I=SUM CBM
    """
    total_row = ITEM_START_ROW + num_items
    last_item_row = total_row - 1

    # E열: Qty 합계
    ws.range(f'E{total_row}').formula = f"=SUM(E{ITEM_START_ROW}:E{last_item_row})"

    # G열: 단위 라벨
    ws.range(f'G{total_row}').value = "KGS"

    # H열: Gross Weight 합계
    ws.range(f'H{total_row}').formula = f"=SUM(H{ITEM_START_ROW}:H{last_item_row})"

    # I열: CBM 합계
    ws.range(f'I{total_row}').formula = f"=SUM(I{ITEM_START_ROW}:I{last_item_row})"

    logger.debug(f"Total 수식 업데이트: row {total_row}")


def _fill_items_batch(
    ws: xw.Sheet,
    items_df: pd.DataFrame,
) -> None:
    """아이템 데이터 배치 쓰기 (성능 최적화)

    PL은 열이 불연속적이므로(A, E, F, H, I) 열별로 배치 쓰기 수행
    """
    num_items = len(items_df)
    end_row = ITEM_START_ROW + num_items - 1

    names = []
    qtys = []
    net_weights = []
    gross_weights = []
    cbms = []

    for item_idx, (_, item) in enumerate(items_df.iterrows()):
        # 품목명: Model number + Item name
        raw_model = get_value(item, 'model', '')
        model = _to_text(raw_model)
        item_name = get_value(item, 'item_name', '')
        if not item_name:
            item_name = item.get('Item', '') if 'Item' in item.index else ''
        if model and item_name:
            full_name = f"{model} {item_name}"
        elif model:
            full_name = model
        else:
            full_name = str(item_name) if item_name else ''
        names.append(full_name)

        # 수량
        raw_qty = get_value(item, 'item_qty', '')
        if not raw_qty or (isinstance(raw_qty, str) and raw_qty == ''):
            raw_qty = item.get('Qty', 1) if 'Qty' in item.index else 1
        try:
            qty = int(raw_qty) if pd.notna(raw_qty) else 1
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 수량 변환 실패 '{raw_qty}' -> 기본값 1 사용")
            qty = 1
        qtys.append(qty)

        # Net Weight (KG/PC)
        raw_nw = get_value(item, 'weight_per_unit', '')
        try:
            nw = float(raw_nw) if raw_nw and pd.notna(raw_nw) else ''
        except (ValueError, TypeError):
            nw = ''
        net_weights.append(nw)

        # Gross Weight (Kg)
        raw_gw = get_value(item, 'gross_weight', '')
        try:
            gw = float(raw_gw) if raw_gw and pd.notna(raw_gw) else ''
        except (ValueError, TypeError):
            gw = ''
        gross_weights.append(gw)

        # CBM (Measurement)
        raw_cbm = get_value(item, 'cbm', '')
        try:
            cbm = float(raw_cbm) if raw_cbm and pd.notna(raw_cbm) else ''
        except (ValueError, TypeError):
            cbm = ''
        cbms.append(cbm)

    # 열별 배치 쓰기 (5회 COM 호출)
    ws.range(f'{COL_ITEM_NAME}{ITEM_START_ROW}:{COL_ITEM_NAME}{end_row}').value = [[n] for n in names]
    ws.range(f'{COL_QTY}{ITEM_START_ROW}:{COL_QTY}{end_row}').value = [[q] for q in qtys]
    ws.range(f'{COL_NET_WEIGHT}{ITEM_START_ROW}:{COL_NET_WEIGHT}{end_row}').value = [[n] for n in net_weights]
    ws.range(f'{COL_GROSS_WEIGHT}{ITEM_START_ROW}:{COL_GROSS_WEIGHT}{end_row}').value = [[g] for g in gross_weights]
    ws.range(f'{COL_CBM}{ITEM_START_ROW}:{COL_CBM}{end_row}').value = [[c] for c in cbms]

    # 아이템 영역 행 높이 자동 조정
    ws.range(f'{ITEM_START_ROW}:{end_row}').rows.autofit()

    logger.debug(f"PL 아이템 배치 쓰기 완료: {num_items}개")
