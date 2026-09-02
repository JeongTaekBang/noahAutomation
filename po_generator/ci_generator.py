"""
Commercial Invoice 생성 모듈 (xlwings 기반)
=============================================

xlwings를 사용하여 템플릿 기반으로 Commercial Invoice를 생성합니다.
이미지, 서식 등이 완벽하게 보존됩니다.

DN_해외 데이터를 사용합니다.
셀 레이아웃은 Proforma Invoice와 동일하나, 데이터 소스와 일부 셀이 다릅니다.
"""

from __future__ import annotations

import logging
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.config import CI_HS_WIDTH_DONORS, CI_TEMPLATE_FILE
from po_generator.hs_code import HS_COLUMN
from po_generator.utils import get_value, to_text
from po_generator.excel_helpers import (
    XlConstants,
    apply_hs_layout,
    xlwings_app_context,
    prepare_template,
    cleanup_temp_file,
    delete_rows_range,
    find_text_in_column_batch,
    insert_copied_rows,
    layout_item_rows,
)

logger = logging.getLogger(__name__)


# === 셀 매핑 (Commercial Invoice) ===
# Header — Bill to 3줄 확장으로 Row 13 이하 1행씩 밀림
CELL_BILL_TO_1 = 'A9'
CELL_BILL_TO_2 = 'A10'
CELL_BILL_TO_3 = 'A11'
CELL_FROM = 'B14'
CELL_DESTINATION = 'B15'
CELL_INVOICE_NO = 'G4'
CELL_INCOTERMS = 'G5'
CELL_INVOICE_DATE = 'I4'
CELL_PAYMENT_TERMS = 'I5'
CELL_PO_NO = 'G16'
CELL_PO_DATE = 'I16'

# CI 고유: 아이템 시작 행 (Row 19 = Electric Actuator 카테고리 라벨)
ITEM_START_ROW = 20

# 아이템 열 (E=Customer PO, F=Qty, G=Unit Price, H=Currency, I=Amount)
COL_ITEM_NAME = 'A'
# 라인별 HS CODE 판에서만 쓰는 열. 표준판에서는 D가 품목명 병합(A:D) 안이라
# 여기에 쓰지 않는다 — `apply_hs_layout()`이 병합을 A:C로 줄여 D를 열어 준 뒤에만 쓴다.
COL_HS_CODE = 'D'
# 템플릿에 박힌 문서 단위 HS(밸브 부품 코드). 라인별 판에서는 비운다 —
# 한 장이 서로 다른 두 HS를 주장하면 통관에서 어느 쪽을 믿을지 알 수 없다.
CELLS_DOC_HS = ('H12', 'I12')
COL_CUSTOMER_PO = 'E'
COL_QTY = 'F'
COL_UNIT_PRICE = 'G'
COL_CURRENCY = 'H'
COL_AMOUNT = 'I'


# Shipping Mark 영역
CELL_SHIPPING_MARK_NAME = 'A32'   # Customer Name (Shipping Mark 아래)
CELL_SHIPPING_MARK_BILLTO3 = 'A33'  # Bill to 3 (Customer Name 아래)
CELL_SHIPPING_MARK_PO = 'C34'     # Customer PO (PO No: 값)


def create_ci_xlwings(
    template_path: Path,
    output_path: Path,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> None:
    """xlwings로 Commercial Invoice 생성

    Args:
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        order_data: 주문 데이터 (첫 번째 아이템 또는 단일 아이템)
        items_df: 다중 아이템인 경우 전체 아이템 DataFrame
    """
    hs_mode = items_df is not None and HS_COLUMN in items_df.columns
    temp_template, temp_output = prepare_template(
        template_path, "ci_hs" if hs_mode else "ci",
    )

    try:
        with xlwings_app_context() as app:
            wb = app.books.open(str(temp_template))
            ws = wb.sheets[0]

            # 1. 헤더 정보 채우기
            _fill_header(ws, order_data, items_df)

            # 2. 아이템 데이터 채우기
            _fill_items(ws, order_data, items_df)

            wb.save(str(temp_output))
            logger.info(f"Commercial Invoice 생성 완료 (임시): {temp_output}")

    finally:
        cleanup_temp_file(temp_template)

    shutil.move(str(temp_output), str(output_path))
    logger.info(f"Commercial Invoice 저장 완료: {output_path}")


def _collect_customer_pos(order_data: pd.Series, items_df: pd.DataFrame | None) -> str:
    """아이템 전체에서 고유 Customer PO를 순서대로 콤마 결합"""
    source = items_df if items_df is not None else pd.DataFrame([order_data])
    seen: list[str] = []
    for _, item in source.iterrows():
        po = to_text(get_value(item, 'customer_po', ''))
        if po and po not in seen:
            seen.append(po)
    if not seen:
        return to_text(get_value(order_data, 'customer_po', ''))
    return ", ".join(seen)


def _fill_header(ws: xw.Sheet, order_data: pd.Series, items_df: pd.DataFrame | None = None) -> None:
    """헤더 정보 채우기

    CI는 Invoice No = dn_id, Invoice Date = dispatch_date (없으면 today)
    """
    # Invoice No (dn_id)
    dn_id = get_value(order_data, 'dn_id', '')
    ws.range(CELL_INVOICE_NO).value = dn_id

    # Invoice Date (dispatch_date, fallback today)
    dispatch_date = get_value(order_data, 'dispatch_date', '')
    if dispatch_date and pd.notna(dispatch_date):
        if isinstance(dispatch_date, datetime):
            ws.range(CELL_INVOICE_DATE).value = dispatch_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_INVOICE_DATE).value = str(dispatch_date)
    else:
        ws.range(CELL_INVOICE_DATE).value = datetime.now().strftime("%Y-%m-%d")

    # Bill to 정보 (3줄)
    bill_to_1 = get_value(order_data, 'bill_to_1', '')
    bill_to_2 = get_value(order_data, 'bill_to_2', '')
    bill_to_3 = get_value(order_data, 'bill_to_3', '')
    ws.range(CELL_BILL_TO_1).value = bill_to_1
    ws.range(CELL_BILL_TO_2).value = bill_to_2
    ws.range(CELL_BILL_TO_3).value = bill_to_3

    # 운송 정보
    ws.range(CELL_FROM).value = "INCHEON, KOREA"
    ws.range(CELL_DESTINATION).value = bill_to_3

    # Customer PO (복수 PO는 콤마 결합)
    customer_po = _collect_customer_pos(order_data, items_df)
    po_date = get_value(order_data, 'po_receipt_date', '')
    ws.range(CELL_PO_NO).value = customer_po
    if po_date and pd.notna(po_date):
        if isinstance(po_date, datetime):
            ws.range(CELL_PO_DATE).value = po_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_PO_DATE).value = str(po_date)

    # Shipping Mark 영역
    customer_name = get_value(order_data, 'customer_name', '')
    ws.range(CELL_SHIPPING_MARK_NAME).value = customer_name
    ws.range(CELL_SHIPPING_MARK_BILLTO3).value = bill_to_3
    ws.range(CELL_SHIPPING_MARK_PO).value = customer_po

    # Incoterms (G5) - DN_해외 → SO_해외 JOIN
    incoterms = get_value(order_data, 'incoterms', '')
    if incoterms:
        ws.range(CELL_INCOTERMS).value = incoterms

    # Payment terms (I5) - Customer_해외
    payment_terms = get_value(order_data, 'payment_terms', '')
    if payment_terms:
        ws.range(CELL_PAYMENT_TERMS).value = payment_terms


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

    # 라인별 HS 판이면 임시 사본의 아이템 격자에 D열을 만들어 낸다.
    # 행을 지우거나 삽입하기 전에 해야 Total 행이 아직 템플릿 좌표에 있다.
    if HS_COLUMN in items_df.columns:
        apply_hs_layout(
            ws,
            header_row=ITEM_START_ROW - 2,
            first_item_row=ITEM_START_ROW,
            last_row=total_row,
            width_donors=CI_HS_WIDTH_DONORS,
            doc_hs_cells=CELLS_DOC_HS,
        )

    if num_items < template_item_count:
        rows_to_delete = template_item_count - num_items
        delete_rows_range(ws, ITEM_START_ROW + num_items, rows_to_delete)
        _restore_item_borders(ws, num_items)

    elif num_items > template_item_count:
        original_last_row = ITEM_START_ROW + template_item_count - 1
        ws.range(f'A{original_last_row}:I{original_last_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlNone

        insert_copied_rows(
            ws, ITEM_START_ROW + template_item_count,
            num_items - template_item_count, source_row=ITEM_START_ROW,
        )

        _restore_item_borders(ws, num_items)

    # 아이템 데이터 배치 쓰기 → 병합 보장 + 행 높이 교정
    # (복사·삽입한 행은 병합을 잃을 수 있고, 병합 칸은 autofit이 먹지 않는다.
    #  CI와 PL은 늘 같이 첨부되므로 같은 공용 헬퍼를 그대로 쓴다 — 규칙이 갈리면 안 된다)
    names = _fill_items_batch(ws, items_df)
    layout_item_rows(ws, ITEM_START_ROW, names, hs_column=HS_COLUMN in items_df.columns)

    _update_total_row(ws, num_items, order_data)

    return num_items - template_item_count if num_items > template_item_count else 0


def _update_total_row(ws: xw.Sheet, num_items: int, order_data: pd.Series) -> None:
    """Total 행의 수식과 Currency 업데이트

    CI Total: F=SUM qty, G="EA", H=currency, I=SUM amount
    """
    total_row = ITEM_START_ROW + num_items
    last_item_row = total_row - 1

    # F열: Qty 합계
    qty_formula = f"=SUM(F{ITEM_START_ROW}:F{last_item_row})"
    ws.range(f'F{total_row}').formula = qty_formula

    # G열: 단위
    ws.range(f'G{total_row}').value = "EA"

    # H열: Currency
    currency = get_value(order_data, 'currency', '')
    if currency:
        ws.range(f'H{total_row}').value = currency
        logger.debug(f"Currency 업데이트: H{total_row} = {currency}")

    # I열: Amount 합계
    sum_formula = f"=SUM(I{ITEM_START_ROW}:I{last_item_row})"
    ws.range(f'I{total_row}').formula = sum_formula
    logger.debug(f"Total 수식 업데이트: I{total_row} = {sum_formula}")


def _fill_items_batch(
    ws: xw.Sheet,
    items_df: pd.DataFrame,
) -> list[str]:
    """아이템 데이터 배치 쓰기 (성능 최적화)

    CI는 열이 불연속적이므로(A, E, F, G, H, I) 열별로 배치 쓰기 수행
    """
    num_items = len(items_df)
    end_row = ITEM_START_ROW + num_items - 1

    names = []
    customer_pos = []
    qtys = []
    prices = []
    currencies = []
    amounts = []

    for item_idx, (_, item) in enumerate(items_df.iterrows()):
        # 품목명: Model number + Item name (model number 있으면 앞에 붙임)
        raw_model = get_value(item, 'model', '')
        model = to_text(raw_model)
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

        # Customer PO (행별)
        row_customer_po = get_value(item, 'customer_po', '')
        customer_pos.append(to_text(row_customer_po))

        # 수량
        raw_qty = get_value(item, 'item_qty', '')
        if not raw_qty or (isinstance(raw_qty, str) and raw_qty == ''):
            raw_qty = item.get('Qty', 1) if 'Qty' in item.index else 1
        try:
            # int(float(...)): '2.0' 같은 숫자형 문자열도 안전하게 파싱 (int('2.0')은 ValueError)
            qty = int(float(raw_qty)) if pd.notna(raw_qty) else 1
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 수량 변환 실패 '{raw_qty}' -> 기본값 1 사용")
            qty = 1
        qtys.append(qty)

        # 단가: unit_price 우선, fallback sales_unit_price
        # 정상 0원 DN 단가를 누락으로 오인하지 않도록 부재/공백만 SO 판매가로 대체
        raw_price = get_value(item, 'unit_price', None)
        if raw_price is None or (isinstance(raw_price, str) and raw_price.strip() == ''):
            raw_price = get_value(item, 'sales_unit_price', 0)
        try:
            unit_price = float(raw_price) if pd.notna(raw_price) else 0
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 단가 변환 실패 '{raw_price}' -> 기본값 0 사용")
            unit_price = 0
        prices.append(unit_price)

        # 통화 (각 행에 currency 표시)
        currency = get_value(item, 'currency', '')
        currencies.append(str(currency) if currency else '')

        # 금액
        amounts.append(qty * unit_price)

    # 열별 배치 쓰기 (6회 COM 호출)
    ws.range(f'{COL_ITEM_NAME}{ITEM_START_ROW}:{COL_ITEM_NAME}{end_row}').value = [[n] for n in names]
    if HS_COLUMN in items_df.columns:
        # 코드는 items_df에 열로 붙어 있다 — 정렬을 거쳐도 행과 함께 움직이므로
        # 별도 리스트로 넘길 때처럼 순서가 어긋날 수 없다.
        hs_codes = [to_text(code) for code in items_df[HS_COLUMN]]
        ws.range(f'{COL_HS_CODE}{ITEM_START_ROW}:{COL_HS_CODE}{end_row}').value = [[c] for c in hs_codes]
    ws.range(f'{COL_CUSTOMER_PO}{ITEM_START_ROW}:{COL_CUSTOMER_PO}{end_row}').value = [[p] for p in customer_pos]
    ws.range(f'{COL_QTY}{ITEM_START_ROW}:{COL_QTY}{end_row}').value = [[q] for q in qtys]
    ws.range(f'{COL_UNIT_PRICE}{ITEM_START_ROW}:{COL_UNIT_PRICE}{end_row}').value = [[p] for p in prices]
    ws.range(f'{COL_CURRENCY}{ITEM_START_ROW}:{COL_CURRENCY}{end_row}').value = [[c] for c in currencies]
    ws.range(f'{COL_AMOUNT}{ITEM_START_ROW}:{COL_AMOUNT}{end_row}').value = [[a] for a in amounts]

    # 행 높이는 여기서 만지지 않는다 — 품목명 칸이 A:D 병합이라 `rows.autofit()`이 먹지 않고,
    # 오히려 1줄로 줄여 긴 품목명을 잘라낸다. 호출부가 `layout_item_rows()`로 처리한다.
    logger.debug(f"CI 아이템 배치 쓰기 완료: {num_items}개")
    return names
