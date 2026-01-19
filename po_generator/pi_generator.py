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
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.config import PI_TEMPLATE_FILE
from po_generator.utils import get_safe_value

logger = logging.getLogger(__name__)


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
    if not template_path.exists():
        raise FileNotFoundError(f"템플릿 파일이 없습니다: {template_path}")

    # 날짜
    today = datetime.now()
    today_str = today.strftime("%Y-%m-%d")

    # 경로에 한글이 포함된 경우 임시 폴더에서 작업 (xlwings COM 인터페이스 호환성)
    temp_dir = Path(tempfile.gettempdir())
    temp_template = temp_dir / f"pi_template_{today.strftime('%Y%m%d%H%M%S')}.xlsx"
    temp_output = temp_dir / f"pi_output_{today.strftime('%Y%m%d%H%M%S')}.xlsx"

    # 템플릿을 임시 폴더로 복사
    shutil.copy(template_path, temp_template)

    # Excel 앱 시작 (백그라운드)
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        # 임시 템플릿 열기
        wb = app.books.open(str(temp_template))
        ws = wb.sheets[0]

        # 1. 헤더 정보 채우기
        _fill_header(ws, order_data, today_str)

        # 2. 아이템 데이터 채우기
        inserted_rows = _fill_items(ws, order_data, items_df)

        # 3. Shipping Mark (동적으로 찾아서 채우기)
        _fill_shipping_mark(ws, order_data)

        # 임시 위치에 저장
        wb.save(str(temp_output))
        logger.info(f"Proforma Invoice 생성 완료 (임시): {temp_output}")

    finally:
        # 정리
        try:
            wb.close()
        except NameError:
            pass
        app.quit()

        # 임시 템플릿 삭제
        try:
            temp_template.unlink()
        except Exception:
            pass

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
    # Invoice No (SO_ID 사용)
    so_id = get_safe_value(order_data, 'SO_ID', '')
    ws.range(CELL_INVOICE_NO).value = so_id

    # Invoice Date
    ws.range(CELL_INVOICE_DATE).value = today_str

    # Customer 정보
    customer_name = get_safe_value(order_data, 'Customer name', '')
    customer_address = get_safe_value(order_data, 'Customer address', '')
    customer_country = get_safe_value(order_data, 'Customer country', '')
    customer_tel = get_safe_value(order_data, 'Customer TEL', '')
    customer_fax = get_safe_value(order_data, 'Customer FAX', '')

    # Consigned to (고객명 + 주소)
    consigned_to = f"{customer_name}\n{customer_address}" if customer_address else customer_name
    ws.range(CELL_CONSIGNED_TO).value = consigned_to
    ws.range(CELL_CONSIGNED_COUNTRY).value = customer_country
    ws.range(CELL_CONSIGNED_TEL).value = customer_tel
    ws.range(CELL_CONSIGNED_FAX).value = customer_fax

    # 운송 정보
    ws.range(CELL_FROM).value = "INCHEON, KOREA"
    ws.range(CELL_DESTINATION).value = customer_country

    # Customer PO
    customer_po = get_safe_value(order_data, 'Customer PO', '')
    po_date = get_safe_value(order_data, 'PO receipt date', '')
    ws.range(CELL_PO_NO).value = customer_po
    if po_date and pd.notna(po_date):
        if isinstance(po_date, datetime):
            ws.range(CELL_PO_DATE).value = po_date.strftime("%Y-%m-%d")
        else:
            ws.range(CELL_PO_DATE).value = str(po_date)

    # Incoterms
    incoterms = get_safe_value(order_data, 'Incoterms', '')
    # Incoterms는 별도 셀이 있으면 추가 (현재 매핑에 없음)

    # L/C 정보 (있으면)
    lc_no = get_safe_value(order_data, 'L/C No', '')
    lc_date = get_safe_value(order_data, 'L/C date', '')
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
    # xlwings 상수
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlContinuous = 1
    xlThin = 2

    # 마지막 아이템 행 (Total 바로 위)
    last_item_row = ITEM_START_ROW + num_items - 1

    # 헤더 아래 행 (첫 번째 아이템 행 바로 위 = Row 17)의 아래 테두리
    header_bottom_row = ITEM_START_ROW - 1
    ws.range(f'A{header_bottom_row}:I{header_bottom_row}').api.Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.range(f'A{header_bottom_row}:I{header_bottom_row}').api.Borders(xlEdgeBottom).Weight = xlThin

    # 마지막 아이템 행의 아래 테두리
    ws.range(f'A{last_item_row}:I{last_item_row}').api.Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.range(f'A{last_item_row}:I{last_item_row}').api.Borders(xlEdgeBottom).Weight = xlThin

    logger.debug(f"테두리 복원: Row {header_bottom_row} 하단, Row {last_item_row} 하단")


def _find_total_row(ws: xw.Sheet, start_row: int, max_search: int = 20) -> int:
    """'Total' 텍스트가 있는 행 찾기

    Args:
        ws: xlwings Sheet 객체
        start_row: 검색 시작 행
        max_search: 최대 검색 행 수

    Returns:
        Total 행 번호 (못 찾으면 start_row + 10)
    """
    for row in range(start_row, start_row + max_search):
        cell_value = ws.range(f'A{row}').value
        if cell_value and 'Total' in str(cell_value):
            return row
    return start_row + 10  # 기본값


def _fill_items(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None,
) -> int:
    """아이템 데이터 채우기

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
        # 뒤에서부터 삭제 (Total 행 바로 위부터)
        for _ in range(rows_to_delete):
            delete_row = ITEM_START_ROW + num_items
            ws.range(f'{delete_row}:{delete_row}').api.Delete(Shift=-4162)  # xlUp
        logger.debug(f"{rows_to_delete}개 초과 행 삭제")

        # 테두리 복원: 행 삭제로 사라진 테두리 다시 그리기
        _restore_item_borders(ws, num_items)

    # 행 수 조정: 템플릿 예시보다 실제 아이템이 많으면 행 삽입
    elif num_items > template_item_count:
        rows_to_insert = num_items - template_item_count
        source_row = ITEM_START_ROW
        for i in range(rows_to_insert):
            insert_row = ITEM_START_ROW + template_item_count + i
            ws.range(f'{source_row}:{source_row}').api.Copy()
            ws.range(f'{insert_row}:{insert_row}').api.Insert(Shift=-4121)  # xlDown
        logger.debug(f"{rows_to_insert}개 행 삽입")

    # 아이템 데이터 채우기
    total_amount = 0
    currency = get_safe_value(order_data, 'Currency', 'USD')

    for idx, (_, item) in enumerate(items_df.iterrows()):
        row_num = ITEM_START_ROW + idx

        # 품목명
        item_name = get_safe_value(item, 'Item name', '')
        ws.range(f'{COL_ITEM_NAME}{row_num}').value = item_name

        # 수량
        qty = get_safe_value(item, 'Item qty', 1)
        try:
            qty = int(qty) if pd.notna(qty) else 1
        except (ValueError, TypeError):
            qty = 1
        ws.range(f'{COL_QTY}{row_num}').value = qty

        # 단가 (Sales Unit Price 사용)
        unit_price = get_safe_value(item, 'Sales Unit Price', 0)
        try:
            unit_price = float(unit_price) if pd.notna(unit_price) else 0
        except (ValueError, TypeError):
            unit_price = 0
        ws.range(f'{COL_UNIT_PRICE}{row_num}').value = unit_price

        # 금액
        amount = qty * unit_price
        ws.range(f'{COL_AMOUNT}{row_num}').value = amount
        total_amount += amount

    return num_items - template_item_count if num_items > template_item_count else 0


def _fill_shipping_mark(ws: xw.Sheet, order_data: pd.Series) -> None:
    """Shipping Mark 영역 채우기 (동적으로 찾아서)

    각 레이블 텍스트를 찾아서 해당 위치에 값을 채웁니다.

    Args:
        ws: xlwings Sheet 객체
        order_data: 주문 데이터
    """
    # 데이터 추출
    customer_name = get_safe_value(order_data, 'Customer name', '')
    customer_country = get_safe_value(order_data, 'Customer country', '')
    customer_po = get_safe_value(order_data, 'Customer PO', '')

    # 검색 범위 (아이템 행 삭제로 위치가 변할 수 있으므로 넓게 검색)
    search_start = 20
    search_end = 100

    # 1. Shipping Mark 헤더 찾아서 +1행에 Customer name
    for row in range(search_start, search_end):
        cell_value = ws.range(f'A{row}').value
        if cell_value and 'Shipping Mark' in str(cell_value):
            ws.range(f'A{row + 1}').value = customer_name
            ws.range(f'A{row + 2}').value = customer_country
            logger.debug(f"Shipping Mark 발견 Row {row}: {customer_name}, {customer_country}")
            break

    # 2. "PO No" 텍스트 찾아서 같은 행 C열에 PO 값
    for row in range(search_start, search_end):
        for col in ['A', 'B']:
            cell_value = ws.range(f'{col}{row}').value
            if cell_value and 'PO No' in str(cell_value):
                ws.range(f'C{row}').value = customer_po
                logger.debug(f"PO No 발견 {col}{row}: {customer_po} -> C{row}")
                return

    logger.warning("PO No 레이블을 찾을 수 없습니다.")
