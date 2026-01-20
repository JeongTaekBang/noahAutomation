"""
Excel 생성 모듈 (xlwings 기반)
===============================

xlwings를 사용하여 템플릿 기반으로 Purchase Order 및 Description 시트를 생성합니다.
이미지, 서식 등이 완벽하게 보존됩니다.
"""

from __future__ import annotations

import logging
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.config import (
    TOTAL_COLUMNS,
    ITEM_START_ROW_FALLBACK,
    VAT_RATE_DOMESTIC,
    SPEC_FIELDS,
    OPTION_FIELDS,
    PO_TEMPLATE_FILE,
)
from po_generator.utils import (
    get_value,
    escape_excel_formula,
)
from po_generator.excel_helpers import (
    find_item_start_row_xlwings,
    PO_HEADER_LABELS,
)

logger = logging.getLogger(__name__)


# === 템플릿 셀 매핑 ===
CELL_TITLE = 'A1'
CELL_DATE = 'A5'
CELL_DELIVERY_ADDR = 'C5'
CELL_CUSTOMER_PO = 'C7'
CELL_CUSTOMER_NAME = 'A10'


def _find_item_start_row_xlwings(
    ws: xw.Sheet,
    search_labels: tuple[str, ...] = PO_HEADER_LABELS,
    max_search_rows: int = 30,
    fallback_row: int = ITEM_START_ROW_FALLBACK,
) -> int:
    """템플릿에서 아이템 시작 행을 동적으로 찾기 (xlwings 버전)

    excel_helpers.find_item_start_row_xlwings의 래퍼입니다.
    하위 호환성을 위해 유지됩니다.

    Args:
        ws: xlwings Sheet 객체
        search_labels: 검색할 헤더 레이블
        max_search_rows: 최대 검색 행 수
        fallback_row: 헤더를 찾지 못했을 때 기본값

    Returns:
        아이템 시작 행 번호
    """
    return find_item_start_row_xlwings(
        ws,
        search_labels=search_labels,
        max_search_rows=max_search_rows,
        fallback_row=fallback_row,
    )


def _ensure_template_exists() -> None:
    """템플릿 파일이 없으면 오류 발생"""
    if not PO_TEMPLATE_FILE.exists():
        raise FileNotFoundError(
            f"템플릿 파일이 없습니다: {PO_TEMPLATE_FILE}\n"
            "templates/purchase_order.xlsx 파일을 생성해 주세요."
        )


def _fill_header_data(
    ws: xw.Sheet,
    order_data: pd.Series,
    rck_order_no: str,
    today_str: str,
) -> None:
    """헤더 섹션에 데이터 채움

    Args:
        ws: xlwings Sheet 객체
        order_data: 주문 데이터
        rck_order_no: RCK Order No.
        today_str: 오늘 날짜 문자열
    """
    customer_name = get_value(order_data, 'customer_name')
    customer_po = get_value(order_data, 'customer_po')
    delivery_addr = get_value(order_data, 'delivery_address')

    # 데이터 채움 (수식 인젝션 방지 적용)
    ws.range(CELL_TITLE).value = f"Purchase Order - {escape_excel_formula(rck_order_no)}"
    ws.range(CELL_DATE).value = f"Date:  {today_str}"
    ws.range(CELL_DELIVERY_ADDR).value = escape_excel_formula(delivery_addr)
    ws.range(CELL_CUSTOMER_PO).value = escape_excel_formula(customer_po)
    ws.range(CELL_CUSTOMER_NAME).value = escape_excel_formula(customer_name)


def _fill_item_data(
    ws: xw.Sheet,
    row_num: int,
    item_idx: int,
    item_data: pd.Series,
    currency: str = 'KRW',
    is_export: bool = False,
) -> None:
    """아이템 행에 데이터 채움

    Args:
        ws: xlwings Sheet 객체
        row_num: 행 번호
        item_idx: 아이템 인덱스 (0부터 시작)
        item_data: 아이템 데이터
        currency: 통화 코드 (KRW 또는 USD)
        is_export: 해외 여부
    """
    number_format = '₩#,##0' if currency == 'KRW' else '$#,##0.00'

    # 데이터 추출
    item_name = get_value(item_data, 'item_name')

    # Description: 해외는 Model number + Item name, 국내는 Item name만
    if is_export:
        model_number = get_value(item_data, 'model')
        if model_number and item_name:
            description = f"{model_number} {item_name}"
        elif item_name:
            description = item_name
        elif model_number:
            description = model_number
        else:
            description = ''
    else:
        description = item_name if item_name else ''

    # 수량
    try:
        qty = int(float(get_value(item_data, 'item_qty', 1)))
    except (ValueError, TypeError):
        qty = 1

    # 단가
    try:
        ico_unit = float(get_value(item_data, 'ico_unit', 0))
    except (ValueError, TypeError):
        ico_unit = 0

    # 납기일
    requested_date = get_value(item_data, 'delivery_date')
    requested_date_str = ''
    if requested_date and not pd.isna(requested_date):
        try:
            if isinstance(requested_date, datetime):
                requested_date_str = requested_date.strftime("%Y-%m-%d")
            else:
                requested_date_str = str(requested_date)[:10]
        except (ValueError, TypeError):
            requested_date_str = ''

    # 데이터 입력 (수식 인젝션 방지 적용)
    ws.range(f'A{row_num}').value = item_idx + 1
    ws.range(f'B{row_num}').value = escape_excel_formula(description)
    ws.range(f'F{row_num}').value = qty
    ws.range(f'G{row_num}').value = "EA"
    ws.range(f'H{row_num}').value = ico_unit
    ws.range(f'H{row_num}').number_format = number_format
    ws.range(f'I{row_num}').value = requested_date_str
    ws.range(f'J{row_num}').formula = f"=H{row_num}*F{row_num}"
    ws.range(f'J{row_num}').number_format = number_format


def _fill_footer_data(
    ws: xw.Sheet,
    order_data: pd.Series,
    footer_start_row: int,
    is_export: bool = False,
) -> int:
    """푸터 섹션에 데이터 채움

    Args:
        ws: xlwings Sheet 객체
        order_data: 주문 데이터
        footer_start_row: 푸터 시작 행 (합계 섹션 다음)
        is_export: 해외 여부

    Returns:
        마지막 행 번호
    """
    remark = get_value(order_data, 'remark')
    incoterms = 'EXW' if is_export else get_value(order_data, 'incoterms')
    currency = 'KRW'

    # 프로젝트 정보
    opportunity = get_value(order_data, 'opportunity')
    sector = get_value(order_data, 'sector')
    industry_code = get_value(order_data, 'industry_code')

    r = footer_start_row
    ws.range(f'D{r}').value = escape_excel_formula(opportunity)
    r += 1
    ws.range(f'D{r}').value = escape_excel_formula(sector)
    r += 1
    ws.range(f'D{r}').value = escape_excel_formula(industry_code)
    r += 1
    ws.range(f'C{r}').value = f"Note. {escape_excel_formula(remark)}" if remark else "Note."
    r += 1
    ws.range(f'B{r}').value = currency
    r += 1
    ws.range(f'B{r}').value = incoterms

    # 마지막 행 (청록색 푸터)은 r + 3
    return r + 3


def _find_totals_row(ws: xw.Sheet, start_row: int, max_search: int = 20) -> int:
    """'Total net amount' 텍스트가 있는 행 찾기

    Args:
        ws: xlwings Sheet 객체
        start_row: 검색 시작 행
        max_search: 최대 검색 행 수

    Returns:
        Total net amount 행 번호 (못 찾으면 start_row)
    """
    for row in range(start_row, start_row + max_search):
        cell_value = ws.range(f'I{row}').value
        if cell_value and 'Total net' in str(cell_value):
            return row
    return start_row


def _create_purchase_order(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> None:
    """Purchase Order 시트 생성 (xlwings 기반)

    Args:
        ws: xlwings Sheet 객체 (템플릿에서 복사된 시트)
        order_data: 첫 번째 아이템 데이터 (공통 정보 추출용)
        items_df: 다중 아이템인 경우 DataFrame, 단일이면 None
    """
    logger.info("Purchase Order 시트 생성 중 (xlwings)...")

    # 아이템 목록 준비
    if items_df is not None:
        items_list = [row for _, row in items_df.iterrows()]
    else:
        items_list = [order_data]

    num_items = len(items_list)

    # 공통 데이터
    rck_order_no = get_value(order_data, 'order_no')
    today_str = datetime.now().strftime("%d/%b/%Y").upper()
    currency = 'KRW'

    # 해외(수출) 건 여부 확인
    sheet_type = get_value(order_data, 'sheet_type', '')
    is_export = sheet_type == '해외'

    # 1. 헤더 데이터 채움
    _fill_header_data(ws, order_data, rck_order_no, today_str)

    # 2. 아이템 시작 행 동적 탐지
    template_row = _find_item_start_row_xlwings(ws, fallback_row=ITEM_START_ROW_FALLBACK)

    # 3. 템플릿 아이템 행 수 계산 (Total net amount 행 찾기)
    totals_row = _find_totals_row(ws, template_row)
    template_item_count = totals_row - template_row
    logger.debug(f"템플릿 아이템 수: {template_item_count}, 실제 아이템 수: {num_items}")

    # 4. 행 수 조정
    if num_items < template_item_count:
        # 초과 행 삭제
        rows_to_delete = template_item_count - num_items
        for _ in range(rows_to_delete):
            delete_row = template_row + num_items
            ws.range(f'{delete_row}:{delete_row}').api.Delete(Shift=-4162)  # xlUp
        logger.debug(f"{rows_to_delete}개 초과 행 삭제")
    elif num_items > template_item_count:
        # 행 삽입
        rows_to_insert = num_items - template_item_count
        for i in range(rows_to_insert):
            ws.range(f'{template_row}:{template_row}').api.Copy()
            insert_row = template_row + template_item_count + i
            ws.range(f'{insert_row}:{insert_row}').api.Insert(Shift=-4121)  # xlDown
        logger.debug(f"{rows_to_insert}개 행 삽입")

    # 5. 아이템 데이터 채움
    for item_idx, item_data in enumerate(items_list):
        row_num = template_row + item_idx
        _fill_item_data(ws, row_num, item_idx, item_data, currency, is_export)

    item_last_row = template_row + num_items - 1

    # 6. 합계 섹션 업데이트 (동적 위치)
    totals_start_row = item_last_row + 1
    row_total_net = totals_start_row
    row_vat = totals_start_row + 1
    row_order_total = totals_start_row + 2

    # SUM 공식 범위 업데이트
    ws.range(f'J{row_total_net}').formula = f"=SUM(J{template_row}:J{item_last_row})"

    # VAT 처리 (해외는 0)
    if is_export:
        ws.range(f'J{row_vat}').value = 0
    else:
        ws.range(f'J{row_vat}').formula = f"=J{row_total_net}*{VAT_RATE_DOMESTIC}"

    # Order Total 공식 업데이트
    ws.range(f'J{row_order_total}').formula = f"=SUM(J{row_total_net}:J{row_vat})"

    # 7. 푸터 데이터 채움
    footer_start_row = row_order_total + 1
    last_row = _fill_footer_data(ws, order_data, footer_start_row, is_export)

    # 8. 인쇄 영역 업데이트
    ws.api.PageSetup.PrintArea = f'$A$1:$J${last_row}'

    logger.info(f"Purchase Order 시트 생성 완료 (아이템 {num_items}개)")


def _create_description_sheet(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> None:
    """Description 시트 생성 (xlwings 기반)

    Args:
        ws: xlwings Sheet 객체 (템플릿에서 복사된 시트)
        order_data: 첫 번째 아이템 데이터
        items_df: 다중 아이템인 경우 DataFrame, 단일이면 None
    """
    logger.info("Description 시트 생성 중 (xlwings)...")

    # 아이템 목록 준비
    if items_df is not None:
        items_list = [row for _, row in items_df.iterrows()]
    else:
        items_list = [order_data]

    num_items = len(items_list)

    # 첫 번째 아이템 데이터 (B열)
    try:
        qty_first = int(float(get_value(items_list[0], 'item_qty', 1)))
    except (ValueError, TypeError):
        qty_first = 1
    ws.range('B2').value = qty_first

    # 추가 아이템 열 생성 (B열은 이미 템플릿에 있음)
    if num_items > 1:
        for idx in range(1, num_items):
            col_letter = chr(ord('C') + idx - 1)  # C, D, E...

            # Row 1: Line No
            ws.range(f'{col_letter}1').value = idx + 1

            # Row 2: Qty
            try:
                qty = int(float(get_value(items_list[idx], 'item_qty', 1)))
            except (ValueError, TypeError):
                qty = 1
            ws.range(f'{col_letter}2').value = qty

    # SPEC_FIELDS 데이터 채움 (수식 인젝션 방지 적용)
    row_idx = 3
    for field in SPEC_FIELDS:
        for idx, item_data in enumerate(items_list):
            col_letter = chr(ord('B') + idx)
            value = get_value(item_data, field, default='')
            escaped_value = escape_excel_formula(value) if value else None
            ws.range(f'{col_letter}{row_idx}').value = escaped_value
        row_idx += 1

    # OPTION_FIELDS 데이터 채움 (수식 인젝션 방지 적용)
    for field in OPTION_FIELDS:
        for idx, item_data in enumerate(items_list):
            col_letter = chr(ord('B') + idx)
            value = get_value(item_data, field, default='')
            escaped_value = escape_excel_formula(value) if value else None
            ws.range(f'{col_letter}{row_idx}').value = escaped_value
        row_idx += 1

    logger.info("Description 시트 생성 완료")


def _create_po_workbook_impl(
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> Path:
    """xlwings 기반으로 PO Workbook 생성 (내부 구현)

    Args:
        order_data: 주문 데이터
        items_df: 다중 아이템 DataFrame (선택)

    Returns:
        생성된 임시 파일 경로
    """
    # 템플릿 확인
    _ensure_template_exists()

    # 날짜
    today = datetime.now()

    # 경로에 한글이 포함된 경우 임시 폴더에서 작업
    temp_dir = Path(tempfile.gettempdir())
    temp_template = temp_dir / f"po_template_{today.strftime('%Y%m%d%H%M%S')}.xlsx"
    temp_output = temp_dir / f"po_output_{today.strftime('%Y%m%d%H%M%S')}.xlsx"

    # 템플릿을 임시 폴더로 복사
    shutil.copy(PO_TEMPLATE_FILE, temp_template)

    # Excel 앱 시작 (백그라운드)
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        # 임시 템플릿 열기
        wb = app.books.open(str(temp_template))

        # Purchase Order 시트
        ws_po = wb.sheets['Purchase Order']
        _create_purchase_order(ws_po, order_data, items_df)

        # Description 시트
        ws_desc = wb.sheets['Description']
        _create_description_sheet(ws_desc, order_data, items_df)

        # Purchase Order 시트를 먼저 보이도록 활성화
        ws_po.activate()

        # 임시 위치에 저장
        wb.save(str(temp_output))
        logger.info(f"PO 생성 완료 (임시): {temp_output}")

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

    return temp_output


class POWorkbook:
    """PO Workbook 래퍼 클래스 (기존 API 호환용)

    create_po.py에서 wb.save(output_file) 패턴을 유지하기 위한 래퍼입니다.
    """

    def __init__(self, temp_file: Path):
        self.temp_file = temp_file

    def save(self, output_file: Path) -> None:
        """임시 파일을 최종 출력 경로로 이동"""
        shutil.move(str(self.temp_file), str(output_file))
        logger.info(f"PO 저장 완료: {output_file}")


def create_po_workbook(
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> POWorkbook:
    """PO Workbook 생성 (공개 API)

    xlwings로 생성 후 POWorkbook 래퍼를 반환합니다.
    사용법: wb = create_po_workbook(order_data)
           wb.save(output_file)

    Args:
        order_data: 주문 데이터
        items_df: 다중 아이템 DataFrame (선택)

    Returns:
        POWorkbook 래퍼 (save 메서드 제공)
    """
    temp_file = _create_po_workbook_impl(order_data, items_df)
    return POWorkbook(temp_file)
