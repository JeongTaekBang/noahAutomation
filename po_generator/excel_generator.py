"""
Excel 생성 모듈 (xlwings 기반)
===============================

xlwings를 사용하여 템플릿 기반으로 Purchase Order 및 Description 시트를 생성합니다.
이미지, 서식 등이 완벽하게 보존됩니다.
"""

from __future__ import annotations

import logging
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.config import (
    TOTAL_COLUMNS,
    ITEM_START_ROW_FALLBACK,
    VAT_RATE_DOMESTIC,
    PO_TEMPLATE_FILE,
)
from po_generator.utils import (
    get_value,
    escape_excel_formula,
    get_spec_option_fields,
)
from po_generator.excel_helpers import (
    XlConstants,
    xlwings_app_context,
    prepare_template,
    cleanup_temp_file,
    find_item_start_row_xlwings,
    PO_HEADER_LABELS,
    batch_write_rows,
    delete_rows_range,
    find_text_in_column_batch,
)

logger = logging.getLogger(__name__)


# === 템플릿 셀 매핑 ===
CELL_TITLE = 'A1'
CELL_DATE = 'A5'
CELL_DELIVERY_ADDR = 'C5'
CELL_CUSTOMER_PO = 'C7'
CELL_CUSTOMER_NAME = 'A10'


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
    """'Total net amount' 텍스트가 있는 행 찾기 (배치 읽기 최적화)

    Args:
        ws: xlwings Sheet 객체
        start_row: 검색 시작 행
        max_search: 최대 검색 행 수

    Returns:
        Total net amount 행 번호 (못 찾으면 start_row)
    """
    # 배치 읽기로 20회 COM 호출 → 1회로 감소
    end_row = start_row + max_search - 1
    row = find_text_in_column_batch(ws, 'I', 'Total net', start_row, end_row)
    return row if row is not None else start_row


def _create_purchase_order(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> None:
    """Purchase Order 시트 생성 (xlwings 기반) - 배치 쓰기 최적화

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
    template_row = find_item_start_row_xlwings(
        ws,
        search_labels=PO_HEADER_LABELS,
        fallback_row=ITEM_START_ROW_FALLBACK,
    )

    # 3. 템플릿 아이템 행 수 계산 (Total net amount 행 찾기)
    totals_row = _find_totals_row(ws, template_row)
    template_item_count = totals_row - template_row
    logger.debug(f"템플릿 아이템 수: {template_item_count}, 실제 아이템 수: {num_items}")

    # 4. 행 수 조정
    if num_items < template_item_count:
        # 초과 행 삭제 - 범위 삭제로 N회 COM 호출 → 1회로 감소
        rows_to_delete = template_item_count - num_items
        delete_rows_range(ws, template_row + num_items, rows_to_delete)
    elif num_items > template_item_count:
        # 행 삽입
        rows_to_insert = num_items - template_item_count
        for i in range(rows_to_insert):
            ws.range(f'{template_row}:{template_row}').api.Copy()
            insert_row = template_row + template_item_count + i
            ws.range(f'{insert_row}:{insert_row}').api.Insert(Shift=XlConstants.xlShiftDown)
        logger.debug(f"{rows_to_insert}개 행 삽입")

    # 5. 아이템 데이터 배치 채움 (N개 아이템 * 9열 COM 호출 → 최적화)
    _fill_items_batch_po(ws, template_row, items_list, currency, is_export)

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


def _fill_items_batch_po(
    ws: xw.Sheet,
    start_row: int,
    items_list: list,
    currency: str,
    is_export: bool,
) -> None:
    """PO 아이템 데이터 배치 쓰기 (성능 최적화)

    PO는 열이 불연속적이고 수식/포맷이 있어서 그룹별로 배치 처리

    Args:
        ws: xlwings Sheet 객체
        start_row: 아이템 시작 행
        items_list: 아이템 리스트
        currency: 통화 코드
        is_export: 해외 여부
    """
    num_items = len(items_list)
    end_row = start_row + num_items - 1
    number_format = '₩#,##0' if currency == 'KRW' else '$#,##0.00'

    # 데이터 준비
    col_a = []  # No.
    col_b = []  # Description
    col_f = []  # Qty
    col_g = []  # Unit (EA)
    col_h = []  # Unit Price
    col_i = []  # Requested Date

    for item_idx, item_data in enumerate(items_list):
        # No.
        col_a.append(item_idx + 1)

        # Description
        item_name = get_value(item_data, 'item_name')
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
        col_b.append(escape_excel_formula(description))

        # Qty
        raw_qty = get_value(item_data, 'item_qty', 1)
        try:
            qty = int(float(raw_qty)) if raw_qty is not None else 1
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 수량 변환 실패 '{raw_qty}' -> 기본값 1 사용")
            qty = 1
        col_f.append(qty)

        # Unit
        col_g.append("EA")

        # Unit Price
        raw_price = get_value(item_data, 'ico_unit', 0)
        try:
            ico_unit = float(raw_price) if raw_price is not None else 0
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 단가 변환 실패 '{raw_price}' -> 기본값 0 사용")
            ico_unit = 0
        col_h.append(ico_unit)

        # Requested Date
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
        col_i.append(requested_date_str)

    # 배치 쓰기 (6회 COM 호출 - 아이템 수에 관계없이 고정)
    ws.range(f'A{start_row}:A{end_row}').value = [[v] for v in col_a]
    ws.range(f'B{start_row}:B{end_row}').value = [[v] for v in col_b]
    ws.range(f'F{start_row}:F{end_row}').value = [[v] for v in col_f]
    ws.range(f'G{start_row}:G{end_row}').value = [[v] for v in col_g]
    ws.range(f'H{start_row}:H{end_row}').value = [[v] for v in col_h]
    ws.range(f'I{start_row}:I{end_row}').value = [[v] for v in col_i]

    # H열 숫자 포맷 (범위 포맷은 1회 COM 호출)
    ws.range(f'H{start_row}:H{end_row}').number_format = number_format

    # J열 수식 채우기 (수식은 개별 셀에 설정해야 함)
    # 하지만 배치로 2D 리스트의 수식을 쓸 수 있음
    formulas = [[f'=H{start_row + i}*F{start_row + i}'] for i in range(num_items)]
    ws.range(f'J{start_row}:J{end_row}').formula = formulas
    ws.range(f'J{start_row}:J{end_row}').number_format = number_format

    logger.debug(f"PO 아이템 배치 쓰기 완료: {num_items}개")


def _apply_description_borders(
    ws: xw.Sheet,
    num_items: int,
    num_rows: int,
) -> None:
    """Description 시트에 테두리 적용

    Args:
        ws: xlwings Sheet 객체
        num_items: 아이템(열) 수
        num_rows: 행 수 (레이블 포함)
    """
    # 열 범위 계산 (B열부터)
    end_col_num = ord('B') + num_items - 1
    if end_col_num <= ord('Z'):
        end_col = chr(end_col_num)
    else:
        end_col = 'A' + chr(ord('A') + (end_col_num - ord('Z') - 1))

    # 전체 데이터 영역에 테두리 적용 (A1부터)
    data_range = ws.range(f'A1:{end_col}{num_rows}')

    # 외곽 테두리
    for edge in [XlConstants.xlEdgeLeft, XlConstants.xlEdgeTop,
                 XlConstants.xlEdgeBottom, XlConstants.xlEdgeRight]:
        data_range.api.Borders(edge).LineStyle = XlConstants.xlContinuous
        data_range.api.Borders(edge).Weight = XlConstants.xlThin

    # 내부 세로선
    if num_items > 1:
        data_range.api.Borders(XlConstants.xlInsideVertical).LineStyle = XlConstants.xlContinuous
        data_range.api.Borders(XlConstants.xlInsideVertical).Weight = XlConstants.xlThin

    # 내부 가로선
    if num_rows > 1:
        data_range.api.Borders(XlConstants.xlInsideHorizontal).LineStyle = XlConstants.xlContinuous
        data_range.api.Borders(XlConstants.xlInsideHorizontal).Weight = XlConstants.xlThin

    logger.debug(f"Description 테두리 적용: A1:{end_col}{num_rows}")


def _create_description_sheet(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> None:
    """Description 시트 생성 (xlwings 기반) - 배치 쓰기 최적화

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

    # 시트 구분 확인 (국내/해외)
    sheet_type = get_value(order_data, 'sheet_type', '국내')

    # 동적으로 SPEC/OPTION 필드 가져오기
    spec_fields, option_fields = get_spec_option_fields(sheet_type)
    all_fields = spec_fields + option_fields

    # 헤더 데이터 준비 (Row 1: Line No, Row 2: Qty)
    header_row1 = []  # Line No
    header_row2 = []  # Qty
    for idx, item_data in enumerate(items_list):
        header_row1.append(idx + 1)
        try:
            qty = int(float(get_value(item_data, 'item_qty', 1)))
        except (ValueError, TypeError):
            qty = 1
        header_row2.append(qty)

    # A열에 레이블 쓰기 (동적 필드 기준)
    labels = ['Line No', 'Qty'] + all_fields
    labels_2d = [[label] for label in labels]
    ws.range(f'A1:A{len(labels)}').value = labels_2d

    # B열부터 값 쓰기
    end_col = chr(ord('B') + num_items - 1) if num_items <= 24 else None
    if num_items > 24:
        # 26개 이상인 경우 Excel 열 문자 계산
        end_col_num = ord('B') + num_items - 1
        if end_col_num <= ord('Z'):
            end_col = chr(end_col_num)
        else:
            # AA, AB, ... 처리
            end_col = 'A' + chr(ord('A') + (end_col_num - ord('Z') - 1))

    # 헤더 배치 쓰기 (2회 COM 호출)
    ws.range(f'B1:{end_col}1').value = [header_row1]
    ws.range(f'B2:{end_col}2').value = [header_row2]

    # 필드 데이터 준비
    data_2d = []
    for field in all_fields:
        row_data = []
        for item_data in items_list:
            value = get_value(item_data, field, default='')
            escaped_value = escape_excel_formula(value) if value else None
            row_data.append(escaped_value)
        data_2d.append(row_data)

    # SPEC/OPTION 필드 배치 쓰기 (1회 COM 호출 - 30+N 필드 * M 아이템)
    if data_2d:
        num_rows = len(data_2d)
        end_row = 3 + num_rows - 1
        ws.range(f'B3:{end_col}{end_row}').value = data_2d

    # 테두리 적용 (전체 데이터 영역)
    _apply_description_borders(ws, num_items, len(labels))

    logger.info(f"Description 시트 생성 완료 ({len(all_fields)}개 필드 x {num_items}개 아이템)")


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
    # 템플릿 준비 (임시 폴더로 복사)
    temp_template, temp_output = prepare_template(PO_TEMPLATE_FILE, "po")

    try:
        # xlwings App 생명주기 관리
        with xlwings_app_context() as app:
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
        # 임시 템플릿 삭제
        cleanup_temp_file(temp_template)

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
