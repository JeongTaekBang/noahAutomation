"""
Excel 생성 모듈
===============

Purchase Order 및 Description 시트를 생성합니다.
"""

from __future__ import annotations

import logging
from datetime import datetime
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

from po_generator.config import (
    COLORS,
    COLUMN_WIDTHS,
    TOTAL_COLUMNS,
    MAX_ITEMS_PER_PO,
    ITEM_START_ROW,
    ITEM_END_ROW,
    SPEC_FIELDS,
    OPTION_FIELDS,
)
from po_generator.utils import get_safe_value

logger = logging.getLogger(__name__)


# === 스타일 정의 ===
THIN_BORDER = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin'),
)

RED_FILL = PatternFill(start_color=COLORS.RED, end_color=COLORS.RED, fill_type="solid")
RED_BRIGHT_FILL = PatternFill(
    start_color=COLORS.RED_BRIGHT, end_color=COLORS.RED_BRIGHT, fill_type="solid"
)
GRAY_FILL = PatternFill(start_color=COLORS.GRAY, end_color=COLORS.GRAY, fill_type="solid")
TEAL_FILL = PatternFill(start_color=COLORS.TEAL, end_color=COLORS.TEAL, fill_type="solid")
GREEN_FILL = PatternFill(start_color=COLORS.GREEN, end_color=COLORS.GREEN, fill_type="solid")


def _create_header_section(
    ws: Worksheet,
    order_data: pd.Series,
    rck_order_no: str,
    today_str: str,
) -> None:
    """헤더 섹션 생성 (Row 1-11)"""
    customer_name = get_safe_value(order_data, 'Customer name')
    incoterms = get_safe_value(order_data, 'Incoterms')

    # 배송 주소 찾기
    delivery_addr = ''
    for col in order_data.index:
        if '납품' in str(col) or '주소' in str(col):
            val = order_data.get(col, '')
            if pd.notna(val) and str(val) != 'nan':
                delivery_addr = str(val)
                break

    # Row 1: 타이틀 (빨간 배경)
    ws['A1'] = f"Purchase Order - {rck_order_no}"
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].alignment = Alignment(vertical='center')
    for col in range(1, TOTAL_COLUMNS + 1):
        ws.cell(row=1, column=col).fill = RED_FILL

    # Row 2-3: Vendor Info + Rotork 로고
    ws['A2'] = "Vendor Name:  NOAH Actuation"
    ws['A2'].font = Font(size=10)
    ws['J2'] = "rotork"
    ws['J2'].font = Font(bold=True, size=14, color=COLORS.RED, italic=True)
    ws['J2'].alignment = Alignment(horizontal='right', vertical='center')

    ws['A3'] = "Vendor Number: -"
    ws['A3'].font = Font(size=10)
    ws['J3'] = "Rotork Korea"
    ws['J3'].font = Font(size=10)
    ws['J3'].alignment = Alignment(horizontal='right', vertical='top')

    # Row 4-5: Vendor Reference + Delivery Address
    ws['A4'] = "Vendor Reference: -"
    ws['A4'].font = Font(size=10)
    ws['C4'] = "Delivery Address:"
    ws['C4'].font = Font(size=9, bold=True, italic=True)

    ws['A5'] = f"Date:  {today_str}"
    ws['A5'].font = Font(size=10)
    ws['C5'] = delivery_addr
    ws['C5'].font = Font(size=9, italic=True)
    ws['C5'].alignment = Alignment(wrap_text=True)

    # Row 6-10: Supplier/Customer Info
    ws['A6'] = "Supplier address:"
    ws['A6'].font = Font(size=9, bold=True)

    ws['A7'] = "NOAH Actuation"
    ws['A7'].font = Font(size=9)

    ws['A8'] = "11 Jeongseojin-9ro, Seo-gu, Incheon, Korea(22850)"
    ws['A8'].font = Font(size=9)

    ws['A9'] = "Final Customer"
    ws['A9'].font = Font(size=9, bold=True)

    ws['A10'] = customer_name
    ws['A10'].font = Font(size=9)

    ws['A11'] = "Order Details"
    ws['A11'].font = Font(size=9, bold=True)


def _create_item_header(ws: Worksheet) -> None:
    """아이템 헤더 생성 (Row 12)"""
    ws.merge_cells('B12:E12')

    headers = [
        ('A12', 'Item\nNumber'),
        ('B12', 'Description'),
        ('F12', 'Qty'),
        ('G12', 'Unit'),
        ('H12', 'Unit price'),
        ('I12', 'Delivery\nRequired'),
        ('J12', 'Amount'),
    ]

    header_font = Font(bold=True, size=10, color="FFFFFF")

    for cell_addr, text in headers:
        ws[cell_addr] = text
        ws[cell_addr].font = header_font
        ws[cell_addr].fill = GRAY_FILL
        ws[cell_addr].alignment = Alignment(
            wrap_text=True, horizontal='center', vertical='center'
        )
        ws[cell_addr].border = THIN_BORDER

    # Row 12 전체 회색 배경
    for col in range(1, TOTAL_COLUMNS + 1):
        ws.cell(row=12, column=col).fill = GRAY_FILL
        ws.cell(row=12, column=col).border = THIN_BORDER


def _create_item_rows(
    ws: Worksheet,
    items_list: list[pd.Series],
    currency: str = 'KRW',
) -> None:
    """아이템 데이터 행 생성 (Row 13-19)"""
    number_format = '₩#,##0' if currency == 'KRW' else '$#,##0.00'

    for item_idx, item_data in enumerate(items_list[:MAX_ITEMS_PER_PO]):
        row_num = ITEM_START_ROW + item_idx

        # 데이터 추출
        model = get_safe_value(item_data, 'Model')
        power = get_safe_value(item_data, 'Power supply')
        item_name = get_safe_value(item_data, 'Item name')

        # Description 조합
        desc_parts = []
        if item_name:
            desc_parts.append(item_name)
        elif model:
            desc_parts.append(model)

        if power:
            desc_parts.append(power.replace('-1Ph-', ', ').replace('-3Ph-', ', '))

        if str(get_safe_value(item_data, 'ALS')).upper() == 'Y':
            desc_parts.append('ALS')

        description = ', '.join([p for p in desc_parts if p])

        # 수량
        try:
            qty = int(float(get_safe_value(item_data, 'Item qty', 1)))
        except (ValueError, TypeError):
            qty = 1

        # 단가
        try:
            ico_unit = float(get_safe_value(item_data, 'ICO Unit', 0))
        except (ValueError, TypeError):
            ico_unit = 0

        # 납기일
        requested_date = get_safe_value(item_data, 'Requested delivery date')
        requested_date_str = ''
        if requested_date and not pd.isna(requested_date):
            try:
                if isinstance(requested_date, datetime):
                    requested_date_str = requested_date.strftime("%Y-%m-%d")
                else:
                    requested_date_str = str(requested_date)[:10]
            except (ValueError, TypeError):
                requested_date_str = ''

        # 셀 병합 및 데이터 입력
        ws.merge_cells(f'B{row_num}:E{row_num}')

        ws[f'A{row_num}'] = item_idx + 1
        ws[f'A{row_num}'].alignment = Alignment(horizontal='center')

        ws[f'B{row_num}'] = description

        ws[f'F{row_num}'] = qty
        ws[f'F{row_num}'].alignment = Alignment(horizontal='center')

        ws[f'G{row_num}'] = "EA"
        ws[f'G{row_num}'].alignment = Alignment(horizontal='center')

        ws[f'H{row_num}'] = ico_unit
        ws[f'H{row_num}'].number_format = number_format
        ws[f'H{row_num}'].alignment = Alignment(horizontal='right')

        ws[f'I{row_num}'] = requested_date_str
        ws[f'I{row_num}'].alignment = Alignment(horizontal='center')

        ws[f'J{row_num}'] = f"=H{row_num}*F{row_num}"
        ws[f'J{row_num}'].number_format = number_format
        ws[f'J{row_num}'].alignment = Alignment(horizontal='right')

        # 테두리
        for col in range(1, TOTAL_COLUMNS + 1):
            ws.cell(row=row_num, column=col).border = THIN_BORDER

    # 빈 아이템 행 (나머지)
    start_empty_row = ITEM_START_ROW + len(items_list[:MAX_ITEMS_PER_PO])
    for row in range(start_empty_row, ITEM_END_ROW + 1):
        ws.merge_cells(f'B{row}:E{row}')
        for col in range(1, TOTAL_COLUMNS + 1):
            ws.cell(row=row, column=col).border = THIN_BORDER


def _create_totals_section(ws: Worksheet, currency: str = 'KRW') -> None:
    """합계 섹션 생성 (Row 20-22)"""
    number_format = '₩#,##0' if currency == 'KRW' else '$#,##0.00'
    title_font = Font(bold=True, size=11)

    # Row 20: Total net amount
    ws['I20'] = "Total net amount"
    ws['I20'].font = title_font
    ws['I20'].alignment = Alignment(horizontal='right')
    ws['J20'] = "=SUM(J13:J19)"
    ws['J20'].number_format = number_format
    ws['J20'].alignment = Alignment(horizontal='right')
    ws['J20'].border = THIN_BORDER

    # Row 21: VAT
    ws['I21'] = "VAT"
    ws['I21'].alignment = Alignment(horizontal='right')
    ws['J21'] = "=J20*0.1"
    ws['J21'].number_format = number_format
    ws['J21'].alignment = Alignment(horizontal='right')
    ws['J21'].border = THIN_BORDER

    # Row 22: Order Total
    ws['I22'] = "Order Total"
    ws['I22'].font = title_font
    ws['I22'].alignment = Alignment(horizontal='right')
    ws['J22'] = "=SUM(J20:J21)"
    ws['J22'].number_format = number_format
    ws['J22'].font = Font(bold=True)
    ws['J22'].alignment = Alignment(horizontal='right')
    ws['J22'].border = THIN_BORDER


def _create_footer_section(ws: Worksheet, order_data: pd.Series) -> None:
    """푸터 섹션 생성 (Row 23-31)"""
    remark = get_safe_value(order_data, 'Remark')
    incoterms = get_safe_value(order_data, 'Incoterms')
    currency = 'KRW'

    # Row 23-25: Project Info
    ws['A23'] = "D365CEProject:"
    ws['C23'] = "Industry:"
    ws['A24'] = "Project name:"
    ws['C24'] = "Valve type:"
    ws['C25'] = "Final country:"

    # Row 26: Additional information
    ws['A26'] = "Additional information"
    ws['C26'] = f"Note. {remark}" if remark else "Note."
    ws['C26'].alignment = Alignment(wrap_text=True, vertical='top')
    ws['H26'] = "On behalf of Rotork"

    # Row 27: Currency
    ws['A27'] = "Order Currency:"
    ws['B27'] = currency
    ws['H27'] = "Contact:"

    # Row 28: Delivery Terms
    ws['A28'] = "Delivery Terms:"
    ws['B28'] = incoterms

    # Row 29: Delivery mode
    ws['A29'] = "Delivery mode:"
    ws['H29'] = "Email:"

    # Row 30: Payment Terms
    ws['A30'] = "Terms of payment:"
    ws['H30'] = "Tel:"

    # Row 31: Footer (청록색 배경)
    ws['A31'] = "Keeping the World flowing for Future Generations"
    ws['A31'].font = Font(size=12, color="FFFFFF", italic=True)
    ws['A31'].alignment = Alignment(vertical='center')
    ws['J31'] = "1 of 1"
    ws['J31'].font = Font(size=12, color="FFFFFF")
    ws['J31'].alignment = Alignment(horizontal='right', vertical='center')

    for col in range(1, TOTAL_COLUMNS + 1):
        ws.cell(row=31, column=col).fill = TEAL_FILL


def _apply_layout_settings(ws: Worksheet) -> None:
    """레이아웃 설정 적용 (열 너비, 행 높이, 인쇄 설정)"""
    # 열 너비
    for col, width in COLUMN_WIDTHS.as_dict().items():
        ws.column_dimensions[col].width = width

    # 행 높이
    ws.row_dimensions[1].height = 25
    ws.row_dimensions[12].height = 30
    ws.row_dimensions[31].height = 25

    # 인쇄 설정
    ws.print_area = 'A1:J31'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.orientation = 'portrait'
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5


def create_purchase_order(
    ws: Worksheet,
    order_data: pd.Series,
    items_df: Optional[pd.DataFrame] = None,
) -> None:
    """Purchase Order 시트 생성

    Args:
        ws: 워크시트
        order_data: 첫 번째 아이템 데이터 (공통 정보 추출용)
        items_df: 다중 아이템인 경우 DataFrame, 단일이면 None
    """
    logger.info("Purchase Order 시트 생성 중...")

    # 아이템 목록 준비
    if items_df is not None:
        items_list = [row for _, row in items_df.iterrows()]
    else:
        items_list = [order_data]

    # 공통 데이터
    rck_order_no = get_safe_value(order_data, 'RCK Order no.')
    today_str = datetime.now().strftime("%d/%b/%Y").upper()
    currency = 'KRW'

    # 각 섹션 생성
    _create_header_section(ws, order_data, rck_order_no, today_str)
    _create_item_header(ws)
    _create_item_rows(ws, items_list, currency)
    _create_totals_section(ws, currency)
    _create_footer_section(ws, order_data)
    _apply_layout_settings(ws)

    logger.info("Purchase Order 시트 생성 완료")


def create_description_sheet(
    ws: Worksheet,
    order_data: pd.Series,
    items_df: Optional[pd.DataFrame] = None,
) -> None:
    """Description 시트 생성 (사양 정보)

    Args:
        ws: 워크시트
        order_data: 첫 번째 아이템 데이터
        items_df: 다중 아이템인 경우 DataFrame, 단일이면 None
    """
    logger.info("Description 시트 생성 중...")

    white_bold_font = Font(bold=True, color="FFFFFF")

    # 아이템 목록 준비
    if items_df is not None:
        items_list = [row for _, row in items_df.iterrows()]
    else:
        items_list = [order_data]

    num_items = len(items_list)

    # Row 1: Line No 헤더
    ws['A1'] = "Line No"
    ws['A1'].font = white_bold_font
    ws['A1'].fill = GREEN_FILL
    ws['A1'].border = THIN_BORDER
    ws['A1'].alignment = Alignment(horizontal='center')

    # 각 아이템에 대해 Line No 출력
    for idx in range(num_items):
        col_num = 2 + idx
        ws.cell(row=1, column=col_num, value=idx + 1)
        ws.cell(row=1, column=col_num).border = THIN_BORDER
        ws.cell(row=1, column=col_num).alignment = Alignment(horizontal='center')

    row_idx = 2

    # 액추에이터 사양 필드 (초록 배경)
    for field in SPEC_FIELDS:
        ws.cell(row=row_idx, column=1, value=field)
        ws.cell(row=row_idx, column=1).font = white_bold_font
        ws.cell(row=row_idx, column=1).fill = GREEN_FILL
        ws.cell(row=row_idx, column=1).border = THIN_BORDER

        for idx, item_data in enumerate(items_list):
            col_num = 2 + idx
            value = get_safe_value(item_data, field)
            ws.cell(row=row_idx, column=col_num, value=value if value else None)
            ws.cell(row=row_idx, column=col_num).border = THIN_BORDER

        row_idx += 1

    # 옵션 필드 (빨간 배경)
    for field in OPTION_FIELDS:
        ws.cell(row=row_idx, column=1, value=field)
        ws.cell(row=row_idx, column=1).font = white_bold_font
        ws.cell(row=row_idx, column=1).fill = RED_BRIGHT_FILL
        ws.cell(row=row_idx, column=1).border = THIN_BORDER

        for idx, item_data in enumerate(items_list):
            col_num = 2 + idx
            value = get_safe_value(item_data, field)
            ws.cell(row=row_idx, column=col_num, value=value if value else None)
            ws.cell(row=row_idx, column=col_num).border = THIN_BORDER

        row_idx += 1

    # 열 너비 조정
    ws.column_dimensions['A'].width = 25
    for idx in range(num_items):
        col_letter = get_column_letter(2 + idx)
        ws.column_dimensions[col_letter].width = 15

    logger.info("Description 시트 생성 완료")
