"""
템플릿 엔진 모듈
================

템플릿 파일을 로드하고, 동적 행 처리, 공식 조정 등을 담당합니다.
"""

from __future__ import annotations

import re
import logging
from copy import copy
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.utils import get_column_letter

from po_generator.config import PO_TEMPLATE_FILE, TEMPLATE_DIR

logger = logging.getLogger(__name__)


def load_template(template_path: Optional[Path] = None) -> Workbook:
    """템플릿 파일 로드 (이미지 포함)

    Args:
        template_path: 템플릿 파일 경로 (None이면 기본 PO 템플릿)

    Returns:
        로드된 Workbook

    Raises:
        FileNotFoundError: 템플릿 파일이 없는 경우
    """
    if template_path is None:
        template_path = PO_TEMPLATE_FILE

    if not template_path.exists():
        raise FileNotFoundError(f"템플릿 파일을 찾을 수 없습니다: {template_path}")

    logger.info(f"템플릿 로드: {template_path}")
    # 이미지/차트 등 모든 요소를 포함하여 로드
    return load_workbook(template_path, data_only=False)


def copy_cell_style(source: Cell, target: Cell) -> None:
    """셀 스타일 복사 (폰트, 배경, 테두리, 정렬, 숫자 포맷)

    Args:
        source: 원본 셀
        target: 대상 셀
    """
    if source.has_style:
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)


def clone_row(
    ws: Worksheet,
    source_row: int,
    target_row: int,
    max_col: int = 10,
) -> None:
    """행 복제 (값, 스타일, 병합 포함)

    Args:
        ws: 워크시트
        source_row: 원본 행 번호
        target_row: 대상 행 번호
        max_col: 복제할 최대 열 수
    """
    # 원본 행의 병합 정보 수집
    merges_to_add = []
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row == source_row and merged_range.max_row == source_row:
            # 같은 행 내 병합 (예: B13:E13)
            new_merge = f"{get_column_letter(merged_range.min_col)}{target_row}:" \
                        f"{get_column_letter(merged_range.max_col)}{target_row}"
            merges_to_add.append(new_merge)

    # 셀 복사
    for col in range(1, max_col + 1):
        source_cell = ws.cell(row=source_row, column=col)
        target_cell = ws.cell(row=target_row, column=col)

        # MergedCell이 아닌 경우에만 값 복사
        if not isinstance(source_cell, MergedCell):
            # 공식인 경우 행 번호 조정
            if source_cell.value and isinstance(source_cell.value, str) \
               and source_cell.value.startswith('='):
                target_cell.value = adjust_formula_row(
                    source_cell.value, source_row, target_row
                )
            else:
                target_cell.value = source_cell.value

            # 스타일 복사
            copy_cell_style(source_cell, target_cell)

    # 새 병합 적용
    for merge_range in merges_to_add:
        ws.merge_cells(merge_range)

    # 행 높이 복사
    if ws.row_dimensions[source_row].height:
        ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def adjust_formula_row(formula: str, old_row: int, new_row: int) -> str:
    """공식 내 행 참조를 조정

    예: =H13*F13 → =H14*F14 (old_row=13, new_row=14)

    Args:
        formula: 원본 공식
        old_row: 원본 행 번호
        new_row: 새 행 번호

    Returns:
        조정된 공식
    """
    # 행 번호 패턴: 열 문자 + 숫자 (예: A13, J13)
    pattern = r'([A-Z]+)' + str(old_row) + r'(?![0-9])'
    replacement = r'\g<1>' + str(new_row)
    return re.sub(pattern, replacement, formula, flags=re.IGNORECASE)


def shift_formulas_in_range(
    ws: Worksheet,
    start_row: int,
    end_row: int,
    shift_by: int,
    reference_start_row: int,
    max_col: int = 10,
) -> None:
    """특정 범위 내 공식의 행 참조를 이동

    Args:
        ws: 워크시트
        start_row: 시작 행
        end_row: 끝 행
        shift_by: 이동할 행 수
        reference_start_row: 이 행 이상의 참조만 이동
        max_col: 최대 열 수
    """
    for row in range(start_row, end_row + 1):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                cell.value = shift_formula_references(
                    cell.value, shift_by, reference_start_row
                )


def shift_formula_references(
    formula: str,
    shift_by: int,
    reference_start_row: int,
) -> str:
    """공식 내 행 참조를 shift_by만큼 이동

    reference_start_row 이상의 행 참조만 이동합니다.

    Args:
        formula: 원본 공식
        shift_by: 이동할 행 수 (양수: 아래로, 음수: 위로)
        reference_start_row: 이 행 이상의 참조만 이동

    Returns:
        조정된 공식
    """
    def replace_row(match):
        col = match.group(1)
        row = int(match.group(2))
        if row >= reference_start_row:
            return f"{col}{row + shift_by}"
        return match.group(0)

    # 패턴: 열 문자 + 숫자 (예: A13, J15)
    pattern = r'([A-Z]+)(\d+)'
    return re.sub(pattern, replace_row, formula, flags=re.IGNORECASE)


def _backup_images(ws: Worksheet) -> list:
    """워크시트의 이미지 정보를 백업

    Args:
        ws: 워크시트

    Returns:
        이미지 객체 리스트
    """
    images = []
    if hasattr(ws, '_images'):
        # 이미지 객체와 앵커 정보를 복사
        for img in ws._images:
            images.append(img)
    return images


def _restore_images(ws: Worksheet, images: list) -> None:
    """백업된 이미지를 워크시트에 복원

    Args:
        ws: 워크시트
        images: 백업된 이미지 리스트
    """
    # 기존 이미지 제거 후 복원
    ws._images = images


def insert_rows_with_template(
    ws: Worksheet,
    template_row: int,
    count: int,
    max_col: int = 10,
) -> int:
    """템플릿 행을 복제하여 여러 행 삽입

    템플릿 행 아래에 (count - 1)개의 행을 삽입하고 스타일을 복제합니다.
    기존 행들은 아래로 이동합니다.
    이미지는 원래 위치에 유지됩니다.

    Args:
        ws: 워크시트
        template_row: 템플릿 행 번호 (이 행이 첫 번째 아이템 행이 됨)
        count: 필요한 총 행 수
        max_col: 최대 열 수

    Returns:
        마지막 아이템 행 번호
    """
    if count <= 1:
        return template_row

    rows_to_insert = count - 1
    insert_position = template_row + 1

    logger.debug(f"행 삽입: template_row={template_row}, count={count}, "
                 f"rows_to_insert={rows_to_insert}")

    # 1. 이미지 백업 (insert_rows가 이미지를 손상시킬 수 있음)
    images_backup = _backup_images(ws)

    # 2. 새 행 삽입 (기존 행들이 아래로 밀림)
    ws.insert_rows(insert_position, rows_to_insert)

    # 3. 이미지 복원
    if images_backup:
        _restore_images(ws, images_backup)

    # 4. 삽입된 행들에 템플릿 스타일 복제
    for i in range(rows_to_insert):
        target_row = insert_position + i
        clone_row(ws, template_row, target_row, max_col)

    return template_row + rows_to_insert


def update_sum_formula(
    ws: Worksheet,
    cell_address: str,
    item_start_row: int,
    item_end_row: int,
    column: str = 'J',
) -> None:
    """SUM 공식의 범위 업데이트

    Args:
        ws: 워크시트
        cell_address: 공식이 있는 셀 주소 (예: 'J14')
        item_start_row: 아이템 시작 행
        item_end_row: 아이템 끝 행
        column: 합계할 열 (기본: 'J')
    """
    formula = f"=SUM({column}{item_start_row}:{column}{item_end_row})"
    ws[cell_address] = formula
    logger.debug(f"SUM 공식 업데이트: {cell_address} = {formula}")


def ensure_template_dir() -> None:
    """템플릿 디렉토리가 없으면 생성"""
    TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)


def generate_po_template() -> Path:
    """Purchase Order 템플릿 파일 생성

    현재 코드 기반 레이아웃으로 템플릿을 생성합니다.
    이미지는 사용자가 직접 추가할 수 있습니다.

    Returns:
        생성된 템플릿 파일 경로
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    from po_generator.config import (
        COLORS, COLUMN_WIDTHS, TOTAL_COLUMNS,
        SPEC_FIELDS, OPTION_FIELDS,
    )

    ensure_template_dir()

    wb = Workbook()
    ws = wb.active
    ws.title = "Purchase Order"

    # 스타일 정의
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    red_fill = PatternFill(start_color=COLORS.RED, end_color=COLORS.RED, fill_type="solid")
    gray_fill = PatternFill(start_color=COLORS.GRAY, end_color=COLORS.GRAY, fill_type="solid")
    teal_fill = PatternFill(start_color=COLORS.TEAL, end_color=COLORS.TEAL, fill_type="solid")
    green_fill = PatternFill(start_color=COLORS.GREEN, end_color=COLORS.GREEN, fill_type="solid")
    red_bright_fill = PatternFill(
        start_color=COLORS.RED_BRIGHT, end_color=COLORS.RED_BRIGHT, fill_type="solid"
    )

    white_font = Font(color="FFFFFF")
    white_bold_font = Font(bold=True, color="FFFFFF")

    # === Row 1: 타이틀 (빨간 배경) ===
    ws['A1'] = "Purchase Order - "  # RCK Order no. 가 붙음
    ws['A1'].font = Font(bold=True, size=14, color="FFFFFF")
    ws['A1'].alignment = Alignment(vertical='center')
    for col in range(1, TOTAL_COLUMNS + 1):
        ws.cell(row=1, column=col).fill = red_fill

    # === Row 2-3: Vendor Info + Rotork 로고 ===
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

    # === Row 4-5: Vendor Reference + Delivery Address ===
    ws['A4'] = "Vendor Reference: -"
    ws['A4'].font = Font(size=10)
    ws['C4'] = "Delivery Address:"
    ws['C4'].font = Font(size=9, bold=True, italic=True)

    ws['A5'] = "Date:  "  # 날짜가 붙음
    ws['A5'].font = Font(size=10)
    ws['C5'] = ""  # Delivery Address 값
    ws['C5'].font = Font(size=9, italic=True)
    ws['C5'].alignment = Alignment(wrap_text=True)

    # === Row 6-11: Supplier/Customer Info ===
    ws['A6'] = "Supplier address:"
    ws['A6'].font = Font(size=9, bold=True)
    ws['C6'] = "Customer PO:"
    ws['C6'].font = Font(size=9, bold=True)

    ws['A7'] = "NOAH Actuation"
    ws['A7'].font = Font(size=9)
    ws['C7'] = ""  # Customer PO 값
    ws['C7'].font = Font(size=9)

    ws['A8'] = "11 Jeongseojin-9ro, Seo-gu, Incheon, Korea(22850)"
    ws['A8'].font = Font(size=9)

    ws['A9'] = "Final Customer"
    ws['A9'].font = Font(size=9, bold=True)

    ws['A10'] = ""  # Customer name 값
    ws['A10'].font = Font(size=9)

    ws['A11'] = "Order Details"
    ws['A11'].font = Font(size=9, bold=True)

    # === Row 12: 아이템 헤더 (회색 배경) ===
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
        ws[cell_addr].fill = gray_fill
        ws[cell_addr].alignment = Alignment(
            wrap_text=True, horizontal='center', vertical='center'
        )
        ws[cell_addr].border = thin_border

    for col in range(1, TOTAL_COLUMNS + 1):
        ws.cell(row=12, column=col).fill = gray_fill
        ws.cell(row=12, column=col).border = thin_border

    # === Row 13: 아이템 템플릿 행 (복제될 행) ===
    ws.merge_cells('B13:E13')
    ws['A13'] = 1
    ws['A13'].alignment = Alignment(horizontal='center')
    ws['B13'] = ""  # Description
    ws['F13'] = 1   # Qty
    ws['F13'].alignment = Alignment(horizontal='center')
    ws['G13'] = "EA"
    ws['G13'].alignment = Alignment(horizontal='center')
    ws['H13'] = 0   # Unit price
    ws['H13'].number_format = '₩#,##0'
    ws['H13'].alignment = Alignment(horizontal='right')
    ws['I13'] = ""  # Delivery date
    ws['I13'].alignment = Alignment(horizontal='center')
    ws['J13'] = "=H13*F13"  # Amount 공식
    ws['J13'].number_format = '₩#,##0'
    ws['J13'].alignment = Alignment(horizontal='right')

    for col in range(1, TOTAL_COLUMNS + 1):
        ws.cell(row=13, column=col).border = thin_border

    # === Row 14-16: 합계 섹션 ===
    title_font = Font(bold=True, size=11)

    ws['I14'] = "Total net amount"
    ws['I14'].font = title_font
    ws['I14'].alignment = Alignment(horizontal='right')
    ws['J14'] = "=SUM(J13:J13)"  # 동적으로 조정됨
    ws['J14'].number_format = '₩#,##0'
    ws['J14'].alignment = Alignment(horizontal='right')
    ws['J14'].border = thin_border

    ws['I15'] = "VAT"
    ws['I15'].alignment = Alignment(horizontal='right')
    ws['J15'] = "=J14*0.1"  # 해외는 0으로 변경됨
    ws['J15'].number_format = '₩#,##0'
    ws['J15'].alignment = Alignment(horizontal='right')
    ws['J15'].border = thin_border

    ws['I16'] = "Order Total"
    ws['I16'].font = title_font
    ws['I16'].alignment = Alignment(horizontal='right')
    ws['J16'] = "=SUM(J14:J15)"
    ws['J16'].number_format = '₩#,##0'
    ws['J16'].font = Font(bold=True)
    ws['J16'].alignment = Alignment(horizontal='right')
    ws['J16'].border = thin_border

    # === Row 17-25: 푸터 섹션 ===
    r = 17
    ws[f'A{r}'] = "D365CEProject:"
    ws[f'C{r}'] = "Opportunity:"
    ws[f'D{r}'] = ""  # Opportunity 값

    r += 1
    ws[f'A{r}'] = "Project name:"
    ws[f'C{r}'] = "Sector:"
    ws[f'D{r}'] = ""  # Sector 값

    r += 1
    ws[f'C{r}'] = "Industry code:"
    ws[f'D{r}'] = ""  # Industry code 값

    r += 1
    ws[f'A{r}'] = "Additional information"
    ws[f'C{r}'] = "Note."  # Remark가 붙음
    ws[f'C{r}'].alignment = Alignment(wrap_text=True, vertical='top')
    ws[f'H{r}'] = "On behalf of Rotork"

    r += 1
    ws[f'A{r}'] = "Order Currency:"
    ws[f'B{r}'] = "KRW"  # Currency 값
    ws[f'H{r}'] = "Contact:"

    r += 1
    ws[f'A{r}'] = "Delivery Terms:"
    ws[f'B{r}'] = ""  # Incoterms 값

    r += 1
    ws[f'A{r}'] = "Delivery mode:"
    ws[f'H{r}'] = "Email:"

    r += 1
    ws[f'A{r}'] = "Terms of payment:"
    ws[f'H{r}'] = "Tel:"

    r += 1  # Row 25: 청록색 푸터
    ws[f'A{r}'] = "Keeping the World flowing for Future Generations"
    ws[f'A{r}'].font = Font(size=12, color="FFFFFF", italic=True)
    ws[f'A{r}'].alignment = Alignment(vertical='center')
    ws[f'J{r}'] = "1 of 1"
    ws[f'J{r}'].font = Font(size=12, color="FFFFFF")
    ws[f'J{r}'].alignment = Alignment(horizontal='right', vertical='center')

    for col in range(1, TOTAL_COLUMNS + 1):
        ws.cell(row=r, column=col).fill = teal_fill

    # === 레이아웃 설정 ===
    for col, width in COLUMN_WIDTHS.as_dict().items():
        ws.column_dimensions[col].width = width

    ws.row_dimensions[1].height = 25
    ws.row_dimensions[12].height = 30
    ws.row_dimensions[r].height = 25

    ws.print_area = f'A1:J{r}'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.page_setup.orientation = 'portrait'
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5

    # === Description 시트 템플릿 ===
    ws_desc = wb.create_sheet("Description")

    # Row 1: Line No
    ws_desc['A1'] = "Line No"
    ws_desc['A1'].font = white_bold_font
    ws_desc['A1'].fill = green_fill
    ws_desc['A1'].border = thin_border
    ws_desc['A1'].alignment = Alignment(horizontal='center')

    ws_desc['B1'] = 1  # 첫 번째 아이템
    ws_desc['B1'].border = thin_border
    ws_desc['B1'].alignment = Alignment(horizontal='center')

    # Row 2: Qty
    ws_desc['A2'] = "Qty"
    ws_desc['A2'].font = white_bold_font
    ws_desc['A2'].fill = green_fill
    ws_desc['A2'].border = thin_border
    ws_desc['A2'].alignment = Alignment(horizontal='center')

    ws_desc['B2'] = 1
    ws_desc['B2'].border = thin_border
    ws_desc['B2'].alignment = Alignment(horizontal='center')

    # Row 3+: SPEC_FIELDS (초록 배경)
    row_idx = 3
    for field in SPEC_FIELDS:
        ws_desc.cell(row=row_idx, column=1, value=field)
        ws_desc.cell(row=row_idx, column=1).font = white_bold_font
        ws_desc.cell(row=row_idx, column=1).fill = green_fill
        ws_desc.cell(row=row_idx, column=1).border = thin_border
        ws_desc.cell(row=row_idx, column=2).border = thin_border
        row_idx += 1

    # OPTION_FIELDS (빨간 배경)
    for field in OPTION_FIELDS:
        ws_desc.cell(row=row_idx, column=1, value=field)
        ws_desc.cell(row=row_idx, column=1).font = white_bold_font
        ws_desc.cell(row=row_idx, column=1).fill = red_bright_fill
        ws_desc.cell(row=row_idx, column=1).border = thin_border
        ws_desc.cell(row=row_idx, column=2).border = thin_border
        row_idx += 1

    # 열 너비
    ws_desc.column_dimensions['A'].width = 25
    ws_desc.column_dimensions['B'].width = 15

    # 저장
    wb.save(PO_TEMPLATE_FILE)
    logger.info(f"PO 템플릿 생성 완료: {PO_TEMPLATE_FILE}")

    return PO_TEMPLATE_FILE
