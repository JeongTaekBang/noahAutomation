"""
Excel Generator 테스트
======================

excel_generator.py 모듈의 테스트
"""

from datetime import datetime, timedelta
from pathlib import Path
import tempfile

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from po_generator.excel_generator import (
    create_purchase_order,
    create_description_sheet,
)
from po_generator.config import SPEC_FIELDS, OPTION_FIELDS


class TestCreatePurchaseOrder:
    """Purchase Order 시트 생성 테스트"""

    def test_creates_sheet_with_correct_title(self, valid_order_data):
        """시트 타이틀이 올바르게 설정되는지 테스트"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Purchase Order"

        create_purchase_order(ws, valid_order_data)

        # A1 셀에 주문번호가 포함되어야 함
        assert "Purchase Order" in str(ws['A1'].value)
        assert valid_order_data['RCK Order no.'] in str(ws['A1'].value)

    def test_creates_header_section(self, valid_order_data):
        """헤더 섹션이 올바르게 생성되는지 테스트"""
        wb = Workbook()
        ws = wb.active

        create_purchase_order(ws, valid_order_data)

        # Vendor Name
        assert "NOAH Actuation" in str(ws['A2'].value)
        # Customer name
        assert valid_order_data['Customer name'] == ws['A10'].value
        # Date 포맷 확인
        assert "Date:" in str(ws['A5'].value)

    def test_creates_item_rows(self, valid_order_data):
        """아이템 행이 올바르게 생성되는지 테스트"""
        wb = Workbook()
        ws = wb.active

        create_purchase_order(ws, valid_order_data)

        # Row 13에 첫 번째 아이템이 있어야 함
        assert ws['A13'].value == 1  # Item Number
        assert ws['F13'].value == valid_order_data['Item qty']  # Qty
        assert ws['G13'].value == "EA"  # Unit

    def test_creates_totals_section(self, valid_order_data):
        """합계 섹션이 올바르게 생성되는지 테스트"""
        wb = Workbook()
        ws = wb.active

        create_purchase_order(ws, valid_order_data)

        # Total net amount (Row 20)
        assert "Total net amount" in str(ws['I20'].value)
        # 수식 확인
        assert "=SUM(J13:J19)" in str(ws['J20'].value)

        # VAT (Row 21)
        assert "VAT" in str(ws['I21'].value)
        assert "=J20*0.1" in str(ws['J21'].value)

        # Order Total (Row 22)
        assert "Order Total" in str(ws['I22'].value)

    def test_creates_footer_section(self, valid_order_data):
        """푸터 섹션이 올바르게 생성되는지 테스트"""
        wb = Workbook()
        ws = wb.active

        create_purchase_order(ws, valid_order_data)

        # Footer text (Row 31)
        assert "Keeping the World flowing" in str(ws['A31'].value)
        # Currency
        assert ws['B27'].value == 'KRW'
        # Incoterms
        assert ws['B28'].value == valid_order_data['Incoterms']

    def test_handles_multiple_items(self, multiple_items_df):
        """다중 아이템 처리 테스트"""
        wb = Workbook()
        ws = wb.active
        order_data = multiple_items_df.iloc[0]

        create_purchase_order(ws, order_data, multiple_items_df)

        # 두 개의 아이템이 Row 13, 14에 있어야 함
        assert ws['A13'].value == 1
        assert ws['A14'].value == 2
        # 세 번째 행은 비어있어야 함 (아이템 없음)
        assert ws['A15'].value is None

    def test_saves_valid_excel_file(self, valid_order_data):
        """유효한 Excel 파일로 저장되는지 테스트"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Purchase Order"

        create_purchase_order(ws, valid_order_data)

        # 임시 파일에 저장
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            temp_path = Path(f.name)

        try:
            wb.save(temp_path)
            # 저장된 파일 다시 로드
            loaded_wb = load_workbook(temp_path)
            loaded_ws = loaded_wb.active

            # 데이터가 유지되는지 확인
            assert "Purchase Order" in str(loaded_ws['A1'].value)
            loaded_wb.close()
        finally:
            temp_path.unlink()

    def test_applies_print_settings(self, valid_order_data):
        """인쇄 설정이 적용되는지 테스트"""
        wb = Workbook()
        ws = wb.active

        create_purchase_order(ws, valid_order_data)

        # 인쇄 영역 (openpyxl은 시트 이름과 절대 참조 형식으로 반환: 'Sheet'!$A$1:$J$31)
        print_area_str = str(ws.print_area)
        assert '$A$1' in print_area_str and '$J$31' in print_area_str
        # 페이지 설정
        assert ws.page_setup.fitToPage is True


class TestCreateDescriptionSheet:
    """Description 시트 생성 테스트"""

    def test_creates_line_no_header(self, valid_order_data):
        """Line No 헤더가 올바르게 생성되는지 테스트"""
        wb = Workbook()
        ws = wb.create_sheet("Description")

        create_description_sheet(ws, valid_order_data)

        assert ws['A1'].value == "Line No"
        assert ws.cell(row=1, column=2).value == 1  # 첫 번째 아이템

    def test_creates_spec_fields(self, valid_order_data):
        """사양 필드가 올바르게 생성되는지 테스트"""
        wb = Workbook()
        ws = wb.create_sheet("Description")

        create_description_sheet(ws, valid_order_data)

        # SPEC_FIELDS가 Row 2부터 시작
        for idx, field in enumerate(SPEC_FIELDS):
            row = idx + 2
            assert ws.cell(row=row, column=1).value == field

    def test_creates_option_fields(self, valid_order_data):
        """옵션 필드가 올바르게 생성되는지 테스트"""
        wb = Workbook()
        ws = wb.create_sheet("Description")

        create_description_sheet(ws, valid_order_data)

        # OPTION_FIELDS는 SPEC_FIELDS 다음에 시작
        start_row = len(SPEC_FIELDS) + 2
        for idx, field in enumerate(OPTION_FIELDS):
            row = start_row + idx
            assert ws.cell(row=row, column=1).value == field

    def test_populates_data_values(self, valid_order_data):
        """데이터 값이 올바르게 채워지는지 테스트"""
        wb = Workbook()
        ws = wb.create_sheet("Description")

        create_description_sheet(ws, valid_order_data)

        # Power supply 값 확인 (SPEC_FIELDS의 첫 번째)
        power_supply_row = 2  # Row 2
        assert ws.cell(row=power_supply_row, column=2).value == valid_order_data['Power supply']

        # ALS 옵션 값 확인
        als_row = len(SPEC_FIELDS) + 2 + OPTION_FIELDS.index('ALS')
        assert ws.cell(row=als_row, column=2).value == valid_order_data['ALS']

    def test_handles_multiple_items(self, multiple_items_df):
        """다중 아이템 처리 테스트"""
        wb = Workbook()
        ws = wb.create_sheet("Description")
        order_data = multiple_items_df.iloc[0]

        create_description_sheet(ws, order_data, multiple_items_df)

        # Line No: 1, 2가 있어야 함
        assert ws.cell(row=1, column=2).value == 1
        assert ws.cell(row=1, column=3).value == 2

    def test_column_widths_applied(self, valid_order_data):
        """열 너비가 설정되는지 테스트"""
        wb = Workbook()
        ws = wb.create_sheet("Description")

        create_description_sheet(ws, valid_order_data)

        # A열 너비
        assert ws.column_dimensions['A'].width == 25


class TestExcelGeneratorIntegration:
    """Excel Generator 통합 테스트"""

    def test_full_workbook_generation(self, valid_order_data):
        """전체 워크북 생성 통합 테스트"""
        wb = Workbook()

        # Purchase Order 시트
        ws_po = wb.active
        ws_po.title = "Purchase Order"
        create_purchase_order(ws_po, valid_order_data)

        # Description 시트
        ws_desc = wb.create_sheet("Description")
        create_description_sheet(ws_desc, valid_order_data)

        # 두 시트가 존재하는지 확인
        assert "Purchase Order" in wb.sheetnames
        assert "Description" in wb.sheetnames

        # 저장 및 로드 테스트
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            temp_path = Path(f.name)

        try:
            wb.save(temp_path)
            loaded_wb = load_workbook(temp_path, data_only=True)

            # 시트 확인
            assert len(loaded_wb.sheetnames) == 2
            loaded_wb.close()
        finally:
            temp_path.unlink()

    def test_multiple_items_workbook(self, multiple_items_df):
        """다중 아이템 워크북 생성 테스트"""
        wb = Workbook()
        order_data = multiple_items_df.iloc[0]

        ws_po = wb.active
        ws_po.title = "Purchase Order"
        create_purchase_order(ws_po, order_data, multiple_items_df)

        ws_desc = wb.create_sheet("Description")
        create_description_sheet(ws_desc, order_data, multiple_items_df)

        # 저장 테스트
        with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
            temp_path = Path(f.name)

        try:
            wb.save(temp_path)
            assert temp_path.exists()
            assert temp_path.stat().st_size > 0
        finally:
            temp_path.unlink()
