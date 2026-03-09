"""
NOAH Purchase Order Auto-Generator Package
==========================================

RCK Order No.를 입력하면 NOAH_PO_Lists.xlsx에서 해당 데이터를 읽어
자동으로 발주서(Purchase Order + Description)를 생성합니다.
"""

from po_generator.config import BASE_DIR, OUTPUT_DIR, HISTORY_FILE
from po_generator.utils import load_noah_po_lists, find_order_data, get_value
from po_generator.validators import validate_order_data
from po_generator.history import check_duplicate_order, save_to_history, sanitize_filename
from po_generator.excel_generator import create_po_workbook

__version__ = "2.5.0"
__all__ = [
    "load_noah_po_lists",
    "find_order_data",
    "get_value",
    "validate_order_data",
    "check_duplicate_order",
    "save_to_history",
    "sanitize_filename",
    "create_po_workbook",
]
