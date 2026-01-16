"""
설정 및 상수 정의
================

경로, 색상, 필드 정의 등 프로젝트 전역 설정값을 관리합니다.
"""

from pathlib import Path
from dataclasses import dataclass
from typing import Final


# === 경로 설정 ===
BASE_DIR: Final[Path] = Path(__file__).parent.parent
DATA_DIR: Final[Path] = BASE_DIR.parent  # 상위 폴더 (NOAH ACTUATION)
NOAH_PO_LISTS_FILE: Final[Path] = DATA_DIR / "NOAH_PO_Lists.xlsx"
OUTPUT_DIR: Final[Path] = BASE_DIR / "generated_po"
HISTORY_FILE: Final[Path] = BASE_DIR / "po_history.xlsx"  # Legacy (하위 호환)
HISTORY_DIR: Final[Path] = BASE_DIR / "po_history"  # 새로운 폴더 방식

# === 템플릿 설정 ===
TEMPLATE_DIR: Final[Path] = BASE_DIR / "templates"
PO_TEMPLATE_FILE: Final[Path] = TEMPLATE_DIR / "purchase_order.xlsx"
TS_TEMPLATE_FILE: Final[Path] = TEMPLATE_DIR / "transaction_statement.xlsx"

# === 거래명세표 출력 설정 ===
TS_OUTPUT_DIR: Final[Path] = BASE_DIR / "generated_ts"


# === 시트 설정 ===
DOMESTIC_SHEET_INDEX: Final[int] = 0  # 국내
EXPORT_SHEET_INDEX: Final[int] = 1    # 해외


# === Excel 레이아웃 상수 (Purchase Order) ===
TOTAL_COLUMNS: Final[int] = 10
ITEM_START_ROW: Final[int] = 13
# MAX_ITEMS_PER_PO, ITEM_END_ROW 제거됨 - 아이템 수 제한 없이 동적 처리

# === 거래명세표 레이아웃 상수 ===
TS_TOTAL_COLUMNS: Final[int] = 9  # A-I (9열)
TS_ITEM_START_ROW: Final[int] = 13  # 아이템 시작 행
TS_HEADER_ROW: Final[int] = 12  # 헤더 행


# === 검증 설정 ===
MIN_LEAD_TIME_DAYS: Final[int] = 7


# === 메시지 마커 ===
MSG_ERROR: Final[str] = "[오류]"
MSG_WARNING: Final[str] = "[경고]"
MSG_NOTICE: Final[str] = "[주의]"


# === Excel 셀 참조 (이력 추출용 - 헤더 영역 고정 위치) ===
CELL_TITLE: Final[str] = "A1"
CELL_DATE: Final[str] = "A5"
CELL_CUSTOMER_NAME: Final[str] = "A10"


# === 파일명/출력 설정 ===
ORDER_LIST_DISPLAY_LIMIT: Final[int] = 20  # 주문 목록 출력 제한
HISTORY_CUSTOMER_DISPLAY_LENGTH: Final[int] = 15  # 이력 조회 시 고객명 표시 길이
HISTORY_DESC_DISPLAY_LENGTH: Final[int] = 20  # 이력 조회 시 설명 표시 길이
HISTORY_DATE_DISPLAY_LENGTH: Final[int] = 10  # 이력 조회 시 날짜 표시 길이


# === 필수 필드 (내부 키 사용) ===
REQUIRED_FIELDS: Final[tuple[str, ...]] = (
    'customer_name',
    'customer_po',
    'item_qty',
    'model',
    'ico_unit',
)


# === 액추에이터 사양 필드 (Description 시트) ===
SPEC_FIELDS: Final[tuple[str, ...]] = (
    'Power supply', 'Motor(kW)', 'BASE', 'ACT Flange', 'Operating time',
    'Handwheel', 'RPM', 'Turns', 'Bushing', 'MOV', 'Gearbox model',
    'Gearbox Flange', 'Gearbox ratio', 'Gearbox position', 'Operating mode',
    'Fail action', 'Enclosure', 'Cable entry', 'Paint', 'Cover tube(mm)',
    'WD code', 'Test report', 'Version', 'Note',
)


# === 옵션 필드 (Y 체크 시 가격 반영) ===
OPTION_FIELDS: Final[tuple[str, ...]] = (
    'Model', 'Bush', 'ALS', 'EXT', 'DC24V', 'Modbus, Profibus', 'LCU', 'PIU',
    'CPT+PIU', 'PCU+PIU', '-40', '-60', 'SCP', 'EXP', 'Bush-SQ', 'Bush-STAR',
    'INTEGRAL', 'IMS', 'BLDC', 'HART, Foundation Fieldbus', 'ATS',
)


@dataclass(frozen=True)
class Colors:
    """Excel 셀 배경색 (RGB hex)"""
    RED: str = "C00000"
    RED_BRIGHT: str = "FF0000"
    GRAY: str = "808080"
    TEAL: str = "008080"
    GREEN: str = "00B050"
    WHITE: str = "FFFFFF"


@dataclass(frozen=True)
class ColumnWidths:
    """Purchase Order 시트 열 너비"""
    A: int = 18
    B: int = 20
    C: int = 10
    D: int = 8
    E: int = 8
    F: int = 6
    G: int = 6
    H: int = 14
    I: int = 14
    J: int = 16

    def as_dict(self) -> dict[str, int]:
        return {
            'A': self.A, 'B': self.B, 'C': self.C, 'D': self.D, 'E': self.E,
            'F': self.F, 'G': self.G, 'H': self.H, 'I': self.I, 'J': self.J,
        }


# 인스턴스 생성
COLORS: Final[Colors] = Colors()
COLUMN_WIDTHS: Final[ColumnWidths] = ColumnWidths()


# === 컬럼 별칭 (Column Alias) ===
# NOAH_PO_Lists.xlsx 컬럼명이 변경되어도 자동으로 대응
# key: 내부 키, value: 가능한 컬럼명들 (첫 번째가 기본값)
COLUMN_ALIASES: Final[dict[str, tuple[str, ...]]] = {
    # 핵심 필드
    'order_no': ('RCK Order no.', 'RCK Order No', 'RCK Order no', 'Order No', '주문번호'),
    'customer_name': ('Customer name', 'Customer Name', 'customer name', '고객명', '고객사'),
    'customer_po': ('Customer PO', 'Customer PO No', 'customer po', '고객 PO', '고객PO'),
    'item_qty': ('Item qty', 'Item Qty', 'item qty', 'Qty', '수량'),
    'ico_unit': ('ICO Unit', 'ICO unit', 'ico unit', 'Unit Price', '단가'),
    'sales_unit_price': ('Sales Unit Price', 'Sales unit price', 'sales unit price', '판매단가'),
    'model': ('Model', 'MODEL', 'model', '모델'),
    'delivery_date': ('Requested delivery date', 'Delivery Date', 'delivery date', '납기일', '요청납기일'),
    'item_name': ('Item name', 'Item Name', 'item name', '품목명'),
    'remark': ('Remark', 'REMARK', 'remark', '비고'),
    'incoterms': ('Incoterms', 'INCOTERMS', 'incoterms', '인코텀즈'),
    'opportunity': ('Opportunity', 'OPPORTUNITY', 'opportunity', '프로젝트'),
    'sector': ('Sector', 'SECTOR', 'sector', '섹터'),
    'industry_code': ('Industry code', 'Industry Code', 'industry code', '산업코드'),
    'sheet_type': ('_시트구분',),  # 내부 컬럼
    # 사양 필드
    'power_supply': ('Power supply', 'Power Supply', 'power supply', '전원'),
    'als': ('ALS', 'als'),
}


# === 공급자 정보 (로토크 코리아) ===
@dataclass(frozen=True)
class SupplierInfo:
    """거래명세표 공급자 정보"""
    name: str = '로토크 콘트롤즈 코리아㈜'
    rep_name: str = '이민수'
    business_no: str = '220-81-21175'
    address: str = '경기도 성남시 분당구 장미로 42'
    address2: str = '야탑리더스빌딩 515'
    business_type: str = '도매업, 제조, 도매'
    business_item: str = '기타운수및기계장비, 밸브류, 무역'


SUPPLIER_INFO: Final[SupplierInfo] = SupplierInfo()


# === 거래명세표 열 너비 ===
@dataclass(frozen=True)
class TSColumnWidths:
    """거래명세표 시트 열 너비"""
    A: int = 8   # 월/일
    B: int = 22  # DESCRIPTION
    C: int = 10  # 비고
    D: int = 8   # 규격 SIZE
    E: int = 8   # 수량 QTY
    F: int = 14  # 단가 UNIT/PRICE
    G: int = 14  # 금액 AMOUNT
    H: int = 14  # 세액 TAXABLE AMOUNT
    I: int = 8   # 여유 열

    def as_dict(self) -> dict[str, int]:
        return {
            'A': self.A, 'B': self.B, 'C': self.C, 'D': self.D,
            'E': self.E, 'F': self.F, 'G': self.G, 'H': self.H, 'I': self.I,
        }


TS_COLUMN_WIDTHS: Final[TSColumnWidths] = TSColumnWidths()
