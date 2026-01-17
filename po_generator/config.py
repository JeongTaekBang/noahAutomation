"""
설정 및 상수 정의
================

경로, 색상, 필드 정의 등 프로젝트 전역 설정값을 관리합니다.
사용자 설정은 user_settings.py에서 관리합니다.
"""

from pathlib import Path
from dataclasses import dataclass
from typing import Final


# === 경로 설정 ===
BASE_DIR: Final[Path] = Path(__file__).parent.parent

# 사용자 설정에서 값 가져오기 (없으면 기본값 사용)
try:
    from user_settings import DATA_FOLDER
    DATA_DIR: Final[Path] = Path(DATA_FOLDER)
except ImportError:
    DATA_DIR: Final[Path] = BASE_DIR.parent

# 출력 폴더 기본 경로 (user_settings에서 설정 가능)
try:
    from user_settings import OUTPUT_BASE_DIR
    _OUTPUT_BASE: Path | None = Path(OUTPUT_BASE_DIR) if OUTPUT_BASE_DIR else None
except ImportError:
    _OUTPUT_BASE = None
# 새 데이터베이스 파일 (SO/PO/DN 분리 구조)
NOAH_SO_PO_DN_FILE: Final[Path] = DATA_DIR / "NOAH_SO_PO_DN.xlsx"
# 기존 파일 (하위 호환 - deprecated)
NOAH_PO_LISTS_FILE: Final[Path] = DATA_DIR / "NOAH_PO_Lists.xlsx"

# 출력 폴더 (OUTPUT_BASE_DIR 설정 시 해당 경로 사용, 없으면 프로젝트 폴더)
_OUT_BASE: Path = _OUTPUT_BASE if _OUTPUT_BASE else BASE_DIR
OUTPUT_DIR: Final[Path] = _OUT_BASE / "generated_po"
HISTORY_FILE: Final[Path] = _OUT_BASE / "po_history.xlsx"  # Legacy (하위 호환)
HISTORY_DIR: Final[Path] = _OUT_BASE / "po_history"  # 새로운 폴더 방식

# === 템플릿 설정 ===
TEMPLATE_DIR: Final[Path] = BASE_DIR / "templates"
PO_TEMPLATE_FILE: Final[Path] = TEMPLATE_DIR / "purchase_order.xlsx"
TS_TEMPLATE_FILE: Final[Path] = TEMPLATE_DIR / "ts_template_local.xlsx"

# === 거래명세표 출력 설정 ===
TS_OUTPUT_DIR: Final[Path] = _OUT_BASE / "generated_ts"

# === Proforma Invoice 설정 ===
PI_TEMPLATE_FILE: Final[Path] = TEMPLATE_DIR / "proforma_invoice.xlsx"
PI_OUTPUT_DIR: Final[Path] = _OUT_BASE / "generated_pi"

# === Commercial Invoice 설정 ===
CI_TEMPLATE_FILE: Final[Path] = TEMPLATE_DIR / "commercial_invoice.xlsx"
CI_OUTPUT_DIR: Final[Path] = _OUT_BASE / "generated_ci"


# === 시트 설정 (NOAH_SO_PO_DN.xlsx) ===
# 국내 시트
SO_DOMESTIC_SHEET: Final[str] = 'SO_국내'
PO_DOMESTIC_SHEET: Final[str] = 'PO_국내'
DN_DOMESTIC_SHEET: Final[str] = 'DN_국내'
PMT_DOMESTIC_SHEET: Final[str] = 'PMT_국내'
# 해외 시트
SO_EXPORT_SHEET: Final[str] = 'SO_해외'
PO_EXPORT_SHEET: Final[str] = 'PO_해외'
DN_EXPORT_SHEET: Final[str] = 'DN_해외'

# 기존 설정 (하위 호환 - deprecated)
DOMESTIC_SHEET_INDEX: Final[int] = 0  # 국내
EXPORT_SHEET_INDEX: Final[int] = 1    # 해외


# === Excel 레이아웃 상수 (Purchase Order) ===
TOTAL_COLUMNS: Final[int] = 10
# ITEM_START_ROW 제거됨 - find_item_start_row()로 동적 탐지
ITEM_START_ROW_FALLBACK: Final[int] = 13  # 동적 탐지 실패 시 기본값

# === 거래명세표 레이아웃 상수 ===
TS_TOTAL_COLUMNS: Final[int] = 9  # A-I (9열)
TS_HEADER_ROW: Final[int] = 12  # 헤더 행
# TS_ITEM_START_ROW 제거됨 - find_item_start_row()로 동적 탐지

# === 비즈니스 규칙 상수 ===
try:
    from user_settings import VAT_RATE_DOMESTIC as _VAT_RATE
    VAT_RATE_DOMESTIC: Final[float] = _VAT_RATE
except ImportError:
    VAT_RATE_DOMESTIC: Final[float] = 0.1  # 국내 부가세율 10%

# === 안전 장치 상수 ===
MAX_HEADER_SEARCH_ROWS: Final[int] = 30  # 동적 헤더 탐지 최대 범위
HISTORY_MAX_SEARCH_ROWS: Final[int] = 100  # 이력 아이템 검색 제한


# === 검증 설정 ===
try:
    from user_settings import MIN_LEAD_TIME_DAYS as _MIN_LEAD
    MIN_LEAD_TIME_DAYS: Final[int] = _MIN_LEAD
except ImportError:
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

try:
    from user_settings import HISTORY_CUSTOMER_DISPLAY_LENGTH as _HIST_CUST_LEN
    HISTORY_CUSTOMER_DISPLAY_LENGTH: Final[int] = _HIST_CUST_LEN
except ImportError:
    HISTORY_CUSTOMER_DISPLAY_LENGTH: Final[int] = 15  # 이력 조회 시 고객명 표시 길이

try:
    from user_settings import HISTORY_DESC_DISPLAY_LENGTH as _HIST_DESC_LEN
    HISTORY_DESC_DISPLAY_LENGTH: Final[int] = _HIST_DESC_LEN
except ImportError:
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
# NOAH_SO_PO_DN.xlsx 컬럼명이 변경되어도 자동으로 대응
# key: 내부 키, value: 가능한 컬럼명들 (첫 번째가 기본값)
COLUMN_ALIASES: Final[dict[str, tuple[str, ...]]] = {
    # 핵심 필드 (새 구조: PO_ID가 발주번호, SO_ID가 연결키)
    'order_no': ('PO_ID', 'RCK Order no.', 'RCK Order No', 'RCK Order no', 'Order No', '주문번호'),
    'so_id': ('SO_ID', 'SO ID', 'so_id'),
    'noah_oc_no': ('NOAH O.C No.', 'NOAH O.C No', 'NOAH OC No', '공장발주번호'),
    'customer_name': ('Customer name', 'Customer Name', 'customer name', '고객명', '고객사'),
    'customer_po': ('Customer PO', 'Customer PO No', 'customer po', '고객 PO', '고객PO'),
    'item_qty': ('Item qty', 'Item Qty', 'item qty', 'Qty', '수량'),
    'ico_unit': ('ICO Unit', 'ICO unit', 'ico unit', 'Unit Price', '단가'),
    'total_ico': ('Total ICO', 'Total ico', 'total_ico', '총ICO'),
    'sales_unit_price': ('Sales Unit Price', 'Sales unit price', 'sales unit price', '판매단가'),
    'model': ('Model', 'MODEL', 'model', '모델'),
    'delivery_date': ('예상 EXW date', '예상 납품 날짜', 'Requested delivery date', 'Delivery Date', 'delivery date', '납기일', '요청납기일'),
    'item_name': ('Item name', 'Item Name', 'item name', 'Item', '품목명'),
    'remark': ('Note', 'Remark', 'REMARK', 'remark', '비고'),
    'incoterms': ('Incoterms', 'INCOTERMS', 'incoterms', '인코텀즈'),
    'opportunity': ('Opportunity', 'OPPORTUNITY', 'opportunity', '프로젝트'),
    'sector': ('Sector', 'SECTOR', 'sector', '섹터'),
    'industry_code': ('Industry code', 'Industry Code', 'industry code', '산업코드'),
    'sheet_type': ('_시트구분',),  # 내부 컬럼
    'status': ('Status', 'STATUS', 'status', '상태'),
    # 사양 필드
    'power_supply': ('Power supply', 'Power Supply', 'power supply', '전원'),
    'als': ('ALS', 'als'),
    # DN (납품) 필드
    'dn_id': ('DN_ID', 'DN ID', 'dn_id', '납품번호'),
    'unit_price': ('Unit Price', 'unit price', '단가'),
    'total_sales': ('Total Sales', 'total sales', '판매금액'),
    'tax_invoice_no': ('세금계산서', '세금계산서번호', 'Tax Invoice No'),
    # PMT (입금) 필드
    'advance_id': ('선수금_ID', 'ADV_ID', '선수금ID'),
    'expected_amount': ('입금 예정 금액', '예정금액'),
    'paid_amount': ('입금액', 'Paid Amount', '입금금액'),
    'paid_date': ('입금일', 'Paid Date'),
    'tax_invoice_date': ('세금계산서 발행일', '발행일'),
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


# user_settings에서 공급자 정보 가져오기
try:
    from user_settings import SUPPLIER_INFO as _USER_SUPPLIER
    SUPPLIER_INFO: Final[SupplierInfo] = SupplierInfo(
        name=_USER_SUPPLIER.get('name', '로토크 콘트롤즈 코리아㈜'),
        rep_name=_USER_SUPPLIER.get('rep_name', '이민수'),
        business_no=_USER_SUPPLIER.get('business_no', '220-81-21175'),
        address=_USER_SUPPLIER.get('address', '경기도 성남시 분당구 장미로 42'),
        address2=_USER_SUPPLIER.get('address2', '야탑리더스빌딩 515'),
        business_type=_USER_SUPPLIER.get('business_type', '도매업, 제조, 도매'),
        business_item=_USER_SUPPLIER.get('business_item', '기타운수및기계장비, 밸브류, 무역'),
    )
except ImportError:
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
