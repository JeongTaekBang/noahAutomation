"""
설정 및 상수 정의
================

경로, 색상, 필드 정의 등 프로젝트 전역 설정값을 관리합니다.
사용자 설정은 user_settings.py에서 관리합니다.
"""

import configparser
from pathlib import Path
from dataclasses import dataclass
from typing import Any, Final


# === 배포판 설정 파일 (noah_config.ini) ===
# 개발 PC는 user_settings.py를 쓰고, 배포판은 앱 폴더의 noah_config.ini를 쓴다.
# ini는 GUI(noah_gui.py) 첫 실행 마법사가 만든다.
_INI_FILE: Final[Path] = Path(__file__).parent.parent / "noah_config.ini"

# ini 키 → user_settings.py 변수명 (ini에서 지원하는 항목만)
_INI_KEYS: Final[dict[str, str]] = {
    'DATA_FOLDER': 'data_folder',
    'OUTPUT_BASE_DIR': 'output_base_dir',
}


def _read_ini() -> dict[str, str]:
    """noah_config.ini의 [paths] 섹션 로드 (없으면 빈 dict)"""
    if not _INI_FILE.exists():
        return {}
    # interpolation=None — 경로에 '%'가 있어도 깨지지 않게
    # utf-8-sig — 메모장으로 저장하면 BOM이 붙는다. BOM이 있으면 첫 섹션 헤더가
    #             '﻿[paths]'가 되어 섹션을 못 찾고, 설정이 조용히 무시된다.
    parser = configparser.ConfigParser(interpolation=None)
    try:
        parser.read(_INI_FILE, encoding='utf-8-sig')
    except (configparser.Error, OSError):
        return {}
    if not parser.has_section('paths'):
        return {}
    return {k: v.strip() for k, v in parser.items('paths') if v.strip()}


_INI_VALUES: Final[dict[str, str]] = _read_ini()

_MISSING: Final[object] = object()


# === 사용자 설정 로딩 헬퍼 ===
def _load_user_setting(name: str, default: Any) -> Any:
    """설정값 로드

    우선순위: user_settings.py → noah_config.ini → 기본값

    user_settings.py에 이름이 있으면 값이 None이어도 그것이 최종값이다
    (`OUTPUT_BASE_DIR = None`은 "프로젝트 폴더 사용"이라는 의미 있는 설정).
    따라서 개발 PC에 ini가 생겨도 기존 동작은 그대로다.
    배포판에는 user_settings.py가 없어 ini가 쓰인다.
    """
    try:
        import user_settings
        value = getattr(user_settings, name, _MISSING)
        if value is not _MISSING:
            return value
    except ImportError:
        pass

    ini_key = _INI_KEYS.get(name)
    if ini_key and ini_key in _INI_VALUES:
        return _INI_VALUES[ini_key]

    return default


# === 경로 설정 ===
BASE_DIR: Final[Path] = Path(__file__).parent.parent

# 사용자 설정에서 값 가져오기 (없으면 기본값 사용)
_data_folder = _load_user_setting('DATA_FOLDER', None)
DATA_DIR: Final[Path] = Path(_data_folder) if _data_folder else BASE_DIR.parent

# 출력 폴더 기본 경로 (user_settings에서 설정 가능)
_output_base_dir = _load_user_setting('OUTPUT_BASE_DIR', None)
_OUTPUT_BASE: Path | None = Path(_output_base_dir) if _output_base_dir else None
# 새 데이터베이스 파일 (SO/PO/DN 분리 구조)
NOAH_SO_PO_DN_FILE: Final[Path] = DATA_DIR / "NOAH_SO_PO_DN.xlsx"
# SQLite 백업 DB
DB_FILE: Final[Path] = DATA_DIR / "noah_data.db"
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

# === Final Invoice 설정 (대금 청구용) ===
FI_TEMPLATE_FILE: Final[Path] = TEMPLATE_DIR / "final_invoice.xlsx"
FI_OUTPUT_DIR: Final[Path] = _OUT_BASE / "generated_fi"

# === Packing List 설정 ===
PL_TEMPLATE_FILE: Final[Path] = TEMPLATE_DIR / "packing_list.xlsx"
PL_OUTPUT_DIR: Final[Path] = _OUT_BASE / "generated_pl"

# === Order Confirmation 설정 ===
OC_TEMPLATE_FILE: Final[Path] = TEMPLATE_DIR / "order_confirmation.xlsx"
OC_OUTPUT_DIR: Final[Path] = _OUT_BASE / "generated_oc"

# === 거래처 납기현황 회신 설정 (템플릿 없음 — 조회 결과를 새 통합문서로 출력) ===
DS_OUTPUT_DIR: Final[Path] = _OUT_BASE / "generated_ds"


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
CUSTOMER_DOMESTIC_SHEET: Final[str] = 'Customer_국내'
CUSTOMER_EXPORT_SHEET: Final[str] = 'Customer_해외'
WEIGHT_SHEET: Final[str] = 'Weight'
# 월별 환율 (가로형: 헤더 `FX | 2026-01 | 2026-02 | ...`, 행 = 통화)
# 해외 매출을 선적월 환율로 재환산할 때 사용 — Order Book / AX_매출대사 공통 기준
FX_SHEET: Final[str] = 'FX'

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
VAT_RATE_DOMESTIC: Final[float] = _load_user_setting('VAT_RATE_DOMESTIC', 0.1)

# === 안전 장치 상수 ===
# MAX_HEADER_SEARCH_ROWS 제거됨 - 미사용
# HISTORY_MAX_SEARCH_ROWS는 history.py에서 함수 기본값으로 이동


# === 검증 설정 ===
MIN_LEAD_TIME_DAYS: Final[int] = _load_user_setting('MIN_LEAD_TIME_DAYS', 7)


# === 메시지 마커 ===
MSG_ERROR: Final[str] = "[오류]"
MSG_WARNING: Final[str] = "[경고]"
MSG_NOTICE: Final[str] = "[주의]"


# === Excel 셀 참조 - history.py로 이동됨 ===
# CELL_TITLE, CELL_DATE, CELL_CUSTOMER_NAME은 history.py에서만 사용


# === 파일명/출력 설정 ===
ORDER_LIST_DISPLAY_LIMIT: Final[int] = 20  # 주문 목록 출력 제한
HISTORY_CUSTOMER_DISPLAY_LENGTH: Final[int] = _load_user_setting('HISTORY_CUSTOMER_DISPLAY_LENGTH', 15)
HISTORY_DESC_DISPLAY_LENGTH: Final[int] = _load_user_setting('HISTORY_DESC_DISPLAY_LENGTH', 20)
HISTORY_DATE_DISPLAY_LENGTH: Final[int] = 10  # 이력 조회 시 날짜 표시 길이


# === 필수 필드 (내부 키 사용) ===
REQUIRED_FIELDS: Final[tuple[str, ...]] = (
    'customer_name',
    'customer_po',
    'item_qty',
    'model',
    'ico_unit',
)


# === 메타 컬럼 (Description 시트에서 제외) ===
# PO 시트에서 사양/옵션이 아닌 메타 정보 컬럼들
META_COLUMNS: Final[frozenset[str]] = frozenset({
    'PO_ID', 'SO_ID', 'NOAH O.C No.', 'Customer name', 'Customer PO',
    'Item name', 'Item qty', 'ICO Unit', 'Total ICO',
    '예상 납품 날짜', '예상 EXW date', 'Status',
    # 내부 컬럼
    '_시트구분', '_문서유형',
})

# === 사양 필드 시작 마커 ===
# 이 컬럼부터 사양 필드 시작 (동적 추출 시 사용)
SPEC_START_COLUMN: Final[str] = 'Power supply'

# === 옵션 필드 시작 마커 ===
# 이 컬럼부터 옵션 필드 시작 (Status 다음 컬럼)
OPTION_START_COLUMN: Final[str] = 'Model'

# === 액추에이터 사양 필드 (Description 시트) - Fallback용 ===
# 동적 추출 실패 시 사용되는 기본값
SPEC_FIELDS: Final[tuple[str, ...]] = (
    'Power supply', 'Motor(kW)', 'BASE', 'ACT Flange', 'Operating time',
    'Handwheel', 'RPM', 'Turns', 'Bushing', 'MOV', 'Gearbox model',
    'Gearbox Flange', 'Gearbox ratio', 'Gearbox position', 'Operating mode',
    'Fail action', 'Enclosure', 'Cable entry', 'Paint', 'Cover tube(mm)',
    'WD code', 'Test report', 'Version', 'Note',
)


# === 옵션 필드 (Y 체크 시 가격 반영) - Fallback용 ===
# 동적 추출 실패 시 사용되는 기본값
OPTION_FIELDS: Final[tuple[str, ...]] = (
    'Model', 'Bush', 'ALS', 'EXT', 'DC24V', 'Modbus, Profibus', 'LCU', 'PIU',
    'CPT+PIU', 'PCU+PIU', '-40', '-60', 'SCP', 'EXP', 'Bush-SQ', 'Bush-STAR',
    'INTEGRAL', 'IMS', 'BLDC', 'HART, Foundation Fieldbus', 'ATS',
    'MOV조립', 'VALVE 가격',
)


# === Weight 매핑 설정 (Packing List Net Weight) ===
# PO_해외 옵션열(Y 체크) → Weight 시트 MODEL 코드 접미사
# 예: Model 'NA006' + IMS 옵션 → '006' + 'IM' → Weight 코드 '006IM'
# 여기 없는 옵션(Bush/ALS/EXT/DC24V/PIU 등)은 무게에 영향 없음 → 매핑 미사용
WEIGHT_OPTION_SUFFIX: Final[dict[str, str]] = {
    'INTEGRAL': 'IN',
    'IMS': 'IM',
    'LCU': 'L',
    'PCU+PIU': 'P',
    'SCP': 'S',
    'EXP': 'X',
}

# 한 라인에 무게 영향 옵션이 복수로 Y일 때 적용 우선순위 (앞이 우선)
# Weight 시트엔 단일 옵션 행만 존재하므로 우선순위 최상위 1개로 코드를 만든다.
# (예외: LCU + PCU+PIU 동시면 결합코드 '…LP'를 우선 시도)
WEIGHT_OPTION_PRIORITY: Final[tuple[str, ...]] = (
    'INTEGRAL', 'IMS', 'LCU', 'PCU+PIU', 'SCP', 'EXP',
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
    'customer_name': ('Customer name', 'Customer Name', 'customer name', '고객명', '고객사', '거래처명'),
    'customer_po': ('Customer PO', 'Customer PO No', 'customer po', '고객 PO', '고객PO'),
    'item_qty': ('Item qty', 'Item Qty', 'item qty', 'Qty', '수량'),
    'ico_unit': ('ICO Unit', 'ICO unit', 'ico unit', 'Unit Price', '단가'),
    'total_ico': ('Total ICO', 'Total ico', 'total_ico', '총ICO'),
    'sales_unit_price': ('Sales Unit Price', 'Sales unit price', 'sales unit price', '판매단가'),
    'model': ('Model', 'MODEL', 'model', '모델', 'Model number'),
    'delivery_date': ('예상 EXW date', '예상 납품 날짜', 'Requested delivery date', 'Delivery Date', 'delivery date', '납기일', '요청납기일'),
    'delivery_address': ('납품 주소', '납품주소', 'Delivery Address', 'Delivery address', 'delivery address', '배송주소', '배송 주소'),
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
    'dispatch_date': ('출고일', 'Dispatch Date', 'dispatch_date', '출하일', '선적일'),
    'unit_price': ('Unit Price', 'unit price', '단가'),
    'total_sales': ('Total Sales', 'total sales', '판매금액'),
    'tax_invoice_no': ('세금계산서', '세금계산서번호', 'Tax Invoice No'),
    # PMT (입금) 필드
    'advance_id': ('선수금_ID', 'ADV_ID', '선수금ID'),
    'expected_amount': ('입금 예정 금액', '예정금액'),
    'paid_amount': ('입금액', 'Paid Amount', '입금금액'),
    'paid_date': ('입금일', 'Paid Date'),
    'tax_invoice_date': ('세금계산서 발행일', '발행일'),
    # PI/TS (해외) 필드
    'customer_address': ('Customer address', 'Customer Address', 'customer address', '고객주소'),
    'customer_country': ('Customer country', 'Customer Country', 'customer country', '고객국가'),
    'customer_tel': ('Customer TEL', 'Customer Tel', 'customer tel', '고객전화'),
    'customer_fax': ('Customer FAX', 'Customer Fax', 'customer fax', '고객팩스'),
    'currency': ('Currency', 'CURRENCY', 'currency', '통화'),
    'po_receipt_date': ('PO receipt date', 'PO Receipt Date', 'po receipt date', 'PO수령일'),
    'lc_no': ('L/C No', 'LC No', 'lc no', 'LC번호'),
    'lc_date': ('L/C date', 'LC date', 'lc date', 'LC발행일'),
    # Final Invoice (대금 청구) 필드
    # `SO_해외`의 'Business registration number'에는 사업자번호가 아니라 고객코드
    # (C-0054)가 들어 있다 — 컬럼 이름만 국내 시트와 같다. `Customer_해외`에서는
    # 'C-code by 해외'가 그 짝이다.
    # 주의: `Customer_해외`의 '고객코드' 컬럼은 **다른 값**(AX 번호, 2352 같은 숫자)이라
    # 여기 별칭에 넣으면 안 된다 — 앞 별칭이 없는 시트에서 조용히 엉뚱한 컬럼으로 풀려
    # 수신자 조인이 전 건 미매칭이 된다.
    'customer_code': ('Business registration number', 'C-code by 해외'),
    'bill_to_1': ('Bill to 1', 'bill to 1'),
    'bill_to_2': ('Bill to 2', 'bill to 2'),
    'bill_to_3': ('Bill to 3', 'bill to 3'),
    'payment_terms': ('Payment terms', 'Payment Terms', 'payment terms', '결제조건'),
    'rck_po': ('RCK PO', 'RCK PO No', 'rck_po'),
    # Order Confirmation 필드
    'exw_noah': ('EXW NOAH', 'EXW Noah', 'exw_noah', 'EXW date'),
    'shipping_method': ('Shipping method', 'Shipping Method', 'shipping_method', '배송방법'),
    # Customer_국내 (거래명세표 메일 발송) 필드
    'biz_no': ('Business registration number', '사업자번호', '사업자등록번호', 'biz_no', 'BRN'),
    'customer_name_en': ('Customer Name ENG', 'Customer name ENG', '거래처명(영문)', '영문 거래처명'),
    # 주의: '참조 이메일'도 '이메일'을 포함하므로 부분일치를 쓰면 안 된다.
    #       두 목록 모두 완전일치 전용이며, 참조용 이름은 customer_email에 넣지 말 것.
    'customer_email': (
        '수신자 이메일', '수신자이메일', '수신 이메일', '수신메일',
        '이메일', '이메일 주소', '메일', '담당자 이메일',
        'Email', 'E-mail', 'EMAIL', 'email',
    ),
    'customer_email_cc': (
        '참조 이메일', '참조이메일', '참조메일', '참조 메일', 'CC 이메일',
        'CC', 'Cc', 'cc', 'Email CC', '참조',
    ),
    # Packing List 필드
    'model_code': ('Model code', 'AX Project number', 'model_code'),
    'weight_per_unit': ('Weight per unit', 'Weight/Unit', 'KG/PC', 'weight_per_unit', '단위중량'),
    'gross_weight': ('Gross Weight', 'Weight', 'gross_weight', 'Total Weight', '총중량'),
    'cbm': ('CBM', 'cbm', 'Cubic Meter', '체적'),
}


# === 공급자 정보 (로토크 코리아) ===
@dataclass(frozen=True)
class SupplierInfo:
    """거래명세표 공급자 정보"""
    name: str = '로토크 콘트롤즈 코리아㈜'
    # 영문 상호 — 해외 고객에게 나가는 메일(OC) 서명용.
    # `name`(한글)을 영문 본문에 쓸 수 없어 별도로 둔다. 값은 해외 문서 템플릿
    # 머리글(order_confirmation.xlsx A3)과 같아야 한다 — 메일과 첨부의 발신자가 갈리면 안 된다.
    name_en: str = 'Rotork Controls Korea Co., Ltd.'
    rep_name: str = '이민수'
    business_no: str = '220-81-21175'
    address: str = '경기도 성남시 분당구 장미로 42'
    address2: str = '야탑리더스빌딩 515'
    business_type: str = '도매업, 제조, 도매'
    business_item: str = '기타운수및기계장비, 밸브류, 무역'


# user_settings에서 공급자 정보 가져오기
_user_supplier = _load_user_setting('SUPPLIER_INFO', None)
if _user_supplier:
    SUPPLIER_INFO: Final[SupplierInfo] = SupplierInfo(
        name=_user_supplier.get('name', '로토크 콘트롤즈 코리아㈜'),
        name_en=_user_supplier.get('name_en', 'Rotork Controls Korea Co., Ltd.'),
        rep_name=_user_supplier.get('rep_name', '이민수'),
        business_no=_user_supplier.get('business_no', '220-81-21175'),
        address=_user_supplier.get('address', '경기도 성남시 분당구 장미로 42'),
        address2=_user_supplier.get('address2', '야탑리더스빌딩 515'),
        business_type=_user_supplier.get('business_type', '도매업, 제조, 도매'),
        business_item=_user_supplier.get('business_item', '기타운수및기계장비, 밸브류, 무역'),
    )
else:
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


# === 거래명세표 메일 발송 설정 ===
# 수신자(To)는 Customer_국내에서 사업자번호로 조회하고, 아래 CC는 항상 붙는 고정 참조자.
_ts_mail_cc = _load_user_setting('TS_MAIL_CC', ())
TS_MAIL_CC: Final[tuple[str, ...]] = tuple(_ts_mail_cc) if _ts_mail_cc else ()

# 첨부 형식: 'pdf' | 'xlsx' | 'both'
TS_MAIL_ATTACH_FORMAT: Final[str] = _load_user_setting('TS_MAIL_ATTACH_FORMAT', 'pdf')

# 메일 작성 방식: 'auto' | 'outlook' | 'eml'
#   outlook — Outlook COM. classic Outlook 전용 (새 Outlook은 COM 미지원)
#   eml     — .eml 초안 파일을 만들어 기본 메일 앱으로 열기. 새 Outlook에서도 동작
#   auto    — COM을 한 번 시도해보고 안 되면 eml로 자동 전환 (기본)
TS_MAIL_BACKEND: Final[str] = _load_user_setting('TS_MAIL_BACKEND', 'auto')

# 제목/본문 템플릿
# 치환자: {customer} 한글 거래처명, {customer_en} 영문 거래처명(없으면 한글명),
#         {customer_po} 거래처 발주번호(여러 건이면 쉼표 구분, 없으면 N/A),
#         {doc_id} DN 번호, {date} 출고일, {supplier} 공급자명(한글)
TS_MAIL_SUBJECT: Final[str] = _load_user_setting(
    'TS_MAIL_SUBJECT',
    '[거래명세표] {customer} - {date}',
)
TS_MAIL_BODY: Final[str] = _load_user_setting(
    'TS_MAIL_BODY',
    '{customer} 귀중\n\n'
    '{date}자 출고분 거래명세표를 첨부와 같이 송부합니다.\n\n'
    '발주번호: {customer_po}\n\n'
    '본 메일은 자동 발송된 메일입니다.\n',
)


# === 납기현황 회신 메일 설정 (delivery_status.py) ===
# 수신자(To)는 거래명세표와 같은 경로로 찾는다 — 조회 기준인 사업자번호로 Customer_국내 조인.
DS_MAIL_CC: Final[tuple[str, ...]] = tuple(_load_user_setting('DS_MAIL_CC', ()) or ())

# 첨부 형식: 'xlsx' | 'pdf' | 'both'
# 기본이 xlsx인 이유 — 납기현황은 고객이 자기 시스템에 옮겨 넣거나 정렬해 보는 표라서
# PDF보다 원본이 쓸모 있다. 거래명세표(PDF 기본)와 성격이 다르다.
DS_MAIL_ATTACH_FORMAT: Final[str] = _load_user_setting('DS_MAIL_ATTACH_FORMAT', 'xlsx')

# 제목/본문 템플릿
# 치환자: {customer} 한글 거래처명, {customer_en} 영문 거래처명, {date} 기준일,
#         {count} 미출고 건수, {qty} 미출고 수량 합, {supplier} 공급자명,
#         {table} 납기현황 표 (본문 전용 — 평문/HTML 각각 알맞은 형태로 치환된다)
DS_MAIL_SUBJECT: Final[str] = _load_user_setting(
    'DS_MAIL_SUBJECT',
    '[Delivery Schedule] {customer} - {date} 기준',
)
DS_MAIL_BODY: Final[str] = _load_user_setting(
    'DS_MAIL_BODY',
    '{customer} 귀중\n\n'
    '{date} 기준 미출고 {count}건의 납기현황을 아래와 같이 송부하오니 참고 바랍니다.\n\n'
    '{table}\n\n'
    '본 메일은 자동 발송된 메일입니다.\n',
)


# === Order Confirmation 메일 설정 (create_oc.py) — 해외 전용 ===
# 수신자(To)는 `Customer_해외`에서 **고객코드**로 조회한다. 국내 문서(거래명세표·납기현황)가
# 쓰는 사업자번호가 아니다 — `SO_해외`의 'Business registration number' 컬럼에 실제로는
# `C-0054` 같은 고객코드가 들어 있고, 이것이 `Customer_해외.C-code by 해외`와 맞물린다.
OC_MAIL_CC: Final[tuple[str, ...]] = tuple(_load_user_setting('OC_MAIL_CC', ()) or ())

# 첨부 형식: 'pdf' | 'xlsx' | 'both'
# OC는 고객이 보관·회신용으로 받는 확정 통지라 원본(xlsx)을 줄 이유가 없다 — 거래명세표와 같이 PDF.
OC_MAIL_ATTACH_FORMAT: Final[str] = _load_user_setting('OC_MAIL_ATTACH_FORMAT', 'pdf')

# 제목/본문 템플릿 — 해외 고객이 받으므로 **영문**이다.
# 본문은 인사말 + 발주번호 확인 한 줄 + 자동발송 안내 — 안내문·서명 없이 짧게 간다
# (2026-08-05 사용자 결정. 제목이 발신 조직을 밝히고, 서명은 발송자 메일 클라이언트 몫).
# 치환자: {customer} 거래처명(Customer_해외.고객명, 이미 영문), {customer_po} 고객 발주번호,
#         {doc_id} SO_ID(= O.C. No), {date} 발행일, {supplier_en} 영문 상호
#         — 오버라이드에서 서명을 넣는다면 {supplier}(한글)가 아니라 {supplier_en}을 쓸 것
OC_MAIL_SUBJECT: Final[str] = _load_user_setting(
    'OC_MAIL_SUBJECT',
    '[Rotork Controls Korea] Order Confirmation - Your PO: {customer_po}',
)
OC_MAIL_BODY: Final[str] = _load_user_setting(
    'OC_MAIL_BODY',
    'Dear {customer},\n\n'
    'Please find attached our Order Confirmation for your purchase order {customer_po}.\n\n'
    '* This email has been sent automatically.*\n',
)
