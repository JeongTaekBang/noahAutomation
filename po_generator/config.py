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
NOAH_PO_LISTS_FILE: Final[Path] = BASE_DIR / "NOAH_PO_Lists.xlsx"
OUTPUT_DIR: Final[Path] = BASE_DIR / "generated_po"
HISTORY_FILE: Final[Path] = BASE_DIR / "po_history.xlsx"


# === 시트 설정 ===
DOMESTIC_SHEET_INDEX: Final[int] = 0  # 국내
EXPORT_SHEET_INDEX: Final[int] = 1    # 해외


# === Excel 레이아웃 상수 ===
TOTAL_COLUMNS: Final[int] = 10
MAX_ITEMS_PER_PO: Final[int] = 7
ITEM_START_ROW: Final[int] = 13
ITEM_END_ROW: Final[int] = 19


# === 검증 설정 ===
MIN_LEAD_TIME_DAYS: Final[int] = 7


# === 필수 필드 ===
REQUIRED_FIELDS: Final[tuple[str, ...]] = (
    'Customer name',
    'Customer PO',
    'Item qty',
    'Model',
    'ICO Unit',
)


# === 액추에이터 사양 필드 (Description 시트) ===
SPEC_FIELDS: Final[tuple[str, ...]] = (
    'Power supply', 'Motor(kW)', 'BASE', 'ACT Flange', 'Operating time',
    'Handwheel', 'RPM', 'Turns', 'Bushing', 'MOV', 'Gearbox model',
    'Gearbox Flange', 'Gearbox ratio', 'Gearbox position', 'Operating mode',
    'Fail action', 'Enclosure', 'Cable entry', 'Paint', 'Cover tube(mm)',
    'WD code', 'Test report', 'Version', 'Remark2',
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
