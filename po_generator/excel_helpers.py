"""
Excel 헬퍼 함수 모듈
====================

여러 모듈에서 공통으로 사용하는 Excel 관련 유틸리티 함수를 제공합니다.

- XlConstants: Excel COM 매직 넘버 상수
- xlwings_app_context: xlwings App 생명주기 관리 컨텍스트 매니저
- prepare_template: 템플릿 파일을 임시 폴더로 복사
- find_item_start_row_openpyxl: openpyxl 워크시트용
- find_item_start_row_xlwings: xlwings 워크시트용
- batch_write_rows: 2D 리스트를 한 번에 쓰기 (성능 최적화)
- batch_read_column: 열의 값을 한 번에 읽기 (성능 최적화)
- delete_rows_range: 연속 행을 한 번에 삭제 (성능 최적화)
"""

from __future__ import annotations

import logging
import shutil
import tempfile
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Generator

import xlwings as xw

if TYPE_CHECKING:
    from openpyxl.worksheet.worksheet import Worksheet as OpenpyxlWorksheet

from po_generator.config import ITEM_START_ROW_FALLBACK

logger = logging.getLogger(__name__)


# === Excel COM 상수 클래스 ===

class XlConstants:
    """Excel COM 인터페이스 매직 넘버 상수

    xlwings에서 Excel COM API를 직접 호출할 때 사용하는 상수들입니다.
    매직 넘버 대신 이 상수들을 사용하여 코드 가독성을 높입니다.

    참고: https://docs.microsoft.com/en-us/office/vba/api/excel.xlshiftdirection
    """
    # Shift 방향
    xlShiftUp = -4162      # 행 삭제 시 아래 행이 위로 올라옴
    xlShiftDown = -4121    # 행 삽입 시 기존 행이 아래로 내려옴

    # 테두리 위치
    xlEdgeLeft = 7         # 왼쪽 테두리
    xlEdgeTop = 8          # 상단 테두리
    xlEdgeBottom = 9       # 하단 테두리
    xlEdgeRight = 10       # 오른쪽 테두리
    xlInsideVertical = 11  # 내부 세로선
    xlInsideHorizontal = 12  # 내부 가로선

    # 테두리 스타일
    xlContinuous = 1       # 실선
    xlNone = -4142         # 테두리 없음
    xlThin = 2             # 얇은 선
    xlMedium = -4138       # 중간 두께


# === xlwings 앱 컨텍스트 매니저 ===

@contextmanager
def xlwings_app_context(
    visible: bool = False,
    display_alerts: bool = False,
    screen_updating: bool = False,
) -> Generator[xw.App, None, None]:
    """xlwings App 생명주기를 안전하게 관리하는 컨텍스트 매니저

    리소스 누수 방지를 위해 오류 발생 시에도 Excel 프로세스를 정리합니다.

    Args:
        visible: Excel 창 표시 여부 (기본: False)
        display_alerts: Excel 알림 표시 여부 (기본: False)
        screen_updating: 화면 업데이트 여부 (기본: False, 성능 향상)

    Yields:
        xw.App: xlwings App 객체

    Example:
        with xlwings_app_context() as app:
            wb = app.books.open(str(template_path))
            ws = wb.sheets[0]
            # ... 작업 수행 ...
            wb.save(str(output_path))
        # 컨텍스트 종료 시 자동으로 정리됨
    """
    app = None
    try:
        app = xw.App(visible=visible)
        app.display_alerts = display_alerts
        app.screen_updating = screen_updating
        yield app
    finally:
        if app is not None:
            # 모든 워크북 닫기 시도
            try:
                for wb in app.books:
                    try:
                        wb.close()
                    except Exception:
                        pass
            except Exception:
                pass
            # App 종료
            try:
                app.quit()
            except Exception:
                pass


# === 템플릿 준비 헬퍼 ===

def prepare_template(template_path: Path, prefix: str = "template") -> tuple[Path, Path]:
    """템플릿 파일을 임시 디렉토리에 복사하고 경로 반환

    xlwings COM 인터페이스는 한글 경로에서 문제가 발생할 수 있습니다.
    템플릿을 임시 폴더로 복사하여 이 문제를 우회합니다.

    Args:
        template_path: 원본 템플릿 파일 경로
        prefix: 임시 파일 접두사 (기본: "template")

    Returns:
        (temp_template_path, temp_output_path): 임시 템플릿 경로와 출력 경로

    Raises:
        FileNotFoundError: 템플릿 파일이 없는 경우

    Example:
        temp_template, temp_output = prepare_template(PO_TEMPLATE_FILE, "po")
        try:
            # ... 작업 수행 ...
        finally:
            cleanup_temp_file(temp_template)
    """
    if not template_path.exists():
        raise FileNotFoundError(f"템플릿 파일이 없습니다: {template_path}")

    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    temp_dir = Path(tempfile.gettempdir())

    temp_template = temp_dir / f"{prefix}_template_{timestamp}.xlsx"
    temp_output = temp_dir / f"{prefix}_output_{timestamp}.xlsx"

    shutil.copy(template_path, temp_template)
    logger.debug(f"템플릿 복사 완료: {template_path} -> {temp_template}")

    return temp_template, temp_output


def cleanup_temp_file(temp_file: Path) -> None:
    """임시 파일을 안전하게 삭제

    Args:
        temp_file: 삭제할 임시 파일 경로
    """
    try:
        if temp_file.exists():
            temp_file.unlink()
            logger.debug(f"임시 파일 삭제: {temp_file}")
    except Exception as e:
        logger.warning(f"임시 파일 삭제 실패: {temp_file} - {e}")


# === 헤더 라벨 프리셋 ===

# Purchase Order 헤더 라벨 (PO, 이력 추출용)
PO_HEADER_LABELS: tuple[str, ...] = (
    'No.',
    'Item Number',
    'Item\nNumber',
    'Item',
)

# 거래명세표 헤더 라벨
TS_HEADER_LABELS: tuple[str, ...] = (
    '월/일',
    '품명',
    'DESCRIPTION',
)

# Proforma Invoice 헤더 라벨
PI_HEADER_LABELS: tuple[str, ...] = (
    'No.',
    'Description',
    'DESCRIPTION',
)

# Final Invoice 헤더 라벨
FI_HEADER_LABELS: tuple[str, ...] = (
    'No.',
    'Description',
    'DESCRIPTION',
)

# 기본 헤더 라벨 (모든 문서 유형에 공통)
DEFAULT_HEADER_LABELS: tuple[str, ...] = (
    'No.',
    'Item Number',
    'Item\nNumber',
    '품명',
    'Item',
)


def find_item_start_row_openpyxl(
    ws: OpenpyxlWorksheet,
    search_labels: tuple[str, ...] = DEFAULT_HEADER_LABELS,
    max_search_rows: int = 30,
    max_search_cols: int = 9,
    fallback_row: int = ITEM_START_ROW_FALLBACK,
) -> int:
    """템플릿에서 아이템 시작 행을 동적으로 찾기 (openpyxl 버전)

    헤더 레이블을 찾아서 그 다음 행이 아이템 시작 위치입니다.

    Args:
        ws: openpyxl Worksheet 객체
        search_labels: 검색할 헤더 레이블
        max_search_rows: 최대 검색 행 수
        max_search_cols: 최대 검색 열 수 (기본: 9, A-I)
        fallback_row: 헤더를 찾지 못했을 때 기본값

    Returns:
        아이템 시작 행 번호
    """
    for row in range(1, max_search_rows + 1):
        for col in range(1, max_search_cols + 1):
            cell_value = ws.cell(row=row, column=col).value
            if cell_value and any(
                label in str(cell_value) for label in search_labels
            ):
                logger.debug(
                    f"헤더 발견: Row {row}, 값='{cell_value}' -> 아이템 시작 Row {row + 1}"
                )
                return row + 1  # 레이블 다음 행이 데이터 시작

    logger.debug(f"헤더를 찾지 못함 -> 기본값 Row {fallback_row} 사용")
    return fallback_row


def find_item_start_row_xlwings(
    ws: xw.Sheet,
    search_labels: tuple[str, ...] = DEFAULT_HEADER_LABELS,
    max_search_rows: int = 30,
    columns: tuple[str, ...] = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'),
    fallback_row: int = ITEM_START_ROW_FALLBACK,
) -> int:
    """템플릿에서 아이템 시작 행을 동적으로 찾기 (xlwings 버전)

    헤더 레이블을 찾아서 그 다음 행이 아이템 시작 위치입니다.

    Args:
        ws: xlwings Sheet 객체
        search_labels: 검색할 헤더 레이블
        max_search_rows: 최대 검색 행 수
        columns: 검색할 열 문자 튜플
        fallback_row: 헤더를 찾지 못했을 때 기본값

    Returns:
        아이템 시작 행 번호
    """
    for row in range(1, max_search_rows + 1):
        for col in columns:
            cell_value = ws.range(f'{col}{row}').value
            if cell_value and any(
                label in str(cell_value) for label in search_labels
            ):
                logger.debug(
                    f"헤더 발견: Row {row}, 값='{cell_value}' -> 아이템 시작 Row {row + 1}"
                )
                return row + 1

    logger.debug(f"헤더를 찾지 못함 -> 기본값 Row {fallback_row} 사용")
    return fallback_row


# === 배치 연산 헬퍼 함수 (성능 최적화) ===

def batch_write_rows(
    ws: xw.Sheet,
    start_cell: str,
    data_2d: list[list],
) -> None:
    """2D 리스트를 한 번에 쓰기 (xlwings)

    여러 셀에 데이터를 쓸 때 셀 단위 COM 호출 대신 범위 쓰기를 사용하여
    성능을 크게 개선합니다.

    Args:
        ws: xlwings Sheet 객체
        start_cell: 시작 셀 주소 (예: 'A10')
        data_2d: 2D 리스트 (각 행은 리스트, [[row1], [row2], ...])

    Example:
        # 50개 아이템 * 8열 = 400회 COM 호출 → 1회로 감소
        data = [[date, name, remark, "EA", qty, price, amount, tax] for item in items]
        batch_write_rows(ws, 'A10', data)
    """
    if not data_2d:
        return

    num_rows = len(data_2d)
    num_cols = len(data_2d[0]) if data_2d else 0

    if num_cols == 0:
        return

    # 시작 셀에서 열 문자와 행 번호 추출
    col = ''.join(c for c in start_cell if c.isalpha())
    row = int(''.join(c for c in start_cell if c.isdigit()))

    # 끝 열 계산
    end_col = chr(ord(col) + num_cols - 1)
    end_row = row + num_rows - 1

    # 한 번에 쓰기
    ws.range(f'{col}{row}:{end_col}{end_row}').value = data_2d
    logger.debug(f"배치 쓰기 완료: {col}{row}:{end_col}{end_row} ({num_rows}행 x {num_cols}열)")


def batch_read_column(
    ws: xw.Sheet,
    col: str,
    start_row: int,
    end_row: int,
) -> list:
    """열의 값을 한 번에 읽기 (xlwings)

    라벨 검색 등에서 셀 단위 COM 호출 대신 범위 읽기를 사용하여
    성능을 크게 개선합니다.

    Args:
        ws: xlwings Sheet 객체
        col: 열 문자 (예: 'A')
        start_row: 시작 행 번호
        end_row: 끝 행 번호

    Returns:
        값 리스트 (None 포함 가능)

    Example:
        # 36회 COM 호출 → 1회로 감소
        values = batch_read_column(ws, 'A', 15, 50)
        for idx, val in enumerate(values):
            if val and 'PO No' in str(val):
                return 15 + idx
    """
    values = ws.range(f'{col}{start_row}:{col}{end_row}').value

    # 단일 셀인 경우 리스트로 변환
    if not isinstance(values, list):
        values = [values]

    return values


def delete_rows_range(
    ws: xw.Sheet,
    start_row: int,
    count: int,
) -> None:
    """연속 행을 한 번에 삭제 (xlwings)

    반복 삭제 대신 범위 삭제를 사용하여 성능을 개선합니다.

    Args:
        ws: xlwings Sheet 객체
        start_row: 삭제 시작 행 번호
        count: 삭제할 행 수

    Example:
        # 5회 COM 호출 → 1회로 감소
        delete_rows_range(ws, 10, 5)  # Row 10-14 삭제
    """
    if count <= 0:
        return

    end_row = start_row + count - 1
    # xlShiftUp: 삭제 후 아래 행이 위로 올라옴
    ws.range(f'{start_row}:{end_row}').api.Delete(Shift=XlConstants.xlShiftUp)
    logger.debug(f"범위 삭제 완료: Row {start_row}-{end_row} ({count}행)")


def find_text_in_column_batch(
    ws: xw.Sheet,
    col: str,
    search_text: str,
    start_row: int,
    end_row: int,
) -> int | None:
    """배치 읽기로 열에서 텍스트 찾기 (xlwings)

    셀 단위 검색 대신 범위 읽기 후 Python에서 검색하여 성능을 개선합니다.

    Args:
        ws: xlwings Sheet 객체
        col: 검색할 열 (예: 'A')
        search_text: 찾을 텍스트 (부분 일치)
        start_row: 검색 시작 행
        end_row: 검색 끝 행

    Returns:
        찾은 행 번호 또는 None

    Example:
        # 36회 COM 호출 → 1회로 감소
        row = find_text_in_column_batch(ws, 'A', 'PO No', 15, 50)
    """
    values = batch_read_column(ws, col, start_row, end_row)

    for idx, val in enumerate(values):
        if val and search_text in str(val):
            return start_row + idx

    return None


