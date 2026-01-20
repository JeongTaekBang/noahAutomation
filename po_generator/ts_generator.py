"""
거래명세표 생성 모듈 (xlwings 기반)
====================================

xlwings를 사용하여 템플릿 기반으로 거래명세표를 생성합니다.
이미지, 서식 등이 완벽하게 보존됩니다.

지원 문서 유형:
- DN: 납품 거래명세표 (DN_국내 데이터 사용)
- PMT: 선수금 거래명세표 (PMT_국내 데이터 사용)
"""

from __future__ import annotations

import logging
import shutil
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.config import (
    TS_TEMPLATE_FILE,
    ITEM_START_ROW_FALLBACK,
    VAT_RATE_DOMESTIC,
)
from po_generator.utils import get_value
from po_generator.excel_helpers import (
    find_item_start_row_xlwings,
    TS_HEADER_LABELS,
)

logger = logging.getLogger(__name__)


# === 셀 매핑 (템플릿 기준) ===
CELL_DATE = 'B2'           # DATE : 날짜
CELL_CUSTOMER = 'B7'       # 고객명 귀하

# 레이블 검색 범위 (행 삽입 후에도 이 범위 내에 있음)
LABEL_SEARCH_START = 15
LABEL_SEARCH_END = 50

# 기본 행 위치 (동적 탐지 실패 시 폴백값)
# ITEM_START_ROW는 config의 ITEM_START_ROW_FALLBACK 사용
BASE_PO_ROW = 23           # PO No. 행 (폴백값)
BASE_TOTAL_ROW = 25        # 합계 행 (폴백값)


def _find_ts_item_start_row(
    ws: xw.Sheet,
    search_labels: tuple[str, ...] = TS_HEADER_LABELS,
    max_search_rows: int = 30,
) -> int:
    """거래명세표 템플릿에서 아이템 시작 행을 동적으로 찾기 (xlwings)

    excel_helpers.find_item_start_row_xlwings의 래퍼입니다.
    거래명세표 전용 헤더 라벨과 검색 열을 사용합니다.

    Args:
        ws: xlwings Sheet 객체
        search_labels: 검색할 헤더 레이블
        max_search_rows: 최대 검색 행 수

    Returns:
        아이템 시작 행 번호
    """
    return find_item_start_row_xlwings(
        ws,
        search_labels=search_labels,
        max_search_rows=max_search_rows,
        columns=('A', 'B', 'C', 'D'),  # 거래명세표 헤더는 앞쪽 열에 위치
        fallback_row=ITEM_START_ROW_FALLBACK,
    )


def _find_label_row(ws: xw.Sheet, col: str, search_text: str) -> int | None:
    """지정된 열에서 텍스트를 포함하는 셀의 행 번호 찾기

    Args:
        ws: xlwings Sheet 객체
        col: 검색할 열 (예: 'A', 'E')
        search_text: 찾을 텍스트

    Returns:
        행 번호 또는 None
    """
    for row in range(LABEL_SEARCH_START, LABEL_SEARCH_END + 1):
        cell_value = ws.range(f'{col}{row}').value
        if cell_value and search_text in str(cell_value):
            return row
    return None


def create_ts_xlwings(
    template_path: Path,
    output_path: Path,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
    doc_type: str = 'DN',
) -> None:
    """xlwings로 거래명세표 생성

    Args:
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        order_data: 주문 데이터 (첫 번째 아이템 또는 단일 아이템)
        items_df: 다중 아이템인 경우 전체 아이템 DataFrame
        doc_type: 문서 유형 ('DN' 또는 'PMT')
    """
    if not template_path.exists():
        raise FileNotFoundError(f"템플릿 파일이 없습니다: {template_path}")

    # 날짜
    today = datetime.now()
    today_str = today.strftime("%Y. %m. %d")

    # 경로에 한글이 포함된 경우 임시 폴더에서 작업 (xlwings COM 인터페이스 호환성)
    temp_dir = Path(tempfile.gettempdir())
    temp_template = temp_dir / f"ts_template_{today.strftime('%Y%m%d%H%M%S')}.xlsx"
    temp_output = temp_dir / f"ts_output_{today.strftime('%Y%m%d%H%M%S')}.xlsx"

    # 템플릿을 임시 폴더로 복사
    shutil.copy(template_path, temp_template)

    # Excel 앱 시작 (백그라운드)
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        # 임시 템플릿 열기
        wb = app.books.open(str(temp_template))
        ws = wb.sheets[0]

        # 1. 헤더 정보
        ws.range(CELL_DATE).value = f"DATE : {today_str}"
        customer_name = get_value(order_data, 'customer_name', '')
        ws.range(CELL_CUSTOMER).value = f"{customer_name} 귀하"

        # DN/ADV 공통 처리 (remark만 다름)
        remark = '선수금' if doc_type == 'ADV' else ''
        _fill_ts_data(ws, order_data, items_df, today, remark)

        # 임시 위치에 저장
        wb.save(str(temp_output))
        logger.info(f"거래명세표 생성 완료 (임시): {temp_output}")

    finally:
        # 정리 (wb가 정의되어 있을 때만)
        try:
            wb.close()
        except NameError:
            pass
        app.quit()

        # 임시 템플릿 삭제
        try:
            temp_template.unlink()
        except Exception:
            pass

    # 최종 출력 경로로 이동
    shutil.move(str(temp_output), str(output_path))
    logger.info(f"거래명세표 저장 완료: {output_path}")


def _find_ts_subtotal_row(ws: xw.Sheet, start_row: int, max_search: int = 15) -> int:
    """소계(SUM) 수식이 있는 행 찾기

    Args:
        ws: xlwings Sheet 객체
        start_row: 검색 시작 행
        max_search: 최대 검색 행 수

    Returns:
        소계 행 번호 (못 찾으면 start_row + 3)
    """
    for row in range(start_row, start_row + max_search):
        e_formula = ws.range(f'E{row}').formula
        if e_formula and '=SUM' in str(e_formula):
            return row
    return start_row + 3  # 기본값


def _restore_ts_item_borders(ws: xw.Sheet, item_start_row: int, num_items: int) -> None:
    """행 삭제 후 아이템 영역 테두리 복원

    Args:
        ws: xlwings Sheet 객체
        item_start_row: 아이템 시작 행
        num_items: 실제 아이템 수
    """
    # xlwings 상수
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlContinuous = 1
    xlThin = 2

    # 마지막 아이템 행 (소계 바로 위)
    last_item_row = item_start_row + num_items - 1

    # 헤더 아래 행 (첫 번째 아이템 행 바로 위)의 아래 테두리
    header_bottom_row = item_start_row - 1
    ws.range(f'A{header_bottom_row}:H{header_bottom_row}').api.Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.range(f'A{header_bottom_row}:H{header_bottom_row}').api.Borders(xlEdgeBottom).Weight = xlThin

    # 마지막 아이템 행의 아래 테두리
    ws.range(f'A{last_item_row}:H{last_item_row}').api.Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.range(f'A{last_item_row}:H{last_item_row}').api.Borders(xlEdgeBottom).Weight = xlThin

    logger.debug(f"테두리 복원: Row {header_bottom_row} 하단, Row {last_item_row} 하단")


def _fill_ts_data(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None,
    today: datetime,
    remark: str = '',
) -> None:
    """거래명세표 데이터 채우기 (DN/ADV 공통)

    Args:
        ws: xlwings Sheet 객체
        order_data: 주문 데이터 (첫 번째 아이템)
        items_df: 다중 아이템인 경우 DataFrame
        today: 오늘 날짜
        remark: 비고 텍스트 (예: '선수금')
    """
    # 아이템 준비
    if items_df is None:
        items_df = pd.DataFrame([order_data])
    num_items = len(items_df)

    # 아이템 시작 행 동적 탐지
    item_start_row = _find_ts_item_start_row(ws)

    # 템플릿의 기존 아이템 행 수 계산 (소계 행 찾기)
    subtotal_row = _find_ts_subtotal_row(ws, item_start_row)
    template_item_count = subtotal_row - item_start_row
    logger.debug(f"템플릿 아이템 수: {template_item_count}, 실제 아이템 수: {num_items}")

    # 행 수 조정: 템플릿 예시보다 실제 아이템이 적으면 초과 행 삭제
    if num_items < template_item_count:
        rows_to_delete = template_item_count - num_items
        # 같은 위치에서 반복 삭제 - xlUp으로 아래 행이 올라오므로 연속 행 삭제됨
        for _ in range(rows_to_delete):
            delete_row = item_start_row + num_items
            ws.range(f'{delete_row}:{delete_row}').api.Delete(Shift=-4162)  # xlUp
        logger.debug(f"{rows_to_delete}개 초과 행 삭제")

        # 테두리 복원: 행 삭제로 사라진 테두리 다시 그리기
        _restore_ts_item_borders(ws, item_start_row, num_items)

    # 행 수 조정: 템플릿 예시보다 실제 아이템이 많으면 행 삽입
    elif num_items > template_item_count:
        rows_to_insert = num_items - template_item_count
        source_row = item_start_row
        for i in range(rows_to_insert):
            insert_row = item_start_row + template_item_count + i
            ws.range(f'{source_row}:{source_row}').api.Copy()
            ws.range(f'{insert_row}:{insert_row}').api.Insert(Shift=-4121)  # xlDown
        logger.debug(f"{rows_to_insert}개 행 삽입")

    # 기존 아이템 행 데이터 초기화 (서식은 유지)
    for row in range(item_start_row, item_start_row + num_items):
        ws.range(f'A{row}:H{row}').value = None

    # 아이템 데이터 채우기
    total_amount = 0
    total_tax = 0

    for idx, (_, item) in enumerate(items_df.iterrows()):
        row_num = item_start_row + idx
        item_amount, item_tax = _fill_item_row(ws, row_num, item, today, remark)
        total_amount += item_amount
        total_tax += item_tax

    # 소계 행 수식 업데이트 (다중 아이템인 경우)
    subtotal_row = item_start_row + num_items
    if num_items > 1:
        last_item_row = item_start_row + num_items - 1
        ws.range(f'E{subtotal_row}').formula = f'=SUM(E{item_start_row}:E{last_item_row})'
        ws.range(f'G{subtotal_row}').formula = f'=SUM(G{item_start_row}:G{last_item_row})'
        ws.range(f'H{subtotal_row}').formula = f'=SUM(H{item_start_row}:H{last_item_row})'

    # PO No. 채우기 (레이블 위치를 찾아서 같은 행에 값 입력)
    customer_po = get_value(order_data, 'customer_po', '')
    po_row = _find_label_row(ws, 'A', 'PO No')
    if po_row is None:
        po_row = BASE_PO_ROW + (num_items - 1) if num_items > 1 else BASE_PO_ROW
    ws.range(f'B{po_row}').value = customer_po

    # 합계 채우기 (레이블 위치를 찾아서 같은 행에 값 입력)
    grand_total = total_amount + total_tax
    total_row = _find_label_row(ws, 'E', '합 계')
    if total_row is None:
        total_row = BASE_TOTAL_ROW + (num_items - 1) if num_items > 1 else BASE_TOTAL_ROW
    ws.range(f'G{total_row}').value = grand_total


def _fill_item_row(
    ws: xw.Sheet,
    row_num: int,
    item: pd.Series,
    today: datetime,
    remark: str = '',
) -> tuple[int, int]:
    """아이템 행 데이터 채우기 (DN/ADV 공통)

    Args:
        ws: xlwings Sheet 객체
        row_num: 행 번호
        item: 아이템 데이터 (SO_국내에서 JOIN된 데이터)
        today: 오늘 날짜
        remark: 비고 텍스트 (예: '선수금')

    Returns:
        (금액, 세액) 튜플
    """
    # 월/일
    ws.range(f'A{row_num}').value = f"{today.month}/{today.day}"

    # Description (item_name 키 사용)
    item_name = get_value(item, 'item_name', '')
    ws.range(f'B{row_num}').value = item_name

    # 비고
    ws.range(f'C{row_num}').value = remark

    # 규격
    ws.range(f'D{row_num}').value = "EA"

    # 수량 (item_qty 키 사용)
    qty = get_value(item, 'item_qty', 1)
    try:
        qty = int(qty) if pd.notna(qty) else 1
    except (ValueError, TypeError):
        qty = 1
    ws.range(f'E{row_num}').value = qty

    # 단가 (sales_unit_price 키 사용)
    unit_price = get_value(item, 'sales_unit_price', 0)
    try:
        unit_price = int(float(unit_price)) if pd.notna(unit_price) else 0
    except (ValueError, TypeError):
        unit_price = 0
    ws.range(f'F{row_num}').value = unit_price

    # 금액 (수량 * 단가)
    amount = qty * unit_price
    ws.range(f'G{row_num}').value = amount

    # 세액 (국내 VAT)
    tax = int(amount * VAT_RATE_DOMESTIC)
    ws.range(f'H{row_num}').value = tax

    return amount, tax
