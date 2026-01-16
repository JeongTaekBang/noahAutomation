"""
거래명세표 생성 모듈 (xlwings 기반)
====================================

xlwings를 사용하여 템플릿 기반으로 거래명세표를 생성합니다.
이미지, 서식 등이 완벽하게 보존됩니다.
"""

from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
import xlwings as xw

from po_generator.config import (
    TS_TEMPLATE_FILE,
    TS_ITEM_START_ROW,
)
from po_generator.utils import get_safe_value

logger = logging.getLogger(__name__)


# === 셀 매핑 (템플릿 기준) ===
CELL_DATE = 'B2'           # DATE : 날짜
CELL_CUSTOMER = 'B7'       # 고객명 귀하
CELL_PO_NO = 'B22'         # PO No.
CELL_GRAND_TOTAL = 'G24'   # 합계 (VAT 포함)

# 아이템 시작 행 (A~H 열)
ITEM_START_ROW = 13
# 소계 행 (아이템 1개일 때 Row 14)
SUBTOTAL_BASE_ROW = 14


def create_ts_xlwings(
    template_path: Path,
    output_path: Path,
    order_data: pd.Series,
    items_df: Optional[pd.DataFrame] = None,
) -> None:
    """xlwings로 거래명세표 생성

    Args:
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        order_data: 주문 데이터 (첫 번째 아이템 또는 단일 아이템)
        items_df: 다중 아이템인 경우 전체 아이템 DataFrame
    """
    if not template_path.exists():
        raise FileNotFoundError(f"템플릿 파일이 없습니다: {template_path}")

    # 아이템 준비
    if items_df is None:
        items_df = pd.DataFrame([order_data])
    num_items = len(items_df)

    # 시트 구분 (국내/해외)
    sheet_type = get_safe_value(order_data, '_시트구분', '국내')
    is_domestic = sheet_type == '국내'

    # 날짜
    today = datetime.now()
    today_str = today.strftime("%Y. %m. %d")

    # Excel 앱 시작 (백그라운드)
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        # 템플릿 열기
        wb = app.books.open(str(template_path))
        ws = wb.sheets[0]

        # 1. 헤더 정보
        ws.range(CELL_DATE).value = f"DATE : {today_str}"
        customer_name = get_safe_value(order_data, 'Customer name', '')
        ws.range(CELL_CUSTOMER).value = f"{customer_name} 귀하"

        # 2. 아이템 행 삽입 (다중 아이템인 경우)
        if num_items > 1:
            # Row 14 위에 (num_items - 1)개 행 삽입
            insert_count = num_items - 1
            for _ in range(insert_count):
                ws.range(f'{SUBTOTAL_BASE_ROW}:{SUBTOTAL_BASE_ROW}').insert('down')

        # 3. 아이템 데이터 채우기
        total_amount = 0
        total_tax = 0

        for idx, (_, item) in enumerate(items_df.iterrows()):
            row_num = ITEM_START_ROW + idx
            item_amount, item_tax = _fill_item_row(ws, row_num, item, today, is_domestic)
            total_amount += item_amount
            total_tax += item_tax

        # 4. 소계 행 수식 업데이트 (다중 아이템인 경우)
        subtotal_row = ITEM_START_ROW + num_items
        if num_items > 1:
            last_item_row = ITEM_START_ROW + num_items - 1
            ws.range(f'E{subtotal_row}').formula = f'=SUM(E{ITEM_START_ROW}:E{last_item_row})'
            ws.range(f'G{subtotal_row}').formula = f'=SUM(G{ITEM_START_ROW}:G{last_item_row})'
            ws.range(f'H{subtotal_row}').formula = f'=SUM(H{ITEM_START_ROW}:H{last_item_row})'

        # 5. PO No. 채우기 (행이 밀렸으므로 위치 조정)
        customer_po = get_safe_value(order_data, 'Customer PO', '')
        po_row = 22 + (num_items - 1) if num_items > 1 else 22
        ws.range(f'B{po_row}').value = customer_po

        # 6. 합계 채우기 (행이 밀렸으므로 위치 조정)
        grand_total = total_amount + total_tax
        total_row = 24 + (num_items - 1) if num_items > 1 else 24
        ws.range(f'G{total_row}').value = grand_total

        # 저장
        wb.save(str(output_path))
        logger.info(f"거래명세표 생성 완료: {output_path}")

    finally:
        # 정리
        wb.close()
        app.quit()


def _fill_item_row(
    ws,
    row_num: int,
    item: pd.Series,
    today: datetime,
    is_domestic: bool,
) -> tuple[int, int]:
    """아이템 행 데이터 채우기

    Args:
        ws: xlwings Worksheet
        row_num: 행 번호
        item: 아이템 데이터
        today: 오늘 날짜
        is_domestic: 국내 여부

    Returns:
        (금액, 세액) 튜플
    """
    # 월/일
    ws.range(f'A{row_num}').value = f"{today.month}/{today.day}"

    # Description (Model + Item name)
    model = get_safe_value(item, 'Model', '')
    item_name = get_safe_value(item, 'Item name', '')
    description = model if model else item_name
    if model and item_name and model != item_name:
        description = f"{model} - {item_name}"
    ws.range(f'B{row_num}').value = description

    # 비고
    remark = get_safe_value(item, 'Remark', '')
    ws.range(f'C{row_num}').value = remark

    # 규격
    ws.range(f'D{row_num}').value = "EA"

    # 수량
    qty = get_safe_value(item, 'Item qty', 1)
    try:
        qty = int(qty) if pd.notna(qty) else 1
    except (ValueError, TypeError):
        qty = 1
    ws.range(f'E{row_num}').value = qty

    # 단가 (Sales Unit Price 사용)
    unit_price = get_safe_value(item, 'Sales Unit Price', 0)
    try:
        unit_price = int(float(unit_price)) if pd.notna(unit_price) else 0
    except (ValueError, TypeError):
        unit_price = 0
    ws.range(f'F{row_num}').value = unit_price

    # 금액 (수량 * 단가)
    amount = qty * unit_price
    ws.range(f'G{row_num}').value = amount

    # 세액 (국내: 10%, 해외: 0%)
    if is_domestic:
        tax = int(amount * 0.1)
    else:
        tax = 0
    ws.range(f'H{row_num}').value = tax

    return amount, tax
