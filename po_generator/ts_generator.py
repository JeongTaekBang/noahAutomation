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
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.config import (
    TS_TEMPLATE_FILE,
    TS_COLUMN_WIDTHS,
    TS_PO_REMARK_CUSTOMERS,
    ITEM_START_ROW_FALLBACK,
    VAT_RATE_DOMESTIC,
)
from po_generator.utils import get_value
from po_generator.excel_helpers import (
    XlConstants,
    xlwings_app_context,
    prepare_template,
    cleanup_temp_file,
    find_item_start_row_xlwings,
    TS_HEADER_LABELS,
    batch_write_rows,
    delete_rows_range,
    find_text_in_column_batch,
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

# SO 비고(호선명) 컬럼 — 문서 경로마다 이름이 다르다.
# DN 경로는 load_dn_data()가 DN 자체 Remarks(월합 세금계산서 문구)와 구분하려고
# 'SO Remarks'로 개명해 병합하고, 선수금(ADV) 경로는 SO_국내 행 그대로라 'Remarks'다.
# 컬럼 유무로 추측하면 안 된다 — DN 아이템에도 'Remarks'는 있는데 그건 월합 문구다.
SO_REMARK_COLUMNS = {'DN': 'SO Remarks', 'ADV': 'Remarks'}


def resolve_so_remark_column(
    order_data: pd.Series,
    items_df: pd.DataFrame | None,
    doc_type: str = 'DN',
) -> str | None:
    """호선명을 표기할 문서인지 판정 — 맞으면 SO 비고 컬럼명, 아니면 None

    하단 PO No. 병기와 본문 비고(C열) 표기가 같은 관문을 쓴다 —
    지정 거래처(config.TS_PO_REMARK_CUSTOMERS, 고객명 부분일치)이고
    경로별 SO 비고 컬럼(SO_REMARK_COLUMNS)이 실제로 있을 때만.
    """
    if items_df is None:
        return None
    remark_col = SO_REMARK_COLUMNS.get(doc_type)
    customer_name = str(get_value(order_data, 'customer_name', '') or '')
    if (
        remark_col is not None
        and remark_col in items_df.columns
        and any(keyword in customer_name for keyword in TS_PO_REMARK_CUSTOMERS)
    ):
        return remark_col
    return None


def build_row_remark(
    item: pd.Series,
    remark: str = '',
    use_po_as_remark: bool = False,
    so_remark_col: str | None = None,
) -> str:
    """아이템 행 비고(C열) 텍스트

    우선순위 — 기존 표기를 밀어내지 않는다:
    1. 월합(use_po_as_remark): 행별 Customer PO (하단에서 발주번호↔호선 짝을 보인다)
    2. 공통 remark: 선수금 문서의 '선수금' 표기
    3. 지정 거래처의 호선명(SO 비고) — 비고가 비어 있을 자리에만 들어간다.
       한 문서에 호선이 여럿이면(실측 DND-2026-0790: 12행에 배 3척) 하단 나열만으로는
       어느 행이 어느 배인지 알 수 없어서 행마다 적는다
    """
    if use_po_as_remark:
        return str(get_value(item, 'customer_po', '') or '')
    if remark:
        return remark
    if so_remark_col is not None:
        value = item.get(so_remark_col)
        return '' if pd.isna(value) else str(value).strip()
    return ''


def build_po_cell_text(
    order_data: pd.Series,
    items_df: pd.DataFrame | None,
    doc_type: str = 'DN',
) -> str:
    """하단 PO No. 칸 텍스트 — 발주번호 나열 + (지정 거래처만) SO 비고(호선명) 병기

    조선 기자재 거래처(config.TS_PO_REMARK_CUSTOMERS)는 발주가 호선 단위라,
    거래명세표에도 호선명이 있어야 고객이 어느 배 물량인지 안다.
    형식은 발주번호별 괄호 병기 — 'OR26060022 (H-8327)'. 한 발주에 호선이 여럿이면
    나열한다: 'SCT2605-134 (한화-H4394, 한화-H4395, 한화-H4396)' (2026-08 실측).
    비고가 빈 발주번호는 지금처럼 번호만 적는다.

    Args:
        order_data: 대표 행 (거래처 판정용 고객명)
        items_df: 문서에 실린 전체 아이템 (None이면 order_data 단건)
        doc_type: 'DN' 또는 'ADV' — SO 비고 컬럼명이 갈린다 (SO_REMARK_COLUMNS)

    Returns:
        PO No. 칸에 쓸 문자열 (여러 발주번호면 콤마로 구분)
    """
    if items_df is None or 'Customer PO' not in items_df.columns:
        return str(get_value(order_data, 'customer_po', '') or '')

    # 발주번호: 등장 순서 유지 + 중복 제거 (기존 동작)
    po_values = list(dict.fromkeys(
        str(v).strip() for v in items_df['Customer PO'].dropna() if str(v).strip()
    ))

    remark_col = resolve_so_remark_column(order_data, items_df, doc_type)
    if remark_col is None:
        return ', '.join(po_values)

    # 발주번호 ↔ 호선명 짝짓기 (호선명은 SO 라인 단위라 아이템 행에서 그대로 짝이 나온다)
    remarks_by_po: dict[str, list[str]] = {}
    for _, item in items_df.iterrows():
        po = item.get('Customer PO')
        remark = item.get(remark_col)
        po_key = '' if pd.isna(po) else str(po).strip()
        text = '' if pd.isna(remark) else str(remark).strip()
        if po_key and text and text not in remarks_by_po.setdefault(po_key, []):
            remarks_by_po[po_key].append(text)

    return ', '.join(
        f"{po} ({', '.join(remarks_by_po[po])})" if remarks_by_po.get(po) else po
        for po in po_values
    )


def create_ts_xlwings(
    template_path: Path,
    output_path: Path,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
    doc_type: str = 'DN',
    use_po_as_remark: bool = False,
) -> None:
    """xlwings로 거래명세표 생성

    Args:
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        order_data: 주문 데이터 (첫 번째 아이템 또는 단일 아이템)
        items_df: 다중 아이템인 경우 전체 아이템 DataFrame
        doc_type: 문서 유형 ('DN' 또는 'PMT')
        use_po_as_remark: 월합 거래명세표 — 각 행의 비고(C열)에 해당 아이템의 Customer PO 표기
    """
    # 템플릿 준비 (임시 폴더로 복사)
    temp_template, temp_output = prepare_template(template_path, "ts")

    # 출고일 가져오기 (없으면 오늘 날짜 사용)
    dispatch_date = get_value(order_data, 'dispatch_date', None)
    if dispatch_date is None or pd.isna(dispatch_date):
        dispatch_date = datetime.now()
    elif not isinstance(dispatch_date, datetime):
        try:
            dispatch_date = pd.to_datetime(dispatch_date)
        except (ValueError, TypeError):
            dispatch_date = datetime.now()
    dispatch_date_str = dispatch_date.strftime("%Y. %m. %d")

    try:
        # xlwings App 생명주기 관리
        with xlwings_app_context() as app:
            # 임시 템플릿 열기
            wb = app.books.open(str(temp_template))
            ws = wb.sheets[0]

            # 0. 열 너비 — config가 단일 소스 (템플릿 값 무시). 줄바꿈·행 높이
            # 자동 조정(_wrap_item_text)이 열 너비에 좌우되므로 데이터보다 먼저 적용
            _apply_column_widths(ws)

            # 1. 헤더 정보 (출고일 사용)
            ws.range(CELL_DATE).value = f"DATE : {dispatch_date_str}"
            customer_name = get_value(order_data, 'customer_name', '')
            ws.range(CELL_CUSTOMER).value = f"{customer_name} 귀하"

            # DN/ADV 공통 처리 (remark만 다름)
            remark = '선수금' if doc_type == 'ADV' else ''
            _fill_ts_data(
                ws, order_data, items_df, dispatch_date, remark, use_po_as_remark,
                doc_type=doc_type,
            )

            # 임시 위치에 저장
            wb.save(str(temp_output))
            logger.info(f"거래명세표 생성 완료 (임시): {temp_output}")

    finally:
        # 임시 템플릿 삭제
        cleanup_temp_file(temp_template)

    # 최종 출력 경로로 이동
    shutil.move(str(temp_output), str(output_path))
    logger.info(f"거래명세표 저장 완료: {output_path}")


def _apply_column_widths(ws: xw.Sheet) -> None:
    """열 너비 적용 — config.TS_COLUMN_WIDTHS가 단일 소스

    템플릿 파일의 열 너비를 매번 덮어쓴다. 템플릿 이진 파일 속 값은 리뷰가 안 보여서
    한번 어긋나면 계속 어긋난 채 나간다. 너비 근거(품명·비고를 넓히되 F·H는 상단
    공급자 박스가 폭을 고정한다)는 config의 TSColumnWidths 주석이 소유한다.
    """
    for col, width in TS_COLUMN_WIDTHS.as_dict().items():
        ws.range(f'{col}1').api.EntireColumn.ColumnWidth = width


def _find_ts_subtotal_row(ws: xw.Sheet, start_row: int, max_search: int = 15) -> int:
    """소계(SUM) 수식이 있는 행 찾기 (배치 읽기 최적화)

    Args:
        ws: xlwings Sheet 객체
        start_row: 검색 시작 행
        max_search: 최대 검색 행 수

    Returns:
        소계 행 번호 (못 찾으면 start_row + 3)
    """
    # 배치 읽기로 15회 COM 호출 → 1회로 감소
    end_row = start_row + max_search - 1
    formulas = ws.range(f'E{start_row}:E{end_row}').formula

    # xlwings 범위 읽기는 tuple of tuples 반환: (('val1',), ('val2',), ...)
    # 단일 셀은 문자열 반환
    if isinstance(formulas, (list, tuple)) and formulas and isinstance(formulas[0], (list, tuple)):
        # 2D → 1D 평탄화 (각 행의 첫 번째 값만 추출)
        formulas = [f[0] if f else '' for f in formulas]
    elif not isinstance(formulas, (list, tuple)):
        formulas = [formulas]

    for idx, formula in enumerate(formulas):
        if formula and '=SUM' in str(formula):
            return start_row + idx

    return start_row + 3  # 기본값


def _restore_ts_item_borders(ws: xw.Sheet, item_start_row: int, num_items: int) -> None:
    """행 삭제 후 아이템 영역 테두리 복원

    Args:
        ws: xlwings Sheet 객체
        item_start_row: 아이템 시작 행
        num_items: 실제 아이템 수
    """
    # 마지막 아이템 행 (소계 바로 위)
    last_item_row = item_start_row + num_items - 1

    # 헤더 아래 행 (첫 번째 아이템 행 바로 위)의 아래 테두리
    header_bottom_row = item_start_row - 1
    ws.range(f'A{header_bottom_row}:H{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{header_bottom_row}:H{header_bottom_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    # 마지막 아이템 행의 아래 테두리
    ws.range(f'A{last_item_row}:H{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).LineStyle = XlConstants.xlContinuous
    ws.range(f'A{last_item_row}:H{last_item_row}').api.Borders(XlConstants.xlEdgeBottom).Weight = XlConstants.xlThin

    logger.debug(f"테두리 복원: Row {header_bottom_row} 하단, Row {last_item_row} 하단")


def _wrap_item_text(ws: xw.Sheet, item_start_row: int, end_row: int) -> None:
    """아이템 행의 품명(B)·비고(C) 줄바꿈 + 행 높이 자동 조정

    긴 품명이 열 너비에서 잘리는 것을 막습니다. 숫자 열은 줄바꿈 대상이 아니며,
    행이 세로로 늘어나므로 값들이 위로 붙지 않도록 세로 가운데 정렬합니다.

    Args:
        ws: xlwings Sheet 객체
        item_start_row: 아이템 시작 행
        end_row: 아이템 마지막 행
    """
    try:
        ws.range(f'B{item_start_row}:C{end_row}').api.WrapText = True
        ws.range(f'A{item_start_row}:H{end_row}').api.VerticalAlignment = XlConstants.xlCenter
        ws.range(f'A{item_start_row}:A{end_row}').api.EntireRow.AutoFit()
        logger.debug(f"아이템 행 줄바꿈/높이 조정: {item_start_row}~{end_row}")
    except Exception as e:
        # 서식 조정 실패가 문서 생성 자체를 막지는 않게 한다
        logger.warning(f"아이템 행 줄바꿈 설정 실패 (내용은 정상): {e}")


def _fill_ts_data(
    ws: xw.Sheet,
    order_data: pd.Series,
    items_df: pd.DataFrame | None,
    dispatch_date: datetime,
    remark: str = '',
    use_po_as_remark: bool = False,
    doc_type: str = 'DN',
) -> None:
    """거래명세표 데이터 채우기 (DN/ADV 공통) - 배치 쓰기 최적화

    Args:
        ws: xlwings Sheet 객체
        order_data: 주문 데이터 (첫 번째 아이템)
        items_df: 다중 아이템인 경우 DataFrame
        dispatch_date: 출고일
        remark: 비고 텍스트 (예: '선수금')
        use_po_as_remark: 월합 거래명세표 — 각 행 비고(C열)에 해당 아이템의 Customer PO 표기
        doc_type: 'DN' 또는 'ADV' — PO No. 칸의 호선명 병기 시 SO 비고 컬럼명이 갈린다
    """
    # 아이템 준비
    if items_df is None:
        items_df = pd.DataFrame([order_data])
    num_items = len(items_df)

    # 아이템 시작 행 동적 탐지
    item_start_row = find_item_start_row_xlwings(
        ws,
        search_labels=TS_HEADER_LABELS,
        columns=('A', 'B', 'C', 'D'),
        fallback_row=ITEM_START_ROW_FALLBACK,
    )

    # 템플릿의 기존 아이템 행 수 계산 (소계 행 찾기)
    subtotal_row = _find_ts_subtotal_row(ws, item_start_row)
    template_item_count = subtotal_row - item_start_row
    logger.debug(f"템플릿 아이템 수: {template_item_count}, 실제 아이템 수: {num_items}")

    # 행 수 조정: 템플릿 예시보다 실제 아이템이 적으면 초과 행 삭제
    if num_items < template_item_count:
        rows_to_delete = template_item_count - num_items
        # 범위 삭제로 N회 COM 호출 → 1회로 감소
        delete_rows_range(ws, item_start_row + num_items, rows_to_delete)

        # 테두리 복원: 행 삭제로 사라진 테두리 다시 그리기
        _restore_ts_item_borders(ws, item_start_row, num_items)

    # 행 수 조정: 템플릿 예시보다 실제 아이템이 많으면 행 삽입
    elif num_items > template_item_count:
        rows_to_insert = num_items - template_item_count
        source_row = item_start_row
        for i in range(rows_to_insert):
            insert_row = item_start_row + template_item_count + i
            ws.range(f'{source_row}:{source_row}').api.Copy()
            ws.range(f'{insert_row}:{insert_row}').api.Insert(Shift=XlConstants.xlShiftDown)
        logger.debug(f"{rows_to_insert}개 행 삽입")

    # 기존 아이템 행 데이터 초기화 (서식은 유지) - 배치 초기화
    end_row = item_start_row + num_items - 1
    ws.range(f'A{item_start_row}:H{end_row}').value = None

    # 아이템 데이터 배치 쓰기 (N개 아이템 * 8열 COM 호출 → 1회로 감소)
    so_remark_col = resolve_so_remark_column(order_data, items_df, doc_type)
    total_amount, total_tax = _fill_items_batch(
        ws, item_start_row, items_df, dispatch_date, remark, use_po_as_remark,
        so_remark_col=so_remark_col,
    )

    # 품명·비고 줄바꿈 + 행 높이 자동 조정
    # 품명은 최대 64자(전체의 27%가 B열 너비 22를 넘음), 비고는 36자까지 들어오는데
    # 줄바꿈이 꺼져 있으면 열 너비에서 잘려 거래처가 품목을 알 수 없다.
    # 인쇄가 fitToPage(축소)라 열을 넓히면 글자만 작아지므로, 너비는 그대로 두고
    # 세로로 늘린다. 아이템 행에는 병합 셀이 없어 AutoFit이 안전하다.
    _wrap_item_text(ws, item_start_row, end_row)

    # 소계 행 수식 업데이트 (다중 아이템인 경우)
    subtotal_row = item_start_row + num_items
    if num_items > 1:
        last_item_row = item_start_row + num_items - 1
        ws.range(f'E{subtotal_row}').formula = f'=SUM(E{item_start_row}:E{last_item_row})'
        ws.range(f'G{subtotal_row}').formula = f'=SUM(G{item_start_row}:G{last_item_row})'
        ws.range(f'H{subtotal_row}').formula = f'=SUM(H{item_start_row}:H{last_item_row})'

    # 행 삽입/삭제로 라벨이 이동한 양 (음수면 라벨이 위로 올라옴)
    row_shift = num_items - template_item_count
    label_search_end = LABEL_SEARCH_END + max(0, row_shift)

    # PO No. 채우기 (여러 발주번호면 콤마로 구분, 지정 거래처는 SO 비고의 호선명 병기)
    customer_po = build_po_cell_text(order_data, items_df, doc_type)
    po_row = find_text_in_column_batch(ws, 'A', 'PO No', LABEL_SEARCH_START, label_search_end)
    if po_row is None:
        po_row = BASE_PO_ROW + row_shift
    ws.range(f'B{po_row}').value = customer_po

    # 합계 채우기 (레이블 위치를 찾아서 같은 행에 값 입력)
    grand_total = total_amount + total_tax
    total_row = find_text_in_column_batch(ws, 'E', '합 계', LABEL_SEARCH_START, label_search_end)
    if total_row is None:
        total_row = BASE_TOTAL_ROW + row_shift
    ws.range(f'G{total_row}').value = grand_total


def _fill_items_batch(
    ws: xw.Sheet,
    item_start_row: int,
    items_df: pd.DataFrame,
    dispatch_date: datetime,
    remark: str = '',
    use_po_as_remark: bool = False,
    so_remark_col: str | None = None,
) -> tuple[int, int]:
    """아이템 데이터 배치 쓰기 (성능 최적화)

    N개 아이템 * 8열 = N*8회 COM 호출 → 1회로 감소

    Args:
        ws: xlwings Sheet 객체
        item_start_row: 아이템 시작 행
        items_df: 아이템 DataFrame
        dispatch_date: 출고일
        remark: 비고 텍스트
        use_po_as_remark: True면 행별 비고를 해당 아이템의 Customer PO로 채움
        so_remark_col: 지정 거래처의 호선명 컬럼 (resolve_so_remark_column 결과, 아니면 None)

    Returns:
        (총 금액, 총 세액)
    """
    data_2d = []
    total_amount = 0
    total_tax = 0
    default_date_str = f"{dispatch_date.month}월 {dispatch_date.day}일"

    for item_idx, (_, item) in enumerate(items_df.iterrows()):
        # 아이템별 출고일 (없으면 기본 출고일 사용)
        item_date = item.get('출고일', None)
        if item_date is not None and pd.notna(item_date):
            try:
                if not isinstance(item_date, (datetime, pd.Timestamp)):
                    item_date = pd.to_datetime(item_date)
                date_str = f"{item_date.month}월 {item_date.day}일"
            except (ValueError, TypeError):
                date_str = default_date_str
        else:
            date_str = default_date_str

        # 수량 ('2.0' 같은 숫자형 문자열도 안전하게 파싱 — int('2.0')은 ValueError)
        raw_qty = get_value(item, 'item_qty', 1)
        try:
            qty = int(float(raw_qty)) if pd.notna(raw_qty) else 1
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 수량 변환 실패 '{raw_qty}' -> 기본값 1 사용")
            qty = 1

        # 단가
        raw_price = get_value(item, 'sales_unit_price', 0)
        try:
            unit_price = int(float(raw_price)) if pd.notna(raw_price) else 0
        except (ValueError, TypeError):
            logger.warning(f"Item {item_idx+1}: 단가 변환 실패 '{raw_price}' -> 기본값 0 사용")
            unit_price = 0

        # 금액/세액 계산
        amount = qty * unit_price
        tax = int(amount * VAT_RATE_DOMESTIC)
        total_amount += amount
        total_tax += tax

        # 비고: 월합=행별 PO > 선수금 > 지정 거래처 호선명 (build_row_remark)
        row_remark = build_row_remark(item, remark, use_po_as_remark, so_remark_col)

        # 행 데이터: A(월/일), B(품명), C(비고), D(규격), E(수량), F(단가), G(금액), H(세액)
        data_2d.append([
            date_str,
            get_value(item, 'item_name', ''),
            row_remark,
            "EA",
            qty,
            unit_price,
            amount,
            tax,
        ])

    # 한 번에 쓰기
    batch_write_rows(ws, f'A{item_start_row}', data_2d)

    return total_amount, total_tax


