"""
Certificate of Origin 생성 모듈 (xlwings 기반)
================================================

DN_해외 한 건으로 **수출자(RCK) 원산지 선언서**를 만든다 — 선언·서명은 RCK(CI의 Seller),
생산 공장 NOAH는 Manufacturer로 밝힌다. 근거는 NOAH가 RCK에 내주는 원산지확인서다.
양식(회사 정보·선언문·서명 이미지)은 템플릿 파일이 소유하고, 생성기는 값만 쓴다.

CI·PL과 다른 점:
- 품목 행에 **병합이 없다**. 그래서 행 높이는 Excel 자동 맞춤이 그대로 먹고,
  행 복사 삽입이 병합을 잃는 문제(`excel_helpers.ensure_row_merges`)도 생기지 않는다
  — `layout_item_rows()`가 필요 없는 이유다.
- **운임 줄은 싣지 않는다.** 원산지를 증명할 물건이 아니다 (CI는 HS 빈칸으로 싣는다).
- 표의 주문 참조는 **Customer PO**(업체가 아는 번호, CI의 PO No.와 같다)이고 번호는 **1부터 순번**이다.
  SO `Line item`은 고객 PO 라인과 다른 우리 번호라 Customer PO와 짝지으면 거짓말이 된다 — 실측:
  Watergates PO `WGDBS2601233`을 SO 5개로 받아 `DNO-2026-0020`에 'PO / Line 1'이 다섯 번 나온다.
  우리 SO 번호(SO_ID)는 헤더 Order No.에 남는다.

값을 정하는 부분(`build_coo_lines`·`build_coo_header`)은 COM 없는 순수 함수라
실데이터 없이 테스트한다.
"""

from __future__ import annotations

import logging
import shutil
from dataclasses import dataclass, replace
from datetime import datetime
from pathlib import Path

import pandas as pd
import xlwings as xw

from po_generator.hs_code import HS_COLUMN, is_non_goods
from po_generator.utils import get_value, normalize_line_item, to_text
from po_generator.excel_helpers import (
    MIN_ITEM_ROW_HEIGHT,
    cleanup_temp_file,
    insert_copied_rows,
    layout_address_rows,
    prepare_template,
    xlwings_app_context,
)

logger = logging.getLogger(__name__)


# === 셀 매핑 (Certificate of Origin) — 값 칸은 C:F 병합의 좌상단 ===
CELL_CERTIFICATE_NO = 'C7'
CELL_ISSUE_DATE = 'C8'

# Consignee 3줄 (C15:C17) — Bill to 1/2/3 (CI의 'Consigned to'와 같은 출처, 해외 DN 전부 3줄)
CONSIGNEE_START_ROW = 15
CONSIGNEE_ROW_COUNT = 3

# 참조 정보 5줄 (C19:C23) — 순서가 곧 행 순서다
REFERENCE_START_ROW = 19  # Commercial Invoice No. / Order No. / Customer Order No. /
                          # Means of Transport / Country of Destination

VALUE_COLS = 'CDEF'

# 표
HEADER_ROW = 26
ITEM_START_ROW = 27
LAST_COL = 'F'
ITEM_HEADERS = ('No.', 'Customer PO', 'Item No.', 'Model No. / Description', 'HS Code', 'Qty')
TOTAL_LABEL_COL = 'D'
QTY_COL = 'F'

# 템플릿 마지막 행(서명 이미지 자리의 끝) — 품목이 늘면 그만큼 아래로 밀린다.
# Total부터 여기까지가 '맺음 블록'이라 페이지가 갈리면 안 된다 (`_keep_closing_block`).
TEMPLATE_LAST_ROW = 40

# 선적항 — CI의 From(B14)과 같은 값
PORT_OF_LOADING = 'INCHEON, KOREA'


@dataclass(frozen=True)
class CooLine:
    """표 한 줄 (A~F 순서대로 `cells()`)

    `order_no`(SO_ID)·`line`(SO Line item)은 표에 찍지 않는다 — 정렬과 헤더 Order No.에만 쓴다.
    `no`는 정렬이 끝난 뒤 1부터 매긴다 (`build_coo_lines`).
    """

    order_no: str
    line: str
    item_no: str
    description: str
    hs_code: str
    qty: int | float
    customer_po: str = ''
    shipping_method: str = ''
    no: int = 0

    def cells(self) -> list:
        return [str(self.no), self.customer_po, self.item_no, self.description, self.hs_code, self.qty]


@dataclass(frozen=True)
class CooHeader:
    certificate_no: str
    issue_date: str
    consignee: tuple[str, ...]
    invoice_no: str
    order_nos: str
    customer_pos: str
    transport: str
    destination: str

    def references(self) -> list[str]:
        """C19:C23에 위에서부터 쓰는 값"""
        return [self.invoice_no, self.order_nos, self.customer_pos,
                self.transport, self.destination]


def certificate_no(dn_id: str) -> str:
    """증명서 번호 — DN 하나에 COO 하나라 DN_ID에서 바로 만든다 (CI Invoice No.와 짝)"""
    return f"COO-{dn_id}"


def _unique_in_order(values) -> list[str]:
    seen: list[str] = []
    for value in values:
        text = to_text(value).strip()
        if text and text not in seen:
            seen.append(text)
    return seen


def _parse_qty(raw: object, label: str) -> int | float:
    """수량 — 증명서에 추정치를 찍지 않는다 (CI처럼 1로 메우지 않고 멈춘다)"""
    try:
        qty = float(raw)
    except (TypeError, ValueError):
        raise ValueError(f"수량을 읽을 수 없습니다: {label} = {raw!r}") from None
    if pd.isna(qty):
        raise ValueError(f"수량이 비어 있습니다: {label}")
    return int(qty) if qty.is_integer() else qty


def _line_sort_key(line: CooLine) -> tuple:
    """Customer PO → SO → SO 라인 — 업체가 PO별로 묶어 읽고, 한 PO 안에서는 우리 입력 순서"""
    number = int(line.line) if line.line.isdigit() else float('inf')
    return (line.customer_po, line.order_no, number, line.line)


def build_coo_lines(items_df: pd.DataFrame) -> list[CooLine]:
    """DN 아이템 → 표 줄 (운임 제외, Customer PO → SO → Line 순, 1부터 순번)

    Args:
        items_df: DN 아이템 (Model number 보강 + `_hs_code` 열이 붙은 상태)

    Returns:
        CooLine 목록 — 비어 있으면 실을 물품이 없는 DN이다
    """
    lines: list[CooLine] = []
    for _, item in items_df.iterrows():
        description = to_text(get_value(item, 'item_name', ''))
        if is_non_goods(description):
            continue
        order_no = to_text(get_value(item, 'so_id', ''))
        line = normalize_line_item(item.get('Line item'))
        lines.append(CooLine(
            order_no=order_no,
            line=line,
            item_no=to_text(get_value(item, 'model', '')),
            description=description,
            hs_code=to_text(item.get(HS_COLUMN, '')),
            qty=_parse_qty(get_value(item, 'item_qty', None), f"{order_no} L{line}"),
            customer_po=to_text(get_value(item, 'customer_po', '')),
            shipping_method=to_text(get_value(item, 'shipping_method', '')),
        ))
    ordered = sorted(lines, key=_line_sort_key)
    return [replace(line, no=i) for i, line in enumerate(ordered, start=1)]


def _date_text(value: object) -> str:
    """CI Invoice Date와 같은 규칙 — 출고일, 없으면 오늘 (같은 봉투의 두 장이 같은 날짜)"""
    if value is None or value == '' or pd.isna(value):
        return datetime.now().strftime("%Y-%m-%d")
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    return str(value)


def build_coo_header(order_data: pd.Series, lines: list[CooLine]) -> CooHeader:
    """헤더 값 — 번호 목록(Order/Customer PO/운송)은 표에 실린 줄에서 모은다

    Args:
        order_data: DN 첫 행 (Bill to·출고일·국가 등 DN 단위 값)
        lines: `build_coo_lines()` 결과
    """
    consignee = _unique_in_order(
        get_value(order_data, key, '') for key in ('bill_to_1', 'bill_to_2', 'bill_to_3')
    )
    if not consignee:
        # Customer_해외 조인이 비면 이름이라도 — 수하인이 빈 증명서는 내보내지 않는다
        consignee = _unique_in_order([get_value(order_data, 'customer_name', '')])

    methods = _unique_in_order(line.shipping_method for line in lines)
    transport = ' / '.join(methods)
    transport = f"{transport}, from {PORT_OF_LOADING}" if transport else f"From {PORT_OF_LOADING}"

    destination = to_text(get_value(order_data, 'customer_country', '')) \
        or to_text(get_value(order_data, 'bill_to_3', ''))

    dn_id = to_text(get_value(order_data, 'dn_id', ''))
    return CooHeader(
        certificate_no=certificate_no(dn_id),
        issue_date=_date_text(get_value(order_data, 'dispatch_date', None)),
        consignee=tuple(consignee[:CONSIGNEE_ROW_COUNT]),
        invoice_no=dn_id,
        order_nos=', '.join(_unique_in_order(line.order_no for line in lines)),
        customer_pos=', '.join(_unique_in_order(line.customer_po for line in lines)),
        transport=transport,
        destination=destination,
    )


def create_coo_xlwings(
    template_path: Path,
    output_path: Path,
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> None:
    """xlwings로 Certificate of Origin 생성

    Args:
        template_path: 템플릿 파일 경로
        output_path: 출력 파일 경로
        order_data: DN 첫 행
        items_df: DN 전체 아이템 (`_hs_code` 열 포함)

    Raises:
        ValueError: 실을 물품이 없거나(운임만) 수량을 못 읽을 때, 템플릿이 다를 때
    """
    if items_df is None:
        items_df = pd.DataFrame([order_data])

    # 값부터 다 정한다 — 데이터 문제로 멈출 거면 Excel을 띄우기 전에 멈춘다
    lines = build_coo_lines(items_df)
    if not lines:
        raise ValueError("원산지증명서에 실을 물품 줄이 없습니다 (운임 줄만 있는 DN)")
    header = build_coo_header(order_data, lines)

    temp_template, temp_output = prepare_template(template_path, "coo")

    try:
        # Book1 없이 띄운다 — 맺음 블록 유지가 읽는 페이지 나누기가 실제 PDF와 같아지려면
        # (`xlwings_app_context`의 add_book 설명)
        with xlwings_app_context(add_book=False) as app:
            wb = app.books.open(str(temp_template))
            ws = wb.sheets[0]

            assert_coo_template(ws)
            _fill_header(ws, header)
            _fill_items(ws, lines)

            wb.save(str(temp_output))
            logger.info(f"Certificate of Origin 생성 완료 (임시): {temp_output}")

    finally:
        cleanup_temp_file(temp_template)

    shutil.move(str(temp_output), str(output_path))
    logger.info(f"Certificate of Origin 저장 완료: {output_path}")


def assert_coo_template(ws: xw.Sheet) -> None:
    """템플릿 표 모양이 이 생성기의 열 매핑과 같은지 (값 쓰기 전, COM 읽기 2회)

    양식은 템플릿 파일이 소유한다 — 누가 열을 옮기거나 행을 끼우면 값이 엉뚱한 칸에
    조용히 들어간다. 머리글 행과 Total 위치로 막는다.

    Raises:
        ValueError: 머리글이나 Total 위치가 다를 때
    """
    actual = tuple(to_text(v).strip() for v in ws.range(
        f'A{HEADER_ROW}:{LAST_COL}{HEADER_ROW}'
    ).value)
    if actual != ITEM_HEADERS:
        raise ValueError(
            f"COO 템플릿 표 머리글이 생성기와 다릅니다 (행 {HEADER_ROW}): "
            f"{actual} != {ITEM_HEADERS}"
        )
    total_cell = f'{TOTAL_LABEL_COL}{ITEM_START_ROW + 1}'
    if to_text(ws.range(total_cell).value).strip() != 'Total':
        raise ValueError(f"COO 템플릿의 Total 행 위치가 다릅니다 ({total_cell})")


def _fill_header(ws: xw.Sheet, header: CooHeader) -> None:
    """헤더 값 쓰기 + 긴 값이 병합 경계에서 잘리지 않게 행 높이 확보"""
    ws.range(CELL_CERTIFICATE_NO).value = header.certificate_no
    ws.range(CELL_ISSUE_DATE).value = header.issue_date

    consignee = list(header.consignee) + [''] * (CONSIGNEE_ROW_COUNT - len(header.consignee))
    references = header.references()

    # 병합 칸(C:F)마다 좌상단(C)에만 쓴다 — 열 하나 세로 배치라 행별 쓰기가 같은 범위다
    for start, values in ((CONSIGNEE_START_ROW, consignee), (REFERENCE_START_ROW, references)):
        for offset, value in enumerate(values):
            ws.range(f'{VALUE_COLS[0]}{start + offset}').value = value
        # 병합 칸은 넘친 글자를 옆으로 흘리지 않고 경계에서 자른다 (OC·FI와 같은 문제)
        layout_address_rows(ws, start, values, '', left_cols=VALUE_COLS, right_cols=VALUE_COLS)


def _fill_items(ws: xw.Sheet, lines: list[CooLine]) -> None:
    """표 채우기 — 행 삽입(타일 복사) → 값 한 번에 쓰기 → 높이 자동 맞춤 → 합계"""
    num_items = len(lines)
    end_row = ITEM_START_ROW + num_items - 1

    # 템플릿은 품목 1행이다. 그 행을 복사해 늘린다 (격자·서식·텍스트 형식이 따라온다)
    insert_copied_rows(ws, ITEM_START_ROW + 1, num_items - 1, source_row=ITEM_START_ROW)

    ws.range(f'A{ITEM_START_ROW}:{LAST_COL}{end_row}').value = [line.cells() for line in lines]

    _fit_item_rows(ws, ITEM_START_ROW, end_row)

    total_row = end_row + 1
    ws.range(f'{QTY_COL}{total_row}').formula = (
        f'=SUM({QTY_COL}{ITEM_START_ROW}:{QTY_COL}{end_row})'
    )
    logger.debug(f"COO 아이템 {num_items}줄 (Row {ITEM_START_ROW}~{end_row}), Total Row {total_row}")

    _keep_closing_block(ws, end_row, TEMPLATE_LAST_ROW + num_items - 1)


def _fit_item_rows(ws: xw.Sheet, start_row: int, end_row: int) -> None:
    """품목 행 높이 = 자동 맞춤, 단 한 줄짜리는 최소 높이까지

    품목 행엔 병합이 없어 자동 맞춤이 먹는다 (CI·PL의 A:C 병합과 다르다 — 모듈
    docstring). 다만 한 줄은 Arial 10 기본 높이(12.75pt)로 붙어 격자가 답답해서,
    다른 문서와 같은 `MIN_ITEM_ROW_HEIGHT`까지는 띄운다.
    """
    rows = ws.range(f'{start_row}:{end_row}')
    rows.rows.autofit()

    # 다중 행 RowHeight는 전부 같으면 그 값, 다르면 None — 짧은 품목명뿐이면 COM 1회
    uniform = rows.api.RowHeight
    if isinstance(uniform, (int, float)):
        heights = [float(uniform)] * (end_row - start_row + 1)
    else:
        heights = [ws.range(f'{r}:{r}').api.RowHeight for r in range(start_row, end_row + 1)]

    # 최소 높이에 못 미치는 연속 구간을 묶어 한 번씩 쓴다
    run_start = None
    for offset, height in enumerate(heights + [MIN_ITEM_ROW_HEIGHT]):
        short = height < MIN_ITEM_ROW_HEIGHT
        if short and run_start is None:
            run_start = offset
        elif not short and run_start is not None:
            ws.range(f'{start_row + run_start}:{start_row + offset - 1}').api.RowHeight = \
                MIN_ITEM_ROW_HEIGHT
            run_start = None


def closing_block_split(breaks: list[int], last_item_row: int, block_end_row: int) -> bool:
    """맺음 블록(Total ~ 서명)이 페이지 경계에 걸리는가 (순수 함수)

    Args:
        breaks: Excel 페이지 나누기 행 번호들 (그 행이 새 페이지의 첫 행)
        last_item_row: 마지막 품목 행
        block_end_row: 맺음 블록 마지막 행 (서명 이미지 자리의 끝)

    Returns:
        새 페이지가 Total 행부터 서명란 사이에서 시작하면 True — Total부터 시작하는
        경우도 포함한다(그 페이지엔 품목 없이 머리글·합계만 남는다)
    """
    return any(last_item_row < b <= block_end_row for b in breaks)


def _keep_closing_block(ws: xw.Sheet, last_item_row: int, block_end_row: int) -> None:
    """서명 블록이 다음 장으로 잘려 나가지 않게 — 마지막 품목 행과 함께 넘긴다

    실측(DNO-2026-0135, 7줄): Name/Title만 2쪽으로 넘어가고, 그 쪽엔 반복 머리글만 덩그러니
    찍혔다. Excel엔 '다음과 함께 유지'가 없어서, 자동 나누기가 맺음 블록에 걸리면 **마지막
    품목 행 앞**에 수동 나누기를 넣는다 — 새 쪽은 머리글 + 품목 1줄 이상 + 합계 + 선언문 +
    서명이 된다. 템플릿은 '폭만 한 장(높이 자동)'이라 수동 나누기가 지켜진다 — 높이까지
    장수로 고정하면 Excel이 수동 나누기를 무시한다 (2026-10-14 실측: 폭 맞춤 + 수동 나누기 → PDF 반영).
    """
    breaks = read_page_breaks(ws)
    if closing_block_split(breaks, last_item_row, block_end_row):
        # 품목이 한 줄뿐이면 표 머리글째 넘긴다 — 앞 장 끝에 빈 머리글만 남지 않게
        before = last_item_row if last_item_row > ITEM_START_ROW else HEADER_ROW
        ws.api.HPageBreaks.Add(ws.range(f'A{before}').api)
        logger.debug(f"맺음 블록 유지: 자동 나누기 {breaks} → Row {before} 앞 수동 나누기")


def read_page_breaks(ws: xw.Sheet, max_reads: int = 5) -> list[int]:
    """가로 페이지 나누기 행 번호 — **같은 결과가 두 번 연속 나올 때까지** 다시 읽는다

    Excel은 페이지 나누기를 접근할 때 조금씩 계산한다. 첫 열거는 부분 결과다
    (실측: 5줄 DN이 PDF 2쪽인데 첫 읽기 [] → 둘째 [47], 33줄 DN은 [43] → [43, 73]).
    첫 값을 믿으면 서명 블록이 잘리는 걸 못 본다 — 실제로 그렇게 12건 중 6건이 잘려 나갔다.
    """
    previous: list[int] | None = None
    for _ in range(max_reads):
        page_breaks = ws.api.HPageBreaks
        current = [page_breaks.Item(i).Location.Row for i in range(1, page_breaks.Count + 1)]
        if current == previous:
            return current
        previous = current
    logger.warning(f"페이지 나누기가 {max_reads}번 읽는 동안 안정되지 않음: {previous}")
    return previous or []
