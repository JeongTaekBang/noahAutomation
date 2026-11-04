#!/usr/bin/env python
"""
AX Sales order book — IR + Item transaction으로 만드는 오더북
===========================================================

AX export 두 개(Item requirement, Item transaction)로 회계팀 오더북 Raw 형식의 오더북을 만든다.
AX Sales Order Book(Table 50229)처럼 **달마다 계산해 저장**하고, 다음 달은 저장된 End에서 시작한다
(계산 규칙은 po_generator/ax_orderbook.py, 저장은 po_generator/ax_orderbook_store.py → ax_orderbook/ax_orderbook.db).

    저장소가 비어 있으면   기준선 — P01~마감월을 IR·IT 한 벌로 한 번에 (Start(P01) = 회계 P01 Start)
    저장된 다음 달         저장된 직전 달 End·청구 누계에서 시작 — 그 사이의 청구·취소·변경은 이 달로, 지난 달은 그대로
    저장된 마지막 달(Open)  다시 돌리면 다시 계산해 바꿔 넣는다(AX Last calculated). 그 앞 달(Closed)은 거부
기준선부터 다시 만들려면 저장소 파일을 지운다(또는 --store로 다른 파일).
회계 오더북은 기준선의 P01 Start와 키 매핑·속성에 쓰고, 같은 달 행이 있으면 대조해 차이와 원인을 붙인다.
결과는 그 달 폴더에 엑셀과 같은 이름의 HTML 두 개 — HTML은 Edge에서 더블클릭해 여는 인터랙티브 보고서
(필터·차트·드릴다운, 인터넷 없이 열린다 — po_generator/ax_orderbook_report.py).

폴더 — 달마다 ax_orderbook/연도/P월/ (예: ax_orderbook/2026/P11/)에 그 달 입력을 넣고 결과도 거기 쓴다.
여러 달이 같이 쓰는 것(ax_orderbook.db, fx_history.csv, 통화지정.csv)은 ax_orderbook/ 바로 아래.

준비물 (그 달 폴더, 없으면 ax_orderbook/ 바로 아래 — **마감 직후(다음 달 1~7일) 같은 날** 한 쌍):
    Item requirement_YYYYMMDD.xlsx   AX Item requirement — Show=All, Created date ≥ 01/01/2024 하나만
    Item transaction_YYYYMMDD.xlsx   AX Item transaction(프로젝트 품목 거래) — Project date ≥ 01/01/2024 하나만
    Sales order book_MM_YYYY.xlsx    회계 오더북 (ax_orderbook/ 아래 어디든 가장 새 것 — 기준선은 P01 행 필수, 그 뒤엔
                                     키 매핑·속성용이라 지난 해 파일도 된다)
    환율: NOAH_SO_PO_DN FX 시트(sync_db.py로 DB에) + Finance Hub 마감환율 파일(그 달 폴더에 둔 사본도 읽는다 —
          fx_history.csv에 쌓는다)

사용법:
    python build_ax_orderbook.py P10                  # 저장소가 비었으면 P01~P10 기준선, 아니면 P10 한 달
    python build_ax_orderbook.py P11 --dry-run        # 계산하고 결과 파일(_미리보기)만 — 저장하지 않는다
    python build_ax_orderbook.py P10 --year 2026
    python build_ax_orderbook.py P10 --fx JPY=8.6157  # 마감월 환율을 직접 넣기 (CUR=YYYY-MM=RATE도 됨)
    python build_ax_orderbook.py P10 --ir 파일 --it 파일 --orderbook 파일 --store 저장소.db
"""

from __future__ import annotations

import argparse
import getpass
import logging
import re
import shutil
import sys
import warnings
from dataclasses import replace
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

from po_generator.ax_orderbook import (
    DIFF_CAUSES, RAW_COLUMNS, OrderbookResult, build_month, build_orderbook, classify_differences,
    compare_with_official, empty_official, keys_from_raw, load_fx_book, month_bounds, parse_period,
    read_currency_overrides, read_ir, read_it, read_raw,
)
from po_generator.ax_orderbook_report import report_data, write_report
from po_generator.ax_orderbook_store import BASELINE, MONTHLY, OrderbookStore, Plan
from po_generator.config import AX_ORDERBOOK_DB, AX_ORDERBOOK_DIR, DB_FILE, FINANCE_HUB_DIR
from po_generator.logging_config import setup_logging

logger = logging.getLogger(__name__)

# 여러 달이 같이 쓰는 파일 — ax_orderbook/ 바로 아래에 둔다
FX_HISTORY_NAME = 'fx_history.csv'
# 외화 거래처 목록에 없는 외화 주문 — `코드,통화` (거래처 C번호 또는 프로젝트 번호). 확인 목록 '통화 확인'을 보고 추가
CURRENCY_OVERRIDE_NAME = '통화지정.csv'
_ORDERBOOK_RE = re.compile(r'Sales order book_(\d{1,2})_(\d{4})', re.IGNORECASE)
_DATED_RE = re.compile(r'_(\d{8})$')
# 다음 달 이 날짜까지 뽑은 추출본을 그 달 마감 자료로 본다 — 더 늦으면 다음 달 청구·취소가 이 달로 섞인다
SNAPSHOT_GRACE_DAYS = 7
PERIOD_SHEET_COLUMNS: dict[str, str] = {
    'Period': 'Period', 'status': '상태', 'basis': '방식', 'calculated_at': '계산 시각', 'calculated_by': '계산한 사람',
    'ir_file': 'IR 파일', 'it_file': 'IT 파일', 'snapshot': '추출일', 'row_count': 'Raw 행', 'billed_rows': '청구 누계 행',
}


def default_year(month: int, today: date | None = None) -> int:
    """마감월이 이번 달보다 뒤면 작년 (1월에 P12를 돌리는 경우)."""
    today = today or date.today()
    return today.year - 1 if month > today.month else today.year


def parse_fx_overrides(items: list[str], close_ym: str) -> dict[tuple[str, str], float]:
    """'JPY=8.6157' (마감월) 또는 'JPY=2026-10=8.6157'"""
    out: dict[tuple[str, str], float] = {}
    for item in items or []:
        parts = item.split('=')
        if len(parts) == 2:
            cur, month, rate = parts[0], close_ym, parts[1]
        elif len(parts) == 3:
            cur, month, rate = parts
        else:
            raise ValueError(f"--fx 형식: JPY=8.6157 또는 JPY=2026-10=8.6157 (받은 값: {item})")
        out[(cur.strip().upper(), month.strip())] = float(rate)
    return out


def _latest_dated(folder: Path, prefix: str) -> list[Path]:
    """`{prefix}_YYYYMMDD.xlsx` 중 가장 최근 추출본 하나 (파일 이름의 날짜, 없으면 수정 시각)."""
    files = [p for p in folder.glob(f'{prefix}_*.xlsx') if not p.name.startswith('~$')]
    if not files:
        return []
    def stamp(p: Path):
        m = _DATED_RE.search(p.stem)
        return (m.group(1) if m else '', p.stat().st_mtime)
    return [max(files, key=stamp)]


def period_dir(root: Path, year: int, month: int) -> Path:
    """그 달 폴더 — 입력을 넣고 결과를 쓰는 곳 (ax_orderbook/2026/P11)"""
    return root / str(year) / f"P{month:02d}"


def find_ir_files(root: Path, year: int, month: int) -> list[Path]:
    """그 달 폴더의 최신 추출본, 없으면 ax_orderbook 바로 아래 — 다른 달 폴더는 보지 않는다(그 달 추출본이어야 한다)"""
    return _latest_dated(period_dir(root, year, month), 'Item requirement') or _latest_dated(root, 'Item requirement')


def find_it_files(root: Path, year: int, month: int) -> list[Path]:
    return _latest_dated(period_dir(root, year, month), 'Item transaction') or _latest_dated(root, 'Item transaction')


def _outside_old(root: Path, p: Path) -> bool:
    return 'Old' not in p.relative_to(root).parts


def find_orderbooks(root: Path) -> list[Path]:
    """회계 오더북 — ax_orderbook 아래 어디든(달 폴더 포함, Old 제외). 파일 이름의 (연, 월)이 새것부터, 같으면 최근 수정부터"""
    files = [p for p in root.rglob('Sales order book_*.xlsx') if not p.name.startswith('~$') and _outside_old(root, p)]
    def stamp(p: Path):
        m = _ORDERBOOK_RE.search(p.stem)
        return ((int(m.group(2)), int(m.group(1))) if m else (0, 0), p.stat().st_mtime)
    return sorted(files, key=stamp, reverse=True)


def shared_file(root: Path, name: str) -> tuple[Path, Path | None]:
    """여러 달이 같이 쓰는 파일의 자리(ax_orderbook 바로 아래)와, 거기 없을 때 달 폴더에 들어간 가장 최근 사본"""
    home = root / name
    if home.exists():
        return home, None
    copies = [p for p in root.glob(f'*/P*/{name}') if _outside_old(root, p)]
    return home, max(copies, key=lambda p: p.stat().st_mtime) if copies else None


def load_official(paths: list[Path], year: int, upto: int) -> tuple[pd.DataFrame, Path]:
    """기준선용 — P01 행이 있는 가장 새 회계 오더북 → (P01~upto 행, 파일). P01 Start가 기초값이라 P01 행이 없으면 못 쓴다."""
    for path in paths:
        m = _ORDERBOOK_RE.search(path.stem)
        if m and int(m.group(2)) != year:
            continue
        raw = read_raw(path)
        if (raw['pn'] == 1).any():
            return raw[raw['pn'] <= upto].copy(), path
    raise FileNotFoundError(f"{year}년 P01 행이 있는 회계 오더북(Sales order book_*.xlsx)을 못 찾았습니다")


def load_reference(paths: list[Path], year: int, upto: int) -> tuple[pd.DataFrame, Path | None]:
    """다음 달 계산용 — 키 매핑·속성만 쓰므로 그 해 가장 새 파일, 없으면 지난 해 가장 새 파일. 없으면 빈 표."""
    for path in paths:                       # find_orderbooks가 (연, 월) 내림차순으로 준다
        m = _ORDERBOOK_RE.search(path.stem)
        if m and int(m.group(2)) > year:
            continue
        raw = read_raw(path)
        same_year = pd.to_numeric(raw['year'], errors='coerce').fillna(year).eq(year)
        return raw[~same_year | (raw['pn'] <= upto)].copy(), path
    return empty_official(), None


def official_of_year(official: pd.DataFrame, year: int) -> pd.DataFrame:
    """대조용 — 그 해 회계 행만 (지난 해 파일을 매핑용으로 읽었으면 대조할 게 없다)"""
    if 'year' not in official.columns:
        return official
    return official[pd.to_numeric(official['year'], errors='coerce').fillna(year).eq(year)]


def snapshot_date(paths: list[Path]) -> date | None:
    """파일 이름의 추출일(`_YYYYMMDD`) — 여러 개면 가장 늦은 날, 하나라도 날짜가 없으면 None"""
    days = []
    for p in paths:
        m = _DATED_RE.search(p.stem)
        try:
            days.append(datetime.strptime(m.group(1), '%Y%m%d').date())
        except (AttributeError, ValueError):
            return None
    return max(days) if days else None


def check_snapshot(ir_paths: list[Path], it_paths: list[Path], year: int, month: int
                   ) -> tuple[date | None, list[str], list[str]]:
    """추출 시점 점검 → (추출일, 오류, 경고). 오류가 있으면 저장하지 않는다(--dry-run은 경고로 보고 계산만).

    IR과 IT는 같은 날 뽑은 한 쌍이어야 하고(날이 다르면 출고와 수주 라인이 어긋난다), 그 달이 끝난 뒤 뽑은 것이어야
    한다(그 전이면 그 달 출고·청구가 덜 들어 있다). 저장한 달은 다시 못 고치니 저장할 때만 막는다."""
    errors: list[str] = []
    warns: list[str] = []
    ir_day, it_day = snapshot_date(ir_paths), snapshot_date(it_paths)
    if ir_day is None or it_day is None:
        warns.append("파일 이름에 추출일(_YYYYMMDD)이 없어 추출 시점을 점검하지 못했습니다")
        return ir_day or it_day, errors, warns
    if ir_day != it_day:
        errors.append(f"IR({ir_day:%Y-%m-%d})과 IT({it_day:%Y-%m-%d})의 추출일이 다릅니다 — 같은 날 뽑은 한 쌍이어야 "
                      "출고와 수주 라인이 맞물립니다")
    close_day = month_bounds(year, month)[1].date()        # 다음 달 1일
    day = min(ir_day, it_day)
    if day < close_day:
        errors.append(f"추출일 {day:%Y-%m-%d}은 {year} P{month:02d}이 끝나기 전입니다 — {close_day:%Y-%m-%d} 이후에 "
                      "뽑은 IR·IT로 계산하세요")
    elif day >= close_day + timedelta(days=SNAPSHOT_GRACE_DAYS):
        warns.append(f"추출일 {day:%Y-%m-%d}은 마감 뒤 {(day - close_day).days + 1}일째입니다 — 그 사이 청구된 이 달 "
                     "출고는 이 달 Output으로, 그 사이의 취소·변경은 이 달로 섞입니다(AX는 다음 달)")
    return day, errors, warns


def describe_plan(plan: Plan, year: int, upto: int, preview: bool = False) -> str:
    """[1/4] 줄 — 이번 실행이 무엇을 하고 Start를 어디서 가져오는지. 기간 뒤에 조사를 붙이지 않는다(P10을/P12를이 갈린다)"""
    this = f"{year} P{upto:02d}"
    if plan.kind == 'baseline':
        return (f"기준선 — 저장소가 비어 있어 {year} P01~P{upto:02d} 전체를 한 번에 계산 (P01 Start = 회계 P01 Start) · "
                + ('미리보기: 저장 안 함' if preview else f"저장 뒤 마지막 달 P{upto:02d}만 Open"))
    prev = f"{plan.prev[0]} P{plan.prev[1]:02d}"
    if plan.kind == 'next':
        return (f"새 달 {this} 계산 — Start = 저장된 {prev} End · "
                + ('미리보기: 저장 안 함' if preview else f"저장하면 {prev} 닫힘(Closed)"))
    return (f"마지막 달 {this} 다시 계산 — Start = 저장된 {prev} End ({prev}까지는 그대로) · "
            + ('미리보기: 바꿔 넣지 않음' if preview else f"{this} 자리에 바꿔 넣음"))


def method_note(plan: Plan, upto: int) -> str:
    """요약 시트의 계산 방식 설명"""
    common = 'Output = 청구된 출고, End = 미출고 + 청구 전 출고, Variance = 차액.'
    if plan.kind == 'baseline':
        return (f"기준선: P01~P{upto:02d}을 같은 IR·IT로 한 번에 계산해 저장했다(Start(P01) = 회계 P01 Start). "
                f"다음 달부터는 저장된 직전 End에서 시작한다. {common}")
    return (f"월 계산: 저장된 직전 달 End·청구 누계에서 시작 — 그 사이 청구된 것은 이 달 Output, 취소·단가/수량 변경·환차는 "
            f"이 달 Variance. P{upto:02d} 앞 달은 저장된 값 그대로다. {common}")


def period_table(store: OrderbookStore, year: int, months: list[int], basis: str, files: dict[str, str],
                 snapshot: str, res: OrderbookResult, saved: bool) -> pd.DataFrame:
    """기간 시트 — 저장소의 기간 기록(AX의 Report 묶음). 미리보기면 이번 달을 '미리보기'로 덧붙인다."""
    p = store.periods(year)
    if not saved:
        p = p[~p['pn'].isin(months)]
        rows = res.raw_rows['pn'].value_counts()
        billed = res.billed['pn'].value_counts() if len(res.billed) else pd.Series(dtype=int)
        preview = pd.DataFrame([{
            'year': year, 'pn': m, 'status': '미리보기(저장 안 함)', 'basis': basis,
            'calculated_at': f"{datetime.now():%Y-%m-%d %H:%M:%S}", 'calculated_by': getpass.getuser(),
            'ir_file': files['Item requirement'], 'it_file': files['Item transaction'], 'snapshot': snapshot,
            'row_count': int(rows.get(m, 0)), 'billed_rows': int(billed.get(m, 0))} for m in months])
        p = preview if p.empty else pd.concat([p, preview], ignore_index=True)
    p = p.assign(Period=p['pn'].map(lambda m: f"P{int(m):02d}"))
    return p[list(PERIOD_SHEET_COLUMNS)].rename(columns=PERIOD_SHEET_COLUMNS)


# ---------------------------------------------------------------------------
# 결과 파일
# ---------------------------------------------------------------------------
def month_table(res: OrderbookResult, official: pd.DataFrame) -> pd.DataFrame:
    """달별 우리 합계 vs 회계 합계 + 프로젝트 일치율"""
    rows = []
    for m in range(1, res.upto + 1):
        k = res.keys[res.keys['pn'] == m]
        row = {'Period': f"P{m:02d}"}
        for label, col in (('Start', 'Sv'), ('Input', 'Iv'), ('Output', 'Ov'), ('Variance', 'Vv'), ('End', 'Ev')):
            row[f'우리 {label}'] = k[col].sum()
        om = official[official['pn'] == m]
        if len(om):
            c = compare_with_official(k, om)
            t = c['totals'].set_index('항목')
            for label in ('Start', 'Input', 'Output', 'Variance', 'End'):
                row[f'회계 {label}'] = t.loc[label, '회계']
            for label in ('Input', 'Output', 'End'):
                row[f'차이 {label}'] = t.loc[label, '차이']
            row['프로젝트 일치% Output'] = c['project_match']['Output']
            row['프로젝트 일치% End'] = c['project_match']['End']
        rows.append(row)
    return pd.DataFrame(rows)


def cause_table(diff: pd.DataFrame) -> pd.DataFrame:
    """달 × 원인 — End 차이 합계와 키 수"""
    if diff.empty:
        return pd.DataFrame()
    t = diff.pivot_table(index='원인', columns='Period', values='d_End', aggfunc='sum').fillna(0.0)
    t = t.reindex([c for c in DIFF_CAUSES if c in t.index])
    t.insert(0, '키 수', diff.groupby('원인').size().reindex(t.index))
    return t.reset_index()


def _summary_rows(res: OrderbookResult, months: pd.DataFrame, causes: pd.DataFrame, files: dict[str, str],
                  note: str) -> list[list]:
    rows: list[list] = [
        [f"AX 오더북 (IR + Item transaction) — {res.year}년 P01~{res.period}"],
        [f"만든 시각: {datetime.now():%Y-%m-%d %H:%M}"],
        *[[f"{label}: {name}"] for label, name in files.items()],
        [note],
        [res.next_filter_note],
        [],
        list(months.columns),
        *months.values.tolist(),
    ]
    if len(causes):
        rows += [[], ['회계와의 End 차이 — 원인별 (원)'] + [''] * (len(causes.columns) - 1), list(causes.columns)]
        rows += causes.values.tolist()
    rows += [[], ['확인 목록', '건수', '금액']]
    for kind, g in res.review.groupby('구분', sort=False):
        rows.append([kind, len(g), g['금액'].sum()])
    if len(res.pending):
        rows += [[], ['출고 후 미청구(마감월 말)', len(res.pending), res.pending['금액(원)'].sum()]]
    return rows


def write_workbook(res: OrderbookResult, out_path: Path, months: pd.DataFrame, diff: pd.DataFrame,
                   causes: pd.DataFrame, files: dict[str, str], periods: pd.DataFrame, note: str) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    sheets: list[tuple[str, pd.DataFrame]] = [('기간', periods), ('Raw', res.raw_rows[list(RAW_COLUMNS)]),
                                              ('차이_월별', months)]
    if len(causes):
        sheets.append(('차이_원인', causes))
    if len(diff):
        sheets.append(('차이_키별', diff))
    sheets += [('출고후미청구', res.pending), ('확인목록', res.review), (f'라인내역_{res.period}', res.lines),
               ('환율', res.fx_used)]

    target = out_path
    try:
        writer = pd.ExcelWriter(target, engine='openpyxl')
    except PermissionError:
        target = out_path.with_name(f"{out_path.stem}_{datetime.now():%H%M%S}{out_path.suffix}")
        writer = pd.ExcelWriter(target, engine='openpyxl')
    with writer:
        pd.DataFrame(_summary_rows(res, months, causes, files, note)).to_excel(
            writer, sheet_name='요약', index=False, header=False)
        for name, df in sheets:
            df.to_excel(writer, sheet_name=name, index=False)
        for ws in writer.book.worksheets:
            _format_sheet(ws, header=ws.title != '요약')
    return target


def _format_sheet(ws, header: bool = True) -> None:
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    if header and ws.max_row > 1:
        ws.freeze_panes = 'A2'
        for cell in ws[1]:
            cell.font = Font(bold=True)
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        width = 10
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 200), min_col=col, max_col=col):
            v = row[0].value
            if isinstance(v, float):
                row[0].number_format = '#,##0.##' if abs(v) < 1000 else '#,##0'
            elif isinstance(v, (datetime, pd.Timestamp)):
                row[0].number_format = 'yyyy-mm-dd'
            width = max(width, min(len(str(v)) + 2 if v is not None else 0, 60))
        for row in ws.iter_rows(min_row=201, min_col=col, max_col=col):
            v = row[0].value
            if isinstance(v, float):
                row[0].number_format = '#,##0.##' if abs(v) < 1000 else '#,##0'
            elif isinstance(v, (datetime, pd.Timestamp)):
                row[0].number_format = 'yyyy-mm-dd'
        ws.column_dimensions[letter].width = width


def write_html(res: OrderbookResult, workbook: Path, diff: pd.DataFrame, official: pd.DataFrame,
               periods: pd.DataFrame, files: dict[str, str], note: str, preview: bool) -> Path | None:
    """엑셀 옆에 같은 이름의 인터랙티브 HTML — Edge에서 더블클릭으로 연다. HTML은 덤이라 실패해도 엑셀 결과는 그대로 두고
    경고만 한다(예: Plotly가 없는 PC)."""
    try:
        return write_report(report_data(res, diff, official, periods, files, note, preview=preview),
                            workbook.with_suffix('.html'))
    except Exception as exc:          # noqa: BLE001 — 어떤 실패든 엑셀 결과를 막지 않는다
        logger.debug("결과 HTML 실패", exc_info=True)
        print(f"[경고] 결과 HTML을 못 만들었습니다 (엑셀은 그대로): {exc}")
        return None


def print_summary(res: OrderbookResult, out_path: Path, months: pd.DataFrame, causes: pd.DataFrame,
                  saved_note: str, html_path: Path | None = None) -> None:
    print(f"\n[{res.year} P01~{res.period}] Raw 행 {len(res.raw_rows):,}개")
    cols = ['Period', '우리 Output', '우리 End'] + [c for c in ('차이 Output', '차이 End', '프로젝트 일치% Output',
                                                               '프로젝트 일치% End') if c in months.columns]
    for r in months[cols].itertuples(index=False):
        text = f"  {r[0]}  Output {r[1]:>16,.0f}  End {r[2]:>17,.0f}"
        if len(r) > 3 and pd.notna(r[3]):
            text += (f"  | 회계와 차이 Output {r[3]:>13,.0f}  End {r[4]:>14,.0f}"
                     f"  (프로젝트 일치 {r[5]}% / {r[6]}%)")
        print(text)
    if len(causes) and res.period in causes.columns:
        print("\n회계와의 End 차이 원인 (마감월):")
        for r in causes[['원인', res.period]].itertuples(index=False):
            if abs(r[1]) >= 1:
                print(f"  {r[0]:24s} {r[1]:>16,.0f}")
    if len(res.review):
        print("\n확인 목록:")
        for kind, g in res.review.groupby('구분', sort=False):
            print(f"  {kind}: {len(g)}건")
    if len(res.pending):
        print(f"\n출고 후 미청구 (마감월 말): {len(res.pending)}라인 {res.pending['금액(원)'].sum():,.0f}원")
    if res.next_filter_note:
        print(f"\n추출 범위: {res.next_filter_note}")
    print(f"\n{saved_note}")
    print(f"결과: {out_path}")
    if html_path:
        print(f"      {html_path}  (Edge로 열기)")


def create_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description='AX Sales order book — Item requirement + Item transaction으로 달마다 계산해 저장하는 오더북',
        epilog=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('period', help='계산할 달 (예: P10) — 저장소가 비었으면 P01부터 이 달까지 기준선')
    parser.add_argument('--year', type=int, help='연도 (생략 시 마감월 기준 올해/작년)')
    parser.add_argument('--ir', nargs='+', type=Path, help='Item requirement export (생략 시 그 달 폴더의 최신 파일)')
    parser.add_argument('--it', nargs='+', type=Path, help='Item transaction export (생략 시 그 달 폴더의 최신 파일)')
    parser.add_argument('--orderbook', nargs='+', type=Path, help='회계 오더북 파일 (생략 시 ax_orderbook 아래에서 찾음)')
    parser.add_argument('--fx', action='append', default=[], metavar='CUR=RATE',
                        help='환율 직접 입력 — 마감월이면 JPY=8.6157, 다른 달이면 JPY=2026-09=8.6968 (여러 번 가능)')
    parser.add_argument('--store', type=Path, help=f'저장소 파일 (기본: {AX_ORDERBOOK_DB.name})')
    parser.add_argument('--dry-run', action='store_true', help='계산하고 결과 파일(_미리보기)만 — 저장하지 않는다')
    parser.add_argument('--out', type=Path, help='결과 폴더 (기본: 그 달 폴더 ax_orderbook/연도/P월)')
    parser.add_argument('-v', '--verbose', action='store_true', help='상세 로그')
    return parser


def main(argv: list[str] | None = None) -> int:
    args = create_argument_parser().parse_args(argv)
    setup_logging(verbose=args.verbose)
    try:
        upto = parse_period(args.period)
        year = args.year or default_year(upto)
        close_ym = f"{year}-{upto:02d}"
        store = OrderbookStore(args.store or AX_ORDERBOOK_DB)
        plan = store.plan(year, upto)
        root = AX_ORDERBOOK_DIR
        here = period_dir(root, year, upto)
        ir_paths = args.ir or find_ir_files(root, year, upto)
        it_paths = args.it or find_it_files(root, year, upto)
        ob_paths = args.orderbook or find_orderbooks(root)
        needed = [('Item requirement', ir_paths, 'Item requirement_YYYYMMDD.xlsx', f"{here} (없으면 {root})"),
                  ('Item transaction', it_paths, 'Item transaction_YYYYMMDD.xlsx', f"{here} (없으면 {root})")]
        if plan.kind == 'baseline':
            needed.append(('회계 오더북', ob_paths, 'Sales order book_MM_YYYY.xlsx — 기준선은 P01 행 필수', f"{root} 아래"))
        for label, paths, name, where in needed:
            if not paths:
                print(f"[오류] {label} 파일이 없습니다 ({name}): {where}")
                return 1

        day, errors, warns = check_snapshot(list(ir_paths), list(it_paths), year, upto)
        for text in warns:
            print(f"[경고] {text}")
        for text in errors:
            print(f"[{'경고' if args.dry_run else '오류'}] {text}")
        if errors and not args.dry_run:
            print("저장하지 않았습니다 — 계산만 보려면 --dry-run")
            return 1

        print(f"[1/4] {describe_plan(plan, year, upto, preview=args.dry_run)}")
        if plan.kind == 'baseline':
            official, ob_path = load_official(list(ob_paths), year, upto)
        else:
            official, ob_path = load_reference(list(ob_paths), year, upto)
            if ob_path is None:
                print("[경고] 회계 오더북이 없어 부서 매핑은 규칙으로, 속성은 저장된 행에서만 가져옵니다")
        print(f"[2/4] IR·IT 읽는 중… ({', '.join(p.name for p in [*ir_paths, *it_paths])})")
        ir = read_ir(list(ir_paths))
        it = read_it(list(it_paths))
        fx_file, fx_copy = shared_file(root, FX_HISTORY_NAME)
        if fx_copy:              # 환율 이력은 달마다 이어 쓰는 파일 — 자리로 복사해 거기서 이어 간다
            shutil.copy2(fx_copy, fx_file)
            print(f"[안내] {FX_HISTORY_NAME}를 {fx_copy.parent}에서 {root}로 복사해 이어 씁니다 (여러 달이 같이 쓰는 파일)")
        hub_dirs = [d for d in (FINANCE_HUB_DIR, here) if d is not None]   # 그 달 폴더에 둔 마감환율 파일도 읽는다
        fx = load_fx_book(DB_FILE, fx_file, hub_dirs, parse_fx_overrides(args.fx, close_ym))
        currency_file, currency_copy = shared_file(root, CURRENCY_OVERRIDE_NAME)
        if currency_copy:
            print(f"[경고] {CURRENCY_OVERRIDE_NAME}가 {root}에 없어 {currency_copy}를 씁니다 — 여러 달이 같이 쓰는 "
                  f"파일이라 {root}로 옮기세요")
            currency_file = currency_copy
        overrides = read_currency_overrides(currency_file)

        print("[3/4] 계산 중…")
        if plan.kind == 'baseline':
            res = build_orderbook(official, ir, it, year, upto, fx, overrides)
            months, basis = list(range(1, upto + 1)), BASELINE
        else:
            res = build_month(official, ir, it, year, upto, fx, store.rows(*plan.prev), store.billed(*plan.prev),
                              store.rows(), overrides)
            months, basis = [upto], MONTHLY
        files = {'회계 오더북': ob_path.name if ob_path else '(없음)',
                 'Item requirement': ', '.join(p.name for p in ir_paths),
                 'Item transaction': ', '.join(p.name for p in it_paths)}
        snapshot = f"{day:%Y-%m-%d}" if day else ''
        earlier = store.rows(year)
        earlier = earlier[~earlier['pn'].isin(months)]
        raw_all = pd.concat([earlier, res.raw_rows], ignore_index=True) if len(earlier) else res.raw_rows
        if args.dry_run:
            saved_note = "미리보기 — 저장하지 않았습니다"
        else:
            store.save(year, months, res.raw_rows, res.billed, basis, files['Item requirement'],
                       files['Item transaction'], snapshot)
            saved_note = f"저장: {store.path} — {year} P{upto:02d} Open, 그 앞 달은 Closed"
        periods = period_table(store, year, months, basis, files, snapshot, res, saved=not args.dry_run)

        res_all = replace(res, raw_rows=raw_all, keys=keys_from_raw(raw_all))
        official_y = official_of_year(official, year)
        diff = classify_differences(res_all, official_y, ir)
        table = month_table(res_all, official_y)
        causes = cause_table(diff)
        out_dir = args.out or here
        out_path = out_dir / f"AX_오더북_{year}_P{upto:02d}{'_미리보기' if args.dry_run else ''}.xlsx"
        print("[4/4] 결과 파일 쓰는 중…")
        files['저장소'] = str(store.path)
        note = method_note(plan, upto)
        written = write_workbook(res_all, out_path, table, diff, causes, files, periods, note)
        html_path = write_html(res_all, written, diff, official_y, periods, files, note, preview=args.dry_run)
        print_summary(res_all, written, table, causes, saved_note, html_path)
        return 0
    except (ValueError, FileNotFoundError, KeyError) as exc:
        print(f"[오류] {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
