#!/usr/bin/env python
"""
AX Sales order book — IR + Item transaction으로 만드는 오더북
===========================================================

AX export 두 개(Item requirement, Item transaction)로 회계팀 오더북 Raw 형식의 연속 오더북(P01~마감월)을
만든다. NOAH 파워쿼리 Order_Book 컨셉 — 원천 데이터로 매번 전부 다시 계산한다(계산 규칙은 po_generator/ax_orderbook.py).
회계 오더북은 P01 Start(기초)·키 매핑·속성에 쓰고, 같은 달 행이 있으면 대조해 차이와 원인을 붙인다.

준비물 (ax_orderbook/ 폴더, 마감 직후 같은 날 한 쌍):
    Item requirement_YYYYMMDD.xlsx   AX Item requirement — Show=All, Created date ≥ 01/01/2024 하나만
    Item transaction_YYYYMMDD.xlsx   AX Item transaction(프로젝트 품목 거래) — Project date ≥ 01/01/2024 하나만
    Sales order book_MM_YYYY.xlsx    회계 오더북 (P01 행 필수)
    환율: NOAH_SO_PO_DN FX 시트(sync_db.py로 DB에) + Finance Hub 마감환율 파일(fx_history.csv에 쌓는다)

사용법:
    python build_ax_orderbook.py P10                  # 연도 생략 시 마감월 기준 올해/작년 → P01~P10
    python build_ax_orderbook.py P10 --year 2026
    python build_ax_orderbook.py P10 --fx JPY=8.6157  # 마감월 환율을 직접 넣기 (CUR=YYYY-MM=RATE도 됨)
    python build_ax_orderbook.py P10 --ir 파일 --it 파일 --orderbook 파일
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
import warnings
from datetime import date, datetime
from pathlib import Path

import pandas as pd

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

from po_generator.ax_orderbook import (
    DIFF_CAUSES, RAW_COLUMNS, OrderbookResult, build_orderbook, classify_differences, compare_with_official,
    load_fx_book, parse_period, read_currency_overrides, read_ir, read_it, read_raw,
)
from po_generator.config import AX_ORDERBOOK_DIR, DB_FILE, FINANCE_HUB_DIR
from po_generator.logging_config import setup_logging

logger = logging.getLogger(__name__)

FX_HISTORY_FILE = AX_ORDERBOOK_DIR / 'fx_history.csv'
# 외화 거래처 목록에 없는 외화 주문 — `코드,통화` (거래처 C번호 또는 프로젝트 번호). 확인 목록 '통화 확인'을 보고 추가
CURRENCY_OVERRIDE_FILE = AX_ORDERBOOK_DIR / '통화지정.csv'
_ORDERBOOK_RE = re.compile(r'Sales order book_(\d{1,2})_(\d{4})', re.IGNORECASE)
_DATED_RE = re.compile(r'_(\d{8})$')


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


def find_ir_files(folder: Path) -> list[Path]:
    return _latest_dated(folder, 'Item requirement')


def find_it_files(folder: Path) -> list[Path]:
    return _latest_dated(folder, 'Item transaction')


def find_orderbooks(folder: Path) -> list[Path]:
    files = [p for p in folder.glob('Sales order book_*.xlsx') if not p.name.startswith('~$')]
    def stamp(p: Path):
        m = _ORDERBOOK_RE.search(p.stem)
        return (int(m.group(2)), int(m.group(1))) if m else (0, 0)
    return sorted(files, key=stamp, reverse=True)


def load_official(paths: list[Path], year: int, upto: int) -> tuple[pd.DataFrame, Path]:
    """P01 행이 있는 가장 새 회계 오더북 → (P01~upto 행, 파일). P01 Start가 기초값이라 P01 행이 없으면 못 쓴다."""
    for path in paths:
        m = _ORDERBOOK_RE.search(path.stem)
        if m and int(m.group(2)) != year:
            continue
        raw = read_raw(path)
        if (raw['pn'] == 1).any():
            return raw[raw['pn'] <= upto].copy(), path
    raise FileNotFoundError(f"{year}년 P01 행이 있는 회계 오더북(Sales order book_*.xlsx)을 못 찾았습니다")


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


def _summary_rows(res: OrderbookResult, months: pd.DataFrame, causes: pd.DataFrame, files: dict[str, str]) -> list[list]:
    rows: list[list] = [
        [f"AX 오더북 (IR + Item transaction) — {res.year}년 P01~{res.period}"],
        [f"만든 시각: {datetime.now():%Y-%m-%d %H:%M}"],
        *[[f"{label}: {name}"] for label, name in files.items()],
        ['Start(P01) = 회계 P01 Start, 그 뒤로는 직전 달 우리 End. Output = 청구된 출고(출고월), End = 미출고 + 청구 전 출고.'],
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
                   causes: pd.DataFrame, files: dict[str, str]) -> Path:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    sheets: list[tuple[str, pd.DataFrame]] = [('Raw', res.raw_rows[list(RAW_COLUMNS)]), ('차이_월별', months)]
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
        pd.DataFrame(_summary_rows(res, months, causes, files)).to_excel(
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


def print_summary(res: OrderbookResult, out_path: Path, months: pd.DataFrame, causes: pd.DataFrame) -> None:
    print(f"\n[{res.year} P01~{res.period}] Raw 행 {len(res.raw_rows):,}개")
    cols = ['Period', '우리 Output', '우리 End'] + [c for c in ('차이 Output', '차이 End', '프로젝트 일치% Output',
                                                               '프로젝트 일치% End') if c in months.columns]
    for r in months[cols].itertuples(index=False):
        text = f"  {r[0]}  Output {r[1]:>16,.0f}  End {r[2]:>17,.0f}"
        if len(r) > 3:
            text += (f"  | 회계와 차이 Output {r[3]:>13,.0f}  End {r[4]:>14,.0f}"
                     f"  (프로젝트 일치 {r[5]}% / {r[6]}%)")
        print(text)
    if len(causes):
        print("\n회계와의 End 차이 원인 (마감월):")
        last = causes[['원인', res.period]] if res.period in causes.columns else causes.iloc[:, :2]
        for r in last.itertuples(index=False):
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
    print(f"\n결과: {out_path}")


def create_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description='AX Sales order book — Item requirement + Item transaction으로 P01~마감월 오더북',
        epilog=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('period', help='마감월 (예: P10) — P01부터 이 달까지 만든다')
    parser.add_argument('--year', type=int, help='연도 (생략 시 마감월 기준 올해/작년)')
    parser.add_argument('--ir', nargs='+', type=Path, help='Item requirement export (생략 시 폴더의 최신 파일)')
    parser.add_argument('--it', nargs='+', type=Path, help='Item transaction export (생략 시 폴더의 최신 파일)')
    parser.add_argument('--orderbook', nargs='+', type=Path, help='회계 오더북 파일 (생략 시 폴더에서 찾음)')
    parser.add_argument('--fx', action='append', default=[], metavar='CUR=RATE',
                        help='환율 직접 입력 — 마감월이면 JPY=8.6157, 다른 달이면 JPY=2026-09=8.6968 (여러 번 가능)')
    parser.add_argument('--out', type=Path, help='결과 폴더 (기본: ax_orderbook/연도)')
    parser.add_argument('-v', '--verbose', action='store_true', help='상세 로그')
    return parser


def main(argv: list[str] | None = None) -> int:
    args = create_argument_parser().parse_args(argv)
    setup_logging(verbose=args.verbose)
    try:
        upto = parse_period(args.period)
        year = args.year or default_year(upto)
        close_ym = f"{year}-{upto:02d}"
        ir_paths = args.ir or find_ir_files(AX_ORDERBOOK_DIR)
        it_paths = args.it or find_it_files(AX_ORDERBOOK_DIR)
        ob_paths = args.orderbook or find_orderbooks(AX_ORDERBOOK_DIR)
        for label, paths, name in (('Item requirement', ir_paths, 'Item requirement_YYYYMMDD.xlsx'),
                                   ('Item transaction', it_paths, 'Item transaction_YYYYMMDD.xlsx'),
                                   ('회계 오더북', ob_paths, 'Sales order book_MM_YYYY.xlsx')):
            if not paths:
                print(f"[오류] {label} 파일이 없습니다 ({name}): {AX_ORDERBOOK_DIR}")
                return 1
        print(f"[1/4] 회계 오더북 읽는 중… ({', '.join(p.name for p in ob_paths)})")
        official, ob_path = load_official(list(ob_paths), year, upto)
        print(f"[2/4] IR·IT 읽는 중… ({', '.join(p.name for p in [*ir_paths, *it_paths])})")
        ir = read_ir(list(ir_paths))
        it = read_it(list(it_paths))
        fx = load_fx_book(DB_FILE, FX_HISTORY_FILE, FINANCE_HUB_DIR, parse_fx_overrides(args.fx, close_ym))
        print(f"[3/4] P01~P{upto:02d} 계산 중…")
        res = build_orderbook(official, ir, it, year, upto, fx, read_currency_overrides(CURRENCY_OVERRIDE_FILE))
        diff = classify_differences(res, official, ir)
        months = month_table(res, official)
        causes = cause_table(diff)
        out_dir = args.out or (AX_ORDERBOOK_DIR / str(year))
        out_path = out_dir / f"AX_오더북_{year}_P{upto:02d}.xlsx"
        print("[4/4] 결과 파일 쓰는 중…")
        files = {'회계 오더북': ob_path.name, 'Item requirement': ', '.join(p.name for p in ir_paths),
                 'Item transaction': ', '.join(p.name for p in it_paths)}
        written = write_workbook(res, out_path, months, diff, causes, files)
        print_summary(res, written, months, causes)
        return 0
    except (ValueError, FileNotFoundError, KeyError) as exc:
        print(f"[오류] {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
