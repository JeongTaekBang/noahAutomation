#!/usr/bin/env python
"""
Industry Code 대사 (Industry Code Reconciliation)
===================================================

Orderbook 파일의 빈 Industry code를
NOAH_SO_PO_DN.xlsx (PO → SO) 매핑을 통해 채워서 새 파일로 출력합니다.
추가 컬럼: NOAH Sector, 매핑상태

매핑 체인:
    Orderbook 발주번호 → PO시트 NOAH O.C No. → SO_ID → SO시트 Industry code, Sector

사용법:
    python reconcile_ind.py P03           # P03 대사
    python reconcile_ind.py P03 -v        # 상세 로그
"""

from __future__ import annotations

import argparse
import logging
import sys
import warnings
from pathlib import Path

import pandas as pd
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

from po_generator.config import (
    NOAH_SO_PO_DN_FILE, BASE_DIR,
    SO_DOMESTIC_SHEET, SO_EXPORT_SHEET,
    PO_DOMESTIC_SHEET, PO_EXPORT_SHEET,
)
from po_generator.logging_config import setup_logging

logger = logging.getLogger(__name__)

RECON_DIR = BASE_DIR / "ind_code_reconciliation"

# NOAH_SO_PO_DN 컬럼
PO_OC_COL = 'NOAH O.C No.'
PO_SO_ID_COL = 'SO_ID'
SO_SECTOR_COL = 'Sector'
SO_IND_COL = 'Industry code'


def find_orderbook_file(period_code: str) -> Path | None:
    """ind_code_reconciliation/{period_code}/ 에서 Orderbook 파일 찾기"""
    period_dir = RECON_DIR / period_code.upper()
    if not period_dir.exists():
        return None
    for f in period_dir.iterdir():
        if f.suffix == '.xlsx' and not f.name.startswith('~'):
            if 'orderbook' in f.name.lower() and '결과' not in f.name:
                return f
    return None


def get_ob_sheet_name(period_code: str) -> str:
    """기간 코드로 Orderbook 시트명 생성 (P03 → P03_start)"""
    return f"{period_code.upper()}_start"


def build_mapping() -> dict[str, tuple[str, object, str | None]]:
    """발주번호 → (SO_ID, Industry code, Sector) 매핑 딕셔너리 생성"""
    # PO: NOAH O.C No. → SO_ID
    po_dom = pd.read_excel(
        NOAH_SO_PO_DN_FILE, sheet_name=PO_DOMESTIC_SHEET,
        usecols=[PO_OC_COL, PO_SO_ID_COL],
    )
    po_exp = pd.read_excel(
        NOAH_SO_PO_DN_FILE, sheet_name=PO_EXPORT_SHEET,
        usecols=[PO_OC_COL, PO_SO_ID_COL],
    )
    po_map = pd.concat([po_dom, po_exp], ignore_index=True)
    po_map = po_map.dropna(subset=[PO_OC_COL, PO_SO_ID_COL])
    po_map[PO_OC_COL] = po_map[PO_OC_COL].astype(str).str.strip()
    po_map[PO_SO_ID_COL] = po_map[PO_SO_ID_COL].astype(str).str.strip()
    po_map = po_map.drop_duplicates(subset=[PO_OC_COL])

    # SO: SO_ID → (Industry code, Sector)
    so_dom = pd.read_excel(
        NOAH_SO_PO_DN_FILE, sheet_name=SO_DOMESTIC_SHEET,
        usecols=[PO_SO_ID_COL, SO_IND_COL, SO_SECTOR_COL],
    )
    so_exp = pd.read_excel(
        NOAH_SO_PO_DN_FILE, sheet_name=SO_EXPORT_SHEET,
        usecols=[PO_SO_ID_COL, SO_IND_COL, SO_SECTOR_COL],
    )
    so_map = pd.concat([so_dom, so_exp], ignore_index=True)
    so_map = so_map.dropna(subset=[PO_SO_ID_COL])
    so_map[PO_SO_ID_COL] = so_map[PO_SO_ID_COL].astype(str).str.strip()
    so_map = so_map.drop_duplicates(subset=[PO_SO_ID_COL])

    # Join & build dict
    full = po_map.merge(so_map, on=PO_SO_ID_COL, how='left')
    logger.debug("매핑: %d건 (Industry code 있음: %d건)",
                 len(full), full[SO_IND_COL].notna().sum())

    result = {}
    for _, row in full.iterrows():
        oc = row[PO_OC_COL]
        so_id = row[PO_SO_ID_COL]
        ind = row[SO_IND_COL] if pd.notna(row[SO_IND_COL]) else None
        sector = row[SO_SECTOR_COL] if pd.notna(row[SO_SECTOR_COL]) else None
        result[oc] = (so_id, ind, sector)

    return result


def fill_industry_code(
    ob: pd.DataFrame, mapping: dict, ind_col: str,
) -> tuple[pd.DataFrame, int, int, int, int]:
    """Orderbook에 Industry code, Sector, 매핑상태 채우기

    Returns:
        (result_df, total_null, filled, so_missing, po_missing)
    """
    result = ob.copy()
    result['NOAH Sector'] = None
    result['매핑상태'] = None

    null_mask = result[ind_col].isna()
    total_null = int(null_mask.sum())
    filled = 0
    so_missing = 0
    po_missing = 0

    for idx in result[null_mask].index:
        order_no = result.at[idx, '발주번호']
        if pd.isna(order_no):
            continue
        order_no = str(order_no).strip()

        if order_no in mapping:
            so_id, ind_code, sector = mapping[order_no]
            if ind_code is not None:
                result.at[idx, ind_col] = ind_code
                result.at[idx, 'NOAH Sector'] = sector
                result.at[idx, '매핑상태'] = '매칭'
                filled += 1
            else:
                result.at[idx, 'NOAH Sector'] = sector
                result.at[idx, '매핑상태'] = 'SO에 Industry code 없음'
                so_missing += 1
        else:
            result.at[idx, '매핑상태'] = 'PO에 발주번호 없음'
            po_missing += 1

    return result, total_null, filled, so_missing, po_missing


def write_output(result: pd.DataFrame, output_file: Path) -> None:
    """결과 Excel 파일 출력"""

    def _add_table(writer, sheet_name: str, display_name: str) -> None:
        ws = writer.sheets[sheet_name]
        if ws.max_row < 2:
            return
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        tbl = Table(displayName=display_name, ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tbl)

    legend = pd.DataFrame([
        ['매칭', '발주번호 → PO → SO 체인으로 Industry code 찾아 채움'],
        ['SO에 Industry code 없음', 'PO에서 SO_ID 찾았으나 SO시트에 Industry code 비어있음'],
        ['PO에 발주번호 없음', 'NOAH_SO_PO_DN PO시트에 해당 NOAH O.C No. 없음'],
    ], columns=['매칭상태', '설명'])

    output_file.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        result.to_excel(writer, sheet_name='결과', index=False)
        _add_table(writer, '결과', '결과')

        legend.to_excel(writer, sheet_name='범례', index=False)
        _add_table(writer, '범례', '범례')


def print_summary(
    total_null: int, filled: int, so_missing: int, po_missing: int,
) -> None:
    """콘솔 요약 출력"""
    print()
    print("Industry Code 대사 결과")
    print("=" * 60)
    print(f"  빈 Industry code: {total_null}행")
    print(f"    매칭 (채움):             {filled}")
    if so_missing > 0:
        print(f"    SO에 Industry code 없음: {so_missing}")
    if po_missing > 0:
        print(f"    PO에 발주번호 없음:       {po_missing}")
    print(f"  채움 후 남은 빈 행:        {total_null - filled}")


def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성"""
    parser = argparse.ArgumentParser(
        prog='reconcile_ind',
        description='Industry Code 대사 — Orderbook의 빈 Industry code를 SO에서 채워 새 파일 생성',
        epilog='예시: python reconcile_ind.py P03',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        'period',
        help='대사 월 코드 (예: P03, P04)',
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='상세 로그 출력',
    )

    return parser


def main() -> int:
    """메인 함수"""
    parser = create_argument_parser()
    args = parser.parse_args()

    setup_logging(verbose=args.verbose)
    period = args.period.upper()

    # 1. 파일 찾기
    ob_file = find_orderbook_file(period)
    if not ob_file:
        print(f"[오류] Orderbook 파일을 찾을 수 없습니다 (ind_code_reconciliation/{period}/*orderbook*)")
        return 1

    if not NOAH_SO_PO_DN_FILE.exists():
        print(f"[오류] NOAH_SO_PO_DN.xlsx를 찾을 수 없습니다: {NOAH_SO_PO_DN_FILE}")
        return 1

    ob_sheet = get_ob_sheet_name(period)
    print(f"Orderbook: {ob_file.name} [{ob_sheet}]")
    print(f"NOAH:      {NOAH_SO_PO_DN_FILE.name}")

    # 2. 데이터 로드
    try:
        ob = pd.read_excel(ob_file, sheet_name=ob_sheet)
    except Exception as e:
        print(f"[오류] Orderbook 로드 실패: {e}")
        return 1

    # Industry code 컬럼 탐색 (typo 대응)
    ind_col = None
    for candidate in ['Indusry code', 'Industry code']:
        if candidate in ob.columns:
            ind_col = candidate
            break
    if ind_col is None:
        print(f"[오류] Industry code 컬럼을 찾을 수 없습니다")
        return 1

    total_null = int(ob[ind_col].isna().sum())
    print(f"전체 행:   {len(ob)}행")
    print(f"빈 Industry code: {total_null}행")

    if total_null == 0:
        print("\n모든 Industry code가 이미 채워져 있습니다.")
        return 0

    # 3. 매핑 구성
    try:
        mapping = build_mapping()
    except Exception as e:
        print(f"[오류] 매핑 데이터 로드 실패: {e}")
        return 1

    print(f"매핑 건수: {len(mapping)}건 (PO→SO)")

    # 4. Industry code 채우기
    result, total_null, filled, so_missing, po_missing = fill_industry_code(
        ob, mapping, ind_col,
    )

    # 5. Excel 출력
    output_file = RECON_DIR / period / f"ind_code_결과_{period}.xlsx"
    write_output(result, output_file)
    print(f"\n출력: {output_file}")

    # 6. 콘솔 요약
    print_summary(total_null, filled, so_missing, po_missing)

    return 0


if __name__ == "__main__":
    sys.exit(main())
