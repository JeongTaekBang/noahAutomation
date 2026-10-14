#!/usr/bin/env python
"""
Certificate of Origin(원산지증명서) 생성기
==========================================

DN_ID를 입력하면 NOAH_SO_PO_DN.xlsx의 DN_해외 시트에서 해당 데이터를 읽어
제조사(NOAH) 원산지 선언서를 xlsx + PDF로 생성합니다.

사용법:
    python create_coo.py DNO-2026-0001              # 단일 생성
    python create_coo.py DNO-2026-0001 DNO-2026-0002 # 여러 건 동시 생성
"""

from __future__ import annotations

import argparse
import logging
import sys

import pandas as pd

from po_generator.config import COO_OUTPUT_DIR
from po_generator.hs_code import is_non_goods
from po_generator.utils import load_dn_export_data, get_value
from po_generator.logging_config import setup_logging
from po_generator.services import DocumentService, GenerationStatus

logger = logging.getLogger(__name__)


def print_available_ids(df_dn: pd.DataFrame, limit: int = 15) -> None:
    """사용 가능한 DN_ID 목록 출력"""
    print("\n" + "=" * 60)
    print("사용 가능한 DN_ID 목록 (DN_해외)")
    print("=" * 60)

    dn_ids = df_dn['DN_ID'].dropna().unique().tolist()
    print(f"\n[Certificate of Origin] DN_ID ({len(dn_ids)}건)")
    print("-" * 40)
    for dn_id in dn_ids[:limit]:
        rows = df_dn[df_dn['DN_ID'] == dn_id]
        customer = rows['Customer name'].iloc[0] if len(rows) > 0 else ''
        print(f"  {dn_id:<20} {str(customer)[:25] if customer else ''}")
    if len(dn_ids) > limit:
        print(f"  ... 외 {len(dn_ids) - limit}건")

    print("\n" + "=" * 60)
    print("위 DN_ID 중 하나를 입력하여 Certificate of Origin을 생성하세요.")
    print("=" * 60)


def generate_coo(dn_id: str) -> tuple[bool, list[str]]:
    """Certificate of Origin 생성

    Returns:
        (성공 여부, 사람이 확인해야 할 메모 — 마지막에 다시 모아 찍는다)
    """
    print(f"\n{'=' * 60}")
    print(f"Certificate of Origin 생성: {dn_id}")
    print('=' * 60)

    service = DocumentService()

    order_data = service.finder.find_dn_export(dn_id)
    if order_data is None:
        print(f"  [오류] '{dn_id}'를 찾을 수 없습니다.")
        return False, []

    print(f"  고객: {order_data.get_value('customer_name', 'N/A')}")
    bill_to_1 = order_data.get_value('bill_to_1', '')
    if bill_to_1:
        print(f"  Consignee: {bill_to_1}")

    goods = 0
    for _, item in order_data.all_items.iterrows():
        item_name = get_value(item, 'item_name', '')
        if is_non_goods(item_name):
            print(f"    - {item_name} (운임 — 원산지증명서에 싣지 않음)")
            continue
        goods += 1
        print(f"    {goods}. {item_name} x {get_value(item, 'item_qty', '')}")

    result = service.generate_coo(dn_id)

    if result.success:
        print(f"  -> Certificate of Origin 생성 완료: {result.output_file.name}")
        for extra in result.extra_files:
            print(f"  -> PDF: {extra.name}")
        for warning in result.warnings:
            print(f"  [확인 필요] {warning}")
        for error in result.errors:
            print(f"  [실패] {error}")
        # PDF가 빠지면 보낼 본이 없다 — 성공으로 보고하지 않는다
        return not result.errors, result.warnings + result.errors

    if result.status == GenerationStatus.FILE_ERROR:
        print(f"  [오류] {result.errors[0] if result.errors else result.message}")
    else:
        print(f"  [오류] {result.message}")
    logger.error(f"Certificate of Origin 생성 실패: {result.message}")
    return False, []


def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성"""
    description = """
Certificate of Origin(원산지증명서) 생성기
==========================================

NOAH_SO_PO_DN.xlsx의 DN_해외 시트에서 데이터를 읽어
제조사(NOAH) 원산지 선언서를 xlsx + PDF로 생성합니다.
"""

    epilog = """
사용 예시:
  python create_coo.py DNO-2026-0001              # COO 1건 생성
  python create_coo.py DNO-2026-0001 DNO-2026-0002 # 여러 건 동시 생성

인자 없이 실행하면 사용 가능한 DN_ID 목록을 표시합니다.
"""

    parser = argparse.ArgumentParser(
        prog='create_coo',
        description=description,
        epilog=epilog,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        'dn_ids',
        nargs='*',
        metavar='DN_ID',
        help='DN_ID (예: DNO-2026-0001)',
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

    print("NOAH_SO_PO_DN.xlsx 로딩 중...")
    try:
        df_dn = load_dn_export_data()
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        return 1

    if not args.dn_ids:
        parser.print_help()
        print_available_ids(df_dn)
        return 0

    print(f"DN 해외: {len(df_dn)}건 로드 완료")

    success_count = 0
    notices: list[str] = []
    for dn_id in args.dn_ids:
        ok, dn_notices = generate_coo(dn_id)
        if ok:
            success_count += 1
        notices.extend(f"{dn_id}  {note}" for note in dn_notices)

    print(f"\n{'=' * 60}")
    print(f"완료: {success_count}/{len(args.dn_ids)}건 Certificate of Origin 생성")
    # 경고는 문서가 길면 스크롤 위로 사라진다 — 요약에서 한 번 더 보여 준다
    for note in notices:
        print(f"[확인 필요] {note}")
    print(f"출력 폴더: {COO_OUTPUT_DIR}")
    print('=' * 60)

    return 0 if success_count == len(args.dn_ids) else 1


if __name__ == "__main__":
    sys.exit(main())
