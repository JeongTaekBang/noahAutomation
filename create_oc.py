#!/usr/bin/env python
"""
Order Confirmation 생성기
=========================

SO_ID를 입력하면 NOAH_SO_PO_DN.xlsx의 SO_해외 시트에서 해당 데이터를 읽어
자동으로 Order Confirmation을 생성합니다.
Dispatch date는 SO_해외의 'EXW NOAH' 컬럼 값을 사용합니다.

사용법:
    python create_oc.py SOO-2026-0001              # 단일 생성
    python create_oc.py SOO-2026-0001 SOO-2026-0002 # 여러 건 동시 생성
"""

from __future__ import annotations

import argparse
import logging
import sys

import pandas as pd

from po_generator.config import (
    OC_OUTPUT_DIR,
    OC_TEMPLATE_FILE,
)
from po_generator.utils import (
    load_so_export_with_customer,
    get_value,
)
from po_generator.logging_config import setup_logging
from po_generator.services import DocumentService, GenerationStatus

logger = logging.getLogger(__name__)


def print_available_ids(df_so: pd.DataFrame, limit: int = 15) -> None:
    """사용 가능한 SO_ID 목록 출력"""
    print("\n" + "=" * 60)
    print("사용 가능한 SO_ID 목록 (SO_해외)")
    print("=" * 60)

    so_ids = df_so['SO_ID'].dropna().unique().tolist()
    print(f"\n[Order Confirmation] SO_ID ({len(so_ids)}건)")
    print("-" * 40)
    for so_id in so_ids[:limit]:
        customer = df_so[df_so['SO_ID'] == so_id]['Customer name'].iloc[0] if len(df_so[df_so['SO_ID'] == so_id]) > 0 else ''
        customer_short = str(customer)[:25] if customer else ''
        print(f"  {so_id:<20} {customer_short}")
    if len(so_ids) > limit:
        print(f"  ... 외 {len(so_ids) - limit}건")

    print("\n" + "=" * 60)
    print("위 SO_ID 중 하나를 입력하여 Order Confirmation을 생성하세요.")
    print("=" * 60)


def generate_oc(so_id: str, df_so: pd.DataFrame) -> bool:
    """Order Confirmation 생성"""
    print(f"\n{'=' * 60}")
    print(f"Order Confirmation 생성: {so_id}")
    print('=' * 60)

    service = DocumentService()

    # 1. SO 데이터 검색 및 정보 출력
    order_data = service.finder.find_so_export_with_customer(so_id)
    if order_data is None:
        print(f"  [오류] '{so_id}'를 찾을 수 없습니다.")
        return False

    # 2. 기본 정보 출력
    print(f"  고객: {order_data.get_value('customer_name', 'N/A')}")
    bill_to_1 = order_data.get_value('bill_to_1', '')
    if bill_to_1:
        print(f"  Bill to: {bill_to_1}")

    payment_terms = order_data.get_value('payment_terms', '')
    if payment_terms:
        print(f"  Payment Terms: {payment_terms}")

    if order_data.is_multi_item:
        print(f"  [다중 아이템] {order_data.item_count}개 아이템 발견")
        for idx, (_, item) in enumerate(order_data.items_df.iterrows()):
            item_name = get_value(item, 'item_name', 'N/A')
            item_qty = get_value(item, 'item_qty', 'N/A')
            unit_price = get_value(item, 'sales_unit_price', 0)
            exw_noah = get_value(item, 'exw_noah', '')
            try:
                line = f"    {idx + 1}. {item_name} x {item_qty} @ {float(unit_price):,.2f}"
                if exw_noah:
                    line += f"  (EXW: {exw_noah})"
                print(line)
            except (ValueError, TypeError):
                print(f"    {idx + 1}. {item_name} x {item_qty}")
    else:
        item_name = order_data.get_value('item_name', 'N/A')
        print(f"  품목: {item_name}")

    # 3. 문서 생성
    result = service.generate_oc(so_id)

    # 4. 결과 처리
    if result.success:
        print(f"  -> Order Confirmation 생성 완료: {result.output_file.name}")
        return True
    else:
        if result.status == GenerationStatus.FILE_ERROR:
            print(f"  [오류] {result.errors[0] if result.errors else result.message}")
        else:
            print(f"  [오류] {result.message}")
        logger.error(f"Order Confirmation 생성 실패: {result.message}")
        return False


def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성"""
    description = """
Order Confirmation 생성기
=========================

NOAH_SO_PO_DN.xlsx의 SO_해외 시트에서 데이터를 읽어
Order Confirmation을 자동 생성합니다.
Dispatch date는 SO_해외의 EXW NOAH 컬럼을 사용합니다.
"""

    epilog = """
사용 예시:
  python create_oc.py SOO-2026-0001              # OC 1건 생성
  python create_oc.py SOO-2026-0001 SOO-2026-0002 # 여러 건 동시 생성

인자 없이 실행하면 사용 가능한 SO_ID 목록을 표시합니다.
"""

    parser = argparse.ArgumentParser(
        prog='create_oc',
        description=description,
        epilog=epilog,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        'so_ids',
        nargs='*',
        metavar='SO_ID',
        help='SO_ID (예: SOO-2026-0001)',
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
        df_so = load_so_export_with_customer()
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        return 1

    if not args.so_ids:
        parser.print_help()
        print_available_ids(df_so)
        return 0

    print(f"SO 해외: {len(df_so)}건 로드 완료")

    success_count = 0
    for so_id in args.so_ids:
        if generate_oc(so_id, df_so):
            success_count += 1

    print(f"\n{'=' * 60}")
    print(f"완료: {success_count}/{len(args.so_ids)}건 Order Confirmation 생성")
    print(f"출력 폴더: {OC_OUTPUT_DIR}")
    print('=' * 60)

    return 0 if success_count == len(args.so_ids) else 1


if __name__ == "__main__":
    sys.exit(main())
