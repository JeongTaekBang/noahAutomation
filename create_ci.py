#!/usr/bin/env python
"""
Commercial Invoice 생성기
=========================

DN_ID를 입력하면 NOAH_SO_PO_DN.xlsx의 DN_해외 시트에서 해당 데이터를 읽어
자동으로 Commercial Invoice를 생성합니다.

사용법:
    python create_ci.py DNO-2026-0001              # 단일 생성
    python create_ci.py DNO-2026-0001 DNO-2026-0002 # 여러 건 동시 생성
"""

from __future__ import annotations

import argparse
import logging
import sys

import pandas as pd

from po_generator.config import (
    CI_OUTPUT_DIR,
    CI_TEMPLATE_FILE,
)
from po_generator.utils import (
    load_dn_export_data,
    get_value,
)
from po_generator.logging_config import setup_logging
from po_generator.services import DocumentService, GenerationStatus

logger = logging.getLogger(__name__)


def print_available_ids(df_dn: pd.DataFrame, limit: int = 15) -> None:
    """사용 가능한 DN_ID 목록 출력"""
    print("\n" + "=" * 60)
    print("사용 가능한 DN_ID 목록 (DN_해외)")
    print("=" * 60)

    dn_ids = df_dn['DN_ID'].dropna().unique().tolist()
    print(f"\n[Commercial Invoice] DN_ID ({len(dn_ids)}건)")
    print("-" * 40)
    for dn_id in dn_ids[:limit]:
        customer = df_dn[df_dn['DN_ID'] == dn_id]['Customer name'].iloc[0] if len(df_dn[df_dn['DN_ID'] == dn_id]) > 0 else ''
        customer_short = str(customer)[:25] if customer else ''
        print(f"  {dn_id:<20} {customer_short}")
    if len(dn_ids) > limit:
        print(f"  ... 외 {len(dn_ids) - limit}건")

    print("\n" + "=" * 60)
    print("위 DN_ID 중 하나를 입력하여 Commercial Invoice를 생성하세요.")
    print("=" * 60)


def generate_ci(dn_id: str, df_dn: pd.DataFrame) -> bool:
    """Commercial Invoice 생성

    Args:
        dn_id: DN_ID
        df_dn: DN 해외 데이터 (하위 호환용)

    Returns:
        성공 여부
    """
    print(f"\n{'=' * 60}")
    print(f"Commercial Invoice 생성: {dn_id}")
    print('=' * 60)

    service = DocumentService()

    # 1. DN 데이터 검색 및 정보 출력
    order_data = service.finder.find_dn_export(dn_id)
    if order_data is None:
        print(f"  [오류] '{dn_id}'를 찾을 수 없습니다.")
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
            item_name = get_value(item, 'item_name', '')
            if not item_name:
                item_name = item.get('Item', 'N/A') if 'Item' in item.index else 'N/A'
            item_qty = get_value(item, 'item_qty', '')
            if not item_qty:
                item_qty = item.get('Qty', 'N/A') if 'Qty' in item.index else 'N/A'
            unit_price = get_value(item, 'unit_price', 0)
            try:
                print(f"    {idx + 1}. {item_name} x {item_qty} @ {float(unit_price):,.2f}")
            except (ValueError, TypeError):
                print(f"    {idx + 1}. {item_name} x {item_qty}")
    else:
        item_name = order_data.get_value('item_name', '')
        if not item_name:
            item_name = order_data.first_item.get('Item', 'N/A') if 'Item' in order_data.first_item.index else 'N/A'
        print(f"  품목: {item_name}")

    # 3. 문서 생성 (서비스 사용)
    result = service.generate_ci(dn_id)

    # 4. 결과 처리
    if result.success:
        print(f"  -> Commercial Invoice 생성 완료: {result.output_file.name}")
        return True
    else:
        if result.status == GenerationStatus.FILE_ERROR:
            print(f"  [오류] {result.errors[0] if result.errors else result.message}")
        else:
            print(f"  [오류] {result.message}")
        logger.error(f"Commercial Invoice 생성 실패: {result.message}")
        return False


def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성"""
    description = """
Commercial Invoice 생성기
=========================

NOAH_SO_PO_DN.xlsx의 DN_해외 시트에서 데이터를 읽어
Commercial Invoice를 자동 생성합니다.
"""

    epilog = """
사용 예시:
  python create_ci.py DNO-2026-0001              # CI 1건 생성
  python create_ci.py DNO-2026-0001 DNO-2026-0002 # 여러 건 동시 생성

인자 없이 실행하면 사용 가능한 DN_ID 목록을 표시합니다.
"""

    parser = argparse.ArgumentParser(
        prog='create_ci',
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
    for dn_id in args.dn_ids:
        if generate_ci(dn_id, df_dn):
            success_count += 1

    print(f"\n{'=' * 60}")
    print(f"완료: {success_count}/{len(args.dn_ids)}건 Commercial Invoice 생성")
    print(f"출력 폴더: {CI_OUTPUT_DIR}")
    print('=' * 60)

    return 0 if success_count == len(args.dn_ids) else 1


if __name__ == "__main__":
    sys.exit(main())
