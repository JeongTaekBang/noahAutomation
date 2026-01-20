#!/usr/bin/env python
"""
Proforma Invoice 생성기
=======================

SO_ID를 입력하면 NOAH_SO_PO_DN.xlsx의 SO_해외 시트에서 해당 데이터를 읽어
자동으로 Proforma Invoice를 생성합니다.

사용법:
    python create_pi.py SOO-2026-0001              # 단일 생성
    python create_pi.py SOO-2026-0001 SOO-2026-0002 # 여러 건 동시 생성
"""

from __future__ import annotations

import argparse
import logging
import sys

import pandas as pd

from po_generator.config import (
    PI_OUTPUT_DIR,
    PI_TEMPLATE_FILE,
)
from po_generator.utils import (
    load_so_export_data,
    get_value,
)
from po_generator.pi_generator import create_pi_xlwings
from po_generator.cli_common import validate_output_path, generate_output_filename
from po_generator.logging_config import setup_logging
from po_generator.services import DocumentService, GenerationStatus

logger = logging.getLogger(__name__)


def print_available_ids(df_so: pd.DataFrame, limit: int = 15) -> None:
    """사용 가능한 SO_ID 목록 출력

    Args:
        df_so: SO 해외 데이터
        limit: 출력 제한 수
    """
    print("\n" + "=" * 60)
    print("사용 가능한 SO_ID 목록 (SO_해외)")
    print("=" * 60)

    so_ids = df_so['SO_ID'].dropna().unique().tolist()
    print(f"\n[Proforma Invoice] SO_ID ({len(so_ids)}건)")
    print("-" * 40)
    for so_id in so_ids[:limit]:
        # 고객명도 함께 표시
        customer = df_so[df_so['SO_ID'] == so_id]['Customer name'].iloc[0] if len(df_so[df_so['SO_ID'] == so_id]) > 0 else ''
        customer_short = str(customer)[:25] if customer else ''
        print(f"  {so_id:<20} {customer_short}")
    if len(so_ids) > limit:
        print(f"  ... 외 {len(so_ids) - limit}건")

    print("\n" + "=" * 60)
    print("위 SO_ID 중 하나를 입력하여 Proforma Invoice를 생성하세요.")
    print("=" * 60)


def generate_pi(so_id: str, df_so: pd.DataFrame) -> bool:
    """Proforma Invoice 생성

    DocumentService를 사용하여 Proforma Invoice를 생성합니다.

    Args:
        so_id: SO_ID
        df_so: SO 해외 데이터 (하위 호환용, 실제로는 사용하지 않음)

    Returns:
        성공 여부
    """
    print(f"\n{'=' * 60}")
    print(f"Proforma Invoice 생성: {so_id}")
    print('=' * 60)

    service = DocumentService()

    # 1. SO 데이터 검색 및 정보 출력
    order_data = service.finder.find_so_export(so_id)
    if order_data is None:
        print(f"  [오류] '{so_id}'를 찾을 수 없습니다.")
        return False

    # 2. 기본 정보 출력
    if order_data.is_multi_item:
        print(f"  [다중 아이템] {order_data.item_count}개 아이템 발견")
        for idx, (_, item) in enumerate(order_data.items_df.iterrows()):
            item_name = get_value(item, 'item_name', 'N/A')
            item_qty = get_value(item, 'item_qty', 'N/A')
            unit_price = get_value(item, 'sales_unit_price', 0)
            print(f"    {idx + 1}. {item_name} x {item_qty} @ {unit_price:,.2f}")

    print(f"  고객: {order_data.get_value('customer_name', 'N/A')}")
    if not order_data.is_multi_item:
        print(f"  품목: {order_data.get_value('item_name', 'N/A')}")
        print(f"  수량: {order_data.get_value('item_qty', 'N/A')}")
        unit_price = order_data.get_value('sales_unit_price', 0)
        currency = order_data.get_value('currency', 'USD')
        print(f"  단가: {currency} {unit_price:,.2f}")

    # 3. 문서 생성 (서비스 사용)
    result = service.generate_pi(so_id)

    # 4. 결과 처리
    if result.success:
        print(f"  -> Proforma Invoice 생성 완료: {result.output_file.name}")
        return True
    else:
        if result.status == GenerationStatus.FILE_ERROR:
            print(f"  [오류] {result.errors[0] if result.errors else result.message}")
        else:
            print(f"  [오류] {result.message}")
        logger.error(f"Proforma Invoice 생성 실패: {result.message}")
        return False


def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성

    Returns:
        설정된 ArgumentParser
    """
    description = """
Proforma Invoice 생성기 (해외 전용)
===================================

NOAH_SO_PO_DN.xlsx의 SO_해외 시트에서 데이터를 읽어
Proforma Invoice를 자동 생성합니다.
"""

    epilog = """
사용 예시:
  python create_pi.py SOO-2026-0001              # PI 1건 생성
  python create_pi.py SOO-2026-0001 SOO-2026-0002 # 여러 건 동시 생성

인자 없이 실행하면 사용 가능한 SO_ID 목록을 표시합니다.
"""

    parser = argparse.ArgumentParser(
        prog='create_pi',
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
    """메인 함수

    Returns:
        종료 코드 (0: 성공, 1: 실패)
    """
    parser = create_argument_parser()
    args = parser.parse_args()

    # 로깅 설정
    setup_logging(verbose=args.verbose)

    # 데이터 로드
    print("NOAH_SO_PO_DN.xlsx 로딩 중...")
    try:
        df_so = load_so_export_data()
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        return 1

    # 인자 없으면 도움말 + 사용 가능한 ID 출력
    if not args.so_ids:
        parser.print_help()
        print_available_ids(df_so)
        return 0

    print(f"SO 해외: {len(df_so)}건 로드 완료")

    # 각 SO_ID에 대해 Proforma Invoice 생성
    success_count = 0
    for so_id in args.so_ids:
        if generate_pi(so_id, df_so):
            success_count += 1

    # 결과 출력
    print(f"\n{'=' * 60}")
    print(f"완료: {success_count}/{len(args.so_ids)}건 Proforma Invoice 생성")
    print(f"출력 폴더: {PI_OUTPUT_DIR}")
    print('=' * 60)

    return 0 if success_count == len(args.so_ids) else 1


if __name__ == "__main__":
    sys.exit(main())
