#!/usr/bin/env python
"""
거래명세표(Transaction Statement) 생성기
=========================================

DN_ID 또는 선수금_ID를 입력하면 NOAH_SO_PO_DN.xlsx에서 해당 데이터를 읽어
자동으로 거래명세표를 생성합니다.

사용법:
    python create_ts.py DN-2026-0001              # 납품 거래명세표
    python create_ts.py ADV_2026-0001             # 선수금 거래명세표
    python create_ts.py DN-2026-0001 DN-2026-0002 # 여러 건 동시 생성
"""

from __future__ import annotations

import argparse
import logging
import sys

import pandas as pd

from po_generator.config import (
    TS_OUTPUT_DIR,
    TS_TEMPLATE_FILE,
)
from po_generator.utils import (
    load_dn_data,
    load_pmt_data,
    get_value,
)
from po_generator.ts_generator import create_ts_xlwings
from po_generator.cli_common import validate_output_path, generate_output_filename
from po_generator.logging_config import setup_logging
from po_generator.services import DocumentService, GenerationStatus

logger = logging.getLogger(__name__)


def detect_id_type(doc_id: str) -> str:
    """ID 유형 감지

    Args:
        doc_id: 문서 ID

    Returns:
        'DN' 또는 'ADV'
    """
    doc_id_upper = doc_id.upper()
    if doc_id_upper.startswith('DN'):
        return 'DN'
    elif doc_id_upper.startswith('ADV'):
        return 'ADV'
    else:
        # 기본값: DN으로 처리
        return 'DN'


def print_available_ids(df_dn: pd.DataFrame, df_pmt: pd.DataFrame, limit: int = 10) -> None:
    """사용 가능한 ID 목록 출력

    Args:
        df_dn: DN 데이터
        df_pmt: PMT 데이터
        limit: 출력 제한 수
    """
    print("\n" + "=" * 50)
    print("사용 가능한 ID 목록 (DN_국내 / PMT_국내)")
    print("=" * 50)

    # DN 목록
    dn_ids = df_dn['DN_ID'].dropna().unique().tolist()
    print(f"\n[납품 거래명세표] DN_ID ({len(dn_ids)}건)")
    print("-" * 30)
    for dn_id in dn_ids[:limit]:
        # 고객명도 함께 표시
        customer = df_dn[df_dn['DN_ID'] == dn_id]['Customer name'].iloc[0] if len(df_dn[df_dn['DN_ID'] == dn_id]) > 0 else ''
        customer_short = str(customer)[:20] if customer else ''
        print(f"  {dn_id:<18} {customer_short}")
    if len(dn_ids) > limit:
        print(f"  ... 외 {len(dn_ids) - limit}건")

    # PMT 목록
    pmt_ids = df_pmt['선수금_ID'].dropna().unique().tolist()
    print(f"\n[선수금 거래명세표] 선수금_ID ({len(pmt_ids)}건)")
    print("-" * 30)
    for pmt_id in pmt_ids[:limit]:
        # 고객명도 함께 표시
        customer = df_pmt[df_pmt['선수금_ID'] == pmt_id]['Customer name'].iloc[0] if len(df_pmt[df_pmt['선수금_ID'] == pmt_id]) > 0 else ''
        customer_short = str(customer)[:20] if customer else ''
        print(f"  {pmt_id:<18} {customer_short}")
    if len(pmt_ids) > limit:
        print(f"  ... 외 {len(pmt_ids) - limit}건")

    print("\n" + "=" * 50)
    print("위 ID 중 하나를 입력하여 거래명세표를 생성하세요.")
    print("=" * 50)


def generate_ts_from_dn(dn_id: str, df_dn: pd.DataFrame) -> bool:
    """DN 기반 거래명세표 생성

    DocumentService를 사용하여 거래명세표를 생성합니다.

    Args:
        dn_id: DN_ID
        df_dn: DN 데이터 (하위 호환용, 실제로는 사용하지 않음)

    Returns:
        성공 여부
    """
    print(f"\n{'=' * 50}")
    print(f"거래명세표 생성 (납품): {dn_id}")
    print('=' * 50)

    service = DocumentService()

    # 1. DN 데이터 검색 및 정보 출력
    order_data = service.finder.find_dn(dn_id)
    if order_data is None:
        print(f"  [오류] '{dn_id}'를 찾을 수 없습니다.")
        return False

    # 2. 기본 정보 출력
    if order_data.is_multi_item:
        print(f"  [다중 아이템] {order_data.item_count}개 아이템 발견")
        for idx, (_, item) in enumerate(order_data.items_df.iterrows()):
            item_name = get_value(item, 'item_name', 'N/A')
            item_qty = get_value(item, 'item_qty', 'N/A')
            print(f"    {idx + 1}. {item_name} x {item_qty}")

    print(f"  고객: {order_data.get_value('customer_name', 'N/A')}")
    if not order_data.is_multi_item:
        print(f"  품목: {order_data.get_value('item_name', 'N/A')}")
        print(f"  수량: {order_data.get_value('item_qty', 'N/A')}")
        unit_price = order_data.get_value('sales_unit_price', 0)
        if unit_price:
            print(f"  단가: {unit_price:,}")

    # 3. 문서 생성 (서비스 사용)
    result = service.generate_ts(dn_id, doc_type='DN')

    # 4. 결과 처리
    if result.success:
        print(f"  -> 거래명세표 생성 완료: {result.output_file.name}")
        return True
    else:
        if result.status == GenerationStatus.FILE_ERROR:
            print(f"  [오류] {result.errors[0] if result.errors else result.message}")
        else:
            print(f"  [오류] {result.message}")
        return False


def generate_ts_from_adv(advance_id: str) -> bool:
    """선수금 거래명세표 생성 (SO_국내 데이터 사용)

    DocumentService를 사용하여 선수금 거래명세표를 생성합니다.

    Args:
        advance_id: 선수금_ID

    Returns:
        성공 여부
    """
    print(f"\n{'=' * 50}")
    print(f"거래명세표 생성 (선수금): {advance_id}")
    print('=' * 50)

    service = DocumentService()

    # 1. SO 데이터 로드 (선수금_ID -> SO_ID -> SO 아이템들)
    result = service.finder.find_so_for_advance(advance_id)
    if result is None:
        print(f"  [오류] '{advance_id}'를 찾을 수 없습니다.")
        return False

    pmt_data, order_data = result

    # 2. 기본 정보 출력
    customer_name = order_data.get_value('customer_name', 'N/A')
    print(f"  고객: {customer_name}")
    print(f"  SO_ID: {order_data.get_value('so_id', 'N/A')}")

    if order_data.is_multi_item:
        print(f"  [다중 아이템] {order_data.item_count}개 아이템 발견")
        for idx, (_, item) in enumerate(order_data.items_df.iterrows()):
            item_name = get_value(item, 'item_name', 'N/A')
            item_qty = get_value(item, 'item_qty', 'N/A')
            unit_price = get_value(item, 'sales_unit_price', 0)
            print(f"    {idx + 1}. {item_name} x {item_qty} @ {unit_price:,.0f}")
    else:
        print(f"  품목: {order_data.get_value('item_name', 'N/A')}")
        print(f"  수량: {order_data.get_value('item_qty', 'N/A')}")
        print(f"  단가: {order_data.get_value('sales_unit_price', 0):,.0f}")

    # 3. 문서 생성 (서비스 사용)
    gen_result = service.generate_ts(advance_id, doc_type='ADV')

    # 4. 결과 처리
    if gen_result.success:
        print(f"  -> 선수금 거래명세표 생성 완료: {gen_result.output_file.name}")
        return True
    else:
        if gen_result.status == GenerationStatus.FILE_ERROR:
            print(f"  [오류] {gen_result.errors[0] if gen_result.errors else gen_result.message}")
        else:
            print(f"  [오류] {gen_result.message}")
        return False


def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성

    Returns:
        설정된 ArgumentParser
    """
    description = """
거래명세표 생성기 (국내 전용)
============================

NOAH_SO_PO_DN.xlsx의 DN_국내 또는 PMT_국내 시트에서 데이터를 읽어
거래명세표를 자동 생성합니다.

지원 ID 유형:
  - DN_ID (예: DN-2026-0001)    : 납품 거래명세표
  - 선수금_ID (예: ADV_2026-0001) : 선수금 거래명세표
"""

    epilog = """
사용 예시:
  python create_ts.py DN-2026-0001              # 납품 거래명세표 1건
  python create_ts.py ADV_2026-0001             # 선수금 거래명세표 1건
  python create_ts.py DN-2026-0001 DN-2026-0002 # 여러 건 동시 생성
  python create_ts.py DN-2026-0001 ADV_2026-0001  # DN + 선수금 혼합

인자 없이 실행하면 사용 가능한 ID 목록을 표시합니다.
"""

    parser = argparse.ArgumentParser(
        prog='create_ts',
        description=description,
        epilog=epilog,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        'doc_ids',
        nargs='*',
        metavar='ID',
        help='DN_ID (DN-XXXX-XXXX) 또는 선수금_ID (ADV_XXXX-XXXX)',
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
        df_dn = load_dn_data()
        df_pmt = load_pmt_data()
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        return 1

    # 인자 없으면 도움말 + 사용 가능한 ID 출력
    if not args.doc_ids:
        parser.print_help()
        print_available_ids(df_dn, df_pmt)
        return 0

    print(f"DN: {len(df_dn)}건, PMT: {len(df_pmt)}건 로드 완료")

    # 각 ID에 대해 거래명세표 생성
    success_count = 0
    for doc_id in args.doc_ids:
        id_type = detect_id_type(doc_id)

        if id_type == 'DN':
            if generate_ts_from_dn(doc_id, df_dn):
                success_count += 1
        else:  # ADV
            if generate_ts_from_adv(doc_id):
                success_count += 1

    # 결과 출력
    print(f"\n{'=' * 50}")
    print(f"완료: {success_count}/{len(args.doc_ids)}건 거래명세표 생성")
    print(f"출력 폴더: {TS_OUTPUT_DIR}")
    print('=' * 50)

    return 0 if success_count == len(args.doc_ids) else 1


if __name__ == "__main__":
    sys.exit(main())
