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
    find_dn_data,
    load_so_for_advance,
    get_value,
)
from po_generator.ts_generator import create_ts_xlwings
from po_generator.cli_common import validate_output_path, generate_output_filename
from po_generator.logging_config import setup_logging

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

    Args:
        dn_id: DN_ID
        df_dn: DN 데이터

    Returns:
        성공 여부
    """
    print(f"\n{'=' * 50}")
    print(f"거래명세표 생성 (납품): {dn_id}")
    print('=' * 50)

    # 1. DN 데이터 검색
    dn_result = find_dn_data(df_dn, dn_id)
    if dn_result is None:
        print(f"  [오류] '{dn_id}'를 찾을 수 없습니다.")
        return False

    # 2. 다중/단일 아이템 처리
    if isinstance(dn_result, pd.DataFrame):
        items_df = dn_result
        dn_data = items_df.iloc[0]
        num_items = len(items_df)
        print(f"  [다중 아이템] {num_items}개 아이템 발견")
        for idx, (_, item) in enumerate(items_df.iterrows()):
            item_name = get_value(item, 'item_name', 'N/A')
            item_qty = get_value(item, 'item_qty', 'N/A')
            print(f"    {idx + 1}. {item_name} x {item_qty}")
    else:
        items_df = None
        dn_data = dn_result
        num_items = 1

    # 3. 기본 정보 출력
    print(f"  고객: {get_value(dn_data, 'customer_name', 'N/A')}")
    if num_items == 1:
        print(f"  품목: {get_value(dn_data, 'item_name', 'N/A')}")
        print(f"  수량: {get_value(dn_data, 'item_qty', 'N/A')}")
        print(f"  단가: {get_value(dn_data, 'sales_unit_price', 'N/A'):,}")

    # 4. 템플릿 확인
    if not TS_TEMPLATE_FILE.exists():
        print(f"  [오류] 템플릿 파일이 없습니다: {TS_TEMPLATE_FILE}")
        return False

    # 5. 출력 디렉토리 생성
    TS_OUTPUT_DIR.mkdir(exist_ok=True)

    # 6. 파일명 생성 및 경로 검증
    customer_name = get_value(dn_data, 'customer_name', 'Unknown')
    output_file = generate_output_filename("TS", dn_id, customer_name, TS_OUTPUT_DIR)

    if not validate_output_path(output_file, TS_OUTPUT_DIR):
        return False

    # 7. 거래명세표 생성 (xlwings)
    try:
        create_ts_xlwings(
            template_path=TS_TEMPLATE_FILE,
            output_path=output_file,
            order_data=dn_data,
            items_df=items_df,
            doc_type='DN',
        )
        print(f"  -> 거래명세표 생성 완료: {output_file.name}")

    except FileNotFoundError as e:
        print(f"  [오류] {e}")
        return False
    except PermissionError:
        print(f"  [오류] 파일 저장 실패: {output_file.name} (파일이 열려있거나 권한 없음)")
        return False
    except Exception as e:
        print(f"  [오류] 거래명세표 생성 실패: {e}")
        return False

    return True


def generate_ts_from_adv(advance_id: str) -> bool:
    """선수금 거래명세표 생성 (SO_국내 데이터 사용)

    PMT_국내에서 선수금_ID → SO_ID 매핑을 찾고,
    SO_국내에서 품목/금액 정보를 가져와 거래명세표를 생성합니다.

    Args:
        advance_id: 선수금_ID

    Returns:
        성공 여부
    """
    print(f"\n{'=' * 50}")
    print(f"거래명세표 생성 (선수금): {advance_id}")
    print('=' * 50)

    # 1. SO 데이터 로드 (선수금_ID -> SO_ID -> SO 아이템들)
    result = load_so_for_advance(advance_id)
    if result is None:
        print(f"  [오류] '{advance_id}'를 찾을 수 없습니다.")
        return False

    pmt_data, so_items = result
    num_items = len(so_items)

    # 2. 기본 정보 출력
    first_item = so_items.iloc[0]
    customer_name = get_value(first_item, 'customer_name', 'N/A')
    print(f"  고객: {customer_name}")
    print(f"  SO_ID: {get_value(first_item, 'so_id', 'N/A')}")

    if num_items > 1:
        print(f"  [다중 아이템] {num_items}개 아이템 발견")
        for idx, (_, item) in enumerate(so_items.iterrows()):
            item_name = get_value(item, 'item_name', 'N/A')
            item_qty = get_value(item, 'item_qty', 'N/A')
            unit_price = get_value(item, 'sales_unit_price', 0)
            print(f"    {idx + 1}. {item_name} x {item_qty} @ {unit_price:,.0f}")
    else:
        print(f"  품목: {get_value(first_item, 'item_name', 'N/A')}")
        print(f"  수량: {get_value(first_item, 'item_qty', 'N/A')}")
        print(f"  단가: {get_value(first_item, 'sales_unit_price', 0):,.0f}")

    # 3. 템플릿 확인
    if not TS_TEMPLATE_FILE.exists():
        print(f"  [오류] 템플릿 파일이 없습니다: {TS_TEMPLATE_FILE}")
        return False

    # 4. 출력 디렉토리 생성
    TS_OUTPUT_DIR.mkdir(exist_ok=True)

    # 5. 파일명 생성 및 경로 검증
    output_file = generate_output_filename("TS_ADV", advance_id, customer_name, TS_OUTPUT_DIR)

    if not validate_output_path(output_file, TS_OUTPUT_DIR):
        return False

    # 6. 거래명세표 생성 (xlwings) - DN과 동일한 로직, doc_type='ADV'
    try:
        create_ts_xlwings(
            template_path=TS_TEMPLATE_FILE,
            output_path=output_file,
            order_data=first_item,
            items_df=so_items if num_items > 1 else None,
            doc_type='ADV',
        )
        print(f"  -> 선수금 거래명세표 생성 완료: {output_file.name}")

    except FileNotFoundError as e:
        print(f"  [오류] {e}")
        return False
    except PermissionError:
        print(f"  [오류] 파일 저장 실패: {output_file.name} (파일이 열려있거나 권한 없음)")
        return False
    except Exception as e:
        print(f"  [오류] 거래명세표 생성 실패: {e}")
        return False

    return True


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
