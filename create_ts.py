#!/usr/bin/env python
"""
거래명세표(Transaction Statement) 생성기
=========================================

RCK Order No.를 입력하면 NOAH_PO_Lists.xlsx에서 해당 데이터를 읽어
자동으로 거래명세표를 생성합니다.

사용법:
    python create_ts.py ND-0001
    python create_ts.py ND-0001 ND-0002 ND-0003  # 여러 건 동시 생성
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

from po_generator.config import (
    TS_OUTPUT_DIR,
    TS_TEMPLATE_FILE,
    ORDER_LIST_DISPLAY_LIMIT,
)
from po_generator.utils import (
    load_noah_po_lists,
    find_order_data,
    get_safe_value,
)
from po_generator.history import sanitize_filename
from po_generator.ts_generator import create_ts_xlwings


def generate_ts(order_no: str, df: pd.DataFrame) -> bool:
    """거래명세표 생성 메인 함수

    Args:
        order_no: RCK Order No.
        df: 전체 주문 데이터

    Returns:
        성공 여부
    """
    print(f"\n{'=' * 50}")
    print(f"거래명세표 생성: {order_no}")
    print('=' * 50)

    # 1. 주문 데이터 검색
    order_result = find_order_data(df, order_no)
    if order_result is None:
        print(f"  [오류] '{order_no}'를 찾을 수 없습니다.")
        return False

    # 2. 다중/단일 아이템 처리
    if isinstance(order_result, pd.DataFrame):
        items_df = order_result
        order_data = items_df.iloc[0]
        num_items = len(items_df)
        print(f"  [다중 아이템] {num_items}개 아이템 발견")
        for idx, (_, item) in enumerate(items_df.iterrows()):
            item_name = get_safe_value(item, 'Item name', 'N/A')
            item_qty = get_safe_value(item, 'Item qty', 'N/A')
            print(f"    {idx + 1}. {item_name} x {item_qty}")
    else:
        items_df = None
        order_data = order_result
        num_items = 1

    # 3. 기본 정보 출력
    print(f"  고객: {get_safe_value(order_data, 'Customer name', 'N/A')}")
    if num_items == 1:
        print(f"  품목: {get_safe_value(order_data, 'Item name', 'N/A')}")
        print(f"  수량: {get_safe_value(order_data, 'Item qty', 'N/A')}")
    print(f"  시트: {get_safe_value(order_data, '_시트구분', 'N/A')}")

    # 4. 템플릿 확인
    if not TS_TEMPLATE_FILE.exists():
        print(f"  [오류] 템플릿 파일이 없습니다: {TS_TEMPLATE_FILE}")
        return False

    # 5. 출력 디렉토리 생성
    TS_OUTPUT_DIR.mkdir(exist_ok=True)

    # 6. 파일명 생성
    today = datetime.now().strftime("%y%m%d")
    customer_name_raw = get_safe_value(order_data, 'Customer name', 'Unknown')
    customer_name_safe = sanitize_filename(customer_name_raw)
    order_no_safe = sanitize_filename(order_no)
    output_file = TS_OUTPUT_DIR / f"TS_{order_no_safe}_{customer_name_safe}_{today}.xlsx"

    # Path Traversal 방지
    try:
        if not output_file.resolve().is_relative_to(TS_OUTPUT_DIR.resolve()):
            print(f"  [오류] 잘못된 파일 경로: {output_file}")
            return False
    except ValueError:
        if str(TS_OUTPUT_DIR.resolve()) not in str(output_file.resolve()):
            print(f"  [오류] 잘못된 파일 경로: {output_file}")
            return False

    # 7. 거래명세표 생성 (xlwings)
    try:
        create_ts_xlwings(
            template_path=TS_TEMPLATE_FILE,
            output_path=output_file,
            order_data=order_data,
            items_df=items_df,
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


def print_available_orders(df: pd.DataFrame, limit: int = ORDER_LIST_DISPLAY_LIMIT) -> None:
    """사용 가능한 주문번호 목록 출력

    Args:
        df: 주문 데이터
        limit: 출력 제한 수
    """
    orders = df['RCK Order no.'].dropna().unique().tolist()
    print("\n사용 가능한 RCK Order No. 목록:")
    for order in orders[:limit]:
        print(f"  - {order}")
    if len(orders) > limit:
        print(f"  ... 외 {len(orders) - limit}건")


def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성

    Returns:
        설정된 ArgumentParser
    """
    parser = argparse.ArgumentParser(
        prog='create_ts',
        description='거래명세표 생성기 - RCK Order No.로 거래명세표 자동 생성',
        epilog='예시: python create_ts.py ND-0001 ND-0002',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        'order_numbers',
        nargs='*',
        metavar='ORDER_NO',
        help='생성할 RCK Order No. (여러 개 가능)',
    )

    return parser


def main() -> int:
    """메인 함수

    Returns:
        종료 코드 (0: 성공, 1: 실패)
    """
    parser = create_argument_parser()
    args = parser.parse_args()

    # 인자 없으면 도움말 + 사용 가능한 주문번호 출력
    if not args.order_numbers:
        parser.print_help()

        try:
            df = load_noah_po_lists()
            print_available_orders(df)
        except FileNotFoundError as e:
            print(f"\n[오류] {e}")
            return 1

        return 0

    # 데이터 로드
    print("NOAH_PO_Lists.xlsx 로딩 중...")
    try:
        df = load_noah_po_lists()
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        return 1

    print(f"총 {len(df)}건의 주문 데이터 로드 완료")

    # 각 Order No.에 대해 거래명세표 생성
    success_count = 0
    for order_no in args.order_numbers:
        if generate_ts(order_no, df):
            success_count += 1

    # 결과 출력
    print(f"\n{'=' * 50}")
    print(f"완료: {success_count}/{len(args.order_numbers)}건 거래명세표 생성")
    print(f"출력 폴더: {TS_OUTPUT_DIR}")
    print('=' * 50)

    return 0 if success_count == len(args.order_numbers) else 1


if __name__ == "__main__":
    sys.exit(main())
