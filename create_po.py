#!/usr/bin/env python
"""
NOAH Purchase Order Auto-Generator
===================================

RCK Order No.를 입력하면 NOAH_PO_Lists.xlsx에서 해당 데이터를 읽어
자동으로 발주서(Purchase Order + Description)를 생성합니다.

사용법:
    python create_po.py ND-0001
    python create_po.py ND-0001 ND-0002 ND-0003  # 여러 건 동시 생성
    python create_po.py ND-0001 --force          # 중복 발주/검증 오류 무시

검증 항목:
    - 필수 필드: Customer name, Customer PO, Item qty, Model, ICO Unit
    - ICO Unit: 0 또는 음수인 경우 오류
    - 납기일: 과거이면 오류, 7일 이내면 경고
"""

from __future__ import annotations

import sys
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from po_generator.config import OUTPUT_DIR
from po_generator.utils import (
    load_noah_po_lists,
    find_order_data,
    get_safe_value,
)
from po_generator.validators import validate_order_data, validate_multiple_items
from po_generator.history import check_duplicate_order, save_to_history
from po_generator.excel_generator import create_purchase_order, create_description_sheet

# 경고 필터링
import warnings
warnings.filterwarnings('ignore', category=UserWarning)
warnings.filterwarnings('ignore', category=FutureWarning)


def setup_logging(verbose: bool = False) -> None:
    """로깅 설정

    Args:
        verbose: 상세 로깅 여부
    """
    level = logging.DEBUG if verbose else logging.INFO

    # 콘솔 핸들러
    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    console_format = logging.Formatter('%(message)s')
    console_handler.setFormatter(console_format)

    # 루트 로거 설정
    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    root_logger.addHandler(console_handler)

    # po_generator 로거 설정
    pkg_logger = logging.getLogger('po_generator')
    pkg_logger.setLevel(level)


logger = logging.getLogger(__name__)


def generate_po(order_no: str, df: pd.DataFrame, force: bool = False) -> bool:
    """발주서 생성 메인 함수

    Args:
        order_no: RCK Order No.
        df: 전체 주문 데이터
        force: 강제 생성 여부

    Returns:
        성공 여부
    """
    print(f"\n{'=' * 50}")
    print(f"발주서 생성: {order_no}")
    print('=' * 50)

    # 1. 중복 발주 체크
    if not force:
        dup_info = check_duplicate_order(order_no)
        if dup_info:
            print(f"  [경고] 이미 발주된 건입니다!")
            print(f"         이전 발주일: {dup_info['생성일시']}")
            print(f"         이전 파일: {Path(dup_info['생성파일']).name}")

            response = input("  계속 진행하시겠습니까? (Y/N): ").strip().upper()
            if response != 'Y':
                print("  -> 발주 취소됨")
                return False

    # 2. 주문 데이터 검색
    order_result = find_order_data(df, order_no)
    if order_result is None:
        print(f"  [오류] '{order_no}'를 찾을 수 없습니다.")
        return False

    # 3. 다중/단일 아이템 처리
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

    # 4. 기본 정보 출력
    print(f"  고객: {get_safe_value(order_data, 'Customer name', 'N/A')}")
    if num_items == 1:
        print(f"  품목: {get_safe_value(order_data, 'Item name', 'N/A')}")
        print(f"  수량: {get_safe_value(order_data, 'Item qty', 'N/A')}")
    print(f"  시트: {get_safe_value(order_data, '_시트구분', 'N/A')}")

    # 5. 데이터 검증
    if items_df is not None:
        validation = validate_multiple_items(items_df)
    else:
        validation = validate_order_data(order_data)

    # 경고 출력
    for warn in validation.warnings:
        print(f"  [주의] {warn}")

    # 오류 출력 및 처리
    if validation.has_errors:
        for err in validation.errors:
            print(f"  [오류] {err}")

        if not force:
            response = input("  오류가 있습니다. 그래도 진행하시겠습니까? (Y/N): ").strip().upper()
            if response != 'Y':
                print("  -> 발주 취소됨")
                return False
        else:
            print("  -> --force 옵션으로 오류 무시하고 진행")

    # 6. 출력 디렉토리 생성
    OUTPUT_DIR.mkdir(exist_ok=True)

    # 7. 파일명 생성
    today = datetime.now().strftime("%y%m%d")
    customer_name_safe = get_safe_value(order_data, 'Customer name', 'Unknown')[:10]
    customer_name_safe = customer_name_safe.replace(' ', '_')
    output_file = OUTPUT_DIR / f"PO_{order_no}_{customer_name_safe}_{today}.xlsx"

    # 8. 워크북 생성
    wb = Workbook()

    # Purchase Order 시트
    ws_po = wb.active
    ws_po.title = "Purchase Order"
    create_purchase_order(ws_po, order_data, items_df)

    # Description 시트
    ws_desc = wb.create_sheet("Description")
    create_description_sheet(ws_desc, order_data, items_df)

    # 9. 저장
    wb.save(output_file)
    print(f"  -> 발주서 생성 완료: {output_file.name}")

    # 10. 이력 저장
    save_to_history(order_data, output_file)

    return True


def print_available_orders(df: pd.DataFrame, limit: int = 20) -> None:
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


def main() -> int:
    """메인 함수

    Returns:
        종료 코드 (0: 성공, 1: 실패)
    """
    # 로깅 설정
    setup_logging(verbose='--verbose' in sys.argv or '-v' in sys.argv)

    # 옵션 파싱
    force = '--force' in sys.argv
    args = [
        arg for arg in sys.argv[1:]
        if arg not in ('--force', '--verbose', '-v')
    ]

    # 인자 없으면 도움말 출력
    if len(args) < 1:
        print(__doc__)

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

    # 각 Order No.에 대해 발주서 생성
    success_count = 0
    for order_no in args:
        if generate_po(order_no, df, force):
            success_count += 1

    # 결과 출력
    print(f"\n{'=' * 50}")
    print(f"완료: {success_count}/{len(args)}건 발주서 생성")
    print(f"출력 폴더: {OUTPUT_DIR}")
    print('=' * 50)

    return 0 if success_count == len(args) else 1


if __name__ == "__main__":
    sys.exit(main())
