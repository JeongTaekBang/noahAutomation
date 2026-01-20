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
    python create_po.py --history                # 발주 이력 조회
    python create_po.py --history --export       # 이력을 Excel로 내보내기

검증 항목:
    - 필수 필드: Customer name, Customer PO, Item qty, Model, ICO Unit
    - ICO Unit: 0 또는 음수인 경우 오류
    - 납기일: 과거이면 오류, 7일 이내면 경고
"""

from __future__ import annotations

import argparse
import sys
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd

from po_generator.config import (
    OUTPUT_DIR,
    ORDER_LIST_DISPLAY_LIMIT,
    HISTORY_CUSTOMER_DISPLAY_LENGTH,
    HISTORY_DESC_DISPLAY_LENGTH,
    HISTORY_DATE_DISPLAY_LENGTH,
)
from po_generator.utils import get_value
from po_generator.validators import validate_order_data, validate_multiple_items
from po_generator.history import check_duplicate_order, save_to_history, get_all_history, get_history_count, get_current_month_info
from po_generator.excel_generator import create_po_workbook
from po_generator.cli_common import print_available_orders, validate_output_path, generate_output_filename
from po_generator.logging_config import setup_logging
from po_generator.services import DocumentService, GenerationStatus

# 경고 필터링 (openpyxl/pandas 관련 경고만 선택적으로 무시)
import warnings
# openpyxl의 스타일 관련 UserWarning 무시 (예: 알 수 없는 확장 기능)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
# pandas의 FutureWarning 무시 (버전 호환성 관련)
warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')

logger = logging.getLogger(__name__)


def generate_po(order_no: str, df: pd.DataFrame, force: bool = False) -> bool:
    """발주서 생성 메인 함수

    DocumentService를 사용하여 발주서를 생성합니다.
    사용자 상호작용(중복 확인, 검증 오류 확인)은 CLI에서 처리합니다.

    Args:
        order_no: RCK Order No.
        df: 전체 주문 데이터 (하위 호환용, 실제로는 사용하지 않음)
        force: 강제 생성 여부

    Returns:
        성공 여부
    """
    print(f"\n{'=' * 50}")
    print(f"발주서 생성: {order_no}")
    print('=' * 50)

    service = DocumentService()

    # 1. 중복 발주 체크 (CLI에서 사용자 상호작용 필요)
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

    # 2. 데이터 검색 및 정보 출력
    order_data = service.finder.find_po(order_no)
    if order_data is None:
        print(f"  [오류] '{order_no}'를 찾을 수 없습니다.")
        return False

    # 3. 기본 정보 출력
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
    print(f"  시트: {order_data.get_value('sheet_type', 'N/A')}")

    # 4. 데이터 검증 (CLI에서 사용자 상호작용 필요)
    if order_data.is_multi_item:
        validation = validate_multiple_items(order_data.items_df)
    else:
        validation = validate_order_data(order_data.first_item)

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

    # 5. 문서 생성 (서비스 사용, 중복 체크/검증은 이미 완료)
    result = service.generate_po(order_no, force=True, skip_history=False)

    # 6. 결과 처리
    if result.success:
        print(f"  -> 발주서 생성 완료: {result.output_file.name}")

        # 경고 출력 (이미 위에서 출력했으므로 중복 방지)
        # for warn in result.warnings:
        #     print(f"  [주의] {warn}")

        return True
    else:
        if result.status == GenerationStatus.FILE_ERROR:
            print(f"  [오류] 파일 저장 실패: {result.errors[0] if result.errors else '알 수 없는 오류'}")
        else:
            print(f"  [오류] {result.message}")
        return False


def show_history(export: bool = False) -> int:
    """발주 이력 조회 및 내보내기 (현재 월만)

    Args:
        export: Excel 파일로 내보내기 여부

    Returns:
        종료 코드 (0: 성공, 1: 실패)
    """
    month_str, month_dir = get_current_month_info()

    print("\n" + "=" * 60)
    print(f"  발주 이력 조회 ({month_str})")
    print("=" * 60)

    count = get_history_count()
    if count == 0:
        print(f"\n  {month_str} 발주 이력이 없습니다.")
        print(f"  폴더: {month_dir}")
        return 0

    print(f"\n  {month_str}: 총 {count}건의 발주 이력")

    df = get_all_history()
    if df.empty:
        print("  이력 데이터를 불러올 수 없습니다.")
        return 1

    print("\n  이력 목록:")
    print("-" * 60)

    # 전체 표시 (월별이므로 건수가 적음)
    for idx, row in df.iterrows():
        order_no = row.get('RCK Order no.', 'N/A')
        customer = str(row.get('Customer name', 'N/A'))[:HISTORY_CUSTOMER_DISPLAY_LENGTH]
        desc = str(row.get('Description', row.get('Model', 'N/A')))[:HISTORY_DESC_DISPLAY_LENGTH]
        created = str(row.get('생성일시', 'N/A'))[:HISTORY_DATE_DISPLAY_LENGTH]
        total = row.get('Total net amount', row.get('Order Total', ''))
        total_str = f"{int(total):,}" if pd.notna(total) and total != '' else ''
        print(f"  {created} | {order_no} | {customer} | {desc} | {total_str}")

    print("-" * 60)

    # Excel 내보내기 (po_history 루트에 저장)
    if export:
        from po_generator.config import HISTORY_DIR
        HISTORY_DIR.mkdir(parents=True, exist_ok=True)
        today = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_file = HISTORY_DIR / f"발주이력_{month_str.replace(' ', '')}_{today}.xlsx"
        df.to_excel(export_file, index=False)
        print(f"\n  -> Excel 내보내기 완료: {export_file.name}")
        print(f"     저장 위치: {HISTORY_DIR}")
        print(f"     저장된 컬럼: {len(df.columns)}개")
        print(f"     저장된 행: {len(df)}건")

    return 0


def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성

    Returns:
        설정된 ArgumentParser
    """
    parser = argparse.ArgumentParser(
        prog='create_po',
        description='NOAH Purchase Order Auto-Generator - RCK Order No.로 발주서 자동 생성',
        epilog='예시: python create_po.py ND-0001 ND-0002 --force',
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        'order_numbers',
        nargs='*',
        metavar='ORDER_NO',
        help='생성할 RCK Order No. (여러 개 가능)',
    )

    parser.add_argument(
        '-f', '--force',
        action='store_true',
        help='중복 발주 및 검증 오류 무시하고 강제 생성',
    )

    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='상세 로그 출력',
    )

    parser.add_argument(
        '--history',
        action='store_true',
        help='현재 월 발주 이력 조회',
    )

    parser.add_argument(
        '--export',
        action='store_true',
        help='이력을 Excel 파일로 내보내기 (--history와 함께 사용)',
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

    # 이력 조회 모드
    if args.history:
        return show_history(export=args.export)

    # 인자 없으면 도움말 + 사용 가능한 주문번호 출력
    if not args.order_numbers:
        parser.print_help()

        try:
            service = DocumentService()
            df = service.finder.load_po_data()
            print_available_orders(df)
        except FileNotFoundError as e:
            print(f"\n[오류] {e}")
            return 1

        return 0

    # 데이터 로드 (서비스에서 자동 로드하므로 메시지만 출력)
    print("NOAH_PO_Lists.xlsx 로딩 중...")
    try:
        service = DocumentService()
        df = service.finder.load_po_data()
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        return 1

    print(f"총 {len(df)}건의 주문 데이터 로드 완료")

    # 각 Order No.에 대해 발주서 생성
    success_count = 0
    for order_no in args.order_numbers:
        if generate_po(order_no, df, args.force):
            success_count += 1

    # 결과 출력
    print(f"\n{'=' * 50}")
    print(f"완료: {success_count}/{len(args.order_numbers)}건 발주서 생성")
    print(f"출력 폴더: {OUTPUT_DIR}")
    print('=' * 50)

    return 0 if success_count == len(args.order_numbers) else 1


if __name__ == "__main__":
    sys.exit(main())
