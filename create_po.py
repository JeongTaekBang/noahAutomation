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
from po_generator.utils import (
    load_noah_po_lists,
    find_order_data,
    get_safe_value,
)
from po_generator.validators import validate_order_data, validate_multiple_items
from po_generator.history import check_duplicate_order, save_to_history, get_all_history, get_history_count, get_current_month_info
from po_generator.excel_generator import create_po_workbook
from po_generator.cli_common import print_available_orders, validate_output_path, generate_output_filename

# 경고 필터링 (openpyxl/pandas 관련 경고만 선택적으로 무시)
import warnings
# openpyxl의 스타일 관련 UserWarning 무시 (예: 알 수 없는 확장 기능)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
# pandas의 FutureWarning 무시 (버전 호환성 관련)
warnings.filterwarnings('ignore', category=FutureWarning, module='pandas')


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

    # 7. 파일명 생성 및 경로 검증
    customer_name_raw = get_safe_value(order_data, 'Customer name', 'Unknown')
    output_file = generate_output_filename("PO", order_no, customer_name_raw, OUTPUT_DIR)

    if not validate_output_path(output_file, OUTPUT_DIR):
        return False

    # 8. 워크북 생성 및 저장 (트랜잭션, 템플릿 기반)
    try:
        wb = create_po_workbook(order_data, items_df)

        # 9. 저장
        wb.save(output_file)
        print(f"  -> 발주서 생성 완료: {output_file.name}")

    except PermissionError:
        print(f"  [오류] 파일 저장 실패: {output_file.name} (파일이 열려있거나 권한 없음)")
        return False
    except Exception as e:
        print(f"  [오류] 발주서 생성 실패: {e}")
        # 롤백: 부분적으로 생성된 파일 삭제
        if output_file.exists():
            try:
                output_file.unlink()
                print("  -> 생성된 파일 롤백 완료")
            except (IOError, OSError, PermissionError) as rollback_error:
                logger.warning(f"롤백 실패: {rollback_error}")
        return False

    # 10. 이력 저장 (발주서 파일에서 데이터 추출)
    order_no_val = get_safe_value(order_data, 'RCK Order no.')
    customer_name_val = get_safe_value(order_data, 'Customer name')
    history_saved = save_to_history(output_file, order_no_val, customer_name_val)

    if not history_saved:
        print("  [주의] 이력 저장 실패 - 발주서는 정상 생성됨")

    return True


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
