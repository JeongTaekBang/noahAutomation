#!/usr/bin/env python
"""
Order Confirmation 생성기
=========================

SO_ID를 입력하면 NOAH_SO_PO_DN.xlsx의 SO_해외 시트에서 해당 데이터를 읽어
자동으로 Order Confirmation을 생성합니다.
Dispatch date는 SO_해외의 'EXW NOAH' 컬럼 값을 사용합니다.

생성 후 "이메일 발송?"을 묻고, y면 **영문** 메일 초안을 만듭니다 (해외 고객 전용 문서).

사용법:
    python create_oc.py SOO-2026-0001              # 단일 생성
    python create_oc.py SOO-2026-0001 SOO-2026-0002 # 여러 건 동시 생성
    python create_oc.py SOO-2026-0001 --mail       # 확인 없이 메일 초안
    python create_oc.py SOO-2026-0001 --no-mail    # 묻지 않고 문서만 (배치용)
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

import pandas as pd

from po_generator.config import (
    CUSTOMER_EXPORT_SHEET,
    OC_MAIL_ATTACH_FORMAT,
    OC_MAIL_BODY,
    OC_MAIL_CC,
    OC_MAIL_SUBJECT,
    OC_OUTPUT_DIR,
    OC_TEMPLATE_FILE,
)
from po_generator.utils import (
    load_customer_overseas,
    load_so_export_with_customer,
    get_value,
)
from po_generator.logging_config import setup_logging
from po_generator.mail_cli import (
    MailOptions,
    add_mail_arguments,
    collect_customer_po,
    confirm_recipient,
    format_mail_date,
    prepare_mail_options as _prepare_mail_options,
    report_mail_result,
)
from po_generator.mailer import (
    MailConfigError,
    create_document_mail,
    find_recipient_overseas_for_order,
)
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


def prepare_mail_options(args: argparse.Namespace) -> MailOptions:
    """CLI 인자로 OC 메일 옵션 구성

    판정 규칙은 `mail_cli.prepare_mail_options()`가 소유하고, 여기서는 OC 상수만
    넘깁니다. 국내 문서와 달리 **수신자 마스터가 `Customer_해외`** 입니다.

    Args:
        args: 파싱된 CLI 인자

    Returns:
        MailOptions
    """
    return _prepare_mail_options(
        args,
        attach_format=OC_MAIL_ATTACH_FORMAT,
        fixed_cc=OC_MAIL_CC,
        loader=load_customer_overseas,
        sheet_label=CUSTOMER_EXPORT_SHEET,
    )


def _mail_oc(
    order_data: pd.Series,
    output_file: Path,
    so_id: str,
    opts: MailOptions,
    items_df: pd.DataFrame | None = None,
) -> bool:
    """생성된 Order Confirmation을 메일로 발송/초안 생성

    수신자 확인 관문(조회 → 표시 → y/N)은 `mail_cli.confirm_recipient()`가 소유한다.
    메일 실패는 OC 생성 성공을 뒤엎지 않습니다 (경고만 출력).

    수신자는 **고객코드**로 `Customer_해외`에서 찾습니다 — 거래명세표가 쓰는
    사업자번호가 아니다. 어느 별칭 키를 읽는지는 mailer의
    `find_recipient_overseas_for_order()`가 안다.

    Args:
        order_data: 주문 데이터 (고객코드/고객명 포함)
        output_file: 생성된 OC 경로
        so_id: SO_ID (= O.C. No)
        opts: 메일 옵션
        items_df: 문서에 실린 전체 아이템 (발주번호 수집용)

    Returns:
        메일 생성/발송 성공 여부
    """
    if not opts.enabled:
        return True

    customer_code = get_value(order_data, 'customer_code', '')
    customer_name = str(get_value(order_data, 'customer_name', ''))
    customer_po = collect_customer_po(order_data, items_df)
    date_str = format_mail_date(None)  # OC 발행일 = 오늘

    recipient = confirm_recipient(
        opts,
        lambda df: find_recipient_overseas_for_order(order_data, df),
        missing_lines=[
            f"  [메일 생략] 수신자 미등록 — {customer_name} / {customer_code or '(고객코드 없음)'}",
            f"             {opts.sheet_label} 시트에 해당 고객코드의 이메일을 입력하세요.",
        ],
        info_lines=[f"  발주번호: {customer_po}"],
    )
    if recipient is None:
        return False

    try:
        result = create_document_mail(
            xlsx_path=output_file,
            recipient=recipient,
            doc_id=so_id,
            subject_template=OC_MAIL_SUBJECT,
            body_template=OC_MAIL_BODY,
            date_str=date_str,
            send=opts.send,
            attach_format=OC_MAIL_ATTACH_FORMAT,
            backend=opts.backend,
            customer_po=customer_po,
            doc_label='Order Confirmation',
            draft_prefix='oc_draft',
        )
    except MailConfigError as e:
        print(f"  [메일 오류] {e}")
        return False

    return report_mail_result(result, want_send=opts.send)


def generate_oc(so_id: str, mail_opts: MailOptions | None = None) -> bool:
    """Order Confirmation 생성

    Args:
        so_id: SO_ID
        mail_opts: 메일 발송 옵션 (None이면 발송 안 함)

    Returns:
        성공 여부 (메일 실패는 성공 여부에 영향 없음)
    """
    mail_opts = mail_opts or MailOptions.disabled()
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
        _mail_oc(order_data.first_item, result.output_file, so_id, mail_opts,
                 items_df=order_data.items_df)
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

메일 발송 (해외 고객 — 본문은 영문):
  생성이 끝나면 받는사람/참조를 보여주고 "이메일을 발송하시겠습니까? [y/N]"을 묻습니다.
  y를 누르면 PDF를 첨부한 영문 메일 초안이 열리고, 최종 [보내기]는 직접 누릅니다.
  수신자는 **고객코드**로 Customer_해외에서 조회합니다 (국내 문서의 사업자번호가 아님).
  고정 참조(CC)는 user_settings.py의 OC_MAIL_CC로 관리합니다.

  python create_oc.py SOO-2026-0001 --mail     # 확인 없이 메일 초안
  python create_oc.py SOO-2026-0001 --send     # 확인 없이 즉시 발송
  python create_oc.py SOO-2026-0001 --no-mail  # 묻지 않고 문서만

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

    add_mail_arguments(parser, 'Order Confirmation')

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

    # 메일 옵션 (수신자 마스터는 실제 발송 시점에 지연 로딩)
    mail_opts = prepare_mail_options(args)

    success_count = 0
    for so_id in args.so_ids:
        if generate_oc(so_id, mail_opts):
            success_count += 1

    print(f"\n{'=' * 60}")
    print(f"완료: {success_count}/{len(args.so_ids)}건 Order Confirmation 생성")
    print(f"출력 폴더: {OC_OUTPUT_DIR}")
    print('=' * 60)

    return 0 if success_count == len(args.so_ids) else 1


if __name__ == "__main__":
    sys.exit(main())
