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

생성 후 "이메일을 발송하시겠습니까? [y/N]"을 묻고, y면 Outlook 메일 창을 띄웁니다.
    python create_ts.py DN-2026-0001 --mail       # 확인 없이 바로 Outlook 창
    python create_ts.py DN-2026-0001 --no-mail    # 묻지 않고 문서만
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

from po_generator.config import (
    CUSTOMER_DOMESTIC_SHEET,
    TS_MAIL_ATTACH_FORMAT,
    TS_MAIL_CC,
    TS_OUTPUT_DIR,
    TS_TEMPLATE_FILE,
)
from po_generator.utils import (
    load_dn_data,
    load_pmt_data,
    get_value,
    resolve_column,
)
from po_generator.ts_generator import create_ts_xlwings
from po_generator.cli_common import validate_output_path, generate_output_filename
from po_generator.logging_config import setup_logging
from po_generator.mailer import (
    NO_VALUE,
    MailBackend,
    MailConfigError,
    create_ts_mail,
    find_recipient_for_order,
    resolve_backend,
)
# 메일 CLI 배선은 delivery_status.py와 공유한다 (po_generator/mail_cli.py).
# 여기서 다시 노출하는 이유는 `from create_ts import MailMode` 하는 기존 호출부·테스트를
# 그대로 두기 위해서다.
from po_generator.mail_cli import (  # noqa: F401  (재노출)
    MailMode,
    MailOptions,
    add_mail_arguments,
    confirm as _confirm,
    report_mail_result,
    resolve_mail_mode,
    show_recipient,
)
from po_generator.services import DocumentService, GenerationStatus

logger = logging.getLogger(__name__)


def _format_mail_date(value: object) -> str:
    """메일 제목/본문에 쓸 날짜 문자열

    출고일이 비어 있거나(선수금 문서) 파싱되지 않으면 오늘 날짜를 씁니다.
    빈 값은 None/NaT/NaN 모두로 들어올 수 있어 한 곳에서 흡수합니다.

    Args:
        value: 출고일 값

    Returns:
        'YYYY-MM-DD'
    """
    try:
        stamp = pd.to_datetime(value)
    except (ValueError, TypeError):
        stamp = None

    if stamp is None or pd.isna(stamp):
        return datetime.now().strftime('%Y-%m-%d')
    return stamp.strftime('%Y-%m-%d')


def _collect_customer_po(
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> str:
    """거래처 발주번호 표기 문자열

    월합 거래명세표는 DN마다 발주번호가 다를 수 있으므로 **전체 아이템에서** 모읍니다.
    첫 건만 쓰면 여러 발주가 묶인 문서에 엉뚱한 번호 하나만 나가게 됩니다.

    Args:
        order_data: 주문 데이터 (items_df가 없을 때 폴백)
        items_df: 문서에 실린 전체 아이템

    Returns:
        발주번호 (여러 건이면 ', ' 구분, 없으면 'N/A')
    """
    values: list = []
    if items_df is not None and not items_df.empty:
        col = resolve_column(items_df.columns, 'customer_po')
        if col is not None:
            values = items_df[col].tolist()
    if not values:
        values = [get_value(order_data, 'customer_po', '')]

    unique: dict[str, None] = {}
    for value in values:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            continue
        text = str(value).strip()
        if text and text.lower() != 'nan':
            unique.setdefault(text, None)

    return ', '.join(unique) if unique else NO_VALUE


def _mail_ts(
    order_data: pd.Series,
    output_file: Path,
    doc_id: str,
    opts: MailOptions,
    items_df: pd.DataFrame | None = None,
) -> bool:
    """생성된 거래명세표를 메일로 발송/초안 생성

    메일 실패는 거래명세표 생성 성공을 뒤엎지 않습니다 (경고만 출력).

    Args:
        order_data: 주문 데이터 (사업자번호/고객명/출고일 포함)
        output_file: 생성된 거래명세표 경로
        doc_id: DN_ID 또는 선수금_ID
        opts: 메일 옵션
        items_df: 문서에 실린 전체 아이템 (발주번호 수집용)

    Returns:
        메일 생성/발송 성공 여부
    """
    if not opts.enabled:
        return True

    df_customer = opts.customer_master()
    if df_customer is None:
        return False

    try:
        recipient = find_recipient_for_order(order_data, df_customer)
    except MailConfigError as e:
        print(f"  [메일 오류] {e}")
        return False

    if recipient is None:
        biz_no = get_value(order_data, 'biz_no', '(사업자번호 없음)')
        customer = get_value(order_data, 'customer_name', '')
        print(f"  [메일 생략] 수신자 미등록 — {customer} / {biz_no}")
        print(f"             {CUSTOMER_DOMESTIC_SHEET} 시트에 해당 사업자번호의 이메일을 입력하세요.")
        return False

    # 제목/본문 날짜: 출고일 우선, 없으면 오늘
    # (선수금 거래명세표는 SO_국내 기반이라 출고일 자체가 없다)
    date_str = _format_mail_date(get_value(order_data, 'dispatch_date', None))
    customer_po = _collect_customer_po(order_data, items_df)

    # 누구에게 나가는지 먼저 보여주고 확인받는다 (오발송 차단)
    show_recipient(recipient)
    print(f"  발주번호: {customer_po}")

    if opts.ask and not _confirm("  이메일을 발송하시겠습니까? [y/N]: "):
        print("  -> 메일 생략")
        return False

    try:
        result = create_ts_mail(
            xlsx_path=output_file,
            recipient=recipient,
            doc_id=doc_id,
            date_str=date_str,
            send=opts.send,
            backend=opts.backend,
            customer_po=customer_po,
        )
    except MailConfigError as e:
        print(f"  [메일 오류] {e}")
        return False

    return report_mail_result(result, want_send=opts.send)


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


def generate_ts_from_dn(
    dn_id: str,
    df_dn: pd.DataFrame,
    mail_opts: MailOptions | None = None,
) -> bool:
    """DN 기반 거래명세표 생성

    DocumentService를 사용하여 거래명세표를 생성합니다.

    Args:
        dn_id: DN_ID
        df_dn: DN 데이터 (하위 호환용, 실제로는 사용하지 않음)
        mail_opts: 메일 발송 옵션 (None이면 발송 안 함)

    Returns:
        성공 여부 (메일 실패는 성공 여부에 영향 없음)
    """
    mail_opts = mail_opts or MailOptions.disabled()
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
        _mail_ts(order_data.first_item, result.output_file, dn_id, mail_opts,
                 items_df=order_data.items_df)
        return True
    else:
        if result.status == GenerationStatus.FILE_ERROR:
            print(f"  [오류] {result.errors[0] if result.errors else result.message}")
        else:
            print(f"  [오류] {result.message}")
        return False


def generate_merged_ts(dn_ids: list[str], mail_opts: MailOptions | None = None) -> bool:
    """여러 DN을 합쳐서 월합 거래명세표 생성

    Args:
        dn_ids: DN_ID 목록
        mail_opts: 메일 발송 옵션 (None이면 발송 안 함)

    Returns:
        성공 여부
    """
    mail_opts = mail_opts or MailOptions.disabled()
    print(f"\n{'=' * 50}")
    print(f"월합 거래명세표 생성: {len(dn_ids)}건")
    print('=' * 50)

    service = DocumentService()
    all_items = []
    first_order_data = None
    first_dn_id = None
    customer_names = set()
    latest_dispatch_date = None

    # 1. 모든 DN 데이터 수집
    for dn_id in dn_ids:
        order_data = service.finder.find_dn(dn_id)
        if order_data is None:
            print(f"  [경고] '{dn_id}'를 찾을 수 없습니다. 건너뜁니다.")
            continue

        # 첫 번째 유효한 데이터 저장
        if first_order_data is None:
            first_order_data = order_data
            first_dn_id = dn_id

        # 고객명 수집
        customer_name = order_data.get_value('customer_name', '')
        if customer_name:
            customer_names.add(customer_name)

        # 출고일 비교 (가장 최근 출고일 사용)
        dispatch_date = order_data.get_value('dispatch_date', None)
        if dispatch_date is not None:
            try:
                if not isinstance(dispatch_date, pd.Timestamp):
                    dispatch_date = pd.to_datetime(dispatch_date)
                if latest_dispatch_date is None or dispatch_date > latest_dispatch_date:
                    latest_dispatch_date = dispatch_date
            except (ValueError, TypeError):
                pass

        # 아이템 수집
        if order_data.is_multi_item:
            all_items.append(order_data.items_df)
            print(f"  {dn_id}: {order_data.item_count}개 아이템")
        else:
            all_items.append(pd.DataFrame([order_data.first_item]))
            print(f"  {dn_id}: 1개 아이템")

    # 2. 유효성 검사
    if not all_items:
        print("  [오류] 유효한 DN이 없습니다.")
        return False

    if len(customer_names) > 1:
        print(f"  [경고] 고객이 여러 명입니다: {sorted(customer_names)}")
        actual = first_order_data.get_value('customer_name', 'Unknown')
        print(f"  -> '{actual}' 기준으로 생성합니다.")

    # 3. 아이템 합치기
    merged_items_df = pd.concat(all_items, ignore_index=True)
    print(f"\n  총 {len(merged_items_df)}개 아이템")

    # 4. 출고일 업데이트 (가장 최근 출고일)
    if latest_dispatch_date is not None:
        first_order_data.first_item['출고일'] = latest_dispatch_date
        print(f"  출고일: {latest_dispatch_date.strftime('%Y-%m-%d')}")

    # 5. 파일명 생성 (월합_DN_고객명_날짜) — 충돌 시 _1,_2 자동 접미사 + 경로 검증
    customer_name = first_order_data.get_value('customer_name', 'Unknown')
    output_path = generate_output_filename(
        "월합", first_dn_id or "merge", customer_name, TS_OUTPUT_DIR
    )
    if not validate_output_path(output_path, TS_OUTPUT_DIR):
        return False
    output_filename = output_path.name

    # 6. 거래명세표 생성
    try:
        create_ts_xlwings(
            template_path=TS_TEMPLATE_FILE,
            output_path=output_path,
            order_data=first_order_data.first_item,
            items_df=merged_items_df,
            doc_type='DN',
            use_po_as_remark=True,
        )
        print(f"\n  -> 월합 거래명세표 생성 완료: {output_filename}")
    except Exception as e:
        print(f"  [오류] 거래명세표 생성 실패: {e}")
        return False

    # 고객이 섞인 월합 문서는 메일 발송 금지 — 타 거래처 라인이 노출됨
    if mail_opts.enabled and len(customer_names) > 1:
        print(f"  [메일 중단] 고객이 {len(customer_names)}곳 섞여 있어 발송하지 않습니다.")
        print(f"             {sorted(customer_names)}")
        return True

    # 월합은 DN마다 발주번호가 다를 수 있으므로 합쳐진 전체 아이템을 넘긴다
    _mail_ts(first_order_data.first_item, output_path, first_dn_id or "merge", mail_opts,
             items_df=merged_items_df)
    return True


def generate_ts_from_adv(advance_id: str, mail_opts: MailOptions | None = None) -> bool:
    """선수금 거래명세표 생성 (SO_국내 데이터 사용)

    DocumentService를 사용하여 선수금 거래명세표를 생성합니다.

    Args:
        advance_id: 선수금_ID
        mail_opts: 메일 발송 옵션 (None이면 발송 안 함)

    Returns:
        성공 여부 (메일 실패는 성공 여부에 영향 없음)
    """
    mail_opts = mail_opts or MailOptions.disabled()
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
        _mail_ts(order_data.first_item, gen_result.output_file, advance_id, mail_opts,
                 items_df=order_data.items_df)
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

월합 거래명세표 (여러 DN을 한 장으로):
  python create_ts.py DN-2026-0001 DN-2026-0002 DN-2026-0003 --merge

Outlook 메일 발송:
  생성이 끝나면 받는사람/참조를 보여주고 "이메일을 발송하시겠습니까? [y/N]"을 묻습니다.
  y를 누르면 PDF를 첨부한 Outlook 메일 창이 뜨고, 최종 [보내기]는 직접 누릅니다.
  수신자는 사업자번호로 Customer_국내에서 조회하고, 고정 참조(CC)는 user_settings.py의
  TS_MAIL_CC로 관리합니다.

  python create_ts.py DN-2026-0001 --mail     # 확인 없이 바로 Outlook 창
  python create_ts.py DN-2026-0001 --send     # 확인 없이 즉시 발송
  python create_ts.py DN-2026-0001 --no-mail  # 묻지 않고 문서만

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

    parser.add_argument(
        '-m', '--merge',
        action='store_true',
        help='여러 DN을 한 장의 거래명세표로 합침 (월합)',
    )

    parser.add_argument(
        '-i', '--interactive',
        action='store_true',
        help='대화형 모드 (여러 줄 입력 지원)',
    )

    parser.add_argument(
        '--mail',
        action='store_true',
        help='확인 없이 Outlook 메일 초안 열기 (기본은 건별 y/N 확인)',
    )

    parser.add_argument(
        '--send',
        action='store_true',
        help='확인 없이 즉시 발송',
    )

    parser.add_argument(
        '--no-mail',
        action='store_true',
        help='메일 확인 없이 문서만 생성 (배치용)',
    )

    return parser


def prepare_mail_options(args: argparse.Namespace) -> MailOptions:
    """CLI 인자로 메일 옵션 구성

    마스터 로딩은 실제로 메일이 필요한 시점까지 미룹니다
    (메일을 쓰지 않는 실행에 Excel 로딩 비용을 물리지 않기 위함).

    Args:
        args: 파싱된 CLI 인자

    Returns:
        MailOptions
    """
    mode = resolve_mail_mode(args, is_tty=sys.stdin.isatty())
    if mode is MailMode.OFF:
        return MailOptions.disabled()

    label = {
        MailMode.ASK: '건별 확인 후 발송',
        MailMode.DRAFT: '확인 없이 초안 열기',
        MailMode.SEND: '확인 없이 즉시 발송',
    }[mode]
    # auto는 용도에 따라 갈린다 — 초안은 .eml(사용자 기본 메일 앱), 즉시 발송만 COM
    backend = resolve_backend(send=mode is MailMode.SEND)
    backend_label = 'Outlook COM' if backend is MailBackend.OUTLOOK else '.eml 초안'
    # 참조자는 거래처마다 달라서 Customer_국내의 '참조 이메일'로 건별 관리한다.
    # TS_MAIL_CC는 모든 거래처에 공통으로 붙일 주소가 있을 때만 쓰는 선택 항목.
    cc_note = f" / 고정 CC {len(TS_MAIL_CC)}명" if TS_MAIL_CC else ""
    print(f"\n메일: {label} / 첨부 {TS_MAIL_ATTACH_FORMAT.upper()} / "
          f"방식 {backend_label}{cc_note}")

    # .eml은 작성 창을 띄우는 방식이라 자동 발송이 불가능하다
    if mode is MailMode.SEND and backend is MailBackend.EML:
        print("  [주의] .eml 방식은 자동 발송을 지원하지 않습니다 — 초안까지만 진행됩니다.")
        print("         (Outlook COM 사용 불가 환경 — 클래식 Outlook이 없거나 실행 실패)")

    return MailOptions(mode=mode, backend=backend)


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

    # 대화형 모드: 여러 줄 입력 받기
    if args.interactive:
        print("\nDN_ID를 입력하세요 (한 줄에 하나씩, 빈 줄 입력 시 완료):")
        doc_ids = []
        while True:
            try:
                line = input().strip()
                if not line:
                    break
                doc_ids.append(line)
            except EOFError:
                break

        if not doc_ids:
            print("[오류] ID가 입력되지 않았습니다.")
            return 1

        print(f"\n{len(doc_ids)}개 ID 입력됨")
        args.doc_ids = doc_ids

    # 인자 없으면 도움말 + 사용 가능한 ID 출력
    if not args.doc_ids:
        parser.print_help()
        print_available_ids(df_dn, df_pmt)
        return 0

    print(f"DN: {len(df_dn)}건, PMT: {len(df_pmt)}건 로드 완료")

    # 메일 옵션 (마스터는 실제 발송 시점에 지연 로딩)
    mail_opts = prepare_mail_options(args)

    # --merge 옵션: 여러 DN을 한 장으로 합침
    if args.merge:
        # DN만 merge 가능 (ADV는 제외)
        dn_ids = [doc_id for doc_id in args.doc_ids if detect_id_type(doc_id) == 'DN']
        adv_ids = [doc_id for doc_id in args.doc_ids if detect_id_type(doc_id) == 'ADV']

        if adv_ids:
            print(f"\n[경고] 선수금({len(adv_ids)}건)은 merge에서 제외됩니다: {adv_ids}")

        # 중복 DN_ID 제거 (입력 순서 보존) — 같은 DN을 두 번 넘기면 금액 이중 합산되므로 차단
        _deduped = list(dict.fromkeys(dn_ids))
        if len(_deduped) != len(dn_ids):
            _dups = [d for d in dict.fromkeys(dn_ids) if dn_ids.count(d) > 1]
            print(f"\n[경고] 중복 DN_ID 제거됨: {_dups}")
            dn_ids = _deduped

        if len(dn_ids) < 2:
            print("\n[오류] --merge 옵션은 2개 이상의 DN_ID가 필요합니다.")
            return 1

        success = generate_merged_ts(dn_ids, mail_opts)
        return 0 if success else 1

    # 일반 모드: 각 ID에 대해 거래명세표 생성
    success_count = 0
    for doc_id in args.doc_ids:
        id_type = detect_id_type(doc_id)

        if id_type == 'DN':
            if generate_ts_from_dn(doc_id, df_dn, mail_opts):
                success_count += 1
        else:  # ADV
            if generate_ts_from_adv(doc_id, mail_opts):
                success_count += 1

    # 결과 출력
    print(f"\n{'=' * 50}")
    print(f"완료: {success_count}/{len(args.doc_ids)}건 거래명세표 생성")
    print(f"출력 폴더: {TS_OUTPUT_DIR}")
    print('=' * 50)

    return 0 if success_count == len(args.doc_ids) else 1


if __name__ == "__main__":
    sys.exit(main())
