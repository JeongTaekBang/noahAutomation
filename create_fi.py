#!/usr/bin/env python
"""
Final Invoice 생성기 (대금 청구용)
===================================

NOAH_SO_PO_DN.xlsx의 DN_해외 시트에서 데이터를 읽어 Final Invoice를 생성합니다
(xlsx + PDF). Invoice Date = DN_해외 선적일.

**Invoice No = DN_해외의 INV 번호** — 같은 번호를 넣은 행 전부가 인보이스 한 장입니다.
DN_ID·발주번호는 그 인보이스를 찾아가는 입구라, 어느 쪽으로 넣어도 같은 문서가 나옵니다.
INV 번호가 없는 DN은 예전처럼 DN_ID를 Invoice No로 씁니다 (복수 PO면 발주번호별 분리).

사용법:
    python create_fi.py INV-2026-0001              # INV 번호 한 건 = 인보이스 한 장
    python create_fi.py DNO-2026-0001              # 그 DN의 인보이스 전부 (INV 번호별)
    python create_fi.py DNO-2026-0001 DNO-2026-0002 # 여러 건 동시 생성
    python create_fi.py --po 26KPO00144            # 발주번호 기준 (그 발주가 실린 인보이스)
"""

from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime

import pandas as pd

from po_generator.config import (
    FI_OUTPUT_DIR,
    FI_TEMPLATE_FILE,
    NOAH_SO_PO_DN_FILE,
)
from po_generator.fi_invoice import (
    INVOICE_COLUMN_NAME,
    invoice_numbers,
    invoices_of,
    suggest_invoice_numbers,
    unique_texts,
)
from po_generator.utils import (
    get_value,
    resolve_column,
)
from po_generator.logging_config import setup_logging
from po_generator.services import DocumentService, GenerationStatus

logger = logging.getLogger(__name__)


def _saved_at() -> str:
    """읽은 워크북의 저장 시각 ('MM-DD HH:MM', 알 수 없으면 '')"""
    try:
        return datetime.fromtimestamp(NOAH_SO_PO_DN_FILE.stat().st_mtime).strftime('%m-%d %H:%M')
    except OSError:
        return ''


def invoice_hint(items: pd.DataFrame) -> str:
    """INV 번호를 넣으라는 안내의 뒷말 — 컬럼이 없으면 컬럼부터, 있으면 넣을 번호까지

    **저장된 파일을 읽는다**는 것을 같이 알린다. Excel에 번호를 넣고 저장하기 전에 돌리면
    번호가 없는 것으로 읽혀 DN 번호로 나온다 (2026-11-11 실사용: 14:18 생성, 저장은 14:21 —
    "DNO로 나오네"). 읽은 파일의 저장 시각을 보여 주면 그 어긋남이 바로 보인다.

    Args:
        items: 지금 만들려는 라인 (DN 한 건이면 그 DN 번호를 따른 번호를 권한다)
    """
    if resolve_column(items.columns, 'invoice_no') is None:
        return f"DN_해외에 '{INVOICE_COLUMN_NAME}' 컬럼을 만들고 번호를 넣으세요"

    details = []
    saved = _saved_at()
    if saved:
        details.append(f"읽은 파일은 {saved} 저장본")
    suggested = suggest_invoice_numbers(items)
    if len(suggested) == 1:
        details.append(f"권장 번호: {suggested[0][0]}")
    elif suggested:
        # 복수 PO 선적은 Customer PO마다 한 장 — 어느 PO에 어느 번호인지까지
        details.append("권장 번호: " + ', '.join(f"{no} = {po}" for no, po in suggested))
    tail = f" ({', '.join(details)})" if details else ''
    return f"DN_해외 {INVOICE_COLUMN_NAME}에 번호를 넣고 저장한 뒤 다시 생성하세요{tail}"


def no_invoice_notice(items: pd.DataFrame, shown_as: str) -> str:
    """INV 번호 없이 예전 방식으로 나갈 때의 알림 — 화면에 찍고 요약에도 싣는다

    Args:
        items: 지금 만들려는 라인
        shown_as: Invoice No가 어떻게 나가는지 (예: 'Invoice No가 DNO-2026-0191로 나갑니다')
    """
    notice = f"INV 번호가 없어 {shown_as} — {invoice_hint(items)}"
    print(f"  [확인 필요] {notice}")
    return notice


def print_available_ids(df_dn: pd.DataFrame, limit: int = 15) -> None:
    """사용 가능한 DN_ID 목록 출력

    Args:
        df_dn: DN 해외 데이터
        limit: 출력 제한 수
    """
    print("\n" + "=" * 60)
    print("사용 가능한 DN_ID 목록 (DN_해외)")
    print("=" * 60)

    dn_ids = df_dn['DN_ID'].dropna().unique().tolist()
    print(f"\n[Final Invoice] DN_ID ({len(dn_ids)}건)")
    print("-" * 40)
    for dn_id in dn_ids[:limit]:
        # 고객명도 함께 표시
        customer = df_dn[df_dn['DN_ID'] == dn_id]['Customer name'].iloc[0] if len(df_dn[df_dn['DN_ID'] == dn_id]) > 0 else ''
        customer_short = str(customer)[:25] if customer else ''
        print(f"  {dn_id:<20} {customer_short}")
    if len(dn_ids) > limit:
        print(f"  ... 외 {len(dn_ids) - limit}건")

    print("\n" + "=" * 60)
    print("위 DN_ID나 INV 번호를 입력하여 Final Invoice를 생성하세요.")
    if resolve_column(df_dn.columns, 'invoice_no') is None:
        print(f"INV 번호: DN_해외에 '{INVOICE_COLUMN_NAME}' 컬럼이 없습니다 — Invoice No는 DN_ID로 나갑니다.")
    else:
        used = sorted(set(invoice_numbers(df_dn)) - {''})
        print(f"INV 번호: {len(used)}건 사용 중 — 번호는 DN 번호를 따릅니다 "
              f"(DNO-2026-0215 → INV-2026-0215, Customer PO가 여럿이면 -1, -2 …)")
    print("=" * 60)


def _print_item_info(order_data, label: str = '') -> None:
    """아이템 정보 출력 헬퍼"""
    if order_data.is_multi_item:
        print(f"  {label}[다중 아이템] {order_data.item_count}개 아이템 발견")
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
        print(f"  {label}품목: {item_name}")


def print_available_customer_pos(df_dn: pd.DataFrame, limit: int = 20) -> None:
    """사용 가능한 Customer PO(발주번호) 목록 출력

    Args:
        df_dn: DN 해외 데이터
        limit: 출력 제한 수
    """
    cpo_col = resolve_column(df_dn.columns, 'customer_po')
    if not cpo_col:
        print("  [오류] Customer PO 컬럼을 찾을 수 없습니다.")
        return

    print("\n" + "=" * 60)
    print("사용 가능한 발주번호 목록 (DN_해외)")
    print("=" * 60)

    cpos = df_dn[cpo_col].dropna().unique().tolist()
    cpos = sorted(str(p) for p in cpos)
    print(f"\n[Final Invoice] 발주번호 ({len(cpos)}건)")
    print("-" * 55)
    for po in cpos[:limit]:
        # 해당 PO의 DN 목록과 고객명
        po_rows = df_dn[df_dn[cpo_col].astype(str).str.strip() == str(po).strip()]
        dn_ids = po_rows['DN_ID'].dropna().unique().tolist()
        customer = po_rows['Customer name'].iloc[0] if len(po_rows) > 0 else ''
        customer_short = str(customer)[:20] if customer else ''
        dn_list = ', '.join(str(d) for d in dn_ids[:3])
        if len(dn_ids) > 3:
            dn_list += f' 외 {len(dn_ids) - 3}건'
        print(f"  {po:<20} {customer_short:<22} DN: {dn_list}")
    if len(cpos) > limit:
        print(f"  ... 외 {len(cpos) - limit}건")

    print("\n" + "=" * 60)
    print("위 발주번호 중 하나를 입력하여 Final Invoice를 생성하세요.")
    print("  예: python create_fi.py --po 26KPO00144")
    print("=" * 60)


def _report(result, label: str = '') -> tuple[bool, list[str]]:
    """생성 결과 출력 — (성공 여부, 요약에 다시 보여줄 알림)

    xlsx와 PDF를 같이 만든다. PDF가 빠지면 보낼 본이 없으므로 성공으로 치지 않는다
    (create_coo.py와 같은 규약). 경고는 Invoice Date 판단(선적일 공란·라인마다 다름)과
    인보이스 묶음 판단(여러 DN·여러 PO)이다.
    """
    suffix = f" ({label})" if label else ''
    if not result.success:
        # 사유가 errors에 있는 상태(파일·검증)는 그것을, 나머지는 message를 보여 준다
        if result.status in (GenerationStatus.FILE_ERROR, GenerationStatus.VALIDATION_ERROR):
            reasons = result.errors or [result.message]
        else:
            reasons = [result.message]
        for reason in reasons:
            print(f"  [오류]{suffix} {reason}")
        logger.error(f"Final Invoice 생성 실패{suffix}: {'; '.join(reasons)}")
        return False, []

    print(f"  -> Final Invoice 생성 완료{suffix}: {result.output_file.name}")
    for extra in result.extra_files:
        print(f"  -> PDF: {extra.name}")
    for warning in result.warnings:
        print(f"  [확인 필요] {warning}")
    for error in result.errors:
        print(f"  [실패] {error}")
    return not result.errors, result.warnings + result.errors


def generate_invoice(
    service: DocumentService,
    invoice_no: str,
    done: set[str],
) -> tuple[bool, list[str]]:
    """INV 번호 한 건 = 인보이스 한 장 (그 번호가 적힌 행 전부 — DN을 가리지 않는다)

    Args:
        service: 문서 서비스 (조회 캐시를 공유한다)
        invoice_no: INV 번호
        done: 이번 실행에서 이미 만든 INV 번호 — DN 두 건이 같은 인보이스에 걸려 있을 때
            같은 문서를 두 번 만들지 않는다

    Returns:
        (성공 여부, 알림 목록)
    """
    order_data = service.finder.find_dn_export_by_invoice(invoice_no)
    if order_data is None:
        print(f"  [오류] INV 번호 '{invoice_no}'를 찾을 수 없습니다.")
        return False, []

    items = order_data.all_items
    number = invoice_numbers(items).iloc[0]
    print(f"\n  --- {number} ---")
    if number in done:
        print("  [건너뜀] 이번 실행에서 이미 만들었습니다.")
        return True, []
    done.add(number)

    print(f"  고객: {order_data.get_value('customer_name', 'N/A')}")
    print(f"  DN: {', '.join(unique_texts(items, 'dn_id')) or 'N/A'}")
    print(f"  Customer PO: {', '.join(unique_texts(items, 'customer_po')) or 'N/A'}")
    _print_item_info(order_data)

    return _report(service.generate_fi_by_invoice(number), label=number)


def _generate_invoices(
    service: DocumentService,
    numbers: list[str],
    done: set[str],
) -> tuple[bool, list[str]]:
    """INV 번호 여러 건을 차례로 — 알림에는 어느 인보이스 것인지 붙인다"""
    all_success = True
    notices: list[str] = []
    for number in numbers:
        ok, inv_notices = generate_invoice(service, number, done)
        all_success = all_success and ok
        notices.extend(f"{number}  {note}" for note in inv_notices)
    return all_success, notices


def generate_fi_by_po(
    service: DocumentService,
    customer_po: str,
    done: set[str],
) -> tuple[bool, list[str]]:
    """발주번호(Customer PO) 기준 Final Invoice 생성

    INV 번호가 있으면 그 발주가 실린 인보이스를 만든다. 없으면 예전 방식 — 복수 DN에
    걸친 동일 발주번호 아이템을 한 장으로 통합한다 (Invoice No = 첫 DN_ID).

    Returns:
        (성공 여부, 알림 목록)
    """
    print(f"\n{'=' * 60}")
    print(f"Final Invoice 생성 (발주번호): {customer_po}")
    print('=' * 60)

    # 1. Customer PO로 데이터 검색
    order_data = service.finder.find_dn_export_by_customer_po(customer_po)
    if order_data is None:
        print(f"  [오류] 발주번호 '{customer_po}'를 찾을 수 없습니다.")
        return False, []

    # 2. 매칭된 DN 정보 출력
    items_df = order_data.all_items
    dn_ids = items_df['DN_ID'].dropna().unique().tolist() if 'DN_ID' in items_df.columns else []
    print(f"  고객: {order_data.get_value('customer_name', 'N/A')}")
    if dn_ids:
        print(f"  관련 DN: {', '.join(str(d) for d in dn_ids)} ({len(dn_ids)}건)")
    print(f"  아이템: {order_data.item_count}개")

    # 3. INV 번호가 있으면 인보이스 단위로
    numbers, error = invoices_of(items_df, customer_po)
    if error:
        print(f"  [오류] {error}")
        return False, []
    if numbers:
        print(f"\n  [INV 번호] 이 발주가 실린 인보이스 {len(numbers)}건: {', '.join(numbers)}")
        return _generate_invoices(service, numbers, done)

    _print_item_info(order_data, label='')

    # 4. 예전 방식 — 복수 DN 통합 1장 (Invoice No = 첫 DN_ID)
    notice = no_invoice_notice(items_df, "Invoice No가 DN 번호로 나갑니다")
    ok, notices = _report(service.generate_fi_by_customer_po(customer_po))
    return ok, [notice] + notices


def generate_fi(
    service: DocumentService,
    doc_id: str,
    done: set[str],
) -> tuple[bool, list[str]]:
    """Final Invoice 생성 — DN_ID 또는 INV 번호

    - INV 번호 → 그 인보이스 한 장
    - DN_ID → 그 DN의 라인에 INV 번호가 있으면 인보이스별로, 없으면 예전 방식
      (Invoice No = DN_ID, RCK PO가 복수개면 발주번호별 분리 생성)

    Args:
        service: 문서 서비스
        doc_id: DN_ID 또는 INV 번호
        done: 이번 실행에서 이미 만든 INV 번호

    Returns:
        (성공 여부, 알림 목록)
    """
    print(f"\n{'=' * 60}")
    print(f"Final Invoice 생성: {doc_id}")
    print('=' * 60)

    # 1. INV 번호인가 — 먼저 본다 (DN 조회는 못 찾으면 경고 로그를 남긴다)
    if service.finder.find_dn_export_by_invoice(doc_id) is not None:
        return generate_invoice(service, doc_id, done)

    order_data = service.finder.find_dn_export(doc_id)
    if order_data is None:
        print(f"  [오류] '{doc_id}'를 찾을 수 없습니다 (DN_ID·INV 번호 어느 쪽에도 없음).")
        return False, []
    dn_id = doc_id

    # 2. 기본 정보 출력
    print(f"  고객: {order_data.get_value('customer_name', 'N/A')}")
    bill_to_1 = order_data.get_value('bill_to_1', '')
    if bill_to_1:
        print(f"  Bill to: {bill_to_1}")

    payment_terms = order_data.get_value('payment_terms', '')
    if payment_terms:
        print(f"  Payment Terms: {payment_terms}")

    # 3. INV 번호가 있으면 인보이스 단위로
    items_df = order_data.all_items
    numbers, error = invoices_of(items_df, dn_id)
    if error:
        print(f"  [오류] {error}")
        return False, []
    if numbers:
        print(f"\n  [INV 번호] 이 DN의 인보이스 {len(numbers)}건: {', '.join(numbers)}")
        return _generate_invoices(service, numbers, done)

    _print_item_info(order_data)

    # 4. RCK PO 그룹 감지 (INV 번호가 없는 DN — 예전 방식)
    rck_po_col = resolve_column(items_df.columns, 'rck_po')

    if rck_po_col:
        rck_pos = items_df[rck_po_col].dropna().unique().tolist()
    else:
        rck_pos = []

    # 5. 단일 PO (또는 PO 컬럼 없음) → 기존 방식
    if len(rck_pos) <= 1:
        notice = no_invoice_notice(items_df, f"Invoice No가 {dn_id}로 나갑니다")
        ok, notices = _report(service.generate_fi(dn_id))
        return ok, [notice] + notices

    # 6. 복수 PO → 발주번호별 분리 생성
    print(f"\n  [발주번호 분리] {len(rck_pos)}개 RCK PO 발견 → 발주번호별 FI 생성")
    # 여러 장이 같은 Invoice No(DN_ID)로 나간다 — 업체 AP가 서로 다른 청구를 같은 번호로 받는다
    same_no_notice = no_invoice_notice(
        items_df, f"{len(rck_pos)}장이 모두 Invoice No {dn_id}로 나갑니다")
    # RCK PO 공란 라인은 어떤 발주번호 필터(== po)에도 안 잡혀 모든 FI에서 누락됨 → 과소청구.
    # 조용히 빠지지 않도록 경고 + 건수 정합성(분리합 vs 전체) 점검.
    _blank_mask = (items_df[rck_po_col].isna()
                   | (items_df[rck_po_col].astype(str).str.strip() == ''))
    _n_blank = int(_blank_mask.sum())
    if _n_blank:
        print(f"  [경고] RCK PO 공란 라인 {_n_blank}건은 발주번호별 FI에서 누락됩니다 "
              f"(과소청구 위험) — 해당 라인의 RCK PO 입력 후 재생성 필요")
        logger.warning("FI 분리 생성: RCK PO 공란 %d건 누락 위험 (dn_id=%s)", _n_blank, dn_id)
    for po in sorted(rck_pos):
        po_items = items_df[items_df[rck_po_col] == po]
        print(f"    {po}: {len(po_items)}개 아이템")

    all_success = True
    notices: list[str] = [same_no_notice]
    for po in sorted(rck_pos):
        print(f"\n  --- {po} ---")
        ok, po_notices = _report(service.generate_fi(dn_id, rck_po=po), label=po)
        all_success = all_success and ok
        notices.extend(f"{po}  {note}" for note in po_notices)

    return all_success, notices


def _print_summary(headline: str, notices: list[str]) -> None:
    """마지막 요약 — 경고는 문서가 길면 스크롤 위로 사라지므로 한 번 더 보여 준다"""
    print(f"\n{'=' * 60}")
    print(headline)
    for note in notices:
        print(f"[확인 필요] {note}")
    print(f"출력 폴더: {FI_OUTPUT_DIR} (xlsx + PDF)")
    print('=' * 60)


def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성

    Returns:
        설정된 ArgumentParser
    """
    description = f"""
Final Invoice 생성기 (대금 청구용)
===================================

NOAH_SO_PO_DN.xlsx의 DN_해외 시트에서 데이터를 읽어
Final Invoice를 xlsx + PDF로 자동 생성합니다.
Invoice Date는 DN_해외의 선적일입니다 (공란이면 오늘 날짜 + 경고).

Invoice No는 DN_해외의 INV 번호('{INVOICE_COLUMN_NAME}' 컬럼)입니다 — 같은 번호를 넣은
행 전부가 인보이스 한 장입니다. INV 번호가 없는 DN은 DN_ID가 Invoice No로 나갑니다.

입력은 세 가지입니다 (어느 쪽으로 넣어도 같은 INV 번호는 같은 문서가 나옵니다):
  1. INV 번호: 그 인보이스 한 장
  2. DN_ID: 그 DN의 인보이스 전부 (INV 번호가 없으면 DN 단위, 복수 PO는 발주번호별 분리)
  3. 발주번호(--po): 그 발주가 실린 인보이스 (INV 번호가 없으면 복수 DN 통합 1장)
"""

    epilog = """
사용 예시:
  python create_fi.py INV-2026-0001              # INV 번호로 FI 생성
  python create_fi.py DNO-2026-0001              # DN_ID로 FI 생성
  python create_fi.py DNO-2026-0001 INV-2026-0002 # 여러 건 동시 생성
  python create_fi.py --po 26KPO00144            # 발주번호 기준 FI 생성
  python create_fi.py --po 26KPO00144 26KPO00145 # 여러 발주번호 동시 생성
  python create_fi.py --po                       # 사용 가능한 발주번호 목록 표시

인자 없이 실행하면 사용 가능한 DN_ID 목록과 INV 번호 현황을 표시합니다.
INV 번호는 DN 번호를 따릅니다 (DNO-2026-0215 → INV-2026-0215, Customer PO가 여럿이면 -1, -2 …).
"""

    parser = argparse.ArgumentParser(
        prog='create_fi',
        description=description,
        epilog=epilog,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        'ids',
        nargs='*',
        metavar='ID',
        help='DN_ID 또는 INV 번호 (예: DNO-2026-0001, INV-2026-0001)',
    )

    parser.add_argument(
        '--po',
        nargs='*',
        metavar='CUSTOMER_PO',
        help='발주번호(Customer PO) 기준 FI 생성. 인자 없이 --po만 입력 시 목록 표시.',
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

    # 데이터 로드 — 서비스 하나가 읽은 것을 끝까지 같이 쓴다 (건마다 다시 읽지 않는다)
    print("NOAH_SO_PO_DN.xlsx 로딩 중...")
    service = DocumentService()
    try:
        df_dn = service.finder.load_dn_export_data()
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        return 1

    # 이번 실행에서 만든 INV 번호 — 같은 인보이스를 두 번 만들지 않는다
    done: set[str] = set()

    # --po 모드: 발주번호 기준 생성
    if args.po is not None:
        # ID 위치 인자와 --po 동시 입력 방지 (positional이 조용히 무시되는 것 방지)
        if args.ids:
            print("[오류] DN_ID·INV 번호와 --po는 동시에 사용할 수 없습니다. 둘 중 하나만 지정하세요.")
            return 1
        if not args.po:
            # --po만 입력 (인자 없음) → 발주번호 목록 표시
            print_available_customer_pos(df_dn)
            return 0

        print(f"DN 해외: {len(df_dn)}건 로드 완료")

        success_count = 0
        notices: list[str] = []
        for customer_po in args.po:
            ok, po_notices = generate_fi_by_po(service, customer_po, done)
            if ok:
                success_count += 1
            notices.extend(f"{customer_po}  {note}" for note in po_notices)

        _print_summary(f"완료: {success_count}/{len(args.po)}건 Final Invoice 생성 (발주번호 기준)",
                       notices)
        return 0 if success_count == len(args.po) else 1

    # DN_ID·INV 번호 모드 (기본)
    if not args.ids:
        parser.print_help()
        print_available_ids(df_dn)
        return 0

    print(f"DN 해외: {len(df_dn)}건 로드 완료")

    # 각 ID에 대해 Final Invoice 생성
    success_count = 0
    notices: list[str] = []
    for doc_id in args.ids:
        ok, id_notices = generate_fi(service, doc_id, done)
        if ok:
            success_count += 1
        notices.extend(f"{doc_id}  {note}" for note in id_notices)

    _print_summary(f"완료: {success_count}/{len(args.ids)}건 Final Invoice 생성", notices)
    return 0 if success_count == len(args.ids) else 1


if __name__ == "__main__":
    sys.exit(main())
