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

하루치 출고를 거래처별 메일 한 통으로 묶기 (문서는 DN별 1장 그대로):
    python create_ts.py --date 2026-08-06 --customer 씨앤케이   # 그 거래처만
    python create_ts.py --date 2026-08-06                      # 그날 전체, 거래처별 1통씩
    python create_ts.py DN-2026-0001 DN-2026-0002 --one-mail   # 명시 ID를 묶어 1통
"""

from __future__ import annotations

import argparse
import logging
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Sequence

import pandas as pd

from po_generator.config import (
    TS_MAIL_ATTACH_FORMAT,
    TS_MAIL_CC,
    TS_OUTPUT_DIR,
    TS_TEMPLATE_FILE,
)
from po_generator.utils import (
    BIZ_NO_MIN_DIGITS,
    load_dn_data,
    load_pmt_data,
    get_value,
    normalize_biz_no,
    resolve_column,
)
from po_generator.ts_generator import create_ts_xlwings
from po_generator.cli_common import validate_output_path, generate_output_filename
from po_generator.logging_config import setup_logging
from po_generator.mailer import (
    MailConfigError,
    as_paths,
    create_ts_mail,
    find_recipient_for_order,
)
# 메일 CLI 배선은 delivery_status.py·create_oc.py와 공유한다 (po_generator/mail_cli.py).
from po_generator.mail_cli import (
    MailOptions,
    add_mail_arguments,
    collect_customer_po as _collect_customer_po,
    confirm_recipient,
    format_mail_date as _format_mail_date,
    prepare_mail_options as _prepare_mail_options,
    report_mail_result,
)
from po_generator.services import DocumentService, GenerationStatus

logger = logging.getLogger(__name__)


def foreign_biz_numbers(order_data: pd.Series, items_df: pd.DataFrame | None) -> list[str]:
    """문서에 실린 아이템 중 **대표 거래처가 아닌** 사업자번호들

    DN 번호를 잘못 재사용하면 한 DN에 두 거래처가 들어간다 — 2026-08-06 실측:
    `DND-2026-0748`에 씨앤케이(SOD-2026-0729)와 오토밸브(SOD-2026-0743)가 함께 있고,
    `DND-2026-0328`(2026-04-14)도 마찬가지다. 그 문서를 그대로 첨부하면 **남의 거래
    내역(품목·수량·단가)이 고객에게 나간다.**

    Args:
        order_data: 대표 행 (이 문서의 주인)
        items_df: 문서에 실린 전체 아이템

    Returns:
        섞여 들어온 사업자번호 목록 (없으면 빈 목록)
    """
    if items_df is None or items_df.empty:
        return []

    biz_col = resolve_column(items_df.columns, 'biz_no')
    if biz_col is None:
        return []

    own = normalize_biz_no(get_value(order_data, 'biz_no', ''))
    others = {normalize_biz_no(value) for value in items_df[biz_col]}
    others.discard('')
    others.discard(own)
    return sorted(others)


def describe_foreign(items_df: pd.DataFrame, biz_numbers: Sequence[str]) -> str:
    """섞여 들어온 거래처 표기 — '오토밸브(1178176942)'

    번호만 찍으면 사람이 어느 줄을 고쳐야 할지 시트에서 찾기 어렵다.
    """
    name_col = resolve_column(items_df.columns, 'customer_name')
    biz_col = resolve_column(items_df.columns, 'biz_no')

    labels: list[str] = []
    for digits in biz_numbers:
        name = ''
        if name_col is not None and biz_col is not None:
            hits = items_df[items_df[biz_col].map(normalize_biz_no) == digits]
            if not hits.empty:
                name = str(hits.iloc[0][name_col]).strip()
        labels.append(f"{name}({digits})" if name and name.lower() != 'nan' else digits)
    return ', '.join(labels)


def _mail_ts(
    order_data: pd.Series,
    output_file: Path | Sequence[Path],
    doc_id: str,
    opts: MailOptions,
    items_df: pd.DataFrame | None = None,
    date_str: str | None = None,
) -> bool:
    """생성된 거래명세표를 메일로 발송/초안 생성

    수신자 확인 관문(조회 → 표시 → y/N)은 `mail_cli.confirm_recipient()`가 소유한다.
    메일 실패는 거래명세표 생성 성공을 뒤엎지 않습니다 (경고만 출력).

    첨부는 **여러 장**일 수 있다 — 같은 날 같은 거래처로 나간 DN이 여러 건이면
    문서는 DN별 1장이되 메일은 한 통이다 (`generate_ts_batch`).

    Args:
        order_data: 주문 데이터 (사업자번호/고객명/출고일 포함)
        output_file: 생성된 거래명세표 경로 (여러 장이면 목록)
        doc_id: DN_ID 또는 선수금_ID (묶음이면 'DND-... 외 N건')
        opts: 메일 옵션
        items_df: 문서에 실린 전체 아이템 (발주번호 수집용)
        date_str: 제목/본문 날짜 (없으면 order_data의 출고일)

    Returns:
        메일 생성/발송 성공 여부
    """
    if not opts.enabled:
        return True

    # 첨부에 남의 거래처 라인이 섞여 있으면 보내지 않는다 — 네 경로(단건·월합·묶음·선수금)가
    # 모두 여기로 모이므로 관문도 여기 하나만 둔다.
    foreign = foreign_biz_numbers(order_data, items_df)
    if foreign:
        print("  [메일 중단] 문서에 다른 거래처 라인이 섞여 있습니다: "
              f"{describe_foreign(items_df, foreign)}")
        print(f"             같은 DN 번호를 두 거래처에 쓰지 않았는지 시트를 확인하세요 ({doc_id}).")
        return False

    attachments = as_paths(output_file)
    biz_no = get_value(order_data, 'biz_no', '(사업자번호 없음)')
    customer = get_value(order_data, 'customer_name', '')
    # 제목/본문 날짜: 출고일 우선, 없으면 오늘
    # (선수금 거래명세표는 SO_국내 기반이라 출고일 자체가 없다)
    date_str = date_str or _format_mail_date(get_value(order_data, 'dispatch_date', None))
    customer_po = _collect_customer_po(order_data, items_df)

    recipient = confirm_recipient(
        opts,
        lambda df: find_recipient_for_order(order_data, df),
        missing_lines=[
            f"  [메일 생략] 수신자 미등록 — {customer} / {biz_no}",
            f"             {opts.sheet_label} 시트에 해당 사업자번호의 이메일을 입력하세요.",
        ],
        info_lines=[f"  발주번호: {customer_po}"],
    )
    if recipient is None:
        return False

    try:
        result = create_ts_mail(
            xlsx_path=attachments,
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


@dataclass(frozen=True, eq=False)  # DataFrame을 담고 있어 값 비교는 쓰지 않는다
class BuiltTS:
    """생성이 끝난 거래명세표 한 장 — 묶음 메일에 필요한 것만"""
    doc_id: str
    output_file: Path
    order_data: pd.Series      # 대표 행 (사업자번호·고객명·출고일)
    # 문서에 실린 전체 아이템 (발주번호 수집용). **단일 아이템이어도 1행짜리 DataFrame**이다 —
    # `OrderData.items_df`는 단일이면 None이라 그대로 담으면 묶음 메일의 concat이 죽는다.
    items_df: pd.DataFrame


def _report_generation_error(doc_id: str, result) -> None:
    """생성 실패 사유 출력 (단건·묶음 공통)"""
    if result.status == GenerationStatus.FILE_ERROR:
        print(f"  [오류] {result.errors[0] if result.errors else result.message}")
    else:
        print(f"  [오류] {result.message}")


def _build_ts_from_dn(dn_id: str, service: DocumentService) -> BuiltTS | None:
    """DN 하나로 거래명세표를 만든다 (메일은 하지 않는다)

    단건 실행과 묶음 실행이 **같은 생성 경로**를 쓰도록 떼어 놓은 부분이다 —
    갈라지면 "묶어 보낼 때만 다른 문서가 나간다"가 된다.

    Args:
        dn_id: DN_ID
        service: 문서 서비스 (묶음 실행에서 재사용)

    Returns:
        BuiltTS 또는 None (조회/생성 실패 — 사유는 화면에 출력됨)
    """
    print(f"\n{'=' * 50}")
    print(f"거래명세표 생성 (납품): {dn_id}")
    print('=' * 50)

    # 1. DN 데이터 검색 및 정보 출력
    order_data = service.finder.find_dn(dn_id)
    if order_data is None:
        print(f"  [오류] '{dn_id}'를 찾을 수 없습니다.")
        return None

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
    if not result.success:
        _report_generation_error(dn_id, result)
        return None

    print(f"  -> 거래명세표 생성 완료: {result.output_file.name}")
    return BuiltTS(
        doc_id=dn_id,
        output_file=result.output_file,
        order_data=order_data.first_item,
        items_df=order_data.all_items,
    )


def generate_ts_from_dn(
    dn_id: str,
    df_dn: pd.DataFrame,
    mail_opts: MailOptions | None = None,
) -> bool:
    """DN 기반 거래명세표 생성 + 건별 메일

    Args:
        dn_id: DN_ID
        df_dn: DN 데이터 (하위 호환용, 실제로는 사용하지 않음)
        mail_opts: 메일 발송 옵션 (None이면 발송 안 함)

    Returns:
        성공 여부 (메일 실패는 성공 여부에 영향 없음)
    """
    built = _build_ts_from_dn(dn_id, DocumentService())
    if built is None:
        return False

    _mail_ts(built.order_data, built.output_file, dn_id,
             mail_opts or MailOptions.disabled(), items_df=built.items_df)
    return True


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


def _build_ts_from_adv(advance_id: str, service: DocumentService) -> BuiltTS | None:
    """선수금 거래명세표를 만든다 (메일은 하지 않는다) — `_build_ts_from_dn`의 선수금판

    Args:
        advance_id: 선수금_ID
        service: 문서 서비스 (묶음 실행에서 재사용)

    Returns:
        BuiltTS 또는 None (조회/생성 실패 — 사유는 화면에 출력됨)
    """
    print(f"\n{'=' * 50}")
    print(f"거래명세표 생성 (선수금): {advance_id}")
    print('=' * 50)

    # 1. SO 데이터 로드 (선수금_ID -> SO_ID -> SO 아이템들)
    result = service.finder.find_so_for_advance(advance_id)
    if result is None:
        print(f"  [오류] '{advance_id}'를 찾을 수 없습니다.")
        return None

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
    if not gen_result.success:
        _report_generation_error(advance_id, gen_result)
        return None

    print(f"  -> 선수금 거래명세표 생성 완료: {gen_result.output_file.name}")
    return BuiltTS(
        doc_id=advance_id,
        output_file=gen_result.output_file,
        order_data=order_data.first_item,
        items_df=order_data.all_items,
    )


def generate_ts_from_adv(advance_id: str, mail_opts: MailOptions | None = None) -> bool:
    """선수금 거래명세표 생성 + 건별 메일 (SO_국내 데이터 사용)

    Args:
        advance_id: 선수금_ID
        mail_opts: 메일 발송 옵션 (None이면 발송 안 함)

    Returns:
        성공 여부 (메일 실패는 성공 여부에 영향 없음)
    """
    built = _build_ts_from_adv(advance_id, DocumentService())
    if built is None:
        return False

    _mail_ts(built.order_data, built.output_file, advance_id,
             mail_opts or MailOptions.disabled(), items_df=built.items_df)
    return True


# === 하루치 출고를 거래처별 한 통으로 (--date / --one-mail) ===

# 연도까지 적은 형식 / 연도를 생략한 형식 (생략하면 올해)
_DATE_FORMATS_WITH_YEAR: tuple[str, ...] = ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y%m%d')
_DATE_FORMATS_NO_YEAR: tuple[str, ...] = ('%m-%d', '%m/%d', '%m.%d')


def parse_dispatch_date(text: str) -> pd.Timestamp | None:
    """출고일 조회어 → 날짜

    '2026-08-06' · '2026/8/6' · '20260806' · '08-06' · '8/6'을 받습니다.
    연도를 생략하면 올해로 읽습니다 (출고 회신은 대개 당일·전날 것이라).

    Args:
        text: 날짜 문자열

    Returns:
        pd.Timestamp (자정 기준) 또는 None (형식 불명)
    """
    raw = (text or '').strip()
    if not raw:
        return None

    for fmt in _DATE_FORMATS_WITH_YEAR:
        try:
            return pd.Timestamp(datetime.strptime(raw, fmt))
        except ValueError:
            continue

    for fmt in _DATE_FORMATS_NO_YEAR:
        try:
            parsed = datetime.strptime(raw, fmt)
        except ValueError:
            continue
        return pd.Timestamp(datetime(datetime.now().year, parsed.month, parsed.day))

    return None


def filter_customer(rows: pd.DataFrame, query: str) -> pd.DataFrame:
    """조회어(사업자번호 또는 거래처명 부분일치)로 행 좁히기

    판정 규칙은 `delivery_status.resolve_customer()`와 같습니다 —
    숫자가 `BIZ_NO_MIN_DIGITS`자리 이상이면 사업자번호, 아니면 이름 부분일치.

    Args:
        rows: DN 행
        query: 조회어

    Returns:
        좁혀진 행 (컬럼을 못 찾으면 빈 DataFrame)
    """
    digits = normalize_biz_no(query)
    if len(digits) >= BIZ_NO_MIN_DIGITS:
        biz_col = resolve_column(rows.columns, 'biz_no')
        if biz_col is None:
            return rows.iloc[0:0]
        return rows[rows[biz_col].map(normalize_biz_no) == digits]

    name_col = resolve_column(rows.columns, 'customer_name')
    if name_col is None:
        return rows.iloc[0:0]
    keyword = query.strip().lower()
    return rows[
        rows[name_col].astype(str).str.lower().str.contains(keyword, regex=False, na=False)
    ]


def select_dn_ids(
    df_dn: pd.DataFrame,
    target_date: pd.Timestamp,
    customer: str = '',
) -> list[str]:
    """출고일(+거래처)로 DN_ID 고르기 — 대상 목록을 화면에 보여준다

    Args:
        df_dn: DN 데이터
        target_date: 출고일
        customer: 거래처 조회어 (빈 값이면 그날 전체)

    Returns:
        DN_ID 목록 (시트 순서, 중복 제거)
    """
    date_col = resolve_column(df_dn.columns, 'dispatch_date')
    dn_col = resolve_column(df_dn.columns, 'dn_id')
    name_col = resolve_column(df_dn.columns, 'customer_name')
    if date_col is None or dn_col is None:
        print("  [오류] DN 시트에서 출고일/DN_ID 컬럼을 찾을 수 없습니다.")
        return []

    dates = pd.to_datetime(df_dn[date_col], errors='coerce').dt.normalize()
    rows = df_dn[dates == target_date.normalize()]
    if customer:
        rows = filter_customer(rows, customer)

    date_label = target_date.strftime('%Y-%m-%d')
    if rows.empty:
        scope = f"{date_label} 출고분" + (f" / '{customer}'" if customer else "")
        print(f"\n[안내] {scope}에 해당하는 DN이 없습니다.")
        recent = sorted(dates.dropna().unique())[-5:]
        if len(recent) > 0:
            labels = ', '.join(pd.Timestamp(d).strftime('%Y-%m-%d') for d in recent)
            print(f"       최근 출고일: {labels}")
        return []

    dn_ids = list(dict.fromkeys(rows[dn_col].dropna().astype(str)))

    print(f"\n{'=' * 50}")
    print(f"{date_label} 출고분: DN {len(dn_ids)}건", end='')
    if name_col is not None:
        counts = (
            rows.drop_duplicates(subset=[dn_col])[name_col]
            .astype(str).value_counts()
        )
        print(f" / 거래처 {len(counts)}곳")
        for customer_name, count in counts.items():
            print(f"  {customer_name}  DN {count}건")
    else:
        print()
    print('=' * 50)

    return dn_ids


def group_by_customer(built: Sequence[BuiltTS]) -> list[list[BuiltTS]]:
    """생성된 거래명세표를 거래처별로 묶는다 (첫 등장 순서 유지)

    묶는 키는 **정규화한 사업자번호**다. 이름으로 묶으면 표기 차이('(주)' 유무, 공백)로
    같은 거래처가 갈라져 메일이 두 통 가고, 반대로 이름이 비슷한 남남이 한 통에 담긴다.
    사업자번호가 비어 있는 건은 서로 묶지 않는다 — 모르는 것끼리 합치면 남의 명세표가 붙는다.

    Args:
        built: 생성된 거래명세표들

    Returns:
        거래처별 묶음 목록
    """
    groups: dict[str, list[BuiltTS]] = {}
    for idx, one in enumerate(built):
        biz_no = normalize_biz_no(get_value(one.order_data, 'biz_no', ''))
        key = biz_no or f"_unknown_{idx}"
        groups.setdefault(key, []).append(one)
    return list(groups.values())


def _group_doc_label(doc_ids: Sequence[str]) -> str:
    """묶음 메일의 문서 식별자 표기 ('DND-2026-0742 외 7건')"""
    if len(doc_ids) == 1:
        return doc_ids[0]
    return f"{doc_ids[0]} 외 {len(doc_ids) - 1}건"


def _group_date_str(group: Sequence[BuiltTS]) -> str | None:
    """묶음 제목/본문에 쓸 날짜 — 묶음 내 **가장 늦은 출고일** (월합과 같은 규칙)

    Returns:
        'YYYY-MM-DD' 또는 None (출고일이 하나도 없으면 호출부가 오늘로 폴백)
    """
    stamps = [
        stamp for stamp in (
            pd.to_datetime(get_value(one.order_data, 'dispatch_date', None), errors='coerce')
            for one in group
        )
        if not pd.isna(stamp)
    ]
    return _format_mail_date(max(stamps)) if stamps else None


def _mail_group(group: Sequence[BuiltTS], mail_opts: MailOptions) -> bool:
    """한 거래처 묶음을 메일 한 통으로 (첨부 = 묶음 전체)"""
    first = group[0]
    customer = get_value(first.order_data, 'customer_name', '(고객명 없음)')

    print(f"\n{'-' * 50}")
    print(f"메일 1통: {customer} — 거래명세표 {len(group)}장")
    for one in group:
        print(f"  {one.doc_id}  {one.output_file.name}")

    return _mail_ts(
        first.order_data,
        [one.output_file for one in group],
        _group_doc_label([one.doc_id for one in group]),
        mail_opts,
        # 발주번호는 묶음 전체에서 모은다 (mail_cli.collect_customer_po)
        items_df=pd.concat([one.items_df for one in group], ignore_index=True),
        date_str=_group_date_str(group),
    )


def generate_ts_batch(
    doc_ids: Sequence[str],
    mail_opts: MailOptions | None = None,
) -> bool:
    """여러 건을 생성한 뒤 **거래처별로 묶어** 메일 한 통씩

    문서는 지금과 똑같이 DN(선수금)마다 1장이다 — 달라지는 건 메일뿐이다.
    같은 날 한 거래처로 8건이 나가면 첨부 8개짜리 한 통이 된다.

    Args:
        doc_ids: DN_ID/선수금_ID 목록
        mail_opts: 메일 발송 옵션 (None이면 발송 안 함)

    Returns:
        전 건 생성 성공 여부 (메일 실패는 영향 없음)
    """
    mail_opts = mail_opts or MailOptions.disabled()
    service = DocumentService()

    built: list[BuiltTS] = []
    for doc_id in doc_ids:
        if detect_id_type(doc_id) == 'ADV':
            one = _build_ts_from_adv(doc_id, service)
        else:
            one = _build_ts_from_dn(doc_id, service)
        if one is not None:
            built.append(one)

    if not built:
        print("\n[오류] 생성된 거래명세표가 없습니다.")
        return False

    # 남의 거래처 라인이 섞인 문서는 첨부에서 뺀다 — 나머지는 그대로 나간다.
    # (문서 자체는 만들어 둔다: 사람이 열어 보고 시트를 고쳐야 하니까)
    mailable: list[BuiltTS] = []
    for one in built:
        foreign = foreign_biz_numbers(one.order_data, one.items_df)
        if foreign:
            print(f"\n[경고] {one.doc_id}: 다른 거래처 라인이 섞여 있습니다 — "
                  f"{describe_foreign(one.items_df, foreign)} (메일 첨부에서 제외)")
            print("       같은 DN 번호를 두 거래처에 쓰지 않았는지 시트를 확인하세요.")
            continue
        mailable.append(one)

    groups = group_by_customer(mailable)

    print(f"\n{'=' * 50}")
    summary = f"완료: {len(built)}/{len(doc_ids)}건 거래명세표 생성"
    if mail_opts.enabled:
        summary += f" / 메일 {len(groups)}통 (거래처별)"
        if len(mailable) != len(built):
            summary += f" / 첨부 제외 {len(built) - len(mailable)}장"
    print(summary)
    print(f"출력 폴더: {TS_OUTPUT_DIR}")
    print('=' * 50)

    if mail_opts.enabled:
        for group in groups:
            _mail_group(group, mail_opts)

    return len(built) == len(doc_ids)


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

하루치 출고를 거래처별 메일 한 통으로 (문서는 DN별 1장 그대로, 첨부만 여러 개):
  python create_ts.py --date 2026-08-06 --customer 씨앤케이   # 그 거래처만
  python create_ts.py --date 2026-08-06                      # 그날 전체, 거래처별 1통씩
  python create_ts.py DN-2026-0001 DN-2026-0002 --one-mail   # 명시 ID를 묶어 1통

  --merge는 '문서'를 한 장으로 합치고, --one-mail은 문서는 그대로 두고 '메일'만 묶습니다.

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
        '--date',
        metavar='YYYY-MM-DD',
        help='그날 출고분 전체를 대상으로 (예: 2026-08-06, 8/6). 메일은 거래처별 1통',
    )

    parser.add_argument(
        '--customer',
        metavar='조회어',
        help='--date 안에서 거래처 한 곳만 (사업자번호 또는 거래처명 부분일치)',
    )

    parser.add_argument(
        '--one-mail',
        action='store_true',
        help='문서는 건별로 만들고 메일만 거래처별 한 통으로 묶기 (첨부 여러 개)',
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
    """CLI 인자로 거래명세표 메일 옵션 구성

    판정 규칙은 `mail_cli.prepare_mail_options()`가 소유하고, 여기서는 거래명세표
    상수만 넘깁니다 (수신자 마스터는 기본값인 `Customer_국내`).

    Args:
        args: 파싱된 CLI 인자

    Returns:
        MailOptions
    """
    return _prepare_mail_options(
        args,
        attach_format=TS_MAIL_ATTACH_FORMAT,
        fixed_cc=TS_MAIL_CC,
    )


def validate_selection_args(args: argparse.Namespace) -> str | None:
    """대상 선택 인자들의 조합 검사

    `--merge`는 **문서**를 한 장으로 합치고, `--one-mail`/`--date`는 문서는 그대로 두고
    **메일**만 묶는다. 뜻이 정반대라 섞이면 무엇이 나갔는지 사람이 알 수 없다.

    Args:
        args: 파싱된 CLI 인자

    Returns:
        오류 메시지 또는 None (문제 없음)
    """
    if args.merge and (args.one_mail or args.date):
        return (
            "--merge와 --one-mail/--date는 함께 쓸 수 없습니다.\n"
            "       --merge    : 여러 DN을 '문서' 한 장으로 합침\n"
            "       --one-mail : 문서는 DN별 1장 그대로, '메일'만 거래처별 한 통"
        )
    if args.customer and not args.date:
        return (
            "--customer는 --date와 함께 씁니다 "
            "(날짜 없이 거래처만 주면 그 거래처의 과거 DN 전부가 대상이 됩니다)."
        )
    if args.date and args.doc_ids:
        return "--date와 ID를 함께 줄 수 없습니다 (날짜로 고르거나, ID를 직접 주거나)."
    return None


def main() -> int:
    """메인 함수

    Returns:
        종료 코드 (0: 성공, 1: 실패)
    """
    parser = create_argument_parser()
    args = parser.parse_args()

    error = validate_selection_args(args)
    if error:
        print(f"[오류] {error}")
        return 1

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

    # --date: 그날 출고분을 자동으로 고른다 (DN 번호를 손으로 적지 않는다)
    if args.date:
        target_date = parse_dispatch_date(args.date)
        if target_date is None:
            print(f"[오류] 날짜 형식을 알 수 없습니다: '{args.date}' (예: 2026-08-06, 8/6)")
            return 1

        print(f"DN: {len(df_dn)}건, PMT: {len(df_pmt)}건 로드 완료")
        dn_ids = select_dn_ids(df_dn, target_date, args.customer or '')
        if not dn_ids:
            return 1

        # 날짜로 고른다는 것 자체가 "그날치를 거래처별로 묶는다"는 뜻이다
        return 0 if generate_ts_batch(dn_ids, prepare_mail_options(args)) else 1

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

    # --one-mail 옵션: 문서는 건별로, 메일만 거래처별 한 통으로
    if args.one_mail:
        doc_ids = list(dict.fromkeys(args.doc_ids))
        if len(doc_ids) != len(args.doc_ids):
            dups = [d for d in doc_ids if args.doc_ids.count(d) > 1]
            print(f"\n[경고] 중복 ID 제거됨: {dups}")

        return 0 if generate_ts_batch(doc_ids, mail_opts) else 1

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
