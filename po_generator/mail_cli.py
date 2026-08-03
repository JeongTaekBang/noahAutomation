"""
메일 발송 CLI 공통 배선
========================

문서를 만든 뒤 "메일 보낼까요?"를 묻는 CLI들이 공유하는 상태 기계입니다.
`create_ts.py`(거래명세표)·`delivery_status.py`(납기현황)·`create_oc.py`(OC)가
같은 규칙을 씁니다.

여기에 모아 둔 이유
-------------------
모드 판정 규칙(비대화형이면 묻지 않기, `--send`가 `--mail`을 이기고 `--no-mail`이 전부를
이기기)과 수신자 마스터 지연 로딩은 CLI마다 똑같이 필요하다. 복사해 두면 한쪽만 고쳐져
갈라지고, 그 갈라짐이 **고객에게 메일이 나가는 경로**에서 벌어진다.

문서마다 다른 것(어느 마스터를 읽는지, 첨부 형식, 고정 CC)은 전부 인자다 —
국내 문서는 `Customer_국내`, OC는 `Customer_해외`를 읽지만 그 위의 규칙은 같다.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import datetime
from enum import Enum
from typing import Callable, Sequence

import pandas as pd

from po_generator.config import CUSTOMER_DOMESTIC_SHEET
from po_generator.mailer import (
    NO_VALUE,
    MailBackend,
    MailConfigError,
    MailResult,
    Recipient,
    resolve_backend,
)
from po_generator.utils import get_value, load_customer_domestic, resolve_column


class MailMode(str, Enum):
    """메일 동작 모드"""
    OFF = 'off'      # --no-mail 또는 비대화형 실행: 묻지도 보내지도 않음
    ASK = 'ask'      # 기본: 수신자를 보여주고 y/N 확인
    DRAFT = 'draft'  # --mail: 확인 없이 초안 열기
    SEND = 'send'    # --send: 확인 없이 즉시 발송


# 긍정 응답 — 'ㅛ'는 한글 IME 상태에서 y를 누른 경우
_YES_ANSWERS: frozenset[str] = frozenset({'y', 'yes', 'ㅛ', '네', 'ㅇ'})


@dataclass
class MailOptions:
    """메일 발송 옵션

    수신자 마스터는 **실제로 필요할 때 한 번만** 로드합니다.
    (메일을 안 쓰는 실행에 Excel 로딩 비용을 물리지 않기 위함)

    어느 마스터를 읽을지는 `loader`/`sheet_label`로 주입합니다 — 기본은 국내
    (`Customer_국내`)이고, OC는 `Customer_해외` 로더를 넘깁니다.
    """
    mode: MailMode = MailMode.OFF
    backend: MailBackend | None = None  # None이면 mailer가 자동 판정
    df_customer: pd.DataFrame | None = None
    # None = 국내 기본 로더. 기본값에 함수를 **직접 박지 않는다** — 박으면 클래스 정의
    # 시점에 함수 객체가 굳어, 모듈 속성을 갈아끼우는 쪽(테스트의 monkeypatch, 런타임
    # 교체)이 조용히 무시된다. 실제로 그렇게 만들었다가 기존 테스트 3건이 깨졌다.
    # 설정을 호출 시점에 읽는 `mailer.resolve_backend()`와 같은 취지.
    loader: Callable[[], pd.DataFrame] | None = None
    sheet_label: str = CUSTOMER_DOMESTIC_SHEET
    _skip_reason: str = ''  # 설정 문제로 이번 실행 내내 메일을 접은 이유

    @classmethod
    def disabled(cls) -> MailOptions:
        return cls()

    @property
    def enabled(self) -> bool:
        return self.mode is not MailMode.OFF

    @property
    def send(self) -> bool:
        return self.mode is MailMode.SEND

    @property
    def ask(self) -> bool:
        return self.mode is MailMode.ASK

    def customer_master(self) -> pd.DataFrame | None:
        """수신자 마스터 (지연 로딩 + 캐시)

        로딩/설정 실패는 이번 실행 전체에 대해 한 번만 안내하고 이후 조용히 건너뜁니다.

        Returns:
            DataFrame 또는 None (사용 불가)
        """
        if self._skip_reason:
            return None
        if self.df_customer is not None:
            return self.df_customer

        loader = self.loader or load_customer_domestic
        try:
            df = loader()
        except (FileNotFoundError, ValueError) as e:
            self._skip_reason = str(e)
            print(f"  [메일 생략] 수신자 마스터를 읽을 수 없습니다: {e}")
            return None

        if resolve_column(df.columns, 'customer_email') is None:
            self._skip_reason = 'no-email-column'
            print(f"  [메일 생략] '{self.sheet_label}' 시트에 이메일 컬럼이 없습니다.")
            print("             시트 끝에 '수신자 이메일' 컬럼을 추가하면 메일 발송을 물어봅니다.")
            return None

        self.df_customer = df
        return df


def confirm(question: str) -> bool:
    """y/N 확인 (기본값 N)

    파이프/리다이렉트로 stdin이 닫혀 있거나 Ctrl+C면 '아니오'로 처리합니다.

    Args:
        question: 표시할 질문

    Returns:
        사용자가 긍정했는지 여부
    """
    try:
        answer = input(question).strip().lower()
    except (EOFError, KeyboardInterrupt):
        print()
        return False
    return answer in _YES_ANSWERS


def resolve_mail_mode(args: argparse.Namespace, is_tty: bool) -> MailMode:
    """CLI 인자 + 실행 환경으로 메일 모드 결정

    비대화형(배치/파이프) 실행에서 기본 ASK를 그대로 두면 input()에서 멈추므로
    OFF로 낮춥니다. 명시적 --mail/--send는 그대로 존중합니다.

    Args:
        args: 파싱된 CLI 인자
        is_tty: stdin이 터미널인지

    Returns:
        MailMode
    """
    if args.no_mail:
        return MailMode.OFF
    if args.send:
        return MailMode.SEND
    if args.mail:
        return MailMode.DRAFT
    return MailMode.ASK if is_tty else MailMode.OFF


def show_recipient(recipient: Recipient) -> None:
    """수신자/참조를 화면에 보여준다 — 발송 확인 직전의 오발송 차단 장치

    고객에게 나가는 메일은 항상 사람이 "누구에게"를 눈으로 한 번 본다.
    CLI마다 이 표시가 갈리면 확인 습관도 갈리므로 한 곳에 둔다.
    """
    print(f"  받는사람: {recipient.to_line}")
    if recipient.cc:
        print(f"  참조    : {recipient.cc_line}")


def report_mail_result(result: MailResult, want_send: bool) -> bool:
    """메일 생성/발송 결과를 콘솔에 보고

    Args:
        result: mailer가 돌려준 결과
        want_send: 사용자가 즉시 발송을 원했는지 (.eml은 초안으로 강등되므로 안내가 필요)

    Returns:
        성공 여부 (호출부의 반환값으로 그대로 쓴다)
    """
    if result.success:
        attach_names = ', '.join(p.name for p in result.attachments)
        verb = "메일 발송 완료" if result.sent else "메일 초안 생성 (메일 창에서 [보내기] 확인)"
        print(f"  -> {verb}: {attach_names}")
        if want_send and not result.sent:
            print("     [주의] 자동 발송이 안 되는 방식이라 초안까지만 진행했습니다.")
        return True

    print(f"  [메일 실패] {result.message}")
    return False


def add_mail_arguments(parser: argparse.ArgumentParser, doc_label: str) -> None:
    """`--mail` / `--send` / `--no-mail` 인자 추가

    Args:
        parser: 대상 파서
        doc_label: 도움말에 쓸 문서 이름 (예: '거래명세표', '납기현황')
    """
    parser.add_argument(
        '--mail',
        action='store_true',
        help=f'확인 없이 {doc_label} 메일 초안 열기',
    )
    parser.add_argument(
        '--send',
        action='store_true',
        help='확인 없이 즉시 발송 (Outlook COM 환경에서만 동작)',
    )
    parser.add_argument(
        '--no-mail',
        action='store_true',
        help='메일을 묻지도 보내지도 않음 (배치용)',
    )


def prepare_mail_options(
    args: argparse.Namespace,
    attach_format: str,
    fixed_cc: Sequence[str] = (),
    loader: Callable[[], pd.DataFrame] | None = None,
    sheet_label: str = CUSTOMER_DOMESTIC_SHEET,
) -> MailOptions:
    """CLI 인자로 메일 옵션 구성 + 이번 실행의 동작을 화면에 알린다

    마스터 로딩은 실제로 메일이 필요한 시점까지 미룹니다
    (메일을 쓰지 않는 실행에 Excel 로딩 비용을 물리지 않기 위함).

    Args:
        args: 파싱된 CLI 인자 (`add_mail_arguments`로 붙인 것)
        attach_format: 첨부 형식 표기용 ('pdf' | 'xlsx' | 'both')
        fixed_cc: 고정 참조 목록 (표시용 — 실제 적용은 mailer가 한다)
        loader: 수신자 마스터 로더 (None이면 국내 기본)
        sheet_label: 마스터 시트 이름 (안내 메시지용)

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
    # 참조자는 거래처마다 달라서 마스터의 '참조 이메일'로 건별 관리한다.
    # 고정 CC는 모든 거래처에 공통으로 붙일 주소가 있을 때만 쓰는 선택 항목.
    cc_note = f" / 고정 CC {len(fixed_cc)}명" if fixed_cc else ""
    print(f"\n메일: {label} / 첨부 {attach_format.upper()} / "
          f"방식 {backend_label}{cc_note}")

    # .eml은 작성 창을 띄우는 방식이라 자동 발송이 불가능하다
    if mode is MailMode.SEND and backend is MailBackend.EML:
        print("  [주의] .eml 방식은 자동 발송을 지원하지 않습니다 — 초안까지만 진행됩니다.")
        print("         (Outlook COM 사용 불가 환경 — 클래식 Outlook이 없거나 실행 실패)")

    return MailOptions(
        mode=mode, backend=backend, loader=loader, sheet_label=sheet_label,
    )


def collect_customer_po(
    order_data: pd.Series,
    items_df: pd.DataFrame | None = None,
) -> str:
    """거래처 발주번호 표기 문자열

    한 문서에 여러 발주가 묶일 수 있으므로(월합 거래명세표, 다중 아이템 SO)
    **전체 아이템에서** 모읍니다. 첫 건만 쓰면 엉뚱한 번호 하나만 나간다.

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


def format_mail_date(value: object) -> str:
    """메일 제목/본문에 쓸 날짜 문자열

    날짜가 비어 있거나(선수금·OC처럼 기준일이 없는 문서) 파싱되지 않으면 오늘을 씁니다.
    빈 값은 None/NaT/NaN 모두로 들어올 수 있어 한 곳에서 흡수합니다. 날짜 표기
    형식('YYYY-MM-DD')의 단일 소유처 — 세 CLI가 제각기 찍으면 표기가 갈립니다.

    Args:
        value: 날짜 값 (None이면 오늘)

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


def confirm_recipient(
    opts: MailOptions,
    find: Callable[[pd.DataFrame], Recipient | None],
    missing_lines: Sequence[str],
    info_lines: Sequence[str] = (),
) -> Recipient | None:
    """수신자 조회 → 화면 표시 → (ASK면) y/N 확인 — 오발송 차단 관문

    세 CLI(거래명세표·납기현황·OC)가 같은 순서로 밟는 단계라 한 곳에 둔다:
    마스터 로딩 실패/설정 오류/미등록/사용자 거부 어느 쪽이든 **None = 보내지 않는다**
    이고, 사유는 여기서 이미 화면에 안내했으므로 호출부는 False만 돌려주면 된다.

    Args:
        opts: 메일 옵션 (마스터 지연 로딩 포함)
        find: 마스터 DataFrame을 받아 수신자를 찾는 함수 (조인키는 호출부가 안다)
        missing_lines: 수신자 미등록일 때 출력할 안내 줄들
        info_lines: 수신자 아래에 보여줄 추가 정보 줄들 (발주번호 등)

    Returns:
        확인까지 끝난 Recipient, 또는 None (보내지 않음)
    """
    df_customer = opts.customer_master()
    if df_customer is None:
        return None

    try:
        recipient = find(df_customer)
    except MailConfigError as e:
        print(f"  [메일 오류] {e}")
        return None

    if recipient is None:
        for line in missing_lines:
            print(line)
        return None

    # 누구에게 나가는지 먼저 보여주고 확인받는다 (오발송 차단)
    show_recipient(recipient)
    for line in info_lines:
        print(line)

    if opts.ask and not confirm("  이메일을 발송하시겠습니까? [y/N]: "):
        print("  -> 메일 생략")
        return None

    return recipient
