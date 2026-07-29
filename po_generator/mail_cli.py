"""
메일 발송 CLI 공통 배선
========================

문서를 만든 뒤 "메일 보낼까요?"를 묻는 CLI들이 공유하는 상태 기계입니다.
`create_ts.py`(거래명세표)와 `delivery_status.py`(납기현황)가 같은 규칙을 씁니다.

여기에 모아 둔 이유
-------------------
모드 판정 규칙(비대화형이면 묻지 않기, `--send`가 `--mail`을 이기고 `--no-mail`이 전부를
이기기)과 수신자 마스터 지연 로딩은 CLI마다 똑같이 필요하다. 복사해 두면 한쪽만 고쳐져
갈라지고, 그 갈라짐이 **고객에게 메일이 나가는 경로**에서 벌어진다.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from enum import Enum

import pandas as pd

from po_generator.config import CUSTOMER_DOMESTIC_SHEET
from po_generator.mailer import MailBackend
from po_generator.utils import load_customer_domestic, resolve_column


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

    Customer_국내 마스터는 **실제로 필요할 때 한 번만** 로드합니다.
    (메일을 안 쓰는 실행에 Excel 로딩 비용을 물리지 않기 위함)
    """
    mode: MailMode = MailMode.OFF
    backend: MailBackend | None = None  # None이면 mailer가 자동 판정
    df_customer: pd.DataFrame | None = None
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
        """Customer_국내 마스터 (지연 로딩 + 캐시)

        로딩/설정 실패는 이번 실행 전체에 대해 한 번만 안내하고 이후 조용히 건너뜁니다.

        Returns:
            DataFrame 또는 None (사용 불가)
        """
        if self._skip_reason:
            return None
        if self.df_customer is not None:
            return self.df_customer

        try:
            df = load_customer_domestic()
        except (FileNotFoundError, ValueError) as e:
            self._skip_reason = str(e)
            print(f"  [메일 생략] 수신자 마스터를 읽을 수 없습니다: {e}")
            return None

        if resolve_column(df.columns, 'customer_email') is None:
            self._skip_reason = 'no-email-column'
            print(f"  [메일 생략] '{CUSTOMER_DOMESTIC_SHEET}' 시트에 이메일 컬럼이 없습니다.")
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
