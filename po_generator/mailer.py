"""거래명세표 메일 발송 모듈 (Outlook COM 기반)
================================================

거래명세표를 발행한 뒤, 사업자번호로 `Customer_국내`에서 수신자를 조회해
PDF를 첨부한 Outlook 메일을 만듭니다.

동작 원칙:
- 기본은 **초안 열기**(Display). `send=True`일 때만 즉시 발송.
- 메일 실패가 거래명세표 생성 성공을 뒤엎지 않도록, 호출부에서 별도 단계로 처리.
- Outlook 미설치 환경에서 import만으로 죽지 않도록 win32com은 지연 import.
"""

from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
from dataclasses import dataclass
from datetime import datetime
from email.message import EmailMessage
from enum import Enum
from html import escape as html_escape
from pathlib import Path

import pandas as pd

from po_generator.config import (
    CUSTOMER_DOMESTIC_SHEET,
    SUPPLIER_INFO,
    TS_MAIL_ATTACH_FORMAT,
    TS_MAIL_BACKEND,
    TS_MAIL_BODY,
    TS_MAIL_CC,
    TS_MAIL_SUBJECT,
)
from po_generator.excel_helpers import cleanup_temp_file, xlwings_app_context
from po_generator.utils import (
    get_value,
    load_customer_domestic,
    normalize_biz_no,
    resolve_column,
)

logger = logging.getLogger(__name__)

# Outlook MailItem 상수 (olMailItem)
OL_MAIL_ITEM = 0

# 값이 없을 때 본문에 표기할 문자열 (빈칸으로 두면 누락인지 없음인지 구분이 안 됨)
NO_VALUE = 'N/A'

# 메일 주소 구분자: 세미콜론/쉼표/줄바꿈 혼용 허용
_EMAIL_SPLIT_RE = re.compile(r'[;,\n\r]+')
# 형식 검증: 사람이 손으로 넣는 셀이라 오타를 걸러내되 과하게 엄격하지 않게
_EMAIL_RE = re.compile(r'^[^@\s]+@[^@\s]+\.[^@\s]+$')


class MailConfigError(Exception):
    """메일 설정/마스터 데이터가 준비되지 않은 경우"""


class MailBackend(str, Enum):
    """메일 작성 방식"""
    OUTLOOK = 'outlook'  # Outlook COM (classic Outlook 전용)
    EML = 'eml'          # .eml 초안 파일 + 기본 메일 앱으로 열기


@dataclass(frozen=True)
class Recipient:
    """거래명세표 메일 수신자"""
    biz_no: str
    customer_name: str
    to: tuple[str, ...]
    cc: tuple[str, ...]
    customer_name_en: str = ''  # 영문 메일용 (없으면 한글명으로 폴백)

    @property
    def display_name_en(self) -> str:
        """영문 표기 이름 (미등록이면 한글명)"""
        return self.customer_name_en or self.customer_name

    @property
    def to_line(self) -> str:
        return '; '.join(self.to)

    @property
    def cc_line(self) -> str:
        return '; '.join(self.cc)


@dataclass(frozen=True)
class MailResult:
    """메일 생성/발송 결과"""
    success: bool
    sent: bool           # True=즉시 발송, False=초안만 열림
    recipient: Recipient | None
    attachments: tuple[Path, ...] = ()
    message: str = ''
    backend: MailBackend | None = None


# === 메일 주소 파싱 ===

def split_emails(value: object) -> tuple[str, ...]:
    """셀 값에서 메일 주소 목록 추출

    한 셀에 여러 주소가 `;` `,` 줄바꿈으로 들어있어도 모두 분리합니다.
    형식이 깨진 값은 버리고 경고를 남깁니다.

    Args:
        value: 셀 값

    Returns:
        유효한 메일 주소 튜플 (중복 제거, 입력 순서 유지)
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ()

    text = str(value).strip()
    if not text or text.lower() == 'nan':
        return ()

    seen: dict[str, None] = {}
    for token in _EMAIL_SPLIT_RE.split(text):
        addr = token.strip().strip('<>').strip()
        if not addr:
            continue
        if not _EMAIL_RE.match(addr):
            logger.warning(f"메일 주소 형식 오류로 제외: '{addr}'")
            continue
        seen.setdefault(addr, None)

    return tuple(seen)


# === 수신자 조회 ===

def find_recipient(
    biz_no: object,
    df_customer: pd.DataFrame | None = None,
    fallback_name: str = '',
) -> Recipient | None:
    """사업자번호로 Customer_국내에서 메일 수신자 조회

    Args:
        biz_no: 사업자번호 (DN_국내의 Business registration number)
        df_customer: Customer_국내 DataFrame (없으면 로드)
        fallback_name: 마스터에 거래처명이 없을 때 쓸 이름 (보통 DN의 Customer name)

    Returns:
        Recipient (고정 CC 포함) 또는 None (사업자번호 미매칭/메일 미등록)

    Raises:
        MailConfigError: Customer_국내에 이메일 컬럼 자체가 없는 경우
    """
    key = normalize_biz_no(biz_no)
    if not key:
        logger.warning("사업자번호가 비어 있어 수신자를 조회할 수 없습니다.")
        return None

    if df_customer is None:
        df_customer = load_customer_domestic()

    email_col = resolve_column(df_customer.columns, 'customer_email')
    if email_col is None:
        raise MailConfigError(
            f"'{CUSTOMER_DOMESTIC_SHEET}' 시트에 이메일 컬럼이 없습니다.\n"
            f"  시트 끝에 '수신자 이메일' 컬럼을 추가하고 거래처별 주소를 입력하세요.\n"
            f"  (인식 가능한 헤더: 수신자 이메일 / 이메일 / 메일 / Email / E-mail / 담당자 이메일)"
        )
    cc_col = resolve_column(df_customer.columns, 'customer_email_cc')
    name_col = resolve_column(df_customer.columns, 'customer_name') or '거래처명'
    name_en_col = resolve_column(df_customer.columns, 'customer_name_en')

    # 정규화 컬럼은 load_customer_domestic()이 붙여주지만, 원본 시트를 그대로
    # 넘겨도 KeyError 대신 동작하도록 보강한다.
    if '_사업자번호_정규화' not in df_customer.columns:
        biz_col = resolve_column(df_customer.columns, 'biz_no')
        if biz_col is None:
            raise MailConfigError(
                f"'{CUSTOMER_DOMESTIC_SHEET}' 시트에서 사업자번호 컬럼을 찾을 수 없습니다."
            )
        df_customer = df_customer.assign(
            _사업자번호_정규화=df_customer[biz_col].map(normalize_biz_no)
        )

    matched = df_customer[df_customer['_사업자번호_정규화'] == key]
    if matched.empty:
        logger.warning(
            f"{CUSTOMER_DOMESTIC_SHEET}에 사업자번호 미등록: {biz_no} ({fallback_name})"
        )
        return None

    row = matched.iloc[0]
    customer_name = str(row[name_col]) if name_col in matched.columns else ''
    if not customer_name or customer_name.lower() == 'nan':
        customer_name = fallback_name

    customer_name_en = ''
    if name_en_col is not None:
        raw_en = row[name_en_col]
        if raw_en is not None and not pd.isna(raw_en):
            customer_name_en = str(raw_en).strip()

    to = split_emails(row[email_col])
    if not to:
        logger.warning(
            f"{CUSTOMER_DOMESTIC_SHEET}에 메일 주소 미입력: {biz_no} ({customer_name})"
        )
        return None

    # 고정 CC + 거래처별 참조메일 (중복 제거, To에 이미 있는 주소는 제외)
    cc_values: list[str] = list(TS_MAIL_CC)
    if cc_col is not None:
        cc_values.extend(split_emails(row[cc_col]))

    seen_lower = {addr.lower() for addr in to}
    cc: list[str] = []
    for addr in cc_values:
        addr = addr.strip()
        if addr and addr.lower() not in seen_lower:
            seen_lower.add(addr.lower())
            cc.append(addr)

    return Recipient(
        biz_no=str(biz_no),
        customer_name=customer_name,
        to=to,
        cc=tuple(cc),
        customer_name_en=customer_name_en,
    )


def find_recipient_for_order(
    order_data: pd.Series,
    df_customer: pd.DataFrame | None = None,
) -> Recipient | None:
    """주문(DN) 데이터에서 사업자번호를 뽑아 수신자 조회

    Args:
        order_data: DN 행 (Business registration number 포함)
        df_customer: Customer_국내 DataFrame (없으면 로드)

    Returns:
        Recipient 또는 None
    """
    biz_no = get_value(order_data, 'biz_no', '')
    customer_name = str(get_value(order_data, 'customer_name', ''))
    return find_recipient(biz_no, df_customer, fallback_name=customer_name)


# === PDF 변환 ===

def export_pdf(xlsx_path: Path, pdf_path: Path | None = None) -> Path:
    """거래명세표 xlsx를 PDF로 변환

    xlwings COM은 한글 경로에서 실패할 수 있으므로 임시 폴더에서 변환한 뒤
    최종 경로로 옮깁니다 (ts_generator와 동일한 패턴).

    Args:
        xlsx_path: 원본 xlsx 경로
        pdf_path: 출력 PDF 경로 (기본: xlsx와 같은 위치, 확장자만 .pdf)

    Returns:
        생성된 PDF 경로

    Raises:
        FileNotFoundError: 원본 파일이 없는 경우
    """
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"변환할 파일이 없습니다: {xlsx_path}")

    if pdf_path is None:
        pdf_path = xlsx_path.with_suffix('.pdf')
    pdf_path = Path(pdf_path)

    stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
    temp_dir = Path(tempfile.gettempdir())
    temp_xlsx = temp_dir / f"ts_pdf_{stamp}.xlsx"
    temp_pdf = temp_dir / f"ts_pdf_{stamp}.pdf"

    shutil.copy2(xlsx_path, temp_xlsx)
    try:
        with xlwings_app_context() as app:
            wb = app.books.open(str(temp_xlsx))
            try:
                wb.to_pdf(str(temp_pdf))
            finally:
                wb.close()
    finally:
        cleanup_temp_file(temp_xlsx)

    if not temp_pdf.exists():
        raise RuntimeError(f"PDF 변환 실패: {xlsx_path.name}")

    pdf_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(temp_pdf), str(pdf_path))
    logger.info(f"PDF 변환 완료: {pdf_path}")
    return pdf_path


def build_attachments(
    xlsx_path: Path,
    attach_format: str = TS_MAIL_ATTACH_FORMAT,
) -> tuple[Path, ...]:
    """첨부 파일 목록 준비 (필요 시 PDF 변환)

    Args:
        xlsx_path: 생성된 거래명세표 xlsx 경로
        attach_format: 'pdf' | 'xlsx' | 'both'

    Returns:
        첨부할 파일 경로 튜플
    """
    fmt = (attach_format or 'pdf').strip().lower()
    xlsx_path = Path(xlsx_path)

    if fmt == 'xlsx':
        return (xlsx_path,)
    if fmt == 'both':
        return (export_pdf(xlsx_path), xlsx_path)
    return (export_pdf(xlsx_path),)


# === Outlook 메일 생성 ===

def _get_outlook():
    """Outlook Application COM 객체 획득 (지연 import)

    Raises:
        MailConfigError: pywin32 미설치 또는 Outlook 사용 불가
    """
    try:
        import win32com.client
    except ImportError as e:
        raise MailConfigError(
            "pywin32가 설치되어 있지 않습니다. `pip install pywin32` 후 다시 시도하세요."
        ) from e

    try:
        return win32com.client.Dispatch('Outlook.Application')
    except Exception as e:
        raise MailConfigError(
            f"Outlook을 실행할 수 없습니다. Outlook이 설치·로그인되어 있는지 확인하세요. ({e})"
        ) from e


# COM 사용 가능 여부 판정 결과 (프로세스당 1회만 조사)
_com_available: bool | None = None


def outlook_com_available() -> bool:
    """Outlook COM을 쓸 수 있는지 (결과 캐시)

    새 Outlook(olk.exe)은 COM 자동화를 지원하지 않아 여기서 False가 됩니다.
    한 번 실패하면 같은 실행 안에서 다시 시도하지 않습니다.

    Returns:
        사용 가능 여부
    """
    global _com_available
    if _com_available is None:
        try:
            _get_outlook()
            _com_available = True
        except MailConfigError as e:
            logger.info(f"Outlook COM 사용 불가 — .eml 방식으로 전환합니다: {e}")
            _com_available = False
    return _com_available


def resolve_backend(configured: str | None = None) -> MailBackend:
    """설정값 + 환경으로 메일 작성 방식 결정

    설정값은 호출 시점에 읽습니다 (기본 인자로 박아두면 import 시점에 고정돼
    테스트나 런타임 변경이 반영되지 않음).

    Args:
        configured: 'auto' | 'outlook' | 'eml' (None이면 TS_MAIL_BACKEND 사용)

    Returns:
        MailBackend
    """
    choice = (configured or TS_MAIL_BACKEND or 'auto').strip().lower()
    if choice == 'outlook':
        return MailBackend.OUTLOOK
    if choice == 'eml':
        return MailBackend.EML
    return MailBackend.OUTLOOK if outlook_com_available() else MailBackend.EML


def _body_to_html(body: str) -> str:
    """평문 본문을 최소 HTML로 변환

    서식을 입히지 않고 줄바꿈만 유지합니다. 거래처명에 `&`나 `<`가 있어도
    깨지지 않도록 이스케이프를 먼저 합니다 (예: 'S&T중공업').

    Args:
        body: 평문 본문

    Returns:
        HTML 문자열
    """
    escaped = html_escape(body)
    return '<html><body>' + escaped.replace('\n', '<br>\n') + '</body></html>'


def build_eml(
    recipient: Recipient,
    subject: str,
    body: str,
    attachments: tuple[Path, ...],
    output_dir: Path | None = None,
) -> Path:
    """발송 대기 상태(.eml) 초안 파일 생성

    `X-Unsent: 1` 헤더가 핵심입니다. 이게 있어야 Outlook이 '받은 메일'이 아니라
    [보내기] 버튼이 있는 **작성 창**으로 엽니다. Date 헤더는 넣지 않습니다
    (넣으면 수신 메일로 취급될 수 있음).

    Args:
        recipient: 수신자 정보
        subject: 제목
        body: 본문 (평문)
        attachments: 첨부 파일 경로
        output_dir: 저장 폴더 (기본: 임시 폴더)

    Returns:
        생성된 .eml 경로
    """
    msg = EmailMessage()
    msg['To'] = ', '.join(recipient.to)
    if recipient.cc:
        msg['Cc'] = ', '.join(recipient.cc)
    msg['Subject'] = subject
    msg['X-Unsent'] = '1'
    msg.set_content(body)
    # HTML 대체본을 함께 넣는다. 평문만 보내면 Outlook이 서명을 본문 '위'에 끼워넣어
    # 서명 → 인사말 순서가 되지만, HTML이면 본문 '아래'에 정상적으로 붙는다.
    msg.add_alternative(_body_to_html(body), subtype='html')

    for path in attachments:
        path = Path(path)
        subtype = 'pdf' if path.suffix.lower() == '.pdf' else 'octet-stream'
        msg.add_attachment(
            path.read_bytes(),
            maintype='application',
            subtype=subtype,
            filename=path.name,
        )

    target_dir = Path(output_dir) if output_dir else Path(tempfile.gettempdir())
    target_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d%H%M%S%f')
    eml_path = target_dir / f"ts_draft_{stamp}.eml"
    eml_path.write_bytes(msg.as_bytes())
    logger.info(f".eml 초안 생성: {eml_path}")
    return eml_path


def open_eml(eml_path: Path) -> None:
    """.eml 파일을 기본 메일 앱으로 열기

    Args:
        eml_path: .eml 경로

    Raises:
        MailConfigError: 연결된 앱이 없거나 열기 실패
    """
    try:
        os.startfile(str(Path(eml_path).resolve()))
    except OSError as e:
        raise MailConfigError(
            f".eml 초안을 열 수 없습니다. .eml에 연결된 메일 앱이 있는지 확인하세요. ({e})"
        ) from e


def render_template(
    template: str,
    recipient: Recipient,
    doc_id: str,
    date_str: str,
    customer_po: str = NO_VALUE,
) -> str:
    """제목/본문 템플릿 치환

    Args:
        template: 치환자를 포함한 템플릿 문자열
        recipient: 수신자 정보
        doc_id: DN_ID 등 문서 ID
        date_str: 표시용 날짜
        customer_po: 거래처 발주번호 (여러 건이면 쉼표 구분, 없으면 'N/A')

    Returns:
        치환된 문자열 (알 수 없는 치환자가 있으면 원본 유지)
    """
    values = {
        'customer': recipient.customer_name,
        'customer_en': recipient.display_name_en,
        'customer_po': customer_po or NO_VALUE,
        'doc_id': doc_id,
        'date': date_str,
        'supplier': SUPPLIER_INFO.name,
    }
    try:
        return template.format(**values)
    except (KeyError, IndexError) as e:
        logger.warning(f"메일 템플릿 치환자 오류({e}) — 원본 문자열 사용")
        return template


def create_ts_mail(
    xlsx_path: Path,
    recipient: Recipient,
    doc_id: str,
    date_str: str | None = None,
    send: bool = False,
    attach_format: str = TS_MAIL_ATTACH_FORMAT,
    backend: MailBackend | None = None,
    customer_po: str = NO_VALUE,
) -> MailResult:
    """거래명세표 첨부 메일 생성 (기본: 초안 열기)

    Args:
        xlsx_path: 생성된 거래명세표 경로
        recipient: 수신자 정보
        doc_id: DN_ID 등 문서 ID
        date_str: 제목/본문에 쓸 날짜 (기본: 오늘)
        send: True면 즉시 발송 (Outlook COM에서만 가능)
        attach_format: 'pdf' | 'xlsx' | 'both'
        backend: 작성 방식 (기본: 설정+환경으로 자동 판정)
        customer_po: 거래처 발주번호 (여러 건이면 쉼표 구분, 없으면 'N/A')

    Returns:
        MailResult
    """
    if date_str is None:
        date_str = datetime.now().strftime('%Y-%m-%d')
    if backend is None:
        backend = resolve_backend()

    try:
        attachments = build_attachments(xlsx_path, attach_format)
    except Exception as e:
        logger.exception("첨부 파일 준비 실패")
        return MailResult(
            success=False, sent=False, recipient=recipient,
            message=f"첨부 파일 준비 실패: {e}", backend=backend,
        )

    subject = render_template(TS_MAIL_SUBJECT, recipient, doc_id, date_str, customer_po)
    body = render_template(TS_MAIL_BODY, recipient, doc_id, date_str, customer_po)

    if backend is MailBackend.EML:
        return _create_via_eml(recipient, subject, body, attachments, doc_id, send)
    return _create_via_outlook(recipient, subject, body, attachments, doc_id, send)


def _create_via_outlook(
    recipient: Recipient,
    subject: str,
    body: str,
    attachments: tuple[Path, ...],
    doc_id: str,
    send: bool,
) -> MailResult:
    """Outlook COM으로 메일 작성 (classic Outlook 전용)"""
    try:
        outlook = _get_outlook()
        mail = outlook.CreateItem(OL_MAIL_ITEM)
        mail.To = recipient.to_line
        if recipient.cc:
            mail.CC = recipient.cc_line
        mail.Subject = subject
        mail.Body = body

        for path in attachments:
            mail.Attachments.Add(str(Path(path).resolve()))

        if send:
            mail.Send()
            logger.info(f"거래명세표 메일 발송: {doc_id} -> {recipient.to_line}")
        else:
            mail.Display()
            logger.info(f"거래명세표 메일 초안 생성: {doc_id} -> {recipient.to_line}")

    except MailConfigError:
        raise
    except Exception as e:
        logger.exception("Outlook 메일 생성 실패")
        return MailResult(
            success=False, sent=False, recipient=recipient,
            attachments=attachments, message=f"Outlook 메일 생성 실패: {e}",
            backend=MailBackend.OUTLOOK,
        )

    return MailResult(
        success=True, sent=send, recipient=recipient, attachments=attachments,
        message='발송 완료' if send else '초안 생성 완료',
        backend=MailBackend.OUTLOOK,
    )


def _create_via_eml(
    recipient: Recipient,
    subject: str,
    body: str,
    attachments: tuple[Path, ...],
    doc_id: str,
    send: bool,
) -> MailResult:
    """.eml 초안을 만들어 기본 메일 앱으로 열기 (새 Outlook 호환)

    .eml은 '작성 창을 띄우는' 방식이라 자동 발송이 불가능합니다.
    send=True로 들어와도 초안까지만 하고 그 사실을 결과에 남깁니다.
    """
    note = ''
    if send:
        note = ' (.eml 방식은 자동 발송을 지원하지 않아 초안까지만 진행)'
        logger.warning(".eml 방식에서는 즉시 발송이 불가능합니다 — 초안으로 대체")

    try:
        eml_path = build_eml(recipient, subject, body, attachments)
        open_eml(eml_path)
    except MailConfigError:
        raise
    except Exception as e:
        logger.exception(".eml 초안 생성 실패")
        return MailResult(
            success=False, sent=False, recipient=recipient,
            attachments=attachments, message=f".eml 초안 생성 실패: {e}",
            backend=MailBackend.EML,
        )

    logger.info(f"거래명세표 .eml 초안 열기: {doc_id} -> {recipient.to_line}")
    return MailResult(
        success=True, sent=False, recipient=recipient, attachments=attachments,
        message=f'초안 생성 완료{note}', backend=MailBackend.EML,
    )
