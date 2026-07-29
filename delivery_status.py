#!/usr/bin/env python
"""
거래처 납기현황 조회 (Delivery Status)
=======================================

거래처가 "언제 나오냐"고 물을 때 회신할 미출고 현황표를 만듭니다.
조회 기준은 **사업자등록번호**(Business registration number), 대상은 국내(`SO_국내`)입니다.

출고 여부 판정
--------------
시트의 `Status` 컬럼을 그대로 읽지 않고 `DN_국내` 출고수량으로 **직접 계산**합니다.

`SO_국내.Status`는 수기 입력이 아니라 파워쿼리 결과를 끌어온 캐시값이라서
(`XLOOKUP(SO_ID&Line item, SO_통합[...], SO_통합[출고완료])`), 새로고침 전에 파일을 읽으면
옛 판정이 그대로 나옵니다. 실제로 이미 납품한 라인이 '미출고'로 남아 있는 경우가 있고,
그대로 회신하면 **고객에게 "아직 안 나갔습니다"라고 잘못 알리게** 됩니다.

`SO_통합[출고완료]`의 정의 자체가 DN 출고수량 기반이므로, 여기서 같은 식을 계산하면
결과는 같으면서 새로고침 여부에 의존하지 않습니다 (`DN_국내`는 직접 입력 시트라 캐시가 없음).

    출고수량 없음            → 미출고
    주문수량 - 출고수량 > 0   → 부분 출고
    출고일 없음              → 공장 출고
    그 외                    → 출고 완료   ← 회신 대상에서 제외

캐시 `Status`와 계산 결과가 어긋나면 실행 시 새로고침 안내를 띄웁니다. 회신 문서는 어차피
맞게 나가지만, 대시보드·피벗 등 같은 파일을 보는 다른 산출물도 낡았다는 신호이기 때문입니다.

분할 납기
---------
한 주문 안에서 `EXW NOAH`가 갈리면 **날짜별로 행을 나눕니다.** 실제로 흔합니다
(`SOD-2026-0264`는 32라인이 7개 날짜에 걸쳐 2027년까지). 나뉜 주문은 같은 Customer PO가
여러 줄로 보이므로, 어느 품목이 어느 날짜인지 비고에 품목명을 덧붙입니다.

메일 회신
---------
생성 후 수신자를 보여주고 "이메일을 발송하시겠습니까?"를 묻습니다. 수신자는 조회 기준인
사업자번호로 `Customer_국내`에서 찾습니다(거래명세표와 같은 경로). 본문에는 납기 표를
그대로 싣고 xlsx를 첨부합니다 — 고객이 첨부를 열지 않아도 답을 볼 수 있어야 하고,
원래 손으로 회신하던 형식도 표를 본문에 붙이는 방식이었습니다.

사용법:
    python delivery_status.py 615-81-88675        # 사업자번호 (하이픈 유무 무관)
    python delivery_status.py 엔이에스             # 거래처명 부분일치
    python delivery_status.py --list              # 미출고가 있는 거래처 목록
    python delivery_status.py 615-81-88675 --all  # 출고완료 포함 전체
    python delivery_status.py 615-81-88675 --mail # 확인 없이 메일 초안 열기
    python delivery_status.py 615-81-88675 --no-mail  # 묻지 않고 문서만 (배치용)
    python delivery_status.py 615-81-88675 -v     # 상세 로그
"""

from __future__ import annotations

import argparse
import datetime as dt
import logging
import sys
import unicodedata
import warnings
from html import escape as html_escape
from pathlib import Path

import pandas as pd

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

from po_generator.cli_common import generate_output_filename
from po_generator.config import (
    CUSTOMER_DOMESTIC_SHEET,
    DS_MAIL_ATTACH_FORMAT,
    DS_MAIL_BODY,
    DS_MAIL_CC,
    DS_MAIL_SUBJECT,
    DS_OUTPUT_DIR,
    DN_DOMESTIC_SHEET,
    NOAH_SO_PO_DN_FILE,
    SO_DOMESTIC_SHEET,
)
from po_generator.logging_config import setup_logging
from po_generator.mail_cli import (
    MailOptions,
    add_mail_arguments,
    confirm,
    resolve_mail_mode,
)
from po_generator.mailer import (
    MailConfigError,
    create_document_mail,
    find_recipient,
    render_template,
)
from po_generator.utils import normalize_biz_no, normalize_line_item

logger = logging.getLogger(__name__)


# 회신 대상 출고상태 — '출고 완료'만 빠진다
PENDING_STATUSES: tuple[str, ...] = ('미출고', '부분 출고', '공장 출고')

# 회신에서 통째로 빼는 주문 (캐시 여부와 무관하게 시트 Status로만 판단)
# 취소/보류는 DN이 영영 안 생기므로 수량으로는 구분할 수 없다.
EXCLUDED_SO_STATUSES: frozenset[str] = frozenset({'Cancelled', 'Hold'})

# EXW NOAH가 비어 있을 때 고객 회신에 찍을 문구
DATE_TBD_LABEL = '처리중'

# 한 주문의 납기가 갈려 행이 나뉠 때, 비고에 덧붙일 품목명의 최대 길이 (넘으면 '외 N종')
# 개수가 아니라 길이로 자른다 — 품목명이 'MA02/0.75kW/43RPM/ON-OFF/SBWG-04-1SM'처럼
# 긴 거래처(피엠에스)가 있어서 "3개까지"로는 회신 표가 깨진다.
SPLIT_ITEM_NOTE_MAX_CHARS = 60

# 사업자등록번호 자릿수 — 조회어가 번호인지 이름인지 가르는 기준
BIZ_NO_MIN_DIGITS = 8

# 공장 출고일 컬럼명 — **'예정일'이라고 쓴다.**
# `EXW NOAH`는 확정 약속이 아니라 계획이다. 고객에게 '출고일'로 나가면 그날 안 나갔을 때
# 약속을 어긴 것이 되고, 실제로 요청납기 초과가 164행 중 23행이다.
COL_EXW: str = 'NOAH 공장 출고 예정일'

# 납기현황 시트(요약) 컬럼 — 요청납기와 출고 예정일을 나란히 둬서 한눈에 대조되게 한다
SUMMARY_COLUMNS: tuple[str, ...] = (
    'Customer PO', 'Remarks', '수량',
    'Requested delivery date', COL_EXW,
    'Sales 금액', 'PO receipt date',
)

# 메일 본문 표에 넣을 컬럼 (PO receipt date는 첨부 xlsx에만 — 고객이 이미 아는 날짜다)
MAIL_TABLE_COLUMNS: tuple[str, ...] = tuple(
    c for c in SUMMARY_COLUMNS if c != 'PO receipt date'
)

# 오른쪽 정렬할 숫자 컬럼 / 가운데 정렬할 날짜 컬럼
_MAIL_TABLE_NUMERIC: frozenset[str] = frozenset({'수량', 'Sales 금액'})
_MAIL_TABLE_CENTER: frozenset[str] = frozenset({'Requested delivery date', COL_EXW})

# 상세 시트 컬럼 — 날짜는 접수 → 요청납기 → 공장출고 → 예상납기 순으로 읽히게 둔다
DETAIL_COLUMNS: tuple[str, ...] = (
    'SO_ID', 'Line item', 'Customer PO', 'Remarks', 'Item name',
    '주문수량', '출고수량', '미출고수량', 'Sales Unit Price', '미출고금액',
    'PO receipt date', 'Requested delivery date', 'EXW NOAH',
    'Expected delivery date', '출고상태',
)

# 본문 템플릿의 {table}을 HTML 표로 바꿔치기할 때 쓰는 자리표시자.
# 평문을 HTML로 이스케이프한 뒤에 넣어야 표 태그가 살아남으므로, 이스케이프에도
# 형태가 변하지 않는 토큰이어야 한다 (영숫자 + 퍼센트).
_HTML_TABLE_TOKEN = '%%DELIVERY_STATUS_TABLE%%'


# === 값 정규화 =============================================================

def as_date(value) -> pd.Timestamp | None:
    """날짜 셀 → Timestamp (빈 값이면 None)

    `EXW NOAH`는 빈 칸이 NaN이 아니라 `datetime.time(0, 0)`으로 읽힙니다
    (셀에 0 값이 들어 있어 컬럼 dtype이 object로 떨어진다).
    NaN만 검사하면 미정 납기가 1900-01-00 같은 날짜로 둔갑하므로 타입까지 본다.
    """
    if value is None or value is pd.NaT:
        return None
    if isinstance(value, dt.time):  # 빈 칸이 시간 0으로 읽힌 경우
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if isinstance(value, str) and not value.strip():
        return None
    try:
        ts = pd.Timestamp(value)
    except (ValueError, TypeError):
        return None
    if pd.isna(ts):
        return None
    return ts


def format_date(value) -> str:
    """날짜 → 'YYYY-MM-DD' (빈 값이면 '')"""
    ts = as_date(value)
    return ts.strftime('%Y-%m-%d') if ts is not None else ''




def _clean_text(value) -> str:
    """셀 값 → 표시용 문자열 (NaN/None은 빈 문자열)"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    text = str(value).strip()
    return '' if text.lower() == 'nan' else text


def _display_width(text: str) -> int:
    """콘솔 표시 폭 — 한글/한자는 2칸을 차지한다"""
    return sum(2 if unicodedata.east_asian_width(ch) in 'WF' else 1 for ch in text)


def _pad(text: str, width: int) -> str:
    """표시 폭 기준 왼쪽 정렬 패딩 (넘치면 '…'로 자른다)

    f-string의 `:<32`는 글자 수로 세기 때문에 거래처명·Remarks가 한글이면 열이 어긋난다.
    """
    if _display_width(text) > width:
        out = ''
        for ch in text:
            if _display_width(out) + _display_width(ch) > width - 1:
                break
            out += ch
        text = out + '…'
    return text + ' ' * max(0, width - _display_width(text))


# === 데이터 로드 ===========================================================

def load_sheets() -> tuple[pd.DataFrame, pd.DataFrame]:
    """SO_국내 / DN_국내 로드

    Returns:
        (SO DataFrame, DN DataFrame)

    Raises:
        FileNotFoundError: 데이터 파일이 없는 경우
    """
    if not NOAH_SO_PO_DN_FILE.exists():
        raise FileNotFoundError(f"소스 파일을 찾을 수 없습니다: {NOAH_SO_PO_DN_FILE}")

    logger.info("데이터 로드: %s", NOAH_SO_PO_DN_FILE.name)
    with pd.ExcelFile(NOAH_SO_PO_DN_FILE) as xl:
        so = pd.read_excel(xl, sheet_name=SO_DOMESTIC_SHEET)
        dn = pd.read_excel(xl, sheet_name=DN_DOMESTIC_SHEET)

    so = so[so['SO_ID'].notna()].copy()
    dn = dn[dn['DN_ID'].notna()].copy()
    logger.info("SO_국내 %d행 / DN_국내 %d행", len(so), len(dn))
    return so, dn


def aggregate_dn(dn: pd.DataFrame) -> pd.DataFrame:
    """DN_국내를 SO_ID + Line item 단위 출고수량/최종출고일로 집계

    한 SO 라인을 여러 번에 나눠 출고(분할 납품)하면 DN 행이 여러 개이므로 합산한다.
    """
    columns = ['_so_id', '_line', '출고수량', '_last_ship', '_dn_match']
    if dn.empty:
        # dtype을 명시한다 — 빈 프레임을 columns만으로 만들면 전부 object가 되고,
        # merge 뒤 fillna에서 pandas 다운캐스팅 경고가 난다.
        return pd.DataFrame({
            '_so_id': pd.Series(dtype='object'),
            '_line': pd.Series(dtype='object'),
            '출고수량': pd.Series(dtype='float64'),
            '_last_ship': pd.Series(dtype='datetime64[ns]'),
            '_dn_match': pd.Series(dtype='object'),
        })

    work = pd.DataFrame({
        '_so_id': dn['SO_ID'].map(_clean_text),
        '_line': dn['Line item'].map(normalize_line_item),
        '출고수량': pd.to_numeric(dn['Qty'], errors='coerce').fillna(0.0),
        # as_date로 한 번 거른 뒤 datetime64로 — 빈 칸이 time(0,0)으로 읽히는 시트가 섞여도
        # groupby('max')가 object dtype에서 깨지지 않게 한다.
        '_ship_date': pd.to_datetime(dn['출고일'].map(as_date), errors='coerce'),
    })
    grouped = work.groupby(['_so_id', '_line'], as_index=False).agg(
        출고수량=('출고수량', 'sum'),
        _last_ship=('_ship_date', 'max'),
    )
    # DN 행의 '존재' 자체를 표시 — 파워쿼리의 `[출고수량] = null` 판정과 맞추기 위해서다.
    # 수량 0짜리 DN 행이 있으면 '미출고'가 아니라 '부분 출고'가 맞다.
    grouped['_dn_match'] = True
    return grouped[columns]


def attach_ship_status(so: pd.DataFrame, dn: pd.DataFrame) -> pd.DataFrame:
    """SO 라인에 출고수량·미출고수량·출고상태를 붙인다

    출고상태 판정은 파워쿼리 `SO_통합[출고완료]`와 같은 식이다 (모듈 docstring 참조).
    """
    dn_agg = aggregate_dn(dn)

    work = so.copy()
    work['_so_id'] = work['SO_ID'].map(_clean_text)
    work['_line'] = work['Line item'].map(normalize_line_item)
    work['_biz'] = work['Business registration number'].map(normalize_biz_no)
    work['_주문수량'] = pd.to_numeric(work['Item qty'], errors='coerce').fillna(0.0)
    work['_단가'] = pd.to_numeric(work['Sales Unit Price'], errors='coerce').fillna(0.0)

    # SO 라인이 (SO_ID, Line item)으로 유일하지 않으면 merge가 fan-out되어 수량이 부풀려진다.
    dup = int(work.duplicated(subset=['_so_id', '_line']).sum())
    if dup:
        logger.warning(
            "%s: (SO_ID, Line item) 중복 %d건 — 출고수량이 중복 집계될 수 있습니다",
            SO_DOMESTIC_SHEET, dup,
        )

    work = work.merge(dn_agg, on=['_so_id', '_line'], how='left')
    work['출고수량'] = pd.to_numeric(work['출고수량'], errors='coerce').fillna(0.0)
    # merge 결과의 _dn_match는 True 아니면 NaN이라 notna()가 곧 '매칭됨'이다
    # (object 컬럼에 fillna(False)를 쓰면 pandas 다운캐스팅 경고가 난다)
    work['_has_dn'] = work['_dn_match'].notna()
    work['_미출고수량'] = (work['_주문수량'] - work['출고수량']).clip(lower=0.0)

    def _status(row) -> str:
        if not row['_has_dn']:
            return '미출고'
        if row['_주문수량'] - row['출고수량'] > 0.001:
            return '부분 출고'
        if pd.isna(row['_last_ship']):
            return '공장 출고'
        return '출고 완료'

    # 빈 프레임에 apply(axis=1)을 걸면 Series가 아니라 DataFrame이 돌아와 대입이 깨진다
    work['_출고상태'] = (
        work.apply(_status, axis=1) if len(work)
        else pd.Series(dtype='object', index=work.index)
    )
    work['_미출고금액'] = work['_단가'] * work['_미출고수량']
    return work


def cached_status(df: pd.DataFrame) -> pd.Series:
    """시트의 `Status` 컬럼을 문자열 Series로 (컬럼이 없으면 전부 빈 문자열)

    Status는 파워쿼리 결과의 캐시라 언제든 통째로 빠질 수 있다 (수식 오류, 시트 개편).
    없다고 조회가 멈추면 안 되므로 '판정 정보 없음'으로 취급한다.
    """
    if 'Status' not in df.columns:
        return pd.Series('', index=df.index, dtype=object)
    return df['Status'].map(_clean_text)


def count_stale_status(work: pd.DataFrame) -> int:
    """시트의 캐시 `Status`와 계산된 출고상태가 어긋나는 행 수

    파워쿼리 새로고침이 밀렸는지 알려주는 신호. 취소/보류 주문은 애초에 판정 대상이
    아니므로 제외하고, 수식이 빈 문자열을 반환한 행(NaN)도 '아직 계산 안 됨'이라 제외한다.
    """
    cached = cached_status(work)
    comparable = (cached != '') & (~cached.isin(EXCLUDED_SO_STATUSES))
    return int((comparable & (cached != work['_출고상태'])).sum())


# === 거래처 조회 ===========================================================

def customer_directory(work: pd.DataFrame) -> pd.DataFrame:
    """사업자번호 → 거래처명 목록 (SO에 등장하는 거래처)

    같은 사업자번호에 표기가 여러 개면 가장 많이 쓰인 이름을 대표로 삼는다.
    """
    rows = work[work['_biz'] != ''].copy()
    rows['_name'] = rows['Customer name'].map(_clean_text)
    if rows.empty:
        return pd.DataFrame(columns=['_biz', '_name'])
    top = (
        rows.groupby(['_biz', '_name'], as_index=False)
            .size()
            .sort_values(['_biz', 'size'], ascending=[True, False])
            .drop_duplicates(subset='_biz', keep='first')
    )
    return top[['_biz', '_name']].reset_index(drop=True)


def resolve_customer(work: pd.DataFrame, query: str) -> tuple[list[tuple[str, str]], str]:
    """조회어(사업자번호 또는 거래처명) → 후보 [(사업자번호, 거래처명)]

    Returns:
        (후보 목록, 조회 방식 설명)
    """
    directory = customer_directory(work)
    digits = normalize_biz_no(query)

    if len(digits) >= BIZ_NO_MIN_DIGITS:
        hits = directory[directory['_biz'] == digits]
        return list(hits.itertuples(index=False, name=None)), '사업자번호로'

    keyword = query.strip().lower()
    hits = directory[directory['_name'].str.lower().str.contains(keyword, regex=False, na=False)]
    return list(hits.itertuples(index=False, name=None)), '거래처명으로'


def format_biz_no(digits: str) -> str:
    """정규화된 사업자번호 → 'NNN-NN-NNNNN' 표기 (10자리가 아니면 원문 유지)"""
    if len(digits) == 10:
        return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"
    return digits


# === 집계 =================================================================

def select_rows(work: pd.DataFrame, biz_no: str, include_shipped: bool) -> pd.DataFrame:
    """회신 대상 SO 라인 선별"""
    rows = work[work['_biz'] == biz_no].copy()
    rows = rows[~cached_status(rows).isin(EXCLUDED_SO_STATUSES)]

    if not include_shipped:
        rows = rows[rows['_출고상태'].isin(PENDING_STATUSES)]
    return rows


def item_note(group: pd.DataFrame) -> str:
    """행에 묶인 품목명 — 비고에 붙여 "어느 품목이 이 날짜인지" 알려준다

    품목이 많으면 회신 표가 읽히지 않으므로 길이가 찰 때까지만 쓰고 나머지는 종수로 접는다.
    첫 품목은 아무리 길어도 반드시 남긴다 (전부 '외 N종'만 나오면 단서가 없다).
    """
    names: list[str] = []
    for name in group['Item name'].map(_clean_text):
        if name and name not in names:
            names.append(name)
    if not names:
        return ''

    shown: list[str] = []
    used = 0
    for name in names:
        if shown and used + len(name) + 2 > SPLIT_ITEM_NOTE_MAX_CHARS:
            break
        shown.append(name)
        used += len(name) + 2

    text = ', '.join(shown)
    if len(names) > len(shown):
        text += f" 외 {len(names) - len(shown)}종"
    return f"품목: {text}"


def build_summary(rows: pd.DataFrame) -> pd.DataFrame:
    """주문 × 납기 단위 요약 — 고객에게 보내는 표

    기본 단위는 이미 메일로 회신하던 표와 같은 주문(SO_ID) 1행이고, 여기에 Sales 금액과
    PO receipt date를 더했다.

    다만 **한 주문 안에서 납기가 갈리면 날짜별로 행을 나눈다.** 기준은 두 날짜 모두다 —
    고객 요청납기(`Requested delivery date`)와 공장 출고일(`EXW NOAH`).
    분할 납기가 실제로 있다 (세진밸브 `SOD-2026-0264`는 32라인이 7개 출고일에 걸쳐 2027년까지).
    대표 날짜 하나로 뭉개면 나머지 납기가 통째로 사라지거나 "그날 다 나온다"는 틀린 약속이 된다.
    납기가 하나뿐인 주문(대부분)은 예전 그대로 1행이다.

    나뉜 주문은 같은 Customer PO가 여러 줄로 보이므로 어느 행이 무엇인지 알려줘야 하는데,
    **비고만으로 구분이 안 될 때만** 품목명을 덧붙인다. 비고가 이미 호선별로 갈려 있으면
    (세진밸브 `H2734`/`H2735`…) 품목은 잡음이고, 피엠에스(두 행 모두 `묘도 GS`)·
    티에스엔텍(비고 없음)처럼 겹칠 때만 품목이 단서가 된다.
    """
    if rows.empty:
        return pd.DataFrame(columns=list(SUMMARY_COLUMNS))

    work = rows.copy()
    # 리스트를 그대로 컬럼에 넣으면 안 된다 — None이 섞이면 pandas가 datetime64로 캐스팅해
    # None이 NaT가 되고, `is not None`을 통과한 NaT가 strftime에서 터진다.
    # 날짜 계산은 컬럼에 담기 전에 끝내고, 컬럼에는 이미 만든 문자열/정렬키만 넣는다.
    exw_dates = [as_date(v) for v in work['EXW NOAH']]
    work['_exw_key'] = [
        d.strftime('%Y-%m-%d') if d is not None else DATE_TBD_LABEL for d in exw_dates
    ]
    work['_sort_exw'] = [d if d is not None else pd.Timestamp.max for d in exw_dates]
    work['_req_key'] = [format_date(v) for v in work['Requested delivery date']]

    records = []
    for _so_id, so_group in work.groupby('_so_id', sort=False):
        # 요청납기·공장출고일 **둘 중 하나라도** 다르면 다른 행이다
        date_groups = list(so_group.groupby(['_req_key', '_exw_key'], sort=False))
        # Remarks가 라인마다 갈리는 주문이 있다 — 행마다 첫 비어있지 않은 값을 쓴다
        base_remarks = [
            next((r for r in g['Remarks'].map(_clean_text) if r), '')
            for _, g in date_groups
        ]
        # 품목은 "비고만으로 어느 행인지 모를 때"만 붙인다.
        # 세진밸브(SOD-2026-0264)처럼 비고가 이미 호선별로 갈려 있으면 품목은 잡음일 뿐이고,
        # 피엠에스(두 행 모두 '묘도 GS')·티에스엔텍(비고 없음)처럼 겹치면 품목이 유일한 단서다.
        need_note = len(date_groups) > 1 and (
            len(set(base_remarks)) < len(base_remarks) or any(not r for r in base_remarks)
        )

        for ((req_key, exw_key), group), remarks in zip(date_groups, base_remarks):
            if need_note:
                note = item_note(group)
                if note:
                    remarks = f"{remarks} ({note})" if remarks else note
            first = group.iloc[0]
            records.append({
                'Customer PO': _clean_text(first['Customer PO']),
                'Remarks': remarks,
                '수량': float(group['_미출고수량'].sum()),
                'Requested delivery date': req_key,
                COL_EXW: exw_key,
                'Sales 금액': float(group['_미출고금액'].sum()),
                'PO receipt date': format_date(first['PO receipt date']),
                '_sort_exw': first['_sort_exw'],
            })

    summary = pd.DataFrame(records)
    # 납기 확정 건을 앞에, '처리중'을 뒤에 — 회신 받는 쪽이 먼저 보고 싶은 순서.
    # 같은 주문이 나뉜 행들은 요청납기 순으로 붙어 보이게 한다.
    summary = summary.sort_values(
        ['_sort_exw', 'PO receipt date', 'Customer PO', 'Requested delivery date'],
    ).reset_index(drop=True)
    return summary[list(SUMMARY_COLUMNS)]


def build_detail(rows: pd.DataFrame) -> pd.DataFrame:
    """SO 라인 단위 상세 — 내부 확인용"""
    columns = list(DETAIL_COLUMNS)
    if rows.empty:
        return pd.DataFrame(columns=columns)

    detail = pd.DataFrame({
        'SO_ID': rows['_so_id'],
        'Line item': rows['_line'],
        'Customer PO': rows['Customer PO'].map(_clean_text),
        'Remarks': rows['Remarks'].map(_clean_text),
        'Item name': rows['Item name'].map(_clean_text),
        '주문수량': rows['_주문수량'],
        '출고수량': rows['출고수량'],
        '미출고수량': rows['_미출고수량'],
        'Sales Unit Price': rows['_단가'],
        '미출고금액': rows['_미출고금액'],
        'PO receipt date': rows['PO receipt date'].map(format_date),
        'Requested delivery date': rows['Requested delivery date'].map(format_date),
        'EXW NOAH': rows['EXW NOAH'].map(format_date),
        'Expected delivery date': rows['Expected delivery date'].map(format_date),
        '출고상태': rows['_출고상태'],
    })
    detail['_line_sort'] = pd.to_numeric(detail['Line item'], errors='coerce')
    detail = detail.sort_values(['SO_ID', '_line_sort']).reset_index(drop=True)
    return detail.drop(columns=['_line_sort'])[columns]


def build_customer_list(work: pd.DataFrame) -> pd.DataFrame:
    """미출고 잔량이 있는 거래처 목록 (--list)"""
    pending = work[
        (~cached_status(work).isin(EXCLUDED_SO_STATUSES))
        & (work['_출고상태'].isin(PENDING_STATUSES))
        & (work['_biz'] != '')
    ].copy()
    if pending.empty:
        return pd.DataFrame(columns=['사업자번호', '거래처명', '주문건수', '미출고수량', '미출고금액'])

    directory = customer_directory(work).set_index('_biz')['_name'].to_dict()
    agg = pending.groupby('_biz').agg(
        주문건수=('_so_id', 'nunique'),
        미출고수량=('_미출고수량', 'sum'),
        미출고금액=('_미출고금액', 'sum'),
    ).reset_index()
    agg['거래처명'] = agg['_biz'].map(directory).fillna('')
    agg['사업자번호'] = agg['_biz'].map(format_biz_no)
    agg = agg.sort_values('미출고금액', ascending=False).reset_index(drop=True)
    return agg[['사업자번호', '거래처명', '주문건수', '미출고수량', '미출고금액']]


# === 출력 =================================================================

def write_output(
    summary: pd.DataFrame,
    detail: pd.DataFrame,
    output_file: Path,
    customer_name: str,
    biz_no: str,
) -> None:
    """납기현황 xlsx 출력 (납기현황 / 상세 2시트)"""
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    output_file.parent.mkdir(parents=True, exist_ok=True)

    def _add_table(writer, sheet_name: str, display_name: str) -> None:
        ws = writer.sheets[sheet_name]
        if ws.max_row < 2:
            return
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        tbl = Table(displayName=display_name, ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tbl)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='납기현황', index=False)
        _add_table(writer, '납기현황', '납기현황')

        detail.to_excel(writer, sheet_name='상세', index=False)
        _add_table(writer, '상세', '상세')

        # 열 너비·숫자 서식은 **컬럼명**으로 지정한다.
        # 컬럼 letter를 박아두면 컬럼이 하나 늘 때마다 서식이 조용히 한 칸씩 밀린다.
        widths = {
            'Customer PO': 22, 'Remarks': 52, '수량': 8,
            'Requested delivery date': 20, COL_EXW: 22,
            'Sales 금액': 16, 'PO receipt date': 16,
            'SO_ID': 16, 'Line item': 8, 'Item name': 40,
            '주문수량': 10, '출고수량': 10, '미출고수량': 11,
            'Sales Unit Price': 14, '미출고금액': 14,
            'EXW NOAH': 14, 'Expected delivery date': 20, '출고상태': 11,
        }
        money = {'Sales 금액', 'Sales Unit Price', '미출고금액'}

        for sheet, df in (('납기현황', summary), ('상세', detail)):
            ws = writer.sheets[sheet]
            for idx, name in enumerate(df.columns, start=1):
                letter = get_column_letter(idx)
                if name in widths:
                    ws.column_dimensions[letter].width = widths[name]
                if name in money:
                    for row in range(2, ws.max_row + 1):
                        ws[f"{letter}{row}"].number_format = '#,##0'
            for cell in ws[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center', vertical='center')

        # 거래처/기준일은 인쇄 머리글로 — 표 자체는 A1부터 시작해야 필터/복사가 깔끔하다
        head = writer.sheets['납기현황']
        head.oddHeader.left.text = f"{customer_name} ({format_biz_no(biz_no)})"
        head.oddHeader.right.text = f"기준일 {dt.date.today():%Y-%m-%d}"


def print_summary(summary: pd.DataFrame, customer_name: str, biz_no: str) -> None:
    """콘솔 요약 출력"""
    print()
    print(f"납기현황 — {customer_name} ({format_biz_no(biz_no)})")
    print("=" * 88)

    if summary.empty:
        print("  미출고 건이 없습니다.")
        return

    print(f"  {_pad('Customer PO', 22)} {_pad('Remarks', 32)} {'수량':>6} {_pad('출고 예정일', 13)} {'Sales 금액':>14}")
    print("  " + "-" * 86)
    for _, row in summary.iterrows():
        print(
            f"  {_pad(row['Customer PO'], 22)} {_pad(row['Remarks'], 32)} {row['수량']:>6,.0f} "
            f"{_pad(row[COL_EXW], 13)} {row['Sales 금액']:>14,.0f}"
        )
    print("  " + "-" * 86)
    print(
        f"  {_pad('합계', 22)} {_pad('', 32)} {summary['수량'].sum():>6,.0f} "
        f"{_pad('', 13)} {summary['Sales 금액'].sum():>14,.0f}"
    )
    tbd = int((summary[COL_EXW] == DATE_TBD_LABEL).sum())
    if tbd:
        print(f"\n  출고 예정일 미정({DATE_TBD_LABEL}): {tbd}건 — EXW NOAH가 비어 있습니다.")


def print_customer_list(listing: pd.DataFrame) -> None:
    """--list 출력"""
    print()
    print("미출고 잔량이 있는 거래처")
    print("=" * 76)
    if listing.empty:
        print("  없습니다.")
        return
    print(f"  {_pad('사업자번호', 16)} {_pad('거래처명', 28)} {'주문':>5} {'수량':>7} {'미출고금액':>15}")
    print("  " + "-" * 74)
    for _, row in listing.iterrows():
        print(
            f"  {_pad(row['사업자번호'], 16)} {_pad(row['거래처명'], 28)} {row['주문건수']:>5} "
            f"{row['미출고수량']:>7,.0f} {row['미출고금액']:>15,.0f}"
        )
    print("  " + "-" * 74)
    print(f"  거래처 {len(listing)}곳 / 미출고금액 합계 {listing['미출고금액'].sum():,.0f}")


# === 메일 본문 ============================================================

def is_late(row) -> bool:
    """공장 출고일이 고객 요청납기보다 늦은 행인가

    둘 다 'YYYY-MM-DD' 문자열이라 사전순 비교가 곧 날짜 비교다.
    출고일이 미정(`처리중`)이거나 요청납기가 비어 있으면 **판정하지 않는다** —
    모르는 것을 늦었다고 표시하면 안 된다.
    """
    exw = str(row.get(COL_EXW, '') or '')
    req = str(row.get('Requested delivery date', '') or '')
    if not req or not exw or exw == DATE_TBD_LABEL:
        return False
    return exw > req


def build_text_table(summary: pd.DataFrame) -> str:
    """평문 본문용 표

    HTML을 못 보는 메일 클라이언트가 받는 대체본이다. 데이터가 빠지면 안 되므로
    "첨부 참조"로 때우지 않고 같은 내용을 문자표로 그린다.
    요청납기 초과는 색을 못 쓰므로 `*` 표식과 각주로 대신한다 (HTML과 정보량을 맞춘다).
    """
    cols = [c for c in MAIL_TABLE_COLUMNS if c in summary.columns]
    if summary.empty or not cols:
        return ''

    late_flags = [is_late(row) for _, row in summary.iterrows()]

    def _text(row, col, late: bool) -> str:
        text = _cell_text(row[col])
        return f"{text} *" if late and col == COL_EXW else text

    widths = {
        c: max(
            [_display_width(c)]
            + [_display_width(_text(row, c, late))
               for (_, row), late in zip(summary.iterrows(), late_flags)]
        )
        for c in cols
    }
    lines = [
        '  '.join(_pad(c, widths[c]) for c in cols),
        '-' * (sum(widths.values()) + 2 * (len(cols) - 1)),
    ]
    for (_, row), late in zip(summary.iterrows(), late_flags):
        lines.append('  '.join(_pad(_text(row, c, late), widths[c]) for c in cols))

    if any(late_flags):
        lines.append('')
        lines.append('* 요청 납기일보다 공장 출고 예정일이 늦은 건')
    return '\n'.join(lines)


def build_html_table(summary: pd.DataFrame) -> str:
    """HTML 본문용 표

    메일 클라이언트는 <style> 블록을 자주 지우므로 인라인 스타일만 쓴다.
    거래처명·비고에 `&`나 `<`가 들어와도 깨지지 않게 값은 모두 이스케이프한다
    (예: 'S&T중공업', '<긴급>').
    """
    cols = [c for c in MAIL_TABLE_COLUMNS if c in summary.columns]
    cell = 'padding:6px 10px;border:1px solid #c9c9c9;'
    head = (
        '<tr>'
        + ''.join(
            f'<th style="{cell}background:#f2f2f2;text-align:center;">{html_escape(c)}</th>'
            for c in cols
        )
        + '</tr>'
    )

    rows = []
    for _, row in summary.iterrows():
        late = is_late(row)
        cells = []
        for c in cols:
            text = html_escape(_cell_text(row[c]))
            if c in _MAIL_TABLE_NUMERIC:
                align = 'right'
            elif c in _MAIL_TABLE_CENTER:
                align = 'center'
            else:
                align = 'left'
            # 출고일 미정은 눈에 띄어야 한다 — 고객이 가장 먼저 물어보는 줄이다.
            # 요청납기를 넘긴 출고일은 굵게까지 — 색만으로 구분하면 색각 이상에서 안 보인다.
            if row[c] == DATE_TBD_LABEL:
                emphasis = 'color:#c00000;'
            elif late and c == COL_EXW:
                emphasis = 'color:#c00000;font-weight:bold;'
            else:
                emphasis = ''
            cells.append(f'<td style="{cell}text-align:{align};{emphasis}">{text}</td>')
        rows.append('<tr>' + ''.join(cells) + '</tr>')

    return (
        '<table style="border-collapse:collapse;font-family:맑은 고딕,sans-serif;'
        'font-size:13px;">'
        + head + ''.join(rows) +
        '</table>'
    )


def build_html_body(body_text: str, html_table: str) -> str:
    """본문 평문 → HTML

    **Outlook은 본문의 첫 블록 요소 바로 뒤에 자동 서명을 끼워 넣는다** (2026-07-29 실측 2회).
    처음엔 `<br>`로 이은 인라인 텍스트였고 서명이 표(첫 블록) 앞에 들어갔다. 문단을 `<p>`로
    바꿨더니 이번엔 첫 `<p>` 뒤에 들어갔다. `<div>`로 감싸도 Outlook은 그 안으로 파고든다.

    그래서 **본문 전체를 표 한 칸(`<table><tr><td>`) 안에 넣어 최상위 블록을 하나로 만든다.**
    삽입 지점이 그 블록 뒤 = 본문 맨 끝이 되어 서명이 정상 위치에 붙는다.
    (메일 레이아웃에서 흔히 쓰는 방식이라 클라이언트 호환성도 넓다)

    Args:
        body_text: 치환이 끝난 평문 본문 (표 자리에 `_HTML_TABLE_TOKEN`)
        html_table: 끼워 넣을 표 HTML

    Returns:
        HTML 문자열
    """
    blocks: list[str] = []
    for para in body_text.split('\n\n'):
        para = para.strip('\n')
        if not para:
            continue
        if para.strip() == _HTML_TABLE_TOKEN:
            blocks.append(html_table)
            continue
        # 문단 안의 줄바꿈만 <br>로 살린다
        blocks.append(
            '<p style="margin:0 0 12px 0;">'
            + html_escape(para).replace('\n', '<br>')
            + '</p>'
        )

    return (
        '<html><body>'
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0" '
        'style="border-collapse:collapse;"><tr><td '
        'style="font-family:맑은 고딕,sans-serif;font-size:13px;line-height:1.6;">'
        + ''.join(blocks) +
        '</td></tr></table>'
        '</body></html>'
    )


def _cell_text(value) -> str:
    """표 셀 표시값 (수량은 정수, 그 외는 문자열)"""
    if isinstance(value, float):
        return f"{value:,.0f}"
    return _clean_text(value)


def mail_summary(
    summary: pd.DataFrame,
    output_file: Path,
    biz_no: str,
    customer_name: str,
    opts: MailOptions,
) -> bool:
    """생성된 납기현황을 메일로 발송/초안 생성

    메일 실패는 문서 생성 성공을 뒤엎지 않습니다 (경고만 출력).

    Args:
        summary: 납기현황 요약 (본문 표)
        output_file: 생성된 xlsx 경로
        biz_no: 정규화된 사업자번호 (조회 기준 — 수신자 조회에 그대로 쓴다)
        customer_name: SO 기준 거래처명 (마스터에 이름이 없을 때 폴백)
        opts: 메일 옵션

    Returns:
        메일 생성/발송 성공 여부
    """
    if not opts.enabled:
        return True

    df_customer = opts.customer_master()
    if df_customer is None:
        return False

    try:
        recipient = find_recipient(
            biz_no, df_customer, fallback_name=customer_name, fixed_cc=DS_MAIL_CC,
        )
    except MailConfigError as e:
        print(f"  [메일 오류] {e}")
        return False

    if recipient is None:
        print(f"  [메일 생략] 수신자 미등록 — {customer_name} / {format_biz_no(biz_no)}")
        print(f"             {CUSTOMER_DOMESTIC_SHEET} 시트에 해당 사업자번호의 이메일을 입력하세요.")
        return False

    # 누구에게 나가는지 먼저 보여주고 확인받는다 (오발송 차단)
    print()
    print(f"  받는사람: {recipient.to_line}")
    if recipient.cc:
        print(f"  참조    : {recipient.cc_line}")
    print(f"  내용    : 미출고 {len(summary)}건 / 첨부 {output_file.name}")

    if opts.ask and not confirm("  이메일을 발송하시겠습니까? [y/N]: "):
        print("  -> 메일 생략")
        return False

    date_str = dt.date.today().strftime('%Y-%m-%d')
    extra = {
        'count': len(summary),
        'qty': f"{summary['수량'].sum():,.0f}",
        'table': build_text_table(summary),
    }
    # HTML 본문은 표 자리를 토큰으로 둔 채 렌더링하고, 문단을 블록으로 감싸며 표를 끼워 넣는다.
    # (표를 먼저 넣으면 이스케이프가 <table>을 통째로 문자로 만들어 버린다)
    body_html = build_html_body(
        render_template(
            DS_MAIL_BODY, recipient, '', date_str,
            extra=dict(extra, table=_HTML_TABLE_TOKEN),
        ),
        build_html_table(summary),
    )

    try:
        result = create_document_mail(
            xlsx_path=output_file,
            recipient=recipient,
            doc_id=format_biz_no(biz_no),
            subject_template=DS_MAIL_SUBJECT,
            body_template=DS_MAIL_BODY,
            date_str=date_str,
            send=opts.send,
            attach_format=DS_MAIL_ATTACH_FORMAT,
            backend=opts.backend,
            extra=extra,
            body_html=body_html,
            doc_label='납기현황',
            draft_prefix='ds_draft',
        )
    except MailConfigError as e:
        print(f"  [메일 오류] {e}")
        return False

    if result.success:
        attach_names = ', '.join(p.name for p in result.attachments)
        verb = "메일 발송 완료" if result.sent else "메일 초안 생성 (메일 창에서 [보내기] 확인)"
        print(f"  -> {verb}: {attach_names}")
        if opts.send and not result.sent:
            print("     [주의] 자동 발송이 안 되는 방식이라 초안까지만 진행했습니다.")
        return True

    print(f"  [메일 실패] {result.message}")
    return False


def warn_if_stale(stale_count: int) -> None:
    """캐시 Status가 계산 결과와 어긋나면 새로고침 안내"""
    if stale_count <= 0:
        return
    print()
    print(f"[주의] SO_국내 Status가 실제 출고와 다른 행 {stale_count}건 — 파워쿼리 새로고침이 밀렸습니다.")
    print("       이 회신 문서는 DN_국내 출고수량으로 직접 계산하므로 영향이 없지만,")
    print("       같은 파일을 보는 대시보드·피벗은 낡은 값을 씁니다. 엑셀에서 [모두 새로 고침] 하세요.")


# === CLI ==================================================================

def create_argument_parser() -> argparse.ArgumentParser:
    """CLI 인자 파서 생성"""
    parser = argparse.ArgumentParser(
        prog='delivery_status',
        description='거래처 납기현황 조회 — 사업자등록번호 기준 미출고 현황 xlsx 생성',
        epilog=(
            '예시:\n'
            '  python delivery_status.py 615-81-88675\n'
            '  python delivery_status.py 엔이에스\n'
            '  python delivery_status.py --list'
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        'customer',
        nargs='?',
        help='사업자등록번호 (하이픈 유무 무관) 또는 거래처명 일부',
    )
    parser.add_argument(
        '-l', '--list',
        action='store_true',
        dest='show_list',
        help='미출고 잔량이 있는 거래처 목록만 출력',
    )
    parser.add_argument(
        '-a', '--all',
        action='store_true',
        help='출고완료 건까지 포함 (기본은 미출고만)',
    )
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='상세 로그 출력',
    )
    add_mail_arguments(parser, '납기현황')
    return parser


def main() -> int:
    """메인 함수"""
    parser = create_argument_parser()
    args = parser.parse_args()
    setup_logging(verbose=args.verbose)

    # 비대화형 실행(배치·파이프)에서는 묻지 않는다 — input()에서 멈추면 안 된다
    mail_opts = MailOptions(mode=resolve_mail_mode(args, sys.stdin.isatty()))

    if not args.show_list and not args.customer:
        parser.print_help()
        print("\n[오류] 사업자등록번호 또는 거래처명을 입력하세요 (목록은 --list).")
        return 1

    # 1. 데이터 로드 + 출고상태 계산
    try:
        so, dn = load_sheets()
    except FileNotFoundError as e:
        print(f"[오류] {e}")
        return 1
    except PermissionError:
        print(f"[오류] 파일이 열려 있습니다. 엑셀을 닫고 다시 실행하세요: {NOAH_SO_PO_DN_FILE.name}")
        return 1

    missing = [c for c in ('SO_ID', 'Line item', 'Business registration number') if c not in so.columns]
    if missing:
        print(f"[오류] {SO_DOMESTIC_SHEET} 시트에 필요한 컬럼이 없습니다: {missing}")
        return 1

    work = attach_ship_status(so, dn)
    stale = count_stale_status(work)

    # 2. 목록 모드
    if args.show_list:
        print_customer_list(build_customer_list(work))
        warn_if_stale(stale)
        return 0

    # 3. 거래처 확정
    candidates, lookup_kind = resolve_customer(work, args.customer)
    if not candidates:
        print(f"[오류] {lookup_kind} '{args.customer}'에 해당하는 거래처를 찾을 수 없습니다.")
        print("       `python delivery_status.py --list`로 거래처 목록을 확인하세요.")
        return 1
    if len(candidates) > 1:
        print(f"[오류] '{args.customer}'에 해당하는 거래처가 {len(candidates)}곳입니다. 사업자번호로 지정하세요:")
        for biz, name in candidates:
            print(f"  - {format_biz_no(biz)}  {name}")
        return 1

    biz_no, customer_name = candidates[0]

    # 4. 집계
    rows = select_rows(work, biz_no, include_shipped=args.all)
    summary = build_summary(rows)
    detail = build_detail(rows)

    if summary.empty:
        scope = '주문' if args.all else '미출고 건'
        print(f"\n{customer_name} ({format_biz_no(biz_no)}) — 해당하는 {scope}이 없습니다.")
        warn_if_stale(stale)
        return 0

    # 5. 출력
    output_file = generate_output_filename('납기현황', customer_name, format_biz_no(biz_no), DS_OUTPUT_DIR)
    try:
        write_output(summary, detail, output_file, customer_name, biz_no)
    except PermissionError:
        print(f"[오류] 출력 파일이 열려 있습니다. 닫고 다시 실행하세요: {output_file.name}")
        return 1

    print_summary(summary, customer_name, biz_no)
    print(f"\n출력: {output_file}")

    # 6. 메일 (실패해도 문서 생성 성공을 뒤엎지 않는다)
    mail_summary(summary, output_file, biz_no, customer_name, mail_opts)

    warn_if_stale(stale)
    return 0


if __name__ == "__main__":
    sys.exit(main())
