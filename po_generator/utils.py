"""
유틸리티 함수
=============

데이터 로드, 값 추출 등 공통 유틸리티 함수를 제공합니다.
"""

from __future__ import annotations

import logging
from typing import Any

import pandas as pd

from po_generator.config import (
    NOAH_SO_PO_DN_FILE,
    NOAH_PO_LISTS_FILE,
    SO_DOMESTIC_SHEET,
    PO_DOMESTIC_SHEET,
    DN_DOMESTIC_SHEET,
    PMT_DOMESTIC_SHEET,
    SO_EXPORT_SHEET,
    PO_EXPORT_SHEET,
    DN_EXPORT_SHEET,
    CUSTOMER_EXPORT_SHEET,
    COLUMN_ALIASES,
    META_COLUMNS,
    SPEC_START_COLUMN,
    OPTION_START_COLUMN,
    SPEC_FIELDS,
    OPTION_FIELDS,
)

logger = logging.getLogger(__name__)

# Excel 수식 인젝션 방지용 문자
FORMULA_ESCAPE_CHARS: tuple[str, ...] = ('=', '+', '-', '@')


def escape_excel_formula(value: Any) -> Any:
    """Excel 수식 인젝션 방지

    사용자 입력이 =, +, -, @로 시작하면 Excel에서 수식으로 해석될 수 있음.
    이를 방지하기 위해 앞에 작은따옴표(')를 추가.

    Args:
        value: 원본 값

    Returns:
        이스케이프된 값 (문자열인 경우만 처리)
    """
    if isinstance(value, str) and value and value[0] in FORMULA_ESCAPE_CHARS:
        return "'" + value
    return value


def _get_safe_value(
    order_data: pd.Series,
    key: str,
    default: Any = ''
) -> Any:
    """안전하게 값 가져오기 (NaN 처리) - 내부용

    외부에서는 get_value()를 사용하세요. 별칭 매핑이 자동으로 적용됩니다.

    Args:
        order_data: 주문 데이터 Series
        key: 가져올 컬럼 키 (실제 컬럼명)
        default: 기본값 (값이 없거나 NaN인 경우)

    Returns:
        해당 키의 값 또는 기본값
    """
    value = order_data.get(key, default)

    # None 또는 pandas NaN/NA 체크
    if value is None or pd.isna(value):
        return default

    # 문자열 'nan' 체크 (대소문자 무관) - pandas가 가끔 이런 문자열을 반환
    if isinstance(value, str) and value.strip().lower() == 'nan':
        return default

    return value


# 하위 호환성을 위한 별칭 - Deprecated, get_value() 사용 권장
get_safe_value = _get_safe_value


def resolve_column(
    columns: pd.Index | list[str],
    key: str,
) -> str | None:
    """별칭에서 실제 컬럼명 찾기

    Args:
        columns: DataFrame의 컬럼 목록 (df.columns)
        key: 내부 키 (예: 'customer_name') 또는 실제 컬럼명

    Returns:
        실제 컬럼명 또는 None (찾지 못한 경우)
    """
    # 1. key가 이미 실제 컬럼명인 경우
    if key in columns:
        return key

    # 2. key가 내부 키인 경우, 별칭에서 찾기
    aliases = COLUMN_ALIASES.get(key)
    if aliases:
        for alias in aliases:
            if alias in columns:
                return alias

    # 3. 대소문자 무시 검색 (fallback)
    key_lower = key.lower()
    for col in columns:
        if col.lower() == key_lower:
            return col

    return None


def get_value(
    order_data: pd.Series,
    key: str,
    default: Any = '',
) -> Any:
    """내부 키로 값 가져오기 (별칭 지원) - 표준 API

    COLUMN_ALIASES에 정의된 내부 키 또는 실제 컬럼명을 사용할 수 있습니다.
    외부에서 데이터에 접근할 때는 이 함수를 사용하세요.

    Args:
        order_data: 주문 데이터 Series
        key: 내부 키 (예: 'customer_name') 또는 실제 컬럼명
        default: 기본값 (값이 없거나 NaN인 경우)

    Returns:
        해당 키의 값 또는 기본값
    """
    # 실제 컬럼명 찾기
    actual_col = resolve_column(order_data.index, key)

    if actual_col is None:
        return default

    return _get_safe_value(order_data, actual_col, default)


def _load_and_merge_sheets(
    xl: pd.ExcelFile,
    so_sheet: str,
    po_sheet: str,
    sheet_type: str,
) -> pd.DataFrame:
    """SO와 PO 시트를 로드하고 SO_ID로 병합

    Args:
        xl: ExcelFile 객체
        so_sheet: SO 시트명
        po_sheet: PO 시트명
        sheet_type: 시트 구분 ('국내' 또는 '해외')

    Returns:
        병합된 DataFrame (PO 기준, SO 정보 포함)
    """
    # SO 시트 로드 (고객 정보)
    df_so = pd.read_excel(xl, sheet_name=so_sheet)

    # PO 시트 로드 (발주 정보 + 사양)
    df_po = pd.read_excel(xl, sheet_name=po_sheet)

    # SO에서 필요한 컬럼만 선택 (PO에 없는 것들)
    # PO_ID가 없는 행(빈 행) 제외
    df_po = df_po[df_po['PO_ID'].notna()].copy()

    # SO_ID로 병합 (PO 기준 left join)
    # SO에서 가져올 컬럼: Customer PO, Customer name, Incoterms 등 (PO에 없는 것들)
    so_cols_to_merge = ['SO_ID', 'Customer PO', 'Customer name', 'Incoterms',
                        'Opportunity', 'Sector', 'Industry code',
                        'Sales Unit Price', 'Sales amount', 'Currency',
                        'PO receipt date', 'Requested delivery date', '납품 주소',
                        'Model number', 'Item name']
    # SO에 실제 존재하는 컬럼만 선택
    so_cols_to_merge = [c for c in so_cols_to_merge if c in df_so.columns]

    df_so_subset = df_so[so_cols_to_merge].copy()

    # SO_ID가 같은 행이 여러 개일 수 있음 (다중 아이템)
    # SO에서 SO_ID별 첫 행만 가져옴 (Customer name, Customer PO 등은 동일하므로)
    df_so_unique = df_so_subset.drop_duplicates(subset='SO_ID', keep='first')

    # PO 기준으로 left join
    df_merged = df_po.merge(df_so_unique, on='SO_ID', how='left', suffixes=('', '_SO'))

    # 시트 구분 추가
    df_merged['_시트구분'] = sheet_type

    return df_merged


def load_noah_po_lists() -> pd.DataFrame:
    """NOAH_SO_PO_DN.xlsx에서 데이터 로드

    국내/해외의 SO+PO 시트를 병합하여 하나의 DataFrame으로 합칩니다.
    새 파일이 없으면 기존 NOAH_PO_Lists.xlsx에서 로드합니다.

    Returns:
        합쳐진 주문 데이터 DataFrame

    Raises:
        FileNotFoundError: 소스 파일이 없는 경우
    """
    # 새 파일 우선 사용
    if NOAH_SO_PO_DN_FILE.exists():
        logger.info(f"데이터 로드: {NOAH_SO_PO_DN_FILE.name}")
        with pd.ExcelFile(NOAH_SO_PO_DN_FILE) as xl:
            # 국내 데이터 로드 및 병합
            df_domestic = _load_and_merge_sheets(
                xl, SO_DOMESTIC_SHEET, PO_DOMESTIC_SHEET, '국내'
            )

            # 해외 데이터 로드 및 병합
            df_export = _load_and_merge_sheets(
                xl, SO_EXPORT_SHEET, PO_EXPORT_SHEET, '해외'
            )

        # concat: pandas가 자동으로 없는 컬럼에 NaN 채움
        dfs = [df for df in [df_domestic, df_export] if len(df) > 0]
        df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

        logger.info(f"총 {len(df)}건의 주문 데이터 로드 완료")
        return df

    # 기존 파일로 폴백
    if NOAH_PO_LISTS_FILE.exists():
        logger.info(f"데이터 로드 (Legacy): {NOAH_PO_LISTS_FILE.name}")
        with pd.ExcelFile(NOAH_PO_LISTS_FILE) as xl:
            df_domestic = pd.read_excel(xl, sheet_name=0)
            df_export = pd.read_excel(xl, sheet_name=1)

        df_domestic['_시트구분'] = '국내'
        df_export['_시트구분'] = '해외'

        # concat: pandas가 자동으로 없는 컬럼에 NaN 채움
        dfs = [df for df in [df_domestic, df_export] if len(df) > 0]
        df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

        logger.info(f"총 {len(df)}건의 주문 데이터 로드 완료")
        return df

    raise FileNotFoundError(
        f"소스 파일을 찾을 수 없습니다: {NOAH_SO_PO_DN_FILE} 또는 {NOAH_PO_LISTS_FILE}"
    )


def _find_data_by_id(
    df: pd.DataFrame,
    column_key: str,
    id_value: str,
    id_label: str,
    allow_multiple: bool = True,
) -> pd.Series | pd.DataFrame | None:
    """ID로 데이터 검색 (공통 로직)

    Args:
        df: 검색 대상 DataFrame
        column_key: COLUMN_ALIASES의 키 (예: 'order_no', 'dn_id')
        id_value: 검색할 ID 값
        id_label: 로그 메시지용 라벨 (예: '주문번호', 'DN_ID')
        allow_multiple: True면 다중 아이템 시 DataFrame 반환

    Returns:
        - 단일 아이템: pd.Series
        - 다중 아이템 (allow_multiple=True): pd.DataFrame
        - 없음: None
    """
    col = resolve_column(df.columns, column_key)
    if col is None:
        logger.error(f"{id_label} 컬럼을 찾을 수 없습니다.")
        return None

    mask = df[col] == id_value
    match_count = mask.sum()

    if match_count == 0:
        logger.warning(f"{id_label} '{id_value}'를 찾을 수 없습니다.")
        return None

    if allow_multiple and match_count > 1:
        logger.info(f"{id_label} '{id_value}': {match_count}개 아이템 발견 (다중 아이템)")
        return df[mask]

    logger.info(f"{id_label} '{id_value}': {'발견' if not allow_multiple else '단일 아이템'}")
    return df[mask].iloc[0]


def find_order_data(
    df: pd.DataFrame,
    order_no: str
) -> pd.Series | pd.DataFrame | None:
    """PO_ID (또는 RCK Order No.)로 주문 데이터 검색

    Args:
        df: 전체 주문 데이터
        order_no: 검색할 PO_ID (예: ND-0001, NO-0001)

    Returns:
        - 단일 아이템: pd.Series
        - 다중 아이템: pd.DataFrame
        - 없음: None
    """
    return _find_data_by_id(df, 'order_no', order_no, '주문번호')


def format_currency(value: float, currency: str = 'KRW') -> str:
    """통화 포맷팅

    Args:
        value: 금액
        currency: 통화 코드 (KRW 또는 USD)

    Returns:
        포맷팅된 통화 문자열 (예: "₩1,000,000" 또는 "$1,000.00")
    """
    if currency == 'KRW':
        return f"₩{value:,.0f}"
    return f"${value:,.2f}"


# === DN (납품) 데이터 로드 ===

def load_dn_data() -> pd.DataFrame:
    """DN_국내 데이터 로드 (SO 정보 포함)

    DN 시트의 DN_ID/SO_ID를 기준으로 SO 시트에서 품목/금액 정보를 가져옵니다.
    Single Source of Truth: Item 정보는 SO_국내에서만 관리.

    Returns:
        DN 데이터 DataFrame (SO 정보 포함)

    Raises:
        FileNotFoundError: 소스 파일이 없는 경우
    """
    if not NOAH_SO_PO_DN_FILE.exists():
        raise FileNotFoundError(f"소스 파일을 찾을 수 없습니다: {NOAH_SO_PO_DN_FILE}")

    logger.info(f"DN 데이터 로드: {NOAH_SO_PO_DN_FILE.name}")
    with pd.ExcelFile(NOAH_SO_PO_DN_FILE) as xl:
        # DN_국내 로드 (DN 고유 정보만: DN_ID, SO_ID, 납품일 등)
        df_dn = pd.read_excel(xl, sheet_name=DN_DOMESTIC_SHEET)
        # SO_국내에서 품목/금액/고객 정보 가져오기
        df_so = pd.read_excel(xl, sheet_name=SO_DOMESTIC_SHEET)

    df_dn = df_dn[df_dn['DN_ID'].notna()].copy()
    # DN_ID 중복 제거 (DN_ID는 고유해야 함, SO_ID만 참조용)
    df_dn = df_dn.drop_duplicates(subset='DN_ID', keep='first')
    so_cols = [
        'SO_ID',
        'Customer name',        # 고객명
        'Customer PO',          # PO No.
        'Item name',            # 품목명
        'Item qty',             # 수량
        'Sales Unit Price',     # 판매단가
        'Total Sales',          # 총 판매금액
        'Business registration number',
    ]
    so_cols = [c for c in so_cols if c in df_so.columns]

    # SO는 다중 아이템 가능 → drop_duplicates 하지 않음
    df_so_subset = df_so[so_cols].copy()

    # SO_ID로 조인 (DN 1건에 SO 다중 아이템 → 결과도 다중 행)
    df_merged = df_dn.merge(df_so_subset, on='SO_ID', how='left', suffixes=('', '_SO'))
    df_merged['_시트구분'] = '국내'
    df_merged['_문서유형'] = 'DN'

    logger.info(f"DN 데이터 {len(df_merged)}건 로드 완료")
    return df_merged


def find_dn_data(
    df: pd.DataFrame,
    dn_id: str
) -> pd.Series | pd.DataFrame | None:
    """DN_ID로 납품 데이터 검색

    Args:
        df: DN 데이터
        dn_id: 검색할 DN_ID (예: DN-2026-0001)

    Returns:
        - 단일 아이템: pd.Series
        - 다중 아이템: pd.DataFrame
        - 없음: None
    """
    return _find_data_by_id(df, 'dn_id', dn_id, 'DN_ID')


# === PMT (입금/선수금) 데이터 로드 ===

def load_pmt_data() -> pd.DataFrame:
    """PMT_국내 데이터 로드 (SO 정보 포함)

    PMT 시트와 SO 시트를 SO_ID로 조인하여 반환합니다.

    Returns:
        PMT 데이터 DataFrame (SO 정보 포함)

    Raises:
        FileNotFoundError: 소스 파일이 없는 경우
    """
    if not NOAH_SO_PO_DN_FILE.exists():
        raise FileNotFoundError(f"소스 파일을 찾을 수 없습니다: {NOAH_SO_PO_DN_FILE}")

    logger.info(f"PMT 데이터 로드: {NOAH_SO_PO_DN_FILE.name}")
    with pd.ExcelFile(NOAH_SO_PO_DN_FILE) as xl:
        # PMT_국내 로드
        df_pmt = pd.read_excel(xl, sheet_name=PMT_DOMESTIC_SHEET)
        # SO_국내에서 거래명세표에 필요한 모든 정보 가져오기
        df_so = pd.read_excel(xl, sheet_name=SO_DOMESTIC_SHEET)

    df_pmt = df_pmt[df_pmt['선수금_ID'].notna()].copy()
    # 선수금_ID 중복 제거 (선수금_ID는 고유해야 함)
    df_pmt = df_pmt.drop_duplicates(subset='선수금_ID', keep='first')
    so_cols = [
        'SO_ID',
        'Customer name',        # 고객명
        'Customer PO',          # PO No.
        'Item name',            # 품목명
        'Item qty',             # 수량
        'Sales Unit Price',     # 판매단가
        'Total Sales',          # 총 판매금액
        'Business registration number',
    ]
    so_cols = [c for c in so_cols if c in df_so.columns]
    # 목록 표시용이므로 SO_ID별 첫 행만 사용 (고객명 등 대표 정보만 필요)
    # 실제 ADV 처리는 load_so_for_advance()에서 전체 아이템 로드
    df_so_subset = df_so[so_cols].drop_duplicates(subset='SO_ID', keep='first')

    # SO_ID로 조인
    df_merged = df_pmt.merge(df_so_subset, on='SO_ID', how='left', suffixes=('', '_SO'))
    df_merged['_시트구분'] = '국내'
    df_merged['_문서유형'] = 'PMT'

    logger.info(f"PMT 데이터 {len(df_merged)}건 로드 완료")
    return df_merged


def find_pmt_data(
    df: pd.DataFrame,
    advance_id: str
) -> pd.Series | None:
    """선수금_ID로 입금 데이터 검색

    Args:
        df: PMT 데이터
        advance_id: 검색할 선수금_ID (예: ADV_2026-0001)

    Returns:
        - pd.Series (항상 단일 건)
        - 없음: None
    """
    return _find_data_by_id(df, 'advance_id', advance_id, '선수금_ID', allow_multiple=False)


# === SO (Sales Order) 데이터 로드 (선수금용) ===

def load_so_for_advance(advance_id: str) -> tuple[pd.Series, pd.DataFrame] | None:
    """선수금_ID로 SO 아이템 데이터 로드

    PMT_국내에서 선수금_ID로 SO_ID를 찾고,
    SO_국내에서 해당 SO의 모든 아이템을 반환합니다.

    Args:
        advance_id: 선수금_ID (예: ADV_2026-0008)

    Returns:
        - (pmt_data, so_items_df): PMT 정보와 SO 아이템 DataFrame
        - None: 찾을 수 없는 경우
    """
    if not NOAH_SO_PO_DN_FILE.exists():
        raise FileNotFoundError(f"소스 파일을 찾을 수 없습니다: {NOAH_SO_PO_DN_FILE}")

    with pd.ExcelFile(NOAH_SO_PO_DN_FILE) as xl:
        # 1. PMT_국내에서 선수금_ID로 SO_ID 찾기
        df_pmt = pd.read_excel(xl, sheet_name=PMT_DOMESTIC_SHEET)
        # 2. SO_국내 로드
        df_so = pd.read_excel(xl, sheet_name=SO_DOMESTIC_SHEET)

    df_pmt = df_pmt[df_pmt['선수금_ID'].notna()].copy()

    pmt_row = df_pmt[df_pmt['선수금_ID'] == advance_id]
    if len(pmt_row) == 0:
        logger.warning(f"선수금_ID '{advance_id}'를 찾을 수 없습니다.")
        return None

    pmt_data = pmt_row.iloc[0]
    so_id = pmt_data['SO_ID']
    logger.info(f"선수금_ID '{advance_id}' -> SO_ID: {so_id}")

    # SO_국내에서 해당 SO_ID의 모든 아이템 가져오기
    so_items = df_so[df_so['SO_ID'] == so_id].copy()

    if len(so_items) == 0:
        logger.warning(f"SO_ID '{so_id}'에 해당하는 SO 데이터를 찾을 수 없습니다.")
        return None

    # SO 아이템에 PMT 정보 추가 (선수금_ID)
    so_items['선수금_ID'] = advance_id
    so_items['_시트구분'] = '국내'
    so_items['_문서유형'] = 'ADV'

    logger.info(f"SO_ID '{so_id}': {len(so_items)}개 아이템 발견")
    return pmt_data, so_items


# === SO 해외 데이터 로드 (Proforma Invoice / Commercial Invoice용) ===

def load_so_export_data() -> pd.DataFrame:
    """SO_해외 데이터 로드

    Returns:
        SO_해외 DataFrame

    Raises:
        FileNotFoundError: 소스 파일이 없는 경우
    """
    if not NOAH_SO_PO_DN_FILE.exists():
        raise FileNotFoundError(f"소스 파일을 찾을 수 없습니다: {NOAH_SO_PO_DN_FILE}")

    logger.info(f"SO 해외 데이터 로드: {NOAH_SO_PO_DN_FILE.name}")
    with pd.ExcelFile(NOAH_SO_PO_DN_FILE) as xl:
        # SO_해외 로드 (Model 컬럼을 문자열로 읽어 앞 0 보존)
        df_so = pd.read_excel(
            xl,
            sheet_name=SO_EXPORT_SHEET,
            dtype={'Model': str, 'Model number': str}
        )

    df_so = df_so[df_so['SO_ID'].notna()].copy()
    df_so['_시트구분'] = '해외'

    logger.info(f"SO 해외 데이터 {len(df_so)}건 로드 완료")
    return df_so


def find_so_export_data(
    df: pd.DataFrame,
    so_id: str
) -> pd.Series | pd.DataFrame | None:
    """SO_ID로 해외 SO 데이터 검색

    Args:
        df: SO 해외 데이터
        so_id: 검색할 SO_ID (예: SOO-2026-0001)

    Returns:
        - 단일 아이템: pd.Series
        - 다중 아이템: pd.DataFrame
        - 없음: None
    """
    return _find_data_by_id(df, 'so_id', so_id, 'SO_ID')


# === DN 해외 데이터 로드 (Final Invoice용) ===

def load_dn_export_data() -> pd.DataFrame:
    """DN_해외 데이터 로드 (SO_해외 + Customer_해외 JOIN)

    3시트 JOIN: DN_해외 → SO_해외(SO_ID) → Customer_해외(고객코드)
    Bill to 1/2/3, Payment terms 컬럼이 추가됩니다.

    Returns:
        DN 해외 데이터 DataFrame (고객 청구 주소 포함)

    Raises:
        FileNotFoundError: 소스 파일이 없는 경우
    """
    if not NOAH_SO_PO_DN_FILE.exists():
        raise FileNotFoundError(f"소스 파일을 찾을 수 없습니다: {NOAH_SO_PO_DN_FILE}")

    logger.info(f"DN 해외 데이터 로드: {NOAH_SO_PO_DN_FILE.name}")
    with pd.ExcelFile(NOAH_SO_PO_DN_FILE) as xl:
        # 1. DN_해외 로드
        df_dn = pd.read_excel(xl, sheet_name=DN_EXPORT_SHEET)
        # 2. SO_해외 로드 (고객코드 추출용)
        df_so = pd.read_excel(xl, sheet_name=SO_EXPORT_SHEET)
        # 3. Customer_해외 로드 (Bill to, Payment terms)
        df_cust = pd.read_excel(xl, sheet_name=CUSTOMER_EXPORT_SHEET)

    # DN_ID가 있는 행만 사용
    df_dn = df_dn[df_dn['DN_ID'].notna()].copy()

    # SO_해외에서 SO_ID별 고객코드 + Customer PO + PO receipt date 추출
    so_code_cols = ['SO_ID', 'Business registration number', 'Customer PO', 'PO receipt date']
    so_code_cols = [c for c in so_code_cols if c in df_so.columns]
    df_so_codes = df_so[so_code_cols].drop_duplicates(subset='SO_ID', keep='first')

    # DN_해외 + SO_해외 JOIN (SO_ID) → 고객코드 획득
    df_dn = df_dn.merge(df_so_codes, on='SO_ID', how='left', suffixes=('', '_SO'))

    # Customer_해외에서 필요한 컬럼만 추출
    cust_cols = ['C-code by 해외', 'Bill to 1', 'Bill to 2', 'Bill to 3', 'Payment terms']
    cust_cols = [c for c in cust_cols if c in df_cust.columns]
    df_cust_subset = df_cust[cust_cols].copy()

    # DN_해외+SO → Customer_해외 JOIN (고객코드)
    if 'Business registration number' in df_dn.columns and 'C-code by 해외' in df_cust_subset.columns:
        df_dn = df_dn.merge(
            df_cust_subset,
            left_on='Business registration number',
            right_on='C-code by 해외',
            how='left',
            suffixes=('', '_CUST'),
        )

    df_dn['_시트구분'] = '해외'
    df_dn['_문서유형'] = 'FI'

    logger.info(f"DN 해외 데이터 {len(df_dn)}건 로드 완료 (Customer JOIN 포함)")
    return df_dn


def find_dn_export_data(
    df: pd.DataFrame,
    dn_id: str
) -> pd.Series | pd.DataFrame | None:
    """DN_ID로 해외 DN 데이터 검색

    Args:
        df: DN 해외 데이터
        dn_id: 검색할 DN_ID (예: DNO-2026-0001)

    Returns:
        - 단일 아이템: pd.Series
        - 다중 아이템: pd.DataFrame
        - 없음: None
    """
    return _find_data_by_id(df, 'dn_id', dn_id, 'DN_ID (해외)')


# === 동적 사양/옵션 필드 추출 ===

# 캐시: 시트별 필드 목록 저장
_spec_option_fields_cache: dict[str, tuple[list[str], list[str]]] = {}



def get_spec_option_fields(
    sheet_type: str = '국내',
    force_reload: bool = False,
) -> tuple[list[str], list[str]]:
    """PO 시트에서 동적으로 사양/옵션 필드 추출

    META_COLUMNS에 정의된 메타 컬럼을 제외한 나머지를
    SPEC_START_COLUMN ~ Status 전까지 → SPEC_FIELDS
    OPTION_START_COLUMN ~ 끝까지 → OPTION_FIELDS
    로 분류합니다.

    Args:
        sheet_type: '국내' 또는 '해외'
        force_reload: True이면 캐시 무시하고 다시 로드

    Returns:
        (spec_fields, option_fields) 튜플
    """
    # 캐시 확인
    if not force_reload and sheet_type in _spec_option_fields_cache:
        return _spec_option_fields_cache[sheet_type]

    # 시트명 결정
    po_sheet = PO_DOMESTIC_SHEET if sheet_type == '국내' else PO_EXPORT_SHEET

    try:
        if not NOAH_SO_PO_DN_FILE.exists():
            logger.warning(f"데이터 파일 없음, 기본 필드 사용: {NOAH_SO_PO_DN_FILE}")
            return list(SPEC_FIELDS), list(OPTION_FIELDS)

        # PO 시트 컬럼 목록 가져오기
        with pd.ExcelFile(NOAH_SO_PO_DN_FILE) as xl:
            df = pd.read_excel(xl, sheet_name=po_sheet, nrows=0)  # 헤더만 읽기
            columns = list(df.columns)

        # 마커 컬럼 위치 찾기
        spec_start_idx = None
        option_start_idx = None

        for idx, col in enumerate(columns):
            if col == SPEC_START_COLUMN:
                spec_start_idx = idx
            elif col == OPTION_START_COLUMN:
                option_start_idx = idx

        if spec_start_idx is None or option_start_idx is None:
            logger.warning("마커 컬럼을 찾을 수 없음, 기본 필드 사용")
            return list(SPEC_FIELDS), list(OPTION_FIELDS)

        # SPEC: spec_start_idx ~ option_start_idx 전까지 (META 제외)
        # OPTION: option_start_idx ~ 끝까지 (META 제외)
        spec_fields = []
        option_fields = []

        for idx, col in enumerate(columns):
            if col in META_COLUMNS:
                continue

            if spec_start_idx <= idx < option_start_idx:
                spec_fields.append(col)
            elif idx >= option_start_idx:
                option_fields.append(col)

        logger.info(f"동적 필드 추출 완료 ({sheet_type}): SPEC {len(spec_fields)}개, OPTION {len(option_fields)}개")

        # 캐시 저장
        _spec_option_fields_cache[sheet_type] = (spec_fields, option_fields)

        return spec_fields, option_fields

    except Exception as e:
        logger.warning(f"동적 필드 추출 실패: {e}, 기본 필드 사용")
        return list(SPEC_FIELDS), list(OPTION_FIELDS)
