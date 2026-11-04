"""
AX Sales order book — Item requirement(IR) + Item transaction(IT)로 만드는 오더북
==============================================================================

회계팀 오더북 파일(`Sales order book_MM_YYYY.xlsx`)의 **Raw** 시트 형식으로, AX export 두 개만 가지고 오더북을 만든다.
AX Sales Order Book(Table 50229)처럼 **달마다 계산해 저장**하고, 다음 달은 저장된 End에서 시작한다 — 저장은
ax_orderbook_store.py가 하고 여기는 계산만 한다(저장소를 모른다). 처음 한 번은 같은 IR·IT로 P01~마감월을 한 번에
만든다(기준선 — NOAH 파워쿼리 `Order_Book`처럼 수주 라인을 월별로 굴린다).

    IR (Item requirement) = 수주 라인: 생성일·수량·단가·상태
    IT (Item transaction) = 출고 1건 = 1행: Project date·수량·금액·통화, Transaction status = 청구 여부

    Start    = 기준선 P01은 회계 P01 Start(키별), 그 뒤로는 직전 달 End(기준선 안에선 계산값, 그 뒤엔 저장값)
    Input    = 그달 생성된 라인(KST) × 수주월 환율
    Output   = 그달 말까지 청구된 출고 누계 − 이미 Output으로 인식한 누계, × 그달 환율. 기준선은 같은 추출본의 지난 달 말
               누계를 빼므로 '그달 출고분 중 청구된 것'이고(IT엔 청구 날짜가 없어 출고월), 그 뒤로는 저장된 누계를 빼므로
               지난 마감 때 청구 전이던 출고가 청구되면 그 달 Output이다(AX처럼 청구한 달)
    End      = (라인 수량 − 월말까지 출고 수량) × 단가 × 수주월 환율 + 월말까지 출고했지만 청구 전(Posted)인 금액
    Variance = End − Start − Input + Output   (환율 재평가·수주/청구 금액 차이·지난 계산 뒤의 취소와 변경 — AX도 같은 뜻)

IR의 `Invoiced`는 고객 청구가 아니라 **출고됨**이다(프로젝트 품목요구는 출고만 올려도 Invoiced). 청구 여부는 IT에만 있다.
Cancelled 라인은 처음부터 뺀다(NOAH Order_Book이 Cancelled·Hold를 빼는 것과 같다). 기준선 안의 지난 달은 지금 데이터로
계산되고, 저장한 뒤의 취소·변경은 그 일이 보인 달의 Variance로 들어간다.
키 = Project + Cost center + Prod/Dept(Revenue). 키 매핑과 Raw 속성은 회계 오더북에서 배우고, 숫자는 IR·IT로만 낸다.
실측 근거는 tasks/todo.md. Excel/COM은 쓰지 않는다(읽기는 pandas, 환율 파일만 openpyxl).
"""

from __future__ import annotations

import logging
import re
import sqlite3
from collections import Counter
from collections.abc import Sequence
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)

# === Raw 시트 (회계팀 오더북, 48열) ===
RAW_COLUMNS: tuple[str, ...] = (
    'Project', 'Report Start', 'Period', 'Account number', 'Customer Name', 'Cost center', 'Key',
    'Adjusted orderbook code', 'Category', 'Restate', 'Contract Engineer', 'Contract Engineer Name',
    'CE Project', 'Project group', 'Name', 'Qty Start', 'Qty Input', 'Qty Output', 'Qty End',
    'Qty Variance', 'Value Start', 'Value Input', 'Value Output', 'Value End', 'Value Variance',
    'Estimate Margin (%)', 'Total Qty input', 'Total Value input', 'Receipt date',
    'Project creation date', 'Requested ship date', 'Market', 'Order Book Code', 'Industry',
    'Final destination country', 'Location', 'Prod/Dept (Revenue)', 'Period status',
    'Account Manager', 'Account Manager Name', 'Country', 'Sanctioned', 'MCA Number', 'd',
    'Sector Check', 'OS Industry', 'OS Product', ' Inter Co',
)
# 회계팀이 수식으로 채우는 열 — 비워 둔다. 붙여 넣은 뒤 위 행의 수식을 끌어내리면 된다
FORMULA_COLUMNS: frozenset[str] = frozenset({
    'Key', 'Adjusted orderbook code', 'Category', 'Restate', 'Contract Engineer Name',
    'Total Qty input', 'Total Value input', 'Account Manager Name', 'Country', 'Sanctioned',
    'Sector Check', 'OS Industry', 'OS Product', ' Inter Co',
})
NUM_COLUMNS: tuple[str, ...] = (
    'Qty Start', 'Qty Input', 'Qty Output', 'Qty End', 'Qty Variance',
    'Value Start', 'Value Input', 'Value Output', 'Value End', 'Value Variance',
)
# 기존 키는 회계 오더북의 마지막 행에서 그대로 물려받는 AX 원본 속성
CARRY_COLUMNS: tuple[str, ...] = (
    'Account number', 'Customer Name', 'Cost center', 'Contract Engineer', 'CE Project',
    'Project group', 'Name', 'Estimate Margin (%)', 'Receipt date', 'Project creation date',
    'Requested ship date', 'Market', 'Order Book Code', 'Industry', 'Final destination country',
    'Location', 'Prod/Dept (Revenue)', 'Account Manager', 'MCA Number', 'd',
)
# 같은 프로젝트의 다른 키에서 빌려 올 수 있는 프로젝트 단위 속성 (Order Book Code는 키 단위라 제외)
PROJECT_LEVEL_COLUMNS: tuple[str, ...] = (
    'Account number', 'Customer Name', 'Contract Engineer', 'CE Project', 'Project group', 'Name',
    'Market', 'Industry', 'Final destination country', 'Account Manager', 'MCA Number',
)
# 신규 키에서 사람이 채워야 하는 속성 — 비어 있으면 확인 목록에 올린다
NEW_KEY_REQUIRED: tuple[str, ...] = ('Market', 'Industry', 'Order Book Code', 'Account Manager', 'Name')
_KEY_VALUES: tuple[str, ...] = ('Sq', 'Iq', 'Oq', 'Eq', 'Vq', 'Sv', 'Iv', 'Ov', 'Ev', 'Vv')

# === Item requirement export ===
IR_REQUIRED: tuple[str, ...] = (
    'Project', 'Line number', 'Created date', 'Customer', 'Customer name', 'Project group',
    'Item number', 'Item name', 'Quantity', 'Unit price', 'Net amount', 'Line status',
    'Actual ship date', 'Deliver remainder', 'Division', 'Prod/Dept (Material)',
    'Contract engineer', 'VAT group',
)
IR_TEXT_COLUMNS: tuple[str, ...] = (
    'Customer', 'Project', 'Line number', 'Item number', 'Division', 'Prod/Dept (Material)',
    'VAT group', 'Line status', 'Customer name', 'Contract engineer', 'Project group',
)
# AX createdDateTime은 UTC로 찍힌다 — 프로젝트 생성일과 KST 98.9% / UTC 90.6% 일치 (2026-10-29 실측)
KST_OFFSET = pd.Timedelta(hours=9)

# === Item transaction export (프로젝트 품목 거래 — 출고 1건 = 1행) ===
IT_REQUIRED: tuple[str, ...] = (
    'Project date', 'Project', 'Line number', 'Item number', 'Item name', 'Quantity', 'Net amount',
    'Sales currency', 'Line property', 'Transaction status',
)
IT_TEXT_COLUMNS: tuple[str, ...] = (
    'Project', 'Item number', 'Sales currency', 'Line property', 'Transaction status', 'Customer account',
)
# Transaction status: Invoiced = 고객 청구 완료 / Posted = 출고됐지만 청구 전(크레딧·반품도 음수 Posted로 남는다) /
# Adjusted = 정정으로 대체된 행(합계 0) — 2026-11-02 실측(82-15300.001 Posted = AX에서 미청구 확인)
IT_BILLED, IT_PENDING, IT_ADJUSTED = 'Invoiced', 'Posted', 'Adjusted'

# IR엔 통화 열이 없다. 외화 거래처는 '공식 오더북 Output ÷ IR Net amount'가 그 달 월환율과
# 1e-6까지 맞는 것으로 역추적했다(2026-10-29). 새 외화 거래처가 생기면 여기에 한 줄.
FOREIGN_CUSTOMER_CURRENCY: dict[str, str] = {
    'C00014': 'EUR', 'C00033': 'USD', 'C00037': 'USD', 'C00054': 'GBP', 'C02286': 'JPY',
    'C02642': 'USD', 'C02643': 'EUR', 'C02644': 'EUR', 'C02645': 'USD', 'C02646': 'USD',
    'C02647': 'USD', 'C02648': 'USD', 'C02649': 'USD', 'C02650': 'EUR', 'C02651': 'EUR',
    'C02652': 'USD', 'C02653': 'EUR', 'C02654': 'USD', 'C02732': 'USD', 'C02753': 'USD',
    'C02814': 'EUR',
}
# 외화 거래처는 전부 VAT group ZR/00이다. 통화표에 없는 ZR/00 거래처는 외화일 수 있어 확인 목록에 올리되,
# 원화로 확인된 곳은 뺀다(단가가 원화 규모 — 2026-11-02 실측).
KRW_ZERO_RATED_CUSTOMERS: frozenset[str] = frozenset({
    'C01011', 'C01040', 'C01047', 'C01059', 'C02370', 'C02493',
})
ZERO_RATED_VAT_GROUPS: frozenset[str] = frozenset({'ZR', '00'})

# NOAH FX 시트엔 USD/EUR/GBP만 있다. JPY는 공식 오더북 Motoyama Output ÷ IR 금액으로 역산했다(소수 4자리 그대로).
FX_SEED: dict[tuple[str, str], float] = {
    ('JPY', '2026-01'): 9.2159, ('JPY', '2026-02'): 9.3580, ('JPY', '2026-03'): 9.1261,
    ('JPY', '2026-04'): 9.4347, ('JPY', '2026-05'): 9.2479, ('JPY', '2026-06'): 9.4041,
    ('JPY', '2026-07'): 9.5269, ('JPY', '2026-08'): 8.8645, ('JPY', '2026-09'): 8.6968,
    ('JPY', '2026-10'): 8.6157,
}
MANUAL_FX_SOURCE = '직접 입력'
_CLOSING_FILE_RE = re.compile(r'Closing rates\s*-\s*P(\d{1,2})\s+(\d{4})', re.IGNORECASE)

BIG_VARIANCE = 10_000_000      # 확인 목록에 올리는 Variance 크기 (원)
_EPS = 1e-6


# ---------------------------------------------------------------------------
# 기간
# ---------------------------------------------------------------------------
def parse_period(text: str) -> int:
    """'P10' / '10' → 10"""
    m = re.fullmatch(r'[Pp]?(\d{1,2})', text.strip())
    if not m or not 1 <= int(m.group(1)) <= 12:
        raise ValueError(f"기간 형식이 아닙니다: {text!r} (예: P10)")
    return int(m.group(1))


def month_bounds(year: int, month: int) -> tuple[pd.Timestamp, pd.Timestamp]:
    """(마감월 1일, 다음 달 1일)"""
    start = pd.Timestamp(year, month, 1)
    return start, start + pd.offsets.MonthBegin(1)


def ym(ts: pd.Timestamp) -> str:
    return f"{ts.year}-{ts.month:02d}"


def normalize_line_number(value) -> str | None:
    """IR은 '1.001'·'14'(문자), IT는 1.001·14.0(숫자)으로 같은 라인을 적는다 → '1.001'·'14'로 맞춘다."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    text = str(value).strip()
    try:
        return f"{float(text):.3f}".rstrip('0').rstrip('.')
    except ValueError:
        return text or None


# ---------------------------------------------------------------------------
# 읽기
# ---------------------------------------------------------------------------
def read_raw(path: Path) -> pd.DataFrame:
    """회계팀 오더북 파일의 Raw 시트 → 숫자 열은 0 채움, 'pn'(기간 번호) 추가"""
    text_cols = ['Project', 'Account number', 'Cost center', 'Key', 'Contract Engineer', 'Location',
                 'Prod/Dept (Revenue)', 'Account Manager', 'Order Book Code', 'Industry', 'Period']
    df = pd.read_excel(path, sheet_name='Raw', dtype={c: str for c in text_cols})
    missing = [c for c in RAW_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"{path.name} Raw 시트에 없는 열: {missing}")
    for c in NUM_COLUMNS:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
    df['Period'] = df['Period'].astype(str).str.strip()
    df['pn'] = pd.to_numeric(df['Period'].str.extract(r'(\d+)')[0], errors='coerce').astype('Int64')
    # 이 파일의 연도 — Report Start가 그 달 1일이다(다른 달이 한두 행 섞여 있어 최빈값으로)
    years = pd.to_datetime(df['Report Start'], errors='coerce').dt.year.dropna()
    df['year'] = int(years.mode().iloc[0]) if len(years) else pd.NA
    # Key는 수식(Project & Cost center & Prod/Dept) — 캐시가 비어도 다시 만들 수 있게
    df['Key'] = df['Project'].fillna('') + df['Cost center'].fillna('') + df['Prod/Dept (Revenue)'].fillna('')
    return df


def read_ir(paths: list[Path]) -> pd.DataFrame:
    """IR export(여러 개면 합친다) → 라인키로 중복 제거(나중 파일 우선), KST 생성일 추가.

    라인번호는 재사용된다(삭제 후 같은 번호로 재생성) — 라인키는 Project + Line number + Created date.
    """
    frames = []
    for order, path in enumerate(paths):
        df = pd.read_excel(path, dtype={c: str for c in IR_TEXT_COLUMNS})
        missing = [c for c in IR_REQUIRED if c not in df.columns]
        if missing:
            raise ValueError(f"{path.name}에 없는 열: {missing} — AX 화면에서 이 열을 숨기지 마세요")
        df['_file_order'] = order
        frames.append(df)
    return prepare_ir(pd.concat(frames, ignore_index=True))


def prepare_ir(ir: pd.DataFrame) -> pd.DataFrame:
    """형 맞추기 + 라인키 + KST 생성일. `_file_order`가 있으면 같은 라인은 나중 파일이 이긴다."""
    ir = ir.copy()
    if '_file_order' not in ir.columns:
        ir['_file_order'] = 0
    for c in IR_TEXT_COLUMNS:
        if c in ir.columns:
            ir[c] = ir[c].where(ir[c].isna(), ir[c].astype(str).str.strip())
    for c in ('Quantity', 'Unit price', 'Net amount', 'Deliver remainder'):
        ir[c] = pd.to_numeric(ir[c], errors='coerce')
    for c in ('Created date', 'Actual ship date', 'Ship date', 'Requested receipt date'):
        if c in ir.columns:
            ir[c] = pd.to_datetime(ir[c], errors='coerce')
    ir['line_key'] = ir['Project'] + '|' + ir['Line number'] + '|' + ir['Created date'].astype(str)
    ir = (ir.sort_values('_file_order', kind='stable')
            .drop_duplicates('line_key', keep='last')
            .drop(columns='_file_order')
            .reset_index(drop=True))
    ir['created_kst'] = ir['Created date'] + KST_OFFSET
    ir['ln'] = ir['Line number'].map(normalize_line_number)
    return ir


def read_it(paths: list[Path]) -> pd.DataFrame:
    """Item transaction export(여러 개면 합친다) → `prepare_it`"""
    frames = []
    for path in paths:
        df = pd.read_excel(path, dtype={c: str for c in IT_TEXT_COLUMNS})
        missing = [c for c in IT_REQUIRED if c not in df.columns]
        if missing:
            raise ValueError(f"{path.name}에 없는 열: {missing} — AX 화면에서 이 열을 숨기지 마세요")
        frames.append(df)
    return prepare_it(pd.concat(frames, ignore_index=True))


def prepare_it(it: pd.DataFrame) -> pd.DataFrame:
    """형 맞추기 + 라인번호 정규화. `Adjusted`(정정으로 대체된 행, 합계 0)는 뺀다.

    청구 상태는 행이 아니라 **라인 합계**로 봐야 한다 — 크레딧 노트·반품은 음수 Non-Invoiceable Posted 행으로 남아
    원래 Posted 행을 상쇄한다(82-12893.001: +2·12.1M과 −2·−12.1M, 사용자 확인 크레딧 노트).
    """
    it = it.copy()
    for c in IT_TEXT_COLUMNS:
        if c in it.columns:
            it[c] = it[c].where(it[c].isna(), it[c].astype(str).str.strip())
    it['Project date'] = pd.to_datetime(it['Project date'], errors='coerce')
    for c in ('Quantity', 'Net amount'):
        it[c] = pd.to_numeric(it[c], errors='coerce').fillna(0.0)
    it['Sales currency'] = it['Sales currency'].fillna('KRW').str.upper()
    it = it[it['Transaction status'] != IT_ADJUSTED].reset_index(drop=True)
    it['ln'] = it['Line number'].map(normalize_line_number)
    # 청구할 게 없는 무상 출고(Non-Invoiceable·0원)는 출고와 동시에 끝난 것으로 본다 — 안 그러면 수량이 영영 '청구 전'에 남는다
    free = it['Line property'].eq('Non-Invoiceable') & it['Net amount'].abs().lt(0.5)
    it['billed'] = it['Transaction status'].eq(IT_BILLED) | free
    it['pending'] = it['Transaction status'].eq(IT_PENDING) & ~free
    it['it_id'] = np.arange(len(it))
    return it


# ---------------------------------------------------------------------------
# 환율
# ---------------------------------------------------------------------------
def parse_closing_rates(path: Path) -> tuple[str, dict[str, float]]:
    """그룹 월 마감환율 파일 → (적용 월 'YYYY-MM', {통화: 원화 환율}).

    파일 기간 P{NN}은 그 달 22일 무렵 환율이고 AX는 **다음 달** 거래에 쓴다(P10 파일 = 11월, 사용자 확인).
    원화 환율 = (GBP당 KRW) ÷ (GBP당 통화) — 'Rate v GBP' 열의 4자리 값 그대로라 AX 값이 재현된다
    (P09 USD 1,883.3885 ÷ 1.3595 = 1,385.3538).
    """
    import openpyxl  # 이 함수만 쓰므로 여기서

    m = _CLOSING_FILE_RE.search(path.stem)
    if not m:
        raise ValueError(f"마감환율 파일 이름이 아닙니다: {path.name}")
    applies = pd.Period(f"{int(m.group(2))}-{int(m.group(1)):02d}", 'M') + 1
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        per_gbp: dict[str, float] = {}
        for row in wb.worksheets[0].iter_rows(values_only=True):
            if len(row) < 3 or not isinstance(row[1], str) or not isinstance(row[2], (int, float)):
                continue
            code = row[1].strip().upper()
            if re.fullmatch(r'[A-Z]{3}', code) and row[2]:
                per_gbp.setdefault(code, float(row[2]))
    finally:
        wb.close()
    if 'KRW' not in per_gbp:
        raise ValueError(f"{path.name}에서 KRW 환율을 못 찾았습니다")
    krw_per_gbp = per_gbp.pop('KRW')
    rates = {cur: krw_per_gbp / v for cur, v in per_gbp.items()}
    rates['GBP'] = krw_per_gbp
    return str(applies), rates


def _month_index(ym_text: str) -> int:
    y, m = ym_text.split('-')
    return int(y) * 12 + int(m)


@dataclass
class FxBook:
    """(통화, 'YYYY-MM') → 원화 환율. 없는 달은 가장 가까운 달로 대체하고 기록한다."""
    rates: dict[tuple[str, str], tuple[float, str]] = field(default_factory=dict)
    used: dict[tuple[str, str], tuple[float, str]] = field(default_factory=dict)
    fallbacks: dict[tuple[str, str], str] = field(default_factory=dict)

    def put(self, cur: str, month: str, rate: float, source: str) -> None:
        self.rates[(cur.upper(), month)] = (float(rate), source)

    def get(self, cur: str, month: str) -> float:
        if cur == 'KRW':
            return 1.0
        hit = self.rates.get((cur, month))
        if hit is None:
            same = [(abs(_month_index(k[1]) - _month_index(month)), k[1]) for k in self.rates if k[0] == cur]
            if not same:
                raise KeyError(f"{cur} 환율이 하나도 없습니다 — NOAH_SO_PO_DN FX 시트에 넣거나 --fx {cur}={month}=환율")
            near = min(same)[1]
            hit = self.rates[(cur, near)]
            self.fallbacks[(cur, month)] = near
        self.used[(cur, month)] = hit
        return hit[0]

    def rate_series(self, currencies: pd.Series, months: pd.Series) -> np.ndarray:
        """행마다 환율 — 같은 (통화, 월)은 한 번만 묻는다"""
        pairs = list(zip(currencies, months))
        cache = {p: self.get(*p) for p in set(pairs)}
        return np.array([cache[p] for p in pairs], dtype=float)


def _read_noah_fx(db_file: Path) -> pd.DataFrame:
    """NOAH DB의 fx 테이블(= NOAH_SO_PO_DN FX 시트, USD/EUR/GBP 월환율 = AX 월환율). OneDrive 원본은 읽기 전용 사본으로 연다."""
    if not db_file.exists():
        return pd.DataFrame(columns=['currency', 'ym', 'rate'])
    src = sqlite3.connect(f"file:{db_file.as_posix()}?mode=ro", uri=True)
    mem = sqlite3.connect(':memory:')
    try:
        src.backup(mem)
        return pd.read_sql('SELECT currency, ym, rate FROM fx', mem)
    except sqlite3.Error as exc:
        logger.warning("NOAH DB 환율을 못 읽었습니다: %s", exc)
        return pd.DataFrame(columns=['currency', 'ym', 'rate'])
    finally:
        src.close()
        mem.close()


def load_fx_book(db_file: Path | None, history_csv: Path | None, hub_dir: Path | Sequence[Path] | None,
                 overrides: dict[tuple[str, str], float] | None = None) -> FxBook:
    """우선순위(높은 쪽이 이긴다): 직접 입력 > 마감환율 파일·이력 > NOAH FX 시트 > 역산 시드.

    hub_dir = 마감환율 파일을 찾을 폴더 — Finance Hub, 그리고 사용자가 사본을 넣어 두는 그 달 폴더(여럿이면 리스트).
    Finance Hub 폴더엔 현재 기간 파일만 남으므로, 읽은 값을 이력 CSV에 쌓아 다음 달에도 쓰게 한다. 직접 입력(--fx)도
    이력에 남긴다 — 저장한 달의 End는 라인을 수주월 환율로 매긴 값이라, 다음 달에 그 환율이 빠지면 같은 라인이 다른
    환율로 다시 매겨져 가짜 Variance가 생긴다(10월 JPY를 --fx로만 넣었더니 11월에 82-15517.001이 +1.1M).
    지난 해 환율(2025년 생성 외화 라인의 수주월)은 NOAH FX 시트에 월 열을 더해 `sync_db.py`로 넣는다.
    """
    book = FxBook()
    for (cur, month), rate in FX_SEED.items():
        book.put(cur, month, rate, '공식 오더북 역산')
    if db_file is not None:
        for r in _read_noah_fx(db_file).itertuples():
            book.put(r.currency, r.ym, r.rate, 'NOAH FX 시트')

    history = pd.DataFrame(columns=['ym', 'currency', 'rate', 'source'])
    if history_csv is not None and history_csv.exists():
        history = pd.read_csv(history_csv, dtype={'ym': str, 'currency': str, 'source': str})
    changed = False
    hub_dirs = [hub_dir] if isinstance(hub_dir, Path) else list(hub_dir or [])
    closing_files = {p for d in hub_dirs if d is not None and d.is_dir()
                     for p in d.glob('FX Rates - Closing rates - P* *.xlsx') if not p.name.startswith('~$')}
    for path in sorted(closing_files, key=lambda p: p.name):      # 같은 파일이 두 폴더에 있어도 그 달은 한 번만 쌓인다
        try:
            applies, rates = parse_closing_rates(path)
        except (ValueError, OSError) as exc:
            logger.warning("마감환율 파일을 못 읽었습니다: %s (%s)", path.name, exc)
            continue
        if (history['ym'].eq(applies) & history['source'].ne(MANUAL_FX_SOURCE)).any():
            continue
        new = pd.DataFrame([{'ym': applies, 'currency': c, 'rate': r, 'source': path.name}
                            for c, r in sorted(rates.items())])
        history = new if history.empty else pd.concat([history, new], ignore_index=True)
        changed = True
        logger.info("마감환율 이력에 %s 환율을 추가했습니다 (%s)", applies, path.name)
    if overrides:
        manual = pd.DataFrame([{'ym': m, 'currency': c, 'rate': r, 'source': MANUAL_FX_SOURCE}
                               for (c, m), r in sorted(overrides.items())])
        replaced = ((history['ym'] + '|' + history['currency']).isin(manual['ym'] + '|' + manual['currency'])
                    & history['source'].eq(MANUAL_FX_SOURCE))
        history = manual if history.empty else pd.concat([history[~replaced], manual], ignore_index=True)
        changed = True
    if changed and history_csv is not None:
        history_csv.parent.mkdir(parents=True, exist_ok=True)
        history.to_csv(history_csv, index=False, encoding='utf-8-sig')
    manual_rows = history['source'].eq(MANUAL_FX_SOURCE)          # 직접 입력은 이력 안에서도 마지막(우선)
    for part in (history[~manual_rows], history[manual_rows]):
        for r in part.itertuples():
            book.put(r.currency, r.ym, r.rate, r.source)
    return book


# ---------------------------------------------------------------------------
# 매핑 (회계 오더북에서 배운다)
# ---------------------------------------------------------------------------
def learn_dept_map(history: pd.DataFrame, ir: pd.DataFrame) -> dict[str, str]:
    """Prod/Dept(Material) → (Revenue). 산술식 하나가 아니다(1034→0134, 1172→0272, 3000→0108) —
    공식 오더북과 IR 양쪽에서 부서가 하나뿐인 프로젝트끼리 짝지어 가장 많이 나온 짝을 쓴다."""
    rev = history.groupby('Project')['Prod/Dept (Revenue)'].agg(lambda s: frozenset(s.dropna()))
    live = ir[ir['Line status'] != 'Cancelled']
    mat = live.groupby('Project')['Prod/Dept (Material)'].agg(lambda s: frozenset(s.dropna()))
    both = pd.concat([rev.rename('rev'), mat.rename('mat')], axis=1, join='inner')
    single = both[(both['rev'].map(len) == 1) & (both['mat'].map(len) == 1)]
    pairs = Counter((next(iter(m)), next(iter(r))) for m, r in zip(single['mat'], single['rev']))
    mapping: dict[str, str] = {}
    for (m, r), _ in pairs.most_common():
        mapping.setdefault(m, r)
    return mapping


def dept_by_rule(material: str | None) -> str | None:
    """학습된 짝이 없을 때 — 1000~1799는 −900, 1800번대는 −1000.
    P01~P09에서 배운 45쌍 중 42쌍이 이 규칙이다(예외 1000→0020, 3000·3010→0108은 학습 쪽이 이긴다)."""
    try:
        v = int(str(material))
    except (TypeError, ValueError):
        return None
    if 1800 <= v < 1900:
        return f"{v - 1000:04d}"
    if 1000 < v < 1800:
        return f"{v - 900:04d}"
    return None


def _latest_by(history: pd.DataFrame, key_cols: list[str], value_col: str) -> dict:
    h = history.sort_values('pn').dropna(subset=[value_col])
    h = h.drop_duplicates(key_cols, keep='last')
    if len(key_cols) == 1:
        return dict(zip(h[key_cols[0]], h[value_col]))
    return dict(zip(map(tuple, h[key_cols].to_numpy()), h[value_col]))


def read_currency_overrides(path: Path) -> dict[str, str]:
    """통화지정.csv — `코드,통화` (코드 = 거래처 C번호 또는 프로젝트 번호). 없으면 빈 dict.

    통화는 주문 단위로 갈리기도 한다 — 한국수자원공사는 원화 거래처인데 82-17677.001만 EUR였다(P10 실측).
    """
    if not path.exists():
        return {}
    df = pd.read_csv(path, dtype=str, encoding='utf-8-sig').dropna(how='all')
    df.columns = [c.strip() for c in df.columns]
    code_col, cur_col = df.columns[:2]
    return {str(c).strip(): str(v).strip().upper() for c, v in zip(df[code_col], df[cur_col])
            if isinstance(c, str) and isinstance(v, str) and c.strip()}


def assign_currency(customers: pd.Series, projects: pd.Series, overrides: dict[str, str] | None = None,
                    customer_map: dict[str, str] | None = None) -> pd.Series:
    """프로젝트 지정 > 거래처 지정(통화지정.csv) > 외화 거래처 목록 > KRW"""
    overrides = overrides or {}
    customer_map = {**(FOREIGN_CUSTOMER_CURRENCY if customer_map is None else customer_map), **overrides}
    by_project = projects.map(overrides)
    return by_project.fillna(customers.map(customer_map)).fillna('KRW')


def line_frame(ir: pd.DataFrame, official: pd.DataFrame, currency_overrides: dict[str, str] | None = None
               ) -> pd.DataFrame:
    """계산에 쓰는 IR 라인 — Cancelled 제외, 키·통화·단가·수주월을 붙인다.

    키로 거르지 않는다: 우리 매핑이 공식과 다른 키를 만들어도 라인은 남고 '부서·CC 귀속 차이'로 보인다
    (예전엔 공식 Start에 없는 키를 대상에서 뺐다가 P10에서 크레인·MOD6N2 등 라인 5개 9.4M이 조용히 빠졌다).
    """
    dept_map = learn_dept_map(official, ir)
    cc_map = _latest_by(official, ['Project', 'Prod/Dept (Revenue)'], 'Cost center')
    lines = ir[ir['Line status'] != 'Cancelled'].copy()
    lines['cur'] = assign_currency(lines['Customer'], lines['Project'], currency_overrides)
    lines['RevDept'] = lines['Prod/Dept (Material)'].map(dept_map)
    unmapped = lines['RevDept'].isna()
    lines.loc[unmapped, 'RevDept'] = lines.loc[unmapped, 'Prod/Dept (Material)'].map(dept_by_rule)
    lines['dept_by_rule'] = unmapped & lines['RevDept'].notna()
    lines['CC'] = [cc_map.get((p, d), dv) for p, d, dv in
                   zip(lines['Project'], lines['RevDept'], lines['Division'])]
    lines['Key'] = lines['Project'] + lines['CC'].fillna('') + lines['RevDept'].fillna('????')
    lines['unit'] = lines['Net amount'] / lines['Quantity'].where(lines['Quantity'] != 0)
    lines['cre_ym'] = lines['created_kst'].dt.strftime('%Y-%m')
    return lines.reset_index(drop=True)


def link_shipments(lines: pd.DataFrame, it: pd.DataFrame, official: pd.DataFrame | None = None) -> pd.DataFrame:
    """IT 행 → IR 라인. (Project, 라인번호)가 같으면 **출고일 이전에 생긴 가장 최근 라인**에 붙인다 — 라인번호는
    재사용된다(2026-11-02 실측 1,902쌍). 그런 라인이 없으면 가장 가까운 뒤 라인, 그래도 없으면 프로젝트 주 키에 붙이고
    `linked=False`로 둔다(2024년 이전 생성 라인, 라인번호 0 등)."""
    cols = ['Project', 'ln', 'line_key', 'Key', 'cur', 'cre_ym', 'created_kst']
    lk = lines[cols].assign(cre_day=lines['created_kst'].dt.normalize(), ln=lines['ln'].fillna('?'))
    lk = lk.sort_values('cre_day')
    t = it[it['Project date'].notna()].assign(ln=it['ln'].fillna('?')).sort_values('Project date')
    back = pd.merge_asof(t, lk, left_on='Project date', right_on='cre_day', by=['Project', 'ln'],
                         direction='backward')
    miss = back['line_key'].isna().to_numpy()
    if miss.any():
        fwd = pd.merge_asof(t[miss], lk, left_on='Project date', right_on='cre_day', by=['Project', 'ln'],
                            direction='forward')
        back = pd.concat([back[~miss], fwd], ignore_index=True)
    back = _relink_reused(back, lines)
    back['linked'] = back['line_key'].notna()
    # 라인 없는 출고: 그 프로젝트에서 가장 흔한 키 → 없으면 공식 오더북의 키
    proj_key = lines.groupby('Project')['Key'].agg(lambda s: s.mode().iloc[0])
    if official is not None and len(official):
        off_key = official.sort_values('pn').drop_duplicates('Project', keep='last').set_index('Project')['Key']
        proj_key = proj_key.combine_first(off_key)
    back['Key'] = back['Key'].fillna(back['Project'].map(proj_key))
    back['cre_ym'] = back['cre_ym'].fillna(back['Project date'].dt.strftime('%Y-%m'))
    back['group'] = back['line_key'].fillna('~' + back['Project'] + '|' + back['ln'].fillna('?'))
    # 청구 누계 묶음 — 추출본이 바뀌어도 같은 출고를 같은 묶음으로 보게 IR 라인키 대신 원천 값으로 만든다.
    # 라인번호가 재사용돼도 품목이 다르면 갈린다(82-15514.001: 108번이 압력계와 포지셔너)
    back['bgroup'] = (back['Project'] + '|' + back['ln'].fillna('?') + '|'
                      + back['Item number'].fillna('?').astype(str))
    return back.drop(columns=['cre_day']).sort_values('it_id').reset_index(drop=True)


def _relink_reused(linked: pd.DataFrame, lines: pd.DataFrame) -> pd.DataFrame:
    """같은 (Project, 라인번호)에 IR 라인이 여럿이면 날짜만으로는 틀린다 — 1월 라인이 3월에 나가는데 2월에 같은 번호로
    다른 품목 라인이 생기면, '출고일 이전 가장 최근 라인'은 2월 라인을 고른다(82-15514.001 등에서 실측). 그래서
    **단가가 맞는 라인**을 먼저 고르고(±0.5%), 그중 출고일 이전 가장 최근, 없으면 가장 이른 라인. 단가까지 같은 라인끼리
    섞이는 건 키·금액이 같아 숫자에 영향이 없다."""
    multi = lines.groupby(['Project', 'ln']).size()
    multi = multi[multi > 1]
    if multi.empty:
        return linked
    cand = lines[lines.set_index(['Project', 'ln']).index.isin(multi.index)]
    by_group = {k: g.sort_values('created_kst') for k, g in cand.groupby(['Project', 'ln'])}
    mask = linked.set_index(['Project', 'ln']).index.isin(multi.index)
    fields = ['line_key', 'Key', 'cur', 'cre_ym', 'created_kst']
    linked = linked.copy()
    for i in np.flatnonzero(mask):
        r = linked.iloc[i]
        g = by_group[(r['Project'], r['ln'])]
        pool = g
        if abs(r['Quantity']) > _EPS and abs(r['Net amount']) > 0.5:
            unit = r['Net amount'] / r['Quantity']
            ok = (g['unit'] - unit).abs() <= g['unit'].abs() * 0.005 + 0.5
            if ok.any():
                pool = g[ok]
        before = pool[pool['created_kst'].dt.normalize() <= r['Project date']]
        pick = before.iloc[-1] if len(before) else pool.iloc[0]
        for f in fields:
            linked.iat[i, linked.columns.get_loc(f)] = pick[f]
    return linked


# ---------------------------------------------------------------------------
# 계산
# ---------------------------------------------------------------------------
BILLED_COLUMNS: tuple[str, ...] = ('Project', 'Key', 'cur', 'qty', 'amt')


@dataclass
class OrderbookResult:
    year: int
    upto: int
    raw_rows: pd.DataFrame          # Raw 시트 행 (계산한 달들, 48열 + _key·pn)
    keys: pd.DataFrame              # 달·키별 계산 (pn, Key, Project, Sq..Vv)
    lines: pd.DataFrame             # 마감월(upto) 라인별 기여
    pending: pd.DataFrame           # 마감월 말 출고 후 미청구 (라인별)
    review: pd.DataFrame            # 확인 목록
    fx_used: pd.DataFrame
    next_filter_note: str
    key_currency: dict[str, str] = field(default_factory=dict)   # 키 → 통화 (차이 원인 '환율' 판단용)
    billed: pd.DataFrame = field(default_factory=pd.DataFrame)   # 계산한 달마다 그달 말까지 인식한 청구 누계 (pn 열)

    @property
    def period(self) -> str:
        return f"P{self.upto:02d}"


def _is_pending_balance(qty: pd.Series, net: pd.Series) -> pd.Series:
    """라인 합계로 '청구 전'이 남았는가 — 실물(수량)과 금액이 둘 다 남아야 한다. 크레딧 노트는 수량·금액이 함께 0으로
    상쇄되고(82-12893.001), 수량만 상쇄되고 금액이 남는 이상한 쌍(82-14677.001, 176k)은 청구 대기가 아니다."""
    return qty.gt(_EPS) & net.ge(0.5)


def empty_billed() -> pd.DataFrame:
    return pd.DataFrame({c: pd.Series(dtype=float if c in ('qty', 'amt') else object) for c in BILLED_COLUMNS},
                        index=pd.Index([], name='bgroup'))


def billed_cumulative(ship: pd.DataFrame, pend: pd.Timestamp) -> pd.DataFrame:
    """pend 전까지 출고된 IT 행 중 청구된 것의 누계 — 묶음(프로젝트|라인번호|품목번호)별 수량·금액(거래통화).

    Output은 이 누계가 '이미 인식한 누계'보다 늘어난 만큼이다. 같은 추출본의 지난 달 말 누계와 빼면 그달 출고분이 되고
    (기준선), 저장된 지난 마감 때 누계와 빼면 그 사이에 청구된 것 전부가 된다 — 지난 마감 때 청구 전이던 출고가 이번에
    청구되면 이번 달 Output이다(AX Sales Order Book처럼 청구한 달)."""
    rows = ship[ship['billed'] & ship['Project date'].lt(pend)]
    if rows.empty:
        return empty_billed()
    return rows.groupby('bgroup').agg(Project=('Project', 'first'), Key=('Key', 'first'),
                                      cur=('Sales currency', 'first'), qty=('Quantity', 'sum'),
                                      amt=('Net amount', 'sum'))


def _output_delta(now: pd.DataFrame, prev: pd.DataFrame, fx: FxBook, close_ym: str) -> pd.DataFrame:
    """Output = 청구 누계 − 이미 인식한 누계(묶음별), 그달 환율. 누계가 줄면(청구 뒤 크레딧) 음수 Output."""
    d = now[['qty', 'amt']].sub(prev[['qty', 'amt']], fill_value=0.0)
    d = d.join(now[['Project', 'Key', 'cur']].combine_first(prev[['Project', 'Key', 'cur']]))
    d = d[d['qty'].abs().gt(_EPS) | d['amt'].abs().ge(0.5)]
    if d.empty:
        return pd.DataFrame(columns=['group', 'Key', 'Project', 'out_q', 'out_v'])
    return pd.DataFrame({'group': 'B|' + d.index.astype(str), 'Key': d['Key'].to_numpy(),
                         'Project': d['Project'].to_numpy(), 'out_q': d['qty'].to_numpy(),
                         'out_v': d['amt'].to_numpy() * fx.rate_series(d['cur'], pd.Series([close_ym] * len(d)))})


def _month_contrib(lines: pd.DataFrame, ship: pd.DataFrame, year: int, month: int, fx: FxBook,
                   prev_billed: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """월 m의 그룹별 Input·Output·End + 그달 말 청구 누계. 그룹 = IR 라인키(라인 없는 출고는 '~프로젝트|번호'),
    Output은 청구 누계 묶음('B|프로젝트|번호|품목')."""
    pstart, pend = month_bounds(year, month)
    close_ym = ym(pstart)
    live = lines[lines['created_kst'] < pend]
    shipped = ship[ship['Project date'] < pend]

    # End ① 아직 안 나간 수량 — 라인 수량 − 월말까지 IT 출고 수량. IR이 '출고 끝'(Invoiced)이라 하고 마지막 출고일이
    # 월말 전이면 0 — 재사용된 라인번호로 출고가 다른 라인에 붙어도 유령 잔량이 남지 않게
    shipped_q = shipped.groupby('group')['Quantity'].sum()
    open_q = (live['Quantity'].fillna(0.0) - live['line_key'].map(shipped_q).fillna(0.0)).clip(lower=0.0)
    done = live['Line status'].eq('Invoiced') & live['Actual ship date'].lt(pend)
    open_q = open_q.where(~done, 0.0)
    has_open = open_q.gt(_EPS) & live['unit'].notna()
    open_v = np.zeros(len(live))
    if has_open.any():
        sub = live[has_open]
        open_v[has_open.to_numpy()] = (open_q[has_open] * sub['unit']).to_numpy() * fx.rate_series(sub['cur'], sub['cre_ym'])
    frame = pd.DataFrame({'group': live['line_key'], 'Key': live['Key'], 'Project': live['Project'],
                          'end_q': open_q.to_numpy(), 'end_v': open_v})

    # Input — 그달 생성된 라인
    cin = live['created_kst'].ge(pstart)
    in_v = np.zeros(len(live))
    if cin.any():
        sub = live[cin]
        in_v[cin.to_numpy()] = sub['Net amount'].fillna(0.0).to_numpy() * fx.rate_series(sub['cur'], sub['cre_ym'])
    frame['in_q'] = np.where(cin, live['Quantity'].fillna(0.0), 0.0)
    frame['in_v'] = in_v

    # End ② 출고했지만 청구 전 — 라인 합계로(크레딧·반품 상쇄), 수주월 환율
    pend_rows = shipped[shipped['pending']]
    if len(pend_rows):
        g = pend_rows.groupby('group').agg(q=('Quantity', 'sum'), net=('Net amount', 'sum'), Key=('Key', 'first'),
                                           Project=('Project', 'first'), cur=('Sales currency', 'first'),
                                           cre_ym=('cre_ym', 'first'))
        g = g[_is_pending_balance(g['q'], g['net'])]
        if len(g):
            g['v'] = g['net'].clip(lower=0.0) * fx.rate_series(g['cur'], g['cre_ym'])
            frame = pd.concat([frame, pd.DataFrame({'group': g.index, 'Key': g['Key'].to_numpy(),
                                                    'Project': g['Project'].to_numpy(), 'end_q': g['q'].to_numpy(),
                                                    'end_v': g['v'].to_numpy(), 'in_q': 0.0, 'in_v': 0.0})],
                              ignore_index=True)

    # Output — 그달 말 청구 누계 − 이미 인식한 누계, 그달 환율
    billed_now = billed_cumulative(ship, pend)
    out = _output_delta(billed_now, prev_billed, fx, close_ym)
    if len(out):
        frame = pd.concat([frame, out], ignore_index=True)
    for c in ('in_q', 'in_v', 'out_q', 'out_v', 'end_q', 'end_v'):
        if c not in frame.columns:
            frame[c] = 0.0
        frame[c] = frame[c].astype(float).fillna(0.0)
    contrib = frame.groupby(['group', 'Key', 'Project'], as_index=False)[
        ['in_q', 'in_v', 'out_q', 'out_v', 'end_q', 'end_v']].sum()
    return contrib, billed_now


def _key_month(start: pd.DataFrame, contrib: pd.DataFrame, month: int) -> pd.DataFrame:
    """키별 Start(q, v) + 그달 기여 → 그달 키별 계산 (Variance = 차액)"""
    agg = contrib.groupby('Key')[['in_q', 'in_v', 'out_q', 'out_v', 'end_q', 'end_v']].sum()
    k = pd.DataFrame(index=start.index.union(agg.index))
    k.index.name = 'Key'
    k['Sq'] = start['q'].reindex(k.index).fillna(0.0)
    k['Sv'] = start['v'].reindex(k.index).fillna(0.0)
    k = k.join(agg).fillna(0.0).rename(columns={'in_q': 'Iq', 'in_v': 'Iv', 'out_q': 'Oq', 'out_v': 'Ov',
                                                'end_q': 'Eq', 'end_v': 'Ev'})
    k['Vq'] = k['Eq'] - k['Sq'] - k['Iq'] + k['Oq']
    k['Vv'] = k['Ev'] - k['Sv'] - k['Iv'] + k['Ov']
    # 금액이 하나도 없는 키(무상 라인 수량만 남은 옛 프로젝트)는 뺀다 — 회계 Raw에도 8979행 중 8행뿐이다
    k = k[k[['Sv', 'Iv', 'Ov', 'Ev', 'Vv']].abs().ge(0.5).any(axis=1)].copy()
    k['pn'] = month
    return k.reset_index()


def _prepare(ir: pd.DataFrame, it: pd.DataFrame, official: pd.DataFrame,
             overrides: dict[str, str]) -> tuple[pd.DataFrame, pd.DataFrame]:
    lines = line_frame(ir, official, overrides)
    return lines, link_shipments(lines, it, official)


def build_orderbook(official: pd.DataFrame, ir: pd.DataFrame, it: pd.DataFrame, year: int, upto: int,
                    fx: FxBook, currency_overrides: dict[str, str] | None = None) -> OrderbookResult:
    """기준선 — IR·IT 한 벌로 P01~upto를 한 번에. official = 회계 오더북 Raw 행(P01 필수 — 기초값, 키 매핑·속성 학습용).
    달마다 '이미 인식한 청구 누계'를 같은 추출본의 지난 달 말 누계로 두므로 Output은 그달 출고분이다."""
    p01 = official[official['pn'] == 1]
    if p01.empty:
        raise ValueError("회계 오더북에 P01 행이 없습니다 — 기초(P01 Start)를 못 잡습니다")
    overrides = currency_overrides or {}
    lines, ship = _prepare(ir, it, official, overrides)

    start = p01.groupby('Key')[['Qty Start', 'Value Start']].sum()
    start.columns = ['q', 'v']
    prev_billed = billed_cumulative(ship, pd.Timestamp(year, 1, 1))       # 올해 전에 이미 인식된 청구
    months, billed, last = [], [], None
    for m in range(1, upto + 1):
        last, prev_billed = _month_contrib(lines, ship, year, m, fx, prev_billed)
        k = _key_month(start, last, m)
        months.append(k)
        billed.append(prev_billed.reset_index().assign(pn=m))
        start = k.set_index('Key')[['Eq', 'Ev']].rename(columns={'Eq': 'q', 'Ev': 'v'})
    return _finish(pd.concat(months, ignore_index=True), last, lines, ship, ir, fx, year, upto, official,
                   attribute_history(official, None, year), overrides, pd.concat(billed, ignore_index=True))


def build_month(official: pd.DataFrame, ir: pd.DataFrame, it: pd.DataFrame, year: int, month: int, fx: FxBook,
                prev_rows: pd.DataFrame, prev_billed: pd.DataFrame, history: pd.DataFrame | None = None,
                currency_overrides: dict[str, str] | None = None) -> OrderbookResult:
    """다음 달 하나 — AX Sales Order Book처럼 **저장된 직전 달**에서 시작한다.

    Start = 저장된 직전 End, '이미 인식한 청구 누계' = 저장된 직전 달 누계. 그래서 직전 계산 뒤의 취소·단가/수량 변경은
    이번 달 Variance로, 지난 마감 때 청구 전이던 출고가 이번에 청구되면 이번 달 Output으로 잡히고 지난 달은 그대로다.
    history = 저장된 Raw 행(속성 물려받기용 — 계산할 달보다 앞의 것), official = 회계 오더북(키 매핑·속성 — 지난 해
    파일이어도 되고, 없으면 `empty_official()`)."""
    overrides = currency_overrides or {}
    lines, ship = _prepare(ir, it, official, overrides)
    start = prev_rows.groupby('_key')[['Qty End', 'Value End']].sum()
    start.columns = ['q', 'v']
    start.index.name = 'Key'
    last, billed_now = _month_contrib(lines, ship, year, month, fx, prev_billed)
    keys = _key_month(start, last, month)
    if history is not None and len(history):        # 다시 계산할 때 그 달의 옛 계산이 속성 이력에 끼지 않게
        history = history[history['year'] * 100 + history['pn'] < year * 100 + month]
    attr_history = attribute_history(official, history, year)
    return _finish(keys, last, lines, ship, ir, fx, year, month, official, attr_history, overrides,
                   billed_now.reset_index().assign(pn=month))


def empty_official() -> pd.DataFrame:
    """회계 오더북이 없을 때 — 부서 매핑은 규칙으로, 속성은 저장된 행에서만"""
    return pd.DataFrame(columns=[*RAW_COLUMNS, 'pn', 'year'])


def attribute_history(official: pd.DataFrame, stored: pd.DataFrame | None, year: int) -> pd.DataFrame:
    """속성을 물려받을 이력 — 저장된 우리 Raw 행 + 회계 Raw 행, 기간 순. 같은 기간이면 회계 행이 이긴다.
    'pn'을 연·월 순번(연×100+월)으로 바꿔 둔다 — 아래 함수들이 'pn' 순으로 마지막 값을 고르기 때문. 회계 행의 연도는
    파일의 연도('year' 열 — read_raw), 없으면 계산하는 해."""
    parts = []
    if stored is not None and len(stored):
        s = stored.copy()
        s['Key'] = s['_key']
        s['_ord'] = s['year'] * 100 + s['pn']
        s['_src'] = 0
        parts.append(s)
    if official is not None and len(official):
        o = official.copy()
        o_year = pd.to_numeric(o['year'], errors='coerce').fillna(year).astype(int) if 'year' in o.columns else year
        o['_ord'] = o_year * 100 + o['pn']
        o['_src'] = 1
        parts.append(o)
    if not parts:
        return pd.DataFrame(columns=list(RAW_COLUMNS) + ['pn'])
    # 빈 열(수식 열 등)은 빼고 합친 뒤 되살린다 — 빈 열이 열 타입을 정하지 않게(pandas가 바꾸려는 동작을 미리 고정)
    columns = list(dict.fromkeys(c for p in parts for c in p.columns))
    h = (pd.concat([p.dropna(axis=1, how='all') for p in parts], ignore_index=True)
           .reindex(columns=columns).sort_values(['_ord', '_src'], kind='stable'))
    h['pn'] = h['_ord']
    return h.drop(columns=['_ord', '_src']).reset_index(drop=True)


def keys_from_raw(raw_rows: pd.DataFrame) -> pd.DataFrame:
    """Raw 행 → 달·키별 계산 표(pn, Key, Project, Sq..Vv) — 저장된 달까지 합쳐 요약·대조할 때"""
    names = {'Qty Start': 'Sq', 'Qty Input': 'Iq', 'Qty Output': 'Oq', 'Qty End': 'Eq', 'Qty Variance': 'Vq',
             'Value Start': 'Sv', 'Value Input': 'Iv', 'Value Output': 'Ov', 'Value End': 'Ev',
             'Value Variance': 'Vv'}
    k = raw_rows.rename(columns=names)
    k['Key'] = raw_rows['_key']
    return k[['pn', 'Key', 'Project'] + list(_KEY_VALUES)].reset_index(drop=True)


def _finish(keys: pd.DataFrame, last: pd.DataFrame, lines: pd.DataFrame, ship: pd.DataFrame, ir: pd.DataFrame,
            fx: FxBook, year: int, upto: int, official: pd.DataFrame, attr_history: pd.DataFrame,
            overrides: dict[str, str], billed: pd.DataFrame) -> OrderbookResult:
    project_of = {**dict(zip(attr_history['Key'], attr_history['Project'])),
                  **dict(zip(lines['Key'], lines['Project'])), **dict(zip(ship['Key'], ship['Project']))}
    keys['Project'] = keys['Key'].map(project_of).fillna('')
    # 신규 = 회계 오더북에 없는 키 — 지난 달에 저장했어도 회계가 속성을 채우기 전까지는 확인 목록에 남는다
    keys['new_key'] = ~keys['Key'].isin(set(official['Key']))

    raw_rows = _assemble_raw(keys, lines, attr_history, year)
    pending = _pending_list(ship, lines, year, upto, fx)
    line_detail = last.merge(lines[['line_key', 'Line number', 'Item number', 'Item name', 'Customer', 'cur',
                                    'Line status', 'created_kst', 'Actual ship date', 'Quantity',
                                    'Deliver remainder', 'Net amount']],
                             left_on='group', right_on='line_key', how='left')
    out_rows = line_detail['group'].str.startswith('B|')
    if out_rows.any():
        parts = line_detail.loc[out_rows, 'group'].str.split('|', n=3, expand=True)
        line_detail.loc[out_rows, 'Line number'] = parts[2]
        line_detail.loc[out_rows, 'Item number'] = parts[3]
    line_detail = line_detail[line_detail[['in_v', 'out_v', 'end_v', 'in_q', 'out_q', 'end_q']].abs().gt(_EPS)
                              .any(axis=1)].drop(columns='line_key').reset_index(drop=True)
    review = _review(keys, lines, ship, raw_rows, ir, fx, {**FOREIGN_CUSTOMER_CURRENCY, **overrides},
                     year, upto, attr_history)
    fx_used = pd.DataFrame([{'통화': c, '월': mo, '환율': r, '출처': s, '대체한 달': fx.fallbacks.get((c, mo), '')}
                            for (c, mo), (r, s) in sorted(fx.used.items())])
    note = _next_filter_note(ir, raw_rows[raw_rows['pn'] == upto])
    key_currency = lines.groupby('Key')['cur'].agg(lambda s: s.mode().iloc[0]).to_dict()
    return OrderbookResult(year, upto, raw_rows, keys, line_detail, pending, review, fx_used, note, key_currency,
                           billed)


def _pending_list(ship: pd.DataFrame, lines: pd.DataFrame, year: int, upto: int, fx: FxBook) -> pd.DataFrame:
    """마감월 말 기준 출고 후 미청구 — 라인 합계가 남는 것만(크레딧·반품 상쇄 뒤)."""
    _, pend = month_bounds(year, upto)
    rows = ship[ship['pending'] & ship['Project date'].lt(pend)]
    if rows.empty:
        return pd.DataFrame(columns=['Project', 'Key', 'Line number', 'Item name', '첫 출고일', '마지막 출고일',
                                     '경과 개월', '수량', '통화', '금액(통화)', '금액(원)'])
    g = rows.groupby('group').agg(Project=('Project', 'first'), Key=('Key', 'first'), ln=('ln', 'first'),
                                  item=('Item name', 'first'), first=('Project date', 'min'),
                                  last=('Project date', 'max'), q=('Quantity', 'sum'), net=('Net amount', 'sum'),
                                  cur=('Sales currency', 'first'), cre_ym=('cre_ym', 'first'))
    g = g[_is_pending_balance(g['q'], g['net'])].copy()
    g['krw'] = g['net'] * fx.rate_series(g['cur'], g['cre_ym'])
    months = ((pend.year - g['first'].dt.year) * 12 + (pend.month - g['first'].dt.month) - 1).clip(lower=0)
    out = pd.DataFrame({'Project': g['Project'], 'Key': g['Key'], 'Line number': g['ln'], 'Item name': g['item'],
                        '첫 출고일': g['first'], '마지막 출고일': g['last'], '경과 개월': months, '수량': g['q'],
                        '통화': g['cur'], '금액(통화)': g['net'], '금액(원)': g['krw']})
    return out.sort_values(['첫 출고일', 'Project']).reset_index(drop=True)


def _mode(s: pd.Series):
    s = s.dropna()
    return s.mode().iloc[0] if len(s) else None


def _assemble_raw(keys: pd.DataFrame, lines: pd.DataFrame, history: pd.DataFrame, year: int) -> pd.DataFrame:
    """달·키별 숫자 + 속성 → Raw 48열. 속성은 키마다 한 번 정한다: 이력(회계 오더북·저장된 우리 행)에 있는 키는
    마지막 행 그대로, 새 키는 같은 프로젝트의 프로젝트 단위 속성 + IR."""
    hist = history.sort_values('pn', kind='stable')
    key_attr = hist.drop_duplicates('Key', keep='last').set_index('Key')
    proj_attr = hist.drop_duplicates('Project', keep='last').set_index('Project')
    name_by_account = _latest_by(history, ['Account number'], 'Customer Name')
    ce_code_by_name = _latest_by(history, ['Contract Engineer Name'], 'Contract Engineer')
    by_proj = lines.groupby('Project')
    proj_created = by_proj['created_kst'].min().dt.normalize()
    proj_receipt = by_proj['Ship date'].max() if 'Ship date' in lines.columns else pd.Series(dtype='datetime64[ns]')
    proj_req = (by_proj['Requested receipt date'].max() if 'Requested receipt date' in lines.columns
                else pd.Series(dtype='datetime64[ns]'))
    lines_of = dict(tuple(lines.groupby('Key')))

    attrs: dict[str, dict] = {}
    for key, project in keys.drop_duplicates('Key')[['Key', 'Project']].itertuples(index=False):
        row = {c: None for c in RAW_COLUMNS}
        if key in key_attr.index:
            src = key_attr.loc[key]
            for c in CARRY_COLUMNS:
                row[c] = src[c]
        else:
            g = lines_of.get(key, lines.iloc[0:0])
            if project in proj_attr.index:          # 기존 프로젝트의 새 부서 — 프로젝트 속성은 빌려 온다
                for c in PROJECT_LEVEL_COLUMNS:
                    row[c] = proj_attr.loc[project, c]
            account = row['Account number'] or _mode(g['Customer'])
            row['Account number'] = account
            row['Customer Name'] = row['Customer Name'] or name_by_account.get(account) or _mode(g['Customer name'])
            row['Project group'] = row['Project group'] or _mode(g['Project group'])
            row['Contract Engineer'] = row['Contract Engineer'] or ce_code_by_name.get(_mode(g['Contract engineer']))
            row['Cost center'] = _mode(g['CC'])
            row['Prod/Dept (Revenue)'] = _mode(g['RevDept'])
            row['Project creation date'] = proj_created.get(project)
            row['Receipt date'] = proj_receipt.get(project)
            row['Requested ship date'] = proj_req.get(project)
            row['Location'] = '4080'
        row['Project'] = project
        attrs[key] = row

    rows = []
    for k in keys.itertuples(index=False):
        row = dict(attrs[k.Key])
        row['Report Start'] = pd.Timestamp(year, int(k.pn), 1)
        row['Period'] = f"P{int(k.pn):02d}"
        row['Period status'] = 'Open'
        for c, src_col in (('Qty Start', 'Sq'), ('Qty Input', 'Iq'), ('Qty Output', 'Oq'), ('Qty End', 'Eq'),
                           ('Qty Variance', 'Vq'), ('Value Start', 'Sv'), ('Value Input', 'Iv'),
                           ('Value Output', 'Ov'), ('Value End', 'Ev'), ('Value Variance', 'Vv')):
            row[c] = round(float(getattr(k, src_col)), 6)
        row['_key'] = k.Key
        row['pn'] = int(k.pn)
        rows.append(row)
    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame(columns=list(RAW_COLUMNS) + ['_key', 'pn'])
    out = out.sort_values(['pn', 'Cost center', 'Prod/Dept (Revenue)', 'Project'], na_position='last')
    return out.reset_index(drop=True)


def _order_book_code_hints(history: pd.DataFrame, min_share: float = 0.8) -> dict[tuple[str, str], str]:
    """(Cost center, Prod/Dept) → 'IQ3 (97%)' — 이력에서 그 부서의 Order Book Code가 한쪽으로 쏠릴 때만.

    Raw에 바로 채우지는 않는다: 95% 이상 쏠린 부서만 골라도 P10 신규 키 211개 중 7개(275.7M)가 틀렸다.
    """
    h = history.drop_duplicates('Key', keep='last').dropna(subset=['Order Book Code'])
    hints = {}
    for (cc, dept), s in h.groupby(['Cost center', 'Prod/Dept (Revenue)'])['Order Book Code']:
        share = s.value_counts(normalize=True)
        if share.iloc[0] >= min_share:
            hints[(cc, dept)] = f"{share.index[0]} ({share.iloc[0]:.0%})"
    return hints


def _review(keys: pd.DataFrame, lines: pd.DataFrame, ship: pd.DataFrame, raw_rows: pd.DataFrame, ir: pd.DataFrame,
            fx: FxBook, currency_map: dict[str, str], year: int, upto: int, history: pd.DataFrame) -> pd.DataFrame:
    items: list[dict] = []
    ystart, _ = month_bounds(year, 1)
    _, pend = month_bounds(year, upto)
    obc_hints = _order_book_code_hints(history)

    def add(kind, key='', project='', customer='', amount=np.nan, detail=''):
        items.append({'구분': kind, 'Key': key, 'Project': project, '고객': customer, '금액': amount, '내용': detail})

    # 회계 오더북에 없는 키 — 사람이 속성을 채워야 한다
    last = raw_rows.drop_duplicates('_key', keep='last').set_index('_key')
    for key, g in keys[keys['new_key']].groupby('Key'):
        a = last.loc[key]
        missing = [c for c in NEW_KEY_REQUIRED if pd.isna(a[c]) or a[c] in ('', None)]
        if missing:
            detail = '비어 있음: ' + ', '.join(missing)
            hint = obc_hints.get((a['Cost center'], a['Prod/Dept (Revenue)']))
            if hint and 'Order Book Code' in missing:
                detail += f" / Order Book Code 추천 {hint}"
            add('신규 키 — 속성 입력 필요', key, a['Project'], a['Customer Name'], g['Iv'].sum(), detail)

    # 마감월 큰 Variance
    km = keys[(keys['pn'] == upto) & keys['Vv'].abs().ge(BIG_VARIANCE)]
    for k in km.itertuples(index=False):
        add('큰 Variance', k.Key, k.Project, '', k.Vv, '환율 재평가 또는 수주·청구 금액 차이(단가·수량 변경)')

    # 수주·청구 금액 차이 — 출고·청구가 끝났는데 IR 금액 ≠ 청구 금액 (NOAH 'SO-DN 금액 차이 감지'와 같은 생각).
    # (Project, 라인번호) 묶음으로 본다 — 라인번호 재사용 때문에 라인 단위로 보면 출고가 옆 라인에 붙어 헛경보가 난다
    grp = lines.groupby(['Project', 'ln']).agg(net=('Net amount', 'sum'), all_done=('Line status', lambda s: s.eq('Invoiced').all()),
                                               last=('Actual ship date', 'max'), cust=('Customer name', 'first'),
                                               cur=('cur', 'first'))
    billed = ship[ship['billed']].groupby(['Project', 'ln'])['Net amount'].sum()
    pending_grp = set(map(tuple, ship.loc[ship['pending'], ['Project', 'ln']].to_numpy()))
    grp['billed'] = billed.reindex(grp.index)
    done = grp[grp['all_done'] & grp['billed'].notna() & grp['last'].ge(ystart) & grp['last'].lt(pend)
               & ~grp.index.isin(pending_grp) & (grp['billed'] - grp['net']).abs().gt(1.0)]
    for proj, g in done.groupby(level='Project'):
        add('수주·청구 금액 차이', '', proj, g['cust'].iloc[0], (g['billed'] - g['net']).sum(),
            f"라인번호 {', '.join(g.index.get_level_values('ln'))[:60]} — 청구 {g['billed'].sum():,.0f} vs 수주 "
            f"{g['net'].sum():,.0f} (통화 {g['cur'].iloc[0]}). Variance로 들어간다")

    # 라인 없는 출고(이 해에 금액이 남는 것만)
    nolink = ship[~ship['linked'] & ship['Project date'].ge(ystart) & ship['Project date'].lt(pend)]
    for proj, g in nolink.groupby('Project'):
        if abs(g['Net amount'].sum()) >= 0.5:
            add('IR 라인 없는 IT 출고', g['Key'].iloc[0], proj, '', g['Net amount'].sum(),
                f"{len(g)}행 — 라인번호 {', '.join(sorted(set(g['ln'].dropna())))[:60]}: 프로젝트 주 키에 붙임")

    # 통화 — IT 판매 통화와 우리 지정이 다르면
    linked = ship[ship['linked'] & ship['Project date'].ge(ystart)]
    bad_cur = linked[(linked['Sales currency'] != linked['cur']) & linked['Net amount'].abs().ge(0.5)]
    for proj, g in bad_cur.groupby('Project'):
        add('통화 확인', '', proj, '', g['Net amount'].sum(),
            f"IT 판매 통화 {', '.join(sorted(set(g['Sales currency'])))} ≠ 지정 {_mode(g['cur'])} — 통화지정.csv 확인")
    active = lines[(lines['created_kst'] >= ystart) | lines['Line status'].eq('Open order')
                   | lines['Actual ship date'].ge(ystart)]
    zr = active[active['VAT group'].isin(ZERO_RATED_VAT_GROUPS) & ~active['Customer'].isin(currency_map)
                & ~active['Project'].isin(currency_map) & ~active['Customer'].isin(KRW_ZERO_RATED_CUSTOMERS)
                & active['Net amount'].abs().ge(0.5)]
    for cust, g in zr.groupby('Customer'):
        add('통화 확인', '', ', '.join(sorted(g['Project'].unique())[:5]), _mode(g['Customer name']),
            g['Net amount'].sum(), f"VAT {_mode(g['VAT group'])}인데 외화 목록에 없음 — 원화로 계산함 "
                                   f"(중앙 단가 {g['Unit price'].median():,.0f})")

    unmapped = active[active['RevDept'].isna()]
    for m, g in unmapped.groupby('Prod/Dept (Material)'):
        add('부서 매핑 없음', '', ', '.join(sorted(g['Project'].unique())[:5]), '', g['Net amount'].sum(),
            f"Prod/Dept(Material) {m} → Revenue 부서를 모름 (라인 {len(g)}개)")
    ruled = active[active['dept_by_rule']]
    for m, g in ruled.groupby('Prod/Dept (Material)'):
        add('부서 규칙으로 추정', '', ', '.join(sorted(g['Project'].unique())[:5]), '', g['Net amount'].sum(),
            f"Prod/Dept(Material) {m} → {g['RevDept'].iloc[0]} — 배운 짝이 없어 규칙 적용 (라인 {len(g)}개)")

    for (cur, month), near in sorted(fx.fallbacks.items()):
        add('환율 없음 — FX 시트에 추가', '', '', '', np.nan,
            f"{cur} {month} 환율이 없어 {near} 환율로 계산함 — NOAH_SO_PO_DN FX 시트에 {month} 열을 넣고 sync_db.py 후 다시 실행")
    return pd.DataFrame(items, columns=['구분', 'Key', 'Project', '고객', '금액', '내용'])


def _next_filter_note(ir: pd.DataFrame, last_rows: pd.DataFrame) -> str:
    """추출 필터(Created date ≥ X)가 충분한지 — 잔고가 남은 가장 오래된 프로젝트보다 앞이어야 한다."""
    cutoff = ir['Created date'].min()
    with_balance = last_rows[(last_rows['Value End'].abs() > _EPS) | (last_rows['Qty End'].abs() > _EPS)]
    oldest = pd.to_datetime(with_balance['Project creation date'], errors='coerce').min()
    if pd.isna(oldest) or pd.isna(cutoff):
        return ''
    msg = (f"IR 생성일 최솟값 {cutoff:%Y-%m-%d} / 잔고가 남은 가장 오래된 프로젝트 생성일 {oldest:%Y-%m-%d}. ")
    if oldest < cutoff.normalize():
        return msg + "추출 필터 날짜를 더 이르게 잡아야 합니다 (오래된 라인이 빠졌을 수 있음)."
    return msg + f"필터 날짜는 {oldest:%Y-%m-01} 이전이면 충분합니다."


# ---------------------------------------------------------------------------
# 회계 오더북과 대조
# ---------------------------------------------------------------------------
def compare_with_official(keys: pd.DataFrame, official_rows: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """한 달치 대조 — keys = 그달 우리 키별 계산(Key, Project, Sq..Vv), official_rows = 그달 회계 Raw 행."""
    off = official_rows.groupby('Key')[list(NUM_COLUMNS)].sum()
    ours = keys.set_index('Key')[['Project'] + list(_KEY_VALUES)]
    j = ours.join(off, how='outer')
    for c in list(_KEY_VALUES) + list(NUM_COLUMNS):
        j[c] = j[c].fillna(0.0)
    proj_off = official_rows.drop_duplicates('Key').set_index('Key')['Project']
    j['Project'] = j['Project'].fillna(proj_off)
    pairs = {'Start': ('Sv', 'Value Start'), 'Input': ('Iv', 'Value Input'), 'Output': ('Ov', 'Value Output'),
             'End': ('Ev', 'Value End'), 'Variance': ('Vv', 'Value Variance')}
    for name, (a, b) in pairs.items():
        j[f'd_{name}'] = j[a] - j[b]
    j['d_Orders'] = (j['Iv'] + j['Vv']) - (j['Value Input'] + j['Value Variance'])
    j['in_ours'] = j.index.isin(ours.index)
    j['in_official'] = j.index.isin(off.index)
    tot_rows = [{'항목': name, '회계': j[b].sum(), '계산': j[a].sum()} for name, (a, b) in pairs.items()]
    tot_rows.append({'항목': 'Orders (Input+Variance)', '회계': (j['Value Input'] + j['Value Variance']).sum(),
                     '계산': (j['Iv'] + j['Vv']).sum()})
    totals = pd.DataFrame(tot_rows)
    totals['차이'] = totals['계산'] - totals['회계']
    totals['차이%'] = np.where(totals['회계'].abs() > 0, totals['차이'] / totals['회계'].abs() * 100, np.nan)
    match = []
    for name, (a, b) in pairs.items():
        act = j[(j[a].abs() > 1) | (j[b].abs() > 1)]
        exact = int((act[a] - act[b]).abs().le(1).sum())
        match.append({'항목': name, '대상 키': len(act), '원 단위 일치': exact,
                      '일치율%': round(exact / len(act) * 100, 1) if len(act) else np.nan})
    proj = j.groupby('Project')[['Ov', 'Value Output', 'Ev', 'Value End']].sum()
    proj = proj[(proj.abs() > 1).any(axis=1)]
    proj_match = {'Output': round(((proj['Ov'] - proj['Value Output']).abs() <= 1).mean() * 100, 1) if len(proj) else np.nan,
                  'End': round(((proj['Ev'] - proj['Value End']).abs() <= 1).mean() * 100, 1) if len(proj) else np.nan,
                  '프로젝트': len(proj)}
    return {'totals': totals, 'match': pd.DataFrame(match), 'keys': j.reset_index(), 'project_match': proj_match}


DIFF_CAUSES: tuple[str, ...] = (
    '회계 조정', '취소 주문(회계는 그달 Variance)', '지난 달 차이 이월', '부서·CC 귀속 차이', '청구 월',
    '취소 주문(처음부터 제외)', '수주 사후 변경(IR은 지금 값)', '기초 차이', '환율', '기타',
)
# 결과 HTML의 원인 설명 — 판정 규칙은 classify_differences 독스트링이 소유하고, 여기는 읽는 사람용 한 줄
DIFF_CAUSE_NOTES: dict[str, str] = {
    '회계 조정': '회계가 그 키의 Start를 손으로 고쳤다 (회계 Start ≠ 회계 직전 달 End)',
    '취소 주문(회계는 그달 Variance)': '그달 흐름은 같고 Start·Variance만 다르다 — 회계는 취소한 달에 빼고 우리는 처음부터 뺐다',
    '지난 달 차이 이월': '그달 Input·Output·End는 같다 — 지난 달까지의 차이가 Start로 넘어온 것',
    '부서·CC 귀속 차이': '프로젝트 합계는 같고, 금액이 귀속된 키(Cost center·부서)만 다르다',
    '청구 월': 'Output 차이가 앞뒤 달과 상쇄된다 — 우리는 출고월, AX는 청구월(월말 출고·다음 달 청구)',
    '취소 주문(처음부터 제외)': '우리 End가 작고 그 프로젝트에 취소 라인이 있다 — 우리는 취소 주문을 처음부터 뺀다',
    '수주 사후 변경(IR은 지금 값)': 'End 차이가 Input 차이로 설명된다 — IR은 바뀐 지금 값, AX는 수주 당시 값으로 넣고 바뀐 달에 고친다',
    '기초 차이': 'P01 — 회계 P01 Start와 IR·IT가 본 전년 12월 말 잔고의 차이',
    '환율': '외화 키의 환율 차이',
    '기타': '위 어디에도 맞지 않는 차이',
}


def classify_differences(result: OrderbookResult, official: pd.DataFrame, ir: pd.DataFrame) -> pd.DataFrame:
    """달·키별 우리 vs 회계 차이 + 원인. 원인은 위에서부터 처음 맞는 것:

    회계 조정        공식 Start ≠ 공식 직전 달 End (회계가 키를 손으로 고쳤다)
    (그달은 맞음)    Input·Output·End는 같고 Start·Variance만 다르다 — 지난 달 차이가 넘어온 것.
                     그 프로젝트에 취소 라인이 있으면 '취소 주문(회계는 그달 Variance)', 아니면 '지난 달 차이 이월'
    부서·CC 귀속 차이  프로젝트 합계는 맞는데 금액이 귀속된 키(부서·CC)만 다르다
    청구 월          Output 차이가 앞뒤 달 Output 차이와 상쇄된다 — 출고월 인식 vs AX 청구월(월말 출고·다음 달 청구)
    취소 주문        우리 End가 작고 그 프로젝트에 월말 전 생성된 취소 라인이 있다 — 우리는 처음부터 뺀다
    수주 사후 변경   End 차이가 그 키의 지금까지 Input 차이로 설명된다 — IR은 수량·단가·삭제가 반영된 지금 값이고,
                     AX는 수주 당시 값으로 넣고 바뀐 달에 Variance로 고친다(82-16670.001 +35.4M, 82-17444.001 −28.8M)
    기초 차이        P01 — 회계 P01 Start와 IR·IT가 본 12월 말 잔고의 차이
    환율             외화 키
    기타
    """
    frames = []
    for m in range(1, result.upto + 1):
        om = official[official['pn'] == m]
        if om.empty:
            continue
        cmp = compare_with_official(result.keys[result.keys['pn'] == m], om)['keys']
        cmp['pn'] = m
        frames.append(cmp)
    if not frames:
        return pd.DataFrame()
    j = pd.concat(frames, ignore_index=True)
    dcols = ['d_Start', 'd_Input', 'd_Output', 'd_End', 'd_Variance']
    j = j[j[dcols].abs().gt(1).any(axis=1)].copy()

    proj = j.groupby(['pn', 'Project'])[['d_Input', 'd_Output', 'd_End']].sum()
    proj_split = proj.abs().le(1).all(axis=1)
    proj_out = proj['d_Output']
    off_end = official.groupby(['pn', 'Key'])['Value End'].sum()
    off_start = official.groupby(['pn', 'Key'])['Value Start'].sum()
    cancelled = ir[ir['Line status'] == 'Cancelled'].groupby('Project')['created_kst'].min()
    foreign = {k for k, c in result.key_currency.items() if c != 'KRW'}
    j = j.sort_values(['Key', 'pn'])
    j['cum_dI'] = j.groupby('Key')['d_Input'].cumsum()

    causes = []
    for r in j.itertuples(index=False):
        m, key, project = r.pn, r.Key, r.Project
        _, pend = month_bounds(result.year, m)
        has_cancel = project in cancelled.index and cancelled[project] < pend
        if m > 1 and (m - 1, key) in off_end.index and abs(off_start.get((m, key), 0.0) - off_end[(m - 1, key)]) > 1:
            causes.append('회계 조정')
        elif max(abs(r.d_Input), abs(r.d_Output), abs(r.d_End)) <= 1:
            causes.append('취소 주문(회계는 그달 Variance)' if has_cancel and r.d_Start < 0 else '지난 달 차이 이월')
        elif proj_split.get((m, project), False):
            causes.append('부서·CC 귀속 차이')
        elif abs(r.d_Output) > 1 and any(
                abs(proj_out.get((m, project), 0.0) + proj_out.get((m + s, project), 0.0))
                <= max(1.0, 0.01 * abs(proj_out.get((m, project), 0.0))) for s in (-1, 1)):
            causes.append('청구 월')
        elif r.d_End < -1 and has_cancel:
            causes.append('취소 주문(처음부터 제외)')
        elif abs(r.cum_dI) > 1 and abs(r.d_End - r.cum_dI) <= max(1.0, 0.001 * abs(r.cum_dI)):
            causes.append('수주 사후 변경(IR은 지금 값)')
        elif m == 1:
            causes.append('기초 차이')
        elif key in foreign:
            causes.append('환율')
        else:
            causes.append('기타')
    j['원인'] = causes
    j['Period'] = j['pn'].map(lambda p: f"P{int(p):02d}")
    j = j.drop(columns='cum_dI')
    order = ['Period', 'Key', 'Project', '원인'] + dcols + ['Sv', 'Iv', 'Ov', 'Ev', 'Vv', 'Value Start', 'Value Input',
                                                         'Value Output', 'Value End', 'Value Variance',
                                                         'in_ours', 'in_official']
    return j[order].sort_values(['Period', '원인', 'd_End']).reset_index(drop=True)
