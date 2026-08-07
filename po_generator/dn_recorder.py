"""
DN_국내 출고기록 계산 (Delivery Note Recorder)
==============================================

공장 출고리스트(`po_reconciliation/{year}/{period}/2026리스트_RCK_{period}.xlsx`)를 읽어
`DN_국내`에 추가할 행을 계산합니다. **Excel COM을 쓰지 않습니다** — 쓰기는 `dn_writer.py` 몫.

라인/수량을 어디서 가져오는가
------------------------------
출고리스트는 "어느 주문이 언제 나갔는지"만 알려주고 **어느 라인이 몇 개 나갔는지는 없다**.
그래서 `PO_국내`의 `Status`를 본다 — 공장이 그 달에 계산서를 끊은 라인이 `Invoiced P{XX}`다.

    pending = 이 SO_ID의 PO 라인 중 Status ∉ {Invoiced P01..P{XX}, Cancelled}

    pending 없음  →  마지막 출고: `SO_국내` 전 라인의 잔량 (= Item qty − 기출고 DN 누계)
    pending 있음  →  부분 출고  : `PO_국내` Invoiced P{XX} 라인의 Item qty 그대로

`PO_국내`만 보면 안 되는 이유: **PO는 부속을 1라인에 합쳐 적는다.**
`SA09X-MA + ADAPTER`, `NA015 (...) / 부싱가공 (별각 22X22)` 처럼 SO 2라인이 PO 1라인이 되는데,
PO 기준으로만 뽑으면 부속 라인이 통째로 빠진다 (2026-03~08 실측 6건, 최대 360,000원).
반대로 `Cancelled` 라인을 안 빼면 취소분이 출고로 잡힌다 (같은 기간 2건).

자기검증 — 통과분만 자동 입력
------------------------------
    (a) Σ(PO Invoiced P{XX}의 Total ICO) == 출고리스트 계산서금액
    (b) 출고리스트에 같은 SO_ID가 여러 날짜로 있지 않을 것

(b)는 PO 라인을 어느 날짜에 배분할지 알 방법이 없어 자동화가 불가능하다.
하나라도 걸리면 `Review`로 빼고 사람에게 넘긴다 — 2026-03~08 561건 리플레이 기준
자동 입력 549건 중 548건 정확(99.8%), 걸러낸 12건은 전부 진짜 확인 대상이었다(헛경보 0).

멱등성
------
`(SO_ID, 출고일)`이 이미 `DN_국내`에 있으면 건너뛴다. 같은 기간을 두 번 돌려도 중복 입력되지
않는다 — 출고리스트는 월중에 갱신되며 여러 번 실행하는 것이 정상 사용 방식이다.
"""

from __future__ import annotations

import datetime as dt
import logging
import re
from dataclasses import dataclass, field
from itertools import combinations
from pathlib import Path

import pandas as pd

from po_generator.config import EXCLUDED_SO_STATUSES
from po_generator.recon_paths import resolve_period_dir

logger = logging.getLogger(__name__)

# 출고리스트 Delivery 시트
DELIVERY_SHEET = 'Delivery'
COL_SO_ID = 'SO_ID'
COL_SHIP_DATE = '납품완료'
COL_INVOICE_AMOUNT = '계산서금액'
COL_CUSTOMER = 'Customer'
COL_RCK_ORDER = 'RCK ODER'

# 금액 비교 허용 오차 (원) — reconcile_po.py와 같은 기준
AMOUNT_TOLERANCE = 1.0

# 월합 세금계산서 거래처 판별에 쓰는 Remarks 키워드
MONTHLY_CLOSE_KEYWORDS = ('월합', '마감')

_DN_ID_RE = re.compile(r'^DND-(\d{4})-(\d+)$')
_PERIOD_RE = re.compile(r'^P(\d{1,2})$', re.IGNORECASE)


@dataclass(frozen=True)
class DnLine:
    """DN_국내에 추가할 행 하나."""
    dn_id: str
    so_id: str
    line_item: int
    qty: int
    currency: str
    ship_date: dt.datetime
    tax_date: dt.datetime | None
    remarks: str | None
    seq: int
    # 아래는 표시/검증용 — 시트에서는 수식이 채운다
    item_name: str = ''
    customer_name: str = ''


@dataclass(frozen=True)
class Review:
    """자동 입력에서 빼고 사람이 봐야 하는 출고 건."""
    so_id: str
    ship_date: dt.datetime | None
    customer: str
    reason: str
    detail: str


@dataclass
class Plan:
    """한 기간(period)의 계산 결과."""
    period: str
    delivery_file: Path
    lines: list[DnLine] = field(default_factory=list)
    reviews: list[Review] = field(default_factory=list)
    skipped: list[Review] = field(default_factory=list)   # 이미 입력된 건

    @property
    def dn_ids(self) -> list[str]:
        seen: dict[str, None] = {}
        for line in self.lines:
            seen.setdefault(line.dn_id, None)
        return list(seen)


# === 소스 로딩 =============================================================

def find_delivery_file(recon_dir: Path, period: str) -> Path | None:
    """기간 폴더에서 출고리스트 파일 찾기 (`reconcile_po.find_file`과 같은 규칙)"""
    period_dir = resolve_period_dir(recon_dir, period)
    if period_dir is None:
        return None
    for keyword in ('리스트', 'RCK'):
        for f in sorted(period_dir.iterdir()):
            if f.suffix == '.xlsx' and not f.name.startswith('~'):
                if keyword.upper() in f.name.upper():
                    return f
    return None


def load_source_frames(workbook: Path,
                       sheets: tuple[str, ...]) -> tuple[pd.DataFrame, ...]:
    """워크북에서 시트들을 읽고 **파일 핸들을 반드시 닫는다**

    `pd.ExcelFile`은 닫을 때까지 파일을 붙잡고 있다. 그대로 두면 뒤이어 Excel에게
    같은 파일을 쓰기로 열라고 할 때 **우리 프로세스가 우리를 막는다** —
    Excel은 읽기 전용으로 열거나(저장이 조용히 무시된다) 아예 열기에 실패한다.
    2026-08-07 실사용에서 이 순서로 두 번 다 터졌다. 읽고 쓰는 CLI에서는
    읽기 핸들을 넘기지 않는 것이 규약이다.
    """
    with pd.ExcelFile(workbook) as xf:
        return tuple(pd.read_excel(xf, sheet) for sheet in sheets)


def load_delivery(delivery_file: Path) -> pd.DataFrame:
    """출고리스트 Delivery 시트에서 SO_ID가 있는 행만 (N/A = 서비스 출고라 DN 대상 아님)"""
    df = pd.read_excel(delivery_file, sheet_name=DELIVERY_SHEET)
    if COL_SO_ID not in df.columns:
        raise ValueError(
            f"출고리스트에 '{COL_SO_ID}' 컬럼이 없습니다: {delivery_file.name}")
    df = df[df[COL_SO_ID].notna()].copy()
    df[COL_SO_ID] = df[COL_SO_ID].astype(str).str.strip()
    df = df[df[COL_SO_ID] != ''].copy()
    df[COL_SHIP_DATE] = pd.to_datetime(df[COL_SHIP_DATE], errors='coerce')
    logger.debug("출고리스트 로드: %d건 (%s)", len(df), delivery_file.name)
    return df


# === 채번 =================================================================

def next_dn_seq(dn_df: pd.DataFrame, year: int) -> int:
    """해당 연도의 다음 DN 일련번호 (없으면 1)"""
    best = 0
    for value in dn_df.get('DN_ID', pd.Series(dtype=object)).dropna():
        m = _DN_ID_RE.match(str(value).strip())
        if m and int(m.group(1)) == year:
            best = max(best, int(m.group(2)))
    return best + 1


def format_dn_id(year: int, seq: int) -> str:
    return f"DND-{year}-{seq:04d}"


def next_row_seq(dn_df: pd.DataFrame) -> int:
    """`Seq` 열의 다음 값 — 정렬용 일련번호라 단순 증가면 된다."""
    seq = pd.to_numeric(dn_df.get('Seq', pd.Series(dtype=float)), errors='coerce')
    return int(seq.max()) + 1 if seq.notna().any() else 1


def period_month(period: str) -> int:
    m = _PERIOD_RE.match(period.strip())
    if not m:
        raise ValueError(f"기간 코드 형식이 아닙니다 (예: P08): {period}")
    month = int(m.group(1))
    if not 1 <= month <= 12:
        raise ValueError(f"기간 코드의 월이 1~12가 아닙니다: {period}")
    return month


def settled_statuses(month: int) -> set[str]:
    """'더 나갈 것이 없다'고 볼 수 있는 PO Status 집합"""
    return {f"Invoiced P{m:02d}" for m in range(1, month + 1)} | {'Cancelled'}


# === 월합 세금계산서 거래처 ==================================================

def monthly_close_customers(dn_df: pd.DataFrame) -> dict[str, str]:
    """사업자번호 → 월합 Remarks (기존 DN에서 유도)

    월합 거래처는 출고 시점에 세금계산서를 끊지 않으므로 발행일을 비우고 Remarks를 물려준다.
    2026-08-06 실측: 이 규칙이 P08 수기 입력분(엔이에스 615-81-88675만 공란)을 그대로 재현한다.
    """
    if 'Remarks' not in dn_df.columns or 'Business registration number' not in dn_df.columns:
        return {}
    remarks = dn_df['Remarks'].astype(str)
    mask = remarks.str.contains('|'.join(MONTHLY_CLOSE_KEYWORDS), na=False)
    out: dict[str, str] = {}
    for biz, grp in dn_df[mask].groupby('Business registration number'):
        key = normalize_biz_no(biz)
        if not key:
            continue
        counts = grp['Remarks'].dropna().astype(str).value_counts()
        if len(counts):
            out[key] = counts.index[0]
    return out


def normalize_biz_no(value: object) -> str:
    """사업자번호에서 숫자만 남긴다 (하이픈 표기 차이로 갈리지 않게)"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ''
    return re.sub(r'\D', '', str(value))


# === 핵심 계산 =============================================================

def build_plan(
    period: str,
    delivery: pd.DataFrame,
    so_df: pd.DataFrame,
    po_df: pd.DataFrame,
    dn_df: pd.DataFrame,
    *,
    delivery_file: Path | None = None,
    fill_tax_date: bool = True,
) -> Plan:
    """출고리스트 → DN_국내 추가 행 계산

    Args:
        period: 기간 코드 (예: 'P08')
        delivery: `load_delivery()` 결과
        so_df/po_df/dn_df: `SO_국내` / `PO_국내` / `DN_국내`
        fill_tax_date: 세금계산서 발행일을 출고일과 같게 채울지
                       (월합 거래처는 이 값과 무관하게 항상 공란)
    """
    period = period.upper()
    month = period_month(period)
    status_col = po_df['Status'].astype(str).str.strip() if 'Status' in po_df else None
    invoiced_now = f"Invoiced {period}"
    settled = settled_statuses(month)

    plan = Plan(period=period, delivery_file=delivery_file or Path())

    # --- 조회용 인덱스 -----------------------------------------------------
    so_lines = _so_lines_by_order(so_df)
    so_meta = _so_meta_by_order(so_df)
    po_by_so = _po_by_order(po_df, status_col)
    already = _existing_shipments(dn_df)
    shipped = _shipped_qty(dn_df)
    # 중복 방어는 **실행 전 시트 상태**만 본다. `shipped`는 잔량 계산용이라 실행 중
    # 갱신되는데, 그걸 같이 보면 한 SO가 여러 날 나갈 때 앞 출고 때문에 뒤 출고가
    # "이미 기록됨"으로 걸려 통째로 사라진다 (2026-05 SOD-2026-0188 실측).
    recorded_maps = _shipment_maps(dn_df)
    monthly = monthly_close_customers(dn_df)

    # 출고 이벤트는 (SO_ID, 출고일) 단위 — 같은 주문·같은 날이 두 줄로 적히기도 한다
    # (2026-05 SOD-2026-0467 실측). 행마다 돌면 같은 출고를 두 번 기록하게 된다.
    events = _delivery_events(delivery)
    ship_dates = events.groupby(COL_SO_ID)[COL_SHIP_DATE].nunique().to_dict()
    invoice_total = events.groupby(COL_SO_ID)['_invoice'].sum().to_dict()

    # 같은 SO가 여러 날 나갔으면 PO 행의 ICO 금액을 날짜별 계산서금액에 배정해 본다
    splits, corrections = _split_multi_date(
        events, po_by_so, invoiced_now, ship_dates)

    year_hint = _year_from_delivery(delivery)
    dn_seq = next_dn_seq(dn_df, year_hint)
    row_seq = next_row_seq(dn_df)

    # 출고일 순 — 파일 순서가 이미 날짜순이지만 stable sort로 못 박는다
    ordered = events.sort_values(
        COL_SHIP_DATE, kind='stable', na_position='last')

    for _, row in ordered.iterrows():
        so_id = str(row[COL_SO_ID]).strip()
        ship_date = row[COL_SHIP_DATE]
        customer = str(row.get('_customer', '') or '')

        if pd.isna(ship_date):
            plan.reviews.append(Review(
                so_id, None, customer, '출고일 없음',
                f"출고리스트의 '{COL_SHIP_DATE}'가 비어 있습니다"))
            continue

        if (so_id, ship_date.normalize()) in already:
            plan.skipped.append(Review(
                so_id, ship_date, customer, '이미 입력됨',
                f"DN_국내에 {so_id} / {ship_date:%Y-%m-%d} 행이 이미 있습니다"))
            continue

        if so_id not in so_lines:
            plan.reviews.append(Review(
                so_id, ship_date, customer, 'SO 없음',
                f"SO_국내에 {so_id}가 없습니다"))
            continue

        po_rows = po_by_so.get(so_id)
        if po_rows is None or po_rows.empty:
            plan.reviews.append(Review(
                so_id, ship_date, customer, 'PO 없음',
                f"PO_국내에 {so_id}가 없습니다"))
            continue

        invoiced_rows = po_rows[po_rows['_status'] == invoiced_now]
        if invoiced_rows.empty:
            statuses = ', '.join(sorted(set(po_rows['_status']))) or '(없음)'
            plan.reviews.append(Review(
                so_id, ship_date, customer, f'{invoiced_now} 없음',
                f"PO 상태가 {statuses} — 공장 계산서가 이 기간에 안 잡혔습니다"))
            continue

        if ship_date.normalize() in corrections.get(so_id, ()):
            # 출고가 아니라 단가 정정 행 — 금액은 첫 출고에 합산돼 이미 반영된다
            plan.skipped.append(Review(
                so_id, ship_date, customer, '단가 정정 행',
                '출고가 아니라 금액만 정정한 행입니다 (계산서금액은 첫 출고에 합산)'))
            continue

        if ship_dates.get(so_id, 1) > 1 and so_id not in corrections:
            # 날짜별 계산서금액에 PO 행을 금액으로 배정한다. PO는 분할출고마다 행을
            # 따로 두므로 대개 정확히 나뉜다 (2026-03~07 12건 중 10건, 모호 0건).
            assigned = splits.get(so_id)
            if assigned is None:
                negative = any(
                    v < 0 for v in events.loc[
                        events[COL_SO_ID] == so_id, '_invoice'])
                reason = ('반품 포함' if negative else '같은 SO 여러 날 출고')
                detail = (
                    '계산서금액이 음수인 행이 있습니다 — 출고/반품/재출고를 어떻게 '
                    '적을지는 사람이 정해야 합니다'
                    if negative else
                    f"출고리스트에 {ship_dates[so_id]}개 날짜인데 PO 금액이 "
                    "날짜별 계산서금액으로 안 나뉩니다 — 어느 라인이 어느 날 "
                    "나갔는지 판단할 수 없습니다")
                plan.reviews.append(Review(
                    so_id, ship_date, customer, reason, detail))
                continue
            # 배정은 PO 전 라인으로 하고(금액이 맞아야 하므로), DN에 쓸 때 SO에 없는
            # 라인을 뺀다 — 2026-05 SOD-2026-0188이 정확히 이 모양이다
            qty_by_line = sales_only(
                assigned.get(ship_date.normalize(), {}), so_lines[so_id])
        else:
            ico_sum = float(pd.to_numeric(
                invoiced_rows['Total ICO'], errors='coerce').fillna(0).sum())
            inv_amt = float(invoice_total.get(so_id, 0) or 0)
            if abs(ico_sum - inv_amt) >= AMOUNT_TOLERANCE:
                plan.reviews.append(Review(
                    so_id, ship_date, customer, '금액 불일치',
                    f"PO ICO 합 {ico_sum:,.0f} vs 계산서금액 {inv_amt:,.0f} "
                    f"(차이 {ico_sum - inv_amt:,.0f})"))
                continue

            pending = po_rows[~po_rows['_status'].isin(settled)]
            is_final = pending.empty
            qty_by_line = (
                _final_shipment_lines(
                    so_id, so_lines[so_id], po_rows, shipped, ship_date.normalize())
                if is_final
                else _partial_shipment_lines(invoiced_rows, so_lines[so_id])
            )

        if not qty_by_line:
            # 잔량이 0 = 이미 다 나갔다는 뜻이지 확인할 거리가 아니다. 사람이 출고일을
            # 하루 이틀 다르게 적어 두면 이 경로로 온다 (2026-07 SOD-2026-0790:
            # 출고리스트 7/29, DN 7/27). 확인 목록에 올리면 헛경보만 쌓인다.
            done = _recorded_summary(recorded_maps, so_id)
            if done:
                plan.skipped.append(Review(
                    so_id, ship_date, customer, '이미 출고 완료',
                    f"잔량이 없습니다 — 기록된 출고: {done}"))
            else:
                plan.reviews.append(Review(
                    so_id, ship_date, customer, '기록할 라인 없음',
                    '대상 라인이 모두 취소/보류이거나 SO에 없습니다'))
            continue

        # 출고리스트 날짜와 다른 날로 이미 기록해 둔 건 (사람이 실제 출고일을 알고 고쳐 적은 경우).
        # (SO_ID, 출고일) 키만 보면 못 걸러 지난 기간을 다시 돌릴 때 통째로 중복 입력된다
        # (2026-03 SOD-2026-0156 실측: 출고리스트는 3/23인데 DN은 5/19로 기록돼 있다).
        recorded = _recorded_date(recorded_maps, so_id, qty_by_line)
        if recorded:
            plan.skipped.append(Review(
                so_id, ship_date, customer, '이미 입력됨(다른 출고일)',
                f"라인·수량이 똑같은 출고가 {recorded}에 이미 기록돼 있습니다"))
            continue

        meta = so_meta.get(so_id, {})
        biz_no = normalize_biz_no(meta.get('biz_no'))
        remarks = monthly.get(biz_no)
        tax_date = None if (remarks or not fill_tax_date) else ship_date

        dn_id = format_dn_id(year_hint, dn_seq)
        dn_seq += 1
        for line_item in sorted(qty_by_line):
            plan.lines.append(DnLine(
                dn_id=dn_id,
                so_id=so_id,
                line_item=int(line_item),
                qty=int(qty_by_line[line_item]),
                currency=str(meta.get('currency') or 'KRW'),
                ship_date=ship_date,
                tax_date=tax_date,
                remarks=remarks,
                seq=row_seq,
                item_name=str(so_lines[so_id].get(line_item, ('', 0))[0]),
                customer_name=str(meta.get('customer_name') or customer),
            ))
            row_seq += 1
            # 같은 실행 안에서 뒤 출고가 앞 출고를 잔량에서 빼도록 누계 갱신
            shipped.setdefault((so_id, int(line_item)), []).append(
                (ship_date.normalize(), int(qty_by_line[line_item])))

    return plan


def _final_shipment_lines(
    so_id: str,
    lines: dict[int, tuple[str, int]],
    po_rows: pd.DataFrame,
    shipped: ShipLog,
    ship_date: pd.Timestamp,
) -> dict[int, int]:
    """마지막 출고 — SO 전 라인의 잔량. PO가 전부 취소된 라인은 뺀다."""
    cancelled = {
        int(li) for li, grp in po_rows.groupby('_line')
        if set(grp['_status']) <= {'Cancelled'}
    }
    out: dict[int, int] = {}
    for line_item, (_name, qty) in lines.items():
        if line_item in cancelled:
            continue
        remaining = int(qty) - _prior_qty(shipped, so_id, line_item, ship_date)
        if remaining > 0:
            out[line_item] = remaining
    return out


def _delivery_events(delivery: pd.DataFrame) -> pd.DataFrame:
    """출고리스트를 (SO_ID, 출고일) 단위로 합친다

    같은 주문이 같은 날 두 줄로 적히는 경우가 있다(2026-05 SOD-2026-0467).
    행마다 돌면 같은 출고를 두 번 기록하므로 여기서 한 번 합친다.
    `dropna=False`여야 출고일이 빈 행도 남아 '출고일 없음'으로 걸린다.
    """
    amount = (pd.to_numeric(delivery.get(COL_INVOICE_AMOUNT), errors='coerce')
              if COL_INVOICE_AMOUNT in delivery.columns else 0)
    customer = (delivery[COL_CUSTOMER] if COL_CUSTOMER in delivery.columns else '')
    df = pd.DataFrame({
        COL_SO_ID: delivery[COL_SO_ID],
        COL_SHIP_DATE: delivery[COL_SHIP_DATE],
        '_invoice': amount,
        '_customer': customer,
    })
    return (df.groupby([COL_SO_ID, COL_SHIP_DATE], dropna=False, sort=False)
            .agg(_invoice=('_invoice', 'sum'), _customer=('_customer', 'first'))
            .reset_index())


def _split_multi_date(
    events: pd.DataFrame,
    po_by_so: dict[str, pd.DataFrame],
    invoiced_now: str,
    ship_dates: dict[str, int],
) -> tuple[dict[str, dict[pd.Timestamp, dict[int, int]] | None],
           dict[str, set[pd.Timestamp]]]:
    """여러 날 나간 SO를 푼다

    Returns:
        (splits, corrections)
        splits[so_id]     = {출고일: {라인: 수량}} — 날짜별로 갈랐을 때. 못 풀면 None
        corrections[so_id] = {정정 행의 날짜, ...} — 출고가 아니라 단가 정정인 행

    `PO_국내`는 **분할출고마다 행을 따로 둔다** — 그래서 PO 행들의 `Total ICO`를
    날짜별 `계산서금액`에 정확히 나눠 담을 수 있다 (2026-03~07 실측 12건 중 10건 유일 배정,
    모호 0건).

    안 나뉘는 경우 중 하나는 **출고가 두 번이 아니라 단가를 정정한 것**이다.
    2026-05 `SOD-2026-0306`: 5/11 계산서 9,956,592 → 5/13에 `L260441-1R`로 103,616.
    뒤 행은 `AMOUNT`가 10,060,208로 바뀌었고 103,616은 그 **차액**이다.
    출고는 5/11 한 번(8개)이고 5/13은 돈만 오간 행이라, 합계가 PO ICO와 맞으면
    한 출고로 보고 뒤 행은 건너뛴다.

    반품(계산서금액 음수)이 섞인 건은 손대지 않는다 — 출고/반품/재출고를 어떻게 적을지는
    사람이 정할 일이다 (2026-03 `SOD-2026-0280`).
    """
    splits: dict[str, dict[pd.Timestamp, dict[int, int]] | None] = {}
    corrections: dict[str, set[pd.Timestamp]] = {}

    for so_id, count in ship_dates.items():
        if count <= 1:
            continue
        po_rows = po_by_so.get(so_id)
        if po_rows is None:
            splits[so_id] = None
            continue
        rows = po_rows[po_rows['_status'] == invoiced_now]
        items = [
            (int(r['_line']),
             int(pd.to_numeric(r['Item qty'], errors='coerce') or 0),
             float(pd.to_numeric(r.get('Total ICO'), errors='coerce') or 0.0))
            for _, r in rows.iterrows()
        ]
        grp = (events[events[COL_SO_ID] == so_id]
               .dropna(subset=[COL_SHIP_DATE])
               .sort_values(COL_SHIP_DATE))
        dates = [d.normalize() for d in grp[COL_SHIP_DATE]]
        targets = [float(v) for v in grp['_invoice']]

        solution, why = _solve_split(items, targets)
        if solution is not None:
            splits[so_id] = dict(zip(dates, solution))
            continue

        # **나눌 방법이 아예 없을 때만** 단가 정정을 의심한다. 'ambiguous'는 출고가
        # 여러 번인데 어느 쪽인지 모르는 것이라 한 건으로 뭉치면 안 된다.
        splits[so_id] = None
        if why != 'no_split' or any(t < 0 for t in targets):
            continue                        # 모호하거나 반품 — 사람 몫
        ico_total = sum(i[2] for i in items)
        if abs(ico_total - sum(targets)) < AMOUNT_TOLERANCE:
            del splits[so_id]
            corrections[so_id] = set(dates[1:])   # 첫 날이 실제 출고
            logger.debug("단가 정정으로 판단 — %s: 출고 %s, 정정 %s",
                         so_id, dates[0].date(), [d.date() for d in dates[1:]])
    return splits, corrections


# 배정 탐색 상한 — 조합 탐색이라 항목이 많으면 폭발한다. 넘으면 '못 풀었다'로
# 처리해 사람에게 넘긴다 (오래 도는 것보다 확인 목록에 뜨는 편이 낫다).
MAX_SPLIT_ITEMS = 24
MAX_SPLIT_NODES = 200_000


def split_by_amount(
    items: list[tuple[int, int, float]],
    targets: list[float],
) -> list[dict[int, int]] | None:
    """PO 행들을 날짜별 금액에 정확히 나눠 담는다 (못 하면 None)

    Args:
        items: [(Line item, Item qty, Total ICO), ...] — PO 행 하나당 한 항목
        targets: 날짜순 계산서금액

    Returns:
        targets와 같은 순서의 [{라인: 수량}, ...].
        나눌 방법이 없거나 **결과가 갈리는 방법이 둘 이상이면** None (= 사람 확인).
    """
    return _solve_split(items, targets)[0]


def _solve_split(
    items: list[tuple[int, int, float]],
    targets: list[float],
) -> tuple[list[dict[int, int]] | None, str]:
    """`split_by_amount` + **왜 못 했는지**

    사유를 구분해야 하는 이유: '나눌 방법이 아예 없다'는 출고가 한 번이었다는 뜻일 수
    있지만(단가 정정 행이 섞인 경우), '결과가 갈린다'는 출고가 여러 번인데 어느 쪽인지
    모른다는 뜻이다. 둘을 같이 취급하면 후자를 한 건으로 뭉쳐 버린다.

    Returns:
        (해답 또는 None, 사유) — 사유는 'ok' | 'ambiguous' | 'no_split' | 'too_many'
    """
    if not targets or not items:
        return None, 'no_split'
    if len(items) > MAX_SPLIT_ITEMS:
        return None, 'too_many'

    budget = [MAX_SPLIT_NODES]
    found: list[tuple[tuple, list[dict[int, int]]]] = []
    for groups in _search_split(items, targets, budget):
        result = [_qty_map(g) for g in groups]
        key = tuple(tuple(sorted(m.items())) for m in result)
        if key not in (k for k, _ in found):
            found.append((key, result))
        if len(found) > 1:          # 결과가 갈린다 = 모호
            return None, 'ambiguous'
    if budget[0] <= 0:              # 탐색을 다 못 돌았으면 유일하다고 말할 수 없다
        return None, 'too_many'
    if not found:
        return None, 'no_split'
    return found[0][1], 'ok'


def _search_split(items, targets, budget):
    """items를 targets 금액에 맞게 나누는 모든 조합 (yield)"""
    if len(targets) == 1:
        # 마지막 칸은 조합을 볼 필요가 없다 — 남은 게 전부 들어가야 한다
        if abs(sum(i[2] for i in items) - targets[0]) < AMOUNT_TOLERANCE:
            yield [items]
        return
    target = targets[0]
    for size in range(len(items) + 1):
        for pick in combinations(range(len(items)), size):
            budget[0] -= 1
            if budget[0] <= 0:
                return
            if abs(sum(items[i][2] for i in pick) - target) >= AMOUNT_TOLERANCE:
                continue
            chosen = set(pick)
            rest = [it for j, it in enumerate(items) if j not in chosen]
            for tail in _search_split(rest, targets[1:], budget):
                yield [[items[i] for i in pick]] + tail


def _qty_map(group: list[tuple[int, int, float]]) -> dict[int, int]:
    out: dict[int, int] = {}
    for line_item, qty, _ico in group:
        if qty:
            out[line_item] = out.get(line_item, 0) + qty
    return out


def _partial_shipment_lines(invoiced_rows: pd.DataFrame,
                            so_lines: dict[int, tuple[str, int]]) -> dict[int, int]:
    """부분 출고 — 이 기간에 공장이 계산서를 끊은 PO 라인/수량 그대로.

    같은 라인이 여러 행으로 쪼개져 있을 수 있어(분할발주) 합산한다.
    `SO_국내`에 없는 PO 라인은 뺀다 (`sales_only`) — 매입만 발생한 건이다.
    """
    qty = pd.to_numeric(invoiced_rows['Item qty'], errors='coerce').fillna(0)
    grouped = qty.groupby(invoiced_rows['_line']).sum()
    return sales_only(
        {int(k): int(v) for k, v in grouped.items() if int(v) > 0}, so_lines)


def sales_only(qty_by_line: dict[int, int],
               so_lines: dict[int, tuple[str, int]]) -> dict[int, int]:
    """`SO_국내`에 없는 라인을 뺀다 — DN은 매출 장부다

    PO에만 있고 SO에 없는 라인은 **고객에게 판 것이 아니라 매입만 발생한 것**이다
    (2026-05 `SOD-2026-0188` L4 'De-cluch Gear Box Bushing 하부 가공' 450,000원:
    PO는 4라인인데 SO는 3라인). 실데이터 불변식 — DN 1,431행 중 SO에 없는 라인은 0건.

    금액 대조에서는 이 라인도 살려 둬야 한다. 공장은 그것까지 계산서를 끊으므로
    출고리스트 `계산서금액`에 포함돼 있다 (위 건: 14,610,800원에 450,000원이 들어 있다).
    빼는 건 **DN에 쓸 라인을 고를 때뿐**이다.
    """
    dropped = [li for li in qty_by_line if li not in so_lines]
    if dropped:
        logger.debug("SO에 없는 PO 라인 제외 (매입만 발생): %s", dropped)
    return {li: q for li, q in qty_by_line.items() if li in so_lines}


# === 인덱스 헬퍼 ===========================================================

def _so_lines_by_order(so_df: pd.DataFrame) -> dict[str, dict[int, tuple[str, int]]]:
    """SO_ID → {Line item: (Item name, Item qty)} — 취소/보류 라인은 뺀다

    `SO_국내.Status`는 파워쿼리 캐시라 출고 여부 판정에는 쓰면 안 되지만
    (그건 DN 수량으로 계산한다), **취소/보류만은 이 컬럼으로만 알 수 있다** —
    취소 건은 DN이 영영 안 생겨 수량으로 구분할 방법이 없다.

    PO 쪽 `Cancelled`와는 다른 얘기다. 판매가 취소돼도 공장은 이미 만든 것을
    계산서로 넘기기도 한다 (2026-07 `SOD-2026-0364` L5: SO는 `Cancelled`인데
    PO L5 'IP66 TEST 시료 값'은 `Invoiced P07` 1,000,000원). 그건 매입만 발생한 것이라
    DN에 들어가면 안 된다 — 실데이터 불변식: SO Cancelled 44라인 중 DN에 있는 것 0건.

    SO_ID 키 자체는 전 라인이 취소돼도 남긴다 — 'SO 없음'(=오타)과
    '기록할 라인 없음'(=전부 취소)은 사람에게 다른 얘기다.
    """
    out: dict[str, dict[int, tuple[str, int]]] = {}
    status = (so_df['Status'].astype(str).str.strip()
              if 'Status' in so_df.columns
              else pd.Series('', index=so_df.index))
    for (_, r), st in zip(so_df.iterrows(), status):
        so_id, line_item = r.get('SO_ID'), r.get('Line item')
        if pd.isna(so_id) or pd.isna(line_item):
            continue
        bucket = out.setdefault(str(so_id).strip(), {})
        if st in EXCLUDED_SO_STATUSES:
            logger.debug("SO 취소/보류 라인 제외: %s L%s (%s)", so_id, line_item, st)
            continue
        qty = pd.to_numeric(r.get('Item qty'), errors='coerce')
        if pd.isna(qty):
            continue
        bucket[int(line_item)] = (str(r.get('Item name') or ''), int(qty))
    return out


def _so_meta_by_order(so_df: pd.DataFrame) -> dict[str, dict[str, object]]:
    """SO_ID → 사업자번호/고객명/통화 (첫 행 기준)"""
    out: dict[str, dict[str, object]] = {}
    for _, r in so_df.iterrows():
        so_id = r.get('SO_ID')
        if pd.isna(so_id):
            continue
        key = str(so_id).strip()
        if key in out:
            continue
        out[key] = {
            'biz_no': r.get('Business registration number'),
            'customer_name': r.get('Customer name'),
            'currency': r.get('Currency'),
        }
    return out


def _po_by_order(po_df: pd.DataFrame, status_col: pd.Series | None
                 ) -> dict[str, pd.DataFrame]:
    """SO_ID → PO 행들 (`_status`, `_line` 정규화 열 추가)"""
    if not {'SO_ID', 'Line item'} <= set(po_df.columns):
        return {}
    df = po_df.copy()
    df['_status'] = (status_col if status_col is not None
                     else df.get('Status', pd.Series(dtype=object))
                     ).astype(str).str.strip()
    df['_line'] = pd.to_numeric(df['Line item'], errors='coerce')
    df = df[df['_line'].notna() & df['SO_ID'].notna()].copy()
    df['_line'] = df['_line'].astype(int)
    df['SO_ID'] = df['SO_ID'].astype(str).str.strip()
    return {so_id: grp for so_id, grp in df.groupby('SO_ID')}


def _existing_shipments(dn_df: pd.DataFrame) -> set[tuple[str, pd.Timestamp]]:
    """이미 기록된 (SO_ID, 출고일) — 멱등성 판정용"""
    if 'SO_ID' not in dn_df.columns or '출고일' not in dn_df.columns:
        return set()
    dates = pd.to_datetime(dn_df['출고일'], errors='coerce')
    return {
        (str(so_id).strip(), date.normalize())
        for so_id, date in zip(dn_df['SO_ID'], dates)
        if pd.notna(so_id) and pd.notna(date)
    }


ShipLog = dict[tuple[str, int], list[tuple[pd.Timestamp, int]]]


def _shipped_qty(dn_df: pd.DataFrame) -> ShipLog:
    """(SO_ID, Line item) → [(출고일, 수량), ...]

    누계를 미리 합치지 않고 날짜를 남긴다. 잔량은 **그 출고일 이전 분만** 빼야 하기 때문이다.
    미리 합치면 지난 달을 다시 돌릴 때 아직 일어나지도 않은 다음 달 출고가 잔량에서 빠져
    라인이 통째로 사라진다.
    """
    out: ShipLog = {}
    if not {'SO_ID', 'Line item', 'Qty', '출고일'} <= set(dn_df.columns):
        return out
    dates = pd.to_datetime(dn_df['출고일'], errors='coerce')
    for (_, r), date in zip(dn_df.iterrows(), dates):
        so_id, line_item = r.get('SO_ID'), r.get('Line item')
        qty = pd.to_numeric(r.get('Qty'), errors='coerce')
        if pd.isna(so_id) or pd.isna(line_item) or pd.isna(qty) or pd.isna(date):
            continue
        out.setdefault((str(so_id).strip(), int(line_item)), []).append(
            (date.normalize(), int(qty)))
    return out


def _prior_qty(log: ShipLog, so_id: str, line_item: int,
               before: pd.Timestamp) -> int:
    """해당 출고일 **이전**에 나간 수량 합계"""
    return sum(q for date, q in log.get((so_id, line_item), ())
               if date < before)


def _shipment_maps(dn_df: pd.DataFrame
                   ) -> dict[str, dict[pd.Timestamp, dict[int, int]]]:
    """SO_ID → {출고일: {라인: 수량}} — '같은 출고를 다른 날로 적었나' 판정용"""
    out: dict[str, dict[pd.Timestamp, dict[int, int]]] = {}
    if not {'SO_ID', 'Line item', 'Qty', '출고일'} <= set(dn_df.columns):
        return out
    dates = pd.to_datetime(dn_df['출고일'], errors='coerce')
    for (_, r), date in zip(dn_df.iterrows(), dates):
        so_id, line_item = r.get('SO_ID'), r.get('Line item')
        qty = pd.to_numeric(r.get('Qty'), errors='coerce')
        if pd.isna(so_id) or pd.isna(line_item) or pd.isna(qty) or pd.isna(date):
            continue
        bucket = out.setdefault(str(so_id).strip(), {}).setdefault(
            date.normalize(), {})
        line_item = int(line_item)
        bucket[line_item] = bucket.get(line_item, 0) + int(qty)
    return out


def _recorded_summary(maps: dict[str, dict[pd.Timestamp, dict[int, int]]],
                      so_id: str) -> str:
    """이 SO로 이미 기록된 출고일들 (없으면 빈 문자열)"""
    return ', '.join(f"{d:%Y-%m-%d}" for d in sorted(maps.get(so_id, {})))


def _recorded_date(maps: dict[str, dict[pd.Timestamp, dict[int, int]]],
                   so_id: str, qty_by_line: dict[int, int]) -> str:
    """라인·수량이 **똑같은** 출고가 다른 날짜로 이미 있으면 그 날짜

    '라인마다 누계가 이만큼 있다'로 보면 안 된다 — 분할출고에서 뒤 회차가 앞 회차
    수량에 가려 통째로 사라진다 (2026-07 `SOD-2026-0713`: 7/23 L1 1개를 8/6의 L1 2개가
    덮어 버렸다). 같은 출고를 사람이 하루 이틀 다른 날로 적은 경우만 잡아야 하므로
    **한 날짜의 라인·수량 조합이 정확히 일치**할 때만 건너뛴다.
    """
    for date, recorded in maps.get(so_id, {}).items():
        if recorded == qty_by_line:
            return f"{date:%Y-%m-%d}"
    return ''


def _year_from_delivery(delivery: pd.DataFrame) -> int:
    """DN_ID 연도 — 출고일에서 얻는다 (기간 폴더 연도와 어긋날 수 있어 실제 날짜 우선)"""
    dates = delivery[COL_SHIP_DATE].dropna()
    if len(dates):
        return int(dates.min().year)
    return dt.date.today().year
