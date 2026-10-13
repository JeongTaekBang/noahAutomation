"""
데이터 정합성 검사 — 0이어야 하는 것들
=====================================

DB(`noah_data.db`)를 읽어 **정상이라면 한 건도 없어야 하는 것**을 찾는다.
`check_data.py`(CLI)와 대시보드 'SO/DN 불일치' 카드가 같은 함수를 부른다 — 산식을
두 곳에 두면 한쪽만 고쳐져 "대시보드는 0건인데 CLI는 3건"이 된다 (`v_dn_revenue`로
DN 매출 산식을 모은 것과 같은 이유).

**처리할 일 목록은 여기 넣지 않는다.** 세금계산서 미발행·PO 확정 지연·EXW 미출고는
늘 몇 건씩 있는 게 정상이라, 섞으면 결과가 항상 빨갛고 결국 아무도 안 본다.
그런 것은 대시보드에 있다.

검사를 더할 때는 실데이터에서 몇 건이 걸리는지 먼저 본다. 'DN 단가 0원'은 후보였지만
2026-10 실측 12행이 전부 SO 단가도 0원(무상 라인)이라 넣지 않았다 — 걸리는 게 전부
정상이면 그건 검사가 아니라 잡음이다.

각 검사는 `sqlite3.Connection → 문제 행 DataFrame` 순수 함수다 (빈 DataFrame = 정상).
열 이름은 사람이 읽는 한글이다 — 콘솔과 결과 xlsx에 그대로 나간다.
"""

from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass
from typing import Callable

import pandas as pd

from po_generator.config import EXCLUDED_SO_STATUSES


# === SO ↔ DN 라인 (과다출고 / 단가 불일치) ==================================

def so_dn_line_anomalies(conn: sqlite3.Connection) -> pd.DataFrame:
    """SO ↔ DN 라인 단위 이상(불일치) 검출 — 대시보드 'SO/DN 불일치' 카드의 원본.

    각 (SO_ID, Line item) 단위로 두 가지 무결성 체크:
      • 과다출고: SUM(DN.Qty) > SO.[Item qty]
      • 단가 불일치: SO.[Sales Unit Price] × SUM(DN.Qty) ≠ SUM(DN.[Total Sales])
                    (abs ≥ 1 AND 상대 ≥ 0.1% — 반올림 노이즈 제거)
    둘 중 하나라도 해당되는 라인만 반환. 부분납품(DN 누계 < SO 수량)은 정상.
    원화/외화 비교는 각 통화 원본끼리 — FX 환산 노이즈 제거.
    취소 SO 라인은 대조하지 않는다 — 거기 DN이 있는 것 자체가 `dn_on_excluded_so`가 잡는 문제다.
    """
    df = pd.read_sql_query("""
        WITH so_lines AS (
            SELECT SO_ID,
                   CAST([Line item] AS INTEGER) AS line_item,
                   CAST([Sales Unit Price] AS REAL) AS so_unit_price,
                   CAST([Item qty] AS REAL)         AS so_qty,
                   COALESCE([Customer name], '')    AS customer_name,
                   COALESCE(Sector, '')             AS sector,
                   COALESCE([OS name], '')          AS os_name,
                   COALESCE([Item name], '')        AS item_name,
                   COALESCE(Currency, 'KRW')        AS currency,
                   '국내' AS market
            FROM so_domestic
            WHERE COALESCE(Status, '') != 'Cancelled'
            UNION ALL
            SELECT SO_ID,
                   CAST([Line item] AS INTEGER),
                   CAST([Sales Unit Price] AS REAL),
                   CAST([Item qty] AS REAL),
                   COALESCE([Customer name], ''),
                   COALESCE(Sector, ''),
                   COALESCE([OS name], ''),
                   COALESCE([Item name], ''),
                   COALESCE(Currency, 'USD'),
                   '해외'
            FROM so_export
            WHERE COALESCE(Status, '') != 'Cancelled'
        ),
        dn_lines AS (
            SELECT SO_ID,
                   CAST([Line item] AS INTEGER) AS line_item,
                   CAST(Qty AS REAL)            AS dn_qty,
                   CAST([Unit Price] AS REAL)   AS dn_unit_price,
                   CAST([Total Sales] AS REAL)  AS dn_amount,
                   '국내' AS market
            FROM dn_domestic
            UNION ALL
            SELECT SO_ID,
                   CAST([Line item] AS INTEGER),
                   CAST(Qty AS REAL),
                   CAST([Unit Price] AS REAL),
                   CAST([Total Sales] AS REAL),
                   '해외'
            FROM dn_export
        )
        SELECT s.SO_ID, s.line_item, s.market, s.customer_name, s.sector,
               s.os_name, s.item_name, s.currency,
               s.so_unit_price, s.so_qty,
               SUM(d.dn_qty)     AS dn_qty_total,
               SUM(d.dn_amount)  AS dn_amount_total,
               MIN(d.dn_unit_price) AS dn_unit_min,
               MAX(d.dn_unit_price) AS dn_unit_max,
               COUNT(*) AS dn_line_count
        FROM so_lines s
        INNER JOIN dn_lines d
          ON d.SO_ID = s.SO_ID
         AND d.line_item = s.line_item
         AND d.market = s.market
        GROUP BY s.SO_ID, s.line_item, s.market
    """, conn)

    if df.empty:
        return df

    for c in ("so_unit_price", "so_qty", "dn_qty_total", "dn_amount_total",
             "dn_unit_min", "dn_unit_max"):
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # 과다출고: SO 수량 입력된 라인에서 DN 누계 > SO qty (0.001 tolerance)
    has_so_qty = df["so_qty"].notna() & (df["so_qty"] > 0)
    over_qty = df["dn_qty_total"] - df["so_qty"]
    df["over_delivered"] = has_so_qty & (over_qty > 0.001)
    df["over_qty"] = over_qty.where(df["over_delivered"], 0)

    # 단가 불일치: SO 단가 입력된 라인에서 (expected vs actual) 임계값 초과
    has_so_unit = df["so_unit_price"].notna() & (df["so_unit_price"] > 0)
    df["expected_amount"] = df["so_unit_price"] * df["dn_qty_total"]
    df["diff"] = df["dn_amount_total"] - df["expected_amount"]
    df["abs_diff"] = df["diff"].abs()
    exp_abs = df["expected_amount"].abs()
    df["rel_diff"] = df["abs_diff"] / exp_abs.where(exp_abs > 0)
    df["price_mismatch"] = (
        has_so_unit & (df["abs_diff"] >= 1) & (df["rel_diff"] >= 0.001)
    )

    # 하나라도 이상이 있는 라인만 → 과다출고 먼저, 그 다음 금액 차이 큰 순
    result = df[df["over_delivered"] | df["price_mismatch"]].copy()
    if result.empty:
        return result
    result["_sort"] = (
        result["over_delivered"].astype(int) * 1e15
        + result["abs_diff"].fillna(0)
    )
    return result.sort_values("_sort", ascending=False).drop(columns="_sort").reset_index(drop=True)


_LINE_LABELS = {
    "market": "시장", "SO_ID": "SO_ID", "line_item": "Line",
    "customer_name": "고객", "item_name": "품목",
}


def find_over_delivery(conn: sqlite3.Connection) -> pd.DataFrame:
    df = so_dn_line_anomalies(conn)
    if df.empty:
        return pd.DataFrame()
    df = df[df["over_delivered"]].sort_values("over_qty", ascending=False)
    return df[[*_LINE_LABELS, "so_qty", "dn_qty_total", "over_qty"]].rename(columns={
        **_LINE_LABELS, "so_qty": "SO 수량", "dn_qty_total": "DN 누계", "over_qty": "초과",
    }).reset_index(drop=True)


def find_price_mismatch(conn: sqlite3.Connection) -> pd.DataFrame:
    df = so_dn_line_anomalies(conn)
    if df.empty:
        return pd.DataFrame()
    df = df[df["price_mismatch"]].sort_values("abs_diff", ascending=False)
    cols = ["currency", "so_unit_price", "dn_qty_total", "expected_amount",
            "dn_amount_total", "diff"]
    return df[[*_LINE_LABELS, *cols]].rename(columns={
        **_LINE_LABELS, "currency": "통화", "so_unit_price": "SO 단가",
        "dn_qty_total": "DN 누계", "expected_amount": "기대 금액",
        "dn_amount_total": "DN 금액", "diff": "차이",
    }).reset_index(drop=True)


# === DN → SO 참조 ==========================================================

# (시장, DN 테이블, SO 테이블). SO_ID 체계가 시장별로 따로라 국내 DN은 국내 SO에서만 찾는다.
_MARKETS = (("국내", "dn_domestic", "so_domestic"), ("해외", "dn_export", "so_export"))

_SAME_LINE = ("s.SO_ID = d.SO_ID "
              "AND CAST(s.[Line item] AS INTEGER) = CAST(d.[Line item] AS INTEGER)")


def _dn_line_columns(market: str) -> str:
    return (f"'{market}' AS 시장, d.DN_ID, d.SO_ID, "
            "CAST(d.[Line item] AS INTEGER) AS Line, "
            "d.[Customer name] AS 고객, d.Item AS 품목, CAST(d.Qty AS REAL) AS 수량")


def find_dn_without_so_line(conn: sqlite3.Connection) -> pd.DataFrame:
    """DN의 (SO_ID, Line item)이 SO 시트에 없는 라인"""
    parts = [
        f"SELECT {_dn_line_columns(market)}, SUBSTR(d.[출고일], 1, 10) AS 출고일 "
        f"FROM {dn} d WHERE NOT EXISTS (SELECT 1 FROM {so} s WHERE {_SAME_LINE})"
        for market, dn, so in _MARKETS
    ]
    return pd.read_sql_query(
        " UNION ALL ".join(parts) + " ORDER BY 1, 2, 3, 4", conn)


def find_dn_on_excluded_so(conn: sqlite3.Connection) -> pd.DataFrame:
    """SO Status가 Cancelled/Hold인 라인에 붙은 DN"""
    statuses = sorted(EXCLUDED_SO_STATUSES)
    marks = ", ".join("?" * len(statuses))
    parts = [
        f"SELECT {_dn_line_columns(market)}, TRIM(s.Status) AS [SO Status] "
        f"FROM {dn} d JOIN {so} s ON {_SAME_LINE} WHERE TRIM(s.Status) IN ({marks})"
        for market, dn, so in _MARKETS
    ]
    return pd.read_sql_query(
        " UNION ALL ".join(parts) + " ORDER BY 1, 2, 3, 4", conn,
        params=statuses * len(parts))


# === SO ↔ PO 참조 ==========================================================

# (시장, SO 테이블, PO 테이블). PO는 SO_ID로 SO와 이어진다 (라인 번호는 1:1이 아니다).
_SO_PO = (("국내", "so_domestic", "po_domestic"), ("해외", "so_export", "po_export"))


def _short_list(names: pd.Series) -> str:
    """품목 이름 앞 셋 + '외 N' — 수십 라인짜리 주문이 표를 밀어내지 않게"""
    names = [str(n) for n in names.dropna()]
    head = ", ".join(names[:3])
    return f"{head} 외 {len(names) - 3}" if len(names) > 3 else head


def find_so_without_po(conn: sqlite3.Connection) -> pd.DataFrame:
    """살아 있는 SO인데 PO 시트에 그 SO_ID가 한 줄도 없는 주문

    대시보드 '발주 커버리지'의 'PO 미등록'과 같은 정의다 — PO 행이 하나라도 있으면
    (Cancelled여도) 등록된 것으로 본다. 다만 대시보드는 발주할 일을 보는 화면이라 출고
    완료 건을 빼지만, 여기서는 **이미 나간 주문도 센다** — PO가 없으면 매입이 기록되지
    않은 것이고 매입대사에서야 드러난다.

    **라인 단위로는 보지 않는다.** PO는 부속·운임을 한 라인에 합쳐 적어 SO 라인과 1:1이
    아니다 — 2026-10 실측 SO 라인 중 같은 라인 번호의 PO가 없는 32라인은 거의 전부
    운임(FEDEX COST·운임비)·합산 부속(부싱 가공·ADAPTER)·할인·보증 라인이었다.

    SO와 PO는 대개 같은 날 들어온다 (동기화 이력 1,322주문 중 99.2%가 같은 동기화,
    나머지도 열흘 안) — 그래서 유예 기간을 두지 않는다. 걸리면 PO 입력이 밀린 것이다.
    """
    statuses = sorted(EXCLUDED_SO_STATUSES)
    marks = ", ".join("?" * len(statuses))
    parts = [
        f"SELECT '{market}' AS 시장, s.SO_ID, s.[Customer name] AS 고객, "
        "s.[Customer PO] AS [Customer PO], SUBSTR(s.[PO receipt date], 1, 10) AS 접수일, "
        "CAST(s.[Line item] AS INTEGER) AS line, s.[Item name] AS item "
        f"FROM {so} s WHERE COALESCE(TRIM(s.Status), '') NOT IN ({marks}) "
        f"AND NOT EXISTS (SELECT 1 FROM {po} p WHERE p.SO_ID = s.SO_ID)"
        for market, so, po in _SO_PO
    ]
    lines = pd.read_sql_query(" UNION ALL ".join(parts), conn,
                              params=statuses * len(parts))
    if lines.empty:
        return pd.DataFrame()

    out = (lines.sort_values(["시장", "SO_ID", "line"])
                .groupby(["시장", "SO_ID"], sort=False)
                .agg(고객=("고객", "first"), **{"Customer PO": ("Customer PO", "first")},
                     접수일=("접수일", "min"), **{"라인 수": ("line", "size")},
                     품목=("item", _short_list))
                .reset_index())
    return out.sort_values(["시장", "접수일", "SO_ID"], na_position="last").reset_index(drop=True)


def find_po_without_so(conn: sqlite3.Connection) -> pd.DataFrame:
    """PO에 적힌 SO_ID가 SO 시트에 없는 발주 — 'PO에 없는 SO'의 반대 방향

    SO를 취소 처리(Status=Cancelled) 대신 지웠거나, PO의 SO_ID를 잘못 적은 경우다.
    대시보드 발주 커버리지·마진은 SO에서 PO를 찾아가므로 이 발주는 어디에도 안 잡힌다.
    **취소된 PO도 센다** — SO 행을 지우면 그 주문이 취소됐다는 기록까지 같이 사라진다.
    SO_ID가 빈 PO도 여기 걸린다 (어느 주문의 발주인지 모른다).
    """
    parts = [
        f"SELECT '{market}' AS 시장, p.PO_ID, "
        "COALESCE(NULLIF(TRIM(p.SO_ID), ''), '(공란)') AS SO_ID, p.[Customer name] AS 고객, "
        "CAST(p.[Line item] AS INTEGER) AS line, p.[Item name] AS item, "
        "TRIM(COALESCE(p.Status, '')) AS status "
        f"FROM {po} p WHERE NOT EXISTS (SELECT 1 FROM {so} s WHERE s.SO_ID = p.SO_ID)"
        for market, so, po in _SO_PO
    ]
    lines = pd.read_sql_query(" UNION ALL ".join(parts), conn)
    if lines.empty:
        return pd.DataFrame()

    def _statuses(values: pd.Series) -> str:
        return ", ".join(sorted({v for v in values if v}))

    return (lines.sort_values(["시장", "PO_ID", "line"])
                 .groupby(["시장", "PO_ID", "SO_ID"], sort=False)
                 .agg(고객=("고객", "first"), **{"라인 수": ("line", "size")},
                      품목=("item", _short_list), Status=("status", _statuses))
                 .reset_index())


# === 한 DN에 거래처 둘 ======================================================

def _customer_key(market: str, value: object) -> str:
    """거래처 키 정규화 — 국내 사업자번호는 숫자만, 해외 고객코드는 공백·하이픈 제거.

    빈 값은 ''로 두고 세지 않는다 — 모르는 것을 '다른 거래처'로 치면 정상 DN이 걸린다
    (묶음 메일이 빈 사업자번호끼리 묶지 않는 것과 같은 규칙).
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value)
    return re.sub(r"\D", "", text) if market == "국내" else re.sub(r"[\s-]", "", text).upper()


def find_dn_mixed_customers(conn: sqlite3.Connection) -> pd.DataFrame:
    """한 DN_ID에 거래처가 둘 이상 — DN 번호 재사용

    거래처 키는 그 시장의 문서가 쓰는 것을 따른다:
      - 국내: `DN_국내.Business registration number` — 거래명세표 메일 차단
        (`create_ts.foreign_biz_numbers`)과 같은 열. SO 쪽으로 보면 안 된다:
        `SOD-2026-0339`(삼성전자 FOC)는 한 주문 안에서 라인마다 사업자번호가 달라
        멀쩡한 DN이 걸린다 (2026-10 실측)
      - 해외: DN에 거래처 키가 없어 `SO_해외`의 고객코드(이름은 `Business registration
        number`)로 본다 — OC 메일 조인키와 같다
    """
    frames = [
        pd.read_sql_query(
            "SELECT DN_ID, SO_ID, [Business registration number] AS raw, "
            "[Customer name] AS name FROM dn_domestic", conn).assign(시장="국내"),
        pd.read_sql_query(
            "SELECT d.DN_ID, d.SO_ID, s.[Business registration number] AS raw, "
            f"s.[Customer name] AS name FROM dn_export d JOIN so_export s ON {_SAME_LINE}",
            conn).assign(시장="해외"),
    ]
    df = pd.concat(frames, ignore_index=True)
    if df.empty:
        return pd.DataFrame()
    df["key"] = [_customer_key(m, v) for m, v in zip(df["시장"], df["raw"])]

    rows = []
    for (market, dn_id), group in df.groupby(["시장", "DN_ID"], sort=True):
        known = group[group["key"] != ""].drop_duplicates("key").sort_values("key")
        if len(known) < 2:
            continue
        rows.append({
            "시장": market,
            "DN_ID": dn_id,
            "거래처": " / ".join(f"{r.name} ({r.raw})" for r in known.itertuples()),
            "SO_ID": ", ".join(sorted(group["SO_ID"].dropna().unique())),
        })
    return pd.DataFrame(rows, columns=["시장", "DN_ID", "거래처", "SO_ID"])


# === 환율 없는 외화 DN ======================================================

def find_fx_missing(conn: sqlite3.Connection) -> pd.DataFrame:
    """선적했는데 그 달 환율이 `fx`에 없는 해외 DN — (통화, 선적월)별

    조건은 `v_dn_revenue`의 해외 CASE를 그대로 뒤집은 것이다: KRW가 아니고 선적월이
    있는데 환율 조인이 비면 뷰는 **오류 없이** 시트의 `Total Sales KRW`(수주시점 환율)로
    폴백한다. 통화가 공란이어도 같은 폴백을 탄다.
    """
    df = pd.read_sql_query("""
        SELECT COALESCE(NULLIF(TRIM(e.Currency), ''), '(공란)') AS 통화,
               SUBSTR(e.[선적일], 1, 7) AS 선적월,
               e.DN_ID,
               CAST(e.[Total Sales] AS REAL) AS amount
        FROM dn_export e
        LEFT JOIN fx f
            ON f.currency = e.Currency
           AND f.ym = SUBSTR(e.[선적일], 1, 7)
        WHERE COALESCE(e.Currency, '') != 'KRW'
          AND TRIM(COALESCE(e.[선적일], '')) != ''
          AND f.rate IS NULL
    """, conn)
    if df.empty:
        return pd.DataFrame()
    out = (df.groupby(["통화", "선적월"], sort=True)
             .agg(**{"행 수": ("DN_ID", "size"),
                     "외화 금액": ("amount", "sum"),
                     "DN_ID": ("DN_ID", lambda s: ", ".join(sorted(s.dropna().unique())))})
             .reset_index())
    return out[["통화", "선적월", "행 수", "외화 금액", "DN_ID"]]


# === 등록표 · 실행 ==========================================================

@dataclass(frozen=True)
class Check:
    key: str
    title: str
    why: str   # 걸리면 무슨 일이 생기고 무엇을 고치나 — 콘솔과 결과 파일에 그대로 나간다
    find: Callable[[sqlite3.Connection], pd.DataFrame]


CHECKS: tuple[Check, ...] = (
    Check("over_delivery", "SO/DN 과다출고",
          "DN 출고수량 누계가 SO 주문수량보다 많다. DN 수량 오기이거나 다른 라인 번호로 "
          "적어야 할 출고다.",
          find_over_delivery),
    Check("price_mismatch", "SO/DN 단가 불일치",
          "DN 금액이 SO 단가 × 출고수량과 다르다(1원·0.1% 미만은 무시). 매출이 SO와 다르게 "
          "잡히므로 DN 단가나 SO 단가 중 틀린 쪽을 고친다.",
          find_price_mismatch),
    Check("dn_without_so_line", "SO에 없는 DN 라인",
          "DN의 (SO_ID, Line item)이 SO 시트에 없다. DN 단가 XLOOKUP이 비고 Order Book에서 "
          "매출이 어느 주문에도 안 붙는다. SO_ID·Line item 오기를 고친다.",
          find_dn_without_so_line),
    Check("dn_on_excluded_so", "취소·보류 SO 라인의 DN",
          "SO Status가 Cancelled/Hold인 라인에 출고 기록이 있다. 실제로 나갔으면 SO Status를, "
          "아니면 DN 행을 고친다 (보류 전에 일부 나간 라인이면 그대로 둬도 된다).",
          find_dn_on_excluded_so),
    Check("so_without_po", "PO에 없는 SO",
          "SO는 있는데 PO 시트에 그 SO_ID가 한 줄도 없다. 공장 발주가 기록되지 않아 매입대사와 "
          "발주 커버리지에서 빠진다. PO를 등록하거나, 주문이 취소됐으면 SO Status를 Cancelled로 한다.",
          find_so_without_po),
    Check("po_without_so", "SO에 없는 PO",
          "PO에 적힌 SO_ID가 SO 시트에 없다. SO를 취소 처리 대신 지웠거나 PO의 SO_ID 오기다. "
          "발주 커버리지·마진은 SO에서 PO를 찾아가므로 이 발주가 빠진다. 주문이 취소됐으면 SO를 "
          "되살려 Status를 Cancelled로, 아니면 PO의 SO_ID를 고친다.",
          find_po_without_so),
    Check("dn_mixed_customers", "한 DN에 거래처 둘",
          "DN 번호를 재사용해 한 DN에 거래처가 둘 이상이다. 거래명세표·인보이스에 남의 "
          "품목·단가가 찍힌다. 시트에서 DN_ID를 나눈다.",
          find_dn_mixed_customers),
    Check("fx_missing", "환율 없는 외화 DN",
          "선적월 환율이 FX 시트에 없어 매출이 수주시점 환율(Total Sales KRW)로 잡혔다. "
          "FX 시트에 그 달 환율을 넣고 DB Sync를 다시 돌린다.",
          find_fx_missing),
)


@dataclass(frozen=True)
class CheckResult:
    check: Check
    rows: pd.DataFrame | None = None   # None = 검사 자체가 실패
    error: str | None = None

    @property
    def ok(self) -> bool:
        return self.error is None and self.rows is not None and self.rows.empty

    @property
    def count(self) -> int:
        return 0 if self.rows is None else len(self.rows)


def run_checks(conn: sqlite3.Connection,
               checks: tuple[Check, ...] = CHECKS) -> list[CheckResult]:
    """검사를 전부 돌린다. 한 검사가 죽어도 나머지는 돈다.

    실패한 검사는 `rows=None` + `error`로 남긴다 — 0건과 구분돼야 한다
    ("문제 없음"과 "확인 못 함"은 다른 사건이다).
    """
    results = []
    for check in checks:
        try:
            rows = check.find(conn)
        except Exception as e:  # 컬럼 개명·테이블 누락 등 — 무엇이든 '실패'로 보고
            results.append(CheckResult(check, error=f"{type(e).__name__}: {e}"))
        else:
            results.append(CheckResult(check, rows=rows.reset_index(drop=True)))
    return results
