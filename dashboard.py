"""
NOAH 대시보드
============
수주/출고 현황, 제품/섹터/고객 분석, Order Book(백로그) 등 핵심 KPI.
데이터 소스: noah_data.db (SQLite, sync_db.py로 동기화)

Usage:
    streamlit run dashboard.py
"""

import calendar
import logging
import sqlite3
from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import streamlit as st

from po_generator.config import DB_FILE
from po_generator.db_schema import get_sync_metadata

logger = logging.getLogger(__name__)

# ── 상수 ──────────────────────────────────────────────────────
_TODAY = datetime.today()
_TODAY_DATE = _TODAY.date()
_THIS_MONTH = _TODAY.strftime("%Y-%m")
_PREV_MONTH = (_TODAY.replace(day=1) - timedelta(days=1)).strftime("%Y-%m")
_WEEK_START = (_TODAY - timedelta(days=_TODAY.weekday())).date()
_WEEK_END = _WEEK_START + timedelta(days=6)

C_INPUT = "#1f77b4"
C_OUTPUT = "#ff7f0e"
C_ENDING = "#2ca02c"
C_DANGER = "#d62728"
C_PURPLE = "#9467bd"


# ═══════════════════════════════════════════════════════════════
# 포맷 유틸
# ═══════════════════════════════════════════════════════════════
def fmt_krw(v: float) -> str:
    """KRW 포맷 (억/만 자동)"""
    if pd.isna(v) or v == 0:
        return "₩0"
    a, s = abs(v), "-" if v < 0 else ""
    if a >= 1e8:
        return f"{s}₩{a / 1e8:,.1f}억"
    if a >= 1e4:
        return f"{s}₩{a / 1e4:,.0f}만"
    return f"{s}₩{a:,.0f}"


def fmt_num(v) -> str:
    """숫자 콤마 포맷 (테이블용)"""
    if pd.isna(v) or v == 0:
        return "0"
    return f"{v:,.0f}"


def fmt_qty(v) -> str:
    return "0" if pd.isna(v) else f"{int(v):,}"


def fmt_date(v) -> str:
    if pd.isna(v):
        return ""
    return v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else str(v)[:10]


# ═══════════════════════════════════════════════════════════════
# DB 연결
# ═══════════════════════════════════════════════════════════════
def _conn():
    return sqlite3.connect(str(DB_FILE)) if DB_FILE.exists() else None


# ═══════════════════════════════════════════════════════════════
# 로더 에러 수집 — session_state 기반 (캐시 히트 시에도 유지)
# ═══════════════════════════════════════════════════════════════
def _record_load_error(source: str, err: Exception) -> None:
    """로더 실패를 session_state에 수집 — 페이지 렌더 시 배너로 표시"""
    if "_load_errors" not in st.session_state:
        st.session_state["_load_errors"] = []
    msg = f"{source}: {type(err).__name__}: {err}"
    if msg not in st.session_state["_load_errors"]:
        st.session_state["_load_errors"].append(msg)


def _show_load_errors() -> None:
    """수집된 로더 에러가 있으면 경고 배너 표시 후 초기화"""
    errors = st.session_state.get("_load_errors", [])
    if errors:
        st.warning(
            "일부 데이터 로드 실패\n\n"
            + "\n".join(f"- {e}" for e in errors)
        )
        st.session_state["_load_errors"] = []


# ═══════════════════════════════════════════════════════════════
# 데이터 로더 (캐시 5분)
# ═══════════════════════════════════════════════════════════════
@st.cache_data(ttl=300)
def load_so() -> pd.DataFrame:
    """SO 국내+해외 통합 (Status, EXW NOAH 포함)"""
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    try:
        df = pd.read_sql_query("""
            SELECT SO_ID,
                   [Customer name] AS customer_name,
                   [Item name]     AS item_name,
                   [OS name]       AS os_name,
                   CAST([Line item] AS INTEGER) AS line_item,
                   CAST([Item qty] AS REAL)     AS qty,
                   CAST([Sales amount] AS REAL) AS amount_krw,
                   Period  AS period,
                   [Model code] AS model_code,
                   Sector  AS sector,
                   [Expected delivery date] AS delivery_date,
                   [Requested delivery date] AS requested_date,
                   [EXW NOAH] AS exw_noah,
                   [PO receipt date] AS po_receipt_date,
                   COALESCE(Status, '') AS status,
                   COALESCE([Customer PO], '') AS customer_po,
                   '' AS incoterms,
                   '' AS shipping_method,
                   '국내' AS market
            FROM so_domestic
            WHERE COALESCE(Status, '') != 'Cancelled'
              AND Period IS NOT NULL AND TRIM(Period) != ''
            UNION ALL
            SELECT SO_ID, [Customer name], [Item name], [OS name],
                   CAST([Line item] AS INTEGER),
                   CAST([Item qty] AS REAL),
                   CAST([Sales amount KRW] AS REAL),
                   Period, [Model code], Sector,
                   [Expected delivery date],
                   [Requested delivery date],
                   [EXW NOAH],
                   [PO receipt date],
                   COALESCE(Status, ''),
                   COALESCE([Customer PO], ''),
                   COALESCE(Incoterms, ''),
                   COALESCE([Shipping method], ''),
                   '해외'
            FROM so_export
            WHERE COALESCE(Status, '') != 'Cancelled'
              AND Period IS NOT NULL AND TRIM(Period) != ''
        """, conn)
    except Exception as e:
        logger.warning("데이터 로드 실패: %s", e)
        _record_load_error("SO", e)
        return pd.DataFrame()
    finally:
        conn.close()
    df["delivery_date"] = pd.to_datetime(df["delivery_date"], errors="coerce")
    df["requested_date"] = pd.to_datetime(df["requested_date"], errors="coerce")
    df.loc[df["requested_date"].dt.year <= 1900, "requested_date"] = pd.NaT
    df["exw_noah"] = pd.to_datetime(df["exw_noah"], errors="coerce")
    df.loc[df["exw_noah"].dt.year <= 1900, "exw_noah"] = pd.NaT
    df["po_receipt_date"] = pd.to_datetime(df["po_receipt_date"], errors="coerce")
    for c in ("os_name", "sector", "customer_name"):
        df[c] = df[c].fillna("")
    return df


@st.cache_data(ttl=300)
def load_dn() -> pd.DataFrame:
    """DN 국내+해외 통합 (매출 기준: 국내=출고일, 해외=선적일)"""
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    try:
        df = pd.read_sql_query("""
            SELECT DN_ID, SO_ID,
                   CAST([Line item] AS INTEGER) AS line_item,
                   CAST(Qty AS REAL)            AS qty,
                   CAST([Total Sales] AS REAL)  AS amount_krw,
                   [출고일] AS dispatch_date, '국내' AS market
            FROM dn_domestic
            WHERE [출고일] IS NOT NULL AND TRIM(COALESCE([출고일], '')) != ''
            UNION ALL
            SELECT DN_ID, SO_ID,
                   CAST([Line item] AS INTEGER),
                   CAST(Qty AS REAL),
                   CAST([Total Sales KRW] AS REAL),
                   [선적일], '해외'
            FROM dn_export
            WHERE [선적일] IS NOT NULL AND TRIM(COALESCE([선적일], '')) != ''
        """, conn)
    except Exception as e:
        logger.warning("데이터 로드 실패: %s", e)
        _record_load_error("DN", e)
        return pd.DataFrame()
    finally:
        conn.close()
    df["dispatch_date"] = pd.to_datetime(df["dispatch_date"], errors="coerce")
    df["dispatch_month"] = df["dispatch_date"].dt.strftime("%Y-%m")
    return df


@st.cache_data(ttl=300)
def load_dn_export_shipping() -> pd.DataFrame:
    """해외 DN 선적 파이프라인 (공장 출고 이후 전체 현황)"""
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    try:
        df = pd.read_sql_query("""
            SELECT DN_ID, SO_ID,
                   [Customer name] AS customer_name,
                   [Item] AS item_name,
                   CAST(Qty AS REAL) AS qty,
                   CAST([Total Sales KRW] AS REAL) AS amount_krw,
                   [출고일]       AS factory_date,
                   [공장 픽업일]   AS pickup_date,
                   [선적 예정일]   AS expected_ship_date,
                   [선적일]       AS ship_date,
                   [B/L]          AS bl_no,
                   [운송 업체]     AS carrier
            FROM dn_export
            WHERE [출고일] IS NOT NULL AND TRIM(COALESCE([출고일], '')) != ''
        """, conn)
    except Exception as e:
        logger.warning("데이터 로드 실패: %s", e)
        _record_load_error("DN 해외선적", e)
        return pd.DataFrame()
    finally:
        conn.close()
    for c in ("factory_date", "pickup_date", "expected_ship_date", "ship_date"):
        df[c] = pd.to_datetime(df[c], errors="coerce")
    return df


@st.cache_data(ttl=300)
def load_po_status() -> pd.DataFrame:
    """PO 국내+해외 SO_ID별 Status 집계 (Invoiced 여부 판단용)."""
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    try:
        df = pd.read_sql_query("""
            SELECT SO_ID, COALESCE(Status, '') AS po_status
            FROM po_domestic
            UNION ALL
            SELECT SO_ID, COALESCE(Status, '')
            FROM po_export
        """, conn)
    except Exception as e:
        logger.warning("데이터 로드 실패: %s", e)
        _record_load_error("PO Status", e)
        return pd.DataFrame()
    finally:
        conn.close()
    return df


@st.cache_data(ttl=300)
def load_po_detail() -> pd.DataFrame:
    """PO SO_ID 단위 집계 — 발주 커버리지·마진 분석용.

    PO line_item은 SO line_item과 1:1 대응하지 않음 (본체+부속을 합쳐서 발주하는 경우 등).
    따라서 SO_ID 단위로 집계합니다.
    """
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    try:
        df = pd.read_sql_query("""
            SELECT p.SO_ID,
                   SUM(CAST(p.[Item qty] AS REAL))  AS po_qty,
                   SUM(CAST(p.[Total ICO] AS REAL))  AS po_total_ico,
                   GROUP_CONCAT(DISTINCT COALESCE(p.Status, '')) AS po_statuses,
                   GROUP_CONCAT(DISTINCT p.PO_ID) AS po_ids,
                   (SELECT GROUP_CONCAT(DISTINCT o.PO_ID)
                    FROM po_domestic o
                    WHERE o.SO_ID = p.SO_ID AND COALESCE(o.Status, '') = 'Open'
                   ) AS open_po_ids,
                   MIN(p.[공장 발주 날짜]) AS factory_order_date,
                   MIN(NULLIF(p.[공장 EXW date], '')) AS factory_exw,
                   '국내' AS market
            FROM po_domestic p
            WHERE COALESCE(p.Status, '') != 'Cancelled'
            GROUP BY p.SO_ID
            UNION ALL
            SELECT p.SO_ID,
                   SUM(CAST(p.[Item qty] AS REAL))  AS po_qty,
                   SUM(CAST(p.[Total ICO] AS REAL))  AS po_total_ico,
                   GROUP_CONCAT(DISTINCT COALESCE(p.Status, '')) AS po_statuses,
                   GROUP_CONCAT(DISTINCT p.PO_ID) AS po_ids,
                   (SELECT GROUP_CONCAT(DISTINCT o.PO_ID)
                    FROM po_export o
                    WHERE o.SO_ID = p.SO_ID AND COALESCE(o.Status, '') = 'Open'
                   ) AS open_po_ids,
                   MIN(p.[공장 발주 날짜]) AS factory_order_date,
                   MIN(NULLIF(p.[공장 EXW date], '')) AS factory_exw,
                   '해외' AS market
            FROM po_export p
            WHERE COALESCE(p.Status, '') != 'Cancelled'
            GROUP BY p.SO_ID
        """, conn)
    except Exception as e:
        logger.warning("데이터 로드 실패: %s", e)
        _record_load_error("PO Detail", e)
        return pd.DataFrame()
    finally:
        conn.close()
    df["po_qty"] = pd.to_numeric(df["po_qty"], errors="coerce").fillna(0)
    df["po_total_ico"] = pd.to_numeric(df["po_total_ico"], errors="coerce").fillna(0)
    df["po_statuses"] = df["po_statuses"].fillna("")
    df["po_ids"] = df["po_ids"].fillna("")
    df["open_po_ids"] = df["open_po_ids"].fillna("")
    df["factory_order_date"] = pd.to_datetime(df["factory_order_date"], errors="coerce")
    df["factory_exw"] = pd.to_datetime(df["factory_exw"], errors="coerce")
    return df


@st.cache_data(ttl=300)
def load_po_sent_pending() -> pd.DataFrame:
    """PO Status='Sent' (확정 대기) 건 — 공장 발주 날짜 포함."""
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    try:
        df = pd.read_sql_query("""
            SELECT SO_ID,
                   PO_ID,
                   COALESCE([Item name], '') AS item_name,
                   COALESCE([공장 발주 날짜], '') AS order_date,
                   COALESCE([공장 EXW date], '') AS factory_exw,
                   CAST([Item qty] AS REAL) AS po_qty,
                   CAST([Total ICO] AS REAL) AS po_total_ico,
                   COALESCE([NOAH O.C No.], '') AS noah_oc,
                   '국내' AS market
            FROM po_domestic
            WHERE COALESCE(Status, '') = 'Sent'
            UNION ALL
            SELECT SO_ID,
                   PO_ID,
                   COALESCE([Item name], '') AS item_name,
                   COALESCE([공장 발주 날짜], '') AS order_date,
                   COALESCE([공장 EXW date], '') AS factory_exw,
                   CAST([Item qty] AS REAL) AS po_qty,
                   CAST([Total ICO] AS REAL) AS po_total_ico,
                   COALESCE([NOAH O.C No.], '') AS noah_oc,
                   '해외' AS market
            FROM po_export
            WHERE COALESCE(Status, '') = 'Sent'
        """, conn)
    except Exception as e:
        logger.warning("PO sent pending 로드 실패: %s", e)
        _record_load_error("PO Sent Pending", e)
        return pd.DataFrame()
    finally:
        conn.close()
    if not df.empty:
        df["po_qty"] = pd.to_numeric(df["po_qty"], errors="coerce").fillna(0)
        df["po_total_ico"] = pd.to_numeric(df["po_total_ico"], errors="coerce").fillna(0)
        df["order_date"] = pd.to_datetime(df["order_date"], errors="coerce")
        df["factory_exw"] = pd.to_datetime(df["factory_exw"], errors="coerce")
    return df


@st.cache_data(ttl=300)
def load_po_exw_pending() -> pd.DataFrame:
    """PO 공장 EXW date 경과 & 미Invoiced 라인 — line item 단위."""
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    try:
        df = pd.read_sql_query("""
            SELECT SO_ID, PO_ID,
                   CAST([Line item] AS INTEGER) AS line_item,
                   COALESCE([Item name], '') AS item_name,
                   CAST([Item qty] AS REAL) AS po_qty,
                   CAST([Total ICO] AS REAL) AS po_total_ico,
                   COALESCE([공장 발주 날짜], '') AS order_date,
                   COALESCE([공장 EXW date], '') AS factory_exw,
                   COALESCE(Status, '') AS po_status,
                   COALESCE([NOAH O.C No.], '') AS noah_oc,
                   '국내' AS market
            FROM po_domestic
            WHERE COALESCE([공장 EXW date], '') != ''
              AND NOT COALESCE(Status, '') LIKE 'Invoiced%'
              AND COALESCE(Status, '') != 'Cancelled'
            UNION ALL
            SELECT SO_ID, PO_ID,
                   CAST([Line item] AS INTEGER),
                   COALESCE([Item name], ''),
                   CAST([Item qty] AS REAL),
                   CAST([Total ICO] AS REAL),
                   COALESCE([공장 발주 날짜], ''),
                   COALESCE([공장 EXW date], ''),
                   COALESCE(Status, ''),
                   COALESCE([NOAH O.C No.], ''),
                   '해외'
            FROM po_export
            WHERE COALESCE([공장 EXW date], '') != ''
              AND NOT COALESCE(Status, '') LIKE 'Invoiced%'
              AND COALESCE(Status, '') != 'Cancelled'
        """, conn)
    except Exception as e:
        logger.warning("PO EXW pending 로드 실패: %s", e)
        _record_load_error("PO EXW Pending", e)
        return pd.DataFrame()
    finally:
        conn.close()
    if not df.empty:
        df["po_qty"] = pd.to_numeric(df["po_qty"], errors="coerce").fillna(0)
        df["po_total_ico"] = pd.to_numeric(df["po_total_ico"], errors="coerce").fillna(0)
        df["order_date"] = pd.to_datetime(df["order_date"], errors="coerce")
        df["factory_exw"] = pd.to_datetime(df["factory_exw"], errors="coerce")
    return df


@st.cache_data(ttl=300)
def load_dn_tax_pending() -> pd.DataFrame:
    """국내 DN 세금계산서 미발행 건 — 출고 완료 but 세금계산서/선수금 세금계산서 모두 미발행."""
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    try:
        df = pd.read_sql_query("""
            SELECT DN_ID, SO_ID, [Customer name] AS customer_name,
                   [Item] AS item_name,
                   CAST([Line item] AS INTEGER) AS line_item,
                   CAST(Qty AS REAL) AS qty,
                   CAST([Total Sales] AS REAL) AS amount_krw,
                   [출고일] AS dispatch_date
            FROM dn_domestic
            WHERE [출고일] IS NOT NULL AND TRIM(COALESCE([출고일], '')) != ''
              AND (TRIM(COALESCE([세금계산서 발행일], '')) = '' OR [세금계산서 발행일] IS NULL)
              AND UPPER(TRIM(COALESCE([세금계산서 발행일], ''))) != 'N/A'
              AND ([선수금 세금계산서 발행일] IS NULL
                   OR TRIM(COALESCE([선수금 세금계산서 발행일], '')) = ''
                   OR UPPER(TRIM(COALESCE([선수금 세금계산서 발행일], ''))) = 'N/A')
              AND CAST(COALESCE([Total Sales], 0) AS REAL) > 0
        """, conn)
    except Exception as e:
        logger.warning("데이터 로드 실패: %s", e)
        _record_load_error("DN 세금계산서", e)
        return pd.DataFrame()
    finally:
        conn.close()
    df["dispatch_date"] = pd.to_datetime(df["dispatch_date"], errors="coerce")
    return df


@st.cache_data(ttl=300)
def load_backlog() -> pd.DataFrame:
    """현재 백로그 (Ending > 0) — order_book_backlog.sql 패턴"""
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    try:
        df = pd.read_sql_query("""
        WITH
        so_combined AS (
            SELECT SO_ID, [Customer name] AS customer_name,
                   [OS name] AS os_name,
                   CAST([Line item] AS INTEGER) AS line_item,
                   CAST([Item qty] AS REAL) AS qty,
                   CAST([Sales amount] AS REAL) AS amount,
                   [Model code] AS model_code, Sector AS sector,
                   [Expected delivery date] AS delivery_date, '국내' AS market
            FROM so_domestic
            WHERE COALESCE(Status, '') != 'Cancelled'
              AND Period IS NOT NULL AND TRIM(Period) != ''
            UNION ALL
            SELECT SO_ID, [Customer name], [OS name],
                   CAST([Line item] AS INTEGER),
                   CAST([Item qty] AS REAL),
                   CAST([Sales amount KRW] AS REAL),
                   [Model code], Sector, [Expected delivery date], '해외'
            FROM so_export
            WHERE COALESCE(Status, '') != 'Cancelled'
              AND Period IS NOT NULL AND TRIM(Period) != ''
        ),
        dn_combined AS (
            SELECT SO_ID, CAST([Line item] AS INTEGER) AS line_item,
                   CAST(Qty AS REAL) AS out_qty,
                   CAST([Total Sales] AS REAL) AS out_amt
            FROM dn_domestic
            WHERE [출고일] IS NOT NULL AND TRIM(COALESCE([출고일], '')) != ''
            UNION ALL
            SELECT SO_ID, CAST([Line item] AS INTEGER),
                   CAST(Qty AS REAL), CAST([Total Sales KRW] AS REAL)
            FROM dn_export
            WHERE [선적일] IS NOT NULL AND TRIM(COALESCE([선적일], '')) != ''
        ),
        events AS (
            SELECT SO_ID, customer_name, os_name, line_item,
                   model_code, sector, delivery_date, market,
                   qty AS in_qty, amount AS in_amt,
                   0 AS out_qty, 0 AS out_amt
            FROM so_combined
            UNION ALL
            SELECT d.SO_ID,
                   COALESCE(s.customer_name, 'UNKNOWN') AS customer_name,
                   COALESCE(s.os_name, 'UNKNOWN')       AS os_name,
                   d.line_item,
                   COALESCE(s.model_code, '')            AS model_code,
                   COALESCE(s.sector, '')                AS sector,
                   COALESCE(s.delivery_date, '')         AS delivery_date,
                   COALESCE(s.market, '')                AS market,
                   0, 0, d.out_qty, d.out_amt
            FROM dn_combined d
            LEFT JOIN so_combined s
              ON d.SO_ID = s.SO_ID AND d.line_item = s.line_item
        )
        SELECT SO_ID, os_name, delivery_date,
               MIN(customer_name) AS customer_name,
               MIN(market)        AS market,
               MIN(sector)        AS sector,
               GROUP_CONCAT(DISTINCT model_code) AS model_code,
               SUM(in_qty  - out_qty) AS ending_qty,
               SUM(in_amt  - out_amt) AS ending_amount
        FROM events
        GROUP BY SO_ID, os_name, delivery_date
        HAVING SUM(in_amt - out_amt) > 0
        ORDER BY market, SO_ID, os_name
        """, conn)
    except Exception as e:
        logger.warning("데이터 로드 실패: %s", e)
        _record_load_error("Backlog", e)
        return pd.DataFrame()
    finally:
        conn.close()
    df["delivery_date"] = pd.to_datetime(df["delivery_date"], errors="coerce")
    return df


@st.cache_data(ttl=300)
def load_order_book() -> pd.DataFrame:
    """월별 Order Book (order_book.sql 실행) — Start/Input/Output/Ending"""
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    from pathlib import Path
    sql_path = Path(__file__).parent / "sql" / "order_book.sql"
    if not sql_path.exists():
        return pd.DataFrame()
    try:
        sql = sql_path.read_text(encoding="utf-8")
        df = pd.read_sql_query(sql, conn)
    except Exception as e:
        logger.warning("Order Book SQL 실행 실패: %s", e)
        _record_load_error("Order Book", e)
        return pd.DataFrame()
    finally:
        conn.close()
    return df


@st.cache_data(ttl=300)
def load_sync_meta() -> dict:
    conn = _conn()
    if not conn:
        return {}
    try:
        return get_sync_metadata(conn)
    except Exception as e:
        _record_load_error("Sync Meta", e)
        return {}
    finally:
        conn.close()


@st.cache_data(ttl=300)
def load_snapshot_meta() -> pd.DataFrame:
    conn = _conn()
    if not conn:
        return pd.DataFrame()
    try:
        return pd.read_sql_query(
            "SELECT * FROM ob_snapshot_meta WHERE is_active=1 ORDER BY period DESC",
            conn,
        )
    except Exception as e:
        logger.warning("데이터 로드 실패: %s", e)
        _record_load_error("Snapshot Meta", e)
        return pd.DataFrame()
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════════════
# 헬퍼
# ═══════════════════════════════════════════════════════════════
def enrich_dn(dn: pd.DataFrame, so: pd.DataFrame) -> pd.DataFrame:
    """DN에 SO 메타데이터(고객, 섹터, 품목) 추가"""
    if dn.empty or so.empty:
        return dn
    meta = so[["SO_ID", "line_item", "customer_name", "sector", "os_name"]].drop_duplicates(
        subset=["SO_ID", "line_item"]
    )
    return dn.merge(meta, on=["SO_ID", "line_item"], how="left")


def filt(df, market, sectors, customers,
         *, period_col="period", year=None, month=None):
    """사이드바 필터 적용"""
    if df.empty:
        return df
    f = df
    if market != "전체" and "market" in f.columns:
        f = f[f["market"] == market]
    if sectors and "sector" in f.columns:
        f = f[f["sector"].isin(sectors)]
    if customers and "customer_name" in f.columns:
        f = f[f["customer_name"].isin(customers)]
    if period_col and period_col in f.columns:
        if year and year != "전체":
            f = f[f[period_col].astype(str).str.startswith(year)]
            if month and month != "전체":
                f = f[f[period_col] == f"{year}-{month}"]
    return f


def _status_icon(status: str, overdue: bool = False) -> str:
    """상태 아이콘"""
    if overdue and status not in ("출고 완료",):
        return f"🔴 {status or '미출고'}"
    m = {"출고 완료": "🟢", "부분 출고": "🟡", "공장 출고": "🔵", "미출고": "⚪"}
    return f"{m.get(status, '⚪')} {status or '미출고'}"


def _render_cards(items: list[dict], cols_per_row: int = 3):
    """카드 격자 렌더 — items = [{"title": "🔴 ND-001 고객A", "lines": [...]}]"""
    for i in range(0, len(items), cols_per_row):
        cols = st.columns(cols_per_row)
        for j, item in enumerate(items[i:i + cols_per_row]):
            with cols[j]:
                with st.container(border=True):
                    st.markdown(item["title"])
                    for line in item["lines"]:
                        st.caption(line)


# 지연 버킷 (PO 확정 지연 / EXW 완료 미출고 공통)
_OVERDUE_BUCKETS: list[tuple[int | None, str, str]] = [
    (7,    "🟡", "~7일"),
    (14,   "🟠", "8~14일"),
    (30,   "🔴", "15~30일"),
    (None, "⚫", "30일+"),
]


def _assign_bucket(days: int) -> tuple[int, str, str]:
    """경과일 → (sort_key, icon, label)"""
    for i, (threshold, icon, label) in enumerate(_OVERDUE_BUCKETS):
        if threshold is None or days <= threshold:
            return (i, icon, label)
    return (len(_OVERDUE_BUCKETS) - 1, "⚫", "30일+")


def _render_bucketed_cards(items_with_bucket: list[dict], cols_per_row: int = 2):
    """버킷별 그룹 헤더 + 카드 렌더링.

    items_with_bucket: [{"bucket_key": int, "bucket_icon": str, "bucket_label": str, "title": ..., "lines": [...]}]
    """
    if not items_with_bucket:
        return
    items_with_bucket.sort(key=lambda x: (-x["bucket_key"], x.get("days", 0)))
    # 심각도 높은 버킷부터 (역순)
    from itertools import groupby as _groupby
    for bkey, grp in _groupby(
        sorted(items_with_bucket, key=lambda x: x["bucket_key"], reverse=True),
        key=lambda x: (x["bucket_key"], x["bucket_icon"], x["bucket_label"]),
    ):
        grp_list = list(grp)
        _, icon, label = bkey
        st.caption(f"{icon} **{label}** ({len(grp_list)}건)")
        _render_cards(grp_list, cols_per_row=cols_per_row)


# ═══════════════════════════════════════════════════════════════
# 메인
# ═══════════════════════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="NOAH 대시보드", page_icon="📊",
        layout="wide", initial_sidebar_state="expanded",
    )

    if not DB_FILE.exists():
        st.error("데이터베이스 파일이 없습니다.")
        st.code("python sync_db.py", language="bash")
        st.info("위 명령어로 Excel → SQLite 동기화를 먼저 수행하세요.")
        return

    # ── 테마 전환 (시스템 테마 감지 → 토글로 반대 테마) ──
    try:
        sys_dark = st.context.theme.type == "dark"
    except AttributeError:
        sys_dark = st.get_option("theme.base") != "light"
    toggle_label = ":sunny: Light Mode" if sys_dark else ":crescent_moon: Dark Mode"
    override = st.sidebar.toggle(toggle_label, value=False, key="theme_toggle")
    want_dark = (not sys_dark) if override else sys_dark

    _DARK_CSS = """
    <style>
    .stApp, [data-testid="stAppViewContainer"] { background-color: #0e1117 !important; color: #fafafa !important; }
    [data-testid="stSidebar"], [data-testid="stSidebar"] > div { background-color: #1a1d24 !important; color: #fafafa !important; }
    [data-testid="stHeader"] { background-color: #0e1117 !important; }
    [data-testid="stMetric"], [data-testid="stMetricValue"],
    [data-testid="stMetricLabel"], [data-testid="stMetricDelta"] { color: #fafafa !important; }
    .stMarkdown, .stMarkdown p, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4, .stMarkdown li,
    label, .stSelectbox label, .stMultiSelect label, .stRadio label, [data-testid="stWidgetLabel"] { color: #fafafa !important; }
    [data-testid="stExpander"] { border-color: #333 !important; }
    [data-testid="stExpander"] summary { color: #fafafa !important; }
    .stDataFrame, .stTable { color: #fafafa !important; }
    .stTabs [data-baseweb="tab-list"] button { color: #ccc !important; }
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] { color: #fafafa !important; }
    hr { border-color: #333 !important; }
    </style>
    """
    _LIGHT_CSS = """
    <style>
    /* 메인 배경 */
    .stApp, [data-testid="stAppViewContainer"] { background-color: #ffffff !important; color: #1a1a1a !important; }
    /* 사이드바 */
    [data-testid="stSidebar"], [data-testid="stSidebar"] > div { background-color: #f0f2f6 !important; color: #1a1a1a !important; }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stRadio label,
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"],
    [data-testid="stSidebar"] [data-baseweb="radio"] label,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] p { color: #1a1a1a !important; }
    /* 헤더 */
    [data-testid="stHeader"] { background-color: #ffffff !important; }
    /* metric */
    [data-testid="stMetric"], [data-testid="stMetricValue"],
    [data-testid="stMetricLabel"], [data-testid="stMetricDelta"] { color: #1a1a1a !important; }
    /* 텍스트 전체 */
    .stMarkdown, .stMarkdown p, .stMarkdown h1, .stMarkdown h2, .stMarkdown h3, .stMarkdown h4, .stMarkdown li,
    label, .stSelectbox label, .stMultiSelect label, .stRadio label,
    [data-testid="stWidgetLabel"], span, p { color: #1a1a1a !important; }
    /* selectbox / multiselect / date_input — 테두리 + 텍스트 */
    [data-baseweb="select"] { background-color: #ffffff !important; }
    [data-baseweb="select"] > div { background-color: #ffffff !important; border-color: #ccc !important; color: #1a1a1a !important; }
    [data-baseweb="select"] span, [data-baseweb="select"] div { color: #1a1a1a !important; }
    [data-baseweb="input"] { background-color: #ffffff !important; border-color: #ccc !important; color: #1a1a1a !important; }
    [data-baseweb="input"] input { color: #1a1a1a !important; }
    [data-baseweb="popover"] li { color: #1a1a1a !important; background-color: #ffffff !important; }
    [data-baseweb="popover"] li:hover { background-color: #e8e8e8 !important; }
    /* date input */
    [data-testid="stDateInput"] input { background-color: #ffffff !important; border-color: #ccc !important; color: #1a1a1a !important; }
    [data-testid="stDateInput"] [data-baseweb="calendar"] { background-color: #ffffff !important; color: #1a1a1a !important; }
    /* multiselect 태그 */
    [data-baseweb="tag"] { background-color: #e0e0e0 !important; color: #1a1a1a !important; }
    /* expander */
    [data-testid="stExpander"] { border-color: #ddd !important; }
    [data-testid="stExpander"] summary,
    [data-testid="stExpander"] summary span { color: #1a1a1a !important; }
    /* 데이터프레임/테이블 */
    .stDataFrame, .stTable { color: #1a1a1a !important; }
    /* 탭 */
    .stTabs [data-baseweb="tab-list"] button { color: #555 !important; }
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] { color: #1a1a1a !important; }
    /* divider */
    hr { border-color: #ddd !important; }
    /* caption */
    .stCaption, [data-testid="stCaptionContainer"] { color: #666 !important; }
    /* 버튼 (새로고침, 이전달/다음달 등) */
    .stButton > button,
    [data-testid="stSidebar"] .stButton > button,
    [data-testid="baseButton-secondary"] {
        background-color: #f0f2f6 !important;
        color: #1a1a1a !important;
        border: 1px solid #ccc !important;
    }
    .stButton > button:hover,
    [data-testid="stSidebar"] .stButton > button:hover,
    [data-testid="baseButton-secondary"]:hover {
        background-color: #e0e2e6 !important;
        border-color: #999 !important;
    }
    /* 아이콘 버튼 (캘린더 화살표 등) */
    [data-baseweb="button-group"] button,
    [data-baseweb="calendar"] button,
    button[kind="icon"] {
        color: #1a1a1a !important;
        background-color: transparent !important;
    }
    [data-baseweb="button-group"] button:hover,
    [data-baseweb="calendar"] button:hover {
        background-color: #e0e2e6 !important;
    }
    </style>
    """
    # Plotly 템플릿 전역 설정
    pio.templates.default = "plotly" if not want_dark else "plotly_dark"
    if override:
        st.markdown(_DARK_CSS if want_dark else _LIGHT_CSS, unsafe_allow_html=True)

    # ── 사이드바 ──
    st.sidebar.title("NOAH 대시보드")
    page = st.sidebar.radio("페이지", [
        "오늘의 현황", "수주/출고 현황", "제품 분석",
        "섹터 분석", "고객 분석", "발주 커버리지",
        "수익성 분석", "Order Book",
    ])
    st.sidebar.divider()
    market = st.sidebar.radio("시장 구분", ["전체", "국내", "해외"], horizontal=True)

    so_raw = load_so()
    years = sorted(so_raw["period"].dropna().str[:4].unique()) if not so_raw.empty else []
    year = st.sidebar.selectbox("연도", ["전체"] + list(years), index=0)
    month = st.sidebar.selectbox(
        "월", ["전체"] + [f"{m:02d}" for m in range(1, 13)], index=0,
        disabled=(year == "전체"),
        help="연도를 먼저 선택하세요" if year == "전체" else None,
    )

    all_sectors = sorted(s for s in so_raw["sector"].dropna().unique() if s) if not so_raw.empty else []
    sectors = st.sidebar.multiselect("섹터", all_sectors)

    with st.sidebar.expander("고객 필터"):
        all_custs = sorted(c for c in so_raw["customer_name"].dropna().unique() if c) if not so_raw.empty else []
        customers = st.multiselect("고객", all_custs)

    if st.sidebar.button("🔄 데이터 새로고침"):
        st.cache_data.clear()
        st.rerun()

    meta = load_sync_meta()
    if meta:
        latest = max((v["last_sync"] for v in meta.values() if v.get("last_sync")), default="")
        if latest:
            st.sidebar.caption(f"마지막 동기화: {latest}")

    # ── 로더 에러 배너 ──
    _show_load_errors()

    # ── 라우팅 ──
    kw = dict(market=market, sectors=sectors, customers=customers, year=year, month=month)
    {
        "오늘의 현황": pg_today,
        "수주/출고 현황": pg_orders,
        "제품 분석": pg_product,
        "섹터 분석": pg_sector,
        "고객 분석": pg_customer,
        "발주 커버리지": pg_po_coverage,
        "수익성 분석": pg_margin,
        "Order Book": pg_orderbook,
    }[page](**kw)


# ═══════════════════════════════════════════════════════════════
# 납기 캘린더 (pg_today 내부 사용)
# ═══════════════════════════════════════════════════════════════
def build_calendar_data(so_pending: pd.DataFrame, dn: pd.DataFrame,
                        year: int, month: int,
                        ship_df: pd.DataFrame | None = None) -> pd.DataFrame:
    """월별 납기/출고/EXW/픽업 집계 — 날짜별 건수·금액 반환 (테스트 가능한 순수 함수)."""
    import calendar as _cal
    _, n_days = _cal.monthrange(year, month)
    days = pd.DataFrame({"day": range(1, n_days + 1)})

    # SO 납기 집계
    so_agg = pd.DataFrame(columns=["day", "so_count", "so_amount"])
    if not so_pending.empty and "delivery_date" in so_pending.columns:
        sp = so_pending[
            (so_pending["delivery_date"].dt.year == year)
            & (so_pending["delivery_date"].dt.month == month)
        ].copy()
        if not sp.empty:
            sp["day"] = sp["delivery_date"].dt.day
            so_agg = sp.groupby("day").agg(
                so_count=("SO_ID", "nunique"),
                so_amount=("amount_krw", "sum"),
            ).reset_index()

    # DN 출고 집계
    dn_agg = pd.DataFrame(columns=["day", "dn_count", "dn_amount"])
    if not dn.empty and "dispatch_date" in dn.columns:
        dp = dn[
            (dn["dispatch_date"].dt.year == year)
            & (dn["dispatch_date"].dt.month == month)
        ].copy()
        if not dp.empty:
            dp["day"] = dp["dispatch_date"].dt.day
            dn_agg = dp.groupby("day").agg(
                dn_count=("DN_ID", "nunique"),
                dn_amount=("amount_krw", "sum"),
            ).reset_index()

    # EXW 출고 예정 집계
    exw_agg = pd.DataFrame(columns=["day", "exw_count"])
    if not so_pending.empty and "exw_noah" in so_pending.columns:
        ep = so_pending[
            so_pending["exw_noah"].notna()
            & (so_pending["exw_noah"].dt.year == year)
            & (so_pending["exw_noah"].dt.month == month)
        ].copy()
        if not ep.empty:
            ep["day"] = ep["exw_noah"].dt.day
            exw_agg = ep.groupby("day").agg(
                exw_count=("SO_ID", "nunique"),
            ).reset_index()

    # 공장 픽업 집계
    pk_agg = pd.DataFrame(columns=["day", "pk_count"])
    if ship_df is not None and not ship_df.empty and "pickup_date" in ship_df.columns:
        pp = ship_df[
            ship_df["pickup_date"].notna()
            & (ship_df["pickup_date"].dt.year == year)
            & (ship_df["pickup_date"].dt.month == month)
        ].copy()
        if not pp.empty:
            pp["day"] = pp["pickup_date"].dt.day
            pk_agg = pp.groupby("day").agg(
                pk_count=("DN_ID", "nunique"),
            ).reset_index()

    # 해외 선적 예정 집계
    ship_agg = pd.DataFrame(columns=["day", "ship_count"])
    if ship_df is not None and not ship_df.empty and "expected_ship_date" in ship_df.columns:
        sd = ship_df[
            ship_df["expected_ship_date"].notna()
            & (ship_df["expected_ship_date"].dt.year == year)
            & (ship_df["expected_ship_date"].dt.month == month)
        ].copy()
        if not sd.empty:
            sd["day"] = sd["expected_ship_date"].dt.day
            ship_agg = sd.groupby("day").agg(
                ship_count=("DN_ID", "nunique"),
            ).reset_index()

    merged = (days
              .merge(so_agg, on="day", how="left")
              .merge(dn_agg, on="day", how="left")
              .merge(exw_agg, on="day", how="left")
              .merge(pk_agg, on="day", how="left")
              .merge(ship_agg, on="day", how="left"))
    for c in ("so_count", "so_amount", "dn_count", "dn_amount", "exw_count", "pk_count", "ship_count"):
        merged[c] = pd.to_numeric(merged[c], errors="coerce").fillna(0)
    return merged


def _render_delivery_calendar(so_pending: pd.DataFrame, dn: pd.DataFrame):
    """납기 캘린더 — Plotly Heatmap 기반, 월 네비게이션 + 날짜 클릭 드릴다운."""
    st.subheader("납기 캘린더")

    # Session state — 월 네비게이션
    if "cal_year" not in st.session_state:
        st.session_state.cal_year = _TODAY.year
    if "cal_month" not in st.session_state:
        st.session_state.cal_month = _TODAY.month

    cy, cm = st.session_state.cal_year, st.session_state.cal_month

    # 네비게이션 버튼
    nav_l, nav_c, nav_r = st.columns([1, 3, 1])
    with nav_l:
        if st.button("◀ 이전 달", key="cal_prev"):
            if cm == 1:
                st.session_state.cal_year = cy - 1
                st.session_state.cal_month = 12
            else:
                st.session_state.cal_month = cm - 1
            st.rerun()
    with nav_c:
        st.markdown(f"<h3 style='text-align:center;margin:0'>{cy}년 {cm}월</h3>",
                    unsafe_allow_html=True)
    with nav_r:
        if st.button("다음 달 ▶", key="cal_next"):
            if cm == 12:
                st.session_state.cal_year = cy + 1
                st.session_state.cal_month = 1
            else:
                st.session_state.cal_month = cm + 1
            st.rerun()

    # 데이터 준비
    ship_df = load_dn_export_shipping()
    cal_data = build_calendar_data(so_pending, dn, cy, cm, ship_df=ship_df)

    # 달력 그리드 구성
    cal_obj = calendar.Calendar(firstweekday=6)  # 일요일 시작
    weeks = cal_obj.monthdayscalendar(cy, cm)
    n_weeks = len(weeks)
    weekdays = ["일", "월", "화", "수", "목", "금", "토"]

    z_vals = []       # heatmap z
    text_vals = []    # 셀 텍스트
    customdata = []   # 클릭 이벤트용 실제 날짜

    today_cell = None  # (row, col) 오늘 위치

    for wi, week in enumerate(weeks):
        z_row, text_row, cd_row = [], [], []
        for di, day in enumerate(week):
            if day == 0:
                z_row.append(None)
                text_row.append("")
                cd_row.append("")
            else:
                row = cal_data[cal_data["day"] == day]
                so_cnt = int(row["so_count"].iloc[0]) if not row.empty else 0
                so_amt = float(row["so_amount"].iloc[0]) if not row.empty else 0
                dn_cnt = int(row["dn_count"].iloc[0]) if not row.empty else 0
                dn_amt = float(row["dn_amount"].iloc[0]) if not row.empty else 0
                exw_cnt = int(row["exw_count"].iloc[0]) if not row.empty else 0
                pk_cnt = int(row["pk_count"].iloc[0]) if not row.empty else 0
                ship_cnt = int(row["ship_count"].iloc[0]) if not row.empty else 0

                # Z값: 양수=미래/오늘 납기, 음수=과납기
                from datetime import date as _date_cls
                cell_date = _date_cls(cy, cm, day)
                if so_cnt > 0 and cell_date < _TODAY_DATE:
                    z_val = -so_cnt  # 과납기
                else:
                    z_val = so_cnt

                # 텍스트 조립
                lines = [f"<b>{day}</b>"]
                if exw_cnt > 0:
                    lines.append(f"\U0001F3ED {exw_cnt}건")  # 🏭 EXW 출고
                if pk_cnt > 0:
                    lines.append(f"\U0001F69B {pk_cnt}건")   # 🚛 픽업
                if so_cnt > 0:
                    lines.append(f"\U0001F4E6 {so_cnt}건 {fmt_krw(so_amt)}")
                if ship_cnt > 0:
                    lines.append(f"\U0001F6A2 {ship_cnt}건")   # 🚢 해외 선적
                if dn_cnt > 0:
                    lines.append(f"\U0001F69A {dn_cnt}건 {fmt_krw(dn_amt)}")

                z_row.append(z_val)
                text_row.append("<br>".join(lines))
                cd_row.append(f"{cy}-{cm:02d}-{day:02d}")

                if cell_date == _TODAY_DATE:
                    today_cell = (wi, di)

        z_vals.append(z_row)
        text_vals.append(text_row)
        customdata.append(cd_row)

    # Plotly Heatmap
    fig_cal = go.Figure(data=go.Heatmap(
        z=z_vals,
        text=text_vals,
        texttemplate="%{text}",
        textfont=dict(size=11),
        customdata=customdata,
        x=weekdays,
        y=[f"{i+1}주" for i in range(n_weeks)],
        colorscale=[
            [0, "#d62728"],     # 빨강 (과납기)
            [0.5, "#ffffff"],   # 흰색 (0건)
            [1, "#1f77b4"],     # 파랑 (미래 납기)
        ],
        zmid=0,
        showscale=False,
        hoverinfo="text",
        xgap=2,
        ygap=2,
    ))

    # 오늘 테두리 강조
    if today_cell is not None:
        tr, tc = today_cell
        fig_cal.add_shape(
            type="rect",
            x0=tc - 0.5, x1=tc + 0.5,
            y0=tr - 0.5, y1=tr + 0.5,
            line=dict(color="#ff6600", width=3),
        )

    fig_cal.update_layout(
        height=50 + n_weeks * 80,
        margin=dict(l=40, r=20, t=10, b=20),
        yaxis=dict(autorange="reversed"),
        xaxis=dict(side="top"),
        plot_bgcolor="white",
    )

    st.plotly_chart(fig_cal, key="cal_heatmap", use_container_width=True)

    # 날짜 선택 → 상세
    # 현재 달이면 오늘 날짜 기본 선택, 다른 달이면 1일
    if cy == _TODAY.year and cm == _TODAY.month:
        default_date = _TODAY_DATE
    else:
        default_date = datetime(cy, cm, 1).date()
    sel_date = st.date_input(
        "날짜 선택",
        value=default_date, key="cal_date_pick",
        min_value=datetime(cy, cm, 1).date(),
        max_value=datetime(cy, cm, calendar.monthrange(cy, cm)[1]).date(),
    )
    sel_date_str = sel_date.strftime("%Y-%m-%d") if sel_date else None

    if sel_date_str:
        # 각 섹션 건수 미리 계산 (expander 라벨용)
        if not so_pending.empty and "exw_noah" in so_pending.columns:
            day_exw = so_pending[so_pending["exw_noah"].dt.date == sel_date]
        else:
            day_exw = pd.DataFrame()
        ship_all = load_dn_export_shipping()
        if not ship_all.empty:
            pickup_dns = ship_all.loc[ship_all["pickup_date"].dt.date == sel_date, "DN_ID"].unique()
            day_pickup = ship_all[ship_all["DN_ID"].isin(pickup_dns)]
        else:
            day_pickup = pd.DataFrame()
        if not so_pending.empty:
            day_so = so_pending[so_pending["delivery_date"].dt.date == sel_date]
        else:
            day_so = pd.DataFrame()
        if not dn.empty:
            day_dn = dn[dn["dispatch_date"].dt.date == sel_date]
        else:
            day_dn = pd.DataFrame()
        if not ship_all.empty and "expected_ship_date" in ship_all.columns:
            ship_dns = ship_all.loc[ship_all["expected_ship_date"].dt.date == sel_date, "DN_ID"].unique()
            day_ship = ship_all[ship_all["DN_ID"].isin(ship_dns)]
        else:
            day_ship = pd.DataFrame()
        n_exw = day_exw["SO_ID"].nunique() if not day_exw.empty else 0
        n_pickup = day_pickup["DN_ID"].nunique() if not day_pickup.empty else 0
        n_delivery = day_so["SO_ID"].nunique() if not day_so.empty else 0
        n_dispatch = day_dn["DN_ID"].nunique() if not day_dn.empty else 0
        n_ship = day_ship["DN_ID"].nunique() if not day_ship.empty else 0
        summary_parts = []
        if n_exw: summary_parts.append(f"EXW {n_exw}")
        if n_pickup: summary_parts.append(f"픽업 {n_pickup}")
        if n_ship: summary_parts.append(f"선적 {n_ship}")
        if n_delivery: summary_parts.append(f"납기 {n_delivery}")
        if n_dispatch: summary_parts.append(f"출고 {n_dispatch}")
        summary = " · ".join(summary_parts) if summary_parts else "해당 건 없음"

        with st.expander(f"📅 {sel_date_str} 상세 — {summary}", expanded=sel_date == _TODAY_DATE):

            # ── (A) EXW 출고 예정 (SO 국내+해외 EXW NOAH 기준) ──
            st.markdown("**🏭 EXW 출고 예정** — 공장 출고 예정 오더")
            if not day_exw.empty:
                agg = dict(
                    고객명=("customer_name", "first"),
                    섹터=("sector", "first"),
                    품목수=("line_item", "nunique") if "line_item" in day_exw.columns else ("os_name", "count"),
                    총수량=("qty", "sum"),
                    총금액=("amount_krw", "sum"),
                    요청납기=("requested_date", "min"),
                    납품예정일=("delivery_date", "min"),
                    마켓=("market", "first"),
                    Status=("status", "first"),
                )
                if "customer_po" in day_exw.columns:
                    agg["고객PO"] = ("customer_po", "first")
                g = day_exw.groupby("SO_ID").agg(**agg).reset_index()
                items = []
                for _, r in g.iterrows():
                    icon = _status_icon(r["Status"], False)
                    mkt_tag = "🇰🇷" if r["마켓"] == "국내" else "🌏"
                    po_info = f" · PO: {r['고객PO']}" if r.get("고객PO") else ""
                    req = fmt_date(r["요청납기"]) if pd.notna(r["요청납기"]) else "ASAP"
                    lines = [
                        f"품목 {r['품목수']}건 · 수량 {int(r['총수량']):,} · {fmt_krw(r['총금액'])}{po_info}",
                        f"📅 납기 {req}",
                    ]
                    if pd.notna(r["납품예정일"]):
                        lines.append(f"📦 납품 예정일 {fmt_date(r['납품예정일'])}")
                    sec_tag = f" · {r['섹터']}" if r["섹터"] else ""
                    items.append({
                        "title": f"{icon} {mkt_tag} **{r['SO_ID']}**  {r['고객명']}{sec_tag}",
                        "lines": lines,
                    })
                _render_cards(items, cols_per_row=2)
            else:
                st.info("EXW 출고 예정 건 없음")

            # ── (B) 공장 픽업 (DN_해외 공장 픽업일 기준) ──
            st.markdown("**🚛 공장 픽업** — 해외 DN 공장 픽업 예정")
            if not day_pickup.empty:
                day_pickup["carrier"] = day_pickup["carrier"].fillna("")
                so_meta = load_so()[["SO_ID", "customer_po", "sector", "incoterms", "shipping_method"]].drop_duplicates(subset=["SO_ID"])
                day_pickup = day_pickup.merge(so_meta, on="SO_ID", how="left")
                day_pickup["customer_po"] = day_pickup["customer_po"].fillna("")
                day_pickup["sector"] = day_pickup["sector"].fillna("")
                day_pickup["incoterms"] = day_pickup["incoterms"].fillna("")
                day_pickup["shipping_method"] = day_pickup["shipping_method"].fillna("")
                pk = day_pickup.groupby("DN_ID").agg(
                    고객명=("customer_name", "first"),
                    섹터=("sector", "first"),
                    고객PO=("customer_po", "first"),
                    품목수=("item_name", "nunique"),
                    총수량=("qty", "sum"),
                    총금액=("amount_krw", "sum"),
                    공장출고일=("factory_date", "min"),
                    선적예정일=("expected_ship_date", "min"),
                    운송업체=("carrier", "first"),
                    SO_ID=("SO_ID", lambda x: " / ".join(x.unique())),
                    Incoterms=("incoterms", "first"),
                    운송방식=("shipping_method", "first"),
                ).reset_index()
                items = []
                for _, r in pk.iterrows():
                    po_info = f" · PO: {r['고객PO']}" if r["고객PO"] else ""
                    lines = [
                        f"SO: {r['SO_ID']} · 품목 {r['품목수']}건 · 수량 {int(r['총수량']):,} · {fmt_krw(r['총금액'])}{po_info}",
                        f"출고 {fmt_date(r['공장출고일'])} → 선적예정 {fmt_date(r['선적예정일'])}",
                    ]
                    terms_parts = []
                    if r["Incoterms"]:
                        terms_parts.append(r["Incoterms"])
                    if r["운송방식"]:
                        terms_parts.append(r["운송방식"])
                    if terms_parts:
                        lines.append(f"📦 {' · '.join(terms_parts)}")
                    if r["운송업체"]:
                        lines.append(f"🚛 운송 업체: {r['운송업체']}")
                    sec_tag = f" · {r['섹터']}" if r["섹터"] else ""
                    items.append({
                        "title": f"🚛 **{r['DN_ID']}**  {r['고객명']}{sec_tag}",
                        "lines": lines,
                    })
                _render_cards(items, cols_per_row=2)
            else:
                st.info("공장 픽업 예정 건 없음")

            # ── (E) 해외 선적 예정 (DN_해외 선적 예정일 기준) ──
            st.markdown("**🚢 해외 선적 예정** — 해외 DN 선적 예정 현황")
            if not day_ship.empty:
                so_meta2 = load_so()[["SO_ID", "customer_po", "sector", "incoterms", "shipping_method"]].drop_duplicates(subset=["SO_ID"])
                day_ship = day_ship.merge(so_meta2, on="SO_ID", how="left")
                day_ship["customer_po"] = day_ship["customer_po"].fillna("")
                day_ship["sector"] = day_ship["sector"].fillna("")
                day_ship["incoterms"] = day_ship["incoterms"].fillna("")
                day_ship["shipping_method"] = day_ship["shipping_method"].fillna("")
                day_ship["bl_no"] = day_ship["bl_no"].fillna("")
                day_ship["carrier"] = day_ship["carrier"].fillna("")
                sh = day_ship.groupby("DN_ID").agg(
                    고객명=("customer_name", "first"),
                    섹터=("sector", "first"),
                    고객PO=("customer_po", "first"),
                    품목수=("item_name", "nunique"),
                    총수량=("qty", "sum"),
                    총금액=("amount_krw", "sum"),
                    공장출고일=("factory_date", "min"),
                    픽업일=("pickup_date", "min"),
                    선적예정일=("expected_ship_date", "min"),
                    BL=("bl_no", "first"),
                    운송업체=("carrier", "first"),
                    SO_ID=("SO_ID", lambda x: " / ".join(x.unique())),
                    Incoterms=("incoterms", "first"),
                    운송방식=("shipping_method", "first"),
                ).reset_index()
                items = []
                for _, r in sh.iterrows():
                    po_info = f" · PO: {r['고객PO']}" if r["고객PO"] else ""
                    lines = [
                        f"SO: {r['SO_ID']} · 품목 {r['품목수']}건 · 수량 {int(r['총수량']):,} · {fmt_krw(r['총금액'])}{po_info}",
                        f"출고 {fmt_date(r['공장출고일'])} → 픽업 {fmt_date(r['픽업일'])} → 선적예정 {fmt_date(r['선적예정일'])}",
                    ]
                    terms_parts = []
                    if r["Incoterms"]:
                        terms_parts.append(r["Incoterms"])
                    if r["운송방식"]:
                        terms_parts.append(r["운송방식"])
                    if terms_parts:
                        lines.append(f"📦 {' · '.join(terms_parts)}")
                    if r["BL"]:
                        lines.append(f"📄 B/L: {r['BL']}")
                    if r["운송업체"]:
                        lines.append(f"🚛 운송 업체: {r['운송업체']}")
                    sec_tag = f" · {r['섹터']}" if r["섹터"] else ""
                    items.append({
                        "title": f"🚢 **{r['DN_ID']}**  {r['고객명']}{sec_tag}",
                        "lines": lines,
                    })
                _render_cards(items, cols_per_row=2)
            else:
                st.info("해외 선적 예정 건 없음")

            d_a, d_b = st.columns(2)

            # (C) 납기 예정
            with d_a:
                st.markdown("**📦 납기 예정**")
                if not day_so.empty:
                    agg = dict(
                        고객명=("customer_name", "first"),
                        섹터=("sector", "first"),
                        품목수=("line_item", "nunique") if "line_item" in day_so.columns else ("os_name", "count"),
                        총수량=("qty", "sum"),
                        총금액=("amount_krw", "sum"),
                        공장출고일=("exw_noah", "min"),
                        Status=("status", "first"),
                    )
                    if "customer_po" in day_so.columns:
                        agg["고객PO"] = ("customer_po", "first")
                    g = day_so.groupby("SO_ID").agg(**agg).reset_index()
                    overdue_flag = sel_date < _TODAY_DATE
                    items = []
                    for _, r in g.iterrows():
                        icon = _status_icon(r["Status"], overdue_flag)
                        po_info = f" · PO: {r['고객PO']}" if r.get("고객PO") else ""
                        sec_tag = f" · {r['섹터']}" if r.get("섹터") else ""
                        items.append({
                            "title": f"{icon}  **{r['SO_ID']}**  {r['고객명']}{sec_tag}",
                            "lines": [
                                f"품목 {r['품목수']}건 · 수량 {int(r['총수량']):,} · {fmt_krw(r['총금액'])}{po_info}",
                                f"📅 EXW {fmt_date(r['공장출고일'])}",
                            ],
                        })
                    _render_cards(items, cols_per_row=2)
                else:
                    st.info("납기 예정 건 없음")

            # (D) 출고 실적
            with d_b:
                st.markdown("**🚚 출고 실적**")
                if not day_dn.empty:
                    # SO에서 customer_po 조인
                    if not so_pending.empty and "customer_po" in so_pending.columns:
                        so_po = so_pending[["SO_ID", "customer_po"]].drop_duplicates(subset=["SO_ID"])
                        day_dn = day_dn.merge(so_po, on="SO_ID", how="left")
                    if "customer_po" not in day_dn.columns:
                        day_dn["customer_po"] = ""
                    day_dn["customer_po"] = day_dn["customer_po"].fillna("")
                    agg_dict = dict(
                        총수량=("qty", "sum"),
                        총금액=("amount_krw", "sum"),
                        고객PO=("customer_po", "first"),
                    )
                    if "customer_name" in day_dn.columns:
                        agg_dict["고객명"] = ("customer_name", "first")
                    if "sector" in day_dn.columns:
                        agg_dict["섹터"] = ("sector", "first")
                    agg_dict["SO_ID"] = ("SO_ID", "first")
                    if "line_item" in day_dn.columns:
                        agg_dict["품목수"] = ("line_item", "nunique")
                    tbl = day_dn.groupby("DN_ID").agg(**agg_dict).reset_index()
                    items = []
                    for _, r in tbl.iterrows():
                        cust = r.get("고객명", "")
                        n_items = r.get("품목수", "?")
                        po_info = f" · PO: {r['고객PO']}" if r.get("고객PO") else ""
                        sec = r.get("섹터", "")
                        sec_tag = f" · {sec}" if sec else ""
                        items.append({
                            "title": f"📦 **{r['DN_ID']}**  {cust}{sec_tag}",
                            "lines": [
                                f"SO: {r['SO_ID']} · 품목 {n_items}건{po_info}",
                                f"수량 {int(r['총수량']):,} · {fmt_krw(r['총금액'])}",
                            ],
                        })
                    _render_cards(items, cols_per_row=2)
                else:
                    st.info("출고 실적 없음")


# ═══════════════════════════════════════════════════════════════
# Page 1: 오늘의 현황
# ═══════════════════════════════════════════════════════════════
def pg_today(market, sectors, customers, **_):
    st.title("오늘의 현황")
    if _.get("year", "전체") != "전체" or _.get("month", "전체") != "전체":
        st.caption("ℹ️ 이 페이지는 현재 시점 기준입니다 — 연도/월 필터는 적용되지 않습니다.")

    so = filt(load_so(), market, sectors, customers)
    dn = filt(enrich_dn(load_dn(), load_so()), market, sectors, customers, period_col=None)
    backlog = filt(load_backlog(), market, sectors, customers, period_col=None)

    # 납기 계산 (미완료 건만)
    completed = ("출고 완료",)
    so_pending = so[(so["delivery_date"].notna()) & (~so["status"].isin(completed))] if not so.empty else pd.DataFrame()

    today_due = so_pending[so_pending["delivery_date"].dt.date == _TODAY_DATE] if not so_pending.empty else pd.DataFrame()
    week_due = so_pending[
        (so_pending["delivery_date"].dt.date >= _WEEK_START)
        & (so_pending["delivery_date"].dt.date <= _WEEK_END)
    ] if not so_pending.empty else pd.DataFrame()
    overdue = so_pending[so_pending["delivery_date"].dt.date < _TODAY_DATE] if not so_pending.empty else pd.DataFrame()

    month_so = so[so["period"] == _THIS_MONTH] if not so.empty else pd.DataFrame()
    month_dn = dn[dn["dispatch_month"] == _THIS_MONTH] if not dn.empty else pd.DataFrame()

    # KPI 카드
    today_n = len(today_due)
    week_n = len(week_due)
    overdue_n = overdue["SO_ID"].nunique() if not overdue.empty else 0
    overdue_amt = overdue["amount_krw"].sum() if not overdue.empty else 0
    so_amt = month_so["amount_krw"].sum() if not month_so.empty else 0
    dn_amt = month_dn["amount_krw"].sum() if not month_dn.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        with st.container(border=True):
            st.markdown(f"{'🔴' if today_n else '🟢'} **오늘 납기**")
            st.markdown(f"### {today_n}건")
    with c2:
        with st.container(border=True):
            st.markdown(f"{'🔴' if overdue_n else '🟢'} **이번 주 납기**")
            st.markdown(f"### {week_n}건")
            if overdue_n:
                st.caption(f"🔴 지연 {overdue_n}건 / {fmt_krw(overdue_amt)}")
    with c3:
        with st.container(border=True):
            st.markdown("📥 **금월 수주**")
            st.markdown(f"### {fmt_krw(so_amt)}")
            prev_so = so[so["period"] == _PREV_MONTH] if not so.empty else pd.DataFrame()
            prev_so_amt = prev_so["amount_krw"].sum() if not prev_so.empty else 0
            if prev_so_amt:
                delta_pct = (so_amt - prev_so_amt) / prev_so_amt * 100
                st.caption(f"{'↑' if delta_pct >= 0 else '↓'} 전월 대비 {delta_pct:+.0f}%")
    with c4:
        with st.container(border=True):
            st.markdown("📤 **금월 출고**")
            st.markdown(f"### {fmt_krw(dn_amt)}")
            if so_amt > 0:
                ratio = min(dn_amt / so_amt, 1.0)
                st.progress(ratio)
                st.caption(f"수주 대비 {dn_amt / so_amt * 100:.0f}%")

    # ── 납기 캘린더 ──
    _render_delivery_calendar(so_pending, dn)

    # ── PO 확정 지연 (Sent → Confirmed 미전환) ──
    st.subheader("📋 PO 확정 지연")
    po_sent = load_po_sent_pending()
    if not po_sent.empty:
        # 사이드바 필터 적용
        if market != "전체":
            po_sent = po_sent[po_sent["market"] == market]
        # SO 메타 조인 (customer_name, sector)
        so_meta = so[["SO_ID", "customer_name", "sector"]].drop_duplicates(subset=["SO_ID"])
        po_sent = po_sent.merge(so_meta, on="SO_ID", how="left")
        po_sent["customer_name"] = po_sent["customer_name"].fillna("")
        po_sent["sector"] = po_sent["sector"].fillna("")
        if sectors:
            po_sent = po_sent[po_sent["sector"].isin(sectors)]
        if customers:
            po_sent = po_sent[po_sent["customer_name"].isin(customers)]

    if not po_sent.empty and po_sent["order_date"].notna().any():
        # SO_ID + PO_ID 기준 집계
        g = po_sent.groupby(["SO_ID", "PO_ID"]).agg(
            고객명=("customer_name", "first"),
            섹터=("sector", "first"),
            마켓=("market", "first"),
            품목수=("po_qty", "count"),
            총수량=("po_qty", "sum"),
            총ICO=("po_total_ico", "sum"),
            발주일=("order_date", "min"),
            공장EXW=("factory_exw", "min"),
            OC번호=("noah_oc", "first"),
        ).reset_index()
        g = g[g["발주일"].notna()].copy()
        g["경과일"] = g["발주일"].apply(lambda d: (_TODAY_DATE - d.date()).days if pd.notna(d) else 0)
        g = g[g["경과일"] > 0].copy()

        if not g.empty:
            st.caption(f"공장 발주(Sent) 후 확정(Confirmed) 미전환 건: **{len(g)}건**")
            g = g.copy()
            g[["bk", "bk_icon", "bk_label"]] = g["경과일"].apply(
                lambda d: pd.Series(_assign_bucket(d))
            )

            tab_dom, tab_exp = st.tabs(["🇰🇷 국내", "🌏 해외"])
            for tab, mkt in [(tab_dom, "국내"), (tab_exp, "해외")]:
                with tab:
                    mkt_g = g[g["마켓"] == mkt]
                    if mkt_g.empty:
                        st.info(f"{mkt} 해당 건 없음")
                        continue
                    mkt_g = mkt_g.sort_values(["bk", "경과일"], ascending=[False, False])

                    from itertools import groupby as _igroupby
                    for bkey, grp in _igroupby(
                        mkt_g.itertuples(),
                        key=lambda r: (r.bk, r.bk_icon, r.bk_label),
                    ):
                        grp_list = list(grp)
                        _, icon, label = bkey
                        st.caption(f"{icon} **{label}** ({len(grp_list)}건)")
                        for r in grp_list:
                            so_id = r.SO_ID
                            po_id = r.PO_ID
                            cust = getattr(r, "고객명", "")
                            n_items = getattr(r, "품목수", 0)
                            tot_qty = int(getattr(r, "총수량", 0))
                            tot_ico = getattr(r, "총ICO", 0)
                            ord_dt = getattr(r, "발주일", pd.NaT)
                            exw_dt = getattr(r, "공장EXW", pd.NaT)
                            oc = getattr(r, "OC번호", "")
                            days = getattr(r, "경과일", 0)
                            sec = getattr(r, "섹터", "")
                            sec_tag = f" [{sec}]" if sec else ""
                            oc_info = f" · OC: {oc}" if oc else ""
                            exw_info = f" · EXW {fmt_date(exw_dt)}" if pd.notna(exw_dt) else ""
                            header = (
                                f"{icon} **{po_id}**  {cust}{sec_tag} — "
                                f"SO: {so_id} · 품목 {n_items}건 · 수량 {tot_qty:,} · {fmt_krw(tot_ico)}{oc_info} · "
                                f"발주일 {fmt_date(ord_dt)} (**{days}일**){exw_info}"
                            )
                            with st.expander(header):
                                detail = po_sent[
                                    (po_sent["SO_ID"] == so_id) & (po_sent["PO_ID"] == po_id)
                                ][["SO_ID", "PO_ID", "item_name", "po_qty", "po_total_ico", "order_date", "factory_exw", "noah_oc"]].copy()
                                detail.columns = ["SO_ID", "PO_ID", "품목명", "수량", "ICO 금액", "발주일", "공장 EXW", "OC No."]
                                detail["ICO 금액"] = detail["ICO 금액"].apply(lambda v: f"{int(v):,}" if pd.notna(v) else "")
                                detail["수량"] = detail["수량"].apply(lambda v: int(v) if pd.notna(v) else 0)
                                detail["발주일"] = detail["발주일"].apply(lambda v: fmt_date(v) if pd.notna(v) else "")
                                detail["공장 EXW"] = detail["공장 EXW"].apply(lambda v: fmt_date(v) if pd.notna(v) else "")
                                st.dataframe(detail, use_container_width=True, hide_index=True)
        else:
            st.success("PO 확정 지연 건 없음")
    else:
        st.success("PO 확정 지연 건 없음")

    # ── 미발주 현황 (공장 PO 미발주 SO) ──
    st.subheader("📋 미발주 현황")
    po_detail = load_po_detail()
    po_all_status = load_po_status()
    # 출고 완료 제외
    so_active = so[so["status"] != "출고 완료"] if not so.empty else pd.DataFrame()
    cov = calc_coverage(so_active, po_detail, po_all_status=po_all_status)
    if not cov.empty:
        unordered = cov[cov["coverage_status"].isin(["PO 미등록", "미발주", "부분 발주"])].copy()
    else:
        unordered = pd.DataFrame()

    if not unordered.empty:
        # 수주일 기준 경과일 (수주일 없으면 -1로 표시, 필터에서 제외하지 않음)
        unordered["경과일"] = unordered["po_receipt_date"].apply(
            lambda d: (_TODAY_DATE - d.date()).days if pd.notna(d) else -1
        )

    if not unordered.empty:
        n_unreg = len(unordered[unordered["coverage_status"] == "PO 미등록"])
        n_unord = len(unordered[unordered["coverage_status"] == "미발주"])
        n_part = len(unordered[unordered["coverage_status"] == "부분 발주"])
        parts = []
        if n_unreg:
            parts.append(f"🔴 PO미등록 {n_unreg}건")
        if n_unord:
            parts.append(f"미발주 {n_unord}건")
        if n_part:
            parts.append(f"부분발주 {n_part}건")
        st.caption(f"공장 발주 필요: **{' · '.join(parts)}**")

        unordered[["bk", "bk_icon", "bk_label"]] = unordered["경과일"].apply(
            lambda d: pd.Series((-1, "⚪", "수주일 미입력")) if d < 0 else pd.Series(_assign_bucket(d))
        )

        tab_dom, tab_exp = st.tabs(["🇰🇷 국내", "🌏 해외"])
        for tab, mkt in [(tab_dom, "국내"), (tab_exp, "해외")]:
            with tab:
                mkt_u = unordered[unordered["market"] == mkt]
                if mkt_u.empty:
                    st.info(f"{mkt} 미발주 건 없음")
                    continue
                mkt_u = mkt_u.sort_values(["bk", "경과일"], ascending=[False, False])

                from itertools import groupby as _igroupby
                for bkey, grp in _igroupby(
                    mkt_u.itertuples(),
                    key=lambda r: (r.bk, r.bk_icon, r.bk_label),
                ):
                    grp_list = list(grp)
                    _, icon, label = bkey
                    st.caption(f"{icon} **{label}** ({len(grp_list)}건)")
                    for r in grp_list:
                        so_id = r.SO_ID
                        cust = getattr(r, "customer_name", "")
                        sec = getattr(r, "sector", "")
                        sec_tag = f" [{sec}]" if sec else ""
                        os_nm = getattr(r, "os_name", "")
                        qty_val = int(getattr(r, "qty", 0))
                        amt = getattr(r, "amount_krw", 0)
                        ico = getattr(r, "po_total_ico", 0)
                        rcv_dt = getattr(r, "po_receipt_date", pd.NaT)
                        dlv_dt = getattr(r, "delivery_date", pd.NaT)
                        days = getattr(r, "경과일", 0)
                        cov_st = getattr(r, "coverage_status", "")
                        po_ids_val = getattr(r, "open_po_ids", "")
                        po_tag = f" · PO: {po_ids_val}" if po_ids_val else ""
                        ico_tag = f" · ICO {fmt_krw(ico)}" if ico else ""
                        cov_tag = "PO미등록" if cov_st == "PO 미등록" else ("부분발주" if cov_st == "부분 발주" else "미발주")
                        dlv_info = f" · 납기 {fmt_date(dlv_dt)}" if pd.notna(dlv_dt) else ""
                        rcv_info = f"수주일 {fmt_date(rcv_dt)} (**{days}일**)" if days >= 0 else "수주일 미입력"
                        header = (
                            f"{icon} **{so_id}**  {cust}{sec_tag} — "
                            f"{os_nm} · 수량 {qty_val:,} · {fmt_krw(amt)}{ico_tag}{po_tag} · "
                            f"{rcv_info}{dlv_info} · `{cov_tag}`"
                        )
                        with st.expander(header):
                            detail = so_active[so_active["SO_ID"] == so_id][
                                ["SO_ID", "line_item", "item_name", "os_name", "qty", "amount_krw", "po_receipt_date", "delivery_date", "status"]
                            ].drop_duplicates(subset=["SO_ID", "line_item"]).copy()
                            detail = detail.drop(columns=["line_item"])
                            detail.columns = ["SO_ID", "품목명", "OS name", "수량", "매출금액", "수주일", "납기일", "Status"]
                            detail["PO_ID"] = po_ids_val
                            detail["공장발주일"] = ""
                            detail = detail[["SO_ID", "PO_ID", "품목명", "OS name", "수량", "매출금액", "수주일", "공장발주일", "납기일", "Status"]]
                            detail["수량"] = detail["수량"].apply(fmt_qty)
                            detail["매출금액"] = detail["매출금액"].apply(fmt_num)
                            detail["수주일"] = detail["수주일"].apply(fmt_date)
                            detail["납기일"] = detail["납기일"].apply(fmt_date)
                            st.dataframe(detail, use_container_width=True, hide_index=True)
    else:
        st.success("미발주 건 없음")

    # ── EXW 완료 미출고 (PO 공장 EXW date < 오늘 & 미Invoiced) ──
    # PO line item 단위: 공장 EXW 경과했는데 Invoiced 안 된 라인
    st.subheader("🚨 EXW 완료 미출고")
    po_exw = load_po_exw_pending()
    if not po_exw.empty:
        # 공장 EXW < 오늘 필터
        po_exw = po_exw[po_exw["factory_exw"].notna() & (po_exw["factory_exw"].dt.date < _TODAY_DATE)].copy()
        # 사이드바 필터 적용
        if market != "전체":
            po_exw = po_exw[po_exw["market"] == market]
        # SO 메타 조인
        so_meta = so[["SO_ID", "customer_name", "sector", "customer_po", "delivery_date"]].drop_duplicates(subset=["SO_ID"])
        po_exw = po_exw.merge(so_meta, on="SO_ID", how="left")
        po_exw["customer_name"] = po_exw["customer_name"].fillna("")
        po_exw["sector"] = po_exw["sector"].fillna("")
        if sectors:
            po_exw = po_exw[po_exw["sector"].isin(sectors)]
        if customers:
            po_exw = po_exw[po_exw["customer_name"].isin(customers)]

    if not po_exw.empty:
        # SO_ID 단위 집계 (PO line 기준)
        g = po_exw.groupby("SO_ID").agg(
            고객명=("customer_name", "first"),
            섹터=("sector", "first"),
            고객PO=("customer_po", "first"),
            PO_ID=("PO_ID", "first"),
            OC번호=("noah_oc", "first"),
            마켓=("market", "first"),
            품목수=("line_item", "nunique"),
            총수량=("po_qty", "sum"),
            총ICO=("po_total_ico", "sum"),
            EXW=("factory_exw", "min"),
            납기=("delivery_date", "min"),
        ).reset_index().sort_values("EXW")
        g["경과일"] = g["EXW"].apply(lambda d: (_TODAY_DATE - d.date()).days if pd.notna(d) else 0)

        st.caption(
            f"EXW 일자 경과 후 Invoiced 미완료 건: **{len(g)}건** "
            "— 공장에 EXW date 재확인 필요"
        )

        tab_dom, tab_exp = st.tabs(["🇰🇷 국내", "🌏 해외"])
        for tab, mkt in [(tab_dom, "국내"), (tab_exp, "해외")]:
            with tab:
                mkt_g = g[g["마켓"] == mkt]
                if mkt_g.empty:
                    st.info(f"{mkt} 해당 건 없음")
                    continue
                mkt_g = mkt_g.copy()
                mkt_g[["bk", "bk_icon", "bk_label"]] = mkt_g["경과일"].apply(
                    lambda d: pd.Series(_assign_bucket(d))
                )
                mkt_g = mkt_g.sort_values(["bk", "경과일"], ascending=[False, False])
                mkt_lines = po_exw[po_exw["market"] == mkt]

                from itertools import groupby as _igroupby
                for bkey, grp in _igroupby(
                    mkt_g.itertuples(),
                    key=lambda r: (r.bk, r.bk_icon, r.bk_label),
                ):
                    grp_list = list(grp)
                    _, icon, label = bkey
                    st.caption(f"{icon} **{label}** ({len(grp_list)}건)")
                    for r in grp_list:
                        so_id = r.SO_ID
                        cust = getattr(r, "고객명", "")
                        sec = getattr(r, "섹터", "")
                        cust_po = getattr(r, "고객PO", "")
                        po_id = r.PO_ID
                        oc = r.OC번호
                        n_items = getattr(r, "품목수", 0)
                        tot_qty = getattr(r, "총수량", 0)
                        tot_ico = getattr(r, "총ICO", 0)
                        exw_dt = r.EXW
                        days = getattr(r, "경과일", 0)
                        del_dt = getattr(r, "납기", pd.NaT)
                        po_info = f" · PO: {cust_po}" if cust_po else ""
                        oc_info = f" · OC: {oc}" if oc else ""
                        sec_tag = f" [{sec}]" if sec else ""
                        req = fmt_date(del_dt) if pd.notna(del_dt) else "—"
                        header = (
                            f"{icon} **{so_id}** ({po_id})  {cust}{sec_tag} — "
                            f"품목 {n_items}건 · 수량 {int(tot_qty):,} · {fmt_krw(tot_ico)}{po_info}{oc_info} · "
                            f"EXW {fmt_date(exw_dt)} (**{days}일**)"
                        )
                        with st.expander(header):
                            detail = mkt_lines[mkt_lines["SO_ID"] == so_id][
                                ["line_item", "item_name", "po_qty", "po_total_ico", "factory_exw", "po_status"]
                            ].copy()
                            detail.columns = ["Line", "품목명", "수량", "ICO 금액", "공장 EXW", "Status"]
                            detail["ICO 금액"] = detail["ICO 금액"].apply(lambda v: f"{int(v):,}" if pd.notna(v) else "")
                            detail["공장 EXW"] = detail["공장 EXW"].apply(lambda v: fmt_date(v) if pd.notna(v) else "")
                            detail = detail.sort_values("Line").reset_index(drop=True)
                            st.dataframe(detail, use_container_width=True, hide_index=True)
    else:
        st.success("EXW 미출고 건 없음")

    # ── 납기 현황 (납기 경과 & DN 미생성/부분출고) ──
    # 납품 추적: 납기 지났는데 DN이 없거나 출고 수량이 부족한 건
    # → 납기 경과 라인이 있는 SO 식별 후, 해당 SO의 **전체 라인** 집계
    st.subheader("📦 납기 현황 (미완료 건)")

    # Step 1: 납기 경과 라인에서 SO_ID 후보 식별
    _so_due_lines = so[
        so["delivery_date"].notna()
        & (so["delivery_date"].dt.date < _TODAY_DATE)
        & (so["status"] != "출고 완료")
    ] if not so.empty else pd.DataFrame()

    if not _so_due_lines.empty:
        # DN qty+금액 집계
        if not dn.empty:
            dn_agg = dn.groupby(["SO_ID", "line_item"]).agg(
                dn_qty=("qty", "sum"),
                dn_amount=("amount_krw", "sum"),
            ).reset_index()
        else:
            dn_agg = pd.DataFrame(columns=["SO_ID", "line_item", "dn_qty", "dn_amount"])

        # Step 2: 납기 경과 라인에 DN 매칭 → 잔여 있는 SO_ID만, 납기 경과 라인만
        due_check = _so_due_lines.merge(dn_agg, on=["SO_ID", "line_item"], how="left")
        due_check["dn_qty"] = due_check["dn_qty"].fillna(0)
        due_check["dn_amount"] = due_check["dn_amount"].fillna(0)
        due_check["remaining_qty"] = due_check["qty"] - due_check["dn_qty"]
        due_check["remaining_amount"] = due_check["amount_krw"] - due_check["dn_amount"]

        # Step 2a: SO exw_noah 누락 시 PO factory_exw로 보충
        # (PO 라인이 SO 라인과 1:1 대응 안 되는 경우 SO_ID 단위로 매칭)
        if not po_detail.empty and "factory_exw" in po_detail.columns:
            _po_exw = po_detail[["SO_ID", "factory_exw"]].dropna(subset=["factory_exw"])
            if not _po_exw.empty:
                due_check = due_check.merge(_po_exw, on="SO_ID", how="left")
                _mask = due_check["exw_noah"].isna() & due_check["factory_exw"].notna()
                due_check.loc[_mask, "exw_noah"] = due_check.loc[_mask, "factory_exw"]
                due_check.drop(columns=["factory_exw"], inplace=True)
        so_ids_with_remaining = set(due_check.loc[due_check["remaining_qty"] > 0, "SO_ID"])

        # 납기 경과 라인 중 잔여 있는 SO만 (미래 납기 라인 제외)
        due_pending = due_check[due_check["SO_ID"].isin(so_ids_with_remaining)].copy() if so_ids_with_remaining else pd.DataFrame()

        if not due_pending.empty:
            _n_so = due_pending["SO_ID"].nunique()
            st.caption(
                f"납기 경과 후 DN 미생성 또는 부분출고 건: **{_n_so}건** "
                "— DN 발급 또는 납기 일정 확인 필요"
            )
            tab_dom, tab_exp = st.tabs(["🇰🇷 국내", "🌏 해외"])
            for tab, mkt in [(tab_dom, "국내"), (tab_exp, "해외")]:
                with tab:
                    mkt_df = due_pending[due_pending["market"] == mkt]
                    if mkt_df.empty:
                        st.info(f"{mkt} 납기 지연 건 없음")
                        continue
                    g = mkt_df.groupby("SO_ID").agg(
                        고객명=("customer_name", "first"),
                        섹터=("sector", "first"),
                        고객PO=("customer_po", "first"),
                        품목수=("line_item", "nunique"),
                        총수량=("qty", "sum"),
                        출고수량=("dn_qty", "sum"),
                        잔여수량=("remaining_qty", "sum"),
                        총금액=("amount_krw", "sum"),
                        잔여금액=("remaining_amount", "sum"),
                        납기일=("delivery_date", "min"),
                        공장출고일=("exw_noah", "min"),
                        Status=("status", "first"),
                    ).reset_index().sort_values("납기일")
                    g["경과일"] = g["납기일"].apply(
                        lambda d: (_TODAY_DATE - d.date()).days if pd.notna(d) else 0
                    )
                    # 버킷 분류
                    g = g.copy()
                    g[["bk", "bk_icon", "bk_label"]] = g["경과일"].apply(
                        lambda d: pd.Series(_assign_bucket(d))
                    )
                    g = g.sort_values(["bk", "경과일"], ascending=[False, False])

                    from itertools import groupby as _igroupby
                    for bkey, grp in _igroupby(
                        g.itertuples(),
                        key=lambda r: (r.bk, r.bk_icon, r.bk_label),
                    ):
                        grp_list = list(grp)
                        _, icon, label = bkey
                        st.caption(f"{icon} **{label}** ({len(grp_list)}건)")
                        for r in grp_list:
                            so_id = r.SO_ID
                            cust = getattr(r, "고객명", "")
                            sec = getattr(r, "섹터", "")
                            cust_po = getattr(r, "고객PO", "")
                            n_items = getattr(r, "품목수", 0)
                            tot_qty = int(getattr(r, "총수량", 0))
                            dn_shipped = int(getattr(r, "출고수량", 0))
                            remaining = int(getattr(r, "잔여수량", 0))
                            tot_amt = getattr(r, "총금액", 0)
                            rem_amt = getattr(r, "잔여금액", 0)
                            del_dt = getattr(r, "납기일", pd.NaT)
                            exw_dt = getattr(r, "공장출고일", pd.NaT)
                            days = getattr(r, "경과일", 0)
                            po_info = f" · PO: {cust_po}" if cust_po else ""
                            sec_tag = f" [{sec}]" if sec else ""
                            qty_tag = f"잔여 {remaining:,}/{tot_qty:,}" if dn_shipped > 0 else f"수량 {tot_qty:,}"
                            amt_tag = f"{fmt_krw(rem_amt)}/{fmt_krw(tot_amt)}" if dn_shipped > 0 else fmt_krw(tot_amt)
                            header = (
                                f"{icon} **{so_id}**  {cust}{sec_tag} — "
                                f"{qty_tag} · {amt_tag}{po_info} · "
                                f"납기 {fmt_date(del_dt)} (**{days}일**)"
                            )
                            with st.expander(header):
                                detail = mkt_df[mkt_df["SO_ID"] == so_id][
                                    ["line_item", "item_name", "qty", "dn_qty", "remaining_qty",
                                     "amount_krw", "dn_amount", "remaining_amount", "delivery_date", "exw_noah"]
                                ].copy()
                                detail.columns = [
                                    "Line", "품목명", "주문", "출고", "잔여",
                                    "주문금액", "출고금액", "잔여금액", "납기", "EXW",
                                ]
                                detail["주문금액"] = detail["주문금액"].apply(lambda v: f"{int(v):,}" if pd.notna(v) else "")
                                detail["출고금액"] = detail["출고금액"].apply(lambda v: f"{int(v):,}" if v else "")
                                detail["잔여금액"] = detail["잔여금액"].apply(lambda v: f"{int(v):,}" if v else "")
                                detail["주문"] = detail["주문"].apply(lambda v: int(v) if pd.notna(v) else 0)
                                detail["출고"] = detail["출고"].apply(lambda v: int(v) if v else 0)
                                detail["잔여"] = detail["잔여"].apply(lambda v: int(v) if v else 0)
                                detail["납기"] = detail["납기"].apply(lambda v: fmt_date(v) if pd.notna(v) else "")
                                detail["EXW"] = detail["EXW"].apply(lambda v: fmt_date(v) if pd.notna(v) else "")
                                detail = detail.sort_values("Line").reset_index(drop=True)
                                st.dataframe(detail, use_container_width=True, hide_index=True)
        else:
            st.success("납기 지연 건 없음")
    else:
        st.info("납기 지연 건 없음")

    # ── 해외 선적 Action Items ──
    if market != "국내":
        st.subheader("🚢 해외 선적 Action Items")
        ship_df = load_dn_export_shipping()
        if not ship_df.empty:
            # SO 메타 조인으로 sector, customer_po, incoterms, shipping_method 추가
            so_meta = load_so()[["SO_ID", "sector", "customer_po", "incoterms", "shipping_method"]].drop_duplicates(subset=["SO_ID"])
            ship_df = ship_df.merge(so_meta, on="SO_ID", how="left")
            ship_df["sector"] = ship_df["sector"].fillna("")
            ship_df["customer_po"] = ship_df["customer_po"].fillna("")
            ship_df["incoterms"] = ship_df["incoterms"].fillna("")
            ship_df["shipping_method"] = ship_df["shipping_method"].fillna("")
        if sectors and not ship_df.empty:
            ship_df = ship_df[ship_df["sector"].isin(sectors)]
        if customers and not ship_df.empty:
            ship_df = ship_df[ship_df["customer_name"].isin(customers)]
        if not ship_df.empty:
            # 공장 출고 완료 but 선적 미완료
            pending_ship = ship_df[ship_df["ship_date"].isna()].copy()
            if not pending_ship.empty:
                pending_ship["carrier"] = pending_ship["carrier"].fillna("")

                # DN 단위 집계
                ps = pending_ship.groupby("DN_ID").agg(
                    고객명=("customer_name", "first"),
                    섹터=("sector", "first"),
                    품목수=("item_name", "nunique"),
                    총수량=("qty", "sum"),
                    총금액=("amount_krw", "sum"),
                    공장출고일=("factory_date", "min"),
                    픽업일=("pickup_date", "min"),
                    선적예정일=("expected_ship_date", "min"),
                    운송업체=("carrier", "first"),
                    SO_ID=("SO_ID", "first"),
                    Incoterms=("incoterms", "first"),
                    운송방식=("shipping_method", "first"),
                ).reset_index()

                # 경과일 (공장출고일 기준)
                ps["경과일"] = ps["공장출고일"].apply(
                    lambda d: (_TODAY_DATE - d.date()).days if pd.notna(d) else 0
                )

                # 파이프라인 단계 분류
                def _ship_stage(r):
                    if r["운송업체"] == "":
                        return "🔴 포워더 미정"
                    if pd.isna(r["픽업일"]):
                        return "🟡 픽업 대기"
                    return "🟠 선적 대기"
                ps["단계"] = ps.apply(_ship_stage, axis=1)

                # ── (1) KPI 카드 ──
                _stage_order = ["🔴 포워더 미정", "🟡 픽업 대기", "🟠 선적 대기"]
                kpi_cols = st.columns(3)
                for i, stage in enumerate(_stage_order):
                    stage_df = ps[ps["단계"] == stage]
                    n = len(stage_df)
                    amt = stage_df["총금액"].sum()
                    kpi_cols[i].metric(stage, f"{n}건", delta=fmt_krw(amt), delta_color="off")

                # ── (2) 고객별 / 운송방식별 요약 ──
                tab_cust, tab_ship = st.tabs(["고객별 현황", "운송방식별 현황"])

                with tab_cust:
                    cust_rows = []
                    for cust, cg in ps.groupby("고객명"):
                        row = {"고객명": cust, "DN건수": len(cg)}
                        for stage in _stage_order:
                            label = stage[2:]  # "포워더 미정" etc.
                            row[label] = int((cg["단계"] == stage).sum())
                        row["총수량"] = int(cg["총수량"].sum())
                        row["총금액"] = fmt_krw(cg["총금액"].sum())
                        row["최대경과일"] = f"{int(cg['경과일'].max())}일"
                        cust_rows.append(row)
                    cust_tbl = pd.DataFrame(cust_rows).sort_values("DN건수", ascending=False)
                    st.dataframe(cust_tbl, use_container_width=True, hide_index=True)

                with tab_ship:
                    ship_method_rows = []
                    for method, mg in ps.groupby("운송방식"):
                        label = method if method else "(미지정)"
                        row = {"운송방식": label, "DN건수": len(mg)}
                        for stage in _stage_order:
                            stg_label = stage[2:]
                            row[stg_label] = int((mg["단계"] == stage).sum())
                        row["총수량"] = int(mg["총수량"].sum())
                        row["총금액"] = fmt_krw(mg["총금액"].sum())
                        row["최대경과일"] = f"{int(mg['경과일'].max())}일"
                        ship_method_rows.append(row)
                    ship_tbl = pd.DataFrame(ship_method_rows).sort_values("DN건수", ascending=False)
                    st.dataframe(ship_tbl, use_container_width=True, hide_index=True)

                # ── (3) DN 상세 테이블 (접기) ──
                with st.expander(f"📋 DN 상세 ({len(ps)}건)", expanded=False):
                    detail = ps[["단계", "DN_ID", "고객명", "SO_ID", "Incoterms",
                                 "운송방식", "품목수", "총수량", "총금액",
                                 "공장출고일", "픽업일", "선적예정일",
                                 "운송업체", "경과일"]].copy()
                    detail = detail.sort_values(["단계", "경과일"], ascending=[True, False])
                    detail["총금액"] = detail["총금액"].apply(lambda v: fmt_krw(v))
                    detail["총수량"] = detail["총수량"].apply(lambda v: int(v))
                    for dc in ("공장출고일", "픽업일", "선적예정일"):
                        detail[dc] = detail[dc].apply(lambda v: fmt_date(v) if pd.notna(v) else "")
                    detail["경과일"] = detail["경과일"].apply(lambda v: f"{int(v)}일")
                    detail["운송업체"] = detail["운송업체"].replace("", "-")
                    detail["Incoterms"] = detail["Incoterms"].replace("", "-")
                    detail["운송방식"] = detail["운송방식"].replace("", "-")
                    st.dataframe(detail, use_container_width=True, hide_index=True)
            else:
                st.success("선적 대기 건 없음")

        else:
            st.info("해외 출고 데이터 없음")

    # ── 세금계산서 미발행 (국내) ──
    if market != "해외":
        st.subheader("🧾 세금계산서 미발행 (국내)")
        tax_pending = load_dn_tax_pending()
        if not tax_pending.empty:
            # 고객/섹터 필터 적용
            if customers:
                tax_pending = tax_pending[tax_pending["customer_name"].isin(customers)]

            if not tax_pending.empty:
                # Aging 계산
                tax_pending["경과일"] = (_TODAY_DATE - tax_pending["dispatch_date"].dt.date).apply(lambda d: d.days if d else 0)

                def _tax_aging(d):
                    if d <= 7:
                        return "① 7일 이내"
                    if d <= 14:
                        return "② 7~14일"
                    if d <= 30:
                        return "③ 14~30일"
                    return "④ 30일+"

                tax_pending["aging"] = tax_pending["경과일"].apply(_tax_aging)

                # DN별 집계
                dn_agg = tax_pending.groupby("DN_ID").agg(
                    customer_name=("customer_name", "first"),
                    품목수=("line_item", "nunique"),
                    amount_krw=("amount_krw", "sum"),
                    dispatch_date=("dispatch_date", "min"),
                    경과일=("경과일", "max"),
                    aging=("aging", "max"),
                ).reset_index()

                # KPI 카드
                total_pending = len(dn_agg)
                total_amt = dn_agg["amount_krw"].sum()
                max_days = dn_agg["경과일"].max()
                over_30 = len(dn_agg[dn_agg["경과일"] > 30])

                tc1, tc2, tc3, tc4 = st.columns(4)
                with tc1:
                    with st.container(border=True):
                        st.markdown("**미발행 건수**")
                        st.markdown(f"### {total_pending}건")
                with tc2:
                    with st.container(border=True):
                        st.markdown("**미발행 금액**")
                        st.markdown(f"### {fmt_krw(total_amt)}")
                with tc3:
                    with st.container(border=True):
                        st.markdown(f"{'🔴' if max_days > 30 else '🟡'} **최장 경과**")
                        st.markdown(f"### {max_days}일")
                with tc4:
                    with st.container(border=True):
                        st.markdown(f"{'🔴' if over_30 else '🟢'} **30일 초과**")
                        st.markdown(f"### {over_30}건")

                # Aging 바 + 고객별 금액
                tax_col1, tax_col2 = st.columns(2)
                with tax_col1:
                    aging_agg = dn_agg.groupby("aging").agg(
                        건수=("DN_ID", "count"),
                        금액=("amount_krw", "sum"),
                    ).reset_index().sort_values("aging")
                    aging_colors = {
                        "① 7일 이내": C_ENDING, "② 7~14일": "#ff9800",
                        "③ 14~30일": C_PURPLE, "④ 30일+": C_DANGER,
                    }
                    fig_ta = px.bar(aging_agg, x="aging", y="금액", color="aging",
                                    labels={"aging": "출고 후 경과", "금액": "미발행 금액"},
                                    color_discrete_map=aging_colors,
                                    text="건수")
                    fig_ta.update_traces(
                        texttemplate="%{text}건",
                        hovertemplate="<b>%{x}</b><br>₩%{y:,.0f} (%{text}건)<extra></extra>",
                    )
                    fig_ta.update_layout(height=350, margin=dict(t=30, b=30), showlegend=False)
                    st.plotly_chart(fig_ta, use_container_width=True)

                with tax_col2:
                    cust_agg = dn_agg.groupby("customer_name").agg(
                        건수=("DN_ID", "count"),
                        금액=("amount_krw", "sum"),
                    ).sort_values("금액", ascending=False).head(10).reset_index()
                    cust_agg.columns = ["고객", "건수", "금액"]
                    fig_tc = px.bar(cust_agg, y="고객", x="금액", orientation="h",
                                    color_discrete_sequence=[C_INPUT], text="건수")
                    fig_tc.update_traces(
                        texttemplate="%{text}건",
                        hovertemplate="<b>%{y}</b><br>₩%{x:,.0f} (%{text}건)<extra></extra>",
                    )
                    fig_tc.update_layout(height=350, margin=dict(t=30, l=150),
                                         yaxis=dict(autorange="reversed"), showlegend=False)
                    st.plotly_chart(fig_tc, use_container_width=True)

                # 상세 (expander)
                with st.expander(f"미발행 상세 {total_pending}건"):
                    detail = dn_agg[["DN_ID", "customer_name", "품목수", "amount_krw", "dispatch_date", "경과일"]].copy()
                    detail.columns = ["DN_ID", "고객명", "품목수", "금액", "출고일", "경과일"]
                    detail = detail.sort_values("경과일", ascending=False)
                    detail["출고일"] = detail["출고일"].apply(fmt_date)
                    detail["금액"] = detail["금액"].apply(fmt_num)
                    detail["경과일"] = detail["경과일"].apply(lambda d: f"{d}일")
                    st.dataframe(detail, use_container_width=True, hide_index=True)
            else:
                st.success("세금계산서 미발행 건 없음 (필터 기준)")
        else:
            st.success("세금계산서 미발행 건 없음")


# ═══════════════════════════════════════════════════════════════
# Page 2: 수주/출고 현황
# ═══════════════════════════════════════════════════════════════
def pg_orders(market, sectors, customers, year, month):
    st.title("수주/출고 현황")

    # 전체 (KPI용, 연도/월 필터 무시)
    so_all = filt(load_so(), market, sectors, customers)
    dn_ej = enrich_dn(load_dn(), load_so())
    dn_all = filt(dn_ej, market, sectors, customers, period_col=None)

    # 차트용 (연도/월 필터 추가 적용 — so_all은 이미 m/s/c 필터 적용 상태)
    so = filt(load_so(), market, sectors, customers, year=year, month=month)
    dn = dn_all.copy()
    if not dn.empty:
        if year and year != "전체":
            dn = dn[dn["dispatch_month"].astype(str).str.startswith(year)]
            if month and month != "전체":
                dn = dn[dn["dispatch_month"] == f"{year}-{month}"]

    # KPI 기준 월 결정
    if year and year != "전체" and month and month != "전체":
        # 연도+월 모두 선택
        kpi_month = f"{year}-{month}"
        y, m = int(year), int(month)
        kpi_prev = f"{y}-{m - 1:02d}" if m > 1 else f"{y - 1}-12"
        kpi_label, prev_label = kpi_month, kpi_prev
    elif year and year != "전체":
        # 연도만 선택 → 해당 연도 내 최신 월 기준
        year_periods = so_all[so_all["period"].astype(str).str.startswith(year)]["period"] if not so_all.empty else pd.Series(dtype=str)
        if not year_periods.empty:
            kpi_month = str(year_periods.max())
            try:
                y, m = int(kpi_month[:4]), int(kpi_month[5:7])
                kpi_prev = f"{y}-{m - 1:02d}" if m > 1 else f"{y - 1}-12"
            except (ValueError, IndexError):
                kpi_month, kpi_prev = _THIS_MONTH, _PREV_MONTH
            kpi_label, prev_label = kpi_month, kpi_prev
        else:
            kpi_month, kpi_prev = _THIS_MONTH, _PREV_MONTH
            kpi_label, prev_label = "금월", "전월"
    else:
        kpi_month, kpi_prev = _THIS_MONTH, _PREV_MONTH
        kpi_label, prev_label = "금월", "전월"

    # 월별 집계
    month_so = so_all[so_all["period"] == kpi_month] if not so_all.empty else pd.DataFrame()
    prev_so = so_all[so_all["period"] == kpi_prev] if not so_all.empty else pd.DataFrame()
    month_dn = dn_all[dn_all["dispatch_month"] == kpi_month] if not dn_all.empty else pd.DataFrame()
    prev_dn = dn_all[dn_all["dispatch_month"] == kpi_prev] if not dn_all.empty else pd.DataFrame()

    this_so_amt = month_so["amount_krw"].sum() if not month_so.empty else 0
    prev_so_amt = prev_so["amount_krw"].sum() if not prev_so.empty else 0
    this_dn_amt = month_dn["amount_krw"].sum() if not month_dn.empty else 0
    prev_dn_amt = prev_dn["amount_krw"].sum() if not prev_dn.empty else 0

    # ── KPI (2행) ──
    st.subheader("수주")
    c1, c2, c3 = st.columns(3)
    so_delta = f"{(this_so_amt - prev_so_amt) / prev_so_amt * 100:+.0f}%" if prev_so_amt else None
    c1.metric(f"{kpi_label} 수주", fmt_krw(this_so_amt), delta=so_delta)
    c2.metric(f"{prev_label} 수주", fmt_krw(prev_so_amt))
    c3.metric("누적 수주", fmt_krw(so_all["amount_krw"].sum() if not so_all.empty else 0))

    st.subheader("출고")
    c4, c5, c6 = st.columns(3)
    dn_delta = f"{(this_dn_amt - prev_dn_amt) / prev_dn_amt * 100:+.0f}%" if prev_dn_amt else None
    c4.metric(f"{kpi_label} 출고", fmt_krw(this_dn_amt), delta=dn_delta)
    c5.metric(f"{prev_label} 출고", fmt_krw(prev_dn_amt))
    if this_dn_amt:
        btb = this_so_amt / this_dn_amt
        c6.metric(f"Book-to-Bill ({kpi_label})", f"{btb:.2f}", help="수주/출고 비율 (>1: 수주 우세)")
    else:
        c6.metric(f"Book-to-Bill ({kpi_label})", "N/A", help="출고 0 — 산출 불가")

    # ── 월별 수주/출고 금액 + 누적매출 ──
    st.subheader("월별 수주/출고 금액 추이")
    so_m = (
        so.groupby("period").agg(수주금액=("amount_krw", "sum")).reset_index()
        if not so.empty else pd.DataFrame(columns=["period", "수주금액"])
    )
    dn_m = (
        dn.groupby("dispatch_month").agg(출고금액=("amount_krw", "sum"))
        .reset_index().rename(columns={"dispatch_month": "period"})
        if not dn.empty else pd.DataFrame(columns=["period", "출고금액"])
    )
    merged = pd.merge(so_m, dn_m, on="period", how="outer").fillna(0).sort_values("period")

    if not merged.empty:
        merged["누적매출"] = merged["출고금액"].cumsum()
        fig = go.Figure()
        fig.add_trace(go.Bar(x=merged["period"], y=merged["수주금액"],
                             name="수주", marker_color=C_INPUT,
                             hovertemplate="<b>%{x}</b><br>수주: ₩%{y:,.0f}<extra></extra>"))
        fig.add_trace(go.Bar(x=merged["period"], y=merged["출고금액"],
                             name="출고", marker_color=C_OUTPUT,
                             hovertemplate="<b>%{x}</b><br>출고: ₩%{y:,.0f}<extra></extra>"))
        fig.add_trace(go.Scatter(
            x=merged["period"], y=merged["누적매출"], name="누적매출",
            mode="lines+markers", yaxis="y2",
            line=dict(color=C_ENDING, width=2),
            hovertemplate="<b>%{x}</b><br>누적매출: ₩%{y:,.0f}<extra></extra>",
        ))
        fig.update_layout(
            barmode="group", height=400, margin=dict(t=30, b=30),
            xaxis=dict(type="category", dtick=1, rangeslider=dict(visible=True)),
            yaxis2=dict(title="누적매출", overlaying="y", side="right"),
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("데이터 없음")

    # ── Book-to-Bill 월별 추이 ──
    if not merged.empty and (merged["출고금액"] > 0).any():
        st.subheader("Book-to-Bill 월별 추이")
        btb_m = merged[merged["출고금액"] > 0].copy()
        btb_m["BtB"] = btb_m["수주금액"] / btb_m["출고금액"]
        fig_btb = go.Figure()
        fig_btb.add_trace(go.Scatter(
            x=btb_m["period"], y=btb_m["BtB"], mode="lines+markers",
            line=dict(color=C_PURPLE, width=2), name="Book-to-Bill",
            hovertemplate="<b>%{x}</b><br>BtB: %{y:.2f}<extra></extra>",
        ))
        fig_btb.add_hline(y=1.0, line_dash="dash", line_color="gray",
                          annotation_text="균형선 (1.0)")
        fig_btb.update_layout(
            height=280, margin=dict(t=30, b=30),
            xaxis=dict(type="category"),
            yaxis=dict(title="수주/출고 비율"),
        )
        st.plotly_chart(fig_btb, use_container_width=True)

    # ── 일별 수주/출고 현황 ──
    st.subheader("일별 수주/출고 현황")

    # 월 선택 — SO/DN에 존재하는 월 목록
    available_months = set()
    if not so_all.empty:
        available_months.update(so_all["period"].dropna().unique())
    if not dn_all.empty:
        available_months.update(dn_all["dispatch_month"].dropna().unique())
    available_months = sorted(available_months)
    if not available_months:
        st.info("선택된 필터 조건에 해당하는 수주/출고 데이터가 없습니다.")
        daily_month = None
    else:
        default_idx = available_months.index(kpi_month) if kpi_month in available_months else len(available_months) - 1
        daily_month = st.selectbox("월 선택", available_months, index=default_idx, key="daily_month")

    if daily_month:
        dc1, dc2 = st.columns(2)

        # 일별 수주 (PO receipt date 기준)
        with dc1:
            st.markdown(f"**{daily_month} 일별 수주**")
            if not so_all.empty and "po_receipt_date" in so_all.columns:
                so_with_date = so_all[so_all["po_receipt_date"].notna()].copy()
                so_with_date["receipt_month"] = so_with_date["po_receipt_date"].dt.strftime("%Y-%m")
                m_so = so_with_date[so_with_date["receipt_month"] == daily_month].copy()
                if not m_so.empty:
                    m_so["day"] = m_so["po_receipt_date"].dt.strftime("%m-%d")
                    daily_so = m_so.groupby("day")["amount_krw"].sum().reset_index()
                    total_so_amt = daily_so["amount_krw"].sum()
                    st.metric("수주 합계", fmt_krw(total_so_amt), help=f"{m_so['SO_ID'].nunique()}건")
                    fig_so_d = px.bar(daily_so, x="day", y="amount_krw",
                                      labels={"day": "날짜", "amount_krw": "수주금액"},
                                      color_discrete_sequence=[C_INPUT])
                    fig_so_d.update_traces(hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>")
                    fig_so_d.update_layout(height=300, margin=dict(t=10, b=30),
                                           xaxis=dict(type="category"))
                    st.plotly_chart(fig_so_d, use_container_width=True)
                else:
                    st.info("수주 데이터 없음")

        # 일별 출고
        with dc2:
            st.markdown(f"**{daily_month} 일별 출고**")
            if not dn_all.empty:
                m_dn = dn_all[dn_all["dispatch_month"] == daily_month].copy()
                if not m_dn.empty:
                    m_dn["day"] = m_dn["dispatch_date"].dt.strftime("%m-%d")
                    daily_dn = m_dn.groupby("day")["amount_krw"].sum().reset_index()
                    total_dn_amt = daily_dn["amount_krw"].sum()
                    st.metric("출고 합계", fmt_krw(total_dn_amt), help=f"{m_dn['DN_ID'].nunique()}건")
                    fig_dn_d = px.bar(daily_dn, x="day", y="amount_krw",
                                      labels={"day": "날짜", "amount_krw": "출고금액"},
                                      color_discrete_sequence=[C_OUTPUT])
                    fig_dn_d.update_traces(hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>")
                    fig_dn_d.update_layout(height=300, margin=dict(t=10, b=30),
                                           xaxis=dict(type="category"))
                    st.plotly_chart(fig_dn_d, use_container_width=True)
                else:
                    st.info("출고 데이터 없음")


# ═══════════════════════════════════════════════════════════════
# Page 3: 제품 분석
# ═══════════════════════════════════════════════════════════════
def pg_product(market, sectors, customers, year, month):
    st.title("제품 분석")
    so = filt(load_so(), market, sectors, customers, year=year, month=month)
    if so.empty:
        st.info("데이터 없음")
        return

    by_amt = so.groupby("os_name")["amount_krw"].sum().sort_values(ascending=False)
    by_qty = so.groupby("os_name")["qty"].sum().sort_values(ascending=False)

    # KPI
    c1, c2, c3 = st.columns(3)
    c1.metric("총 제품 종류", f"{so['os_name'].nunique()}")
    c2.metric("최다 판매 제품", by_qty.index[0] if len(by_qty) else "-")
    c3.metric("최고 매출 제품", by_amt.index[0] if len(by_amt) else "-")

    # Top 15 매출
    st.subheader("제품별 매출 Top 15")
    top15 = by_amt.head(15).reset_index()
    top15.columns = ["제품", "매출"]
    fig = px.bar(top15, y="제품", x="매출", orientation="h",
                 color_discrete_sequence=[C_INPUT])
    fig.update_traces(hovertemplate="<b>%{y}</b><br>매출: ₩%{x:,.0f}<extra></extra>")
    fig.update_layout(
        height=max(400, len(top15) * 32),
        margin=dict(t=30, l=200),
        yaxis=dict(autorange="reversed"),
    )
    event_prod = st.plotly_chart(fig, use_container_width=True, on_select="rerun", key="product_top15")
    selected_product = None
    if event_prod and event_prod.selection and event_prod.selection.points:
        selected_product = event_prod.selection.points[0]["y"]
    if selected_product:
        st.subheader(f"📌 {selected_product} 상세")
        sub_p = so[so["os_name"] == selected_product]
        dc1, dc2 = st.columns(2)
        with dc1:
            st.markdown("**월별 매출 추이**")
            pm = sub_p.groupby("period")["amount_krw"].sum().reset_index()
            fig_pm = px.bar(pm, x="period", y="amount_krw",
                            labels={"period": "월", "amount_krw": "매출"},
                            color_discrete_sequence=[C_INPUT])
            fig_pm.update_traces(hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>")
            fig_pm.update_layout(height=300, margin=dict(t=30, b=30),
                                 xaxis=dict(type="category"))
            st.plotly_chart(fig_pm, use_container_width=True)
        with dc2:
            st.markdown("**섹터별 비중**")
            ps = sub_p.groupby("sector")["amount_krw"].sum().reset_index()
            fig_ps = px.pie(ps, names="sector", values="amount_krw", hole=0.4)
            fig_ps.update_traces(hovertemplate="<b>%{label}</b><br>₩%{value:,.0f} (%{percent})<extra></extra>")
            fig_ps.update_layout(height=300, margin=dict(t=30, b=30))
            st.plotly_chart(fig_ps, use_container_width=True)
        st.markdown("**주요 고객 Top 5**")
        pc = sub_p.groupby("customer_name")["amount_krw"].sum().nlargest(5).reset_index()
        pc.columns = ["고객", "매출"]
        fig_pc = px.bar(pc, y="고객", x="매출", orientation="h",
                        color_discrete_sequence=[C_OUTPUT])
        fig_pc.update_traces(hovertemplate="<b>%{y}</b><br>₩%{x:,.0f}<extra></extra>")
        fig_pc.update_layout(height=250, margin=dict(t=30, l=150),
                             yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_pc, use_container_width=True)

    # 구성비 + 월별 추이
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("제품 구성비")
        top8 = by_amt.head(8)
        etc = by_amt.iloc[8:].sum()
        parts = pd.concat([top8, pd.Series({"기타": etc})])
        fig2 = px.pie(values=parts.values, names=parts.index, hole=0.4)
        fig2.update_traces(hovertemplate="<b>%{label}</b><br>₩%{value:,.0f} (%{percent})<extra></extra>")
        fig2.update_layout(height=400, margin=dict(t=30, b=30))
        st.plotly_chart(fig2, use_container_width=True)

    with col2:
        st.subheader("제품별 월별 추이 (Top 5)")
        top5 = by_amt.head(5).index.tolist()
        sub = so[so["os_name"].isin(top5)]
        if not sub.empty:
            m = sub.groupby(["period", "os_name"])["amount_krw"].sum().reset_index()
            fig3 = px.area(m, x="period", y="amount_krw", color="os_name",
                           labels={"amount_krw": "매출", "period": "월"})
            fig3.update_traces(hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>")
            fig3.update_layout(height=400, margin=dict(t=30, b=30),
                               xaxis=dict(type="category"))
            st.plotly_chart(fig3, use_container_width=True)

    # ── 제품별 평균 단가 ──
    st.subheader("제품별 평균 단가 (Top 15)")
    avg_price = so.groupby("os_name").agg(
        총금액=("amount_krw", "sum"), 총수량=("qty", "sum")
    )
    avg_price["평균단가"] = avg_price["총금액"] / avg_price["총수량"].replace(0, pd.NA)
    avg_top = avg_price.nlargest(15, "총금액")[["평균단가"]].reset_index()
    avg_top.columns = ["제품", "평균단가"]
    fig4 = px.bar(avg_top, y="제품", x="평균단가", orientation="h",
                  color_discrete_sequence=[C_ENDING],
                  labels={"평균단가": "평균단가 (원)"})
    fig4.update_traces(hovertemplate="<b>%{y}</b><br>평균단가: ₩%{x:,.0f}<extra></extra>")
    fig4.update_layout(
        height=max(350, len(avg_top) * 28), margin=dict(t=30, l=200),
        yaxis=dict(autorange="reversed"),
        xaxis=dict(tickformat=",.0f"),
    )
    st.plotly_chart(fig4, use_container_width=True)

    # ── 제품별 Backlog ──
    backlog = filt(load_backlog(), market, sectors, customers, period_col=None)
    if not backlog.empty:
        st.subheader("제품별 Backlog Top 10")
        bl_prod = backlog.groupby("os_name")["ending_amount"].sum().nlargest(10).reset_index()
        bl_prod.columns = ["제품", "Backlog 금액"]
        fig5 = px.bar(bl_prod, y="제품", x="Backlog 금액", orientation="h",
                      color_discrete_sequence=[C_DANGER])
        fig5.update_traces(hovertemplate="<b>%{y}</b><br>Backlog: ₩%{x:,.0f}<extra></extra>")
        fig5.update_layout(
            height=max(300, len(bl_prod) * 28), margin=dict(t=30, l=200),
            yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig5, use_container_width=True)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 제품 심층 분석
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    st.divider()
    st.header("제품 심층 분석")

    # ── 1. 제품 ABC 세그먼트 ──
    st.subheader("제품 세그먼트 (ABC)")
    total_sales = by_amt.sum()
    if total_sales > 0:
        abc_p = by_amt.reset_index()
        abc_p.columns = ["제품", "매출"]
        abc_p["누적비율"] = abc_p["매출"].cumsum() / total_sales * 100
        abc_p["등급"] = abc_p["누적비율"].apply(
            lambda p: "A" if p <= 80 else ("B" if p <= 95 else "C")
        )
        grade_p = abc_p.groupby("등급").agg(제품수=("제품", "count"), 매출합=("매출", "sum")).reset_index()
        grade_p["매출비중"] = (grade_p["매출합"] / total_sales * 100).round(1)

        pc1, pc2 = st.columns(2)
        with pc1:
            colors_abc = {"A": C_INPUT, "B": "#ff9800", "C": "#999999"}
            fig_pabc = px.bar(grade_p, x="등급", y="제품수", color="등급",
                              text="제품수", color_discrete_map=colors_abc)
            fig_pabc.update_traces(texttemplate="%{text}종", textposition="outside",
                                   hovertemplate="<b>%{x}등급</b><br>%{y}종<extra></extra>")
            fig_pabc.update_layout(height=300, margin=dict(t=30, b=30), showlegend=False)
            st.plotly_chart(fig_pabc, use_container_width=True)
        with pc2:
            for _, r in grade_p.iterrows():
                st.metric(f"{r['등급']}등급", f"{r['제품수']}종 · 매출 {r['매출비중']}%")

    # ── 2. 제품 성장률 ──
    st.subheader("제품 성장률")
    if so["period"].nunique() >= 2:
        periods = sorted(so["period"].unique())
        latest_p = periods[-1]
        prev_p = periods[-2]
        cur_p = so[so["period"] == latest_p].groupby("os_name")["amount_krw"].sum()
        prev_pp = so[so["period"] == prev_p].groupby("os_name")["amount_krw"].sum()
        growth_p = pd.DataFrame({"현재": cur_p, "이전": prev_pp}).fillna(0)
        growth_p["증감"] = growth_p["현재"] - growth_p["이전"]
        prev_safe = growth_p["이전"].where(growth_p["이전"] != 0)
        growth_p["성장률"] = (growth_p["증감"] / prev_safe * 100).round(1)
        growth_p = growth_p.reset_index().rename(columns={"os_name": "제품"})

        gp1, gp2 = st.columns(2)
        with gp1:
            st.markdown(f"**성장 Top 5** ({prev_p} → {latest_p})")
            top_gp = growth_p.dropna(subset=["성장률"]).nlargest(5, "성장률")
            if not top_gp.empty:
                fig_gp = px.bar(top_gp, y="제품", x="성장률", orientation="h",
                                color_discrete_sequence=[C_ENDING])
                fig_gp.update_traces(hovertemplate="<b>%{y}</b><br>%{x:+.1f}%<extra></extra>")
                fig_gp.update_layout(height=250, margin=dict(t=10, l=150),
                                     yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_gp, use_container_width=True)
        with gp2:
            st.markdown(f"**감소 Top 5** ({prev_p} → {latest_p})")
            bot_gp = growth_p.dropna(subset=["성장률"]).nsmallest(5, "성장률")
            bot_gp = bot_gp[bot_gp["성장률"] < 0]
            if not bot_gp.empty:
                fig_gpd = px.bar(bot_gp, y="제품", x="성장률", orientation="h",
                                 color_discrete_sequence=[C_DANGER])
                fig_gpd.update_traces(hovertemplate="<b>%{y}</b><br>%{x:+.1f}%<extra></extra>")
                fig_gpd.update_layout(height=250, margin=dict(t=10, l=150),
                                      yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_gpd, use_container_width=True)
            else:
                st.info("감소 제품 없음")

    # ── 3. 제품 × 고객 매트릭스 ──
    st.subheader("제품 × 고객 매트릭스")
    top_prods_hm = by_amt.head(10).index.tolist()
    top_custs_hm = so.groupby("customer_name")["amount_krw"].sum().nlargest(10).index.tolist()
    sub_hm = so[so["os_name"].isin(top_prods_hm) & so["customer_name"].isin(top_custs_hm)]
    if not sub_hm.empty:
        pivot_hm = sub_hm.pivot_table(
            index="customer_name", columns="os_name",
            values="amount_krw", aggfunc="sum", fill_value=0,
        )
        hover_hm = pivot_hm.map(lambda v: fmt_krw(v) if v > 0 else "-")
        fig_hm = go.Figure(data=go.Heatmap(
            z=pivot_hm.values,
            x=pivot_hm.columns.tolist(),
            y=pivot_hm.index.tolist(),
            text=hover_hm.values,
            texttemplate="%{text}",
            hovertemplate="<b>%{y}</b> × %{x}<br>₩%{z:,.0f}<extra></extra>",
            colorscale="Blues",
        ))
        fig_hm.update_layout(
            height=max(300, 50 + len(pivot_hm) * 35),
            margin=dict(t=20, b=30, l=150),
            xaxis=dict(title="제품", side="top"),
            yaxis=dict(title="고객", autorange="reversed"),
        )
        st.plotly_chart(fig_hm, use_container_width=True)
        st.caption("빈 칸(0원) = Cross-sell 기회 — 해당 고객에 해당 제품을 제안해 볼 수 있습니다")

    # ── 4. 제품별 납기 준수율 (OTD) ──
    st.subheader("제품별 납기 준수율 (OTD)")
    dn_raw = load_dn()
    dn_ej = enrich_dn(dn_raw, load_so())
    dn_f = filt(dn_ej, market, sectors, customers, period_col=None)
    so_del = so[["SO_ID", "line_item", "delivery_date"]].drop_duplicates()
    if not dn_f.empty and not so_del.empty:
        otd_raw = dn_f.merge(so_del, on=["SO_ID", "line_item"], how="left")
        otd_raw = otd_raw[otd_raw["delivery_date"].notna() & otd_raw["dispatch_date"].notna()]
        if not otd_raw.empty and "os_name" in otd_raw.columns:
            otd_raw["on_time"] = otd_raw["dispatch_date"] <= otd_raw["delivery_date"]
            otd_prod = otd_raw.groupby("os_name").agg(
                총건수=("on_time", "count"),
                정시건수=("on_time", "sum"),
            ).reset_index()
            otd_prod["OTD"] = (otd_prod["정시건수"] / otd_prod["총건수"] * 100).round(1)
            otd_prod = otd_prod.rename(columns={"os_name": "제품"})
            otd_prod = otd_prod[otd_prod["총건수"] >= 2].sort_values("OTD")

            if not otd_prod.empty:
                fig_otd_p = px.bar(otd_prod.head(15), y="제품", x="OTD", orientation="h",
                                   color="OTD",
                                   color_continuous_scale=["#d62728", "#ff9800", "#2ca02c"],
                                   range_color=[0, 100])
                fig_otd_p.update_traces(hovertemplate="<b>%{y}</b><br>OTD: %{x:.1f}%<extra></extra>")
                fig_otd_p.update_layout(height=max(300, len(otd_prod.head(15)) * 28),
                                        margin=dict(t=30, l=150),
                                        yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_otd_p, use_container_width=True)
                st.caption("OTD 낮은 제품 = 생산 병목 가능성 — 공정 리드타임 점검 필요")

    # ── 5. 제품별 수익성 ──
    st.subheader("제품별 수익성")
    po_data = load_po_detail()
    if not so.empty and not po_data.empty:
        margin_p = calc_margin(so, po_data)
        costed_p = margin_p[margin_p["has_cost"]]
        if not costed_p.empty:
            prod_margin = costed_p.groupby("os_name").agg(
                매출=("amount_krw", "sum"),
                원가=("po_total_ico", "sum"),
            ).reset_index()
            prod_margin["마진"] = prod_margin["매출"] - prod_margin["원가"]
            sales_safe = prod_margin["매출"].where(prod_margin["매출"] != 0)
            prod_margin["마진율"] = (prod_margin["마진"] / sales_safe * 100).round(1)
            prod_margin = prod_margin.rename(columns={"os_name": "제품"})

            pm1, pm2 = st.columns(2)
            with pm1:
                st.markdown("**고매출-저마진 제품** (마진율 < 20%)")
                low_mp = prod_margin[prod_margin["마진율"] < 20].sort_values("매출", ascending=False).head(10)
                if not low_mp.empty:
                    fig_lmp = px.bar(low_mp, y="제품", x="마진율", orientation="h",
                                     color="마진율",
                                     color_continuous_scale=["#d62728", "#ff9800", "#2ca02c"],
                                     range_color=[0, 50])
                    fig_lmp.update_traces(hovertemplate="<b>%{y}</b><br>마진율: %{x:.1f}%<extra></extra>")
                    fig_lmp.update_layout(height=300, margin=dict(t=10, l=150),
                                          yaxis=dict(autorange="reversed"))
                    st.plotly_chart(fig_lmp, use_container_width=True)
                else:
                    st.success("저마진 제품 없음")
            with pm2:
                st.markdown("**고마진 제품 Top 10**")
                top_mp = prod_margin.nlargest(10, "마진")
                fig_tmp = px.bar(top_mp, y="제품", x="마진", orientation="h",
                                 color_discrete_sequence=[C_ENDING])
                fig_tmp.update_traces(hovertemplate="<b>%{y}</b><br>마진: ₩%{x:,.0f}<extra></extra>")
                fig_tmp.update_layout(height=300, margin=dict(t=10, l=150),
                                      yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_tmp, use_container_width=True)

    # ── 6. 제품 집중도 ──
    st.subheader("제품 집중도")
    if total_sales > 0:
        top1_pct = by_amt.iloc[0] / total_sales * 100 if len(by_amt) >= 1 else 0
        top3_pct = by_amt.head(3).sum() / total_sales * 100 if len(by_amt) >= 1 else 0
        # HHI (Herfindahl-Hirschman Index)
        shares = by_amt / total_sales * 100
        hhi = (shares ** 2).sum()

        cc1, cc2, cc3 = st.columns(3)
        cc1.metric("Top 1 비중", f"{top1_pct:.1f}%", help=by_amt.index[0] if len(by_amt) else "")
        cc2.metric("Top 3 비중", f"{top3_pct:.1f}%")
        cc3.metric("HHI 지수", f"{hhi:.0f}",
                    help="< 1500: 분산 · 1500~2500: 보통 · > 2500: 집중")

    # ── 7. 신규 제품 추이 ──
    st.subheader("신규 제품 추이")
    so_all_raw = filt(load_so(), market, sectors, customers)
    if not so_all_raw.empty:
        first_prod = so_all_raw.groupby("os_name")["period"].min().reset_index()
        first_prod.columns = ["제품", "첫수주월"]
        monthly_new_p = first_prod.groupby("첫수주월").size().reset_index(name="신규제품수")
        monthly_active_p = so_all_raw.groupby("period")["os_name"].nunique().reset_index()
        monthly_active_p.columns = ["period", "활성제품수"]
        merged_np = monthly_active_p.merge(monthly_new_p, left_on="period", right_on="첫수주월", how="left").fillna(0)
        # 필터 기간과 일치하도록 표시 기간 한정
        display_periods = set(so["period"].unique())
        merged_np = merged_np[merged_np["period"].isin(display_periods)]

        fig_np = go.Figure()
        fig_np.add_trace(go.Bar(x=merged_np["period"], y=merged_np["신규제품수"],
                                name="신규 제품", marker_color=C_ENDING,
                                hovertemplate="<b>%{x}</b><br>신규: %{y}<extra></extra>"))
        fig_np.add_trace(go.Scatter(x=merged_np["period"], y=merged_np["활성제품수"],
                                    name="활성 제품 수", mode="lines+markers",
                                    line=dict(color=C_INPUT, width=2),
                                    hovertemplate="<b>%{x}</b><br>활성: %{y}<extra></extra>"))
        fig_np.update_layout(height=300, margin=dict(t=30, b=30),
                             xaxis=dict(type="category", title="월"),
                             yaxis=dict(title="제품 수"))
        st.plotly_chart(fig_np, use_container_width=True)

    # ── 8. 제품별 고객 수 ──
    st.subheader("제품별 고객 수")
    if not so.empty:
        prod_cust = so.groupby("os_name").agg(
            고객수=("customer_name", "nunique"),
            매출=("amount_krw", "sum"),
        ).reset_index().rename(columns={"os_name": "제품"})

        pc1, pc2 = st.columns(2)
        with pc1:
            single_cust = prod_cust[prod_cust["고객수"] == 1].sort_values("매출", ascending=False)
            st.markdown(f"**단일 고객 의존 제품: {len(single_cust)}종**")
            if not single_cust.empty:
                # 해당 고객명 추가
                single_ids = single_cust["제품"].tolist()
                sc_detail = so[so["os_name"].isin(single_ids)].groupby("os_name").agg(
                    고객=("customer_name", "first"), 매출=("amount_krw", "sum"),
                ).reset_index().rename(columns={"os_name": "제품"}).sort_values("매출", ascending=False)
                sc_detail["매출"] = sc_detail["매출"].apply(fmt_krw)
                st.dataframe(sc_detail.head(10), use_container_width=True, hide_index=True)
        with pc2:
            top_spread = prod_cust.nlargest(15, "고객수")
            fig_pcs = px.bar(top_spread, y="제품", x="고객수", orientation="h",
                             color_discrete_sequence=[C_INPUT])
            fig_pcs.update_traces(hovertemplate="<b>%{y}</b><br>%{x}개사<extra></extra>")
            fig_pcs.update_layout(height=max(300, len(top_spread) * 28),
                                  margin=dict(t=30, l=150),
                                  yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig_pcs, use_container_width=True)


# ═══════════════════════════════════════════════════════════════
# Page 4: 섹터 분석
# ═══════════════════════════════════════════════════════════════
def pg_sector(market, sectors, customers, year, month):
    st.title("섹터 분석")
    so = filt(load_so(), market, sectors, customers, year=year, month=month)
    if so.empty:
        st.info("데이터 없음")
        return

    by_sector = so.groupby("sector")["amount_krw"].sum().sort_values(ascending=False)
    total = by_sector.sum()

    # KPI (최대 6개 섹터)
    cols = st.columns(min(len(by_sector), 6))
    for i, (sec, amt) in enumerate(by_sector.items()):
        if i >= len(cols):
            break
        pct = amt / total * 100 if total else 0
        cols[i].metric(str(sec) or "(미분류)", f"{fmt_krw(amt)} ({pct:.1f}%)")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("섹터별 매출 비율")
        fig = px.pie(values=by_sector.values, names=by_sector.index)
        fig.update_traces(hovertemplate="<b>%{label}</b><br>₩%{value:,.0f} (%{percent})<extra></extra>")
        fig.update_layout(height=400, margin=dict(t=30, b=30))
        event_sec = st.plotly_chart(fig, use_container_width=True, on_select="rerun", key="sector_pie")

    with col2:
        st.subheader("섹터별 월별 추이")
        m = so.groupby(["period", "sector"])["amount_krw"].sum().reset_index()
        if not m.empty:
            fig2 = px.bar(m, x="period", y="amount_krw", color="sector",
                          barmode="stack",
                          labels={"amount_krw": "매출", "period": "월"})
            fig2.update_traces(hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>")
            fig2.update_layout(height=400, margin=dict(t=30, b=30),
                               xaxis=dict(type="category"))
            st.plotly_chart(fig2, use_container_width=True)

    # 섹터 드릴다운
    selected_sector = None
    if event_sec and event_sec.selection and event_sec.selection.points:
        selected_sector = event_sec.selection.points[0]["label"]
    if selected_sector:
        st.subheader(f"📌 {selected_sector} 상세")
        sub_sec = so[so["sector"] == selected_sector]
        dc1, dc2 = st.columns(2)
        with dc1:
            st.markdown("**제품 믹스**")
            sp = sub_sec.groupby("os_name")["amount_krw"].sum().nlargest(10).reset_index()
            fig_sp = px.bar(sp, y="os_name", x="amount_krw", orientation="h",
                            labels={"os_name": "제품", "amount_krw": "매출"},
                            color_discrete_sequence=[C_INPUT])
            fig_sp.update_traces(hovertemplate="<b>%{y}</b><br>₩%{x:,.0f}<extra></extra>")
            fig_sp.update_layout(height=350, margin=dict(t=30, l=150),
                                 yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig_sp, use_container_width=True)
        with dc2:
            st.markdown("**월별 추이**")
            sm = sub_sec.groupby("period")["amount_krw"].sum().reset_index()
            fig_sm = px.bar(sm, x="period", y="amount_krw",
                            labels={"period": "월", "amount_krw": "매출"},
                            color_discrete_sequence=[C_INPUT])
            fig_sm.update_traces(hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>")
            fig_sm.update_layout(height=350, margin=dict(t=30, b=30),
                                 xaxis=dict(type="category"))
            st.plotly_chart(fig_sm, use_container_width=True)
        st.markdown("**주요 고객**")
        sc = sub_sec.groupby("customer_name")["amount_krw"].sum().nlargest(5).reset_index()
        sc.columns = ["고객", "매출"]
        fig_sc = px.bar(sc, y="고객", x="매출", orientation="h",
                        color_discrete_sequence=[C_OUTPUT])
        fig_sc.update_traces(hovertemplate="<b>%{y}</b><br>₩%{x:,.0f}<extra></extra>")
        fig_sc.update_layout(height=250, margin=dict(t=30, l=150),
                             yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_sc, use_container_width=True)

    # 섹터별 제품 믹스
    st.subheader("섹터별 제품 믹스")
    top_prods = so.groupby("os_name")["amount_krw"].sum().nlargest(8).index
    mix = (so[so["os_name"].isin(top_prods)]
           .groupby(["sector", "os_name"])["amount_krw"].sum().reset_index())
    if not mix.empty:
        fig3 = px.bar(mix, x="sector", y="amount_krw", color="os_name",
                      barmode="group", labels={"amount_krw": "매출"})
        fig3.update_traces(hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>")
        fig3.update_layout(height=400, margin=dict(t=30, b=30))
        st.plotly_chart(fig3, use_container_width=True)

    # ── 섹터별 Backlog ──
    backlog = filt(load_backlog(), market, sectors, customers, period_col=None)
    if not backlog.empty:
        st.subheader("섹터별 Backlog 현황")
        bl_sec = backlog.groupby("sector").agg(
            건수=("SO_ID", "nunique"),
            금액=("ending_amount", "sum"),
        ).sort_values("금액", ascending=False).reset_index()
        bl_sec["비중"] = (bl_sec["금액"] / bl_sec["금액"].sum() * 100).round(1).astype(str) + "%"
        bl_sec["금액"] = bl_sec["금액"].apply(fmt_krw)
        bl_sec.rename(columns={"sector": "섹터"}, inplace=True)
        st.dataframe(bl_sec, use_container_width=True, hide_index=True)

    # ── 섹터별 평균 주문 규모 ──
    st.subheader("섹터별 평균 주문 규모")
    sec_avg = so.groupby("sector").agg(
        총금액=("amount_krw", "sum"), 주문건수=("SO_ID", "nunique"),
    )
    sec_avg["평균주문액"] = sec_avg["총금액"] / sec_avg["주문건수"]
    sec_avg = sec_avg.sort_values("평균주문액", ascending=False).reset_index()
    fig4 = px.bar(sec_avg, x="sector", y="평균주문액",
                  labels={"sector": "섹터", "평균주문액": "평균 주문 금액"},
                  color_discrete_sequence=[C_PURPLE])
    fig4.update_traces(hovertemplate="<b>%{x}</b><br>평균주문액: ₩%{y:,.0f}<extra></extra>")
    fig4.update_layout(height=350, margin=dict(t=30, b=30))
    st.plotly_chart(fig4, use_container_width=True)


# ═══════════════════════════════════════════════════════════════
# Page 5: 고객 분석
# ═══════════════════════════════════════════════════════════════
def pg_customer(market, sectors, customers, year, month):
    st.title("고객 분석")
    so = filt(load_so(), market, sectors, customers, year=year, month=month)
    if so.empty:
        st.info("데이터 없음")
        return

    by_cust = so.groupby("customer_name")["amount_krw"].sum().sort_values(ascending=False)

    # KPI
    c1, c2, c3 = st.columns(3)
    c1.metric("총 고객 수", f"{so['customer_name'].nunique()}")
    c2.metric("Top 고객", by_cust.index[0] if len(by_cust) else "-")
    if year and year != "전체" and month and month != "전체":
        # 특정 월 선택 — so가 이미 해당 월만 포함
        c3.metric(f"{year}-{month} 수주 고객 수", f"{so['customer_name'].nunique()}")
    elif year and year != "전체":
        # 연도만 선택 — so가 해당 연도만 포함, 최근 월 기준
        latest = so["period"].max() if not so.empty else None
        if latest:
            month_custs = so[so["period"] == latest]["customer_name"].nunique()
            c3.metric(f"{latest} 수주 고객 수", f"{month_custs}")
        else:
            c3.metric(f"{year} 수주 고객 수", "0")
    else:
        month_custs = so[so["period"] == _THIS_MONTH]["customer_name"].nunique()
        c3.metric("금월 수주 고객 수", f"{month_custs}")

    # Top 15
    st.subheader("고객별 매출 Top 15")
    top15 = by_cust.head(15).reset_index()
    top15.columns = ["고객", "매출"]
    fig = px.bar(top15, y="고객", x="매출", orientation="h",
                 color_discrete_sequence=[C_INPUT])
    fig.update_traces(hovertemplate="<b>%{y}</b><br>매출: ₩%{x:,.0f}<extra></extra>")
    fig.update_layout(
        height=max(400, len(top15) * 32),
        margin=dict(t=30, l=200),
        yaxis=dict(autorange="reversed"),
    )
    event_cust = st.plotly_chart(fig, use_container_width=True, on_select="rerun", key="customer_top15")
    selected_customer = None
    if event_cust and event_cust.selection and event_cust.selection.points:
        selected_customer = event_cust.selection.points[0]["y"]
    if selected_customer:
        st.subheader(f"📌 {selected_customer} 상세")
        sub_cust = so[so["customer_name"] == selected_customer]
        dc1, dc2 = st.columns(2)
        with dc1:
            st.markdown("**월별 매출 추이**")
            cm = sub_cust.groupby("period")["amount_krw"].sum().reset_index()
            fig_cm = px.bar(cm, x="period", y="amount_krw",
                            labels={"period": "월", "amount_krw": "매출"},
                            color_discrete_sequence=[C_INPUT])
            fig_cm.update_traces(hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>")
            fig_cm.update_layout(height=300, margin=dict(t=30, b=30),
                                 xaxis=dict(type="category"))
            st.plotly_chart(fig_cm, use_container_width=True)
        with dc2:
            st.markdown("**제품 믹스**")
            cp = sub_cust.groupby("os_name")["amount_krw"].sum().nlargest(8).reset_index()
            fig_cp = px.pie(cp, names="os_name", values="amount_krw", hole=0.4)
            fig_cp.update_traces(hovertemplate="<b>%{label}</b><br>₩%{value:,.0f} (%{percent})<extra></extra>")
            fig_cp.update_layout(height=300, margin=dict(t=30, b=30))
            st.plotly_chart(fig_cp, use_container_width=True)
        # 백로그 현황
        backlog_c = filt(load_backlog(), market, sectors, customers, period_col=None)
        cust_bl = backlog_c[backlog_c["customer_name"] == selected_customer] if not backlog_c.empty else pd.DataFrame()
        if not cust_bl.empty:
            st.markdown(f"**Backlog 현황** — {len(cust_bl)}건 / {fmt_krw(cust_bl['ending_amount'].sum())}")
            bl_t = cust_bl[["SO_ID", "os_name", "delivery_date", "ending_qty", "ending_amount"]].copy()
            bl_t.columns = ["SO_ID", "품목", "납기일", "잔여수량", "잔여금액"]
            bl_t["납기일"] = bl_t["납기일"].apply(fmt_date)
            bl_t["잔여수량"] = bl_t["잔여수량"].apply(lambda x: f"{int(x):,}")
            bl_t["잔여금액"] = bl_t["잔여금액"].apply(fmt_num)
            st.dataframe(bl_t, use_container_width=True, hide_index=True)

    # Pareto (상위 20)
    st.subheader("고객 집중도 (Pareto)")
    _all_cust_total = by_cust.sum()
    pareto = by_cust.head(20).reset_index()
    pareto.columns = ["고객", "매출"]
    pareto["누적비율"] = pareto["매출"].cumsum() / _all_cust_total * 100 if _all_cust_total else 0
    fig2 = go.Figure()
    fig2.add_trace(go.Bar(x=pareto["고객"], y=pareto["매출"],
                          name="매출", marker_color=C_INPUT,
                          hovertemplate="<b>%{x}</b><br>매출: ₩%{y:,.0f}<extra></extra>"))
    fig2.add_trace(go.Scatter(
        x=pareto["고객"], y=pareto["누적비율"], name="누적 %",
        yaxis="y2", mode="lines+markers", line=dict(color=C_DANGER),
        hovertemplate="<b>%{x}</b><br>누적: %{y:.1f}%<extra></extra>",
    ))
    fig2.update_layout(
        yaxis2=dict(title="누적 %", overlaying="y", side="right", range=[0, 105]),
        height=400, margin=dict(t=30, b=30),
    )
    st.plotly_chart(fig2, use_container_width=True)

    # 고객 상세 + Backlog
    st.subheader("고객 상세")
    detail = (
        so.groupby("customer_name")
        .agg(주문건수=("SO_ID", "nunique"), 총수량=("qty", "sum"),
             총금액=("amount_krw", "sum"), 최근수주월=("period", "max"))
        .sort_values("총금액", ascending=False)
        .reset_index()
    )
    detail["평균주문액"] = detail["총금액"] / detail["주문건수"]
    # Backlog 병합
    backlog = filt(load_backlog(), market, sectors, customers, period_col=None)
    if not backlog.empty:
        bl_cust = backlog.groupby("customer_name").agg(
            Backlog건수=("SO_ID", "nunique"),
            Backlog금액=("ending_amount", "sum"),
        ).reset_index()
        detail = detail.merge(bl_cust, on="customer_name", how="left").fillna(0)
        detail["Backlog금액"] = detail["Backlog금액"].apply(fmt_krw)
        detail["Backlog건수"] = detail["Backlog건수"].apply(lambda x: f"{int(x)}")
    detail.rename(columns={"customer_name": "고객명"}, inplace=True)
    detail["총금액"] = detail["총금액"].apply(fmt_krw)
    detail["평균주문액"] = detail["평균주문액"].apply(fmt_krw)
    detail["총수량"] = detail["총수량"].apply(lambda x: f"{int(x):,}")
    st.dataframe(detail, use_container_width=True, hide_index=True)

    # ── 고객별 월별 매출 추이 (Top 5) ──
    st.subheader("고객별 월별 매출 추이 (Top 5)")
    top5_cust = by_cust.head(5).index.tolist()
    sub = so[so["customer_name"].isin(top5_cust)]
    if not sub.empty:
        m = sub.groupby(["period", "customer_name"])["amount_krw"].sum().reset_index()
        fig3 = px.line(m, x="period", y="amount_krw", color="customer_name",
                       markers=True, labels={"amount_krw": "매출", "period": "월"})
        fig3.update_traces(hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>")
        fig3.update_layout(height=400, margin=dict(t=30, b=30),
                           xaxis=dict(type="category"))
        st.plotly_chart(fig3, use_container_width=True)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 고객 심층 분석
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    st.divider()
    st.header("고객 심층 분석")

    # ── 1. ABC 세그먼트 ──
    st.subheader("고객 세그먼트 (ABC)")
    abc = by_cust.reset_index()
    abc.columns = ["고객", "매출"]
    total_sales = abc["매출"].sum()
    if total_sales > 0:
        abc["누적비율"] = abc["매출"].cumsum() / total_sales * 100
        abc["등급"] = abc["누적비율"].apply(
            lambda p: "A" if p <= 80 else ("B" if p <= 95 else "C")
        )
        grade_summary = abc.groupby("등급").agg(고객수=("고객", "count"), 매출합=("매출", "sum")).reset_index()
        grade_summary["매출비중"] = (grade_summary["매출합"] / total_sales * 100).round(1)

        gc1, gc2 = st.columns(2)
        with gc1:
            colors_abc = {"A": C_INPUT, "B": "#ff9800", "C": "#999999"}
            fig_abc = px.bar(grade_summary, x="등급", y="고객수", color="등급",
                             text="고객수", color_discrete_map=colors_abc)
            fig_abc.update_traces(texttemplate="%{text}개사", textposition="outside",
                                  hovertemplate="<b>%{x}등급</b><br>%{y}개사<extra></extra>")
            fig_abc.update_layout(height=300, margin=dict(t=30, b=30), showlegend=False)
            st.plotly_chart(fig_abc, use_container_width=True)
        with gc2:
            for _, r in grade_summary.iterrows():
                st.metric(f"{r['등급']}등급", f"{r['고객수']}개사 · 매출 {r['매출비중']}%")

        with st.expander("ABC 상세"):
            abc_tbl = abc.copy()
            abc_tbl["매출"] = abc_tbl["매출"].apply(fmt_krw)
            st.dataframe(abc_tbl[["고객", "매출", "등급"]], use_container_width=True, hide_index=True)

    # ── 2. 고객 성장률 ──
    st.subheader("고객 성장률")
    if not so.empty and so["period"].nunique() >= 2:
        periods = sorted(so["period"].unique())
        # 최근 2개 기간 비교
        if len(periods) >= 2:
            latest_p = periods[-1]
            prev_p = periods[-2]
            cur = so[so["period"] == latest_p].groupby("customer_name")["amount_krw"].sum()
            prev = so[so["period"] == prev_p].groupby("customer_name")["amount_krw"].sum()
            growth = pd.DataFrame({"현재": cur, "이전": prev}).fillna(0)
            growth["증감"] = growth["현재"] - growth["이전"]
            prev_safe = growth["이전"].where(growth["이전"] != 0)
            growth["성장률"] = (growth["증감"] / prev_safe * 100).round(1)
            growth = growth.reset_index().rename(columns={"customer_name": "고객"})

            gr1, gr2 = st.columns(2)
            with gr1:
                st.markdown(f"**성장 Top 5** ({prev_p} → {latest_p})")
                top_growth = growth.dropna(subset=["성장률"]).nlargest(5, "성장률")
                if not top_growth.empty:
                    fig_gr = px.bar(top_growth, y="고객", x="성장률", orientation="h",
                                    color_discrete_sequence=[C_ENDING])
                    fig_gr.update_traces(hovertemplate="<b>%{y}</b><br>%{x:+.1f}%<extra></extra>")
                    fig_gr.update_layout(height=250, margin=dict(t=10, l=150),
                                         yaxis=dict(autorange="reversed"))
                    st.plotly_chart(fig_gr, use_container_width=True)
            with gr2:
                st.markdown(f"**감소 Top 5** ({prev_p} → {latest_p})")
                bot_growth = growth.dropna(subset=["성장률"]).nsmallest(5, "성장률")
                bot_growth = bot_growth[bot_growth["성장률"] < 0]
                if not bot_growth.empty:
                    fig_gd = px.bar(bot_growth, y="고객", x="성장률", orientation="h",
                                    color_discrete_sequence=[C_DANGER])
                    fig_gd.update_traces(hovertemplate="<b>%{y}</b><br>%{x:+.1f}%<extra></extra>")
                    fig_gd.update_layout(height=250, margin=dict(t=10, l=150),
                                         yaxis=dict(autorange="reversed"))
                    st.plotly_chart(fig_gd, use_container_width=True)
                else:
                    st.info("감소 고객 없음")

    # ── 3. 신규 vs 기존 고객 ──
    st.subheader("신규 vs 기존 고객")
    so_all_raw = filt(load_so(), market, sectors, customers)
    if not so_all_raw.empty:
        first_order = so_all_raw.groupby("customer_name")["period"].min().reset_index()
        first_order.columns = ["고객", "첫수주월"]
        monthly_new = first_order.groupby("첫수주월").size().reset_index(name="신규고객수")
        # 월별 전체 활성 고객 수
        monthly_active = so_all_raw.groupby("period")["customer_name"].nunique().reset_index()
        monthly_active.columns = ["period", "활성고객수"]
        merged_nc = monthly_active.merge(monthly_new, left_on="period", right_on="첫수주월", how="left").fillna(0)
        merged_nc["기존고객수"] = merged_nc["활성고객수"] - merged_nc["신규고객수"]
        # 필터 기간과 일치하도록 표시 기간 한정
        display_periods = set(so["period"].unique())
        merged_nc = merged_nc[merged_nc["period"].isin(display_periods)]

        fig_nc = go.Figure()
        fig_nc.add_trace(go.Bar(x=merged_nc["period"], y=merged_nc["기존고객수"],
                                name="기존 고객", marker_color=C_INPUT,
                                hovertemplate="<b>%{x}</b><br>기존: %{y}<extra></extra>"))
        fig_nc.add_trace(go.Bar(x=merged_nc["period"], y=merged_nc["신규고객수"],
                                name="신규 고객", marker_color=C_ENDING,
                                hovertemplate="<b>%{x}</b><br>신규: %{y}<extra></extra>"))
        fig_nc.update_layout(barmode="stack", height=300, margin=dict(t=30, b=30),
                             xaxis=dict(type="category", title="월"),
                             yaxis=dict(title="고객 수"))
        st.plotly_chart(fig_nc, use_container_width=True)

    # ── 4. 고객 리텐션 ──
    st.subheader("고객 리텐션")
    if not so_all_raw.empty and so_all_raw["period"].nunique() >= 3:
        all_periods = sorted(so_all_raw["period"].unique())
        retention_data = []
        for i, p in enumerate(all_periods):
            if i < 1:
                continue
            prev_custs = set(so_all_raw[so_all_raw["period"] < p]["customer_name"].unique())
            if not prev_custs:
                continue
            curr_custs = set(so_all_raw[so_all_raw["period"] == p]["customer_name"].unique())
            retained = prev_custs & curr_custs
            retention_rate = len(retained) / len(prev_custs) * 100 if prev_custs else 0
            retention_data.append({"월": p, "리텐션율": round(retention_rate, 1),
                                   "활성": len(curr_custs), "재구매": len(retained)})
        if retention_data:
            ret_df = pd.DataFrame(retention_data)
            # 필터 기간과 일치하도록 표시 기간 한정
            ret_df = ret_df[ret_df["월"].isin(display_periods)]
            if not ret_df.empty:
                fig_ret = go.Figure()
                fig_ret.add_trace(go.Scatter(
                    x=ret_df["월"], y=ret_df["리텐션율"], mode="lines+markers",
                    line=dict(color=C_PURPLE, width=2),
                    hovertemplate="<b>%{x}</b><br>리텐션: %{y:.1f}%<extra></extra>",
                ))
                fig_ret.update_layout(height=300, margin=dict(t=30, b=30),
                                      xaxis=dict(type="category"),
                                      yaxis=dict(title="리텐션율 (%)", range=[0, 105]))
                st.plotly_chart(fig_ret, use_container_width=True)
                st.caption("리텐션율 = 이전까지 수주 이력이 있는 고객 중 해당 월에도 수주한 비율")

    # ── 5. RFM 스코어 ──
    st.subheader("RFM 분석")
    if not so.empty:
        # Recency 기준: 데이터 최신월 (DB 동기화 지연/과거 필터 대응)
        ref_period = so["period"].max()
        rfm = so.groupby("customer_name").agg(
            최근수주=("period", "max"),
            주문횟수=("SO_ID", "nunique"),
            총금액=("amount_krw", "sum"),
        ).reset_index()
        rfm.columns = ["고객", "최근수주", "Frequency", "Monetary"]
        # Recency: 최근 수주로부터 몇 개월
        rfm["Recency"] = rfm["최근수주"].apply(
            lambda p: (pd.Timestamp(ref_period + "-01") - pd.Timestamp(p + "-01")).days // 30
        )
        # 점수 (각 지표를 4분위로 스코어링, 고객 4명 미만 시 분위 수 축소)
        n_customers = len(rfm)
        q = min(4, n_customers) if n_customers >= 2 else 0
        if q >= 2:
            labels = list(range(1, q + 1))
            for col, ascending in [("Recency", True), ("Frequency", False), ("Monetary", False)]:
                rfm[f"{col}_점수"] = pd.qcut(rfm[col].rank(method="first"), q=q, labels=labels).astype(int)
                if ascending:
                    rfm[f"{col}_점수"] = (q + 1) - rfm[f"{col}_점수"]
        else:
            for col in ["Recency", "Frequency", "Monetary"]:
                rfm[f"{col}_점수"] = 2  # 고객 1명이면 중간 점수
        rfm["RFM점수"] = rfm["Recency_점수"] + rfm["Frequency_점수"] + rfm["Monetary_점수"]
        # 등급 경계를 q에 비례하여 산출 (q=4 기준 10/7/4 → 83%/58%/33%)
        _max = q * 3 if q >= 2 else 6
        _vip_th = round(_max * 10 / 12)
        _good_th = round(_max * 7 / 12)
        _normal_th = round(_max * 4 / 12)
        rfm["등급"] = rfm["RFM점수"].apply(
            lambda s: "VIP" if s >= _vip_th else ("우수" if s >= _good_th else ("보통" if s >= _normal_th else "관심"))
        )

        rc1, rc2 = st.columns(2)
        with rc1:
            rfm_summary = rfm["등급"].value_counts().reindex(["VIP", "우수", "보통", "관심"]).fillna(0).astype(int)
            rfm_colors = {"VIP": C_INPUT, "우수": C_ENDING, "보통": "#ff9800", "관심": C_DANGER}
            fig_rfm = px.bar(x=rfm_summary.index, y=rfm_summary.values, color=rfm_summary.index,
                             color_discrete_map=rfm_colors, text=rfm_summary.values)
            fig_rfm.update_traces(texttemplate="%{text}개사", textposition="outside",
                                   hovertemplate="<b>%{x}</b><br>%{y}개사<extra></extra>")
            fig_rfm.update_layout(height=300, margin=dict(t=30, b=30), showlegend=False,
                                  xaxis=dict(title=""), yaxis=dict(title="고객 수"))
            st.plotly_chart(fig_rfm, use_container_width=True)
        with rc2:
            st.markdown(
                f"**RFM 등급 기준** (만점 {_max}점)\n"
                f"- **VIP** ({_vip_th}점 이상): 최근 구매, 자주, 고액\n"
                f"- **우수** ({_good_th}점 이상): 양호한 고객\n"
                f"- **보통** ({_normal_th}점 이상): 평균적 고객\n"
                f"- **관심** ({_normal_th}점 미만): 이탈 위험"
            )
        with st.expander("RFM 상세"):
            rfm_tbl = rfm[["고객", "최근수주", "Recency", "Frequency", "Monetary", "RFM점수", "등급"]].copy()
            rfm_tbl = rfm_tbl.sort_values("RFM점수", ascending=False)
            rfm_tbl["Monetary"] = rfm_tbl["Monetary"].apply(fmt_krw)
            rfm_tbl.rename(columns={"Recency": "최근(개월)", "Frequency": "주문횟수"}, inplace=True)
            st.dataframe(rfm_tbl, use_container_width=True, hide_index=True)

    # ── 6. 고객별 제품 다양성 ──
    st.subheader("고객별 제품 다양성")
    if not so.empty:
        diversity = so.groupby("customer_name").agg(
            제품종류=("os_name", "nunique"),
            매출=("amount_krw", "sum"),
        ).reset_index().rename(columns={"customer_name": "고객"})
        diversity = diversity.sort_values("매출", ascending=False)

        dv1, dv2 = st.columns(2)
        with dv1:
            single = diversity[diversity["제품종류"] == 1]
            st.metric("단일 제품 고객", f"{len(single)}개사",
                      help="1가지 제품만 구매 → Cross-sell 기회")
            multi = diversity[diversity["제품종류"] >= 3]
            st.metric("다품목 고객 (3+)", f"{len(multi)}개사")
        with dv2:
            fig_dv = px.scatter(diversity.head(30), x="매출", y="제품종류",
                                text="고객", size="매출",
                                labels={"매출": "매출(KRW)", "제품종류": "제품 종류 수"},
                                color_discrete_sequence=[C_INPUT])
            fig_dv.update_traces(textposition="top center", textfont_size=9,
                                  hovertemplate="<b>%{text}</b><br>매출: ₩%{x:,.0f}<br>제품: %{y}종<extra></extra>")
            fig_dv.update_layout(height=350, margin=dict(t=30, b=30))
            st.plotly_chart(fig_dv, use_container_width=True)

    # ── 7. 고객별 납기 준수율 ──
    st.subheader("고객별 납기 준수율 (OTD)")
    dn_raw = load_dn()
    dn_ej = enrich_dn(dn_raw, load_so())
    dn_f = filt(dn_ej, market, sectors, customers, period_col=None)
    so_del = so[["SO_ID", "line_item", "delivery_date"]].drop_duplicates()
    if not dn_f.empty and not so_del.empty:
        otd_raw = dn_f.merge(so_del, on=["SO_ID", "line_item"], how="left")
        otd_raw = otd_raw[otd_raw["delivery_date"].notna() & otd_raw["dispatch_date"].notna()]
        if not otd_raw.empty:
            otd_raw["on_time"] = otd_raw["dispatch_date"] <= otd_raw["delivery_date"]
            otd_cust = otd_raw.groupby("customer_name").agg(
                총건수=("on_time", "count"),
                정시건수=("on_time", "sum"),
            ).reset_index()
            otd_cust["OTD"] = (otd_cust["정시건수"] / otd_cust["총건수"] * 100).round(1)
            otd_cust = otd_cust.rename(columns={"customer_name": "고객"})
            otd_cust = otd_cust[otd_cust["총건수"] >= 2].sort_values("OTD")  # 2건 이상만

            if not otd_cust.empty:
                avg_otd = otd_cust["정시건수"].sum() / otd_cust["총건수"].sum() * 100
                st.metric("전체 OTD", f"{avg_otd:.1f}%", help="약속 납기 이내 출고 비율")

                fig_otd = px.bar(otd_cust.head(20), y="고객", x="OTD", orientation="h",
                                 color="OTD",
                                 color_continuous_scale=["#d62728", "#ff9800", "#2ca02c"],
                                 range_color=[0, 100])
                fig_otd.update_traces(hovertemplate="<b>%{y}</b><br>OTD: %{x:.1f}%<extra></extra>")
                fig_otd.update_layout(height=max(300, len(otd_cust.head(20)) * 28),
                                      margin=dict(t=30, l=150),
                                      yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_otd, use_container_width=True)
            else:
                st.info("OTD 분석 대상 없음")
        else:
            st.info("납기/출고 데이터 부족")

    # ── 8. 고객 수익성 ──
    st.subheader("고객 수익성")
    po_data = load_po_detail()
    if not so.empty and not po_data.empty:
        margin = calc_margin(so, po_data)
        costed = margin[margin["has_cost"]]
        if not costed.empty:
            cust_margin = costed.groupby("customer_name").agg(
                매출=("amount_krw", "sum"),
                원가=("po_total_ico", "sum"),
            ).reset_index()
            cust_margin["마진"] = cust_margin["매출"] - cust_margin["원가"]
            sales_safe = cust_margin["매출"].where(cust_margin["매출"] != 0)
            cust_margin["마진율"] = (cust_margin["마진"] / sales_safe * 100).round(1)
            cust_margin = cust_margin.rename(columns={"customer_name": "고객"})
            cust_margin = cust_margin.sort_values("매출", ascending=False)

            cm1, cm2 = st.columns(2)
            with cm1:
                st.markdown("**고매출-저마진 고객** (마진율 < 20%)")
                low = cust_margin[cust_margin["마진율"] < 20].head(10)
                if not low.empty:
                    fig_lm = px.bar(low, y="고객", x="마진율", orientation="h",
                                    color="마진율",
                                    color_continuous_scale=["#d62728", "#ff9800", "#2ca02c"],
                                    range_color=[0, 50])
                    fig_lm.update_traces(hovertemplate="<b>%{y}</b><br>마진율: %{x:.1f}%<extra></extra>")
                    fig_lm.update_layout(height=300, margin=dict(t=10, l=150),
                                         yaxis=dict(autorange="reversed"))
                    st.plotly_chart(fig_lm, use_container_width=True)
                else:
                    st.success("저마진 고객 없음")
            with cm2:
                st.markdown("**고마진 고객 Top 10**")
                top_m = cust_margin.nlargest(10, "마진")
                fig_tm = px.bar(top_m, y="고객", x="마진", orientation="h",
                                color_discrete_sequence=[C_ENDING])
                fig_tm.update_traces(hovertemplate="<b>%{y}</b><br>마진: ₩%{x:,.0f}<extra></extra>")
                fig_tm.update_layout(height=300, margin=dict(t=10, l=150),
                                     yaxis=dict(autorange="reversed"))
                st.plotly_chart(fig_tm, use_container_width=True)
        else:
            st.info("원가 데이터 없음")


# ═══════════════════════════════════════════════════════════════
# 커버리지 / 마진 순수 계산 함수 (테스트 가능)
# ═══════════════════════════════════════════════════════════════
def calc_coverage(so: pd.DataFrame, po: pd.DataFrame,
                  po_all_status: pd.DataFrame | None = None) -> pd.DataFrame:
    """SO_ID 단위로 SO↔PO 커버리지 계산. 순수 함수.

    커버리지 판정: PO 존재 여부 + PO Status 기반 (수량 비교 아님).
    PO line_item은 SO와 1:1 대응하지 않으므로 (본체+부속 합산 발주 등) 수량 비교 불가.

    - PO 미등록: SO시트에만 있고 PO시트에 아예 없음 (데이터 입력 누락 리스크)
    - 미발주: PO Open 상태만 (PO 생성은 됐으나 공장 발주 전)
    - 발주 진행중: PO 존재, Open/Sent 포함
    - 발주 확정: PO 존재, 모두 Confirmed/Invoiced
    - 발주취소: 모든 PO가 Cancelled (po_all_status로 판별, 결과에서 제외)
    """
    if so.empty:
        return pd.DataFrame()
    # 누락 컬럼 기본값 보장
    if "po_receipt_date" not in so.columns:
        so = so.assign(po_receipt_date=pd.NaT)
    for _col in ("po_ids", "open_po_ids"):
        if _col not in po.columns:
            po = po.assign(**{_col: ""})
    if "factory_order_date" not in po.columns:
        po = po.assign(factory_order_date=pd.NaT)
    # 발주취소 SO 제외 (모든 PO가 Cancelled인 SO)
    if po_all_status is not None and not po_all_status.empty:
        cancelled_only = (
            po_all_status.groupby("SO_ID")["po_status"]
            .apply(lambda s: s.astype(str).str.startswith("Cancelled").all())
            .reset_index()
        )
        cancelled_so_ids = set(cancelled_only[cancelled_only["po_status"]]["SO_ID"])
        if cancelled_so_ids:
            so = so[~so["SO_ID"].isin(cancelled_so_ids)]
    if so.empty:
        return pd.DataFrame()
    # SO를 SO_ID 단위로 집계
    so_agg = so.groupby("SO_ID").agg(
        customer_name=("customer_name", "first"),
        os_name=("os_name", lambda x: ", ".join(sorted(x.unique()))),
        sector=("sector", "first"),
        market=("market", "first"),
        period=("period", "first"),
        qty=("qty", "sum"),
        amount_krw=("amount_krw", "sum"),
        delivery_date=("delivery_date", "min"),
        po_receipt_date=("po_receipt_date", "min"),
        status=("status", "first"),
    ).reset_index()
    # PO 조인 (PO는 이미 SO_ID 단위, Cancelled 제외됨)
    if not po.empty:
        po_dedup = po.groupby("SO_ID").agg(
            po_qty=("po_qty", "sum"),
            po_total_ico=("po_total_ico", "sum"),
            po_statuses=("po_statuses", lambda x: ",".join(sorted(set(",".join(x).split(",")) - {""}))),
            po_ids=("po_ids", lambda x: ",".join(sorted(set(",".join(x).split(",")) - {""}))),
            open_po_ids=("open_po_ids", lambda x: ",".join(sorted(set(",".join(x).split(",")) - {""}))),
            factory_order_date=("factory_order_date", "min"),
        ).reset_index()
        merged = so_agg.merge(po_dedup, on="SO_ID", how="left")
    else:
        merged = so_agg.copy()
    merged["po_qty"] = merged["po_qty"].fillna(0) if "po_qty" in merged.columns else 0
    merged["po_total_ico"] = merged["po_total_ico"].fillna(0) if "po_total_ico" in merged.columns else 0
    merged["po_statuses"] = merged["po_statuses"].fillna("") if "po_statuses" in merged.columns else ""
    merged["po_ids"] = merged["po_ids"].fillna("") if "po_ids" in merged.columns else ""
    merged["open_po_ids"] = merged["open_po_ids"].fillna("") if "open_po_ids" in merged.columns else ""
    if "factory_order_date" not in merged.columns:
        merged["factory_order_date"] = pd.NaT
    # 커버리지 판정: PO 존재 + Status 기반
    def _coverage_status(row):
        # PO 미등록 = SO시트에만 있고 PO시트에 아예 없음 (데이터 입력 누락)
        # Open = PO 생성만 됨, 공장 발주 전 → 미발주
        # Sent = 공장에 발주함 → 발주 진행중
        # Confirmed/Invoiced = 공장 확인/출고 → 발주 확정
        if not row["po_statuses"]:
            return "PO 미등록"
        statuses = {s.strip() for s in row["po_statuses"].split(",") if s.strip()}
        has_open = "Open" in statuses
        has_ordered = any(
            s.startswith("Sent") or s.startswith("Confirmed") or s.startswith("Invoiced")
            for s in statuses
        )
        if has_open and has_ordered:
            return "부분 발주"  # Open + 발주된 PO 혼합
        if has_open:
            return "미발주"    # Open만
        all_confirmed = all(
            s.startswith("Confirmed") or s.startswith("Invoiced") for s in statuses
        )
        return "발주 확정" if all_confirmed else "발주 진행중"
    merged["coverage_status"] = merged.apply(_coverage_status, axis=1)
    return merged


def calc_margin(so: pd.DataFrame, po: pd.DataFrame) -> pd.DataFrame:
    """SO_ID 단위로 SO↔PO 마진 계산. 순수 함수.

    반환값은 SO_ID별 1행.
    """
    if so.empty:
        return pd.DataFrame()
    # SO를 SO_ID 단위로 집계
    so_agg = so.groupby("SO_ID").agg(
        customer_name=("customer_name", "first"),
        os_name=("os_name", lambda x: ", ".join(sorted(x.unique()))),
        sector=("sector", "first"),
        market=("market", "first"),
        period=("period", "first"),
        qty=("qty", "sum"),
        amount_krw=("amount_krw", "sum"),
    ).reset_index()
    # PO 조인
    if not po.empty:
        po_dedup = po.groupby("SO_ID").agg(
            po_total_ico=("po_total_ico", "sum"),
        ).reset_index()
        merged = so_agg.merge(po_dedup, on="SO_ID", how="left")
    else:
        merged = so_agg.copy()
    merged["po_total_ico"] = merged["po_total_ico"].fillna(0) if "po_total_ico" in merged.columns else 0
    merged["has_cost"] = merged["po_total_ico"] > 0
    merged["margin_amount"] = merged["amount_krw"] - merged["po_total_ico"]
    margin_raw = merged["margin_amount"] / merged["amount_krw"].where(merged["amount_krw"] != 0) * 100
    merged["margin_pct"] = margin_raw.where(margin_raw.notna(), 0.0)
    return merged


# ═══════════════════════════════════════════════════════════════
# Page 6: 발주 커버리지
# ═══════════════════════════════════════════════════════════════
def pg_po_coverage(market, sectors, customers, year, month):
    st.title("발주 커버리지")
    so = filt(load_so(), market, sectors, customers, year=year, month=month)
    # 출고 완료 건 제외 — 이미 끝난 건은 커버리지 분석 불필요
    if not so.empty:
        so = so[so["status"] != "출고 완료"]
    po = load_po_detail()
    po_all = load_po_status()  # Cancelled 포함 전체 PO Status
    if so.empty:
        st.info("데이터 없음")
        return

    merged = calc_coverage(so, po, po_all_status=po_all)
    if merged.empty:
        st.info("데이터 없음")
        return

    # ── KPI 카드 ──
    unreg = merged[merged["coverage_status"] == "PO 미등록"]
    unordered = merged[merged["coverage_status"] == "미발주"]
    partial = merged[merged["coverage_status"] == "부분 발주"]
    in_progress = merged[merged["coverage_status"] == "발주 진행중"]
    confirmed = merged[merged["coverage_status"] == "발주 확정"]
    total_lines = len(merged)
    # PO미등록 + 미발주 + 부분발주 = 발주 필요 건
    need_order = pd.concat([unreg, unordered, partial])
    need_order_amt = need_order["amount_krw"].sum()

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    with c1:
        with st.container(border=True):
            st.markdown("**:red[PO 미등록]**")
            st.markdown(f"### {len(unreg)}건")
            st.caption("SO만 있고 PO 없음")
    with c2:
        with st.container(border=True):
            st.markdown("**미발주**")
            st.markdown(f"### {len(unordered)}건")
            st.caption("PO Open (발주 전)")
    with c3:
        with st.container(border=True):
            st.markdown("**부분 발주**")
            st.markdown(f"### {len(partial)}건")
            st.caption("일부 Open + 일부 발주")
    with c4:
        with st.container(border=True):
            st.markdown("**발주 진행중**")
            st.markdown(f"### {len(in_progress)}건")
            st.caption("Sent — 공장 확인 대기")
    with c5:
        with st.container(border=True):
            st.markdown("**발주 확정**")
            st.markdown(f"### {len(confirmed)}건")
            st.caption("Confirmed/Invoiced")
    with c6:
        with st.container(border=True):
            st.markdown("**발주 필요 금액**")
            st.markdown(f"### {fmt_krw(need_order_amt)}")
            st.caption(f"미등록+미발주+부분 {len(need_order)}건")

    # ── 커버리지 요약 Stacked Bar ──
    st.subheader("커버리지 요약")
    status_counts = merged["coverage_status"].value_counts()
    status_order = ["PO 미등록", "미발주", "부분 발주", "발주 진행중", "발주 확정"]
    status_colors = {"PO 미등록": "#b71c1c", "미발주": C_DANGER, "부분 발주": "#e65100", "발주 진행중": "#ff9800", "발주 확정": C_ENDING}
    fig_stack = go.Figure()
    for s in status_order:
        cnt = status_counts.get(s, 0)
        pct = cnt / total_lines * 100 if total_lines else 0
        fig_stack.add_trace(go.Bar(
            x=[pct], y=["커버리지"], orientation="h", name=s,
            marker_color=status_colors[s], text=f"{s} {cnt}건 ({pct:.0f}%)",
            textposition="inside",
            hovertemplate=f"<b>{s}</b><br>{cnt}건 ({pct:.1f}%)<extra></extra>",
        ))
    fig_stack.update_layout(
        barmode="stack", height=100, margin=dict(t=10, b=10, l=80),
        xaxis=dict(title="비율 (%)", range=[0, 100]),
        showlegend=True, legend=dict(orientation="h", y=-0.5),
    )
    st.plotly_chart(fig_stack, use_container_width=True)

    # ── PO 미등록 상세 테이블 ──
    st.subheader("🔴 PO 미등록 상세")
    st.caption("SO시트에만 있고 PO시트에 아예 없음 — 데이터 입력 누락 확인 필요")
    if not unreg.empty:
        tab_d, tab_e = st.tabs(["국내", "해외"])
        for tab, mkt in [(tab_d, "국내"), (tab_e, "해외")]:
            with tab:
                sub = unreg[unreg["market"] == mkt]
                if not sub.empty:
                    tbl = sub[["SO_ID", "customer_name", "os_name", "qty", "amount_krw", "po_receipt_date", "delivery_date"]].copy()
                    tbl.columns = ["SO_ID", "고객명", "품목", "수량", "매출금액", "수주일", "납기일"]
                    tbl = tbl.sort_values("수주일")
                    tbl["수주일"] = tbl["수주일"].apply(fmt_date)
                    tbl["납기일"] = tbl["납기일"].apply(fmt_date)
                    tbl["수량"] = tbl["수량"].apply(fmt_qty)
                    tbl["매출금액"] = tbl["매출금액"].apply(fmt_num)
                    st.dataframe(tbl, use_container_width=True, hide_index=True)
                else:
                    st.success(f"{mkt} PO 미등록 건 없음")
    else:
        st.success("PO 미등록 건 없음")

    # ── 미발주 상세 테이블 ──
    st.subheader("미발주 상세")
    st.caption("PO Open 상태 (PO 생성은 됐으나 공장 발주 전)")
    if not unordered.empty:
        tab_d, tab_e = st.tabs(["국내", "해외"])
        for tab, mkt in [(tab_d, "국내"), (tab_e, "해외")]:
            with tab:
                sub = unordered[unordered["market"] == mkt]
                if not sub.empty:
                    tbl = sub[["SO_ID", "po_ids", "customer_name", "os_name", "qty", "amount_krw", "po_receipt_date", "factory_order_date", "delivery_date"]].copy()
                    tbl.columns = ["SO_ID", "PO_ID", "고객명", "품목", "수량", "매출금액", "수주일", "공장발주일", "납기일"]
                    tbl = tbl.sort_values("수주일")
                    tbl["수주일"] = tbl["수주일"].apply(fmt_date)
                    tbl["공장발주일"] = tbl["공장발주일"].apply(fmt_date)
                    tbl["납기일"] = tbl["납기일"].apply(fmt_date)
                    tbl["수량"] = tbl["수량"].apply(fmt_qty)
                    tbl["매출금액"] = tbl["매출금액"].apply(fmt_num)
                    st.dataframe(tbl, use_container_width=True, hide_index=True)
                else:
                    st.success(f"{mkt} 미발주 건 없음")
    else:
        st.success("미발주 건 없음")

    # ── 부분 발주 상세 테이블 ──
    st.subheader("부분 발주 상세")
    st.caption("일부 PO는 발주(Sent/Confirmed), 일부는 Open — Open 부분 발주 필요")
    if not partial.empty:
        tab_d, tab_e = st.tabs(["국내", "해외"])
        for tab, mkt in [(tab_d, "국내"), (tab_e, "해외")]:
            with tab:
                sub = partial[partial["market"] == mkt]
                if not sub.empty:
                    tbl = sub[["SO_ID", "po_ids", "customer_name", "os_name", "qty", "amount_krw", "po_receipt_date", "factory_order_date", "delivery_date", "po_statuses"]].copy()
                    tbl.columns = ["SO_ID", "PO_ID", "고객명", "품목", "수량", "매출금액", "수주일", "공장발주일", "납기일", "PO Status"]
                    tbl = tbl.sort_values("수주일")
                    tbl["수주일"] = tbl["수주일"].apply(fmt_date)
                    tbl["공장발주일"] = tbl["공장발주일"].apply(fmt_date)
                    tbl["납기일"] = tbl["납기일"].apply(fmt_date)
                    tbl["수량"] = tbl["수량"].apply(fmt_qty)
                    tbl["매출금액"] = tbl["매출금액"].apply(fmt_num)
                    st.dataframe(tbl, use_container_width=True, hide_index=True)
                else:
                    st.success(f"{mkt} 부분 발주 건 없음")
    else:
        st.success("부분 발주 건 없음")

    # ── 발주 진행중 상세 테이블 ──
    st.subheader("발주 진행중 상세")
    st.caption("공장에 발주(Sent)했으나 Confirmed 전 단계")
    if not in_progress.empty:
        tab_d, tab_e = st.tabs(["국내", "해외"])
        for tab, mkt in [(tab_d, "국내"), (tab_e, "해외")]:
            with tab:
                sub = in_progress[in_progress["market"] == mkt]
                if not sub.empty:
                    tbl = sub[["SO_ID", "po_ids", "customer_name", "os_name", "qty", "amount_krw", "po_receipt_date", "factory_order_date", "delivery_date", "po_statuses"]].copy()
                    tbl.columns = ["SO_ID", "PO_ID", "고객명", "품목", "수량", "매출금액", "수주일", "공장발주일", "납기일", "PO Status"]
                    tbl = tbl.sort_values("수주일")
                    tbl["수주일"] = tbl["수주일"].apply(fmt_date)
                    tbl["공장발주일"] = tbl["공장발주일"].apply(fmt_date)
                    tbl["납기일"] = tbl["납기일"].apply(fmt_date)
                    tbl["수량"] = tbl["수량"].apply(fmt_qty)
                    tbl["매출금액"] = tbl["매출금액"].apply(fmt_num)
                    st.dataframe(tbl, use_container_width=True, hide_index=True)
                else:
                    st.success(f"{mkt} 발주 진행중 건 없음")
    else:
        st.success("발주 진행중 건 없음")

    # ── 고객별 미발주금액 Top 10 | 섹터별 커버리지율 ──
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("고객별 발주필요 금액 Top 10")
        if not need_order.empty:
            cust_unord = need_order.groupby("customer_name")["amount_krw"].sum().nlargest(10).reset_index()
            cust_unord.columns = ["고객", "발주필요금액"]
            fig_cu = px.bar(cust_unord, y="고객", x="발주필요금액", orientation="h",
                            color_discrete_sequence=[C_DANGER])
            fig_cu.update_traces(hovertemplate="<b>%{y}</b><br>₩%{x:,.0f}<extra></extra>")
            fig_cu.update_layout(height=350, margin=dict(t=30, l=150),
                                 yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig_cu, use_container_width=True)
        else:
            st.info("미발주 건 없음")

    with col2:
        st.subheader("섹터별 커버리지율")
        sec_cov = merged.groupby("sector").agg(
            총SO수량=("qty", "sum"), 총PO수량=("po_qty", "sum"),
        ).reset_index()
        sec_qty_safe = sec_cov["총SO수량"].replace(0, pd.NA)
        sec_cov["커버리지%"] = (sec_cov["총PO수량"] / sec_qty_safe * 100).fillna(100).round(1)
        sec_cov = sec_cov.sort_values("커버리지%", ascending=False)
        fig_sc = px.bar(sec_cov, x="sector", y="커버리지%",
                        labels={"sector": "섹터"},
                        color_discrete_sequence=[C_INPUT])
        fig_sc.update_traces(hovertemplate="<b>%{x}</b><br>커버리지: %{y:.1f}%<extra></extra>")
        fig_sc.add_hline(y=100, line_dash="dash", line_color="gray",
                         annotation_text="100%")
        fig_sc.update_layout(height=350, margin=dict(t=30, b=30))
        st.plotly_chart(fig_sc, use_container_width=True)

    # ── PO Status 파이프라인 ──
    st.subheader("PO Status 파이프라인")
    po_with_status = merged[merged["po_statuses"] != ""]
    if not po_with_status.empty:
        def _classify_po_status(s):
            if not s:
                return "No PO"
            for token in s.split(","):
                token = token.strip()
                if token.startswith("Invoiced"):
                    return "Invoiced"
            if "Confirmed" in s:
                return "Confirmed"
            if "Sent" in s:
                return "Sent"
            if "Open" in s:
                return "Open"
            return s.split(",")[0].strip() or "Other"

        po_with_status = po_with_status.copy()
        po_with_status["po_stage"] = po_with_status["po_statuses"].apply(_classify_po_status)
        stage_counts = po_with_status["po_stage"].value_counts().reset_index()
        stage_counts.columns = ["Stage", "건수"]
        stage_order = ["Open", "Sent", "Confirmed", "Invoiced"]
        stage_counts["sort_key"] = stage_counts["Stage"].apply(
            lambda x: stage_order.index(x) if x in stage_order else len(stage_order)
        )
        stage_counts = stage_counts.sort_values("sort_key")
        fig_pipe = px.bar(stage_counts, x="Stage", y="건수",
                          color_discrete_sequence=[C_PURPLE])
        fig_pipe.update_traces(hovertemplate="<b>%{x}</b><br>%{y}건<extra></extra>")
        fig_pipe.update_layout(height=300, margin=dict(t=30, b=30))
        st.plotly_chart(fig_pipe, use_container_width=True)
    else:
        st.info("PO 데이터 없음")


# ═══════════════════════════════════════════════════════════════
# Page 7: 수익성 분석
# ═══════════════════════════════════════════════════════════════
def pg_margin(market, sectors, customers, year, month):
    st.title("수익성 분석")
    so = filt(load_so(), market, sectors, customers, year=year, month=month)
    po = load_po_detail()
    if so.empty:
        st.info("데이터 없음")
        return

    merged = calc_margin(so, po)

    # DN 출고금액 조인 (미출고금액 계산용) — SO_ID 단위
    dn = load_dn()
    if not dn.empty:
        dn_agg = dn.groupby("SO_ID").agg(
            dn_amount=("amount_krw", "sum"),
        ).reset_index()
        merged = merged.merge(dn_agg, on="SO_ID", how="left")
    merged["dn_amount"] = merged["dn_amount"].fillna(0) if "dn_amount" in merged.columns else 0
    merged["unbilled_amount"] = merged["amount_krw"] - merged["dn_amount"]

    # 원가 확정 건만 마진 계산
    costed = merged[merged["has_cost"]]
    total_sales = costed["amount_krw"].sum()
    total_ico = costed["po_total_ico"].sum()
    total_margin = total_sales - total_ico
    margin_rate = total_margin / total_sales * 100 if total_sales else 0

    # ── KPI 카드 ──
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        with st.container(border=True):
            st.markdown("**총매출**")
            st.markdown(f"### {fmt_krw(total_sales)}")
    with c2:
        with st.container(border=True):
            st.markdown("**총원가(ICO)**")
            st.markdown(f"### {fmt_krw(total_ico)}")
    with c3:
        with st.container(border=True):
            st.markdown("**총마진**")
            st.markdown(f"### {fmt_krw(total_margin)}")
    with c4:
        with st.container(border=True):
            st.markdown("**마진율**")
            st.markdown(f"### {margin_rate:.1f}%")
            if total_sales:
                st.progress(max(0.0, min(margin_rate / 100, 1.0)))

    # 원가 미확정 건 안내
    no_cost = merged[~merged["has_cost"]]
    if not no_cost.empty:
        st.caption(f"원가 미확정 {len(no_cost)}건 (매출 {fmt_krw(no_cost['amount_krw'].sum())})은 마진 계산에서 제외")

    # ── 월별 마진 추이 ──
    st.subheader("월별 마진 추이")
    if not costed.empty:
        monthly = costed.groupby("period").agg(
            매출=("amount_krw", "sum"),
            원가=("po_total_ico", "sum"),
        ).reset_index().sort_values("period")
        monthly["마진"] = monthly["매출"] - monthly["원가"]
        monthly_safe = monthly["매출"].replace(0, pd.NA)
        monthly["마진율"] = (monthly["마진"] / monthly_safe * 100).fillna(0)

        fig_mm = go.Figure()
        fig_mm.add_trace(go.Bar(x=monthly["period"], y=monthly["매출"],
                                name="매출", marker_color=C_INPUT,
                                hovertemplate="<b>%{x}</b><br>매출: ₩%{y:,.0f}<extra></extra>"))
        fig_mm.add_trace(go.Bar(x=monthly["period"], y=monthly["원가"],
                                name="원가(ICO)", marker_color=C_OUTPUT,
                                hovertemplate="<b>%{x}</b><br>원가: ₩%{y:,.0f}<extra></extra>"))
        fig_mm.add_trace(go.Scatter(
            x=monthly["period"], y=monthly["마진율"], name="마진율(%)",
            mode="lines+markers", yaxis="y2",
            line=dict(color=C_ENDING, width=2),
            hovertemplate="<b>%{x}</b><br>마진율: %{y:.1f}%<extra></extra>",
        ))
        fig_mm.update_layout(
            barmode="group", height=400, margin=dict(t=30, b=30),
            xaxis=dict(type="category"),
            yaxis2=dict(title="마진율(%)", overlaying="y", side="right"),
        )
        st.plotly_chart(fig_mm, use_container_width=True)
    else:
        st.info("원가 확정 건 없음")

    # ── 탭: 고객별 | 섹터별 | 모델별 ──
    tab_cust, tab_sec, tab_model = st.tabs(["고객별", "섹터별", "모델별"])

    def _margin_analysis(df, group_col, label, tab):
        """그룹별 마진 분석 — Top 15 bar + 상세 테이블"""
        with tab:
            if df.empty:
                st.info("데이터 없음")
                return
            grp = df.groupby(group_col).agg(
                매출=("amount_krw", "sum"),
                원가=("po_total_ico", "sum"),
                미출고금액=("unbilled_amount", "sum"),
            ).reset_index()
            grp["마진"] = grp["매출"] - grp["원가"]
            sales_safe = grp["매출"].replace(0, pd.NA)
            grp["마진율"] = (grp["마진"] / sales_safe * 100).fillna(0).round(1)
            grp = grp.sort_values("매출", ascending=False)

            # Top 15 마진율 bar
            top15 = grp.head(15)
            fig = px.bar(top15, y=group_col, x="마진율", orientation="h",
                         labels={group_col: label, "마진율": "마진율(%)"},
                         color="마진율",
                         color_continuous_scale=["#d62728", "#ff9800", "#2ca02c"],
                         range_color=[0, 50])
            fig.update_traces(hovertemplate=f"<b>%{{y}}</b><br>마진율: %{{x:.1f}}%<extra></extra>")
            fig.update_layout(height=max(350, len(top15) * 28),
                              margin=dict(t=30, l=200),
                              yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig, use_container_width=True)

            # 상세 테이블
            tbl = grp.copy()
            tbl["매출"] = tbl["매출"].apply(fmt_num)
            tbl["원가"] = tbl["원가"].apply(fmt_num)
            tbl["마진"] = tbl["마진"].apply(fmt_num)
            tbl["미출고금액"] = tbl["미출고금액"].apply(fmt_num)
            tbl.rename(columns={group_col: label}, inplace=True)
            st.dataframe(tbl, use_container_width=True, hide_index=True)

    _margin_analysis(merged[merged["has_cost"]], "customer_name", "고객", tab_cust)
    _margin_analysis(merged[merged["has_cost"]], "sector", "섹터", tab_sec)
    _margin_analysis(merged[merged["has_cost"]], "os_name", "모델", tab_model)

    # ── 저마진 경보 Top 10 ──
    st.subheader("저마진 경보 Top 10")
    if not costed.empty:
        so_margin = costed.groupby("SO_ID").agg(
            고객명=("customer_name", "first"),
            품목=("os_name", lambda x: ", ".join(x.unique())),
            매출=("amount_krw", "sum"),
            원가=("po_total_ico", "sum"),
        ).reset_index()
        so_margin["마진"] = so_margin["매출"] - so_margin["원가"]
        sales_safe = so_margin["매출"].replace(0, pd.NA)
        so_margin["마진율"] = (so_margin["마진"] / sales_safe * 100).fillna(0).round(1)
        low_margin = so_margin[so_margin["마진율"] < 20].nlargest(10, "매출")
        if not low_margin.empty:
            lm = low_margin.copy()
            lm["매출"] = lm["매출"].apply(fmt_num)
            lm["원가"] = lm["원가"].apply(fmt_num)
            lm["마진"] = lm["마진"].apply(fmt_num)
            st.dataframe(lm, use_container_width=True, hide_index=True,
                         column_config={
                             "마진율": st.column_config.ProgressColumn(
                                 "마진율(%)", format="%.1f%%", min_value=0, max_value=100),
                         })
        else:
            st.success("마진율 20% 미만 건 없음")

    # ── 미출고금액 Top 10 (고객별) ──
    st.subheader("미출고금액 Top 10 (고객별)")
    unbilled = merged.groupby("customer_name")["unbilled_amount"].sum().nlargest(10).reset_index()
    unbilled.columns = ["고객", "미출고금액"]
    if not unbilled.empty and unbilled["미출고금액"].sum() > 0:
        fig_ub = px.bar(unbilled, y="고객", x="미출고금액", orientation="h",
                        color_discrete_sequence=[C_PURPLE])
        fig_ub.update_traces(hovertemplate="<b>%{y}</b><br>미출고: ₩%{x:,.0f}<extra></extra>")
        fig_ub.update_layout(height=350, margin=dict(t=30, l=150),
                             yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_ub, use_container_width=True)
    else:
        st.info("미출고 건 없음")


# ═══════════════════════════════════════════════════════════════
# Page 8: Order Book (백로그) — 3탭 구조
# ═══════════════════════════════════════════════════════════════
def pg_orderbook(market, sectors, customers, **_):
    st.title("Order Book")
    if _.get("year", "전체") != "전체" or _.get("month", "전체") != "전체":
        st.caption("ℹ️ 이 페이지는 현재 잔고 기준입니다 — 연도/월 필터는 적용되지 않습니다.")

    backlog = filt(load_backlog(), market, sectors, customers, period_col=None)

    today_ts = pd.Timestamp.today().normalize()

    # ── KPI 카드 (상단 고정) ──
    c1, c2, c3, c4 = st.columns(4)
    if not backlog.empty:
        bl_total = backlog["ending_amount"].sum()
        overdue = backlog[backlog["delivery_date"] < today_ts]
        imminent = backlog[
            (backlog["delivery_date"] >= today_ts)
            & (backlog["delivery_date"] <= today_ts + timedelta(days=14))
        ]
        c1.metric("Backlog 금액", fmt_krw(bl_total))
        c2.metric("Backlog 건수", f"{len(backlog)}건")
        c3.metric("납기 지연", f"{len(overdue)}건 / {fmt_krw(overdue['ending_amount'].sum())}",
                  help="납기일 < 오늘")
        c4.metric("납기 임박 (2주)", f"{len(imminent)}건",
                  help="오늘 ~ 2주 이내 납기")
    else:
        c1.metric("Backlog 금액", "₩0")
        c2.metric("Backlog 건수", "0건")
        c3.metric("납기 지연", "0건")
        c4.metric("납기 임박", "0건")

    # ── OB 데이터 로드 (탭 공용) ──
    ob = load_order_book()
    ob_monthly = pd.DataFrame()
    if not ob.empty:
        ob = ob.rename(columns={"구분": "market", "Sector": "sector", "Customer name": "customer_name"})
        ob = filt(ob, market, sectors, customers, period_col=None)
        _line_key = ["SO_ID", "OS name", "Expected delivery date"]
        ob_flow = ob.groupby("Period").agg(
            Input=("Value_Input_amount", "sum"),
            Output=("Value_Output_amount", "sum"),
        ).reset_index()
        _line_end = ob.groupby(_line_key + ["Period"])["Value_Ending_amount"].sum().reset_index()
        _pivot = _line_end.pivot_table(
            index=_line_key, columns="Period", values="Value_Ending_amount",
        ).sort_index(axis=1).ffill(axis=1).fillna(0)
        _monthly_ending = _pivot.sum(axis=0).reset_index()
        _monthly_ending.columns = ["Period", "Ending"]
        ob_monthly = ob_flow.merge(_monthly_ending, on="Period", how="outer").fillna(0).sort_values("Period")

    # ═══ 3탭 구조 ═══
    tab_exec, tab_risk, tab_conv = st.tabs(["Executive", "Risk", "Conversion"])

    # ──────────────────────────────────────────────
    # Executive 탭
    # ──────────────────────────────────────────────
    with tab_exec:
        # 워터폴 차트 — Opening → Input → Output → Ending
        if not ob_monthly.empty:
            st.subheader("Order Book 워터폴")

            # 뷰 모드 토글 + 월 선택
            wf_col1, wf_col2 = st.columns([1, 2])
            with wf_col1:
                wf_mode = st.radio("보기", ["월별", "누적"], horizontal=True, key="wf_mode")
            with wf_col2:
                if wf_mode == "월별":
                    periods = ob_monthly["Period"].tolist()
                    wf_sel_period = st.selectbox(
                        "기준 월", periods, index=len(periods) - 1, key="wf_period",
                    )
                else:
                    wf_sel_period = None

            if wf_mode == "월별":
                # 선택 월 워터폴
                sel_row = ob_monthly[ob_monthly["Period"] == wf_sel_period].iloc[0]
                opening = sel_row["Ending"] - sel_row["Input"] + sel_row["Output"]

                wf_labels = ["Opening", "수주(Input)", "출고(Output)", "Ending"]
                wf_values = [opening, sel_row["Input"], -sel_row["Output"], sel_row["Ending"]]
                wf_measures = ["absolute", "relative", "relative", "total"]

                fig_wf = go.Figure(go.Waterfall(
                    x=wf_labels, y=wf_values, measure=wf_measures,
                    connector=dict(line=dict(color="gray", dash="dot")),
                    increasing=dict(marker=dict(color=C_INPUT)),
                    decreasing=dict(marker=dict(color=C_OUTPUT)),
                    totals=dict(marker=dict(color=C_ENDING)),
                    textposition="outside",
                    text=[fmt_krw(abs(v)) for v in wf_values],
                    hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>",
                ))
                fig_wf.update_layout(
                    height=400, margin=dict(t=30, b=30),
                    yaxis=dict(title="금액 (KRW)"),
                    title=dict(text=f"기준 월: {wf_sel_period}", font=dict(size=14)),
                )
                st.plotly_chart(fig_wf, use_container_width=True)

            else:  # 누적
                total_input = ob_monthly["Input"].sum()
                total_output = ob_monthly["Output"].sum()
                first_opening = ob_monthly.iloc[0]["Ending"] - ob_monthly.iloc[0]["Input"] + ob_monthly.iloc[0]["Output"]
                final_ending = ob_monthly.iloc[-1]["Ending"]
                period_range = f"{ob_monthly['Period'].min()} ~ {ob_monthly['Period'].max()}"

                wf_labels = ["Opening", "총 수주(Input)", "총 출고(Output)", "Ending"]
                wf_values = [first_opening, total_input, -total_output, final_ending]
                wf_measures = ["absolute", "relative", "relative", "total"]

                fig_wf = go.Figure(go.Waterfall(
                    x=wf_labels, y=wf_values, measure=wf_measures,
                    connector=dict(line=dict(color="gray", dash="dot")),
                    increasing=dict(marker=dict(color=C_INPUT)),
                    decreasing=dict(marker=dict(color=C_OUTPUT)),
                    totals=dict(marker=dict(color=C_ENDING)),
                    textposition="outside",
                    text=[fmt_krw(abs(v)) for v in wf_values],
                    hovertemplate="<b>%{x}</b><br>₩%{y:,.0f}<extra></extra>",
                ))
                fig_wf.update_layout(
                    height=400, margin=dict(t=30, b=30),
                    yaxis=dict(title="금액 (KRW)"),
                    title=dict(text=f"누적 기간: {period_range}", font=dict(size=14)),
                )
                st.plotly_chart(fig_wf, use_container_width=True)

            latest_period = ob_monthly["Period"].max()
            latest = ob_monthly[ob_monthly["Period"] == latest_period].iloc[0]

            # 3대 KPI
            kc1, kc2, kc3 = st.columns(3)
            # Backlog Cover (months)
            recent_3 = ob_monthly.tail(3)
            avg_output_3m = recent_3["Output"].mean()
            backlog_cover = latest["Ending"] / avg_output_3m if avg_output_3m else float("inf")
            kc1.metric("Backlog Cover",
                       f"{backlog_cover:.1f}개월" if backlog_cover != float("inf") else "N/A",
                       help="Ending / 최근 3개월 평균 Output")
            # Past Due Ratio
            if not backlog.empty:
                overdue_amt = backlog[backlog["delivery_date"] < today_ts]["ending_amount"].sum()
                past_due_pct = overdue_amt / bl_total * 100 if bl_total else 0
                kc2.metric("Past Due Ratio", f"{past_due_pct:.1f}%",
                           help="지연 Backlog / 총 Backlog")
            else:
                kc2.metric("Past Due Ratio", "0%")
            # Book-to-Bill
            btb_val = latest["Input"] / latest["Output"] if latest["Output"] else float("inf")
            kc3.metric("Book-to-Bill",
                       f"{btb_val:.2f}" if btb_val != float("inf") else "N/A",
                       help="Input / Output (>1: 수주 우세)")

            # 월별 Input/Output/Ending 추이
            st.subheader("월별 수주잔고 추이")
            st.caption("현재 DB 상태 기준 계산 (마감 확정치 아님)")
            fig_ob = go.Figure()
            fig_ob.add_trace(go.Bar(x=ob_monthly["Period"], y=ob_monthly["Input"],
                                    name="수주(Input)", marker_color=C_INPUT,
                                    hovertemplate="<b>%{x}</b><br>수주: ₩%{y:,.0f}<extra></extra>"))
            fig_ob.add_trace(go.Bar(x=ob_monthly["Period"], y=ob_monthly["Output"],
                                    name="출고(Output)", marker_color=C_OUTPUT,
                                    hovertemplate="<b>%{x}</b><br>출고: ₩%{y:,.0f}<extra></extra>"))
            fig_ob.add_trace(go.Scatter(
                x=ob_monthly["Period"], y=ob_monthly["Ending"], name="잔고(Ending)",
                mode="lines+markers", line=dict(color=C_ENDING, width=2),
                hovertemplate="<b>%{x}</b><br>잔고: ₩%{y:,.0f}<extra></extra>",
            ))
            fig_ob.update_layout(
                barmode="group", height=400, margin=dict(t=30, b=30),
                xaxis=dict(type="category", title="Period", rangeslider=dict(visible=True)),
                yaxis=dict(title="금액 (KRW)"),
            )
            st.plotly_chart(fig_ob, use_container_width=True)

            # 섹터별 / 고객별 Backlog
            if not backlog.empty:
                col3, col4 = st.columns(2)
                with col3:
                    st.subheader("섹터별 Backlog")
                    bl_sec = backlog.groupby("sector")["ending_amount"].sum().sort_values(ascending=False).reset_index()
                    bl_sec.columns = ["섹터", "금액"]
                    fig_s = px.bar(bl_sec, x="섹터", y="금액", color_discrete_sequence=[C_INPUT])
                    fig_s.update_traces(hovertemplate="<b>%{x}</b><br>Backlog: ₩%{y:,.0f}<extra></extra>")
                    fig_s.update_layout(height=350, margin=dict(t=30, b=30))
                    st.plotly_chart(fig_s, use_container_width=True)
                with col4:
                    st.subheader("고객별 Backlog Top 10")
                    bl_cust = backlog.groupby("customer_name")["ending_amount"].sum().nlargest(10).reset_index()
                    bl_cust.columns = ["고객", "금액"]
                    fig_c = px.bar(bl_cust, y="고객", x="금액", orientation="h",
                                   color_discrete_sequence=[C_OUTPUT])
                    fig_c.update_traces(hovertemplate="<b>%{y}</b><br>Backlog: ₩%{x:,.0f}<extra></extra>")
                    fig_c.update_layout(height=350, margin=dict(t=30, l=150),
                                        yaxis=dict(autorange="reversed"))
                    st.plotly_chart(fig_c, use_container_width=True)
        else:
            st.info("Order Book 데이터 없음")

    # ──────────────────────────────────────────────
    # Risk 탭
    # ──────────────────────────────────────────────
    with tab_risk:
        # Aging 분석
        if not backlog.empty:
            st.subheader("Backlog Aging 분석")
            bl = backlog.copy()
            bl["days"] = (bl["delivery_date"] - today_ts).dt.days

            def _aging(d):
                if pd.isna(d):
                    return "날짜없음"
                if d < -30:
                    return "① 30일+ 지연"
                if d < 0:
                    return "② 0~30일 지연"
                if d <= 14:
                    return "③ 2주 이내"
                if d <= 30:
                    return "④ 2주~1개월"
                if d <= 90:
                    return "⑤ 1~3개월"
                return "⑥ 3개월+"

            bl["aging"] = bl["days"].apply(_aging)
            aging_agg = bl.groupby("aging").agg(
                건수=("SO_ID", "count"),
                금액=("ending_amount", "sum"),
            ).reset_index().sort_values("aging")

            col1, col2 = st.columns(2)
            with col1:
                fig_a = px.bar(aging_agg, x="aging", y="금액",
                               color="aging", labels={"aging": "구간", "금액": "Backlog 금액"},
                               color_discrete_sequence=px.colors.sequential.RdBu_r)
                fig_a.update_traces(hovertemplate="<b>%{x}</b><br>Backlog: ₩%{y:,.0f}<extra></extra>")
                fig_a.update_layout(height=350, margin=dict(t=30, b=30), showlegend=False)
                event_aging = st.plotly_chart(fig_a, use_container_width=True,
                                              on_select="rerun", key="aging_bar")
            with col2:
                fig_a2 = px.pie(aging_agg, names="aging", values="금액", hole=0.4)
                fig_a2.update_traces(hovertemplate="<b>%{label}</b><br>₩%{value:,.0f} (%{percent})<extra></extra>")
                fig_a2.update_layout(height=350, margin=dict(t=30, b=30))
                st.plotly_chart(fig_a2, use_container_width=True)

            # Aging 드릴다운
            selected_aging = None
            if event_aging and event_aging.selection and event_aging.selection.points:
                selected_aging = event_aging.selection.points[0]["x"]
            if selected_aging:
                st.subheader(f"📌 {selected_aging} 상세")
                aging_detail = bl[bl["aging"] == selected_aging]
                ad = aging_detail[["market", "SO_ID", "customer_name", "os_name",
                                   "delivery_date", "ending_qty", "ending_amount"]].copy()
                ad.columns = ["구분", "SO_ID", "고객명", "품목", "납기일", "잔여수량", "잔여금액"]
                ad["납기일"] = ad["납기일"].apply(fmt_date)
                ad["잔여수량"] = ad["잔여수량"].apply(lambda x: f"{int(x):,}")
                ad["잔여금액"] = ad["잔여금액"].apply(fmt_num)
                st.dataframe(ad, use_container_width=True, hide_index=True)

            # 고금액 위험건 Top 10
            st.subheader("고금액 위험건 Top 10")
            overdue_bl = bl[bl["delivery_date"] < today_ts].nlargest(10, "ending_amount")
            if not overdue_bl.empty:
                risk = overdue_bl[["market", "SO_ID", "customer_name", "os_name",
                                   "delivery_date", "ending_qty", "ending_amount", "days"]].copy()
                risk.columns = ["구분", "SO_ID", "고객명", "품목", "납기일", "잔여수량", "잔여금액", "지연일"]
                risk["지연일"] = risk["지연일"].apply(lambda d: f"{abs(int(d))}일")
                risk["납기일"] = risk["납기일"].apply(fmt_date)
                risk["잔여수량"] = risk["잔여수량"].apply(lambda x: f"{int(x):,}")
                risk["잔여금액"] = risk["잔여금액"].apply(fmt_num)
                st.dataframe(risk, use_container_width=True, hide_index=True)
            else:
                st.success("납기 지연 건 없음")

            # 납기 분포 히트맵
            st.subheader("납기 분포 히트맵")
            st.caption(f"금월 ~ {today_ts.year}년 말 — 월별 x 섹터 납기 예정 금액")
            bl_heat = backlog.copy()
            bl_heat["delivery_date"] = pd.to_datetime(bl_heat["delivery_date"], errors="coerce")
            month_start = today_ts.replace(day=1)
            month_end = pd.Timestamp(today_ts.year, 12, 31)
            bl_heat = bl_heat[
                (bl_heat["delivery_date"] >= month_start)
                & (bl_heat["delivery_date"] <= month_end)
            ]
            if not bl_heat.empty:
                bl_heat["month"] = bl_heat["delivery_date"].dt.to_period("M").astype(str)
                bl_heat["sector"] = bl_heat["sector"].fillna("미분류")
                pivot = bl_heat.pivot_table(
                    index="sector", columns="month",
                    values="ending_amount", aggfunc="sum", fill_value=0,
                )
                hover_text = pivot.map(lambda v: fmt_krw(v))
                fig_hm = go.Figure(data=go.Heatmap(
                    z=pivot.values,
                    x=pivot.columns.tolist(),
                    y=pivot.index.tolist(),
                    text=hover_text.values,
                    texttemplate="%{text}",
                    hovertemplate="<b>%{y}</b> · %{x}<br>₩%{z:,.0f}<extra></extra>",
                    colorscale="YlOrRd",
                ))
                fig_hm.update_layout(
                    height=max(250, 50 + len(pivot) * 40),
                    margin=dict(t=20, b=30),
                    xaxis=dict(type="category", title="납기월"),
                    yaxis=dict(title="섹터", autorange="reversed"),
                )
                st.plotly_chart(fig_hm, use_container_width=True)
            else:
                st.info("연말까지 납기 예정 건 없음")
        else:
            st.info("Backlog 데이터 없음")

    # ──────────────────────────────────────────────
    # Conversion 탭
    # ──────────────────────────────────────────────
    with tab_conv:
        so = load_so()
        so_f = filt(so, market, sectors, customers)
        dn_raw = load_dn()
        dn_ej = enrich_dn(dn_raw, so)
        dn_f = filt(dn_ej, market, sectors, customers, period_col=None)
        po = load_po_detail()

        # 전환 퍼널 — 필터된 SO_ID 범위로 제한
        st.subheader("전환 퍼널")
        so_ids = set(so_f["SO_ID"].unique()) if not so_f.empty else set()
        po_so_ids = set(po["SO_ID"].unique()) & so_ids if not po.empty else set()
        dn_so_ids = set(dn_f["SO_ID"].unique()) & so_ids if not dn_f.empty else set()

        so_cnt = len(so_ids)
        po_cnt = len(po_so_ids)
        dn_cnt = len(dn_so_ids)

        if so_cnt > 0:
            fig_funnel = go.Figure(go.Funnel(
                y=["SO 수주", "PO 발행", "DN 출고"],
                x=[so_cnt, po_cnt, dn_cnt],
                textinfo="value+percent initial",
                marker=dict(color=[C_INPUT, C_PURPLE, C_OUTPUT]),
                hovertemplate="<b>%{y}</b><br>%{x}건<extra></extra>",
            ))
            fig_funnel.update_layout(height=300, margin=dict(t=30, b=30))
            st.plotly_chart(fig_funnel, use_container_width=True)

            # 전환율 메트릭
            mc1, mc2 = st.columns(2)
            so_po_pct = po_cnt / so_cnt * 100 if so_cnt else 0
            po_dn_pct = dn_cnt / po_cnt * 100 if po_cnt else 0
            mc1.metric("SO → PO 전환율", f"{so_po_pct:.1f}%")
            mc2.metric("PO → DN 전환율", f"{po_dn_pct:.1f}%")
        else:
            st.info("SO 데이터 없음")

        # 리드타임 분석 — 수주→출고
        st.subheader("리드타임 분석 (수주 → 출고)")
        if not so_f.empty and not dn_f.empty:
            # SO Period(1일 기준) → DN dispatch_date (필터 범위 내)
            so_period = so_f[["SO_ID", "period"]].dropna(subset=["period"])
            so_period = so_period.sort_values("period").drop_duplicates(subset=["SO_ID"], keep="first")
            so_period["so_date"] = pd.to_datetime(so_period["period"] + "-01", errors="coerce")
            dn_first = dn_f.groupby("SO_ID")["dispatch_date"].min().reset_index()
            dn_first.columns = ["SO_ID", "first_dispatch"]
            lt = so_period.merge(dn_first, on="SO_ID", how="inner")
            lt["lead_days"] = (lt["first_dispatch"] - lt["so_date"]).dt.days
            lt = lt[lt["lead_days"] >= 0]  # 음수 제거

            if not lt.empty:
                # KPI 카드 — 핵심 수치 먼저 표시
                lt_mean = lt["lead_days"].mean()
                lt_median = lt["lead_days"].median()
                lt_min = lt["lead_days"].min()
                lt_max = lt["lead_days"].max()
                lk1, lk2, lk3, lk4 = st.columns(4)
                lk1.metric("평균", f"{lt_mean:.0f}일")
                lk2.metric("중앙값", f"{lt_median:.0f}일")
                lk3.metric("최단", f"{lt_min:.0f}일")
                lk4.metric("최장", f"{lt_max:.0f}일")

                col_lt1, col_lt2 = st.columns(2)
                with col_lt1:
                    fig_box = px.box(lt, y="lead_days",
                                     labels={"lead_days": "리드타임 (일)"},
                                     color_discrete_sequence=[C_INPUT])
                    fig_box.update_layout(height=350, margin=dict(t=30, b=30),
                                          title="수주→출고 리드타임 분포")
                    st.plotly_chart(fig_box, use_container_width=True)

                with col_lt2:
                    # 월별 평균 리드타임
                    lt["so_month"] = lt["period"]
                    lt_monthly = lt.groupby("so_month")["lead_days"].mean().reset_index()
                    lt_monthly.columns = ["월", "평균리드타임"]
                    lt_monthly = lt_monthly.sort_values("월")
                    fig_lt = px.line(lt_monthly, x="월", y="평균리드타임",
                                    markers=True,
                                    labels={"평균리드타임": "평균 리드타임 (일)"},
                                    color_discrete_sequence=[C_PURPLE])
                    fig_lt.update_traces(hovertemplate="<b>%{x}</b><br>%{y:.0f}일<extra></extra>")
                    fig_lt.update_layout(height=350, margin=dict(t=30, b=30),
                                         xaxis=dict(type="category"),
                                         title="월별 평균 리드타임 추이")
                    st.plotly_chart(fig_lt, use_container_width=True)

                q1 = lt["lead_days"].quantile(0.25)
                q3 = lt["lead_days"].quantile(0.75)
                iqr = q3 - q1
                outlier_threshold = q3 + 1.5 * iqr
                n_outliers = int((lt["lead_days"] > outlier_threshold).sum())
                spread = "안정적" if iqr <= 30 else ("다소 편차 있음" if iqr <= 60 else "편차가 큼")
                skew = "빠른 쪽에 집중" if lt_mean > lt_median else ("느린 쪽에 집중" if lt_mean < lt_median else "고르게 분포")

                lines = [
                    f"**수주→출고 리드타임** {len(lt)}건 분석 결과:\n",
                    f"- 절반의 건이 **{lt_median:.0f}일 이내**에 출고됩니다 (중앙값)",
                    f"- 대부분(75%)의 건이 **{q1:.0f}~{q3:.0f}일** 사이에 완료됩니다 (상자 범위)",
                    f"- 가장 빨리 출고된 건은 **{lt_min:.0f}일**, 가장 오래 걸린 건은 **{lt_max:.0f}일**",
                    f"- 분포 특성: **{spread}** (상자 폭 {iqr:.0f}일) · {skew}",
                ]
                if n_outliers > 0:
                    lines.append(f"- **이상치 {n_outliers}건** — {outlier_threshold:.0f}일 초과, 개별 원인 확인 필요")
                else:
                    lines.append("- 이상치 없음 — 리드타임이 일정하게 관리되고 있습니다")
                st.markdown("\n".join(lines))

                # 이상치 상세 (expander)
                if n_outliers > 0:
                    lt_outliers = lt[lt["lead_days"] > outlier_threshold].copy()
                    # 고객명 조인
                    so_cust = so_f[["SO_ID", "customer_name"]].drop_duplicates(subset=["SO_ID"])
                    lt_outliers = lt_outliers.merge(so_cust, on="SO_ID", how="left")
                    lt_outliers = lt_outliers.sort_values("lead_days", ascending=False)
                    with st.expander(f"이상치 {n_outliers}건 상세"):
                        otbl = lt_outliers[["SO_ID", "customer_name", "period", "first_dispatch", "lead_days"]].copy()
                        otbl.columns = ["SO_ID", "고객명", "수주월", "최초출고일", "리드타임(일)"]
                        otbl["최초출고일"] = otbl["최초출고일"].apply(fmt_date)
                        otbl["리드타임(일)"] = otbl["리드타임(일)"].apply(lambda x: f"{int(x)}일")
                        st.dataframe(otbl, use_container_width=True, hide_index=True)
            else:
                st.info("리드타임 분석 가능한 데이터 없음")

        # 해외 물류 리드타임 (출고→픽업→선적) — 국내 필터 시 비표시
        if market != "국내":
            ship_df = load_dn_export_shipping()
            if not ship_df.empty:
                # SO 메타 조인으로 sector 필터 적용 (customer_name은 ship_df에 이미 존재)
                so_meta = so[["SO_ID", "sector"]].drop_duplicates(subset=["SO_ID"])
                ship_df = ship_df.merge(so_meta, on="SO_ID", how="left")
                ship_df["sector"] = ship_df["sector"].fillna("")
                if sectors:
                    ship_df = ship_df[ship_df["sector"].isin(sectors)]
                if customers:
                    ship_df = ship_df[ship_df["customer_name"].isin(customers)]

            if not ship_df.empty:
                st.subheader("해외 물류 리드타임 (출고 → 픽업 → 선적)")
                st.caption("공장 출고 후 각 구간별 소요일 — 어디서 병목이 생기는지 비교")
                ship = ship_df.copy()
                ship["출고_픽업"] = (ship["pickup_date"] - ship["factory_date"]).dt.days
                ship["픽업_선적"] = (ship["ship_date"] - ship["pickup_date"]).dt.days

                # KPI 카드 — 구간별 평균
                seg_stats = []
                lt_data = []
                for label, col in [("출고→픽업", "출고_픽업"), ("픽업→선적", "픽업_선적")]:
                    valid = ship[ship[col].notna() & (ship[col] >= 0)]
                    if not valid.empty:
                        seg_stats.append((label, valid[col].mean(), valid[col].median(), len(valid)))
                        for _, r in valid.iterrows():
                            lt_data.append({"구간": label, "리드타임(일)": r[col]})

                if seg_stats:
                    sk_cols = st.columns(len(seg_stats) * 2)
                    for i, (label, avg, med, cnt) in enumerate(seg_stats):
                        sk_cols[i * 2].metric(f"{label} 평균", f"{avg:.0f}일")
                        sk_cols[i * 2 + 1].metric(f"{label} 중앙값", f"{med:.0f}일", help=f"{cnt}건 기준")

                if lt_data:
                    lt_df = pd.DataFrame(lt_data)
                    fig_lt_box = px.box(lt_df, x="구간", y="리드타임(일)",
                                        color="구간",
                                        color_discrete_sequence=[C_INPUT, C_OUTPUT])
                    fig_lt_box.update_layout(height=350, margin=dict(t=30, b=30),
                                              showlegend=False)
                    st.plotly_chart(fig_lt_box, use_container_width=True)
                    # 구간별 상세 통계 기반 동적 해석
                    seg_details = {}
                    for label, col in [("출고→픽업", "출고_픽업"), ("픽업→선적", "픽업_선적")]:
                        valid = ship[ship[col].notna() & (ship[col] >= 0)][col]
                        if not valid.empty:
                            sq1, sq3 = valid.quantile(0.25), valid.quantile(0.75)
                            s_iqr = sq3 - sq1
                            s_outlier_th = sq3 + 1.5 * s_iqr
                            seg_details[label] = {
                                "median": valid.median(), "q1": sq1, "q3": sq3,
                                "iqr": s_iqr, "min": valid.min(), "max": valid.max(),
                                "n_outliers": int((valid > s_outlier_th).sum()),
                                "count": len(valid),
                            }

                    lines = ["**해외 물류 구간별 분석**\n"]
                    total_ship_outliers = 0
                    for label, s in seg_details.items():
                        spread = "안정적" if s["iqr"] <= 5 else ("다소 편차 있음" if s["iqr"] <= 15 else "편차가 큼")
                        lines.append(f"**{label}** ({s['count']}건)")
                        lines.append(f"- 절반이 **{s['median']:.0f}일 이내** · 대부분 **{s['q1']:.0f}~{s['q3']:.0f}일** 사이")
                        lines.append(f"- 최단 {s['min']:.0f}일 ~ 최장 {s['max']:.0f}일 · {spread}")
                        if s["n_outliers"] > 0:
                            lines.append(f"- 이상치 {s['n_outliers']}건 — 비정상적 지연 발생")
                            total_ship_outliers += s["n_outliers"]
                        lines.append("")

                    # 구간 간 비교
                    if len(seg_details) == 2:
                        seg_labels = list(seg_details.keys())
                        s0, s1 = seg_details[seg_labels[0]], seg_details[seg_labels[1]]
                        if s0["median"] > s1["median"]:
                            bottleneck = seg_labels[0]
                        elif s1["median"] > s0["median"]:
                            bottleneck = seg_labels[1]
                        else:
                            bottleneck = None
                        if bottleneck:
                            lines.append(f"**병목 구간: {bottleneck}** — 중앙값 기준으로 이 구간이 더 오래 걸립니다.")
                        else:
                            lines.append("두 구간의 소요일이 비슷합니다.")

                    st.markdown("\n".join(lines))

                    # 이상치 상세 (expander)
                    if total_ship_outliers > 0:
                        outlier_rows = []
                        for label, col in [("출고→픽업", "출고_픽업"), ("픽업→선적", "픽업_선적")]:
                            if label not in seg_details:
                                continue
                            s = seg_details[label]
                            s_th = s["q3"] + 1.5 * s["iqr"]
                            ov = ship[ship[col].notna() & (ship[col] > s_th)].copy()
                            if not ov.empty:
                                ov["구간"] = label
                                ov["소요일"] = ov[col]
                                outlier_rows.append(ov)
                        if outlier_rows:
                            ship_otbl = pd.concat(outlier_rows, ignore_index=True)
                            ship_otbl = ship_otbl.sort_values("소요일", ascending=False)
                            with st.expander(f"이상치 {total_ship_outliers}건 상세"):
                                disp = ship_otbl[["구간", "DN_ID", "SO_ID", "customer_name",
                                                   "factory_date", "pickup_date", "ship_date", "소요일"]].copy()
                                disp.columns = ["구간", "DN_ID", "SO_ID", "고객명",
                                                "공장출고일", "픽업일", "선적일", "소요일"]
                                for dc in ("공장출고일", "픽업일", "선적일"):
                                    disp[dc] = disp[dc].apply(fmt_date)
                                disp["소요일"] = disp["소요일"].apply(lambda x: f"{int(x)}일")
                                st.dataframe(disp, use_container_width=True, hide_index=True)
                else:
                    st.info("해외 물류 리드타임 데이터 없음")


if __name__ == "__main__":
    main()
