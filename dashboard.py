"""
NOAH 대시보드
============
수주/출고 현황, 제품/섹터/고객 분석, Order Book(백로그) 등 핵심 KPI.
데이터 소스: noah_data.db (SQLite, sync_db.py로 동기화)

Usage:
    streamlit run dashboard.py
"""

import sqlite3
from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from po_generator.config import DB_FILE
from po_generator.db_schema import get_sync_metadata

# ── 상수 ──────────────────────────────────────────────────────
_TODAY = datetime.today()
_TODAY_DATE = _TODAY.date()
_THIS_MONTH = _TODAY.strftime("%Y-%m")
_WEEK_START = (_TODAY - timedelta(days=_TODAY.weekday())).date()
_WEEK_END = _WEEK_START + timedelta(days=6)

C_INPUT = "#1f77b4"
C_OUTPUT = "#ff7f0e"
C_ENDING = "#2ca02c"
C_DANGER = "#d62728"


# ═══════════════════════════════════════════════════════════════
# 포맷 유틸
# ═══════════════════════════════════════════════════════════════
def fmt_krw(v: float) -> str:
    """KRW 포맷 (억/만 자동 단위)"""
    if pd.isna(v) or v == 0:
        return "₩0"
    a, s = abs(v), "-" if v < 0 else ""
    if a >= 1e8:
        return f"{s}₩{a / 1e8:,.1f}억"
    if a >= 1e4:
        return f"{s}₩{a / 1e4:,.0f}만"
    return f"{s}₩{a:,.0f}"


def fmt_qty(v) -> str:
    return "0" if pd.isna(v) else f"{int(v):,}"


# ═══════════════════════════════════════════════════════════════
# DB 연결
# ═══════════════════════════════════════════════════════════════
def _conn():
    return sqlite3.connect(str(DB_FILE)) if DB_FILE.exists() else None


# ═══════════════════════════════════════════════════════════════
# 데이터 로더 (캐시 5분)
# ═══════════════════════════════════════════════════════════════
@st.cache_data(ttl=300)
def load_so() -> pd.DataFrame:
    """SO 국내+해외 통합"""
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
                   [Expected delivery date], '해외'
            FROM so_export
            WHERE COALESCE(Status, '') != 'Cancelled'
              AND Period IS NOT NULL AND TRIM(Period) != ''
        """, conn)
    finally:
        conn.close()
    df["delivery_date"] = pd.to_datetime(df["delivery_date"], errors="coerce")
    for c in ("os_name", "sector", "customer_name"):
        df[c] = df[c].fillna("")
    return df


@st.cache_data(ttl=300)
def load_dn() -> pd.DataFrame:
    """DN 국내+해외 통합"""
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
    finally:
        conn.close()
    df["dispatch_date"] = pd.to_datetime(df["dispatch_date"], errors="coerce")
    df["dispatch_month"] = df["dispatch_date"].dt.strftime("%Y-%m")
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
            SELECT s.SO_ID, s.customer_name, s.os_name, s.line_item,
                   s.model_code, s.sector, s.delivery_date, s.market,
                   0, 0, d.out_qty, d.out_amt
            FROM dn_combined d
            INNER JOIN so_combined s
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
    finally:
        conn.close()
    df["delivery_date"] = pd.to_datetime(df["delivery_date"], errors="coerce")
    return df


@st.cache_data(ttl=300)
def load_sync_meta() -> dict:
    conn = _conn()
    if not conn:
        return {}
    try:
        return get_sync_metadata(conn)
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
    except Exception:
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

    # ── 사이드바 ──
    st.sidebar.title("NOAH 대시보드")
    page = st.sidebar.radio("페이지", [
        "오늘의 현황", "수주/출고 현황", "제품 분석",
        "섹터 분석", "고객 분석", "Order Book",
    ])
    st.sidebar.divider()
    market = st.sidebar.radio("시장 구분", ["전체", "국내", "해외"], horizontal=True)

    so_raw = load_so()
    years = sorted(so_raw["period"].dropna().str[:4].unique()) if not so_raw.empty else []
    year = st.sidebar.selectbox("연도", ["전체"] + list(years), index=0)
    month = st.sidebar.selectbox(
        "월", ["전체"] + [f"{m:02d}" for m in range(1, 13)], index=0,
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

    # ── 라우팅 ──
    kw = dict(market=market, sectors=sectors, customers=customers, year=year, month=month)
    {
        "오늘의 현황": pg_today,
        "수주/출고 현황": pg_orders,
        "제품 분석": pg_product,
        "섹터 분석": pg_sector,
        "고객 분석": pg_customer,
        "Order Book": pg_orderbook,
    }[page](**kw)


# ═══════════════════════════════════════════════════════════════
# Page 1: 오늘의 현황
# ═══════════════════════════════════════════════════════════════
def pg_today(market, sectors, customers, **_):
    st.title("오늘의 현황")

    so = filt(load_so(), market, sectors, customers)
    dn = filt(enrich_dn(load_dn(), load_so()), market, sectors, customers, period_col=None)
    backlog = filt(load_backlog(), market, sectors, customers, period_col=None)

    # 납기 계산
    so_d = so[so["delivery_date"].notna()] if not so.empty else pd.DataFrame()
    today_due = so_d[so_d["delivery_date"].dt.date == _TODAY_DATE] if not so_d.empty else pd.DataFrame()
    week_due = so_d[
        (so_d["delivery_date"].dt.date >= _WEEK_START)
        & (so_d["delivery_date"].dt.date <= _WEEK_END)
    ] if not so_d.empty else pd.DataFrame()

    month_so = so[so["period"] == _THIS_MONTH] if not so.empty else pd.DataFrame()
    month_dn = dn[dn["dispatch_month"] == _THIS_MONTH] if not dn.empty else pd.DataFrame()

    # KPI
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("오늘 납기 건수", fmt_qty(len(today_due)))
    c2.metric("이번 주 납기 건수", fmt_qty(len(week_due)))
    c3.metric("금월 수주 금액", fmt_krw(month_so["amount_krw"].sum() if not month_so.empty else 0))
    c4.metric("금월 출고 금액", fmt_krw(month_dn["amount_krw"].sum() if not month_dn.empty else 0))

    # 이번 주 납기 예정
    st.subheader("이번 주 납기 예정")
    if not week_due.empty:
        t = week_due[["SO_ID", "customer_name", "os_name", "qty",
                       "amount_krw", "delivery_date", "market"]].copy()
        t.columns = ["SO_ID", "고객명", "품목", "수량", "금액", "납기일", "구분"]
        t["금액"] = t["금액"].apply(fmt_krw)
        t["납기일"] = t["납기일"].dt.strftime("%Y-%m-%d")
        st.dataframe(t, use_container_width=True, hide_index=True)
    else:
        st.info("이번 주 납기 예정 없음")

    # 최근 출고 5건
    st.subheader("최근 출고 5건")
    if not dn.empty:
        r = dn.nlargest(5, "dispatch_date")
        t2 = r[["DN_ID", "SO_ID", "qty", "amount_krw", "dispatch_date", "market"]].copy()
        t2.columns = ["DN_ID", "SO_ID", "수량", "금액", "출고일", "구분"]
        t2["금액"] = t2["금액"].apply(fmt_krw)
        t2["출고일"] = t2["출고일"].dt.strftime("%Y-%m-%d")
        st.dataframe(t2, use_container_width=True, hide_index=True)
    else:
        st.info("출고 데이터 없음")

    # 백로그 요약
    st.subheader("백로그 요약")
    if not backlog.empty:
        b1, b2 = st.columns(2)
        dom = backlog[backlog["market"] == "국내"]
        ovs = backlog[backlog["market"] == "해외"]
        b1.metric("국내 백로그", f"{len(dom)}건 / {fmt_krw(dom['ending_amount'].sum())}")
        b2.metric("해외 백로그", f"{len(ovs)}건 / {fmt_krw(ovs['ending_amount'].sum())}")
    else:
        st.info("백로그 데이터 없음")


# ═══════════════════════════════════════════════════════════════
# Page 2: 수주/출고 현황
# ═══════════════════════════════════════════════════════════════
def pg_orders(market, sectors, customers, year, month):
    st.title("수주/출고 현황")

    # 전체 (KPI용, 연도/월 필터 무시)
    so_all = filt(load_so(), market, sectors, customers)
    dn_ej = enrich_dn(load_dn(), load_so())
    dn_all = filt(dn_ej, market, sectors, customers, period_col=None)

    # 차트용 (연도/월 필터 적용)
    so = filt(so_all, "전체", [], [], year=year, month=month)
    dn = dn_all.copy()
    if not dn.empty:
        if year and year != "전체":
            dn = dn[dn["dispatch_month"].astype(str).str.startswith(year)]
            if month and month != "전체":
                dn = dn[dn["dispatch_month"] == f"{year}-{month}"]

    # ── 수주 KPI ──
    st.subheader("수주")
    month_so = so_all[so_all["period"] == _THIS_MONTH] if not so_all.empty else pd.DataFrame()
    c1, c2, c3 = st.columns(3)
    c1.metric("오늘 수주", "— (월별 집계)")
    c2.metric(
        "금월 수주",
        f"{len(month_so)}건 / {fmt_krw(month_so['amount_krw'].sum() if not month_so.empty else 0)}",
    )
    c3.metric(
        "누적 수주",
        f"{len(so_all)}건 / {fmt_krw(so_all['amount_krw'].sum() if not so_all.empty else 0)}",
    )

    # ── 출고 KPI ──
    st.subheader("출고")
    today_dn = dn_all[dn_all["dispatch_date"].dt.date == _TODAY_DATE] if not dn_all.empty else pd.DataFrame()
    month_dn = dn_all[dn_all["dispatch_month"] == _THIS_MONTH] if not dn_all.empty else pd.DataFrame()
    c4, c5, c6 = st.columns(3)
    c4.metric(
        "오늘 출고",
        f"{len(today_dn)}건 / {fmt_krw(today_dn['amount_krw'].sum() if not today_dn.empty else 0)}",
    )
    c5.metric(
        "금월 출고",
        f"{len(month_dn)}건 / {fmt_krw(month_dn['amount_krw'].sum() if not month_dn.empty else 0)}",
    )
    c6.metric(
        "누적 출고",
        f"{len(dn_all)}건 / {fmt_krw(dn_all['amount_krw'].sum() if not dn_all.empty else 0)}",
    )

    # ── 월별 수주/출고 금액 추이 ──
    st.subheader("월별 수주/출고 금액 추이")
    so_m = (
        so.groupby("period").agg(수주금액=("amount_krw", "sum"), 수주수량=("qty", "sum")).reset_index()
        if not so.empty else pd.DataFrame(columns=["period", "수주금액", "수주수량"])
    )
    dn_m = (
        dn.groupby("dispatch_month").agg(출고금액=("amount_krw", "sum"), 출고수량=("qty", "sum"))
        .reset_index().rename(columns={"dispatch_month": "period"})
        if not dn.empty else pd.DataFrame(columns=["period", "출고금액", "출고수량"])
    )
    merged = pd.merge(so_m, dn_m, on="period", how="outer").fillna(0).sort_values("period")

    if not merged.empty:
        merged["누적잔고"] = (merged["수주금액"] - merged["출고금액"]).cumsum()
        fig = go.Figure()
        fig.add_trace(go.Bar(x=merged["period"], y=merged["수주금액"],
                             name="수주", marker_color=C_INPUT))
        fig.add_trace(go.Bar(x=merged["period"], y=merged["출고금액"],
                             name="출고", marker_color=C_OUTPUT))
        fig.add_trace(go.Scatter(
            x=merged["period"], y=merged["누적잔고"], name="누적잔고",
            mode="lines+markers", line=dict(color=C_ENDING, width=2),
        ))
        fig.update_layout(barmode="group", height=400, margin=dict(t=30, b=30))
        st.plotly_chart(fig, use_container_width=True)

        # 수량 추이
        st.subheader("월별 수주/출고 수량 추이")
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(x=merged["period"], y=merged["수주수량"],
                              name="수주", marker_color=C_INPUT))
        fig2.add_trace(go.Bar(x=merged["period"], y=merged["출고수량"],
                              name="출고", marker_color=C_OUTPUT))
        fig2.update_layout(barmode="group", height=350, margin=dict(t=30, b=30))
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("데이터 없음")

    # 금월 일별 출고
    st.subheader("금월 일별 출고 현황")
    if not dn_all.empty:
        m_dn = dn_all[dn_all["dispatch_month"] == _THIS_MONTH].copy()
        if not m_dn.empty:
            m_dn["day"] = m_dn["dispatch_date"].dt.strftime("%m-%d")
            daily = m_dn.groupby("day")["amount_krw"].sum().reset_index()
            fig3 = px.bar(daily, x="day", y="amount_krw",
                          labels={"day": "날짜", "amount_krw": "출고금액(KRW)"})
            fig3.update_layout(height=300, margin=dict(t=30, b=30))
            st.plotly_chart(fig3, use_container_width=True)
        else:
            st.info("금월 출고 데이터 없음")


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
    fig.update_layout(
        height=max(400, len(top15) * 32),
        margin=dict(t=30, l=200),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig, use_container_width=True)

    # 구성비 + 월별 추이
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("제품 구성비")
        top8 = by_amt.head(8)
        etc = by_amt.iloc[8:].sum()
        parts = pd.concat([top8, pd.Series({"기타": etc})])
        fig2 = px.pie(values=parts.values, names=parts.index, hole=0.4)
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
            fig3.update_layout(height=400, margin=dict(t=30, b=30))
            st.plotly_chart(fig3, use_container_width=True)


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
        fig.update_layout(height=400, margin=dict(t=30, b=30))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("섹터별 월별 추이")
        m = so.groupby(["period", "sector"])["amount_krw"].sum().reset_index()
        if not m.empty:
            fig2 = px.bar(m, x="period", y="amount_krw", color="sector",
                          barmode="stack",
                          labels={"amount_krw": "매출", "period": "월"})
            fig2.update_layout(height=400, margin=dict(t=30, b=30))
            st.plotly_chart(fig2, use_container_width=True)

    # 섹터별 제품 믹스
    st.subheader("섹터별 제품 믹스")
    top_prods = so.groupby("os_name")["amount_krw"].sum().nlargest(8).index
    mix = (so[so["os_name"].isin(top_prods)]
           .groupby(["sector", "os_name"])["amount_krw"].sum().reset_index())
    if not mix.empty:
        fig3 = px.bar(mix, x="sector", y="amount_krw", color="os_name",
                      barmode="group", labels={"amount_krw": "매출"})
        fig3.update_layout(height=400, margin=dict(t=30, b=30))
        st.plotly_chart(fig3, use_container_width=True)


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
    month_custs = so[so["period"] == _THIS_MONTH]["customer_name"].nunique()

    # KPI
    c1, c2, c3 = st.columns(3)
    c1.metric("총 고객 수", f"{so['customer_name'].nunique()}")
    c2.metric("Top 고객", by_cust.index[0] if len(by_cust) else "-")
    c3.metric("금월 수주 고객 수", f"{month_custs}")

    # Top 15
    st.subheader("고객별 매출 Top 15")
    top15 = by_cust.head(15).reset_index()
    top15.columns = ["고객", "매출"]
    fig = px.bar(top15, y="고객", x="매출", orientation="h",
                 color_discrete_sequence=[C_INPUT])
    fig.update_layout(
        height=max(400, len(top15) * 32),
        margin=dict(t=30, l=200),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig, use_container_width=True)

    # Pareto (상위 20)
    st.subheader("고객 집중도 (Pareto)")
    pareto = by_cust.head(20).reset_index()
    pareto.columns = ["고객", "매출"]
    pareto["누적비율"] = pareto["매출"].cumsum() / pareto["매출"].sum() * 100

    fig2 = go.Figure()
    fig2.add_trace(go.Bar(x=pareto["고객"], y=pareto["매출"],
                          name="매출", marker_color=C_INPUT))
    fig2.add_trace(go.Scatter(
        x=pareto["고객"], y=pareto["누적비율"], name="누적 %",
        yaxis="y2", mode="lines+markers", line=dict(color=C_DANGER),
    ))
    fig2.update_layout(
        yaxis2=dict(title="누적 %", overlaying="y", side="right", range=[0, 105]),
        height=400, margin=dict(t=30, b=30),
    )
    st.plotly_chart(fig2, use_container_width=True)

    # 고객 상세 테이블
    st.subheader("고객 상세")
    detail = (
        so.groupby("customer_name")
        .agg(주문건수=("SO_ID", "nunique"), 총수량=("qty", "sum"),
             총금액=("amount_krw", "sum"), 최근수주월=("period", "max"))
        .sort_values("총금액", ascending=False)
        .reset_index()
    )
    detail["평균주문액"] = detail["총금액"] / detail["주문건수"]
    detail.rename(columns={"customer_name": "고객명"}, inplace=True)
    detail["총금액"] = detail["총금액"].apply(fmt_krw)
    detail["평균주문액"] = detail["평균주문액"].apply(fmt_krw)
    detail["총수량"] = detail["총수량"].apply(lambda x: f"{int(x):,}")
    st.dataframe(detail, use_container_width=True, hide_index=True)


# ═══════════════════════════════════════════════════════════════
# Page 6: Order Book (백로그)
# ═══════════════════════════════════════════════════════════════
def pg_orderbook(market, sectors, customers, **_):
    st.title("Order Book (백로그)")

    backlog = filt(load_backlog(), market, sectors, customers, period_col=None)
    snap_meta = load_snapshot_meta()

    # KPI
    c1, c2, c3, c4 = st.columns(4)
    if not backlog.empty:
        c1.metric("Backlog 금액", fmt_krw(backlog["ending_amount"].sum()))
        c2.metric("Backlog 수량", fmt_qty(backlog["ending_qty"].sum()))
        c3.metric("Backlog 건수", f"{len(backlog)}건")
    else:
        c1.metric("Backlog 금액", "₩0")
        c2.metric("Backlog 수량", "0")
        c3.metric("Backlog 건수", "0건")
    c4.metric("마지막 마감 월",
              snap_meta.iloc[0]["period"] if not snap_meta.empty else "없음")

    # Backlog 추이 (스냅샷 데이터)
    if not snap_meta.empty:
        st.subheader("Backlog 추이 (마감 기준)")
        conn = _conn()
        if conn:
            try:
                snap = pd.read_sql_query("""
                    SELECT snapshot_period,
                           SUM(ending_qty)    AS ending_qty,
                           SUM(ending_amount) AS ending_amount
                    FROM ob_snapshot
                    GROUP BY snapshot_period
                    ORDER BY snapshot_period
                """, conn)
            finally:
                conn.close()
            if not snap.empty:
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=snap["snapshot_period"], y=snap["ending_amount"],
                    name="Backlog 금액", marker_color=C_INPUT,
                ))
                fig.add_trace(go.Scatter(
                    x=snap["snapshot_period"], y=snap["ending_qty"],
                    name="Backlog 수량", yaxis="y2",
                    mode="lines+markers", line=dict(color=C_OUTPUT),
                ))
                fig.update_layout(
                    yaxis2=dict(title="수량", overlaying="y", side="right"),
                    height=350, margin=dict(t=30, b=30),
                )
                st.plotly_chart(fig, use_container_width=True)

    # Backlog 상세 테이블
    st.subheader("Backlog 상세")
    if not backlog.empty:
        t = backlog[["market", "SO_ID", "customer_name", "os_name",
                      "delivery_date", "ending_qty", "ending_amount",
                      "model_code", "sector"]].copy()
        t.columns = ["구분", "SO_ID", "고객명", "품목", "납기일",
                      "잔여수량", "잔여금액", "Model code", "Sector"]
        t["납기일"] = t["납기일"].dt.strftime("%Y-%m-%d").fillna("")
        t["잔여수량"] = t["잔여수량"].apply(lambda x: f"{int(x):,}")
        t["잔여금액"] = t["잔여금액"].apply(fmt_krw)
        st.dataframe(t, use_container_width=True, hide_index=True)
    else:
        st.info("백로그 데이터 없음")


if __name__ == "__main__":
    main()
