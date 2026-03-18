"""
NOAH 대시보드
============
수주/출고 현황, 제품/섹터/고객 분석, Order Book(백로그) 등 핵심 KPI.
데이터 소스: noah_data.db (SQLite, sync_db.py로 동기화)

Usage:
    streamlit run dashboard.py
"""

import calendar
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
                   [EXW NOAH] AS exw_noah,
                   COALESCE(Status, '') AS status,
                   COALESCE([Customer PO], '') AS customer_po,
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
                   [EXW NOAH],
                   COALESCE(Status, ''),
                   COALESCE([Customer PO], ''),
                   '해외'
            FROM so_export
            WHERE COALESCE(Status, '') != 'Cancelled'
              AND Period IS NOT NULL AND TRIM(Period) != ''
        """, conn)
    except Exception:
        return pd.DataFrame()
    finally:
        conn.close()
    df["delivery_date"] = pd.to_datetime(df["delivery_date"], errors="coerce")
    df["exw_noah"] = pd.to_datetime(df["exw_noah"], errors="coerce")
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
    except Exception:
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
                   [B/L]          AS bl_no
            FROM dn_export
            WHERE [출고일] IS NOT NULL AND TRIM(COALESCE([출고일], '')) != ''
        """, conn)
    except Exception:
        return pd.DataFrame()
    finally:
        conn.close()
    for c in ("factory_date", "pickup_date", "expected_ship_date", "ship_date"):
        df[c] = pd.to_datetime(df[c], errors="coerce")
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
    except Exception:
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
    except Exception:
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
# 납기 캘린더 (pg_today 내부 사용)
# ═══════════════════════════════════════════════════════════════
def build_calendar_data(so_pending: pd.DataFrame, dn: pd.DataFrame,
                        year: int, month: int) -> pd.DataFrame:
    """월별 납기/출고 집계 — 날짜별 건수·금액 반환 (테스트 가능한 순수 함수)."""
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

    merged = days.merge(so_agg, on="day", how="left").merge(dn_agg, on="day", how="left")
    for c in ("so_count", "so_amount", "dn_count", "dn_amount"):
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
    cal_data = build_calendar_data(so_pending, dn, cy, cm)

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

                # Z값: 양수=미래/오늘 납기, 음수=과납기
                from datetime import date as _date_cls
                cell_date = _date_cls(cy, cm, day)
                if so_cnt > 0 and cell_date < _TODAY_DATE:
                    z_val = -so_cnt  # 과납기
                else:
                    z_val = so_cnt

                # 텍스트 조립
                lines = [f"<b>{day}</b>"]
                if so_cnt > 0:
                    lines.append(f"\U0001F4E6 {so_cnt}건 {fmt_krw(so_amt)}")
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
    sel_date = st.date_input(
        "날짜 선택",
        value=None, key="cal_date_pick",
        min_value=datetime(cy, cm, 1).date(),
        max_value=datetime(cy, cm, calendar.monthrange(cy, cm)[1]).date(),
    )
    sel_date_str = sel_date.strftime("%Y-%m-%d") if sel_date else None

    if sel_date_str:
        st.markdown(f"---\n#### {sel_date_str} 상세")
        d_a, d_b = st.columns(2)

        # (A) 납기 예정
        with d_a:
            st.markdown("**📦 납기 예정**")
            if not so_pending.empty:
                day_so = so_pending[so_pending["delivery_date"].dt.date == sel_date]
            else:
                day_so = pd.DataFrame()

            if not day_so.empty:
                agg = dict(
                    고객명=("customer_name", "first"),
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
                    items.append({
                        "title": f"{icon}  **{r['SO_ID']}**  {r['고객명']}",
                        "lines": [
                            f"품목 {r['품목수']}건 · 수량 {int(r['총수량']):,} · {fmt_krw(r['총금액'])}{po_info}",
                            f"📅 EXW {fmt_date(r['공장출고일'])}",
                        ],
                    })
                _render_cards(items, cols_per_row=2)
            else:
                st.info("납기 예정 건 없음")

        # (B) 출고 실적
        with d_b:
            st.markdown("**🚚 출고 실적**")
            if not dn.empty:
                day_dn = dn[dn["dispatch_date"].dt.date == sel_date]
            else:
                day_dn = pd.DataFrame()

            if not day_dn.empty:
                # SO에서 customer_po 조인
                so_po = so[["SO_ID", "customer_po"]].drop_duplicates(subset=["SO_ID"])
                day_dn = day_dn.merge(so_po, on="SO_ID", how="left")
                day_dn["customer_po"] = day_dn["customer_po"].fillna("")
                agg_dict = dict(
                    총수량=("qty", "sum"),
                    총금액=("amount_krw", "sum"),
                    고객PO=("customer_po", "first"),
                )
                if "customer_name" in day_dn.columns:
                    agg_dict["고객명"] = ("customer_name", "first")
                agg_dict["SO_ID"] = ("SO_ID", "first")
                if "line_item" in day_dn.columns:
                    agg_dict["품목수"] = ("line_item", "nunique")
                tbl = day_dn.groupby("DN_ID").agg(**agg_dict).reset_index()
                items = []
                for _, r in tbl.iterrows():
                    cust = r.get("고객명", "")
                    n_items = r.get("품목수", "?")
                    po_info = f" · PO: {r['고객PO']}" if r.get("고객PO") else ""
                    items.append({
                        "title": f"📦 **{r['DN_ID']}**  {cust}",
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

    # ── 납기 현황 (국내/해외 탭) ──
    st.subheader("납기 현황 (미완료 건)")
    # 이번 주 + 지연 건 합산
    show_due = pd.concat([overdue, week_due]).drop_duplicates() if not overdue.empty or not week_due.empty else pd.DataFrame()

    if not show_due.empty:
        tab_dom, tab_exp = st.tabs(["🇰🇷 국내", "🌏 해외"])
        for tab, mkt in [(tab_dom, "국내"), (tab_exp, "해외")]:
            with tab:
                mkt_df = show_due[show_due["market"] == mkt]
                if mkt_df.empty:
                    st.info(f"{mkt} 납기 예정/지연 건 없음")
                    continue
                # SO_ID별 그룹
                g = mkt_df.groupby("SO_ID").agg(
                    고객명=("customer_name", "first"),
                    고객PO=("customer_po", "first"),
                    품목수=("line_item", "nunique"),
                    총수량=("qty", "sum"),
                    총금액=("amount_krw", "sum"),
                    납기일=("delivery_date", "min"),
                    공장출고일=("exw_noah", "min"),
                    Status=("status", "first"),
                ).reset_index().sort_values("납기일")
                g["지연"] = g["납기일"].dt.date < _TODAY_DATE
                items = []
                for _, r in g.iterrows():
                    icon = _status_icon(r["Status"], r["지연"])
                    po_info = f" · PO: {r['고객PO']}" if r["고객PO"] else ""
                    items.append({
                        "title": f"{icon}  **{r['SO_ID']}**  {r['고객명']}",
                        "lines": [
                            f"품목 {r['품목수']}건 · 수량 {int(r['총수량']):,} · {fmt_krw(r['총금액'])}{po_info}",
                            f"📅 납기 {fmt_date(r['납기일'])} · EXW {fmt_date(r['공장출고일'])}",
                        ],
                    })
                _render_cards(items)
    else:
        st.info("납기 예정/지연 건 없음")

    # ── 해외 선적 Action Items ──
    if market != "국내":
        st.subheader("🚢 해외 선적 Action Items")
        ship_df = load_dn_export_shipping()
        if not ship_df.empty:
            # SO 메타 조인으로 sector, customer_po 추가
            so_meta = load_so()[["SO_ID", "sector", "customer_po"]].drop_duplicates(subset=["SO_ID"])
            ship_df = ship_df.merge(so_meta, on="SO_ID", how="left")
            ship_df["sector"] = ship_df["sector"].fillna("")
            ship_df["customer_po"] = ship_df["customer_po"].fillna("")
        if sectors and not ship_df.empty:
            ship_df = ship_df[ship_df["sector"].isin(sectors)]
        if customers and not ship_df.empty:
            ship_df = ship_df[ship_df["customer_name"].isin(customers)]
        if not ship_df.empty:
            # 공장 출고 완료 but 선적 미완료 → CI/PL 작성 + 포워더 필요
            pending_ship = ship_df[ship_df["ship_date"].isna()].copy()
            if not pending_ship.empty:
                st.markdown("**선적 대기** — CI/PL 작성 및 포워더 arrange 필요")
                ps = pending_ship.groupby("DN_ID").agg(
                    고객명=("customer_name", "first"),
                    고객PO=("customer_po", "first"),
                    품목수=("item_name", "nunique"),
                    총수량=("qty", "sum"),
                    총금액=("amount_krw", "sum"),
                    공장출고일=("factory_date", "min"),
                    픽업일=("pickup_date", "min"),
                    선적예정일=("expected_ship_date", "min"),
                    BL=("bl_no", "first"),
                ).reset_index().sort_values("공장출고일")
                items = []
                for _, r in ps.iterrows():
                    bl = r["BL"] if pd.notna(r["BL"]) else ""
                    po_info = f" · PO: {r['고객PO']}" if r["고객PO"] else ""
                    lines = [
                        f"품목 {r['품목수']}건 · 수량 {int(r['총수량']):,} · {fmt_krw(r['총금액'])}{po_info}",
                        f"출고 {fmt_date(r['공장출고일'])} → 픽업 {fmt_date(r['픽업일'])} → 선적예정 {fmt_date(r['선적예정일'])}",
                    ]
                    if bl:
                        lines.append(f"B/L: {bl}")
                    items.append({
                        "title": f"⏳ **{r['DN_ID']}**  {r['고객명']}",
                        "lines": lines,
                    })
                _render_cards(items)
            else:
                st.success("선적 대기 건 없음")

        else:
            st.info("해외 출고 데이터 없음")

    # ── 백로그 요약 ──
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
            kpi_month = year_periods.max()
            y, m = int(kpi_month[:4]), int(kpi_month[5:7])
            kpi_prev = f"{y}-{m - 1:02d}" if m > 1 else f"{y - 1}-12"
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

    # ── 일별 출고 ──
    st.subheader(f"{kpi_label} 일별 출고 현황")
    if not dn_all.empty:
        m_dn = dn_all[dn_all["dispatch_month"] == kpi_month].copy()
        if not m_dn.empty:
            m_dn["day"] = m_dn["dispatch_date"].dt.strftime("%m-%d")
            daily = m_dn.groupby("day")["amount_krw"].sum().reset_index()
            fig3 = px.bar(daily, x="day", y="amount_krw",
                          labels={"day": "날짜", "amount_krw": "출고금액(KRW)"})
            fig3.update_traces(hovertemplate="<b>%{x}</b><br>출고: ₩%{y:,.0f}<extra></extra>")
            fig3.update_layout(height=300, margin=dict(t=30, b=30),
                               xaxis=dict(type="category"))
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
    pareto = by_cust.head(20).reset_index()
    pareto.columns = ["고객", "매출"]
    _pareto_total = pareto["매출"].sum()
    pareto["누적비율"] = pareto["매출"].cumsum() / _pareto_total * 100 if _pareto_total else 0
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


# ═══════════════════════════════════════════════════════════════
# Page 6: Order Book (백로그)
# ═══════════════════════════════════════════════════════════════
def pg_orderbook(market, sectors, customers, **_):
    st.title("Order Book (백로그)")
    if _.get("year", "전체") != "전체" or _.get("month", "전체") != "전체":
        st.caption("ℹ️ 이 페이지는 현재 잔고 기준입니다 — 연도/월 필터는 적용되지 않습니다.")

    backlog = filt(load_backlog(), market, sectors, customers, period_col=None)

    today_ts = pd.Timestamp.today().normalize()

    # KPI
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

    # ── 월별 Input / Output / Ending 추이 (order_book.sql) ──
    try:
        ob = load_order_book()
    except Exception as e:
        st.error(f"Order Book 데이터 로드 실패: {e}")
        ob = pd.DataFrame()
    if not ob.empty:
        ob = ob.rename(columns={"구분": "market", "Sector": "sector", "Customer name": "customer_name"})
        ob = filt(ob, market, sectors, customers, period_col=None)
        st.subheader("월별 수주잔고 추이 — 실시간")
        st.caption("현재 DB 상태 기준 계산 (마감 확정치 아님)")
        # Input/Output은 flow → 단순 합산, Ending은 stock → 라인별 forward-fill 후 합산
        _line_key = ["SO_ID", "OS name", "Expected delivery date"]
        ob_flow = ob.groupby("Period").agg(
            Input=("Value_Input_amount", "sum"),
            Output=("Value_Output_amount", "sum"),
        ).reset_index()
        # Ending: 라인별 마지막 이벤트 Ending을 빈 월로 전파 후 합산
        _line_end = ob.groupby(_line_key + ["Period"])["Value_Ending_amount"].sum().reset_index()
        _pivot = _line_end.pivot_table(
            index=_line_key, columns="Period", values="Value_Ending_amount",
        ).sort_index(axis=1).ffill(axis=1).fillna(0)
        _monthly_ending = _pivot.sum(axis=0).reset_index()
        _monthly_ending.columns = ["Period", "Ending"]
        ob_monthly = ob_flow.merge(_monthly_ending, on="Period", how="outer").fillna(0).sort_values("Period")
        if not ob_monthly.empty:
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

    # ── Aging 분석 ──
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

    # ── 섹터별 / 고객별 Backlog ──
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

    # ── 납기 분포 히트맵 ──
    if not backlog.empty:
        st.subheader("납기 분포 히트맵")
        st.caption(f"금월 ~ {today_ts.year}년 말 — 월별 × 섹터 납기 예정 금액")
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
            # hover에 금액 포맷 표시
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


if __name__ == "__main__":
    main()
