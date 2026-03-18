"""
대시보드 스모크 테스트
=====================
필터 정합성, 기간 로직, 0-나눗셈 방어 등 핵심 로직만 검증.
Streamlit UI 렌더링은 테스트하지 않음 — 순수 데이터 로직만.
"""

import pandas as pd
import pytest

from dashboard import build_calendar_data, enrich_dn, filt, fmt_krw


# ═══════════════════════════════════════════════════════════════
# 공통 픽스처
# ═══════════════════════════════════════════════════════════════
@pytest.fixture
def sample_so():
    """SO 샘플 — 국내/해외, 2개 섹터, 3개월"""
    return pd.DataFrame({
        "SO_ID": ["ND-001", "ND-001", "ND-002", "NO-001", "NO-002"],
        "line_item": [1, 2, 1, 1, 1],
        "customer_name": ["고객A", "고객A", "고객B", "고객C", "고객C"],
        "os_name": ["IQ3", "CVA", "IQ3", "IQ3", "CVA"],
        "sector": ["Oil&Gas", "Oil&Gas", "Water", "Power", "Power"],
        "market": ["국내", "국내", "국내", "해외", "해외"],
        "period": ["2025-01", "2025-01", "2025-02", "2025-01", "2026-03"],
        "qty": [10, 5, 20, 15, 8],
        "amount_krw": [1_000_000, 500_000, 2_000_000, 3_000_000, 800_000],
        "delivery_date": pd.to_datetime(
            ["2025-02-01", "2025-02-01", "2025-03-15", "2025-04-01", "2026-04-01"]
        ),
        "exw_noah": pd.NaT,
        "status": ["", "", "", "", ""],
    })


@pytest.fixture
def sample_dn():
    """DN 샘플"""
    return pd.DataFrame({
        "DN_ID": ["DND-001", "DND-002", "DNO-001"],
        "SO_ID": ["ND-001", "ND-002", "NO-001"],
        "line_item": [1, 1, 1],
        "qty": [10, 10, 15],
        "amount_krw": [1_000_000, 1_000_000, 3_000_000],
        "dispatch_date": pd.to_datetime(["2025-02-10", "2025-03-05", "2025-03-20"]),
        "dispatch_month": ["2025-02", "2025-03", "2025-03"],
        "market": ["국내", "국내", "해외"],
    })


@pytest.fixture
def sample_backlog():
    """Backlog 샘플"""
    return pd.DataFrame({
        "SO_ID": ["ND-001", "NO-002"],
        "os_name": ["CVA", "CVA"],
        "customer_name": ["고객A", "고객C"],
        "market": ["국내", "해외"],
        "sector": ["Oil&Gas", "Power"],
        "ending_qty": [5, 8],
        "ending_amount": [500_000, 800_000],
        "delivery_date": pd.to_datetime(["2025-02-01", "2026-04-01"]),
    })


# ═══════════════════════════════════════════════════════════════
# filt() 필터 함수 테스트
# ═══════════════════════════════════════════════════════════════
class TestFilt:
    """filt() 헬퍼 단위 테스트"""

    def test_market_filter(self, sample_so):
        """시장 필터 — 국내만"""
        result = filt(sample_so, "국내", [], [])
        assert (result["market"] == "국내").all()
        assert len(result) == 3

    def test_market_all(self, sample_so):
        """시장=전체 — 필터 미적용"""
        result = filt(sample_so, "전체", [], [])
        assert len(result) == len(sample_so)

    def test_sector_filter(self, sample_so):
        """섹터 필터"""
        result = filt(sample_so, "전체", ["Water"], [])
        assert (result["sector"] == "Water").all()
        assert len(result) == 1

    def test_customer_filter(self, sample_so):
        """고객 필터"""
        result = filt(sample_so, "전체", [], ["고객C"])
        assert (result["customer_name"] == "고객C").all()
        assert len(result) == 2

    def test_year_filter(self, sample_so):
        """연도 필터"""
        result = filt(sample_so, "전체", [], [], year="2025")
        assert all(r.startswith("2025") for r in result["period"])
        assert len(result) == 4

    def test_year_month_filter(self, sample_so):
        """연도+월 필터"""
        result = filt(sample_so, "전체", [], [], year="2025", month="01")
        assert (result["period"] == "2025-01").all()
        assert len(result) == 3

    def test_month_without_year_ignored(self, sample_so):
        """연도=전체 + 월=03 → 월 필터 무시 (filt 동작 확인)"""
        result = filt(sample_so, "전체", [], [], year="전체", month="03")
        assert len(result) == len(sample_so)

    def test_combined_filters(self, sample_so):
        """시장+섹터+연월 복합 필터"""
        result = filt(sample_so, "국내", ["Oil&Gas"], [], year="2025", month="01")
        assert len(result) == 2
        assert (result["market"] == "국내").all()
        assert (result["sector"] == "Oil&Gas").all()
        assert (result["period"] == "2025-01").all()

    def test_empty_df(self):
        """빈 DataFrame 입력 — 에러 없이 빈 DF 반환"""
        result = filt(pd.DataFrame(), "국내", ["Oil&Gas"], ["고객A"])
        assert result.empty

    def test_period_col_none(self, sample_backlog):
        """period_col=None — 기간 필터 스킵 (backlog용)"""
        result = filt(sample_backlog, "국내", [], [], period_col=None, year="2025")
        # period_col=None이면 year 필터 적용 안 됨, market만 적용
        assert len(result) == 1
        assert result.iloc[0]["market"] == "국내"


# ═══════════════════════════════════════════════════════════════
# 필터 일관성 테스트 (페이지 간)
# ═══════════════════════════════════════════════════════════════
class TestFilterConsistency:
    """모든 데이터 소스에 동일 필터가 적용되는지 검증"""

    def test_backlog_respects_market_filter(self, sample_backlog):
        """Backlog에 시장 필터 적용 — 제품/섹터/고객 분석 공통"""
        filtered = filt(sample_backlog, "해외", [], [], period_col=None)
        assert len(filtered) == 1
        assert filtered.iloc[0]["market"] == "해외"

    def test_backlog_respects_sector_filter(self, sample_backlog):
        """Backlog에 섹터 필터 적용"""
        filtered = filt(sample_backlog, "전체", ["Power"], [], period_col=None)
        assert len(filtered) == 1
        assert filtered.iloc[0]["sector"] == "Power"

    def test_backlog_respects_customer_filter(self, sample_backlog):
        """Backlog에 고객 필터 적용"""
        filtered = filt(sample_backlog, "전체", [], ["고객A"], period_col=None)
        assert len(filtered) == 1
        assert filtered.iloc[0]["customer_name"] == "고객A"

    def test_so_and_backlog_same_market_scope(self, sample_so, sample_backlog):
        """SO와 Backlog에 같은 필터 적용 시 같은 시장 범위"""
        kw = dict(market="해외", sectors=[], customers=[], period_col=None)
        so_markets = filt(sample_so, **kw)["market"].unique()
        bl_markets = filt(sample_backlog, **kw)["market"].unique()
        assert set(so_markets) == set(bl_markets)


# ═══════════════════════════════════════════════════════════════
# enrich_dn 테스트
# ═══════════════════════════════════════════════════════════════
class TestEnrichDN:
    """DN에 SO 메타데이터 병합"""

    def test_enrich_adds_customer(self, sample_dn, sample_so):
        result = enrich_dn(sample_dn, sample_so)
        assert "customer_name" in result.columns
        assert result.loc[result["DN_ID"] == "DND-001", "customer_name"].iloc[0] == "고객A"

    def test_enrich_adds_sector(self, sample_dn, sample_so):
        result = enrich_dn(sample_dn, sample_so)
        assert "sector" in result.columns

    def test_enrich_empty_dn(self, sample_so):
        result = enrich_dn(pd.DataFrame(), sample_so)
        assert result.empty


# ═══════════════════════════════════════════════════════════════
# 기간 로직 테스트
# ═══════════════════════════════════════════════════════════════
class TestPeriodLogic:
    """KPI 기준 월 결정 로직 (pg_orders, pg_customer에서 사용)"""

    def test_kpi_month_specific_selection(self):
        """year+month 선택 시 해당 월이 KPI 기준"""
        year, month = "2025", "06"
        kpi_month = f"{year}-{month}"
        y, m = int(year), int(month)
        kpi_prev = f"{y}-{m - 1:02d}" if m > 1 else f"{y - 1}-12"
        assert kpi_month == "2025-06"
        assert kpi_prev == "2025-05"

    def test_kpi_month_january_wraps(self):
        """1월 선택 시 전월 = 전년 12월"""
        year, month = "2025", "01"
        y, m = int(year), int(month)
        kpi_prev = f"{y}-{m - 1:02d}" if m > 1 else f"{y - 1}-12"
        assert kpi_prev == "2024-12"

    def test_customer_kpi_year_only(self, sample_so):
        """연도만 선택 시 (월=전체) — so 내 최근 월 기준"""
        so = filt(sample_so, "전체", [], [], year="2025")
        latest = so["period"].max()
        assert latest == "2025-02"
        month_custs = so[so["period"] == latest]["customer_name"].nunique()
        assert month_custs == 1  # 고객B만

    def test_customer_kpi_year_month(self, sample_so):
        """연도+월 선택 시 — so 전체가 해당 월이므로 nunique 사용"""
        so = filt(sample_so, "전체", [], [], year="2025", month="01")
        assert so["customer_name"].nunique() == 2  # 고객A, 고객C

    def test_kpi_month_year_only(self, sample_so):
        """연도만 선택(월=전체) → 해당 연도 내 최신 월이 KPI 기준"""
        year, month = "2025", "전체"
        so_all = filt(sample_so, "전체", [], [])
        # dashboard.py의 year-only 분기 미러링
        if year and year != "전체" and (not month or month == "전체"):
            year_periods = so_all[so_all["period"].str.startswith(year)]["period"]
            kpi_month = year_periods.max()
            y, m = int(kpi_month[:4]), int(kpi_month[5:7])
            kpi_prev = f"{y}-{m - 1:02d}" if m > 1 else f"{y - 1}-12"
        assert kpi_month == "2025-02"   # 2025 내 최신
        assert kpi_prev == "2025-01"

    def test_kpi_month_year_only_no_data(self, sample_so):
        """존재하지 않는 연도 → 빈 periods → 금월 fallback"""
        so_all = filt(sample_so, "전체", [], [])
        year_periods = so_all[so_all["period"].str.startswith("2020")]["period"]
        assert year_periods.empty  # fallback 조건 확인


# ═══════════════════════════════════════════════════════════════
# 엣지 케이스 / 0 나눗셈 방어
# ═══════════════════════════════════════════════════════════════
class TestEdgeCases:
    """0 나눗셈, 빈 데이터 등 극단값 방어"""

    def test_avg_price_zero_qty(self):
        """수량 0인 제품 — 평균단가 NaN (에러 아님)"""
        df = pd.DataFrame({
            "os_name": ["IQ3", "CVA"],
            "amount_krw": [1_000_000, 500_000],
            "qty": [0, 10],
        })
        avg = df.groupby("os_name").agg(총금액=("amount_krw", "sum"), 총수량=("qty", "sum"))
        avg["평균단가"] = avg["총금액"] / avg["총수량"].replace(0, pd.NA)
        assert pd.isna(avg.loc["IQ3", "평균단가"])
        assert avg.loc["CVA", "평균단가"] == 50_000

    def test_pareto_zero_total(self):
        """매출 합계 0 — Pareto 누적비율 0"""
        pareto = pd.DataFrame({"고객": ["A"], "매출": [0]})
        total = pareto["매출"].sum()
        pareto["누적비율"] = pareto["매출"].cumsum() / total * 100 if total else 0
        assert pareto["누적비율"].iloc[0] == 0

    def test_btb_zero_output_displays_na(self):
        """출고 0 → 표시값 'N/A' (dashboard.py의 분기 로직 미러링)"""
        this_dn_amt = 0
        this_so_amt = 1_000_000
        display = f"{this_so_amt / this_dn_amt:.2f}" if this_dn_amt else "N/A"
        assert display == "N/A"

    def test_btb_normal_display(self):
        """정상 BtB — 소수점 2자리 포맷"""
        this_dn_amt = 1_000_000
        this_so_amt = 2_000_000
        display = f"{this_so_amt / this_dn_amt:.2f}" if this_dn_amt else "N/A"
        assert display == "2.00"


# ═══════════════════════════════════════════════════════════════
# 포맷 유틸 테스트
# ═══════════════════════════════════════════════════════════════
class TestFmtKrw:
    """KRW 포맷 헬퍼"""

    def test_zero(self):
        assert fmt_krw(0) == "₩0"

    def test_nan(self):
        assert fmt_krw(float("nan")) == "₩0"

    def test_억(self):
        assert "억" in fmt_krw(1.5e8)

    def test_만(self):
        assert "만" in fmt_krw(5e4)

    def test_negative(self):
        assert fmt_krw(-2e8).startswith("-")


# ═══════════════════════════════════════════════════════════════
# 드릴다운 데이터 필터링 로직 테스트
# ═══════════════════════════════════════════════════════════════
class TestDrillDown:
    """드릴다운 클릭 시 표시되는 서브 데이터 필터링 검증"""

    # ── 제품 분석 드릴다운 ──
    def test_product_drilldown_monthly(self, sample_so):
        """제품 클릭 → 월별 매출 데이터"""
        sub = sample_so[sample_so["os_name"] == "IQ3"]
        pm = sub.groupby("period")["amount_krw"].sum().reset_index()
        assert len(pm) == 2  # 2025-01, 2025-02
        assert pm["amount_krw"].sum() == 6_000_000  # 1M + 2M + 3M

    def test_product_drilldown_sector_mix(self, sample_so):
        """제품 클릭 → 섹터별 비중"""
        sub = sample_so[sample_so["os_name"] == "IQ3"]
        ps = sub.groupby("sector")["amount_krw"].sum().reset_index()
        assert set(ps["sector"]) == {"Oil&Gas", "Water", "Power"}

    def test_product_drilldown_top_customers(self, sample_so):
        """제품 클릭 → 주요 고객 Top 5"""
        sub = sample_so[sample_so["os_name"] == "IQ3"]
        pc = sub.groupby("customer_name")["amount_krw"].sum().nlargest(5).reset_index()
        assert pc.iloc[0]["customer_name"] == "고객C"  # 3M 최대
        assert pc.iloc[0]["amount_krw"] == 3_000_000

    # ── 섹터 분석 드릴다운 ──
    def test_sector_drilldown_product_mix(self, sample_so):
        """섹터 클릭 → 제품 믹스"""
        sub = sample_so[sample_so["sector"] == "Oil&Gas"]
        sp = sub.groupby("os_name")["amount_krw"].sum().nlargest(10).reset_index()
        assert set(sp["os_name"]) == {"IQ3", "CVA"}

    def test_sector_drilldown_customers(self, sample_so):
        """섹터 클릭 → 주요 고객"""
        sub = sample_so[sample_so["sector"] == "Power"]
        sc = sub.groupby("customer_name")["amount_krw"].sum().nlargest(5).reset_index()
        assert len(sc) == 1
        assert sc.iloc[0]["customer_name"] == "고객C"

    def test_sector_drilldown_monthly(self, sample_so):
        """섹터 클릭 → 월별 추이"""
        sub = sample_so[sample_so["sector"] == "Oil&Gas"]
        sm = sub.groupby("period")["amount_krw"].sum().reset_index()
        assert len(sm) == 1  # 2025-01만
        assert sm["amount_krw"].sum() == 1_500_000  # 1M + 500K

    # ── 고객 분석 드릴다운 ──
    def test_customer_drilldown_monthly(self, sample_so):
        """고객 클릭 → 월별 매출 추이"""
        sub = sample_so[sample_so["customer_name"] == "고객A"]
        cm = sub.groupby("period")["amount_krw"].sum().reset_index()
        assert len(cm) == 1  # 2025-01
        assert cm["amount_krw"].sum() == 1_500_000  # 1M + 500K

    def test_customer_drilldown_product_mix(self, sample_so):
        """고객 클릭 → 제품 믹스"""
        sub = sample_so[sample_so["customer_name"] == "고객A"]
        cp = sub.groupby("os_name")["amount_krw"].sum().nlargest(8).reset_index()
        assert set(cp["os_name"]) == {"IQ3", "CVA"}

    def test_customer_drilldown_backlog(self, sample_backlog):
        """고객 클릭 → 백로그 현황"""
        cust_bl = sample_backlog[sample_backlog["customer_name"] == "고객A"]
        assert len(cust_bl) == 1
        assert cust_bl["ending_amount"].sum() == 500_000

    # ── Aging 드릴다운 ──
    def test_aging_drilldown(self, sample_backlog):
        """Aging 구간 클릭 → 해당 건 필터링"""
        bl = sample_backlog.copy()
        today_ts = pd.Timestamp("2026-03-18")
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
        # ND-001: 2025-02-01 → 30일+ 지연
        overdue = bl[bl["aging"] == "① 30일+ 지연"]
        assert len(overdue) == 1
        assert overdue.iloc[0]["SO_ID"] == "ND-001"
        # NO-002: 2026-04-01 → 2주 이내
        imminent = bl[bl["aging"] == "③ 2주 이내"]
        assert len(imminent) == 1
        assert imminent.iloc[0]["SO_ID"] == "NO-002"

    def test_drilldown_empty_result(self, sample_so):
        """존재하지 않는 항목 클릭 → 빈 결과"""
        sub = sample_so[sample_so["os_name"] == "NONEXISTENT"]
        assert sub.empty

    def test_drilldown_single_item(self, sample_so):
        """단일 건만 있는 제품 드릴다운"""
        sub = sample_so[sample_so["os_name"] == "CVA"]
        pm = sub.groupby("period")["amount_krw"].sum().reset_index()
        # CVA: 2025-01 (500K) + 2026-03 (800K)
        assert len(pm) == 2
        assert pm["amount_krw"].sum() == 1_300_000


# ═══════════════════════════════════════════════════════════════
# 납기 캘린더 데이터 테스트
# ═══════════════════════════════════════════════════════════════
class TestCalendarData:
    """build_calendar_data() 순수 함수 단위 테스트"""

    @pytest.fixture
    def cal_so(self):
        """2026-03 납기 SO 샘플"""
        return pd.DataFrame({
            "SO_ID": ["ND-A", "ND-A", "ND-B"],
            "customer_name": ["고객A", "고객A", "고객B"],
            "os_name": ["IQ3", "CVA", "IQ3"],
            "qty": [10, 5, 20],
            "amount_krw": [1_000_000, 500_000, 2_000_000],
            "delivery_date": pd.to_datetime(["2026-03-10", "2026-03-10", "2026-03-25"]),
            "exw_noah": pd.NaT,
            "status": ["", "", ""],
            "market": ["국내", "국내", "국내"],
        })

    @pytest.fixture
    def cal_dn(self):
        """2026-03 출고 DN 샘플"""
        return pd.DataFrame({
            "DN_ID": ["DND-X", "DND-Y"],
            "SO_ID": ["ND-A", "ND-A"],
            "qty": [10, 5],
            "amount_krw": [1_000_000, 500_000],
            "dispatch_date": pd.to_datetime(["2026-03-10", "2026-03-15"]),
            "dispatch_month": ["2026-03", "2026-03"],
            "market": ["국내", "국내"],
        })

    def test_all_days_present(self, cal_so, cal_dn):
        """3월 = 31일 → 31행 반환"""
        result = build_calendar_data(cal_so, cal_dn, 2026, 3)
        assert len(result) == 31
        assert list(result["day"]) == list(range(1, 32))

    def test_so_aggregation(self, cal_so, cal_dn):
        """3/10에 SO 2건(ND-A unique 1건) → so_count=1, so_amount=1.5M"""
        result = build_calendar_data(cal_so, cal_dn, 2026, 3)
        day10 = result[result["day"] == 10].iloc[0]
        assert day10["so_count"] == 1  # ND-A (nunique)
        assert day10["so_amount"] == 1_500_000

    def test_dn_aggregation(self, cal_so, cal_dn):
        """3/10에 DN 1건, 3/15에 DN 1건"""
        result = build_calendar_data(cal_so, cal_dn, 2026, 3)
        day10 = result[result["day"] == 10].iloc[0]
        assert day10["dn_count"] == 1
        assert day10["dn_amount"] == 1_000_000
        day15 = result[result["day"] == 15].iloc[0]
        assert day15["dn_count"] == 1
        assert day15["dn_amount"] == 500_000

    def test_empty_day_zero_fill(self, cal_so, cal_dn):
        """이벤트 없는 날 → 모두 0"""
        result = build_calendar_data(cal_so, cal_dn, 2026, 3)
        day1 = result[result["day"] == 1].iloc[0]
        assert day1["so_count"] == 0
        assert day1["so_amount"] == 0
        assert day1["dn_count"] == 0
        assert day1["dn_amount"] == 0

    def test_both_empty(self):
        """so_pending + dn 모두 빈 경우 → 빈 달력(0으로 채움)"""
        result = build_calendar_data(pd.DataFrame(), pd.DataFrame(), 2026, 2)
        assert len(result) == 28  # 2026-02 = 28일
        assert result["so_count"].sum() == 0
        assert result["dn_count"].sum() == 0

    def test_dn_multiline_counts_unique(self):
        """같은 DN_ID 다중 라인 → dn_count=1 (nunique), dn_amount=합산"""
        dn_multi = pd.DataFrame({
            "DN_ID": ["DND-A", "DND-A", "DND-B"],
            "SO_ID": ["ND-1", "ND-1", "ND-2"],
            "qty": [10, 5, 20],
            "amount_krw": [1_000_000, 500_000, 2_000_000],
            "dispatch_date": pd.to_datetime(["2026-03-10", "2026-03-10", "2026-03-10"]),
            "dispatch_month": ["2026-03", "2026-03", "2026-03"],
            "market": ["국내", "국내", "국내"],
        })
        result = build_calendar_data(pd.DataFrame(), dn_multi, 2026, 3)
        day10 = result[result["day"] == 10].iloc[0]
        assert day10["dn_count"] == 2  # DND-A, DND-B (nunique)
        assert day10["dn_amount"] == 3_500_000
