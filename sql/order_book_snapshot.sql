-- ═══════════════════════════════════════════════════════════════
-- Order Book — 스냅샷 기반 (params.period로 임의 월 조회)
-- ═══════════════════════════════════════════════════════════════
-- params CTE의 period 한 줄만 바꿔 조회:
--   • 마감(closed)월 → ob_snapshot 동결값 그대로
--                      (월말 총 백로그 = 무활동 그룹 포함 전 그룹, close_period --list 총계와 일치)
--   • 미마감(open)월 → 라이브 롤링 (Start=마지막스냅샷Ending, Variance=소급변경분)
--                      스냅샷 없으면 순수 롤링 (활동 있는 월만)
--
-- 전제: sync_db.py + close_period.py로 스냅샷 생성

WITH
params AS (SELECT '2026-05' AS period),   -- ← 조회할 월 (yyyy-MM). 마감월/미마감월 모두 가능
-- ─── 1~5: 이벤트 기반 롤링 계산 ───
so_combined AS (
    SELECT
        SO_ID, [Customer name], [Customer PO], [Item name], [OS name],
        CAST([Line item] AS INTEGER) AS [Line item],
        CAST([Item qty] AS REAL) AS [Item qty],
        ROUND(CAST([Sales amount] AS REAL)) AS [Sales amount KRW],
        Period, [AX Period], [Model code], Sector,
        [Business registration number], [Industry code],
        [Expected delivery date], '국내' AS 구분
    FROM so_domestic
    WHERE COALESCE(Status, '') NOT IN ('Cancelled', 'Hold')
      AND Period IS NOT NULL AND TRIM(Period) != ''
    UNION ALL
    SELECT
        SO_ID, [Customer name], [Customer PO], [Item name], [OS name],
        CAST([Line item] AS INTEGER),
        CAST([Item qty] AS REAL),
        ROUND(CAST([Sales amount KRW] AS REAL)),
        Period, [AX Period], [Model code], Sector,
        [Business registration number], [Industry code],
        [Expected delivery date], '해외'
    FROM so_export
    WHERE COALESCE(Status, '') NOT IN ('Cancelled', 'Hold')
      AND Period IS NOT NULL AND TRIM(Period) != ''
),
dn_combined AS (
    SELECT SO_ID, CAST([Line item] AS INTEGER) AS [Line item],
        CAST(Qty AS REAL) AS Qty, ROUND(CAST([Total Sales] AS REAL)) AS 출고금액,
        SUBSTR([출고일], 1, 7) AS 출고월,
        [Customer name] AS dn_cust, [Item] AS dn_item,
        [Customer PO] AS dn_po, [Business registration number] AS dn_brn,
        '국내' AS dn_market
    FROM dn_domestic
    WHERE [출고일] IS NOT NULL AND TRIM(COALESCE([출고일], '')) != ''
    UNION ALL
    SELECT SO_ID, CAST([Line item] AS INTEGER),
        CAST(Qty AS REAL), ROUND(CAST([Total Sales KRW] AS REAL)),
        SUBSTR([선적일], 1, 7),
        [Customer name], [Item], [Customer PO], '', '해외'
    FROM dn_export
    WHERE [선적일] IS NOT NULL AND TRIM(COALESCE([선적일], '')) != ''
),
dn_by_month AS (
    SELECT SO_ID, [Line item], 출고월,
        SUM(Qty) AS Output_qty, SUM(출고금액) AS Output_amount,
        MIN(dn_cust) AS dn_cust, MIN(dn_item) AS dn_item,
        MIN(dn_po) AS dn_po, MIN(dn_brn) AS dn_brn, MIN(dn_market) AS dn_market
    FROM dn_combined WHERE 출고월 IS NOT NULL AND 출고월 != ''
    GROUP BY SO_ID, [Line item], 출고월
),
events_line_item AS (
    SELECT s.SO_ID, s.[Customer name], s.[Customer PO], s.[Item name],
        s.[OS name], s.[Line item], s.[Item qty], s.[Sales amount KRW],
        s.Period AS 등록Period, s.[AX Period], s.[Model code],
        s.Sector, s.[Business registration number], s.[Industry code],
        s.[Expected delivery date], s.구분,
        s.Period AS event_period,
        s.[Item qty] AS Value_Input_qty, s.[Sales amount KRW] AS Value_Input_amount,
        0 AS Value_Output_qty, 0 AS Value_Output_amount
    FROM so_combined s
    UNION ALL
    SELECT dm.SO_ID,
        COALESCE(s.[Customer name], NULLIF(dm.dn_cust, '0'), 'UNKNOWN') AS [Customer name],
        COALESCE(s.[Customer PO], NULLIF(dm.dn_po, '0'), '')    AS [Customer PO],
        COALESCE(s.[Item name], dm.dn_item, '')                 AS [Item name],
        COALESCE(s.[OS name], dm.dn_item, 'UNKNOWN')            AS [OS name],
        dm.[Line item],
        COALESCE(s.[Item qty], 0)               AS [Item qty],
        COALESCE(s.[Sales amount KRW], 0)       AS [Sales amount KRW],
        COALESCE(s.Period, '')  AS 등록Period,
        COALESCE(s.[AX Period], '')             AS [AX Period],
        COALESCE(s.[Model code], '')            AS [Model code],
        COALESCE(s.Sector, '')                  AS Sector,
        COALESCE(s.[Business registration number], NULLIF(dm.dn_brn, '0'), '') AS [Business registration number],
        COALESCE(s.[Industry code], '')         AS [Industry code],
        COALESCE(s.[Expected delivery date], '') AS [Expected delivery date],
        COALESCE(s.구분, dm.dn_market, '')      AS 구분,
        dm.출고월, 0, 0, dm.Output_qty, dm.Output_amount
    FROM dn_by_month dm
    LEFT JOIN so_combined s ON dm.SO_ID = s.SO_ID AND dm.[Line item] = s.[Line item]
),
os_grouped AS (
    SELECT SO_ID, [OS name], [Expected delivery date], event_period AS Period,
        MIN([Customer name]) AS [Customer name], MIN([Customer PO]) AS [Customer PO],
        MIN([Item name]) AS [Item name], MIN(구분) AS 구분, MIN(등록Period) AS 등록Period,
        MIN(Sector) AS Sector,
        MIN([Business registration number]) AS [Business registration number],
        MIN([Industry code]) AS [Industry code],
        GROUP_CONCAT(DISTINCT [AX Period]) AS [AX Period],
        GROUP_CONCAT(DISTINCT [Model code]) AS [Model code],
        SUM(Value_Input_qty) AS Value_Input_qty, SUM(Value_Input_amount) AS Value_Input_amount,
        SUM(Value_Output_qty) AS Value_Output_qty, SUM(Value_Output_amount) AS Value_Output_amount
    FROM events_line_item
    GROUP BY SO_ID, [OS name], [Expected delivery date], event_period
),
rolling AS (
    SELECT *,
        COALESCE(SUM(Value_Input_qty - Value_Output_qty) OVER w_prev, 0) AS Value_Start_qty,
        SUM(Value_Input_qty - Value_Output_qty) OVER w_curr AS Value_Ending_qty,
        COALESCE(SUM(Value_Input_amount - Value_Output_amount) OVER w_prev, 0) AS Value_Start_amount,
        SUM(Value_Input_amount - Value_Output_amount) OVER w_curr AS Value_Ending_amount
    FROM os_grouped
    WINDOW
        w_prev AS (PARTITION BY SO_ID, [OS name], [Expected delivery date] ORDER BY Period ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING),
        w_curr AS (PARTITION BY SO_ID, [OS name], [Expected delivery date] ORDER BY Period ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)
),

-- ─── 6. 스냅샷 메타 (마감된 period 목록) ───
active_snapshots AS (
    SELECT period FROM ob_snapshot_meta WHERE is_active = 1
),
last_snapshot AS (
    SELECT MAX(period) AS last_period FROM active_snapshots
),
next_open_period AS (
    SELECT
        CASE WHEN CAST(SUBSTR(lp.last_period, 6, 2) AS INTEGER) = 12
            THEN CAST(CAST(SUBSTR(lp.last_period, 1, 4) AS INTEGER) + 1 AS TEXT) || '-01'
            ELSE SUBSTR(lp.last_period, 1, 5) || PRINTF('%02d', CAST(SUBSTR(lp.last_period, 6, 2) AS INTEGER) + 1)
        END AS period
    FROM last_snapshot lp
    WHERE lp.last_period IS NOT NULL
),

-- ─── 7. 스냅샷 시점 누적 Ending 재계산 (correlated subquery 대체) ───
cumul_at_snapshot AS (
    SELECT
        SO_ID, [OS name], [Expected delivery date],
        SUM(Value_Input_qty - Value_Output_qty) AS recalc_ending_qty,
        SUM(Value_Input_amount - Value_Output_amount) AS recalc_ending_amount
    FROM os_grouped
    WHERE Period <= (SELECT last_period FROM last_snapshot)
    GROUP BY SO_ID, [OS name], [Expected delivery date]
),

-- ─── 8. Open Period (스냅샷 이후) — Variance 반영 롤링 ───
open_periods AS (
    SELECT
        r.Period,
        r.등록Period,
        r.구분,
        r.SO_ID,
        r.[Customer name],
        r.[Customer PO],
        r.[Item name],
        r.[OS name],
        r.[Expected delivery date],
        r.[AX Period],
        r.[Model code],
        r.Sector,
        r.[Business registration number],
        r.[Industry code],
        -- Start = 이전 스냅샷 ending (있으면), 없으면 롤링 Start
        COALESCE(snap.ending_qty, r.Value_Start_qty) AS Value_Start_qty,
        r.Value_Input_qty,
        r.Value_Output_qty,
        -- Variance: 첫 open period에만 표시
        CASE WHEN r.Period = (SELECT period FROM next_open_period)
        THEN COALESCE(c2.recalc_ending_qty - COALESCE(snap.ending_qty, 0), 0)
        ELSE 0
        END AS Value_Variance_qty,
        -- Ending = Start + Input + Variance - Output
        COALESCE(snap.ending_qty, r.Value_Start_qty)
            + r.Value_Input_qty
            + CASE WHEN r.Period = (SELECT period FROM next_open_period)
              THEN COALESCE(c2.recalc_ending_qty - COALESCE(snap.ending_qty, 0), 0)
              ELSE 0 END
            - r.Value_Output_qty
        AS Value_Ending_qty,
        -- Amount
        COALESCE(snap.ending_amount, r.Value_Start_amount) AS Value_Start_amount,
        r.Value_Input_amount,
        r.Value_Output_amount,
        CASE WHEN r.Period = (SELECT period FROM next_open_period)
        THEN COALESCE(c2.recalc_ending_amount - COALESCE(snap.ending_amount, 0), 0)
        ELSE 0
        END AS Value_Variance_amount,
        COALESCE(snap.ending_amount, r.Value_Start_amount)
            + r.Value_Input_amount
            + CASE WHEN r.Period = (SELECT period FROM next_open_period)
              THEN COALESCE(c2.recalc_ending_amount - COALESCE(snap.ending_amount, 0), 0)
              ELSE 0 END
            - r.Value_Output_amount
        AS Value_Ending_amount
    FROM rolling r
    LEFT JOIN ob_snapshot snap
        ON snap.snapshot_period = (SELECT last_period FROM last_snapshot)
       AND snap.SO_ID = r.SO_ID
       AND snap.[OS name] = r.[OS name]
       AND snap.[Expected delivery date] = COALESCE(r.[Expected delivery date], '')
    LEFT JOIN cumul_at_snapshot c2
        ON c2.SO_ID = r.SO_ID
       AND c2.[OS name] = r.[OS name]
       AND COALESCE(c2.[Expected delivery date], '') = COALESCE(r.[Expected delivery date], '')
    WHERE r.Period = (SELECT period FROM params)
      AND r.Period NOT IN (SELECT period FROM active_snapshots)
),

-- ─── 9. 마감(closed) Period — ob_snapshot 동결값 (무활동 그룹 포함 전 그룹) ───
closed_periods AS (
    SELECT
        s.snapshot_period AS Period,
        s.등록Period,
        s.구분,
        s.SO_ID,
        s.customer_name   AS [Customer name],
        ''                AS [Customer PO],
        s.item_name       AS [Item name],
        s.[OS name],
        s.[Expected delivery date],
        s.[AX Period],
        s.[Model code],
        s.Sector,
        ''                AS [Business registration number],
        ''                AS [Industry code],
        s.start_qty       AS Value_Start_qty,
        s.input_qty       AS Value_Input_qty,
        s.output_qty      AS Value_Output_qty,
        s.variance_qty    AS Value_Variance_qty,
        s.ending_qty      AS Value_Ending_qty,
        s.start_amount    AS Value_Start_amount,
        s.input_amount    AS Value_Input_amount,
        s.output_amount   AS Value_Output_amount,
        s.variance_amount AS Value_Variance_amount,
        s.ending_amount   AS Value_Ending_amount
    FROM ob_snapshot s, params p
    WHERE s.snapshot_period = p.period
      AND p.period IN (SELECT period FROM active_snapshots)
)

-- ═══ 최종: params.period가 마감월이면 동결 스냅샷, 미마감월이면 라이브 롤링 ═══
SELECT * FROM closed_periods
UNION ALL
SELECT * FROM open_periods
ORDER BY Period DESC, 구분, SO_ID, [OS name];
