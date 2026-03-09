-- ═══════════════════════════════════════════════════════════════
-- 현재 Backlog 현황 (스냅샷 기반 Order Book 요약)
-- 마지막 Period 기준, Ending > 0인 건만 표시
-- ═══════════════════════════════════════════════════════════════
-- 사용법: order_book_snapshot.sql과 동일한 구조, Backlog 필터 적용
-- 스냅샷이 있으면 Variance 반영, 없으면 롤링 fallback

WITH RECURSIVE
so_combined AS (
    SELECT
        SO_ID, [Customer name], [Customer PO], [Item name], [OS name],
        CAST([Line item] AS INTEGER) AS [Line item],
        CAST([Item qty] AS REAL) AS [Item qty],
        CAST([Sales amount] AS REAL) AS [Sales amount KRW],
        Period, [AX Period], [AX Project number], Sector,
        [Business registration number], [Industry code],
        [Expected delivery date], '국내' AS 구분
    FROM so_domestic
    WHERE COALESCE(Status, '') != 'Cancelled'
      AND Period IS NOT NULL AND TRIM(Period) != ''
    UNION ALL
    SELECT
        SO_ID, [Customer name], [Customer PO], [Item name], [OS name],
        CAST([Line item] AS INTEGER),
        CAST([Item qty] AS REAL),
        CAST([Sales amount KRW] AS REAL),
        Period, [AX Period], [AX Project number], Sector,
        [Business registration number], [Industry code],
        [Expected delivery date], '해외'
    FROM so_export
    WHERE COALESCE(Status, '') != 'Cancelled'
      AND Period IS NOT NULL AND TRIM(Period) != ''
),
dn_combined AS (
    SELECT SO_ID, CAST([Line item] AS INTEGER) AS [Line item],
        CAST(Qty AS REAL) AS Qty, CAST([Total Sales] AS REAL) AS 출고금액,
        SUBSTR([출고일], 1, 7) AS 출고월
    FROM dn_domestic
    WHERE [출고일] IS NOT NULL AND TRIM(COALESCE([출고일], '')) != ''
    UNION ALL
    SELECT SO_ID, CAST([Line item] AS INTEGER),
        CAST(Qty AS REAL), CAST([Total Sales KRW] AS REAL),
        SUBSTR([선적일], 1, 7)
    FROM dn_export
    WHERE [선적일] IS NOT NULL AND TRIM(COALESCE([선적일], '')) != ''
),
dn_by_month AS (
    SELECT SO_ID, [Line item], 출고월,
        SUM(Qty) AS Output_qty, SUM(출고금액) AS Output_amount
    FROM dn_combined WHERE 출고월 IS NOT NULL AND 출고월 != ''
    GROUP BY SO_ID, [Line item], 출고월
),
dn_last_month AS (
    SELECT SO_ID, [Line item], MAX(출고월) AS last_출고월
    FROM dn_combined WHERE 출고월 IS NOT NULL AND 출고월 != ''
    GROUP BY SO_ID, [Line item]
),
period_bounds AS (
    SELECT MIN(p) AS min_period, MAX(p) AS max_period
    FROM (SELECT Period AS p FROM so_combined UNION SELECT 출고월 FROM dn_by_month WHERE 출고월 IS NOT NULL)
),
month_series(m) AS (
    SELECT min_period FROM period_bounds
    UNION ALL
    SELECT CASE WHEN CAST(SUBSTR(m, 6, 2) AS INTEGER) = 12
        THEN CAST(CAST(SUBSTR(m, 1, 4) AS INTEGER) + 1 AS TEXT) || '-01'
        ELSE SUBSTR(m, 1, 5) || PRINTF('%02d', CAST(SUBSTR(m, 6, 2) AS INTEGER) + 1) END
    FROM month_series, period_bounds WHERE m < max_period
),
so_with_dn AS (
    SELECT s.*, d.last_출고월
    FROM so_combined s LEFT JOIN dn_last_month d ON s.SO_ID = d.SO_ID AND s.[Line item] = d.[Line item]
),
so_expanded AS (
    SELECT s.SO_ID, s.[Customer name], s.[Customer PO], s.[Item name], s.[OS name],
        s.[Line item], s.[Item qty], s.[Sales amount KRW],
        s.Period AS 등록Period, s.[AX Period], s.[AX Project number],
        s.Sector, s.[Business registration number], s.[Industry code],
        s.[Expected delivery date], s.구분, ms.m AS Period
    FROM so_with_dn s
    JOIN month_series ms ON ms.m >= s.Period
        AND ms.m <= MAX(COALESCE(s.last_출고월, (SELECT max_period FROM period_bounds)), s.Period)
),
with_input AS (
    SELECT *,
        CASE WHEN Period = 등록Period THEN [Item qty] ELSE 0 END AS Value_Input_qty,
        CASE WHEN Period = 등록Period THEN [Sales amount KRW] ELSE 0 END AS Value_Input_amount
    FROM so_expanded
),
with_output AS (
    SELECT wi.*,
        COALESCE(dm.Output_qty, 0) AS Value_Output_qty,
        COALESCE(dm.Output_amount, 0) AS Value_Output_amount
    FROM with_input wi
    LEFT JOIN dn_by_month dm ON wi.SO_ID = dm.SO_ID AND wi.[Line item] = dm.[Line item] AND wi.Period = dm.출고월
),
os_grouped AS (
    SELECT SO_ID, [OS name], [Expected delivery date], Period,
        MIN([Customer name]) AS [Customer name], MIN([Customer PO]) AS [Customer PO],
        MIN([Item name]) AS [Item name], MIN(구분) AS 구분, MIN(등록Period) AS 등록Period,
        MIN(Sector) AS Sector,
        MIN([Business registration number]) AS [Business registration number],
        MIN([Industry code]) AS [Industry code],
        GROUP_CONCAT(DISTINCT [AX Period]) AS [AX Period],
        GROUP_CONCAT(DISTINCT [AX Project number]) AS [AX Project number],
        SUM(Value_Input_qty) AS Value_Input_qty, SUM(Value_Input_amount) AS Value_Input_amount,
        SUM(Value_Output_qty) AS Value_Output_qty, SUM(Value_Output_amount) AS Value_Output_amount
    FROM with_output
    GROUP BY SO_ID, [OS name], [Expected delivery date], Period
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
last_snapshot AS (
    SELECT MAX(period) AS last_period FROM ob_snapshot_meta WHERE is_active = 1
),
last_p AS (SELECT MAX(Period) AS lp FROM rolling),

-- Backlog: 스냅샷 기반 Ending 또는 롤링 Ending
backlog AS (
    SELECT
        r.구분,
        r.SO_ID,
        r.[Customer name],
        r.[OS name],
        r.[Expected delivery date] AS 납기일,
        -- Ending: 스냅샷 이후면 Variance 반영, 아니면 롤링
        CASE
            WHEN (SELECT last_period FROM last_snapshot) IS NOT NULL THEN
                COALESCE(snap.ending_qty, r.Value_Start_qty)
                + r.Value_Input_qty
                - r.Value_Output_qty
                + COALESCE(
                    (SELECT rl.Value_Ending_qty FROM rolling rl
                     WHERE rl.SO_ID = r.SO_ID AND rl.[OS name] = r.[OS name]
                       AND COALESCE(rl.[Expected delivery date], '') = COALESCE(r.[Expected delivery date], '')
                       AND rl.Period = (SELECT last_period FROM last_snapshot))
                    - COALESCE(snap.ending_qty, 0), 0)
            ELSE r.Value_Ending_qty
        END AS 잔여수량_raw,
        CASE
            WHEN (SELECT last_period FROM last_snapshot) IS NOT NULL THEN
                COALESCE(snap.ending_amount, r.Value_Start_amount)
                + r.Value_Input_amount
                - r.Value_Output_amount
                + COALESCE(
                    (SELECT rl.Value_Ending_amount FROM rolling rl
                     WHERE rl.SO_ID = r.SO_ID AND rl.[OS name] = r.[OS name]
                       AND COALESCE(rl.[Expected delivery date], '') = COALESCE(r.[Expected delivery date], '')
                       AND rl.Period = (SELECT last_period FROM last_snapshot))
                    - COALESCE(snap.ending_amount, 0), 0)
            ELSE r.Value_Ending_amount
        END AS 잔여금액_raw,
        r.[AX Project number],
        r.Sector,
        r.[Industry code]
    FROM rolling r, last_p
    LEFT JOIN ob_snapshot snap
        ON snap.snapshot_period = (SELECT last_period FROM last_snapshot)
       AND snap.SO_ID = r.SO_ID
       AND snap.[OS name] = r.[OS name]
       AND snap.[Expected delivery date] = COALESCE(r.[Expected delivery date], '')
    WHERE r.Period = last_p.lp
)

-- ═══ Backlog 현황: Ending > 0 ═══
SELECT
    구분,
    SO_ID,
    [Customer name],
    [OS name],
    납기일,
    CAST(잔여수량_raw AS INTEGER) AS 잔여수량,
    PRINTF('%,.0f', 잔여금액_raw) AS 잔여금액,
    [AX Project number],
    Sector,
    [Industry code]
FROM backlog
WHERE 잔여금액_raw > 0
ORDER BY 구분, SO_ID, [OS name];
