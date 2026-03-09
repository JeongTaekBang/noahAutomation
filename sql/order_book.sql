-- ═══════════════════════════════════════════════════════════════
-- Order Book (수주잔고 롤링 원장)
-- Power Query M 코드 → SQLite 변환
-- DB Browser for SQLite > Execute SQL 탭에서 실행
-- ═══════════════════════════════════════════════════════════════
-- 동작: SO(수주) × Period(월) 확장 → Input/Output 계산 → 롤링 잔고
-- 전제: sync_db.py로 동기화된 noah_data.db 사용

WITH RECURSIVE
-- ─── 1. SO 통합 (국내 + 해외, Cancelled·빈 Period 제외) ───
so_combined AS (
    SELECT
        SO_ID,
        [Customer name],
        [Customer PO],
        [Item name],
        [OS name],
        CAST([Line item] AS INTEGER) AS [Line item],
        CAST([Item qty] AS REAL)     AS [Item qty],
        CAST([Sales amount] AS REAL) AS [Sales amount KRW],
        Period,
        [AX Period],
        [AX Project number],
        Sector,
        [Business registration number],
        [Industry code],
        [Expected delivery date],
        '국내' AS 구분
    FROM so_domestic
    WHERE COALESCE(Status, '') != 'Cancelled'
      AND Period IS NOT NULL AND TRIM(Period) != ''

    UNION ALL

    SELECT
        SO_ID,
        [Customer name],
        [Customer PO],
        [Item name],
        [OS name],
        CAST([Line item] AS INTEGER),
        CAST([Item qty] AS REAL),
        CAST([Sales amount KRW] AS REAL),
        Period,
        [AX Period],
        [AX Project number],
        Sector,
        [Business registration number],
        [Industry code],
        [Expected delivery date],
        '해외'
    FROM so_export
    WHERE COALESCE(Status, '') != 'Cancelled'
      AND Period IS NOT NULL AND TRIM(Period) != ''
),

-- ─── 2. DN 통합 (출고월 계산: 국내=출고일, 해외=선적일) ───
dn_combined AS (
    SELECT
        SO_ID,
        CAST([Line item] AS INTEGER) AS [Line item],
        CAST(Qty AS REAL)            AS Qty,
        CAST([Total Sales] AS REAL)  AS 출고금액,
        SUBSTR([출고일], 1, 7)        AS 출고월
    FROM dn_domestic
    WHERE [출고일] IS NOT NULL AND TRIM(COALESCE([출고일], '')) != ''

    UNION ALL

    SELECT
        SO_ID,
        CAST([Line item] AS INTEGER),
        CAST(Qty AS REAL),
        CAST([Total Sales KRW] AS REAL),
        SUBSTR([선적일], 1, 7)
    FROM dn_export
    WHERE [선적일] IS NOT NULL AND TRIM(COALESCE([선적일], '')) != ''
),

-- ─── 3. DN 월별 집계 (분할 출고 대응) ───
dn_by_month AS (
    SELECT SO_ID, [Line item], 출고월,
           SUM(Qty)    AS Output_qty,
           SUM(출고금액) AS Output_amount
    FROM dn_combined
    WHERE 출고월 IS NOT NULL AND 출고월 != ''
    GROUP BY SO_ID, [Line item], 출고월
),

-- ─── 4. DN 마지막 출고월 (Period 확장 끝점 결정) ───
dn_last_month AS (
    SELECT SO_ID, [Line item], MAX(출고월) AS last_출고월
    FROM dn_combined
    WHERE 출고월 IS NOT NULL AND 출고월 != ''
    GROUP BY SO_ID, [Line item]
),

-- ─── 5. 전체 기간 범위 ───
period_bounds AS (
    SELECT MIN(p) AS min_period, MAX(p) AS max_period
    FROM (
        SELECT Period AS p FROM so_combined
        UNION
        SELECT 출고월 FROM dn_by_month WHERE 출고월 IS NOT NULL
    )
),

-- ─── 6. 연속 월 생성 (min~max, 재귀 CTE) ───
month_series(m) AS (
    SELECT min_period FROM period_bounds
    UNION ALL
    SELECT
        CASE
            WHEN CAST(SUBSTR(m, 6, 2) AS INTEGER) = 12
            THEN CAST(CAST(SUBSTR(m, 1, 4) AS INTEGER) + 1 AS TEXT) || '-01'
            ELSE SUBSTR(m, 1, 5) || PRINTF('%02d', CAST(SUBSTR(m, 6, 2) AS INTEGER) + 1)
        END
    FROM month_series, period_bounds
    WHERE m < max_period
),

-- ─── 7. SO + 마지막출고월 JOIN ───
so_with_dn AS (
    SELECT s.*, d.last_출고월
    FROM so_combined s
    LEFT JOIN dn_last_month d
        ON s.SO_ID = d.SO_ID AND s.[Line item] = d.[Line item]
),

-- ─── 8. SO × Period 확장 (등록월 ~ 출고월/현재월) ───
--   출고 완료: 마지막 출고월에서 끊음
--   미출고:    max_period까지 (Backlog)
so_expanded AS (
    SELECT
        s.SO_ID, s.[Customer name], s.[Customer PO], s.[Item name],
        s.[OS name], s.[Line item], s.[Item qty], s.[Sales amount KRW],
        s.Period AS 등록Period, s.[AX Period], s.[AX Project number],
        s.Sector, s.[Business registration number], s.[Industry code],
        s.[Expected delivery date], s.구분,
        ms.m AS Period
    FROM so_with_dn s
    JOIN month_series ms
        ON ms.m >= s.Period
       AND ms.m <= MAX(
            COALESCE(s.last_출고월, (SELECT max_period FROM period_bounds)),
            s.Period
       )
),

-- ─── 9. Input 계산 (등록Period에만 수주 금액 기록) ───
with_input AS (
    SELECT *,
        CASE WHEN Period = 등록Period THEN [Item qty]          ELSE 0 END AS Value_Input_qty,
        CASE WHEN Period = 등록Period THEN [Sales amount KRW]  ELSE 0 END AS Value_Input_amount
    FROM so_expanded
),

-- ─── 10. Output 조인 (DN 출고를 해당 월에 매칭) ───
with_output AS (
    SELECT
        wi.*,
        COALESCE(dm.Output_qty, 0)    AS Value_Output_qty,
        COALESCE(dm.Output_amount, 0) AS Value_Output_amount
    FROM with_input wi
    LEFT JOIN dn_by_month dm
        ON wi.SO_ID = dm.SO_ID
       AND wi.[Line item] = dm.[Line item]
       AND wi.Period = dm.출고월
),

-- ─── 11. OS name 그룹화 (같은 제품+납기일 합산) ───
os_grouped AS (
    SELECT
        SO_ID, [OS name], [Expected delivery date], Period,
        MIN([Customer name])  AS [Customer name],
        MIN([Customer PO])    AS [Customer PO],
        MIN([Item name])      AS [Item name],
        MIN(구분)              AS 구분,
        MIN(등록Period)        AS 등록Period,
        MIN(Sector)           AS Sector,
        MIN([Business registration number]) AS [Business registration number],
        MIN([Industry code])  AS [Industry code],
        GROUP_CONCAT(DISTINCT [AX Period])         AS [AX Period],
        GROUP_CONCAT(DISTINCT [AX Project number]) AS [AX Project number],
        SUM(Value_Input_qty)     AS Value_Input_qty,
        SUM(Value_Input_amount)  AS Value_Input_amount,
        SUM(Value_Output_qty)    AS Value_Output_qty,
        SUM(Value_Output_amount) AS Value_Output_amount
    FROM with_output
    GROUP BY SO_ID, [OS name], [Expected delivery date], Period
)

-- ─── 12. 롤링 계산 (Window function: Start/Ending 전파) ───
SELECT
    Period,
    등록Period,
    구분,
    SO_ID,
    [Customer name],
    [Customer PO],
    [Item name],
    [OS name],
    [Expected delivery date],
    [AX Period],
    [AX Project number],
    Sector,
    [Business registration number],
    [Industry code],
    -- Start = 이전 Period까지의 누적 (Input - Output)
    COALESCE(SUM(Value_Input_qty - Value_Output_qty) OVER (
        PARTITION BY SO_ID, [OS name], [Expected delivery date]
        ORDER BY Period ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING
    ), 0) AS Value_Start_qty,
    Value_Input_qty,
    Value_Output_qty,
    0 AS Value_Variance_qty,
    -- Ending = 현재까지의 누적
    SUM(Value_Input_qty - Value_Output_qty) OVER (
        PARTITION BY SO_ID, [OS name], [Expected delivery date]
        ORDER BY Period ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
    ) AS Value_Ending_qty,
    COALESCE(SUM(Value_Input_amount - Value_Output_amount) OVER (
        PARTITION BY SO_ID, [OS name], [Expected delivery date]
        ORDER BY Period ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING
    ), 0) AS Value_Start_amount,
    Value_Input_amount,
    Value_Output_amount,
    0 AS Value_Variance_amount,
    SUM(Value_Input_amount - Value_Output_amount) OVER (
        PARTITION BY SO_ID, [OS name], [Expected delivery date]
        ORDER BY Period ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
    ) AS Value_Ending_amount
FROM os_grouped
ORDER BY Period DESC, 구분, SO_ID, [OS name];
