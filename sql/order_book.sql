-- ═══════════════════════════════════════════════════════════════
-- Order Book (수주잔고 이벤트 기반 원장)
-- 이벤트 기반: SO 등록(Input), DN 출고(Output) 발생 월만 행 생성
-- DB Browser for SQLite > Execute SQL 탭에서 실행
-- ═══════════════════════════════════════════════════════════════
-- 동작: SO(수주) Input + DN(출고) Output 이벤트 → 롤링 잔고 계산
-- 빈 월(활동 없는 월)은 행을 생성하지 않음
-- 전제: sync_db.py로 동기화된 noah_data.db 사용

WITH
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
        [Model code],
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
        [Model code],
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

-- ─── 4. 이벤트 통합 (Input: SO 등록 + Output: DN 출고) ───
events_line_item AS (
    -- Input: SO 등록 이벤트
    SELECT
        s.SO_ID, s.[Customer name], s.[Customer PO], s.[Item name],
        s.[OS name], s.[Line item], s.[Item qty], s.[Sales amount KRW],
        s.Period AS 등록Period, s.[AX Period], s.[Model code],
        s.Sector, s.[Business registration number], s.[Industry code],
        s.[Expected delivery date], s.구분,
        s.Period AS event_period,
        s.[Item qty]          AS Value_Input_qty,
        s.[Sales amount KRW]  AS Value_Input_amount,
        0 AS Value_Output_qty,
        0 AS Value_Output_amount
    FROM so_combined s

    UNION ALL

    -- Output: DN 출고 이벤트 (LEFT JOIN — 취소/누락 SO의 DN도 보존)
    SELECT
        dm.SO_ID,
        COALESCE(s.[Customer name], 'UNKNOWN') AS [Customer name],
        COALESCE(s.[Customer PO], '')           AS [Customer PO],
        COALESCE(s.[Item name], '')             AS [Item name],
        COALESCE(s.[OS name], 'UNKNOWN')        AS [OS name],
        dm.[Line item],
        COALESCE(s.[Item qty], 0)               AS [Item qty],
        COALESCE(s.[Sales amount KRW], 0)       AS [Sales amount KRW],
        COALESCE(s.Period, '')  AS 등록Period,
        COALESCE(s.[AX Period], '')             AS [AX Period],
        COALESCE(s.[Model code], '')            AS [Model code],
        COALESCE(s.Sector, '')                  AS Sector,
        COALESCE(s.[Business registration number], '') AS [Business registration number],
        COALESCE(s.[Industry code], '')         AS [Industry code],
        COALESCE(s.[Expected delivery date], '') AS [Expected delivery date],
        COALESCE(s.구분, '')                    AS 구분,
        dm.출고월 AS event_period,
        0, 0,
        dm.Output_qty,
        dm.Output_amount
    FROM dn_by_month dm
    LEFT JOIN so_combined s ON dm.SO_ID = s.SO_ID AND dm.[Line item] = s.[Line item]
),

-- ─── 5. OS name 그룹화 (같은 제품+납기일 합산) ───
os_grouped AS (
    SELECT
        SO_ID, [OS name], [Expected delivery date], event_period AS Period,
        MIN([Customer name])  AS [Customer name],
        MIN([Customer PO])    AS [Customer PO],
        MIN([Item name])      AS [Item name],
        MIN(구분)              AS 구분,
        MIN(등록Period)        AS 등록Period,
        MIN(Sector)           AS Sector,
        MIN([Business registration number]) AS [Business registration number],
        MIN([Industry code])  AS [Industry code],
        GROUP_CONCAT(DISTINCT [AX Period])         AS [AX Period],
        GROUP_CONCAT(DISTINCT [Model code]) AS [Model code],
        SUM(Value_Input_qty)     AS Value_Input_qty,
        SUM(Value_Input_amount)  AS Value_Input_amount,
        SUM(Value_Output_qty)    AS Value_Output_qty,
        SUM(Value_Output_amount) AS Value_Output_amount
    FROM events_line_item
    GROUP BY SO_ID, [OS name], [Expected delivery date], event_period
)

-- ─── 6. 롤링 계산 (Window function: Start/Ending 전파) ───
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
    [Model code],
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
