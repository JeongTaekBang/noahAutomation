-- ═══════════════════════════════════════════════════════════════
-- 현재 Backlog 현황 (Order Book 요약)
-- 전체 이벤트 합산 기준, Ending > 0인 건만 표시
-- ═══════════════════════════════════════════════════════════════
-- 이벤트 기반: 전체 Input/Output 합산으로 잔고 계산
-- 빈 월 확장 없이 직접 합산하므로 고성능

WITH
so_combined AS (
    SELECT SO_ID, [Customer name], [Customer PO], [Item name], [OS name],
        CAST([Line item] AS INTEGER) AS [Line item],
        CAST([Item qty] AS REAL) AS [Item qty],
        CAST([Sales amount] AS REAL) AS [Sales amount KRW],
        Period, [AX Period], [Model code], Sector,
        [Business registration number], [Industry code],
        [Expected delivery date], '국내' AS 구분
    FROM so_domestic
    WHERE COALESCE(Status, '') != 'Cancelled'
      AND Period IS NOT NULL AND TRIM(Period) != ''
    UNION ALL
    SELECT SO_ID, [Customer name], [Customer PO], [Item name], [OS name],
        CAST([Line item] AS INTEGER),
        CAST([Item qty] AS REAL),
        CAST([Sales amount KRW] AS REAL),
        Period, [AX Period], [Model code], Sector,
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
events_line_item AS (
    SELECT s.SO_ID, s.[Customer name], s.[Customer PO], s.[Item name],
        s.[OS name], s.[Line item],
        s.Period AS 등록Period, s.[AX Period], s.[Model code],
        s.Sector, s.[Business registration number], s.[Industry code],
        s.[Expected delivery date], s.구분,
        s.[Item qty] AS Value_Input_qty, s.[Sales amount KRW] AS Value_Input_amount,
        0 AS Value_Output_qty, 0 AS Value_Output_amount
    FROM so_combined s
    UNION ALL
    SELECT dm.SO_ID,
        COALESCE(s.[Customer name], 'UNKNOWN') AS [Customer name],
        COALESCE(s.[Customer PO], '')           AS [Customer PO],
        COALESCE(s.[Item name], '')             AS [Item name],
        COALESCE(s.[OS name], 'UNKNOWN')        AS [OS name],
        dm.[Line item],
        COALESCE(s.Period, '')                  AS Period,
        COALESCE(s.[AX Period], '')             AS [AX Period],
        COALESCE(s.[Model code], '')            AS [Model code],
        COALESCE(s.Sector, '')                  AS Sector,
        COALESCE(s.[Business registration number], '') AS [Business registration number],
        COALESCE(s.[Industry code], '')         AS [Industry code],
        COALESCE(s.[Expected delivery date], '') AS [Expected delivery date],
        COALESCE(s.구분, '')                    AS 구분,
        0, 0, dm.Output_qty, dm.Output_amount
    FROM dn_by_month dm
    LEFT JOIN so_combined s ON dm.SO_ID = s.SO_ID AND dm.[Line item] = s.[Line item]
),
-- ─── Backlog: 전체 이벤트 합산, Ending > 0 ───
backlog AS (
    SELECT
        SO_ID, [OS name], [Expected delivery date],
        MIN([Customer name]) AS [Customer name],
        MIN(구분) AS 구분,
        MIN(Sector) AS Sector,
        MIN([Industry code]) AS [Industry code],
        GROUP_CONCAT(DISTINCT [Model code]) AS [Model code],
        SUM(Value_Input_qty - Value_Output_qty) AS Value_Ending_qty,
        SUM(Value_Input_amount - Value_Output_amount) AS Value_Ending_amount
    FROM events_line_item
    GROUP BY SO_ID, [OS name], [Expected delivery date]
    HAVING SUM(Value_Input_amount - Value_Output_amount) > 0
)

-- ═══ Backlog 현황 ═══
SELECT
    구분,
    SO_ID,
    [Customer name],
    [OS name],
    [Expected delivery date] AS 납기일,
    CAST(Value_Ending_qty AS INTEGER) AS 잔여수량,
    PRINTF('%,.0f', Value_Ending_amount) AS 잔여금액,
    [Model code],
    Sector,
    [Industry code]
FROM backlog
ORDER BY 구분, SO_ID, [OS name];


-- ═══ Backlog 요약 (구분별 합계) ═══
/*
SELECT
    구분,
    COUNT(DISTINCT SO_ID) AS 주문건수,
    SUM(Value_Ending_qty) AS 총잔여수량,
    PRINTF('%,.0f', SUM(Value_Ending_amount)) AS 총잔여금액
FROM backlog
GROUP BY 구분;
*/
