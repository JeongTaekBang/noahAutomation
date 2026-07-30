-- ═══════════════════════════════════════════════════════════════
-- 현재 Backlog 현황 (스냅샷 기반 Order Book 요약)
-- 전체 이벤트 합산 기준, Ending > 0인 건만 표시
-- ═══════════════════════════════════════════════════════════════
-- 이벤트 기반: 전체 Input/Output 합산으로 잔고 계산
-- 스냅샷 유무와 무관하게 최종 Ending은 동일
-- (snap_ending + variance + post_events = total_events)

WITH
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
-- DN 월별 집계는 v_dn_by_month 뷰 (매출 인식 필터 + 분할 출고 합산의 단일 정의)
events_line_item AS (
    SELECT s.SO_ID, s.[Customer name], s.[Customer PO], s.[Item name],
        s.[OS name], s.[Line item], s.[Model code],
        s.Sector, s.[Business registration number], s.[Industry code],
        s.[Expected delivery date], s.구분,
        s.[Item qty] AS Value_Input_qty, s.[Sales amount KRW] AS Value_Input_amount,
        0 AS Value_Output_qty, 0 AS Value_Output_amount, 0 AS Value_FX_amount
    FROM so_combined s
    UNION ALL
    SELECT dm.SO_ID,
        COALESCE(s.[Customer name], 'UNKNOWN') AS [Customer name],
        COALESCE(s.[Customer PO], '')           AS [Customer PO],
        COALESCE(s.[Item name], '')             AS [Item name],
        COALESCE(s.[OS name], 'UNKNOWN')        AS [OS name],
        dm.[Line item],
        COALESCE(s.[Model code], '')            AS [Model code],
        COALESCE(s.Sector, '')                  AS Sector,
        COALESCE(s.[Business registration number], '') AS [Business registration number],
        COALESCE(s.[Industry code], '')         AS [Industry code],
        COALESCE(s.[Expected delivery date], '') AS [Expected delivery date],
        COALESCE(s.구분, '')                    AS 구분,
        0, 0, dm.Output_qty, dm.Output_amount, dm.Output_fx
    FROM v_dn_by_month dm
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
        SUM(Value_Input_amount - Value_Output_amount + Value_FX_amount) AS Value_Ending_amount
    FROM events_line_item
    GROUP BY SO_ID, [OS name], [Expected delivery date]
    -- 잔여수량 또는 잔여금액 중 하나라도 남으면 open (수량 양수/금액 0 라인 보존)
    HAVING ABS(SUM(Value_Input_qty - Value_Output_qty)) > 0.001
        OR SUM(Value_Input_amount - Value_Output_amount + Value_FX_amount) > 0.5
)

-- ═══ Backlog 현황: Ending > 0 ═══
SELECT
    구분,
    SO_ID,
    [Customer name],
    [OS name],
    [Expected delivery date] AS 납기일,
    CAST(ROUND(Value_Ending_qty) AS INTEGER) AS 잔여수량,
    PRINTF('%,.0f', Value_Ending_amount) AS 잔여금액,
    [Model code],
    Sector,
    [Industry code]
FROM backlog
ORDER BY 구분, SO_ID, [OS name];
