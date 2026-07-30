-- ═══════════════════════════════════════════════════════════════
-- Order Book (수주잔고 이벤트 기반 원장)
-- 이벤트 기반: SO 등록(Input), DN 매출(Output) 발생 월만 행 생성
-- DB Browser for SQLite > Execute SQL 탭에서 실행
-- ═══════════════════════════════════════════════════════════════
-- 동작: SO(수주) Input + DN(매출) Output 이벤트 → 롤링 잔고 계산
-- 빈 월(활동 없는 월)은 행을 생성하지 않음
-- 전제: sync_db.py로 동기화된 noah_data.db 사용 (fx 테이블 + v_dn_revenue 뷰 포함)
--
-- Output 귀속월/금액은 `v_dn_revenue` 뷰가 단일 정의한다 (db_schema.py):
--   국내 = 세금계산서 발행월 (N/A=발행불필요면 출고월, 선수금+출고면 출고월)
--   해외 = 선적월, KRW는 선적월 환율로 재환산
-- Variance = 환율 재평가분(재환산액 − 시트 KRW) → Ending은 재환산 전과 동일하게 유지된다.
-- Period 필터 → SUM(Value_Output_amount)는 Excel `AX_매출대사` 같은 월 합계와 일치한다.

WITH
-- ─── 1. SO 통합 (국내 + 해외, Cancelled·Hold·빈 Period 제외) ───
so_combined AS (
    SELECT
        SO_ID,
        [Customer name],
        [Customer PO],
        [Item name],
        [OS name],
        CAST([Line item] AS INTEGER) AS [Line item],
        CAST([Item qty] AS REAL)     AS [Item qty],
        ROUND(CAST([Sales amount] AS REAL)) AS [Sales amount KRW],
        Period,
        [AX Period],
        [Model code],
        Sector,
        [Business registration number],
        [Industry code],
        COALESCE([Expected delivery date], '') AS [Expected delivery date],
        '국내' AS 구분
    FROM so_domestic
    WHERE COALESCE(Status, '') NOT IN ('Cancelled', 'Hold')
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
        ROUND(CAST([Sales amount KRW] AS REAL)),
        Period,
        [AX Period],
        [Model code],
        Sector,
        [Business registration number],
        [Industry code],
        COALESCE([Expected delivery date], '') AS [Expected delivery date],
        '해외'
    FROM so_export
    WHERE COALESCE(Status, '') NOT IN ('Cancelled', 'Hold')
      AND Period IS NOT NULL AND TRIM(Period) != ''
),

-- ─── 2. 이벤트 통합 (Input: SO 등록 + Output: DN 매출) ───
-- DN 월별 집계는 v_dn_by_month 뷰 (매출 인식 필터 + 분할 출고 합산의 단일 정의)
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
        0 AS Value_Output_amount,
        0 AS Value_Variance_amount
    FROM so_combined s

    UNION ALL

    -- Output: DN 매출 이벤트 (LEFT JOIN — 취소/누락 SO의 DN도 보존)
    SELECT
        dm.SO_ID,
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
        dm.매출월 AS event_period,
        0, 0,
        dm.Output_qty,
        dm.Output_amount,
        dm.Output_fx
    FROM v_dn_by_month dm
    LEFT JOIN so_combined s ON dm.SO_ID = s.SO_ID AND dm.[Line item] = s.[Line item]
),

-- ─── 3. OS name 그룹화 (같은 제품+납기일 합산) ───
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
        SUM(Value_Output_amount) AS Value_Output_amount,
        SUM(Value_Variance_amount) AS Value_Variance_amount
    FROM events_line_item
    GROUP BY SO_ID, [OS name], [Expected delivery date], event_period
)

-- ─── 4. 롤링 계산 (Window function: Start/Ending 전파) ───
-- 금액 Ending = 누적(Input − Output + Variance) — Variance가 환율차를 상쇄하므로
-- 재환산 도입 전과 같은 값이 나온다 (수량엔 환율 영향이 없어 Variance 없음)
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
    COALESCE(SUM(Value_Input_amount - Value_Output_amount + Value_Variance_amount) OVER (
        PARTITION BY SO_ID, [OS name], [Expected delivery date]
        ORDER BY Period ROWS BETWEEN UNBOUNDED PRECEDING AND 1 PRECEDING
    ), 0) AS Value_Start_amount,
    Value_Input_amount,
    Value_Output_amount,
    Value_Variance_amount,
    SUM(Value_Input_amount - Value_Output_amount + Value_Variance_amount) OVER (
        PARTITION BY SO_ID, [OS name], [Expected delivery date]
        ORDER BY Period ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
    ) AS Value_Ending_amount
FROM os_grouped
ORDER BY Period DESC, 구분, SO_ID, [OS name];
