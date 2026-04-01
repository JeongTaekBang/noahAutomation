-- ═══════════════════════════════════════════════════════════════
-- Order Book Variance 분석
-- 마감 스냅샷 간 소급 변경 내역 + 변동이유 자동 분류
-- ═══════════════════════════════════════════════════════════════
-- 사용법: params CTE의 period 값을 분석할 마감 기간으로 변경
-- 전제: close_period.py로 해당 period 마감 완료
--
-- 변동이유 분류:
--   환율차이    — 해외 건 Sales amount KRW 소급 변경 (수량 불변, 금액만 변동)
--   판매가변경  — 국내 건 Sales amount 소급 변경 (수량 불변, 금액만 변동)
--   수량변경    — SO 수량 소급 수정 또는 라인 추가/삭제
--   반올림      — KRW 환산 소수점 ±1원 이내
--
-- 제외: 납기변경 (EDD 수정은 그룹키 이동일 뿐 금액/수량 변동 아님)

WITH params AS (SELECT '2026-03' AS period),

-- ─── 1. Variance 발생 행 추출 ───
var_raw AS (
    SELECT
        s.SO_ID, s.[OS name], s.[Expected delivery date],
        s.구분, s.customer_name, s.등록Period,
        s.start_amount, s.input_amount, s.output_amount,
        s.variance_qty, s.variance_amount,
        s.ending_amount
    FROM ob_snapshot s, params p
    WHERE s.snapshot_period = p.period
      AND (ABS(s.variance_qty) > 0.001 OR ABS(s.variance_amount) > 0.5)
),

-- ─── 2. 납기변경 감지용 윈도우 계산 ───
-- 같은 SO_ID + OS name에 음/양 Variance가 동시 존재 → EDD 변경 (제외 대상)
var_with_flags AS (
    SELECT *,
        SUM(CASE WHEN variance_amount < -0.5 THEN 1 ELSE 0 END)
            OVER (PARTITION BY SO_ID, [OS name]) AS _neg_cnt,
        SUM(CASE WHEN variance_amount > 0.5 THEN 1 ELSE 0 END)
            OVER (PARTITION BY SO_ID, [OS name]) AS _pos_cnt
    FROM var_raw
),

-- ─── 3. 변동이유 분류 (납기변경 제외) ───
classified AS (
    SELECT
        SO_ID,
        [OS name],
        [Expected delivery date] AS 납기일,
        구분,
        customer_name AS 고객명,
        등록Period,
        variance_qty AS Var_수량,
        variance_amount AS Var_금액,
        start_amount AS Start_금액,
        ending_amount AS Ending_금액,
        CASE
            WHEN ABS(variance_amount) <= 1 AND ABS(variance_qty) <= 0.001
                THEN '반올림'
            WHEN 구분 = '해외' AND ABS(variance_qty) <= 0.001
                THEN '환율차이'
            WHEN 구분 = '국내' AND ABS(variance_qty) <= 0.001
                THEN '판매가변경'
            WHEN ABS(variance_qty) > 0.001
                THEN '수량변경'
            ELSE '금액변경'
        END AS 변동이유
    FROM var_with_flags
    WHERE NOT (_neg_cnt > 0 AND _pos_cnt > 0)  -- 납기변경 제외
)

-- ═══ 상세 내역 ═══
SELECT
    변동이유,
    SO_ID,
    [OS name],
    납기일,
    구분,
    고객명,
    CAST(Var_수량 AS INTEGER) AS Var_수량,
    PRINTF('%,.0f', Var_금액) AS Var_금액,
    PRINTF('%,.0f', Start_금액) AS Start,
    PRINTF('%,.0f', Ending_금액) AS Ending
FROM classified
ORDER BY
    CASE 변동이유
        WHEN '환율차이'   THEN 1
        WHEN '판매가변경' THEN 2
        WHEN '수량변경'   THEN 3
        WHEN '반올림'     THEN 4
        ELSE 5
    END,
    ABS(Var_금액) DESC;

-- ═══ 유형별 요약 ═══
/*
SELECT
    변동이유,
    COUNT(*) AS 건수,
    PRINTF('%,.0f', SUM(Var_금액)) AS 합계
FROM classified
GROUP BY 변동이유
ORDER BY ABS(SUM(Var_금액)) DESC;
*/
