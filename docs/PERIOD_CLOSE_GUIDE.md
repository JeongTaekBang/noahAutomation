# Order Book 월 마감 (Period Close) 운용 가이드

Order Book의 월별 수주잔고를 스냅샷으로 저장하여, Start를 고정하고 소급 변경분을 Variance로 자동 감지하는 기능.

---

## 개요

### 왜 필요한가?

기존 `order_book.sql`은 매번 SO/DN raw 데이터에서 롤링 재계산하므로, 과거 Period의 값이 데이터 수정에 따라 소급 변동됨. AX2009의 월별 마감과 같이 **스냅샷(고정값)**을 저장하면:

- **과거 데이터 확정**: 마감된 Period의 Start/Input/Output/Ending은 변하지 않음
- **소급 변경 감지**: 마감 후 원 데이터가 수정되면 Variance로 자동 검출
- **감사/보고**: 월말 기준 수주잔고를 확정된 숫자로 보고 가능

### 핵심 공식

```
Ending = Start(스냅샷) + Input + Variance - Output
```

| 항목 | 의미 |
|------|------|
| Start | 전월 마감 스냅샷의 Ending (고정) |
| Input | 당월 신규 수주 (등록Period = 당월) |
| Output | 당월 출고 (DN 출고월 = 당월) |
| Variance | 전월 이전 데이터 소급 변경분 (자동 계산) |
| Ending | 당월 말 수주잔고 |

---

## 시스템 도식

### 데이터 흐름 (전체)

```
NOAH_SO_PO_DN.xlsx  (원본 — 수동 입력)
        │
        │  sync_db.py
        ▼
┌─────────────────────────────────────────────────────┐
│  noah_data.db (SQLite)                              │
│                                                     │
│  ┌─── Raw 테이블 (sync_db가 관리) ───┐              │
│  │ so_domestic  │ so_export          │              │
│  │ dn_domestic  │ dn_export          │              │
│  │ po_domestic  │ po_export          │              │
│  │ pmt_domestic │                    │              │
│  └──────────────┴────────────────────┘              │
│        │                                            │
│        │  close_period.py (take_snapshot)            │
│        │  롤링 CTE 실행 → Variance 계산 → 저장       │
│        ▼                                            │
│  ┌─── 스냅샷 테이블 (close_period가 관리) ───┐       │
│  │ ob_snapshot       │ 마감 데이터 (고정값)  │       │
│  │ ob_snapshot_meta  │ 마감 메타 (일시/상태) │       │
│  └───────────────────┴───────────────────────┘       │
│        │                                            │
│        │  order_book_snapshot.sql (DB Browser 조회)   │
│        ▼                                            │
│  ┌─── 조회 결과 ─────────────────────────────┐       │
│  │ Open Period만 표시                         │       │
│  │ → Start=스냅샷Ending, Variance 반영        │       │
│  │ (마감 Period는 close_period.py --list)     │       │
│  └────────────────────────────────────────────┘       │
└─────────────────────────────────────────────────────┘
```

### 테이블 관계

```
ob_snapshot_meta (마감 메타)
┌──────────┬────────────────────┬──────────┐
│ period   │ closed_at          │ is_active│
│ (PK)     │                    │          │
├──────────┼────────────────────┼──────────┤
│ 2026-01  │ 2026-02-03T09:00   │ 1        │
│ 2026-02  │ 2026-03-03T09:30   │ 1        │
└──────────┴────────────────────┴──────────┘
      │ 1
      │
      ▼ N
ob_snapshot (마감 데이터)
┌──────────┬──────────┬─────────┬────────┬───────┬───────┬────────┬─────────┬────────┐
│ period   │ SO_ID    │ OS name │ Start  │ Input │Output │Variance│ Ending  │ ...    │
│ (PK)     │ (PK)     │ (PK)    │        │       │       │        │         │        │
├──────────┼──────────┼─────────┼────────┼───────┼───────┼────────┼─────────┼────────┤
│ 2026-01  │ SOD-0001 │ IQ10    │ 0      │ 500만 │ 100만 │ 0      │ 400만   │        │
│ 2026-01  │ SOD-0002 │ MA      │ 0      │ 300만 │ 0     │ 0      │ 300만   │        │
│ 2026-02  │ SOD-0001 │ IQ10    │ 400만  │ 200만 │ 150만 │ 0      │ 450만   │        │
│ ...      │          │         │        │       │       │        │         │        │
└──────────┴──────────┴─────────┴────────┴───────┴───────┴────────┴─────────┴────────┘
```

### 마감 로직 (take_snapshot)

```
close_period.py 2026-02 실행
        │
        ▼
  ① 형식 검증 (yyyy-MM)
        │
        ▼
  ② 순차 검증 ── 이전 마감 = 2026-01? ──── NO → 거부
        │                                         "순차 마감 필요"
       YES
        │
        ▼
  ③ 롤링 CTE 실행 (order_book.sql과 동일한 계산)
     → 2026-02의 모든 SO별 Start/Input/Output/Ending 추출
        │
        ▼
  ④ Variance 계산 (이전 마감이 있는 경우)
     ┌──────────────────────────────────────────┐
     │  현재 raw로 2026-01 재계산 → Ending=1,100│
     │  2026-01 스냅샷 저장값    → Ending=1,000 │
     │  ──────────────────────────────────────  │
     │  Variance = 1,100 - 1,000 = +100        │
     └──────────────────────────────────────────┘
        │
        ▼
  ⑤ Start 결정
     Start = 이전 스냅샷(2026-01)의 Ending (고정값)
     (롤링 Start가 아님!)
        │
        ▼
  ⑥ Ending 계산
     Ending = Start + Input + Variance - Output
        │
        ▼
  ⑦ ob_snapshot에 INSERT (전체 행)
     ob_snapshot_meta에 INSERT (마감 기록)
        │
        ▼
     완료: "2026-02 마감 완료 (294건), Variance 감지 3건"
```

### Variance 발생 시나리오

```
[마감 시점]                       [마감 후 소급 수정]
                                        │
1월 마감 (스냅샷)                        │  누군가 1월 SO 금액 수정
  SOD-0001: Ending = 1,000              │  (1,000 → 1,100)
                                        ▼
                                  sync_db.py 실행
                                  → so_domestic 업데이트
                                        │
                                  ob_snapshot은?
                                  → 여전히 1,000 (불변!)
                                        │
                              2월 마감 (close_period 2026-02)
                                        │
                                  ④ Variance 계산:
                                    raw 재계산 1월 = 1,100
                                    스냅샷 1월     = 1,000
                                    Variance       = +100
                                        │
                                        ▼
                              ┌─────────────────────────┐
                              │ 2월 결과:                │
                              │   Start    = 1,000      │
                              │   Input    = 500        │
                              │   Variance = +100       │
                              │   Output   = 200        │
                              │   Ending   = 1,400      │
                              │                         │
                              │ (Variance 없었다면       │
                              │  Ending = 1,300)        │
                              └─────────────────────────┘
```

### SQL 조회 분기 (order_book_snapshot.sql)

**Open Period만 표시** — 마감된 Period는 `close_period.py --list`로 조회.

```
order_book_snapshot.sql 실행
        │
        ▼
  ob_snapshot_meta에 활성 마감 있나?
        │
   ┌────┴────┐
  YES        NO
   │          │
   │          └─→ 전체 롤링 계산 (order_book.sql과 동일)
   │
   ▼
  Open Period만 출력:
   │
   ├─ 첫 번째 Open Period (마감 직후 월)
   │    → Start = 마지막 스냅샷 Ending
   │    → Variance = raw 재계산 Ending - 스냅샷 Ending
   │    → Ending = Start + Input + Variance - Output
   │
   └─ 이후 Open Period
        → 롤링 계산 (Variance는 이미 Start에 흡수)
```

---

## 사전 조건

### 1. DB 동기화 필수

마감 전에 **반드시** `sync_db.py`를 실행하여 Excel 데이터를 SQLite에 동기화해야 합니다. 스냅샷은 DB의 데이터를 기반으로 계산합니다.

```bash
python sync_db.py --changes    # Excel → SQLite 동기화 (변경 내역 표시)
```

bat 메뉴에서는 `[8] DB Sync` 선택.

### 2. noah_data.db 파일 존재

`sync_db.py`를 한 번이라도 실행하면 DATA_DIR에 `noah_data.db`가 생성됩니다. 마감 시 이 파일이 없으면 오류가 발생합니다.

---

## 마감 순서

**순차 마감이 강제됩니다.** 1월을 마감하지 않으면 2월을 마감할 수 없습니다.

```
1월 마감 → 2월 마감 → 3월 마감 → ...
```

첫 마감은 아무 월이나 가능하지만, 이후에는 반드시 직전 마감의 다음 월만 마감할 수 있습니다.

### 권장 월초 루틴

```
1. Excel 데이터 최신화 (NOAH_SO_PO_DN.xlsx)
2. python sync_db.py --changes        # DB 동기화
3. python close_period.py --list      # 현재 마감 현황 확인
4. python close_period.py 2026-MM     # 전월 마감
5. python close_period.py --list      # 마감 결과 확인
```

---

## CLI 사용법

### 월 마감

```bash
python close_period.py 2026-01                    # 1월 마감
python close_period.py 2026-02 --note "정기 마감"   # 비고 포함
python close_period.py 2026-03 -v                  # 상세 로그
```

### 마감 취소 (Undo)

```bash
python close_period.py --undo 2026-02              # 2월 마감 취소
```

**주의**: 최신 활성 마감만 취소 가능합니다. 1월, 2월이 마감된 상태에서 1월만 취소할 수 없습니다. 1월을 취소하려면 먼저 2월을 취소해야 합니다.

### 마감 현황 조회

```bash
python close_period.py --list                      # 전체 마감 현황 (금액 포함)
```

### 현재 상태 요약

```bash
python close_period.py --status                    # 마지막 마감, 다음 마감 가능 여부
```

---

## BAT 메뉴 사용법

`create_po.bat` 실행 후 메인 메뉴에서 `[9] Order Book Close (월 마감)` 선택:

```
[1] 월 마감           → Period 입력 + 비고 입력 → 마감 실행
[2] 마감 취소 (최신만) → Period 입력 → 최신 마감 취소
[3] 마감 현황 조회     → --list 실행 (금액 합계 테이블)
[4] 현재 상태         → --status 실행 (마지막 마감, 다음 마감)
[0] 메뉴로 돌아가기
```

---

## `--list` 출력 해석

### 출력 예시

```
Order Book 마감 현황
================================================================================================
Period   건수    Start       Input       Output    Variance    Ending   마감일시
------------------------------------------------------------------------------------------------
2026-01  166        0    1,597.2M      45.7M          -    1,551.6M   2026-03-09 16:04
2026-02  294  1,551.6M   2,165.6M     570.9M          -    3,146.4M   2026-03-09 16:08
------------------------------------------------------------------------------------------------
합계                      3,762.9M     616.5M          -    3,146.4M

  2026-02: 정기 마감
```

### 각 컬럼 의미

| 컬럼 | 의미 | 비고 |
|------|------|------|
| Period | 마감 월 (yyyy-MM) | |
| 건수 | 스냅샷에 저장된 행 수 | SO_ID + OS name + 납기일 조합 |
| Start | 월초 수주잔고 (금액) | 전월 Ending과 동일해야 함 |
| Input | 당월 신규 수주 금액 | 등록Period = 해당 월인 SO 합계 |
| Output | 당월 출고 금액 | 출고월 = 해당 월인 DN 합계 |
| Variance | 소급 변경분 | 전월 이전 데이터 수정 시 발생 |
| Ending | 월말 수주잔고 | = Start + Input + Variance - Output |
| 마감일시 | 마감 실행 시각 | |

**금액 단위**: 백만 단위는 `M` 접미사로 표시 (예: `1,597.2M` = 약 15.97억). 1,000 미만은 원 단위 표시.

### 정합성 체크 (Cross-period Validation)

전월 Ending과 당월 Start가 일치하지 않으면 경고가 표시됩니다:

```
2026-02  294  1,551.6M   ...
           ** Start != 전월 Ending (!차이 5.0M)
```

이 경고가 나타나면 데이터 수정이 있었음을 의미합니다. Variance 항목을 확인하세요.

### 합계 행

- **Input 합계**: 전체 기간 신규 수주 총액
- **Output 합계**: 전체 기간 출고 총액
- **Ending**: 마지막 활성 마감 Period의 Ending (현재 수주잔고)
- 취소된 Period는 합계에 포함되지 않음

### 비고 (Notes)

마감 시 `--note`로 입력한 비고는 테이블 하단에 표시됩니다.

---

## Variance 발생 시 대응

### Variance란?

마감 후 과거 데이터가 수정되었을 때, 다음 마감 시 자동으로 감지되는 차이분입니다.

**계산 방식**: 현재 raw 데이터로 이전 Period를 재계산한 Ending - 이전 스냅샷의 Ending

### 발생 원인 (일반적)

| 원인 | 예시 |
|------|------|
| SO 수량 변경 | 주문 수량 100 → 120으로 수정 |
| SO 금액 변경 | 단가 변경, 할인 적용 |
| SO 취소 | Status를 Cancelled로 변경 |
| DN 수정 | 출고 수량/금액 정정 |
| 신규 SO 과거 등록 | 1월 데이터를 2월에 입력 |

### 대응 절차

1. `--list`에서 Variance 발생 확인
2. 마감 결과 메시지에서 Variance 건수 확인
3. DB Browser에서 상세 조회 (아래 SQL 참고)
4. 원인 파악 후 정상 변경이면 → 그대로 진행
5. 비정상이면 → Excel 데이터 수정 → DB 동기화 → 해당 Period 마감 취소 → 재마감

### Variance 상세 조회 SQL

```sql
-- Variance가 있는 행만 조회 (특정 Period)
SELECT SO_ID, [OS name], customer_name, item_name,
       variance_qty, variance_amount,
       start_amount, input_amount, output_amount, ending_amount
FROM ob_snapshot
WHERE snapshot_period = '2026-02'
  AND ABS(variance_amount) > 0.5
ORDER BY ABS(variance_amount) DESC;
```

---

## Undo 사용법과 주의사항

### 기본 사용

```bash
python close_period.py --undo 2026-02    # 2월 마감 취소
```

### 동작

1. 해당 Period의 `ob_snapshot` 데이터 전체 삭제
2. `ob_snapshot_meta`의 `is_active`를 0으로 변경
3. 이후 해당 Period를 다시 마감할 수 있음

### 제약사항

- **최신 활성 마감만 취소 가능**: 1월, 2월이 마감된 상태에서는 2월만 취소 가능
- 중간 Period를 취소하려면 이후 Period부터 역순으로 취소해야 함
- 취소 후에도 `ob_snapshot_meta`에 기록은 남음 (`is_active = 0`)

### 재마감 시나리오

```bash
# 2월 데이터 오류 발견 시
python close_period.py --undo 2026-02     # 2월 취소
# Excel 데이터 수정
python sync_db.py --changes               # DB 동기화
python close_period.py 2026-02            # 2월 재마감
```

---

## DB Browser에서 직접 조회

DB Browser for SQLite (https://sqlitebrowser.org/dl/) 에서 `noah_data.db`를 열어 직접 조회할 수 있습니다.

### 테이블 구조

**`ob_snapshot`** — 스냅샷 데이터

| 컬럼 | 설명 |
|------|------|
| `snapshot_period` | 마감 Period (PK 일부) |
| `SO_ID` | 판매 주문 ID (PK 일부) |
| `OS name` | 기종명 (PK 일부) |
| `Expected delivery date` | 납기일 (PK 일부) |
| `start_qty/amount` | 월초 잔고 |
| `input_qty/amount` | 당월 수주 |
| `output_qty/amount` | 당월 출고 |
| `variance_qty/amount` | 소급 변경분 |
| `ending_qty/amount` | 월말 잔고 |
| `customer_name`, `item_name`, `구분` 등 | 컨텍스트 정보 |

**`ob_snapshot_meta`** — 마감 메타

| 컬럼 | 설명 |
|------|------|
| `period` | 마감 Period (PK) |
| `is_active` | 활성 여부 (1=활성, 0=취소) |
| `closed_at` | 마감 시각 (ISO) |
| `note` | 비고 |

### SQL 예시

```sql
-- 전체 마감 현황 (금액 합계 포함)
SELECT m.period, m.closed_at, m.is_active,
       COUNT(s.SO_ID) AS 건수,
       PRINTF('%.1fM', SUM(s.start_amount) / 1000000.0) AS Start,
       PRINTF('%.1fM', SUM(s.input_amount) / 1000000.0) AS Input,
       PRINTF('%.1fM', SUM(s.output_amount) / 1000000.0) AS Output,
       PRINTF('%.1fM', SUM(s.ending_amount) / 1000000.0) AS Ending
FROM ob_snapshot_meta m
LEFT JOIN ob_snapshot s ON m.period = s.snapshot_period
WHERE m.is_active = 1
GROUP BY m.period
ORDER BY m.period;

-- 특정 Period의 고객별 수주잔고
SELECT customer_name, 구분,
       SUM(ending_qty) AS 잔여수량,
       SUM(ending_amount) AS 잔여금액
FROM ob_snapshot
WHERE snapshot_period = '2026-01'
GROUP BY customer_name, 구분
ORDER BY SUM(ending_amount) DESC;

-- 국내/해외 구분별 잔고 추이
SELECT snapshot_period, 구분,
       SUM(start_amount) AS Start,
       SUM(input_amount) AS Input,
       SUM(output_amount) AS Output,
       SUM(ending_amount) AS Ending
FROM ob_snapshot
GROUP BY snapshot_period, 구분
ORDER BY snapshot_period, 구분;

-- Variance 발생 이력
SELECT snapshot_period, SO_ID, customer_name, item_name,
       variance_qty, variance_amount
FROM ob_snapshot
WHERE ABS(variance_amount) > 0.5
ORDER BY snapshot_period, ABS(variance_amount) DESC;
```

---

## 운용 가이드

### 월초 마감 루틴 (권장)

매월 초 (영업일 1~2일차)에 전월 마감을 수행합니다.

```bash
# 1단계: 데이터 최신화
#   - NOAH_SO_PO_DN.xlsx의 전월 데이터 입력 완료 확인
#   - 특히 DN (출고) 데이터가 전월 말까지 모두 입력되었는지 확인

# 2단계: DB 동기화
python sync_db.py --changes

# 3단계: 현재 상태 확인
python close_period.py --list           # 기존 마감 현황
python close_period.py --status         # 다음 마감 가능 Period 확인

# 4단계: 마감 실행
python close_period.py 2026-MM --note "M월 정기 마감"

# 5단계: 결과 확인
python close_period.py --list           # Variance 발생 여부 확인
```

### 마감 전 체크리스트

- [ ] 전월 SO 데이터 입력 완료
- [ ] 전월 DN (출고) 데이터 입력 완료
- [ ] Cancelled 주문 Status 반영 완료
- [ ] `sync_db.py` 실행 완료
- [ ] `--status`에서 다음 마감 가능 Period 확인

### 주의사항

1. **마감 후 데이터 수정은 Variance로 반영됨**: 마감 후 과거 데이터를 수정하면 다음 마감 시 Variance로 자동 감지됩니다. 스냅샷 자체는 변경되지 않습니다.

2. **DB 동기화 없이 마감하면 안 됨**: 마감은 SQLite DB의 데이터를 기준으로 계산합니다. Excel을 수정한 후 동기화 없이 마감하면 이전 데이터 기준으로 마감됩니다.

3. **순차 마감 건너뛰기 불가**: 1월을 마감하지 않고 2월을 마감할 수 없습니다.

4. **Undo 후 재마감 시 결과가 달라질 수 있음**: 그 사이에 데이터가 변경되었으면 재마감 결과가 달라집니다 (정상 동작).

---

## 소스 코드

| 파일 | 역할 |
|------|------|
| `close_period.py` | CLI 진입점 (마감/취소/현황/상태) |
| `po_generator/snapshot.py` | SnapshotEngine — 스냅샷 생성/취소/조회 |
| `po_generator/db_schema.py` | `create_snapshot_tables()` (ob_snapshot, ob_snapshot_meta 테이블) |
| `sql/order_book_snapshot.sql` | 스냅샷 기반 Order Book SQL |
| `sql/order_book_snapshot_backlog.sql` | 스냅샷 기반 Backlog 뷰 |

---

## 에러 처리

| 상황 | 메시지 | 대응 |
|------|--------|------|
| DB 파일 없음 | `DB 파일이 없습니다` | `sync_db.py` 먼저 실행 |
| 잘못된 형식 | `잘못된 형식: 'xxx' (yyyy-MM 필요)` | `2026-01` 형식으로 입력 |
| 이미 마감됨 | `'2026-01'는 이미 마감되었습니다` | `--list`로 확인 |
| 순차 위반 | `순차 마감 필요: 마지막 마감='2026-01', 다음 마감 가능='2026-02'` | 이전 Period부터 마감 |
| 데이터 없음 | `해당하는 Order Book 데이터가 없습니다` | Excel에 해당 월 데이터 존재 여부 확인 |
| 취소 불가 | `최신 마감만 취소 가능합니다` | 최신 Period부터 역순 취소 |
