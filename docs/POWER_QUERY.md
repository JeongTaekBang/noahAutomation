# Power Query 가이드

NOAH_SO_PO_DN.xlsx 파일에서 사용하는 파워 쿼리 정리.

---

## 데이터 구조

### ERP vs NOAH_SO_PO_DN.xlsx

#### 일반적인 ERP 테이블 구조 (정규화된 관계형)
```
[Customers] 1──N [Sales Orders] 1──N [Sales Lines]
                        │                   │
                        │                   │
[Vendors] 1──N [Purchase Orders] 1──N [Purchase Lines]
                                            │
                                       [Inventory Transactions]
                                            │
                                    [Delivery Notes / Shipments]
```

#### NOAH_SO_PO_DN.xlsx 구조 (반정규화된 Flat 구조)
```
SO_국내 / SO_해외   ← Sales Order + Line (한 행 = 한 아이템)
PO_국내 / PO_해외   ← Purchase Order + Line (한 행 = 한 아이템)
DN_국내 / DN_해외   ← Delivery Note + Line (한 행 = 한 아이템)
```

#### 구조 비교

| 측면 | ERP | NOAH_SO_PO_DN.xlsx |
|------|-----|---------------------|
| 정규화 | 정규화 (Header/Line 분리) | 반정규화 (Header+Line 합쳐짐) |
| 키 관계 | PK-FK (자동 무결성) | SO_ID + Line item (수동 조인) |
| 중복 | 없음 | Customer name 등 반복 |
| 조인 | 자동 (관계 정의됨) | Power Query로 수동 |
| 무결성 | DB 레벨 강제 | 없음 (사람이 실수 가능) |

### 왜 반정규화(Flat) 구조인가?

**담당자가 직접 수기 입력하는 환경**이기 때문.

| 구조 | 장점 | 단점 |
|------|------|------|
| **정규화 (Header/Line 분리)** | 중복 없음, 데이터 무결성 | 입력 번거로움, 여러 시트 오가야 함 |
| **반정규화 (Flat)** | 한 행에 모든 정보, 입력 빠름 | Customer name 등 반복 입력 |

Header/Line 분리 시 문제점:
1. SO Header 시트에서 SO_ID 생성
2. SO Line 시트로 이동해서 SO_ID 참조하며 아이템 입력
3. 같은 주문이면 또 SO_ID 찾아서 입력...

→ 입력 속도 저하 + 실수 증가

**결론**: 수기 입력 환경에서는 Flat 구조가 현실적. ERP 통합 전 임시 운영이므로 효율성 우선.

### 원가 계산 비교

| 항목 | ERP | NOAH_SO_PO_DN.xlsx |
|------|-----|---------------------|
| 원가 출처 | BOM + 구매단가 + 노무비 등 | PO 시트의 `ICO Unit`, `Total ICO` |
| 매출-원가 매칭 | Order/Item 기준 자동 연결 | SO_ID + Line item 기준 조인 |
| 마진 계산 | Sales - COGS (자동) | `Sales amount KRW - 원가` (Power Query) |

```
ERP 원가 계산:
  SO Line ──(자동 연결)── PO Line ──(자동 연결)── Inventory Transaction

NOAH 엑셀 원가 계산:
  SO 시트 ──(Power Query JOIN on SO_ID + Line item)── PO 시트
```

---

## 개요

| 쿼리 | 용도 |
|------|------|
| DN_원가포함 | 출고 내역 + 원가 |
| SO_통합 | 주문 현황 + 원가 + 마진 + 출고 상태 |
| PO_현황 | 발주 현황 + Status별 집계 + 매입금액 |
| **PO_매입월별** | **월별 매입 집계 (IC Balance Confirmation용)** |
| **PO_AX대사** | **Period + AX Project + AX PO별 GRN 금액 집계 (회계 마감 대사용)** |
| **PO_미출고** | **Invoiced인데 DN 미매칭 건 (데이터 점검용)** |
| **PO_출고** | **출고(DN)는 됐는데 계상 라인(Invoiced·외주비)이 전무한 PO (PO_ID 단위, 분할출고·번들 오탐 배제)** |
| **PO_Industry** | **PO_ID별 Industry code + Opportunity + AX Project (분석용)** |
| Inventory_Transaction | 입출고 트랜잭션 (감사 추적용) |
| **Order_Book** | **월별 수주잔고 (Backlog) 롤링 원장 - AX 오더북 형식** |
| **AX_매출대사** | **매출(DN) 라인별 집계 — 국내=세금계산서발행일(선수금건은 출고일) / 해외=선적일(FX 선적월 재환산) 기준 매출월, AX 매출 대사용** |

---

## SO ↔ PO 관계

### 기존 구조 (행 1:1)

초기 설계는 SO_국내와 PO_국내가 **같은 행**으로 연동되는 구조였음:

```
SO_국내 행 1  ←→  PO_국내 행 1
SO_국내 행 2  ←→  PO_국내 행 2
```

PO_국내의 일부 컬럼이 SO_국내를 참조:
- Customer name: `=XLOOKUP(SO_ID, SO_국내[SO_ID], SO_국내[Customer name])`
- Item name: `=XLOOKUP(SO_ID & Line item, ...)`

**문제**: SO 수량 변경 시 PO에 추가 발주가 필요하면 행이 안 맞음.

### 변경된 구조 (1:1 관계)

SO와 PO를 **1:1 매칭**으로 관리:

```
SO_국내: 고객 주문 (행 단위 = 발주 단위)
├── SOD-0001, Line 1, Item A, 수량 10  (1차 주문)
├── SOD-0001, Line 2, Item A, 수량 5   (추가 주문 → 새 Line)

PO_국내: 공장 발주 (SO Line과 1:1)
├── POD-0001, SOD-0001, Line 1, Item A, 수량 10
├── POD-0002, SOD-0001, Line 2, Item A, 수량 5
```

**관계**: SO_ID + Line item 기준 **1:1**

### 운용 규칙

| 상황 | SO_국내 | PO_국내 |
|------|---------|---------|
| 최초 발주 | 새 행, 수량 10 | 새 행, 수량 10 (1:1) |
| 추가 발주 | **새 Line item 추가**, 수량 5 | **새 행 추가**, 수량 5 (1:1) |
| 수량 감소 | 해당 행 수량 수정 또는 Status=Cancelled | 해당 행 Status=Cancelled |
| 전체 취소 | Status=Cancelled | 해당 PO 행 Status=Cancelled |

### PO_Status 정의

| Status | 설명 | 발주 단계 |
|--------|------|----------|
| **Open** | 발주서 등록만, 공장 발주 전 | 1단계 |
| **Sent** | 공장 발주 완료, O.C. 대기 | 2단계 |
| **Confirmed** | 공장 O.C. 수령 | 3단계 |
| **Invoiced** | 공장 출고 완료 | 4단계 (완료) |
| **Cancelled** | 발주 취소 | 제외 |

---

## 비즈니스 배경

### 엔티티 관계
```
NOAH (Factory)                    RCK (Sales Office)
─────────────────                 ─────────────────
제조/출고                    →    판매/고객 대응
AR - RCK (IC)                     AP - NOAH (IC)
```

### 거래 흐름
1. **RCK → NOAH 발주** (PO): RCK가 NOAH에 제품 주문
2. **NOAH 출고** (DN): NOAH가 제품 생산 완료 후 Final Invoice 발행
3. **고객 납품**: RCK는 재고를 보유하지 않음 (Pass-through)
   - NOAH 출고 = RCK 입고 = 고객 납품 (동시 발생)

### DN 시트의 비즈니스 성격

RCK는 재고를 보유하지 않는 **Pass-through 구조**이므로, DN 기록의 의미:

```
NOAH 생산 완료 → NOAH 출고 → (RCK 통과) → 고객 납품
                   ↑
                 DN 발생 시점
```

| 시점 | NOAH | RCK | 고객 |
|------|------|-----|------|
| DN 발생 전 | 생산 중 (WIP) | - | - |
| DN 발생 | 출고 완료 | 입고=출고 (동시) | 수령 |

**DN 시트에 기록되면 다음 세 가지가 동시에 발생한 것으로 가정:**
1. NOAH가 생산 완료했다
2. NOAH가 출고했다 (RCK에 Invoice 발행)
3. 고객이 받았다

**단, 해외 오더의 경우 DN 발생과 고객 수령 사이에 시차가 있음:**
- **국내**: DN 발생 → 다음날 고객 도착 (출고일 바로 입력)
- **해외**: DN 발생(공장 출고) → 인코텀즈에 따라 운송 기간 → 고객 도착 (출고일은 실제 선적 시 입력)
- 따라서 출고금액은 있지만 출고일이 없는 상태 = "공장 출고" (NOAH→RCK 출고 완료, 고객 선적 전)

| 관점 | DN의 의미 |
|------|----------|
| NOAH 관점 | Final Invoice 발행 (AR 인식) |
| RCK 관점 | 매입과 매출이 동시 발생 |
| 물류 관점 | 고객이 물건을 받은 시점 |

### Pass-through 구조 용어 정리

| 용어 | 설명 | 사용 분야 |
|------|------|----------|
| **Pass-through** | 중간에서 그냥 통과시킴 (재고 없이) | 물류, 회계 |
| **Drop Shipping** | 판매자가 재고 없이 제조사→고객 직배송 | 이커머스, 유통 |
| **Cross-docking** | 입고 즉시 출고 (창고 보관 없음) | 물류센터 |
| **Back-to-back Order** | 고객 주문 받으면 바로 공급자에 발주 | 무역, 구매 |
| **Intercompany Pass-through** | 그룹사 간 재고 없이 거래 통과 | 다국적기업 회계 |

RCK-NOAH 구조는 **Back-to-back Order** 또는 **Intercompany Pass-through**가 가장 적합:

```
고객 주문 → RCK (SO 생성) → NOAH (PO 발행) → 생산 → 고객 납품
              ↑
         재고 보유 안 함
         마진만 취함
```

### Drop Shipping과의 차이

| 항목 | Drop Shipping | RCK-NOAH 구조 |
|------|---------------|---------------|
| 배송 | 제조사 → 고객 (직접) | NOAH → 고객 (RCK 명의) |
| 송장 | 제조사가 발행 | RCK가 고객에게 발행 |
| 관계 | 독립 회사 간 | 같은 그룹사 (Intercompany) |
| 용어 | Drop Shipping | Intercompany Back-to-back / Pass-through Trading |

### Inventory Transaction 이해

#### ERP의 Inventory Transaction 생성 시점

**PO Line 등록만으로는 Transaction이 생기지 않는다.**

```
PO 생성 (발주)     → Transaction 없음 (아직 물건이 안 왔음)
     ↓
PO Receipt (입고)  → Inventory Transaction 생성 (+ 재고)
     ↓
SO Shipment (출고) → Inventory Transaction 생성 (- 재고)
```

| 이벤트 | Transaction 생성 | 재고 영향 |
|--------|------------------|----------|
| PO 생성 | X | 없음 |
| PO Receipt (입고 확인) | O | +입고 |
| SO 생성 | X | 없음 (예약만) |
| SO Shipment (출고) | O | -출고 |
| Transfer (창고 이동) | O | A창고-, B창고+ |
| Adjustment (재고 조정) | O | +/- |

**핵심**: 문서 생성이 아니라 **물리적 이벤트**(입고/출고)가 Transaction을 만듦.

#### NOAH_SO_PO_DN.xlsx에서 Inventory Transaction 추출

NOAH_SO_PO_DN.xlsx에는 명시적인 Inventory Transaction 테이블이 없지만, **DN 시트가 사실상 Transaction 역할**을 한다.

Pass-through 구조에서:
```
RCK 관점:
- 입고 = DN 발생 시점 (NOAH에서 받음)
- 출고 = DN 발생 시점 (고객에게 넘김)
- 입고와 출고가 동시 → 재고 잔액 항상 0
```

Power Query로 DN 데이터를 **두 개의 Transaction으로 분리**하면 추출 가능.

### 문제 상황: AX2009 Item 미등록
- NOAH Final Invoice 발행 시점에 RCK의 AX2009에 Item이 등록되지 않은 경우
- AX2009에서 정식 매출 트랜잭션 처리 불가
- 하지만 **Intercompany Balance Confirmation**을 위해 AP 인식 필요

### 임시 회계 처리 (GL 수기 분개)

**DN 발생 시점** (Item 미등록):
```
Dr. Inventory           xxx    ← 원가 (Total ICO)
    Cr. AP - NOAH (IC)      xxx
```
- 물리적으로는 고객에게 납품 완료
- 회계상으로는 재고로 대기 (매출 인식 전)
- **목적**: IC Balance 맞추기 위해 AP 선인식

**Item 등록 후** (AX2009 정식 처리):
```
1) GL 역분개
   Dr. AP - NOAH (IC)    xxx
       Cr. Inventory         xxx

2) AX2009 매출 처리
   → AR/Sales, COGS/Inventory 자동 생성
```

### 요약

| 시점 | 물리적 상태 | 회계 상태 |
|------|-------------|-----------|
| DN 발생 (Item 미등록) | 고객 보유 | Inventory / AP-IC |
| Item 등록 후 | 고객 보유 | AR/Sales, COGS 정리 |

- Timing difference이지만 IC Balance Confirmation을 위해 필요
---

## Inventory_Transaction

### 목적
- DN_국내 + DN_해외 통합
- PO에서 원가 조인 (SO_ID + Item 기준)
- 각 DN을 Receipt(입고) + Issue(출고) 두 개의 Transaction으로 분리
- 감사 추적, 입출고 건수 집계, COGS 계산용

### 결과 컬럼
| 컬럼 | 설명 |
|------|------|
| 출고일 | 국내: 출고일, 해외: 선적일 |
| 구분 | 국내/해외 |
| DN_ID | 출고 번호 |
| SO_ID | 주문 번호 |
| Customer name | 고객명 |
| Business registration number | 사업자등록번호 (SO에서) |
| Item | 아이템명 |
| Line item | 라인 번호 |
| Txn_Type | Receipt (입고) / Issue (출고) |
| From_To | NOAH → RCK / RCK → Customer |
| Qty | 원래 수량 |
| Qty_Change | 재고 변동 (+입고, -출고) |
| 원가_단가 | ICO Unit (PO에서 조인) |
| 원가_합계 | Total ICO (PO에서 조인) |
| Cost_Change | 원가 변동 (+입고, -출고) |

### M 코드
```
let
    // ========== DN 원본 로드 ==========
    DN_국내_Raw = Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content],
    DN_해외_Raw = Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content],

    // ========== 국내: 컬럼 정리 + 태그 ==========
    DN_국내 = Table.SelectColumns(DN_국내_Raw, {"DN_ID", "SO_ID", "Customer name", "Item", "Line item", "Qty", "출고일"}),
    DN_국내_Tagged = Table.AddColumn(DN_국내, "구분", each "국내"),

    // ========== 해외: 컬럼 정리 + 선적일→출고일 + 태그 ==========
    DN_해외 = Table.SelectColumns(DN_해외_Raw, {"DN_ID", "SO_ID", "Customer name", "Item", "Line item", "Qty", "선적일"}),
    DN_해외_Renamed = Table.RenameColumns(DN_해외, {{"선적일", "출고일"}}),
    DN_해외_Tagged = Table.AddColumn(DN_해외_Renamed, "구분", each "해외"),

    // ========== DN 통합 ==========
    DN_Combined = Table.Combine({DN_국내_Tagged, DN_해외_Tagged}),

    // ========== PO 원가 (SO_ID + Line item 기준 합산) ==========
    PO_국내_Raw = Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content],
    PO_해외_Raw = Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content],

    PO_국내 = Table.SelectColumns(PO_국내_Raw, {"SO_ID", "Line item", "ICO Unit", "Total ICO"}),
    PO_해외 = Table.SelectColumns(PO_해외_Raw, {"SO_ID", "Line item", "ICO Unit", "Total ICO"}),
    PO_Combined = Table.Group(Table.Combine({PO_국내, PO_해외}), {"SO_ID", "Line item"}, {
        {"ICO Unit", each List.Average([ICO Unit]), type number},
        {"Total ICO", each List.Sum([Total ICO]), type number}
    }),

    // ========== SO (SO_ID 기준 중복 제거) ==========
    SO_국내_Raw = Excel.CurrentWorkbook(){[Name="SO_국내"]}[Content],
    SO_해외_Raw = Excel.CurrentWorkbook(){[Name="SO_해외"]}[Content],

    SO_국내 = Table.SelectColumns(SO_국내_Raw, {"SO_ID", "Business registration number"}),
    SO_해외 = Table.SelectColumns(SO_해외_Raw, {"SO_ID", "Business registration number"}),
    SO_Combined = Table.Distinct(Table.Combine({SO_국내, SO_해외}), {"SO_ID"}),

    // ========== DN + PO 원가 조인 ==========
    WithCost = Table.NestedJoin(DN_Combined, {"SO_ID", "Line item"}, PO_Combined, {"SO_ID", "Line item"}, "PO_Data", JoinKind.LeftOuter),
    WithCostExpanded = Table.ExpandTableColumn(WithCost, "PO_Data", {"ICO Unit", "Total ICO"}, {"원가_단가", "원가_합계"}),

    // ========== 조인: + SO (Business registration number) ==========
    WithBRN = Table.NestedJoin(WithCostExpanded, {"SO_ID"}, SO_Combined, {"SO_ID"}, "SO_Data", JoinKind.LeftOuter),
    WithBRNExpanded = Table.ExpandTableColumn(WithBRN, "SO_Data", {"Business registration number"}, {"Business registration number"}),

    // ========== Receipt Transaction (NOAH → RCK 입고) ==========
    Receipt = Table.AddColumn(WithBRNExpanded, "Txn_Type", each "Receipt"),
    Receipt_Qty = Table.AddColumn(Receipt, "Qty_Change", each [Qty]),
    Receipt_Cost = Table.AddColumn(Receipt_Qty, "Cost_Change", each [원가_합계]),
    Receipt_Final = Table.AddColumn(Receipt_Cost, "From_To", each "NOAH → RCK"),

    // ========== Issue Transaction (RCK → 고객 출고) ==========
    Issue = Table.AddColumn(WithBRNExpanded, "Txn_Type", each "Issue"),
    Issue_Qty = Table.AddColumn(Issue, "Qty_Change", each -[Qty]),
    Issue_Cost = Table.AddColumn(Issue_Qty, "Cost_Change", each if [원가_합계] = null then null else -[원가_합계]),
    Issue_Final = Table.AddColumn(Issue_Cost, "From_To", each "RCK → Customer"),

    // ========== 통합 + 정렬 ==========
    Combined = Table.Combine({Receipt_Final, Issue_Final}),
    Sorted = Table.Sort(Combined, {
        {"출고일", Order.Ascending},
        {"SO_ID", Order.Ascending},
        {"Txn_Type", Order.Descending}
    }),

    // ========== 컬럼 순서 정리 ==========
    Reordered = Table.ReorderColumns(Sorted, {
        "출고일", "구분", "DN_ID", "SO_ID", "Customer name", "Business registration number", "Item", "Line item",
        "Txn_Type", "From_To", "Qty", "Qty_Change", "원가_단가", "원가_합계", "Cost_Change"
    }),

    // ========== 타입 변환 ==========
    Result = Table.TransformColumnTypes(Reordered, {
        {"출고일", type date},
        {"Qty", Int64.Type},
        {"Qty_Change", Int64.Type},
        {"원가_단가", Currency.Type},
        {"원가_합계", Currency.Type},
        {"Cost_Change", Currency.Type}
    })
in
    Result
```

### 결과 예시

| 출고일 | 구분 | DN_ID | SO_ID | Customer name | Business registration number | Item | Txn_Type | From_To | Qty | Qty_Change | 원가_단가 | 원가_합계 | Cost_Change |
|--------|------|-------|-------|---------------|------------------------------|------|----------|---------|-----|------------|-----------|-----------|-------------|
| 2026-01-15 | 국내 | DN-001 | ND-0001 | 삼성전자 | 123-45-67890 | IQ3 | Receipt | NOAH → RCK | 2 | 2 | 500,000 | 1,000,000 | 1,000,000 |
| 2026-01-15 | 국내 | DN-001 | ND-0001 | 삼성전자 | 123-45-67890 | IQ3 | Issue | RCK → Customer | 2 | -2 | 500,000 | 1,000,000 | -1,000,000 |
| 2026-01-20 | 해외 | DN-002 | NE-0001 | ABC Corp | 987-65-43210 | CVA | Receipt | NOAH → RCK | 1 | 1 | 800,000 | 800,000 | 800,000 |
| 2026-01-20 | 해외 | DN-002 | NE-0001 | ABC Corp | 987-65-43210 | CVA | Issue | RCK → Customer | 1 | -1 | 800,000 | 800,000 | -800,000 |

### 활용 예시

| 분석 | 방법 |
|------|------|
| 월별 출고 건수 | `Txn_Type = "Issue"` 필터 → 출고일 기준 그룹화 |
| 월별 COGS | `Txn_Type = "Issue"` → Cost_Change 합계 (부호 반전) |
| 고객별 입출고 이력 | Customer name 필터 |
| 고객별 원가 | Customer name 그룹화 → 원가_합계 합계 |
| 국내/해외 비율 | 구분 컬럼 피벗 |
| 감사 추적 | 전체 데이터 시간순 정렬 |
| 재고 가치 | Cost_Change 누적 합계 (Pass-through라 항상 0) |

### 실용성 판단

| 질문 | 답변 |
|------|------|
| 기술적으로 가능? | O |
| 실용적 가치? | O (원가 포함으로 COGS 분석 가능) |
| 언제 유용? | 입출고 건수 집계, COGS 계산, 감사 추적 |

**참고**: Pass-through 구조에서는 재고 잔액/가치가 항상 0이지만, 원가 정보가 포함되어 COGS 분석과 고객별 원가 집계에 활용 가능.

---

## DN_원가포함

### 목적
- DN_국내 + DN_해외 통합
- PO에서 원가 조인 (SO_ID + Item 기준)
- SO에서 Model code 조인

### 결과 컬럼
| 컬럼 | 설명 |
|------|------|
| DN_ID | 출고 번호 |
| SO_ID | 주문 번호 |
| Customer name | 고객명 |
| Item | 아이템명 |
| Qty | 수량 |
| Unit Price | 판매 단가 |
| Total Sales KRW | 매출 (KRW) |
| 출고일 | 국내: 출고일, 해외: 선적일 |
| 구분 | 국내/해외 |
| 원가_단가 | ICO Unit (PO에서) |
| 원가_합계 | Total ICO (PO에서) |
| Opportunity | Opportunity 번호 (SO에서) |
| Customer PO | 고객 발주번호 (SO에서) |
| OS name | OneStream Item name (SO에서) |
| Business registration number | 사업자등록번호 (SO에서) |
| Model code | ERP 프로젝트 번호 (SO에서) |
| AX PO | AX 발주번호 (PO에서) |
| Currency | 통화 (SO에서) |
### 용도
- **IC Balance Confirmation** → 원가_합계 합계 = RCK AP-NOAH (IC)
  ```
  NOAH AR Statement  vs  DN_원가포함 (원가_합계 SUM)
  ```

### M 코드
```
let
    // ========== DN 테이블 ==========
    DN_국내_Raw = Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content],
    DN_해외_Raw = Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content],

    // 국내: 컬럼 정리 + 구분 태그
    DN_국내 = Table.SelectColumns(DN_국내_Raw, {"DN_ID", "SO_ID", "Customer name", "Item", "Line item", "Qty", "Unit Price", "Total Sales", "출고일"}),
    DN_국내_Renamed = Table.RenameColumns(DN_국내, {{"Total Sales", "Total Sales KRW"}}),
    DN_국내_Tagged = Table.AddColumn(DN_국내_Renamed, "구분", each "국내"),

    // 해외: 컬럼 정리 + 선적일 → 출고일로 rename + 구분 태그
    DN_해외 = Table.SelectColumns(DN_해외_Raw, {"DN_ID", "SO_ID", "Customer name", "Item", "Line item", "Qty", "Unit Price", "Total Sales KRW", "선적일"}),
    DN_해외_Renamed = Table.RenameColumns(DN_해외, {{"선적일", "출고일"}}),
    DN_해외_Tagged = Table.AddColumn(DN_해외_Renamed, "구분", each "해외"),

    // DN 통합
    DN_Combined = Table.Combine({DN_국내_Tagged, DN_해외_Tagged}),

    // ========== PO 원가 (SO_ID + Line item 기준 합산) ==========
    PO_국내_Raw = Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content],
    PO_해외_Raw = Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content],

    PO_국내 = Table.SelectColumns(PO_국내_Raw, {"SO_ID", "Line item", "ICO Unit", "Total ICO", "AX PO"}),
    PO_해외 = Table.SelectColumns(PO_해외_Raw, {"SO_ID", "Line item", "ICO Unit", "Total ICO", "AX PO"}),
    PO_Combined = Table.Group(Table.Combine({PO_국내, PO_해외}), {"SO_ID", "Line item"}, {
        {"ICO Unit", each List.Average([ICO Unit]), type number},
        {"Total ICO", each List.Sum([Total ICO]), type number},
        {"AX PO", each Text.Combine(List.Distinct(List.RemoveNulls([AX PO])), ", "), type text}
    }),

    // ========== SO (SO_ID + Line item 기준 중복 제거) ==========
    SO_국내_Raw = Excel.CurrentWorkbook(){[Name="SO_국내"]}[Content],
    SO_해외_Raw = Excel.CurrentWorkbook(){[Name="SO_해외"]}[Content],

    SO_국내 = Table.SelectColumns(SO_국내_Raw, {"SO_ID", "Line item", "Opportunity", "Customer PO", "OS name", "Business registration number", "Model code", "Currency"}),
    SO_해외 = Table.SelectColumns(SO_해외_Raw, {"SO_ID", "Line item", "Opportunity", "Customer PO", "OS name", "Business registration number", "Model code", "Currency"}),
    SO_Combined = Table.Distinct(Table.Combine({SO_국내, SO_해외}), {"SO_ID", "Line item"}),

    // ========== 조인: DN + PO (원가) ==========
    WithCost = Table.NestedJoin(DN_Combined, {"SO_ID", "Line item"}, PO_Combined, {"SO_ID", "Line item"}, "PO_Data", JoinKind.LeftOuter),
    WithCostExpanded = Table.ExpandTableColumn(WithCost, "PO_Data", {"ICO Unit", "Total ICO", "AX PO"}, {"원가_단가", "원가_합계", "AX PO"}),

    // ========== 조인: + SO (SO_ID + Line item 기준) ==========
    WithAX = Table.NestedJoin(WithCostExpanded, {"SO_ID", "Line item"}, SO_Combined, {"SO_ID", "Line item"}, "SO_Data", JoinKind.LeftOuter),
    WithAXExpanded = Table.ExpandTableColumn(WithAX, "SO_Data", {"Opportunity", "Customer PO", "OS name", "Business registration number", "Model code", "Currency"}, {"Opportunity", "Customer PO", "OS name", "Business registration number", "Model code", "Currency"}),

    // ========== 타입 변환 ==========
    Result = Table.TransformColumnTypes(WithAXExpanded, {
        {"출고일", type date},
        {"Total Sales KRW", Currency.Type},
        {"원가_단가", Currency.Type},
        {"원가_합계", Currency.Type}
    })
in
    Result
```

---

## SO_통합

### 목적
- SO_국내 + SO_해외 통합
- PO에서 발주번호(PO_ID·NOAH O.C No.) + 원가 조인 → 마진/마진율 계산
- DN에서 출고금액 조인 → 출고완료 여부, 미출고금액 계산
- Cancelled·Hold 건 제외

### 결과 컬럼
| 컬럼 | 설명 |
|------|------|
| SO_ID | 주문 번호 |
| AX Period | AX 기간 (yyyy-MM) |
| Customer name | 고객명 |
| Item name | 아이템명 |
| Sales amount | 매출 외화 (해외만, 국내는 null) |
| Sales amount KRW | 매출 (KRW) |
| 구분 | 국내/해외 |
| PO_ID | 발주 번호 (PO에서 조인, 분할 발주 시 콤마 연결) |
| NOAH O.C No. | 공장 발주번호 (PO에서 조인, 분할 발주 시 콤마 연결) |
| 원가_단가 | ICO Unit |
| 원가 | Total ICO |
| DN_ID | 출고 번호 (분할 출고 시 콤마 연결) |
| 출고수량 | DN에서 출고된 수량 (분할 출고 합산) |
| 출고금액 | DN에서 출고된 금액 |
| 출고일 | DN에서 출고된 날짜 |
| 마진 | Sales - 원가 |
| 마진율 | 마진 / Sales (%) |
| 출고완료 | 미출고/부분 출고/공장 출고/출고 완료 (**수량 기준** 판정 — 무상공급 대응) |
| 매출연월 | 출고일 기준 연월 (yyyy-MM) |
| 미출고금액 | Sales - 출고금액 |

### 용도
- **마진율 정렬** → 수익성 낮은 건 파악
- **PO_ID·NOAH O.C No.** → 이 주문 라인이 어느 발주로 나갔는지 역추적 (PO 시트·공장 문의)
- **출고완료 = 미출고/공장 출고** → 미출고·미선적 현황
- **미출고금액 합계** → 백로그 파악
- **매출연월 그룹화** → 월별 매출 집계

### M 코드
```
let
    // ========== SO 원본 ==========
    SO_국내_Raw = Excel.CurrentWorkbook(){[Name="SO_국내"]}[Content],
    SO_해외_Raw = Excel.CurrentWorkbook(){[Name="SO_해외"]}[Content],

    // 공통 컬럼 찾기 (Status, 납품 주소 제외)
    국내_Columns = Table.ColumnNames(SO_국내_Raw),
    해외_Columns = Table.ColumnNames(SO_해외_Raw),
    CommonColumns = List.Intersect({국내_Columns, 해외_Columns}),
    CommonColumns_Filtered = List.RemoveItems(CommonColumns, {"Sales amount", "Sales amount KRW", "Status", "납품 주소"}),

    // 국내: 공통컬럼 + Sales amount → Sales amount KRW로 rename
    SO_국내_Selected = Table.SelectColumns(SO_국내_Raw, CommonColumns_Filtered & {"Sales amount", "Status"}),
    SO_국내_Renamed = Table.RenameColumns(SO_국내_Selected, {{"Sales amount", "Sales amount KRW"}}),
    SO_국내_Tagged = Table.AddColumn(SO_국내_Renamed, "구분", each "국내"),

    // 해외: 공통컬럼 + Sales amount (외화) + Sales amount KRW (원화)
    SO_해외_Selected = Table.SelectColumns(SO_해외_Raw, CommonColumns_Filtered & {"Sales amount", "Sales amount KRW", "Status"}),
    SO_해외_Tagged = Table.AddColumn(SO_해외_Selected, "구분", each "해외"),

    // SO 합치기 + 에러 치환 + Cancelled·Hold 제외 (null은 포함)
    SO_Combined = Table.Combine({SO_국내_Tagged, SO_해외_Tagged}),
    SO_CleanErrors = Table.ReplaceErrorValues(SO_Combined,
        List.Transform(Table.ColumnNames(SO_Combined), each {_, null})
    ),
    SO_Filtered = Table.SelectRows(SO_CleanErrors, each [Status] = null or not List.Contains({"Cancelled", "Hold"}, [Status])),
    SO_Final = Table.RemoveColumns(SO_Filtered, {"Status"}),

    // ========== PO 원가 (SO_ID + Line item 기준 합산) ==========
    PO_국내_Raw = Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content],
    PO_해외_Raw = Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content],

    PO_국내_Select = Table.SelectColumns(PO_국내_Raw, {"SO_ID", "Line item", "PO_ID", "NOAH O.C No.", "ICO Unit", "Total ICO"}),
    PO_국내 = Table.ReplaceErrorValues(PO_국내_Select,
        List.Transform(Table.ColumnNames(PO_국내_Select), each {_, null})
    ),
    PO_해외_Select = Table.SelectColumns(PO_해외_Raw, {"SO_ID", "Line item", "PO_ID", "NOAH O.C No.", "ICO Unit", "Total ICO"}),
    PO_해외 = Table.ReplaceErrorValues(PO_해외_Select,
        List.Transform(Table.ColumnNames(PO_해외_Select), each {_, null})
    ),
    // PO_ID·NOAH O.C No.는 원가와 같은 행 집합에서 뽑는다 (Status 필터 없음) —
    // 원가가 있는데 PO_ID가 비는 불일치를 막기 위함. 분할발주로 한 SO 라인에
    // PO가 여럿이면 콤마로 연결(PO_현황과 동일 규칙), 공란은 List.RemoveNulls로 제거.
    PO_Combined = Table.Group(Table.Combine({PO_국내, PO_해외}), {"SO_ID", "Line item"}, {
        {"PO_ID", each Text.Combine(List.Distinct(List.RemoveNulls([PO_ID])), ", "), type text},
        {"NOAH O.C No.", each Text.Combine(List.Distinct(List.RemoveNulls([#"NOAH O.C No."])), ", "), type text},
        {"ICO Unit", each List.Average([ICO Unit]), type number},
        {"Total ICO", each List.Sum([Total ICO]), type number}
    }),

    // ========== DN 출고 (SO_ID + Line item 기준 합산) - 분할 출고 대응 ==========
    DN_국내_Raw = Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content],
    DN_해외_Raw = Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content],

    DN_국내_Select = Table.SelectColumns(DN_국내_Raw, {"SO_ID", "Line item", "DN_ID", "Qty", "Total Sales", "출고일"}),
    DN_국내 = Table.ReplaceErrorValues(DN_국내_Select,
        List.Transform(Table.ColumnNames(DN_국내_Select), each {_, null})
    ),
    DN_국내_Renamed = Table.RenameColumns(DN_국내, {{"Total Sales", "출고금액"}, {"Qty", "출고수량"}}),
    DN_해외_Select = Table.SelectColumns(DN_해외_Raw, {"SO_ID", "Line item", "DN_ID", "Qty", "Total Sales KRW", "선적일"}),
    DN_해외 = Table.ReplaceErrorValues(DN_해외_Select,
        List.Transform(Table.ColumnNames(DN_해외_Select), each {_, null})
    ),
    DN_해외_Renamed = Table.RenameColumns(DN_해외, {{"Total Sales KRW", "출고금액"}, {"선적일", "출고일"}, {"Qty", "출고수량"}}),
    DN_Combined = Table.Group(Table.Combine({DN_국내_Renamed, DN_해외_Renamed}), {"SO_ID", "Line item"}, {
        {"DN_ID", each Text.Combine(List.Distinct(List.RemoveNulls([DN_ID])), ", "), type text},
        {"출고수량", each List.Sum([출고수량]), type number},
        {"출고금액", each List.Sum([출고금액]), Currency.Type},
        {"출고일", each List.Max([출고일]), type nullable date}
    }),

    // ========== SO에 원가 조인 (SO_ID + Line item) ==========
    WithCost = Table.NestedJoin(SO_Final, {"SO_ID", "Line item"}, PO_Combined, {"SO_ID", "Line item"}, "PO", JoinKind.LeftOuter),
    WithCostExpanded = Table.ExpandTableColumn(WithCost, "PO", {"PO_ID", "NOAH O.C No.", "ICO Unit", "Total ICO"}, {"PO_ID", "NOAH O.C No.", "원가_단가", "원가"}),

    // ========== SO에 출고 조인 (SO_ID + Line item) - 출고일 포함 ==========
    WithShip = Table.NestedJoin(WithCostExpanded, {"SO_ID", "Line item"}, DN_Combined, {"SO_ID", "Line item"}, "DN", JoinKind.LeftOuter),
    WithShipExpanded = Table.ExpandTableColumn(WithShip, "DN", {"DN_ID", "출고수량", "출고금액", "출고일"}, {"DN_ID", "출고수량", "출고금액", "출고일"}),

    // ========== 계산 컬럼 추가 ==========
    WithMargin = Table.AddColumn(WithShipExpanded, "마진", each [Sales amount KRW] - (if [원가] = null then 0 else [원가]), type number),
    WithMarginRate = Table.AddColumn(WithMargin, "마진율", each if [Sales amount KRW] = 0 or [Sales amount KRW] = null then null else [마진] / [Sales amount KRW], Percentage.Type),
    // 출고 상태는 금액이 아닌 "수량"으로 판정 — 무상공급(단가 0)·환율차 케이스 대응
    WithShipStatus = Table.AddColumn(WithMarginRate, "출고완료", each
        let 발주수량 = if [Item qty] = null then 0 else [Item qty] in
            if [출고수량] = null then "미출고"
            else if 발주수량 - [출고수량] > 0 then "부분 출고"
            else if [출고일] = null then "공장 출고"
            else "출고 완료",
        type text),
    WithSalesMonth = Table.AddColumn(WithShipStatus, "매출연월", each
        if [출고일] = null then null
        else Text.From(Date.Year([출고일])) & "-" & Text.PadStart(Text.From(Date.Month([출고일])), 2, "0"),
        type text),
    WithRemaining = Table.AddColumn(WithSalesMonth, "미출고금액", each [Sales amount KRW] - (if [출고금액] = null then 0 else [출고금액]), type number),

    // ========== 타입 변환 ==========
    Result = Table.TransformColumnTypes(WithRemaining, {
        {"PO receipt date", type date},
        {"Requested delivery date", type date},
        {"출고일", type date},
        {"Sales amount", Currency.Type},
        {"Sales amount KRW", Currency.Type},
        {"원가_단가", Currency.Type},
        {"원가", Currency.Type},
        {"출고금액", Currency.Type},
        {"마진", Currency.Type},
        {"미출고금액", Currency.Type}
    }),
    #"다시 정렬한 열 수" = Table.ReorderColumns(Result,{"SO_ID", "PO receipt date", "Period", "AX Period", "AX Project number", "CS담당자", "Business registration number", "Customer name", "Customer PO", "Order type", "Opportunity", "Sector", "Industry code", "Model code", "Item name", "OS name", "Currency", "Line item", "Item qty", "Sales Unit Price", "Incoterms", "Requested delivery date", "EXW NOAH", "Expected delivery date", "영업 담당", "Remarks", "Sales amount KRW", "구분", "Sales amount", "PO_ID", "NOAH O.C No.", "원가_단가", "원가", "DN_ID", "출고수량", "출고금액", "출고일", "마진", "마진율", "출고완료", "매출연월", "미출고금액"})
in
    #"다시 정렬한 열 수"
```

---

## PO_현황

### 목적
- SO_국내 + SO_해외의 주문 수량과 PO 발주 현황 비교
- Status별 발주 수량/금액 집계
- 미발주 현황 파악
- **IC Balance Confirmation** 활용 (Invoiced 금액 = AP-NOAH)

### PO_Status 참고

| Status | 의미 | 발주 집계 | 매입 집계 |
|--------|------|----------|----------|
| Open | 발주서 등록만, 공장 발주 전 | 제외 | 제외 |
| Sent | 공장 발주 완료, O.C. 대기 | 포함 | 제외 |
| Confirmed | 공장 O.C. 수령 | 포함 | 제외 |
| **Invoiced P01/P02** | **공장 출고 완료 → RCK AP 인식** | 포함 | **포함** |
| Holding | 발주 보류 | 제외 | 제외 |
| Cancelled | 발주 취소 | 제외 | 제외 |

**Invoiced 의미**: 공장(NOAH)이 Final Invoice 발행 = RCK 입장에서 매입(AP-NOAH IC) 발생
- P01 = 1차 출고, P02 = 2차 출고 (분할 출고 시)
- **매입금액 합계 = IC Balance Confirmation 대상**

### 결과 컬럼

| 컬럼 | 설명 |
|------|------|
| SO_ID | 주문 번호 |
| Customer name | 고객명 |
| Item name | 아이템명 |
| 구분 | 국내/해외 |
| SO수량 | SO의 주문 수량 (같은 SO_ID+Item 합계) |
| PO_ID | 발주 번호 (여러 개면 콤마로 연결) |
| 발주수량 | PO 합계 (Sent/Confirmed/Invoiced만) |
| 발주금액 | Total ICO 합계 |
| 최근발주일 | 마지막 발주 날짜 |
| AX PO | AX 발주번호 (여러 개면 콤마로 연결) |
| **매입수량** | Invoiced 건 수량 (공장 출고 완료 = RCK AP) |
| **매입금액** | Invoiced 건 금액 (IC Balance 대상) |
| 미발주수량 | SO수량 - 발주수량 |
| 발주완료 | Y/N |

### 용도

| 필터/분석 | 용도 |
|----------|------|
| 발주완료 = N | 추가 발주 필요한 건 |
| 미발주수량 > 0 | 부분 발주된 건 |
| **매입금액 합계** | **IC Balance Confirmation** (RCK AP-NOAH = NOAH AR) |

**참고**: 출고 현황은 **SO_통합** 쿼리에서 확인

### M 코드

```
let
    // ========== SO 원본 ==========
    SO_국내_Raw = Excel.CurrentWorkbook(){[Name="SO_국내"]}[Content],
    SO_해외_Raw = Excel.CurrentWorkbook(){[Name="SO_해외"]}[Content],

    SO_국내 = Table.SelectColumns(SO_국내_Raw, {"SO_ID", "Customer name", "Item name", "Line item", "Item qty"}),
    SO_국내_Tagged = Table.AddColumn(SO_국내, "구분", each "국내"),

    SO_해외 = Table.SelectColumns(SO_해외_Raw, {"SO_ID", "Customer name", "Item name", "Line item", "Item qty"}),
    SO_해외_Tagged = Table.AddColumn(SO_해외, "구분", each "해외"),

    SO_Combined = Table.Combine({SO_국내_Tagged, SO_해외_Tagged}),

    // SO_ID + Line item 기준 그룹화 (같은 조합의 수량 합계)
    SO_Grouped = Table.Group(SO_Combined, {"SO_ID", "Line item", "구분"}, {
        {"SO수량", each List.Sum([Item qty]), type number},
        {"Customer name", each List.First([Customer name]), type text},
        {"Item name", each List.First([Item name]), type text}
    }),

    // ========== PO 원본 (Open, Cancelled, Holding 제외) ==========
    PO_국내_Raw = Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content],
    PO_해외_Raw = Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content],

    PO_국내 = Table.SelectColumns(PO_국내_Raw, {"PO_ID", "SO_ID", "Line item", "Item qty", "Total ICO", "Status", "공장 발주 날짜", "AX PO"}),
    PO_해외 = Table.SelectColumns(PO_해외_Raw, {"PO_ID", "SO_ID", "Line item", "Item qty", "Total ICO", "Status", "공장 발주 날짜", "AX PO"}),
    PO_Combined = Table.Combine({PO_국내, PO_해외}),

    // Open, Cancelled, Holding, null 제외 (Sent, Confirmed, Invoiced P01/P02 등 포함)
    PO_Filtered = Table.SelectRows(PO_Combined, each [Status] <> null and not List.Contains({"Open", "Cancelled", "Holding"}, [Status])),

    // SO_ID + Line item 기준 그룹화 (발주 전체 + Invoiced 별도 집계)
    PO_Grouped = Table.Group(PO_Filtered, {"SO_ID", "Line item"}, {
        {"PO_ID", each Text.Combine(List.Distinct([PO_ID]), ", "), type text},
        {"발주수량", each List.Sum([Item qty]), type number},
        {"발주금액", each List.Sum([Total ICO]), type number},
        {"최근발주일", each List.Max([공장 발주 날짜]), type date},
        {"AX PO", each Text.Combine(List.Distinct(List.RemoveNulls([AX PO])), ", "), type text},
        // Invoiced P01, P02 등 = 공장 출고 완료 = RCK 매입(AP) 대상
        {"매입수량", each List.Sum(Table.SelectRows(_, each Text.StartsWith([Status], "Invoiced"))[Item qty]), type number},
        {"매입금액", each List.Sum(Table.SelectRows(_, each Text.StartsWith([Status], "Invoiced"))[Total ICO]), type number}
    }),

    // ========== SO + PO 발주 조인 ==========
    WithPO = Table.NestedJoin(SO_Grouped, {"SO_ID", "Line item"}, PO_Grouped, {"SO_ID", "Line item"}, "PO_Data", JoinKind.LeftOuter),
    WithPOExpanded = Table.ExpandTableColumn(WithPO, "PO_Data", {"PO_ID", "발주수량", "발주금액", "최근발주일", "AX PO", "매입수량", "매입금액"}, {"PO_ID", "발주수량", "발주금액", "최근발주일", "AX PO", "매입수량", "매입금액"}),

    // ========== 계산 컬럼 ==========
    // null 안전 처리: SO수량이 null이면 0으로 대체
    WithRemaining = Table.AddColumn(WithPOExpanded, "미발주수량", each (if [SO수량] = null then 0 else [SO수량]) - (if [발주수량] = null then 0 else [발주수량]), type number),
    // null 비교 시 null <= 0 은 null 반환 → if null then 오류 발생하므로 명시적 null 체크
    WithOrderStatus = Table.AddColumn(WithRemaining, "발주완료", each if [미발주수량] = null or [미발주수량] <= 0 then "Y" else "N", type text),

    // ========== 정렬 ==========
    Sorted = Table.Sort(WithOrderStatus, {{"발주완료", Order.Ascending}, {"SO_ID", Order.Ascending}}),

    // ========== 타입 변환 ==========
    Result = Table.TransformColumnTypes(Sorted, {
        {"SO수량", Int64.Type},
        {"발주수량", Int64.Type},
        {"발주금액", Currency.Type},
        {"매입수량", Int64.Type},
        {"매입금액", Currency.Type},
        {"미발주수량", Int64.Type},
        {"최근발주일", type date}
    })
in
    Result
```

### 결과 예시

| SO_ID | Customer name | Item name | 구분 | SO수량 | PO_ID | 발주수량 | 발주금액 | 최근발주일 | AX PO | 매입수량 | 매입금액 | 미발주수량 | 발주완료 |
|-------|---------------|-----------|------|--------|-------|----------|----------|------------|-------|----------|----------|------------|----------|
| SOD-0001 | 삼성전자 | IQ3 | 국내 | 15 | POD-0001, POD-0005 | 15 | 7,500,000 | 2026-01-20 | PO-001 | 15 | 7,500,000 | 0 | Y |
| SOD-0002 | LG전자 | CVA | 국내 | 10 | POD-0002 | 10 | 8,000,000 | 2026-01-15 | PO-002 | 0 | 0 | 0 | Y |
| SOD-0003 | 현대중공업 | NA028 | 국내 | 20 | POD-0003 | 15 | 6,000,000 | 2026-01-25 | PO-003 | 10 | 4,000,000 | 5 | N |
| SOO-0019 | ASC | ACTEA BUSH | 해외 | 10 | - | - | - | - | - | - | - | 10 | N |

**IC Balance Confirmation 활용**: 매입금액 합계 = RCK AP-NOAH (IC) = NOAH AR Statement

---

## PO_매입월별

### 목적
- **월별 IC Balance Confirmation** 용도
- Invoiced 건(공장 출고 완료)의 매입금액을 PO Period별로 집계

### 데이터 흐름
```
PO_국내 + PO_해외
    │
    │ Invoiced 필터
    ▼
Period + 구분별 그룹화
```

### 결과 컬럼
| 컬럼 | 설명 |
|------|------|
| Period | PO 시트의 Period (yyyy-MM 형식) |
| 구분 | 국내/해외 |
| 매입건수 | Invoiced PO 라인 수 |
| 매입수량 | Item qty 합계 |
| 매입금액 | Total ICO 합계 (= RCK AP) |

### M 코드

> **v3 (2026-03 수정)**: DN 조인 제거. PO 시트의 Period + Invoiced Status만으로 집계.
> 기존 DN 출고일 기반 방식은 Invoiced P01/P02 등 분할 인보이스 시 오집계 발생.

```
let
    PO_국내_Raw = Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content],
    PO_해외_Raw = Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content],

    PO_국내 = Table.SelectColumns(PO_국내_Raw, {"Period", "Item qty", "Total ICO", "Status"}),
    PO_국내_Tagged = Table.AddColumn(PO_국내, "구분", each "국내"),

    PO_해외 = Table.SelectColumns(PO_해외_Raw, {"Period", "Item qty", "Total ICO", "Status"}),
    PO_해외_Tagged = Table.AddColumn(PO_해외, "구분", each "해외"),

    PO_Combined = Table.Combine({PO_국내_Tagged, PO_해외_Tagged}),
    PO_Invoiced = Table.SelectRows(PO_Combined, each [Status] <> null and Text.StartsWith([Status], "Invoiced")),

    Grouped = Table.Group(PO_Invoiced, {"Period", "구분"}, {
        {"매입건수", each Table.RowCount(_), Int64.Type},
        {"매입수량", each List.Sum([Item qty]), Int64.Type},
        {"매입금액", each List.Sum([Total ICO]), Currency.Type}
    }),

    Sorted = Table.Sort(Grouped, {{"Period", Order.Descending}, {"구분", Order.Ascending}})
in
    Sorted
```

### 결과 예시

| Period | 구분 | 매입건수 | 매입수량 | 매입금액 |
|--------|------|----------|----------|----------|
| 2026-02 | 국내 | 5 | 25 | 12,500,000 |
| 2026-02 | 해외 | 2 | 10 | 8,000,000 |
| 2026-01 | 국내 | 8 | 40 | 20,000,000 |
| 2026-01 | 해외 | 3 | 15 | 12,000,000 |

### 용도

| 필터/분석 | 용도 |
|----------|------|
| 특정 Period 필터 | 해당 월 IC Balance Confirmation |
| 구분별 합계 | 국내/해외 매입 비교 |

**IC Balance 확인 방법**:
```
NOAH AR Statement (2026-01월)  vs  PO_매입월별 (Period = 2026-01) 매입금액 합계
```

---

## PO_AX대사

### 목적
- **회계 마감 시 AX GRN 대사** 용도
- Invoiced(GRN 처리 완료) 건만 대상
- SO_ID 기준 flat 구조 — Model code, AX PO를 LEFT JOIN으로 나열
- AX에 입력된 GRN 금액과 엑셀 매입금액 비교

### 대사 프로세스
```
엑셀 (PO_AX대사 쿼리)                                              AX (D365 F&O) GRN
┌─────────────────────────────────────────────────────────┐    ┌─────────────────────────────────┐
│ 2026-01  PRJ-001  P000001  SOD-0001  ₩5,000,000        │──→ │                                 │
│ 2026-01  PRJ-001  P000001  SOD-0005  ₩2,500,000        │──→ │ 2026-01  P000001  ₩7,500,000   │ ✓ 합계 일치
│ 2026-01  PRJ-001  P000002  SOD-0002  ₩8,000,000        │──→ │ 2026-01  P000002  ₩8,000,000   │ ✓ 일치
│ 2026-02  PRJ-002  P000003  SOD-0007  ₩6,000,000        │──→ │ 2026-02  P000003  ₩4,000,000   │ ✗ 불일치
└─────────────────────────────────────────────────────────┘    └─────────────────────────────────┘
  ※ SO_ID별 행이 분리되어 불일치 시 어떤 SO에서 차이인지 즉시 파악 가능
  ※ AX PO별 합계는 Excel 피벗/필터로 확인
```

### 결과 컬럼

| 컬럼 | 설명 |
|------|------|
| Period | 출고일 기준 월 (yyyy-MM 형식) |
| Model code | AX 프로젝트번호 (SO에서 LEFT JOIN) |
| AX PO | AX 발주번호 (PXXXXXX) |
| 구분 | 국내/해외 |
| SO_ID | NOAH SO 번호 (행 기준 키) |
| PO_ID | 포함된 NOAH PO 번호 (여러 개면 콤마로 연결) |
| 건수 | Invoiced PO Line 수 |
| 수량 | Invoiced 수량 합계 |
| 금액 | Invoiced Total ICO 합계 (= AX GRN 금액과 대사 대상) |

### M 코드

```
let
    // ========== PO 원본 (Invoiced 건만 = GRN 처리 완료) ==========
    PO_국내_Raw = Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content],
    PO_해외_Raw = Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content],

    PO_국내 = Table.SelectColumns(PO_국내_Raw, {"PO_ID", "SO_ID", "Line item", "Item qty", "Total ICO", "Status", "AX PO"}),
    PO_국내_Tagged = Table.AddColumn(PO_국내, "구분", each "국내"),

    PO_해외 = Table.SelectColumns(PO_해외_Raw, {"PO_ID", "SO_ID", "Line item", "Item qty", "Total ICO", "Status", "AX PO"}),
    PO_해외_Tagged = Table.AddColumn(PO_해외, "구분", each "해외"),

    PO_Combined = Table.Combine({PO_국내_Tagged, PO_해외_Tagged}),

    // Invoiced만 필터 (GRN 처리 완료 건)
    PO_Invoiced = Table.SelectRows(PO_Combined, each [Status] <> null and Text.StartsWith([Status], "Invoiced")),

    // AX PO 있는 건만 (AX에 입력되어 대사 가능한 건)
    PO_WithAX = Table.SelectRows(PO_Invoiced, each [#"AX PO"] <> null and [#"AX PO"] <> ""),

    // ========== SO 원본 (Model code 조인용) ==========
    SO_국내_Raw = Excel.CurrentWorkbook(){[Name="SO_국내"]}[Content],
    SO_해외_Raw = Excel.CurrentWorkbook(){[Name="SO_해외"]}[Content],

    SO_국내 = Table.SelectColumns(SO_국내_Raw, {"SO_ID", "Line item", "Model code"}),
    SO_해외 = Table.SelectColumns(SO_해외_Raw, {"SO_ID", "Line item", "Model code"}),
    SO_Combined = Table.Combine({SO_국내, SO_해외}),

    // ========== PO + SO 조인 (Model code 가져오기) ==========
    WithProject = Table.NestedJoin(PO_WithAX, {"SO_ID", "Line item"}, SO_Combined, {"SO_ID", "Line item"}, "SO_Data", JoinKind.LeftOuter),
    WithProjectExpanded = Table.ExpandTableColumn(WithProject, "SO_Data", {"Model code"}, {"Model code"}),

    // ========== DN 원본 (출고일 → Period 산정) ==========
    DN_국내_Raw = Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content],
    DN_해외_Raw = Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content],

    DN_국내 = Table.SelectColumns(DN_국내_Raw, {"SO_ID", "Line item", "출고일"}),
    DN_해외 = Table.SelectColumns(DN_해외_Raw, {"SO_ID", "Line item", "출고일"}),

    DN_Combined = Table.Combine({DN_국내, DN_해외}),

    // SO_ID + Line item 기준 출고일 집계 (같은 조합에 여러 DN이 있으면 최신 출고일)
    DN_Grouped = Table.Group(DN_Combined, {"SO_ID", "Line item"}, {
        {"출고일", each List.Max([출고일]), type date}
    }),

    // ========== PO + DN 조인 (출고일 가져오기) ==========
    WithDate = Table.NestedJoin(WithProjectExpanded, {"SO_ID", "Line item"}, DN_Grouped, {"SO_ID", "Line item"}, "DN_Data", JoinKind.LeftOuter),
    WithDateExpanded = Table.ExpandTableColumn(WithDate, "DN_Data", {"출고일"}, {"출고일"}),

    // 출고일 있는 건만 (Period 산정 가능한 건)
    WithValidDate = Table.SelectRows(WithDateExpanded, each
        [출고일] <> null and
        (try Date.Year([출고일]) otherwise null) <> null
    ),

    // ========== Period 추출 ==========
    WithPeriod = Table.AddColumn(WithValidDate, "Period", each
        Text.From(Date.Year([출고일])) & "-" & Text.PadStart(Text.From(Date.Month([출고일])), 2, "0"),
        type text),

    // ========== SO_ID 기준 그룹화 (AX PO, Model code는 LEFT JOIN으로 유지) ==========
    Grouped = Table.Group(WithPeriod, {"Period", "Model code", "AX PO", "구분", "SO_ID"}, {
        {"PO_ID", each Text.Combine(List.Distinct([PO_ID]), ", "), type text},
        {"건수", each Table.RowCount(_), Int64.Type},
        {"수량", each List.Sum([Item qty]), type number},
        {"금액", each List.Sum([Total ICO]), type number}
    }),

    // ========== 컬럼 순서 정리 ==========
    Reordered = Table.ReorderColumns(Grouped, {"Period", "Model code", "AX PO", "구분", "SO_ID", "PO_ID", "건수", "수량", "금액"}),

    // ========== 정렬 ==========
    Sorted = Table.Sort(Reordered, {
        {"Period", Order.Descending},
        {"Model code", Order.Ascending},
        {"AX PO", Order.Ascending},
        {"SO_ID", Order.Ascending}
    }),

    // ========== 타입 변환 ==========
    Result = Table.TransformColumnTypes(Sorted, {
        {"건수", Int64.Type},
        {"수량", Int64.Type},
        {"금액", Currency.Type}
    })
in
    Result
```

### 결과 예시

| Period | Model code | AX PO | 구분 | SO_ID | PO_ID | 건수 | 수량 | 금액 |
|--------|-------------------|-------|------|-------|-------|------|------|------|
| 2026-02 | PRJ-002 | P000003 | 국내 | SOD-0007 | POD-0007 | 2 | 10 | 4,000,000 |
| 2026-02 | PRJ-003 | P000005 | 해외 | SOO-0003 | POO-0003 | 1 | 5 | 3,500,000 |
| 2026-01 | PRJ-001 | P000001 | 국내 | SOD-0001 | POD-0001 | 2 | 8 | 5,000,000 |
| 2026-01 | PRJ-001 | P000001 | 국내 | SOD-0005 | POD-0005 | 2 | 7 | 2,500,000 |
| 2026-01 | PRJ-001 | P000002 | 국내 | SOD-0002 | POD-0002 | 2 | 10 | 8,000,000 |
| 2026-01 | PRJ-003 | P000004 | 해외 | SOO-0001 | POO-0001 | 3 | 20 | 6,000,000 |

### 용도

| 필터/분석 | 용도 |
|----------|------|
| Period = "2026-01" | 해당 월 마감 대사 (월별 필터링) |
| 특정 Model code | 프로젝트 단위 금액 합계 확인 (하위 PO들 합산) |
| 특정 AX PO | AX GRN 금액과 비교 (SO별 행 합산 = AX PO 금액) |
| 특정 SO_ID | 불일치 시 어떤 SO에서 차이인지 즉시 파악 |
| 구분별 소계 | 국내/해외 AP 분리 확인 |

**AX 대사 방법**:
```
PO_AX대사 (Period = 2026-01) 금액 합계  vs  AX D365 F&O GRN (2026-01월) 금액
→ AX PO 필터 후 SO별 행 합산 = PXXXXXX GRN 금액과 대사
→ 불일치 시 SO_ID별로 어디서 차이인지 바로 추적 가능
→ Model code로 동일 프로젝트 내 PO들을 묶어 합산 대사 가능
```

---

## PO_Industry

### 목적
- PO_국내 + PO_해외 통합
- SO에서 Industry code, Opportunity, Model code 조인
- PO_ID별 Industry code 파악용

### 결과 컬럼

| 컬럼 | 설명 |
|------|------|
| PO_ID | 발주 번호 |
| SO_ID | 주문 번호 (여러 건이면 콤마 연결) |
| Customer name | 고객명 (PO에서) |
| Opportunity | Opportunity 번호 (SO에서, 여러 건이면 콤마 연결) |
| Model code | AX 프로젝트 번호 (SO에서, 여러 건이면 콤마 연결) |
| Sector | 섹터 (SO에서, 여러 건이면 콤마 연결) |
| Industry code | 산업코드 (SO에서, 여러 건이면 콤마 연결) |
| 발주날짜 | 공장 발주 날짜 (가장 최근) |
| Status | PO 상태 (PO에서, 여러 건이면 콤마 연결) |
| 구분 | 국내/해외 |

### M 코드

```
let
    // ========== PO 원본 ==========
    PO_국내_Raw = Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content],
    PO_해외_Raw = Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content],

    PO_국내 = Table.SelectColumns(PO_국내_Raw, {"PO_ID", "SO_ID", "Line item", "Customer name", "공장 발주 날짜", "Status"}),
    PO_국내_Tagged = Table.AddColumn(PO_국내, "구분", each "국내"),
    PO_해외 = Table.SelectColumns(PO_해외_Raw, {"PO_ID", "SO_ID", "Line item", "Customer name", "공장 발주 날짜", "Status"}),
    PO_해외_Tagged = Table.AddColumn(PO_해외, "구분", each "해외"),
    PO_Combined = Table.Combine({PO_국내_Tagged, PO_해외_Tagged}),

    // ========== SO (Industry code, Opportunity, Model code) ==========
    SO_국내_Raw = Excel.CurrentWorkbook(){[Name="SO_국내"]}[Content],
    SO_해외_Raw = Excel.CurrentWorkbook(){[Name="SO_해외"]}[Content],
    SO_국내 = Table.SelectColumns(SO_국내_Raw, {"SO_ID", "Line item", "Sector", "Industry code", "Opportunity", "Model code"}),
    SO_해외 = Table.SelectColumns(SO_해외_Raw, {"SO_ID", "Line item", "Sector", "Industry code", "Opportunity", "Model code"}),
    SO_Combined = Table.Distinct(Table.Combine({SO_국내, SO_해외}), {"SO_ID", "Line item"}),

    // ========== PO + SO 조인 (SO_ID + Line item 기준) ==========
    WithSO = Table.NestedJoin(PO_Combined, {"SO_ID", "Line item"}, SO_Combined, {"SO_ID", "Line item"}, "SO_Data", JoinKind.LeftOuter),
    WithSOExpanded = Table.ExpandTableColumn(WithSO, "SO_Data", {"Sector", "Industry code", "Opportunity", "Model code"}, {"Sector", "Industry code", "Opportunity", "Model code"}),

    // ========== PO_ID 기준 그룹화 ==========
    Grouped = Table.Group(WithSOExpanded, {"PO_ID"}, {
        {"SO_ID", each Text.Combine(List.Distinct([SO_ID]), ", "), type text},
        {"Customer name", each List.First([Customer name]), type text},
        {"Opportunity", each Text.Combine(List.Distinct(List.RemoveNulls([Opportunity])), ", "), type text},
        {"Model code", each Text.Combine(List.Distinct(List.RemoveNulls(List.Transform([Model code], Text.From))), ", "), type text},
        {"Sector", each Text.Combine(List.Distinct(List.RemoveNulls([Sector])), ", "), type text},
        {"Industry code", each Text.Combine(List.Distinct(List.RemoveNulls(List.Transform([Industry code], Text.From))), ", "), type text},
        {"발주날짜", each List.Max([공장 발주 날짜]), type date},
        {"Status", each Text.Combine(List.Distinct(List.RemoveNulls([Status])), ", "), type text},
        {"구분", each List.First([구분]), type text}
    }),

    // ========== 정렬 + 타입 ==========
    Sorted = Table.Sort(Grouped, {{"PO_ID", Order.Ascending}}),
    Result = Table.TransformColumnTypes(Sorted, {{"발주날짜", type date}})
in
    Result
```

### 결과 예시

| PO_ID | SO_ID | Customer name | Opportunity | Model code | Sector | Industry code | 발주날짜 | Status | 구분 |
|-------|-------|--------------|-------------|-------------------|--------|--------------|----------|--------|------|
| ND-0001 | SOD-2026-0001 | 삼성전자 | OPP-001 | PRJ-2026-001 | CPI | Power | 2026-01-15 | Invoiced P01 | 국내 |
| ND-0002 | SOD-2026-0002 | LG전자 | OPP-002 | PRJ-2026-002 | W&P | Water | 2026-01-20 | Confirmed | 국내 |
| NO-0001 | SOO-2026-0001 | ABC Corp | OPP-003 | PRJ-2026-003 | O&G | Oil & Gas | 2026-02-01 | Sent | 해외 |

---

## PO_미출고

### 목적
- Invoiced인데 DN에 출고일이 없는 건 상세 목록
- 데이터 불일치 점검용

### 결과 컬럼
| 컬럼 | 설명 |
|------|------|
| PO_ID | 발주 번호 |
| SO_ID | 주문 번호 |
| Customer name | 고객명 (SO에서) |
| Customer PO | 고객 발주번호 (SO에서) |
| Item name | 아이템명 |
| Item qty | 수량 |
| Total ICO | 금액 |
| Status | PO 상태 (Invoiced P01/P02 등) |
| 구분 | 국내/해외 |

### M 코드

```
let
    // ========== PO 원본 (Invoiced 건만) ==========
    PO_국내_Raw = Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content],
    PO_해외_Raw = Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content],

    PO_국내 = Table.SelectColumns(PO_국내_Raw, {"PO_ID", "SO_ID", "Item name", "Line item", "Item qty", "Total ICO", "Status"}),
    PO_국내_Tagged = Table.AddColumn(PO_국내, "구분", each "국내"),

    PO_해외 = Table.SelectColumns(PO_해외_Raw, {"PO_ID", "SO_ID", "Item name", "Line item", "Item qty", "Total ICO", "Status"}),
    PO_해외_Tagged = Table.AddColumn(PO_해외, "구분", each "해외"),

    PO_Combined = Table.Combine({PO_국내_Tagged, PO_해외_Tagged}),

    // Invoiced P01, P02 등만 필터
    PO_Invoiced = Table.SelectRows(PO_Combined, each [Status] <> null and Text.StartsWith([Status], "Invoiced")),

    // ========== SO 원본 (Customer 정보) ==========
    // SO_ID 기준으로 조인 (같은 SO_ID = 같은 고객)
    SO_국내_Raw = Excel.CurrentWorkbook(){[Name="SO_국내"]}[Content],
    SO_해외_Raw = Excel.CurrentWorkbook(){[Name="SO_해외"]}[Content],

    SO_국내 = Table.SelectColumns(SO_국내_Raw, {"SO_ID", "Customer name", "Customer PO"}),
    SO_해외 = Table.SelectColumns(SO_해외_Raw, {"SO_ID", "Customer name", "Customer PO"}),
    SO_Combined = Table.Distinct(Table.Combine({SO_국내, SO_해외}), {"SO_ID"}),

    // ========== DN 원본 (출고일) ==========
    DN_국내_Raw = Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content],
    DN_해외_Raw = Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content],

    DN_국내 = Table.SelectColumns(DN_국내_Raw, {"SO_ID", "Line item", "출고일"}),
    DN_해외 = Table.SelectColumns(DN_해외_Raw, {"SO_ID", "Line item", "출고일"}),

    DN_Combined = Table.Combine({DN_국내, DN_해외}),

    // SO_ID + Line item 기준 출고일 집계
    DN_Grouped = Table.Group(DN_Combined, {"SO_ID", "Line item"}, {
        {"출고일", each List.Max([출고일]), type date}
    }),

    // ========== PO + SO 조인 (Customer 정보, SO_ID 기준) ==========
    WithCustomer = Table.NestedJoin(PO_Invoiced, {"SO_ID"}, SO_Combined, {"SO_ID"}, "SO_Data", JoinKind.LeftOuter),
    WithCustomerExpanded = Table.ExpandTableColumn(WithCustomer, "SO_Data", {"Customer name", "Customer PO"}, {"Customer name", "Customer PO"}),

    // ========== PO + DN 조인 (출고일) ==========
    WithDate = Table.NestedJoin(WithCustomerExpanded, {"SO_ID", "Line item"}, DN_Grouped, {"SO_ID", "Line item"}, "DN_Data", JoinKind.LeftOuter),
    WithDateExpanded = Table.ExpandTableColumn(WithDate, "DN_Data", {"출고일"}, {"출고일"}),

    // ========== 미출고 건만 필터 ==========
    // 출고일이 null이거나 유효하지 않은 날짜인 경우
    미출고 = Table.SelectRows(WithDateExpanded, each
        [출고일] = null or
        (try Date.Year([출고일]) otherwise null) = null
    ),

    // 출고일 컬럼 제거 (어차피 null) + 컬럼 순서 정리
    Cleaned = Table.RemoveColumns(미출고, {"출고일"}),
    Result = Table.ReorderColumns(Cleaned, {"PO_ID", "SO_ID", "Customer name", "Customer PO", "Item name", "Line item", "Item qty", "Total ICO", "Status", "구분"})
in
    Result
```

### 결과 예시

| PO_ID | SO_ID | Customer name | Customer PO | Item name | Item qty | Total ICO | Status | 구분 |
|-------|-------|---------------|-------------|-----------|----------|-----------|--------|------|
| POD-2026-0025 | SOD-2026-0010 | 삼성전자 | 4500012345 | IQ3 | 5 | 2,500,000 | Invoiced P01 | 국내 |
| POO-2026-0008 | SOO-2026-0003 | ABC Corp | PO-2026-001 | CVA | 2 | 1,600,000 | Invoiced P01 | 해외 |

### 점검 방법
1. 결과 목록의 SO_ID + Line item 확인
2. DN 시트에서 해당 조합 검색
3. 불일치 원인 파악:
   - DN 기록 누락 → DN 시트에 추가
   - Line item 불일치 → PO 또는 DN 수정
   - 출고일 비어있음 → DN 시트에서 출고일 입력

---

## PO_출고

### 목적
- **출고(DN)는 됐는데 그 PO에 계상 라인(Invoiced·외주비)이 하나도 없는** 건 — `PO_미출고`의 **정반대 방향** 점검.
- 업무 불변식: **DN에 출고기록이 있으면 해당 PO는 반드시 계상(Invoiced 또는 외주비)되어 있어야 한다.** 이를 어긴(출고됐는데 매입계상 누락) PO를 잡는다. `외주비`(외주가공 매입계상)도 Invoiced와 함께 '계상됨'으로 본다.
- **PO_ID 단위 판정**(라인 단위 아님). 분할 출고는 한 PO가 출고 배치별로 여러 행(계상 분 + Confirmed 잔여)으로 쪼개지는데, 출고된 분이 계상됐으면 위반이 아니다. 또 PO는 부속을 한 라인에 묶고(`SA09X-MA + ADAPTER`) DN은 라인을 분리하므로, Line item 단위 PO↔DN 비교는 번들·반품(Credit Note)·분할 잔여 때문에 **오탐**을 낳는다 → 라인이 아니라 **PO_ID + 출고 여부 + 계상 라인 유무**로 판정한다.
- 국내·해외 모두 대상.

### 결과 컬럼
| 컬럼 | 설명 |
|------|------|
| PO_ID | 발주 번호 (이 PO 전체가 위반) |
| SO_ID | 주문 번호 |
| Customer name / Customer PO | 고객 정보 (SO에서) |
| Item name | 아이템명 (PO) |
| Line item | 라인 번호 |
| Item qty | PO 수량 |
| Total ICO | PO 금액 |
| Status | PO 상태 (이 PO엔 계상 라인이 전무 — Confirmed/Sent/Open 등; Invoiced·외주비는 계상으로 제외) |
| 출고일 | 해당 PO의 DN 출고일 (최댓값) |
| DN_ID | 해당 PO의 출고 DN 번호 |
| 구분 | 국내/해외 |

### M 코드

```
let
    // ========== PO 원본 (국내+해외, Cancelled 제외) — Buffer로 1회만 읽음 ==========
    PO_국내 = Table.AddColumn(Table.SelectColumns(Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content], {"PO_ID", "SO_ID", "Item name", "Line item", "Item qty", "Total ICO", "Status"}), "구분", each "국내"),
    PO_해외 = Table.AddColumn(Table.SelectColumns(Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content], {"PO_ID", "SO_ID", "Item name", "Line item", "Item qty", "Total ICO", "Status"}), "구분", each "해외"),
    PO_Active = Table.Buffer(Table.SelectRows(Table.Combine({PO_국내, PO_해외}), each [Status] <> null and [Status] <> "Cancelled")),

    // ========== PO_ID별 '계상 라인 보유' 여부 → 계상 전무 PO만 ==========
    // '계상' = Invoiced(매입계상) 또는 외주비(외주가공 매입계상). 둘 다 비용으로 잡힌 상태.
    // 분할 출고: 한 PO가 출고 배치별로 여러 행(계상 분 + Confirmed 잔여)으로 나뉜다.
    // 출고된 분이 계상됐으면 그 PO엔 계상 라인이 있으므로 위반 아님 → PO_ID 단위로 본다.
    PO_ID_Inv = Table.Group(PO_Active, {"PO_ID"}, {
        {"has_billed", each List.AnyTrue(List.Transform([Status], each _ <> null and (Text.StartsWith(_, "Invoiced") or Text.StartsWith(_, "외주비")))), type logical}
    }),
    PO_무계상 = Table.SelectColumns(Table.SelectRows(PO_ID_Inv, each [has_billed] = false), {"PO_ID"}),

    // ========== DN에서 '출고(출고일)된 PO_ID' + 출고일/DN_ID (국내=PO_ID, 해외=RCK PO) ==========
    DN_국내 = Table.SelectColumns(Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content], {"PO_ID", "출고일", "DN_ID"}),
    DN_해외 = Table.RenameColumns(Table.SelectColumns(Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content], {"RCK PO", "출고일", "DN_ID"}), {{"RCK PO", "PO_ID"}}),
    DN_출고 = Table.SelectRows(Table.Combine({DN_국내, DN_해외}), each [출고일] <> null and (try Date.Year([출고일]) otherwise null) <> null),
    DN_Grouped = Table.Group(DN_출고, {"PO_ID"}, {
        {"출고일", each List.Max([출고일]), type date},
        {"DN_ID", each Text.Combine(List.Distinct(List.Transform(List.RemoveNulls([DN_ID]), each Text.From(_))), ", "), type text}
    }),

    // ========== 위반 PO = (출고됨) ∩ (계상 라인 전무) — 조인(집합연산)으로 ==========
    // List.Contains/List.Intersect를 행마다 평가하면 위 계산이 매 행 반복돼 매우 느려진다 → Inner Join으로 대체.
    위반PO = Table.RemoveColumns(Table.NestedJoin(DN_Grouped, {"PO_ID"}, PO_무계상, {"PO_ID"}, "chk", JoinKind.Inner), {"chk"}),

    // ========== 위반 PO의 라인 추출 (Inner Join) + 출고 근거(DN) 부착 ==========
    위반라인 = Table.ExpandTableColumn(Table.NestedJoin(PO_Active, {"PO_ID"}, 위반PO, {"PO_ID"}, "DN_Data", JoinKind.Inner), "DN_Data", {"출고일", "DN_ID"}, {"출고일", "DN_ID"}),

    // ========== 고객 정보(SO) 부착 ==========
    SO_국내 = Table.SelectColumns(Excel.CurrentWorkbook(){[Name="SO_국내"]}[Content], {"SO_ID", "Customer name", "Customer PO"}),
    SO_해외 = Table.SelectColumns(Excel.CurrentWorkbook(){[Name="SO_해외"]}[Content], {"SO_ID", "Customer name", "Customer PO"}),
    SO_Combined = Table.Distinct(Table.Combine({SO_국내, SO_해외}), {"SO_ID"}),
    WithCustomer = Table.ExpandTableColumn(Table.NestedJoin(위반라인, {"SO_ID"}, SO_Combined, {"SO_ID"}, "SO_Data", JoinKind.LeftOuter), "SO_Data", {"Customer name", "Customer PO"}, {"Customer name", "Customer PO"}),

    Result = Table.ReorderColumns(WithCustomer, {"PO_ID", "SO_ID", "Customer name", "Customer PO", "Item name", "Line item", "Item qty", "Total ICO", "Status", "출고일", "DN_ID", "구분"})
in
    Result
```

### 점검 방법
1. 결과의 PO_ID 확인 → 출고·매입계상이 끝난 건이면 PO 시트 Status를 **Invoiced(Pxx)**(외주가공이면 **외주비**)로 갱신.
2. 아직 출고 전인데 DN이 잡혔다면(오발행 DN) DN 시트를 점검.
3. **일부만 계상(Invoiced·외주비)된 PO(분할 출고)는 이 쿼리에 안 잡힌다** — 출고된 배치가 계상됐으면 정상이기 때문(예: SOD-2026-0301은 9개 라인 모두 출고분=Invoiced분으로 일치 → 정상). 라인/수량 단위로 더 파고들면 번들(`A + ADAPTER`)·반품(Credit Note)·분할 잔여 때문에 오탐이 급증하는 것이 검증돼, PO_ID 단위가 가장 견고하다.

### 결과 (현재 데이터, 2026-06-29 기준)
- 총 **23라인 / 1건**: `SOO-2026-0165`(NO-0165, Confirmed 23라인).
- `외주비`를 계상으로 포함하면서 `SOD-2026-0046`(ND-0046, 외주비 P02 1라인)은 **계상됨으로 제외**됨(이전 2건 → 1건).
- 제외(정상 판정) 검증: 분할 출고 8건(SOD-2026-0301 등) + 번들(SOD-2026-0231·0331) + 반품 Credit Note(SOO-2026-0032) — 모두 출고분이 계상됐거나 PO↔DN 라인 구조 차이라 **실제 위반 아님**. 라인 단위로 보면 이들이 오탐으로 잡히므로 PO_ID 단위 판정을 채택.

---

## AX_매출대사

### 목적
- **AX ERP 매출 vs NOAH 엑셀 매출** 대사를 위한 **NOAH(엑셀) 측 매출 집계**. 출고(DN) 라인을 매출인식 기준으로 정리한다. AX 실적과의 비교는 Power Pivot 관계 또는 `reconcile_so.py`에서 수행하며, 이 쿼리는 **매출측만** 산출한다.
- **매출인식일**: 국내 = **세금계산서 발행일**(없으면 **선수금 세금계산서가 있고 출고된 건은 출고일** — 선청구 후 출고 시점에 수익인식), 해외 = **선적일**. 출고일이 아니라 이 날짜로 매출월을 귀속한다(고객 cutoff로 출고월과 매출월이 갈리는 경우 대응). 기준일이 비어 있으면(세금계산서·선수금 모두 없이 출고만 됐거나 미선적) `매출인식 = N`으로 표시하고 행은 남긴다.
- **해외 KRW 환산**: FX 시트의 **선적월 환율**을 적용해 `외화금액 × 선적월환율`로 재환산한다. 시트에 이미 있는 `Total Sales KRW`는 `기존_KRW`로 병기하고, 재환산값과의 차이를 `KRW차이`로 노출한다(주문시점 환율 등으로 계산된 기존값과의 괴리 감지).
- **IP 여부**(= AX Invoice Proposal 등록 여부)는 **필터하지 않고 컬럼으로 표시**한다(값은 `Y`로 정규화). '매출인식(세금계산서/선적)'과 'AX 등록(IP)'을 각각 축으로 두어, 인식됐는데 AX 미등록인 건을 대사에서 잡는다.
- **Grain: DN 라인 단위**. `AX Project` + `매출월`로 피벗/합산하여 AX 매출 실적과 대사한다.

### 결과 컬럼
| 컬럼 | 설명 |
|------|------|
| 구분 | 국내 / 해외 |
| AX Project | AX 프로젝트 번호 (국내 `AX Project no` / 해외 `AX Project number` 통합) |
| SO_ID / Line item | 주문번호 / 라인 (SO·FX·매출월 조인키) |
| Customer code | 사업자등록번호 (국내=DN, 해외=SO_해외 조인) |
| Customer name | 고객명 |
| OS name | OneStream Item 분류 (SO에서 `SO_ID`+`Line item` 조인) |
| Item name | 품목명 (DN `Item`) |
| Qty | 출고 수량 |
| Currency | 통화 (KRW/USD/EUR/GBP) |
| 외화금액 | 원 통화 매출액 (해외만; 국내는 공란) |
| 재환산환율 | 적용한 선적월 FX 환율 (해외 비KRW) |
| **매출금액_KRW** | **매출액(KRW) — 국내=`Total Sales`, 해외=`외화금액 × 선적월환율`** |
| 기존_KRW | 시트에 저장된 KRW (국내 `Total Sales` / 해외 `Total Sales KRW`) |
| KRW차이 | `매출금액_KRW − 기존_KRW` (해외 환율차 진단) |
| IP 여부 | AX Invoice Proposal 등록 여부 (`Y` 정규화, 미등록=공란) |
| 매출인식 | `Y`=매출월 확정(세금계산서/선적) / `N`=출고했으나 미인식 |
| 매출일 | 매출인식일 (국내 세금계산서 발행일 · 선수금건은 출고일 / 해외 선적일) |
| 매출연월 | `2026-03` (YYYY-MM) |
| 매출월 | `P03` (Pxx) |
| DN_ID / 출고일 | 출고 참조 |

### M 코드

```
let
    // ========== DN 국내 (매출인식일: 세금계산서 발행일 → 없으면 선수금O·출고일O이면 출고일) ==========
    DN_국내_Sel = Table.SelectColumns(Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content],
        {"DN_ID", "SO_ID", "Line item", "AX Project no", "Business registration number", "Customer name", "Item", "Qty", "Currency", "Unit Price", "Total Sales", "세금계산서 발행일", "선수금 세금계산서 발행일", "출고일", "IP 여부"}),
    // 선수금(선청구) 세금계산서만 있고 출고된 건은 수익인식 시점인 출고일 월을 매출로 본다.
    DN_국내_Rev = Table.AddColumn(DN_국내_Sel, "매출일", each
        if [세금계산서 발행일] <> null then [세금계산서 발행일]
        else if [선수금 세금계산서 발행일] <> null and [출고일] <> null then [출고일]
        else null),
    DN_국내_Std = Table.AddColumn(Table.AddColumn(
        Table.RenameColumns(Table.RemoveColumns(DN_국내_Rev, {"세금계산서 발행일", "선수금 세금계산서 발행일"}), {
            {"AX Project no", "AX Project"},
            {"Business registration number", "Customer code"},
            {"Item", "Item name"},
            {"Total Sales", "기존_KRW"}
        }),
        "구분", each "국내"),
        "외화금액", each null),

    // ========== DN 해외 (매출인식일 = 선적일, KRW = FX 선적월 재환산) ==========
    DN_해외_Sel = Table.SelectColumns(Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content],
        {"DN_ID", "SO_ID", "Line item", "AX Project number", "Customer name", "Item", "Qty", "Currency", "Unit Price", "Total Sales", "Total Sales KRW", "선적일", "출고일", "IP 여부"}),
    DN_해외_Std = Table.AddColumn(Table.AddColumn(
        Table.RenameColumns(DN_해외_Sel, {
            {"AX Project number", "AX Project"},
            {"Item", "Item name"},
            {"선적일", "매출일"},
            {"Total Sales", "외화금액"},
            {"Total Sales KRW", "기존_KRW"}
        }),
        "구분", each "해외"),
        "Customer code", each null),

    // ========== 국내 + 해외 통합 ==========
    DN_All = Table.Combine({DN_국내_Std, DN_해외_Std}),

    // ========== IP 여부 정규화 (Y / y / "Y(6월 DN)" → "Y", 빈값 → null) ==========
    DN_IPNorm = Table.TransformColumns(DN_All, {
        {"IP 여부", each if _ = null then null else (let s = Text.Trim(Text.From(_)) in if s = "" then null else if Text.StartsWith(Text.Upper(s), "Y") then "Y" else s), type text}
    }),

    // ========== 매출일 → 날짜화 + 매출인식 / 매출연월(YYYY-MM) / 매출월(Pxx) 파생 ==========
    DN_Date = Table.RenameColumns(Table.RemoveColumns(
        Table.AddColumn(DN_IPNorm, "매출일_d", each try Date.From([매출일]) otherwise null, type date),
        {"매출일"}), {{"매출일_d", "매출일"}}),
    DN_Period = Table.AddColumn(Table.AddColumn(Table.AddColumn(DN_Date,
        "매출인식", each if [매출일] = null then "N" else "Y", type text),
        "매출연월", each if [매출일] = null then null else Text.From(Date.Year([매출일])) & "-" & Text.PadStart(Text.From(Date.Month([매출일])), 2, "0"), type text),
        "매출월", each if [매출일] = null then null else "P" & Text.PadStart(Text.From(Date.Month([매출일])), 2, "0"), type text),

    // ========== FX 시트 언피벗 (가로 → 세로: Currency + 환율월 + 환율) ==========
    FX_Unpiv = Table.UnpivotOtherColumns(Table.RenameColumns(Excel.CurrentWorkbook(){[Name="FX"]}[Content], {{"FX", "Currency"}}), {"Currency"}, "환율월", "환율"),
    FX_Clean = Table.Buffer(Table.SelectRows(FX_Unpiv, each [Currency] <> null and [환율] <> null and Text.Length(Text.From([환율월])) = 7 and Text.Contains(Text.From([환율월]), "-"))),

    // ========== 해외 KRW = 외화금액 × 선적월 환율 (국내는 이미 KRW) ==========
    Joined_FX = Table.ExpandTableColumn(Table.NestedJoin(DN_Period, {"Currency", "매출연월"}, FX_Clean, {"Currency", "환율월"}, "FX_Match", JoinKind.LeftOuter), "FX_Match", {"환율"}, {"환율"}),
    DN_KRW = Table.AddColumn(Table.AddColumn(Table.AddColumn(Joined_FX,
        "재환산환율", each if [구분] = "해외" and [Currency] <> "KRW" then [환율] else null, type number),
        "매출금액_KRW", each
            if [구분] = "국내" then [기존_KRW]
            else if [Currency] = "KRW" then [외화금액]
            else if [환율] <> null then Number.Round([외화금액] * [환율], 0)
            else null, type number),
        "KRW차이", each if [매출금액_KRW] <> null and [기존_KRW] <> null then [매출금액_KRW] - [기존_KRW] else null, type number),

    // ========== SO 조인 (OS name, 해외 고객코드) — SO_ID + Line item ==========
    SO_Lookup = Table.Buffer(Table.Distinct(Table.Combine({
        Table.SelectColumns(Excel.CurrentWorkbook(){[Name="SO_국내"]}[Content], {"SO_ID", "Line item", "OS name", "Business registration number"}),
        Table.SelectColumns(Excel.CurrentWorkbook(){[Name="SO_해외"]}[Content], {"SO_ID", "Line item", "OS name", "Business registration number"})
    }), {"SO_ID", "Line item"})),
    Joined_SO = Table.ExpandTableColumn(Table.NestedJoin(DN_KRW, {"SO_ID", "Line item"}, SO_Lookup, {"SO_ID", "Line item"}, "SO_Match", JoinKind.LeftOuter), "SO_Match", {"OS name", "Business registration number"}, {"OS name", "SO_고객코드"}),

    // ========== 고객코드 채움 (국내=DN 사업자번호, 해외=SO 조인) ==========
    DN_Code = Table.RenameColumns(Table.RemoveColumns(
        Table.AddColumn(Joined_SO, "고객코드_f", each if [Customer code] <> null and Text.Trim(Text.From([Customer code])) <> "" then [Customer code] else [SO_고객코드]),
        {"Customer code", "SO_고객코드"}), {{"고객코드_f", "Customer code"}}),

    // ========== 최종 컬럼 선택/정렬 + 타입 캐스팅 + 에러 방어 ==========
    Final = Table.SelectColumns(DN_Code, {"구분", "AX Project", "SO_ID", "Line item", "Customer code", "Customer name", "OS name", "Item name", "Qty", "Currency", "외화금액", "재환산환율", "매출금액_KRW", "기존_KRW", "KRW차이", "IP 여부", "매출인식", "매출일", "매출연월", "매출월", "DN_ID", "출고일"}),
    Typed = Table.TransformColumnTypes(Final, {
        {"Qty", Int64.Type},
        {"외화금액", type number},
        {"재환산환율", type number},
        {"매출금액_KRW", Currency.Type},
        {"기존_KRW", Currency.Type},
        {"KRW차이", Currency.Type},
        {"출고일", type date}
    }),
    Result = Table.ReplaceErrorValues(Typed, List.Transform(Table.ColumnNames(Typed), each {_, null}))
in
    Result
```

### 점검 방법
1. `AX Project` + `매출월`로 피벗(합계 `매출금액_KRW`) → AX 매출 실적과 월별·프로젝트별 비교.
2. `매출인식 = N`: 출고했으나 세금계산서 미발행(국내)·미선적(해외) → 아직 매출 미인식. 고객 cutoff/선적 지연 확인.
3. `IP 여부`가 공란인데 `매출인식 = Y`: 매출은 인식됐는데 AX Invoice Proposal 미등록 의심 → AX 등록 점검.
4. `KRW차이`가 큰 해외 건: 시트의 기존 KRW가 선적월 환율과 다른 환율로 계산됨 → 환율/금액 재확인.
5. **전제**: `FX`/`DN_국내`/`DN_해외`/`SO_국내`/`SO_해외`가 통합 문서 안에 **표(또는 이름 정의)** 로 존재해야 `Excel.CurrentWorkbook`이 인식한다. `FX`는 헤더가 `FX | 2026-01 | …`로 승격돼 있어야 하며(아니면 첫 스텝에 `Table.PromoteHeaders` 추가), 월 컬럼명은 `YYYY-MM` 텍스트여야 매출연월과 조인된다.

### 결과 (현재 데이터, 2026-07-01 기준)
- 총 **1,839 라인** (국내 1,161 / 해외 678). 매출인식 **Y 1,765 / N 74** (국내 10·해외 64 미인식).
- 국내 미인식은 세금계산서·선수금 모두 없는 **10건**만 남는다(무상공급 0원·반품 상쇄·당일 출고분). 선수금 세금계산서 + 출고 건 **25건**은 출고일 월로 인식됨.
- OS name·Customer code 조인 결측 **0** (`SO_ID`+`Line item` 완전 매칭 — 해외 고객코드까지 SO에서 채워짐).
- 해외 비KRW 인식건 614 중 **459건이 선적월 환율 재환산 시 기존 시트 KRW와 1,000원 초과 차이**(총 |차이| 8,843만 원) → FX 재적용의 실효 확인. 매출인식 Y·비KRW인데 FX 환율 없는 건 **0**.
- 국내 **출고월 ≠ 세금계산서 발행월 64건** — 현행 `reconcile_so.py`(출고일 기준)와 월귀속이 갈리는 지점.

---

## Order_Book

### 목적
- SO_국내 + SO_해외의 수주 데이터를 **월별 롤링 원장**으로 표현
- AX2009 Order Book과 동일한 형식: **건별 × Period**
- 수주잔고(Backlog) 흐름을 Period별로 추적
- SO-DN 금액 불일치 자동 감지
- **SO_ID + OS name + Expected delivery date 기준 그룹화** (같은 제품의 Line item 합산, 납기일이 다르면 구분)

### 직관적 이해

**택배 추적**이라고 생각하면 됩니다.
주문하면 "배송중"이 되고, 수령하면 "배송완료"가 되듯이,
Order_Book은 그걸 **금액 단위로, 매월** 하는 겁니다.

```
1월: 750만원어치 주문 들어옴, 그중 400만원 출고함
     → 아직 350만원어치 안 나감 (Backlog)

2월: 새 주문 없음, 350만원 출고함
     → 남은 거 없음 (소진)
```

이게 전부입니다. **"이번 달 기준으로 얼마나 밀려있나?"**를 보는 것.

그런데 SO 원본은 이렇게 생겼습니다:

```
SOD-0001, 1월 수주, 750만원   ← 이 1줄이 끝
```

이걸로는 "2월에 얼마 남았지?"를 볼 수가 없습니다. 1월 행밖에 없으니까.
그래서 **달력처럼 펼칩니다**:

```
SOD-0001 × 1월: 들어옴 750만, 나감 400만 → 남음 350만
SOD-0001 × 2월: 들어옴 0,     나감 350만 → 남음 0
```

1줄짜리 주문을 월별로 복제해서 빈 칸을 만들고, 각 칸에 Input/Output을 채운 뒤, 통장처럼 잔고를 굴리는 것 — 이것이 Order_Book 쿼리의 본질입니다.

**6단계 요약:**

| 단계 | 하는 일 | 비유 |
|------|---------|------|
| ① 마지막 매출월 붙이기 | "이 주문 언제 끝나?" 끝점 파악 | 달력을 어디까지 펼칠지 |
| ② Period 확장 | 1줄을 등록~현재월까지 복제 | 빈 달력 만들기 |
| ③ Input 채우기 | 등록월에만 수주 금액 기록 | 입금 기록 |
| ④ Output 채우기 | DN 매출(세금계산서/선적)을 해당 월에 매칭 + 환율 재환산 | 출금 기록 |
| ⑤ Line item 합치기 | 같은 제품+납기일끼리 합산 | 정리 |
| ⑥ 잔고 계산 | Start + Input - Output = Ending | 통장 잔고 |

**①~④ = 빈 달력 만들어서 채우기, ⑤ = 정리, ⑥ = 통장 잔고 계산**

### 개념

```
수주잔고 Order Book = 오더의 흐름을 월별로 추적하는 원장

  Start (전월 이월)
+ Input (당월 수주 = SO 등록, 수주시점 환율)
- Output (당월 매출 = DN 출고/선적, 해외는 선적월 환율 재환산)
+ Variance (환율 재평가 = Output 재환산액 − 시트 KRW)
= Ending (당월 잔고 → 다음 달 Start로 이월)
```

**Variance가 왜 필요한가** — Input과 Output의 환율 시점이 다르기 때문입니다.

```
SOO-2026-0188: USD 1,276 · 6월 수주 · 7월 선적

P06: Input  = 1,914,046  (6월 환율 1,500.036 — SO 시트의 Sales amount KRW)
P07: Output = 1,976,024  (7월 환율 1,548.608 — 매출 인식 환율)

Variance 없이 계산하면 → Ending = 1,914,046 − 1,976,024 = -61,978  ← 있지도 않은 음수 잔고
Variance 넣으면      → Ending = 1,914,046 − 1,976,024 + 61,978 = 0  ← 소진
                                                        ↑
                                            환율 상승분(61,978)만큼 잔고를 재평가한 뒤 소진
```

즉 **Output은 매출 금액(선적월 환율)**, **Variance는 그 환율 재평가분**, **Ending은 재환산 도입 전과 동일**합니다.
환율 노이즈가 Variance로 빠지므로 "Ending ≠ 0 = SO-DN 금액 불일치" 진단이 그대로 유지됩니다.

```
SOD-0001, IQ3를 1월에 수주, 2월에 출고:

P01: Start=0    + Input=500만 - Output=0      = Ending=500만  (Backlog)
P02: Start=500만 + Input=0    - Output=480만  = Ending=20만   (SO-DN 차이)
         ↑                          ↑                  ↑
    P01 Ending             DN 실제 매출 금액    점검 대상 (0이 아님)
```

### 동작 방식

- **버튼/마감 작업 없음** — Ctrl+Alt+F5 새로고침 시 전체 재계산
- SO/DN 원본 데이터에서 매번 처음부터 계산하는 **뷰(View)**
- 스냅샷 저장 없음 (과도기적 사용, 과거 데이터 수정 시 소급 반영)
- Input = SO의 `Sales amount KRW` (수주 금액, 수주시점 환율)
- Output = 국내 `Total Sales` / 해외 `외화금액 × 선적월 환율` (실제 매출 금액) — **`AX_매출대사`와 동일 산식**
- Variance = 해외 `Output 재환산액 − 시트 Total Sales KRW` (환율 재평가분). 국내는 항상 0
- **Output 귀속월 = 매출 인식월** (`AX_매출대사`와 동일):
  - 국내 = **세금계산서 발행월** → `N/A`(발행 불필요: 무상공급·FOC·반품)면 출고월 → 선수금 세금계산서+출고면 출고월 → 아무것도 없으면 **매출 미인식**
  - 해외 = **선적월**
- **미인식 = Backlog 잔류**: 출고했지만 세금계산서가 아직 없으면(월합세금계산서 대기 등) Output이 잡히지 않고 Backlog에 남는다. 발행되면 새로고침만으로 그 달 Output이 된다
- **분할 출고 대응**: 같은 SO+Line item에 DN이 여러 건이면 각 매출월에 해당 수량/금액 배분

### SO-DN 금액 차이 감지

```
출고 후 Ending = 0  → 정상 (SO 금액 = DN 금액)
출고 후 Ending ≠ 0  → SO-DN 금액 불일치 → 데이터 점검 필요

예: SO 수주 500만, DN 출고 480만 → Ending = 20만 (단가 조정 발생?)
```

### OS name + Expected delivery date 기준 그룹화

Line item 단위로 처리하면 행 수가 과도하게 많아지므로, **OS name이 같은 Line item을 합산**하여 표시합니다.
단, **Expected delivery date가 다르면 별도 행**으로 구분합니다.

```
SO_ID = SOD-0001
  Line item 1: IQ3  (OS name: IQ3, 납기: 2/20)  → qty 5, amount 250만
  Line item 2: IQ3  (OS name: IQ3, 납기: 2/20)  → qty 3, amount 150만
  Line item 3: IQ3  (OS name: IQ3, 납기: 3/10)  → qty 2, amount 100만
  Line item 4: CVA  (OS name: CVA, 납기: 2/20)  → qty 2, amount 100만

→ 그룹화 후:
  SOD-0001 × IQ3 × 2/20: qty 8, amount 400만  (Line 1+2 합산, 같은 납기)
  SOD-0001 × IQ3 × 3/10: qty 2, amount 100만  (Line 3, 납기 다름 → 별도)
  SOD-0001 × CVA × 2/20: qty 2, amount 100만  (Line 4 단독)
```

- **그룹화 키**: SO_ID + OS name + Expected delivery date + Period
- **합산 필드**: qty, amount (Input/Output 모두)
- **대표값 필드**: Customer name, Item name, 구분, Sector 등은 첫 번째 값 사용
- **Model code**: 그룹 내 고유값을 `, `로 연결 (예: "P001, P002")
- **처리 순서**: Line item 레벨에서 Input/Output 계산 → OS name + 납기일로 그룹화 → 롤링 계산

### 결과 컬럼

| 컬럼 | 설명 |
|------|------|
| Period | 해당 월 (yyyy-MM, 텍스트) |
| 등록Period | 주문 등록 월 — SO 원본 Period (yyyy-MM, 텍스트) |
| 구분 | 국내/해외 |
| SO_ID | 주문 번호 |
| Customer name | 고객명 |
| Customer PO | 고객 발주번호 (대표값) |
| Item name | 아이템명 (대표값) |
| OS name | OneStream Item name (**그룹화 키**) |
| Expected delivery date | 예상 납기일 (**그룹화 키**, 같은 OS name이라도 납기일 다르면 구분) |
| AX Period | AX 기간 (그룹 내 고유값 연결) |
| Model code | ERP 프로젝트 번호 (그룹 내 고유값 연결) |
| Sector | 사업 부문 |
| Business registration number | 사업자등록번호 |
| Industry code | 산업 코드 |
| Value_Start_qty | 전월 이월 수량 |
| Value_Input_qty | 당월 수주 수량 (등록 Period에만) |
| Value_Output_qty | 당월 매출 수량 (매출월에만, DN 기준) |
| Value_Variance_qty | 수량 조정분 (환율 영향 없음 → 항상 0, 최종 출력에서 제거) |
| Value_Ending_qty | Start + Input - Output |
| Value_Start_amount | 전월 이월 금액 |
| Value_Input_amount | 당월 수주 금액 (SO 기준, 수주시점 환율) |
| Value_Output_amount | 당월 매출 금액 (국내 `Total Sales` / 해외 `외화금액 × 선적월 환율`) |
| **Value_Variance_amount** | **환율 재평가분** = Output 재환산액 − DN 시트 `Total Sales KRW`. 해외 비KRW 건의 출고(선적)월에만 발생, 국내는 0 |
| Value_Ending_amount | Start + Input - Output + Variance |

### 동작 원리 도식

#### 전체 파이프라인

```
┌─────────────┐     ┌─────────────┐
│  SO_국내     │     │  SO_해외     │
│  (수주 원본)  │     │  (수주 원본)  │
└──────┬──────┘     └──────┬──────┘
       └────────┬─────────┘
                ▼
        ┌──────────────┐
        │  SO_Filtered  │  Cancelled·Hold 제외, #N/A 치환
        │  (전체 수주)   │  Period 빈 행 제외
        └──────┬───────┘
               │
               │  ┌─────────────┐     ┌─────────────┐
               │  │  DN_국내     │     │  DN_해외     │
               │  │  (출고 원본)  │     │  (출고 원본)  │
               │  └──────┬──────┘     └──────┬──────┘
               │         └────────┬─────────┘
               │                  ▼
               │          ┌───────────────┐     ┌────────┐
               │          │  DN_Combined  │◀────│   FX    │
               │          │ 매출인식일→매출월│     │ 선적월  │
               │          │ 해외 KRW 재환산 │     │ 환율    │
               │          └──────┬────────┘     └────────┘
               │                 │
               │          ┌──────┴──────────────────┐
               │          ▼                         ▼
               │  ┌──────────────┐         ┌──────────────────┐
               │  │ DN_LastMonth  │         │   DN_ByMonth      │
               │  │ SO+Line별     │         │   SO+Line+매출월별  │
               │  │ 마지막 매출월  │         │   월별 qty/amount  │
               │  └──────┬───────┘         └────────┬─────────┘
               │         │                          │
       ┌───────┴─────────┘                          │
       ▼                                            │
 ══════════════════                                 │
  ① SO + 마지막출고월                                 │
     LEFT JOIN                                      │
 ══════════════════                                 │
       │                                            │
       ▼                                            │
 ══════════════════                                 │
  ② Period 확장                                      │
    끝점 = 항상 LastPeriod (현재월)                     │
                                                    │
    모든 건을 현재월까지 확장                            │
    → 부분출고 잔고 이월 보장                           │
                                                    │
    * 완납 건(Ending=0) 후속 빈 행은                   │
      ⑥ 이후 ZeroFiltered에서 제거                    │
 ══════════════════                                 │
       │                                            │
       ▼                                            │
 ══════════════════                                 │
  ③ Input 계산                                       │
    등록Period에만                                    │
    qty/amount 기록                                  │
 ══════════════════                                 │
       │                                            │
       ├────────────────────────────────────────────┘
       ▼
 ══════════════════
  ④ Output 조인
    DN_ByMonth와
    Period = 매출월 매칭
    → 월별 매출 배분
    → 해외는 선적월 환율
      재환산액을 Output으로,
      시트값과의 차를
      Variance로
 ══════════════════
       │
       ▼
 ══════════════════
  ⑤ OS name 그룹화
    SO_ID + OS name
    + Expected delivery date
    + Period 기준 합산
 ══════════════════
       │
       ▼
 ══════════════════
  ⑥ 롤링 계산
    그룹별 Period 정렬
    Start → Ending 전파
 ══════════════════
       │
       ▼
   ┌────────────┐
   │  결과 테이블  │
   └────────────┘
```

#### 단계별 데이터 변화 (예시)

SOD-0001, IQ3 2개 Line item을 1월 수주, 1월/2월 분할 출고하는 경우:

**원본 데이터**
```
SO 시트:
┌───────────┬───────┬──────┬─────────┬────────┬──────────┬────────┐
│ SO_ID     │ Line  │ OS   │ Period  │ qty    │ amount   │ 납기일  │
├───────────┼───────┼──────┼─────────┼────────┼──────────┼────────┤
│ SOD-0001  │ 1     │ IQ3  │ 2026-01 │ 10     │ 500만    │ 2/20   │
│ SOD-0001  │ 2     │ IQ3  │ 2026-01 │ 5      │ 250만    │ 2/20   │
└───────────┴───────┴──────┴─────────┴────────┴──────────┴────────┘
             SO 원본은 수주 시점에 1행만 존재

DN 시트:
┌───────────┬───────┬──────┬──────────┬──────────┐
│ SO_ID     │ Line  │ qty  │ 출고금액  │ 출고월    │
├───────────┼───────┼──────┼──────────┼──────────┤
│ SOD-0001  │ 1     │ 3    │ 150만    │ 2026-01  │  ← 1월 분할출고
│ SOD-0001  │ 1     │ 7    │ 350만    │ 2026-02  │  ← 2월 잔량출고
│ SOD-0001  │ 2     │ 5    │ 250만    │ 2026-01  │  ← 1월 전량출고
└───────────┴───────┴──────┴──────────┴──────────┘
```

**① SO + 마지막출고월 JOIN** — 각 Line item이 언제까지 활동하는지 끝점을 알아야 함
```
┌───────────┬───────┬──────┬─────────┬────────┬──────────┬──────────┐
│ SO_ID     │ Line  │ OS   │ Period  │ qty    │ amount   │ 출고월    │
├───────────┼───────┼──────┼─────────┼────────┼──────────┼──────────┤
│ SOD-0001  │ 1     │ IQ3  │ 2026-01 │ 10     │ 500만    │ 2026-02  │ ← 마지막 출고월
│ SOD-0001  │ 2     │ IQ3  │ 2026-01 │ 5      │ 250만    │ 2026-01  │ ← 마지막 출고월
└───────────┴───────┴──────┴─────────┴────────┴──────────┴──────────┘
                                                           ↑
                                              ②에서 "어디까지 펼칠지" 결정에 사용
```

**② Period 확장** — SO는 1행뿐인데 월별 잔고를 추적하려면 매월 행이 있어야 함 → 1행을 N행으로 복제
```
┌───────────┬───────┬──────┬──────────┬─────────┐
│ SO_ID     │ Line  │ OS   │ 등록P    │ Period  │  ← 확장된 Period
├───────────┼───────┼──────┼──────────┼─────────┤
│ SOD-0001  │ 1     │ IQ3  │ 2026-01  │ 2026-01 │  ← 등록월~
│ SOD-0001  │ 1     │ IQ3  │ 2026-01  │ 2026-02 │  ←        ~현재월까지
│ SOD-0001  │ 2     │ IQ3  │ 2026-01  │ 2026-01 │  ← 등록월~
│ SOD-0001  │ 2     │ IQ3  │ 2026-01  │ 2026-02 │  ←        ~현재월까지
└───────────┴───────┴──────┴──────────┴─────────┘
  모든 건 = 항상 LastPeriod(현재월)까지 확장
  완납 건(Ending=0)의 후속 빈 행은 ⑥ 이후 ZeroFiltered에서 제거
```

**③ Input 계산** — 수주 금액을 등록월에만 기록 (복제된 행마다 넣으면 중복 집계됨)
```
┌───────────┬───────┬─────────┬───────────┬──────────────┐
│ SO_ID     │ Line  │ Period  │ Input_qty │ Input_amount │
├───────────┼───────┼─────────┼───────────┼──────────────┤
│ SOD-0001  │ 1     │ 2026-01 │ ★ 10      │ ★ 500만      │  ← 등록월이므로 Input 기록
│ SOD-0001  │ 1     │ 2026-02 │ 0         │ 0            │  ← 복제된 행이므로 0
│ SOD-0001  │ 2     │ 2026-01 │ ★ 5       │ ★ 250만      │  ← 등록월이므로 Input 기록
└───────────┴───────┴─────────┴───────────┴──────────────┘
```

**④ Output 조인** — DN 출고를 해당 Period에 매칭 (분할 출고 시 각 월에 해당 금액 배분)
```
DN_ByMonth:                              매칭 결과:
┌──────────┬──────┬─────────┬─────┐     ┌───────┬──────┬─────────┬────────┬──────────┐
│ SO_ID    │ Line │ 출고월   │ qty │     │ SO_ID │ Line │ Period  │ In_qty │ Out_qty  │
├──────────┼──────┼─────────┼─────┤     ├───────┼──────┼─────────┼────────┼──────────┤
│ SOD-0001 │ 1    │ 2026-01 │ 3   │ ──▶ │  0001 │ 1    │ 2026-01 │ 10     │ ★ 3      │
│ SOD-0001 │ 1    │ 2026-02 │ 7   │ ──▶ │  0001 │ 1    │ 2026-02 │ 0      │ ★ 7      │
│ SOD-0001 │ 2    │ 2026-01 │ 5   │ ──▶ │  0001 │ 2    │ 2026-01 │ 5      │ ★ 5      │
└──────────┴──────┴─────────┴─────┘     └───────┴──────┴─────────┴────────┴──────────┘
  Period = 출고월이면 매칭 → Output 기록
```

**⑤ OS name 그룹화** — Line item 단위는 행이 너무 많으므로 같은 제품(OS name)+납기일끼리 합산
```
그룹화 키: SO_ID + OS name + 납기일 + Period
  Line 1 (IQ3, 납기 2/20) + Line 2 (IQ3, 납기 2/20) → 같은 그룹

┌───────────┬──────┬────────┬─────────┬─────────────┬──────────────┬───────────────┬────────────────┐
│ SO_ID     │ OS   │ 납기일  │ Period  │ Input_qty   │ Input_amount │ Output_qty    │ Output_amount  │
├───────────┼──────┼────────┼─────────┼─────────────┼──────────────┼───────────────┼────────────────┤
│ SOD-0001  │ IQ3  │ 2/20   │ 2026-01 │ 10+5 = 15   │ 750만        │ 3+5 = 8       │ 400만          │
│ SOD-0001  │ IQ3  │ 2/20   │ 2026-02 │ 0           │ 0            │ 7             │ 350만          │
└───────────┴──────┴────────┴─────────┴─────────────┴──────────────┴───────────────┴────────────────┘
```

**⑥ 롤링 계산** — 통장처럼 이번 달 잔고가 다음 달 시작으로 이월
```
┌─────────┬────────────┬─────────────┬──────────────┬─────────────┐
│ Period  │ Start      │ Input       │ Output       │ Ending      │
├─────────┼────────────┼─────────────┼──────────────┼─────────────┤
│ 2026-01 │ 0          │ +15 (750만)  │ -8 (400만)   │ = 7 (350만)  │ ← Backlog
│ 2026-02 │ 7 (350만) ◀─ (이월) ──────│──────────────│─────────────│
│         │            │ +0          │ -7 (350만)   │ = 0 (0원)    │ ← 소진
└─────────┴────────────┴─────────────┴──────────────┴─────────────┘
                                                      ↑
  Ending = 0 → 정상 (SO 수주 총액 = DN 출고 총액)
  Ending ≠ 0 → SO-DN 금액 불일치 → 데이터 점검 필요
```

### M 코드

```
let
    // ========== SO 원본 (수주) ==========
    SO_국내_Raw = Excel.CurrentWorkbook(){[Name="SO_국내"]}[Content],
    SO_해외_Raw = Excel.CurrentWorkbook(){[Name="SO_해외"]}[Content],

    SO_국내 = Table.SelectColumns(SO_국내_Raw, {"SO_ID", "Customer name", "Customer PO", "Item name", "OS name", "Line item", "Item qty", "Sales amount", "Period", "Status", "AX Period", "AX Project number", "Sector", "Business registration number", "Industry code", "Expected delivery date"}),
    SO_국내_Renamed = Table.RenameColumns(SO_국내, {{"Sales amount", "Sales amount KRW"}}),
    SO_국내_Tagged = Table.AddColumn(SO_국내_Renamed, "구분", each "국내"),

    SO_해외 = Table.SelectColumns(SO_해외_Raw, {"SO_ID", "Customer name", "Customer PO", "Item name", "OS name", "Line item", "Item qty", "Sales amount KRW", "Period", "Status", "AX Period", "AX Project number", "Sector", "Business registration number", "Industry code", "Expected delivery date"}),
    SO_해외_Tagged = Table.AddColumn(SO_해외, "구분", each "해외"),

    SO_Combined = Table.Combine({SO_국내_Tagged, SO_해외_Tagged}),
    // #N/A 등 에러 값을 null로 치환 — 모든 컬럼 대상 (XLOOKUP 실패 등 lazy evaluation 에러 방지)
    SO_CleanErrors = Table.ReplaceErrorValues(SO_Combined,
        List.Transform(Table.ColumnNames(SO_Combined), each {_, null})
    ),
    // Cancelled, Hold 제외, Period 비어있는 행 제외
    SO_Filtered = Table.SelectRows(SO_CleanErrors, each
        ([Status] = null or not List.Contains({"Cancelled", "Hold"}, [Status])) and
        [Period] <> null and Text.Trim(Text.From([Period])) <> ""
    ),

    // ========== FX 시트 언피벗 (가로 → 세로: Currency + 환율월 + 환율) ==========
    // AX_매출대사와 동일 로직 — 해외 Output을 선적월 환율로 재환산하기 위함
    FX_Unpiv = Table.UnpivotOtherColumns(Table.RenameColumns(Excel.CurrentWorkbook(){[Name="FX"]}[Content], {{"FX", "Currency"}}), {"Currency"}, "환율월", "환율"),
    FX_Clean = Table.Buffer(Table.SelectRows(FX_Unpiv, each [Currency] <> null and [환율] <> null and Text.Length(Text.From([환율월])) = 7 and Text.Contains(Text.From([환율월]), "-"))),

    // ========== DN (매출 인식 시점 + 실제 매출 금액) ==========
    DN_국내_Raw = Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content],
    DN_해외_Raw = Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content],

    // 국내: 매출인식일 기준 (AX_매출대사와 동일 귀속). 우선순위:
    //   ① 세금계산서 발행일  ② 'N/A' 표기(= 발행 불필요: 무상공급·FOC·반품)면 출고일
    //   ③ 선수금 세금계산서 + 출고 완료면 출고일 (선청구 후 출고 시점에 수익인식)
    //   ④ 아무것도 없으면 매출 미인식 → Output 없음 = Backlog에 남는다 (월합세금계산서 대기 등)
    DN_국내_Select = Table.SelectColumns(DN_국내_Raw, {"SO_ID", "Line item", "Qty", "Total Sales", "출고일", "세금계산서 발행일", "선수금 세금계산서 발행일"}),
    DN_국내 = Table.ReplaceErrorValues(DN_국내_Select,
        List.Transform(Table.ColumnNames(DN_국내_Select), each {_, null})
    ),
    DN_국내_Rev = Table.AddColumn(DN_국내, "매출일", each
        let 세금일 = try Date.From([세금계산서 발행일]) otherwise null in
        if 세금일 <> null then 세금일
        // 'N/A'는 날짜가 아니라 "발행하지 않는다"는 표기 → 출고 시점에 매출로 본다
        // (안 그러면 무상공급·FOC 수량이 Backlog에 영구히 남는다)
        else if [세금계산서 발행일] <> null and Text.Upper(Text.Trim(Text.From([세금계산서 발행일]))) = "N/A" then [출고일]
        else if [선수금 세금계산서 발행일] <> null and [출고일] <> null then [출고일]
        else null, type date),
    DN_국내_WithPeriod = Table.AddColumn(DN_국내_Rev, "매출월", each
        if [매출일] = null then null
        else Text.From(Date.Year([매출일])) & "-" & Text.PadStart(Text.From(Date.Month([매출일])), 2, "0"),
        type text),
    DN_국내_Renamed = Table.RenameColumns(
        Table.RemoveColumns(DN_국내_WithPeriod, {"세금계산서 발행일", "선수금 세금계산서 발행일", "매출일"}),
        {{"Total Sales", "출고금액"}}),
    // 국내는 KRW 거래 — 재환산 없음, FX 조정 0
    DN_국내_Final = Table.AddColumn(Table.AddColumn(DN_국내_Renamed,
        "출고금액_재환산", each [출고금액], type number),
        "FX_조정", each 0, type number),

    // 해외: 선적일 기준 (매출 인식 시점), KRW = 외화금액 × 선적월 환율로 재환산
    DN_해외_Select = Table.SelectColumns(DN_해외_Raw, {"SO_ID", "Line item", "Qty", "Currency", "Total Sales", "Total Sales KRW", "선적일"}),
    DN_해외 = Table.ReplaceErrorValues(DN_해외_Select,
        List.Transform(Table.ColumnNames(DN_해외_Select), each {_, null})
    ),
    DN_해외_WithPeriod = Table.AddColumn(DN_해외, "매출월", each
        if [선적일] = null then null
        else Text.From(Date.Year([선적일])) & "-" & Text.PadStart(Text.From(Date.Month([선적일])), 2, "0"),
        type text),
    DN_해외_Renamed = Table.RenameColumns(DN_해외_WithPeriod, {{"Total Sales KRW", "출고금액"}, {"Total Sales", "외화금액"}}),
    // 선적월 환율 조인 → DN 라인 단위로 반올림 (AX_매출대사와 동일 grain·동일 반올림이라 합계가 일치)
    DN_해외_FXJoin = Table.ExpandTableColumn(Table.NestedJoin(DN_해외_Renamed, {"Currency", "매출월"}, FX_Clean, {"Currency", "환율월"}, "FX_Match", JoinKind.LeftOuter), "FX_Match", {"환율"}, {"환율"}),
    DN_해외_FX = Table.AddColumn(Table.AddColumn(DN_해외_FXJoin,
        "출고금액_재환산", each
            if [Currency] = "KRW" then [외화금액]
            else if [환율] <> null and [외화금액] <> null then Number.Round([외화금액] * [환율], 0)
            else [출고금액],   // 선적월 환율·외화금액 결측 → 시트값 유지 (Output이 조용히 0이 되는 것 방지)
        type number),
        "FX_조정", each [출고금액_재환산] - [출고금액], type number),
    DN_해외_Final = Table.RemoveColumns(DN_해외_FX, {"Currency", "외화금액", "환율"}),

    DN_Combined = Table.Combine({DN_국내_Final, DN_해외_Final}),

    // DN 월별 집계 (분할 출고 대응: SO_ID + Line item + 매출월)
    // Output_amount = 선적월 환율 재환산액, Output_fx = 시트값과의 차이(= Variance 재원)
    DN_ByMonth = Table.Group(DN_Combined, {"SO_ID", "Line item", "매출월"}, {
        {"Output_qty", each List.Sum([Qty]), type number},
        {"Output_amount", each List.Sum([출고금액_재환산]), Currency.Type},
        {"Output_fx", each List.Sum([FX_조정]), Currency.Type}
    }),

    // DN 마지막 매출월 (ActivePeriods 범위 결정용)
    DN_LastMonth = Table.Group(DN_Combined, {"SO_ID", "Line item"}, {
        {"매출월", each List.Max(List.RemoveNulls([매출월])), type text}
    }),

    // ========== SO + DN 조인 (기간 범위용, 마지막 매출월만) ==========
    WithDN = Table.NestedJoin(SO_Filtered, {"SO_ID", "Line item"}, DN_LastMonth, {"SO_ID", "Line item"}, "DN_Data", JoinKind.LeftOuter),
    WithDNExpanded = Table.ExpandTableColumn(WithDN, "DN_Data", {"매출월"}),

    // ========== Period 리스트 ==========
    // SO 등록 Period + DN 모든 매출월 (분할 출고 중간 월 누락 방지)
    AllPeriods = List.Buffer(List.Sort(List.Distinct(
        List.RemoveNulls(Table.Column(WithDNExpanded, "Period")) &
        List.RemoveNulls(Table.Column(DN_ByMonth, "매출월"))
    ))),
    LastPeriod = List.Last(AllPeriods),

    // ========== 건별 × Period 확장 ==========
    // 각 SO Line: 등록 Period ~ 마지막 Period까지 행 생성
    // 항상 LastPeriod까지 확장하여 부분출고 잔고 이월 보장
    // 완납 건(Ending=0)의 후속 빈 행은 ZeroFiltered 단계에서 제거
    WithPeriodList = Table.AddColumn(WithDNExpanded, "ActivePeriods", each
        let
            startIdx = List.PositionOf(AllPeriods, [Period]),
            // 항상 LastPeriod까지 확장 (부분출고 건의 잔고 이월 보장)
            // 완납 건(Ending=0)의 후속 빈 행은 롤링 계산 후 ZeroFiltered에서 제거
            endPeriod = LastPeriod,
            endIdx = List.PositionOf(AllPeriods, endPeriod),
            s = if startIdx < 0 then 0 else startIdx,
            e = if endIdx < 0 then List.Count(AllPeriods) - 1
                else if endIdx < s then s else endIdx
        in
            List.Range(AllPeriods, s, e - s + 1)
    ),

    Expanded = Table.ExpandListColumn(WithPeriodList, "ActivePeriods"),
    Renamed = Table.RenameColumns(Expanded, {{"Period", "등록Period"}, {"ActivePeriods", "Period"}}),

    // ========== Input (Line item 레벨) ==========
    // Input = SO 금액 (수주 시점, 등록 Period에만)
    WithInputQty = Table.AddColumn(Renamed, "Value_Input_qty", each
        if [Period] = [등록Period] then [Item qty] else 0, type number),
    WithInputAmt = Table.AddColumn(WithInputQty, "Value_Input_amount", each
        if [Period] = [등록Period] then [Sales amount KRW] else 0, type number),

    // ========== Output (DN 월별 조인) ==========
    // DN_ByMonth와 SO_ID + Line item + Period = 매출월 조인 → 분할 출고 월별 배분
    WithDNOutput = Table.NestedJoin(WithInputAmt, {"SO_ID", "Line item", "Period"}, DN_ByMonth, {"SO_ID", "Line item", "매출월"}, "DN_Output", JoinKind.LeftOuter),
    WithDNOutputExpanded = Table.ExpandTableColumn(WithDNOutput, "DN_Output", {"Output_qty", "Output_amount", "Output_fx"}),
    WithOutputQty = Table.AddColumn(WithDNOutputExpanded, "Value_Output_qty", each
        if [Output_qty] = null then 0 else [Output_qty], type number),
    WithValues = Table.AddColumn(WithOutputQty, "Value_Output_amount", each
        if [Output_amount] = null then 0 else [Output_amount], type number),

    // ========== Variance (환율 재평가) ==========
    // Input은 수주시점 환율의 SO 금액, Output은 선적월 환율 재환산액 → 그 차액을 Variance로 흡수
    // → 매출월에 환율차만큼 Backlog를 재평가하고 소진 ⇒ Ending은 재환산 도입 전과 동일
    WithVariance = Table.AddColumn(WithValues, "Value_Variance_amount", each
        if [Output_fx] = null then 0 else [Output_fx], type number),
    WithValuesCleaned = Table.RemoveColumns(WithVariance, {"Output_qty", "Output_amount", "Output_fx"}),

    // ========== OS name 기준 그룹화 ==========
    // Line item 레벨 → SO_ID + OS name + Expected delivery date + Period 로 합산
    // 같은 OS name의 Line item들이 하나의 행으로 합쳐짐
    OSGrouped = Table.Group(WithValuesCleaned, {"SO_ID", "OS name", "Expected delivery date", "Period"}, {
        {"Customer name", each List.First([Customer name]), type text},
        {"Customer PO", each List.First([Customer PO]), type text},
        {"Item name", each List.First([Item name]), type text},
        {"구분", each List.First([구분]), type text},
        {"등록Period", each List.First([등록Period]), type text},
        {"Sector", each List.First([Sector]), type text},
        {"Business registration number", each List.First([Business registration number]), type text},
        {"Industry code", each List.First([Industry code]), type text},
        {"AX Period", each Text.Combine(List.Distinct(List.RemoveNulls([AX Period])), ", "), type text},
        {"AX Project number", each Text.Combine(List.Distinct(List.RemoveNulls([AX Project number])), ", "), type text},
        {"Value_Input_qty", each List.Sum([Value_Input_qty]), type number},
        {"Value_Input_amount", each List.Sum([Value_Input_amount]), type number},
        {"Value_Output_qty", each List.Sum([Value_Output_qty]), type number},
        {"Value_Output_amount", each List.Sum([Value_Output_amount]), type number},
        {"Value_Variance_amount", each List.Sum([Value_Variance_amount]), type number}
    }),

    // ========== 건별 롤링 계산 ==========
    // SO_ID + OS name 그룹 → 각 그룹 내에서 Start/Ending 전파
    ProcessLine = (lineTable as table) as list =>
        let
            sorted = Table.Sort(lineTable, {{"Period", Order.Ascending}}),
            rows = Table.ToRecords(sorted),
            result = List.Accumulate({0..List.Count(rows)-1}, {}, (state, idx) =>
                let
                    r = rows{idx},
                    sQty = if idx = 0 then 0 else state{idx-1}[Value_Ending_qty],
                    sAmt = if idx = 0 then 0 else state{idx-1}[Value_Ending_amount]
                in
                    state & {[
                        Period = r[Period],
                        구분 = r[구분],
                        #"등록Period" = r[#"등록Period"],
                        SO_ID = r[SO_ID],
                        #"Customer name" = r[#"Customer name"],
                        #"Customer PO" = r[#"Customer PO"],
                        #"Item name" = r[#"Item name"],
                        #"OS name" = r[#"OS name"],
                        #"Expected delivery date" = r[#"Expected delivery date"],
                        #"AX Period" = r[#"AX Period"],
                        #"AX Project number" = r[#"AX Project number"],
                        Sector = r[Sector],
                        #"Business registration number" = r[#"Business registration number"],
                        #"Industry code" = r[#"Industry code"],
                        Value_Start_qty = sQty,
                        Value_Input_qty = r[Value_Input_qty],
                        Value_Output_qty = r[Value_Output_qty],
                        Value_Variance_qty = 0,
                        Value_Ending_qty = sQty + r[Value_Input_qty] - r[Value_Output_qty],
                        Value_Start_amount = sAmt,
                        Value_Input_amount = r[Value_Input_amount],
                        Value_Output_amount = r[Value_Output_amount],
                        Value_Variance_amount = r[Value_Variance_amount],
                        Value_Ending_amount = sAmt + r[Value_Input_amount] - r[Value_Output_amount] + r[Value_Variance_amount]
                    ]}
            )
        in
            result,

    Grouped = Table.Group(OSGrouped, {"SO_ID", "OS name", "Expected delivery date"}, {
        {"Processed", each ProcessLine(_)}
    }),

    // 결과 펼치기
    AllRows = List.Combine(Grouped[Processed]),
    ResultTable = Table.FromRecords(AllRows),

    // ========== 정렬 + 컬럼 정리 + 타입 ==========
    Reordered = Table.ReorderColumns(ResultTable, {
        "Period", "등록Period", "구분", "SO_ID", "Customer name", "Customer PO", "Item name", "OS name",
        "Expected delivery date", "AX Period", "AX Project number", "Sector", "Business registration number", "Industry code",
        "Value_Start_qty", "Value_Input_qty", "Value_Output_qty", "Value_Variance_qty", "Value_Ending_qty",
        "Value_Start_amount", "Value_Input_amount", "Value_Output_amount", "Value_Variance_amount", "Value_Ending_amount"
    }),

    FinalSorted = Table.Sort(Reordered, {
        {"Period", Order.Descending},
        {"구분", Order.Ascending},
        {"SO_ID", Order.Ascending},
        {"OS name", Order.Ascending}
    }),

    Result = Table.TransformColumnTypes(FinalSorted, {
        {"Expected delivery date", type date},
        {"Value_Start_qty", Int64.Type},
        {"Value_Input_qty", Int64.Type},
        {"Value_Output_qty", Int64.Type},
        {"Value_Variance_qty", Int64.Type},
        {"Value_Ending_qty", Int64.Type},
        {"Value_Start_amount", Currency.Type},
        {"Value_Input_amount", Currency.Type},
        {"Value_Output_amount", Currency.Type},
        {"Value_Variance_amount", Currency.Type},
        {"Value_Ending_amount", Currency.Type}
    }),

    // 완납 건(Ending=0) 후속 빈 행 제거: Start=Input=Output=Variance=Ending 모두 0이면 제거
    ZeroFiltered = Table.SelectRows(Result, each
        not ([Value_Start_qty] = 0 and [Value_Input_qty] = 0 and [Value_Output_qty] = 0 and [Value_Ending_qty] = 0
         and [Value_Start_amount] = 0 and [Value_Input_amount] = 0 and [Value_Output_amount] = 0
         and [Value_Variance_amount] = 0 and [Value_Ending_amount] = 0)
    ),

    #"Reordered Columns" = Table.ReorderColumns(ZeroFiltered,{"SO_ID", "AX Project number", "AX Period", "구분", "Period", "등록Period", "Customer name", "Customer PO", "Item name", "OS name", "Sector", "Business registration number", "Industry code", "Value_Start_qty", "Value_Input_qty", "Value_Output_qty", "Value_Variance_qty", "Value_Ending_qty", "Value_Start_amount", "Value_Input_amount", "Value_Output_amount", "Value_Variance_amount", "Value_Ending_amount", "Expected delivery date"}),
    #"Removed Columns" = Table.RemoveColumns(#"Reordered Columns",{"Item name"}),
    #"Reordered Columns1" = Table.ReorderColumns(#"Removed Columns",{"SO_ID", "AX Project number", "AX Period", "구분", "Period", "등록Period", "Business registration number", "Customer name", "Customer PO", "Sector", "Industry code", "OS name", "Value_Start_qty", "Value_Input_qty", "Value_Output_qty", "Value_Variance_qty", "Value_Ending_qty", "Value_Start_amount", "Value_Input_amount", "Value_Output_amount", "Value_Variance_amount", "Value_Ending_amount", "Expected delivery date"}),
    // Value_Variance_qty만 제거 (수량엔 환율 영향이 없어 항상 0) — 금액 Variance는 유지
    #"Removed Columns1" = Table.RemoveColumns(#"Reordered Columns1",{"Value_Variance_qty"})
in
    #"Removed Columns1"
```

### 결과 예시

| Period | 구분 | SO_ID | Customer name | OS name | Expected delivery date | Ending_qty | Ending_amt |
|--------|------|-------|---------------|---------|----------------------|------------|------------|
| 2026-02 | 국내 | SOD-0001 | 삼성전자 | IQ3 | 2026-02-20 | **0** | **0** |
| 2026-02 | 국내 | SOD-0003 | 현대중공업 | NA028 | 2026-03-10 | **20** | **10,000,000** |
| 2026-02 | 해외 | SOO-0001 | ABC Corp | CVA | 2026-03-15 | **5** | **4,000,000** |
| 2026-01 | 국내 | SOD-0002 | LG전자 | CVA | 2026-01-25 | **0** | **0** |
| 2026-01 | 국내 | SOD-0002 | LG전자 | CVA | 2026-02-10 | **3** | **1,500,000** |

```
SOD-0001: 1월 수주(10) → 2월 출고(10) → Ending=0 (소진)
SOD-0002: 같은 CVA지만 납기일 다름 → 1/25분은 출고 완료, 2/10분은 Backlog
SOD-0003: 1월 수주(20) → 2월에도 미출고 → Ending=20 (Backlog)
```

### 활용

| 보고 싶은 것 | 방법 |
|-------------|------|
| **현재 Backlog** | 마지막 Period 필터 → Value_Ending_amount > 0 |
| **월별 요약** | 피벗 테이블: Period 행 → SUM(Input/Output/Ending) |
| **고객별 Backlog** | Value_Ending > 0 필터 → Customer name 그룹화 |
| **국내/해외 split** | 구분 컬럼 필터/슬라이서 |
| **Sector별 분석** | Sector 컬럼 필터/슬라이서 → 사업부문별 Backlog |
| **특정 월 스냅샷** | Period = "2026-01" 필터 → 그 시점의 모든 건 |
| **누적 매출** | Value_Output_amount를 P01~해당월까지 합산 |
| **SO-DN 차이 점검** | 출고 완료 건 중 Value_Ending ≠ 0 필터 |
| **AX 매출 대사** | Period 필터 → SUM(Value_Output_amount) = `AX_매출대사` 같은 월 `매출금액_KRW` 합 (해외는 완전 일치, 국내는 아래 주의) |
| **환율 임팩트** | Period 필터 → SUM(Value_Variance_amount) = 그 달 매출의 환율 재평가액 |

#### AX_매출대사와의 대사 (금액·귀속월 완전 일치)

| 구분 | Order_Book Output 귀속월 | AX_매출대사 매출월 | 금액 |
|------|------------------------|------------------|------|
| 국내 | 세금계산서 발행월 | 세금계산서 발행월 | `Total Sales` (동일) |
| 해외 | DN 선적월 | DN 선적월 | 외화금액 × 선적월 환율 (동일) |

```
Period = "2026-07" 필터 → SUM(Value_Output_amount)
   = AX_매출대사에서 매출월 = "P07" 필터 → SUM(매출금액_KRW)
```

- **환율 축**은 Variance로, **귀속월 축**은 매출인식일로 맞췄다. 2026-01~07 전 월 차이 0원 (실측).
- 대가: Order_Book Backlog는 이제 **"아직 매출로 인식되지 않은 물량"**이다. 출고됐지만 세금계산서 미발행 건은 월말 Backlog에 남는다(2026-07 기준 4라인 / 1,232,420원 — 월합세금계산서 대기분). 물류상 미출고 잔량은 대시보드 `납기현황`/`EXW미출고`를 본다.
- `N/A` 표기(무상공급·FOC·반품)를 출고월 인식으로 처리하는 이유: 세금계산서를 영구히 발행하지 않는 건이라 미인식으로 두면 **수량이 Backlog에 영구히 남는다**. 금액이 0이거나 같은 달 안에서 상쇄되므로 월별 매출 합계는 바뀌지 않는다.

### AX 오더북과의 비교

| 항목 | AX2009 Order Book | NOAH Order_Book |
|------|-------------------|-----------------|
| 마감 | Period 마감 → 잠금 | 없음 (매번 재계산) |
| Start 이월 | DB에 저장된 값 | Power Query가 계산한 값 |
| Variance | 자동 추적 (금액 변경, 취소) | **환율 재평가분** (수량·판가 조정분은 새 Line item으로 추가 → Input에서 넷팅) |
| 스냅샷 | DB에 보존 | 없음 (현재 데이터 기준) |
| 그룹화 | Project number 기준 | SO_ID + OS name 기준 (Line item 합산) |
| Input 기준 | SO 등록일 | SO의 Period 컬럼 (yyyy-MM) |
| Output 기준 | Invoice 일자 | 세금계산서 발행일(국내) / 선적일(해외) — AX와 같은 기준 |
| 금액 기준 | SO 금액 | Input=SO(수주시점 환율), Output=DN(선적월 환율), 차액=Variance |
| 갱신 | 자동 (트랜잭션 기반) | Ctrl+Alt+F5 (수동 새로고침) |

### 전제조건

- SO의 **Period 컬럼**: `yyyy-MM` 형식 텍스트 (예: "2026-01")
- SO의 **Sector 컬럼**: 사업 부문 (예: "Process", "CPI", "Water")
- SO의 **Business registration number 컬럼**: 사업자등록번호
- SO의 **Industry code 컬럼**: 산업 코드
- DN_국내의 **세금계산서 발행일 / 선수금 세금계산서 발행일 / 출고일**, DN_해외의 **선적일**: 날짜 형식 → 쿼리에서 yyyy-MM으로 변환.
  `세금계산서 발행일`에 날짜가 아닌 텍스트가 들어오면 `N/A`만 "발행 불필요"로 인식하고, 그 밖의 텍스트는 **미인식(Backlog 잔류)** 으로 떨어진다 — 매출이 조용히 잡히는 것보다 눈에 보이는 쪽으로 실패한다
- **`FX` 시트**: 통합 문서 안에 표(또는 이름 정의)로 존재하고, 헤더가 `FX | 2026-01 | …`로 승격돼 있어야 한다 (월 컬럼명은 `YYYY-MM` 텍스트). `AX_매출대사`와 같은 전제.
  선적월 환율이 아직 없는 월(예: 마감 전 당월)은 **시트의 `Total Sales KRW`로 폴백**하고 Variance = 0이 된다 → Output이 누락되지 않는다. 환율이 입력되면 새로고침만으로 반영된다.

### 수량/금액 조정 방법 (분개 방식)

원래 SO 행은 수정하지 않고, **조정분을 새 Line item으로 추가**하여 넷팅합니다.

```
P01: SOD-0001, IQ3, Line item 1, qty=10, amount=500만  (원래 수주)
P02: SOD-0001, IQ3, Line item 2, qty=-2, amount=-100만  (조정분)

→ Order_Book 롤링:
  P01: Start=0,  Input=+10, Ending=10
  P02: Start=10, Input=-2,  Ending=8    ← 넷팅
  P03: Start=8,  Output=8,  Ending=0    ← DN 출고, 소진
```

- **원래 행 안 건드림** → SO raw에 이력 보존
- **새 Line item** → 언제, 얼마나 조정했는지 추적 가능
- OS name 그룹화가 자동으로 넷팅 처리
- **수량·판가 조정은 Variance가 아니라 Input으로 넷팅**한다 (Variance는 환율 재평가 전용)

### 한계

| 한계 | 설명 | 대응 |
|------|------|------|
| 마감 잠금 없음 | 과거 SO/DN 수정 시 소급 변경 | 원래 행 수정 금지, 조정은 새 Line item으로 추가 |
| 스냅샷 없음 | 과거 시점 재현 불가 | SO raw 데이터에 원본+조정 이력이 남아 추적 가능 |
| Period 갭 | 활동 없는 월은 행 생성 안됨 | 전월 Ending이 다음 활동월 Start로 정확히 이월됨 |
| 미선적 Backlog는 수주시점 환율 | Variance는 **매출로 인식된 부분만** 재평가한다. 미선적 잔고의 KRW 가치는 환율이 움직여도 그대로 | 실현 시점(선적)에 한 번에 반영되는 구조. 미실현 환평가가 필요해지면 잔고 전체를 매월 재평가하는 방식으로 확장 |
| Backlog ≠ 미출고 물량 | Output이 매출인식 기준이라, 출고했지만 세금계산서 미발행 건은 Backlog에 남는다 | 물류상 미출고는 대시보드 `납기현황`/`EXW미출고` 참조 |
| SQLite 미러와 ±1원 | 재환산 반올림이 정확히 .5인 라인에서 Power Query(`Number.Round`=짝수 반올림)와 SQLite(`ROUND`=올림)가 갈린다 | 실측 2라인, 월 합계 최대 ±1원. `AX_매출대사`와 맞추는 게 목적이라 Excel 쪽 기준을 유지 |

> **향후 확장**: ERP 통합 등으로 스냅샷 기반 조정 추적이 필요해지면, `Value_Variance_amount`에 환율 재평가분과 스냅샷 대비 차이를 구분해 담을 수 있음(예: `Value_Variance_fx` / `Value_Variance_adj`로 분리). 현재 분개 방식과 병행 가능.

### 검증 (2026-07-30 기준 실측)

| 검증 항목 | 결과 |
|-----------|------|
| **Output vs `AX_매출대사`** | **P01~P07 × 국내/해외 14개 조합 전부 차이 0원** (P07 국내 464,170,700 / 해외 789,883,397) |
| 건별 Ending 불변 (환율) | 재환산 도입 전/후 **모든 그룹×Period에서 Ending 동일** (불일치 0건) — 환율차가 Variance로 흡수되어 잔고·SO-DN 진단이 그대로 |
| P07 환율 임팩트 | SUM(Value_Variance_amount) 해외 = **+47,662,819원** (재환산 전 742,220,578 → 789,883,397). 누적 P01~P07 +115,496,363원 |
| 예시 건 SOO-2026-0188 | P06 Input 1,914,046(USD 1,276 × 1,500.036) → P07 Output 1,976,024(× 1,548.608) + Variance 61,978 → **Ending 0** |
| 귀속월 변경의 Backlog 영향 | 국내 **+4개 / +1,232,420원**(세금계산서 미발행 4라인), 해외 **변동 없음**. 그 외 전 그룹 Ending 동일 |
| 국내 환율 | Variance 항상 0 (KRW 거래) |
| 누락 점검 | SO 조인 실패·Cancelled/Hold·SO Period 공란으로 Output이 사라지는 DN 라인 **0건**. DN_해외 통화는 USD/EUR뿐이고 선적월 환율 결측 **0건** |
| SQLite 미러 | `sql/order_book.sql`도 **14개 조합 전부 동일** (뷰 `v_dn_revenue` 공유, ±1원 반올림 제외). DN 테이블 PK에 `_row_seq`를 추가해 분할출고 중복키 행이 삼켜지던 문제까지 해결 |

> **기존 마감 스냅샷 주의** (`ob_snapshot`, 2026-01~06): 옛 기준(출고월·시트 KRW)으로 동결돼 있다.
> 새 기준으로 재계산하면 2026-06 마감 시점 Ending이 **50그룹 / 수량 731 / 220,112,337원** 달라진다
> (대부분 "6월 출고·7월 세금계산서" 건이 6월 말 Backlog로 남게 된 것).
> 그대로 2026-07을 마감하면 이 차이가 **한 번에 Variance로 계상**된다. 기준 변경을 소급 반영하려면
> `close_period.py --undo`로 과거 월을 되돌린 뒤 순서대로 다시 마감해야 한다 — 어느 쪽을 택할지는
> 회계 판단이므로 자동으로 하지 않는다.

---

## Order_Book: Power Query vs SQL 비교

Power Query(Excel)로 구현된 Order_Book 로직을 SQLite로 포팅하면서 **기간 펼치기 → 이벤트 기반**으로 구조를 변경했습니다. 같은 예시 데이터로 두 방식의 차이를 설명합니다.

### 예시 데이터

**SO (수주)**
```
SO_ID = SOD-0001, OS name = IQ3
  Line 1: qty=10, amount=500만, Period=2026-01, 납기=2/20
  Line 2: qty= 5, amount=250만, Period=2026-01, 납기=2/20
```

**DN (출고)**
```
SOD-0001, Line 1: qty=3, 150만, 출고일=2026-01-20  (1월 분할출고)
SOD-0001, Line 1: qty=7, 350만, 출고일=2026-02-15  (2월 잔량출고)
SOD-0001, Line 2: qty=5, 250만, 출고일=2026-01-25  (1월 전량출고)
```

### Power Query 방식 (기간 펼치기)

**핵심 아이디어**: SO 1줄을 달력처럼 월별로 복제(expand)해서 빈 칸을 만들고, 각 칸에 Input/Output을 채운다.

```
                    원본 (1줄)
                   ┌─────────┐
                   │ SOD-0001 │
                   │ 1월 수주 │
                   └────┬────┘
                        │
          ┌─────────────┼──────────────┐
          ▼             ▼              ▼
     ┌─────────┐  ┌─────────┐    ┌─────────┐
     │ × 2026-01│  │ × 2026-02│    │ × 2026-03│   ← 빈 월도 복제
     └─────────┘  └─────────┘    └─────────┘
        (활동)       (활동)         (빈 행)
```

**Step ①~② 마지막 출고월 파악 + Period 확장**

Line 1의 마지막 출고 = 2월 → 1월~2월까지 복제
Line 2의 마지막 출고 = 1월 → 1월만

```
Line 1 × 2026-01  ← 원본
Line 1 × 2026-02  ← 복제 (빈 행)
Line 2 × 2026-01  ← 원본
```
이미 3행. 만약 미출고 건이면 현재월(3월)까지 펼치므로 더 늘어남.

**Step ③④ Input/Output 채우기**

```
Line 1 × 1월: Input=10/500만, Output=3/150만
Line 1 × 2월: Input=0,        Output=7/350만   ← 2월은 복제된 빈 행에 Output만 채움
Line 2 × 1월: Input=5/250만,  Output=5/250만
```

**Step ⑤ OS name 그룹화** (Line 1 + Line 2 합산)

```
┌─────────┬─────────────┬──────────────┬───────────────┬────────────────┐
│ Period  │ Input_qty   │ Input_amount │ Output_qty    │ Output_amount  │
├─────────┼─────────────┼──────────────┼───────────────┼────────────────┤
│ 2026-01 │ 10+5 = 15   │ 750만        │ 3+5 = 8       │ 400만          │
│ 2026-02 │ 0+0 = 0     │ 0            │ 7             │ 350만          │ ← 빈 행이 있어서 가능
└─────────┴─────────────┴──────────────┴───────────────┴────────────────┘
```

**Step ⑥ 롤링 계산**

```
┌─────────┬────────┬────────┬────────┬─────────┐
│ Period  │ Start  │ Input  │ Output │ Ending  │
├─────────┼────────┼────────┼────────┼─────────┤
│ 2026-01 │ 0      │ +15    │ -8     │ = 7     │ Backlog
│ 2026-02 │ 7 ◀────│────(이월)       │         │
│         │        │ +0     │ -7     │ = 0     │ 소진
└─────────┴────────┴────────┴────────┴─────────┘
```

**문제점**: 재귀 CTE로 연속 월을 생성하고(month_series), SO × 월 카테시안 조인으로 빈 행까지 만들어야 함. 데이터가 많아지면 행 수가 폭발적으로 증가.

### SQL 이벤트 기반 방식

**핵심 아이디어**: SO 등록 = Input 이벤트, DN 출고 = Output 이벤트. 이벤트가 발생한 월에만 행을 만든다. 빈 월은 만들지 않는다.

```
    이벤트만 기록
    ┌──────────────────┐     ┌──────────────────┐
    │ 2026-01          │     │ 2026-02          │
    │ Input: 15/750만  │     │ Output: 7/350만  │     ← 2026-03? 이벤트 없으면 행 없음
    │ Output: 8/400만  │     │                  │
    └──────────────────┘     └──────────────────┘
```

**Step 1: events_line_item (UNION ALL)**

Input 이벤트(SO 등록)와 Output 이벤트(DN 출고)를 하나로 합침:

```
┌──────────┬──────┬─────────────┬───────┬────────┬───────┬────────┐
│ SO_ID    │ Line │ event_period│ In_qty│ In_amt │Out_qty│Out_amt │
├──────────┼──────┼─────────────┼───────┼────────┼───────┼────────┤
│ SOD-0001 │ 1    │ 2026-01     │ 10    │ 500만  │ 0     │ 0      │ ← Input (SO 등록)
│ SOD-0001 │ 2    │ 2026-01     │ 5     │ 250만  │ 0     │ 0      │ ← Input (SO 등록)
│ SOD-0001 │ 1    │ 2026-01     │ 0     │ 0      │ 3     │ 150만  │ ← Output (DN 1월)
│ SOD-0001 │ 2    │ 2026-01     │ 0     │ 0      │ 5     │ 250만  │ ← Output (DN 1월)
│ SOD-0001 │ 1    │ 2026-02     │ 0     │ 0      │ 7     │ 350만  │ ← Output (DN 2월)
└──────────┴──────┴─────────────┴───────┴────────┴───────┴────────┘
  Input과 Output이 같은 테이블에 공존 → GROUP BY로 자연 합산
```

**Step 2: os_grouped (GROUP BY)**

```
┌─────────┬─────────────┬──────────────┬───────────────┬────────────────┐
│ Period  │ Input_qty   │ Input_amount │ Output_qty    │ Output_amount  │
├─────────┼─────────────┼──────────────┼───────────────┼────────────────┤
│ 2026-01 │ 15          │ 750만        │ 8             │ 400만          │ ← Input+Output 자연 합산
│ 2026-02 │ 0           │ 0            │ 7             │ 350만          │ ← Output만 있는 이벤트
└─────────┴─────────────┴──────────────┴───────────────┴────────────────┘
  결과가 Power Query ⑤와 동일! 빈 월(2026-03)은 행 자체가 없음
```

**Step 3: Window function 롤링**

```sql
SUM(Input - Output) OVER (ORDER BY Period ROWS UNBOUNDED PRECEDING TO 1 PRECEDING) → Start
SUM(Input - Output) OVER (ORDER BY Period ROWS UNBOUNDED PRECEDING TO CURRENT ROW)  → Ending
```

```
┌─────────┬────────┬────────┬────────┬─────────┐
│ Period  │ Start  │ Input  │ Output │ Ending  │
├─────────┼────────┼────────┼────────┼─────────┤
│ 2026-01 │ 0      │ +15    │ -8     │ = 7     │
│ 2026-02 │ 7      │ +0     │ -7     │ = 0     │  ← 결과 동일!
└─────────┴────────┴────────┴────────┴─────────┘
```

### 핵심 차이 비교

```
┌──────────────────────┬────────────────────────┬────────────────────────┐
│ 항목                 │ Power Query (기간 펼치기) │ SQL (이벤트 기반)       │
├──────────────────────┼────────────────────────┼────────────────────────┤
│ 빈 월 처리           │ 복제해서 빈 행 생성      │ 행 없음 (이벤트만)      │
│                      │                        │                        │
│ CTE 수               │ 12단계                  │ 6단계                  │
│                      │                        │                        │
│ 재귀 CTE            │ 필요 (month_series)      │ 불필요                 │
│                      │                        │                        │
│ 행 수 (SO 100건,    │ ~2,400행                 │ ~400행                 │
│  평균 24개월 span)   │ (100 × 24)              │ (이벤트 있는 월만)      │
│                      │                        │                        │
│ "2026-05 잔고는?"    │ WHERE Period='2026-05'  │ SUM(net) WHERE ≤ 05    │
│                      │ (해당 행이 반드시 존재)   │ (누적 합산으로 계산)    │
│                      │                        │                        │
│ 결과                 │ 동일                    │ 동일 (이벤트 월)        │
│                      │ + 빈 월 filler 행       │                        │
└──────────────────────┴────────────────────────┴────────────────────────┘
```

### "빈 월 조회" 문제와 해결

기간 펼치기에서는 모든 월에 행이 있으므로 `WHERE Period = '2026-05'`로 아무 시점이나 조회 가능했습니다.
이벤트 기반에서는 활동 없는 월에 행이 없으므로, **누적 SUM 패턴**으로 해결합니다:

```sql
-- "2026-05 시점의 잔고를 알려줘"
SELECT SO_ID, [OS name],
    SUM(CASE WHEN Period = '2026-05' THEN Input ELSE 0 END)  AS Input,
    SUM(CASE WHEN Period = '2026-05' THEN Output ELSE 0 END) AS Output,
    SUM(CASE WHEN Period < '2026-05' THEN Input-Output ELSE 0 END) AS Start,
    SUM(Input - Output) AS Ending                    -- ← 전체 누적 = 잔고
FROM os_grouped
WHERE Period <= '2026-05'
GROUP BY SO_ID, [OS name]
```

예시: SOD-0001의 마지막 이벤트가 2월인데 5월 시점을 조회하면?
- 1~2월 이벤트가 `WHERE Period <= '2026-05'`에 포함됨
- `SUM(Input - Output)` = 15 - 15 = 0 → 잔고 없음 (정확!)
- 3~5월에 행이 없어도 누적 합산이므로 문제없음

### 정리

```
Power Query:  SO를 달력처럼 펼쳐서 → 각 칸에 값을 채우고 → 잔고 계산
              (직관적이지만, 빈 칸도 다 만들어야 해서 무거움)

SQL 이벤트:   실제 일어난 일(Input/Output)만 기록 → 필요할 때 누적 합산
              (은행 거래내역처럼, 잔고는 거래를 합산하면 언제든 알 수 있음)
```

두 방식 모두 **같은 결과**를 냅니다. 차이는 "빈 월에 행이 있느냐 없느냐"뿐.

---

## 사용 방법

### 쿼리 생성
1. **데이터** → **데이터 가져오기** → **다른 원본에서** → **빈 쿼리**
2. **홈** → **고급 편집기** → M 코드 붙여넣기
3. **닫기 및 로드**

### 데이터 갱신
- **Ctrl+Alt+F5** (모두 새로 고침)

### 필수 테이블
쿼리 실행 전 아래 시트들이 테이블로 정의되어 있어야 함:
- `SO_국내`, `SO_해외`
- `PO_국내`, `PO_해외`
- `DN_국내`, `DN_해외`

테이블 생성: 시트 선택 → **Ctrl+T** → 테이블 이름 지정

---

## 트러블슈팅

### 중복 행 발생
- **원인**: 조인 키(SO_ID, Line item)에 중복 데이터 존재
- **해결**: `Table.Distinct()` 사용하여 중복 제거

### 원가/출고금액이 null
- **원인**: PO 또는 DN에 해당 SO_ID + Line item 조합이 없음
- **확인**: 원본 시트에서 Line item 일치 여부 확인

### Sales = 0인 행이 SO_통합에서 누락 (2026-01-30 수정)
- **증상**: Sales amount = 0인 SO_ID가 SO_통합 쿼리 결과에서 빠짐
- **원인**: `[Status] <> "Cancelled"` 조건에서 Status가 null인 경우 Power Query가 해당 행을 제외
  - Power Query에서 `null <> "Cancelled"` → `null` 반환 → 행 제외
- **해결**: `[Status] = null or [Status] <> "Cancelled"` 로 수정
  - Cancelled만 제외하고 null 포함 나머지는 모두 포함

### PO_현황에서 null 오류 (2026-02-02 수정)
- **증상**: `Expression.Error: 값 null을(를) Logical 형식으로 변환할 수 없습니다`
- **원인 1**: Status가 null인 행에서 `[Status] = "Sent" or ...` 비교 오류
- **원인 2**: 미발주수량이 null인 행에서 `[미발주수량] <= 0` 비교 오류
  - Power Query에서 `null <= 0` → `null` 반환
  - `if null then` → Logical 변환 오류
- **해결**:
  ```
  // Status 필터링: List.Contains 사용
  each List.Contains({"Sent", "Confirmed", "Invoiced"}, [Status])

  // 미발주수량 계산: null을 0으로 대체
  each (if [SO수량] = null then 0 else [SO수량]) - (if [발주수량] = null then 0 else [발주수량])

  // 발주완료 판단: 명시적 null 체크
  each if [미발주수량] = null or [미발주수량] <= 0 then "Y" else "N"
  ```

### 무상 건(Sales = 0)이 출고완료 N으로 표시 (2026-01-30 수정)
- **증상**: 출고가 완료된 무상 건(SOD-2026-0017 등)이 출고완료 = N으로 표시
- **원인**: `[출고금액] > 0` 조건 때문에 출고금액이 0인 건은 출고완료 = N
- **해결**: `[출고금액] <> null` 로 수정
  - DN에 조인되면 (출고 기록이 있으면) 출고완료 = Y

### 무상 건의 "부분 출고"가 "출고 완료"로 표시 (2026-06-24 수정)
- **증상**: 부분만 출고된 무상공급 라인(SOD-2026-0301 Line 9, Eye bolt 576개 중 220개 출고)이 `출고 완료`로 표시
- **원인**: 출고완료 판정이 **금액 기반**(`[Sales amount KRW] - [출고금액] > 0`)이었음.
  무상공급은 단가 0 → Sales·출고금액 모두 0 → `0 - 0 > 0` = 거짓 → 부분출고를 영원히 감지 못 함.
  2026-01-30 패치(`[출고금액] = null`)는 무상 건의 "출고/미출고" 이분법만 고쳤고 "부분 출고"는 못 고침 — 같은 뿌리(금액 기반)의 미완성 패치.
- **해결**: 판정을 **수량 기반**으로 전환.
  - DN_Combined에 `출고수량 = List.Sum([Qty])` 추가, 판정식을 `[Item qty] - [출고수량] > 0` 로 변경
  - 무상공급뿐 아니라 해외 환율차로 금액이 어긋나는 케이스도 함께 해결됨
  - 금액 기반 `미출고금액` 컬럼은 재무 백로그용으로 그대로 유지 (무상 건은 0이 맞음)
  - **참고**: Order Book SQL(`sql/order_book.sql`)·대시보드 `load_so()`도 동일하게 수량 기준 — 세 군데 모두 통일

### 분할 출고 시 출고금액 일부만 매칭 (2026-02-05 수정 → 2026-02-07 조인 키 변경 → 2026-02-28 쿼리 수정)
- **증상**: SOO-2026-0011처럼 SO에 Line item 1개인데, DN에서 무게 등의 이유로 분할 출고 시 출고금액 일부만 매칭됨
- **원인**: SO_통합 쿼리의 `Table.Distinct(... {"SO_ID", "Line item"})`가 같은 Line item의 DN 행 중 첫 번째만 유지
  ```
  SO: Line item 1 (매출 300)
  DN: Line item 1 (출고 150)  ← 이것만 남음
  DN: Line item 1 (출고 150)  ← Table.Distinct가 버림
  결과: 출고금액 150 → 미출고 150 (오류)
  ```
- **해결**: `Table.Distinct` → `Table.Group`으로 변경하여 같은 Line item의 출고를 합산
  ```
  // Before (첫 번째 행만 유지)
  DN_Combined = Table.Distinct(Table.Combine({DN_국내_Renamed, DN_해외_Renamed}), {"SO_ID", "Line item"}),

  // After (합산)
  DN_Combined = Table.Group(Table.Combine({DN_국내_Renamed, DN_해외_Renamed}), {"SO_ID", "Line item"}, {
      {"출고금액", each List.Sum([출고금액]), Currency.Type},
      {"출고일", each List.Max([출고일]), type nullable date}
  }),
  ```
- **데이터 입력 규칙**: 분할 출고 시 DN의 Line item을 SO와 동일하게 유지 (SO를 분할할 필요 없음)
  ```
  SO: Line item 1 (매출 300)
  DN: Line item 1 (출고 150), Line item 1 (출고 150)  ← SO Line item 유지
  결과: Table.Group으로 합산 → 출고금액 300 → 미출고 0
  ```
- **참고**: 2026-02-07부터 모든 쿼리의 조인 키를 `SO_ID + Item name` → `SO_ID + Line item`으로 변경. Line item이 행의 유니크 키 역할을 하므로 Item name(설명 필드)보다 정확한 매칭 가능.

### 출고완료 상태를 3단계로 변경 (2026-02-28 수정)
- **배경**: 분할 출고 대응(DN Table.Group 합산)으로 부분 출고가 가능해졌으나, 기존 Y/N 이진 판단으로는 부분 출고를 표현할 수 없음
  - 출고금액이 존재하면 무조건 Y → 부분 출고도 "출고완료"로 표시되는 문제
- **해결**: ERP 방식의 3단계 상태로 변경
  ```
  // Before (Y/N)
  if [출고금액] <> null then "Y" else "N"

  // After (3단계)
  if [출고금액] = null then "미출고"
  else if [Sales amount KRW] - [출고금액] > 0 then "부분 출고"
  else "출고 완료"
  ```

### 출고완료 상태를 4단계로 세분화 (2026-02-28 수정)
- **배경**: "출고 완료"가 두 가지 다른 상황을 포함
  - 출고일 있는 출고 완료 = RCK가 **고객**에게 출고 완료
  - 출고일 없는 출고 완료 = **NOAH(공장)**에서 RCK에게 출고 완료 (고객 선적 전)
  - 주로 해외 오더에서 발생: 국내는 출고 다음날 도착하지만, 해외는 인코텀즈에 따라 운송 기간 소요
- **해결**: "공장 출고" 상태 추가
  ```
  // Before (3단계)
  if [출고금액] = null then "미출고"
  else if [Sales amount KRW] - [출고금액] > 0 then "부분 출고"
  else "출고 완료"

  // After (4단계)
  if [출고금액] = null then "미출고"
  else if [Sales amount KRW] - [출고금액] > 0 then "부분 출고"
  else if [출고일] = null then "공장 출고"
  else "출고 완료"
  ```
  | 조건 | 상태 | 설명 |
  |------|------|------|
  | 출고금액 = null | 미출고 | DN 기록 없음 |
  | 출고금액 < 매출 | 부분 출고 | 일부만 출고, 미출고금액 남아있음 |
  | 출고금액 >= 매출 & 출고일 = null | 공장 출고 | NOAH→RCK 출고 완료, 고객 선적 전 |
  | 출고금액 >= 매출 & 출고일 있음 | 출고 완료 | 고객에게 최종 출고 완료 |

### PO 사양 분리 시 원가 누락 (2026-02-27 수정)
- **증상**: SOO-2026-0041처럼 SO Line item 1개에 PO가 사양별로 여러 행인 경우, 원가가 첫 번째 행만 반영됨
  - Line 1: SO qty=3, PO에 SQ19×19(1개)+SQ17×17(2개) → 원가 3,624,865만 표시 (10,874,595여야 함)
  - Line 3: SO qty=21, PO에 SQ14×14(14개)+SQ17×17(7개) → 원가 28,393,442만 표시 (42,590,163여야 함)
- **원인**: `Table.Distinct(... {"SO_ID", "Line item"})` 가 같은 Line item의 PO 행 중 첫 번째만 남기고 나머지를 버림
  ```
  PO: Line 1, SQ19*19, qty=1, ICO=3,624,865  ← 이것만 남음
  PO: Line 1, SQ17*17, qty=2, ICO=7,249,730  ← 버려짐
  ```
- **해결**: `Table.Distinct` → `Table.Group` 으로 변경하여 같은 Line item의 원가를 합산
  ```
  // Before (첫 번째 행만 유지)
  PO_Combined = Table.Distinct(Table.Combine({PO_국내, PO_해외}), {"SO_ID", "Line item"}),

  // After (합산)
  PO_Combined = Table.Group(Table.Combine({PO_국내, PO_해외}), {"SO_ID", "Line item"}, {
      {"ICO Unit", each List.Average([ICO Unit]), type number},
      {"Total ICO", each List.Sum([Total ICO]), type number}
  }),
  ```
- **영향 범위**: SO_통합, DN_원가포함, Inventory_Transaction 세 쿼리 모두 수정
- **배경**: SO는 제품 레벨로 Line item을 관리하지만, PO는 같은 Line item 내에서 사양(밸브 사이즈 등)별로 행을 분리하는 경우가 있음 (1:N 관계)

### DN_해외 Total Sales KRW XLOOKUP 수식 오류 — 분할 출고 시 미출고금액 더블 계산 (2026-03-23 수정)
- **증상**: SOO-2026-0025처럼 SO 1건에 DN이 2건 이상(분할 출고)일 때, 미출고금액이 음수로 나옴 (더블 계산)
  ```
  SO: Line 1, qty=10, Sales KRW = 12,950,601
  DN: DNO-0034, Line 1, qty=5, Total Sales KRW = 12,950,601  ← SO 전체 금액
  DN: DNO-0035, Line 1, qty=5, Total Sales KRW = 12,950,601  ← SO 전체 금액 (또)
  SO_통합 출고금액 합계: 25,901,202 (2배)
  미출고금액 = 12,950,601 - 25,901,202 = -12,950,601 ❌
  ```
- **원인**: DN_해외 테이블의 `Total Sales KRW` 엑셀 수식이 SO의 전체 `Sales amount KRW`를 그대로 반환
  ```excel
  =XLOOKUP(B130&H130,SO_해외!A:A&SO_해외!R:R,SO_해외!V:V)
  ```
  DN의 Qty와 무관하게 SO의 전체 KRW 금액을 가져오므로, 분할 출고 시 각 DN에 전체 금액이 중복 배정됨
- **해결**: Qty 비율을 곱하여 비례 배분
  ```excel
  =XLOOKUP(B130&H130,SO_해외!A:A&SO_해외!R:R,SO_해외!V:V) * I130 / XLOOKUP(B130&H130,SO_해외!A:A&SO_해외!R:R,SO_해외!S:S)
  ```
  - `I130` = DN Qty, `SO_해외!S:S` = SO Item qty
  - `SO Sales KRW × (DN Qty / SO Qty)` = DN 비례 금액
  - 수학적 증명: `SO_Qty × UnitPrice × ExRate × (DN_Qty / SO_Qty) = DN_Qty × UnitPrice × ExRate` (SO_Qty 약분)
  - 전량 1회 출고 시 비율 = 1 → 기존과 동일, 분할 시 정확한 비례 배분

### 출고금액 = Sales KRW인데 "부분 출고"로 표시 (2026-03-23 수정)
- **증상**: SOO-2026-0020 Line 1처럼 SO qty = DN qty, 출고금액 = Sales KRW (표시상 동일)인데 "부분 출고"로 표시됨
- **원인**: SO_해외의 `Sales amount KRW`가 `=qty × unit_price × exchange_rate` 수식 결과로 소수점 이하 값을 가짐 (예: 10,840,025.77). 표시 서식은 정수로 보이지만 내부 값이 정수가 아님 → DN XLOOKUP과 비교 시 미세한 차이 발생 → `> 0` 조건 충족
- **해결**: SO_해외의 `Sales amount KRW` 수식에 `ROUND(..., 0)` 적용하여 정수로 저장
  - SO 원본이 정수 → DN XLOOKUP도 정수 → Power Query 비교 시 차이 = 0 (정확)

---

## 엑셀 수식 vs Power Query

### 왜 Power Query를 쓰는가?

엑셀 수식(VLOOKUP, XLOOKUP)과 Power Query의 핵심 차이는 **관계 처리 능력**이다.

| 관계 | 엑셀 수식 | Power Query |
|------|----------|-------------|
| **1:1** | O (VLOOKUP/XLOOKUP) | O |
| **1:N** | X (첫 번째만 반환) | O |
| **N:1** | O (각 행에서 조회) | O |
| **N:M** | X | O |

**Power Query = 엑셀에서 SQL 쓰는 것**과 같다.

---

### VLOOKUP/XLOOKUP의 한계 (1:1만 가능)

```
VLOOKUP / XLOOKUP 동작:
"찾으면 첫 번째 매칭 값 반환하고 끝"

SOD-2026-0001로 PO 조회하면?
├── POD-0001 (10개) ← 이것만 반환
├── POD-0005 (5개)  ← 무시됨
└── POD-0008 (3개)  ← 무시됨
```

**우회 방법은 있지만 복잡함:**
- SUMIF: 합계만 가능, 상세 내역 못 봄
- FILTER + 배열 수식: 행 펼치기 어려움
- TEXTJOIN + IF: 텍스트 연결만 가능

---

### LEFT JOIN (1:N - 행이 펼쳐짐)

Power Query의 `Table.NestedJoin`은 **매칭되는 모든 행**을 반환한다.

```
SO_국내 LEFT JOIN PO_국내:

┌──────────────┬────────┐      ┌──────────────┬─────┬───────────┐
│ SO_ID        │ Customer│      │ SO_ID        │ Qty │ 발주일     │
├──────────────┼────────┤      ├──────────────┼─────┼───────────┤
│ SOD-2026-0001│ 삼성전자 │  ←→  │ SOD-2026-0001│ 10  │ 1/15      │
└──────────────┴────────┘      │ SOD-2026-0001│ 5   │ 1/20      │
     (1행)                      │ SOD-2026-0001│ 3   │ 1/25      │
                               └──────────────┴─────┴───────────┘
                                    (3행)

결과: 1행이 3행으로 펼쳐짐
┌──────────────┬────────┬─────┬───────────┐
│ SO_ID        │ Customer│ Qty │ 발주일     │
├──────────────┼────────┼─────┼───────────┤
│ SOD-2026-0001│ 삼성전자 │ 10  │ 1/15      │
│ SOD-2026-0001│ 삼성전자 │ 5   │ 1/20      │
│ SOD-2026-0001│ 삼성전자 │ 3   │ 1/25      │
└──────────────┴────────┴─────┴───────────┘
```

**SQL로 표현:**
```sql
SELECT so.*, po.*
FROM SO_국내 so
LEFT JOIN PO_국내 po ON so.SO_ID = po.SO_ID
```

---

### 관계 유형별 이 프로젝트 사례

#### 1:1 - ICO 가격 조회 (XLOOKUP 가능)

```
PO_국내에서 ICO 가격 조회:
=XLOOKUP(Model & Option, ICO[Key], ICO[Price])

┌──────────┬───────┐      ┌──────────┬───────┬─────────┐
│ Model    │ Option│      │ Model    │ Option│ Price   │
├──────────┼───────┤      ├──────────┼───────┼─────────┤
│ IQ10     │ Bush  │ ───→ │ IQ10     │ Bush  │ 50,000  │ 1:1 매칭
└──────────┴───────┘      └──────────┴───────┴─────────┘
```

#### 1:N - SO → PO (추가 발주)

```
┌──────────────┬──────┬───────┐
│ SO_ID        │ Item │ 수량   │     SO 1건에 PO 여러 건
├──────────────┼──────┼───────┤
│ SOD-2026-0001│ IQ10 │ 15    │ ──┬── POD-0001 (10개) 1차 발주
└──────────────┴──────┴───────┘   ├── POD-0005 (5개)  추가 발주
                                  └── POD-0008 (3개)  추가 발주
```

**XLOOKUP**: 10만 반환 (첫 번째만)
**Power Query**: 3건 다 반환 → GROUP BY로 합계 = 18

#### 1:N - SO → DN (분할 납품)

```
┌──────────────┬──────┬───────┐
│ SO_ID        │ Item │ 수량   │     SO 1건에 DN 여러 건
├──────────────┼──────┼───────┤
│ SOD-2026-0002│ NA038│ 100   │ ──┬── DND-0010 (40개) 1차 납품
└──────────────┴──────┴───────┘   ├── DND-0015 (30개) 2차 납품
                                  └── DND-0020 (30개) 3차 납품
```

#### N:1 - DN → PO (같은 원가 참조)

```
┌───────┬──────────────┬──────┐
│ DN_ID │ SO_ID        │ Item │     DN 여러 건이 PO 1건 원가 참조
├───────┼──────────────┼──────┤
│ DN-010│ SOD-2026-0001│ IQ10 │ ──┐
│ DN-015│ SOD-2026-0001│ IQ10 │ ──┼→ ICO Unit: 500,000 (같은 원가)
│ DN-020│ SOD-2026-0001│ IQ10 │ ──┘
└───────┴──────────────┴──────┘
```

**비즈니스 의미**: 분할 납품 3번 했지만, 원가는 발주 시점에 정해진 거 하나

---

### 체인 조인 (여러 테이블 한번에)

#### 엑셀 수식으로 하면

DN_국내 시트에 수식 여러 개 필요:

```
원가 조회:     =XLOOKUP(SO_ID & Item, PO[Key], PO[ICO Unit])
AX번호 조회:   =XLOOKUP(SO_ID & Item, SO[Key], SO[Model code])
고객PO 조회:   =XLOOKUP(SO_ID & Item, SO[Key], SO[Customer PO])
...

→ 가져올 정보가 10개면 수식 10개
→ 시트 구조 변경되면 수식 다 수정
```

#### Power Query로 하면

```
DN_국내 (기준)
    │
    │ 1차 JOIN: PO에서 원가 가져오기
    ▼
┌─────────────────────────────────────────────────────────┐
│ Table.NestedJoin(DN, {"SO_ID", "Line item"},             │
│                  PO, {"SO_ID", "Line item"}, "PO_Data") │
└─────────────────────────────────────────────────────────┘
    │
    │ 2차 JOIN: SO에서 AX 정보 가져오기
    ▼
┌─────────────────────────────────────────────────────────┐
│ Table.NestedJoin(Result, {"SO_ID", "Line item"},         │
│                  SO, {"SO_ID", "Line item"}, "SO_Data") │
└─────────────────────────────────────────────────────────┘
    │
    ▼
최종 결과 (DN + PO + SO 정보가 한 테이블에)
```

**SQL로 표현:**
```sql
SELECT
    dn.*,
    po.ICO_Unit as 원가_단가,
    po.Total_ICO as 원가_합계,
    so.AX_Project_number,
    so.Customer_PO
FROM DN_국내 dn
LEFT JOIN PO_국내 po
    ON dn.SO_ID = po.SO_ID AND dn.Item = po.Item_name
LEFT JOIN SO_국내 so
    ON dn.SO_ID = so.SO_ID AND dn.Item = so.Item_name
```

#### 시각적으로 데이터 흐름

```
Step 1: DN만
┌───────┬──────────────┬──────┬─────┐
│ DN_ID │ SO_ID        │ Item │ Qty │
├───────┼──────────────┼──────┼─────┤
│ DN-010│ SOD-2026-0001│ IQ10 │ 10  │
└───────┴──────────────┴──────┴─────┘

Step 2: DN + PO (원가)
┌───────┬──────────────┬──────┬─────┬───────────┬───────────┐
│ DN_ID │ SO_ID        │ Item │ Qty │ 원가_단가  │ 원가_합계  │
├───────┼──────────────┼──────┼─────┼───────────┼───────────┤
│ DN-010│ SOD-2026-0001│ IQ10 │ 10  │ 500,000   │ 5,000,000 │
└───────┴──────────────┴──────┴─────┴───────────┴───────────┘
                                      ↑ PO에서 가져옴

Step 3: DN + PO + SO (AX 정보)
┌───────┬──────────────┬──────┬─────┬───────────┬─────────────────┬────────┐
│ DN_ID │ SO_ID        │ Item │ Qty │ 원가_합계  │ Model code      │
├───────┼──────────────┼──────┼─────┼───────────┼─────────────────┤
│ DN-010│ SOD-2026-0001│ IQ10 │ 10  │ 5,000,000 │ (없음)          │
└───────┴──────────────┴──────┴─────┴───────────┴─────────────────┘
                                                  ↑ SO에서 가져옴
```

---

### Join 종류 비교

| SQL | Power Query | 엑셀 수식 |
|-----|-------------|----------|
| `INNER JOIN` | `JoinKind.Inner` | X |
| `LEFT JOIN` | `JoinKind.LeftOuter` | VLOOKUP (1:1만) |
| `RIGHT JOIN` | `JoinKind.RightOuter` | X |
| `FULL JOIN` | `JoinKind.FullOuter` | X |
| `CROSS JOIN` | `Table.AddColumn` + 중첩 | X |

---

### 이 프로젝트에서 Power Query 활용 요약

| 쿼리 | 관계 | 목적 |
|------|------|------|
| `PO_현황` | SO ← PO (1:N) | 발주 합계, 미발주 현황 |
| `SO_통합` | SO ← PO (N:1), SO ← DN (N:1) | 마진 계산, 출고 상태 |
| `DN_원가포함` | DN ← PO ← SO (체인) | GL 대상 파악, IC Balance |
| `Inventory_Transaction` | DN → Receipt + Issue (분리) | 입출고 추적 |

---

### 결론: 언제 무엇을 쓰나

| 상황 | 도구 |
|------|------|
| 단순 1:1 조회 | XLOOKUP |
| 1:N 관계 (추가발주, 분할납품) | Power Query |
| 여러 테이블 조인 | Power Query |
| 집계 + 조건 필터 | Power Query |
| 복잡한 비즈니스 질문 | Power Query |

```
"SOD-2026-0001 발주 수량 합계?" → SUMIF 가능
"미발주 건 목록?" → Power Query 필요
"마진율 20% 이하인 건?" → Power Query 필요
"GL 분개 대상 금액?" → Power Query 필요 (3-way JOIN)
```

---

## 파워 쿼리 vs 파워 피벗

### 기능 비교

| 기능 | 파워 쿼리 | 파워 피벗 |
|------|----------|----------|
| 데이터 변환/정제 | O | X |
| 테이블 조인 | O | O |
| 계산 컬럼 | O | O |
| 그룹화/집계 | O | O |
| DAX 수식 | X | O |
| 동적 Measure | X | O |
| 피벗 테이블 연동 | 결과 테이블로 | 데이터 모델로 |

### 언제 파워 피벗이 필요한가

| 상황 | 파워 쿼리 | 파워 피벗 |
|------|----------|----------|
| 고정된 분석 뷰 | O | - |
| 마진, 출고완료 등 계산 | O | - |
| 사용자가 피벗으로 자유롭게 드릴다운 | △ | O |
| YTD, MTD, 전년비 등 시계열 분석 | X | O (DAX) |
| 여러 팩트 테이블 관계 | △ | O |

### 현재 상황 판단: 파워 쿼리로 충분

| 요소 | 현재 상황 | 판단 |
|------|----------|------|
| 데이터 규모 | 소규모 (ERP 통합 전 임시) | 파워 쿼리 OK |
| 분석 목적 | 명확함 (Backlog, 마진) | 고정 뷰로 충분 |
| 관계 구조 | 단순 (SO → PO → DN) | 파워 쿼리 조인 OK |
| 시계열 분석 | 없음 (YTD, 전년비 불필요) | DAX 불필요 |
| 사용자 | 본인 위주 | 동적 피벗 불필요 |

### 파워 쿼리로 가능한 분석

- **Backlog**: 미출고금액 합계 (SO_통합)
- **마진 분석**: 마진율 정렬/필터 (SO_통합)
- **국내/해외 구분**: 구분 컬럼 필터

### 파워 피벗이 필요해지는 시점

- 다른 팀원이 자유롭게 피벗으로 분석해야 할 때
- Period별 누적/비교 분석이 필요할 때
- 데이터가 수천 건 이상으로 늘어날 때

### 결론

```
현재: 파워 쿼리 → 결과 테이블 → 필터/정렬로 분석
미래: ERP 통합되면 이 엑셀 자체가 필요 없어짐
```

ERP 통합 전까지 임시 운영이므로, 파워 쿼리로 빠르게 뽑아 쓰는 게 효율적.
오버엔지니어링할 이유 없음.
