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
| DN_원가포함 | 출고 내역 + 원가 + GL대상 여부 |
| SO_통합 | 주문 현황 + 원가 + 마진 + 출고 상태 |
| PO_현황 | 발주 현황 + Status별 집계 + 매입금액 |
| **PO_매입월별** | **월별 매입 집계 (IC Balance Confirmation용)** |
| **PO_AX대사** | **Period + AX PO(PXXXXXX)별 GRN 금액 집계 (회계 마감 대사용)** |
| **PO_미출고** | **Invoiced인데 DN 미매칭 건 (데이터 점검용)** |
| Inventory_Transaction | 입출고 트랜잭션 (감사 추적용) |
| **Order_Book** | **월별 수주잔고 (Backlog) 롤링 원장 - AX 오더북 형식** |

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
- DN_원가포함 쿼리의 **GL대상 = Y** 필터 → **원가_합계** 합계 = GL 분개 금액

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
- SO에서 AX Project number 조인 → GL대상 여부 판단

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
| AX Project number | ERP 프로젝트 번호 (SO에서) |
| AX PO | AX 발주번호 (PO에서) |
| Currency | 통화 (SO에서) |
| GL대상 | Y = AX 미등록 → 수기 분개 필요 |

### 용도
- **GL대상 = Y 필터** → 원가_합계 합계 = GL 분개 금액
- AX2009에 Item 미등록 시 임시 회계 처리:
  - `Dr. Inventory / Cr. AP-NOAH (IC)`
  - Item 등록 후 역분개
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

    SO_국내 = Table.SelectColumns(SO_국내_Raw, {"SO_ID", "Line item", "Opportunity", "Customer PO", "OS name", "Business registration number", "AX Project number", "Currency"}),
    SO_해외 = Table.SelectColumns(SO_해외_Raw, {"SO_ID", "Line item", "Opportunity", "Customer PO", "OS name", "Business registration number", "AX Project number", "Currency"}),
    SO_Combined = Table.Distinct(Table.Combine({SO_국내, SO_해외}), {"SO_ID", "Line item"}),

    // ========== 조인: DN + PO (원가) ==========
    WithCost = Table.NestedJoin(DN_Combined, {"SO_ID", "Line item"}, PO_Combined, {"SO_ID", "Line item"}, "PO_Data", JoinKind.LeftOuter),
    WithCostExpanded = Table.ExpandTableColumn(WithCost, "PO_Data", {"ICO Unit", "Total ICO", "AX PO"}, {"원가_단가", "원가_합계", "AX PO"}),

    // ========== 조인: + SO (SO_ID + Line item 기준) ==========
    WithAX = Table.NestedJoin(WithCostExpanded, {"SO_ID", "Line item"}, SO_Combined, {"SO_ID", "Line item"}, "SO_Data", JoinKind.LeftOuter),
    WithAXExpanded = Table.ExpandTableColumn(WithAX, "SO_Data", {"Opportunity", "Customer PO", "OS name", "Business registration number", "AX Project number", "Currency"}, {"Opportunity", "Customer PO", "OS name", "Business registration number", "AX Project number", "Currency"}),

    // ========== GL 대상 여부 ==========
    WithGLFlag = Table.AddColumn(WithAXExpanded, "GL대상", each if [AX Project number] = null or Text.Trim(Text.From([AX Project number])) = "" then "Y" else "N", type text),

    // ========== 타입 변환 ==========
    Result = Table.TransformColumnTypes(WithGLFlag, {
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
- PO에서 원가 조인 → 마진/마진율 계산
- DN에서 출고금액 조인 → 출고완료 여부, 미출고금액 계산
- Cancelled 건 제외

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
| 원가_단가 | ICO Unit |
| 원가 | Total ICO |
| 출고금액 | DN에서 출고된 금액 |
| 출고일 | DN에서 출고된 날짜 |
| 마진 | Sales - 원가 |
| 마진율 | 마진 / Sales (%) |
| 출고완료 | 미출고/부분 출고/공장 출고/출고 완료 |
| 매출연월 | 출고일 기준 연월 (yyyy-MM) |
| 미출고금액 | Sales - 출고금액 |

### 용도
- **마진율 정렬** → 수익성 낮은 건 파악
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

    // SO 합치기 + Cancelled 제외 (null은 포함)
    SO_Combined = Table.Combine({SO_국내_Tagged, SO_해외_Tagged}),
    SO_Filtered = Table.SelectRows(SO_Combined, each [Status] = null or [Status] <> "Cancelled"),
    SO_Final = Table.RemoveColumns(SO_Filtered, {"Status"}),

    // ========== PO 원가 (SO_ID + Line item 기준 합산) ==========
    PO_국내_Raw = Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content],
    PO_해외_Raw = Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content],

    PO_국내 = Table.SelectColumns(PO_국내_Raw, {"SO_ID", "Line item", "ICO Unit", "Total ICO"}),
    PO_해외 = Table.SelectColumns(PO_해외_Raw, {"SO_ID", "Line item", "ICO Unit", "Total ICO"}),
    PO_Combined = Table.Group(Table.Combine({PO_국내, PO_해외}), {"SO_ID", "Line item"}, {
        {"ICO Unit", each List.Average([ICO Unit]), type number},
        {"Total ICO", each List.Sum([Total ICO]), type number}
    }),

    // ========== DN 출고 (SO_ID + Line item 기준 합산) - 분할 출고 대응 ==========
    DN_국내_Raw = Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content],
    DN_해외_Raw = Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content],

    DN_국내 = Table.SelectColumns(DN_국내_Raw, {"SO_ID", "Line item", "Total Sales", "출고일"}),
    DN_국내_Renamed = Table.RenameColumns(DN_국내, {{"Total Sales", "출고금액"}}),
    DN_해외 = Table.SelectColumns(DN_해외_Raw, {"SO_ID", "Line item", "Total Sales KRW", "선적일"}),
    DN_해외_Renamed = Table.RenameColumns(DN_해외, {{"Total Sales KRW", "출고금액"}, {"선적일", "출고일"}}),
    DN_Combined = Table.Group(Table.Combine({DN_국내_Renamed, DN_해외_Renamed}), {"SO_ID", "Line item"}, {
        {"출고금액", each List.Sum([출고금액]), type number},
        {"출고일", each List.Max([출고일]), type nullable date}
    }),

    // ========== SO에 원가 조인 (SO_ID + Line item) ==========
    WithCost = Table.NestedJoin(SO_Final, {"SO_ID", "Line item"}, PO_Combined, {"SO_ID", "Line item"}, "PO", JoinKind.LeftOuter),
    WithCostExpanded = Table.ExpandTableColumn(WithCost, "PO", {"ICO Unit", "Total ICO"}, {"원가_단가", "원가"}),

    // ========== SO에 출고 조인 (SO_ID + Line item) - 출고일 포함 ==========
    WithShip = Table.NestedJoin(WithCostExpanded, {"SO_ID", "Line item"}, DN_Combined, {"SO_ID", "Line item"}, "DN", JoinKind.LeftOuter),
    WithShipExpanded = Table.ExpandTableColumn(WithShip, "DN", {"출고금액", "출고일"}, {"출고금액", "출고일"}),

    // ========== 계산 컬럼 추가 ==========
    WithMargin = Table.AddColumn(WithShipExpanded, "마진", each [Sales amount KRW] - (if [원가] = null then 0 else [원가]), type number),
    WithMarginRate = Table.AddColumn(WithMargin, "마진율", each if [Sales amount KRW] = 0 or [Sales amount KRW] = null then null else [마진] / [Sales amount KRW], Percentage.Type),
    WithShipStatus = Table.AddColumn(WithMarginRate, "출고완료", each
        if [출고금액] = null then "미출고"
        else if [Sales amount KRW] - [출고금액] > 0 then "부분 출고"
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
    })
in
    Result
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
- Invoiced 건(공장 출고 완료)의 매입금액을 월별로 집계
- DN 시트의 출고일 기준으로 월 산정

### 데이터 흐름
```
PO (Invoiced 건만)
    │
    │ JOIN (SO_ID + Item)
    ▼
DN (출고일)
    │
    │ 월 추출 + 그룹화
    ▼
월별 매입 집계
```

### 결과 컬럼
| 컬럼 | 설명 |
|------|------|
| 매입월 | 출고일 기준 월 (yyyy-MM 형식) |
| 구분 | 국내/해외 |
| 매입건수 | Invoiced PO 건수 |
| 매입수량 | Invoiced 수량 합계 |
| 매입금액 | Invoiced Total ICO 합계 (= RCK AP) |

### M 코드

```
let
    // ========== PO 원본 (Invoiced 건만) ==========
    PO_국내_Raw = Excel.CurrentWorkbook(){[Name="PO_국내"]}[Content],
    PO_해외_Raw = Excel.CurrentWorkbook(){[Name="PO_해외"]}[Content],

    PO_국내 = Table.SelectColumns(PO_국내_Raw, {"PO_ID", "SO_ID", "Line item", "Item qty", "Total ICO", "Status"}),
    PO_국내_Tagged = Table.AddColumn(PO_국내, "구분", each "국내"),

    PO_해외 = Table.SelectColumns(PO_해외_Raw, {"PO_ID", "SO_ID", "Line item", "Item qty", "Total ICO", "Status"}),
    PO_해외_Tagged = Table.AddColumn(PO_해외, "구분", each "해외"),

    PO_Combined = Table.Combine({PO_국내_Tagged, PO_해외_Tagged}),

    // Invoiced P01, P02 등만 필터 (공장 출고 완료 = RCK 매입)
    PO_Invoiced = Table.SelectRows(PO_Combined, each [Status] <> null and Text.StartsWith([Status], "Invoiced")),

    // ========== DN 원본 (출고일) ==========
    DN_국내_Raw = Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content],
    DN_해외_Raw = Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content],

    DN_국내 = Table.SelectColumns(DN_국내_Raw, {"SO_ID", "Line item", "출고일"}),

    // 해외: 출고일 사용 (= 공장 출고일, 선적일과 별도)
    DN_해외 = Table.SelectColumns(DN_해외_Raw, {"SO_ID", "Line item", "출고일"}),

    DN_Combined = Table.Combine({DN_국내, DN_해외}),

    // SO_ID + Line item 기준 출고일 집계 (같은 조합에 여러 DN이 있으면 최신 출고일)
    DN_Grouped = Table.Group(DN_Combined, {"SO_ID", "Line item"}, {
        {"출고일", each List.Max([출고일]), type date}
    }),

    // ========== PO + DN 조인 (출고일 가져오기) ==========
    WithDate = Table.NestedJoin(PO_Invoiced, {"SO_ID", "Line item"}, DN_Grouped, {"SO_ID", "Line item"}, "DN_Data", JoinKind.LeftOuter),
    WithDateExpanded = Table.ExpandTableColumn(WithDate, "DN_Data", {"출고일"}, {"출고일"}),

    // ========== 미출고 건 제외 (PO_미출고 쿼리에서 별도 확인) ==========
    WithValidDate = Table.SelectRows(WithDateExpanded, each
        [출고일] <> null and
        (try Date.Year([출고일]) otherwise null) <> null
    ),

    // ========== 매입월 추출 ==========
    WithMonth = Table.AddColumn(WithValidDate, "매입월", each
        Text.From(Date.Year([출고일])) & "-" & Text.PadStart(Text.From(Date.Month([출고일])), 2, "0"),
        type text),

    // ========== 월 + 구분 기준 그룹화 ==========
    Grouped = Table.Group(WithMonth, {"매입월", "구분"}, {
        {"매입건수", each Table.RowCount(_), Int64.Type},
        {"매입수량", each List.Sum([Item qty]), Int64.Type},
        {"매입금액", each List.Sum([Total ICO]), Currency.Type}
    }),

    // ========== 정렬 ==========
    Sorted = Table.Sort(Grouped, {{"매입월", Order.Descending}, {"구분", Order.Ascending}}),

    // ========== 타입 변환 ==========
    Result = Table.TransformColumnTypes(Sorted, {
        {"매입건수", Int64.Type},
        {"매입수량", Int64.Type},
        {"매입금액", Currency.Type}
    })
in
    Result
```

### 결과 예시

| 매입월 | 구분 | 매입건수 | 매입수량 | 매입금액 |
|--------|------|----------|----------|----------|
| 2026-02 | 국내 | 5 | 25 | 12,500,000 |
| 2026-02 | 해외 | 2 | 10 | 8,000,000 |
| 2026-01 | 국내 | 8 | 40 | 20,000,000 |
| 2026-01 | 해외 | 3 | 15 | 12,000,000 |

**참고**: 미출고 건은 **PO_미출고** 쿼리에서 확인

### 용도

| 필터/분석 | 용도 |
|----------|------|
| 특정 월 필터 | 해당 월 IC Balance Confirmation |
| 구분별 합계 | 국내/해외 매입 비교 |

**IC Balance 확인 방법**:
```
NOAH AR Statement (2026-01월)  vs  PO_매입월별 (매입월 = 2026-01) 매입금액 합계
```

---

## PO_AX대사

### 목적
- **회계 마감 시 AX GRN 대사** 용도
- Invoiced(GRN 처리 완료) 건만 대상
- Period + AX PO(PXXXXXX)별 금액 집계
- AX에 입력된 GRN 금액과 엑셀 매입금액 비교

### 대사 프로세스
```
엑셀 (PO_AX대사 쿼리)                          AX (D365 F&O) GRN
┌─────────────────────────────────┐        ┌─────────────────────────────────┐
│ 2026-01  P000001  ₩7,500,000   │  ──→   │ 2026-01  P000001  ₩7,500,000   │  ✓ 일치
│ 2026-01  P000002  ₩8,000,000   │  ──→   │ 2026-01  P000002  ₩8,000,000   │  ✓ 일치
│ 2026-02  P000003  ₩6,000,000   │  ──→   │ 2026-02  P000003  ₩4,000,000   │  ✗ 불일치
└─────────────────────────────────┘        └─────────────────────────────────┘
```

### 결과 컬럼

| 컬럼 | 설명 |
|------|------|
| Period | 출고일 기준 월 (yyyy-MM 형식) |
| AX PO | AX 발주번호 (PXXXXXX) |
| 구분 | 국내/해외 |
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
    WithDate = Table.NestedJoin(PO_WithAX, {"SO_ID", "Line item"}, DN_Grouped, {"SO_ID", "Line item"}, "DN_Data", JoinKind.LeftOuter),
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

    // ========== Period + AX PO 기준 그룹화 ==========
    Grouped = Table.Group(WithPeriod, {"Period", "AX PO", "구분"}, {
        {"PO_ID", each Text.Combine(List.Distinct([PO_ID]), ", "), type text},
        {"건수", each Table.RowCount(_), Int64.Type},
        {"수량", each List.Sum([Item qty]), type number},
        {"금액", each List.Sum([Total ICO]), type number}
    }),

    // ========== 정렬 ==========
    Sorted = Table.Sort(Grouped, {
        {"Period", Order.Descending},
        {"AX PO", Order.Ascending}
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

| Period | AX PO | 구분 | PO_ID | 건수 | 수량 | 금액 |
|--------|-------|------|-------|------|------|------|
| 2026-02 | P000003 | 국내 | POD-0007 | 2 | 10 | 4,000,000 |
| 2026-02 | P000005 | 해외 | POO-0003 | 1 | 5 | 3,500,000 |
| 2026-01 | P000001 | 국내 | POD-0001, POD-0005 | 4 | 15 | 7,500,000 |
| 2026-01 | P000002 | 국내 | POD-0002 | 2 | 10 | 8,000,000 |
| 2026-01 | P000004 | 해외 | POO-0001 | 3 | 20 | 6,000,000 |

### 용도

| 필터/분석 | 용도 |
|----------|------|
| Period = "2026-01" | 해당 월 마감 대사 (월별 필터링) |
| 특정 AX PO | AX GRN 금액과 1:1 비교 |
| 구분별 소계 | 국내/해외 AP 분리 확인 |

**AX 대사 방법**:
```
PO_AX대사 (Period = 2026-01) 금액 합계  vs  AX D365 F&O GRN (2026-01월) 금액
→ PXXXXXX별 1:1 대사, 불일치 시 PO_ID로 라인별 추적
```

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
| ① 마지막 출고월 붙이기 | "이 주문 언제 끝나?" 끝점 파악 | 달력을 어디까지 펼칠지 |
| ② Period 확장 | 1줄을 월별로 복제 | 빈 달력 만들기 |
| ③ Input 채우기 | 등록월에만 수주 금액 기록 | 입금 기록 |
| ④ Output 채우기 | DN 출고를 해당 월에 매칭 | 출금 기록 |
| ⑤ Line item 합치기 | 같은 제품+납기일끼리 합산 | 정리 |
| ⑥ 잔고 계산 | Start + Input - Output = Ending | 통장 잔고 |

**①~④ = 빈 달력 만들어서 채우기, ⑤ = 정리, ⑥ = 통장 잔고 계산**

### 개념

```
수주잔고 Order Book = 오더의 흐름을 월별로 추적하는 원장

  Start (전월 이월)
+ Input (당월 수주 = SO 등록)
- Output (당월 매출 = DN 출고/선적)
+ Variance (조정분, 현재 0)
= Ending (당월 잔고 → 다음 달 Start로 이월)
```

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
- Input = SO의 `Sales amount KRW` (수주 금액)
- Output = DN의 `Total Sales` / `Total Sales KRW` (실제 매출 금액)
- 국내: DN 출고일 기준, **해외: DN 선적일 기준** (매출 인식 시점)
- **분할 출고 대응**: 같은 SO+Line item에 DN이 여러 건이면 각 출고월에 해당 수량/금액 배분

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
- **AX Project number**: 그룹 내 고유값을 `, `로 연결 (예: "P001, P002")
- **처리 순서**: Line item 레벨에서 Input/Output 계산 → OS name + 납기일로 그룹화 → 롤링 계산

### 결과 컬럼

| 컬럼 | 설명 |
|------|------|
| Period | 해당 월 (yyyy-MM, 텍스트) |
| 구분 | 국내/해외 |
| SO_ID | 주문 번호 |
| Customer name | 고객명 |
| Customer PO | 고객 발주번호 (대표값) |
| Item name | 아이템명 (대표값) |
| OS name | OneStream Item name (**그룹화 키**) |
| Expected delivery date | 예상 납기일 (**그룹화 키**, 같은 OS name이라도 납기일 다르면 구분) |
| AX Period | AX 기간 (그룹 내 고유값 연결) |
| AX Project number | ERP 프로젝트 번호 (그룹 내 고유값 연결) |
| Sector | 사업 부문 |
| Business registration number | 사업자등록번호 |
| Industry code | 산업 코드 |
| Value_Start_qty | 전월 이월 수량 |
| Value_Input_qty | 당월 수주 수량 (등록 Period에만) |
| Value_Output_qty | 당월 출고 수량 (출고월에만, DN 기준) |
| Value_Variance_qty | 수량 조정분 (현재 0, 향후 확장용) |
| Value_Ending_qty | Start + Input - Output + Variance |
| Value_Start_amount | 전월 이월 금액 |
| Value_Input_amount | 당월 수주 금액 (SO 기준) |
| Value_Output_amount | 당월 매출 금액 (DN 기준) |
| Value_Variance_amount | 금액 조정분 (현재 0, 향후 확장용) |
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
        │  SO_Filtered  │  Cancelled 제외, #N/A 치환
        │  (전체 수주)   │  Period 빈 행 제외
        └──────┬───────┘
               │
               │  ┌─────────────┐     ┌─────────────┐
               │  │  DN_국내     │     │  DN_해외     │
               │  │  (출고 원본)  │     │  (출고 원본)  │
               │  └──────┬──────┘     └──────┬──────┘
               │         └────────┬─────────┘
               │                  ▼
               │          ┌──────────────┐
               │          │  DN_Combined  │  출고일→출고월 변환
               │          └──────┬───────┘
               │                 │
               │          ┌──────┴──────────────────┐
               │          ▼                         ▼
               │  ┌──────────────┐         ┌──────────────────┐
               │  │ DN_LastMonth  │         │   DN_ByMonth      │
               │  │ SO+Line별     │         │   SO+Line+출고월별  │
               │  │ 마지막 출고월  │         │   월별 qty/amount  │
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
    끝점 = 마지막출고월 or 현재월                        │
                                                    │
    출고 완료 → 마지막출고월에서 끊음                     │
    미출고    → 현재월까지 (Backlog)                    │
                                                    │
    * 출고 완료 건이 이후 Period에                      │
      빈 행으로 남는 것을 방지                          │
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
    Period = 출고월 매칭
    → 월별 출고 배분
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
│ SOD-0001  │ 1     │ IQ3  │ 2026-01  │ 2026-02 │  ←        ~출고월 (2월까지 활동)
│ SOD-0001  │ 2     │ IQ3  │ 2026-01  │ 2026-01 │  ← 1월만 (1월 출고완료라 여기서 끊음)
└───────────┴───────┴──────┴──────────┴─────────┘
  출고 완료 건 = 마지막 출고월에서 끊음 (이후 빈 행 방지)
  미출고 건   = 현재월까지 계속 펼침 (Backlog)
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
    // #N/A 등 에러 값을 null로 치환 (XLOOKUP 실패 등)
    SO_CleanErrors = Table.ReplaceErrorValues(SO_Combined, {{"Expected delivery date", null}}),
    // Cancelled 제외, Period 비어있는 행 제외
    SO_Filtered = Table.SelectRows(SO_CleanErrors, each
        ([Status] = null or [Status] <> "Cancelled") and
        [Period] <> null and Text.Trim(Text.From([Period])) <> ""
    ),

    // ========== DN (출고 시점 + 실제 매출 금액) ==========
    DN_국내_Raw = Excel.CurrentWorkbook(){[Name="DN_국내"]}[Content],
    DN_해외_Raw = Excel.CurrentWorkbook(){[Name="DN_해외"]}[Content],

    // 국내: 출고일 기준, Total Sales = 매출
    DN_국내 = Table.SelectColumns(DN_국내_Raw, {"SO_ID", "Line item", "Qty", "Total Sales", "출고일"}),
    DN_국내_WithPeriod = Table.AddColumn(DN_국내, "출고월", each
        if [출고일] = null then null
        else Text.From(Date.Year([출고일])) & "-" & Text.PadStart(Text.From(Date.Month([출고일])), 2, "0"),
        type text),
    DN_국내_Final = Table.RenameColumns(DN_국내_WithPeriod, {{"Total Sales", "출고금액"}}),

    // 해외: 선적일 기준 (매출 인식 시점), Total Sales KRW = 매출
    DN_해외 = Table.SelectColumns(DN_해외_Raw, {"SO_ID", "Line item", "Qty", "Total Sales KRW", "선적일"}),
    DN_해외_WithPeriod = Table.AddColumn(DN_해외, "출고월", each
        if [선적일] = null then null
        else Text.From(Date.Year([선적일])) & "-" & Text.PadStart(Text.From(Date.Month([선적일])), 2, "0"),
        type text),
    DN_해외_Final = Table.RenameColumns(DN_해외_WithPeriod, {{"Total Sales KRW", "출고금액"}}),

    DN_Combined = Table.Combine({DN_국내_Final, DN_해외_Final}),

    // DN 월별 집계 (분할 출고 대응: SO_ID + Line item + 출고월)
    DN_ByMonth = Table.Group(DN_Combined, {"SO_ID", "Line item", "출고월"}, {
        {"Output_qty", each List.Sum([Qty]), type number},
        {"Output_amount", each List.Sum([출고금액]), type number}
    }),

    // DN 마지막 출고월 (ActivePeriods 범위 결정용)
    DN_LastMonth = Table.Group(DN_Combined, {"SO_ID", "Line item"}, {
        {"출고월", each List.Max(List.RemoveNulls([출고월])), type text}
    }),

    // ========== SO + DN 조인 (기간 범위용, 마지막 출고월만) ==========
    WithDN = Table.NestedJoin(SO_Filtered, {"SO_ID", "Line item"}, DN_LastMonth, {"SO_ID", "Line item"}, "DN_Data", JoinKind.LeftOuter),
    WithDNExpanded = Table.ExpandTableColumn(WithDN, "DN_Data", {"출고월"}),

    // ========== Period 리스트 ==========
    // SO 등록 Period + DN 모든 출고월 (분할 출고 중간 월 누락 방지)
    AllPeriods = List.Buffer(List.Sort(List.Distinct(
        List.RemoveNulls(Table.Column(WithDNExpanded, "Period")) &
        List.RemoveNulls(Table.Column(DN_ByMonth, "출고월"))
    ))),
    LastPeriod = List.Last(AllPeriods),

    // ========== 건별 × Period 확장 ==========
    // 각 SO Line: 등록 Period ~ 출고월(또는 마지막 Period)까지 행 생성
    // 출고 완료 건: 출고월까지만 표시
    // 미출고 건: 마지막 Period까지 표시 (Backlog)
    WithPeriodList = Table.AddColumn(WithDNExpanded, "ActivePeriods", each
        let
            startIdx = List.PositionOf(AllPeriods, [Period]),
            endPeriod = if [출고월] <> null then [출고월] else LastPeriod,
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
    // DN_ByMonth와 SO_ID + Line item + Period = 출고월 조인 → 분할 출고 월별 배분
    WithDNOutput = Table.NestedJoin(WithInputAmt, {"SO_ID", "Line item", "Period"}, DN_ByMonth, {"SO_ID", "Line item", "출고월"}, "DN_Output", JoinKind.LeftOuter),
    WithDNOutputExpanded = Table.ExpandTableColumn(WithDNOutput, "DN_Output", {"Output_qty", "Output_amount"}),
    WithOutputQty = Table.AddColumn(WithDNOutputExpanded, "Value_Output_qty", each
        if [Output_qty] = null then 0 else [Output_qty], type number),
    WithValues = Table.AddColumn(WithOutputQty, "Value_Output_amount", each
        if [Output_amount] = null then 0 else [Output_amount], type number),
    WithValuesCleaned = Table.RemoveColumns(WithValues, {"Output_qty", "Output_amount"}),

    // ========== OS name 기준 그룹화 ==========
    // Line item 레벨 → SO_ID + OS name + Expected delivery date + Period 로 합산
    // 같은 OS name의 Line item들이 하나의 행으로 합쳐짐
    OSGrouped = Table.Group(WithValuesCleaned, {"SO_ID", "OS name", "Expected delivery date", "Period"}, {
        {"Customer name", each List.First([Customer name]), type text},
        {"Customer PO", each List.First([Customer PO]), type text},
        {"Item name", each List.First([Item name]), type text},
        {"구분", each List.First([구분]), type text},
        {"Sector", each List.First([Sector]), type text},
        {"Business registration number", each List.First([Business registration number]), type text},
        {"Industry code", each List.First([Industry code]), type text},
        {"AX Period", each Text.Combine(List.Distinct(List.RemoveNulls([AX Period])), ", "), type text},
        {"AX Project number", each Text.Combine(List.Distinct(List.RemoveNulls([AX Project number])), ", "), type text},
        {"Value_Input_qty", each List.Sum([Value_Input_qty]), type number},
        {"Value_Input_amount", each List.Sum([Value_Input_amount]), type number},
        {"Value_Output_qty", each List.Sum([Value_Output_qty]), type number},
        {"Value_Output_amount", each List.Sum([Value_Output_amount]), type number}
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
                        Value_Variance_amount = 0,
                        Value_Ending_amount = sAmt + r[Value_Input_amount] - r[Value_Output_amount]
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
        "Period", "구분", "SO_ID", "Customer name", "Customer PO", "Item name", "OS name",
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
    })
in
    Result
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

### AX 오더북과의 비교

| 항목 | AX2009 Order Book | NOAH Order_Book |
|------|-------------------|-----------------|
| 마감 | Period 마감 → 잠금 | 없음 (매번 재계산) |
| Start 이월 | DB에 저장된 값 | Power Query가 계산한 값 |
| Variance | 자동 추적 (금액 변경, 취소) | 불필요 (조정분을 새 Line item으로 추가 → Input에서 넷팅) |
| 스냅샷 | DB에 보존 | 없음 (현재 데이터 기준) |
| 그룹화 | Project number 기준 | SO_ID + OS name 기준 (Line item 합산) |
| Input 기준 | SO 등록일 | SO의 Period 컬럼 (yyyy-MM) |
| Output 기준 | Invoice 일자 | DN 출고일(국내) / 선적일(해외) |
| 금액 기준 | SO 금액 | Input=SO, Output=DN (차이 감지 가능) |
| 갱신 | 자동 (트랜잭션 기반) | Ctrl+Alt+F5 (수동 새로고침) |

### 전제조건

- SO의 **Period 컬럼**: `yyyy-MM` 형식 텍스트 (예: "2026-01")
- SO의 **Sector 컬럼**: 사업 부문 (예: "Process", "CPI", "Water")
- SO의 **Business registration number 컬럼**: 사업자등록번호
- SO의 **Industry code 컬럼**: 산업 코드
- DN의 **출고일/선적일**: 날짜 형식 → 쿼리에서 yyyy-MM으로 변환

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
- Variance 컬럼 불필요 (스냅샷 없이도 조정 이력 관리 가능)
- Variance 컬럼은 M 코드에 구조만 유지 (현재 0, 향후 스냅샷 도입 시 활용 가능)

### 한계

| 한계 | 설명 | 대응 |
|------|------|------|
| 마감 잠금 없음 | 과거 SO/DN 수정 시 소급 변경 | 원래 행 수정 금지, 조정은 새 Line item으로 추가 |
| 스냅샷 없음 | 과거 시점 재현 불가 | SO raw 데이터에 원본+조정 이력이 남아 추적 가능 |
| Period 갭 | 활동 없는 월은 행 생성 안됨 | 전월 Ending이 다음 활동월 Start로 정확히 이월됨 |

> **향후 확장**: ERP 통합 등으로 스냅샷 기반 Variance 추적이 필요해지면, M 코드의 `Value_Variance_qty/amount` 컬럼(현재 0)에 스냅샷 대비 차이를 계산하는 로직을 추가할 수 있음. 현재 분개 방식과 병행 가능.

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
- **원인**: 조인 키(SO_ID, Item name)에 중복 데이터 존재
- **해결**: `Table.Distinct()` 사용하여 중복 제거

### 원가/출고금액이 null
- **원인**: PO 또는 DN에 해당 SO_ID + Item 조합이 없음
- **확인**: 원본 시트에서 Item name 일치 여부 확인 (띄어쓰기, 괄호 등)

### GL대상 판단 기준
- AX Project number가 없거나 빈 문자열이면 `Y`
- AX에 Item 등록 후 프로젝트 번호가 부여되면 `N`으로 변경됨

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
      {"출고금액", each List.Sum([출고금액]), type number},
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
AX번호 조회:   =XLOOKUP(SO_ID & Item, SO[Key], SO[AX Project number])
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
│ Table.NestedJoin(DN, {"SO_ID", "Item"},                 │
│                  PO, {"SO_ID", "Item name"}, "PO_Data") │
└─────────────────────────────────────────────────────────┘
    │
    │ 2차 JOIN: SO에서 AX 정보 가져오기
    ▼
┌─────────────────────────────────────────────────────────┐
│ Table.NestedJoin(Result, {"SO_ID", "Item"},             │
│                  SO, {"SO_ID", "Item name"}, "SO_Data") │
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
│ DN_ID │ SO_ID        │ Item │ Qty │ 원가_합계  │ AX Project no   │ GL대상 │
├───────┼──────────────┼──────┼─────┼───────────┼─────────────────┼────────┤
│ DN-010│ SOD-2026-0001│ IQ10 │ 10  │ 5,000,000 │ (없음)          │ Y      │
└───────┴──────────────┴──────┴─────┴───────────┴─────────────────┴────────┘
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
| 분석 목적 | 명확함 (Backlog, GL대상, 마진) | 고정 뷰로 충분 |
| 관계 구조 | 단순 (SO → PO → DN) | 파워 쿼리 조인 OK |
| 시계열 분석 | 없음 (YTD, 전년비 불필요) | DAX 불필요 |
| 사용자 | 본인 위주 | 동적 피벗 불필요 |

### 파워 쿼리로 가능한 분석

- **Backlog**: 미출고금액 합계 (SO_통합)
- **마진 분석**: 마진율 정렬/필터 (SO_통합)
- **GL 대상**: GL대상=Y 필터 (DN_원가포함)
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
