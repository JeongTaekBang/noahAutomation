# 데이터 추출 시 테이블 관계 가이드

ERP/Excel에서 데이터를 뽑아 대시보드를 만들 때, 테이블 간 관계를 모르면 데이터가 뻥튀기되거나 누락됩니다.

---

## 1. 테이블 관계 3가지

| 관계 | 예시 | 의미 |
|------|------|------|
| **1:1** | SO 헤더 ↔ SO 주소 | 1건 = 1건 |
| **1:N** | SO 헤더 → SO 라인 | 주문 1건에 품목 여러 개 |
| **N:M** | SO ↔ PO (NOAH) | 다대다 — 중간 매핑 필요 |

---

## 2. Key 먼저 파악하라

데이터를 뽑기 전에 **"이 두 테이블은 어떤 컬럼으로 연결되는가"**를 먼저 확인해야 합니다.

```
SO_ID → PO는 SO_ID로 연결
     → DN도 SO_ID로 연결
     → 하지만 Line item은 1:1이 아닐 수 있음
```

### NOAH에서 겪은 사례

SO↔PO를 `(SO_ID, Line item)`으로 조인 → 실패.
이유: PO에서 본체+부속을 1라인으로 합쳐서 발주하는 경우가 있어 SO Line item과 PO Line item이 1:1 대응하지 않음.
해결: `SO_ID` 단위로 집계 후 PO Status 기반으로 판정.

---

## 3. 헤더 vs 라인 구분

ERP는 대부분 이 구조:

```
주문 헤더 (SO_ID, 고객, 날짜, 상태)              ← 1건
    └── 주문 라인 (SO_ID + Line, 품목, 수량, 금액)  ← N건
```

Excel로 뽑으면 헤더 정보가 라인마다 반복됩니다.
이걸 그대로 합산하면 **헤더 값이 N배로 뻥튀기**됩니다.

```python
# OK — amount는 라인별 값
df.groupby("customer")["line_amount"].sum()

# 위험 — 헤더 값(할인율 등)을 라인에서 합산하면 중복
df.groupby("customer")["header_discount"].sum()  # 라인 수만큼 중복 합산됨
```

---

## 4. 집계 레벨을 맞춰라

두 테이블을 조인할 때 집계 단위가 다르면 데이터가 뻥튀기됩니다:

```
SO (1건) JOIN DN (3건) → 3행 (SO 금액이 3배로 뻥튀기)
```

해결: **조인 전에 한쪽을 먼저 집계**

```python
# DN을 SO_ID 단위로 먼저 집계
dn_agg = dn.groupby("SO_ID")["amount"].sum()

# 그 다음 SO와 조인 → 1:1
so.merge(dn_agg, on="SO_ID", how="left")
```

---

## 5. JOIN 유형별 의미

| JOIN | 결과 | 용도 |
|------|------|------|
| `INNER JOIN` | 양쪽 다 있는 것만 | SO와 DN 모두 있는 건 (출고 완료) |
| `LEFT JOIN` | 왼쪽 전부 + 오른쪽 매칭 | SO 전체 + PO 있으면 조인, 없으면 NULL |
| `RIGHT JOIN` | 오른쪽 전부 + 왼쪽 매칭 | 거의 안 씀 (LEFT로 뒤집으면 됨) |

**LEFT JOIN 후 NULL의 의미:**

```python
# SO LEFT JOIN PO → PO 컬럼이 NULL = 해당 SO에 PO가 없음 = 미발주
merged = so.merge(po, on="SO_ID", how="left")
미발주 = merged[merged["po_qty"].isna()]
```

---

## 6. D365 / AX 테이블 관계 패턴

### 회계 (General Ledger)

```
GeneralJournalEntry (분개 헤더)
    └── GeneralJournalAccountEntry (분개 라인)
            → MainAccount (계정과목 마스터)
            → DimensionAttributeValueCombination (부서/CC/프로젝트)

Trial Balance = 계정별 차변/대변 집계
```

### 판매 (Sales)

```
SalesTable (SO 헤더: SalesId, 고객, 날짜)
    └── SalesLine (SO 라인: SalesId + LineNum, 품목, 수량, 금액)

CustInvoiceJour (송장 헤더: InvoiceId)
    └── CustInvoiceTrans (송장 라인: InvoiceId + LineNum)
         ↑ SalesId로 SO와 연결
```

### 구매 (Purchase)

```
PurchTable (PO 헤더: PurchId)
    └── PurchLine (PO 라인: PurchId + LineNum)

VendInvoiceJour (공급자 송장 헤더)
    └── VendInvoiceTrans (공급자 송장 라인)
         ↑ PurchId로 PO와 연결
```

### 재고 (Inventory)

```
InventTable (품목 마스터: ItemId)
InventSum (재고 집계: ItemId + Dimension)
InventTrans (재고 트랜잭션: 입출고 이력)
```

---

## 7. 실전 체크리스트

데이터 추출 전 반드시 확인:

- [ ] 이 두 테이블의 **연결 키**가 뭔가? (1:1? 1:N?)
- [ ] 집계 단위가 **같은 레벨**인가? (헤더 vs 라인 혼용 아닌지)
- [ ] LEFT JOIN 후 **행 수가 늘어나지 않았나?** (뻥튀기 체크)
- [ ] **NULL 처리** — 조인 후 NULL은 뭘 의미하나?
- [ ] 금액 합산 시 **중복 계산** 안 되는가?

### 뻥튀기 감지 방법

```python
before = len(so)
after = len(so.merge(dn, on="SO_ID", how="left"))
if after > before:
    print(f"주의: {before}행 → {after}행 (뻥튀기 {after - before}행)")
    # → dn을 먼저 집계해야 함
```

---

## 8. NOAH 프로젝트 적용 사례

| 조인 | 키 | 주의점 |
|------|-----|--------|
| SO ↔ PO | `SO_ID` (Line item 아님) | PO가 본체+부속 합산 발주 → Line item 1:1 안 됨 |
| SO ↔ DN | `SO_ID` + `line_item` | DN은 SO Line과 1:1 대응 |
| SO ↔ Backlog | `SO_ID` + `OS name` | Order Book은 제품 단위 집계 |
| Coverage | `SO_ID` 단위 | PO 존재 + Status 기반 (수량 비교 아님) |
| Margin | `SO_ID` 단위 | PO 없으면 "원가 미확정" |
