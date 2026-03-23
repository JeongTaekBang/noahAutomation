# Template Cell Mappings

각 문서 유형별 템플릿 셀 매핑 정보입니다.

---

## 거래명세표 (`templates/ts_template_local.xlsx`)

**생성기**: `ts_generator.py` (xlwings)
**데이터 소스**: DN_국내 (납품) / PMT_국내 (선수금)

### 고정 필드 (Header)

| 셀 | 필드명 | 설명 | 데이터 소스 |
|----|--------|------|-------------|
| B2 | DATE | 날짜 (`DATE : YYYY. MM. DD`) | 출고일 (dispatch_date) |
| B7 | 고객명 | `{고객명} 귀하` | customer_name |

### 동적 필드 (Item List)

아이템 시작 행은 동적 탐지 (`TS_HEADER_LABELS`로 `월/일`, `품명` 등 검색).

| 열 | 필드명 | 설명 |
|----|--------|------|
| A | 월/일 | 출고일 (MM/DD) |
| B~D | 품명 | 품목명 (item_name) |
| E | 수량 | item_qty |
| F | 단가 | unit_price |
| G | 공급가 | 수량 × 단가 |
| H | 세액 | 공급가 × 10% |

### 하단 영역 (동적 위치)

| 위치 | 필드명 | 설명 |
|------|--------|------|
| PO No 행 (B열) | PO No | Customer PO 번호 |
| 소계 행 | SUM | E/G/H열 합계 수식 |
| 합계 행 (G열) | 합계 | 공급가 + 세액 |

---

## PI — Proforma Invoice (`templates/proforma_invoice.xlsx`)

**생성기**: `pi_generator.py` (xlwings)
**데이터 소스**: SO_해외

### 고정 필드 (Header)

| 셀 | 필드명 | 설명 | 데이터 소스 |
|----|--------|------|-------------|
| A9 | Consigned to | 수취인 주소 (Customer address) | SO_해외 |
| A10 | Consigned country | 수취인 국가 | SO_해외 |
| C10 | Consigned TEL | 수취인 전화번호 | SO_해외 |
| E10 | Consigned FAX | 수취인 팩스번호 | SO_해외 |
| A12 | Vessel Name or Flight No | 선박명 또는 항공편 | — |
| B13 | From | 출발지 | 고정값 "INCHEON, KOREA" |
| B14 | Destination country | 도착 국가 | SO_해외 |
| D15 | Departs on or about | 출발 예정일 | — |
| G4 | Invoice No | SO_ID | SO_해외.SO_ID |
| G5 | L/C No | 신용장 번호 | SO_해외 |
| G15 | PO No (Customer) | 고객 PO 번호 | SO_해외.Customer PO |
| G17 | Incoterms | 인코텀즈 | SO_해외.Incoterms |
| I4 | Invoice date | 인보이스 발행일 | 오늘 날짜 |
| I5 | L/C date | 신용장 발행일 | SO_해외 |
| I11 | HS CODE | 관세 코드 | — |
| I15 | PO Date (Customer) | 고객 PO 일자 | SO_해외.PO receipt date |

### 동적 필드 (Item List - Row 18~)

| 열 | 필드명 | 설명 |
|----|--------|------|
| A | Item name | 품목명 |
| E | Quantity | 수량 |
| G | Unit Price | 단가 (sales_unit_price) |
| I | Amount | 금액 (수량 × 단가) |

**2페이지 지원**: 아이템이 많으면 Row 53부터 2페이지 헤더 사용.

---

## CI — Commercial Invoice (`templates/commercial_invoice.xlsx`)

**생성기**: `ci_generator.py` (xlwings)
**데이터 소스**: DN_해외 + Customer_해외 + SO_해외 (Model number 보강)

> PI와 동일한 셀 구조이나, 데이터 소스가 DN_해외이며 아래 차이점이 있습니다.

### PI와의 차이점

| 항목 | PI | CI |
|------|----|----|
| 데이터 소스 | SO_해외 | DN_해외 |
| Invoice No (G4) | SO_ID | DN_ID |
| Invoice Date (I4) | 오늘 날짜 | dispatch_date (없으면 today) |
| ITEM_START_ROW | 18 | **20** (Row 19 = Electric Actuator 카테고리 라벨) |
| G5 | L/C No | **Incoterms** |
| I5 | L/C Date | **Payment Terms** |
| H열 (아이템 행) | — | 각 행에 currency 표시 |
| A열 품목명 | Item name | **Model number + Item name** |
| 아이템 정렬 | — | Model number 오름차순 |
| Total E열 | — | SUM(Qty) |
| Total F열 | — | "EA" |

### 고정 필드 (Header)

| 셀 | 필드명 | 설명 | 데이터 소스 |
|----|--------|------|-------------|
| A9 | Bill to 1 | 청구처 주소 1행 | Customer_해외.Bill to 1 |
| A10 | Bill to 2 | 청구처 주소 2행 | Customer_해외.Bill to 2 |
| A11 | Bill to 3 | 청구처 주소 3행 | Customer_해외.Bill to 3 |
| B14 | From | 출발지 | 고정값 "INCHEON, KOREA" |
| B15 | Destination | 도착 국가 | bill_to_3 (국가명) |
| G4 | Invoice No | DN_ID | DN_해외.DN_ID |
| G5 | Incoterms | 인코텀즈 | SO_해외.Incoterms (DN→SO JOIN) |
| G16 | PO No (Customer) | 고객 PO 번호 | DN_해외.Customer PO |
| I4 | Invoice date | 선적일 | DN_해외.dispatch_date |
| I5 | Payment Terms | 결제 조건 | Customer_해외.Payment terms |
| I12 | HS CODE | 관세 코드 | — |
| I16 | PO Date (Customer) | 고객 PO 일자 | DN_해외.PO receipt date |

### Shipping Mark 영역

| 셀 | 필드명 | 데이터 소스 |
|----|--------|-------------|
| A32 | Customer Name | Customer_해외.customer_name |
| A33 | Bill to 3 | Customer_해외.Bill to 3 |
| C34 | PO No | DN_해외.Customer PO |

### 동적 필드 (Item List - Row 20~)

| 열 | 필드명 | 설명 |
|----|--------|------|
| A | Item name | **Model number + Item name** (SO_해외에서 Model number 보강) |
| E | Quantity | 수량 |
| G | Unit Price | 단가 (unit_price → sales_unit_price fallback) |
| **H** | **Currency** | 통화 (각 행에 표시) |
| I | Amount | 금액 (수량 × 단가) |

### Total 행

| 열 | 내용 |
|----|------|
| E | SUM(Qty) |
| F | "EA" |
| H | Currency |
| I | SUM(Amount) |

### Model number / Model code 보강 로직
- `DocumentService._enrich_with_model_number()` — SO_해외에서 **SO_ID + Line item 복합키**로 Model number 및 Model code (AX Project number) 조회
- DN에 여러 SO_ID가 섞여 있어도 모든 아이템을 매칭
- 품목명: `"{Model number} {Item name}"` (Model number 없으면 Item name만)
- 아이템을 Model number 오름차순으로 정렬

### Weight 보강 로직 (PL 전용)
- `DocumentService._enrich_with_weight()` — Weight 시트의 ITEM→WEIGHT 매핑으로 Net Weight 자동 조회
- Model code (AX Project number) 값을 키로 사용하여 Weight 시트의 ITEM 컬럼과 매칭
- 매칭 결과를 `'Weight per unit'` 컬럼으로 items_df에 추가 → pl_generator가 F열에 자동 출력
- Weight 시트가 없거나 매칭 실패 시 graceful fallback (빈 값)


---

## PL — Packing List (`templates/packing_list.xlsx`)

**생성기**: `pl_generator.py` (xlwings)
**데이터 소스**: DN_해외 + Customer_해외 + SO_해외 (Model number/Model code 보강) + Weight 시트

> CI와 동일한 헤더 구조이나, 아이템 열이 다릅니다 (단가/금액 대신 Weight/CBM).

### 고정 필드 (Header)

| 셀 | 필드명 | 설명 | 데이터 소스 |
|----|--------|------|-------------|
| A9 | Bill to 1 | 청구처 주소 1행 | Customer_해외.Bill to 1 |
| A10 | Bill to 2 | 청구처 주소 2행 | Customer_해외.Bill to 2 |
| A11 | Bill to 3 | 청구처 주소 3행 | Customer_해외.Bill to 3 |
| G4 | Invoice No | DN_ID | DN_해외.DN_ID |
| G5 | Incoterms | 인코텀즈 | SO_해외.Incoterms (DN→SO JOIN) |
| I4 | Invoice date | 선적일 | DN_해외.dispatch_date |
| I12 | HS CODE | 관세 코드 | — |
| G16 | PO No (Customer) | 고객 PO 번호 | DN_해외.Customer PO |
| I16 | PO Date (Customer) | 고객 PO 일자 | DN_해외.PO receipt date |
| B14 | From | 출발지 | 고정값 "INCHEON, KOREA" |
| B15 | Destination | 도착 국가 | bill_to_3 (국가명) |

### Shipping Mark 영역

| 셀 | 필드명 | 데이터 소스 |
|----|--------|-------------|
| A34 | Customer Name | Customer_해외.customer_name |
| A35 | Bill to 3 | Customer_해외.Bill to 3 |
| C36 | PO No | DN_해외.Customer PO |

### 동적 필드 (Item List - Row 20~)

| 열 | 필드명 | 설명 |
|----|--------|------|
| A | Item name | **Model number + Item name** (CI와 동일 보강 로직) |
| E | Quantity | 수량 |
| F | Net Weight | KG/PC — **Weight 시트 기반 자동 조회** (Model code → ITEM 매핑) |
| H | Gross Weight | Kg (gross_weight) |
| I | CBM | Measurement (cbm) |

### Total 행

| 열 | 내용 |
|----|------|
| E | SUM(Qty) |
| G | "KGS" |
| H | SUM(Gross Weight) |
| I | SUM(CBM) |

---

## Final Invoice (`templates/final_invoice.xlsx`)

**생성기**: `fi_generator.py` (xlwings)
**데이터 소스**: DN_해외 + SO_해외 + Customer_해외 (3-way JOIN)

### 고정 필드 (Header)

| 셀 | 필드명 | 설명 | 데이터 소스 |
|----|--------|------|-------------|
| C7 | Customer PO | 고객 PO 번호 (C7:E7 병합) | SO_해외.Customer PO |
| H7 | Invoice No | 인보이스 번호 = DN_ID (H7:I7 병합) | DN_해외.DN_ID |
| C8 | PO Date | 고객 PO 일자 (C8:E8 병합) | SO_해외.PO receipt date |
| H8 | Invoice Date | 인보이스 발행일 = 선적일 (H8:I8 병합) | DN_해외.선적일 (dispatch_date) |
| H9 | Payment Terms | 결제 조건 (H9:I9 병합) | Customer_해외.Payment terms |
| H10 | Delivery Terms | 인코텀즈 (H10:I10 병합) | SO_해외.Incoterms (DN merge 시 SO 우선) |
| A12 | Customer Address 1 | 청구처 주소 1행 | Customer_해외.Bill to 1 |
| A13 | Customer Address 2 | 청구처 주소 2행 | Customer_해외.Bill to 2 |
| A14 | Customer Address 3 | 청구처 주소 3행 | Customer_해외.Bill to 3 |
| G12 | Delivery Address | 배송 주소 | DN_해외.Delivery Address |

### 동적 필드 (Item List - Row 17~)

헤더 행은 Row 16, 데이터는 Row 17부터.

| 열 | 필드명 | 설명 |
|----|--------|------|
| A (A:D 병합) | Item name | 품목명 |
| E | Quantity | 수량 |
| F | Unit Price | 단가 |
| G | Currency | 통화 (신규) |
| I | Amount | 금액 (수량 × 단가) |

### Total 행

| 열 | 내용 |
|----|------|
| E | SUM(Qty) |
| F | "EA" |
| G | Currency |
| I | SUM(Amount) |

---

## Order Confirmation (`templates/order_confirmation.xlsx`)

**생성기**: `oc_generator.py` (xlwings)
**데이터 소스**: SO_해외 + Customer_해외 (2-way JOIN)

> FI와 동일한 레이아웃에 H11(Shipping method), H열(Dispatch date)이 추가된 형태. SO_ID 기반.

### 고정 필드 (Header)

| 셀 | 필드명 | 설명 | 데이터 소스 |
|----|--------|------|-------------|
| C7 | Customer PO | 고객 PO 번호 (C7:E7 병합) | SO_해외.Customer PO |
| H7 | Invoice No | SO_ID (H7:I7 병합) | SO_해외.SO_ID |
| C8 | PO Date | 고객 PO 일자 (C8:E8 병합) | SO_해외.PO receipt date |
| H8 | Invoice Date | OC 발행일 (H8:I8 병합) | 오늘 날짜 |
| H9 | Payment Terms | 결제 조건 (H9:I9 병합) | Customer_해외.Payment terms |
| H10 | Delivery Terms | 인코텀즈 (H10:I10 병합) | SO_해외.Incoterms |
| H11 | Shipping method | 배송 방법 (H11:I11 병합) | SO_해외.Shipping method |
| A12 | Customer Address: | 라벨 (고정) | — |
| A13 | Customer Address 1 | 청구처 주소 1행 | Customer_해외.Bill to 1 |
| A14 | Customer Address 2 | 청구처 주소 2행 | Customer_해외.Bill to 2 |
| A15 | Customer Address 3 | 청구처 주소 3행 | Customer_해외.Bill to 3 |
| G12 | Delivery Address: | 라벨 (고정) | — |
| G13 | Delivery Address | 배송 주소 | SO_해외.납품 주소 |

### 동적 필드 (Item List - Row 18~)

헤더 행은 Row 17, 데이터는 Row 18부터.

| 열 | 필드명 | 설명 |
|----|--------|------|
| A (A:D 병합) | Description | **Model number + Item name** (Model number 없으면 Item name만) |
| E | Qty | 수량 |
| F | Unit Price | 단가 |
| G | Currency | 통화 |
| **H** | **Dispatch date** | **SO_해외.EXW NOAH** (OC 전용) |
| I | Amount | 금액 (수량 × 단가) |

### Total 행

| 열 | 내용 |
|----|------|
| E | SUM(Qty) |
| F | "EA" |
| G | Currency |
| I | SUM(Amount) |

---

## 구현 상태

| 문서 유형 | 상태 | 생성기 | 템플릿 |
|-----------|------|--------|--------|
| PO (Purchase Order) | ✅ 완료 | `excel_generator.py` (openpyxl) | `purchase_order.xlsx` |
| TS (거래명세표) | ✅ 완료 | `ts_generator.py` (xlwings) | `ts_template_local.xlsx` |
| PI (Proforma Invoice) | ✅ 완료 | `pi_generator.py` (xlwings) | `proforma_invoice.xlsx` |
| CI (Commercial Invoice) | ✅ 완료 | `ci_generator.py` (xlwings) | `commercial_invoice.xlsx` |
| FI (Final Invoice) | ✅ 완료 | `fi_generator.py` (xlwings) | `final_invoice.xlsx` |
| OC (Order Confirmation) | ✅ 완료 | `oc_generator.py` (xlwings) | `order_confirmation.xlsx` |
| PL (Packing List) | ✅ 완료 | `pl_generator.py` (xlwings) | `packing_list.xlsx` |

---

## 구현 시 참고사항

1. **라이브러리**: PO는 openpyxl, 나머지(TS/PI/CI/FI/OC/PL)는 xlwings 사용
2. **동적 행 처리**: xlwings `Range.insert/delete` 패턴 (각 generator 참고)
3. **데이터 조회**: `get_value(data, 'internal_key')` 표준 API 사용 (`COLUMN_ALIASES` 매핑)
4. **데이터 소스**: 국내는 PO/DN/PMT_국내, 해외는 SO/DN_해외 + Customer_해외
5. **Excel COM 최적화**: `excel_helpers.py`의 `batch_write_rows()`, `xlwings_app_context()` 활용
