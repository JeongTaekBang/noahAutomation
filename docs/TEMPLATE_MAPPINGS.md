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

## PI/CI — Proforma/Commercial Invoice (`templates/proforma_invoice.xlsx`, `commercial_invoice.xlsx`)

**생성기**: `pi_generator.py` (xlwings)
**데이터 소스**: SO_해외

> PI와 CI는 동일한 셀 구조를 사용합니다 (PI = 견적, CI = 확정).

### 고정 필드 (Header)

| 셀 | 필드명 | 설명 | 데이터 소스 |
|----|--------|------|-------------|
| A9 | Consigned to | 수취인 주소 (Customer address) | SO 또는 DN |
| A10 | Consigned country | 수취인 국가 | SO 또는 DN |
| C10 | Consigned TEL | 수취인 전화번호 | SO 또는 DN |
| E10 | Consigned FAX | 수취인 팩스번호 | SO 또는 DN |
| A12 | Vessel Name or Flight No | 선박명 또는 항공편 | DN |
| B13 | From | 출발지 (ex: INCHEON) | 고정값 또는 DN |
| B14 | Destination country | 도착 국가 (ex: U.S.A.) | SO 또는 DN |
| D15 | Departs on or about | 출발 예정일 | DN |
| G4 | Invoice No | 인보이스 번호 | 자동생성 또는 DN |
| G5 | L/C No | 신용장 번호 | SO 또는 DN |
| G15 | PO No (Customer) | 고객 PO 번호 | SO |
| G17 | Incoterms | 인코텀즈 | SO |
| I4 | Invoice date | 인보이스 발행일 | 생성일 |
| I5 | L/C date | 신용장 발행일 | SO 또는 DN |
| I11 | HS CODE | 관세 코드 | PO 또는 고정값 |
| I15 | PO Date (Customer) | 고객 PO 일자 | SO |

### 동적 필드 (Item List - Row 18~)

| 열 | 필드명 | 설명 |
|----|--------|------|
| A | Item name | 품목명 |
| E | Quantity | 수량 |
| G | Unit Price | 단가 |
| I | Amount | 금액 (수량 × 단가) |

**2페이지 지원**: 아이템이 많으면 Row 53부터 2페이지 헤더 사용.

---

## Final Invoice (`templates/final_invoice.xlsx`)

**생성기**: `fi_generator.py` (xlwings)
**데이터 소스**: DN_해외 + Customer_해외 (JOIN)

### 고정 필드 (Header)

| 셀 | 필드명 | 설명 | 데이터 소스 |
|----|--------|------|-------------|
| G4 | Invoice No | 인보이스 번호 | 자동생성 |
| I4 | Invoice Date | 인보이스 발행일 | 출고일 (dispatch_date) |
| A9 | Bill to (1줄) | 청구처 주소 1행 | bill_to_1 |
| A10 | Bill to (2줄) | 청구처 주소 2행 | bill_to_2 |
| A11 | Bill to (3줄) | 청구처 주소 3행 | bill_to_3 |
| G8 | Payment Terms | 결제 조건 | payment_terms (G8:G9 병합) |
| I8 | Due Date | 결제 기한 | due_date (I8:I9 병합) |
| G10 | PO No | 고객 PO 번호 | customer_po (G10:G11 병합) |
| I10 | PO Date | 고객 PO 일자 | po_date (I10:I11 병합) |

### 동적 필드 (Item List - Row 14~)

헤더 행은 Row 13, 데이터는 Row 14부터.

| 열 | 필드명 | 설명 |
|----|--------|------|
| A | Item name | 품목명 |
| E | Quantity | 수량 |
| G | Unit Price | 단가 |
| I | Amount | 금액 (수량 × 단가) |

---

## 구현 상태

| 문서 유형 | 상태 | 생성기 | 템플릿 |
|-----------|------|--------|--------|
| PO (Purchase Order) | ✅ 완료 | `excel_generator.py` (openpyxl) | `purchase_order.xlsx` |
| TS (거래명세표) | ✅ 완료 | `ts_generator.py` (xlwings) | `ts_template_local.xlsx` |
| PI (Proforma Invoice) | ✅ 완료 | `pi_generator.py` (xlwings) | `proforma_invoice.xlsx` |
| CI (Commercial Invoice) | ✅ 완료 | `pi_generator.py` (xlwings) | `commercial_invoice.xlsx` |
| FI (Final Invoice) | ✅ 완료 | `fi_generator.py` (xlwings) | `final_invoice.xlsx` |
| Packing List | ❌ 미구현 | — | `packing_list.xlsx` |

---

## 구현 시 참고사항

1. **라이브러리**: PO는 openpyxl, 나머지(TS/PI/CI/FI)는 xlwings 사용
2. **동적 행 처리**: xlwings `Range.insert/delete` 패턴 (각 generator 참고)
3. **데이터 조회**: `get_value(data, 'internal_key')` 표준 API 사용 (`COLUMN_ALIASES` 매핑)
4. **데이터 소스**: 국내는 PO/DN/PMT_국내, 해외는 SO/DN_해외 + Customer_해외
5. **Excel COM 최적화**: `excel_helpers.py`의 `batch_write_rows()`, `xlwings_app_context()` 활용
