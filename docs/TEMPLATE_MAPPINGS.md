# Template Cell Mappings

향후 문서 생성 구현 시 참고용 템플릿 셀 매핑 정보입니다.

---

## Commercial Invoice (`templates/commercial_invoice.xlsx`)

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
| I4 | Invoice date | 인보이스 발행일 | 생성일 |
| I5 | L/C date | 신용장 발행일 | SO 또는 DN |
| I11 | HS CODE | 관세 코드 | PO 또는 고정값 |
| G15 | PO No (Customer) | 고객 PO 번호 | SO |
| I15 | PO Date (Customer) | 고객 PO 일자 | SO |

### 동적 필드 (Item List - Row 18~)

| 셀 | 필드명 | 설명 |
|----|--------|------|
| A18~ | Item name | 품목명 (여러 행 가능) |
| E18~ | Quantity | 수량 |
| G18~ | Unit Price | 단가 |

**동적 행 처리 필요**: 아이템 개수에 따라 행 복제 (`pi_generator.py`의 xlwings 패턴 참고)

---

## 구현 상태

### Proforma Invoice
- [x] 구현 완료 (`pi_generator.py`)
- 셀 매핑: `templates/proforma_invoice.xlsx` 템플릿 참조

### Packing List
- [ ] 셀 매핑 정보 추가 예정

---

## 구현 시 참고사항

1. **라이브러리**: xlwings 사용 (모든 문서 생성에 통일됨)
2. **동적 행 처리**: `pi_generator.py`의 행 복제 패턴 참고 (xlwings Range.insert/delete)
3. **데이터 조회**: `get_value(data, 'internal_key')` 표준 API 사용 (COLUMN_ALIASES 매핑)
4. **데이터 소스**: SO_해외, DN_해외 시트에서 SO_ID로 조회
