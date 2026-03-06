# Changelog

개발 이력, 버그 수정, 리팩토링 기록.

---

## TODO (미완료 항목)

### 템플릿 확장
- [x] Proforma Invoice (PI) 구현 완료 ✓
- [x] Final Invoice (FI) 구현 완료 ✓
- [x] Order Confirmation (OC) 구현 완료 ✓
- [x] Commercial Invoice (CI) 구현 완료 ✓
- [x] Packing List (PL) 구현 완료 ✓

### SQL 기반 데이터 분석
- [x] NOAH_SO_PO_DN.xlsx → SQLite DB 동기화 구현 완료 ✓
  - **배경**: Excel 형식의 데이터 유실/변형 취약점 → SQLite 백업
  - DuckDB 분석 연동은 추후 확장 예정

---

## 2026-03-06: PO 테이블 PK에 `_row_seq` 추가 (부분 매입 대응)

- `db_schema.py`: PO_국내/PO_해외의 PK를 `(PO_ID, Line item)` → `(PO_ID, Line item, _row_seq)`로 변경
- `_row_seq`는 같은 `(PO_ID, Line item)` 그룹 내에서 Excel 행 순서대로 자동 부여 (1, 2, 3...)
- 부분 매입 시 같은 Line item이 분할되어도 PK 충돌 없이 정상 동기화
- `db_schema.py`: `migrate_pk_if_changed()` 추가 — 기존 DB의 PK가 설정과 다르면 자동 DROP → 재생성
- `db_sync.py`: 테이블 생성 전 PK 마이그레이션 체크 호출

---

## 2026-03-06: OC 품목명에 Model number 표시

- `oc_generator.py`: 품목명 출력 시 SO_해외의 Model number가 있으면 `"{Model number} {Item name}"` 형태로 표시
- CI와 동일한 로직 적용 (Model number 없으면 Item name만 출력)

---

## 2026-03-06: 내부 코드 최적화 (데이터 조회/서비스 캐시)

### 배경
데이터 조회 병목 분석 후, 출력 결과에 영향 없는 내부 구현 최적화 수행. 기능 회귀 없음 확인 (58 passed, 0 failed).

### 변경 내용

#### 1. `resolve_column()` 캐시 추가 (`utils.py`)
- `id(columns)` + `key` 기반 dict 캐시 도입
- `get_value()` 매 호출마다 반복되던 별칭 검색을 O(1) 조회로 전환
- 문서 1건 생성 시 수십~백 회 불필요한 선형 검색 제거

#### 2. `get_available_*_ids()` O(n²) → O(n) (`finder_service.py`)
- 4개 메서드(`get_available_po_ids`, `get_available_dn_ids`, `get_available_dn_export_ids`, `get_available_so_export_ids`)
- 기존: `unique()` 루프 안에서 `df[df[col] == id]` 반복 필터 → O(n²)
- 변경: `drop_duplicates(subset=..., keep='first').head(limit)` 단일 패스 → O(n)

#### 3. `find_so_for_advance()` 캐시 재사용 (`finder_service.py`)
- 기존: `load_so_for_advance()`가 Excel 파일을 독립적으로 다시 오픈 (PMT+SO 2시트 재로드)
- 변경: `FinderService`의 캐시된 `_pmt_df`와 신규 `_so_domestic_df` 활용, 중복 Excel I/O 제거
- `_load_so_domestic()` 프라이빗 메서드 추가 (SO_국내 lazy cache)

#### 4. `create_po.py` 다건 처리 서비스 공유
- `generate_po()`에 `service` 파라미터 추가 (기본값 `None` → 하위 호환)
- `main()`에서 여러 주문번호 처리 시 단일 `DocumentService` 인스턴스 공유
- DataFrame 재로드 방지

### 검토 후 제외된 항목
| 제안 | 제외 사유 |
|------|----------|
| `iterrows()` → `itertuples()` | 아이템 1~50건 수준이라 마이크로초 차이. 한글 컬럼명이 namedtuple 필드로 변환 실패 → 코드 복잡도만 증가 |
| xlwings COM 호출 추가 축소 | 이미 `batch_write_rows()` 등으로 96~97% 감소 완료. 남은 row insertion은 Excel API 제약으로 배치화 불가 |
| `get_value()` 배치 API | `resolve_column()` 캐시만으로 병목 해소. 별도 API는 blast radius가 큼 |

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `po_generator/utils.py` | `_RESOLVE_SENTINEL`, `_resolve_cache` 추가, `resolve_column()` 캐시 적용 |
| `po_generator/services/finder_service.py` | `_so_domestic_df` 캐시, `_load_so_domestic()` 추가, `get_available_*_ids()` 4개 single-pass 교체, `find_so_for_advance()` 캐시 재사용, `load_so_for_advance` import 제거 |
| `create_po.py` | `generate_po()` 시그니처에 `service` 파라미터 추가, `main()` 서비스 공유 |

---

## 2026-03-06: Commercial Invoice (CI) & Packing List (PL) 생성기 추가

### 배경
해외 출하 시 필요한 Commercial Invoice와 Packing List 생성 기능 추가. 둘 다 DN_해외 데이터를 사용하며, PI/FI와 유사한 셀 레이아웃.

### CI (Commercial Invoice)
PI와 동일한 셀 구조이나, 데이터 소스가 DN_해외이며 아래 차이점 있음:
- `ITEM_START_ROW = 19` (Row 18 = 카테고리 라벨 유지)
- `CELL_INCOTERMS = G18` (PI는 G17)
- A9 = Delivery Address, Shipping Mark (A31=Customer Name, C33=Customer PO)
- H열에 각 행 currency 표시, Total에 Qty 합계(E) + "EA"(F)
- **Model number 보강**: SO_해외에서 Item name 매칭으로 Model number 조회, 품목명 앞에 추가
- **Model number 오름차순 정렬**

### PL (Packing List)
CI와 동일한 헤더 구조이나, 아이템 열이 다름 (단가/금액 대신 Weight/CBM):
- F열: Net Weight (KG/PC), H열: Gross Weight (Kg), I열: CBM
- Shipping Mark: A31=Customer Name, A32=Customer Country, C33=Customer PO
- Model number 보강 및 정렬: CI와 동일

### 신규 파일
| 파일 | 역할 |
|------|------|
| `create_ci.py` | CI CLI 진입점 (`python create_ci.py DNO-2026-0001`) |
| `create_pl.py` | PL CLI 진입점 (`python create_pl.py DNO-2026-0001`) |
| `po_generator/ci_generator.py` | CI 생성기 (xlwings, PI 기반) |
| `po_generator/pl_generator.py` | PL 생성기 (xlwings, CI 기반 + Weight/CBM) |

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `config.py` | `CI_TEMPLATE_FILE`, `CI_OUTPUT_DIR`, `PL_TEMPLATE_FILE`, `PL_OUTPUT_DIR`, weight/cbm 컬럼 별칭 추가 |
| `utils.py` | `load_dn_export_data()`, `load_so_export_with_customer()` — Customer_해외 merge 시 `drop_duplicates()` 추가 (중복 행 방지) |
| `services/document_service.py` | `_enrich_with_model_number()`, `generate_ci()`, `generate_pl()` 메서드 추가 |
| `create_po.bat` | 메뉴에 [6] CI, [7] PL 추가 (기존 DB동기화 [8], 이력조회 [9]) |
| `docs/TEMPLATE_MAPPINGS.md` | CI/PL 셀 매핑 섹션 추가, PI 섹션 분리 |

### 데이터 흐름
```
DN_해외 → Customer_해외 (customer_code JOIN)
       → SO_해외 (SO_ID + Item name → Model number 보강)
```

### Customer_해외 중복 행 수정
`load_dn_export_data()`와 `load_so_export_with_customer()`에서 Customer_해외 merge 시 `drop_duplicates(subset='C-code by 해외', keep='first')` 추가. Customer_해외에 동일 고객코드 중복 행이 있을 때 DN 행이 배수로 늘어나는 버그 수정.

---

## 2026-03-05: Order Confirmation (OC) 생성기 추가

### 배경
해외 고객에게 주문 확인서(Order Confirmation)를 발행하는 기능 추가. Final Invoice와 동일한 레이아웃이지만, H열에 **Dispatch date** 컬럼이 추가된 형태. Dispatch date는 SO_해외의 `EXW NOAH` 컬럼 값을 사용.

### 신규 파일
| 파일 | 역할 |
|------|------|
| `create_oc.py` | CLI 진입점 (`python create_oc.py SOO-2026-0001`) |
| `po_generator/oc_generator.py` | OC 생성기 (xlwings, FI 기반 + Dispatch date) |

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `config.py` | `OC_TEMPLATE_FILE`, `OC_OUTPUT_DIR` 추가, `exw_noah` 컬럼 별칭 추가 |
| `utils.py` | `load_so_export_with_customer()` 신규 (SO_해외+Customer_해외 JOIN) |
| `services/finder_service.py` | `find_so_export_with_customer()` 메서드 추가 |
| `services/document_service.py` | `generate_oc(so_id)` 메서드 추가 |
| `create_po.bat` | 메뉴에 [5] Order Confirmation 추가 (기존 DB동기화 [5]→[6]으로 이동) |

### OC vs FI 차이점
| 항목 | FI | OC |
|------|----|----|
| 제목 | Invoice | Confirmation of Order |
| H열 (Row 17~) | (없음) | Dispatch date = SO_해외.EXW NOAH |
| 나머지 | 동일 | 동일 |

### 데이터 흐름
`SO_해외` → `Customer_해외` (고객코드 JOIN, Bill to/Payment terms 포함)

---

## 2026-03-04: FI 새 템플릿 대응 업데이트

### 배경
`templates/final_invoice.xlsx` 양식 전면 개편으로 셀 매핑, 아이템 열 구조, 신규 필드 대응 필요.

### 변경 파일
| 파일 | 변경 내용 |
|------|----------|
| `config.py` | `dispatch_date` alias에 `'선적일'` 추가, `delivery_address` alias에 `'Delivery address'` 추가 |
| `utils.py` | `load_dn_export_data()` — SO_해외 merge 컬럼에 `Currency`, `Incoterms` 추가, Customer_해외 JOIN 키를 `resolve_column()`으로 동적 탐지, DN-SO 컬럼 충돌 시 SO 우선 (overlap drop) |
| `fi_generator.py` | 셀 매핑 전면 교체, `_fill_header()` 재작성, `_fill_items_batch()` Currency 열 추가, `_update_total_row()` F열 "EA" 추가 |
| `docs/TEMPLATE_MAPPINGS.md` | FI 섹션 새 템플릿 구조로 업데이트 |

### 셀 매핑 변경 요약
| 필드 | OLD → NEW | 데이터 소스 |
|------|-----------|------------|
| Customer PO | G10 → C7 | SO_해외.Customer PO |
| Invoice No | G4 → H7 | DN_해외.DN_ID |
| PO Date | I10 → C8 | SO_해외.PO receipt date |
| Invoice Date | I4 → H8 | DN_해외.선적일 |
| Payment Terms | G8 → H9 | Customer_해외.Payment terms |
| Delivery Terms | (신규) H10 | SO_해외.Incoterms |
| Customer Address | A9~11 → A12~14 | Customer_해외.Bill to 1/2/3 |
| Delivery Address | (신규) G12 | DN_해외.Delivery address |
| Due Date | I8 → (삭제) | — |

### 아이템 열 변경
| 항목 | OLD → NEW |
|------|-----------|
| ITEM_START_ROW | 14 → 17 |
| Unit Price 열 | G → F |
| Currency 열 | (신규) G |

### 데이터 로드 개선 (`load_dn_export_data`)
- **Customer_해외 JOIN 키**: 하드코딩(`'Business registration number'`) → `resolve_column()`으로 동적 탐지
- **DN-SO 컬럼 충돌**: SO_해외에서 가져올 컬럼이 DN_해외에도 존재하면 merge 전 DN쪽 drop (SO 우선)
- **alias 대소문자**: `'Delivery address'`(소문자 a) 추가 — DN_해외 실제 컬럼명과 매칭

---

## 2026-03-03: Excel → SQLite DB 동기화 구현

### 배경
NOAH_SO_PO_DN.xlsx가 사실상 ERP 역할을 하고 있으나, Excel 형식 특성상 데이터 유실/변형에 취약. 수동 입력 시트(SO, PO, DN, PMT)를 SQLite DB에 upsert 방식으로 업로드하여 데이터를 안전하게 백업하고 관리.

### 신규 파일
| 파일 | 역할 |
|------|------|
| `sync_db.py` | CLI 진입점 (--dry-run, --sheets, --info, -v) |
| `po_generator/db_schema.py` | 테이블/PK 정의, DDL 생성, 스키마 관리 |
| `po_generator/db_sync.py` | SyncEngine — upsert 동기화 엔진 |

### 수정 파일
| 파일 | 변경 |
|------|------|
| `po_generator/config.py` | `DB_FILE` 상수 1줄 추가 |

### 테이블 설계 (7개)

| 테이블명 | 소스 시트 | PK | 행 수 |
|----------|----------|-----|------|
| `so_domestic` | SO_국내 | `(SO_ID, Customer PO, Line item)` | 590 |
| `so_export` | SO_해외 | `(SO_ID, Customer PO, Line item)` | 233 |
| `po_domestic` | PO_국내 | `(SO_ID, Customer PO, Line item)` | 589 |
| `po_export` | PO_해외 | `(SO_ID, Customer PO, Line item, _row_seq)` | 235 |
| `dn_domestic` | DN_국내 | `(DN_ID, Line item)` | 283 |
| `dn_export` | DN_해외 | `(DN_ID, SO_ID, Line item)` | 88 |
| `pmt_domestic` | PMT_국내 | `(선수금_ID)` | 33 |

### 사용법
```bash
python sync_db.py                           # 전체 동기화
python sync_db.py -v                        # 상세 로그
python sync_db.py --sheets SO_국내 PO_국내  # 특정 시트만
python sync_db.py --dry-run                 # 시뮬레이션
python sync_db.py --info                    # DB 현황 조회
```

### 핵심 설계
- **DB**: SQLite (Python 내장, 서버 불필요). 위치: `DATA_DIR / "noah_data.db"`
- **Upsert**: PK 기준 INSERT or UPDATE — 재실행 시 기존 데이터 업데이트
- **PO_해외 `_row_seq`**: 같은 SO Line item에 사양 변형 시 자동 순번 부여
- **스키마 진화**: `ensure_columns_exist()`로 Excel 컬럼 추가 시 자동 대응
- **메타 테이블**: `_sync_meta`에 테이블별 마지막 동기화 시간/행 수 기록

---

## 2026-02-28: Power Query 개선 및 문서 정비

### Power Query 수정
- PO 원가 계산: `Table.Distinct` → `Table.Group` 변경 (사양 분리 시 중복 합산 방지)
- DN 분할납품: `Table.Distinct` → `Table.Group` 변경 (분할 출고 금액 정확 집계)
- SO_통합 출고 상태: 3단계 → 4단계 세분화 (미출고/부분 출고/출고 완료/선적 완료)
- PO_AX대사 쿼리 추가: Period + AX PO별 GRN 금액 집계

### 문서 정비
- `DATA_STRUCTURE_DESIGN.md`: ERP 매핑 섹션 추가 (테이블 관계, 조인, 상태 관리)
- `CLAUDE.md`: 아키텍처, 커맨드, 키 패턴 섹션 확장
- `POWER_QUERY.md`: Key Files에 추가

---

## 2026-02-15: Final Invoice 및 Power Query 문서화

### Final Invoice (FI) 생성기 추가
- `create_fi.py` CLI 진입점 추가 (DN_해외 기반)
- `fi_generator.py` 구현 (xlwings) — Bill-to, Payment Terms, Due Date 등
- `create_po.bat` 메뉴에 [4] Final Invoice 추가
- `OPERATION_GUIDE.md` 운용 가이드 추가
- `config.py`: `FI_TEMPLATE_FILE`, `FI_OUTPUT_DIR` 추가

### Power Query 문서화
- `docs/POWER_QUERY.md` 신규 작성 (SO_통합, PO_현황, Order_Book 쿼리)
- Order_Book 파이프라인 다이어그램 및 단계별 데이터 흐름 예시

### 데이터 구조
- SO 컬럼: `Customer PO`, `Expected delivery date` 추가
- Order_Book: 분할 납품 처리 (DN 월별 조인)

---

## 2026-02-08: TS/PI 기능 및 테스트 추가

### 거래명세표/PI 기능 확장
- TS/PI 관련 기능 정리 및 테스트 추가
- `.gitignore` 업데이트 (generated_ts, po_history, Claude 임시 파일)
- README 갱신 (PO, TS, PI 문서 유형 반영)

---

## 2026-01-31: 거래명세표 기능 개선

### 출고일 기준 날짜 표시
- **변경**: 거래명세표 날짜를 오늘 날짜 → **출고일** 기준으로 변경
- `config.py`: `dispatch_date` 별칭 추가 (`'출고일', 'Dispatch Date', 'dispatch_date', '출하일'`)
- `ts_generator.py`: 헤더(B2)와 아이템(A열) 날짜를 출고일로 표시
  - 출고일이 없으면 오늘 날짜를 폴백으로 사용
  - 파라미터명 `today` → `dispatch_date`로 변경

### 월합 거래명세표 기능 추가
고객이 월합으로 거래명세표를 요청할 때, 여러 DN을 한 장으로 합쳐서 생성

**사용법:**
```bash
python create_ts.py DND-2026-0001 DND-2026-0002 DND-2026-0003 --merge
python create_ts.py --interactive --merge
```

**변경 파일:**
| 파일 | 변경 내용 |
|------|----------|
| `config.py` | `dispatch_date` 컬럼 별칭 추가 |
| `ts_generator.py` | 출고일 기준 날짜 표시 |
| `create_ts.py` | `--merge`, `--interactive` 옵션 추가, `generate_merged_ts()` 함수 |
| `create_po.bat` | 거래명세표 메뉴에 [1] 단건 / [2] 월합 선택 추가 |

**월합 거래명세표 동작:**
- 여러 DN의 아이템을 하나의 DataFrame으로 합침
- 출고일: 입력된 DN 중 **가장 최근 출고일** 사용
- 고객명이 다르면 경고 표시 (첫 번째 고객 기준)
- 파일명: `월합_고객명_날짜.xlsx`

---

## 2026-01-21: 코드 리팩토링 (5 Phases)

Code Reflection 결과를 바탕으로 코드 품질 개선 작업 수행.

### Phase 1: excel_helpers.py 인프라 추가
- [x] `XlConstants` 클래스 추가 ✓
  - Excel COM 매직 넘버를 명명된 상수로 정의
  - `xlShiftUp`, `xlShiftDown`, `xlEdgeTop`, `xlEdgeBottom`, `xlContinuous`, `xlThin` 등
  - 코드 가독성 향상, 하드코딩된 -4162, -4121 등 제거
- [x] `xlwings_app_context` 컨텍스트 매니저 추가 ✓
  - xlwings App 생명주기 안전 관리
  - 오류 발생 시에도 Excel 프로세스 자동 정리
  - 리소스 누수 방지
- [x] `prepare_template()`, `cleanup_temp_file()` 헬퍼 추가 ✓
  - 중복되는 템플릿 복사 로직 통합
  - 임시 파일 안전 삭제

### Phase 2: cli_common.py 보안 수정
- [x] Path Traversal 취약점 수정 ✓
  - **문제**: 문자열 포함 검사(`in`)로 경로 탈출 가능
    - `/home/user/documents`가 `/home/user/doc_files/test.xlsx`에 포함
  - **수정**: `relative_to()` 사용으로 정확한 경로 검증
  ```python
  # Before (취약)
  if str(output_dir.resolve()) not in str(output_file.resolve()):

  # After (안전)
  resolved_file.relative_to(resolved_dir)  # ValueError 발생 시 거부
  ```

### Phase 3: Generator 리팩토링
- [x] `excel_generator.py` 리팩토링 ✓ - `xlwings_app_context`, `XlConstants`, 타입 변환 경고 로깅
- [x] `ts_generator.py` 리팩토링 ✓ - 동일 패턴 적용
- [x] `pi_generator.py` 리팩토링 ✓ - 동일 패턴 적용

### Phase 4: 테스트 커버리지 확대
- [x] `tests/test_excel_helpers.py` 신규 생성 ✓ (16개 테스트)
- [x] `tests/test_cli_common.py` 신규 생성 ✓ (11개 테스트)
- [x] `tests/test_config.py` 신규 생성 ✓ (22개 테스트)
- **테스트 결과**: 47 passed, 2 skipped

### Phase 5: utils.py 중복 함수 통합
- [x] `_find_data_by_id()` 공통 헬퍼 추가 ✓
  - ID로 데이터 검색하는 공통 로직 통합
- [x] 4개 find 함수를 wrapper로 변경 ✓
  | 함수 | 변경 전 | 변경 후 |
  |------|--------|--------|
  | `find_order_data()` | 34줄 | 1줄 (wrapper) |
  | `find_dn_data()` | 33줄 | 1줄 (wrapper) |
  | `find_pmt_data()` | 28줄 | 1줄 (wrapper) |
  | `find_so_export_data()` | 33줄 | 1줄 (wrapper) |
- **효과**: ~90줄 중복 제거, 버그 수정 시 단일 지점 수정
- **테스트 결과**: 160 passed, 2 skipped

**변경 파일 요약:**
| 파일 | 변경 내용 |
|------|----------|
| `excel_helpers.py` | +110 lines (XlConstants, context manager, helpers) |
| `cli_common.py` | 보안 버그 수정 |
| `excel_generator.py` | Context manager 적용, 상수화 |
| `ts_generator.py` | Context manager 적용, 상수화 |
| `pi_generator.py` | Context manager 적용, 상수화 |
| `test_excel_helpers.py` | +160 lines (신규) |
| `test_cli_common.py` | +90 lines (신규) |
| `test_config.py` | +140 lines (신규) |

---

## 2026-01-21: xlwings 성능 최적화

### 배치 연산 헬퍼 함수 추가
`excel_helpers.py`에 새 함수:
- `batch_write_rows`: 2D 리스트를 한 번에 쓰기
- `batch_read_column`: 열의 값을 한 번에 읽기
- `batch_read_range`: 범위의 값을 한 번에 읽기
- `delete_rows_range`: 연속 행을 한 번에 삭제
- `find_text_in_column_batch`: 배치 읽기로 텍스트 찾기

### Generator별 최적화
- `ts_generator.py`: `_fill_items_batch` (N*8회→1회), `_find_label_row` (36회→1회), `_find_ts_subtotal_row` (15회→1회)
- `pi_generator.py`: `_fill_items_batch` (N*4회→4회), `_find_total_row` (20회→1회), `_fill_shipping_mark` (80회→2회)
- `excel_generator.py`: `_fill_items_batch_po`, `_create_description_sheet` (30*N회→1회), `_find_totals_row` (20회→1회)

### 예상 성능 개선 (50개 아이템 기준)
| 파일 | COM 호출 (전) | COM 호출 (후) | 감소율 |
|------|--------------|--------------|--------|
| ts_generator.py | ~500회 | ~20회 | 96% |
| pi_generator.py | ~350회 | ~15회 | 96% |
| excel_generator.py | ~1,500회 | ~50회 | 97% |

---

## 2026-01-21: 버그 수정 - xlwings 범위 formula 읽기

**증상**: 거래명세표 생성 시 템플릿의 예시 아이템이 삭제되지 않고 그대로 남아있음

**원인**: `_find_ts_subtotal_row` 함수의 배치 읽기 최적화에서 xlwings의 `.formula` 속성 반환 형식을 잘못 처리

**상세 분석:**
```python
# xlwings 범위 읽기 반환 형식 차이
ws.range('E13').value           # 단일 셀 → float: 8.0
ws.range('E13:E17').value       # 범위 → list: [8.0, 8.0, 16.0, None, None]

ws.range('E15').formula         # 단일 셀 → str: '=SUM(E13:E14)'
ws.range('E13:E17').formula     # 범위 → tuple of tuples: (('8',), ('8',), ('=SUM(E13:E14)',), ('',), ('',))
```

- `.value`: 단일 열 범위 → **1D list** 반환
- `.formula`: 단일 열 범위 → **tuple of tuples** 반환 (2D 형태)

**버그 코드:**
```python
formulas = ws.range(f'E{start_row}:E{end_row}').formula
if not isinstance(formulas, list):
    formulas = [formulas]  # tuple of tuples가 통째로 리스트에 들어감
for idx, formula in enumerate(formulas):
    if formula and '=SUM' in str(formula):  # 전체 tuple을 문자열로 변환
        return start_row + idx  # 항상 index 0 반환
```

결과: `subtotal_row = 13` (실제로는 15) → `template_item_count = 0` → 행 삭제 안됨

**수정 코드** (`ts_generator.py:186-197`):
```python
formulas = ws.range(f'E{start_row}:E{end_row}').formula

# xlwings 범위 읽기는 tuple of tuples 반환: (('val1',), ('val2',), ...)
# 단일 셀은 문자열 반환
if isinstance(formulas, (list, tuple)) and formulas and isinstance(formulas[0], (list, tuple)):
    # 2D → 1D 평탄화 (각 행의 첫 번째 값만 추출)
    formulas = [f[0] if f else '' for f in formulas]
elif not isinstance(formulas, (list, tuple)):
    formulas = [formulas]
```

**영향 범위:**
| 모듈 | 함수 | 사용 속성 | 상태 |
|------|------|----------|------|
| `ts_generator.py` | `_find_ts_subtotal_row` | `.formula` (범위) | **수정됨** |
| `pi_generator.py` | `_find_total_row` | `.value` (범위) | 문제 없음 |
| `excel_generator.py` | `_find_totals_row` | `.value` (범위) | 문제 없음 |
| `excel_helpers.py` | `batch_read_column` | `.value` (범위) | 문제 없음 |

**교훈:**
- xlwings에서 `.value`와 `.formula`는 범위 읽기 시 반환 형식이 다름
- `.value`: 1D list (단일 열)
- `.formula`: 2D tuple of tuples (항상 2D)
- 배치 최적화 시 반환 형식을 실제 테스트로 확인 필요

---

## 2026-01-21: 서비스 레이어 추가

- [x] `excel_helpers.py` 생성 ✓ - `find_item_start_row` 통합, 헤더 라벨 프리셋
- [x] `services/` 디렉토리 생성 ✓ - DocumentService, FinderService, DocumentResult
- [x] CLI 리팩토링 ✓ - 서비스 레이어 사용, 사용자 상호작용은 CLI 유지
- [x] 행 삭제 주석 수정 ✓ - "같은 위치에서 반복 삭제 - xlUp으로 아래 행이 올라옴"
- [x] 통합 테스트 추가 ✓ - 11개 테스트 케이스

---

## 2026-01-21: 버그 수정 - Description 시트 A열 레이블

- **원인**: 템플릿의 고정 레이블에만 의존, 동적 필드(`get_spec_option_fields`)와 불일치
- **수정**: A열에 레이블 명시적 쓰기 (`['Line No', 'Qty'] + all_fields`)
- `_apply_description_borders` 함수 추가 (테두리 적용)
- 국내/해외 모두 동적 필드 사용 (PO_국내: 47개, PO_해외: 45개)

---

## 2026-01-21: 버그 수정 - PI 행 삽입 시 테두리

- **증상**: 템플릿 마지막 행(8행) 테두리가 중간에 남음 + Total 위 선 누락
- **원인**: 행 삽입 케이스에서 `_restore_item_borders` 미호출
- **수정**: 삽입 전 템플릿 원래 마지막 행 테두리 제거 (`XlConstants.xlNone`)
- **수정**: 삽입 후 `_restore_item_borders` 호출로 새 마지막 행 테두리 추가
- `excel_helpers.py`에 `XlConstants.xlNone = -4142` 상수 추가

---

## 2026-01-20: openpyxl → xlwings 전환

- openpyxl → xlwings 전환 (이미지/서식 보존)
- `get_safe_value` → `get_value` 표준 API로 통일

---

## 2026-01-19: 버그 수정 - PO Delivery Address

- Delivery Address 값이 안 나오던 문제 해결
  - `config.py`: `delivery_address` 컬럼 별칭 추가
  - `utils.py`: SO→PO 병합 시 `'납품 주소'` 컬럼 누락 수정
  - `excel_generator.py`: 하드코딩 키워드 검색 → `get_value()` 사용
- 파일 열 때 Description 시트가 먼저 보이던 문제 해결
  - `excel_generator.py`: `wb.active = ws_po` 추가

---

## 2026-01-19: 버그 수정 - 거래명세표/PI 템플릿 예시 아이템 삭제

- 실제 아이템 < 템플릿 예시 시 초과 행 삭제 안되던 문제 해결
- 행 삭제 후 테두리 복원 (`_restore_ts_item_borders`, `_restore_item_borders`)
- PI: Shipping Mark 영역 검색 범위 수정 (40→20 시작)

---

## 완료된 TODO 항목

### OneDrive 공유 폴더 연동
- [x] 회사 랩탑에서 OneDrive 공유 폴더 경로 확인 ✓
- [x] `config.py`에서 경로 설정 외부화 → `user_settings.py` ✓
- [x] 파일 구조 변경 ✓
- [x] po_history 월별 폴더 방식으로 변경 ✓

### 템플릿 기반 문서 생성
- [x] PO (Purchase Order) - openpyxl 기반 ✓
- [x] 거래명세표 (Transaction Statement) - xlwings 기반 ✓
- [x] PI (Proforma Invoice) - xlwings 기반 ✓
- [x] FI (Final Invoice) - xlwings 기반 ✓
- [x] OC (Order Confirmation) - xlwings 기반 ✓
- [x] CI (Commercial Invoice) - xlwings 기반 ✓
- [x] PL (Packing List) - xlwings 기반 ✓

---

## 라이브러리 선택 기준

| 용도 | 라이브러리 | 이유 |
|------|-----------|------|
| PO 생성 | openpyxl | 이미지 불필요, 빠른 생성 |
| TS/PI/FI/OC 생성 | xlwings | 로고/도장 이미지, 복잡한 서식 완벽 보존 |
| 이력 조회/테스트 검증 | openpyxl | COM 인터페이스 없이 안정적인 읽기 |

### 템플릿 동작 방식
- 템플릿 파일의 **데이터는 무시됨** - 코드에서 초기화 후 새로 채움
- 템플릿의 **구조/서식만 유지됨**: 레이아웃, 서식, 이미지, 수식
- 새 템플릿 추가 시: `templates/` 폴더에 양식 파일 추가 후 코드에서 셀 매핑 정의
