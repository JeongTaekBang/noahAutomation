# Changelog

개발 이력, 버그 수정, 리팩토링 기록.

---

## TODO (미완료 항목)

### 템플릿 확장 (해외 오더)
- [ ] Commercial Invoice 템플릿 (xlwings 기반)
- [ ] Packing List 템플릿 (xlwings 기반)

### SQL 기반 데이터 분석
- [ ] NOAH_SO_PO_DN.xlsx → 로컬 SQL DB 연동
  - **배경**: Power Pivot의 양방향 JOIN 제한으로 복잡한 분석 어려움
  - **목표**: Python에서 데이터 로드 → SQL로 자유로운 분석
  - **추천 라이브러리**: DuckDB (설치 불필요, pandas 직접 연동, 분석 특화)
  - **구현 방향**:
    - `po_generator/data_layer.py` - 데이터 레이어 모듈
    - `queries/` 폴더 - 자주 쓰는 분석 쿼리 저장
    - CLI 확장 (`--analyze` 옵션 또는 `create_report.py`)

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
- [x] PO (Purchase Order) - xlwings 기반 ✓
- [x] 거래명세표 (Transaction Statement) - xlwings 기반 ✓
- [x] PI (Proforma Invoice) - xlwings 기반 ✓

---

## 라이브러리 선택 기준

| 용도 | 라이브러리 | 이유 |
|------|-----------|------|
| 문서 생성 (PO/TS/PI) | xlwings | 로고/도장 이미지, 복잡한 서식 완벽 보존 |
| 이력 조회/테스트 검증 | openpyxl | COM 인터페이스 없이 안정적인 읽기 |

### 템플릿 동작 방식
- 템플릿 파일의 **데이터는 무시됨** - 코드에서 초기화 후 새로 채움
- 템플릿의 **구조/서식만 유지됨**: 레이아웃, 서식, 이미지, 수식
- 새 템플릿 추가 시: `templates/` 폴더에 양식 파일 추가 후 코드에서 셀 매핑 정의
