# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

NOAH Purchase Order Auto-Generator - RCK(Rotork Korea Sales Office)에서 NOAH(Intercompany Factory)로 발주서를 자동 생성하는 도구.

### Business Context
- 2025년 3월: Rotork이 한국 액추에이터 업체 NOAH 인수
- 2026년 1월: NOAH 영업기능이 RCK로 이전
- RCK = Selling Entity (D365 CE: Lead → Opportunity → Quote → Sales Order)
- NOAH = Factory (D365 F&O: Sales Order → Works Order → Despatch → Invoice)
- ERP 미통합으로 RCK→NOAH 발주는 엑셀 양식으로 처리

### Workflow
1. 고객 발주 접수 → NOAH_SO_PO_DN.xlsx에 정보 입력
2. 문서 ID 입력 → 해당 문서 자동 생성
3. 생성 이력 po_history/ 폴더에 건별 파일로 기록 (중복 발주 방지, 데이터 스냅샷)

### Data Source & Documents
`NOAH_SO_PO_DN.xlsx` 파일이 데이터베이스 역할 (ERP 통합 전까지)

```
NOAH_SO_PO_DN.xlsx (데이터 소스)
       │
       ├── 국내 시트 ─→ PO, 거래명세표
       │
       └── 해외 시트 ─→ PO, PI, CI, PL
```

| 문서 | 용도 | 상태 |
|------|------|------|
| PO (Purchase Order) | RCK→NOAH 발주서 | 완료 |
| 거래명세표 | 국내 납품/선수금 | 완료 |
| PI (Proforma Invoice) | 해외 견적서 | 완료 |
| CI (Commercial Invoice) | 해외 상업송장 | 예정 |
| PL (Packing List) | 해외 포장명세서 | 예정 |

## Installation (최초 1회)

### 1. Miniconda 설치
- 다운로드: https://docs.conda.io/en/latest/miniconda.html
- Windows 64-bit 버전 설치

### 2. Conda 가상환경 생성
```bash
# 명령 프롬프트(cmd) 또는 Anaconda Prompt에서 실행
conda create -n po-automate python=3.11
conda activate po-automate
pip install -r requirements.txt
```

### 3. 설정 파일 생성
처음 사용 시 아래 2개 파일을 설정해야 합니다.

### 설정 방법 (Sublime Text 사용)
1. `user_settings.example.py` 열기 → **다른 이름으로 저장** → `user_settings.py`
2. `local_config.example.bat` 열기 → **다른 이름으로 저장** → `local_config.bat`
3. 각 파일에서 본인 경로로 수정

### user_settings.py (Python용)
```python
# 필수: 데이터 파일 경로 (본인 OneDrive 경로로 수정)
DATA_FOLDER = r"C:\Users\본인이름\OneDrive - Rotork plc\바탕 화면\업무\NOAH ACTUATION"

# 선택: 출력 폴더 (None이면 프로젝트 폴더)
OUTPUT_BASE_DIR = None

# 선택: 공급자 정보 (거래명세표용)
SUPPLIER_INFO = {
    'name': '로토크 콘트롤즈 코리아㈜',
    'rep_name': '이민수',
    'business_no': '220-81-21175',
    'address': '경기도 성남시 분당구 장미로 42',
    'address2': '야탑리더스빌딩 515',
    'business_type': '도매업, 제조, 도매',
    'business_item': '기타운수및기계장비, 밸브류, 무역',
}

# 선택: 비즈니스 규칙
MIN_LEAD_TIME_DAYS = 7      # 납기일 경고 기준
VAT_RATE_DOMESTIC = 0.1     # 부가세율

# 선택: 이력 표시
HISTORY_CUSTOMER_DISPLAY_LENGTH = 15
HISTORY_DESC_DISPLAY_LENGTH = 20
```

### local_config.bat (배치 파일용)
```batch
@echo off
REM 본인 Python 경로로 수정
set PYTHON_PATH=%LOCALAPPDATA%\miniconda3\envs\po-automate\python.exe
```

**참고**: 두 파일 모두 `.gitignore`에 포함되어 Git에 올라가지 않음

## Commands

Run the PO generator using the conda environment:

**회사 컴퓨터 (Jeongtaek.Bang):**
```bash
%LOCALAPPDATA%\miniconda3\envs\po-automate\python.exe create_po.py <ORDER_NO>
```

**개인 컴퓨터 (since):**
```bash
C:/Users/since/anaconda3/envs/po-automate/python.exe create_po.py <ORDER_NO>
```

Examples:
```bash
# Single order
python create_po.py ND-0001

# Multiple orders
python create_po.py ND-0001 ND-0002 ND-0003

# Force create (skip duplicate warning and validation errors)
python create_po.py ND-0001 --force

# View PO history
python create_po.py --history

# Export history to Excel (전체 데이터 스냅샷 포함)
python create_po.py --history --export
```

### Validation
발주서 생성 시 다음 항목을 자동 검증:
- **필수 필드**: Customer name, Customer PO, Item qty, Model, ICO Unit
- **ICO Unit**: 0 또는 음수이면 오류
- **납기일**: 과거이면 오류, 7일 이내면 경고

Alternatively, use the batch file for interactive mode:
```
create_po.bat
```

## Architecture

### Project Structure (v2.4 - 코드 리팩토링)
```
purchaseOrderAutomation/
├── po_generator/           # 핵심 패키지
│   ├── __init__.py
│   ├── config.py           # 설정/상수 (경로, 색상, 필드, 템플릿 경로)
│   ├── utils.py            # 유틸리티 (데이터 로드, get_value)
│   ├── validators.py       # 데이터 검증 로직
│   ├── history.py          # 이력 관리 (중복 체크, 저장)
│   ├── excel_helpers.py    # Excel 헬퍼 (find_item_start_row 통합)
│   ├── excel_generator.py  # PO Excel 생성 (xlwings 기반)
│   ├── template_engine.py  # PO 템플릿 생성 (Deprecated - xlwings로 대체)
│   ├── ts_generator.py     # 거래명세표 생성 (xlwings 기반)
│   ├── pi_generator.py     # Proforma Invoice 생성 (xlwings 기반)
│   ├── logging_config.py   # 로깅 설정
│   └── services/           # 서비스 레이어
│       ├── __init__.py
│       ├── document_service.py  # 문서 생성 오케스트레이터
│       ├── finder_service.py    # 데이터 조회 서비스
│       └── result.py            # 결과 클래스 (DocumentResult)
├── templates/              # 템플릿 파일
│   ├── purchase_order.xlsx        # PO 템플릿
│   ├── transaction_statement.xlsx # 거래명세표 템플릿
│   └── proforma_invoice.xlsx      # PI 템플릿
├── tests/                  # pytest 테스트
│   ├── conftest.py
│   ├── test_validators.py
│   ├── test_history.py
│   ├── test_utils.py
│   ├── test_excel_generator.py
│   ├── test_excel_helpers.py     # excel_helpers 테스트 (신규)
│   ├── test_cli_common.py        # CLI 공통 함수 테스트 (신규)
│   ├── test_config.py            # 설정 검증 테스트 (신규)
│   └── test_integration.py       # 통합 테스트
├── create_po.py            # PO CLI 진입점
├── create_ts.py            # 거래명세표 CLI 진입점
├── create_pi.py            # Proforma Invoice CLI 진입점
├── create_po.bat           # Windows 배치 파일
├── NOAH_PO_Lists.xlsx      # 소스 데이터 (국내/해외)
├── po_history/             # 발주 이력 (월별 폴더)
│   └── YYYY/M월/           # 연/월별 폴더
│       └── YYYYMMDD_주문번호_고객명.xlsx
├── generated_po/           # 생성된 발주서 폴더
├── generated_ts/           # 생성된 거래명세표 폴더
├── generated_pi/           # 생성된 Proforma Invoice 폴더
├── requirements.txt
└── .gitignore
```

### Data Flow
1. `NOAH_PO_Lists.xlsx` - Source data with two sheets: 국내 (domestic) and 해외 (export)
2. `create_po.py` - CLI that orchestrates the generation process
3. `po_generator/` - Core package with modular components
4. `generated_po/` - Output directory for generated Excel files
5. `po_history/` - 건별 이력 파일 (발주서에서 추출한 DB 형식 스냅샷, 중복 방지)

### Key Modules
| Module | Responsibility |
|--------|----------------|
| `config.py` | 경로, 색상, 필드 정의, 상수, 템플릿 경로, COLUMN_ALIASES |
| `utils.py` | get_value (표준 API), load_noah_po_lists, find_order_data |
| `validators.py` | ValidationResult, validate_order_data, validate_multiple_items |
| `history.py` | check_duplicate_order, save_to_history (발주서→DB 형식), get_all_history |
| `excel_helpers.py` | XlConstants, xlwings_app_context, find_item_start_row, 배치 연산 헬퍼 |
| `template_engine.py` | generate_po_template (Deprecated - xlwings로 대체됨) |
| `excel_generator.py` | create_po_workbook (xlwings 기반 PO 생성) |
| `ts_generator.py` | create_ts_xlwings (xlwings 기반 거래명세표 생성) |
| `pi_generator.py` | create_pi_xlwings (xlwings 기반 Proforma Invoice 생성) |
| `logging_config.py` | 로깅 설정 (DEBUG 레벨 제어) |
| `services/document_service.py` | DocumentService - 문서 생성 오케스트레이터 |
| `services/finder_service.py` | FinderService - 데이터 조회 서비스 (지연 로딩) |
| `services/result.py` | DocumentResult, GenerationStatus - 결과 클래스 |
| `cli_common.py` | validate_output_path (보안), generate_output_filename |

### Excel Template Structure
템플릿 기반으로 발주서를 생성합니다:
- **템플릿 파일**: `templates/purchase_order.xlsx`
- **사용자가 직접 로고/도장 이미지를 템플릿에 추가 가능**
- 템플릿이 없으면 첫 실행 시 자동 생성

Generated files contain two sheets:
- **Purchase Order** - Vendor info, item details, pricing, delivery terms (Rotork format)
- **Description** - Actuator specifications (SPEC_FIELDS) and options (OPTION_FIELDS)

### Template Customization (로고/도장 추가)
1. `templates/purchase_order.xlsx` 파일을 Excel에서 열기
2. 원하는 위치에 로고/도장 이미지 삽입 (예: J2 셀 근처)
3. 저장 후 발주서 생성 시 이미지가 자동 포함됨

**주의**: 템플릿의 셀 구조(Row 1-25)는 변경하지 마세요. 스타일, 색상, 이미지만 수정 가능합니다.

## Dependencies

- pandas
- openpyxl (이력 조회, 테스트 검증용)
- xlwings (모든 문서 생성 - 이미지/서식 완벽 보존)
- pytest (dev)
- pytest-cov (dev)

## Testing

```bash
# Run all tests
python -m pytest tests/ -v

# Run with coverage
python -m pytest tests/ --cov=po_generator
```

## TODO (향후 작업)

### OneDrive 공유 폴더 연동
- [x] 회사 랩탑에서 OneDrive 공유 폴더 경로 확인 ✓
  - 경로: `C:\Users\Jeongtaek.Bang\OneDrive - Rotork plc\바탕 화면\업무\NOAH ACTUATION\purchaseOrderAutomation`
- [x] `config.py`에서 경로 설정 외부화 → `user_settings.py` ✓
- [x] 파일 구조 변경 ✓
  ```
  OneDrive - Rotork plc/바탕 화면/업무/NOAH ACTUATION/
  ├── NOAH_PO_Lists.xlsx              (소스 데이터 - 상위 폴더)
  └── purchaseOrderAutomation/        (PO 생성기)
      ├── po_generator/
      ├── po_history/                 (이력 - 월별 폴더)
      └── generated_po/               (생성된 발주서)
  ```
- [x] po_history 월별 폴더 방식으로 변경 ✓
  - 구조: `po_history/YYYY/M월/YYYYMMDD_주문번호_고객명.xlsx`
  - **월별로 독립 관리** - 누적은 사용자가 수동으로 합침
  - **발주서에서 데이터 추출 → DB 형식(한 행)으로 저장**
    - 메타: 생성일시, RCK Order no., Customer name, 원본파일
    - Purchase Order 시트: 고객명, 금액, 납기일, Incoterms 등
    - Description 시트: 사양 필드, 옵션 필드 전체
  - 현재 월 이력 조회: `python create_po.py --history`
  - 현재 월 Excel 내보내기: `python create_po.py --history --export`
  - 수동 발주서도 같은 폴더에 넣으면 집계됨

### 템플릿 기반 문서 생성
- [x] PO (Purchase Order) - xlwings 기반 ✓
  - `templates/purchase_order.xlsx` 템플릿 파일
  - `excel_generator.py` 모듈 (xlwings로 이미지/서식 완벽 보존)
  - 사용자가 직접 로고/도장 이미지 추가 가능
  - **버그 수정 (2026-01-19)**:
    - Delivery Address 값이 안 나오던 문제 해결
      - `config.py`: `delivery_address` 컬럼 별칭 추가
      - `utils.py`: SO→PO 병합 시 `'납품 주소'` 컬럼 누락 수정
      - `excel_generator.py`: 하드코딩 키워드 검색 → `get_value()` 사용
    - 파일 열 때 Description 시트가 먼저 보이던 문제 해결
      - `excel_generator.py`: `wb.active = ws_po` 추가 (Purchase Order 시트 활성화)
  - **리팩토링 (2026-01-20)**:
    - openpyxl → xlwings 전환 (이미지/서식 보존)
    - `get_safe_value` → `get_value` 표준 API로 통일
  - **버그 수정 (2026-01-21)**:
    - Description 시트 A열 레이블 누락 문제 해결
      - **원인**: 템플릿의 고정 레이블에만 의존, 동적 필드(`get_spec_option_fields`)와 불일치
      - **수정**: A열에 레이블 명시적 쓰기 (`['Line No', 'Qty'] + all_fields`)
      - `_apply_description_borders` 함수 추가 (테두리 적용)
    - 국내/해외 모두 동적 필드 사용 (PO_국내: 47개, PO_해외: 45개)
- [x] 거래명세표 (Transaction Statement) - xlwings 기반 ✓
  - `templates/transaction_statement.xlsx` 템플릿 파일
  - `ts_generator.py` 모듈 (xlwings로 이미지/서식 완벽 보존)
  - `create_ts.py` CLI 진입점
  - **버그 수정 (2026-01-19)**:
    - 템플릿 예시 아이템이 삭제되지 않던 문제 해결 (실제 아이템 < 템플릿 예시 시 초과 행 삭제)
    - 행 삭제 후 테두리 복원 (`_restore_ts_item_borders`) - 헤더 하단/마지막 아이템 하단 테두리
- [x] PI (Proforma Invoice) - xlwings 기반 ✓
  - `templates/proforma_invoice.xlsx` 템플릿 파일
  - `pi_generator.py` 모듈 (xlwings로 이미지/서식 완벽 보존)
  - `create_pi.py` CLI 진입점
  - **버그 수정 (2026-01-19)**:
    - 템플릿 예시 아이템이 삭제되지 않던 문제 해결 (실제 아이템 < 템플릿 예시 시 초과 행 삭제)
    - Shipping Mark 영역 검색 범위 수정 (40→20 시작) - 행 삭제 후 위치 변경 대응
    - 행 삭제 후 테두리 복원 (`_restore_item_borders`) - 헤더 하단/마지막 아이템 하단 테두리
  - **버그 수정 (2026-01-21)**:
    - 행 삽입 시 테두리 문제 해결 (실제 아이템 > 템플릿 예시)
      - **증상**: 템플릿 마지막 행(8행) 테두리가 중간에 남음 + Total 위 선 누락
      - **원인**: 행 삽입 케이스에서 `_restore_item_borders` 미호출
      - **수정**: 삽입 전 템플릿 원래 마지막 행 테두리 제거 (`XlConstants.xlNone`)
      - **수정**: 삽입 후 `_restore_item_borders` 호출로 새 마지막 행 테두리 추가
    - `excel_helpers.py`에 `XlConstants.xlNone = -4142` 상수 추가
- [ ] 추후 확장 예정 (xlwings 사용, 해외 오더):
  - Packing List 템플릿
  - Commercial Invoice 템플릿

### 템플릿 동작 방식
- 템플릿 파일의 **데이터는 무시됨** - 코드에서 초기화 후 새로 채움
- 템플릿의 **구조/서식만 유지됨**:
  - 레이아웃 (행/열 위치)
  - 서식 (폰트, 테두리, 색상)
  - 이미지 (로고, 도장)
  - 수식 (소계 SUM 등)
- 새 템플릿 추가 시: `templates/` 폴더에 양식 파일 추가 후 코드에서 셀 매핑 정의

### 라이브러리 선택 기준
| 용도 | 라이브러리 | 이유 |
|------|-----------|------|
| 문서 생성 (PO/TS/PI) | xlwings | 로고/도장 이미지, 복잡한 서식 완벽 보존 |
| 이력 조회/테스트 검증 | openpyxl | COM 인터페이스 없이 안정적인 읽기 |

### SQL 기반 데이터 분석 (예정)
- [ ] NOAH_SO_PO_DN.xlsx → 로컬 SQL DB 연동
  - **배경**: Power Pivot의 양방향 JOIN 제한으로 복잡한 분석 어려움
  - **목표**: Python에서 데이터 로드 → SQL로 자유로운 분석
  - **추천 라이브러리**: DuckDB (설치 불필요, pandas 직접 연동, 분석 특화)
  - **구현 방향**:
    - `po_generator/data_layer.py` - 데이터 레이어 모듈
    - `queries/` 폴더 - 자주 쓰는 분석 쿼리 저장
    - CLI 확장 (`--analyze` 옵션 또는 `create_report.py`)

### 서비스 레이어 추가 (2026-01-21)
- [x] `excel_helpers.py` 생성 ✓
  - `find_item_start_row` 함수를 단일 모듈로 통합
  - openpyxl/xlwings 버전 모두 지원
  - 헤더 라벨 프리셋: `PO_HEADER_LABELS`, `TS_HEADER_LABELS`, `PI_HEADER_LABELS`
- [x] `services/` 디렉토리 생성 ✓
  - `DocumentService`: 문서 생성 오케스트레이터 (PO, TS, PI)
  - `FinderService`: 데이터 조회 서비스 (지연 로딩)
  - `DocumentResult`: 생성 결과 클래스 (성공/실패/중복/검증오류 등)
- [x] CLI 리팩토링 ✓
  - `create_po.py`, `create_ts.py`, `create_pi.py`에서 `DocumentService` 사용
  - 사용자 상호작용(중복 확인, 검증 오류 확인)은 CLI에서 유지
  - 비즈니스 로직은 서비스 레이어로 분리
- [x] 행 삭제 주석 수정 ✓
  - "뒤에서부터 삭제" → "같은 위치에서 반복 삭제 - xlUp으로 아래 행이 올라옴"
  - `ts_generator.py`, `pi_generator.py` 주석 수정
- [x] 통합 테스트 추가 ✓
  - `tests/test_integration.py`: 11개 테스트 케이스
  - 파일명 시퀀스, find_item_start_row 일관성, 행 삭제 동작 문서화

### xlwings 성능 최적화 (2026-01-21)
- [x] 배치 연산 헬퍼 함수 추가 ✓
  - `excel_helpers.py`에 새 함수 추가:
    - `batch_write_rows`: 2D 리스트를 한 번에 쓰기
    - `batch_read_column`: 열의 값을 한 번에 읽기
    - `batch_read_range`: 범위의 값을 한 번에 읽기
    - `delete_rows_range`: 연속 행을 한 번에 삭제
    - `find_text_in_column_batch`: 배치 읽기로 텍스트 찾기
- [x] `ts_generator.py` 최적화 ✓
  - `_fill_items_batch`: 아이템 데이터 배치 쓰기 (N*8회 → 1회 COM 호출)
  - `_find_label_row`: 배치 읽기 최적화 (36회 → 1회)
  - `_find_ts_subtotal_row`: 배치 읽기 최적화 (15회 → 1회)
  - `delete_rows_range` 사용 (N회 → 1회)
- [x] `pi_generator.py` 최적화 ✓
  - `_fill_items_batch`: 아이템 데이터 열별 배치 쓰기 (N*4회 → 4회)
  - `_find_total_row`: 배치 읽기 최적화 (20회 → 1회)
  - `_fill_shipping_mark`: 배치 읽기 최적화 (80회 → 2회)
  - `delete_rows_range` 사용 (N회 → 1회)
- [x] `excel_generator.py` 최적화 ✓
  - `_fill_items_batch_po`: 아이템 데이터 열별 배치 쓰기 + 수식 배치
  - `_create_description_sheet`: 필드 데이터 2D 배치 쓰기 (30*N회 → 1회)
  - `_find_totals_row`: 배치 읽기 최적화 (20회 → 1회)
  - `delete_rows_range` 사용 (N회 → 1회)
- **예상 성능 개선** (50개 아이템 기준):
  | 파일 | COM 호출 (전) | COM 호출 (후) | 감소율 |
  |------|--------------|--------------|--------|
  | ts_generator.py | ~500회 | ~20회 | 96% |
  | pi_generator.py | ~350회 | ~15회 | 96% |
  | excel_generator.py | ~1,500회 | ~50회 | 97% |

### 버그 수정: xlwings 범위 formula 읽기 (2026-01-21)

**증상**: 거래명세표 생성 시 템플릿의 예시 아이템이 삭제되지 않고 그대로 남아있음

**원인**: `_find_ts_subtotal_row` 함수의 배치 읽기 최적화에서 xlwings의 `.formula` 속성 반환 형식을 잘못 처리

**상세 분석**:
```python
# xlwings 범위 읽기 반환 형식 차이
ws.range('E13').value           # 단일 셀 → float: 8.0
ws.range('E13:E17').value       # 범위 → list: [8.0, 8.0, 16.0, None, None]

ws.range('E15').formula         # 단일 셀 → str: '=SUM(E13:E14)'
ws.range('E13:E17').formula     # 범위 → tuple of tuples: (('8',), ('8',), ('=SUM(E13:E14)',), ('',), ('',))
```

- `.value`: 단일 열 범위 → **1D list** 반환
- `.formula`: 단일 열 범위 → **tuple of tuples** 반환 (2D 형태)

**버그 코드**:
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

**영향 범위 확인**:
| 모듈 | 함수 | 사용 속성 | 상태 |
|------|------|----------|------|
| `ts_generator.py` | `_find_ts_subtotal_row` | `.formula` (범위) | **수정됨** |
| `pi_generator.py` | `_find_total_row` | `.value` (범위) | 문제 없음 |
| `excel_generator.py` | `_find_totals_row` | `.value` (범위) | 문제 없음 |
| `excel_helpers.py` | `batch_read_column` | `.value` (범위) | 문제 없음 |

**교훈**:
- xlwings에서 `.value`와 `.formula`는 범위 읽기 시 반환 형식이 다름
- `.value`: 1D list (단일 열)
- `.formula`: 2D tuple of tuples (항상 2D)
- 배치 최적화 시 반환 형식을 실제 테스트로 확인 필요

### 코드 리팩토링 (2026-01-21)

Code Reflection 결과를 바탕으로 코드 품질 개선 작업 수행.

#### Phase 1: excel_helpers.py 인프라 추가
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

#### Phase 2: cli_common.py 보안 수정
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

#### Phase 3: Generator 리팩토링
- [x] `excel_generator.py` 리팩토링 ✓
  - `xlwings_app_context` 사용으로 리소스 관리 개선
  - `XlConstants` 사용으로 매직 넘버 제거
  - 타입 변환 실패 시 경고 로깅 추가
- [x] `ts_generator.py` 리팩토링 ✓
  - `xlwings_app_context` 사용
  - 로컬 상수 → `XlConstants` 교체
  - 타입 변환 경고 추가
- [x] `pi_generator.py` 리팩토링 ✓
  - `xlwings_app_context` 사용
  - 로컬 상수 → `XlConstants` 교체
  - 타입 변환 경고 추가

#### Phase 4: 테스트 커버리지 확대
- [x] `tests/test_excel_helpers.py` 신규 생성 ✓ (16개 테스트)
  - `XlConstants` 상수값 검증
  - `prepare_template` 기능 테스트
  - `cleanup_temp_file` 기능 테스트
  - `xlwings_app_context` 모킹 테스트
- [x] `tests/test_cli_common.py` 신규 생성 ✓ (11개 테스트)
  - 정상 경로 허용 테스트
  - Path traversal 거부 테스트
  - Substring 공격 거부 테스트
  - 중첩 서브디렉토리 허용 테스트
- [x] `tests/test_config.py` 신규 생성 ✓ (22개 테스트)
  - `COLUMN_ALIASES` 필수 키 존재 검증
  - Alias 튜플 형식 검증
  - `REQUIRED_FIELDS`, `META_COLUMNS` 테스트
  - `Colors`, `ColumnWidths` 데이터클래스 테스트

**테스트 결과**: 47 passed, 2 skipped

**변경 파일 요약**:
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

#### Phase 5: utils.py 중복 함수 통합
- [x] `_find_data_by_id()` 공통 헬퍼 추가 ✓
  - ID로 데이터 검색하는 공통 로직 통합
  - 파라미터: `column_key`, `id_value`, `id_label`, `allow_multiple`
- [x] 4개 find 함수를 wrapper로 변경 ✓
  | 함수 | 변경 전 | 변경 후 |
  |------|--------|--------|
  | `find_order_data()` | 34줄 | 1줄 (wrapper) |
  | `find_dn_data()` | 33줄 | 1줄 (wrapper) |
  | `find_pmt_data()` | 28줄 | 1줄 (wrapper) |
  | `find_so_export_data()` | 33줄 | 1줄 (wrapper) |
- **효과**: ~90줄 중복 제거, 버그 수정 시 단일 지점 수정
- **테스트 결과**: 160 passed, 2 skipped (기존 API 100% 호환)

### 거래명세표 기능 개선 (2026-01-31)

#### 출고일 기준 날짜 표시
- **변경**: 거래명세표 날짜를 오늘 날짜 → **출고일** 기준으로 변경
- `config.py`: `dispatch_date` 별칭 추가 (`'출고일', 'Dispatch Date', 'dispatch_date', '출하일'`)
- `ts_generator.py`: 헤더(B2)와 아이템(A열) 날짜를 출고일로 표시
  - 출고일이 없으면 오늘 날짜를 폴백으로 사용
  - 파라미터명 `today` → `dispatch_date`로 변경

#### 월합 거래명세표 기능 추가
고객이 월합으로 거래명세표를 요청할 때, 여러 DN을 한 장으로 합쳐서 생성

**사용법:**
```bash
# 명령줄에서 직접
python create_ts.py DND-2026-0001 DND-2026-0002 DND-2026-0003 --merge

# 대화형 모드 (여러 줄 붙여넣기 지원)
python create_ts.py --interactive --merge
```

**배치 파일 (create_po.bat):**
```
[2] 거래명세표 생성 → [2] 월합 거래명세표
→ DN_ID 목록 세로로 붙여넣기
→ 빈 줄 입력 (Enter)
→ 생성 완료
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
