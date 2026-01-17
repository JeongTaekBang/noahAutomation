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
1. 고객 발주 접수 → NOAH_PO_Lists.xlsx에 정보 입력
2. RCK Order No. 입력 → 발주서(Purchase Order + Description) 자동 생성
3. 생성 이력 po_history/ 폴더에 건별 파일로 기록 (중복 발주 방지, 데이터 스냅샷)

## Setup (user_settings.py)

프로젝트 루트에 `user_settings.py` 파일을 생성하고 아래 설정을 입력:

```python
# 필수: 데이터 파일 경로
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

**참고**: `user_settings.py`는 `.gitignore`에 포함되어 Git에 올라가지 않음

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

### Project Structure (v2.2 - 템플릿 기반 + 거래명세표)
```
purchaseOrderAutomation/
├── po_generator/           # 핵심 패키지
│   ├── __init__.py
│   ├── config.py           # 설정/상수 (경로, 색상, 필드, 템플릿 경로)
│   ├── utils.py            # 유틸리티 (데이터 로드, get_safe_value)
│   ├── validators.py       # 데이터 검증 로직
│   ├── history.py          # 이력 관리 (중복 체크, 저장)
│   ├── excel_generator.py  # PO Excel 생성 (openpyxl 기반)
│   ├── template_engine.py  # PO 템플릿 처리 (행 복제, 공식 조정)
│   └── ts_generator.py     # 거래명세표 생성 (xlwings 기반)
├── templates/              # 템플릿 파일
│   ├── purchase_order.xlsx        # PO 템플릿
│   └── transaction_statement.xlsx # 거래명세표 템플릿
├── tests/                  # pytest 테스트
│   ├── conftest.py
│   ├── test_validators.py
│   ├── test_history.py
│   ├── test_utils.py
│   └── test_excel_generator.py
├── create_po.py            # PO CLI 진입점
├── create_ts.py            # 거래명세표 CLI 진입점
├── create_po.bat           # Windows 배치 파일
├── NOAH_PO_Lists.xlsx      # 소스 데이터 (국내/해외)
├── po_history/             # 발주 이력 (월별 폴더)
│   └── YYYY/M월/           # 연/월별 폴더
│       └── YYYYMMDD_주문번호_고객명.xlsx
├── generated_po/           # 생성된 발주서 폴더
├── generated_ts/           # 생성된 거래명세표 폴더
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
| `config.py` | 경로, 색상, 필드 정의, 상수, 템플릿 경로 |
| `utils.py` | get_safe_value, load_noah_po_lists, find_order_data |
| `validators.py` | ValidationResult, validate_order_data, validate_multiple_items |
| `history.py` | check_duplicate_order, save_to_history (발주서→DB 형식), get_all_history |
| `template_engine.py` | load_template, clone_row, generate_po_template (openpyxl 템플릿 처리) |
| `excel_generator.py` | create_po_workbook (openpyxl 기반 PO 생성) |
| `ts_generator.py` | create_ts_xlwings (xlwings 기반 거래명세표 생성) |

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
- openpyxl (PO 생성)
- xlwings (거래명세표 및 향후 문서 - 이미지/서식 완벽 보존)
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
- [x] PO (Purchase Order) - openpyxl 기반 ✓
  - `templates/purchase_order.xlsx` 템플릿 파일
  - `template_engine.py` 모듈 (행 복제, 공식 조정)
  - 사용자가 직접 로고/도장 이미지 추가 가능
- [x] 거래명세표 (Transaction Statement) - xlwings 기반 ✓
  - `templates/transaction_statement.xlsx` 템플릿 파일
  - `ts_generator.py` 모듈 (xlwings로 이미지/서식 완벽 보존)
  - `create_ts.py` CLI 진입점
- [ ] 추후 확장 예정 (xlwings 사용, 해외 오더):
  - Packing List 템플릿
  - Proforma Invoice 템플릿
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
| 문서 | 라이브러리 | 이유 |
|------|-----------|------|
| PO | openpyxl | 이미지 없이 잘 동작, 빠름 |
| 거래명세표/Invoice/Packing List | xlwings | 로고/도장 이미지, 복잡한 서식 완벽 보존 |
