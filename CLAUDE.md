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

## Commands

Run the PO generator using the conda environment:
```bash
C:/Users/since/anaconda3/envs/po-automate/python.exe create_po.py <ORDER_NO>
```

Examples:
```bash
# Single order
C:/Users/since/anaconda3/envs/po-automate/python.exe create_po.py ND-0001

# Multiple orders
C:/Users/since/anaconda3/envs/po-automate/python.exe create_po.py ND-0001 ND-0002 ND-0003

# Force create (skip duplicate warning and validation errors)
C:/Users/since/anaconda3/envs/po-automate/python.exe create_po.py ND-0001 --force

# View PO history
C:/Users/since/anaconda3/envs/po-automate/python.exe create_po.py --history

# Export history to Excel (전체 데이터 스냅샷 포함)
C:/Users/since/anaconda3/envs/po-automate/python.exe create_po.py --history --export
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

### Project Structure (v2.0)
```
purchaseOrderAutomation/
├── po_generator/           # 핵심 패키지
│   ├── __init__.py
│   ├── config.py           # 설정/상수 (경로, 색상, 필드 등)
│   ├── utils.py            # 유틸리티 (데이터 로드, get_safe_value)
│   ├── validators.py       # 데이터 검증 로직
│   ├── history.py          # 이력 관리 (중복 체크, 저장)
│   └── excel_generator.py  # Excel 생성 (PO, Description 시트)
├── tests/                  # pytest 테스트
│   ├── conftest.py
│   ├── test_validators.py
│   ├── test_history.py
│   └── test_utils.py
├── create_po.py            # CLI 진입점
├── create_po.bat           # Windows 배치 파일
├── NOAH_PO_Lists.xlsx      # 소스 데이터 (국내/해외)
├── po_history/             # 발주 이력 (월별 폴더)
│   └── YYYY/M월/           # 연/월별 폴더
│       └── YYYYMMDD_주문번호_고객명.xlsx
├── generated_po/           # 생성된 발주서 폴더
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
| `config.py` | 경로, 색상, 필드 정의, 상수 |
| `utils.py` | get_safe_value, load_noah_po_lists, find_order_data |
| `validators.py` | ValidationResult, validate_order_data, validate_multiple_items |
| `history.py` | check_duplicate_order, save_to_history (발주서→DB 형식), get_all_history |
| `excel_generator.py` | create_purchase_order, create_description_sheet |

### Excel Template Structure
Generated files contain two sheets:
- **Purchase Order** - Vendor info, item details, pricing, delivery terms (Rotork format)
- **Description** - Actuator specifications (SPEC_FIELDS) and options (OPTION_FIELDS)

## Dependencies

- pandas
- openpyxl
- pytest (dev)
- pytest-cov (dev)

## Testing

```bash
# Run all tests
C:/Users/since/anaconda3/envs/po-automate/python.exe -m pytest tests/ -v

# Run with coverage
C:/Users/since/anaconda3/envs/po-automate/python.exe -m pytest tests/ --cov=po_generator
```

## TODO (향후 작업)

### OneDrive 공유 폴더 연동
- [ ] 회사 랩탑에서 OneDrive 공유 폴더 경로 확인
- [ ] `config.py`에서 경로 설정 외부화 (환경변수 또는 설정 파일)
- [ ] 파일 구조 변경:
  ```
  OneDrive 공유폴더/
  ├── NOAH_PO_Lists.xlsx    (소스 - 공유)
  ├── po_history.xlsx       (이력 - 공유)
  └── generated_po/         (생성된 발주서 - 공유)
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
