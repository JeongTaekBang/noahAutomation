# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

NOAH Document Auto-Generator — automates creation of inter-company business documents (PO, Transaction Statement, Proforma Invoice, Final Invoice) for RCK (Rotork Korea) → NOAH (Intercompany Factory) transactions. Data lives in Excel (`NOAH_SO_PO_DN.xlsx`) because the two ERP systems (D365 CE and D365 F&O) are not integrated.

**Language:** Python 3.11+ on Windows. **Key libs:** pandas, openpyxl (PO generation), xlwings (TS/PI/FI — needs Excel COM for images/formulas), pytest.

## Commands

```bash
# Environment setup
conda create -n po-automate python=3.11 && conda activate po-automate
pip install -r requirements.txt

# Run document generation
python create_po.py ND-0001              # Single PO
python create_po.py ND-0001 --force      # Skip validation errors
python create_ts.py DND-2026-0001 --merge  # Merged transaction statement
python create_pi.py NO-0001              # Proforma invoice
python create_fi.py ND-0001              # Final invoice

# Tests
pytest                                    # All tests
pytest --cov=po_generator                 # With coverage
pytest tests/test_create_po.py -v         # Single test file
pytest tests/test_create_po.py::test_name -v  # Single test
```

## Architecture

```
CLI entry points (create_po.py, create_ts.py, create_pi.py, create_fi.py)
    ↓
Service layer (po_generator/services/document_service.py, finder_service.py)
    ↓
Generators (excel_generator.py=openpyxl, ts/pi/fi_generator.py=xlwings)
    ↓
Shared: config.py (paths, constants, aliases), utils.py (data loading), validators.py
```

**Data flow:** CLI → FinderService loads Excel data → validators check fields → generator fills template → output saved to `generated_*/` + history snapshot to `po_history/YYYY/M월/`.

### Key Design Patterns

- **Column Alias System** (`config.py: COLUMN_ALIASES`): Maps internal names to multiple possible Excel column headers. `resolve_column()` auto-detects actual names — critical for resilience to Excel schema changes.
- **Dual Library Strategy**: openpyxl for PO (fast, no image needs); xlwings for TS/PI/FI (preserves images, formulas, COM-dependent).
- **Template Engine** (`template_engine.py`): Clones rows for multi-item orders, auto-adjusts SUM formulas after row insertion.
- **History as DB**: `po_history/YYYY/M월/YYYYMMDD_주문번호_고객명.xlsx` — one file per transaction enables duplicate detection without a database.
- **Result Pattern** (`services/result.py`): `DocumentResult` + `GenerationStatus` enum for structured operation outcomes.

### Configuration Split

- `config.py` — Project constants, paths, sheet names, column aliases, business rules (committed)
- `user_settings.py` — User-specific paths (DATA_FOLDER, OUTPUT_BASE_DIR), supplier info (git-ignored, copy from `user_settings.example.py`)
- `local_config.bat` — Local Python/conda path for batch wrapper (git-ignored)

## Key Files

| File | Purpose |
|------|---------|
| `po_generator/config.py` | All constants, paths, sheet names (`SO_국내`, `PO_국내`, `DN_국내`...), column aliases |
| `po_generator/utils.py` | Data loading (`load_noah_po_lists`), value extraction, Excel injection prevention |
| `po_generator/validators.py` | Required field checks, ICO Unit > 0, delivery date validation |
| `po_generator/services/document_service.py` | Orchestrator: find → validate → generate → save |
| `po_generator/services/finder_service.py` | Order lookup across domestic/overseas sheets |
| `docs/ARCHITECTURE.md` | Detailed system design and data flow diagrams |
| `docs/DATA_STRUCTURE_DESIGN.md` | Excel schema (8 sheets), Power Query setup |
| `docs/POWER_QUERY.md` | Power Query 수식, Power Pivot 관계 — 데이터 소스 구조 이해 시 참고 |
| `docs/CHANGELOG.md` | 버전별 변경 이력 |
| `docs/TEMPLATE_MAPPINGS.md` | Excel 템플릿 셀 매핑 — 템플릿/generator 수정 시 참고 |

## Business Rules

- Order numbers: `ND-*` = domestic, `NO-*` = overseas
- DN numbers: `DND-*` = domestic, `DNO-*` = overseas
- Validation blocks generation unless `--force`: missing required fields, ICO Unit ≤ 0, past delivery date
- Warnings (non-blocking): delivery within 7 days, duplicate order in history

## Self-Improvement Loop

- After ANY correction from the user: update `tasks/lessons.md` with the pattern
- Write rules for yourself that prevent the same mistake
- Review lessons at session start for relevant project

## Task Management

1. **Plan First**: Write plan to `tasks/todo.md` with checkable items
2. **Verify Plan**: Check in before starting implementation
3. **Track Progress**: Mark items complete as you go
4. **Explain Changes**: High-level summary at each step
5. **Document Results**: Add review section to `tasks/todo.md`
6. **Capture Lessons**: Update `tasks/lessons.md` after corrections
