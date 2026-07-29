# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

NOAH Document Auto-Generator — automates creation of inter-company business documents (PO, Transaction Statement, Proforma Invoice, Final Invoice, Order Confirmation) for RCK (Rotork Korea) → NOAH (Intercompany Factory) transactions. Data lives in Excel (`NOAH_SO_PO_DN.xlsx`) because the two ERP systems (D365 CE and D365 F&O) are not integrated.

**Language:** Python 3.11+ on Windows. **Key libs:** pandas, openpyxl (PO generation), xlwings (TS/PI/FI/OC — needs Excel COM for images/formulas), pytest.

## Commands

```bash
# Environment setup
conda create -n po-automate python=3.11 && conda activate po-automate
pip install -r requirements.txt

# Run document generation
python create_po.py ND-0001              # Single PO
python create_po.py ND-0001 --force      # Skip validation errors
python create_ts.py DND-2026-0001 --merge  # Merged transaction statement
python create_ts.py DND-2026-0001           # 생성 후 "이메일 발송?" y/N 확인 → y면 Outlook
python create_ts.py DND-2026-0001 --mail    # 확인 없이 바로 Outlook 창
python create_ts.py DND-2026-0001 --send    # 확인 없이 즉시 발송 (Outlook COM 환경만)
python create_ts.py DND-2026-0001 --no-mail # 묻지 않고 문서만 (배치용)
python create_pi.py NO-0001              # Proforma invoice
python create_fi.py DNO-2026-0001        # Final invoice (복수 RCK PO 시 발주번호별 자동 분리)
python create_fi.py --po 26KPO00144     # 발주번호 기준 FI 생성 (복수 DN 통합)
python create_fi.py --po                # 사용 가능한 발주번호 목록 표시
python create_oc.py SOO-2026-0001        # Order confirmation

# DB sync & snapshot
python sync_db.py                         # Excel → SQLite sync
python close_period.py 2026-01            # Monthly close (snapshot)
python close_period.py --undo 2026-01     # Undo last close
python close_period.py --list             # Snapshot history
python close_period.py --status           # Current status

# PO 매입대사
python reconcile_po.py P03                # 3월 대사 (대사결과 + AX_PO_매핑)
python reconcile_po.py P03 -v             # 상세 로그

# SO 매출대사
python reconcile_so.py P03                # 3월 매출대사 (AX vs NOAH DN)
python reconcile_so.py P03 -v             # 상세 로그

# Industry Code 대사 + Sector 검증
python reconcile_ind.py P03               # Industry code 채움 + Sector 검증
python reconcile_ind.py --sector-only     # Sector 검증만 (월 입력 불필요)
python reconcile_ind.py P03 -v            # 상세 로그

# Dashboard
streamlit run dashboard.py                # Streamlit 대시보드

# GUI (문서 생성 7종) / 사내 배포판 빌드
python noah_gui.py                        # tkinter GUI (개발 PC에서도 그대로 실행)
python cli_dist/build_portable_gui.py     # 배포 zip 빌드 → cli_dist/NOAH_문서생성기_배포.zip

# Tests
pytest                                    # All tests
pytest --cov=po_generator                 # With coverage
pytest tests/test_create_po.py -v         # Single test file
pytest tests/test_create_po.py::test_name -v  # Single test
```

## Architecture

```
CLI entry points (create_po.py, create_ts.py, create_pi.py, create_fi.py, create_oc.py)
    ↓
Service layer (po_generator/services/document_service.py, finder_service.py)
    ↓
Generators (excel_generator.py=openpyxl, ts/pi/fi/oc_generator.py=xlwings)
    ↓
Shared: config.py (paths, constants, aliases), utils.py (data loading), validators.py

DB layer:
  sync_db.py → db_sync.py (Excel→SQLite, upsert+prune) → db_schema.py (DDL)
  close_period.py → snapshot.py (SnapshotEngine) → db_schema.py (snapshot tables)
  sql/order_book.sql (이벤트 기반), sql/order_book_snapshot.sql (snapshot-based)

Reconciliation layer:
  reconcile_po.py → 3-source merge (Delivery + Internal PO + GRN) → 대사결과 + AX_PO_매핑
  reconcile_so.py → 2-source merge (AX Sales + NOAH DN) → 대사결과_SO (매출일 필터: 국내=출고일, 해외=선적일)
  reconcile_ind.py → Industry code 채움 (PO→SO 매핑) + SO Sector 검증 (마스터 Category 교차)
  po_reconciliation/{year}/{period}/ — PO input/output files per period (flat `{period}/` 도 자동 호환)
  so_reconciliation/{year}/{period}/ — SO input/output files per period
  ind_code_reconciliation/{year}/{period}/ — Industry code input/output files per period (sector_검증.xlsx은 루트)
```

**Data flow:** CLI → FinderService loads Excel data → validators check fields → generator fills template → output saved to `generated_*/` + history snapshot to `po_history/YYYY/M월/`.

### Key Design Patterns

- **Column Alias System** (`config.py: COLUMN_ALIASES`): Maps internal names to multiple possible Excel column headers. `resolve_column()` auto-detects actual names — critical for resilience to Excel schema changes.
- **Dual Library Strategy**: openpyxl for PO (fast, no image needs); xlwings for TS/PI/FI/OC (preserves images, formulas, COM-dependent).
- **Template Engine** (`template_engine.py`): Clones rows for multi-item orders, auto-adjusts SUM formulas after row insertion.
- **History as DB**: `po_history/YYYY/M월/YYYYMMDD_주문번호_고객명.xlsx` — one file per transaction enables duplicate detection without a database.
- **Result Pattern** (`services/result.py`): `DocumentResult` + `GenerationStatus` enum for structured operation outcomes. `history_saved` field tracks history persistence separately from generation success.
- **Output File Safety** (`cli_common.py`): Generated files auto-suffix on collision (`_1`, `_2`, ...) to prevent silent overwrites. Raises `FileExistsError` if 100+ collisions.
- **DB Sync Prune** (`db_sync.py`): Excel→SQLite sync includes prune step — rows deleted from Excel are also deleted from DB. Works even when sheet is completely empty. `--dry-run` connects to real DB and rollbacks for accurate diff simulation.
- **Dashboard Error Visibility** (`dashboard.py`): Loader failures collected in `session_state` and displayed as `st.warning()` banner, distinguishing "no data" from "query failure".
- **Snapshot Engine** (`snapshot.py`): Monthly close → `ob_snapshot` freezes Ending, subsequent retroactive changes auto-detected as Variance. Sequential close enforced.

### Configuration Split

- `config.py` — Project constants, paths, sheet names, column aliases, business rules (committed)
- `user_settings.py` — User-specific paths (DATA_FOLDER, OUTPUT_BASE_DIR), supplier info (git-ignored, copy from `user_settings.example.py`)
- `local_config.bat` — Local Python/conda path for batch wrapper (git-ignored)
- `noah_config.ini` — 배포판 경로 설정 (git-ignored, GUI 마법사가 생성). 우선순위는 `user_settings.py` → ini → 기본값이며, **user_settings.py에 이름이 있으면 값이 `None`이어도 그것이 최종값**이다 (`OUTPUT_BASE_DIR = None`을 ini가 덮어쓰면 개발 PC 출력 위치가 조용히 바뀌므로)

## Key Files

| File | Purpose |
|------|---------|
| `po_generator/config.py` | All constants, paths, sheet names (`SO_국내`, `PO_국내`, `DN_국내`...), column aliases |
| `po_generator/utils.py` | Data loading (`load_noah_po_lists`), value extraction, Excel injection prevention |
| `po_generator/validators.py` | Required field checks, ICO Unit > 0, delivery date validation |
| `po_generator/services/document_service.py` | Orchestrator: find → validate → generate → save |
| `po_generator/services/finder_service.py` | Order lookup across domestic/overseas sheets |
| `noah_gui.py` | tkinter GUI — 문서 생성 7종. 기존 `create_*.py`를 **자식 프로세스로 실행**하고 stdout을 로그 위젯에 흘린다(CLI 무수정, COM 격리). 데이터 파일 지정 마법사 포함 |
| `cli_dist/build_portable_gui.py` | 사내 배포판 빌드 — python-build-standalone 런타임 + 앱 파일 + `설치.bat` → zip. 빌드는 임시 폴더에서(OneDrive 동기화·MAX_PATH 회피), 프로젝트엔 zip만 남김 |
| `noah_config.ini` | 배포판 경로 설정 (git-ignored). GUI 마법사가 생성. `user_settings.py`가 있으면 그쪽이 우선 |
| `po_generator/mailer.py` | 거래명세표 메일 발송 — 사업자번호로 `Customer_국내` 수신자 조회, xlsx→PDF 변환, 2가지 백엔드(Outlook COM / `.eml` 초안). 새 Outlook은 COM 미지원이라 `auto`가 `.eml`로 전환 |
| `docs/ARCHITECTURE.md` | Detailed system design and data flow diagrams |
| `docs/DATA_STRUCTURE_DESIGN.md` | Excel schema (8 sheets), Power Query setup |
| `docs/POWER_QUERY.md` | Power Query 수식, Power Pivot 관계 — 데이터 소스 구조 이해 시 참고 |
| `docs/ERP_CONCEPTS.md` | ERP 개념 학습 노트 — SO-PO 라인 매칭, BOM/Kit, 정규화 수준 비교 |
| `docs/CHANGELOG.md` | 버전별 변경 이력 |
| `docs/TEMPLATE_MAPPINGS.md` | Excel 템플릿 셀 매핑 — 템플릿/generator 수정 시 참고 |
| `docs/매입대사_가이드.md` | 운영자 관점 시각 가이드 — 데이터 흐름·합계 블록 해석·월별 액션 |
| `po_generator/snapshot.py` | SnapshotEngine — 월별 마감, Variance 추적 |
| `po_generator/db_schema.py` | SQLite DDL, snapshot tables (`ob_snapshot`, `ob_snapshot_meta`), `_sync_runs` + `_sync_log` v2 (record당 1행 + JSON, actor/host/sync_id/snapshot 포함) |
| `migrate_sync_log.py` | `sync_log.csv` → `_sync_log_legacy_v1` 전용 테이블 1회성 적재 (구 v1 형식; 운영 v2 `_sync_log`와 분리, 이후 `migrate_sync_log_v2.py`로 변환) |
| `migrate_sync_log_v2.py` | `_sync_log` v1 → v2 (record 단위 + JSON 압축, _sync_runs 메타 분리, snapshot 추가) |
| `sql/order_book.sql` | 이벤트 기반 Order Book SQL (Input/Output 이벤트 월만 행 생성, 재귀 CTE 없음) |
| `sql/order_book_snapshot.sql` | 스냅샷 기반 Order Book SQL (마감 고정 + Variance) |
| `sql/order_book_variance.sql` | Variance 변동이유 분석 SQL (환율차이/판매가변경/수량변경/반올림 자동 분류, 납기변경 제외) |
| `dashboard.py` | Streamlit 대시보드 (9페이지: 오늘의현황/수주출고/제품/섹터/고객/발주커버리지/수익성/Order Book/동기화로그, PO미등록감지, PO확정지연, EXW미출고, 납기현황(DN qty매칭+PO EXW보충), 납기캘린더(선적예정 포함), 해외선적(Incoterms/운송방식별), 세금계산서미발행, Order Book 3탭, `_sync_log` 변경이력 조회) |
| `reconcile_po.py` | PO 매입대사 — 공장 출고(Delivery) vs 회계 GRN 금액 비교. 출력: `대사결과_{period}.xlsx` (6시트), `AX_PO_매핑_{period}.xlsx` (Delivery+AX PO) |
| `reconcile_so.py` | SO 매출대사 — AX ERP 매출 vs NOAH DN 매출 비교 (국내=출고일, 해외=선적일 기준 월 필터 + FX 환율차이 자동 판별). 출력: `대사결과_SO_{period}.xlsx` (3시트: 대사/상세/범례) |
| `reconcile_ind.py` | Industry Code 대사 — (1) Orderbook 빈 Industry code를 PO→SO 매핑으로 채움 → `ind_code_결과_{period}.xlsx`, (2) SO Sector vs 마스터 Category 교차 검증 → `sector_검증.xlsx`. `--sector-only`로 검증만 실행 가능 |

## Business Rules

- Order numbers: `ND-*` = domestic, `NO-*` = overseas
- DN numbers: `DND-*` = domestic, `DNO-*` = overseas
- Validation blocks generation unless `--force`: missing required fields, ICO Unit ≤ 0, past delivery date
- Warnings (non-blocking): delivery within 7 days, duplicate order in history
- 거래명세표 메일: 수신자(To)는 `Customer_국내.사업자번호` ↔ `DN_국내.Business registration number` 조인으로 결정, 고정 참조(CC)는 `user_settings.py: TS_MAIL_CC`. 기본은 **수신자를 보여준 뒤 y/N 확인**(비대화형 실행은 자동 OFF — 배치가 프롬프트에서 멈추지 않게). 메일 실패는 문서 생성 성공을 뒤엎지 않으며, 고객이 섞인 `--merge` 문서는 발송 차단

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
