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
python create_ts.py --date 2026-08-06 --customer 씨앤케이  # 그날 그 거래처 출고분 → 메일 1통(첨부 N개)
python create_ts.py --date 2026-08-06                     # 그날 전체 → 거래처별 1통씩
python create_ts.py DND-2026-0742 DND-2026-0743 --one-mail # 명시 ID를 묶어 1통
python create_pi.py NO-0001              # Proforma invoice
python create_fi.py DNO-2026-0001        # Final invoice (복수 RCK PO 시 발주번호별 자동 분리)
python create_fi.py --po 26KPO00144     # 발주번호 기준 FI 생성 (복수 DN 통합)
python create_fi.py --po                # 사용 가능한 발주번호 목록 표시
python create_oc.py SOO-2026-0001        # Order confirmation (생성 후 "이메일 발송?" y/N → 영문 메일)
python create_oc.py SOO-2026-0001 --mail    # 확인 없이 메일 초안
python create_oc.py SOO-2026-0001 --no-mail # 묻지 않고 문서만 (배치용)

# DN 출고기록 자동 입력 (출고리스트 → DN_국내)
python create_dn.py P08                   # 미리보기 → y/N → 워크북에 추가
python create_dn.py P08 --dry-run         # 미리보기 파일만 (워크북 안 건드림)
python create_dn.py P08 --yes             # 확인 없이 추가 (배치용)
python create_dn.py P08 --no-tax-date     # 세금계산서 발행일 전부 공란

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

# 거래처 납기현황 회신 (미출고 조회)
python delivery_status.py 615-81-88675    # 생성 후 "이메일 발송?" y/N 확인 → y면 메일 초안
python delivery_status.py 엔이에스         # 거래처명 부분일치 (후보 여럿이면 목록 출력)
python delivery_status.py --list          # 미출고 잔량이 있는 거래처 목록
python delivery_status.py 615-81-88675 -a # 출고완료 포함 전체
python delivery_status.py 615-81-88675 --mail     # 확인 없이 메일 초안
python delivery_status.py 615-81-88675 --no-mail  # 묻지 않고 문서만 (배치용)

# Industry Code 대사 + Sector 검증
python reconcile_ind.py P03               # Industry code 채움 + Sector 검증
python reconcile_ind.py --sector-only     # Sector 검증만 (월 입력 불필요)
python reconcile_ind.py P03 -v            # 상세 로그

# Dashboard
streamlit run dashboard.py                # Streamlit 대시보드
python dashboard_dist/build_portable_dashboard.py  # 배포 zip → dashboard_dist/NOAH_대시보드_배포.zip

# GUI (문서 7종 + 납기현황) / 사내 배포판 빌드
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

Data entry layer (문서 생성이 아니라 데이터 파일에 써 넣는 유일한 경로):
  create_dn.py → dn_recorder.py (계산, COM 없음) → dn_writer.py (xlwings 쓰기)

DB layer:
  sync_db.py → db_sync.py (Excel→SQLite, upsert+prune, FX 시트 언피벗) → db_schema.py (DDL + v_dn_revenue 뷰)
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
- **병합 셀에는 `rows.autofit()`이 먹지 않는다** (`excel_helpers.py`): 해외 문서 5종(OC·FI·PI·CI·PL)의 품목명 칸은 A:D 병합인데, Excel의 자동 맞춤은 **병합 셀을 측정에서 제외한다**. 105자 품목명에 `autofit()`을 부르면 높이가 15pt → **12.75pt로 오히려 줄면서 1줄로 잘린다** (2026-08-03 실측). 그래서 `autofit_merged_rows()`가 인쇄영역 밖 보조 열(Z)에 같은 텍스트를 넣고 Excel에게 재게 한 뒤 그 높이를 되쓴다. 보조 열 폭은 **문자 단위가 아니라 포인트로** 맞춘다 — 열마다 안쪽 여백이 붙어 A:D(198.00pt)와 같은 문자폭의 단일 열(186.75pt)이 11.25pt 어긋난다. COM 왕복은 배치가 원칙: 높이 되쓰기는 같은 높이 연속 구간당 1회, 행 삽입은 `insert_copied_rows`(일괄 Insert + 타일 Copy, 3회). **병합 셀은 넘친 텍스트를 옆 칸으로 흘리지도 않는다** — 경계에서 그냥 클립된다. OC·FI 헤더의 주소 칸(행별 A:E 병합 + 세로 G:I 병합)이 이 경우라, 값을 쓴 뒤 `layout_address_rows()`로 wrap + 행 높이를 확보한다 (2026-08-05 실측: 74자/81자 주소 클립. `tests/test_doc_layout.py`가 호출 여부 감시).
- **행을 복사해 삽입하면 일부 행이 병합을 잃는다** (`excel_helpers.py`): 48아이템 OC에서 삽입한 41행 중 6행의 A:D 병합이 사라져 품목명이 A열에 갇혀 5~7줄로 흘렀다 (2026-08-03 실측 — 이 결함은 행 높이 교정 이전부터 있었다). 그래서 병합 보장(`ensure_row_merges`)과 높이 교정은 늘 세트고, 5종 생성기는 값 채우기 직후 **복합 헬퍼 `layout_item_rows()` 하나만** 부른다 — 여섯 번째 문서가 한쪽만 부르는 실수를 원천 차단 (`tests/test_doc_layout.py`가 감시). 병합 범위 `ITEM_NAME_MERGED_COLS`는 **excel_helpers 한 곳**에만 둔다 — 특히 CI와 PL은 선적서류라 늘 같이 첨부되어 나란히 읽히므로 줄 높이 규칙이 갈리면 바로 눈에 띈다.
- **History as DB**: `po_history/YYYY/M월/YYYYMMDD_주문번호_고객명.xlsx` — one file per transaction enables duplicate detection without a database.
- **Result Pattern** (`services/result.py`): `DocumentResult` + `GenerationStatus` enum for structured operation outcomes. `history_saved` field tracks history persistence separately from generation success.
- **Output File Safety** (`cli_common.py`): Generated files auto-suffix on collision (`_1`, `_2`, ...) to prevent silent overwrites. Raises `FileExistsError` if 100+ collisions.
- **DB Sync Prune** (`db_sync.py`): Excel→SQLite sync includes prune step — rows deleted from Excel are also deleted from DB. Works even when sheet is completely empty. `--dry-run` connects to real DB and rollbacks for accurate diff simulation.
- **`_row_seq` in PK — 자연키는 유일하지 않다** (`db_schema.py: SYNC_SHEETS`): PO·DN 시트는 같은 문서·같은 라인을 **여러 행으로 나눠 적는다**(분할발주/분할출고). 그래서 PO는 `(PO_ID, Line item, _row_seq)`, DN은 `(DN_ID, SO_ID, Line item, _row_seq)`가 PK다. `_row_seq`를 빼면 upsert가 뒤 행으로 앞 행을 덮어써 **매출·수량이 조용히 사라진다** (2026-07-30 실측: DN에서 18,656,000원 누락). 새 시트를 `SYNC_SHEETS`에 추가할 때 자연키 유일성을 실데이터로 확인할 것 — `SELECT COUNT(*)` vs `COUNT(DISTINCT 자연키)`.
- **PK Migration Logging** (`db_sync.py` + `sync_db.py`): PK 정의를 바꾸면 `migrate_pk_if_changed`가 기존 테이블을 `{table}_bak`으로 백업하고 재생성한다. 재적재는 전 행이 '신규'로 잡히므로 `SheetSyncResult.pk_migrated`를 보고 `_sync_log`에 **'재적재' 1건만** 남긴다 (수천 건의 가짜 '신규'가 변경 이력을 덮지 않게 — 재키잉 억제와 같은 취지).
- **Dashboard Error Visibility** (`dashboard.py`): Loader failures collected in `session_state` and displayed as `st.warning()` banner, distinguishing "no data" from "query failure".
- **Snapshot Engine** (`snapshot.py`): Monthly close → `ob_snapshot` freezes Ending, subsequent retroactive changes auto-detected as Variance. Sequential close enforced. `variance_amount`는 **환율 재평가분 + 소급 변경분 합산** (AX Order Book처럼 조정분 단일 컬럼).
- **Single Source for DN Revenue** (`db_schema.py: v_dn_revenue`): 매출 귀속월·선적월 환율 재환산·환율차를 뷰 한 곳에서 정의. Order Book 계열 6개 소비처(sql 4종 + `snapshot.py` + `dashboard.py`)가 모두 이 뷰를 읽는다.
- **`po_generator/__init__.py`는 비워 둔다**: 하위 모듈을 재export하면 `from po_generator.config import DB_FILE` 한 줄에도 패키지 `__init__`이 먼저 돌아 **pandas·openpyxl·xlwings(→COM)가 통째로 딸려온다**. 대시보드는 `config`·`db_schema`만 쓰고 둘 다 표준 라이브러리만 의존하므로, 비워 둔 덕에 배포판에서 Excel 라이브러리가 빠진다. `tests/test_dashboard_dist.py`가 이 세 파일의 최상위 import를 감시한다.
- **OneDrive로 공유하는 SQLite는 원본을 열지 않는다** (`dashboard.py: _db_snapshot`): DB는 `journal_mode=wal`이라 최신 커밋이 `-wal` 사이드카에 먼저 들어가는데, OneDrive는 본체와 사이드카를 **각각 따로** 동기화한다. 원본을 직접 열면 (1) 커밋이 빠진 상태를 보거나 (2) 읽기 잠금 때문에 OneDrive가 파일을 교체 못 해 **"충돌된 사본"** 이 생긴다 — 둘 다 "사람마다 숫자가 다르다"로 뒤늦게 드러난다. 그래서 백업 API로 일관된 사본을 떠서 그것만 읽고, 캐시 키가 `(mtime, size)`라 원본이 갱신되면 자동으로 다시 뜬다. 발행 쪽은 `sync_db.py: checkpoint_wal()`이 `PRAGMA wal_checkpoint(TRUNCATE)`로 `-wal`을 비워 **단일 파일로 완결**시킨다.

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
| `create_po.bat` | **대화형 메뉴 런처 (이름과 달리 PO 전용이 아니다)** — 문서 7종 + DB Sync·마감·대사·대시보드·납기현황을 번호로 고르는 메뉴다. 거래명세표처럼 **하위 메뉴가 있는 항목도 있다**(단건/월합/하루치 묶음/목록 묶음). CLI에 옵션을 추가하면 여기와 `noah_gui.py`에도 넣어야 사용자 눈에 보인다 — `create_ts.py`만 고치고 "CLI에 반영했다"고 한 적이 있다 (2026-08-07, `tasks/lessons.md`) |
| `noah_gui.py` | tkinter GUI — 문서 7종 + 납기현황. 기존 `create_*.py`·`delivery_status.py`를 **자식 프로세스로 실행**하고 stdout을 로그 위젯에 흘린다(CLI 무수정, COM 격리). 데이터 파일 지정 마법사 포함. **DOC_TYPES에 항목을 추가하면 `build_portable_gui.APP_FILES`에도 넣어야 한다** — 안 그러면 배포판에서 그 버튼만 조용히 실패한다 (빌드 `verify()`와 `tests/test_noah_gui.py`가 대조) |
| `build_common.py` | 배포판 빌드 공통부 — 런타임 내려받기·핀 3자 대조·트리밍·BUILD_INFO·zip. 문서생성기와 대시보드 두 빌더가 공유한다 (복사하면 갈라지고, 그 갈라짐은 "한쪽만 낡은 pandas로 나간다"로 늦게 드러난다 — `mail_cli.py`와 같은 이유). **최상위는 상수·함수 정의만** |
| `cli_dist/build_portable_gui.py` | 문서생성기 배포판 — `build_common` 위에 이 배포판만의 것을 얹는다: `APP_FILES`(담을 것) · 런처/설치 스크립트 · `verify()`(`DOC_TYPES` 대조 + CLI 전수 `--help` 스모크). **모듈 최상위는 상수·함수 정의만** (테스트가 경로로 로드하며, 예외는 `build_common`을 찾는 sys.path 한 줄뿐) |
| `dashboard_dist/build_portable_dashboard.py` | 대시보드 배포판 — `dashboard.py`를 **무수정으로** 담고 `po_generator/`·`sql/`을 동봉한다. 예전 `build_dist.py`는 import를 문자열 치환해 standalone 파일을 만들었는데, import 한 줄이 바뀌자 패턴이 안 맞아 빌드가 죽었고 배포본이 4개월 낡았다 — 그래서 재작성을 아예 없앴다. `verify()`가 streamlit을 **실제로 띄워** 실제 DB 사본으로 페이지를 받아 본다(import만으로는 트리밍 사고를 못 잡는다). pyarrow는 flight/parquet/dataset/substrait를 잘라내되 `arrow_compute`는 남긴다 (streamlit이 로드한다 — 실측) |
| `noah_config.ini` | 배포판 경로 설정 (git-ignored). GUI 마법사가 생성. `user_settings.py`가 있으면 그쪽이 우선 |
| `po_generator/mailer.py` | 고객 메일 발송 — 고객 마스터에서 수신자 조회, xlsx→PDF 변환, 2가지 백엔드(Outlook COM / `.eml` 초안). **첨부는 여러 장일 수 있다** — `as_paths()`가 단건/복수를 흡수하고 `export_pdfs()`가 Excel **1회 기동**으로 N장을 변환한다(8장 실측 24초; 장마다 띄우면 그 두 배). **조인키가 국내/해외로 갈린다**: `find_recipient()`는 사업자번호로 `Customer_국내`, `find_recipient_overseas()`는 고객코드로 `Customer_해외` — 둘 다 `_build_recipient()` 하나를 공유한다. **`auto` = 초안은 `.eml`(사용자 기본 메일 앱 — 새 Outlook 포함), 즉시 발송(--send)만 COM** — COM 초안은 항상 클래식 Outlook 창을 띄우므로 초안에 쓰지 않는다. `create_document_mail()`이 일반형이고 `create_ts_mail()`은 TS 상수를 넘기는 래퍼 — 제목/본문 템플릿·첨부형식·고정 CC·HTML 본문이 전부 인자 |
| `po_generator/mail_cli.py` | 메일 CLI 공통 배선 — `MailMode`/`MailOptions`/`resolve_mail_mode`/`add_mail_arguments`/`prepare_mail_options`/`confirm_recipient`(수신자 조회→표시→y/N 관문)/`collect_customer_po`/`format_mail_date`. `create_ts.py`·`delivery_status.py`·`create_oc.py`가 공유(복사하면 갈라지고, 그 갈라짐이 고객 발송 경로에서 터진다). 어느 마스터를 읽을지는 `MailOptions.loader`/`sheet_label` **짝**으로 주입 — 기본은 국내, OC만 해외. **`loader` 기본값에 함수를 박지 말 것**: 클래스 정의 시점에 굳어 모듈 속성 교체(테스트 monkeypatch)가 무시된다 |
| `docs/ARCHITECTURE.md` | Detailed system design and data flow diagrams |
| `docs/DATA_STRUCTURE_DESIGN.md` | Excel schema (8 sheets), Power Query setup |
| `docs/POWER_QUERY.md` | Power Query 수식, Power Pivot 관계 — 데이터 소스 구조 이해 시 참고 |
| `docs/ERP_CONCEPTS.md` | ERP 개념 학습 노트 — SO-PO 라인 매칭, BOM/Kit, 정규화 수준 비교 |
| `docs/CHANGELOG.md` | 버전별 변경 이력 |
| `docs/TEMPLATE_MAPPINGS.md` | Excel 템플릿 셀 매핑 — 템플릿/generator 수정 시 참고 |
| `docs/매입대사_가이드.md` | 운영자 관점 시각 가이드 — 데이터 흐름·합계 블록 해석·월별 액션 |
| `po_generator/snapshot.py` | SnapshotEngine — 월별 마감, Variance 추적 |
| `po_generator/db_schema.py` | SQLite DDL, snapshot tables (`ob_snapshot`, `ob_snapshot_meta`), `_sync_runs` + `_sync_log` v2 (record당 1행 + JSON, actor/host/sync_id/snapshot 포함), `fx(currency, ym, rate)`, **`v_dn_revenue`(라인 매출 산식) + `v_dn_by_month`(인식 필터+월별 합산) 뷰 = DN 매출의 단일 정의**. 뷰는 `sync_db.py`가 매번 DROP+CREATE로 갱신. `SYNC_LOG_CHANGE_TYPES`·`KNOWN_SYNC_SHEETS`도 여기가 소유 |
| `migrate_sync_log.py` | `sync_log.csv` → `_sync_log_legacy_v1` 전용 테이블 1회성 적재 (구 v1 형식; 운영 v2 `_sync_log`와 분리, 이후 `migrate_sync_log_v2.py`로 변환) |
| `migrate_sync_log_v2.py` | `_sync_log` v1 → v2 (record 단위 + JSON 압축, _sync_runs 메타 분리, snapshot 추가) |
| `sql/order_book.sql` | 이벤트 기반 Order Book SQL (Input/Output 이벤트 월만 행 생성, 재귀 CTE 없음). Output/Variance는 `v_dn_revenue` 뷰에서 온다 — **DN 산식을 여기서 다시 쓰지 말 것** (같은 산식이 sql 4종 + snapshot.py + dashboard.py 6곳에 복제돼 있던 걸 뷰로 모았다) |
| `sql/order_book_snapshot.sql` | 스냅샷 기반 Order Book SQL (마감 고정 + Variance) |
| `sql/order_book_variance.sql` | Variance 변동이유 분석 SQL (환율차이/판매가변경/수량변경/반올림 자동 분류, 납기변경 제외) |
| `dashboard.py` | Streamlit 대시보드 (9페이지: 오늘의현황/수주출고/제품/섹터/고객/발주커버리지/수익성/Order Book/동기화로그, PO미등록감지, PO확정지연, EXW미출고, 납기현황(DN qty매칭+PO EXW보충), 납기캘린더(선적예정 포함), 해외선적(Incoterms/운송방식별), 세금계산서미발행, Order Book 3탭, `_sync_log` 변경이력 조회) |
| `reconcile_po.py` | PO 매입대사 — 공장 출고(Delivery) vs 회계 GRN 금액 비교. 출력: `대사결과_{period}.xlsx` (6시트), `AX_PO_매핑_{period}.xlsx` (Delivery+AX PO) |
| `reconcile_so.py` | SO 매출대사 — AX ERP 매출 vs NOAH DN 매출 비교 (국내=출고일, 해외=선적일 기준 월 필터 + FX 환율차이 자동 판별). 출력: `대사결과_SO_{period}.xlsx` (3시트: 대사/상세/범례) |
| `create_dn.py` + `po_generator/dn_recorder.py` / `dn_writer.py` | DN 출고기록 자동 입력 — 공장 출고리스트를 읽어 `DN_국내`에 추가. **문서를 만드는 게 아니라 마스터 워크북에 써 넣는 유일한 기능**이라 미리보기 → y/N → 백업 → 쓰기 순서다. 계산(`dn_recorder`)과 쓰기(`dn_writer`)를 나눈 건 계산을 COM 없이 테스트하기 위함. 아래 Business Rules 참조 |
| `delivery_status.py` | 거래처 납기현황 회신 — 사업자번호로 `SO_국내` 미출고 조회 → `generated_ds/납기현황_*.xlsx` (납기현황/상세 2시트) → 메일 발송. **출고 여부는 시트 `Status`가 아니라 `DN_국내` 출고수량으로 직접 계산** (아래 Business Rules 참조) |
| `reconcile_ind.py` | Industry Code 대사 — (1) Orderbook 빈 Industry code를 PO→SO 매핑으로 채움 → `ind_code_결과_{period}.xlsx`, (2) SO Sector vs 마스터 Category 교차 검증 → `sector_검증.xlsx`. `--sector-only`로 검증만 실행 가능 |

## Business Rules

- Order numbers: `ND-*` = domestic, `NO-*` = overseas
- DN numbers: `DND-*` = domestic, `DNO-*` = overseas
- Validation blocks generation unless `--force`: missing required fields, ICO Unit ≤ 0, past delivery date
- Warnings (non-blocking): delivery within 7 days, duplicate order in history
- **출고 여부는 `SO_국내.Status`를 읽지 말고 `DN_국내` 출고수량으로 계산한다.** `Status`는 수기 입력이 아니라
  `XLOOKUP(..., SO_통합[출고완료])` 캐시라서 파워쿼리 새로고침 전에는 실제 출고와 어긋난다
  (2026-07-29 실측 29행). 판정식은 파워쿼리 `SO_통합[출고완료]`와 동일 —
  `출고수량 없음=미출고 / 주문-출고>0=부분 출고 / 출고일 없음=공장 출고 / 그 외=출고 완료`.
  시트 `Status`는 `Cancelled`/`Hold` 제외에만 쓴다 (취소 건은 DN이 영영 안 생겨 수량으로 구분 불가).
  `dashboard.py: load_so()`, `delivery_status.py`가 같은 규칙을 쓴다
- 고객 메일 발송(거래명세표·납기현황 공통): 수신자(To)는 `Customer_국내.사업자번호` 조인으로 결정.
  기본은 **수신자를 보여준 뒤 y/N 확인**(비대화형 실행은 자동 OFF — 배치가 프롬프트에서 멈추지 않게).
  메일 실패는 문서 생성 성공을 뒤엎지 않는다
  - 거래명세표: `DN_국내.Business registration number`로 조인, 고정 참조는 `TS_MAIL_CC`, 첨부 기본 PDF.
    고객이 섞인 `--merge` 문서는 발송 차단
  - **묶음 메일(`--date`/`--one-mail`)의 단위는 문서가 아니라 거래처다.** 하루에 한 거래처로
    여러 PO가 나가면(2026-08-06 씨앤케이 8건 실측) 문서는 DN별 1장 그대로 두고 메일만 한 통에
    첨부 N개로 묶는다. 묶는 키는 **정규화한 사업자번호** — 이름으로 묶으면 '(주)' 표기 차이로
    같은 거래처가 갈라지고, 사업자번호가 빈 건은 서로 묶지 않는다(모르는 것끼리 합치면 남의
    명세표가 붙는다). y/N 확인도 문서마다가 아니라 **메일마다** 한 번.
    `--merge`(문서를 합침)와 `--one-mail`(메일만 합침)은 동시 지정 불가
  - **한 문서에 두 거래처가 실리면 메일로 내보내지 않는다** (`create_ts.foreign_biz_numbers`).
    DN 번호를 재사용하면 한 DN에 두 SO가 들어가는데(실측: `DND-2026-0748` 씨앤케이+오토밸브,
    `DND-2026-0328` 코콘+한일전자) 그 문서엔 남의 품목·단가가 찍힌다. 묶음 실행에서는 그
    문서만 첨부에서 빼고 나머지는 그대로 보낸다 — 시트 수정은 사람 몫이라 경고로 남긴다
  - 납기현황: 조회 기준인 사업자번호를 그대로 사용(항상 단일 거래처라 섞임 없음), 고정 참조는 `DS_MAIL_CC`,
    첨부 기본 **xlsx**(고객이 정렬·가공해 보는 표라 원본이 쓸모 있다). 본문에 납기 표를 HTML로 싣는다
- **OC 메일은 해외 전용이고, 조인키가 사업자번호가 아니라 고객코드다.** `SO_해외`의
  `Business registration number` 컬럼에는 실제로 `C-0054` 같은 **고객코드**가 들어 있고
  (이름만 국내 시트와 같다), 이것이 `Customer_해외.C-code by 해외`와 맞물린다.
  주의: `Customer_해외`의 `고객코드` 컬럼은 **다른 값**(AX 번호)이라
  `COLUMN_ALIASES['customer_code']`에 별칭으로 넣으면 안 된다 — 앞 별칭이 없는 시트에서
  조용히 엉뚱한 컬럼으로 풀려 전 건이 미매칭된다. 본문은 **영문**이며 인사말(Dear
  {customer}) + 발주번호 확인 한 줄 + 자동발송 안내로 짧게 간다 — 안내문·서명 없음
  (2026-08-05 결정. 제목 `[Rotork Controls Korea] Order Confirmation - Your PO:
  {customer_po}`가 발신 조직을 밝힌다). 서명을 넣는 오버라이드라면 `{supplier}`(한글)가
  아니라 `{supplier_en}`. 첨부 기본 PDF, 고정 참조는 `OC_MAIL_CC`.
  나머지 규약(수신자 확인 후 y/N, 비대화형 자동 OFF, 메일 실패가 문서 생성을 뒤엎지 않음)은
  거래명세표와 동일하다
- **Order Book Output = 매출 인식 기준**이다 (출고 기준이 아니다) — `AX_매출대사`와 동일 산식.
  국내 = 세금계산서 발행월(미발행이면 **미인식 = Backlog 잔류**), 해외 = 선적월 + 선적월 환율 재환산.
  환율 재평가분은 `Value_Variance_amount`로 빠져 **Ending은 재환산 도입 전과 동일** —
  "Ending ≠ 0 = SO-DN 금액 불일치" 진단이 환율 노이즈에 오염되지 않는다.
  결과: `Period` 필터 → `SUM(Value_Output_amount)` = `AX_매출대사` 같은 월 합계.
  **Backlog는 "미출고"가 아니라 "미인식" 물량** — 물류상 미출고는 대시보드 `납기현황`/`EXW미출고`.
  N/A·선수금 폴백 사다리, 환율 결측 폴백 등 정확한 산식은 `v_dn_revenue`/`v_dn_by_month` 뷰
  (db_schema.py)와 `docs/POWER_QUERY.md`(Order_Book)가 소유한다 — 산문 사본을 늘리지 말 것
- **DN 출고기록 자동 입력의 라인/수량은 `PO_국내.Status`가 정한다** (`create_dn.py`).
  출고리스트는 "어느 주문이 언제 나갔는지"만 알려주고 **어느 라인이 몇 개인지는 없다**.
  `Status = Invoiced P{XX}`가 공장이 그 달에 계산서를 끊은 라인 = 출고된 라인이다.
      pending = 이 SO의 PO 라인 중 Status ∉ {Invoiced P01..P{XX}, Cancelled}
      pending 없음 → `SO_국내` 전 라인의 잔량 (= Item qty − 그 출고일 **이전** DN 누계)
      pending 있음 → `PO_국내` Invoiced P{XX} 라인의 Item qty 그대로
  **PO만 보면 안 된다** — PO는 부속을 1라인에 합쳐 적는다(`SA09X-MA + ADAPTER`,
  `NA015 (...) / 부싱가공`). SO 2라인이 PO 1라인이라 PO 기준으로만 뽑으면 부속이 통째로
  빠진다(2026-03~08 실측 6건, 최대 360,000원). 반대로 `Cancelled`를 안 빼면 취소분이
  출고로 잡힌다(같은 기간 2건). 출고일은 출고리스트 `납품완료`, DN_ID는 출고리스트 1행당 1개.
- **DN 라인은 `SO_국내`에 살아 있는 라인만이다** — PO를 그대로 옮기면 안 된다.
  **DN은 매출 장부**라서 담을 수 있는 건 실제로 판 것뿐인데, PO에는 그렇지 않은 것이 섞인다:
  - **SO에 없는 PO 라인** = 매입만 발생 (`dn_recorder.sales_only`).
    외주 가공비 같은 것 — 2026-05 `SOD-2026-0188` L4 'De-cluch Gear Box Bushing 하부 가공'
    450,000원 (PO 4라인 / SO 3라인)
  - **SO Status가 `Cancelled`/`Hold`인 라인** (`config.EXCLUDED_SO_STATUSES`).
    판매가 취소돼도 공장은 이미 만든 것을 계산서로 넘기기도 한다 — 2026-07
    `SOD-2026-0364` L5는 SO가 `Cancelled`인데 PO L5 'IP66 TEST 시료 값'은
    `Invoiced P07` 1,000,000원. `SO_국내.Status`는 파워쿼리 캐시라 출고 여부 판정에는
    쓰면 안 되지만 **취소/보류만은 이 컬럼으로만 알 수 있다**(취소 건은 DN이 영영 안 생겨
    수량으로 구분 불가). 같은 상수를 `delivery_status.py`도 쓴다

  둘 다 실데이터 불변식이다: DN 1,431행 중 SO에 없는 라인 **0건**,
  SO Cancelled 44라인 중 DN에 있는 것 **0건**.
  단, **금액 대조에서는 두 라인 다 살려 둔다** — 공장은 그것까지 계산서를 끊으므로
  출고리스트 `계산서금액`에 포함돼 있다(0188 건: 14,610,800원에 450,000원이 들어 있다).
  빼는 건 DN에 쓸 라인을 고를 때뿐. 이걸 안 빼면 `Unit Price` XLOOKUP이 SO에서 못 찾아
  단가 0원짜리 행이 DN에 남는다
  자동 입력은 **자기검증 통과분만** — 단일 날짜는 `Σ PO Total ICO == 계산서금액`.
  **같은 SO가 여러 날 나갔으면 PO 행의 ICO를 날짜별 계산서금액에 배정한다**
  (`split_by_amount`) — `PO_국내`는 분할출고마다 행을 따로 두므로 대개 정확히 나뉜다
  (2026-03~07 실측 12건 중 10건 유일 배정, 모호 0건).
  **안 나뉘는 이유를 구분해야 한다** (`_solve_split`의 사유):
  - `no_split`(나눌 방법이 아예 없음) + 합계가 PO ICO와 일치 → **출고는 한 번이고
    나머지는 단가 정정 행**이다. 첫 날 한 건으로 기록하고 뒤 행은 건너뛴다.
    2026-05 `SOD-2026-0306`: 5/11 계산서 9,956,592 → 5/13에 `L260441-1R`로 103,616인데,
    그 행은 `AMOUNT`가 10,060,208로 바뀐 **차액**이다 (출고는 5/11 8개 한 번뿐)
  - `ambiguous`(결과가 갈림) → 출고는 여러 번인데 어느 쪽인지 모른다. 한 건으로 뭉치면
    안 되고 '확인 필요'로 뺀다
  - 계산서금액이 **음수**인 행이 섞이면(반품) 무조건 '확인 필요' — 출고/반품/재출고를
    어떻게 적을지는 사람이 정한다 (2026-03 `SOD-2026-0280`)

  `JOB NO`의 `-1R` 접미사로는 못 가른다 — JOB NO는 주문 단위라 정상 분할출고에도 붙는다.
  세금계산서 발행일은 출고일과 동일하되 **월합 거래처**(기존 DN Remarks에서 유도)는 공란 +
  Remarks 상속 — 그 거래처는 출고 시점에 계산서를 끊지 않는다.
  멱등 판정은 `(SO_ID, 출고일)` **와** "라인·수량 조합이 **정확히 같은** 출고가 다른 날짜로
  이미 기록됨" 둘 다 본다 (사람이 실제 출고일을 알고 하루 이틀 다른 날로 적는 경우가 있다 —
  2026-03 `SOD-2026-0156`). 뒤쪽을 '라인별 누계 ≥ 필요수량'으로 보면 안 된다 — 분할출고에서
  **뒤 회차가 앞 회차 수량에 가려 통째로 사라진다** (2026-07 `SOD-2026-0713`: 7/23 L1 1개를
  8/6의 L1 2개가 덮었다).
  뒤쪽 판정은 **실행 전 시트 상태만** 봐야 한다 — 실행 중 추가분까지 보면 한 SO가 여러 날
  나갈 때 앞 출고 때문에 뒤 출고가 통째로 사라진다 (2026-05 `SOD-2026-0188`에서 실제로 났다).
  출고리스트는 같은 주문·같은 날을 두 줄로 적기도 해서(`SOD-2026-0467`) 이벤트를
  `(SO_ID, 출고일)`로 합친 뒤 순회한다
- **`NOAH_SO_PO_DN.xlsx`는 openpyxl로 저장하면 안 된다** (`dn_writer.py`).
  피벗 3개·파워쿼리 연결 14개·쿼리테이블이 들어 있고 openpyxl은 그것들을 읽지도 쓰지도
  못해서 한 번 저장하면 통째로 사라진다. `reconcile_*.py`가 openpyxl을 쓰는 건 전부
  **새 결과 파일**을 만들 때뿐이다. 표에 행을 더할 때는 **마지막 데이터 행을 타일 복사**한다 —
  `DN_국내`에서 Excel이 자동으로 채우는 계산 열은 6개뿐이고 `Item`(XLOOKUP)·`Unit Price`·
  `AX Project no`(배열 수식)는 수식이 있어도 `calculatedColumnFormula`가 없어 행을 늘려도
  비어 있다. 복사하면 상대참조(`B1432` → `B1433`)가 따라오므로 그 위에 값 열만 덮어쓴다
- 납기현황 요약 행은 **주문 × 요청납기 × 공장출고일** 단위다. 한 주문 안에서 두 날짜 중 하나라도
  다르면 행을 나눈다 — 대표값 하나로 접으면 나머지 납기 약속이 회신에서 사라진다.
  나뉜 행이 비고만으로 구분되지 않을 때만 `(품목: ...)`를 덧붙인다

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
