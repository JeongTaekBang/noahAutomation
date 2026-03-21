# NOAH Document Auto-Generator

Sales Office에서 Intercompany Factory로 보내는 업무 문서를 자동 생성하고, 비즈니스 KPI를 대시보드로 모니터링합니다.

```
NOAH_SO_PO_DN.xlsx (데이터 소스)
       │
       ├── 국내 시트 ─→ PO, 거래명세표
       │
       ├── 해외 시트 ─→ PO, PI, FI, OC, CI, PL
       │
       └── DB sync ──→ SQLite ──→ Streamlit 대시보드
```

## 지원 문서

| 문서 | 용도 | CLI |
|------|------|-----|
| **PO** (Purchase Order) | 사내 발주서 | `create_po.py` |
| **거래명세표** (Transaction Statement) | 국내 납품/선수금 명세 | `create_ts.py` |
| **PI** (Proforma Invoice) | 해외 견적서 | `create_pi.py` |
| **FI** (Final Invoice) | 해외 대금 청구서 | `create_fi.py` |
| **OC** (Order Confirmation) | 해외 주문 확인서 | `create_oc.py` |
| **CI** (Commercial Invoice) | 해외 상업 송장 | `create_ci.py` |
| **PL** (Packing List) | 해외 포장 명세서 | `create_pl.py` |

## 추가 기능

| 기능 | 용도 | CLI |
|------|------|-----|
| **DB Sync** | Excel → SQLite 동기화 | `sync_db.py` |
| **월마감** | 월별 스냅샷 & Variance 추적 | `close_period.py` |
| **대시보드** | Streamlit 비즈니스 KPI 모니터링 | `dashboard.py` |

---

## 사용 방법

### 1단계: 주문 정보 입력

`NOAH_SO_PO_DN.xlsx` 파일에 주문 정보를 입력합니다.

| 시트 | 용도 |
|------|------|
| SO_국내 / PO_국내 / DN_국내 | 국내 고객 주문 |
| SO_해외 / PO_해외 / DN_해외 | 해외 고객 주문 |

**필수 입력 항목:**
- `Order no.` - 주문번호 (예: ND-0001, NO-0001)
- `Customer name` - 고객명
- `Customer PO` - 고객 발주번호
- `Item qty` - 수량
- `Model` - 모델명
- `ICO Unit` - 단가

### 2단계: 문서 생성

`create_po.bat` 파일을 더블클릭하면 대화형 메뉴가 나옵니다.

```
========================================
   NOAH Document Generator
========================================

  [국내]
  [1] 발주서 생성 (PO)
  [2] 거래명세표 생성 (DN/선수금)

  [해외]
  [3] Proforma Invoice 생성 (PI)
  [4] Final Invoice 생성 (대금 청구)
  [5] Order Confirmation 생성 (OC)
  [6] Commercial Invoice 생성 (CI)
  [7] Packing List 생성 (PL)

  [기타]
  [8] 발주 이력 조회
  [9] 발주 이력 Excel 내보내기
  [0] 종료
```

또는 명령 프롬프트에서 직접 실행:

```bash
# PO (발주서)
python create_po.py ND-0001               # 단일 발주
python create_po.py ND-0001 ND-0002       # 여러 건 동시 생성
python create_po.py ND-0001 --force       # 강제 생성 (검증 오류 무시)
python create_po.py --history             # 이력 조회 (현재 월)
python create_po.py --history --export    # 이력을 Excel로 내보내기

# 거래명세표
python create_ts.py DND-2026-0001                        # 단건
python create_ts.py DND-2026-0001 DND-2026-0002 --merge  # 월합 (한 장)
python create_ts.py --interactive --merge                 # 대화형 모드

# 해외 문서
python create_pi.py SOO-2026-0001         # Proforma Invoice
python create_fi.py DNO-2026-0001         # Final Invoice
python create_oc.py SOO-2026-0001         # Order Confirmation
python create_ci.py DNO-2026-0001         # Commercial Invoice
python create_pl.py DNO-2026-0001         # Packing List
```

### 3단계: 결과 확인

생성된 문서는 각 폴더에 저장됩니다.

```
generated_po/   ← 발주서
generated_ts/   ← 거래명세표
generated_pi/   ← Proforma Invoice
generated_fi/   ← Final Invoice
generated_oc/   ← Order Confirmation
generated_ci/   ← Commercial Invoice
generated_pl/   ← Packing List
```

---

## DB Sync & 월마감

### Excel → SQLite 동기화

```bash
python sync_db.py    # NOAH_SO_PO_DN.xlsx → noah.db
```

### 월마감 (Period Close)

```bash
python close_period.py 2026-01          # 1월 마감 (스냅샷 생성)
python close_period.py --undo 2026-01   # 마감 취소
python close_period.py --list           # 스냅샷 이력 조회
python close_period.py --status         # 현재 마감 상태
```

마감 스냅샷은 Order Book의 Ending 값을 고정하고, 이후 소급 변경은 Variance로 자동 추적됩니다.

---

## Streamlit 대시보드

```bash
streamlit run dashboard.py
```

8개 페이지로 구성된 비즈니스 KPI 모니터링 대시보드:

| 페이지 | 내용 |
|--------|------|
| 오늘의 현황 | 주요 지표 요약, 납기 캘린더, 세금계산서 미발행 |
| 수주/출고 | 수주·출고 추이, 월별 비교 |
| 제품 분석 | 모델별 실적, 제품 믹스 |
| 섹터 분석 | 산업별 매출 분석 |
| 고객 분석 | 고객별 매출, 집중도 |
| 발주 커버리지 | 발주 대비 수주 커버율 |
| 수익성 | 마진, 수익률 분석 |
| Order Book | Executive/Risk/Conversion 3탭, 리드타임 분석 |

---

## 자동 검증 항목

| 검증 항목 | 조건 | 결과 |
|----------|------|------|
| 필수 필드 | 비어있음 | 오류 |
| ICO Unit | 0 또는 음수 | 오류 |
| 납기일 | 과거 날짜 | 오류 |
| 납기일 | 7일 이내 | 경고 |
| 중복 발주 | 이미 생성된 주문번호 | 경고 (확인 후 진행 가능) |

오류 시 `--force` 옵션으로 강제 생성 가능.

---

## 파일 구조

```
noahAutomation/
├── create_po.bat               ← 더블클릭 실행 (대화형 메뉴)
├── create_po.py                ← PO CLI
├── create_ts.py                ← 거래명세표 CLI
├── create_pi.py                ← Proforma Invoice CLI
├── create_fi.py                ← Final Invoice CLI
├── create_oc.py                ← Order Confirmation CLI
├── create_ci.py                ← Commercial Invoice CLI
├── create_pl.py                ← Packing List CLI
├── sync_db.py                  ← Excel → SQLite 동기화
├── close_period.py             ← 월마감 / 스냅샷
├── dashboard.py                ← Streamlit 대시보드
│
├── po_generator/               ← 핵심 패키지
│   ├── config.py               ← 상수, 경로, 시트명, 컬럼 별칭
│   ├── utils.py                ← 데이터 로드, 값 추출
│   ├── validators.py           ← 필드 검증
│   ├── excel_generator.py      ← PO 생성 (openpyxl)
│   ├── ts_generator.py         ← 거래명세표 (xlwings)
│   ├── pi_generator.py         ← PI (xlwings)
│   ├── fi_generator.py         ← FI (xlwings)
│   ├── oc_generator.py         ← OC (xlwings)
│   ├── ci_generator.py         ← CI (xlwings)
│   ├── pl_generator.py         ← PL (xlwings)
│   ├── template_engine.py      ← 다중 아이템 행 복제, SUM 수식 조정
│   ├── db_sync.py              ← Excel→SQLite 동기화 엔진
│   ├── db_schema.py            ← SQLite DDL, 스냅샷 테이블
│   ├── snapshot.py             ← 월마감 스냅샷 엔진
│   └── services/
│       ├── document_service.py ← 문서 생성 오케스트레이터
│       ├── finder_service.py   ← 주문 검색 서비스
│       └── result.py           ← DocumentResult 패턴
│
├── templates/                  ← 문서 템플릿 (Excel)
│   ├── purchase_order.xlsx
│   ├── transaction_statement.xlsx
│   ├── proforma_invoice.xlsx
│   ├── final_invoice.xlsx
│   ├── order_confirmation.xlsx
│   ├── commercial_invoice.xlsx
│   └── packing_list.xlsx
│
├── sql/                        ← Order Book SQL
│   ├── order_book.sql          ← 이벤트 기반
│   └── order_book_snapshot.sql ← 스냅샷 기반
│
├── docs/                       ← 문서
│   ├── ARCHITECTURE.md
│   ├── CHANGELOG.md
│   ├── DASHBOARD_GUIDE.md
│   ├── TEMPLATE_MAPPINGS.md
│   └── ...
│
├── generated_*/                ← 생성된 문서 (git-ignored)
└── po_history/                 ← 발주 이력 (월별, git-ignored)
```

---

## 다중 아이템 주문

동일한 `Order no.`로 여러 행을 입력하면 하나의 문서에 여러 아이템이 포함됩니다.

| Order no. | Customer name | Item name | Item qty | ICO Unit |
|---------------|---------------|-----------|----------|----------|
| ND-0001 | ABC전자 | NA-100 | 2 | 500,000 |
| ND-0001 | ABC전자 | NA-200 | 1 | 750,000 |

→ `ND-0001` 발주서에 2개 아이템이 포함됨

---

## 설치

### 1. Miniconda 설치
- https://docs.conda.io/en/latest/miniconda.html 에서 Windows 64-bit 버전 설치

### 2. 환경 생성
```bash
conda create -n po-automate python=3.11
conda activate po-automate
pip install -r requirements.txt
```

### 3. 설정 파일 생성
- `user_settings.example.py` → `user_settings.py`로 복사 후 본인 경로 수정
- `local_config.example.bat` → `local_config.bat`으로 복사 후 본인 Python 경로 수정

---

## 문제 해결

### "주문번호를 찾을 수 없습니다"
- `NOAH_SO_PO_DN.xlsx`에 해당 주문번호가 있는지 확인
- 주문번호 앞뒤 공백 확인
- 국내/해외 시트 모두 확인

### "ICO Unit이 0입니다"
- 단가가 입력되었는지 확인
- 숫자 형식인지 확인 (텍스트로 입력되면 인식 안 됨)

### "이미 발주된 건입니다"
- `po_history/` 폴더에서 이전 발주 기록 확인
- 현재 월 이력: `python create_po.py --history`
- 재발주가 필요하면 Y 입력하여 진행
