# Data Pipeline 확장 구상

ERP 데이터를 활용한 대시보드 확장 로드맵.

---

## 현재 구조

```
Excel (NOAH_SO_PO_DN.xlsx)
    ↓ sync_db.py
SQLite (noah_data.db)
    ↓ dashboard.py
Streamlit 대시보드
```

- 데이터 소스: 수동 관리 Excel (SO/PO/DN 8개 시트)
- 동기화: `sync_db.py` — Excel → SQLite 전체 덮어쓰기
- 대시보드: 8페이지, standalone 배포 가능

---

## 확장 구상

### 데이터 소스 추가

ERP(D365 F&O, AX2009)에 직접 DB 접근은 불가하므로, **ERP에서 Excel/CSV로 다운로드 → SQLite 동기화** 방식을 사용한다.

```
ERP (D365 / AX2009)
    ↓ 수동 다운로드 (Excel/CSV)
로컬 파일
    ↓ sync_erp.py (신규)
SQLite (noah_data.db 또는 별도 DB)
    ↓ dashboard.py
Streamlit 대시보드
```

### 대상 데이터

| 데이터 | ERP 소스 | 활용 |
|--------|----------|------|
| **Order Book** | D365 SO/PO/DN export | 현재 Excel 대체 — 동일 대시보드 |
| **Trial Balance (TB)** | D365 F&O GL export | P&L, BS, Cash Flow 대시보드 |
| **AP/AR Aging** | D365 F&O 미수/미지급 | 채권·채무 현황 |
| **Inventory** | D365 F&O 재고 | 재고 회전율, 과잉재고 분석 |

### sync_erp.py 구상

```python
# 기존 sync_db.py와 동일한 패턴
# 1. ERP에서 다운로드한 Excel/CSV 파일 읽기
# 2. 컬럼 매핑 (ERP 컬럼명 → 내부 표준명)
# 3. SQLite 테이블에 UPSERT
# 4. 동기화 메타 기록 (_sync_meta)
```

기존 `db_schema.py`의 SheetConfig 패턴을 그대로 재사용:
- PK 정의, 컬럼 자동 감지, 변경 감지
- 기존 SO/PO/DN 데이터와 같은 DB에 공존 가능

### 대시보드 확장

| 대시보드 | 데이터 | 주요 지표 |
|----------|--------|----------|
| **현재 8페이지** | SO/PO/DN | 수주, 출고, 커버리지, 마진, Order Book |
| **회계 대시보드** | TB | P&L 월별 추이, BS 구성, 부서별 비용 |
| **Cash Flow** | TB + AP/AR | 현금흐름, 미수금 Aging, 미지급 현황 |
| **재고 대시보드** | Inventory | 재고 금액, 회전율, 과잉/부족 경보 |

동일한 `dashboard.py` 패턴:
- `@st.cache_data` 로더 + SQL 쿼리
- `filt()` 공통 필터
- Plotly 차트 + KPI 카드

### 배포

현재 `dashboard_dist/` Portable 빌드 구조를 그대로 사용:
- Python Embedded + 패키지 포함, 무설치 실행
- OneDrive로 DB 파일 공유 → 자동 탐색
- 사용자는 bat 더블클릭만

---

## 구현 우선순위

1. **TB 동기화** — ERP에서 TB export → `sync_erp.py` → SQLite
2. **회계 대시보드** — P&L, BS 월별 추이 (가장 수요 높음)
3. **AP/AR Aging** — 미수금 관리
4. **Order Book ERP 전환** — 현재 Excel → ERP export로 소스 교체

---

## 제약 사항

- ERP DB 직접 접근 불가 → Excel/CSV 다운로드 기반
- 다운로드 주기에 따라 데이터 신선도 결정 (일 1회 권장)
- D365 F&O와 AX2009 컬럼 구조가 다를 수 있음 → 컬럼 매핑 레이어 필요
