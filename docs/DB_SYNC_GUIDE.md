# Excel → SQLite DB 동기화 가이드

NOAH_SO_PO_DN.xlsx 데이터를 SQLite DB에 백업/동기화하는 기능.

---

## 왜 필요한가?

- Excel은 수동 입력 시 데이터 유실/변형에 취약
- DB에 주기적으로 동기화하면 변경 이력 추적 + 데이터 안전 백업
- SQL로 자유로운 데이터 분석 가능 (DB Browser, Python 등)

## 사용법

```bash
python sync_db.py                           # 전체 동기화
python sync_db.py --changes                 # 동기화 + 변경 내역 표시
python sync_db.py --sheets SO_국내 PO_국내  # 특정 시트만
python sync_db.py --dry-run                 # 시뮬레이션 (DB 변경 안 함)
python sync_db.py --info                    # DB 현황 조회
python sync_db.py -v                        # 상세 로그
```

bat 메뉴에서는 `[5] Excel → DB 동기화` 선택.

## 파일 위치

| 파일 | 위치 | 설명 |
|------|------|------|
| `noah_data.db` | DATA_DIR (Excel과 같은 폴더) | SQLite DB |
| `sync_log.csv` | DATA_DIR (Excel과 같은 폴더) | 변경 이력 로그 |

## 테이블 구조

7개 시트 → 7개 테이블. 국내/해외 컬럼 구조가 다르므로 별도 테이블.

| 테이블명 | 소스 시트 | PK | 비고 |
|----------|----------|-----|------|
| `so_domestic` | SO_국내 | `(SO_ID, Line item)` | |
| `so_export` | SO_해외 | `(SO_ID, Line item)` | |
| `po_domestic` | PO_국내 | `(PO_ID, Line item, _row_seq)` | 부분 매입 순번 |
| `po_export` | PO_해외 | `(PO_ID, Line item, _row_seq)` | 부분 매입 순번 |
| `dn_domestic` | DN_국내 | `(DN_ID, SO_ID, Line item)` | |
| `dn_export` | DN_해외 | `(DN_ID, SO_ID, Line item)` | |
| `pmt_domestic` | PMT_국내 | `(선수금_ID)` | |

### PO 테이블 `_row_seq`

부분 매입 시 같은 Line item이 분할되어 중복될 수 있으므로 (예: Line item 1, qty 2 → Line item 1, qty 1 두 행), 같은 `(PO_ID, Line item)` 내에서 Excel 행 순서대로 자동 순번(1,2,3...) 부여. PO_국내/PO_해외 모두 적용.

### 메타 테이블 `_sync_meta`

| 컬럼 | 설명 |
|------|------|
| `table_name` | 테이블명 (PK) |
| `last_sync` | 마지막 동기화 시각 (ISO) |
| `row_count` | 동기화 후 행 수 |

## 동기화 동작

### Upsert 방식

1. Excel 시트 로드 → 빈 행 제거 (required_column 기준)
2. PK로 기존 행 조회
3. **행이 없으면** → INSERT (신규)
4. **행이 있으면** → 기존 값과 비교 → 변경된 필드가 있을 때만 UPDATE (수정)
5. **값이 동일하면** → 스킵 (동일)
6. **DB에만 존재하는 행** → DELETE (삭제/prune) — Excel에서 제거된 행을 DB에서도 정리

### 변경 감지

- 모든 컬럼을 문자열로 비교 (None, 빈 문자열은 동일 취급)
- 실제 값이 바뀐 필드만 수정으로 기록
- 변경 없는 행은 UPDATE 안 함 → DB 부하 최소화

### 삭제 반영 (Prune)

- Upsert 완료 후, DB에만 존재하고 Excel에는 없는 PK를 감지하여 DELETE
- **시트가 완전히 비어있는 경우에도 동작** — DB 테이블의 전체 행이 prune 대상
- 삭제된 PK는 `sync_log.csv`에 "삭제" 유형으로 기록
- `_sync_meta`, `ob_snapshot` 등 비동기화 테이블은 prune 대상 아님

### Dry-run 동작

- `--dry-run`은 실제 DB에 연결하여 upsert/prune을 수행한 뒤 **트랜잭션을 rollback** — DB는 변경되지 않음
- `isolation_level=None` + 명시적 `BEGIN`으로 DDL(테이블 생성/삭제/컬럼 추가)도 트랜잭션 내에서 실행 → rollback 시 완전 원복
- 운영 DB 기준의 정확한 insert/update/prune 건수를 시뮬레이션
- sync_log.csv에는 기록하지 않음

## 변경 이력 로그 (sync_log.csv)

동기화할 때마다 변경 내역이 CSV로 자동 기록됨. Excel에서 바로 열어서 확인 가능.

| 컬럼 | 설명 | 예시 |
|------|------|------|
| 동기화시각 | 실행 시각 | 2026-03-03 09:28:58 |
| 시트 | 소스 시트명 | SO_국내 |
| 유형 | 신규/수정/삭제 | 수정 |
| PK | Primary Key 값 | SOD-2026-0001 \| JK2026... \| 1 |
| 컬럼 | 변경된 컬럼명 | Status |
| 이전값 | DB 기존 값 | (빈값) |
| 변경값 | Excel 새 값 | Cancelled |

- **신규**: 비어있지 않은 필드마다 1행씩 기록 (이전값은 빈칸, 변경값에 입력된 값 표시)
- **수정**: 변경된 필드마다 1행씩 기록 → 필터/정렬 가능
- **삭제**: PK만 기록 (Excel에서 제거되어 DB에서 삭제된 행)
- UTF-8 BOM 포함 → Excel에서 한글 깨짐 없음

## DB 조회 방법

### 1. DB Browser for SQLite (GUI)

https://sqlitebrowser.org/dl/ 에서 설치 후 `noah_data.db` 열기.

### 2. Python

```python
import sqlite3, pandas as pd
conn = sqlite3.connect("noah_data.db")
df = pd.read_sql("SELECT * FROM so_domestic WHERE Status = 'Open'", conn)
print(df)
conn.close()
```

### 3. CLI

```bash
python sync_db.py --info    # 테이블별 행 수, 마지막 동기화 시간
```

## 소스 코드

| 파일 | 역할 |
|------|------|
| `sync_db.py` | CLI 진입점 |
| `po_generator/db_schema.py` | 테이블/PK 정의, DDL, 스키마 관리 |
| `po_generator/db_sync.py` | SyncEngine — upsert + prune 엔진, 변경 감지 |
| `po_generator/config.py` | `DB_FILE` 상수 |

## 에러 처리

| 상황 | 동작 |
|------|------|
| Excel 파일 없음 | exit code 1 |
| 시트 없음 | 경고 후 스킵 |
| PK 필수 컬럼 NaN | 행 스킵 |
| PK 비필수 컬럼 NaN | 빈 문자열로 치환하여 INSERT 허용 |
| 개별 행 에러 | 경고 + 스킵, 요약에 에러 수 표시 |

## 스키마 진화

Excel에 새 컬럼이 추가되면 `ensure_columns_exist()`가 자동으로 ALTER TABLE ADD COLUMN 실행. 기존 데이터는 유지됨.

## DB 동기화의 가치

### 현재 가치

- **데이터 유실 방지**: Excel에서 실수로 행 삭제/덮어쓰기 시 복구 불가 → DB에 백업되어 있으면 복원 가능
- **변경 추적**: sync_log.csv에 "언제, 어떤 행의 어떤 필드가 무엇에서 무엇으로 바뀌었는지" 자동 기록
- **내부통제 증거**: ERP 통합 전까지 "인터컴퍼니 거래 데이터를 별도로 관리/검증하고 있다"는 증빙
- **데이터 정합성**: 동일 데이터를 두 곳(Excel, DB)에 보관 → 불일치 시 문제 감지 가능

### 한계

- 감사에서 원하는 "누가, 왜 변경했는지"까지는 추적 불가 (Excel 자체 한계)
- ERP를 대체하는 것이 아닌 **ERP 통합 전까지의 안전장치 + 분석 도구**

## 향후 활용 계획

### 역할 분리 구상

```
Excel (NOAH_SO_PO_DN.xlsx)  →  입력 도구 (수기 입력 유지)
         ↓ sync_db.py
SQLite (noah_data.db)       →  분석/리포트/검증 도구
```

현재 Power Query(M 코드)로 하고 있는 분석을 SQL로 전환 가능. Power Query는 Excel 내 실시간 피벗/차트 연동용으로 유지하고, DB는 정합성 검증, 감사 대사, 복잡한 크로스 분석을 보완.

### SQL 전환 대상 (기존 Power Query)

| Power Query | 용도 | SQL 전환 시 장점 |
|---|---|---|
| SO_통합 | 주문 현황 + 원가 + 마진 + 출고 상태 | 다중 JOIN 자유, 조건 필터 유연 |
| DN_원가포함 | 출고 내역 + 원가 + GL대상 | subquery로 정합성 검증 가능 |
| PO_현황 | 발주 현황 + Status별 집계 | GROUP BY로 집계 간결 |
| PO_매입월별 | 월별 매입 집계 (IC Balance) | 기간별 피벗 쉬움 |
| PO_AX대사 | Period + AX PO별 GRN 대사 | 회계 마감 대사 자동화 |
| PO_미출고 | Invoiced인데 DN 없는 건 | LEFT JOIN + WHERE NULL 패턴 |
| Order_Book | 월별 수주잔고 롤링 | 윈도우 함수로 롤링 계산 |

### Power Query vs SQL 비교

| 항목 | Power Query | SQL |
|------|------------|-----|
| JOIN | 양방향 불가 | 제한 없음 |
| 디버깅 | M 코드 (읽기 어려움) | SQL (읽기 쉬움) |
| 성능 | 데이터 많으면 느림 | 빠름 |
| 공유 | Excel 파일 안에 묶임 | .sql 파일로 공유 가능 |
| 실시간 차트 | 연동됨 | 별도 내보내기 필요 |

### 감사/내부통제 활용

| 활용 | SQL 쿼리 예시 |
|------|-------------|
| PO ↔ DN 금액 정합성 | PO 발행건 중 DN 미매칭 또는 금액 불일치 검출 |
| PMT 입금 대사 | 입금액 vs PO 총액 차이 검증 |
| 변경 이력 | sync_log.csv로 데이터 변경 추적 |
| ERP 마이그레이션 | 데이터 클렌징/매핑 기초 자료 |
