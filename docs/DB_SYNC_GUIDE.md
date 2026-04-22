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
| `noah_data.db` | DATA_DIR (Excel과 같은 폴더) | SQLite DB (변경 이력 포함) |

> 과거에는 별도 `sync_log.csv` 파일에 변경 이력을 기록했으나, 파일이 커지면서 Excel 열기가 어려워져 `noah_data.db` 내 `_sync_log` 테이블로 일원화됨 (2026-04). 기존 CSV는 `migrate_sync_log.py`로 DB에 이관.

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

### 동기화 세션 메타 `_sync_runs`

한 번의 sync 호출 = 1개 `_sync_runs` row + N개 `_sync_log` row.

| 컬럼 | 설명 |
|------|------|
| `sync_id` | AUTOINCREMENT PK — 모든 `_sync_log` 행의 FK |
| `started_at` | 세션 시작 (`YYYY-MM-DD HH:MM:SS`) |
| `ended_at` | 세션 종료 (변경 0건이면 NULL 가능) |
| `actor` | 실행 사용자 (`USERNAME` env, fallback `os.getlogin()`) |
| `host` | 실행 호스트 (`socket.gethostname()`) |
| `dry_run` | 0=실제 commit, 1=dry-run (현재 dry-run은 _sync_log 안 씀) |
| `total_changes` | 이번 세션 record 수 |
| `note` | `migrated from v1` 등 주석 |

### 변경 이력 테이블 `_sync_log` (v2)

동기화 시 발생한 신규/수정/삭제 이력을 **record 단위**로 누적 저장. 인덱스: `sync_id`, `(sheet_name, sync_id)`, `pk_display`.

| 컬럼 | 설명 |
|------|------|
| `id` | AUTOINCREMENT PK |
| `sync_id` | `_sync_runs.sync_id` FK |
| `sheet_name` | 소스 시트명 |
| `change_type` | `신규` / `수정` / `삭제` |
| `pk_json` | PK JSON 배열, 예 `["SOD-2026-0001","1"]` (구조 보존, JSON 함수로 추출 가능) |
| `pk_display` | PK 표시 문자열, 예 `"SOD-2026-0001 | 1"` (검색·UI 호환) |
| `changes_json` | 신규/수정 정보 (JSON). 삭제 시 NULL. |
| `row_snapshot_json` | 삭제 직전 전체 row JSON. 신규/수정 시 NULL. |

`changes_json` 구조:

| change_type | 형태 | 예시 |
|---|---|---|
| 신규 | `{col: value, ...}` | `{"SO_ID":"SOD-2026-0001","Status":"Open"}` |
| 수정 | `{col: {old, new}, ...}` | `{"Status":{"old":"Open","new":"Closed"}}` |
| 삭제 | NULL | (사용 안 함, snapshot에 보관) |

`row_snapshot_json` 구조 (삭제 시):
```json
{"SO_ID":"SOD-2026-0001","Status":"Open","비고":"...","Customer":"..."}
```

### v1 → v2 마이그레이션

기존 v1 (필드당 1행)을 v2 (record당 1행 + JSON)로 변환. 한 번만 실행.
```bash
python migrate_sync_log_v2.py             # 마이그레이션 실행 — _sync_log_legacy 백업 후 변환
python migrate_sync_log_v2.py --dry-run   # 변환 통계만 출력 (DB 변경 없음)
python migrate_sync_log_v2.py --drop-legacy  # 검증 끝나면 _sync_log_legacy 제거
```
실측: 115,831행 → 15,263행 (86.8% 압축), 76개 세션 복원.

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
- 삭제된 PK는 `_sync_log` 테이블에 "삭제" 유형으로 기록
- `_sync_meta`, `_sync_log`, `ob_snapshot` 등 시스템/메타 테이블은 prune 대상 아님

### Dry-run 동작

- `--dry-run`은 실제 DB에 연결하여 upsert/prune을 수행한 뒤 **트랜잭션을 rollback** — DB는 변경되지 않음
- `isolation_level=None` + 명시적 `BEGIN`으로 DDL(테이블 생성/삭제/컬럼 추가)도 트랜잭션 내에서 실행 → rollback 시 완전 원복
- 운영 DB 기준의 정확한 insert/update/prune 건수를 시뮬레이션
- `_sync_log` 테이블에도 기록하지 않음

## 변경 이력 로그 (`_sync_log` 테이블)

동기화할 때마다 변경 내역이 `noah_data.db`의 `_sync_log` 테이블에 자동 누적됨. Streamlit 대시보드의 **동기화 로그** 페이지에서 필터·검색·CSV 내보내기 가능.

**기록 규칙:**
- **신규**: 비어있지 않은 필드마다 1행씩 기록 (`old_value` = NULL)
- **수정**: 변경된 필드마다 1행씩 기록 (`old_value` / `new_value` 모두 채움)
- **삭제**: PK만 기록 (`column_name` / `old_value` / `new_value` 모두 NULL)

**조회 방법:**

1. **대시보드** — `streamlit run dashboard.py` → 사이드바 "동기화 로그" 페이지
2. **SQL 직접 조회** (v2 스키마):
   ```sql
   -- 최근 7일 수정 이력 (세션 메타와 함께)
   SELECT r.started_at, r.actor, l.sheet_name, l.pk_display, l.changes_json
   FROM _sync_log l
   JOIN _sync_runs r ON l.sync_id = r.sync_id
   WHERE r.started_at >= date('now', '-7 days')
     AND l.change_type = '수정'
   ORDER BY l.id DESC;

   -- 특정 PK 변경 이력 추적
   SELECT r.started_at, l.change_type, l.changes_json, l.row_snapshot_json
   FROM _sync_log l
   JOIN _sync_runs r ON l.sync_id = r.sync_id
   WHERE l.pk_display LIKE 'SOD-2026-0001%'
   ORDER BY l.id;

   -- 삭제된 행 복구 (row_snapshot_json 활용)
   SELECT r.started_at, r.actor, l.sheet_name, l.pk_display,
          json_extract(l.row_snapshot_json, '$.Status') AS deleted_status
   FROM _sync_log l
   JOIN _sync_runs r ON l.sync_id = r.sync_id
   WHERE l.change_type = '삭제'
     AND l.row_snapshot_json IS NOT NULL
   ORDER BY r.started_at DESC;

   -- 사용자/호스트별 동기화 활동
   SELECT actor, host, COUNT(*) AS sessions, SUM(total_changes) AS changes
   FROM _sync_runs
   WHERE started_at >= date('now', '-30 days')
   GROUP BY actor, host;
   ```

**CSV에서 DB로 마이그레이션** (1회성, 이미 완료):
```bash
python migrate_sync_log.py              # 기존 sync_log.csv → _sync_log
python migrate_sync_log.py --dry-run    # 파싱 테스트만
python migrate_sync_log.py --delete     # 성공 시 CSV 파일 삭제
```

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
- **변경 추적**: `_sync_log` 테이블에 "언제, 어떤 행의 어떤 필드가 무엇에서 무엇으로 바뀌었는지" 자동 기록 → 대시보드에서 바로 조회
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
| 변경 이력 | `_sync_log` 테이블로 데이터 변경 추적 (대시보드 조회) |
| ERP 마이그레이션 | 데이터 클렌징/매핑 기초 자료 |
