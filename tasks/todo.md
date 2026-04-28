# Current Tasks

## In Progress
<!-- - [ ] Task description -->

## Pending
- [ ] Power Query → SQL 쿼리 세트 구현 (DB 활용)
  - SO_통합, DN_원가포함, PO_현황, PO_매입월별, PO_AX대사, PO_미출고, Order_Book
  - 기존 Power Query(M 코드) 기준으로 SQL 변환
  - Excel = 입력 도구, DB = 분석/리포트 도구로 역할 분리

## Completed
- [x] **SO 단가/수량 무단 변경 경고** (2026-04-28)
  - `_so_change_ack` 테이블 추가 (`po_generator/db_schema.py`)
  - `load_so_unauth_changes()` + `_ack_so_change()` 추가 (`dashboard.py`)
  - `pg_today()` 에 ⚠️ 섹션 추가 — Customer PO 변경 없이 `Item qty` / `Sales Unit Price` / `Sales amount(KRW)` 가 바뀐 미확인 변경을 expander + 확인완료 버튼으로 dismiss
  - ack 영구 보존 (audit trail), 0건이면 섹션 숨김

## Review
### 2026-04-28: SO 무단 변경 경고
- **검증**:
  - DDL idempotency + INSERT OR IGNORE 정상
  - 1차 구현: 감시필드 4개(`Item qty`/`Sales Unit Price`/`Sales amount`/`Sales amount KRW`) → 557건 — 노이즈 압도적
  - 노이즈 분석: 456건이 `Sales amount KRW` 단독 변경 (해외 환율 자동 재계산), 359건이 변경량 < 1원 (반올림)
  - 2차 필터(B안): `Item qty`/`Sales Unit Price` 만 + 빈값↔값 제외 → **53건**으로 87% 감소
  - 잔존 53건 모두 사용자 검토 가치 있음 (예: `126,000→1,260,000` 자릿수 실수, `450,000→500,000` 단가 인상)
- **설계 결정**:
  - ack 기반 dismiss (단순 N일 윈도우 X) — 매출/세금계산서/매출대사 영향이 있어 한 번이라도 못 보면 안 됨
  - `Sales amount(KRW)` 의도적 제외 — Excel 수식/환율 자동 재계산이라 사람의 액션 신호 아님
  - `None`/빈문자열 ↔ 값 케이스 제외 — 최초 입력/삭제는 변경이 아님
- **감시 필드**: `Item qty`, `Sales Unit Price`
- **허가 신호**: `Customer PO` 함께 변경 시 자동 제외
- **제외**: `change_type='수정'` 만 — 신규/삭제는 무시. `dry_run=1` 도 제외
