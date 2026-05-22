# Current Tasks

## In Progress
<!-- - [ ] Task description -->

## Pending
- [ ] Power Query → SQL 쿼리 세트 구현 (DB 활용)
  - SO_통합, DN_원가포함, PO_현황, PO_매입월별, PO_AX대사, PO_미출고, Order_Book
  - 기존 Power Query(M 코드) 기준으로 SQL 변환
  - Excel = 입력 도구, DB = 분석/리포트 도구로 역할 분리

## Completed
- [x] **Packing List Net Weight — Model+옵션 기반 Weight 매핑** (2026-05-22)
  - `config.py`: `WEIGHT_OPTION_SUFFIX`, `WEIGHT_OPTION_PRIORITY` 추가
  - `utils.py`: `build_weight_map`(미사용) 제거 → `build_model_weight_map`,
    `load_po_export_data`, `resolve_weight_code`, `build_po_line_weight_map`,
    `normalize_line_item` 추가
  - `document_service.py`: `_enrich_with_weight()` — PO_해외 (SO_ID, Line item)
    조인 + Model/옵션 → Weight 매핑으로 교체
  - `test_utils.py`: weight 매핑 단위 테스트 17개 추가

- [x] **SO 단가/수량 무단 변경 경고** (2026-04-28)
  - `_so_change_ack` 테이블 추가 (`po_generator/db_schema.py`)
  - `load_so_unauth_changes()` + `_ack_so_change()` 추가 (`dashboard.py`)
  - `pg_today()` 에 ⚠️ 섹션 추가 — Customer PO 변경 없이 `Item qty` / `Sales Unit Price` / `Sales amount(KRW)` 가 바뀐 미확인 변경을 expander + 확인완료 버튼으로 dismiss
  - ack 영구 보존 (audit trail), 0건이면 섹션 숨김

## Review
### 2026-05-22: Packing List Net Weight 매핑
- **문제**: PL Net Weight(G열)가 SO_해외 `Model code`로 매핑됐는데 그 컬럼이
  전 행 비어 있어 항상 공란이었음. 사실상 신규 기능.
- **데이터 구조**:
  - 액추에이터 Model/옵션은 `PO_해외`에만 존재 (Model=AN열, 옵션 20개=AO~BH열 Y표시)
  - PL은 DN_해외 기반 → PO_해외와 `(SO_ID, Line item)` 복합키 조인 (PO 중복쌍 7건→첫행)
  - `Weight` 시트: `MODEL`(단축코드), `ITEM`(서술형), `WEIGHT` — 매칭은 `MODEL` 사용
- **매핑 규칙**: PO Model에서 `NA`/`SA` 접두어 제거 → base 코드. 무게 영향 옵션
  (INTEGRAL→IN, IMS→IM, LCU→L, PCU+PIU→P, SCP→S, EXP→X)을 접미사로 부착.
  복수 옵션은 우선순위 1개(`INTEGRAL>IMS>LCU>PCU+PIU>SCP>EXP`), LCU+PCU 동시는
  결합코드 `…LP` 우선. 미매칭 시 base Model 폴백 → base도 없으면 공란.
- **검증**:
  - 핵심 로직 단위 테스트 17개 + 전체 279 passed
  - 실제 PL 3건 생성: DNO-2026-0003(base폴백 NA015/028/060/009→14/18.5/27/11),
    DNO-2026-0002(옵션매칭 SA005L→4.1 `005LP`), DNO-2026-0008(SR10P→41 + Model
    없는 라인 공란)
  - 커버리지: PO_해외 Model 553행 중 549행 해결, 미매칭 4건은 비표준 액세서리
- **설계 결정**:
  - 복수 옵션 시 합산 대신 우선순위 1개 — Weight 시트에 결합 행이 없어 합산은 추정치
  - 매칭 키로 `ITEM`(서술형, `INTE`/`Integ` 등 표기 불일치) 대신 `MODEL` 단축코드 사용
  - 옵션→접미사/우선순위는 `config.py` 상수로 분리 — 향후 Weight 시트 변경 시 조정 용이
- **표시 단위**: 행별 G열 = 단위중량(KG/PC). Total 행 G열은
  `SUMPRODUCT(Qty, 단위중량)` = 총 Net Weight (수량 반영). 미매칭 빈 셀은
  SUMPRODUCT가 0으로 처리 → 오류 없음

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
