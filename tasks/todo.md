# Current Tasks

## Completed — LOW 30건 [2026-06-22]
5개 파일그룹 병렬 에이전트로 처리. pytest 281 passed + 라이브DB SQL 실행확인.
- [x] 생성기/CLI/core 9+5건: resolve_column 비문자열가드, escape_excel 선행제어문자, validators 0값,
      to_text 승격+모델코드, get_column_letter(>Z/AA+), 미사용상수, 0단가 fallback, ci 중복블록,
      create_fi --po충돌, create_po force분리, create_ts 고객명시, finder trim, history 정규식
- [x] 대사 6건: FX 상대오차, Customer backfill, 상세 날짜컬럼, build_mapping 우선dedup,
      fill 불일치경고, recon_paths 빈 플랫폴더
- [x] 대시보드 6건: sync-log 정확매칭, OTD groupby, aging 라벨, 미출고 clip+캡션, backlog 캡션, HAVING
- [x] DB/SQL 4건: _sync_meta 빈시트, NULL EDD COALESCE(3파일), backlog 라운딩/HAVING

### 미적용 (의도적 제외 — 사유 명시)
- ts_generator 라인별 VAT(sub-10원, 의도적), FI 모델 prefix(소유자 확인필요),
  finder get_available_*(dead code), snapshot undo 이력테이블(스키마변경), theme CSS(시각회귀위험),
  Book-to-Bill inf 라벨(미관), analyze_sheets(git-ignored), validate_delivery_date dayfirst(조직관례),
  margin basis 재정의(대규모), dashboard 마감스냅샷 배선(설계), snapshot EDD-move/소급(by-design)

## Completed — 감사 수정 (High 1 + Medium 12) [2026-06-22]
멀티에이전트 감사(101 발견→53 확정) 중 High 1 + Medium 12 수정. pytest 281 passed.
- [x] #1  dashboard.py: 날짜 1900 더미 정화 중앙화 (_sanitize_date) — load_so 4컬럼 + load_backlog
- [x] #2  utils.py: _load_and_merge_sheets / load_dn_data 복합키 머지 참조측 dedup + 경고
- [x] #3  utils.py: resolve_column 캐시 키 id(columns) → tuple(columns)
- [x] #4  create_ts.py: --merge 중복 DN_ID dedup + 경고
- [x] #5  ts_generator.py: 수량 파싱 int(raw_qty) → int(float(raw_qty))
- [x] #6  db_sync.py: prune 2곳 rowid 기준 삭제 + rowcount 확인
- [x] #7  reconcile_so.py: FX 월매칭 연도 인식 (연도폴더 → 정확매칭)
- [x] #8  reconcile_so.py: 매출일 결측 월필터 누락 경고
- [x] #9  reconcile_ind.py: ind_code 정규화 헬퍼 공유 (마스터/SO 양측)
- [x] #10 dashboard.py: 납기현황 SO 합계 전체 라인 기준으로 수정 (+PO EXW fan-out 방지)
- [x] #11 order_book_variance.sql: 납기변경 상쇄쌍(net≈0) 매칭
- [x] #12 migrate_sync_log.py: v1 마이그레이터 → _sync_log_legacy_v1 전용 테이블
- [x] #13 create_ts.py: 월합 출력 파일명 충돌 안전장치
- [x] 검증: pytest 281 passed + 순수로직/prune sqlite 시뮬 검증

## Completed — 2차 감사 수정 (재심복구 + 신규 + 회귀) [2026-06-22]
2차 워크플로우: 기각48 재심→7 진짜버그 복구, 신규17 확정, 회귀3. pytest 281 passed.
수정 적용분:
- [x] 회귀#10 dashboard 납기현황: 라인별 잔여 음수(과출고) clamp(>=0) — 부족분 상계 방지
- [x] 회귀#6  sync_db: 삭제 감사로그를 pruned_snapshots 1:1 순회 (중복/스냅샷유실 제거)
- [x] sync_db: 롤백(total_errors>0)시 _sync_log 유령기록 방지 게이팅 (HIGH)
- [x] qty int(raw_qty)→int(float()) — fi/ci/oc/pi/pl 5개 생성기 (HIGH, TS와 동일클래스)
- [x] reconcile_so: AX Project 정규화 헬퍼 양측 적용 (.0 업캐스트 매칭불가→매출누락) (HIGH)
- [x] reconcile_so: 해외 'Total Sales KRW' 결측 경고 (KRW 0집계→불일치 오표시) (HIGH)
- [x] create_pi: 비숫자 단가 :.2f 가드 (배치 전체 중단 방지)

### Completed — 2차 MEDIUM 6건 [2026-06-22]
- [x] reconcile_po: 1:N AX PO 이중계상 → _line_id로 계산서금액 분배 (_agg_delivery, .copy로 격리)
- [x] reconcile_po: ax_service 국내/해외 키 disjoint(이중계상 방지) + 미분류 Product GRN 경고
- [x] snapshot.py: 마감 시 출하·KRW 공란 해외 DN 경고 (phantom backlog 동결 전 surface)
- [x] create_fi: 복수 RCK PO 분리 시 공란 RCK PO 라인 누락 경고
- [x] document_service --po FI: 복수 DN 통합 시 Invoice No 대표DN 경고
- [x] dashboard load_backlog + order_book_backlog/_snapshot_backlog.sql: ROUND 통일 (order_book/snapshot과 tie-out)
- [x] 검증: pytest 281 passed + reconcile_po 분배/disjoint 시뮬 + 라이브DB SQL 실행확인

### 미적용 (2차 LOW — 추후)
- dashboard Order Book 마감 스냅샷 미사용 — 마감월도 라이브값 (LOW, 캡션 보강 권장)
- fi/ci 0단가 fallback이 무상라인을 SO가로 청구 (LOW)
- snapshot EDD-move Start=0 attribution / 과거기간 소급편집 귀속 (LOW, variance.sql이 중화)

### 미적용 (1차 LOW 39건)
대표: create_po --force 검증오류 동반묵살, FI 모델 prefix 누락, margin basis mismatch,
sync-log 부분문자열 매칭, OTD drop_duplicates fan-out 등

## Pending

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
