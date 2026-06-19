# Lessons Learned

Patterns and mistakes to avoid, updated after each correction.

## xlwings
- `.value` 범위 읽기: 단일 열 → 1D list 반환
- `.formula` 범위 읽기: 단일 열 → 2D tuple of tuples 반환 (항상 2D)
- 배치 최적화 시 반환 형식을 실제 테스트로 확인 필요

## Path Handling
- 경로 검증 시 문자열 `in` 검사 대신 `relative_to()` 사용 (Path Traversal 방지)

## Excel Template
- 행 삭제는 "같은 위치에서 반복 삭제" (xlUp으로 아래 행이 올라옴)
- 행 삽입/삭제 후 테두리 복원 함수 반드시 호출

## DN ↔ SO 수량/단가 (거래명세표)
- **거래명세표(TS)는 "주문(SO)"이 아니라 "실제 납품(DN)" 수량/단가가 정답**
  - `DN_국내` 시트는 자체 `Item`/`Qty`/`Unit Price`/`Total Sales`(실제 출고분)를 가진다 — SO에서 가져오지 말 것
  - 부분 납품(주문 70 → 36만 출고)·분할 납품(SO 한 라인 9 = 576을 DN 8 + 212로 나눠 출고) 시 SO 주문 수량과 달라짐
  - 증상: DND-2026-0560 거래명세표가 36→70, 8/212→576/576으로 잘못 출력 (16개 라인 / 12개 DN 영향)
- **버그 원인**: `load_dn_data()`가 SO를 머지하고, `item_qty` 별칭이 `'Item qty'`(SO)를 `'Qty'`(DN)보다 먼저 매칭
  - 우연히 전량 납품 라인은 SO==DN이라 정상으로 보였음 → 부분/분할 납품에서만 드러남
- **수정**: 머지 후 `df['Item qty'] = df['Qty'].combine_first(df['Item qty'])` 식으로 DN 우선·SO 폴백 (별칭 전역 변경 없이 국한)
- **교훈**: "Single Source of Truth"를 시트 단위로 못박지 말 것 — 같은 엔티티라도 *주문값*과 *실현값*은 다른 시트가 정답. 머지로 값을 덮어쓰기 전 "이 컬럼의 진짜 출처가 어디인가" 확인

## 변경 감지 / 알림 설계
- `_sync_log` 기반 알림 만들 때 **사람의 액션이 아닌 자동 재계산 필드(파생값)는 감시에서 제외**
  - SO의 `Sales amount` / `Sales amount KRW` = `Sales Unit Price × Item qty × FX`로 매번 재계산 → 환율/반올림 노이즈가 압도적 (실측: 557건 중 89%가 이 노이즈)
  - 1차 신호(사람이 직접 입력하는 단가·수량)만 감시
- **빈값 ↔ 값 변경은 "최초 입력" 또는 "삭제"** — `None`/빈문자열 양쪽 케이스 모두 제외해야 진짜 변경만 남음
- 새 알림 기능 만들면 **DB 기반으로 분포 분석 먼저** — 시기/필드/변경량 히스토그램. 분석 없이 켜면 노이즈에 묻힘
