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

## 변경 감지 / 알림 설계
- `_sync_log` 기반 알림 만들 때 **사람의 액션이 아닌 자동 재계산 필드(파생값)는 감시에서 제외**
  - SO의 `Sales amount` / `Sales amount KRW` = `Sales Unit Price × Item qty × FX`로 매번 재계산 → 환율/반올림 노이즈가 압도적 (실측: 557건 중 89%가 이 노이즈)
  - 1차 신호(사람이 직접 입력하는 단가·수량)만 감시
- **빈값 ↔ 값 변경은 "최초 입력" 또는 "삭제"** — `None`/빈문자열 양쪽 케이스 모두 제외해야 진짜 변경만 남음
- 새 알림 기능 만들면 **DB 기반으로 분포 분석 먼저** — 시기/필드/변경량 히스토그램. 분석 없이 켜면 노이즈에 묻힘
