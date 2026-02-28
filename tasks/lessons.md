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
