# Current Tasks — Sync 이력 로깅 개선 [2026-06-22]

리뷰(멀티에이전트 + 적대적 검증) 확정 항목을 하나씩 수정. 원칙: 데이터 무결성 유지, 최소 영향, 회귀 없음.

## 🔴 핵심
- [ ] 1. 로그 기록 실패가 동기화 성공을 가림 — `sync_db.py` main에서 try/except 격리, `print_summary` 항상 실행
- [ ] 2/3. 삭제 레코드 SO_ID 검색 누락 + `pk_json[0]`만 매칭 — `dashboard.py` 전체 토큰 매칭 헬퍼로 2곳 통일
- [ ] 4. 빈 PK/깨진 JSON 페이지 크래시 — 같은 헬퍼에 list·try 가드
- [ ] 5. `_sync_runs` 시각이 "로그쓰기 시점" — `started_at` 실제 sync 시작으로, 대시보드 "시작 + 소요(초)"
- [ ] 6. `docs/DB_SYNC_GUIDE.md` 기록규칙 v1 → v2
- [ ] 7. CLI `--log [N]` / `--note` 추가

## 🟡 다듬기
- [ ] PK 정규화 비대칭 — `write_sync_log_to_db`에서 `_normalize_pk` 일괄
- [ ] `--changes` 롤백 표시 — 에러 시 경고 헤더
- [ ] 추이 차트 `errors='coerce'`
- [ ] ack silent failure — bool 반환 + `st.error`
- [ ] actor 빈문자열 — `(unknown)` 정규화
- [ ] `resolve_related_ids` 2-pass → 6
- [ ] 빈시트 vs 정상 prune 스냅샷 컬럼 통일(헬퍼)
- [ ] `print_summary` 전각 정렬 — east_asian_width 패딩
- [ ] dead `note` 배선 + `dry_run` 주석/대시보드 공백컬럼 제거
- [ ] FK 미강제 — `PRAGMA foreign_keys=ON`(선제)
- [ ] 변경 0건 run 미기록 + 침묵 — 정상 실행 항상 run, 0건 안내

## 보너스
- [ ] `migrate_sync_log.py` `input()` isatty 가드

## 검증
- [x] py_compile 전체 OK (6개 파일)
- [x] pytest tests/ — 281 passed, 2 skipped (회귀 없음, 베이스라인 동일)
- [x] 스모크: write_sync_log(pk정규화 1.0→1, 0건 run기록, note, started_at), `--log` CLI, prune(정상+빈시트)

## Review (2026-06-22)

모든 🔴/🟡/보너스 항목 완료. 파일별 요약:
- **sync_db.py**: 로그기록 try/except 격리(요약 항상 출력·종료코드 보존), 전각 정렬 헬퍼,
  롤백 경고 헤더, write_sync_log(pk 정규화 일괄·0건도 run기록·FK ON·note·조회안내), `--log`/`--note`, `show_log`
- **db_schema.py**: create_sync_run `started_at` 인자(실제 sync 시작시각), dry_run 주석
- **db_sync.py**: SyncSummary.started_at, prune 스냅샷 전체 DB 컬럼 통일(`_prune_snapshot_columns`)
- **dashboard.py**: `_pk_tokens` 전체토큰 매칭(삭제 DN SO_ID 누락+크래시 동시해결, 2곳),
  추이 errors=coerce, ack bool+st.error, actors `(unknown)` 정규화, runs_view 소요(초)+dry_run 제거,
  resolve_related_ids 6-pass
- **docs/DB_SYNC_GUIDE.md**: 기록규칙 v1→v2, started_at/dry_run 설명 정정, `--log`/`--note` 추가
- **migrate_sync_log.py**: input() isatty 가드

검증에서 무해 판정된 항목(빈 삭제 레코드 dead분기, 'None' 리터럴, v1→v2 삭제 snapshot,
마이그레이션 세션병합/pk split 등)은 의도적 미적용 — 1회성·도달불가 또는 이미 방어됨.

후속(미커밋): 원하면 docs/CHANGELOG.md 항목 추가 + 커밋.
