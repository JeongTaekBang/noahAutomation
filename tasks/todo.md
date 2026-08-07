# 거래명세표 묶음 발송 — 하루치를 거래처별 메일 한 통으로 (2026-08-07) — 완료

사용자 요청: 8/6처럼 한 거래처에 여러 PO가 나간 날, PDF는 PO(=DN)별로 만들되
메일은 한 통에 첨부 여러 개로. 예시가 `DN_국내` 8/6 출고분 씨앤케이엔지니어링(DN 8건).

- [x] `mailer`: `as_paths` / `export_pdfs`(Excel 1회 기동) / `build_attachments`·
      `create_document_mail`·`create_ts_mail`이 경로 목록 허용 (OC·납기현황 호출부 무수정)
- [x] `create_ts`: `_build_ts_from_dn`·`_build_ts_from_adv` 추출(단건·묶음 같은 경로),
      `--date`/`--customer`/`--one-mail`, `group_by_customer`(정규화 사업자번호),
      `validate_selection_args`(--merge와 동시 지정 차단)
- [x] 거래처 섞임 관문 `foreign_biz_numbers` — 실데이터에 DN 번호 재사용 2건 발견
      (`DND-2026-0748` 씨앤케이+오토밸브 / `DND-2026-0328` 코콘+한일전자).
      묶음은 그 문서만 첨부에서 제외, 단건도 같은 관문
- [x] `utils.BIZ_NO_MIN_DIGITS` 이동 (delivery_status와 조회어 판정 공유)
- [x] `create_po.bat`(대화형 메뉴): 거래명세표 하위 메뉴에 `[3] 하루치 묶음 발송`,
      `[4] 묶음 발송(목록 붙여넣기)` 추가 — **CLI만 고치면 메뉴에서는 안 보인다**
- [x] `noah_gui`: TS 옵션 체크박스 + `--one-mail` 배선
- [x] 테스트: `tests/test_create_ts_batch.py` 신규 56건, `test_mailer` 첨부 복수 7건,
      `test_noah_gui` 1건. 기존 `export_pdf` patch 지점을 `export_pdfs`로 이동
- [x] 검증: pytest 전체 통과 / 실데이터 8장 생성 + PDF 일괄변환 24.2초 + .eml 첨부 8개 확인

**DN 번호 중복 2건 처리 (2026-08-07 결정):**
- `DND-2026-0328`(한일전자 `SOD-2026-0383` 행) — **그대로 두기로 함.** 4월 건이고 회계는
  SO_ID 단위로 이미 맞다. 관문이 메일 발송만 막으므로 실질 피해는 "그 DN으로 거래명세표를
  못 뽑는다" 하나뿐
- `DND-2026-0748`(오토밸브 `SOD-2026-0743` 행) — 미정. 그대로 두면 8/6 씨앤케이 묶음 메일에
  **7장만 붙고 `26071408R0` 명세표가 빠진다**

번호를 새로 줄 때 **뒤 번호를 밀 필요는 없다** — 빈 번호가 21개 있고(`350~361` 12개 연속은
`_sync_log` 35,351행·생성 문서 어디에도 흔적 없음), 번호↔출고일 역전도 이미 22건이라
DN 번호는 시간순 일련번호가 아니다. 파생 시트 5종(`DN_원가포함`·`SO_통합`·`AX_매출대사`·
`PO_출고`·`INV_transaction`)은 파워쿼리라 새로고침으로 따라온다.

---

# OC "한 페이지 빈 행 채우기" 제거 (2026-08-05 오후) — 완료

사용자 보고: 1아이템 OC가 빈 격자 예닐곱 줄을 달고 나감 → 표는 마지막 아이템
바로 다음 Total로 끝나도록 채움 기능 자체를 제거 (8/3 도입분 되돌림).

- [x] oc_generator: `_page_blank_capacity` 제거, `_fill_items`를
      "부족분 삽입 → 값·행높이 → 남는 행 삭제"로 단순화
- [x] excel_helpers: 채움 전용 스택 제거 (`fit_blank_rows`/`printable_height`/
      `sum_row_heights`/`print_area_last_row`/`A4_HEIGHT_PT`)
- [x] tests/test_page_fit.py → test_doc_layout.py 개명(git mv), 채움 테스트 13개 삭제
- [x] CLAUDE.md 참조 2곳·CHANGELOG 갱신
- [x] 검증: pytest 638 passed · 1아이템(SOO-2026-0239) 재생성 = 아이템+Total만 ·
      27아이템(SOO-2026-0235) 재생성 = 행 18-44 + Total 45 (기존과 동일 구조)

---

# OC PDF 레이아웃 결함 2건 (2026-08-05) — 완료

사용자 보고: OC PDF에서 (1) Customer Address 왼쪽 가운데 줄과 오른쪽 Delivery Address가
잘림, (2) 본문 아이템 가로줄이 두꺼워 보임.

## 원인 (실측 완료)

1. **주소 잘림** — 주소 칸이 병합 셀(왼쪽 A13:E13~A15:E15, 오른쪽 G13:I15)인데
   wrap 없이 한 줄로 들어가서 **병합 경계에서 클립**된다. SECTORIEL 실측:
   bill_to_2 74자 > A:E 폭(~54자), 납품주소 81자 > G:I 폭(~42자).
   병합 셀은 넘친 텍스트를 옆 칸으로 흘리지 않는다.
2. **굵은 선** — 아이템 그리드는 **검정 thin**(PDF 실측 0.96pt)인데, 문서 상단
   규칙선은 **회색 #BBBBBB thin**(0.733 gray)이라 그리드만 무겁게 보인다.
   캘리브레이션 실측: hairline=0.12 / thin=0.96 / medium=1.92pt — 회귀 아님,
   템플릿이 원래 검정 thin (FI·PI 동일, CI·PL은 내부선 없음).

## 계획

- [x] excel_helpers: 프로브 측정부를 `_measure_wrapped_heights`로 추출
      (`autofit_merged_rows` 동작 불변)
- [x] excel_helpers: 순수 함수 `address_row_heights` (좌측 행별 필요 높이 +
      우측 3행 병합 부족분 균등 분배)
- [x] excel_helpers: `layout_address_rows` — wrap 켜기 + 측정 + 높이 쓰기 (OC·FI 공용)
- [x] excel_helpers: `ITEM_GRID_INNER_COLOR = 0xBBBBBB` 상수
- [x] oc_generator: `_fill_header`에서 주소 쓰기 후 `layout_address_rows(ws, 13, ...)`
- [x] oc_generator: `_restore_item_borders`에서 내부 가로선만 #BBBBBB로
      (프레임 = 헤더밴드 하단·마지막 행 하단은 검정 유지)
- [x] fi_generator: 주소만 동일 적용 (`layout_address_rows(ws, 12, ...)`) —
      그리드 색은 보고된 OC만 변경
- [x] tests/test_page_fit.py: `address_row_heights` 순수 테스트 + OC·FI가
      `layout_address_rows`를 부르는지 소스 검사
- [x] 검증: pytest tests/ + OC SOO-2026-0235 재생성 → PDF에서 주소 줄바꿈·회색
      내부선·검정 프레임 실측 + FI 재생성 무회귀 확인

## 리뷰 (2026-08-05)

- **pytest 651 passed, 2 skipped** (기존 641 + 신규 10).
- **OC 재생성** (`OC_SOO-2026-0235_SECTORIEL_260805_1.xlsx/.pdf`): 주소 양쪽 2줄
  완전 렌더(A14 높이 15.95→24.6), PDF 스트림 실측으로 내부선 gray 0.733 /
  프레임(표 상단·Total 위아래) gray 0 확인, 보조 열 Z 잔여물 없음.
- **FI 재생성** (`FI_DNO-2026-0152_...260805.xlsx/.pdf`): 105자 납품 주소가
  G12:I14 안에서 3줄 완전 렌더(행 높이는 47.85pt 안에 들어 불변), 아이템
  그리드는 검정 그대로 — FI·PI 외관은 바꾸지 않음(보고된 OC만).
- 설계 노트: `layout_address_rows`는 값을 쓰지 않고 생성기가 쓴 텍스트를 측정만
  한다. 짧은 주소면 행 높이 불변이라 기존 문서와 픽셀 동일. 페이지 채움 계산은
  헤더 높이를 라이브로 읽으므로(`_page_blank_capacity`) 주소가 자라도 정합.
- 남은 관찰(범위 밖, 기존 동작): 27아이템 OC에서 은행정보·약관 블록이 3페이지로
  넘어가 3페이지에 빈 격자 헤더만 남는 페이지네이션은 이번 수정 전부터 동일.
