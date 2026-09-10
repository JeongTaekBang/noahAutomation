# Changelog

개발 이력, 버그 수정, 리팩토링 기록.

---

## 2026-09-10: 메일 자동 서명 — 억제를 Outlook 설정 쪽으로 (코드 변경 없음)

거래명세표 초안에 개인 서명이 붙는다는 보고. 확인 결과 **코드가 붙이는 서명은 없다** —
`TS_MAIL_BODY`에 서명 문구가 없고 본문도 무서명 구조(`wrap_body_html`) 그대로 나간다.
새 Outlook이 `.eml` 작성창을 열며 끼워 넣는 것이고, 2026-07-29에 실측했던 "이 구조면
안 붙는다"가 클라이언트 갱신으로 수명을 다한 것이다 (그때 남긴 주의가 그대로 현실이 됐다).

실측 (2026-09-10):

- `.eml` 연결 = 새 Outlook(AppX), `olk.exe` 구동 중 — 초안은 새 Outlook 작성창으로 열린다
- 클래식 Outlook COM은 동작(16.0.0.20228)하나 **본문 읽기(`Body`/`HTMLBody`)와 `SaveAs`가
  Object Model Guard에 막혀 `E_ABORT`** — 쓰기와 `Save()`는 정상. 즉 COM으로 만든 메일의
  본문을 되읽어 검증할 수 없다
- 서명은 **인스펙터(작성창)가 만들어지는 시점**에 삽입된다 → 창을 만들지 않는 경로
  (COM `Save()`로 임시보관함에 저장, `--send` 즉시 발송)에는 애초에 붙지 않는다

사용자 결정: 새 Outlook 설정에서 '새 메시지' 기본 서명을 (없음)으로 (회신·전달 서명은 유지).
클라이언트 설정이라 코드 변경 없음.

`wrap_body_html` docstring과 `tasks/lessons.md`의 "이 구조면 서명이 안 붙는다" 단언을
실측에 맞게 정정했다 — 틀린 전제를 남겨 두면 다음 사람이 그대로 물려받는다. 래퍼 자체는
유지한다. 서명이 **인사말 앞으로 끼어드는** 것은 여전히 막고, 거래명세표·납기현황이 같은
구조를 쓰게 하는 것도 이 함수 하나다.

---

## 2026-09-09: SO_통합에 PO_ID·NOAH O.C No. 추가

주문 라인에서 "이건 어느 발주로 나갔나"를 바로 볼 수 있게 PO 시트의 발주번호 2종을
SO_통합 결과에 실었다 (`구분`/`Sales amount` 뒤, `원가_단가` 앞).

- `PO_국내_Select`/`PO_해외_Select`에 `PO_ID`, `NOAH O.C No.` 추가
- `PO_Combined` 그룹화에 콤마 결합 집계 2개 추가 — 분할발주로 한 SO 라인에 PO가
  여럿이면 `ND-0001, ND-0002`처럼 이어 붙인다(`PO_현황`과 같은 규칙).
  `List.RemoveNulls`로 공란 제거 — 안 하면 `Text.Combine`이 빈 값을 구분자와 함께 남긴다
- `WithCostExpanded` 확장 + `Table.ReorderColumns`에 두 컬럼 반영

**Status 필터를 두지 않은 것은 의도**다. 이 블록은 `원가`를 만드는 행 집합 그대로를
쓰므로 "원가는 있는데 PO_ID는 빈" 불일치가 생기지 않는다. 실데이터 확인 결과
취소(Cancelled) PO 56행은 전부 취소된 SO 라인에 붙어 있어(살아있는 SO 라인과 겹침 0건)
필터 유무가 결과를 바꾸지 않는다.

### 검증
임시 워크북에서 변경한 M 블록을 **실제 Power Query로 실행**(마스터 워크북 미접촉):
분할발주 2건 → `ND-0001, ND-0002` / `L260006, L260007`, 같은 PO 2행 → Distinct로 1개,
O.C. 공란 → PO_ID만 표시, 해외 행 결합, 기존 `ICO Unit`·`Total ICO` 집계값 불변.
실 DB 3,357 PO행 시뮬레이션: SO 라인 3,335개 중 콤마 결합 대상 0건(현 데이터는 전부 1:1),
O.C. 공란 237건, 발주 없는 SO 라인 30건(= 두 컬럼 null, 미발주).

⚠️ M 코드는 .xlsx 내부 바이너리라 자동 반영 불가 — Power Query 고급 편집기에 수동 붙여넣기 필요.

---

## 2026-08-07: DN — 세금계산서 발행일은 채우지 않는다

사용자 지시: **그 칸은 직접 입력한다.**

발행 시점은 출고와 별개다 — 월합 거래처는 월말에 끊고, 선발행하는 건도 있다.
그리고 이 날짜의 '월'이 Order Book 매출 인식월을 정하므로 잘못 채우면 없는 달에
매출이 잡힌다. 공란이면 미인식(Backlog 잔류)이라 그런 사고가 없다.

- `build_plan()`의 `fill_tax_date` 인자와 `--no-tax-date` 옵션 제거 (항상 공란)
- GUI의 "세금계산서 발행일 비우기" 체크박스도 제거 — 선택지가 아니라 규칙이다
- **월합 거래처 Remarks 상속은 유지** — 나중에 발행일을 채울 때 어느 날짜를 쓸지
  알 수 있어야 한다 (`25일 마감, 월합세금계산서` 등)

---

## 2026-08-07: DN 쓰기 — 파이썬이 자기 파일을 붙잡고 있어 기록이 안 되던 것 수정

`create_dn.py P08`이 "완료: 4행 추가 (DN_국내 1433~1436행, 표 범위 A1:T1436)"이라고
보고했는데 **파일에는 아무것도 안 들어갔다.** 그 뒤 실행은 열기부터 COM 예외로 죽었다.

원인은 `create_dn.py`가 **읽기용 `pd.ExcelFile` 핸들을 닫지 않은 것**이었다.
그 상태로 Excel에게 같은 파일을 열라고 하니 우리 프로세스가 우리를 막았다 — 실측:

    pandas 핸들 없이        → 열기 성공 (ReadOnly=False)
    pd.ExcelFile 열어 둔 채 → 열기 실패      ← 기존 코드
    그 핸들을 닫은 뒤       → 열기 성공

타이밍에 따라 열기 실패 대신 **`ReadOnly=True`로 열리기도 한다.** 그 상태에서도
쓰기는 메모리에서 멀쩡히 되고 표 범위도 늘어나서 코드가 `A1:T1436`을 읽고 성공이라고
판단하는데, `Save()`는 예외 없이 무시된다(`DisplayAlerts=False`라 '다른 이름으로 저장'
대화상자도 안 뜬다). 첫 실행의 "완료했는데 안 써짐"이 이것이다.

- `dn_recorder.load_source_frames()` — `with pd.ExcelFile(...)`로 **읽기 핸들을 반드시 닫는다**
- **순서 교정: 닫혀 있는지 확인 → 백업 → Excel 열기.**
  Excel이 워크북을 열면 **읽기조차 막혀서**(`open(path,'r+b')`도 `shutil.copy2`도
  PermissionError) 백업 사본을 못 뜬다. 워크북이 이미 열려 있으면 붙어서 쓰지 않고
  닫고 오라고 한다 — 되돌릴 곳 없이 마스터를 건드리지 않기 위해
- COM을 띄우기 전에 `open(path, 'r+b')`로 누가 잡고 있는지 판정한다
- `ReadOnly` 확인으로 **쓰기 전에** 막고, 저장 뒤 파일 mtime·size가 실제로 움직였는지 확인
- 열기·저장이 COM 예외로 죽는 경우도 원문 대신 **행동할 수 있는 안내**로 바꾼다
  (`_busy_message`). Excel은 둘 다 같은 문구('cannot access the file' + 가능성 3가지)로
  돌려줘서 그대로 보여 주면 무엇을 해야 하는지 알 수 없다
- `tests/test_dn_writer.py` 신규 10건 + `test_dn_recorder.py`에 핸들 누수 감시 1건
  (읽은 뒤 파일 rename — Windows는 열린 핸들이 있으면 막는다. 누수 버전으로 실패 확인)

검증: 실제 워크북에 4행 기록 확인 — 표 범위 `A1:T1436`, DN 1,435행, 수식 열
(`PO_ID`·`Customer name`·`Item`·`Unit Price`·`Total Sales`·`AX Project no`) 전부 계산,
피벗 6·쿼리테이블 14 보존.

교훈: **읽고 쓰는 CLI는 읽기 핸들을 넘기지 않는다.** 그리고 마스터 파일에 쓰는 경로에서
"COM이 예외를 안 냈다"를 성공으로 치면 안 된다 — 쓴 결과를 되읽어 확인해야 한다.

---

## 2026-08-07: DN 출고기록 자동 입력 (`create_dn.py`)

공장 출고리스트를 보고 `DN_국내`에 손으로 옮겨 적던 작업을 자동화한다.
연월(`P08`)만 넣으면 추가할 행을 계산해 보여주고, 확인 후 워크북에 써 넣는다.

**라인/수량을 어디서 가져오는가** — 출고리스트는 "어느 주문이 언제 나갔는지"만 알려주고
어느 라인이 몇 개 나갔는지는 없다. `PO_국내`의 `Status`가 그 답이다:

    pending = 이 SO의 PO 라인 중 Status ∉ {Invoiced P01..Pxx, Cancelled}
    pending 없음 → SO_국내 전 라인의 잔량 (= Item qty − 기출고 DN 누계)
    pending 있음 → PO_국내 Invoiced Pxx 라인의 Item qty 그대로

PO만 보면 안 되는 이유: **PO는 부속을 1라인에 합쳐 적는다**
(`SA09X-MA + ADAPTER`, `NA015 (...) / 부싱가공`). SO 2라인이 PO 1라인이라
PO 기준으로만 뽑으면 부속 라인이 통째로 빠진다(6개월 6건, 최대 360,000원).
반대로 `Cancelled` 라인을 안 빼면 취소분이 출고로 잡힌다(같은 기간 2건).

- `po_generator/dn_recorder.py` — 위 규칙 + 자기검증. 단일 날짜는 `Σ PO Total ICO == 계산서금액`.
  **같은 SO가 여러 날 나갔으면 PO 행의 ICO를 날짜별 계산서금액에 배정한다**(`split_by_amount`) —
  `PO_국내`가 분할출고마다 행을 따로 두는 덕에 대개 정확히 나뉜다(2026-03~07 12건 중
  10건 유일 배정, 모호 0건). **안 나뉘는 이유를 구분한다**(`_solve_split`):
  '나눌 방법이 아예 없음' + 합계가 PO ICO와 일치 → 출고는 한 번이고 나머지는
  **단가 정정 행**이라 첫 날 한 건으로 기록한다 (`SOD-2026-0306`: 5/11 계산서 9,956,592 →
  5/13 `L260441-1R`로 103,616인데 그 행은 `AMOUNT`가 10,060,208로 바뀐 **차액**이다.
  출고는 5/11 8개 한 번뿐). '결과가 갈림'은 출고가 여러 번인데 어느 쪽인지 모르는 것이라
  뭉치면 안 되고, 계산서금액이 **음수**인 행(반품)이 섞이면 무조건 사람 몫이다
  (`SOD-2026-0280` 출고→반품→재출고). `JOB NO`의 `-1R`로는 못 가른다 —
  JOB NO는 주문 단위라 정상 분할출고에도 붙는다. COM 의존 없음
- **DN 라인은 `SO_국내`에 살아 있는 것만.** DN은 매출 장부라 PO를 그대로 옮기면 안 된다:
  - SO에 없는 PO 라인 = 매입만 발생 (`sales_only`) — 외주 가공비 같은 것
    (`SOD-2026-0188` L4 'De-cluch Gear Box Bushing 하부 가공' 450,000원, PO 4라인/SO 3라인)
  - SO Status가 `Cancelled`/`Hold`인 라인 — 판매가 취소돼도 공장은 이미 만든 것을
    계산서로 넘긴다 (`SOD-2026-0364` L5: SO는 Cancelled, PO L5 'IP66 TEST 시료 값'은
    Invoiced P07 1,000,000원). 상수 `EXCLUDED_SO_STATUSES`를 `config.py`로 올려
    `delivery_status.py`와 공유 (복사하면 갈라진다)

  실데이터 불변식: DN 1,431행 중 SO에 없는 라인 0건, SO Cancelled 44라인 중 DN에 있는 것 0건.
  금액 대조에서는 두 라인 다 살려 둔다 — 공장은 그것까지 계산서를 끊는다
  (0188 건: 14,610,800원에 450,000원 포함). 빼는 건 DN에 쓸 라인을 고를 때뿐
- `po_generator/dn_writer.py` — xlwings 쓰기. **openpyxl로 저장하면 피벗 3개·파워쿼리 14개가
  통째로 사라진다.** `DN_국내` 표의 계산 열은 6개뿐이고 `Item`·`Unit Price`·`AX Project no`는
  수식이 있어도 계산 열로 등록돼 있지 않아 행을 늘려도 비어 있다 — 마지막 행을 타일 복사해
  상대참조를 물려받은 뒤 `ListObject.Resize`, 값 열만 배치 덮어쓰기(열당 COM 1회).
  쓰기 전 `generated_dn/backup/`에 사본(최근 10개)
- `create_dn.py` — `--dry-run`(미리보기만) / `--yes`(확인 생략) / `--no-tax-date`.
  미리보기 xlsx 2시트(추가분 / 확인필요)
- 세금계산서 발행일은 출고일과 같게 채우되 **월합 거래처**(기존 DN Remarks에서 유도, 12곳)는
  비우고 Remarks를 물려준다 — 그 거래처는 출고 시점에 계산서를 끊지 않는다
- 멱등: `(SO_ID, 출고일)` 중복은 물론 **다른 날짜로 이미 기록해 둔 건**도 건너뛴다.
  키만 보면 못 걸러 지난 기간을 다시 돌릴 때 통째로 중복 입력된다
  (2026-03 `SOD-2026-0156`: 출고리스트 3/23인데 DN은 5/19로 기록).
  이 판정은 **실행 전 시트 상태만** 보고, **라인·수량 조합이 정확히 같을 때만** 건너뛴다.
  실행 중 추가분까지 보면 한 SO가 여러 날 나갈 때 앞 출고 때문에 뒤 출고가 사라지고
  (`SOD-2026-0188`), '라인별 누계 ≥ 필요수량'으로 보면 뒤 회차가 앞 회차 수량에 가려
  사라진다(`SOD-2026-0713`: 7/23 L1 1개를 8/6의 L1 2개가 덮었다). 둘 다 실제로 났다.
  출고 이벤트는 `(SO_ID, 출고일)`로 합쳐 순회한다 — 출고리스트가 같은 주문·같은 날을
  두 줄로 적기도 한다(`SOD-2026-0467`)
- `create_po.bat` `[N]` + 하위 메뉴(미리보기 / 확인 후 추가), `noah_gui.py` DOC_TYPES,
  `build_portable_gui.py` APP_FILES

검증: 실데이터 골든 대조 P03~P08(키 `(SO_ID, Line item, 출고일)`) — 매칭 1124라인에서
**Qty·Currency 불일치 0**. P03·P04·P07·P08은 생성만/정답만도 0으로 **완전 재현**.
확인 필요는 6개월 통틀어 **3건** — 반품이 낀 `SOD-2026-0280` 하나뿐이고 헛경보 0.
남는 차이 4라인은 전부 **출고리스트 날짜와 사람이 적은 날짜가 다른 건**이다
(`SOD-2026-0301` 6/30→7/1·7/20, `SOD-2026-0467` 5/7→5/15) ·
쓰기 E2E는 워크북 사본에서 P08 61행을 지우고 다시 써 넣어 17열 값 일치 + 표 범위 `A1:T1432` +
피벗 6/쿼리테이블 14 보존 확인 · `create_dn.py P08` 재실행 시 추가 0건 ·
pytest 734 passed(신규 `tests/test_dn_recorder.py` 30건 포함)

---

## 2026-08-07: 거래명세표 — 하루치 출고를 거래처별 메일 한 통으로 (`--date`/`--one-mail`)

한 거래처에 하루 여러 PO가 나가는 날이 있다(2026-08-06 씨앤케이엔지니어링 DN 8건,
발주번호 전부 다름). 지금까지는 문서 8장에 **메일도 8통**이었고 y/N도 8번이었다.
문서는 DN별 1장 그대로 두고 **메일만 거래처별 한 통(첨부 8개)** 으로 묶는다.

- `create_ts.py` — `--date`(그날 출고분 자동 선택, `8/6`·`2026-08-06` 등 허용) /
  `--customer`(--date 안에서 거래처 한 곳, 사업자번호·이름 부분일치) / `--one-mail`(명시 ID 묶기).
  `_build_ts_from_dn`·`_build_ts_from_adv` 추출로 단건·묶음이 **같은 생성 경로**를 쓴다.
  묶는 키는 정규화 사업자번호(`group_by_customer`), y/N은 메일당 1회,
  `--merge`(문서 합침)와는 동시 지정 불가
- `mailer.py` — `as_paths()`로 단건/복수 흡수, `export_pdfs()`가 **Excel 1회 기동**으로 N장 변환
  (`export_pdf`은 단건 래퍼). `build_attachments`/`create_document_mail`/`create_ts_mail`이
  경로 목록을 받는다 — OC·납기현황 호출부는 무수정
- **거래처 섞임 관문 신설** (`foreign_biz_numbers`): 한 문서에 두 거래처가 실리면 메일로
  안 나간다. DN 번호 재사용으로 실제로 그런 건이 둘 있다 — `DND-2026-0748`(씨앤케이+오토밸브),
  `DND-2026-0328`(코콘+한일전자). 묶음에서는 그 문서만 첨부에서 빼고 나머지는 보낸다.
  단건 경로도 같은 관문을 지난다(`_mail_ts` 한 곳)
- `utils.BIZ_NO_MIN_DIGITS` — 조회어가 번호인지 이름인지 가르는 기준을 utils로 이동
  (`delivery_status.py`와 `create_ts.py --customer`가 같은 규칙을 쓴다)
- `create_po.bat`(대화형 메뉴) — 거래명세표 하위 메뉴에 `[3] 하루치 묶음 발송`(출고일+거래처 입력)과
  `[4] 묶음 발송`(DN 목록 붙여넣기 → `--interactive --one-mail`) 추가. 거래처명에 `(주)`가 흔해서
  `:delivery_status`와 같은 `if not defined` + 라벨 분기를 쓴다(괄호블록 안에서 전개하면 배치가 죽는다).
  `[3]`에서 **거래처를 비우면 `--mail`** 을 붙여 확인 없이 초안을 띄운다 — 그날 전체는 거래처가 여럿이라
  건건이 y/N을 물으면 5번씩 물어야 한다(사용자 요청, 2026-08-07). 거래처를 지정한 경우는 종전대로 y/N 한 번
- `noah_gui.py` — TS 옵션에 "메일만 거래처별 한 통으로 묶기" 체크박스 (날짜 선택은 CLI/배치 메뉴 전용)

검증: pytest 전체 통과(신규 `tests/test_create_ts_batch.py` 59건 포함) ·
배치 메뉴는 실제 분기부를 떼어내 argv까지 확인(`--customer "(주)삼신"`이 인자 하나로 전달) ·
실데이터 `--date 2026-08-06 --customer 씨앤케이 --no-mail` → 8/8장 생성 ·
그 8장 PDF 일괄 변환 24.2초(Excel 1회) → 첨부 8개짜리 .eml 조립 확인 ·
관문이 `DND-2026-0748`/`DND-2026-0328`을 실데이터에서 잡고 정상 건은 통과.

---

## 2026-08-05: OC "한 페이지 빈 행 채우기" 제거 — 표는 아이템에서 끝난다

8/3에 넣은 페이지 채움(아이템이 적으면 남는 높이만큼 빈 행을 넣어 한 페이지를
완성)이 실물에서 역효과였다 — 1아이템 OC(전체의 절반)가 빈 격자 예닐곱 줄을
달고 나갔고, 회색 내부선 적용 후 그 빈 줄들이 더 도드라졌다(사용자 보고).
표는 마지막 아이템 바로 다음 Total로 끝나는 쪽이 깔끔하다.

- `oc_generator._fill_items` — 부족분 삽입 → 값·행 높이 → **남는 템플릿 행 삭제**
  (빈 행 후보 유지 로직·`_page_blank_capacity` 제거)
- `excel_helpers` — 채움 전용이던 `fit_blank_rows`/`printable_height`/
  `sum_row_heights`/`print_area_last_row`/`A4_HEIGHT_PT` 제거 (OC만 쓰던 스택)
- `tests/test_page_fit.py` → **`tests/test_doc_layout.py`** 개명(git mv) —
  채움 테스트 13개는 기능과 함께 삭제, 남은 내용(주소 높이·상수 불변식·
  생성기 소스 감시)에 맞는 이름으로. CLAUDE.md 참조 2곳 갱신

검증: pytest 전체 통과 · 1아이템 OC(SOO-2026-0239) 재생성 — 아이템 1행 + Total,
빈 행 없음 · 27아이템 OC(SOO-2026-0235) 재생성 — 행 수·합계 불변(원래 채움 미적용).

---

## 2026-08-05: OC 메일 제목·본문 간소화

사용자 결정으로 기본 템플릿 변경 (`config.py`, user_settings로 오버라이드 가능):

- 제목: `[Rotork Controls Korea] Order Confirmation - Your PO: {customer_po}` —
  고객은 자기 PO 번호로 메일을 찾으므로 그것만 밝힌다 (SOO 번호는 첨부의 O.C. No에 있음)
- 본문: `Dear {customer},` + 발주번호 확인 한 줄 +
  `* This email has been sent automatically.*` — 검토 안내문과 서명({supplier_en})은
  제거(발신 조직은 제목이 밝힘), 인사말은 같은 날 재검토로 유지

`tests/test_oc_mail.py`의 템플릿 단언을 새 문구로 갱신 (한글 상호 금지 검사는
제목·본문 양쪽으로 확대). CLAUDE.md·user_settings.py 주석 예시 동기화.

---

## 2026-08-05: OC·FI 주소 잘림 수정 + OC 아이템 그리드 톤 조정

사용자 보고(OC PDF): Customer Address 둘째 줄과 Delivery Address가 중간에서 잘리고,
아이템 사이 가로줄이 너무 두껍게 보인다.

### 1. 긴 주소가 병합 경계에서 잘리던 문제 (OC·FI)
주소 칸은 병합 셀이다 — 왼쪽은 행별 한 줄 병합(`A13:E13` 꼴), 오른쪽 납품 주소는
세 행을 세로로 걸친 병합(`G13:I15` 꼴). **병합 셀은 넘친 텍스트를 옆 칸으로 흘리지
않고 경계에서 클립한다.** SECTORIEL 실측: 74자 bill-to가 A:E 끝에서, 81자 납품
주소가 G:I 끝에서 잘린 채 PDF로 나갔다 — 품목명 잘림(8/3)과 같은 "생성은 성공,
문서는 불량" 부류다.

wrap만 켜서는 안 된다 — 병합 셀은 autofit이 먹지 않아 행 높이가 그대로면 둘째
줄이 세로로 숨는다. `excel_helpers.layout_address_rows()` 신규: wrap을 켜고 보조 열
측정(`_measure_wrapped_heights` — `autofit_merged_rows`에서 추출, 아이템 행과 같은
원리)으로 행 높이까지 확보한다. 왼쪽은 접힌 행만 자라고, 오른쪽 블록이 행 합보다
크면 부족분을 세 행에 균등 분배한다(`address_row_heights`, 순수 함수). 주소가 다
짧으면 높이가 그대로라 기존 문서와 모양이 같다. OC(`ADDR_START_ROW=13`)와
FI(`=12`)가 같은 헬퍼를 부른다 — PI·CI·PL은 주소 셀을 채우지 않아 대상 아님.
`tests/test_page_fit.py`가 순수 계산과 두 생성기의 호출 여부를 감시한다.

### 2. OC 아이템 그리드가 무겁게 보이던 문제
그리드 선은 검정 thin인데, PDF로 나가면 **0.96pt 실선**이다(빈 워크북 캘리브레이션
실측: hairline=0.12 / thin=0.96 / medium=1.92pt). 같은 문서의 상단 규칙선이
**회색 #BBBBBB thin**(gray 0.733)이라 대비 때문에 그리드만 유독 두껍게 읽힌다 —
회귀가 아니라 템플릿이 원래 그랬고(FI·PI 동일, CI·PL은 내부선 없음), OC 메일로
PDF를 보기 시작하면서 드러났다.

`_restore_item_borders()`가 행 사이 **내부 가로선만** 상단 규칙선과 같은
회색(`ITEM_GRID_INNER_COLOR = 0xBBBBBB`)으로 누른다 — COM `Borders(xlInsideHorizontal)`
1회. 표 프레임(헤더밴드 하단·마지막 행 하단·Total)은 검정 유지. 고객이 몇 달 받아온
FI·PI의 모양은 바꾸지 않는다(보고된 OC만).

검증: pytest 651 passed · OC 27아이템(SOO-2026-0235) 재생성 — PDF 스트림에서 내부선
gray 0.733/프레임 gray 0, 주소 2줄 줄바꿈(행 높이 15.95→24.6), 보조 열 잔여물 없음 ·
FI 24라인(DNO-2026-0152) 재생성 — 105자 납품 주소 3줄 완전 렌더, 그리드·행 높이 불변.

---

## 2026-08-03: OC 영문 메일 자동생성 + 해외 문서 5종 PDF 레이아웃 교정

### 1. Order Confirmation 메일 (신규)
거래명세표·납기현황과 같은 흐름(수신자 확인 → y/N → 초안/발송)을 OC에 붙였다.
받는 쪽이 해외 고객이므로 **본문은 영문**이고, 서명은 `{supplier}`(한글)가 아니라
새로 만든 `{supplier_en}`을 쓴다.

**수신자 조인키가 국내와 다르다.** `SO_해외`의 `Business registration number` 컬럼에는
실제로 `C-0054` 같은 **고객코드**가 들어 있고(이름만 국내 시트와 같다), 이것이
`Customer_해외.C-code by 해외`와 맞물린다. 실데이터 검증: SO_해외의 고유 코드 15개가
`Customer_해외` 130개에 전부 매칭.

- `mailer.find_recipient_overseas()` 신규 — 기존 `find_recipient()`(국내)와
  `_build_recipient()` 하나를 공유한다. 국내 경로는 동작 무변경
- `utils.load_customer_overseas()` + `normalize_customer_code()` 신규.
  국내 로더와 `_load_customer_master()`를 공유(조인키만 다르고 나머지가 같다)
- `mail_cli.prepare_mail_options()` / `collect_customer_po()` — `create_ts.py`에 있던 것을
  올려 3개 CLI가 공유. `MailOptions`에 `loader`/`sheet_label`을 주입해 해외 마스터도 같은
  지연로딩·1회 안내 규칙을 탄다
- `noah_gui.py` — OC에 "메일 초안 만들기" 체크박스

**주의:** `Customer_해외`의 이메일 컬럼은 현재 198행 전부 비어 있다. 채우기 전에는
"수신자 미등록"으로 건너뛰고 문서만 생성된다(설계대로).

**함정 — `MailOptions.loader` 기본값에 함수를 박으면 안 된다.** 클래스 정의 시점에
함수 객체가 굳어 모듈 속성을 갈아끼우는 쪽(테스트 monkeypatch)이 조용히 무시된다.
실제로 그렇게 만들었다가 기존 테스트 3건이 깨졌다. `None` 기본값 + 호출 시점 해석으로 고침.

### 2. 긴 품목명이 잘리던 문제 (OC·FI·PI·CI·PL)
품목명 칸은 A:D **병합**인데 Excel의 `rows.autofit()`은 병합 셀을 측정에서 제외한다.
105자 품목명에 대해 행 높이가 15pt → **12.75pt로 오히려 줄면서 1줄로 잘렸다**(실측).
생성은 성공으로 끝나므로 PDF를 열어보기 전엔 알 수 없었다.

`excel_helpers.autofit_merged_rows()` — 인쇄영역 밖 보조 열(Z)에 같은 텍스트를 넣고
Excel에게 재게 한 뒤 그 높이를 되쓴다. 폰트 폭을 코드로 추정하지 않는다.
보조 열 폭은 **문자 단위가 아니라 포인트로** 맞춘다 — 열마다 안쪽 여백이 붙어
A:D(198.00pt)와 같은 문자폭의 단일 열(186.75pt)이 11.25pt 어긋나고, 그대로 두면
36자짜리 품목명이 실제로는 한 줄에 들어가는데도 두 줄로 재어졌다.

### 3. 삽입한 행이 병합을 잃던 문제 (기존 결함)
48아이템 OC를 만들어 보니 삽입한 41행 중 **6행의 A:D 병합이 사라져** 품목명이 A열
하나에 갇혀 5~7줄로 흘렀다. 행 높이 교정 이전부터 있던 결함이고, 높이를 재기 시작하면서
드러났다. `ensure_row_merges()`가 범위를 통째로 풀고 `Merge(Across:=True)`로 행마다 다시
병합한다(COM 2회). 병합 보장과 높이 교정은 늘 세트라 **복합 헬퍼 `layout_item_rows()`**
하나로 묶었고, 5종 생성기는 값 채우기 직후 이것만 부른다 — 여섯 번째 문서가 한쪽만
부르는 실수를 구조로 차단 (`tests/test_page_fit.py`가 감시).

병합 범위 `ITEM_NAME_MERGED_COLS`는 `excel_helpers` 한 곳에만 둔다 — **CI와 PL은
선적서류라 늘 같이 첨부되어 나란히 읽히므로** 줄 높이 규칙이 갈리면 바로 눈에 띈다.
실측 확인: 33아이템 DN에서 CI·PL의 행 높이 분포가 `[26, 38, 50]`으로 동일.

### 4. 아이템이 적으면 한 페이지를 채운다 (OC 전용)
기존에는 아이템이 템플릿 행 수(7)보다 적으면 남는 행을 **삭제**해, 가장 흔한 1아이템
문서가 표 한 줄 + 큰 여백 + 하단 은행정보 블록이 되었다(SO_해외 236건 중 118건이 1아이템).
이제 남는 높이만큼 빈 아이템 행을 채운다.

계산은 COM과 분리된 순수 함수 `fit_blank_rows()`가 하고, 용지·여백·행 높이는 런타임에
`PageSetup`에서 읽는다(템플릿을 고쳐도 따라간다). **행 높이 교정 뒤에** 계산해야 한다 —
긴 품목명이 3줄을 먹으면 들어갈 빈 행 수가 줄기 때문. 이미 한 페이지를 넘긴 문서
(48아이템 등)는 채우지 않는다.

OC 템플릿 실측: 인쇄가능 733.89pt = 헤더 265.00 + 아이템 + 하단 335.25
→ 15pt 행 기준 한 페이지에 8행.

> 템플릿에 `FitToPagesWide/Tall=1`이 있지만 `Zoom=100`이라 비활성이다. **활성화하지 않았다**
> — 48아이템 OC가 한 장으로 짓눌려 판독 불가가 된다.

### 5. 리팩터 (같은 날, 4-각도 리뷰 후)
재사용·단순화·효율·설계깊이 리뷰를 돌려 나온 지적을 반영했다.

**COM 왕복 배치화** — 새 레이아웃 코드가 1아이템 OC 기준 왕복을 ~30→~200회로 늘렸던 것을
배치로 되돌림: 행 높이 합은 다중 행 `Range.Height` **1회**(행별 RowHeight 루프 금지 —
`sum_row_heights` docstring), 높이 되쓰기는 같은 높이 연속 구간당 1회 + 균일하면 읽기도
1회(다중 행 `RowHeight`는 균일할 때만 값을 준다), 행 삽입은 `insert_copied_rows`
(일괄 Insert + **타일 Copy** — Destination이 크면 원본이 반복된다, 행당 2회→총 3회).

**OC 단일 델타 흐름** — "남는 행 삭제 → 빈 행 재삽입"(1아이템 문서가 6행 지우고 7행
되삽입)을 없앴다. 남는 템플릿 행을 빈 행 후보로 남겨 두고, 최종 표시 행 수를 정한 뒤
부족분 삽입/초과분 삭제를 **한 번만** 한다. 페이지 기하를 못 읽으면(`print_area_last_row`
→ None) 추정하지 않고 채움을 접는다 — UsedRange 폴백은 인쇄영역 밖 잔여 셀을 집어
하단 블록을 부풀리므로 두지 않는다.

**메일 관문 통일** — 수신자 조회→표시→y/N 확인 흐름이 세 CLI에 복사돼 있던 것을
`mail_cli.confirm_recipient()`로 모았다. `delivery_status.py`도 `prepare_mail_options()`를
타게 해 "3개 CLI 공유"라는 문서 주장과 코드를 일치시켰다(배너·.eml 경고가 DS에도 적용).
날짜 표기는 `format_mail_date()` 한 곳. 행에서 조인키를 읽는 지식은 mailer로 —
`find_recipient_overseas_for_order()`.

**이름·별칭 정리** — `Recipient.biz_no`가 해외에서 고객코드('C-0054')를 담게 되어
`customer_key`로 개명(읽는 곳 0곳 실측 후). `COLUMN_ALIASES['customer_code']`에서
`'고객코드'` 제거 — `Customer_해외`의 그 컬럼은 AX 번호라 **별칭이 아니라 함정**이었다
(전 소비처 5곳 추적으로 확인). 5종 생성기의 중복 `_to_text`는 `utils.to_text`로 통일.

### 검증
- `pytest tests/` 전체 통과 (신규 `test_oc_mail.py`, `test_page_fit.py` 포함;
  중복 테스트는 단일 소유로 정리 — 백엔드 격리 픽스처는 conftest로)
- 1아이템 OC → 7행 채움·1페이지, 89자 품목명이 PDF에 3줄 전부 출력(PDF 텍스트 추출로 확인)
- 48아이템 OC → 3페이지, 병합 깨짐 0, 빈 행 채움 없음, 합계 불변
- FI/CI/PL 33아이템 → 병합 깨짐 0, Total 위치·합계 불변 — 리팩터 전후 동일
- `.eml` 초안 조립 → 영문 제목/본문 + PDF 첨부 + `X-Unsent: 1` 확인 (리팩터 후 재확인)

---

## 2026-07-31: 대시보드 사내 배포판 — 죽어 있던 빌드를 재작성

### 증상
`dashboard_dist/`에 배포 구조가 있었지만 **빌드가 실행되지 않았다.** 동봉된 `dashboard.py`는
220KB(3/27), 원본은 292KB(7/30) — **4개월 낡은 상태**로 Order Book 매출인식 개편과
동기화로그 v2 페이지가 통째로 빠져 있었다.

### 원인
`build_dist.py`가 원본 `dashboard.py`의 import 줄을 **문자열 치환**해 `po_generator` 의존성을
끊는 방식이었다. 치환 대상이 한 줄짜리였는데 원본이 괄호 다중 import로 바뀌면서
(`ensure_so_change_ack_table`·`get_sync_metadata`·`SYNC_LOG_CHANGE_TYPES`) 패턴이 안 맞아
`sys.exit(1)`. 게다가 `.gitignore`가 `dashboard_dist/`를 통째로 제외해 **빌드 레시피가
git에 없었고**, 그래서 깨진 것도 낡은 것도 아무도 몰랐다.

### 수정
소스 재작성을 **없앴다**. `cli_dist`처럼 `po_generator/`를 동봉해 `dashboard.py`를 무수정으로
돌린다 — import가 늘어도 빌드가 안 깨진다. `build_dist.py`·`build_portable.py`·`launcher.py`·
`dashboard_config.ini` 삭제, 설정은 `noah_config.ini`로 통일.

- `build_common.py` — 두 배포판이 공유하는 빌드 공통부 (런타임·핀 3자 대조·트리밍·zip)
- `po_generator/__init__.py`의 재export 제거 — `from po_generator.config import ...` 한 줄에
  pandas·openpyxl·xlwings(→COM)가 딸려오던 것을 끊었다. 그 이름들을 패키지 루트에서 쓰는
  코드는 저장소 전체에 0건이었다. 덕분에 대시보드 배포판에서 Excel 라이브러리가 빠지고
  모든 CLI 기동도 빨라진다
- `verify()`가 streamlit을 **실제로 띄워** 실제 DB 사본으로 페이지를 받아 본다 —
  트리밍을 246MB 했으므로 import 검사만으로는 부족하다
- `.gitignore` — `dashboard_dist/` 통째 제외 → `*.zip`만 제외 (레시피를 커밋한다)

### WAL — OneDrive 공유의 지뢰
DB는 `journal_mode=wal`이고 OneDrive로 공유한다. WAL이면 최신 커밋이 `-wal` 사이드카에 먼저
들어가는데 OneDrive는 본체와 사이드카를 **각각 따로** 올린다. 원본을 직접 열면 커밋이 빠진
상태를 보거나, 읽기 잠금 탓에 OneDrive가 파일을 교체 못 해 **"충돌된 사본"** 이 생긴다.
둘 다 "사람마다 숫자가 다르다"로 뒤늦게 드러나는 형태다.

- 발행: `sync_db.py: checkpoint_wal()` — 동기화 성공 후 `PRAGMA wal_checkpoint(TRUNCATE)`
- 소비: `dashboard.py: _db_snapshot()` — 백업 API로 일관된 사본을 떠서 그것만 연다.
  캐시 키가 `(mtime, size)`라 원본이 갱신되면 자동으로 다시 뜬다
- 표시: 사이드바 "데이터 기준" — 하루가 넘으면 caption이 아니라 경고로 올린다

### 크기 (최적화)
설치 직후 런타임에서 **246MB 트리밍**. 근거는 전부 실측이다.
- `pyarrow` 84MB → flight·parquet·dataset·substrait·acero·gandiva·tests·include·src 제거.
  `import pyarrow`는 lib/ipc/types/util만, streamlit은 여기에 `_compute`를 더 얹으므로
  **`arrow_compute.dll`은 남긴다**
- `plotly` — `labextension`(JupyterLab) + `package_data`(오프라인 HTML용 plotly.min.js).
  `st.plotly_chart`는 streamlit 자기 번들로 그린다. `write_html`·`plotly.offline` 사용 0건
- `pydeck/nbextension`(Jupyter 자산), `tkinter`·`tcl`(대시보드는 GUI 창이 없다)

### 배포 방식에 대한 판단 (실측 근거)
중앙 호스팅(streamlit 한 대 띄우고 URL 공유)이 갱신 문제를 근본적으로 없애지만, 회사 PC에서
**불가**로 확인됐다: 관리자 권한이 없어 방화벽 인바운드 규칙 생성이 `Access is denied`,
Wi-Fi DHCP라 URL 고정 불가, AC 전원 절전 5분. IT 지원 없이는 폴더 배포가 유일한 선택지다.

---

## 2026-07-30: DN 테이블 PK — 분할출고 중복키 행이 조용히 사라지던 버그

### 증상
Excel `DN_국내`는 1,370행인데 SQLite `dn_domestic`은 **1,368행**. Order Book 환율 작업 중
Excel↔SQLite 대사에서 **6월 국내 매출 18,656,000원 + 수량 19**가 비어 발견됐다.

### 원인
동기화 PK가 `(DN_ID, SO_ID, Line item)`이었는데 이 조합이 **유일하지 않다**. 같은 DN 문서·같은
SO 라인을 두 행으로 나눠 적는 **분할출고가 정상 데이터로 존재**한다. upsert라 뒤 행이 앞 행을
덮어써 한 행이 소리 없이 사라졌다.

```
DND-2026-0511 / SOD-2026-0232 / 2   NOS160-MS-FC   Qty 11 (18,656,000) + Qty 12 (20,352,000)
DND-2026-0560 / SOD-2026-0301 / 9   Eye bolt       Qty 8 + Qty 212 (금액 0)
```

`PO_국내`/`PO_해외`는 이미 같은 이유로 `_row_seq`를 PK에 넣어 뒀는데 DN은 빠져 있었다.

### 수정
`dn_domestic`/`dn_export` PK를 `(DN_ID, SO_ID, Line item, _row_seq)`로, PO와 동일한 처방
(`needs_row_seq=True`, `row_seq_group=(DN_ID, SO_ID, Line item)`).

`migrate_pk_if_changed`가 PK 변경을 감지해 기존 테이블을 `_bak`으로 백업하고 재생성한다.
재적재는 전 행이 '신규'로 잡히지만 데이터가 들어온 게 아니라 키 체계가 바뀐 것이므로,
`SheetSyncResult.pk_migrated`를 두고 `write_sync_log_to_db`가 **'재적재' 1건만 기록**한다
(2,146건의 가짜 '신규'로 변경 이력을 덮지 않는다 — 재키잉 억제와 같은 취지).

### 영향 (개선)
- SQLite Order Book이 Excel `Order_book`·`AX_매출대사`와 **14개 월×구분 조합 전부 일치**
  (2026-06 국내 679,679,520 → **698,335,520**)
- 대시보드 출고 상태 판정 교정: `SOD-2026-0232/2`가 **부분 출고(12/54) → 출고 완료(54/54)**.
  DN 수량 누계로 판정하는데 한 행이 빠져 미출고로 잡히고 있었다
- 고객 발송물은 무영향 — `delivery_status.py`·`reconcile_so.py`·문서 생성 CLI는 Excel을
  직접 읽으므로 애초에 정확했다. SQLite 기반 산출물(대시보드·Order Book·스냅샷)만 달라진다

### 테스트
`tests/test_db_sync_rekey.py` 5건 추가 — 국내/해외 중복키 보존, 재동기화 시 '동일' 판정,
재적재 로그 1건 억제, 일반 동기화는 종전대로 행별 기록. 전체 **561 passed / 2 skipped**.

---

## 2026-07-30: Order Book — 환율 임팩트(Value Variance) + 매출 귀속월 통일

### 배경
Order Book Output이 `AX_매출대사`와 두 축에서 어긋나 있었다.

1. **해외 환율** — Output이 DN 시트의 `Total Sales KRW`(수주시점 환율)였다.
   예: `SOO-2026-0188` USD 1,276을 6월에 수주(환율 1,500.036)하고 7월에 선적 →
   매출은 7월 환율(1,548.608)로 인식돼야 하는데 6월 환율 금액이 Output에 잡혔다.
   P07 해외 기준 **47,662,819원** 차이 (누적 P01~P07 115,496,363원).
2. **국내 귀속월** — Output이 출고월이었고 `AX_매출대사`는 세금계산서 발행월.
   금액이 사라지는 게 아니라 월이 밀렸다 (P04 −86.5백만 / P05 +84.4백만).

### 수정
**Output = 매출 인식 기준으로 통일** (`AX_매출대사`와 동일 산식):
- 국내 = 세금계산서 발행월 → `N/A`(발행 불필요)면 출고월 → 선수금+출고면 출고월 → 없으면 미인식
- 해외 = 선적월, KRW는 `외화금액 × 선적월 환율`로 DN 라인 단위 재환산

**Variance = 환율 재평가분** (`재환산액 − 시트 KRW`)을 도입해 `Ending = Start + Input −
Output + Variance`가 **재환산 전과 완전히 동일**하게 유지된다(불일치 0건 실측).
환율 노이즈가 Variance로 빠지므로 "Ending ≠ 0 = SO-DN 금액 불일치" 진단이 살아 있다.

### 범위
- Power Query `Order_Book` (`docs/POWER_QUERY.md`) — FX 언피벗 + 재환산 + Variance,
  `출고월` → `매출월` 개칭
- SQLite: `fx(currency, ym, rate)` 테이블 + **`v_dn_revenue` 뷰**. 같은 DN 산식이 6곳
  (sql 4종 + `snapshot.py` + `dashboard.py`)에 복제돼 있어 뷰 하나로 모았다
- `FX` 시트 언피벗 동기화(`db_sync._sync_fx`) — 환율 변경을 `_sync_log`에 신규/수정/삭제로 기록.
  월 컬럼(`YYYY-MM`)을 못 찾으면 **에러로 세운다** (fx가 비면 뷰가 조용히 시트 KRW로 폴백해
  숫자가 틀린 채 돌아간다)
- `sql/order_book_snapshot*.sql`·`snapshot.py`는 환율 재평가분을 기존 소급변경 Variance에
  합산 (AX Order Book처럼 조정분 단일 컬럼)

### 검증 (2026-07-30 실측)
- **P01~P07 × 국내/해외 14개 조합 전부 `AX_매출대사`와 차이 0원**
- 귀속월 변경의 Backlog 영향: 국내 **+4개 / +1,232,420원**(월합세금계산서 대기 4라인),
  해외 변동 없음. 그 외 전 그룹 Ending 동일
- 결측 방어: 선적월 환율·외화금액이 없으면 시트 KRW 폴백(Variance 0) → Output이 0으로
  사라지지 않는다
- 테스트 21건 신설 (`tests/test_order_book_fx.py`), 전체 556 passed / 2 skipped

### 알려진 잔여 차이
- **DN 테이블 PK**: `(DN_ID, SO_ID, Line item)`이 겹치는 정상 분할출고 2행이 SQLite에서
  조용히 1행으로 덮어써진다 → SQLite 쪽 6월 국내가 18,656,000원 적다. Excel Order_book은
  정상. PO 테이블처럼 `_row_seq`를 PK에 넣어야 하는 별건 이슈
  **(→ 당일 후속 작업으로 해결 — 위 "DN 테이블 PK" 항목 참조)**
- Power Query `Number.Round`(짝수 반올림)와 SQLite `ROUND`(올림)가 정확히 .5인 2라인에서
  갈려 월 합계 최대 ±1원 차이. `AX_매출대사` 일치가 목적이라 Excel 기준 유지

---

## 2026-07-30: 메일 auto 백엔드 — 초안은 .eml, 즉시 발송만 COM

### 배경
배포판 GUI 테스트에서 거래명세표 메일 초안이 **클래식(옛) Outlook** 창으로 떴다.
원인: `auto` 판정이 "COM 되면 COM"이었는데, 이 PC에 클래식 Outlook이 설치·구성되면서
(7/29까지는 없어서 `.eml` 경로였음) COM이 살아났고, COM은 항상 클래식 창을 띄운다.
사용자가 평소 쓰는 건 새 Outlook — 초안이 낯선 옛 창으로 뜨는 회귀.

### 수정 (`mailer.resolve_backend`)
`auto`의 의미를 용도별로 분리:
- **초안** → 항상 `.eml`. OS 연결 프로그램을 따르므로 그 자체가 "사용자가 평소 쓰는
  메일 앱"이다. 새 Outlook이 X-Unsent `.eml`을 편집 가능한 초안으로 여는 것 실측 확인
- **즉시 발송(--send)** → COM 가능할 때만 OUTLOOK (창 없이 보내는 유일한 방법),
  없으면 `.eml` 초안으로 강등 + 안내 (기존과 동일)

명시 설정(`TS_MAIL_BACKEND='outlook'|'eml'`)은 여전히 auto 규칙보다 우선.
회귀 테스트: `test_auto_draft_is_eml_even_with_com` (COM이 있어도 초안은 .eml).

---

## TODO (미완료 항목)

### 템플릿 확장
- [x] Proforma Invoice (PI) 구현 완료 ✓
- [x] Final Invoice (FI) 구현 완료 ✓
- [x] Order Confirmation (OC) 구현 완료 ✓
- [x] Commercial Invoice (CI) 구현 완료 ✓
- [x] Packing List (PL) 구현 완료 ✓

### SQL 기반 데이터 분석
- [x] NOAH_SO_PO_DN.xlsx → SQLite DB 동기화 구현 완료 ✓
  - **배경**: Excel 형식의 데이터 유실/변형 취약점 → SQLite 백업
  - DuckDB 분석 연동은 추후 확장 예정

---

## 2026-07-30: cli_dist 전면 정비 (재현성·다이어트·버전·제거)

### 배경
배포판 빌드를 점검하니 네 가지가 나왔다. ① 직접 의존성 4개만 핀이라 transitive가
빌드마다 떠다녔다(배포판 numpy 2.4.6 vs 개발 env 2.4.3 실측 — "여기서 테스트한 그대로"가
거짓). ② 설치.bat의 robocopy에 `/R /W`가 없어 기본값(재시도 100만×30초) — 앱 실행 중
업데이트하면 무한 대기로 보인다. ③ pip·pythonwin(IDE)·numpy tests 등 런타임 미사용
~60MB가 실려 나갔다. ④ 버전 정체성이 없어 "어느 빌드 쓰세요?" 문의에 답할 수 없었다.

### 구현 (cli_dist/build_portable_gui.py, requirements.txt, noah_gui.py)
- **핀 3자 대조**: requirements.txt에 transitive까지 전부 `==` 고정(개발 env 기준).
  빌드가 핀 = 배포 런타임 설치본 = 개발 env를 대조하고 하나라도 어긋나면 실패
  (PEP 503 정규화로 `et_xmlfile`/`et-xmlfile` 표기 차이 흡수, `parse_pins`는 `>=` 거부)
- **트리밍 확장**: pip/setuptools(+`distutils-precedence.pth`·`_distutils_hack` 세트 —
  고아 .pth는 매 실행 stderr 오류), pythonwin, numpy tests·distutils·f2py, ensurepip,
  Scripts, tix, include, 지운 패키지의 dist-info(유령 패키지 보고 방지).
  트림 루프에 디렉터리 분기 추가 (기존은 파일만 지웠다)
- **CLI 전수 스모크**: verify()가 8종 CLI를 `--help`로 실행 — 트리밍/의존성 문제가
  받는 사람 PC가 아니라 빌드에서 터지게
- **버전 스탬프**: `날짜+git sha[.dirty]` (예: 2026.07.30+326b87e) → `BUILD_INFO.txt`
  (기계가 읽는 건 `version:` 줄 하나 — `noah_gui.read_build_info`), GUI 타이틀·설치.bat
  헤더·README에 표시
- **설치.bat**: `/R:1 /W:1` + 실패 시 "실행 중이면 닫고" 안내
- **제거.bat 신설**: 설치 폴더+바탕화면 바로가기 삭제. 자기 자신이 삭제 대상 안에 있어
  `%TEMP%` 복사 후 `start`(비동기 — `call`이면 부모 cmd의 CWD가 폴더를 잡아 rd 실패).
  taskkill은 안 쓴다(무관한 python까지 잡는다) — 잠기면 안내만
- step 번호 자동화 (`itertools.count` — 하드코딩 "N/7" 제거)

### 설치 왕복 테스트가 잡은 잠복 버그: 바로가기가 en-US 로캘에서 생성 실패
설치.bat의 WScript.Shell(WshShortcut) 바로가기 생성이 **이 PC에서 처음부터 실패**하고
있었다 (2026-07-28 배포판 최초 구축 때 타 PC 검증을 안 해 놓친 것). 회사 PC는 시스템
로캘이 en-US(ANSI CP1252)인데, WshShortcut은 경로를 ANSI로 변환해서 한글("바탕 화면"
KFM 경로, 바로가기 이름)이 전부 `?`가 되어 E_INVALIDARG로 죽는다.
PowerShell `-EncodedCommand`로 인자 전달을 고쳐도 COM 내부에서 다시 깨진다 (실측).

→ 배치의 PowerShell을 버리고 **동봉된 python의 pywin32 `IShellLinkW`**로 교체
(`_make_shortcut.py` — 빌드가 생성해 app/에 동봉, 설치/제거.bat이 호출).
전 구간 유니코드라 로캘 무관, 바탕화면 경로도 `SHGetFolderPath`로 KFM을 따라간다.
`win32com.shell` import를 verify()의 런타임 검사에 추가해 트리밍으로부터 보호.

검증(이 PC에서 실제 왕복): 설치 → 바로가기 생성·대상 확인(IShellLinkW로 판독) →
앱 켠 채 재설치 → **1초 만에** "실행 중이면 닫고" 안내(기존 기본값이면 사실상 무한 대기)
→ 제거.bat → 폴더·바로가기 모두 삭제 확인.

---

## 2026-07-30: 배포판 GUI에 납기현황 추가

### 배경
납기현황 회신(`delivery_status.py`)은 CLI로만 쓸 수 있었다. 사내 배포판에는 파일 자체가
들어가지 않았고(`build_portable_gui.APP_FILES`), GUI에도 항목이 없었다. 정작 이 기능을
가장 자주 쓸 사람은 Python이 없는 PC에서 배포판을 쓰는 영업 담당자다.

### 구현
- `noah_gui.py` — 문서 종류에 **납기현황** 추가 (8번째, 라디오 그리드가 4×2로 정확히 참).
  거래처 조회 / 거래처 목록(`--list`) 라디오, `--all`·`--mail` 체크박스
  - 유일하게 문서 ID가 아니라 **거래처**로 조회하고, CLI가 한 곳만 받으므로
    `multi: False` 표식을 두고 여러 줄 입력을 messagebox로 막는다
    (그냥 넘기면 argparse `unrecognized arguments`로 죽는다)
  - `build_command`의 메일 플래그를 ts 전용에서 "`mail` 옵션이 있는 문서"로 일반화
- `cli_dist/build_portable_gui.py` — `APP_FILES`에 `delivery_status.py`,
  README에 납기현황 사용법. 의존성 추가는 없다 (첨부가 xlsx라 PDF 변환용 Excel COM 불필요)

### 재발 방지
GUI에 문서를 추가하고 `APP_FILES`에 넣는 걸 잊으면 **배포판에서 그 버튼만 조용히 실패**한다.
판정은 `noah_gui.missing_scripts()` 한 곳이 소유하고, 두 시점에서 막는다.
- 빌드 `verify()` — 복사된 앱에서 `DOC_TYPES`의 모든 script가 존재하는지 대조 (zip 만들기 전)
- `tests/test_noah_gui.py::test_doc_scripts_are_packaged` — DOC_TYPES ⊆ APP_FILES

같은 패턴의 대조 테스트 하나 더: GUI가 만든 ds 명령을 `delivery_status.create_argument_parser()`에
그대로 통과시킨다 — `multi: False`는 CLI `customer`(nargs='?')의 미러라서, CLI 인자 정의가
바뀌어 미러가 어긋나면 테스트가 잡는다.

---

## 2026-07-29: 거래명세표 메일도 서명 없이 (무서명 구조를 공용 래퍼로)

### 배경
납기현황 메일은 본문을 `<table><tr><td>` 한 칸에 담는 구조인데, 이 구조에서는 Outlook이
.eml 초안에 **자동 서명을 붙이지 않는 것**이 실측으로 확인됐다. 사용자 결정: 서명 없는 쪽이
낫다 (본문이 이미 "본 메일은 자동 발송된 메일입니다."로 끝난다). 거래명세표 메일은 아직
인라인 `<br>` 본문이라 서명이 끼던 상태 — 같은 구조로 맞춘다.

### 구현
- `mailer.wrap_body_html()` 신설 — 단일 최상위 블록 래퍼를 한 곳에 정의.
  왜 이 구조인지(서명 삽입 위치 실측 이력)도 여기 docstring에 남긴다
- `body_to_html()`(TS 기본 경로)이 래퍼를 쓰고, 납기현황 `build_html_body()`는
  하드코딩돼 있던 래퍼를 같은 함수로 위임 — **구조 정의가 한 곳**이라 두 메일이 갈릴 수 없다
- 회귀 테스트: TS `.eml`의 `<body>` 직계 자식이 1개인지 + `body_to_html == wrap_body_html(...)`

### 주의
서명 억제는 Outlook의 관찰된 동작에 기대는 것이라 보증은 아니다. Outlook 업데이트로
동작이 바뀌면, 확실한 해법은 자동 서명을 끄고 서명 텍스트를 본문 템플릿에 직접 넣는 것.

---

## 2026-07-29: 납기현황·메일 경로 정리 리팩터 (동작 불변)

하루 동안 피드백을 반복 반영하며 쌓인 중복·군더더기를 걷어냈다.
**고객에게 나가는 산출물은 바이트 단위로 동일** — HEAD와 리팩터본을 같은 실데이터에 돌려
파생 2,023행 전부 + 43개 거래처의 요약/상세/평문표/HTML표를 비교, 불일치 0건.

- `mail_cli.py`에 `show_recipient()` / `report_mail_result()` 추가 — 수신자 표시와 결과 보고가
  `create_ts.py`·`delivery_status.py`에 자구까지 동일하게 복사돼 있었다. 오발송 차단 UX가
  CLI마다 갈리면 확인 습관도 갈리므로 한 곳으로 (mail_cli의 존재 이유와 같은 논리)
- `attach_ship_status()` — 수동 센티널 컬럼(`_dn_match=True` + `notna()` 트릭)을
  `merge(indicator=)`로, 행 단위 `apply`를 `np.select`로. 빈 프레임 특례 코드도 함께 삭제
- `print_summary()` — 콘솔 전용 표를 없애고 **메일 평문 표를 그대로** 출력.
  보내기 전 화면에서 확인하는 표가 실제 나가는 표와 달랐다(요청납기·`*` 지연 표식이
  콘솔에서 안 보였음). 이제 화면 = 발송본. 합계는 한 줄로
- 모듈 docstring의 분할 납기 설명을 현행(요청납기+출고예정일 기준)에 맞춤, 죽은 빈 줄 제거

검증: `pytest tests/` 500 passed, 2 skipped + 실데이터 동등성 비교 (위)

---

## 2026-07-29: 납기현황 메일 회신 (본문 표 + xlsx 첨부)

### 배경
납기현황 xlsx를 만든 뒤 거래명세표와 같은 방식으로 고객에게 바로 보낸다.
조회 자체가 사업자번호 기준이라 수신자 조회가 TS보다 단순하다 — 조회 키를 그대로 쓰면 된다.
본문에 표를 싣는 이유는 **원래 손으로 회신하던 형식이 그랬기 때문**이다.
첨부만 보내면 고객이 파일을 열어야 답을 볼 수 있어 후퇴다.

### mailer.py를 문서 중립으로
TS 전용 상수가 함수 안에 박혀 있어 재사용이 안 됐다. 고정된 것만 인자로 뺐다.
- `create_document_mail()` 신설 — 제목/본문 템플릿, 첨부형식, 고정 CC, HTML 본문,
  초안 파일명 접두사, 로그 라벨을 전부 받는다
- `create_ts_mail()`은 TS 상수를 넘기는 얇은 래퍼로. **TS 동작·호출부 불변**
- `render_template(..., extra=)` / `find_recipient(..., fixed_cc=)` / `build_eml(..., prefix=, body_html=)`
- Outlook COM에 `HTMLBody` 경로 추가. `HTMLBody`와 `Body`는 동시에 쓰면 나중 대입이 앞을 지우므로
  표가 있는 문서는 HTML만 설정한다
- `_body_to_html` → `body_to_html` (모듈 밖에서 쓰게 됨)

### mail_cli.py 신설
`MailMode`·`MailOptions`·`resolve_mail_mode`·`confirm`·`add_mail_arguments`를
`create_ts.py`에서 `po_generator/mail_cli.py`로 옮겼다. CLI 두 개가 똑같이 필요한 배선인데
복사하면 반드시 갈라지고, **그 갈라짐이 고객에게 메일이 나가는 경로에서** 벌어진다.
`create_ts.py`는 재노출만 해서 기존 import를 깨지 않는다.

### 구현
- 공장 출고일 컬럼명은 **`NOAH 공장 출고 예정일`**. `EXW NOAH`는 확정 약속이 아니라 계획이고,
  실제로 요청납기 초과가 164행 중 23행이다. '출고일'로 나가면 그날 안 나갔을 때 약속을 어긴 것이 된다.
  본문·첨부 둘 다 같은 고객에게 가므로 한쪽만 바꾸면 완충 표현이 무너져 함께 바꿨다.
  파일 안 9곳에 흩어져 있던 문자열은 `COL_EXW` 상수로 묶었다
- 표 컬럼: Customer PO / Remarks / 수량 / **Requested delivery date** / NOAH 공장 출고 예정일 /
  **Sales 금액** / PO receipt date. 요청납기를 공장출고일 **바로 왼쪽**에 둬서 한눈에 대조된다
  (엔이에스 14건 중 2건이 요청보다 늦다: `2026-07-15-CH-07` 9/14 요청 → 10/22 출고).
  본문 표는 여기서 PO receipt date만 뺀다 (고객이 이미 아는 날짜)
- **행 분리 기준을 두 날짜 모두로.** 기존엔 `EXW NOAH`만 갈리면 나눴는데, 요청납기가 달라도 나눈다.
  대표값(가장 이른 날)으로 접으면 나머지 약속이 사라진다 — 피엠에스 `SOD-2026-0344`는
  요청 10/06인데 출고 11/15인 라인이 10/06 행에 묻혀 있었다.
  실측 139개 주문 → 요약 161행에서 **164행**으로, 3개 주문만 추가 분리된다
- 실데이터의 요청납기에 ISO 문자열로 들어간 셀이 38건 있는데 `as_date`가 흡수한다
- 열 너비·숫자 서식을 **컬럼명 기준**으로 지정하도록 바꿨다. 컬럼 letter를 박아두면
  컬럼이 하나 늘 때마다 서식이 조용히 한 칸씩 밀린다 (이번에 실제로 밀릴 뻔했다)
- **본문 전체를 표 한 칸(`<table><tr><td>`)으로 감싼다.** Outlook이 본문의 **첫 블록 요소 뒤**에
  자동 서명을 끼워 넣기 때문이다. 실측 2회로 규칙을 좁혔다 —
  `<br>`로 이은 인라인 텍스트일 땐 서명이 표(첫 블록) **앞**에, 문단을 `<p>`로 바꿨더니
  첫 `<p>` **뒤**에 들어갔다. `<div>`로 감싸도 Outlook은 그 안으로 파고든다.
  최상위 블록이 하나뿐이면 삽입 지점이 곧 본문 끝이 된다
- 문구는 거래명세표와 같은 톤으로 건조하게. 끝에 "본 메일은 자동 발송된 메일입니다."
- 제목 태그는 `[Delivery Schedule]` (본문·첨부의 '납기현황'과 중복되지 않게)
- **요청납기를 넘긴 공장 출고일은 빨강 + 굵게.** 색만 쓰면 색각 이상에서 안 보이므로 굵기도 같이 준다.
  출고일이 `처리중`이거나 요청납기가 비어 있으면 **판정하지 않는다** — 모르는 것을 늦었다고 하면 안 된다.
  실측 164행 중 초과 23행 / 미정 24행 / 정상 117행 (최대 +94일, 삼성정공 `PO260501`)
- HTML/평문 **둘 다** 같은 데이터를 담는다. HTML을 못 보는 클라이언트가 "첨부 참조"만 받으면 안 된다.
  색을 못 쓰는 평문은 `*` 표식 + 각주(`* 요청 납기일보다 공장 출고일이 늦은 건`)로 정보량을 맞춘다
- `처리중`(출고일 미정) 행은 빨간색 — 고객이 가장 먼저 묻는 줄이다
- 첨부 기본은 **xlsx** (TS는 PDF). 납기현황은 고객이 정렬·가공해 보는 표라 원본이 쓸모 있다
- 본문 값은 전부 이스케이프. 'S&T중공업', '<긴급>' 같은 비고에 표가 깨지지 않는다.
  표 자체는 이스케이프 뒤 토큰 치환으로 되살린다 (순서가 바뀌면 `<table>`이 글자로 나간다 — 회귀 테스트 있음)

### 검증
- `tests/test_mailer.py` 100 passed (TS 경로 회귀 없음 — 3건은 `MailOptions` 이사로 패치 대상만 갱신)
- `tests/test_delivery_status.py` 79 passed (본문 표·HTML 구조·발송 흐름 22건 추가)
- 실제 `.eml` 생성 확인 — 엔이에스 수신자 `nes@neskorea.co.kr` + 참조 5명(기존 회신과 동일),
  `X-Unsent: 1`, xlsx 첨부, text/plain + text/html 양쪽에 14행 표,
  HTML 블록 순서 `p → p → table → p`

### 주의
`Remarks`가 고객에게 **그대로** 나간다. 지금 데이터에도
`2025-357N, JMU S5617, 2026-02-09-CH-05 미출고 Worm gear 사용하여 제작` 같은 내부 생산 메모가 있다.
발송 전 초안에서 한 번 훑는 것을 전제로 한다 (그래서 기본이 즉시 발송이 아니라 초안이다).

---

## 2026-07-29: 거래처 납기현황 회신 CLI (`delivery_status.py`)

### 배경
거래처가 "언제 나오냐"고 물을 때마다 `SO_국내`를 수기로 피벗해 메일로 회신하고 있었다.
(Customer PO / Remarks / 수량 / NOAH 공장 출고일 표) 이걸 CLI 한 줄로 만든다.
조회 기준은 **사업자등록번호**, 대상은 국내(`SO_국내`).

### 핵심: 시트 `Status`를 읽지 않고 DN 출고수량으로 직접 계산
`SO_국내.Status`는 수기 입력이 아니라 파워쿼리 결과의 캐시다:

```
Status = IFERROR(XLOOKUP(SO_ID&Line item, SO_통합[SO_ID]&SO_통합[Line item], SO_통합[출고완료]), "")
```

그래서 **새로고침이 밀리면 실제 출고와 어긋난다.** 2026-07-29 시점 실측으로 29행이 불일치였고,
그중에는 이미 납품(DN 출고일 2026-07-27)했는데 `Status='미출고'`로 남은 라인이 있었다
(삼신 `SOD-2026-0651` 8라인, 키밸브 `SOD-2026-0210` 8라인 등). 캐시를 그대로 믿고 회신하면
**이미 보낸 물건을 "아직 안 나갔습니다"라고 고객에게 알리게 된다.**

`SO_통합[출고완료]`의 정의 자체가 DN 출고수량 기반이므로(2026-06-24 수량 기반 전환),
`DN_국내`에서 같은 식을 계산하면 결과는 같으면서 새로고침 여부에 의존하지 않는다.
`DN_국내`는 직접 입력 시트라 캐시가 없다. `pandas.read_excel`이 수식 셀의 캐시값을 읽는다는
점을 감안하면, 파일을 읽는 도구가 파워쿼리 새로고침을 전제하는 구조 자체가 취약하다.

시트 `Status`는 `Cancelled`/`Hold` 제외에만 쓴다 — 취소 건은 DN이 영영 안 생겨 수량으로는 구분할 수 없다.

### 구현
- `delivery_status.py` (신규)
  - `attach_ship_status()` — 파워쿼리 `SO_통합[출고완료]`와 동일한 4단계 판정.
    DN 행의 **존재**로 미출고를 가른다(`_dn_match`) — 수량 0짜리 DN 행은 미출고가 아니라 부분 출고다
  - **분할 납기 분리** — 한 주문 안에서 `EXW NOAH`가 갈리면 날짜별로 행을 나눈다.
    미출고 주문 139건 중 10건이 갈리고, 세진밸브 `SOD-2026-0264`는 32라인이 **7개 날짜에 걸쳐 2027년까지**다.
    대표 날짜 하나로 뭉개면 나머지 6개 납기가 회신에서 통째로 사라진다
  - `item_note()` — 나뉜 행은 같은 Customer PO가 여러 줄로 보이므로 **비고만으로 구분이 안 될 때만**
    품목명을 덧붙인다. 10건 중 8건은 비고가 이미 호선별로 갈려 있어(`H2734`/`H2735`…) 품목이 잡음이고,
    피엠에스(두 행 모두 `묘도 GS`)·티에스엔텍(비고 없음) 2건만 단서가 필요하다.
    품목명은 **개수가 아니라 길이로** 접는다 — `MA02/0.75kW/43RPM/ON-OFF/SBWG-04-1SM`처럼 긴 거래처가 있어
    "3개까지"로는 표가 깨진다. 첫 품목은 한도를 넘어도 반드시 남긴다
  - `as_date()` — `EXW NOAH`는 빈 칸이 NaN이 아니라 `datetime.time(0,0)`으로 읽힌다.
    NaN만 검사하면 미정 납기가 엉뚱한 날짜로 둔갑한다
  - 금액은 `Sales Unit Price × 미출고수량`. `Sales amount = 단가 × 수량`이 SO 전 행에서 성립하므로
    부분출고 잔량 금액도 안분 없이 정확하다
  - `count_stale_status()` — 캐시 `Status`와 계산 결과가 어긋나면 실행 시 새로고침 안내.
    회신 문서는 어차피 맞게 나가지만 대시보드·피벗은 낡은 값을 쓴다
  - 조회는 사업자번호(`normalize_biz_no`로 표기 차이 흡수) 또는 거래처명 부분일치. 후보가 여럿이면 목록 출력
- 출력 `generated_ds/납기현황_{거래처}_{사업자번호}_{YYMMDD}.xlsx`
  - `납기현황` 시트 — 주문 × 공장 출고일 1행. Customer PO / Remarks / 수량 / NOAH 공장 출고일 / Sales 금액 / PO receipt date.
    수량·금액은 **미출고 잔량 기준**, `EXW NOAH`가 비면 `처리중`. 납기 확정건 먼저, 처리중은 뒤
  - `상세` 시트 — SO 라인 1행. 주문/출고/미출고 수량, 단가, 출고상태까지
- `create_po.bat` — `[C] 거래처 납기현황 조회` 메뉴 (빈 입력이면 `--list`)
- `config.py` — `DS_OUTPUT_DIR`
- `tests/test_delivery_status.py` — 48개

### 검증
- 첨부 메일 표(엔이에스 615-81-88675)와 **14행 전부 일치** — Customer PO / Remarks / 수량(8·10·12) / 공장 출고일
- 부분출고 `SOD-2026-0301` L9 (무상공급 Eye bolt, 576 주문 / 432 출고) → 잔량 **144**로 표시,
  단가 0이라 금액은 0. 금액 기준이었으면 못 잡았을 케이스
- 분할 납기 세진밸브 `SOD-2026-0264` → 7행(`H2734`~`H2741`, 2026-08-14 ~ 2027-12-01)으로 정상 분리, 품목 주석 없음
- `pytest tests/` 회귀 없음

### 알려진 함정 (다음에 또 밟기 쉬움)
`as_date()`가 돌려준 `None`을 리스트로 모아 DataFrame 컬럼에 넣으면 pandas가 datetime64로 캐스팅해
**`None`이 `NaT`가 된다.** `d is not None` 검사를 통과한 `NaT`가 `strftime`에서 터진다
(`ValueError: NaTType does not support strftime`). 날짜 계산은 컬럼에 담기 전에 끝내고,
컬럼에는 완성된 문자열/정렬키만 넣는다.

---

## 2026-07-28: GUI 사내 배포판 (문서 생성 7종)

### 배경
문서 생성이 `create_po.bat` → `%PYTHON_PATH% create_*.py` 구조라 개발 PC에서만 동작.
다른 담당자가 쓰려면 conda 환경 구성 + `user_settings.py` 작성 + `local_config.bat` 설정을 직접 해야 했다.
Python 런타임까지 담은 무설치 배포판 + tkinter GUI로 "압축 풀고 설치.bat 더블클릭"까지 낮춘다.

### 구현
- `noah_gui.py` (신규) — tkinter 창 하나. 문서 종류 라디오 7종 → ID 멀티라인 입력 → 옵션 체크 → 실시간 로그
  - **GUI는 문서를 직접 만들지 않는다.** 기존 `create_*.py`를 자식 프로세스로 실행하고 stdout을 화면에 흘린다.
    덕분에 CLI 7개 파일을 한 줄도 고치지 않았고, Excel COM이 자식에 격리되어 COM 오류가 창을 죽이지 않는다
  - 대화형 프롬프트 → 위젯 대응: 월합 `--merge`, 메일 초안 `--mail`, FI 발주번호 기준 `--po`, 검증 무시 `--force`
  - 데이터 파일 지정 마법사 — OneDrive 자동 탐색(3초 제한) → 파일 선택 → `zipfile`로 시트 검증 → ini 기록
- `po_generator/config.py` — `noah_config.ini` 폴백 추가.
  우선순위는 `user_settings.py` → ini → 기본값이며, **user_settings.py에 이름이 있으면 값이 `None`이어도 그것이 최종값**.
  `OUTPUT_BASE_DIR = None`(프로젝트 폴더 사용)을 ini가 덮어쓰면 개발 PC의 출력 위치가 조용히 바뀌기 때문
- `cli_dist/build_portable_gui.py` (신규) — 배포 zip 빌드. `cli_dist/requirements.txt`로 버전 고정
- `설치.bat` — `%LOCALAPPDATA%`로 복사 + 바탕화면 바로가기. `/XF noah_config.ini`로 업데이트 시 사용자 설정 보존

### 빌드에서 확인한 것들
- **임베디드 Python에는 tkinter가 없다** (`_tkinter.pyd`·`tcl/`·`Lib/tkinter` 전부 부재). NuGet CPython에도 없음(1773 엔트리 중 0건).
  → tkinter·pythonw·pip을 모두 포함하는 python-build-standalone 배포본 사용
- **빌드는 임시 폴더에서 한다** — (1) 프로젝트 폴더가 OneDrive 안이라 런타임 파일 9천 개가 동기화되고,
  (2) 경로가 길어 `python/Lib/site-packages/...`에서 MAX_PATH(260자)를 넘겨 pip이 실패한다(실측). 프로젝트에는 zip 하나만 남긴다
- 배포 크기: 트리밍(pdb·pandas/tests·idlelib) 후 191MB → zip 70MB
- 검증: 배포판 런타임에서 Excel COM 왕복 + 실제 PI 1건 생성(10 아이템, 37KB) 성공

### 버그 수정
- **BOM이 붙은 `noah_config.ini`를 조용히 무시** — 메모장으로 편집·저장하면 UTF-8 BOM이 붙어
  첫 섹션 헤더가 `﻿[paths]`가 되고 설정 전체가 버려졌다. `utf-8-sig`로 읽도록 수정 (배포판 검증 중 발견)

### 배포 범위
문서 생성 7종(PO/TS/PI/FI/OC/CI/PL)만. DB 동기화·월마감·대시보드·대사(8/9/D/R/S/I)는 제외.
`create_po.bat` 콘솔 메뉴는 개발 PC용으로 그대로 유지.

---

## 2026-07-27: 거래명세표 Outlook 메일 발송 (기본 y/N 확인)

### 배경
거래명세표 발행 후 담당자가 직접 파일을 찾아 메일에 첨부하고 수신자를 기억해서 입력하는 수작업.
거래처 메일 주소가 사람 머릿속·개인 주소록에만 있어 담당자가 바뀌면 승계가 안 되고, 오발송 위험도 있음.

### 구현
- `po_generator/mailer.py` (신규)
  - `find_recipient()` — 사업자번호로 `Customer_국내` 조회. `normalize_biz_no()`로 하이픈·공백·숫자셀(`2208121175.0`)을 모두 숫자만 남겨 정규화하므로 표기 차이로 조인이 깨지지 않음
  - `split_emails()` — 한 셀에 `;` `,` 줄바꿈으로 여러 주소가 있어도 분리, 형식 깨진 주소는 경고 후 제외
  - `export_pdf()` — xlwings로 xlsx→PDF. 한글 경로 COM 이슈 회피를 위해 임시 폴더에서 변환 후 이동 (`ts_generator`와 동일 패턴)
  - `create_ts_mail()` — Outlook COM. 기본 `Display()`(초안), `send=True`일 때만 `Send()`
- `create_ts.py` — **기본 동작이 대화형 확인**. 생성 후 받는사람/참조를 출력하고
  `이메일을 발송하시겠습니까? [y/N]`을 물어 y면 Outlook 창을 띄운다.
  `MailMode`(OFF/ASK/DRAFT/SEND)로 정리하고 `--mail`(확인 없이 초안) / `--send`(확인 없이 발송) /
  `--no-mail`(묻지 않음) 플래그로 덮어쓴다. Customer 마스터는 실제 필요 시점에 1회 지연 로딩
- `config.py` — `CUSTOMER_DOMESTIC_SHEET`, `biz_no`/`customer_email`/`customer_email_cc` 컬럼 별칭, `TS_MAIL_*` 설정
- `utils.py` — `normalize_biz_no()`, `load_customer_domestic()` (사업자번호 중복 시 첫 행 + 경고)

### 설계 결정
- **수신자를 보여준 뒤 확인**: 프롬프트 전에 받는사람/참조를 먼저 출력한다. 잘못 매칭된 수신자에게 나간 메일은 되돌릴 수 없으므로, 확인 시점에 "누구에게 가는지"가 보여야 의미가 있음
- **기본은 초안, 발송은 명시적으로**: y를 눌러도 Outlook 창만 뜨고 최종 [보내기]는 사람이 누름. 즉시 발송은 `--send`를 따로 요구
- **비대화형은 자동 OFF**: `sys.stdin.isatty()`가 False면 ASK를 OFF로 낮춘다. 배치/스케줄러에서 프롬프트로 멈추는 사고 방지 (`_confirm()`도 EOFError/KeyboardInterrupt를 '아니오'로 처리해 이중 방어)
- **메일 거절 ≠ 실패**: y/N에서 N을 눌러도 종료 코드는 0. 문서는 정상 생성됐기 때문
- **설정 문제는 실행당 1회만 안내**: 이메일 컬럼 누락 등은 첫 문서에서 한 줄 안내 후 조용히 건너뜀 (10건 생성 시 같은 오류 10번 반복 방지)
- **메일 실패 ≠ 생성 실패**: 문서는 이미 만들어졌으므로 메일 단계 오류는 경고만 출력하고 종료 코드에 영향 없음. Outlook 미설치 환경에서도 import로 죽지 않게 `win32com` 지연 import
- **월합 문서 발송 차단**: `--merge`로 고객이 2곳 이상 섞인 거래명세표는 발송하지 않음 — 타 거래처 라인이 노출됨
- **수신자 미등록은 조용히 넘어가지 않음**: 사업자번호+거래처명을 찍어 어느 거래처를 마스터에 등록해야 하는지 바로 알 수 있게 함

### 새 Outlook 대응 (.eml 백엔드)
운영 PC가 **새 Outlook(olk.exe)** 으로 전환돼 있어(`UseNewOutlook=1`) COM 발송이
`-2146959355 (0x80080005, 서버 실행 실패)`로 실패. 원인은 코드가 아니라 환경:

- 새 Outlook은 **COM 자동화를 지원하지 않는다**
- classic `OUTLOOK.EXE`는 파일이 남아 있고 COM 등록도 살아 있지만, 실행되면 새 Outlook으로
  넘기고 즉시 종료 → COM 서버가 뜨지 못함 (직접 실행해 12초간 관측: 프로세스 미생성)

대응으로 `.eml` 백엔드 추가 (`TS_MAIL_BACKEND = 'auto' | 'outlook' | 'eml'`, 기본 auto):
- `build_eml()` — To/Cc/제목/본문/첨부를 담은 `.eml` 생성 후 `os.startfile()`로 기본 메일 앱에서 열기
- **`X-Unsent: 1` 헤더가 핵심** — 이게 없으면 '받은 메일' 형태로 열려 [보내기]가 없다.
  `Date` 헤더는 넣지 않는다(넣으면 수신 메일로 취급될 수 있음)
- **HTML 대체본(`add_alternative`)을 반드시 함께 넣는다.** 평문만 보내면 새 Outlook이 서명을
  본문 **위**에 끼워넣어 "서명 → 인사말" 순서가 된다. HTML이면 본문 **아래**에 정상적으로 붙고
  로고까지 렌더링된다. 거래처명의 `&`·`<`가 깨지지 않도록 `html.escape()` 후 줄바꿈만 `<br>` 처리
- `auto`는 COM 가용성을 **프로세스당 1회만** 조사하고 캐시 (건마다 재시도하면 느려짐)
- `.eml`은 작성 창을 띄우는 방식이라 **자동 발송 불가** — `--send`로 실행해도 초안까지만
  진행하고 그 사실을 실행 시작 시점과 결과 양쪽에 명시 (조용히 성공한 척하지 않음)

### auto를 쓰지 않고 'eml'로 고정한 이유
검증 중 COM Dispatch가 **성공하는 경우**가 관측됐다. classic Outlook은 여전히 실행되지 않는데,
새 Outlook의 **`olkexthost.exe`**(COM 애드인 호환 호스트)가 떠 있으면 Dispatch가 통한다.
즉 COM 가용성이 실행 시점에 따라 달라지고, 그러면 같은 명령이 COM 경로(평문 본문·서명이 본문 위)와
`.eml` 경로(HTML 본문·서명이 본문 아래) 중 하나로 갈려 **결과물이 매번 달라진다**.
`user_settings.py: TS_MAIL_BACKEND = 'eml'`로 고정. (config 기본값은 다른 PC를 위해 `auto` 유지)

새 Outlook에서 실물 확인: [보내기] 버튼 있는 작성 창, 받는사람·참조 채워짐, PDF 첨부됨,
한글 정상, 서명이 본문 아래에 로고까지 정상 렌더링.

### 전제 (사용자 작업)
- `Customer_국내` 시트에 이메일 컬럼 추가 필요 (기존 시트에 메일 주소 컬럼이 없었음).
  실제 추가된 헤더는 `수신자 이메일` / `참조 이메일` — 별칭에 반영함.
  주의: `참조 이메일`도 '이메일'을 포함하므로 부분일치로 컬럼을 찾으면 To에 참조 주소가 잡힌다.
  두 별칭 목록 모두 **완전일치 전용**으로 유지할 것
- 고정 CC는 `user_settings.py: TS_MAIL_CC`

### 검증
- `tests/test_mailer.py` 62개 — 사업자번호 정규화/다중주소 파싱/오타 제외/미등록·메일없음/컬럼별칭/CC 중복 제거,
  모드 결정(TTY·플래그 우선순위), `_confirm` 긍정·부정·EOF·Ctrl+C, 마스터 지연 로딩 1회 캐시 및 실패 1회 안내,
  Outlook COM mock으로 Display·Send 분기 및 실패 격리
- 실데이터: `Customer_국내` 806건 로드(사업자번호 중복 2건 경고), DN 1330건 중 1324건 조인, 실제 거래명세표 → PDF 77KB 변환 성공
- 종단 검증(실제 `main()`, Outlook mock): 프롬프트에 `n` → Outlook 0회 / `y` → Outlook 1회·Send 0회, 두 경우 모두 종료 코드 0

---

## 2026-06-26: 대시보드 국내/해외 탭에 시장별 건수 배지 추가 (빈 탭 오해 차단)

### 배경
"오늘의 현황 → 미발주현황"은 PO미등록 **2건**으로 표시되는데 "발주 커버리지 → PO 미등록 상세"는 비어 보인다는 제보. 추적 결과 **불일치/버그 아님**:
- 양 페이지 모두 동일한 `calc_coverage()`(SO_ID 단위)를 사용하며 PO미등록 건수도 2건으로 동일.
- 해당 2건(`SOO-2026-0191`, `SOO-2026-0195`, WATERGATES GMBH)은 **모두 해외**. 상세 테이블이 `국내/해외` 탭으로 분리돼 있고 Streamlit은 첫 탭(국내)을 기본 표시 → 국내 탭이 "건 없음"으로 비어 보였을 뿐, **해외 탭에 2건 정상 존재**.
- 보조 요인: 발주커버리지는 연도/월 필터를 적용하지만 오늘의현황은 이를 무시(현재 시점 기준, 페이지 상단 안내문구 존재).

### 변경
**dashboard.py** — *시장 탭 라벨 건수 배지*
- `_market_tabs(df, *, market_col="market", unique_col=None)` 헬퍼 추가: `🇰🇷 국내 (n)` / `🌏 해외 (m)`처럼 탭 라벨에 시장별 건수를 표시해 데이터가 어느 탭에 있는지 즉시 보이게 함. `unique_col` 지정 시 행 수가 아닌 고유값 수(SO_ID 단위 화면용)로 카운트.
- 기존 국내/해외 탭 9곳을 전부 이 헬퍼로 통일: PO확정지연·미발주현황·EXW미출고·납기현황·해외선적불일치(오늘의현황) + PO미등록·미발주·부분발주·발주진행중 상세(발주커버리지). `market_col`이 `마켓`/`market`으로 다른 화면, SO_ID 중복 화면(납기현황) 모두 대응.

### 검증
- `_market_tabs` 라벨 출력 단위 검증 4종 통과: 실 PO미등록 데이터 → `국내 (0)`/`해외 (2)`, `마켓` 컬럼 경로, `unique_col=SO_ID` 중복제거 경로, 빈/무컬럼 DataFrame 안전(0/0).
- `ast.parse` 구문 통과, 잔존 구형 `st.tabs(["국내"…])` 0건 확인.

---

## 2026-06-26: 동기화 로그 "삭제 오해" 근본 차단 (미완성 행 보류 + 재키잉 인식)

### 배경
대시보드 동기화 로그에서 `ND-0673`이 **신규(12:47) → 삭제(13:12)** 로 찍혔는데 Excel·DB엔 데이터가 멀쩡히 존재하는 제보. 추적 결과:
- PO 시트 PK는 복합키 `(PO_ID, Line item, _row_seq)`. `Line item`이 **키 구성요소**라 그 값이 바뀌면 같은 행이 **삭제(옛 키)+신규(새 키)** 로 기록된다(가변 키의 구조적 특성, 버그 아님).
- `ND-0673`은 12:47에 **`Line item` 빈 채로 먼저 동기화**(키 `ND-0673 |  | 1`) → 13:12 전 `Line item=1` 입력 → 키가 `ND-0673 | 1 | 1`로 바뀌며 옛 키 prune(삭제)+새 키 insert(신규). 행은 새 키로 살아있음.
- 실증 조사: `Seq`도 정렬용 보조값이라 **803회 값→값 재번호**되어 키 후보 부적합. 전체 이력에서 PO "삭제" 로그 **1,271건 중 진짜 삭제는 2건**(오타 PO_ID 교정), 나머지는 전부 같은 PO_ID가 살아있는 **재키잉(identity churn)** — 즉 PO "삭제"는 99.8%가 거짓 신호였음.

### 변경
**po_generator/db_sync.py** — *미완성 행 보류*
- upsert 루프의 PK 추출에서, `_row_seq`(자동 생성)를 제외한 자연키 컬럼이 **하나라도 비면 해당 행을 스킵**(동기화 보류)하도록 변경. 키가 채워질 때까지 DB/로그에 들이지 않아 *빈 키 신규 → 삭제* 패턴을 원천 차단. 기존 `required_column` 단일 검사 + `real_pk 전부 빈값` 검사를 이 단일 규칙으로 통합(상위호환). 현재 데이터 영향 0행(전 시트 빈 키 0개).

**sync_db.py** — *재키잉(키변경) 인식*
- `_reconcile_rekeys()` 추가: 같은 sync·같은 시트에서 **문서ID(pk[0]) 동일 + 비키 컬럼 내용 일치**한 삭제↔신규 쌍을 1:1로 묶어 **'수정' 단일 이벤트**로 변환(키 컬럼 변경분을 `changes`에 기록). 모호(후보 0/2+)하거나 내용 불일치면 묶지 않고 삭제 보존(안전 우선). `write_sync_log_to_db()`가 이를 적용해 `_sync_log`에 기록 → `Line item 1→12` 같은 편집도 더는 삭제로 안 뜸.

### 검증
- `tests/test_db_sync_rekey.py` 9건 신규(미완성 보류 실엔진 2 + 재키잉 순수함수/가드 6 + 로그기록 통합 1) 전부 통과.
- 실 DB `--dry-run`: 전 시트 **에러 0 / 스킵 0 / 삭제 0** — 기존 완성 데이터 영향 없음 확인.
- `pytest tests/` **290 passed / 2 skipped**.
- ⚠️ 과거 `_sync_log` 기록은 감사 추적이라 보존(소급 수정 안 함) — 개선은 **이후 동기화부터** 적용.

---

## 2026-06-24: 출고상태 판정을 수량 기반으로 전환 (무상공급 부분출고 오표시 수정)

### 배경
`SOD-2026-0301` Line 9(Eye bolt, FOC 무상공급 576개 중 220개 출고)가 부분출고인데 **"출고 완료"로 표시**되는 제보. 추적 결과 두 가지 원인:
- **대시보드**는 수동 입력 `SO_국내.Status`를 그대로 표시 — 운영자 입력 오류에 그대로 노출. 실 DB 확인 시 수동 Status가 실제 출고와 **116/2281행 불일치**(출고했는데 "미출고" 방치 58건 등).
- **SO_통합 쿼리**의 출고완료 판정이 **금액 기반**(`[Sales amount KRW] - [출고금액] > 0`) — 무상공급은 단가 0이라 Sales·출고금액 모두 0 → `0-0>0`=거짓 → 부분출고를 영원히 감지 못 함. 2026-01-30 금액기반 패치(`[출고금액]=null`)는 무상 건의 출고/미출고 이분법만 고친 미완성 패치였음.
- 반면 Order Book SQL(`sql/order_book.sql`)은 이미 **수량 기반**(`Input_qty - Output_qty`)으로 올바르게 동작 → 수량 기준이 정답임을 입증.

### 변경
**dashboard.py**
- `load_so()`: 수동 Status 대신 **DN 출고수량 누계**로 출고상태 파생 — DN없음=미출고 / `SO수량-DN수량>0`=부분출고 / 출고일없음=공장출고 / 그외=출고완료. 수동 Status는 Cancelled/Hold 주문 제외에만 사용. 금액이 아닌 수량 기준이라 무상공급·해외 환율차 케이스까지 정확. 부수효과로 납기현황(`status != 출고완료` 전치필터)에서 누락되던 무상 부분출고 라인도 정상 포착.

**docs/POWER_QUERY.md (SO_통합)**
- `DN_Combined`에 `출고수량`(=`List.Sum([Qty])`) 추가, 판정식을 `[Item qty] - [출고수량] > 0` 로 수정(null 가드 포함). 결과컬럼 표에 `출고수량` 추가 + `출고완료` "수량 기준" 명시, 트러블슈팅에 2026-06-24 수정 이력(2026-01-30 패치 계보) 기록. 금액 기반 `미출고금액`은 재무 백로그용으로 유지(무상 건 0이 정상).
- ⚠️ M 코드는 .xlsx 내부 바이너리라 자동 반영 불가 — Power Query 고급 편집기에 수동 붙여넣기 필요.

### 검증
실 DB로 `load_so()` 실행 → Line 9 `부분 출고` 정상(Line 4 부분출고·5/7/8 미출고·1/2/3/6 출고완료 일치), 출력 컬럼 19개 유지. `pytest tests/` 277 passed/2 skipped(실패 4건은 `NOAH_SO_PO_DN.xlsx` Excel 잠김 환경 `PermissionError`로 변경과 무관, dashboard 테스트는 전부 통과).

---

## 2026-06-22: Sync 이력 로깅 개선 — 버그·UX·직관성 (멀티에이전트 리뷰 후속)

### 배경
sync 이력(`_sync_log`/`_sync_runs`) 기능을 멀티에이전트 + 적대적 검증으로 재리뷰(37건 확정). 데이터 무결성을 깨는 버그는 없었고, 이력의 "조회·견고성·의미 표현"에 집중된 항목을 수정.

### 변경
**sync_db.py**
- 로그 기록(`write_sync_log_to_db`)을 try/except로 격리 — 로그 저장 실패가 이미 commit된 동기화 요약 출력/정상 종료코드를 가리던 문제 해소(스케줄러 오판 방지).
- PK를 `_normalize_pk`로 일괄 정규화 — 동일 레코드의 신규/수정/삭제 이벤트가 `1.0` vs `1`로 다른 키로 남던 비대칭 제거.
- 변경 0건도 `_sync_runs` 실행 이력 1행 기록(누가/언제 동기화 감사) + 변경 이력 조회 안내 출력.
- `--log [N]`(최근 세션 이력)·`--note`(세션 메모) 신규. `PRAGMA foreign_keys=ON`. 콘솔 표 전각(한글) 정렬 보정. `--changes` 롤백 시 '미적용' 경고 헤더.

**db_schema.py / db_sync.py**
- `create_sync_run(started_at=...)`로 `_sync_runs.started_at`을 실제 동기화 시작시각으로 기록(기존엔 '로그쓰기 시점'이라 시작/종료가 거의 동일했음).
- prune 삭제 스냅샷을 Excel 헤더가 아닌 **전체 DB 컬럼 기준**으로 통일(`_prune_snapshot_columns`) — 빈시트/정상 경로 완전성 일치, 시트에서 사라진 잔류 컬럼도 보존.

**dashboard.py**
- 주문검색 매칭을 첫 토큰 → **PK 전체 토큰 대조**(`_pk_tokens`) — 삭제(prune)된 DN을 SO_ID로 검색 시 누락 + 빈 PK/깨진 JSON 페이지 크래시 동시 해결(타임라인·탐색 2곳).
- 세션 요약에 실제 소요(초) 추가 + 항상 공백이던 dry_run 컬럼 제거, 추이 차트 `errors='coerce'`, ack 성공/실패 반환+`st.error`, 관여자 빈값 `(unknown)` 정규화, `resolve_related_ids` 2→6 pass, `_so_change_ack` 연결 FK 강제.

**문서/마이그레이션**
- `DB_SYNC_GUIDE.md` 기록규칙을 폐기된 v1 → v2로 정정(존재하지 않는 `old_value`/`new_value`/`column_name` 설명 제거), started_at/dry_run 설명·`--log`/`--note` 추가.
- `migrate_sync_log.py` `input()`에 비대화형(`isatty`) 가드.

### 검증
py_compile 6파일 OK, `pytest tests/` 281 passed/2 skipped, 스모크(write 경로 pk정규화·0건 run·note·started_at / prune 정상+빈시트 / 대시보드 변환 실DB 1000행) 전부 통과.

---

## 2026-06-22: 전면 코드 감사 — 정합성 버그 일괄 수정 (High 1 + Medium 18 + LOW 30)

### 배경
멀티 에이전트 적대적 감사 2회 실시. 1차(모듈별 너비): 101건 발견 → 적대적 검증으로 53 확정.
2차(교차/저커버리지 렌즈 + 1차 기각 48건 재심): **7건 false-negative 복구** + 신규 17 + 1차 수정 회귀 2건 검출.
hand-maintained Excel(ERP 미연동) 특성상 중복·공란·이상타입 셀이 현실적이라 **금액/환율/날짜필터/조인 오류가 회계오류로 직결**되는 항목을 우선 수정.

### 변경 — High/Medium (정합성·금액 직결)
**대시보드**
- 날짜 1900 더미값(Excel zero-date) 정화 중앙화(`_sanitize_date`) — `load_so` 4개 날짜컬럼 + `load_backlog`. 라이브 DB 확인 결과 `SOO-2026-0165`가 납기 `1900-01-01`로 **약 126년 지연 오집계**되던 것 제거(허위 납기지연/OTD 0%/경과일 폭주). (High)
- 납기현황 SO 합계를 경과 라인만 → **주문 전체 라인 기준**으로(주석/라벨 일치). 회귀수정: 라인별 잔여 음수(과출고) `clip(0)`으로 부족분 상계 방지. PO EXW 보충 fan-out 방지.

**문서 생성기**
- 수량 파싱 `int(raw_qty)` → `int(float(raw_qty))` — `'2.0'` 문자열이 ValueError로 **수량 1 묵살** → Invoice 금액 손상. TS + fi/ci/oc/pi/pl 6개 생성기 전체(PO와 동일 패턴 통일).
- `create_pi` 비숫자 단가 `:.2f` 가드(배치 전체 중단 방지), `create_fi` 복수 RCK PO 분리 시 공란 RCK PO 라인 누락 경고(과소청구), `--po` FI 복수 DN 통합 시 Invoice No 대표DN 경고, `create_ts` 월합 출력파일명 충돌 안전장치/중복 DN_ID dedup.

**core**
- `utils.resolve_column` 캐시 키 `id(columns)` → `tuple(columns)`(GC id 재사용 오매핑 방지). 복합키 머지(`_load_and_merge_sheets`/`load_dn_data`) 참조측 dedup + 경고(행 fan-out 이중집계 방지).

**DB 동기화/감사로그**
- `db_sync` prune 2경로 **rowid 기준 삭제 + rowcount 집계** — 정규화 PK(`'42'`)가 legacy 원본(`'42.0'`)과 불일치해 0행 삭제하면서 `pruned` 과대집계되던 것 수정.
- `sync_db` 삭제 감사로그를 `pruned_snapshots` 1:1 순회(회귀수정: 중복/스냅샷 유실). **롤백(`total_errors>0`) 시 `_sync_log` 유령기록 방지 게이팅**.
- `migrate_sync_log` 비작동 v1 마이그레이터를 전용 `_sync_log_legacy_v1` 테이블로(v2 오염/크래시 제거). `snapshot` 마감 시 출하·KRW 공란 해외 DN 경고(phantom backlog 동결 전 surface).

**대사(reconcile)**
- `reconcile_po` 1:N AX PO 계산서금액 **이중계상 제거**(`_line_id` 분배, `.copy`로 요약 격리), `ax_service` 국내/해외 키 disjoint + 미분류 Product GRN 경고.
- `reconcile_so` AX Project 정규화 양측 적용(`.0` 업캐스트 매칭불가 → 매출누락 수정), 해외 `Total Sales KRW`/매출일 결측 경고, FX 월매칭 연도폴더 인식. `reconcile_ind` ind_code 정규화 헬퍼 공유.

**Order Book SQL**
- `order_book_variance`: 납기변경 제외를 '음/양 동시존재' → **그룹 순변동 상쇄(net≈0)**로(실제 환율/판매가 변동이 묻혀 사라지는 false-positive 방지).
- `load_backlog`/`order_book_backlog`/`_snapshot_backlog`: 금액 4건 `ROUND` 통일(order_book/snapshot과 tie-out).

### 변경 — LOW (견고화, 30건)
- **생성기/CLI/core**: 컬럼 letter 산술 `get_column_letter`(>Z·AA+ 안전), `escape_excel_formula` 선행 제어문자 우회 차단, `validators` 정당한 0값 '필수누락' 오라벨 수정, `_to_text`→`utils.to_text` 승격+모델코드 적용, 미사용 셀상수 제거, fi/ci 0단가 fallback 센티넬(무상라인 보존), ci 중복 Shipping Mark 제거, `create_po` **중복승인↔검증오류승인 분리**(중복 Y가 검증오류 우회 방지), `create_fi` DN_ID+`--po` 충돌 가드, `history` 중복탐지 정규식 앵커.
- **대사**: FX 임계값 상대오차 `max(100, 0.5%)`, Customer '' backfill, 상세시트 매출일/선적일 추가, `build_mapping` 값있는 코드 우선dedup, `recon_paths` 빈 플랫폴더가 연도폴더 가리지 않도록.
- **대시보드**: sync-log 검색 `pk_json` 정확매칭(부분문자열 오매칭 제거), OTD `groupby min`(fan-out 방지), 세금계산서 aging 라벨 비중첩, 미출고금액 음수 clip+캡션, backlog KPI '라이브 값' 캡션, `load_backlog` HAVING 수량 OR 금액.
- **DB/SQL**: 빈시트 prune 시 `_sync_meta` 갱신, NULL EDD `COALESCE('')` 정규화(3파일), backlog 잔여수량 ROUND·HAVING.
- **의도적 제외**(사유 기록): ts 라인별 VAT(sub-10원·의도적), FI 모델 prefix(소유자 확인필요), dead code, 스키마변경(undo 이력), theme CSS(시각회귀), by-design(snapshot EDD-move/소급) 등.

### 검증
- `pytest tests/` **281 passed, 2 skipped**(전 단계 반복) — test_validators/create_po/utils/history/cli_common 등이 핵심 변경 커버
- 라이브 `noah_data.db`에서 수정 SQL 6종 정상 실행(HAVING 변경으로 누락됐던 개시 라인 2건 노출 확인)
- 순수로직 타깃 검증(날짜정화, prune rowid, FX 연도, ind/project 정규화, reconcile_po 분배/disjoint), 로직 민감 변경 diff 스팟체크
- 2차 회귀 리뷰: 1차 13개 수정 중 11개 안전 확인, 회귀 2건 즉시 수정

### 파일 변경
33개 파일(+702/−343), 브랜치 `audit-fixes-2026-06`(11커밋, High/Medium/LOW/docs 논리 단위 분리). 주요: `dashboard.py`, 생성기 7종, `reconcile_po/so/ind.py`, `db_sync.py`·`sync_db.py`·`snapshot.py`·`migrate_sync_log.py`, `utils.py`·`validators.py`·`history.py`·`recon_paths.py`, `create_*.py`, `sql/order_book*.sql`. 상세 항목·제외사유는 `tasks/todo.md` 참조.

---

## 2026-06-01: Order Book 마감 — 금액 원 단위 ROUND (유령 잔량 / 'Start != 전월 Ending' 경고 제거)

### 배경
`close_period.py --list`(메뉴 [9]→[3])에서 `** Start != 전월 Ending (!차이 1)` 경고가 디테일 없이 떴음. 추적 결과 KRW 금액이 DB에 `REAL`로 저장되는데, 이른 마감 시점에 **소수점이 박힌 값이 스냅샷에 동결**(예: 1월 input `10425305.49251159`)됐고, 이후 원본이 정수로 정리되며 생긴 < 1원 차이를 Variance 임계값(`> 0.5`)이 걸러내 **유령 ending 잔량**으로 남음. 이 잔량이 월별로 누적(비정수 ending 1월 41건→…)되어 5월에 1.39원이 되며 `diff > 1` 경고를 유발. 각 조각이 0.5원 미만이라 Variance 디테일엔 안 잡혀 "디테일 없는 경고"가 됨. KRW는 정수 통화이므로 sub-won 값 자체가 float 인공물.

### 변경
- **소스 금액 원 단위 ROUND** — rolling/snapshot SQL의 SO/DN 금액 컬럼에 `ROUND(CAST(... AS REAL))` 적용. 향후 모든 마감이 정수 유지, 진짜 ≥1원 차이만 Variance(`> 0.5`)·경고(`> 1`)로 노출
  - `Sales amount`/`Sales amount KRW`(SO 국내/해외), `Total Sales`/`Total Sales KRW`(DN 국내/해외)
- **기존 스냅샷 제자리 ROUND 마이그레이션** (`migrate_snapshot_round.py`, 멱등) — 전체 재마감 대신 `ob_snapshot` 금액 5컬럼(start/input/output/variance/ending)만 정수화. **전체 재마감은 마감 간 소급변경 이력(Variance)을 0으로 소실시키므로 채택하지 않음** — `closed_at`이 실제 월별 마감(3/13·4/1·5/1·6/1)이고 Variance 대부분이 납기변경 상쇄쌍·가격변경 등 진짜 감사 이력이라 보존 필요. 드롭된 잔량은 전부 < 0.5원 → `round`=0 → `Σ Start = Σ 전월 Ending` 정확히 성립
- **`order_book_snapshot.sql` 임의 월 조회 파라미터화** — `params(period)` CTE 한 줄로 마감월/미마감월 모두 조회. 마감월 → `ob_snapshot` 동결값(무활동 그룹 포함 전 그룹, `--list` 총계와 일치), 미마감월 → 라이브 롤링. 기존 "open period만 표시"의 한계(마감월은 못 봄) 해소. `fallback_periods`는 `closed_periods`로 대체(미마감·스냅샷 없음 케이스는 `open_periods`가 흡수). ⚠️ 단순 `Period` group 합산은 이벤트 기반이라 활동월 그룹만 잡혀 과소집계 — 월말 총 백로그는 이 SQL(또는 `--list`/대시보드 ffill)로 조회
- **미매칭 출고(orphaned DN) 라벨을 DN 자체 값으로 표시** — SO 라인과 매칭 안 되는 DN(출고) 이벤트의 `Customer name`/`OS name` 등이 `'UNKNOWN'`으로만 떴음(예: SO에 없는 "시운전 SETTING" 라인 출고). `events_line_item` 출력 branch가 `dn_combined`로 전파한 DN 필드를 폴백으로 사용: `Customer name`·`Customer PO`·`Item name`·`OS name`(=DN Item)·`구분`(국내/해외)·사업자번호를 DN 값으로 채움(`NULLIF(...,'0')`로 placeholder 처리). SO 고유 필드(Sector/AX Period/Model code/Industry code/EDD)는 DN에 없어 공란 유지. 매칭되는 정상 출고는 기존과 동일(SO 값 우선) — 회귀 없음. 음수 Ending이 미매칭 신호 역할 유지. 기존 동결 스냅샷의 `UNKNOWN`은 재마감 전까지 그대로(해당 1건은 이미 Ending=0 해결)

### 검증 (실데이터, 2026-01~05)
- `--list` 경고 사라짐 / Start − 전월 Ending 전 구간 `+0.0000` / 비정수 ending 225행 → **0**
- Variance 이력 보존 (3월 -6.7M, 4월 -119.6M, 5월 -21.4M, 정수화만)
- 미래 마감(2026-06) 시뮬 823그룹 비정수 0 / `pytest` dashboard+integration 75 passed

### 파일 변경
| 파일 | 변경 |
|------|------|
| `po_generator/snapshot.py` | `_ORDER_BOOK_BASE_SQL`의 so/dn 금액 4컬럼 `ROUND` + orphaned DN 라벨 DN 폴백 |
| `sql/order_book.sql` | so/dn 금액 4컬럼 `ROUND` (대시보드 Order Book 페이지도 정수 표시) + orphaned DN 라벨 DN 폴백 |
| `sql/order_book_snapshot.sql` | so/dn 금액 4컬럼 `ROUND` + `params(period)` CTE로 임의 월(마감/미마감) 조회 파라미터화 (`fallback_periods`→`closed_periods`) + orphaned DN 라벨 DN 폴백 |
| `migrate_snapshot_round.py` | 신규 — 기존 `ob_snapshot` 금액 원 단위 ROUND (1회성·멱등, `--dry-run` 지원) |

---

## 2026-05-22: Packing List Net Weight — Model+옵션 기반 Weight 매핑

### 배경
PL의 Net Weight(G열)는 `_enrich_with_weight()`가 SO_해외 `Model code`로 Weight 시트를 조회했으나, 해당 컬럼이 전 행 비어 있어 **항상 공란**이었음. 액추에이터 Model/옵션은 `PO_해외`에만 존재.

### 변경
- `_enrich_with_weight()` 재구현 — DN 아이템을 **(SO_ID, Line item) 복합키**로 PO_해외와 조인, PO의 `Model`(AN열) + 옵션(AO~BH열, Y표시)을 Weight 시트와 매칭
- 매핑 규칙: PO Model에서 `NA`/`SA` 접두어 제거 → Weight `MODEL` base 코드. 무게 영향 옵션(INTEGRAL/IMS/LCU/PCU+PIU/SCP/EXP)을 접미사로 부착해 조회. 복수 옵션은 우선순위 1개(LCU+PCU 동시는 결합코드 `…LP` 우선). 미매칭 시 base Model 무게 폴백, base도 없으면 공란
- Total 행 G열 = `SUMPRODUCT(Qty, 단위중량)` — 행별 단위중량(KG/PC)에 수량을 곱한 총 Net Weight (기존 단위중량 단순 합에서 변경)
- 검증: PO_해외 Model 보유 553행 중 549행(99.3%) 해결, 미매칭은 비표준 액세서리 4건(MOTOR/MS01/SCP-SET)

### 파일 변경
| 파일 | 변경 |
|------|------|
| `config.py` | `WEIGHT_OPTION_SUFFIX`, `WEIGHT_OPTION_PRIORITY` 상수 추가 |
| `utils.py` | `build_weight_map()`(ITEM→WEIGHT, 미사용) 제거 → `build_model_weight_map()`, `load_po_export_data()`, `resolve_weight_code()`, `build_po_line_weight_map()`, `normalize_line_item()` 추가 |
| `services/document_service.py` | `_enrich_with_weight()` PO 기반 (SO_ID, Line item) 조인으로 교체 |
| `pl_generator.py` | Total 행 G열 수식 `SUM` → `SUMPRODUCT(Qty,단위중량)` |
| `tests/test_utils.py` | `normalize_line_item`/`_po_base_code`/`resolve_weight_code` 단위 테스트 추가 |

### 후속 수정 — LCU 'L' 이중 인코딩 차단
`resolve_weight_code()`: `SA005L` 처럼 Model명에 이미 LCU('L')가 들어 있는데 옵션열 `LCU=Y` 까지 중복 체크되면 결합분기가 `005LLP` 같은 'L' 이중 코드를 만들 수 있었음. Model명이 'L'로 끝나면 옵션 LCU를 제거하고, 결합코드 `…LP`는 Model명에 'L' 미포함일 때만 별도 생성하도록 수정. 실데이터 영향 0건(LP 결합 34건·전체 매칭 불변), 순수 안전망. 엣지케이스 단위 테스트 2개 추가.

---

## 2026-04-30: PO 매입대사 — Confirmed 출고 제외 + 요약 시트 추가

### Confirmed PO 매핑 출고 제외 — `(서비스 출고만 ∪ Invoiced this-month) − GRN`

GRN_미포함이 Confirmed 상태 PO의 출고까지 잡아 잘못 집계되던 문제. 예: 출고리스트의 `ND-0232` → PO_국내 매핑 → AX PO=`P023207` (Status=Confirmed)이 P04 GRN_미포함에 Service로 분류돼 들어옴. P04 대사 범위 밖인데 노이즈로 보였음.

원인: AX PO 매핑 테이블이 `df_po` (Status 무관, AX PO 있는 모든 PO) 기준으로 만들어져 Confirmed/기타 상태 PO도 매핑 후보. 매핑 자체는 AX_PO_매핑 export 등 다른 출력에 필요하니 유지하되, **미포함 후보 산출 시에만 필터** 적용.

수정 (`build_raw_data`):
```python
service_del_set = del_set - po_all_ax_set    # 어떤 PO에도 등록 안 된 직접 P###### 출고
not_in_grn_ax = (service_del_set | po_set) - grn_set
```

`po_all_ax_set` = `df_po_all` (Cancelled 제외, Status 무관)의 AX PO 전체. 출고 AX PO가 여기 포함되면 어떤 PO에 등록된 건이므로 매핑된 출고 — Invoiced 아니면 미포함 후보에서 제외. 직접 P###### 출고만 service로 카운트.

`build_raw_data` 시그니처에 `df_po_all` 인자 추가, main()에서 전달.

### 요약 시트 — `구분 × {Excel, AX, Diff}` (첫 번째 탭)

PO Invoiced 매입금액 / 회계 GRN 처리금액을 한 화면에 비교. Diff = AX − Excel 로 GRN 처리 잔액 표시.

```
구분          | Excel         | AX            | Diff (AX-Excel)
Product(국내) |   491,707,008 |   387,499,138 |   -104,207,870
Product(해외) | 1,034,921,076 |   999,095,301 |    -35,825,775
Service       |    49,267,526 |    49,267,526 |             0
Total         | 1,575,895,610 | 1,435,861,965 |  -140,033,645
```

- **Excel** = NOAH_SO_PO_DN Invoiced PXX (Product 국내/해외) + 출고리스트 직접 P###### 행
  - 출고리스트의 Type=`YTC` 직접 출고는 Product(국내)에 합산 (서비스 분류 아님)
  - Type=`Service` 또는 그 외 직접 출고는 Service 행
- **AX** = GRN의 `Cost amount physical` — AX PO가 어느 PO 그룹에 속하는지로 분류
- **Diff** = AX − Excel = GRN 처리 필요분 (음수면 미처리, 양수면 초과 처리)
- Diff 합계가 대사 시트 합계 블록의 'GRN 미포함 합계' 와 일치 → 양방향 검증

### 신규/수정 파일

- `reconcile_po.py` — `build_raw_data()` 시그니처 + Confirmed 매핑 출고 필터 / `요약` 시트 생성

---

## 2026-04-29: PO 매입대사 — 연도 폴더 구조 + 양방향 합계 검증

### 폴더 구조: 연도 중첩 지원

`po_reconciliation/`, `so_reconciliation/`, `ind_code_reconciliation/` 모두 `RECON_DIR/{year}/{period}/` 레이아웃을 도입. 신규 헬퍼 `po_generator/recon_paths.py` 의 `resolve_period_dir()` / `iter_period_dirs()` 가 플랫(`{period}/`) / 중첩(`{year}/{period}/`) 두 레이아웃을 모두 자동 인식 — bat 메뉴는 그대로(P04만 인자로 넘김).

- `ind_code_reconciliation/P03/` → `2026/P03/` 이동 (sector_검증.xlsx은 월 무관 파일이라 루트 유지)
- 같은 period가 여러 연도에 있을 때 최신 연도 선택 + `[경고] po_reconciliation/P04 가 여러 연도에 존재 — 2027 사용 (다른 연도: 2026)` 출력
- reconcile_po / reconcile_so / reconcile_ind 세 스크립트 모두 입력/출력 경로가 새 헬퍼로 통일

### GRN 시트 자동 탐색 — `Purchase order` 컬럼 기준

회계팀이 4월 GRN 파일에 빈 시트(`GRN List_04`)를 첫 번째로 추가하면서 `pd.read_excel(..., sheet_name=None)` 이 잘못된 시트를 읽어 `KeyError: 'Purchase order'` 발생. `load_grn()` 이 시트 목록을 훑어 `Purchase order` 컬럼이 있는 시트를 자동 선택하도록 변경 — 시트 순서/이름이 또 바뀌어도 안전.

### 대사결과_PXX.xlsx — 양방향 합계 검증 블록

대사 시트 표 아래(테이블 외부)에 합계 블록 자동 삽입:

```
── GRN 기준 ──
GRN 합계         : 851M (전체 GRN, raw 합계와 일치)
비교가능 합계    : 851M (PO/출고 매칭)
비교불가         : 0M   (서비스/기타: GRN만 존재)

── PO Invoiced P04 기준 ──
PO Invoiced 합계 : 1,379M
  GRN 매칭       : 826M (대사 시트 PO 합계)
  GRN 미포함     : 553M (별도 시트 — AX에 GRN 처리 필요)
  AX PO 미입력   : 0M   (AX_PO_미입력 시트)
```

- **검증식**: `GRN 매칭 + GRN 미포함 + AX PO 미입력 = PO Invoiced 합계` ✓
- 사용자가 Excel에서 수동으로 SUBTOTAL 합계 행을 추가하다 컬럼 참조 오타로 GRN_금액 칸에 비교_금액 합이 표시되던 문제 해소
- 합계 행은 Excel Table 범위 밖이라 정렬·필터 동작에 영향 없음

### GRN_미포함 시트 — PO Invoiced 기준만 (출고-only 분리)

이전에는 `(del_set | po_set) - grn_set` 으로 출고만 있고 PO Invoiced에는 없는 건까지 포함돼, 한 'PO_금액' 컬럼에 출고 금액이 섞여 합계가 부풀어 보였음(P04 기준 599M = 553M + 46M).

매입 대사 의미를 분명히 하려고 `po_set - grn_set` 으로 단순화:

- GRN_미포함 = "PO Invoiced 인데 당월 GRN에 없는 건 → AX에 GRN 처리 필요"
- 컬럼: `AX PO`, `PO_금액` 두 개로 단순화 (47건 → 19건)
- raw_data_미포함 시트도 자동으로 PO 측만 남아 일관성 확보
- 범례·콘솔 요약 문구도 "AX에 GRN 처리 필요"로 명시

### 신규/수정 파일

- `po_generator/recon_paths.py` — 신규. `resolve_period_dir()`, `iter_period_dirs()`
- `reconcile_po.py` — `find_file()` / `load_grn()` / `build_raw_data()` / `_append_totals_block()` 신설
- `reconcile_so.py` — `find_ax_sales_file()` 경로 헬퍼 사용
- `reconcile_ind.py` — `find_orderbook_file()` (period 있을 때 / 없을 때 모두) + `ind_code_결과` 출력 경로
- `CLAUDE.md`, `docs/PO_RECONCILIATION.md` — 폴더 구조 설명 업데이트

---

## 2026-04-28: SO 단가/수량 무단 변동 감지 — 데이터 입력 오류 검출

오늘의 현황 페이지에 ⚠️ Customer PO 미변경 단가/수량 변동 섹션 추가. Excel에서 Customer PO는 그대로 두고 `Item qty` 또는 `Sales Unit Price` 가 바뀐 케이스를 감지 — 자릿수 실수(`126,000 → 1,260,000`), 단위 혼동(`25 → 99`) 같은 입력 오류와 매출/세금계산서/매출대사에 영향 가는 임의 변경을 잡기 위함.

### 데이터 흐름

`_sync_log` v2 (record-level JSON diff) 활용. SO 시트의 `change_type='수정'` 이벤트 중:

1. `Customer PO` 가 함께 변경된 경우 → "허가된 변경"으로 자동 제외
2. `Item qty` 또는 `Sales Unit Price` 가 실제로 변경됨
3. 빈값(`None`/공백) ↔ 값 변경 제외 — 최초 입력/삭제는 변경 아님
4. `dry_run=1` 동기화 제외
5. 이미 ack된 건 제외

### 노이즈 필터링 — 557건 → 53건

1차 구현(감시 필드 4개: Item qty / Sales Unit Price / Sales amount / Sales amount KRW)에서 557건 발생. 분석 결과:

- 456건(82%) — `Sales amount KRW` 단독 변경 (해외 환율 자동 재계산)
- 359건(64%) — 변경량 < 1원 (부동소수점 반올림)
- 단가/수량 실제 변경은 약 60건뿐

→ `Sales amount(KRW)` 는 사람의 액션이 아닌 Excel 수식·환율 자동 재계산 결과라 감시 제외. `Item qty` / `Sales Unit Price` 만 + 빈값↔값 제외 → **53건**으로 87% 감소. 잔존 건 모두 검토 가치 있음.

### Acknowledge 기반 dismiss

매출/세금계산서/매출대사에 영향이 있어 "오늘만 표시 → 내일 사라짐" 패턴은 위험. 단순 N일 윈도우 대신 영구 ack 방식 채택.

- 새 테이블 `_so_change_ack(sync_log_id PRIMARY KEY, acked_at, acked_by, note)` — `po_generator/db_schema.py` 에 `ensure_so_change_ack_table()` 추가
- `LEFT JOIN _so_change_ack ... WHERE a.sync_log_id IS NULL` 로 자동 필터
- 같은 SO·Line에 새로운 변경이 또 일어나면 별개 sync_log_id 라 다시 경고 (의도된 동작)

### UI

`pg_today()` 페이지 **맨 아래** (다른 검증 섹션과 일관). 0건이면 ✅ success 메시지로 명시 표시 ("Customer PO 미변경 단가/수량 변동 없음"). 1건 이상이면 노란 warning 배너 + expander 카드 목록.

카드 구성:
- 타이틀: `[국내] 업체명 · SOD-2026-0344 | 32 · Sales Unit Price: 3609000 → 4159000 · 감지시각`
- 본문 상단: 업체 / Customer PO / SO 한 줄 강조
- 변경 상세 표 (필드 / 이전값 / 변경값)
- 우측 [✅ 확인 완료] 버튼 → ack INSERT + 캐시 무효화 + rerun

### 한계 명시

`sync_time` 은 **Excel에서 변경된 시점이 아니라 sync_db.py 가 변경을 감지한 시점**. Excel 셀 단위 timestamp가 없어 sync 시점이 추적 가능한 가장 정확한 신호. `actor` 도 마찬가지로 "sync 실행자"이지 "Excel 수정자"가 아님.

### 신규/수정 파일

- `po_generator/db_schema.py` — `_so_change_ack` DDL + ensure 헬퍼
- `dashboard.py` — `load_so_unauth_changes()` / `_ack_so_change()` / `_render_so_unauth_changes()` / `_is_blank()` 추가, `pg_today()` 페이지 끝에 호출
- `tasks/todo.md` — Review 섹션 기록
- `tasks/lessons.md` — "변경 감지 / 알림 설계" 섹션 추가 (자동 재계산 파생필드 제외, 빈값↔값 제외, 분포 분석 우선)

---

## 2026-04-23: 동기화 로그 직관성 개선 — 주문 타임라인 탭 + 날짜×컬럼 히트맵

audit log 페이지가 너무 복잡해서 한 주문 검색 시 결과를 머릿속에서 재구성해야 했음. 두 가지 시각화를 추가해 raw event → narrative/pattern 변환을 UI가 대신 하도록 개선.

### 변경 1 — 페이지를 2개 탭으로 분리

기존 단일 페이지를 `📋 탐색`(기존) / `🕐 주문 타임라인`(신규) 2탭으로 재구성.

- 기존 body는 `_render_sync_log_explore()` 헬퍼로 추출 (감사·통계용 — 변경 없음)
- 새 `pg_sync_log()`는 thin wrapper로 탭만 분기

### 변경 2 — 🕐 주문 타임라인 (신규 탭)

**주문번호 하나** 입력 시 SO↔PO↔DN 관련 모든 변경 이력을 시간순 카드로 표시.

- `resolve_related_ids`로 연관 ID 자동 수집
- KPI: 총 이벤트 / 신규·수정·삭제 / 최초·최근 변경 / 관여자
- **날짜별 그룹핑** — `📅 2026-04-10 (Fri)` 헤더 후 해당 날짜 카드들
- 카드: 좌측 시각/사용자/sync#, 우측 시트·PK·변경 요약·전체 JSON expander
- 변경 요약: 첫 3개 컬럼만 inline (`Status: Draft → Confirmed`), 나머지 expander
- 200건 초과 시 최근 200건 표시 + 안내
- 신규 헬퍼: `_timeline_event_summary(row, max_inline=3) → (markdown, full_json)`

### 변경 3 — 📋 탐색 탭 안에 날짜 × 컬럼 변경 히트맵 추가

"어느 필드가 자주 바뀌나" 패턴 파악용 — 기존 시트별 변경 추이 차트 바로 아래.

- y축: 컬럼명 Top N (5~50, 기본 20), x축: 날짜
- 셀: 변경 횟수 (Blues 색상, 색 강도)
- 컨트롤 3종:
  - **시트 선택** — 시트마다 컬럼 집합이 달라 한 번에 한 시트
  - **카운트 모드** — `수정만`(기본) / `신규+수정` 라디오 토글
  - **Top N 슬라이더**
- **날짜 범위 슬라이더** — 시트의 실제 min/max 자동 디폴트, 단일 날짜면 caption만 표시
- x축 모든 날짜 강제 표시(`tickmode="array"`) — plotly 자동 스킵 비활성화
- 삭제는 제외 (row 단위라 컬럼 분석 의미 없음 — caption에 명시)

### 카운트 모드 토글 — 핵심 디자인 결정

신규 1건 시 모든 컬럼이 한꺼번에 +1 → "어느 필드가 사후에 자주 바뀌나" 패턴이 신규 일자 spike에 묻힘. 실데이터 SO_국내 기준:

| 모드 | 총 컬럼-이벤트 | Top 1 | 의미 |
|------|--------------:|-------|------|
| 신규+수정 | 31,627 | EXW NOAH (1,559) | 신규 spike에 운영 패턴 묻힘 |
| **수정만** | 3,311 | **Status (738)** | 사후 변경 패턴 명확 (Status·AX Period·CS담당자 등) |

→ 디폴트를 `수정만`으로. 토글로 활동 전체 보기도 가능.

### 코드 변경 — `dashboard.py`

- `_timeline_event_summary()` (신규) — 신규/수정/삭제별 markdown 요약 생성
- `_render_order_timeline()` (신규) — 타임라인 탭 렌더러
- `_render_sync_log_explore()` (기존 body 추출) — 탐색 탭 렌더러, 끝부분에 히트맵 섹션 추가
- `pg_sync_log()` (재구성) — title/caption + `st.tabs([...])` thin wrapper

---

## 2026-04-23: SO ↔ DN 불일치 검출 — "오늘의 현황" 데이터 무결성 섹션 추가

SO에 등록된 수량·단가와 DN(납품) 실적이 어긋나면 즉시 보이도록 대시보드 "오늘의 현황" 페이지 하단에 **⚠️ SO/DN 불일치** 섹션 추가.

### 검출 로직 — `(SO_ID, Line item)` 단위 2가지 체크

- **📦 과다출고**: `SUM(DN.Qty) > SO.[Item qty]` (tolerance 0.001)
  - 부분납품(DN 누계 < SO 수량)은 정상이므로 초과분만 플래그
- **💰 단가 불일치**: `SO.[Sales Unit Price] × SUM(DN.Qty) ≠ SUM(DN.[Total Sales])`
  - 임계값: `|diff| ≥ 1` **AND** 상대오차 `≥ 0.1%` — 반올림 노이즈 제거
  - 국내/해외 각 원본 통화(KRW/USD 등)끼리 비교 — FX 환산 노이즈 제거
- "완납 라인 총액 불일치"는 위 단가 체크에 수학적으로 흡수되므로 별도 카테고리로 두지 않음 (중복 신호 방지)

### UI — 국내/해외 탭 + 이상 유형 조건부 렌더

- caption: `📦 과다출고 N건 · 💰 단가 불일치 M건 (동시 K건)`
- expander 헤더 아이콘: `🔴📦` / `🔴💰` / `🔴📦💰` (두 이상 동시 발생 시)
- expander 내부는 해당 유형만 조건부 렌더
  - 과다출고: SO 수량 · DN 누계 수량 · 초과분
  - 단가 불일치: SO 단가 · DN 누계 수량 · 예상 금액 · 실제 금액 · 차이(%)
- DN 라인 detail 테이블(출고/선적일·수량·DN단가·DN금액)에 단가 불일치일 때만 "예상 (SO×수량)"·"차이" 컬럼 추가
- 사이드바 필터(market/sectors/customers) 적용

### 코드 변경 — `dashboard.py`

- `load_so_dn_anomalies()` (신설) — SO/DN INNER JOIN 후 `over_delivered`, `price_mismatch` 두 플래그 계산, 하나라도 걸린 라인만 반환 (과다출고 우선 정렬)
- `load_dn_lines_by_so_line()` (신설) — expander detail 용 DN 라인 로더
- `pg_today()` 끝부분에 "⚠️ SO/DN 불일치" 섹션 추가

### 검증

- 실데이터 962 완납 라인 전부 일치 → 0건 감지 (현재 데이터 건강, 향후 오입력 방지용)
- 합성 케이스 5종(과다출고만/단가만/둘다/부분납품/1원 노이즈) — 앞 3개만 플래그, 뒤 2개 정상 제외

---

## 2026-04-22: `_row_seq` float 타입 유출 버그 수정 — phantom 신규/삭제 쌍 해소

PO_해외 시트 전체(474건)가 매 sync마다 "신규 + 삭제" 쌍으로 잘못 기록되던 현상 수정.

### 원인

`_add_row_seq()`의 `cumcount() + 1`이 `Line item` group key에 숫자 셀이 섞여 있으면 `float64`로 승격 → DB TEXT 컬럼에 `"1.0"`으로 저장. 다음 sync에서 Excel이 만든 `1` (int)과 비교 시 타입 불일치로 전체 행이 다른 PK로 인식되어:
- 기존 DB의 `"1.0"` → Excel에 없는 것으로 판정 → **prune DELETE**
- Excel의 `1` → DB에 없는 것으로 판정 → **INSERT**

→ 같은 record가 신규+삭제로 중복 로그 (실제 데이터는 변경 없음).

### 수정 — 3-layer 방어

- `po_generator/db_sync.py: _add_row_seq()` — `.astype(int)` 명시적 cast
- `po_generator/db_sync.py: _sanitize_value()` — float 중 정수값(`1.0`, `2.0`)은 `int(1)`, `int(2)`로 변환
- `po_generator/db_sync.py: _normalize_pk()` — 방어적으로 `"1.0"` → `"1"` 정규화 (과거 오염 데이터 대응)
- DB 클린업: `UPDATE po_export SET _row_seq = SUBSTR(...)` — 기존 474행의 `"1.0"` → `"1"`

### 검증

수정 후 sync 재실행:
- 수정 전: PO_해외 474 phantom pair
- 수정 후: PO_해외 실제 변경 1건(신규+stale 빈 행 삭제)만 기록

---

## 2026-04-22: `_sync_log` v2 스키마 — 세션 메타·JSON 압축·삭제 스냅샷·actor 추가

전날 추가한 `_sync_log` v1 (필드당 1행, sync_time 문자열 그룹핑)을 audit/CDC best practice에
가까운 v2로 재설계. 기존 115,831행이 record 단위로 압축되어 15,263행 (**86.8% 압축**).

### 변경 — `_sync_runs` 신설 + `_sync_log` 재구성

- **`_sync_runs` (신설)** — 동기화 세션 메타
  - `sync_id INTEGER PK AUTO`, `started_at`, `ended_at`, `actor`, `host`, `dry_run`, `total_changes`, `note`
  - 동일 초에 두 번 sync 돌아도 sync_id로 명확히 구분 (이전 sync_time 문자열로는 충돌 가능)
- **`_sync_log` v2** — record 단위 + JSON
  - `id`, `sync_id` (FK → `_sync_runs`), `sheet_name`, `change_type`
  - `pk_json` (JSON 배열 `["A","B"]`) + `pk_display` (검색용 문자열 `"A | B"`) — 구조/검색 둘 다 만족
  - `changes_json` — 신규: `{col: val,…}`; 수정: `{col: {old,new},…}`; 삭제: NULL
  - `row_snapshot_json` — 삭제: 삭제 직전 전체 row JSON; 신규/수정: NULL
- 인덱스: `idx_sync_log_sync_id`, `idx_sync_log_sheet (sheet_name, sync_id)`, `idx_sync_log_pk (pk_display)`, `idx_sync_runs_started`

### 핵심 가치

| Best practice | v1 (어제) | v2 (오늘) |
|---|---|---|
| 세션 식별 | sync_time 문자열 (충돌 가능) | sync_id PK + actor/host 메타 |
| 신규/수정 표현 | 컬럼당 N행 (팽창) | record당 1행 + JSON (압축) |
| 삭제 audit | PK만 기록 (행 손실) | row_snapshot_json에 전체 row 보존 |
| PK 구조 | `"A | B"` 문자열 (역파싱 필요) | pk_json (배열) + pk_display (검색용) 병행 |

### 코드 변경

- `po_generator/db_schema.py`
  - `ensure_sync_log_tables()` — v2 스키마 DDL (구 `ensure_sync_log_table`은 별칭 유지)
  - `create_sync_run(conn, dry_run, note) → sync_id` — 세션 시작 (actor=USERNAME, host=hostname 자동 캡처)
  - `finalize_sync_run(conn, sync_id, total_changes)` — 종료 시각 + 카운트 갱신
- `po_generator/db_sync.py`
  - `SheetSyncResult.pruned_snapshots` 필드 추가 — 삭제 직전 row 캡처
  - 두 prune 분기(빈 시트 / 일반 prune) 모두 DELETE 직전 `SELECT *` 로 스냅샷 확보
- `sync_db.py`
  - `write_sync_log_to_db()` 전면 재작성 — record 단위 INSERT, sync_id로 그룹화, snapshot 포함
  - 한 sync 호출 = 1개 `_sync_runs` row + N개 `_sync_log` row
- `migrate_sync_log_v2.py` (신규)
  - 기존 `_sync_log` (v1) → `_sync_log_legacy` 백업 후 v2 INSERT
  - `--dry-run` (시뮬레이션) / `--drop-legacy` (백업 테이블 삭제)
  - 실행 결과: 115,831 → 15,263행 (86.8% 압축), 76개 sync_runs 복원

### 대시보드 업데이트

- `dashboard.py`
  - `load_sync_log()` — `_sync_log` JOIN `_sync_runs` 조회로 변경 (actor/host 포함)
  - `load_sync_runs()` 신규 — 세션 단위 로드
  - `pg_sync_log()` — KPI 5개 (총 record + 신규/수정/삭제 + 동기화 세션수)
  - **상세 보기 모드 토글** — 요약(record당 1행, "Status, 비고 외 3개") vs 펼침(컬럼당 1행, 기존 형태)
  - 세션 요약 테이블이 `_sync_runs` 기반 → 사용자/호스트/dry_run/시작·종료시각 표시
  - 컬럼명 검색이 changes_json/row_snapshot_json 모두에서 동작

### 문서

- `docs/DB_SYNC_GUIDE.md` — v2 스키마, JSON 예제, snapshot 활용법 추가
- `CLAUDE.md` — `migrate_sync_log_v2.py` + `_sync_runs` 항목 반영

### 마이그레이션 후 정리

`_sync_log_legacy` 테이블은 검증 기간 동안 보존. 검증 끝나면:
```bash
python migrate_sync_log_v2.py --drop-legacy
```

---

## 2026-04-21: 동기화 로그 DB 이관 + 대시보드 조회 페이지 + 라인 번호 UI

`sync_log.csv` (9.14 MB까지 증가해 Excel 열기 곤란)를 SQLite `_sync_log` 테이블로 이관.
대시보드에 변경 이력 조회 페이지 추가 + 오늘의 현황 expand 테이블에 Line 번호 표시.

### `_sync_log` 테이블 신설 + CSV → DB 전환

- `po_generator/db_schema.py` — `ensure_sync_log_table()` 추가
  - 스키마: `id / sync_time / sheet_name / change_type / pk / column_name / old_value / new_value`
  - 인덱스 2개: `sync_time`, `(sheet_name, sync_time)` — 기간/시트 조회 가속
  - 빈 문자열은 NULL로 저장 (SQL 쿼리 일관성)
- `sync_db.py` — 기존 `write_sync_log()` (CSV) 제거 → `write_sync_log_to_db()` 신규
  - `executemany` 배치 INSERT, sync 트랜잭션과 분리된 별도 커넥션
  - 로그 기록 실패가 이미 commit된 sync 결과를 훼손하지 않도록 설계
- `migrate_sync_log.py` 신규 — 기존 CSV → `_sync_log` 1회성 마이그레이션 스크립트
  - `--dry-run` 파싱 테스트 / `--delete` 성공 후 CSV 자동 삭제
  - 실행 결과: **115,141행** 이관 완료 (2026-03-13 ~ 2026-04-21)

### 대시보드 "동기화 로그" 페이지 (9번째 페이지)

- `dashboard.py` — `pg_sync_log()` + `load_sync_log(days)` + `resolve_related_ids(input_id)`
- 기능:
  - 조회 기간 선택 (7일/30일/90일/전체)
  - KPI 4개: 총 건수 / 신규 / 수정 / 삭제
  - 동기화 세션별 요약 테이블 + 시트별 변경 추이 차트 (plotly stacked bar)
  - 필터 2행: (시트, 변경유형) + (주문 통합 검색, PK 원본 검색, 컬럼명)
  - 필터 결과 CSV 다운로드
- **주문 통합 검색** (핵심 UX): PK 구조가 시트마다 달라 (`SO_ID|…`, `PO_ID|…|seq`, `DN_ID|SO_ID|…`) 한 번호로는 일부 시트만 매칭됨. 예) `ND-0429` 입력 시 PO만 22행 → SO 27행 누락. 해결:
  - po_domestic/po_export로 SO_ID ↔ PO_ID 매핑, dn_domestic/dn_export로 SO_ID ↔ DN_ID 매핑을 따라 2-pass로 연관 ID 전개
  - 사용자가 SO/PO/DN 어느 번호를 입력해도 3개 시트의 연관 이력을 모두 표시
  - 검증: `ND-0429` → 2개 ID (SO+PO) → 49행 반환 (이전 22행)

### 오늘의 현황 expand 테이블에 Line 번호 추가

- `dashboard.py` — PO 확정 지연, 미발주 현황 expand에 `Line` 컬럼 추가
  - `load_po_sent_pending()` SQL에 `CAST([Line item] AS INTEGER) AS line_item` 추가
  - 미발주 현황 expand에서 기존에 `drop(columns=["line_item"])`로 지우던 부분 → 유지
  - EXW 완료 미출고, 납기 현황 expand는 이미 `Line` 컬럼 있어 그대로

### 문서

- `docs/DB_SYNC_GUIDE.md` — CSV 관련 기술을 `_sync_log` 테이블 기반으로 전면 교체, SQL 조회 예제 + 대시보드 링크 + 마이그레이션 방법 추가
- `CLAUDE.md` — 대시보드 8→9페이지, `migrate_sync_log.py` 항목 추가, `db_schema.py` 설명에 `_sync_log` 명시

---

## 2026-04-16: CI/PL 템플릿 Customer PO 열 추가 (E열)

Commercial Invoice, Packing List 템플릿이 아이템 행별 Customer PO를 표시하도록 변경됨.
E20부터 행별 Customer PO를 기록하고, 기존 열은 모두 1칸씩 오른쪽으로 이동.
복수 PO가 섞인 통합 인보이스(`create_fi --po` 등)에서 PO 식별이 가능해짐.

### CI (Commercial Invoice)
- E열 = Customer PO (신규), F열 = Qty (E→F), G/H/I = Unit Price/Currency/Amount (유지)
- Total 행: F=SUM qty, G="EA" (한 칸 이동), H/I 유지
- G16 (PO No), C34 (Shipping Mark PO): 복수 PO는 `, ` 로 결합하여 전체 출력

### PL (Packing List)
- E열 = Customer PO (신규), F열 = Qty (E→F), G열 = Net Weight (F→G), H/I = Gross Weight/CBM (유지)
- Total 행: F=SUM qty, G=SUM net weight (구 `"KGS"` 라벨 제거), H/I 유지
- G16 (PO No), C36 (Shipping Mark PO): 복수 PO는 `, ` 로 결합하여 전체 출력

### 수정 파일
- `templates/commercial_invoice.xlsx`, `templates/packing_list.xlsx` — 레이아웃 변경
- `po_generator/ci_generator.py` — 열 상수 재정의, `_collect_customer_pos()` 헬퍼, `_fill_items_batch`/`_update_total_row` 수정
- `po_generator/pl_generator.py` — 동일 패턴 적용

---

## 2026-04-09: 대시보드 PO 미등록 감지 기능 추가

SO시트에만 있고 PO시트에 SO_ID가 아예 없는 건을 "PO 미등록"으로 분리 표시.
기존 "미발주"(PO가 없거나 Open)에서 데이터 입력 누락 리스크를 별도 식별.

### 오늘의 현황
- 미발주현황 카운트에 🔴 PO미등록 건수 별도 집계
- expander 태그에 `PO미등록` 레이블 표시

### 발주커버리지 페이지
- KPI 카드 6열 확장: `:red[PO 미등록]` 카드 추가 (진한 빨강)
- Stacked Bar 차트에 PO 미등록 구간 추가
- PO 미등록 상세 테이블 추가 (SO_ID, 고객명, 품목, 수량, 매출금액, 수주일, 납기일)

### calc_coverage() 상태 분리
- 기존: PO 없음 → "미발주"
- 변경: PO시트에 아예 없음 → "PO 미등록", PO Open만 → "미발주"

### 수정 파일
- `dashboard.py` — `calc_coverage()` 상태 분리, `pg_today()` 미발주현황, 발주커버리지 KPI·차트·상세
- `tests/test_dashboard.py` — PO 미등록 케이스 반영

---

## 2026-04-08: Final Invoice 발주번호(Customer PO) 기준 생성 기능 추가

- 발주번호(Customer PO)를 입력하면 복수 DN에 걸친 동일 발주번호 아이템을 통합하여 FI 생성
- 기존 DN_ID 기준 생성과 병행 가능 (하위 호환)
- CLI: `python create_fi.py --po 26KPO00144` (여러 건: `--po PO1 PO2`)
- `--po` 인자 없이 실행 시 사용 가능한 발주번호 목록 표시 (관련 DN, 고객명 포함)
- BAT 메뉴: [4] Final Invoice → 서브메뉴 분기 [1] DN_ID 기준 / [2] 발주번호 기준

### 수정 파일
- `create_fi.py` — `--po` CLI 인자, `generate_fi_by_po()`, `print_available_customer_pos()` 추가
- `po_generator/services/finder_service.py` — `find_dn_export_by_customer_po()` 추가
- `po_generator/services/document_service.py` — `generate_fi_by_customer_po()` 추가
- `create_po.bat` — FI 메뉴 서브메뉴 분기 (DN_ID / 발주번호)

---

## 2026-04-03: Industry Code 대사 + SO Sector 검증 기능 추가

### Industry Code 채움 (월별)
- Orderbook의 빈 Industry code를 NOAH_SO_PO_DN.xlsx PO→SO 매핑으로 자동 채움
- 매핑 체인: 발주번호 → PO시트 NOAH O.C No. → SO_ID → SO시트 Industry code
- 새 파일(`ind_code_결과_{period}.xlsx`)로 출력, 원본 미수정
- 추가 컬럼: NOAH Sector (SO시트 Sector), 매핑상태 (매칭/SO에 Industry code 없음/PO에 발주번호 없음)

### SO Sector 검증 (전체, period 무관)
- SO시트(국내+해외) Sector가 Industry code 마스터(Orderbook)의 Category와 일치하는지 교차 검증
- 마스터 매핑: Oil & Gas → OG, Water & Power → WAPO, Chemical, Process & Industrial → CPI
- 불일치 건 상세(SO_ID, 현재Sector, 기대Sector, Customer)와 교차표 요약 출력
- 별도 파일(`sector_검증.xlsx`) 생성, `--sector-only` 옵션으로 단독 실행 가능

### BAT 메뉴
- [I] 서브메뉴 분기: [1] 전체(채움+검증, 월 입력), [2] Sector 검증만(월 입력 불필요)

### 수정 파일
- `reconcile_ind.py` — Industry Code 대사 + Sector 검증 CLI (신규)
- `create_po.bat` — 메뉴에 [I] Industry Code 대사 서브메뉴 추가
- `CLAUDE.md` — Commands, Architecture, Key Files 업데이트

---

## 2026-04-03: Final Invoice 발주번호(RCK PO) 단위 자동 분리 생성

- DN_ID 내에 복수 RCK PO가 포함된 경우, 발주번호별로 FI를 자동 분리 생성
- 단일 RCK PO인 경우 기존 동작 유지 (하위 호환)
- 출력 파일명에 RCK PO 포함: `FI_{DN_ID}_{RCK_PO}_{고객명}_{날짜}.xlsx`

### 수정 파일
- `create_fi.py` — RCK PO 그룹 감지 + 발주번호별 루프 생성
- `po_generator/services/document_service.py` — `generate_fi()`에 `rck_po` 필터 파라미터 추가

---

## 2026-04-03: 대시보드 해외 선적 예정 카드에 Incoterms/운송방식 표시

- 오늘의 현황 → 해외 선적 예정 카드에 Incoterms, Shipping method 정보 추가
- 공장 픽업 카드와 동일한 `📦 Incoterms · 운송방식` 형식으로 표시

### 수정 파일
- `dashboard.py` — 해외 선적 예정 섹션에 SO 메타 조인 확장 + 카드 렌더링 추가

---

## 2026-04-01: SO 매출대사 기능 추가 (AX ERP vs NOAH DN)

### 기능 개요
- AX ERP 매출 금액과 NOAH DN 매출 금액을 AX Project 기준으로 비교하여 차이 확인
- `so_reconciliation/PXX/AX_Sales_PXX.xlsx` ↔ `NOAH_SO_PO_DN.xlsx` DN 시트 비교
- bat 메뉴 `[S]` SO 매출대사 옵션 추가

### 매출일 기준 월 필터
- **국내(DN_국내)**: `출고일` 기준으로 대사 월 필터링
- **해외(DN_해외)**: `선적일` 기준으로 대사 월 필터링 (출고일 ≠ 선적일, 매출 인식은 선적 시점)

### FX 환율차이 자동 판별
- DN 등록 시점 환율 vs 대사 월 환율 차이로 인한 불일치 자동 식별
- FX 시트에서 대사 월 환율 로드 → `외화금액 × 대사월 환율 ≈ AX 금액`이면 `일치(환율차이)` 판정
- 대사 시트에 `대사월_환율`, `재계산_KRW` 컬럼 포함

### 매칭상태
| 상태 | 설명 |
|------|------|
| 일치 | AX = NOAH DN (차이 < 1원) |
| 일치(환율차이) | 외화 × 대사월 환율 = AX (등록월 vs 대사월 환율 차이) |
| 불일치 | AX ≠ NOAH DN (환율차이로도 설명 안됨) |
| NOAH에 없음 | AX에 있지만 해당 월 DN에 매칭 안됨 |

### 출력 파일
- `대사결과_SO_{period}.xlsx` — 3시트: 대사(요약), 상세(DN 라인별), 범례

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `reconcile_so.py` | 신규 — SO 매출대사 CLI (AX Sales ↔ NOAH DN 비교, FX 환율차이 판별) |
| `create_po.bat` | 메뉴에 `[S] SO 매출대사` 추가 + `:reconcile_so` 섹션 |
| `CLAUDE.md` | Commands, Architecture, Key Files에 reconcile_so.py 추가 |

---

## 2026-04-01: Order Book Variance 분석 SQL 추가 + 스냅샷 퇴장 행 버그 수정

### Variance 분석 SQL (`sql/order_book_variance.sql`)
- 마감 스냅샷 간 소급 변경 내역을 변동이유별로 자동 분류
- **환율차이**: 해외 건, 수량 불변 금액만 변동 (Sales amount KRW 환율 소급 변경)
- **판매가변경**: 국내 건, 수량 불변 금액만 변동
- **수량변경**: SO 수량 소급 수정 또는 라인 추가/삭제
- **반올림**: KRW 환산 소수점 ±1원 이내
- 납기변경(EDD 수정)은 그룹키 이동일 뿐 금액/수량 변동이 아니므로 제외
- `params` CTE의 period 값을 변경하여 DB Browser에서 사용

### 스냅샷 퇴장 행 버그 수정 (`po_generator/snapshot.py`)
- **문제**: 전월 Ending > 0이었지만 소급 변경으로 사라진 건이 당월 스냅샷에 누락 → Start ≠ 전월 Ending (70.5M 차이 발생)
- **원인**: rolling SQL의 HAVING 필터가 Ending=0 + 당월 활동 없는 그룹을 제외 → 전월 Ending이 당월 Start에 반영되지 않음
- **수정**: 전월 스냅샷에 있지만 rolling 결과에 없는 그룹을 "퇴장 행"으로 추가 (Start=전월Ending, Variance=소급변경분, Ending≈0)
- 수정 후: P03 Start = P02 Ending = 3,174.1M (차이 0), Variance -6.7M 정확히 반영

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `sql/order_book_variance.sql` | 신규 — Variance 변동이유 분석 SQL |
| `po_generator/snapshot.py` | `take_snapshot()`에 퇴장 행 로직 추가 |

---

## 2026-04-01: Order_Book Power Query — 부분출고 잔고 이월 버그 수정

### 문제
- 부분출고된 SO가 다음 Period에 아예 나타나지 않음
- 예: SOO-2026-0025가 P03에서 부분출고(Output=10, Ending=25)됐는데 P04에 행 없음
- **원인**: Period 확장 시 `endPeriod = if [출고월] <> null then [출고월] else LastPeriod` — 출고 이력이 있으면 마지막 출고월에서 끊어버려 부분출고 건도 완납 건과 동일하게 처리

### 수정
- `endPeriod`를 항상 `LastPeriod`로 변경 — 모든 SO Line을 현재월까지 확장
- 롤링 계산 후 `ZeroFiltered` 단계 추가 — Start=Input=Output=Ending 모두 0인 행 제거 (완납 건 정리)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `docs/POWER_QUERY.md` | M 코드 `WithPeriodList` endPeriod 수정, `ZeroFiltered` 단계 추가, 도식/설명 업데이트 |

---

## 2026-04-01: 납기 현황 — PO EXW 보충 로직 추가

### SO exw_noah 누락 시 PO factory_exw로 보충
- SO 라인과 PO 라인이 1:1 대응하지 않는 케이스 대응 (예: SO 2라인 → PO 1라인 합본 발주)
- `load_po_detail()`에 `MIN(NULLIF(p.[공장 EXW date], ''))` 추가 — SO_ID 단위로 PO의 공장 EXW 집계
- 납기 현황 섹션 Step 2a: SO의 `exw_noah`가 NaT인 라인에 PO의 `factory_exw`를 보충

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `load_po_detail()` SQL에 `factory_exw` 컬럼 추가, 납기 현황 Step 2a PO EXW 보충 로직 |

---

## 2026-03-31: PO 매입대사 — AX PO 매핑 파일 추가

### AX_PO_매핑_{period}.xlsx 별도 출력 (2시트)
- `reconcile_po.py`에 `export_delivery_ax_po()` 함수 추가
- **국내_Delivery 시트**: Delivery 원본 행 유지 + `AX PO` 컬럼 추가 (`RCK ODER` 바로 뒤)
- **해외_PO 시트**: PO_해외 Invoiced 데이터 → `RCK ODER`(PO_ID), `AX PO`, `SO_ID`, `Customer`, `계산서금액`(Total ICO)
- 1:N 매핑(ND-xxxx → 복수 P######) 시 콤마로 합쳐서 표시, 행 복제 없음
- PO_ID별 집계: AX PO 콤마, 금액 합산
- **목적**: 회계팀이 AX 시스템에서 PO번호 기준 GRN 대사 작업 시 활용 (국내/해외 모두)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `reconcile_po.py` | `export_delivery_ax_po()` 추가 (국내 Delivery + 해외 PO 2시트), `main()`에서 해외 Invoiced 필터링 후 전달 |

---

## 2026-03-30: 해외선적 Action Items 개선 + DB Sync 출력 순서 변경

### 해외선적 DN 상세에 Incoterms / 운송방식 컬럼 추가
- `load_so()` SQL에 `Incoterms`, `Shipping method` 컬럼 추가 (so_export JOIN)
- DN 상세 테이블에 Incoterms, 운송방식 표시 (SO_ID 다음 위치)
- **운송방식별 현황** 탭 신규 추가 — Air/Sea/Courier 등 방식별 DN건수, 단계별 건수, 총수량, 총금액, 최대경과일

### DB Sync 결과 테이블 출력 순서 변경
- `sync_db.py --changes` 실행 시 변경 상세 → 로그 저장 → **요약 테이블이 맨 마지막**에 출력되도록 변경

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `load_so()` Incoterms/Shipping method 추가, 해외선적 DN 집계·상세에 컬럼 추가, 운송방식별 탭 추가 |
| `sync_db.py` | `print_summary()` 호출을 맨 마지막으로 이동 |

---

## 2026-03-27: 대시보드 캘린더 — 해외 선적 예정 표시

### 납기 캘린더에 해외 선적 예정 정보 추가
- `dn_export` 테이블의 `선적 예정일` 기준으로 캘린더 셀에 🚢 건수 표시
- 날짜 드릴다운 시 **🚢 해외 선적 예정** 섹션 추가 (EXW/픽업 다음, 납기/출고 이전)
  - DN별 고객명, 섹터, 고객PO, 수량/금액
  - 물류 타임라인: 출고 → 픽업 → 선적예정
  - B/L 번호, 운송 업체

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `build_calendar_data()` 선적예정 집계, 캘린더 셀 🚢 아이콘, 날짜 드릴다운 (E) 섹션 |

---

## 2026-03-26: 운영 신뢰성 P0 개선 (4건)

### DB 동기화 삭제 반영 (prune)
- Excel에서 지운 행이 `noah_data.db`에 잔류하던 문제 해결
- sync 시 DB에만 존재하는 PK를 자동 DELETE + 건수 리포트
- `sync_db.py` 출력 테이블에 "삭제" 컬럼 추가, `--changes`에 삭제 상세 표시
- `sync_log.csv`에 삭제 내역 기록, `--dry-run`에서도 삭제 예정 건수 확인 가능

### 출력 파일 덮어쓰기 방지
- 같은 주문을 같은 날 재생성 시 기존 파일 무경고 덮어쓰기되던 문제 해결
- 파일 존재 시 자동으로 `_1`, `_2`, ... 접미사 부여 (history.py의 기존 패턴과 동일)

### 이력 저장 실패 노출
- `save_to_history()` 실패 시 `logger.warning`만 찍고 성공 반환하던 문제 해결
- `DocumentResult.history_saved` 필드 추가 → CLI에서 `[주의]` 경고 표시
- exit code는 0 유지 (문서 자체는 성공)

### 대시보드 로더 실패 가시화
- 12개 데이터 로더에서 예외 발생 시 "데이터 없음"과 구분 불가하던 문제 해결
- `session_state` 기반 에러 수집 → 페이지 상단에 `st.warning()` 배너 표시
- 예: "일부 데이터 로드 실패 — SO: OperationalError: no such table ..."

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `po_generator/db_sync.py` | prune 로직 + `SheetSyncResult.pruned` 필드 |
| `sync_db.py` | 출력 테이블/로그에 삭제 건수 반영 |
| `po_generator/cli_common.py` | 파일 존재 시 접미사 자동 부여 |
| `po_generator/services/result.py` | `DocumentResult.history_saved` 필드 |
| `po_generator/services/document_service.py` | history 실패 → result 반영 |
| `create_po.py` | history 경고 출력 |
| `dashboard.py` | 로더 에러 수집 + 배너 표시 |

### 후속 보완 (3건)
- **빈 시트 prune 누락 수정**: `total_rows == 0`에서 early return하여 DB 잔류 행이 삭제되지 않던 문제 → 빈 시트에서도 prune 수행
- **dry-run 정확도 개선**: `:memory:` DB 대신 실제 DB에 연결 후 rollback 방식으로 변경 → 운영 DB 기준 정확한 diff 시뮬레이션
- **dry-run 트랜잭션 안전성**: `isolation_level=None` + 명시적 `BEGIN`으로 DDL(DROP/CREATE/ALTER TABLE)도 트랜잭션 내 실행 → rollback 시 완전 원복 보장
- **파일명 suffix 오버플로우 방어**: counter > 100 시 기존 파일 경로 반환 → `FileExistsError` 발생으로 변경

---

## 2026-03-25: 거래명세표(TS) 월합 데이터 중복 버그 수정 및 개선

### 버그 수정
- **아이템 중복 버그**: `load_dn_data()`에서 `SO_ID`만으로 DN↔SO merge → SO에 Line item이 많으면 N² 중복 발생
  - **원인**: DN(8행) dedup→1행 × SO(32행) = 32행 (정상은 8행)
  - **수정**: PO 로딩과 동일하게 `['SO_ID', 'Line item']` 복합키로 join

### 개선 사항
- **PO No. 복수 표시**: 월합 거래명세표에서 발주번호가 여러 개일 때 콤마로 구분하여 모두 표시
- **아이템별 출고일**: 월/일 컬럼에 각 아이템의 `출고일`을 개별 표시 (기존: 단일 날짜 일괄 적용)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `po_generator/utils.py` | `load_dn_data()` merge 키를 `['SO_ID', 'Line item']` 복합키로 변경 |
| `po_generator/ts_generator.py` | PO No. 복수 표시, 아이템별 출고일 표시 |

---

## 2026-03-25: 대시보드 발주 커버리지 상세 테이블 개선

### 변경 내용
- 미발주/부분발주/발주진행중 상세 테이블에 **수주일**, **공장발주일** 2개 날짜 컬럼 추가
  - 수주일: SO `PO receipt date` (고객→RCK 발주일)
  - 공장발주일: PO `공장 발주 날짜` (RCK→NOAH 공장 발주일)
- **PO_ID** 컬럼 추가 (SO_ID 옆, 복수 PO 시 쉼표 구분)
- 세 테이블 모두 **국내 | 해외** 탭으로 분리
- 수주일 기준 오래된 순 정렬

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `load_po_detail()` SQL에 `po_ids`, `factory_order_date` 추가, `calc_coverage()`에 `po_receipt_date` 집계 추가, 3개 상세 테이블 국내/해외 탭 분리 및 컬럼 확장 |

---

## 2026-03-25: 오늘의 현황 — 날짜 상세 접기/펼치기

### 변경 내용
- 납기 캘린더 날짜 상세 섹션을 `st.expander`로 변경 (클릭하여 펼침/접기)
- 라벨에 건수 요약 표시 (예: `📅 2026-03-25 상세 — EXW 2 · 납기 3 · 출고 1`)
- 오늘 날짜는 기본 펼침, 다른 날짜 선택 시 접힌 상태

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `_render_delivery_calendar` 하단 날짜 상세를 `st.expander`로 래핑, 건수 미리 계산 |

---

## 2026-03-25: 오늘의 현황 — 미발주 현황 섹션 추가

### 변경 내용
- PO 확정 지연과 EXW 완료 미출고 사이에 **미발주 현황** 섹션 신규 추가
- `calc_coverage()` 재활용하여 미발주 + 부분발주 건 표시
- 수주일 기준 경과일 버킷 분류 (7일/14일/30일+), 수주일 미입력 건은 `⚪ 수주일 미입력` 버킷
- 국내 | 해외 탭 분리
- 카드에 Sales amount + ICO total 금액 동시 표시
- 클릭하면 SO line item 상세 테이블 (품목명, OS name, 수량, 매출금액, 수주일, 공장발주일, 납기일, PO_ID, Status)
- `open_po_ids` 필드 추가: Open 상태 PO만 표시 (이미 발주된 PO 제외)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `load_po_detail()` SQL에 `open_po_ids` 서브쿼리 추가, `calc_coverage()`에 `open_po_ids` 전달, `pg_today()`에 미발주 현황 섹션 추가 |

---

## 2026-03-23: FI Total 행 Currency 열 수정

### 변경 내용
- Final Invoice Total 행의 Currency 표시 열을 H → G로 변경 (아이템 행의 Currency 열과 일치시킴)

### 수정 파일
- `po_generator/fi_generator.py` — `_update_total_row()` 내 Currency 셀 H→G
- `docs/TEMPLATE_MAPPINGS.md` — FI Total 행 매핑 H→G

---

## 2026-03-23: 대시보드 테마 전환 토글 추가

### 변경 내용
- 사이드바 상단에 테마 전환 토글 추가
- 시스템 테마(dark/light) 자동 감지 → 토글 시 반대 테마로 전환
  - 시스템 다크 → ☀️ Light Mode 토글 표시
  - 시스템 라이트 → 🌙 Dark Mode 토글 표시
- CSS injection으로 배경, 사이드바, 텍스트, 버튼, 콤보박스, 캘린더, expander 등 전체 UI 커버
- Plotly 차트 `pio.templates.default`를 `"plotly"` / `"plotly_dark"`로 전역 전환

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | 테마 감지(`st.context.theme` / `st.get_option` fallback), 토글 UI, Light/Dark CSS, Plotly 템플릿 전환 |

---

## 2026-03-23: PO 확정 지연 — 품목명 컬럼 추가

### 변경 내용
- `load_po_sent_pending()` SQL에 `[Item name]` 컬럼 추가 (국내/해외 양쪽)
- PO 확정 지연 expander 내 detail 테이블에 **품목명** 컬럼 표시 (PO_ID 다음 위치)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `load_po_sent_pending()` SQL에 `item_name` 추가, detail 테이블 컬럼에 `품목명` 포함 |

---

## 2026-03-23: 오늘의 현황 — 전 섹션 Sector 표시 추가

### 배경
대시보드 "오늘의 현황" 페이지의 카드/expander에서 어떤 섹터의 건인지 바로 파악할 수 없었음.

### 변경 내용

**Sector 정보 표시 — 5개 영역 일괄 적용**

| 영역 | 표시 위치 | 형식 |
|------|-----------|------|
| 날짜 카드 — EXW 출고 예정 | 카드 타이틀 | `고객명 · Sector` |
| 날짜 카드 — 공장 픽업 | 카드 타이틀 | `고객명 · Sector` |
| 날짜 카드 — 납기 예정 | 카드 타이틀 | `고객명 · Sector` |
| 날짜 카드 — 출고 실적 | 카드 타이틀 | `고객명 · Sector` |
| PO 확정 지연 | expander 헤더 | `고객명 [Sector]` |
| EXW 완료 미출고 | expander 헤더 | `고객명 [Sector]` |
| 납기 현황 (미완료 건) | expander 헤더 | `고객명 [Sector]` |
| 해외 선적 Action Items | expander 헤더 | `고객명 [Sector]` |

- Sector가 비어있는 건은 태그 미표시 (빈 문자열 처리)
- 공장 픽업 카드: SO 메타 조인에 `sector` 컬럼 추가 (기존 `customer_po`만 가져오던 것)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | 8개 섹션 agg에 `섹터` 추가, 타이틀/헤더에 sector 태그 표시 |

---

## 2026-03-22: 오늘의 현황 — PO 확정 지연 / EXW 미출고 / 납기 현황 개선

### 배경
오늘의 현황 페이지에 공장 발주→출고→납품 파이프라인의 병목을 단계별로 모니터링하는 섹션 추가 및 기존 섹션 개선.

### 신규 섹션

**📋 PO 확정 지연 (Sent → Confirmed 미전환)**
- PO 테이블의 `공장 발주 날짜` 기준 경과일 계산
- Status = "Sent"인 PO 라인만 캡처
- 국내/해외 탭 분리, 7/14/30일+ 버킷 그룹화
- 새 로더: `load_po_sent_pending()`

**🚨 EXW 완료 미출고 (PO 공장 EXW < 오늘 & 미Invoiced)**
- 기존 "EXW 출고지연" 대체 — SO 기반 → **PO line item 기반**으로 전면 재작성
- PO 테이블의 `공장 EXW date` + `Status` 기준 (SO의 exw_noah이 아님)
- Invoiced/Cancelled 제외, EXW 경과 라인만 정확히 캡처
- 새 로더: `load_po_exw_pending()`
- 설명: "공장에 EXW date 재확인 필요"

### 기존 섹션 개선

**📦 납기 현황 (미완료 건) — DN qty 레벨 매칭 추가**
- 기존: SO status 기반 단순 표시 → 개선: DN qty 매칭으로 부분출고 정확 반영
- delivery_date < 오늘 AND (DN 미생성 OR 출고 qty < 주문 qty)
- 잔여수량/잔여금액 표시 (예: "잔여 11/14 · ₩1,568만/₩1,788만")
- 설명: "DN 발급 또는 납기 일정 확인 필요"

**🚢 해외 선적 Action Items — 그룹화 개선**
- 선적 대기 / 포워더 미정 탭 분리
- 공장출고일 기준 경과일 버킷 그룹화 (7/14/30일+)

### 공통 변경

**Expander + 테이블 UI 패턴 적용 (4개 섹션 모두)**
- 카드 렌더링 → `st.expander` + `st.dataframe` 전환
- 접었다 펼치면 line item 상세 테이블 표시
- 버킷 헬퍼: `_OVERDUE_BUCKETS`, `_assign_bucket()`, `_render_bucketed_cards()`

### 핵심 설계 판단
- **EXW 섹션은 PO 데이터 기반**: SO의 EXW NOAH은 계획일, PO의 공장 EXW date가 실제 출고일
- **납기 섹션은 SO+DN 데이터 기반**: 납기 경과 라인만 표시 (미래 납기 라인 제외)
- **부분출고 qty 레벨 매칭**: DN line item별 출고수량 vs SO 주문수량 비교

---

## 2026-03-21: 대시보드 제품/고객 분석 버그 수정

### 배경
코드 리뷰에서 제품분석·고객분석 페이지의 엣지 케이스 버그 및 해석 왜곡 5건 발견.

### 수정 내용

**[High] RFM qcut 예외 — 소수 고객 필터 시 크래시**
- `pd.qcut(q=4)` 고정 → `q=min(4, n_customers)` 동적 축소, 1명이면 중간 점수 고정
- 등급 경계도 q 비례 산출 (`_max * 10/12` 등), 설명 문구도 동적 표시

**[Medium] 제품 집중도 Top 3 비중 0% 표시**
- `len(by_amt) >= 3` 조건 → `len(by_amt) >= 1` (head(3)이 자동 truncate)

**[Medium] 고객 Pareto 누적비율 분모 왜곡**
- 분모: Top 20 합계 → 전체 고객 매출 합계(`by_cust.sum()`)

**[Medium] 연/월 필터 불일치 (신규 제품/고객/리텐션)**
- 신규 제품·신규 고객·리텐션: 전체 기간으로 "최초" 계산 → `display_periods`로 표시만 필터 한정
- RFM: `so_all_raw`(전체기간) → `so`(필터 적용) 변경

**[Low] RFM Recency 기준점**
- `_THIS_MONTH` 고정 → `so["period"].max()` (데이터 최신월 기준, 동기화 지연 대응)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | 제품/고객 분석 5건 버그 수정 |
| `dashboard_dist/dashboard.py` | 배포판 리빌드 (동일 수정 반영) |

---

## 2026-03-21: 대시보드 Portable 배포판 빌드

### 배경
대시보드를 다른 팀원에게 배포할 때 Python/패키지 설치 없이 바로 실행 가능한 형태가 필요.

### 변경 내용

**`dashboard_dist/` — 완전 무설치 배포 패키지**
- `build_portable.py`: Python Embedded 3.11.9 다운로드 + pip/streamlit/pandas/plotly 자동 설치 → `NOAH_Dashboard/` 폴더 생성 (~414MB)
- `build_dist.py`: 원본 `dashboard.py`에서 `po_generator` 의존성 제거한 standalone 버전 자동 생성
- `NOAH Dashboard.bat`: 더블클릭으로 Streamlit 실행, 빈 포트 자동 탐색
- `NOAH Dashboard.vbs`: 콘솔 숨김 버전
- `dashboard_config.ini`: DB 경로 설정 (비워두면 자동 탐색)

**DB 경로 자동 탐색**
- `C:\Users\{누구든}\OneDrive*\` 하위에서 `noah_data.db`를 BFS 탐색 (최대 5단계)
- OneDrive 공유 파일 경로가 사용자마다 달라도 자동 인식
- `rglob` 대신 depth-limited BFS 사용 (OneDrive 경로 길이 초과 에러 방지)
- 우선순위: config.ini 명시 경로 → OneDrive 자동 탐색 → 현재 폴더

**배포 방법**
1. 개발자: `python build_portable.py` → `NOAH_Dashboard/` 폴더 생성
2. 배포: 폴더 통째로 복사 (또는 zip)
3. 사용자: `NOAH Dashboard.bat` 더블클릭 — 사전 설치 불필요

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard_dist/build_portable.py` | 신규 — Portable 빌드 스크립트 |
| `dashboard_dist/build_dist.py` | 신규 — standalone dashboard 변환 |
| `dashboard_dist/launcher.py` | 신규 — exe 런처 소스 |
| `dashboard_dist/dashboard_config.ini` | 신규 — DB 경로 설정 |
| `dashboard_dist/requirements.txt` | 신규 — 패키지 목록 |

---

## 2026-03-20: 대시보드 발주 커버리지 + 수익성 분석 + Order Book 3탭 + 세금계산서 미발행

### 배경
기존 대시보드(6페이지)는 매출/출고 중심이며, PO 테이블(발주/원가)을 거의 활용하지 않았음.
SO↔PO 조인으로 발주 커버리지·마진 분석 2개 신규 페이지를 추가하고, Order Book을 제조업 Best Practice 3탭 구조로 재편.
세금계산서 미발행 현황 섹션 추가로 출고 후 후속 조치 누락 방지.

### 변경 내용

**신규 데이터 로더**
- `load_po_detail()`: PO SO_ID 단위 집계 (PO line_item은 SO와 1:1 대응 안 함 — 본체+부속 합산 발주 등), Cancelled 제외
- `load_dn_tax_pending()`: 국내 DN 세금계산서 미발행 건 (출고 완료 but 세금계산서/선수금 세금계산서 미발행, 금액 0원·N/A 제외)

**신규 페이지: 발주 커버리지**
- `calc_coverage()` 순수 함수: SO_ID 단위 집계 → PO 존재 + Status 기반 판정
  - 미발주: PO 없음 또는 Open만 (공장 발주 전)
  - 부분 발주: Open + Sent/Confirmed 혼합 (일부만 발주)
  - 발주 진행중: Sent (공장에 발주, 확인 대기)
  - 발주 확정: Confirmed/Invoiced (공장 확인/출고 완료)
  - 발주취소: 모든 PO Cancelled → 분석에서 제외
  - 출고 완료 SO → 분석에서 제외
- KPI 카드 5개: 미발주/부분발주/발주진행중/발주확정/발주필요금액
- Stacked bar 커버리지 요약
- 미발주/부분발주/발주진행중 상세 테이블
- 고객별 발주필요금액 Top 10, 섹터별 커버리지율
- PO Status 파이프라인 (Open/Sent/Confirmed/Invoiced)

**신규 페이지: 수익성 분석**
- `calc_margin()` 순수 함수: SO_ID 단위 집계 → margin_amount, margin_pct, has_cost
- KPI 카드 4개: 총매출, 총원가(ICO), 총마진, 마진율
- 월별 마진 추이 (매출 vs 원가 bar + 마진율 line)
- 3탭 분석: 고객별 / 섹터별 / 모델별 (Top 15 마진율 bar + 상세 테이블)
- 저마진 경보 Top 10 (마진율 < 20%, ProgressColumn)
- 미출고금액 Top 10 (고객별)

**Order Book 3탭 재구조화**
- Executive 탭:
  - 워터폴 차트: 월별(selectbox) / 누적 토글
  - 3대 KPI: Backlog Cover / Past Due Ratio / Book-to-Bill
  - 월별 추이, 섹터별/고객별 Backlog
- Risk 탭: Aging 분석 (bar+pie+드릴다운), 고금액 위험건 Top 10, 납기 분포 히트맵
- Conversion 탭:
  - 전환 퍼널 (SO→PO→DN), 전환율 메트릭
  - 리드타임 분석: KPI 카드(평균/중앙값/최단/최장) + Box plot + 월별 추이
  - 데이터 기반 동적 차트 해석 (통계 수치·이상치·병목 구간 자동 분석)
  - 이상치 상세 테이블 (expander)
  - 해외 물류 리드타임 (출고→픽업, 픽업→선적) — 구간별 KPI + 비교 분석
  - 사이드바 필터(market/sector/customer) 적용

**오늘의 현황 개선**
- 세금계산서 미발행 섹션 추가 (국내 전용)
  - KPI 카드 4개: 미발행 건수/금액/최장 경과일/30일 초과
  - Aging 바 차트 (7일/14일/30일/30일+ 구간)
  - 고객별 미발행 금액 Top 10
  - 상세 테이블 (expander)
  - 금액 0원, 세금계산서 발행일 N/A 건 제외
- 백로그 요약 섹션 제거

**사이드바**
- 8페이지: 오늘의 현황, 수주/출고 현황, 제품 분석, 섹터 분석, 고객 분석, 발주 커버리지(NEW), 수익성 분석(NEW), Order Book

### 버그 수정
- `excel_generator.py`: 국내 PO Description에서 Item name 빈 경우 Model fallback 누락 → 해외와 동일하게 Model 사용
- `excel_generator.py`: Description 시트 A열 `-40` 등 숫자형 문자열이 xlwings에 의해 정수로 변환 → `number_format='@'` 적용
- `config.py`: `OPTION_FIELDS` 상수 불일치 (`MOV사양`→`MOV조립`, `VALVE 사양`→`VALVE 가격`)
- `dashboard.py`: Conversion 탭 해외 물류 리드타임에 사이드바 필터 미적용 → 필터 적용 + `market=국내` 시 비표시
- `dashboard.py`: Conversion 퍼널/리드타임에서 DN 필터 누락 → `enrich_dn()` + `filt()` 적용
- `dashboard.py`: pandas FutureWarning (`.fillna()` 다운캐스팅) → `.where()` 패턴으로 대체

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `load_po_detail()`, `load_dn_tax_pending()` 로더, `calc_coverage()`/`calc_margin()` 순수 함수, `pg_po_coverage()`/`pg_margin()` 신규 페이지, `pg_orderbook()` 3탭 재구조화(워터폴 토글, 리드타임 동적 해석, 이상치 상세), 세금계산서 미발행 섹션, 사이드바 8페이지, Conversion 필터 수정, 백로그 요약 제거 |
| `po_generator/excel_generator.py` | 국내 Description Model fallback, Description 레이블 텍스트 형식 |
| `po_generator/config.py` | `OPTION_FIELDS` 상수 업데이트 |
| `tests/test_dashboard.py` | `TestCalcCoverage` 8개 + `TestCalcMargin` 5개 테스트 추가 (64 passed) |
| `docs/CHANGELOG.md` | 변경 이력 기록 |

---

## 2026-03-20: 대시보드 오늘의 현황 대폭 강화

### 배경
"오늘의 현황" 페이지에서 공장 출고/픽업 예정, EXW 지연 건을 한눈에 파악할 수 없었음. 날짜 드릴다운에 EXW·픽업 정보 추가, 캘린더에 아이콘 표시, EXW Overdue 독립 섹션 신설.

### 변경 내용

**날짜 드릴다운 섹션 확장 (캘린더 날짜 선택 시)**
- **(A) 🏭 EXW 출고 예정**: SO 국내+해외 `EXW NOAH` 기준, 국내(🇰🇷)/해외(🌏) 구분 태그
  - 납기: `Requested delivery date` 표시, 없거나 1900년이면 "ASAP"
  - `Expected delivery date`는 "납품 예정일"로 별도 표시
- **(B) 🚛 공장 픽업**: DN_해외 `공장 픽업일` 기준, 운송 업체·선적예정일 표시
- (C) 📦 납기 예정, (D) 🚚 출고 실적: 기존 유지

**캘린더 히트맵 아이콘 추가**
- 🏭 `N건` — EXW 출고 예정
- 🚛 `N건` — 공장 픽업 예정
- 📦 / 🚚 — 기존 납기/출고 아이콘 유지
- `build_calendar_data()` — `ship_df` 옵션 파라미터 추가, `exw_count`·`pk_count` 집계

**🔴 EXW 출고 지연 섹션 신설 (캘린더 아래, 독립 영역)**
- `load_po_status()` 신규: PO 국내+해외 `SO_ID`별 Status 로딩
- EXW NOAH < 오늘 AND PO Status ≠ Invoiced → 지연 건 카드 표시
- 카드: 지연 일수(`N일 지연`), 요청납기, PO 상태(Open/Sent/Confirmed 등)

**해외 선적 Action Items 개선**
- 🔴 포워더 미정 / ⏳ 선적 대기 — 운송 업체 유무로 두 그룹 분리 표시
- 운송 업체 빈 값: "arranging..." 표시, 🔴 아이콘으로 시각 구분
- 고객 PO: DN_ID에 여러 PO 포함 시 콤마 구분으로 모두 표시
- 운송 업체(`[운송 업체]` 컬럼) 카드에 항상 표시

**데이터 로더 개선**
- `load_so()`: `Requested delivery date` 컬럼 추가, `EXW NOAH` 1900년 → NaT 처리
- `load_dn_export_shipping()`: `[운송 업체]` 컬럼 추가
- 날짜 선택 기본값: 현재 달이면 오늘 날짜, 다른 달이면 1일

**버그 수정**
- `_render_delivery_calendar()` 내 `so` 미정의 변수 → `so_pending`으로 수정 (`NameError` 해결)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `load_po_status()` 신규, `load_so()` SQL 확장 (requested_date, 1900년 필터), `load_dn_export_shipping()` carrier 추가, `build_calendar_data()` EXW/픽업 집계, 캘린더 아이콘, 날짜 드릴다운 4섹션, EXW Overdue 섹션, 선적 Action Items 그룹 분리, 날짜 기본값, `so` NameError 수정 |

---

## 2026-03-19: 대시보드 카드 UI 전환 및 Order Book 개선

### 배경
"오늘의 현황" 페이지의 메트릭과 테이블이 엑셀과 차별 없음 → 카드 UI로 시각화 개선.
Order Book 페이지에 납기 분포 히트맵 추가, 불필요한 섹션 정리.

### 변경 내용

**카드 UI 전환 (`st.container(border=True)` + 아이콘)**
- `_render_cards()` 헬퍼 추가 — 3열 격자 카드 렌더러
- KPI 4개: `st.metric()` → 아이콘 카드 (🔴/🟢 납기, 📥 수주 전월비, 📤 출고 달성률 progress bar)
- 납기 현황: `st.dataframe()` → SO_ID별 카드 (상태 아이콘 + 품목/수량/금액 + 납기/EXW)
- 선적 대기: `st.dataframe()` → DN_ID별 카드 (⏳ + 출고→픽업→선적 flow + B/L)
- 캘린더 상세: 납기 예정 → SO_ID별 카드, 출고 실적 → DN_ID별 카드 (2열 격자)
- Customer PO 정보 전 카드에 표시 (`load_so()` SQL에 `Customer PO` 컬럼 추가)

**삭제된 섹션**
- 해외 최근 선적 완료 (7일 이내)
- 국내 최근 출고 (7일) — metric + 카드
- Backlog 추이 — 마감 확정치 (스냅샷 차트)
- Backlog 상세 테이블
- 납기 지연 `st.warning()` — KPI 카드에 흡수

**Order Book 개선**
- 납기 분포 히트맵 추가: 금월~연말, 월별 × 섹터, YlOrRd colorscale, 셀에 금액 표시
- `load_order_book()`: 예외를 빈 DataFrame으로 삼키던 버그 수정 → 예외 전파하여 `st.error()` 경로 정상화

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `_render_cards()` 헬퍼, KPI 카드화, 4개 섹션 카드 전환, Customer PO 추가, 납기 히트맵, 섹션 삭제, `load_order_book()` 예외 수정 |

---

## 2026-03-18: 대시보드 납기 캘린더 추가

### 배경
"오늘의 현황" 페이지에 KPI 메트릭과 테이블만 있어 월 단위 시각적 조망이 불가능. Plotly Heatmap 기반 달력으로 납기 예정(SO)과 출고 실적(DN)을 한눈에 파악할 수 있도록 개선.

### 변경 내용
- **`build_calendar_data()`**: 순수 함수 — SO 납기일/DN 출고일을 날짜별로 집계 (so_count, so_amount, dn_count, dn_amount)
- **`_render_delivery_calendar()`**: Plotly Heatmap 캘린더 UI
  - Session state 기반 월 네비게이션 (◀ 이전 달 / 다음 달 ▶)
  - Diverging colorscale: 빨강(과납기) ↔ 흰색(0건) ↔ 파랑(미래 납기)
  - 셀 텍스트: 📦 납기 예정 건수/금액 + 🚚 출고 실적 건수/금액
  - 오늘 날짜 테두리 강조 (`add_shape`)
  - 날짜 클릭 → 드릴다운: 납기 예정 테이블 + 출고 실적 테이블 (2컬럼 레이아웃)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | `import calendar` 추가, `build_calendar_data()` + `_render_delivery_calendar()` 함수 추가, `pg_today()` 내 KPI 직후 캘린더 렌더 호출 |
| `tests/test_dashboard.py` | `TestCalendarData` 클래스 5개 테스트 추가 (47 passed) |

---

## 2026-03-18: 대시보드 Interactive Charts 추가

### 변경 내용
- **Hover Templates**: 전체 21개 Plotly 차트에 KRW 포맷(`₩%{y:,.0f}`), 한글 라벨, `<extra></extra>` 적용
- **드릴다운 (on_select)**: 제품 Top 15 / 섹터 Pie / 고객 Top 15 / Aging Bar 클릭 시 하위 상세 분석 표시
  - 제품: 월별 추이 + 섹터 비중 + 주요 고객 Top 5
  - 섹터: 제품 믹스 + 월별 추이 + 주요 고객
  - 고객: 월별 추이 + 제품 믹스 + Backlog 현황
  - Aging: 해당 구간 Backlog 상세 테이블
- **Rangeslider**: 수주/출고 월별 추이, Order Book 월별 추이에 rangeslider 추가
- **최소 버전**: `streamlit>=1.35.0` (on_select 파라미터 요구)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `dashboard.py` | 21개 차트 hovertemplate, 4개 드릴다운, 2개 rangeslider |
| `tests/test_dashboard.py` | 드릴다운 필터링 로직 12개 테스트 추가 (42 passed) |
| `requirements.txt` | `streamlit>=1.30.0` → `streamlit>=1.35.0` |

---

## 2026-03-18: Streamlit 대시보드 추가 + 개선

### 배경
NOAH_SO_PO_DN.xlsx 기반 문서 자동화 시스템에 비즈니스 현황을 한눈에 파악할 수 있는 대시보드가 없었음. SQLite DB(noah_data.db)를 데이터 소스로 활용하여 Streamlit 대시보드를 구축. 이후 피드백을 반영하여 전면 개선.

### 신규 파일
| 파일 | 역할 |
|------|------|
| `dashboard.py` | Streamlit 대시보드 앱 (6페이지, ~800줄) |

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `requirements.txt` | `streamlit>=1.35.0`, `plotly>=5.18.0` 추가 |
| `CLAUDE.md` | `streamlit run dashboard.py` 커맨드 추가 |
| `create_po.bat` | `[D]` 대시보드 메뉴 추가 |

### 대시보드 페이지 구성 (6페이지)

| 페이지 | 핵심 내용 |
|--------|----------|
| 오늘의 현황 | KPI, **납기 캘린더** (Plotly Heatmap, 월 네비, 날짜 클릭 드릴다운), 납기 현황 (국내/해외 탭, SO_ID 그룹, Status 아이콘, 지연 경고), 해외 선적 Action Items (선적 대기/최근 완료), 국내 최근 출고, 백로그 요약 |
| 수주/출고 현황 | 전월 대비 증감율, Book-to-Bill 비율/추이, 월별 수주/출고 + 누적매출, 금월 일별 출고 |
| 제품 분석 | Top 15 매출, 구성비 도넛, 월별 추이, 제품별 평균 단가, 제품별 Backlog Top 10 |
| 섹터 분석 | 섹터별 비중, 파이/월별 stacked bar/제품 믹스, 섹터별 Backlog, 평균 주문 규모 |
| 고객 분석 | Top 15, Pareto, 고객 상세 (Backlog 병합), 고객별 월별 매출 추이 (Top 5) |
| Order Book | Backlog KPI (지연/임박), 월별 Input/Output/Ending 추이 (`order_book.sql` 직접 실행), Aging 분석 (6구간), 섹터별/고객별 Backlog, 스냅샷 추이, 상세 테이블 |

### 데이터 레이어 (7개 캐시 로더)
- `load_so()` — SO 통합 (Status, EXW NOAH 포함)
- `load_dn()` — DN 통합 (매출 기준: 국내=출고일, 해외=선적일)
- `load_dn_export_shipping()` — 해외 DN 선적 파이프라인 (출고일/픽업일/선적예정일/선적일/B/L)
- `load_backlog()` — 현재 백로그 (`order_book_backlog.sql` 이벤트 패턴)
- `load_order_book()` — 월별 Order Book (`sql/order_book.sql` 파일 직접 실행)
- `load_sync_meta()` — 동기화 메타정보
- `load_snapshot_meta()` — 스냅샷 메타정보

### SQL 파일 활용
| SQL 파일 | 대시보드 활용 |
|----------|-------------|
| `order_book.sql` | `load_order_book()` — 파일 직접 읽어서 실행, 월별 Input/Output/Ending 추이 |
| `order_book_backlog.sql` | `load_backlog()` — 같은 이벤트 기반 패턴 인라인 SQL |

### 사이드바 필터
시장 구분(전체/국내/해외), 연도/월, 섹터 multiselect, 고객 필터, 새로고침 버튼

### 사용법
```bash
streamlit run dashboard.py
# 또는 create_po.bat → [D] 대시보드
```

---

## 2026-03-12: CI/PL 템플릿 셀 위치 전면 업데이트 (Bill to 3줄 확장)

### 배경
CI/PL 템플릿의 Consigned to 영역이 1줄(주소+국가+Tel+Fax) → 3줄(Bill to 1/2/3)로 확장되면서 Row 13 이하가 1행씩 밀림. 코드의 셀 상수들을 현재 템플릿에 맞게 전면 업데이트.

### 변경 내용

#### 1. 셀 상수 변경 (`ci_generator.py`, `pl_generator.py`)
- `CELL_CONSIGNED_TO/COUNTRY/TEL/FAX` 삭제 → `CELL_BILL_TO_1(A9)`, `CELL_BILL_TO_2(A10)`, `CELL_BILL_TO_3(A11)` 신규
- `CELL_FROM`: B13→B14, `CELL_DESTINATION`: B14→B15, `CELL_DEPARTS`: D15→D16
- `CELL_HS_CODE`: I11→I12, `CELL_PO_NO`: G15→G16, `CELL_PO_DATE`: I15→I16
- `ITEM_START_ROW`: 19→20 (Row 19 = Electric Actuator 카테고리 라벨)

#### 2. Shipping Mark 셀 변경
- **CI**: A31→A32, A32→A33, C33→C34
- **PL**: A33→A34, A34→A35, C35→C36

#### 3. `_fill_header()` 로직 변경
- 기존 `customer_name/address/country/tel/fax` + `delivery_address` 로직 제거
- `bill_to_1/2/3` 3줄 기록으로 교체
- Destination(To:)에 `bill_to_3` 사용 (국가명)

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `ci_generator.py` | 셀 상수 전면 변경, `_fill_header()` Bill to 로직 교체 |
| `pl_generator.py` | 셀 상수 전면 변경, `_fill_header()` Bill to 로직 교체 |
| `docs/TEMPLATE_MAPPINGS.md` | CI/PL 셀 매핑 업데이트 |

---

## 2026-03-12: PL 템플릿 G5에 Incoterms 배치

### 배경
PL 템플릿이 수정되어 F5에 `Incoterms:` 라벨이 추가됨. 기존 G5(L/C No), I5(L/C Date) 로직을 제거하고 G5에 Incoterms 값을 기록하도록 변경.

### 변경 내용

#### 1. 셀 상수 변경 (`pl_generator.py`)
- `CELL_LC_NO = 'G5'` → `CELL_INCOTERMS = 'G5'`
- `CELL_LC_DATE = 'I5'` 삭제

#### 2. `_fill_header()` 로직 변경
- L/C No/Date 기록 로직 제거
- G5에 Incoterms 기록 (SO_해외 JOIN 데이터)

#### 3. Shipping Mark 셀 위치 수정
- `A34→A33`, `A35→A34`, `C36→C35` (템플릿 Row 32 기준, 한 행 밀림 버그 수정)

---

## 2026-03-12: CI 템플릿 셀 위치 변경 (Incoterms/Payment Terms)

### 배경
CI 템플릿이 수정되어, 기존 G5(L/C No), I5(L/C Date) 자리에 Incoterms와 Payment Terms를 배치하고, 기존 G18의 Incoterms 로직을 제거.

### 변경 내용

#### 1. 셀 상수 변경 (`ci_generator.py`)
- `CELL_LC_NO = 'G5'` → `CELL_INCOTERMS = 'G5'`
- `CELL_LC_DATE = 'I5'` → `CELL_PAYMENT_TERMS = 'I5'`
- `CELL_INCOTERMS = 'G18'` 삭제 (G18 Incoterms 로직 제거)

#### 2. `_fill_header()` 로직 변경
- L/C No/Date 기록 로직 제거
- G5에 Incoterms 기록 (SO_해외 JOIN)
- I5에 Payment Terms 기록 (Customer_해외 JOIN)

---

## 2026-03-11: PL 생성기 기능 개선

### 배경
Packing List 생성 시 Shipping Mark 영역의 셀 위치가 실제 템플릿과 불일치하던 문제 수정, SO_해외 `AX Project number` → `Model code` 컬럼명 변경 대응, Weight 시트 기반 Net Weight 자동 조회 기능 추가.

### 변경 내용

#### 1. Shipping Mark 영역 수정 (`pl_generator.py`, `ci_generator.py`)
- **PL**: 셀 위치 수정 `A31→A34`, `A32→A35`, `C33→C36` (템플릿 Row 33 기준)
- **CI**: 셀 위치 유지 `A31`, `A32`, `C33` (템플릿 Row 30 기준, PL과 다름)
- 양쪽 모두 **bill_to_3 표시** (기존 customer_country 대체)
- `CELL_SHIPPING_MARK_COUNTRY` 제거 → `CELL_SHIPPING_MARK_BILLTO3` 신규

#### 2. Model code 별칭 추가 + 매핑 로직 개선 (`config.py`, `utils.py`, `document_service.py`)
- `COLUMN_ALIASES`에 `model_code: ('Model code', 'AX Project number', 'model_code')` 추가
- `load_so_export_data()` / `load_so_export_with_customer()` dtype에 `'Model code': str` 추가
- `_enrich_with_model_number()` 전면 개선:
  - 기존: 단일 SO_ID + Item name 매칭 (첫 SO만 매칭, 다중 SO 누락)
  - 변경: **SO_ID + Line item 복합키** 매칭 (DN에 여러 SO_ID가 섞여 있어도 전체 매칭)
  - Model code도 함께 매핑

#### 3. Weight 시트 기반 Net Weight 자동 조회 (`config.py`, `utils.py`, `document_service.py`)
- `WEIGHT_SHEET = 'Weight'` 상수 추가
- `load_weight_data()`, `build_weight_map()` 함수 추가 (ITEM→WEIGHT dict)
- `_enrich_with_weight()` 메서드 추가 — Model code로 Weight 시트 조회 → `Weight per unit` 컬럼 자동 추가
- `generate_pl()`에서 `_enrich_with_weight()` 호출
- Weight 시트 없거나 매칭 실패 시 graceful fallback

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `config.py` | `model_code` 별칭, `WEIGHT_SHEET` 상수 추가 |
| `utils.py` | SO 해외 dtype에 `Model code` 추가, `load_weight_data()`, `build_weight_map()` 추가 |
| `services/document_service.py` | `_enrich_with_model_number()` SO_ID+Line item 복합키 매칭으로 개선 + Model code 매핑, `_enrich_with_weight()` 신규, `generate_pl()` 수정 |
| `pl_generator.py` | Shipping Mark 상수 수정 (A34/A35/C36), bill_to_3 사용 |
| `ci_generator.py` | Shipping Mark에 bill_to_3 추가 (A31/A32/C33, 기존 위치 유지) |
| `docs/TEMPLATE_MAPPINGS.md` | PL Shipping Mark, Net Weight 데이터 소스 업데이트 |

---

## 2026-03-09: Order Book 스냅샷 기반 Variance 추적

### 배경
기존 `order_book.sql`은 매번 SO/DN raw 데이터에서 롤링 재계산. AX2009처럼 월별 마감(스냅샷) → Start를 고정하고, 소급 변경분을 Variance로 자동 감지하는 방식으로 전환.

**핵심 공식 변경**: `Ending = Start(롤링) + Input - Output` → `Ending = Start(스냅샷) + Input + Variance - Output`

### 신규 파일
| 파일 | 역할 |
|------|------|
| `close_period.py` | CLI 진입점 (마감/취소/현황 조회) |
| `po_generator/snapshot.py` | SnapshotEngine — 스냅샷 생성/취소/조회 |
| `sql/order_book_snapshot.sql` | 스냅샷 기반 Order Book SQL |
| `sql/order_book_snapshot_backlog.sql` | 스냅샷 기반 Backlog 뷰 |

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `po_generator/db_schema.py` | `create_snapshot_tables()` 함수 추가 (`ob_snapshot`, `ob_snapshot_meta` 테이블) |

### DB 테이블

**`ob_snapshot`** — 스냅샷 데이터 (PK: `snapshot_period, SO_ID, OS name, Expected delivery date`)
- 마감 Period의 전체 컬럼 고정값 저장 (Start, Input, Output, Variance, Ending × qty/amount)
- 컨텍스트 (customer_name, item_name, 구분, 등록Period, AX Period, Sector 등)

**`ob_snapshot_meta`** — 마감 메타 (PK: `period`)
- `is_active`: 활성 여부 (undo 시 0으로 변경)
- `closed_at`, `note`

### SnapshotEngine 로직

**`take_snapshot(period)`:**
1. period 형식 검증 (yyyy-MM)
2. 순차 마감 검증 (이전 period 마감 필수)
3. 롤링 order_book CTE 실행 → 해당 period 결과 추출
4. Variance 계산: `recalc_ending(현재 raw) - snap_ending(이전 스냅샷)` → 소급 변경분 감지
5. `ob_snapshot` + `ob_snapshot_meta` 저장

**`undo_snapshot(period)`:** 최신 활성 마감만 취소 가능. meta 비활성화 + snapshot 삭제.

### SQL 구조 (`order_book_snapshot.sql`)

**Open Period만 표시** — 마감된 Period는 `close_period.py --list`로 조회.

| 데이터 구간 | 처리 |
|------------|------|
| Open Period (스냅샷 이후) | Start=스냅샷 Ending, Variance=소급변경분, Ending=Start+Input+Var-Output |
| 스냅샷 없음 | 기존 롤링 계산 fallback (order_book.sql과 동일 결과) |

### 사용법

```bash
python close_period.py 2026-01                    # 1월 마감
python close_period.py 2026-02 --note "정기 마감"   # 노트 포함
python close_period.py --undo 2026-02              # 마감 취소 (최신만)
python close_period.py --list                      # 마감 현황
python close_period.py --status                    # 현재 상태
```

### `--list` 출력 보강

`list_snapshots()` 쿼리에 `total_start`, `total_input`, `total_output`, `total_variance` 합계 컬럼 추가. `print_list()` 함수에서 다음을 표시:

- **컬럼**: Period, 건수, Start, Input, Output, Variance, Ending, 마감일시
- **금액 포맷**: 백만 단위 `M` 접미사 (예: `1,597.2M`)
- **정합성 체크**: 전월 Ending != 당월 Start 시 차이 경고 표시
- **합계 행**: 활성 마감 기준 Input/Output/Variance 총합, 마지막 Ending
- **비고**: `--note`로 입력한 비고를 하단에 표시

### `create_po.bat` 메뉴 개편

- `[8]` → `DB Sync (Excel → SQLite)` (라벨 변경)
- `[9]` → `Order Book Close (월 마감)` 추가 (서브메뉴: 마감/취소/현황/상태)
- `[H]` → `발주 이력 조회` (기존 `[9]`에서 이동)

### 설계 결정사항
- 과거 Period: 스냅샷 고정값만 표시
- Variance: 총액만 (세부 구분 불필요)
- 마감 순서: 순차 강제 (1월→2월→3월)
- 기존 `order_book.sql`: 유지 (롤링 버전 병행)

---

## 2026-03-06: PO 테이블 PK에 `_row_seq` 추가 (부분 매입 대응)

- `db_schema.py`: PO_국내/PO_해외의 PK를 `(PO_ID, Line item)` → `(PO_ID, Line item, _row_seq)`로 변경
- `_row_seq`는 같은 `(PO_ID, Line item)` 그룹 내에서 Excel 행 순서대로 자동 부여 (1, 2, 3...)
- 부분 매입 시 같은 Line item이 분할되어도 PK 충돌 없이 정상 동기화
- `db_schema.py`: `migrate_pk_if_changed()` 추가 — 기존 DB의 PK가 설정과 다르면 자동 DROP → 재생성
- `db_sync.py`: 테이블 생성 전 PK 마이그레이션 체크 호출

---

## 2026-03-06: OC 품목명에 Model number 표시

- `oc_generator.py`: 품목명 출력 시 SO_해외의 Model number가 있으면 `"{Model number} {Item name}"` 형태로 표시
- CI와 동일한 로직 적용 (Model number 없으면 Item name만 출력)

---

## 2026-03-06: 내부 코드 최적화 (데이터 조회/서비스 캐시)

### 배경
데이터 조회 병목 분석 후, 출력 결과에 영향 없는 내부 구현 최적화 수행. 기능 회귀 없음 확인 (58 passed, 0 failed).

### 변경 내용

#### 1. `resolve_column()` 캐시 추가 (`utils.py`)
- `id(columns)` + `key` 기반 dict 캐시 도입
- `get_value()` 매 호출마다 반복되던 별칭 검색을 O(1) 조회로 전환
- 문서 1건 생성 시 수십~백 회 불필요한 선형 검색 제거

#### 2. `get_available_*_ids()` O(n²) → O(n) (`finder_service.py`)
- 4개 메서드(`get_available_po_ids`, `get_available_dn_ids`, `get_available_dn_export_ids`, `get_available_so_export_ids`)
- 기존: `unique()` 루프 안에서 `df[df[col] == id]` 반복 필터 → O(n²)
- 변경: `drop_duplicates(subset=..., keep='first').head(limit)` 단일 패스 → O(n)

#### 3. `find_so_for_advance()` 캐시 재사용 (`finder_service.py`)
- 기존: `load_so_for_advance()`가 Excel 파일을 독립적으로 다시 오픈 (PMT+SO 2시트 재로드)
- 변경: `FinderService`의 캐시된 `_pmt_df`와 신규 `_so_domestic_df` 활용, 중복 Excel I/O 제거
- `_load_so_domestic()` 프라이빗 메서드 추가 (SO_국내 lazy cache)

#### 4. `create_po.py` 다건 처리 서비스 공유
- `generate_po()`에 `service` 파라미터 추가 (기본값 `None` → 하위 호환)
- `main()`에서 여러 주문번호 처리 시 단일 `DocumentService` 인스턴스 공유
- DataFrame 재로드 방지

### 검토 후 제외된 항목
| 제안 | 제외 사유 |
|------|----------|
| `iterrows()` → `itertuples()` | 아이템 1~50건 수준이라 마이크로초 차이. 한글 컬럼명이 namedtuple 필드로 변환 실패 → 코드 복잡도만 증가 |
| xlwings COM 호출 추가 축소 | 이미 `batch_write_rows()` 등으로 96~97% 감소 완료. 남은 row insertion은 Excel API 제약으로 배치화 불가 |
| `get_value()` 배치 API | `resolve_column()` 캐시만으로 병목 해소. 별도 API는 blast radius가 큼 |

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `po_generator/utils.py` | `_RESOLVE_SENTINEL`, `_resolve_cache` 추가, `resolve_column()` 캐시 적용 |
| `po_generator/services/finder_service.py` | `_so_domestic_df` 캐시, `_load_so_domestic()` 추가, `get_available_*_ids()` 4개 single-pass 교체, `find_so_for_advance()` 캐시 재사용, `load_so_for_advance` import 제거 |
| `create_po.py` | `generate_po()` 시그니처에 `service` 파라미터 추가, `main()` 서비스 공유 |

---

## 2026-03-06: Commercial Invoice (CI) & Packing List (PL) 생성기 추가

### 배경
해외 출하 시 필요한 Commercial Invoice와 Packing List 생성 기능 추가. 둘 다 DN_해외 데이터를 사용하며, PI/FI와 유사한 셀 레이아웃.

### CI (Commercial Invoice)
PI와 동일한 셀 구조이나, 데이터 소스가 DN_해외이며 아래 차이점 있음:
- `ITEM_START_ROW = 19` (Row 18 = 카테고리 라벨 유지)
- `CELL_INCOTERMS = G18` (PI는 G17)
- A9 = Delivery Address, Shipping Mark (A31=Customer Name, C33=Customer PO)
- H열에 각 행 currency 표시, Total에 Qty 합계(E) + "EA"(F)
- **Model number 보강**: SO_해외에서 Item name 매칭으로 Model number 조회, 품목명 앞에 추가
- **Model number 오름차순 정렬**

### PL (Packing List)
CI와 동일한 헤더 구조이나, 아이템 열이 다름 (단가/금액 대신 Weight/CBM):
- F열: Net Weight (KG/PC), H열: Gross Weight (Kg), I열: CBM
- Shipping Mark: A31=Customer Name, A32=Customer Country, C33=Customer PO
- Model number 보강 및 정렬: CI와 동일

### 신규 파일
| 파일 | 역할 |
|------|------|
| `create_ci.py` | CI CLI 진입점 (`python create_ci.py DNO-2026-0001`) |
| `create_pl.py` | PL CLI 진입점 (`python create_pl.py DNO-2026-0001`) |
| `po_generator/ci_generator.py` | CI 생성기 (xlwings, PI 기반) |
| `po_generator/pl_generator.py` | PL 생성기 (xlwings, CI 기반 + Weight/CBM) |

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `config.py` | `CI_TEMPLATE_FILE`, `CI_OUTPUT_DIR`, `PL_TEMPLATE_FILE`, `PL_OUTPUT_DIR`, weight/cbm 컬럼 별칭 추가 |
| `utils.py` | `load_dn_export_data()`, `load_so_export_with_customer()` — Customer_해외 merge 시 `drop_duplicates()` 추가 (중복 행 방지) |
| `services/document_service.py` | `_enrich_with_model_number()`, `generate_ci()`, `generate_pl()` 메서드 추가 |
| `create_po.bat` | 메뉴에 [6] CI, [7] PL 추가 (기존 DB동기화 [8], 이력조회 [9]) |
| `docs/TEMPLATE_MAPPINGS.md` | CI/PL 셀 매핑 섹션 추가, PI 섹션 분리 |

### 데이터 흐름
```
DN_해외 → Customer_해외 (customer_code JOIN)
       → SO_해외 (SO_ID + Item name → Model number 보강)
```

### Customer_해외 중복 행 수정
`load_dn_export_data()`와 `load_so_export_with_customer()`에서 Customer_해외 merge 시 `drop_duplicates(subset='C-code by 해외', keep='first')` 추가. Customer_해외에 동일 고객코드 중복 행이 있을 때 DN 행이 배수로 늘어나는 버그 수정.

---

## 2026-03-05: Order Confirmation (OC) 생성기 추가

### 배경
해외 고객에게 주문 확인서(Order Confirmation)를 발행하는 기능 추가. Final Invoice와 동일한 레이아웃이지만, H열에 **Dispatch date** 컬럼이 추가된 형태. Dispatch date는 SO_해외의 `EXW NOAH` 컬럼 값을 사용.

### 신규 파일
| 파일 | 역할 |
|------|------|
| `create_oc.py` | CLI 진입점 (`python create_oc.py SOO-2026-0001`) |
| `po_generator/oc_generator.py` | OC 생성기 (xlwings, FI 기반 + Dispatch date) |

### 수정 파일
| 파일 | 변경 내용 |
|------|----------|
| `config.py` | `OC_TEMPLATE_FILE`, `OC_OUTPUT_DIR` 추가, `exw_noah` 컬럼 별칭 추가 |
| `utils.py` | `load_so_export_with_customer()` 신규 (SO_해외+Customer_해외 JOIN) |
| `services/finder_service.py` | `find_so_export_with_customer()` 메서드 추가 |
| `services/document_service.py` | `generate_oc(so_id)` 메서드 추가 |
| `create_po.bat` | 메뉴에 [5] Order Confirmation 추가 (기존 DB동기화 [5]→[6]으로 이동) |

### OC vs FI 차이점
| 항목 | FI | OC |
|------|----|----|
| 제목 | Invoice | Confirmation of Order |
| H열 (Row 17~) | (없음) | Dispatch date = SO_해외.EXW NOAH |
| 나머지 | 동일 | 동일 |

### 데이터 흐름
`SO_해외` → `Customer_해외` (고객코드 JOIN, Bill to/Payment terms 포함)

---

## 2026-03-04: FI 새 템플릿 대응 업데이트

### 배경
`templates/final_invoice.xlsx` 양식 전면 개편으로 셀 매핑, 아이템 열 구조, 신규 필드 대응 필요.

### 변경 파일
| 파일 | 변경 내용 |
|------|----------|
| `config.py` | `dispatch_date` alias에 `'선적일'` 추가, `delivery_address` alias에 `'Delivery address'` 추가 |
| `utils.py` | `load_dn_export_data()` — SO_해외 merge 컬럼에 `Currency`, `Incoterms` 추가, Customer_해외 JOIN 키를 `resolve_column()`으로 동적 탐지, DN-SO 컬럼 충돌 시 SO 우선 (overlap drop) |
| `fi_generator.py` | 셀 매핑 전면 교체, `_fill_header()` 재작성, `_fill_items_batch()` Currency 열 추가, `_update_total_row()` F열 "EA" 추가 |
| `docs/TEMPLATE_MAPPINGS.md` | FI 섹션 새 템플릿 구조로 업데이트 |

### 셀 매핑 변경 요약
| 필드 | OLD → NEW | 데이터 소스 |
|------|-----------|------------|
| Customer PO | G10 → C7 | SO_해외.Customer PO |
| Invoice No | G4 → H7 | DN_해외.DN_ID |
| PO Date | I10 → C8 | SO_해외.PO receipt date |
| Invoice Date | I4 → H8 | DN_해외.선적일 |
| Payment Terms | G8 → H9 | Customer_해외.Payment terms |
| Delivery Terms | (신규) H10 | SO_해외.Incoterms |
| Customer Address | A9~11 → A12~14 | Customer_해외.Bill to 1/2/3 |
| Delivery Address | (신규) G12 | DN_해외.Delivery address |
| Due Date | I8 → (삭제) | — |

### 아이템 열 변경
| 항목 | OLD → NEW |
|------|-----------|
| ITEM_START_ROW | 14 → 17 |
| Unit Price 열 | G → F |
| Currency 열 | (신규) G |

### 데이터 로드 개선 (`load_dn_export_data`)
- **Customer_해외 JOIN 키**: 하드코딩(`'Business registration number'`) → `resolve_column()`으로 동적 탐지
- **DN-SO 컬럼 충돌**: SO_해외에서 가져올 컬럼이 DN_해외에도 존재하면 merge 전 DN쪽 drop (SO 우선)
- **alias 대소문자**: `'Delivery address'`(소문자 a) 추가 — DN_해외 실제 컬럼명과 매칭

---

## 2026-03-03: Excel → SQLite DB 동기화 구현

### 배경
NOAH_SO_PO_DN.xlsx가 사실상 ERP 역할을 하고 있으나, Excel 형식 특성상 데이터 유실/변형에 취약. 수동 입력 시트(SO, PO, DN, PMT)를 SQLite DB에 upsert 방식으로 업로드하여 데이터를 안전하게 백업하고 관리.

### 신규 파일
| 파일 | 역할 |
|------|------|
| `sync_db.py` | CLI 진입점 (--dry-run, --sheets, --info, -v) |
| `po_generator/db_schema.py` | 테이블/PK 정의, DDL 생성, 스키마 관리 |
| `po_generator/db_sync.py` | SyncEngine — upsert 동기화 엔진 |

### 수정 파일
| 파일 | 변경 |
|------|------|
| `po_generator/config.py` | `DB_FILE` 상수 1줄 추가 |

### 테이블 설계 (7개)

| 테이블명 | 소스 시트 | PK | 행 수 |
|----------|----------|-----|------|
| `so_domestic` | SO_국내 | `(SO_ID, Customer PO, Line item)` | 590 |
| `so_export` | SO_해외 | `(SO_ID, Customer PO, Line item)` | 233 |
| `po_domestic` | PO_국내 | `(SO_ID, Customer PO, Line item)` | 589 |
| `po_export` | PO_해외 | `(SO_ID, Customer PO, Line item, _row_seq)` | 235 |
| `dn_domestic` | DN_국내 | `(DN_ID, Line item)` | 283 |
| `dn_export` | DN_해외 | `(DN_ID, SO_ID, Line item)` | 88 |
| `pmt_domestic` | PMT_국내 | `(선수금_ID)` | 33 |

### 사용법
```bash
python sync_db.py                           # 전체 동기화
python sync_db.py -v                        # 상세 로그
python sync_db.py --sheets SO_국내 PO_국내  # 특정 시트만
python sync_db.py --dry-run                 # 시뮬레이션
python sync_db.py --info                    # DB 현황 조회
```

### 핵심 설계
- **DB**: SQLite (Python 내장, 서버 불필요). 위치: `DATA_DIR / "noah_data.db"`
- **Upsert**: PK 기준 INSERT or UPDATE — 재실행 시 기존 데이터 업데이트
- **PO_해외 `_row_seq`**: 같은 SO Line item에 사양 변형 시 자동 순번 부여
- **스키마 진화**: `ensure_columns_exist()`로 Excel 컬럼 추가 시 자동 대응
- **메타 테이블**: `_sync_meta`에 테이블별 마지막 동기화 시간/행 수 기록

---

## 2026-02-28: Power Query 개선 및 문서 정비

### Power Query 수정
- PO 원가 계산: `Table.Distinct` → `Table.Group` 변경 (사양 분리 시 중복 합산 방지)
- DN 분할납품: `Table.Distinct` → `Table.Group` 변경 (분할 출고 금액 정확 집계)
- SO_통합 출고 상태: 3단계 → 4단계 세분화 (미출고/부분 출고/출고 완료/선적 완료)
- PO_AX대사 쿼리 추가: Period + AX PO별 GRN 금액 집계

### 문서 정비
- `DATA_STRUCTURE_DESIGN.md`: ERP 매핑 섹션 추가 (테이블 관계, 조인, 상태 관리)
- `CLAUDE.md`: 아키텍처, 커맨드, 키 패턴 섹션 확장
- `POWER_QUERY.md`: Key Files에 추가

---

## 2026-02-15: Final Invoice 및 Power Query 문서화

### Final Invoice (FI) 생성기 추가
- `create_fi.py` CLI 진입점 추가 (DN_해외 기반)
- `fi_generator.py` 구현 (xlwings) — Bill-to, Payment Terms, Due Date 등
- `create_po.bat` 메뉴에 [4] Final Invoice 추가
- `OPERATION_GUIDE.md` 운용 가이드 추가
- `config.py`: `FI_TEMPLATE_FILE`, `FI_OUTPUT_DIR` 추가

### Power Query 문서화
- `docs/POWER_QUERY.md` 신규 작성 (SO_통합, PO_현황, Order_Book 쿼리)
- Order_Book 파이프라인 다이어그램 및 단계별 데이터 흐름 예시

### 데이터 구조
- SO 컬럼: `Customer PO`, `Expected delivery date` 추가
- Order_Book: 분할 납품 처리 (DN 월별 조인)

---

## 2026-02-08: TS/PI 기능 및 테스트 추가

### 거래명세표/PI 기능 확장
- TS/PI 관련 기능 정리 및 테스트 추가
- `.gitignore` 업데이트 (generated_ts, po_history, Claude 임시 파일)
- README 갱신 (PO, TS, PI 문서 유형 반영)

---

## 2026-01-31: 거래명세표 기능 개선

### 출고일 기준 날짜 표시
- **변경**: 거래명세표 날짜를 오늘 날짜 → **출고일** 기준으로 변경
- `config.py`: `dispatch_date` 별칭 추가 (`'출고일', 'Dispatch Date', 'dispatch_date', '출하일'`)
- `ts_generator.py`: 헤더(B2)와 아이템(A열) 날짜를 출고일로 표시
  - 출고일이 없으면 오늘 날짜를 폴백으로 사용
  - 파라미터명 `today` → `dispatch_date`로 변경

### 월합 거래명세표 기능 추가
고객이 월합으로 거래명세표를 요청할 때, 여러 DN을 한 장으로 합쳐서 생성

**사용법:**
```bash
python create_ts.py DND-2026-0001 DND-2026-0002 DND-2026-0003 --merge
python create_ts.py --interactive --merge
```

**변경 파일:**
| 파일 | 변경 내용 |
|------|----------|
| `config.py` | `dispatch_date` 컬럼 별칭 추가 |
| `ts_generator.py` | 출고일 기준 날짜 표시 |
| `create_ts.py` | `--merge`, `--interactive` 옵션 추가, `generate_merged_ts()` 함수 |
| `create_po.bat` | 거래명세표 메뉴에 [1] 단건 / [2] 월합 선택 추가 |

**월합 거래명세표 동작:**
- 여러 DN의 아이템을 하나의 DataFrame으로 합침
- 출고일: 입력된 DN 중 **가장 최근 출고일** 사용
- 고객명이 다르면 경고 표시 (첫 번째 고객 기준)
- 파일명: `월합_고객명_날짜.xlsx`

---

## 2026-01-21: 코드 리팩토링 (5 Phases)

Code Reflection 결과를 바탕으로 코드 품질 개선 작업 수행.

### Phase 1: excel_helpers.py 인프라 추가
- [x] `XlConstants` 클래스 추가 ✓
  - Excel COM 매직 넘버를 명명된 상수로 정의
  - `xlShiftUp`, `xlShiftDown`, `xlEdgeTop`, `xlEdgeBottom`, `xlContinuous`, `xlThin` 등
  - 코드 가독성 향상, 하드코딩된 -4162, -4121 등 제거
- [x] `xlwings_app_context` 컨텍스트 매니저 추가 ✓
  - xlwings App 생명주기 안전 관리
  - 오류 발생 시에도 Excel 프로세스 자동 정리
  - 리소스 누수 방지
- [x] `prepare_template()`, `cleanup_temp_file()` 헬퍼 추가 ✓
  - 중복되는 템플릿 복사 로직 통합
  - 임시 파일 안전 삭제

### Phase 2: cli_common.py 보안 수정
- [x] Path Traversal 취약점 수정 ✓
  - **문제**: 문자열 포함 검사(`in`)로 경로 탈출 가능
    - `/home/user/documents`가 `/home/user/doc_files/test.xlsx`에 포함
  - **수정**: `relative_to()` 사용으로 정확한 경로 검증
  ```python
  # Before (취약)
  if str(output_dir.resolve()) not in str(output_file.resolve()):

  # After (안전)
  resolved_file.relative_to(resolved_dir)  # ValueError 발생 시 거부
  ```

### Phase 3: Generator 리팩토링
- [x] `excel_generator.py` 리팩토링 ✓ - `xlwings_app_context`, `XlConstants`, 타입 변환 경고 로깅
- [x] `ts_generator.py` 리팩토링 ✓ - 동일 패턴 적용
- [x] `pi_generator.py` 리팩토링 ✓ - 동일 패턴 적용

### Phase 4: 테스트 커버리지 확대
- [x] `tests/test_excel_helpers.py` 신규 생성 ✓ (16개 테스트)
- [x] `tests/test_cli_common.py` 신규 생성 ✓ (11개 테스트)
- [x] `tests/test_config.py` 신규 생성 ✓ (22개 테스트)
- **테스트 결과**: 47 passed, 2 skipped

### Phase 5: utils.py 중복 함수 통합
- [x] `_find_data_by_id()` 공통 헬퍼 추가 ✓
  - ID로 데이터 검색하는 공통 로직 통합
- [x] 4개 find 함수를 wrapper로 변경 ✓
  | 함수 | 변경 전 | 변경 후 |
  |------|--------|--------|
  | `find_order_data()` | 34줄 | 1줄 (wrapper) |
  | `find_dn_data()` | 33줄 | 1줄 (wrapper) |
  | `find_pmt_data()` | 28줄 | 1줄 (wrapper) |
  | `find_so_export_data()` | 33줄 | 1줄 (wrapper) |
- **효과**: ~90줄 중복 제거, 버그 수정 시 단일 지점 수정
- **테스트 결과**: 160 passed, 2 skipped

**변경 파일 요약:**
| 파일 | 변경 내용 |
|------|----------|
| `excel_helpers.py` | +110 lines (XlConstants, context manager, helpers) |
| `cli_common.py` | 보안 버그 수정 |
| `excel_generator.py` | Context manager 적용, 상수화 |
| `ts_generator.py` | Context manager 적용, 상수화 |
| `pi_generator.py` | Context manager 적용, 상수화 |
| `test_excel_helpers.py` | +160 lines (신규) |
| `test_cli_common.py` | +90 lines (신규) |
| `test_config.py` | +140 lines (신규) |

---

## 2026-01-21: xlwings 성능 최적화

### 배치 연산 헬퍼 함수 추가
`excel_helpers.py`에 새 함수:
- `batch_write_rows`: 2D 리스트를 한 번에 쓰기
- `batch_read_column`: 열의 값을 한 번에 읽기
- `batch_read_range`: 범위의 값을 한 번에 읽기
- `delete_rows_range`: 연속 행을 한 번에 삭제
- `find_text_in_column_batch`: 배치 읽기로 텍스트 찾기

### Generator별 최적화
- `ts_generator.py`: `_fill_items_batch` (N*8회→1회), `_find_label_row` (36회→1회), `_find_ts_subtotal_row` (15회→1회)
- `pi_generator.py`: `_fill_items_batch` (N*4회→4회), `_find_total_row` (20회→1회), `_fill_shipping_mark` (80회→2회)
- `excel_generator.py`: `_fill_items_batch_po`, `_create_description_sheet` (30*N회→1회), `_find_totals_row` (20회→1회)

### 예상 성능 개선 (50개 아이템 기준)
| 파일 | COM 호출 (전) | COM 호출 (후) | 감소율 |
|------|--------------|--------------|--------|
| ts_generator.py | ~500회 | ~20회 | 96% |
| pi_generator.py | ~350회 | ~15회 | 96% |
| excel_generator.py | ~1,500회 | ~50회 | 97% |

---

## 2026-01-21: 버그 수정 - xlwings 범위 formula 읽기

**증상**: 거래명세표 생성 시 템플릿의 예시 아이템이 삭제되지 않고 그대로 남아있음

**원인**: `_find_ts_subtotal_row` 함수의 배치 읽기 최적화에서 xlwings의 `.formula` 속성 반환 형식을 잘못 처리

**상세 분석:**
```python
# xlwings 범위 읽기 반환 형식 차이
ws.range('E13').value           # 단일 셀 → float: 8.0
ws.range('E13:E17').value       # 범위 → list: [8.0, 8.0, 16.0, None, None]

ws.range('E15').formula         # 단일 셀 → str: '=SUM(E13:E14)'
ws.range('E13:E17').formula     # 범위 → tuple of tuples: (('8',), ('8',), ('=SUM(E13:E14)',), ('',), ('',))
```

- `.value`: 단일 열 범위 → **1D list** 반환
- `.formula`: 단일 열 범위 → **tuple of tuples** 반환 (2D 형태)

**버그 코드:**
```python
formulas = ws.range(f'E{start_row}:E{end_row}').formula
if not isinstance(formulas, list):
    formulas = [formulas]  # tuple of tuples가 통째로 리스트에 들어감
for idx, formula in enumerate(formulas):
    if formula and '=SUM' in str(formula):  # 전체 tuple을 문자열로 변환
        return start_row + idx  # 항상 index 0 반환
```

결과: `subtotal_row = 13` (실제로는 15) → `template_item_count = 0` → 행 삭제 안됨

**수정 코드** (`ts_generator.py:186-197`):
```python
formulas = ws.range(f'E{start_row}:E{end_row}').formula

# xlwings 범위 읽기는 tuple of tuples 반환: (('val1',), ('val2',), ...)
# 단일 셀은 문자열 반환
if isinstance(formulas, (list, tuple)) and formulas and isinstance(formulas[0], (list, tuple)):
    # 2D → 1D 평탄화 (각 행의 첫 번째 값만 추출)
    formulas = [f[0] if f else '' for f in formulas]
elif not isinstance(formulas, (list, tuple)):
    formulas = [formulas]
```

**영향 범위:**
| 모듈 | 함수 | 사용 속성 | 상태 |
|------|------|----------|------|
| `ts_generator.py` | `_find_ts_subtotal_row` | `.formula` (범위) | **수정됨** |
| `pi_generator.py` | `_find_total_row` | `.value` (범위) | 문제 없음 |
| `excel_generator.py` | `_find_totals_row` | `.value` (범위) | 문제 없음 |
| `excel_helpers.py` | `batch_read_column` | `.value` (범위) | 문제 없음 |

**교훈:**
- xlwings에서 `.value`와 `.formula`는 범위 읽기 시 반환 형식이 다름
- `.value`: 1D list (단일 열)
- `.formula`: 2D tuple of tuples (항상 2D)
- 배치 최적화 시 반환 형식을 실제 테스트로 확인 필요

---

## 2026-01-21: 서비스 레이어 추가

- [x] `excel_helpers.py` 생성 ✓ - `find_item_start_row` 통합, 헤더 라벨 프리셋
- [x] `services/` 디렉토리 생성 ✓ - DocumentService, FinderService, DocumentResult
- [x] CLI 리팩토링 ✓ - 서비스 레이어 사용, 사용자 상호작용은 CLI 유지
- [x] 행 삭제 주석 수정 ✓ - "같은 위치에서 반복 삭제 - xlUp으로 아래 행이 올라옴"
- [x] 통합 테스트 추가 ✓ - 11개 테스트 케이스

---

## 2026-01-21: 버그 수정 - Description 시트 A열 레이블

- **원인**: 템플릿의 고정 레이블에만 의존, 동적 필드(`get_spec_option_fields`)와 불일치
- **수정**: A열에 레이블 명시적 쓰기 (`['Line No', 'Qty'] + all_fields`)
- `_apply_description_borders` 함수 추가 (테두리 적용)
- 국내/해외 모두 동적 필드 사용 (PO_국내: 47개, PO_해외: 45개)

---

## 2026-01-21: 버그 수정 - PI 행 삽입 시 테두리

- **증상**: 템플릿 마지막 행(8행) 테두리가 중간에 남음 + Total 위 선 누락
- **원인**: 행 삽입 케이스에서 `_restore_item_borders` 미호출
- **수정**: 삽입 전 템플릿 원래 마지막 행 테두리 제거 (`XlConstants.xlNone`)
- **수정**: 삽입 후 `_restore_item_borders` 호출로 새 마지막 행 테두리 추가
- `excel_helpers.py`에 `XlConstants.xlNone = -4142` 상수 추가

---

## 2026-01-20: openpyxl → xlwings 전환

- openpyxl → xlwings 전환 (이미지/서식 보존)
- `get_safe_value` → `get_value` 표준 API로 통일

---

## 2026-01-19: 버그 수정 - PO Delivery Address

- Delivery Address 값이 안 나오던 문제 해결
  - `config.py`: `delivery_address` 컬럼 별칭 추가
  - `utils.py`: SO→PO 병합 시 `'납품 주소'` 컬럼 누락 수정
  - `excel_generator.py`: 하드코딩 키워드 검색 → `get_value()` 사용
- 파일 열 때 Description 시트가 먼저 보이던 문제 해결
  - `excel_generator.py`: `wb.active = ws_po` 추가

---

## 2026-01-19: 버그 수정 - 거래명세표/PI 템플릿 예시 아이템 삭제

- 실제 아이템 < 템플릿 예시 시 초과 행 삭제 안되던 문제 해결
- 행 삭제 후 테두리 복원 (`_restore_ts_item_borders`, `_restore_item_borders`)
- PI: Shipping Mark 영역 검색 범위 수정 (40→20 시작)

---

## 완료된 TODO 항목

### OneDrive 공유 폴더 연동
- [x] 회사 랩탑에서 OneDrive 공유 폴더 경로 확인 ✓
- [x] `config.py`에서 경로 설정 외부화 → `user_settings.py` ✓
- [x] 파일 구조 변경 ✓
- [x] po_history 월별 폴더 방식으로 변경 ✓

### 템플릿 기반 문서 생성
- [x] PO (Purchase Order) - openpyxl 기반 ✓
- [x] 거래명세표 (Transaction Statement) - xlwings 기반 ✓
- [x] PI (Proforma Invoice) - xlwings 기반 ✓
- [x] FI (Final Invoice) - xlwings 기반 ✓
- [x] OC (Order Confirmation) - xlwings 기반 ✓
- [x] CI (Commercial Invoice) - xlwings 기반 ✓
- [x] PL (Packing List) - xlwings 기반 ✓

---

## 라이브러리 선택 기준

| 용도 | 라이브러리 | 이유 |
|------|-----------|------|
| PO 생성 | openpyxl | 이미지 불필요, 빠른 생성 |
| TS/PI/FI/OC 생성 | xlwings | 로고/도장 이미지, 복잡한 서식 완벽 보존 |
| 이력 조회/테스트 검증 | openpyxl | COM 인터페이스 없이 안정적인 읽기 |

### 템플릿 동작 방식
- 템플릿 파일의 **데이터는 무시됨** - 코드에서 초기화 후 새로 채움
- 템플릿의 **구조/서식만 유지됨**: 레이아웃, 서식, 이미지, 수식
- 새 템플릿 추가 시: `templates/` 폴더에 양식 파일 추가 후 코드에서 셀 매핑 정의
