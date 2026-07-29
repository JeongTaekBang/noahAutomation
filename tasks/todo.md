# Current Tasks — 납기현황 메일 발송 [2026-07-29]

납기현황을 만든 뒤 거래명세표(TS)와 같은 방식으로 고객에게 메일 발송한다.
수신자는 `Customer_국내`(사업자번호 조인), 본문에 표를 넣고 xlsx를 첨부한다.

**결정사항** (사용자 확인)
- 본문: **HTML 표 + xlsx 첨부** — 기존 회신이 표를 본문에 붙이는 방식이었다. 첨부만 보내면 후퇴
- 첨부 범위: 생성된 **2시트 그대로** (`납기현황` + `상세`)
- 조회가 이미 사업자번호 기준이라 TS와 달리 "고객 섞임" 위험이 없다 — merge 차단 로직 불필요

## 설계

**`mailer.py`를 문서 중립으로.** 지금은 TS 전용 상수가 함수 안에 박혀 있다.
TS 동작은 한 톨도 바뀌지 않게 하고, 고정된 부분만 인자로 뺀다.
- `create_document_mail(...)` 신설 — 제목/본문 템플릿, 첨부형식, 초안 파일명 접두사, HTML 본문을 받는다
- `create_ts_mail(...)`은 TS 상수를 넘기는 얇은 래퍼로 (기존 호출부·테스트 그대로)
- `render_template(..., extra=None)` — 문서별 치환자 추가
- `find_recipient(..., fixed_cc=None)` — 기본값은 `TS_MAIL_CC` (현행 유지)
- `build_eml(..., prefix=...)` / Outlook COM에 `HTMLBody` 지원

**`po_generator/mail_cli.py` 신설.** `MailMode`·`MailOptions`·`resolve_mail_mode`·`confirm`은
CLI 두 개가 똑같이 필요한 배선인데 지금 `create_ts.py` 안에 있다. 복사하면 반드시 갈라진다.
`create_ts.py`는 재노출만 해서 `from create_ts import MailMode` 하는 기존 테스트를 깨지 않는다.

## 구현
- [x] 1. `po_generator/mail_cli.py` — MailMode / MailOptions / resolve_mail_mode / confirm / add_mail_arguments
- [x] 2. `create_ts.py` — 위 심볼을 mail_cli에서 import + 재노출 (동작 불변)
- [x] 3. `mailer.py` — `create_document_mail`, `render_template(extra)`, `find_recipient(fixed_cc)`,
      `build_eml(prefix, body_html)`, Outlook `HTMLBody` 경로, `body_to_html` 공개
- [x] 4. `config.py` — `DS_MAIL_CC` / `DS_MAIL_ATTACH_FORMAT`(기본 xlsx) / `DS_MAIL_SUBJECT` / `DS_MAIL_BODY`
- [x] 5. `delivery_status.py` — `--mail`/`--send`/`--no-mail`, 본문 표(HTML+평문) 생성, 발송 흐름
- [x] 6. 테스트 — `tests/test_delivery_status.py` 15건 추가 (72개)
- [x] 7. 문서 — `CLAUDE.md` / `docs/CHANGELOG.md` / `create_po.bat` 안내 문구
      (`user_settings.example.py`는 이 레포에 없음 — 설정 설명은 `config.py` 주석에 둔다)

## 원칙
- **메일 실패가 문서 생성 성공을 뒤엎지 않는다** (TS와 동일)
- 비대화형 실행은 자동으로 묻지 않는다 — 배치가 프롬프트에서 멈추면 안 된다
- 수신자를 보여주고 y/N 확인이 기본. 고객에게 나가는 것은 항상 사람이 한 번 본다
- TS 경로는 회귀 0 — `tests/test_mailer.py` 100개가 그대로 통과해야 한다

## Review

계획대로 갔다. 리팩터를 먼저 한 판단이 옳았는데, 근거가 구현 중에 드러났다:
`mailer.py`에서 실제로 문서마다 달라야 했던 건 **6가지**(제목·본문 템플릿, 첨부형식, 고정 CC,
HTML 본문, 초안 파일명, 로그 라벨)였다. 복사했다면 642줄짜리 모듈이 두 벌이 됐을 것이다.

구현 중 판단한 것 셋:

1. ~~**본문 표는 4개 컬럼만.**~~ → **피드백으로 Sales 금액 + Requested delivery date 추가.**
   처음엔 "좁은 화면" 이유로 금액을 뺐는데, 초안을 열어본 사용자가 금액도 요청납기도 필요하다고 했다.
   요청납기를 공장출고일 바로 왼쪽에 둬서 대조되게 했다 — 엔이에스 14건 중 2건이 요청보다 늦다.
   PO receipt date만 첨부 전용으로 남긴다 (고객이 이미 아는 날짜).
   요청납기가 라인마다 갈리면 처음엔 가장 이른 날로 접었는데, 사용자가 "다르면 풀어달라"고 했다.
   맞는 지적이다 — 접으면 나머지 약속이 사라진다. 피엠에스 `SOD-2026-0344`는 요청 10/06인데
   출고 11/15인 라인이 10/06 행에 묻혀 있었다. **분리 기준을 요청납기+공장출고일 둘 다로** 바꿨고
   실측 161행 → 164행 (3개 주문만 추가 분리).
2. **HTML과 평문 둘 다 데이터를 담는다.** 평문 대체본을 "첨부 참조"로 때우면 HTML을 막아둔
   클라이언트가 받는 메일엔 답이 없다. `_pad`(표시 폭 기준)를 재사용해 문자표를 그렸다.
3. **표는 이스케이프 뒤에 넣는다.** 본문 값은 전부 이스케이프해야 하는데(`S&T중공업`),
   표를 먼저 끼우면 `<table>`이 통째로 글자가 된다. 토큰(`%%DELIVERY_STATUS_TABLE%%`)을 넣고
   이스케이프한 다음 되돌린다. 순서가 뒤집히면 조용히 깨지는 종류라 회귀 테스트를 붙였다.

Outlook COM 쪽에서 하나 배웠다: `HTMLBody`와 `Body`를 **둘 다 대입하면 나중 것이 앞을 지운다.**
표가 있는 문서는 HTML만 설정한다 (이 PC는 .eml 경로지만 classic Outlook 환경 대비).

**초안을 실제로 열어보고 고친 것 (2026-07-29, 2·3차)**
- **서명 위치 — 두 번 틀리고 규칙을 좁혔다.**
  - 1차: `<br>`로 이은 인라인 텍스트 → 서명이 표(첫 블록) **앞**에
  - 2차: 문단을 `<p>` 블록으로 → 서명이 첫 `<p>` **뒤**에. `<div>` 래퍼는 무력
  - → 규칙은 "**Outlook은 본문의 첫 블록 요소 뒤에 서명을 넣는다**".
    3차: 본문 전체를 `<table><tr><td>` 한 칸에 담아 최상위 블록을 하나로. 삽입 지점 = 본문 끝
  - **3차로 해결됐다** (사용자 초안 확인). 서명이 표 아래, 본문 끝에 붙는다
- 제목 태그 `[납기현황 회신]` → `[Delivery Schedule]` (사용자 선택)
- 표에 `Requested delivery date` 추가, `NOAH 공장 출고일` → **`NOAH 공장 출고 예정일`**
- 요청납기 초과 행은 빨강 + 굵게 (평문은 `*` + 각주)
- 납기가 갈리면 행 분리 — 기준을 요청납기·출고예정일 **둘 다**로
- 본문 문구를 거래명세표 톤으로 건조하게 (`요청하신…회신드립니다` → `…송부하오니 참고 바랍니다`),
  끝에 "본 메일은 자동 발송된 메일입니다."
- 표에 Sales 금액 추가

**남은 위험**: `Remarks`가 고객에게 그대로 나간다. 데이터에 내부 생산 메모가 섞여 있다
(`… 2026-02-09-CH-05 미출고 Worm gear 사용하여 제작`). 초안을 눈으로 훑는 것을 전제로 한다 —
기본이 즉시 발송이 아니라 초안인 이유가 여기에도 있다.

**최종 상태 (2026-07-29)**: 사용자 실사용 확인 완료 — 초안 열기·서명 위치·표·첨부 전부 정상.
커밋 `a25ef3a`(리팩터) + `7945045`(기능), `origin/master` 반영. `pytest tests/` 500 passed, 2 skipped.

검증:
- 실제 `.eml` 생성 — 수신자 `nes@neskorea.co.kr` + 참조 5명이 **기존 회신 메일과 동일하게** 잡혔다.
  `X-Unsent: 1`, xlsx 첨부, text/plain·text/html 양쪽에 14행 표, 토큰 잔존 없음
- `pytest tests/` **476 passed, 2 skipped** — TS 경로 회귀 0
  (`test_mailer.py` 3건은 `MailOptions` 이사로 monkeypatch 대상만 갱신, 동작 검증 내용은 그대로)

남은 것: 실제 발송 1건. `.eml` 초안까지는 확인했지만 [보내기]를 눌러본 적은 없다.

---

# 완료: 거래처 납기현황 회신 [2026-07-29]

업체가 "언제 나오냐"고 물을 때마다 SO_국내를 수기로 피벗해 회신하던 작업을 CLI 한 줄로 만든다.
조회 기준은 **사업자등록번호**(Business registration number), 대상은 **국내(SO_국내)**.
출력은 첨부 메일 표와 같은 형식 + `Sales 금액`·`PO receipt date` 추가.

## 핵심 설계 결정

**출고 여부는 시트의 `Status` 값이 아니라 DN 출고수량 누계로 그 자리에서 계산한다.**

`SO_국내.Status`는 수기 입력이 아니라 파워쿼리 결과를 끌어온 **캐시값**이다:
```
Status = IFERROR(XLOOKUP(SO_ID&Line item, SO_통합[SO_ID]&SO_통합[Line item], SO_통합[출고완료]), "")

SO_통합[출고완료] =  if [출고수량] = null      then "미출고"
                    else if 발주수량-[출고수량] > 0 then "부분 출고"
                    else if [출고일] = null      then "공장 출고"
                    else "출고 완료"
```
`SO_통합[출고완료]` 자체가 **DN 출고수량 기반 판정**이므로, CLI가 `DN_국내`에서 같은 식을 직접 계산하면
결과는 동일하면서 **파워쿼리 새로고침 여부에 의존하지 않는다**. `DN_국내`는 직접 입력 시트라 캐시가 없다.
(`pandas.read_excel`은 수식 셀의 캐시값을 읽으므로, 새로고침 전 파일을 읽으면 옛 판정이 그대로 나온다.)

`dashboard.py: load_so()`도 같은 규칙(2026-06-24 수량 기반 전환). 시트 `Status`는 `Cancelled`/`Hold` 제외에만 사용.

실측 근거 (2026-07-29, 새로고침 전 파일 기준 SO_국내 1,992행):
- 캐시 `Status='미출고'`인데 DN상 출고 완료 — **20라인** (삼신 SOD-2026-0651 8, 키밸브 SOD-2026-0210 8, 오토밸브·굿이엔지 4)
  → 캐시값으로 거르면 **이미 납품한 건을 "미출고"라고 고객에게 회신**한다.
- 캐시 `Status='출고 완료'`인데 실제 부분출고 — 1라인 (삼성정공 SOD-2026-0232 L2, 54 주문 / 43 출고 / 11 잔량)
  → 남은 11개가 회신에서 누락된다.

**부수 효과**: 캐시 `Status`와 계산 결과가 어긋나면 실행 시 "파워쿼리 새로고침 필요" 알림을 띄운다.
회신 문서는 어차피 맞게 나가지만, 대시보드·피벗 등 다른 산출물도 같이 낡았다는 신호이므로 알려준다.

**금액은 `Sales Unit Price × 미출고수량`으로 낸다.**
SO_국내 1,992행 전부 `Sales amount = Sales Unit Price × Item qty`가 성립하므로, 부분출고 잔량 금액도
안분 없이 정확히 계산된다.

**데이터 소스는 Excel 직접 읽기** (SQLite 아님).
고객 회신용이라 최신성이 최우선이고, `sync_db.py` 선행 실행을 전제로 두지 않는다. `create_*.py`와 동일한 경로.

**집계 단위는 SO_ID.**
첨부 메일 표와 대조 검증 완료 — 엔이에스(615-81-88675) 기준 14행이 Customer PO / Remarks / 수량 /
공장 출고일까지 메일과 일치. (`2026-04-20-CH-02`는 Remarks가 2종이지만 같은 SO_ID라 합계 10으로 한 행)

## 구현
- [x] 1. `po_generator/config.py` — `DS_OUTPUT_DIR` (`generated_ds`) 추가
- [x] 2. `delivery_status.py` (신규) — `reconcile_*.py` 패턴의 최상위 CLI
      - `load_sheets()` — Excel에서 SO_국내·DN_국내 로드 (SQLite 아님 — 최신성)
      - `attach_ship_status()` — DN 수량 누계로 미출고/부분출고/공장출고/출고완료 파생
      - `order_exw_date()` — 주문 내 미정 라인이 섞이면 전체 `처리중`
      - `resolve_customer()` — 사업자번호 정규화 조회(`utils.normalize_biz_no` 재사용) + 거래처명 부분일치 폴백
      - `build_summary()` — SO_ID 단위 집계 / `build_detail()` — 라인 단위
      - `write_output()` — 2시트 xlsx (Excel 표 스타일 + 금액 서식 + 인쇄 머리글)
- [x] 3. `create_po.bat` — `[기타]` 그룹에 `[C] 거래처 납기현황 조회` 메뉴 + `:delivery_status` 라벨
- [x] 4. `tests/test_delivery_status.py` — 48개
- [x] 5. `CLAUDE.md` · `docs/CHANGELOG.md` · `.gitignore`(`generated_ds/`) 갱신

## CLI 인터페이스
```
python delivery_status.py 615-81-88675      # 사업자번호 (하이픈 유무 무관)
python delivery_status.py 엔이에스           # 거래처명 부분일치 (후보 여럿이면 목록 출력 후 종료)
python delivery_status.py --list            # 미출고 잔량이 있는 거래처 목록 (사업자번호·건수·금액)
python delivery_status.py 615-81-88675 --all  # 출고완료 포함 전체
```

## 출력 — `generated_ds/납기현황_거래처명_YYMMDD.xlsx`

**시트1 `납기현황`** (고객 회신용, 주문 × 공장 출고일 1행)

| Customer PO | Remarks | 수량 | NOAH 공장 출고일 | Sales 금액 | PO receipt date |
|---|---|---|---|---|---|

- `수량` = 미출고 잔량 합 (전량 미출고면 주문수량과 동일 → 메일 표와 같게 보인다)
- `NOAH 공장 출고일` = `EXW NOAH`, 빈 값이면 **`처리중`**
- `Sales 금액` = `Sales Unit Price × 미출고수량` 합
- **분할 납기**: 한 주문 안에서 `EXW NOAH`가 갈리면 날짜별로 행을 나눈다(139건 중 10건).
  나뉜 행은 비고만으로 구분이 안 될 때만 `(품목: ...)`를 덧붙인다
- 정렬: 공장 출고일 오름차순(처리중은 뒤) → Customer PO

**시트2 `상세`** (내부 확인용, SO 라인 1행)

SO_ID · Line item · Customer PO · Remarks · Item name · 주문수량 · 출고수량 · 미출고수량 ·
Sales Unit Price · 미출고금액 · PO receipt date · EXW NOAH · Expected delivery date · 출고상태

## 원칙
- 고객에게 나가는 문서다 — 이미 납품된 건이 "미출고"로 찍히는 일이 없어야 한다 (캐시 Status 불신)
- 새로고침을 사용자 기억에 의존하지 않는다 — 도구가 스스로 계산하고, 시트가 낡았으면 알려준다
- 기존 코드 재사용 — `normalize_biz_no`, `load_customer_domestic`, `config` 시트 상수
- 국내만. 해외(SO_해외)는 이번 범위 밖. GUI(noah_gui.py)도 이번 범위 밖 — CLI만 (사용자 결정)

## Review

계획 대비 바뀐 것: **`Status`의 성격을 처음에 잘못 봤다.**
"수기 입력이라 못 믿는다"고 적었는데, 실제 셀은
`IFERROR(XLOOKUP(SO_ID&Line item, SO_통합[...], SO_통합[출고완료]), "")` 였다.
사용자 지적("파워쿼리 새로고침을 안 해서 그래")이 맞았다. 다만 결론은 그대로 — 오히려 더 명확해졌다.
`SO_통합[출고완료]`의 정의 자체가 DN 출고수량 기반이므로, CLI가 같은 식을 계산하면 **결과는 같은데
새로고침 의존이 사라진다**. 도구가 사람의 기억("실행 전에 새로고침해야지")에 기대는 구조를 없앤 셈이다.

구현하며 데이터가 알려준 것 셋:

1. **`EXW NOAH`의 빈 칸은 NaN이 아니라 `datetime.time(0,0)`** (2,023행 중 320행).
   컬럼 dtype이 object로 떨어진다. `isna()`만 검사했으면 미정 납기가 엉뚱한 날짜로 찍혔을 것이다.
2. **한 주문 안에서 `EXW NOAH`가 갈린다.** 처음엔 "하나라도 미정이면 전체를 `처리중`"으로 눌렀는데,
   사용자가 "갈리는 아이템을 구분해서 표시하라"고 정정했다. 세어 보니 미출고 139건 중 10건이 갈리고,
   대부분 **분할 납기**였다 — 세진밸브 `SOD-2026-0264`는 32라인이 7개 날짜로 2027년까지 나간다.
   한 날짜로 눌렀으면 나머지 6개 납기가 회신에서 통째로 사라졌을 것이다. → 날짜별로 행을 나눈다.
   구분 표시는 **비고만으로 안 될 때만** 품목을 덧붙였다 — 10건 중 8건은 비고가 이미 호선별로
   갈려 있어(`H2734`/`H2735`…) 품목이 잡음이었다.
3. **금액을 `단가 × 잔량`으로 낼 수 있었다.** `Sales amount = Sales Unit Price × Item qty`가
   SO 전 행(1,992/1,992)에서 성립해서, 부분출고 잔량 금액에 안분 같은 편법이 필요 없었다.

구현 중 밟은 함정: `as_date()`가 돌려준 `None`을 리스트로 모아 DataFrame 컬럼에 넣으면
pandas가 datetime64로 캐스팅해 **`None`이 `NaT`가 된다.** `d is not None`을 통과한 `NaT`가
`strftime`에서 터졌다. 날짜 계산은 컬럼에 담기 전에 끝내는 걸로 수정.

검증:
- 첨부 메일 표(엔이에스)와 **14행 전부 일치** — Customer PO / Remarks / 수량 / 공장 출고일
- 부분출고 + 무상공급 동시 케이스(`SOD-2026-0301` L9, 576 주문 / 432 출고) → 잔량 144, 금액 0.
  금액 기준 판정이었으면 못 잡았을 라인이다
- 분할 납기 세진밸브 7행 분리(품목 주석 없음) / 피엠에스·티에스엔텍만 품목 주석
- `--list` 43개 거래처 / 미출고 18.5억, 이름 조회·다중후보·미존재·인자없음 경로 확인
- `pytest tests/` — 신규 59개 포함 회귀 없음

남은 것: 실사용 후 피드백. GUI 편입은 사용자 결정으로 이번 범위 밖.

---

# 완료: GUI 사내 배포판 [2026-07-28]

문서 생성 1~7번(PO/TS/PI/FI/OC/CI/PL)을 다른 담당자 PC에서도 쓰도록 tkinter GUI + 포터블 폴더로 배포한다.
결정사항: tkinter 네이티브 창 / 포터블 폴더(단일 exe 아님) / 출력·이력은 공유 폴더 / 메일은 `.eml` 초안까지 / `create_po.bat`은 그대로 유지.

**핵심 설계**: GUI가 기존 `create_*.py`를 subprocess로 호출 → CLI 7개 파일 무수정, Excel COM은 자식 프로세스에 격리.

## 구현 (완료)
- [x] 1. `config.py` — `noah_config.ini` 폴백 (user_settings.py → ini → 기본값)
- [x] 2. `noah_gui.py` (신규) — 문서 종류 라디오 + ID 멀티라인 입력 + 옵션 + 실시간 로그
- [x] 3. 데이터 파일 지정 UI — 자동탐색(3초 제한) → 파일선택 → 시트 검증 → ini 저장. 웹 링크(https) 방어
- [x] 4. `tests/test_config_ini.py` — 13개 통과
- [x] 5. `cli_dist/build_portable_gui.py` + `requirements.txt` — 런타임/패키지/트리밍/검증/zip
- [x] 6. `설치.bat` — `%LOCALAPPDATA%`로 강제 설치 + 바탕화면 바로가기 (OneDrive KFM 회피)
- [x] 7. `.gitignore` / `CLAUDE.md` / `docs/CHANGELOG.md`

## 원칙
- 개발 PC 동작은 한 톨도 바뀌지 않는다 — `user_settings.py`가 ini보다 우선
- 생성 로직에 GUI 코드를 섞지 않는다 — CLI와 GUI가 같은 코드 경로를 쓴다
- 회사 PC는 바탕화면·문서가 전부 OneDrive(KFM)다. 런타임을 동기화 폴더에 두지 않는 건 안내가 아니라 설치 스크립트로 강제한다

## Review

계획 대비 바뀐 것 하나: **런타임을 python-build-standalone으로 교체**.
계획은 "NuGet CPython → 안 되면 임베디드 + tkinter 이식"이었는데,
NuGet 패키지(3.11.9, 1773 엔트리)에도 tkinter가 **없었다**. 이식은 conda 레이아웃
(`Library/bin/tcl86t.dll`, `Library/lib/tcl8.6`)을 python.org 임베디드에 억지로 붙이는
일이라 TCL_LIBRARY 조작이 필요하고 깨지기 쉬웠다. python-build-standalone은
tkinter·pythonw·pip을 모두 포함한 완전한 CPython이라 이식 자체가 불필요했다.

빌드를 임시 폴더로 옮긴 것도 계획에 없던 판단인데, 두 가지가 겹쳐서다:
1. **프로젝트 폴더가 OneDrive 안**이다. `cli_dist/` 아래에 런타임을 풀면 9천 개 파일이 동기화된다.
   정작 우리가 배포 대상에게 피하라고 만든 상황을 빌드가 스스로 만드는 셈이었다.
2. **MAX_PATH**. 처음엔 스크래치패드(긴 경로)에서 검증하다 pip이 파일을 못 만들어 실패했고,
   원인이 260자 제한이었다. 프로젝트 경로도 90자라 여유가 없다.
   → 짧은 임시 경로에서 빌드하고 프로젝트에는 zip 하나만 남긴다.

검증 중 잡은 실제 버그: **BOM 붙은 ini를 조용히 무시**.
PowerShell `Out-File -Encoding utf8`(=BOM 포함)로 ini를 만들어 테스트했더니 설정이 안 먹었다.
`configparser`가 첫 섹션을 `﻿[paths]`로 읽어 섹션을 못 찾고 **아무 오류 없이** 기본값으로 돌아간다.
메모장으로 ini를 편집하는 사용자가 똑같이 당한다. `utf-8-sig`로 수정 + 회귀 테스트 추가.

실사용 문제 하나 더: **엑셀을 열어둔 채 생성하면 raw 트레이스백**.
검증 도중 실제로 `PermissionError`가 났다. CLI에서는 개발자가 읽으면 그만이지만
배포판 GUI에서 트레이스백은 사용자가 대응할 수 없는 화면이다.
실행 전 파일 열기 시도로 미리 걸러 안내하고, 그 사이 잠기는 경우까지 대비해
자식 출력에 `PermissionError`가 보이면 끝에 안내를 덧붙인다.

검증 결과:
- `pytest tests/` 403 passed, 2 skipped (회귀 없음)
- 배포판 런타임: tkinter 8.6 / pandas 2.3.3 / xlwings 0.33.21 / pywin32 COM 로드 OK
- 배포판 Excel COM 왕복(한글·수식·저장) OK
- **배포판에서 실제 PI 1건 생성 성공** (SOO-2026-0013, 10 아이템, 37KB)
- GUI 위젯 트리 — 7종 옵션 패널 전환·FI 모드 전환·입력 파싱 OK
- 배포 크기 190MB → zip 70MB (파일 9,246개)

남은 것: **타 PC 검증**. `설치.bat` → 바탕화면 아이콘 → 마법사 → 문서 1건까지 통과해야 배포.
