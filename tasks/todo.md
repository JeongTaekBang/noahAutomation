# Current Tasks — GUI 사내 배포판 [2026-07-28]

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
