# Changelog

개발 이력, 버그 수정, 리팩토링 기록.

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
