# NOAH_PO_Lists 데이터 구조 설계

## 문서 정보
- 작성일: 2026-01-17
- 상태: **구현 완료**

---

## 1. 현재 문제점

### 현재 구조
- `NOAH_PO_Lists.xlsx` 파일의 국내/해외 시트
- **78개 컬럼**이 하나의 행에 모두 포함
- 고객 발주 → 공장 발주 → 납품 → 세금계산서 모든 프로세스를 단일 시트로 관리

### 문제점
| 문제 | 설명 |
|------|------|
| Status 불일치 | 실제 작업은 했는데 Status 업데이트 안 함 |
| 분할 납품 표현 불가 | 1주문 = 1행 구조로는 여러 번 납품 표현 어려움 |
| 데이터 변조/실수 | 누가 언제 바꿨는지 추적 불가 |
| 동시 편집 충돌 | 여러 명이 OneDrive 공유 폴더로 사용 중 |

---

## 2. 설계 방향

### 핵심 원칙
1. **ERP 스타일 정규화**: SO → PO → DN → PMT 분리
2. **복합 키 사용**: `SO_ID + Line item`으로 개별 아이템 식별
3. **발주 스냅샷 보존**: PO_국내는 값 복사 사용 (수식 참조 X)
4. **정렬 안전성**: ID 기반 참조로 정렬해도 데이터 무결성 유지

### 트랜잭션 흐름
```
고객발주 → 선수금 → 공장발주 → 입고 → 납품 → 세금계산서 → 잔금
   SO        PMT       PO              DN                    PMT
```

---

## 3. 시트 구조

### 전체 구조
```
NOAH_PO_Lists.xlsx
│
├── [Dim - 파워쿼리로 생성]
│   ├── SO_header_국내  ← SO_ID 고유값 (파워쿼리: SO_국내 → 중복 제거)
│   ├── SO_header_해외
│   ├── Customer master_국내
│   ├── Customer master_해외
│   ├── ICO            ← 가격표
│   ├── ITEM
│   └── Industry code
│
├── [Fact - 트랜잭션]
│   ├── SO_국내        ← Sales Order 라인 (고객 발주)
│   ├── SO_해외
│   ├── PO_국내        ← Purchase Order 라인 (공장 발주 + 사양/옵션)
│   ├── PO_해외
│   ├── DN_국내        ← Delivery Note (납품 + 세금계산서)
│   ├── DN_해외
│   ├── PMT_국내       ← Payment (선수금/잔금 입금)
│   └── PMT_해외
│
└── 기타 마스터...
```

### 파워피벗 관계도
```
          ┌─────────────────┐
          │ SO_header_국내  │ (Dim)
          │ SO_ID (고유)    │
          └────────┬────────┘
                   │ 1
       ┌───────────┼───────────┬───────────┐
       ▼ N         ▼ N         ▼ N         ▼ N
   ┌───────┐   ┌───────┐   ┌───────┐   ┌───────┐
   │SO_국내│   │PO_국내│   │DN_국내│   │PMT_국내│
   │(라인) │   │(라인) │   │(납품) │   │(입금) │
   └───────┘   └───────┘   └───────┘   └───────┘
```

---

## 4. 파워쿼리 / 파워피벗 설정

### 4.1 표(Table) 지정

각 시트를 표로 변환 (Ctrl+T):

| 시트 | 표 이름 | 비고 |
|------|---------|------|
| SO_국내 | tbl_SO_국내 | 직접 입력 |
| SO_해외 | tbl_SO_해외 | 직접 입력 |
| PO_국내 | tbl_PO_국내 | =SO 참조 + 직접 입력 |
| PO_해외 | tbl_PO_해외 | =SO 참조 + 직접 입력 |
| DN_국내 | tbl_DN_국내 | 직접 입력 |
| DN_해외 | tbl_DN_해외 | 직접 입력 |
| PMT_국내 | tbl_PMT_국내 | 직접 입력 |
| PMT_해외 | tbl_PMT_해외 | 직접 입력 |
| Customer master_국내 | tbl_Customer_국내 | 마스터 |
| Customer master_해외 | tbl_Customer_해외 | 마스터 |
| ICO | tbl_ICO | 가격표 |

### 4.2 파워쿼리로 SO_header 생성

**SO_header_국내 생성 단계:**

1. `tbl_SO_국내` 표 선택
2. 데이터 탭 → **테이블에서** (파워쿼리 실행)
3. 파워쿼리 편집기에서:
   - SO_ID 열만 선택 (다른 열 모두 제거)
   - 홈 탭 → **중복 제거**
4. 홈 탭 → 닫기 및 로드 → **연결만** 선택
5. 데이터 모델에 추가 체크

**파워쿼리 M 코드:**
```
let
    Source = Excel.CurrentWorkbook(){[Name="tbl_SO_국내"]}[Content],
    SelectColumn = Table.SelectColumns(Source, {"SO_ID"}),
    RemoveDuplicates = Table.Distinct(SelectColumn)
in
    RemoveDuplicates
```

### 4.3 파워피벗 관계 설정

**데이터 탭 → 관계 관리 → 새로 만들기:**

| 관계 | Dim (1) | Fact (N) | 키 |
|------|---------|----------|-----|
| 1 | SO_header_국내 | tbl_SO_국내 | SO_ID |
| 2 | SO_header_국내 | tbl_PO_국내 | SO_ID |
| 3 | SO_header_국내 | tbl_DN_국내 | SO_ID |
| 4 | SO_header_국내 | tbl_PMT_국내 | SO_ID |

### 4.4 데이터 새로고침

SO_국내 표에 새 SO_ID 추가 시:
1. 데이터 탭 → **모두 새로 고침**
2. SO_header_국내 자동 갱신
3. 파워피벗 관계 유지

---

## 5. 시트별 상세 설계

### 5.1 SO (Sales Order) - 고객 발주

**역할**: 고객으로부터 받은 발주 정보

**컬럼 목록 (공통)**:
| 컬럼명 | 설명 |
|--------|------|
| SO_ID | 키 (SOD-2026-0001 / SOO-2026-0001 형식) |
| Customer PO | 고객 발주번호 |
| PO receipt date | 발주 접수일 |
| Period | 기간 (yyyy-MM) |
| AX Period | AX 기간 (yyyy-MM) |
| AX Project number | AX 프로젝트번호 |
| AX Item number | AX 품목번호 |
| CS담당자 | CS 담당자 |
| Business registration number | 사업자등록번호 |
| Customer name | 고객명 |
| Order type | 주문 유형 |
| Opportunity | 기회 |
| Sector | 섹터 |
| Industry code | 산업코드 |
| Item name | 품목명 |
| OS name | OneStream Item name |
| Currency | 통화 |
| **Line item** | 아이템 순번 (같은 SO_ID 내 1, 2, 3...) |
| Item qty | 수량 |
| Sales Unit Price | 판매 단가 |
| Sales amount | 판매 금액 |
| Incoterms | 인코텀즈 |
| Requested delivery date | 요청 납기 (고객 희망일) |
| EXW NOAH | EXW 출고 기준일 |
| Expected delivery date | 예상 납기 (실제 예상 도착일) |
| Status | 상태 (Cancelled 등) |
| 납품 주소 | 배송지 |
| 영업 담당 | 영업 담당자 |
| Remarks | 비고 |

**해외 전용 컬럼**:
| 컬럼명 | 설명 |
|--------|------|
| Model number | 모델 번호 |
| Sales amount KRW | 판매 금액 (원화 환산) |
| Shipping method | 운송 방법 |

**SO_ID 규칙**:
- 형식: `SOD-YYYY-NNNN` (국내), `SOO-YYYY-NNNN` (해외)
- 같은 주문의 여러 품목은 **같은 SO_ID + 다른 Line item** 사용
- **복합 키**: `SO_ID + Line item`으로 개별 아이템 고유 식별
- 피벗테이블로 SO_ID 기준 그룹핑 가능
- **참고**: 코드에서는 `SO_ID` 단독으로 조인 (DN/PMT 조회 시 Line item 없이 SO_ID만 사용)

**예시**:
```
┌──────────────┬───────────┬──────────┬────────┬───────┬──────┐
│ SO_ID        │ Line item │ 고객PO   │ 고객명 │ 모델  │ 수량 │
├──────────────┼───────────┼──────────┼────────┼───────┼──────┤
│ SOD-2026-0001│ 1         │ ABC-123  │ A사    │ IQ10  │ 10   │
│ SOD-2026-0001│ 2         │ ABC-123  │ A사    │ IQ18  │ 20   │
│ SOD-2026-0001│ 3         │ ABC-123  │ A사    │ IQ25  │ 5    │
│ SOD-2026-0002│ 1         │ DEF-456  │ B사    │ NA038 │ 100  │
└──────────────┴───────────┴──────────┴────────┴───────┴──────┘
```

---

### 5.2 PO (Purchase Order) - 공장 발주

**역할**: RCK에서 NOAH 공장으로 보내는 발주 정보 + 제품 사양/옵션

**입력 방식**:
- SO 정보는 **값 복사** (Paste Values)로 입력
- `SO_ID + Line item` 복합 키로 SO와 연결
- 사양/옵션은 직접 입력

**⚠️ 값 복사를 사용하는 이유**:
| 수식 참조 | 값 복사 |
|----------|--------|
| SO 수정 시 PO도 자동 변경 | **발주 당시 데이터 그대로 유지** |
| 이미 발주한 내역이 소급 변경될 위험 | **발주 스냅샷 보존** |
| SO 정렬 시 PO 데이터 깨짐 | **정렬해도 무결성 유지** |

> **발주서는 "발주 시점의 기록"** 이므로, 나중에 SO에서 고객명이나 수량을 수정해도 이미 나간 발주 내역은 바뀌면 안 됨.

**컬럼 구성**:

| 구분 | 컬럼 | 입력 방식 |
|------|------|-----------|
| 키 | SO_ID, **Line item** | 값 복사 (SO와 매칭용) |
| SO 정보 | 고객명, 모델, 수량 등 | 값 복사 |
| PO 정보 | RCK Order no., NOAH O.C No., 공장 발주 날짜, 공장 EXW date, Status | 직접 입력 |
| 가격 | ICO Unit, Total ICO | 수식 계산 |
| 사양 | Power supply, Motor(kW), BASE, ACT Flange, Operating time, Handwheel, RPM, Turns, Bushing, MOV, Gearbox model/Flange/ratio/position, Operating mode, Fail action, Enclosure, Cable entry, Paint, Cover tube(mm), WD code, Test report, Version, Note | 직접 입력 |
| 옵션 (Y/N) | Model, Bush, ALS, EXT, DC24V, Modbus/Profibus, LCU, PIU, CPT+PIU, PCU+PIU, -40, -60, SCP, EXP, Bush-SQ, Bush-STAR, INTEGRAL, IMS, BLDC, HART/Foundation Fieldbus, ATS | 직접 입력 |

**동적 사양/옵션 필드 감지**:

코드에서 사양/옵션 컬럼을 하드코딩하지 않고 `get_spec_option_fields()` 함수로 동적으로 감지합니다:

```
PO 시트 전체 컬럼
├── META_COLUMNS (config.py에 정의된 메타 컬럼) → 제외
├── SPEC_START_COLUMN ~ Status 전까지 → SPEC_FIELDS (사양)
└── OPTION_START_COLUMN ~ 끝까지 → OPTION_FIELDS (옵션 Y/N)
```

- `META_COLUMNS`: SO_ID, Customer name, Model 등 메타데이터 컬럼 (frozenset)
- 사양/옵션 컬럼은 PO 시트에 새 컬럼이 추가되면 자동으로 반영됨
- Description 시트 생성 시 동적 감지된 필드 목록 사용

**ICO Unit 계산 수식** (기존 유지):
```
=XLOOKUP($BF2&$BF2,ICO!$E:$E,ICO!$C:$C,0)
 +SUMPRODUCT((BG2:BZ2="Y")*XLOOKUP($BF2&BG$1:BZ$1,ICO!$E:$E,ICO!$C:$C,0))
```

**예시**:
```
┌──────────────┬───────────┬────────┬───────┬───────────┬───────────┬─────────────┐
│ SO_ID        │ Line item │ 고객명 │ 모델  │ RCK Order │ 발주일    │ 사양/옵션...│
│ (값복사)     │ (값복사)  │(값복사)│(값복사)│ (직접)    │ (직접)    │ (직접)      │
├──────────────┼───────────┼────────┼───────┼───────────┼───────────┼─────────────┤
│ SOD-2026-0001│ 1         │ A사    │ IQ10  │ ND-0001   │ 2026-01-15│ ...         │
│ SOD-2026-0001│ 2         │ A사    │ IQ18  │ ND-0001   │ 2026-01-15│ ...         │
│ SOD-2026-0001│ 3         │ A사    │ IQ25  │ ND-0001   │ 2026-01-15│ ...         │
└──────────────┴───────────┴────────┴───────┴───────────┴───────────┴─────────────┘
```

---

### 5.3 DN (Delivery Note) - 납품

**역할**: 납품 기록 + 세금계산서 발행 정보

**특징**:
- SO_ID로 연결 (행 번호 동기화 불필요)
- **분할 납품 가능** (같은 SO_ID로 여러 행)
- 세금계산서 정보 포함

**컬럼 목록**:
| 컬럼명 | 설명 |
|--------|------|
| DN_ID | 키 (자동 또는 수동) |
| SO_ID | Sales Order 참조 |
| 납품일 | 실제 납품 날짜 |
| 납품 수량 | 이번 납품 수량 |
| 거래명세표 파일 | TS_xxx.xlsx |
| 세금계산서 번호 | 발행 시 기록 |
| 세금계산서 발행일 | |
| 비고 | |

**예시**:
```
┌──────────────┬───────────┬──────┬──────────────┬────────────┐
│ SO_ID        │ 납품일    │ 수량 │ 세금계산서   │ 발행일     │
├──────────────┼───────────┼──────┼──────────────┼────────────┤
│ SO-2026-0001 │ 2026-01-20│ 20   │ 2026-00123   │ 2026-01-22 │
│ SO-2026-0001 │ 2026-01-25│ 15   │ 2026-00145   │ 2026-01-26 │
│ SO-2026-0002 │ 2026-01-22│ 100  │              │            │ ← 미발행
└──────────────┴───────────┴──────┴──────────────┴────────────┘
```

---

### 5.4 PMT (Payment) - 입금

**역할**: 선수금/잔금 입금 기록

**특징**:
- SO_ID로 연결
- 선수금은 납품 전에 발생 가능
- 여러 번 분할 입금 가능

**컬럼 목록**:
| 컬럼명 | 설명 |
|--------|------|
| 선수금_ID | 키 (ADV_YYYY-NNNN 형식, `create_ts.py`에서 사용) |
| SO_ID | Sales Order 참조 |
| 구분 | 선수금 / 잔금 / 완납 |
| 금액 | 입금 금액 |
| 입금일 | |
| 비고 | |

**예시**:
```
┌──────────────┬────────┬──────────┬────────────┬──────────┐
│ SO_ID        │ 구분   │ 금액     │ 입금일     │ 비고     │
├──────────────┼────────┼──────────┼────────────┼──────────┤
│ SO-2026-0001 │ 선수금 │ 5,000,000│ 2026-01-12 │ 50%      │
│ SO-2026-0001 │ 잔금   │ 5,000,000│ 2026-01-26 │          │
│ SO-2026-0002 │ 완납   │10,000,000│ 2026-01-30 │ 납품 후  │
└──────────────┴────────┴──────────┴────────────┴──────────┘
```

---

## 6. 관계도 (ERD)

```
┌─────────────────────────────────────────────────────────────────────┐
│                              마스터                                  │
├─────────────────────────────────────────────────────────────────────┤
│  Customer master    ICO (가격표)    ITEM    Industry code           │
└─────────────────────────────────────────────────────────────────────┘
                                │
                                ▼
┌─────────────────────────────────────────────────────────────────────┐
│  SO_국내 (Sales Order)                                              │
│  ┌───────────────────────────────────────────────────────────────┐ │
│  │ SO_ID + Line item (복합PK) │ 고객PO │ 고객명 │ 모델 │ 수량 │...│ │
│  └───────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────────┘
        │                               │                    │
        │ 복합키 (값복사)               │ 1:N               │ 1:N
        ▼                               ▼                    ▼
┌───────────────────┐         ┌─────────────────┐    ┌─────────────────┐
│  PO_국내          │         │  DN_국내        │    │  PMT_국내       │
│  (Purchase Order) │         │  (Delivery)     │    │  (Payment)      │
├───────────────────┤         ├─────────────────┤    ├─────────────────┤
│ SO_ID + Line item │         │ SO_ID (FK)      │    │ SO_ID (FK)      │
│ (값복사, 스냅샷)  │         │ 납품일, 수량    │    │ 구분 (선수금 등)│
│ + 사양/옵션       │         │ 세금계산서      │    │ 금액, 입금일    │
└───────────────────┘         └─────────────────┘    └─────────────────┘
```

### 왜 PO는 값 복사인가?

**발주서 = 발주 시점의 스냅샷**

- SO_국내: 원본 데이터 (수정 가능)
- PO_국내: 발주 당시 데이터 복사본 (변경 불가)
- DN/PMT: SO_ID로 연결 (실시간 참조 가능)

```
SO_국내 수정 → PO_국내 영향 없음 (이미 발주된 내역 보존)
            → DN/PMT는 SO_ID로 조회 시 최신 정보 참조
```

---

## 6.1 ERP 관점에서 본 데이터 구조

이 엑셀 기반 시스템은 ERP의 핵심 원리를 그대로 구현하고 있다.

### 테이블 관계 = ERP의 FK/PK 관계

| ERP 개념 | 엑셀 구현 | 설명 |
|----------|----------|------|
| Primary Key | `SO_ID + Line item` | 각 시트의 행을 고유 식별하는 복합 키 |
| Foreign Key | SO_ID (DN, PMT에서 참조) | 트랜잭션 간 연결 |
| Dimension Table | SO_header (파워쿼리 생성) | SO_ID 고유값 → 1:N 관계의 1 쪽 |
| Fact Table | SO, PO, DN, PMT | 실제 트랜잭션 데이터 |
| Master Table | Customer master, ICO, ITEM | 참조 데이터 |

### 데이터 참조 = ERP의 FK Lookup

```
엑셀 수식:
=XLOOKUP(SO_ID, SO_국내[SO_ID], SO_국내[Customer name])

ERP SQL:
SELECT c.customer_name FROM sales_order s JOIN customer c ON s.customer_id = c.id
```

둘 다 **키 기반으로 다른 테이블의 값을 참조**하는 동일한 원리.

### 데이터 집계 = ERP의 SQL GROUP BY

```
Power Query:
Table.Group(PO_Combined, {"SO_ID", "Line item"}, {
    {"Total ICO", each List.Sum([Total ICO]), type number}
})

ERP SQL:
SELECT SO_ID, Line_item, SUM(Total_ICO) FROM PO GROUP BY SO_ID, Line_item
```

같은 키의 여러 행을 합산하는 로직. PO 사양 분리(1:N)나 DN 분할 출고(N:1) 모두 이 패턴으로 처리.

### 테이블 간 관계 유형

```
SO ──(1:1)── PO     기본 관계 (SO Line 1개 = PO Line 1개)
SO ──(1:N)── PO     사양 분리 시 (SO Line 1개 → PO 여러 행, Table.Group으로 합산)
SO ──(1:N)── DN     분할 출고 시 (SO Line 1개 → DN 여러 행, Table.Group으로 합산)
SO ──(1:N)── PMT    분할 입금 시 (SO 1건 → 선수금/잔금 여러 행)
```

ERP에서도 SO-DN, SO-PO는 1:N이 기본이며, 집계 시 GROUP BY로 합산한다.

### 상태 관리 = ERP Workflow

```
PO 상태:  Open → Sent → Confirmed → Invoiced (→ Cancelled)
출고 상태: 미출고 → 부분 출고 → 출고 완료
```

| ERP 개념 | 엑셀 구현 |
|----------|----------|
| Order Status | PO의 Status 컬럼 (Open/Sent/Confirmed/Invoiced/Cancelled) |
| Delivery Status | SO_통합 쿼리의 출고완료 (미출고/부분 출고/출고 완료) |
| 미출고금액 | Sales amount - 출고금액 (ERP의 "Deliver remainder") |

### 스냅샷 보존 = ERP의 트랜잭션 불변성

| ERP 원칙 | 엑셀 구현 |
|----------|----------|
| 발주서는 발주 시점의 기록 | PO는 SO에서 **값 복사** (수식 참조 X) |
| SO 변경이 기발주에 영향 없음 | PO 행은 발주 후 수정하지 않음 |
| 추가 발주는 새 트랜잭션 | PO에 새 행 추가 (기존 행 수정 X) |
| 취소도 기록으로 남김 | Status=Cancelled (행 삭제 X) |

### 조인 방식 = ERP의 테이블 조인

| Power Query | SQL | 용도 |
|-------------|-----|------|
| `Table.NestedJoin(..., JoinKind.LeftOuter)` | `LEFT OUTER JOIN` | SO 기준으로 PO/DN 매칭 (없어도 SO는 표시) |
| `Table.Group(..., List.Sum)` | `GROUP BY + SUM` | 같은 키의 금액 합산 (분할 출고/사양 분리) |
| `Table.Distinct(..., {"SO_ID"})` | `SELECT DISTINCT` | 중복 제거 (SO_header 생성) |
| `Table.SelectRows(..., [Status] <> "Cancelled")` | `WHERE Status <> 'Cancelled'` | 취소 건 필터링 |

### 차이점: ERP vs 엑셀

| 항목 | ERP | 현재 엑셀 시스템 |
|------|-----|-----------------|
| 데이터 무결성 | RDBMS 제약조건 (NOT NULL, FK, UNIQUE) | 수동 입력에 의존 |
| 트랜잭션 보장 | ACID 트랜잭션 | 없음 (동시 편집 시 충돌 가능) |
| 권한 관리 | 역할별 접근 제어 | 시트 보호 수준 |
| 감사 추적 | 변경 이력 자동 기록 | po_history로 부분 추적 |
| 자동 채번 | 시퀀스/Auto-increment | 수동 ID 입력 |

> **결론**: 테이블 관계, 조인, 집계, 상태 관리, 스냅샷 보존 등 **데이터 처리 원리는 ERP와 동일**하다. 차이는 무결성 보장 수준뿐이며, 이는 ERP 통합 전까지 운용 규칙으로 보완한다.

---

## 6.2 ERP 모듈별 관계 구조 (학습 참고)

ERP 시스템의 핵심은 **하나의 키(SO_ID)가 모든 하위 문서에 FK로 관통**하는 구조이다. 현재 엑셀 시스템도 이 원리를 동일하게 따른다.

### SO_ID 중심의 1:N 관계 — ERP vs 현재 시스템

#### Sales Module (판매)

```
ERP (SAP/D365):
  Sales Order Header (VBELN/SalesId) ← PK
      ├── SO Line Items (N)       ← 분할 품목
      ├── Deliveries (N)          ← 분할 출고
      ├── Invoices (N)            ← 분할 청구
      └── Payments (N)            ← 분할 입금

현재 시스템:
  SO_header_국내 (SO_ID) ← PK
      ├── SO_국내 (N)              ← Line item으로 분할
      ├── PO_국내 (N)              ← 추가 발주 시 새 행
      ├── DN_국내 (N)              ← 분할 납품
      └── PMT_국내 (N)             ← 선수금/잔금
```

**동일한 구조.** SO_ID 하나로 주문의 모든 트랜잭션을 추적할 수 있다.

#### Procurement Module (구매/발주)

```
ERP:
  Purchase Requisition (1) → Purchase Orders (N)    ← 분할 발주
  PO Header (1)            → PO Line Items (N)
                           → Goods Receipts (N)     ← 분할 입고
                           → Vendor Invoices (N)    ← 분할 청구

현재 시스템:
  SO (1)                   → PO (N)                 ← 추가/분할 발주
  PO 각 행                 → 사양/옵션 컬럼          ← ERP의 PO Line 속성
```

ERP에서 구매요청(PR) 1건에 대해 여러 PO를 발행할 수 있듯이, 현재 시스템에서도 SO 수량 변경 시 PO에 새 행을 추가한다 (섹션 7.0 참조).

#### Inventory Module (재고/창고)

```
ERP:
  Delivery Note (1)  → Transfer Orders (N)   ← 창고 내 이동
                     → Picking Lists (N)      ← 출고 지시
                     → Packing Slips (N)      ← 포장 단위
```

현재 시스템에서는 DN_국내가 이 역할을 단순화하여 처리한다.

#### Finance Module (재무/수금)

```
ERP:
  Invoice (1) → Payment Allocations (N)    ← 분할 입금
  Customer (1) → Open Items (N)            ← 미수금 잔액 관리

현재 시스템:
  SO (1) → PMT (N)    ← 선수금/잔금 분리
  SO_통합 쿼리         ← 미출고금액 = Sales amount - 출고금액
```

### Header-Line 패턴 — 모든 ERP 문서의 기본 구조

ERP의 거의 모든 문서는 **Header(1) → Line(N)** 구조를 따르며, **Header ID + Line 번호**가 복합 키(Composite Key)로 개별 행을 식별한다:

```
SAP:   VBAK (SO Header)       → VBAP (SO Item)          ← 1:N
       복합 키: VBELN + POSNR  (예: 10001 + 10, 10001 + 20, 10001 + 30)
       EKKO (PO Header)       → EKPO (PO Item)          ← 1:N
       복합 키: EBELN + EBELP  (예: 45001 + 10, 45001 + 20)
       LIKP (Delivery Header)  → LIPS (Delivery Item)    ← 1:N
       복합 키: VBELN + POSNR

D365:  SalesOrderHeader       → SalesOrderLine           ← 1:N
       복합 키: SalesId + LineNum  (예: SO-0001 + 1, SO-0001 + 2)
       PurchaseOrderHeader    → PurchaseOrderLine         ← 1:N
       복합 키: PurchId + LineNum

현재:  SO_header_국내          → SO_국내 (Line item)      ← 1:N
       복합 키: SO_ID + Line item  (예: SOD-2026-0001 + 1, SOD-2026-0001 + 2)
```

> **SAP 채번 방식**: Line 번호를 10 단위(10, 20, 30...)로 부여한다. 나중에 기존 라인 사이에 행을 삽입할 수 있도록 간격을 두는 것. 현재 시스템은 1, 2, 3 순번이지만, 엑셀에서는 행 삽입이 자유로우므로 문제없다.

#### 현재 시스템과의 차이: Flat Table 방식

ERP는 Header 테이블과 Line 테이블이 **물리적으로 분리**되어 있다:

```
ERP (물리적 2개 테이블):
  VBAK (Header): VBELN=10001, 고객명=A사, 주문일=2026-01-15
  VBAP (Line):   VBELN=10001, POSNR=10, 모델=IQ10, 수량=10
                 VBELN=10001, POSNR=20, 모델=IQ18, 수량=20
  → 고객명을 보려면 Header JOIN 필요
```

현재 시스템은 엑셀 직접 입력이므로 **Flat Table(단일 시트)** 방식이다. Header 정보가 매 행마다 반복된다:

```
현재 (Flat Table — 1개 시트):
  SO_국내:  SOD-0001, Line 1, A사, IQ10, 10   ← 고객명 반복
            SOD-0001, Line 2, A사, IQ18, 20   ← 고객명 반복
            SOD-0002, Line 1, B사, NA038, 100
  → 조인 없이 한 행에 모든 정보가 있음
```

**Flat Table의 장단점**:

| 항목 | ERP (Header-Line 분리) | 현재 (Flat Table) |
|------|----------------------|-------------------|
| 데이터 중복 | 없음 (정규화) | 고객명 등 Header 정보 반복 |
| 입력 편의성 | Header 먼저 → Line 입력 | 한 행에 모두 입력 (엑셀에 자연스러움) |
| 수정 시 일관성 | Header 1곳만 수정 | 같은 SO_ID의 모든 행 수정 필요 |
| 조회 편의성 | JOIN 필요 | 바로 보임 |

> **SO_header_국내** 시트(파워쿼리로 생성)가 ERP의 Header 테이블 역할을 대신한다. Flat Table에서 SO_ID 고유값을 추출하여 Dim 테이블로 만든 것이므로, 파워피벗에서는 ERP와 동일한 1:N 관계가 성립한다.

### Document Flow — 문서 체인

ERP에서는 문서 간 연결을 별도 테이블로 추적한다:

```
SAP:  VBFA (Document Flow) 테이블
      SO 10001 → DN 80001 → Invoice 90001 → Payment 14001
      모든 연결이 VBFA에 기록됨

D365: InventTransOrigin 테이블
      SO → Packing Slip → Invoice → Payment Journal

현재 시스템:
      SO_ID를 FK로 사용하여 동일한 효과
      SO_ID = SOD-2026-0001 → PO, DN, PMT 모두 조회 가능
```

SAP은 별도 추적 테이블(VBFA)을 두지만, 현재 시스템은 SO_ID FK만으로 충분하다. 문서 유형이 4개(SO/PO/DN/PMT)로 단순하기 때문.

### 다대다(M:N) 관계 — ERP의 마스터 데이터

위의 트랜잭션(SO→PO→DN→PMT)은 모두 **1:N** 관계이다. ERP에서 **다대다** 관계는 주로 마스터 데이터 간에 발생하며, 중간 테이블(Junction Table)로 분해한다:

| 다대다 관계 | 중간 테이블 | 추가 속성 |
|------------|-----------|----------|
| 사용자 ↔ 권한 | User_Role | 부여일, 만료일 |
| 제품 ↔ 공급업체 | Item_Vendor | 단가, 리드타임, 우선순위 |
| 완제품 ↔ 부품 (BOM) | BOM_Line | 수량, 공정순서 |
| 창고 ↔ 제품 | Inventory | 수량, 로트번호 |
| 주문 ↔ 할인/프로모션 | Order_Discount | 적용 금액 |

```
다대다 분해 원리:

  Product (1) ──→ BOM_Line (N) ←── Part (1)
                  제품ID(FK)        부품ID(PK)
                  부품ID(FK)
                  수량, 공정순서

  → 다대다를 1:N + N:1 두 개로 분해
  → 중간 테이블이 양쪽 PK를 FK로 보유
  → 관계 자체의 속성(수량, 단가 등)도 중간 테이블에 저장
```

현재 시스템에서 ICO 테이블이 이 패턴에 가깝다:

```
  Model ↔ Option → ICO가 중간 테이블 역할
  IQ10 + Bush 옵션 → ICO 단가
  IQ10 + ALS 옵션 → ICO 단가
  모델과 옵션의 조합마다 가격이 다름
```

BOM 테이블 구현 계획은 섹션 13 참조.

### 스냅샷 보존 = 전기(Posting) 불변성

ERP에서 가장 중요한 원칙 중 하나:

```
SAP:  전기(Posting) 후 문서는 수정 불가
      → 수정이 필요하면 반대 전기(Reversal)로 취소 후 재생성

D365: Posted documents are immutable
      → 수정 시 Credit Note / Correction Journal 생성

현재 시스템:
      PO는 값 복사(스냅샷) → 원본 SO 수정해도 기발주 영향 없음
      취소 시 Status=Cancelled → 행 삭제 X (기록 보존)
      추가 발주 시 새 행 추가 → 기존 행 수정 X
```

**핵심**: 한번 확정된 트랜잭션은 절대 수정하지 않는다. 변경이 필요하면 새 트랜잭션을 만든다.

---

## 7. 운영 규칙

### 7.0 SO 수량 변경 시 운용 방식

#### 배경

초기 설계는 SO_국내와 PO_국내가 **같은 행**으로 연동되는 구조였음:

```
SO_국내 행 1  ←→  PO_국내 행 1
SO_국내 행 2  ←→  PO_국내 행 2
```

**문제**: SO 수량 변경 시 PO에 추가 발주가 필요하면 행이 안 맞음.

#### 변경된 구조 (1:N 관계)

SO와 PO를 **별개 트랜잭션**으로 관리:

```
SO_국내: 고객 주문 마스터 (최신 상태 유지)
├── SOD-0001, Line item 1, Item A, 수량 15  ← 현재 고객 요청 수량

PO_국내: 공장 발주 트랜잭션 (발주 이력)
├── POD-0001, SOD-0001, Line item 1, Item A, 수량 10  (1차 발주)
├── POD-0002, SOD-0001, Line item 1, Item A, 수량 5   (추가 발주)
└── POD-0003, SOD-0001, Line item 1, Item A, 수량 -3  (취소 시)
```

**관계**: SO_ID + Line item 기준 **1:N**

#### SO_통합 쿼리에서 보면?

위 예시가 SO_통합 파워쿼리를 통과하면 **PO의 N개 행이 GROUP BY로 합산**되어 SO 1행에 붙는다:

```
① PO 원가 합산 (SO_ID + Line item 기준 GROUP BY)
┌──────────┬───────────┬────────────────────────────────────────────┐
│ SO_ID    │ Line item │ Total ICO = SUM(각 PO행의 Total ICO)       │
├──────────┼───────────┼────────────────────────────────────────────┤
│ SOD-0001 │ 1         │ 10개분 원가 + 5개분 원가 + (-3개분) = 12개분 │
└──────────┴───────────┴────────────────────────────────────────────┘
  POD-0001 (10개) + POD-0002 (5개) + POD-0003 (-3개) → 합산

② SO_통합 최종 결과 (SO 1행 + 합산된 원가 + DN 출고)
┌──────────┬───────────┬────────┬──────┬──────────┬────────┬────────┬────────────┐
│ SO_ID    │ Line item │ Item   │ 수량  │ Sales KRW│ 원가   │ 마진   │ 출고완료    │
├──────────┼───────────┼────────┼──────┼──────────┼────────┼────────┼────────────┤
│ SOD-0001 │ 1         │ Item A │ 15   │ 1,500만  │ 900만  │ 600만  │ 미출고      │
└──────────┴───────────┴────────┴──────┴──────────┴────────┴────────┴────────────┘
  ↑ PO가 3행이었지만 GROUP BY로 합산 → SO 1행에 원가 1개 값으로 조인
  ↑ DN이 아직 없으면 출고완료 = "미출고"
```

**핵심**: PO에 추가 발주/취소 행이 아무리 많아도, SO_통합에서는 `SO_ID + Line item` 기준으로 합산되어 **SO 1행 = 원가 1값**으로 정리된다. 이것이 1:N 관계를 GROUP BY로 집계하는 ERP의 기본 패턴이다.

#### 운용 규칙

| 상황 | SO_국내 | PO_국내 |
|------|---------|---------|
| 최초 발주 | 수량 10 입력 | 새 행, 수량 10 |
| 추가 발주 | 수량 10→15 수정 | **새 행 추가**, 수량 5 |
| 수량 감소 | 수량 15→12 수정 | **새 행 추가**, 수량 -3 또는 Status=Cancelled |
| 전체 취소 | Status=Cancelled | 해당 PO 행들 Status=Cancelled |

#### PO_현황 쿼리

Power Query로 SO별 발주 현황 집계:
- **발주수량**: PO 합계 (Sent/Confirmed/Invoiced만, Open/Cancelled 제외)
- **미발주수량**: SO수량 - 발주수량
- **발주완료**: 미발주수량 ≤ 0이면 Y

상세 내용은 `docs/POWER_QUERY.md`의 **PO_현황** 섹션 참조.

### 7.1 입력 흐름

```
1. 고객 발주 접수
   → SO_국내에 행 추가 (SO_ID 생성)
   → 같은 주문의 여러 품목은 같은 SO_ID로 여러 행

2. 선수금 입금 (있는 경우)
   → PMT_국내에 행 추가 (구분: 선수금)

3. 공장 발주
   → PO_국내 같은 행에 사양/옵션 입력 (SO 정보는 자동 참조)
   → python create_po.py SO-2026-0001 실행

4. 납품
   → DN_국내에 행 추가
   → python create_ts.py SO-2026-0001 실행

5. 세금계산서 발행
   → DN_국내 해당 행에 세금계산서 번호/발행일 입력

6. 잔금 입금
   → PMT_국내에 행 추가 (구분: 잔금)
```

### 7.2 시트 보호 및 정렬 주의사항

**SO_국내, SO_해외**:
- 행 삽입/삭제: 허용
- **정렬 시 주의**: Line item 순서 유지 필요
- 정렬해도 `SO_ID + Line item` 복합 키로 PO와 매칭 가능

**PO_국내**:
- 값 복사로 저장 (수식 참조 X)
- 발주 후에는 수정하지 않음 (스냅샷 보존)
- 정렬해도 데이터 무결성 유지

**PO_해외**:
- 단일 편집자 사용으로 수식 참조 유지 가능
- 필요 시 값 복사로 전환

**DN_국내, PMT_국내**:
- 행 삽입/삭제: 자유
- SO_ID로 연결되므로 행 번호 무관

### 7.3 현황 파악 방법

| 확인 항목 | 방법 |
|-----------|------|
| 공장 발주 안 된 건 | SO에는 있는데 PO의 RCK Order가 비어있는 행 |
| 납품 안 된 건 | SO에는 있는데 DN에 해당 SO_ID가 없음 |
| 세금계산서 미발행 | DN에서 세금계산서 번호가 비어있는 행 |
| 선수금 안 받은 건 | SO에는 있는데 PMT에 "선수금" 구분이 없음 |
| 잔금 안 받은 건 | DN(납품완료)은 있는데 PMT에 "잔금" 없음 |

---

## 8. 양식 자동화 연결

### Python 명령어 (예정)

```bash
# 발주서 생성
python create_po.py SO-2026-0001
# → SO + PO JOIN → 발주서 템플릿

# 거래명세표 생성
python create_ts.py SO-2026-0001
# → SO + DN JOIN → 거래명세표 템플릿

# Commercial Invoice 생성 (해외)
python create_invoice.py SO-2026-0001
# → SO + PO + DN + Customer master JOIN → Invoice 템플릿

# Packing List 생성 (해외)
python create_packing.py SO-2026-0001
# → SO + PO + ITEM(Weight, CBM) JOIN → Packing 템플릿
```

### JOIN 관계

| 문서 | 필요한 시트 |
|------|-------------|
| 발주서 (PO) | SO + PO + Customer master + ICO |
| 거래명세표 (TS) | SO + DN + Customer master |
| Commercial Invoice | SO + PO + DN + Customer master |
| Packing List | SO + PO + DN + ITEM (Weight, CBM) |

---

## 9. 마이그레이션 계획

### 현재 → 신규 구조

1. **기존 국내/해외 시트 백업**

2. **SO 시트 생성**
   - 기존 컬럼 중 SO 관련만 복사
   - SO_ID 컬럼 추가

3. **PO 시트 생성**
   - SO 참조 수식 설정 (`=SO_국내!A2` 등)
   - 기존 사양/옵션 컬럼 복사

4. **DN/PMT 시트 생성**
   - 빈 시트로 시작
   - 기존 납품/입금 이력은 수동 또는 스크립트로 이전

5. **시트 보호 설정**
   - SO/PO: 행 삭제 금지
   - DN/PMT: 자유

---

## 10. 백업

**Power Automate로 매일 자동 백업**:
- `NOAH_SO_PO_DN.xlsx` → 별도 백업 폴더로 복사
- 일별 스냅샷 보관

---

## 11. 현재 구현 상태

### 시트별 컬럼 수
| 시트 | 컬럼 수 | 역할 |
|------|--------|------|
| SO_국내 | 29개 | 고객 발주 정보 |
| SO_해외 | 32개 | 고객 발주 정보 (해외) |
| PO_국내 | 55개 | 공장 발주 + 사양/옵션 |
| DN_국내 | 9개 | 납품 |
| PMT_국내 | 7개 | 입금 (선수금/잔금) |

### 전체 시트 목록
1. Notes
2. SO_header_국내
3. SO_국내
4. PO_국내
5. DN_국내
6. PMT_국내
7. SO_header_해외
8. SO_해외
9. PO_해외
10. DN_해외
11. Item detail
12. Customer_국내
13. Customer_해외
14. Item-mapping
15. ICO
16. ITEM
17. Industry code

---

## 12. TODO

- [x] 엑셀 시트 구조 생성 (수동)
- [x] 각 시트 표(Table)로 변환 (Ctrl+T)
- [x] 파워쿼리로 SO_header 생성 (중복 제거)
- [x] 파워피벗 관계 설정
- [x] 기존 데이터 마이그레이션
- [x] Line item 컬럼 추가 (SO_국내, SO_해외, PO_국내)
- [x] PO_국내 값 복사 방식으로 전환 (발주 스냅샷 보존)
- [ ] create_po.py SO_ID 기반으로 수정
- [ ] create_ts.py SO_ID 기반으로 수정
- [ ] create_invoice.py 구현 (해외용)
- [ ] create_packing.py 구현 (해외용)
- [ ] BOM (Bill of Materials) 테이블 구현 (아래 상세 설계 참조)

---

## 13. 향후 구현: BOM (Bill of Materials)

### 배경

현재 PO_국내의 사양/옵션 컬럼(Model, Bush, ALS, EXT 등)은 **완제품 단위**로만 관리된다. 완제품을 구성하는 부품 정보(모터, 기어박스, 플랜지 등)는 시스템에 없으므로 부품별 원가 계산이나 소요량 산출이 불가능하다.

### 핵심 개념: 다대다 관계

완제품과 부품은 **다대다(M:N)** 관계이다:
- 하나의 완제품에 여러 부품이 들어간다
- 하나의 부품이 여러 완제품에 쓰인다

이를 **중간 테이블(BOM_Line)**로 1:N + N:1로 분해한다.

### 테이블 설계

**BOM_Part (부품 마스터)**:
```
┌────────┬────────────┬──────────┬────────┐
│ 부품ID  │ 부품명      │ 단가      │ 비고    │
├────────┼────────────┼──────────┼────────┤
│ P001   │ 모터A      │ ₩50,000  │        │
│ P002   │ 기어박스B   │ ₩80,000  │        │
│ P003   │ 플랜지C    │ ₩30,000  │        │
│ P004   │ 기어박스D   │ ₩90,000  │        │
└────────┴────────────┴──────────┴────────┘
```

**BOM_Line (중간 테이블 — 완제품↔부품 연결)**:
```
┌────────┬────────┬──────┬────────┐
│ 제품ID  │ 부품ID  │ 수량  │ 공정순서 │
├────────┼────────┼──────┼────────┤
│ IQ10   │ P001   │ 1    │ 1      │  ← IQ10에 모터A 1개
│ IQ10   │ P002   │ 1    │ 2      │  ← IQ10에 기어박스B 1개
│ IQ10   │ P003   │ 2    │ 3      │  ← IQ10에 플랜지C 2개
│ IQ18   │ P001   │ 1    │ 1      │  ← IQ18에 모터A 1개
│ IQ18   │ P004   │ 1    │ 2      │  ← IQ18에 기어박스D 1개
│ IQ18   │ P003   │ 4    │ 3      │  ← IQ18에 플랜지C 4개
└────────┴────────┴──────┴────────┘
```

### 관계도

```
Product/ITEM (1) ──→ BOM_Line (N) ←── BOM_Part (1)
 제품ID(PK)           제품ID(FK)        부품ID(PK)
                      부품ID(FK)
                      수량
                      공정순서
```

### 기존 시트와의 연결

```
NOAH_PO_Lists.xlsx
├── 기존 시트들...
├── ITEM          ← 기존 제품 마스터 (제품ID = Model number)
├── BOM_Part      ← 신규: 부품 마스터
└── BOM_Line      ← 신규: 중간 테이블 (ITEM.제품ID + BOM_Part.부품ID)
```

### 활용 예시

| 질문 | 조회 방법 |
|------|----------|
| IQ10 만드는데 부품 뭐 필요해? | `BOM_Line WHERE 제품ID = IQ10` |
| 모터A 쓰는 완제품이 뭐야? | `BOM_Line WHERE 부품ID = P001` |
| IQ10 1대 원가? | `BOM_Line JOIN BOM_Part → SUM(수량 × 단가)` |
| SO 주문 전체 부품 소요량? | `SO JOIN BOM_Line → GROUP BY 부품ID, SUM(SO수량 × BOM수량)` |
| 모터A 100개로 IQ10 몇 대? | `BOM_Line WHERE 부품ID=P001, 제품ID=IQ10 → 100 ÷ 수량` |

### 구현 단계

1. BOM_Part 시트 생성 + 표(Table) 변환 (`tbl_BOM_Part`)
2. BOM_Line 시트 생성 + 표(Table) 변환 (`tbl_BOM_Line`)
3. 파워피벗 관계 설정: ITEM(1)→BOM_Line(N), BOM_Part(1)→BOM_Line(N)
4. 부품 데이터 입력
5. 원가 계산 파워쿼리 또는 피벗테이블 구성

---

## 변경 이력

| 날짜 | 내용 |
|------|------|
| 2026-01-17 | 초안 작성 |
| 2026-01-18 | 구현 완료, 백업 정보 추가, 현재 시트 구조 문서화 |
| 2026-01-30 | **데이터 구조 변경**: Line item 컬럼 추가 (SO_국내, SO_해외, PO_국내), PO_국내 값 복사 방식으로 전환 (발주 스냅샷 보존), 복합 키(SO_ID + Line item) 기반 아이템 식별 |
| 2026-01-31 | **SO↔PO 관계 변경**: 행 1:1 → 1:N 관계로 전환, SO 수량 변경 시 PO에 새 행 추가 방식 채택, PO_현황 Power Query 추가 |
| 2026-02-23 | **SO 컬럼 추가**: AX Period, AX Item number 컬럼을 SO_국내/SO_해외 공통 컬럼에 추가, SO_해외 컬럼 수 문서화 (32개) |
| 2026-02-28 | **ERP 매핑 섹션 추가**: 섹션 6.1 - 테이블 관계, 조인, 집계, 상태 관리, 스냅샷 보존 등 ERP 원리와의 대응 관계 문서화 |
| 2026-02-28 | **BOM 설계 추가**: 섹션 13 - 완제품↔부품 다대다 관계, BOM_Part/BOM_Line 중간 테이블 설계, 원가 계산/소요량 산출 활용 방안 |
| 2026-02-28 | **ERP 모듈 비교 추가**: 섹션 6.2 - 모듈별 1:N 관계, Header-Line 패턴, Document Flow, 다대다 관계, 스냅샷 불변성 등 ERP 원리 학습 참고 자료 |
