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
| PMT_ID | 키 (자동 또는 수동) |
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
├── SOD-0001, Item A, 수량 15  ← 현재 고객 요청 수량

PO_국내: 공장 발주 트랜잭션 (발주 이력)
├── POD-0001, SOD-0001, Item A, 수량 10  (1차 발주)
├── POD-0002, SOD-0001, Item A, 수량 5   (추가 발주)
└── POD-0003, SOD-0001, Item A, 수량 -3  (취소 시)
```

**관계**: SO_ID + Item name 기준 **1:N**

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
