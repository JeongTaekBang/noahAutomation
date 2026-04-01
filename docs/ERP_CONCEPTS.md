# ERP 개념 학습 노트

ERP 역설계 과정에서 정리한 개념. NOAH 시스템과 비교하며 이해를 높이기 위한 목적.

---

## 1. SO-PO 라인 매칭 전략

### 표준 ERP의 원칙: 1:1 라인 매칭

대부분의 ERP는 SO Line → PO Line을 **1:1로 강제**한다.

| ERP | 방식 |
|-----|------|
| D365 F&O (Intercompany) | SO Line → Intercompany PO Line 자동 생성, 1:1 강제. 합치는 것 불가능 |
| SAP (SD→MM) | SO Line → Purchase Requisition → PO Line, 1:1 기본. Third-party drop ship도 1:1 유지 |

**이유**: 라인 단위로 납기/수량/금액을 추적해야 하므로, 매칭이 깨지면 출고/입고/정산 전부 꼬임.

### 1:1이 안 맞는 케이스의 처리 방법

ERP에서도 라인 불일치 상황은 존재하며, 주로 3가지 패턴으로 처리한다.

#### (1) BOM / Kit Item

여러 품목을 묶어 **하나의 Kit 품목코드**로 등록. SO/PO 라인은 Kit 단위로 1:1 유지하고, 구성품은 BOM 테이블에서 관리.

**필요한 마스터 테이블:**

```
[Item Master] — 품목 마스터
품목코드         | 품목명                    | 유형
NA015           | 전동쿼터턴구동기 NA015      | 단품
BUSH-SQ-22      | 붓싱 별각 22*22           | 단품
NA015-KIT-001   | NA015 + 부싱별각22 세트    | Kit     ← Kit도 품목코드로 관리

[BOM Table] — 구성품 테이블
Kit 품목코드      | 구성품 코드    | 수량
NA015-KIT-001   | NA015         | 1
NA015-KIT-001   | BUSH-SQ-22    | 1
```

**SO/PO 입력 시:**

```
[SO]  Line 1: NA015-KIT-001 × 2  (835,000원)
[PO]  Line 1: NA015-KIT-001 × 2  (808,800원)
→ 1:1 매칭 유지. 구성품 추적은 BOM에서.
```

**장점**: SO/PO 라인 매칭이 항상 깔끔
**단점**: Kit 하나당 품목 마스터 + BOM 등록 필요 → 조합이 다양하면 마스터 폭발

**적합한 경우**: 동일 조합이 반복되는 표준 세트 (예: 구동기+부싱 세트가 자주 나오는 경우)

#### (2) Parent-Child Line (SAP Item Category 방식)

SO 라인에 **하위 품목(sub-item)**을 종속시키는 방식.

```
[SO]
Line 1    NA015 전동쿼터턴구동기    × 2    820,000원    ← Parent (메인)
Line 1.1  붓싱 별각 22*22          × 2     15,000원    ← Child (부속, 메인에 종속)

[PO]
Line 1    NA015 + 부싱가공          × 2    808,800원
→ PO는 Parent 라인(Line 1)과 매칭. Child는 자동 포함.
```

**장점**: 별도 Kit 마스터 불필요, 유연한 조합 가능
**단점**: 라인 번호 체계가 복잡해짐 (1, 1.1, 1.2, 2, 2.1...)

#### (3) Header-Level Linking (주문 단위 연결)

라인 매칭을 포기하고 **SO_ID ↔ PO_ID 단위(주문 헤더)**로만 연결.

```
SO: SOD-2026-0251 (Line 1, 2)
PO: ND-0251       (Line 1)
→ 라인은 무시, SO_ID로만 연결. 금액/수량은 주문 단위로 비교.
```

**장점**: 구현 가장 단순, 입력 자유도 최대
**단점**: 라인별 세밀한 추적 불가 (어떤 SO 라인이 어떤 PO 라인에 대응하는지 알 수 없음)

---

## 2. NOAH 시스템의 현재 위치

### 상황

- D365 CE (CRM) ↔ D365 F&O (ERP)가 **미연동**
- Excel이 두 시스템의 브릿지 역할
- SO/PO 라인 매칭을 시스템이 강제할 수 없음

### 채택한 전략: 1:1 기본 + Header-Level Fallback

```
SO Line ←→ PO Line    (1:1 매칭 시도, line_item 기준)
    ↓ 실패 시
SO_ID  ←→ PO SO_ID    (주문 단위 fallback)
```

**적용 사례:**

| 위치 | 방식 |
|------|------|
| `load_po_detail()` | SO_ID 단위 집계 (Header-Level). 주석: "PO line_item은 SO line_item과 1:1 대응하지 않음" |
| 납기 현황 DN 매칭 | SO_ID + line_item 기준 (1:1 시도) |
| 납기 현황 EXW 보충 | SO_ID 단위로 PO factory_exw fallback (2026-04-01 추가) |

### 왜 BOM/Kit을 안 쓰는가

- 품목 마스터 테이블이 없음 (Excel 기반, ERP 미연동)
- 비정형 조합이 가끔 발생하는 수준 → Kit 마스터 관리 비용 > 이득
- Header-Level Fallback으로 실용적으로 충분

---

## 3. 참고: ERP 정규화 수준 비교

```
[Level 1] 단일 시트 — 78개 컬럼 한 행 (NOAH 이전 상태)
    ↓
[Level 2] 트랜잭션 분리 — SO/PO/DN/PMT 시트 분리 (NOAH 현재)
    ↓
[Level 3] 라인 정규화 — SO_ID + Line item 복합키 (NOAH 현재)
    ↓
[Level 4] 품목 마스터 — Item Master + BOM 테이블 (ERP 표준, NOAH 미도입)
    ↓
[Level 5] 완전 연동 — Intercompany 자동 PO 생성 (D365 F&O 표준)
```

NOAH는 Level 3까지 구현. Level 4(품목 마스터)는 ERP 연동 없이는 관리 부담이 크므로, 코드에서 fallback 로직으로 대응 중.
