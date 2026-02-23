# NOAH 문서 자동 생성기

RCK(Rotork Korea Sales Office)에서 NOAH(Intercompany Factory)로 보내는 업무 문서를 자동 생성합니다.

```
NOAH_SO_PO_DN.xlsx (데이터 소스)
       │
       ├── 국내 시트 ─→ PO (발주서), 거래명세표
       │
       └── 해외 시트 ─→ PO (발주서), Proforma Invoice
```

| 문서 | 용도 | CLI |
|------|------|-----|
| **PO** (Purchase Order) | RCK→NOAH 발주서 | `create_po.py` |
| **거래명세표** | 국내 납품/선수금 명세 | `create_ts.py` |
| **PI** (Proforma Invoice) | 해외 견적서 | `create_pi.py` |

---

## 사용 방법

### 1단계: 주문 정보 입력

`NOAH_SO_PO_DN.xlsx` 파일에 주문 정보를 입력합니다.

| 시트 | 용도 |
|------|------|
| 국내 | 국내 고객 주문 |
| 해외 | 해외 고객 주문 |

**필수 입력 항목:**
- `RCK Order no.` - 주문번호 (예: ND-0001, NO-0001)
- `Customer name` - 고객명
- `Customer PO` - 고객 발주번호
- `Item qty` - 수량
- `Model` - 모델명
- `ICO Unit` - 단가

### 2단계: 문서 생성

`create_po.bat` 파일을 더블클릭하면 대화형 메뉴가 나옵니다.

```
==============================
 NOAH 문서 생성기
==============================
[1] 발주서(PO) 생성
[2] 거래명세표 생성
[3] Proforma Invoice 생성
```

또는 명령 프롬프트에서 직접 실행:

#### PO (발주서)
```bash
# 단일 발주
python create_po.py ND-0001

# 여러 건 동시 생성
python create_po.py ND-0001 ND-0002 ND-0003

# 강제 생성 (중복 및 검증 오류 무시)
python create_po.py ND-0001 --force

# 이력 조회 (현재 월)
python create_po.py --history

# 이력을 Excel로 내보내기
python create_po.py --history --export
```

#### 거래명세표
```bash
# 단건 생성
python create_ts.py DND-2026-0001

# 월합 거래명세표 (여러 DN을 한 장으로)
python create_ts.py DND-2026-0001 DND-2026-0002 --merge

# 대화형 모드 (DN 목록 붙여넣기)
python create_ts.py --interactive --merge
```

#### Proforma Invoice
```bash
python create_pi.py NO-0001
```

### 3단계: 결과 확인

생성된 문서는 각 폴더에 저장됩니다.

```
generated_po/   ← 발주서
generated_ts/   ← 거래명세표
generated_pi/   ← Proforma Invoice
```

---

## 자동 검증 항목

발주서 생성 시 다음 항목을 자동으로 검증합니다:

| 검증 항목 | 조건 | 결과 |
|----------|------|------|
| 필수 필드 | 비어있음 | 오류 |
| ICO Unit | 0 또는 음수 | 오류 |
| 납기일 | 과거 날짜 | 오류 |
| 납기일 | 7일 이내 | 경고 |
| 중복 발주 | 이미 생성된 주문번호 | 경고 (확인 후 진행 가능) |

오류가 있으면 진행 여부를 확인합니다. 강제로 생성하려면:

```bash
python create_po.py ND-0001 --force
```

---

## 파일 구조

```
noahAutomation/
├── NOAH_SO_PO_DN.xlsx       ← 데이터 소스 (상위 폴더에 위치)
├── create_po.bat            ← 더블클릭으로 실행 (대화형 메뉴)
├── create_po.py             ← PO CLI
├── create_ts.py             ← 거래명세표 CLI
├── create_pi.py             ← Proforma Invoice CLI
├── po_generator/            ← 핵심 패키지
├── templates/               ← 문서 템플릿 (Excel)
│   ├── purchase_order.xlsx
│   ├── transaction_statement.xlsx
│   └── proforma_invoice.xlsx
├── generated_po/            ← 생성된 발주서
├── generated_ts/            ← 생성된 거래명세표
├── generated_pi/            ← 생성된 Proforma Invoice
└── po_history/              ← 발주 이력 (월별)
    └── YYYY/M월/
        └── YYYYMMDD_주문번호_고객명.xlsx
```

---

## 다중 아이템 주문

동일한 `RCK Order no.`로 여러 행을 입력하면 하나의 발주서에 여러 아이템이 포함됩니다.

**NOAH_SO_PO_DN.xlsx 예시:**

| RCK Order no. | Customer name | Item name | Item qty | ICO Unit |
|---------------|---------------|-----------|----------|----------|
| ND-0001 | ABC전자 | NA-100 | 2 | 500,000 |
| ND-0001 | ABC전자 | NA-200 | 1 | 750,000 |

→ `ND-0001` 발주서에 2개 아이템이 포함됨

---

## 설치

### 1. Miniconda 설치
- https://docs.conda.io/en/latest/miniconda.html 에서 Windows 64-bit 버전 설치

### 2. 환경 생성
```bash
conda create -n po-automate python=3.11
conda activate po-automate
pip install -r requirements.txt
```

### 3. 설정 파일 생성
- `user_settings.example.py` → `user_settings.py`로 복사 후 본인 경로 수정
- `local_config.example.bat` → `local_config.bat`으로 복사 후 본인 Python 경로 수정

---

## 문제 해결

### "주문번호를 찾을 수 없습니다"
- `NOAH_SO_PO_DN.xlsx`에 해당 주문번호가 있는지 확인
- 주문번호 앞뒤 공백 확인
- 국내/해외 시트 모두 확인

### "ICO Unit이 0입니다"
- 단가가 입력되었는지 확인
- 숫자 형식인지 확인 (텍스트로 입력되면 인식 안 됨)

### "이미 발주된 건입니다"
- `po_history/` 폴더에서 이전 발주 기록 확인
- 현재 월 이력: `python create_po.py --history`
- 재발주가 필요하면 Y 입력하여 진행
