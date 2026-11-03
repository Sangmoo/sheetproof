---
title: sheetproof
emoji: 🧾
colorFrom: indigo
colorTo: green
sdk: gradio
sdk_version: 5.50.0
python_version: "3.11"
app_file: app.py
pinned: false
license: mit
short_description: 엑셀 수식을 검증된 Python 코드로 변환
---

# sheetproof — 엑셀 수식을 "검증된 코드"로

수식이 수백 개 얽힌 업무용 엑셀(정산·단가·인센티브 계산 등)을 올리면

1. **셀 의존 관계를 분석·시각화**하고
2. **계산 로직을 Python 코드로 변환**한 뒤
3. **엑셀에 저장된 계산 결과(cached value)와 셀 단위로 대조해 "같은 값이 나온다"는 것을 증명**합니다.

기존 도구는 수식을 *설명*하거나 *변환*하는 데서 멈춥니다. sheetproof 의 중심은 **검증된 변환**입니다.
변환 결과가 원본과 다르면 어느 셀이 다른지, 그 불일치가 어디서 시작됐는지(근본 원인 셀)를 정확히 짚어 줍니다.

- GitHub: <https://github.com/Sangmoo/sheetproof>
- Hugging Face Space: <https://huggingface.co/spaces/Sangmoo/sheetproof>

---

## 기능

| 단계 | 내용 |
|---|---|
| **① 분석** | 시트/수식 통계, 셀 자동 분류(입력 = 수식이 참조하는 상수 / 중간 계산 / 최종 출력), 의존 그래프(SVG + Mermaid 코드), 위상 정렬, 순환 참조·외부 링크·휘발성 함수·매크로·피벗·미지원 함수 경고 |
| **② 변환** | 출력 셀(또는 지정한 범위의 상위 셀만)을 계산하는 Python 모듈 생성. 엑셀 의미론을 재현하는 런타임 `xl_runtime.py` 를 함께 사용 |
| **③ 검증** | 생성된 모듈 소스를 **그대로 실행**해 수식 셀마다 엑셀 저장값과 비교. 일치율, 불일치 셀 목록(기대값/실제값/수식/원인 추정), 근본 원인 셀(★) 표시, 상대/절대 허용 오차 설정 |
| **④ LLM 보완 (선택)** | 지원하지 않는 함수가 들어간 수식만 LLM 이 변환. **본인 API Key 를 입력한 경우에만** 호출하고, 결과는 안전 검사 + 같은 검증 엔진을 통과해야 채택 |
| **⑤ 내보내기** | 모델 모듈(.py) + `xl_runtime.py` + pytest 검증 코드 + HTML 검증 보고서, 그리고 이를 묶은 ZIP |

### 지원 범위

- 파일: `.xlsx`, `.xlsm`(매크로는 무시하고 경고). 최대 10MB
- 참조: 셀, 범위, 전체 열/행(`A:A`, `1:3` → 사용 범위로 자름), 다른 시트(`'My Sheet'!A1`), 이름 정의(통합문서/시트 범위, 참조형·상수형)
- 연산자: `+ - * / ^ & = <> < > <= >=`, 단항 `-`/`+`, 후위 `%`, 배열 상수 `{1,2;3,4}`, 배열 연산(SUMPRODUCT 안의 `(A=B)*(C)` 등)
- 함수(약 50개):
  `SUM SUMIF SUMIFS COUNT COUNTA COUNTIF COUNTIFS AVERAGE AVERAGEIF AVERAGEIFS MIN MAX
  ROUND ROUNDUP ROUNDDOWN INT ABS MOD IF IFS IFERROR AND OR NOT CHOOSE
  VLOOKUP HLOOKUP INDEX MATCH XLOOKUP LEFT RIGHT MID LEN TRIM UPPER LOWER CONCATENATE CONCAT TEXT
  DATE YEAR MONTH DAY EDATE EOMONTH ISBLANK ISNUMBER ISTEXT ISERROR ISNA VALUE SUMPRODUCT`
- 엑셀 의미론 재현
  - 빈 셀: 산술에서는 0, 연결에서는 `""`, 비교에서는 상대 타입의 기본값. `=A1`(빈 셀)은 0
  - 오류값 `#N/A #DIV/0! #VALUE! #REF! #NAME? #NUM!` 은 예외가 아니라 값으로 전파 (왼쪽 피연산자 우선)
  - ROUND 계열은 은행가 반올림이 아닌 "0에서 먼 쪽" 반올림, 유효숫자 15자리 기준 (`ROUND(2.675,2)=2.68`)
  - 날짜: 1900 체계 일련번호, `1900-02-29` 버그(일련번호 60) 재현, DATE 의 월/일 넘침 처리
  - 비교: 텍스트는 대소문자 무시, 타입 순서 숫자 < 텍스트 < 논리값
  - 조건(SUMIF 등): 연산자 접두사, 와일드카드 `* ? ~`, 숫자 조건이 숫자 텍스트와도 일치
  - 집계 함수: 범위 안의 텍스트/논리값은 무시, 직접 인수는 변환 (`SUM(A1)` ≠ `SUM("3")`)
  - TEXT: `0`, `0.00`, `#,##0`, `0%`, `"원"` 리터럴, 양수;음수;0 구역, `yyyy-mm-dd` 등 기본 날짜 서식

### 변환하지 않고 "고정값"으로 처리하는 수식 (검증 제외)

| 경우 | 처리 |
|---|---|
| 휘발성 함수 `NOW TODAY RAND RANDBETWEEN OFFSET INDIRECT CELL INFO` | 엑셀 저장값으로 고정, 검증 제외 |
| 미지원 함수 | 저장값으로 고정 (LLM 보완 대상) |
| 동적 배열/LAMBDA (`LET LAMBDA FILTER SORT UNIQUE SEQUENCE …`), 여러 셀 배열 수식 | 저장값으로 고정 |
| 외부 통합문서 참조, 3D 참조, 정의되지 않은 이름 | 저장값으로 고정 |
| 순환 참조(와 그에 의존하는 셀) | 저장값으로 고정, 순환 경로 경고 |
| 저장값이 `#NAME?` (저장한 Excel 버전이 함수를 모름, 예: Excel 2019 의 XLOOKUP) | 저장값으로 고정 |

고정된 셀에 **의존하는** 다른 수식은 고정값을 입력으로 받아 정상적으로 변환·검증됩니다.

---

## 사용법 (HF Space)

1. `.xlsx` 파일을 올리거나 샘플을 선택합니다.
2. (선택) **대상 셀 범위**에 `정산!H2:H8, 주문!F9` 처럼 적으면 그 셀과 상위 셀만 변환합니다.
3. (선택) 허용 오차를 조정합니다. 기본값은 상대·절대 모두 `1e-9`.
4. **분석 · 변환 · 검증 실행**을 누릅니다.
5. 상단 카드에서 일치율을 확인하고, 탭에서 분석/코드/검증 결과를 봅니다.
6. 미지원 함수가 있으면 **④ LLM 보완** 탭에서 공급자·모델·API Key 를 넣고 실행합니다.
7. **⑤ 내보내기**에서 ZIP 을 받습니다.

> **cached value 가 없는 파일**: openpyxl 등 라이브러리로 만든 파일에는 계산 결과가 저장되어 있지 않아 검증할 수 없습니다.
> Excel(또는 LibreOffice)에서 한 번 열어 저장한 뒤 다시 올려 주세요.

### 내보낸 코드 사용 예

```python
import settlement_model as model

values = model.compute()                     # 엑셀과 같은 입력으로 계산
print(values["정산!H8"])

what_if = model.compute({"정산!B2": 1000})   # 입력 셀 값을 바꿔 시뮬레이션
print(what_if["정산!H8"])
```

```bash
pytest test_settlement_model.py   # 엑셀 저장값과 셀 단위 대조 재현
```

---

## 로컬 실행

Python 3.11 이상.

```bash
git clone https://github.com/Sangmoo/sheetproof.git
```

```bash
cd sheetproof
```

```bash
pip install -r requirements.txt pytest
```

```bash
python app.py
```

브라우저에서 <http://localhost:7860> 을 엽니다.

라이브러리로만 쓰려면:

```python
import sys; sys.path.insert(0, "src")
from sheetproof import run, export_bundle

analysis, gen, report = run("samples/settlement.xlsx", targets_text="", rel_tol=1e-9, abs_tol=1e-9)
print(f"{report.match_rate:.2%}", report.count("mismatch"))
export_bundle(analysis, gen, report, "out/")
```

## 테스트

```bash
python -m pytest
```

| 파일 | 내용 |
|---|---|
| `tests/test_parser.py` | 수식 → AST → 수식 왕복, 연산자 우선순위(`-2^2=4`, `^` 왼쪽 결합), 시트/이름/외부 참조, 배열 상수, 파싱 오류 |
| `tests/test_runtime.py` | 함수별 엑셀 의미론: 빈 셀, 오류 전파, ROUND, 날짜(1900 버그), 조건/와일드카드, 조회, TEXT 서식, 비교 |
| `tests/test_samples.py` | **실제 Excel 로 재계산·저장한 샘플 5개**를 분석→변환→검증해 100% 일치 확인, 분류/위상정렬/순환/대상 범위/입력 변경 |
| `tests/test_verify_export.py` | **일부러 틀린 코드를 주입**해 불일치와 근본 원인 셀을 정확히 잡는지, 허용 오차, 내보낸 pytest 를 별도 프로세스에서 실행 |
| `tests/test_llm.py` | 가짜 Anthropic/OpenAI SDK 로 호출 흐름, 셀 값 미전송, 안전 검사, 검증 실패 시 미채택 |
| `tests/test_app.py` | Gradio 핸들러 스모크 테스트, 업로드 임시 파일 삭제, API 노출 차단 |

### 샘플 다시 만들기

`samples/*.xlsx` 는 `scripts/make_samples.py` 로 수식을 쓰고 **실제 Excel 로 재계산해 저장**한 파일입니다.

```bash
python scripts/make_samples.py
```

```bash
powershell -ExecutionPolicy Bypass -File scripts/excel_recalc.ps1
```

Excel 이 없다면 LibreOffice 로도 재계산할 수 있습니다 (결과가 Excel 과 미세하게 다를 수 있음).

```bash
soffice --headless --convert-to xlsx --outdir samples samples/_raw/*.xlsx
```

| 샘플 | 내용 |
|---|---|
| `settlement.xlsx` | 단순 정산표: 수량×단가, 할인, 부가세(이름 정의), ROUND 계열, 연산자 우선순위 |
| `price_lookup.xlsx` | 단가표 조회: VLOOKUP 정확/근사/와일드카드, HLOOKUP, INDEX/MATCH, XLOOKUP, 시트 간 참조 |
| `sumifs_summary.xlsx` | SUMIF(S)/COUNTIF(S)/AVERAGEIF(S)/SUMPRODUCT, 날짜 조건, 와일드카드 |
| `dates.xlsx` | DATE/EDATE/EOMONTH/YEAR/MONTH/DAY/TEXT, 1900-02-29 버그, 월말 처리 |
| `errors_text.xlsx` | 오류값 전파, 텍스트/논리 함수, 휘발성(TODAY)·미지원(PROPER, REPT) 함수 |
| `no_cache.xlsx` | cached value 가 없는 파일 (경고 확인용) |

> 저장소의 샘플은 Excel 2019 로 저장해서 XLOOKUP 셀에 `#NAME?` 이 저장되어 있습니다. sheetproof 는 이를 감지해
> 고정값으로 처리합니다. XLOOKUP 의미론은 `tests/test_runtime.py` 에서 따로 검증합니다.

---

## Hugging Face Space 등록 방법

1. <https://huggingface.co/new-space> 에서 Space 를 만듭니다.
   - SDK: **Gradio**, 하드웨어: **CPU basic** (무료)
   - ZeroGPU 를 고르면 `@spaces.GPU` 함수가 하나도 없을 때 기동에 실패하므로 CPU basic 을 쓰세요.
2. 이 README 맨 위의 YAML front matter 가 Space 설정입니다. `sdk: gradio`, `sdk_version: 5.50.0`,
   `python_version: "3.11"`, `app_file: app.py` 를 그대로 둡니다.
3. **Access Token(write 권한)** 을 <https://huggingface.co/settings/tokens> 에서 발급합니다.
   git push 시 비밀번호 대신 이 토큰을 씁니다. **토큰을 채팅·이슈·커밋에 붙여 넣지 마세요.**
4. Space 저장소를 원격으로 추가하고 푸시합니다.

   ```bash
   git remote add hf https://huggingface.co/spaces/Sangmoo/sheetproof
   ```

   ```bash
   git push hf main
   ```

   사용자 이름은 HF 아이디, 비밀번호 자리에 토큰을 넣습니다 (Git Credential Manager 가 저장).
5. `samples/*.xlsx` 같은 바이너리 파일은 HF 가 Git LFS(Xet)로만 받습니다. 이 저장소의 `.gitattributes` 에
   `*.xlsx filter=lfs` 가 있으므로 `git lfs install` 만 해 두면 됩니다.
6. Space 의 **Logs** 탭에서 빌드가 끝나고 `Running on local URL` 이 보이면 완료입니다.

GitHub 와 Space 를 같이 쓰려면 원격을 두 개 두고 둘 다 푸시합니다.

```bash
git push origin main
```

```bash
git push hf main
```

---

## 구조

```
sheetproof/
├── app.py                      # HF Space 진입점 (Gradio UI, 모든 이벤트 api_name=False)
├── requirements.txt
├── pyproject.toml
├── src/sheetproof/
│   ├── workbook.py             # openpyxl 로 두 번 열기 (수식 / cached value), 날짜→일련번호, 매크로·외부링크 감지
│   ├── parser.py               # Tokenizer 기반 Pratt 파서 → AST, AST → 수식 문자열
│   ├── analysis.py             # 참조 해석(이름·시트·전체 열), 의존 그래프, 분류, 위상 정렬, 순환 탐지(Tarjan), 경고
│   ├── codegen.py              # AST → Python 식, 모듈 소스 생성 (INPUTS / FIXED / FORMULAS / compute)
│   ├── xl_runtime.py           # 엑셀 의미론 런타임 (표준 라이브러리만 사용, 그대로 내보냄)
│   ├── verify.py               # 생성 모듈 실행 → 셀 단위 비교, 원인 추정, 근본 원인 셀
│   ├── viz.py                  # 의존 그래프 SVG / Mermaid
│   ├── llm.py                  # Anthropic / OpenAI / OpenAI 호환, AST 화이트리스트 안전 검사, 검증 후 채택
│   └── export.py               # .py + xl_runtime.py + pytest + HTML 보고서 + ZIP
├── scripts/
│   ├── make_samples.py         # 샘플 수식 작성
│   └── excel_recalc.ps1        # Excel COM 으로 재계산·저장
├── samples/                    # 실제 Excel 로 재계산한 샘플
└── tests/
```

### 동작 흐름

```
xlsx ─┬─ openpyxl(data_only=False) ─→ 수식 ──→ parser ──→ AST ──→ analysis ──→ 의존 그래프 / 분류 / 순서
      └─ openpyxl(data_only=True)  ─→ 엑셀 저장값 ─────────────────────────────┐
                                                                               ▼
                    codegen ──→ Python 모듈 소스 ──→ exec ──→ compute() ──→ verify (셀 단위 대조)
                                                                               │
                         LLM (미지원 수식만, Key 있을 때) ──→ 안전 검사 ──→ verify 통과 시 채택
```

**검증의 핵심**: 검증 엔진은 내보낼 모듈 소스를 그대로 실행합니다. 즉 사용자가 내려받는 코드 = 검증된 코드입니다.
각 셀은 상위 셀의 *계산값*으로부터 끝까지(end-to-end) 계산되며, 불일치가 생기면 위상 순서로 따라가며
"상위 셀이 모두 일치하는데 혼자 다른 셀"을 근본 원인(★)으로 표시합니다.

### LLM 사용 시 데이터 범위

- 전송: 대상 셀 주소, 수식 문자열, 참조 셀의 주소와 타입(숫자/텍스트/날짜/빈 셀/범위 크기)
- 전송하지 않음: 셀 값, 파일 이름, 다른 시트 내용
- 기본 모델: Anthropic `claude-opus-5-5`, OpenAI `gpt-4.1` (화면에서 변경 가능). OpenAI 호환은 Base URL 필요
- LLM 이 돌려준 식은 `ast` 화이트리스트(임포트·람다·dunder·임의 속성 접근 금지)를 통과하고
  `xl.guard()` 로 감싸 실행되며, 해당 셀이 엑셀 저장값과 일치할 때만 채택됩니다.

### 개인정보

- 업로드 파일은 읽은 직후 임시 파일을 삭제하고 서버에 저장하지 않습니다.
- 내보내기 파일은 임시 폴더에 만들어지며 Gradio 캐시 정리(`delete_cache`)로 1시간 이내 삭제됩니다.
- API Key 는 요청 처리에만 쓰이고 저장되지 않습니다.

---

## 한계

- **cached value 가 필요합니다.** 엑셀이 저장한 계산 결과가 없으면 검증할 수 없습니다.
- 동적 배열(스필), LAMBDA/LET, 여러 셀 배열 수식, 암시적 교차(`@`), 참조 연산자(`:` 함수 결과 범위, 공백 교차, 3D 참조)는 지원하지 않습니다.
- 함수 지원은 위 목록으로 한정됩니다. 그 외는 고정값 처리 또는 LLM 보완이 필요합니다.
- TEXT 는 기본 서식만 지원합니다 (조건부 서식 `[>100]`, 색상 `[Red]`, 분수, 지수 서식 제외).
- 근사 일치 조회(VLOOKUP TRUE, MATCH 1/-1)는 정렬된 데이터를 가정합니다. 정렬되지 않은 데이터에서 엑셀의 이진 탐색 결과와 다를 수 있으며, 이 경우 검증에서 불일치로 드러납니다.
- 1904 날짜 체계 통합문서는 지원하지 않습니다.
- 정밀도: 엑셀은 유효숫자 15자리이므로 아주 큰 수나 누적 오차에서 미세한 차이가 날 수 있습니다. 허용 오차로 조정하세요.
- 수만 개 이상의 수식이 있는 통합문서는 느릴 수 있고, 그래프는 150개 노드까지만 그립니다 (대상 범위로 좁혀 보세요).
- SQL 변환, "입력값 바꿔서 시뮬레이션" 화면, 엑셀 간 로직 비교는 다음 단계 과제입니다.

## 라이선스

MIT
