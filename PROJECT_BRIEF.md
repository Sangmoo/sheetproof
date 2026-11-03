# sheetproof — 엑셀 수식을 "검증된 코드"로 바꾸는 도구

## 0. 대화 규칙
- 나에게 답할 때는 **한국어로만** 답할 것.
- 바로 구현하지 말고, 설계·범위가 애매하면 먼저 제안하고 확인받을 것.
- 결과물은 GitHub(`https://github.com/Sangmoo/sheetproof.git`)과
  Hugging Face Space(`https://huggingface.co/spaces/Sangmoo/sheetproof`)에 올린다.
  (저장소와 Space는 내가 먼저 만들고 알려줄 예정)
- README.md는 기능, 사용법, 로컬 실행, 테스트, **HF Space 등록 방법**, 구조, 한계까지 자세하게 작성한다.
- 검증은 테스트 작성 → 실행 → 브라우저 확인까지 마친다.

## 1. 한 줄 정의
수식이 수백 개 얽힌 업무용 엑셀(정산·단가·인센티브 계산 등)을 올리면
**① 셀 의존 관계를 시각화**하고 **② 계산 로직을 Python/SQL 코드로 변환**한 뒤
**③ 엑셀에 저장된 계산 결과와 셀 단위로 대조해 "같은 값이 나온다"는 것을 증명**한다.

## 2. 왜 만드는가 (차별점)
- 회사마다 "아무도 못 건드리는 엑셀"이 있다. 만든 사람은 퇴사했고 로직은 수식 안에만 있다.
- 기존 도구는 수식을 **설명**하거나 **변환**까지만 한다. 변환 결과가 원본과 같은지 **증명하는 도구는 거의 없다.**
- 핵심 가치는 "변환"이 아니라 **"검증된 변환"**이다. 불일치 셀을 정확히 짚어 주는 것이 제품의 중심이다.

## 3. 사용 흐름 (HF Space)
1. `.xlsx` 업로드 (최대 10MB 정도)
2. **분석**: 시트/수식 통계, 입력 셀(상수)·중간 계산·최종 출력 셀 자동 분류, 의존 그래프, 순환 참조·외부 링크·휘발성 함수 경고
3. **변환**: 출력 셀(또는 사용자가 고른 범위)을 계산하는 Python 함수 생성
   - 1차: 규칙 기반 변환 (지원 함수 목록 내)
   - 2차(선택): 지원하지 않는 수식만 LLM이 변환 — **사용자가 본인 API Key를 입력한 경우에만** 호출
4. **검증**: 엑셀에 저장된 계산 결과(cached value)와 생성 코드 실행 결과를 셀 단위로 비교
   → 일치율, 불일치 셀 목록(기대값/실제값/수식/원인 추정), 허용 오차 설정
5. **내보내기**: Python 모듈(.py) + pytest 검증 코드 + 검증 보고서(HTML) 다운로드. 선택적으로 SQL 버전.

## 4. MVP 범위
**포함**
- `.xlsx` / `.xlsm`(매크로는 무시하고 경고) 읽기
- 셀 참조, 범위, 다른 시트 참조, 이름 정의(Defined Names)
- 연산자: 사칙, `^`, `&`, 비교, 단항 `-`, `%`
- 함수 1차 지원(약 40개):
  SUM, SUMIF(S), COUNT(A/IF/IFS), AVERAGE(IF/IFS), MIN, MAX, ROUND/ROUNDUP/ROUNDDOWN,
  INT, ABS, MOD, IF, IFS, IFERROR, AND, OR, NOT, VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP,
  LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, CONCATENATE/CONCAT, TEXT(기본 서식),
  DATE, YEAR, MONTH, DAY, EDATE, EOMONTH, ISBLANK, ISNUMBER, CHOOSE, SUMPRODUCT
- 엑셀 의미론 재현: 빈 셀 = 0/"" 문맥 처리, 오류값(#N/A, #DIV/0!, #VALUE!, #REF!) 전파,
  ROUND의 반올림 방식(은행가 반올림 아님), 날짜 일련번호(1900 체계, 1900-02-29 버그), 부동소수 오차 허용

**제외(경고만)**
- VBA 매크로, 피벗 테이블, 외부 통합문서 링크, 동적 배열/LAMBDA/LET, 구형 `.xls`
- 휘발성 함수(NOW, TODAY, RAND, OFFSET, INDIRECT): 검증 대상에서 제외하거나 고정값 처리

## 5. 기술 구성 (초안, 첫 대화에서 확정)
- Python 3.11, Gradio 5.50.0 (HF Space 진입점 `app.py`), 핵심 로직은 `src/sheetproof/` 패키지
- 읽기: `openpyxl` 로 두 번 연다 — `data_only=False`(수식), `data_only=True`(엑셀이 저장한 계산 결과)
- 수식 파싱: `openpyxl.formula.Tokenizer` 기반 자체 파서 → AST
  (`formulas`, `pycel` 라이브러리는 참고/비교용으로 검토)
- 의존 그래프: AST에서 참조 추출 → DAG, 위상 정렬, 순환 탐지. 시각화는 SVG(또는 Mermaid)
- 코드 생성: AST → Python 코드. 런타임 헬퍼(`xl_runtime.py`: 엑셀 의미론 함수 구현)를 함께 내보냄
- 검증 엔진: 입력 셀 값을 넣고 생성 코드를 실행 → cached value 와 비교
- LLM(선택): Anthropic(기본 `claude-opus-5-5`), OpenAI, OpenAI 호환. 지원 외 수식만 보내고,
  **LLM 결과도 반드시 같은 검증 엔진을 통과해야** 채택. 셀 값(데이터)은 기본적으로 LLM에 보내지 않음

## 6. 반드시 주의할 함정
- **openpyxl로 만든 파일에는 계산 결과(cached value)가 없다.** 테스트 샘플은
  ① 실제 Excel에서 저장한 파일을 쓰거나 ② LibreOffice headless로 재계산해 저장해야 한다.
  cached value 가 없는 업로드는 "엑셀에서 한 번 열어 저장 후 다시 올려 주세요" 안내.
- 엑셀 날짜는 숫자다. 서식 정보로 날짜 여부를 판단해야 한다.
- 공유 수식(shared formula)과 배열 수식은 openpyxl 에서 표현이 다르다.
- 정밀도: 엑셀은 유효숫자 15자리. 비교 시 상대/절대 허용 오차를 둔다.
- 업로드 파일에 개인정보가 있을 수 있음 → 서버에 저장하지 않고, 처리 후 임시파일 삭제, LLM 전송 범위 명시.

## 7. 마일스톤
1. **M1 분석**: 업로드 → 시트/수식 통계, 셀 분류, 의존 그래프, 경고
2. **M2 변환**: 지원 함수 범위 내 Python 코드 생성 + 런타임 헬퍼
3. **M3 검증**: 셀 단위 대조, 일치율, 불일치 리포트 ← **제품의 핵심, 여기까지가 MVP**
4. **M4 LLM 보완**: 미지원 수식만 LLM 변환 → 검증 통과 시 채택
5. **M5 내보내기/배포**: .py + pytest + HTML 보고서, README, GitHub/HF 배포
6. (이후) SQL 변환, "입력값 바꿔서 시뮬레이션" 화면, 엑셀 간 로직 차이 비교

## 8. 테스트 계획
- 실제 Excel 저장 샘플 3~5개: 단순 정산표, VLOOKUP 단가표, SUMIFS 집계, 날짜 계산, 오류값 포함 시트
- 함수별 엑셀 의미론 단위 테스트(빈 셀, 오류 전파, ROUND, 날짜)
- 파서: 다양한 수식 문자열 → AST 왕복 테스트
- 검증 엔진: 일부러 틀린 변환을 넣어 불일치를 잡는지 확인
- 앱 핸들러 스모크 테스트, LLM은 가짜 SDK로 테스트

## 9. 참고: 이전 프로젝트(xplanlens)에서 얻은 교훈
- HF Space 하드웨어가 ZeroGPU 면 `@spaces.GPU` 함수가 하나도 없을 때 기동 실패
  → 빈 placeholder 함수를 두거나 하드웨어를 CPU basic 으로 설정
- HF push 는 비밀번호가 아니라 **Access Token(write)** 으로 인증. 토큰은 채팅에 붙여넣지 말 것
- README 상단에 HF YAML front matter(`sdk: gradio`, `sdk_version`, `app_file: app.py` 등) 필요
- LLM 기능은 Key 미입력 시 호출하지 않고 안내 문구만 반환, Gradio 이벤트는 `api_name=False` 로 API 노출 차단

## 10. 첫 대화에서 할 일
1. 위 기술 구성과 MVP 범위를 검토하고 수정 제안
2. 프로젝트 구조와 모듈 설계안 제시
3. 내가 확인하면 M1부터 구현 시작