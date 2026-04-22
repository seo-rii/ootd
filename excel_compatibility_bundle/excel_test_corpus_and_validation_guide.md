# Excel 테스트 코퍼스와 검증 가이드

이 문서는 “어디서 파일을 모을 것인가”와 “어떻게 정답을 만들 것인가”를 구체화한 문서다.
목표는 fixture 수집을 취미 수준이 아니라 **재현 가능한 테스트 체계**로 만드는 것이다.

---

## 1. 기본 원칙

테스트 파일은 한 종류만으로 충분하지 않다.
최소한 아래 네 부류가 필요하다.

1. **공식 기능 샘플**
   - 특정 기능이 명확히 드러나는 작은 workbook
2. **실무형 과제 파일**
   - 여러 기능이 섞인 중간/대형 workbook
3. **오픈소스 회귀 코퍼스**
   - parser, round-trip, security, regression seed
4. **합성(synthetic) workbook**
   - 원하는 feature matrix를 체계적으로 생성한 파일

---

## 2. 우선 수집할 공식 샘플

### 2.1 Data Validation Examples.xlsx
Microsoft Download Center에서 공식 샘플 workbook을 제공한다.
whole number, decimal, list, date, time, text length, custom validation 예제가 포함되어 있어 data validation 기능 검증에 매우 좋다.

출처  
https://www.microsoft.com/en-us/download/details.aspx?id=53669

### 2.2 Office Scripts sample workbooks
Microsoft Office Scripts 샘플과 시나리오 문서는 각 예제와 함께 sample workbook을 제공하거나 sample workbook 다운로드를 전제로 한다.
테이블, 행/열 표시/숨김, freeze, worksheet 간 링크 등 workbook object model을 실전처럼 다루는 예제가 많다.

출처  
https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/samples-overview

### 2.3 MOS Excel 교육 자료
Microsoft educator materials는 MO-210(Excel Associate)와 MO-211(Excel Expert) 준비용 파일과 활동을 포함한다.
이 자료는 단일 기능 샘플보다 더 실무형이며, 복합 workbook 테스트 seed로 좋다.

출처  
https://learn.microsoft.com/en-us/training/educator-center/instructor-materials/microsoft-365-apps-certification-preparation-materials

---

## 3. 비교 구현 / 회귀 코퍼스

### 3.1 Open XML SDK repository
Open XML SDK는 low-level Open XML/OPC 작업을 위한 프레임워크이며 ISO 29500에 가깝게 동작하도록 설계되어 있다.
repo 안의 sample/test/data 자산은 OOXML 구조 비교와 regression seed로 유용하다.

출처  
https://github.com/dotnet/Open-XML-SDK

### 3.2 Apache POI repository
Apache POI는 Office 바이너리/OOXML/OLE2를 읽고 쓰는 Java 라이브러리다.
Excel 외에도 광범위한 포맷을 다루므로 differential parser/writer 비교에 좋다.

출처  
https://github.com/apache/poi

### 3.3 LibreOffice core repository
LibreOffice core repo의 `sc/qa/unit/data` 계열은 Calc 관련 다양한 회귀 테스트 파일을 포함한다.
이 파일들은 Excel 정답 그 자체는 아니지만, parser robustness, round-trip, fuzz seed로 매우 유용하다.

출처  
https://github.com/LibreOffice/core

---

## 4. synthetic corpus는 직접 만든다

공개 코퍼스만으로는 coverage가 충분하지 않다.
실제 프로젝트에서는 **합성 generator**가 매우 중요하다.

### 4.1 왜 필요한가

공개 workbook은 대체로 다음 문제가 있다.

- 기능 조합이 우연적이다.
- 어떤 feature가 의도적으로 들어갔는지 불명확하다.
- 버전 차이, 저장 옵션, 작성 방식이 섞여 있다.
- diff/reduction이 어렵다.

### 4.2 generator가 만들어야 할 매트릭스 예시

#### sheet/layout
- hidden / very hidden / visible
- freeze panes
- split panes
- page setup / print area / print titles
- merged cells
- outline/grouping

#### cell/value
- blank / string / number / boolean / error
- large/small float
- date serial 1900/1904
- custom number format
- rich text / inline string / shared string

#### formula
- A1 / R1C1
- relative/absolute/mixed ref
- cross-sheet ref
- name ref
- table structured ref
- volatile functions
- dynamic array spill
- implicit intersection
- error-producing formulas

#### workbook metadata
- defined names (global/local)
- calc settings
- theme/style
- workbook protection flags
- external links stub

#### features
- data validation
- conditional formatting
- tables
- hyperlinks
- comments/notes
- images
- charts read-preserve

### 4.3 generator 출력물

하나의 synthetic case는 최소한 다음을 같이 남겨야 한다.

- `input.xlsx`
- `manifest.json`
- `expected-structure.json`
- `excel-oracle-values.json`
- `excel-oracle.pdf`
- 선택적으로 `sdk-validate.json`

---

## 5. fixture 분류 체계

권장 폴더 구조:

```text
fixtures/
  corpus/
    official/
    office-scripts/
    mos/
    openxmlsdk/
    poi/
    libreoffice/
  synthetic/
    values/
    formulas/
    layout/
    validation/
    styles/
    tables/
  golden/
    values/
    render/
    validate/
  manifests/
```

### fixture manifest 권장 필드

```json
{
  "id": "dv_list_basic_001",
  "source": "official",
  "source_ref": "microsoft-download-center:data-validation-examples",
  "format": "xlsx",
  "strictness": "transitional",
  "features": ["data_validation", "drop_down", "style"],
  "requires_excel_oracle": true,
  "requires_pdf_golden": false,
  "risk_tags": ["round_trip", "validation"],
  "notes": "Single sheet validation list example"
}
```

---

## 6. 검증 축 1: 구조 검증

writer가 만든 파일은 먼저 구조적으로 검사해야 한다.
Open XML SDK의 `OpenXmlValidator.Validate(OpenXmlPackage)`는 package 단위 검증을 지원한다.

### 여기서 잡는 것
- schema 위반
- 잘못된 part 관계
- 허용되지 않는 element/attribute
- strict/transitional mismatch 일부

### 여기서 못 잡는 것
- Excel이 실제로 계산을 어떻게 하는지
- layout이 Excel과 동일한지
- unknown ext를 충분히 보존했는지
- Excel이 조용히 복구하는 문제

즉 validator는 필수이지만 충분조건은 아니다.

출처  
https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.validation.openxmlvalidator.validate?view=openxml-3.0.1

---

## 7. 검증 축 2: round-trip 검증

lossless-first 철학에서는 “열고 저장했더니 문서가 망가지지 않았는가”가 핵심이다.

### 검증 방법

1. input workbook load
2. save without logical mutation
3. ZIP/package canonical diff
4. 알려진 volatile 영역만 ignore rule 적용

### ignore rule 예시

- calcChain 재생성
- workbook calc version 변화
- 문서 작성 시간/수정 시간
- relation id 재배치(정말 필요할 때만)

### 반드시 비교할 것

- unknown part 존재 여부
- extLst 보존 여부
- VBA/project bin 보존 여부 (`xlsm`)
- theme/style/drawing rel 보존 여부

---

## 8. 검증 축 3: Excel 오라클 비교

공식 문서상 `Workbooks.Open`으로 프로그램에서 파일을 열면 기본적으로 매크로가 활성화된다.
따라서 오라클 러너는 먼저 `AutomationSecurity`를 설정해야 한다.
또한 `CalculateFullRebuild()`는 모든 open workbook에 대해 full calculation과 dependency rebuild를 강제한다.
오라클 값 추출은 `Value2`가 적합하며, PDF/XPS는 `ExportAsFixedFormat`으로 출력할 수 있다.

### 권장 순서

1. isolated temp dir 준비
2. Excel Application 시작
3. `AutomationSecurity = ForceDisable`
4. alerts / screen updating / prompt 최소화
5. `Workbooks.Open(...)`
6. `Application.CalculateFullRebuild()`
7. 각 sheet의 used range / targeted ranges를 `Value2`로 추출
8. workbook 또는 worksheet PDF export
9. close / quit / cleanup

### 비교 산출물

- values JSON
- formula text JSON
- names JSON
- workbook metadata snapshot
- PDF render snapshot

출처  
https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open  
https://learn.microsoft.com/ko-kr/office/vba/api/excel.application.automationsecurity  
https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild  
https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2  
https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.exportasfixedformat

---

## 9. 검증 축 4: behavior differential

Excel만 보지 말고 비교 구현도 함께 돌리는 것이 좋다.
다만 목적이 다르다.

- **Excel**: behavior oracle
- **Open XML SDK**: 구조/패키지 비교
- **Apache POI**: 다른 구현과의 의미 차이 포착
- **LibreOffice**: robustness 회귀 및 코퍼스 확장

### 실무 팁

동일 fixture에 대해 아래를 동시에 저장해 두면 좋다.

- `our_read.json`
- `excel_oracle.json`
- `poi_read.json`
- `sdk_validate.json`
- `notes.md`

그러면 “우리가 틀린가 / Excel 특수동작인가 / 다른 구현도 다르게 보는가”를 한 번에 판단할 수 있다.

---

## 10. CI 레벨 제안

### PR 레벨
- unit tests
- parser tests
- small official fixtures
- validator tests
- round-trip smoke tests

### nightly
- medium corpus replay
- synthetic matrix generation
- canonical diff suite
- fuzz regression seeds

### gated Windows lane
- Excel oracle tests
- PDF export comparison
- cross-version Excel runner (가능하면 2개 profile)

---

## 11. 위험한 함정

### 함정 1. “공개 workbook 몇 개면 충분하겠지”
아니다. coverage가 매우 편향된다.

### 함정 2. “validator 통과면 Excel에서도 문제없다”
아니다. validator는 구조 검증일 뿐이다.

### 함정 3. “Excel 값만 비교하면 된다”
아니다. 레이아웃/print/export/round-trip도 별도로 검증해야 한다.

### 함정 4. “매크로는 실행 안 할 거니까 무시 가능”
아니다. `xlsm` 보존, macro security, programmatic open 기본 동작을 반드시 고려해야 한다.

### 함정 5. “LibreOffice/POI와 같으면 맞다”
아니다. 그들은 유용한 비교 대상이지만 Excel 정답 대체물이 아니다.

---

## 12. 추천 초기 수집 리스트

### 당장 확보
- Data Validation Examples.xlsx
- Office Scripts sample workbooks 중 table/visibility/freeze 관련 샘플
- MOS Excel Associate / Expert 교육 자료

### 다음 단계 확보
- Open XML SDK test/data
- Apache POI test-data
- LibreOffice `sc/qa/unit/data`

### 직접 생성
- 날짜 시스템 matrix
- dynamic array / spill matrix
- structured ref / table matrix
- validation / conditional formatting matrix
- print layout matrix

---

## 13. 최종 요약

좋은 Excel 테스트 체계는 **파일을 많이 모으는 것**이 아니라, 아래 네 축을 분리하는 것이다.

1. 공식 샘플
2. 실무형 과제 파일
3. 회귀/오픈소스 코퍼스
4. synthetic generator

그리고 정답은 아래 세 방식이 함께 있어야 한다.

- validator 정답
- round-trip 정답
- Excel oracle 정답

이 세 축이 동시에 돌아가야 “읽고 쓸 수 있다”를 넘어서 “Excel 호환성이 있다”고 말할 수 있다.
