# Excel 엔진 원리와 구현 시 주의점

이 문서는 `excel_compatibility_architecture.md`와 `excel_project_structure_and_interfaces.md`를 보완하는 **원리 문서**다.
목표는 “어디를 먼저 구현해야 하는가”가 아니라, **왜 그 구조가 필요한가**를 Excel 동작 원리에 맞춰 설명하는 것이다.

---

## 1. 가장 중요한 전제: Excel은 파일 포맷보다 실행 의미론이 더 어렵다

Excel 호환 라이브러리의 핵심 난도는 XML 태그를 읽고 쓰는 일이 아니다.
실제 난도는 다음 네 가지에 있다.

1. **계산 의미론**
   - 어떤 셀이 언제 다시 계산되는가
   - 동일한 수식이라도 `Formula` / `Formula2`, 동적 배열 지원 여부, implicit intersection 규칙에 따라 결과가 달라질 수 있다.

2. **Object Model 의미론**
   - `Range`는 단순한 직사각형이 아니라, 하나 이상의 contiguous area나 3D range까지 표현한다.
   - `Selection`, `ActiveCell`, `UsedRange`, `CurrentRegion` 같은 개념은 파일 저장 구조가 아니라 런타임 해석 결과다.

3. **호환성 의미론**
   - 같은 `.xlsx`라도 Strict / Transitional, 확장 파트, `extLst`, `AlternateContent`, future function, `_xlfn`, locale 차이, 버전별 함수 지원 범위가 다르다.

4. **Excel 오라클 의미론**
   - 실제 Excel이 버그를 포함한 역사적 동작을 유지한다.
   - 대표적으로 1900 leap-year bug는 “논리적으로 잘못”되었지만 “호환성 측면에서 정답”인 동작이다.

이 프로젝트에서 Object Model을 중심에 두는 이유는 맞지만, 그 Object Model이 곧 내부 저장 구조라고 가정하면 실패한다.
Object Model은 **public façade**이고, 실제 의미론은 runtime + calc + model에 있어야 한다.

---

## 2. 계산 엔진 원리

Microsoft 문서에 따르면 Excel recalculation은 크게 세 단계로 이해할 수 있다.

1. **dependency tree 구성**
2. **calculation chain 구성**
3. **recalculate**

이때 calculation chain은 “수식 셀의 계산 순서”이지, dependency 자체가 아니다.
또한 recalculation 과정에서 Excel은 더 나은 순서를 찾으면 chain을 수정한다.

### 구현 시 결론

- `calcChain.xml`을 정답으로 보면 안 된다.
- `calcChain.xml`은 **힌트**로만 사용한다.
- 진짜 의존성 그래프는 수식 파싱 결과에서 다시 만들어야 한다.
- `dirty propagation`, `partial recalc`, `full rebuild`는 엔진의 핵심 기능으로 따로 가져가야 한다.

### 권장 내부 구성

- `excel-formula`
  - lexer/parser
  - A1/R1C1 reference parser
  - structured reference parser
  - name binding 전 단계 AST
- `excel-calc`
  - binding / dependency graph
  - calc scheduler
  - evaluator
  - volatile function registry
  - recalc mode / dirty set 관리

---

## 3. Formula, Formula2, 동적 배열

Microsoft는 `Range.Formula`와 `Range.Formula2`를 서로 다른 formula dialect로 설명한다.
동적 배열 이전에는 implicit intersection evaluation(IIE)이 기본이었고, 동적 배열 이후에는 array evaluation(AE)이 셀 수식 표면까지 올라온다.

### 구현 시 결론

라이브러리 내부에서 최소한 아래 세 층을 분리하는 편이 안전하다.

1. **원본 표면 문자열**
   - 사용자가 입력한 수식 텍스트
   - round-trip을 위해 필요

2. **정규화 AST**
   - `@`, spill, `_xlfn`, structured ref, locale 문제를 정리한 중간 표현

3. **호환성 프로필 기반 출력 dialect**
   - Excel 2016/2019/365 등 profile에 맞춰 Formula/Formula2로 다시 직렬화

### 왜 필요한가

- 동일한 논리식이 profile에 따라 다른 문자열로 저장될 수 있다.
- dynamic array 미지원 프로필에서는 future function / `_xlfn` 처리가 달라진다.
- `@` implicit intersection을 AST에서 독립 토큰으로 보지 않으면 이후 변환이 불안정해진다.

---

## 4. 날짜, 숫자, 오류값은 “그럴듯한 타입”보다 “Excel식 표현”이 중요하다

### 4.1 1900/1904 날짜 시스템

Excel은 1900 날짜 시스템과 1904 날짜 시스템을 모두 가진다.
또한 1900 시스템에서는 역사적 호환성을 위해 1900년을 윤년으로 잘못 취급한다.

### 구현 시 결론

- 날짜는 내부적으로 `ExcelSerialDate { system, serial }`처럼 보존하는 계층이 있어야 한다.
- 일반 `chrono::DateTime`이나 언어 기본 날짜형으로 곧장 바꾸면 안 된다.
- display/interop 계층에서만 변환한다.

### 4.2 Value vs Value2

Microsoft 문서상 `Value2`는 `Value`와 달리 Currency, Date 타입을 별도로 사용하지 않는다.
실제 오라클 테스트에서는 이 차이가 매우 중요하다.

### 구현 시 결론

오라클 덤프나 비교용 추출은 기본적으로 `Value2`를 사용한다.
내부 값 타입도 아래처럼 계층을 나누는 편이 좋다.

- **storage value**: 셀에 저장된 본래 상태
- **calc value**: evaluator가 다루는 값
- **interop value**: Object Model/JS/Rust API에 노출되는 값

즉, `CellValue`, `CalcValue`, `OmValue`를 분리해야 한다.

### 4.3 숫자와 정밀도

Excel은 IEEE 754 배정밀도 기반 수치 연산과 Excel 고유의 표시/반올림/에러 의미론이 섞여 있다.
따라서 수치 엔진은 “일반적인 부동소수 계산”과 “Excel 표시/비교/직렬화 규칙”을 분리해야 한다.

---

## 5. Range는 저장 객체가 아니라 view여야 한다

Microsoft 문서상 `Range`는 다음을 모두 표현할 수 있다.

- 단일 셀
- 행
- 열
- 하나 이상의 contiguous block
- 3D range

따라서 `Range`를 단일 직사각형으로 가정하면 초반에는 편해 보여도 금방 막힌다.

### 권장 표현

```rust
pub struct RangeRef {
    pub workbook: WorkbookId,
    pub areas: SmallVec<[RectRef; 1]>,
    pub sheet_scope: SheetScope,
}
```

여기서 `SheetScope`는 단일 시트 / 다중 시트 / 3D 범위를 구분한다.

### 구현 시 결론

- `Range`는 façade
- 셀 저장소는 sparse/chunked storage
- range 연산은 view + iterator + materializer로 설계
- `Areas`, `Rows`, `Columns`, `Cells`는 lazy view가 낫다

---

## 6. 파일 포맷 계층 원리: lossless-first

OOXML을 “아는 태그만 파싱해서 model로 만들고 다시 쓰는 구조”로 시작하면,
초기 구현이 간단해 보이지만 실제 문서를 쉽게 깨뜨린다.

Office는 표준 SpreadsheetML 외에도 확장 관계, future record, `extLst`, drawing ext, unknown relationship를 많이 사용한다.
`MS-OI29500`은 Office가 ISO/IEC 29500를 구현하면서 어디를 확장·변형하는지 설명하는 중요한 자료다.

### 구현 시 결론

loader는 두 결과를 동시에 만들어야 한다.

1. **typed model**
   - 현재 구현이 이해하는 필드

2. **opaque preservation layer**
   - 아직 이해하지 못한 element/attribute/part/relationship/ext

writer는 typed model만이 아니라 opaque layer도 함께 다시 방출해야 한다.

### 실전 규칙

- unknown XML attribute는 가능한 한 보존
- unknown part와 rel은 URI/content-type/relationship id까지 통째로 보존
- `mc:AlternateContent`는 무조건 삭제하지 말고 preserve 우선
- style/theme/drawing은 초기에 read-preserve부터 시작

---

## 7. Strict vs Transitional

지원 문서상 Excel Workbook `.xlsx`와 Strict Open XML Spreadsheet는 모두 `.xlsx` 확장자를 사용하지만 의미가 다르다.
Strict는 ISO strict profile이고, 실제 Excel/실무 파일은 Transitional 및 Office 확장을 많이 사용한다.

### 구현 시 결론

- v0.1은 Transitional write 우선
- Strict는 read 지원부터 시작
- Strict writer는 별도 profile flag로 분리
- validator 결과도 strict/transitional profile을 분리해서 저장

---

## 8. 매크로/보안/오라클 자동화 주의점

Microsoft 문서상 `Workbooks.Open`으로 파일을 프로그램에서 열면 기본적으로 매크로가 활성화된다.
이 동작을 제어하려면 `Application.AutomationSecurity`를 사용해야 한다.
또한 이 설정은 Excel 4.0 macro에 대해 완전한 차단이 아니며, 해당 경우 프롬프트가 나타날 수 있다.

### 구현 시 결론

`excel-oracle-win`은 다음 규칙을 강제해야 한다.

1. Excel 프로세스 시작
2. `AutomationSecurity = msoAutomationSecurityForceDisable`
3. alerts/UI 최소화
4. workbook open
5. `CalculateFullRebuild()`
6. `Value2` 덤프 + PDF export
7. 프로세스 종료 및 격리

### 추가 권장

- 오라클 테스트는 별도 Windows runner에서만 수행
- 파일 시스템 sandbox 디렉터리 사용
- 외부 링크 업데이트/프롬프트 차단 옵션 관리
- crash/hang 감지를 위한 timeout wrapper 추가

---

## 9. Open XML SDK, POI, LibreOffice를 어떻게 참고할 것인가

Open XML SDK는 공식 GitHub 설명에서도 low-level OPC/Open XML 조작을 위한 프레임워크이며, 고수준 생산성 API를 직접 제공하려는 의도가 아니라고 밝힌다.
즉, 네 프로젝트의 public API 템플릿으로 삼을 필요는 없지만, **writer/validator 비교 기준**으로는 매우 좋다.

Apache POI는 Java 생태계의 강력한 비교 대상이고, LibreOffice의 테스트 데이터는 parser/round-trip/fuzz seed로 매우 유용하다.
다만 둘 다 “Excel 정답”이 아니라 **보조 비교 구현**으로 써야 한다.

### 권장 포지션

- Excel desktop: **behavior oracle**
- Open XML SDK: **OPC/OpenXML 구조 비교기 + validator 기반 검사기**
- Apache POI: **대체 구현 비교기**
- LibreOffice corpus: **회귀/퍼즈 seed**

---

## 10. 구현 우선순위 원리

초반에 가장 위험한 실수는 “많이 보이는 기능부터” 구현하는 것이다.

### 먼저 해야 할 것

1. **spec source registry 고정**
2. **IDL/codegen 파이프라인 구축**
3. **OPC + lossless xlsx round-trip**
4. **Excel oracle harness 구축**
5. **formula parser + calc graph skeleton**

### 나중에 해야 할 것

- 차트/도형의 완전 렌더링
- `.xls` 지원
- VBA 실행기
- 세밀한 UI 호환 동작

이 순서를 지키면, public API와 내부 엔진이 초반부터 같은 기준선 위에 놓인다.

---

## 11. 설계 체크리스트

### 반드시 분리할 것

- Object Model vs model storage
- formula text vs normalized AST
- cell storage value vs calc value vs OM value
- xlsx codec vs calc engine
- runtime state vs persisted workbook state
- validator/diff/oracle 테스트

### 반드시 보존할 것

- unknown parts
- unknown relationships
- future functions markers
- strict/transitional profile 정보
- workbook calculation settings
- date system

### 반드시 오라클로 검증할 것

- 날짜 직렬화
- dynamic array spill
- implicit intersection
- structured references
- print/export 결과
- named range scope
- validation/conditional formatting round-trip

---

## 12. 공식 참고 자료

### 계산/수식/동적 배열
- Excel Recalculation  
  https://learn.microsoft.com/en-us/office/client-developer/excel/excel-recalculation
- Excel performance: improving calculation performance  
  https://learn.microsoft.com/en-us/office/vba/excel/concepts/excel-performance/excel-improving-calculation-performance
- Multithreaded recalculation in Excel  
  https://learn.microsoft.com/en-us/office/client-developer/excel/multithreaded-recalculation-in-excel
- Formula vs Formula2  
  https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/range-formula-vs-formula2
- Range.Formula2 property  
  https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2
- Dynamic array formulas and spilled array behavior  
  https://support.microsoft.com/en-us/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531

### Object Model / interop
- Excel object model overview  
  https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model
- Range object  
  https://learn.microsoft.com/en-us/office/vba/api/excel.range%28object%29
- Range.Value2 property  
  https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2
- Office primary interop assemblies  
  https://learn.microsoft.com/en-us/visualstudio/vsto/office-primary-interop-assemblies?view=visualstudio

### 포맷/표준
- ECMA-376  
  https://ecma-international.org/publications-and-standards/standards/ecma-376/
- MS-XLSX  
  https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/f780b2d6-8252-4074-9fe3-5d7bc4830968
- MS-OI29500  
  https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/1fd4a662-8623-49c0-82f0-18fa91b413b8
- Open XML formats and file name extensions  
  https://support.microsoft.com/en-us/office/open-xml-formats-and-file-name-extensions-5200d93c-3449-4380-8e11-31ef14555b18
- Save a workbook in another file format  
  https://support.microsoft.com/en-us/office/save-a-workbook-in-another-file-format-6a16c862-4a36-48f9-a300-c2ca0065286e

### 테스트 / 오라클
- Workbooks.Open  
  https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open
- Application.AutomationSecurity  
  https://learn.microsoft.com/ko-kr/office/vba/api/excel.application.automationsecurity
- Application.CalculateFullRebuild  
  https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild
- Workbook.ExportAsFixedFormat  
  https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.exportasfixedformat
- Worksheet.ExportAsFixedFormat  
  https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.exportasfixedformat
- OpenXmlValidator.Validate  
  https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.validation.openxmlvalidator.validate?view=openxml-3.0.1

### 역사적 호환성
- Excel incorrectly assumes that the year 1900 is a leap year  
  https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/wrongly-assumes-1900-is-leap-year
