# Excel 호환 라이브러리 참고 자료 지도와 읽기 순서

이 문서는 “무엇을 읽어야 하는가”를 바로 결정할 수 있도록 만든 **참고 자료 지도**다.
모든 자료를 같은 중요도로 취급하면 오히려 시간을 낭비하므로, 역할별로 나누어 본다.

---

## 1. 먼저 고정할 원칙

이 프로젝트의 자료는 아래 네 계층으로 나누어 읽어야 한다.

1. **API 계약 자료**
   - Object Model 이름, 멤버, enum, optional parameter, collection 구조
2. **포맷 자료**
   - ZIP/OPC/SpreadsheetML/Office 확장 파트
3. **동작 자료**
   - recalculation, formula dialect, dynamic array, 날짜/정밀도 같은 의미론
4. **검증 자료**
   - validator, Excel automation, sample workbook, corpus

각 계층은 서로 대체되지 않는다.
예를 들어 VBA reference만 봐서는 `.xlsx` writer를 만들 수 없고,
OOXML 스펙만 봐서는 `Range.Formula2`와 dynamic array 의미론을 맞출 수 없다.

---

## 2. 꼭 읽어야 하는 자료: 우선순위 A

### A-1. Excel Object Model overview
**왜 읽는가**
- 전체 Excel object model의 entry point
- Application/Workbook/Worksheet/Range 관계를 빠르게 잡을 수 있다.

**어떻게 쓰는가**
- public façade 범위를 정할 때 사용
- `excel-om` codegen 결과를 점검할 때 사용

출처  
https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model

### A-2. Office Primary Interop Assemblies 문서
**왜 읽는가**
- PIA가 COM 기반 Office object model과 managed code를 연결한다는 역할을 분명히 설명한다.
- 이 프로젝트에서 “문서 HTML이 아니라 TLB/PIA를 기계 판독 소스로 쓴다”는 결정을 뒷받침한다.

**어떻게 쓰는가**
- `office-idl` extractor 설계 근거
- codegen source priority 정의 근거

출처  
https://learn.microsoft.com/en-us/visualstudio/vsto/office-primary-interop-assemblies?view=visualstudio

### A-3. ECMA-376
**왜 읽는가**
- OOXML 전체의 가장 큰 표준 축
- Part 2(OPC)는 문서 패키징 구현에 직접 필요하다.

**어떻게 쓰는가**
- `office-opc`
- content types / relationships / part naming
- strict/transitional 구분의 기준선

출처  
https://ecma-international.org/publications-and-standards/standards/ecma-376/

### A-4. MS-XLSX / MS-OI29500
**왜 읽는가**
- ECMA/ISO만으로는 부족한 Excel/Office 구현 차이를 메운다.
- 실제 Excel 파일 호환성을 맞추는 데 가장 중요한 보완 스펙이다.

**어떻게 쓰는가**
- `excel-xlsx`
- unknown ext/part 전략 검토
- Office 구현 차이 문서화

출처  
https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/f780b2d6-8252-4074-9fe3-5d7bc4830968  
https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/1fd4a662-8623-49c0-82f0-18fa91b413b8

### A-5. Excel Recalculation
**왜 읽는가**
- dependency tree / calculation chain / recalc 개념을 공식적으로 설명한다.

**어떻게 쓰는가**
- `excel-calc`의 개념 설계 기준
- `calcChain.xml`을 정답이 아니라 힌트로 취급하는 근거

출처  
https://learn.microsoft.com/en-us/office/client-developer/excel/excel-recalculation

---

## 3. 구현 시작 직전에 읽어야 하는 자료: 우선순위 B

### B-1. Range object
**왜 읽는가**
- Range가 단일 사각형이 아니라 contiguous blocks, row/column, 3D range까지 포괄한다.

**설계에 주는 영향**
- `RangeRef`를 area list 기반 view로 설계해야 한다.

출처  
https://learn.microsoft.com/en-us/office/vba/api/excel.range%28object%29

### B-2. Formula vs Formula2 / Formula2 property
**왜 읽는가**
- dynamic array 이후 formula language가 어떻게 달라졌는지 이해하는 데 필요하다.

**설계에 주는 영향**
- internal AST와 output dialect 분리
- profile별 serialization 전략

출처  
https://learn.microsoft.com/en-us/office/vba/excel/concepts/cells-and-ranges/range-formula-vs-formula2  
https://learn.microsoft.com/en-us/office/vba/api/excel.range.formula2

### B-3. Range.Value2
**왜 읽는가**
- Date/Currency special casing을 피한 비교용 값 추출 기준을 제공한다.

**설계에 주는 영향**
- oracle dump는 `Value2`
- interop value 설계 시 `Value`와 구분

출처  
https://learn.microsoft.com/en-us/office/vba/api/excel.range.value2

### B-4. 1900 leap-year bug
**왜 읽는가**
- “논리적으로 틀린 동작도 호환성 정답일 수 있다”는 대표 사례다.

**설계에 주는 영향**
- 날짜 serial을 일반 날짜형으로 바로 변환하지 말아야 한다.
- 호환성 프로필과 historical behavior 테스트가 필요하다.

출처  
https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/wrongly-assumes-1900-is-leap-year

---

## 4. 테스트/검증 설계 시 바로 읽어야 하는 자료: 우선순위 C

### C-1. Workbooks.Open / AutomationSecurity
**왜 읽는가**
- programmatic open 시 macro security가 기본적으로 위험할 수 있다.

**설계에 주는 영향**
- `excel-oracle-win`은 파일 open 전에 `AutomationSecurity`를 강제해야 한다.

출처  
https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open  
https://learn.microsoft.com/ko-kr/office/vba/api/excel.application.automationsecurity

### C-2. CalculateFullRebuild
**왜 읽는가**
- full recalculation + dependency rebuild를 강제하는 공식 엔트리포인트다.

**설계에 주는 영향**
- 오라클 테스트는 `CalculateFullRebuild()`를 기준선으로 사용

출처  
https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild

### C-3. ExportAsFixedFormat
**왜 읽는가**
- workbook/sheet 단위 PDF/XPS export를 공식적으로 지원한다.

**설계에 주는 영향**
- 레이아웃/print/renderer 비교에 PDF golden을 쓸 수 있다.

출처  
https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.exportasfixedformat  
https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.exportasfixedformat

### C-4. OpenXmlValidator
**왜 읽는가**
- package/part/element 단위 validate API를 제공한다.

**설계에 주는 영향**
- writer CI에서 구조 검증 레이어를 자동화 가능

출처  
https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.validation.openxmlvalidator.validate?view=openxml-3.0.1

---

## 5. 샘플/코퍼스 자료: 우선순위 D

### D-1. Data Validation Examples.xlsx
**가치**
- 공식 workbook
- 데이터 유효성 검사 케이스를 빠르게 커버

출처  
https://www.microsoft.com/en-us/download/details.aspx?id=53669

### D-2. Office Scripts sample workbooks
**가치**
- 실제 workbook object, table, formatting, row/column visibility, 샘플 시나리오가 들어 있는 실전형 파일 확보에 유리

출처  
https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/samples-overview

### D-3. MOS Excel 교육 자료 (MO-210 / MO-211)
**가치**
- 실무형 과제 파일, 중간 난도 이상의 연습 workbook 확보에 적합

출처  
https://learn.microsoft.com/en-us/training/educator-center/instructor-materials/microsoft-365-apps-certification-preparation-materials

### D-4. Open XML SDK repository
**가치**
- low-level Open XML 구조를 비교하기 좋고, SDK test/data를 corpus seed로 쓰기 좋다.

출처  
https://github.com/dotnet/Open-XML-SDK

### D-5. Apache POI repository
**가치**
- Java 쪽 대체 구현 비교 대상
- OOXML + OLE2 관점 비교에 유용

출처  
https://github.com/apache/poi

### D-6. LibreOffice core repository
**가치**
- `sc/qa/unit/data` 계열 테스트 파일을 parser 회귀 테스트와 fuzz seed로 활용 가능

출처  
https://github.com/LibreOffice/core

---

## 6. 추천 읽기 순서

### 1주차: 방향 고정
1. Excel object model overview
2. Office PIA 문서
3. ECMA-376 개요
4. MS-XLSX introduction
5. MS-OI29500 introduction

### 2주차: 엔진 개념 고정
1. Excel recalculation
2. Formula vs Formula2
3. Range object
4. Value2
5. 1900 leap-year bug

### 3주차: 테스트 체계 고정
1. Workbooks.Open
2. AutomationSecurity
3. CalculateFullRebuild
4. ExportAsFixedFormat
5. OpenXmlValidator

### 4주차: corpus 정리
1. Data Validation Examples
2. Office Scripts sample workbooks
3. MOS materials
4. Open XML SDK / Apache POI / LibreOffice corpus 수집

---

## 7. 프로젝트 문서와의 연결

- `sources.toml`
  - 여기 문서에 있는 자료를 버전·역할·우선순위로 pinning
- `office-idl.schema.json`
  - A-1/A-2의 결과를 저장하는 정규화 포맷
- `excel_compatibility_architecture.md`
  - 이 자료 지도를 실제 제품 구조로 바꾼 문서
- `excel_project_structure_and_interfaces.md`
  - 그 구조를 crate / trait / interface 수준으로 내린 문서
- `excel-oracle-win-protocol.md`
  - C 그룹 자료를 실제 테스트 프로토콜로 고정한 문서

---

## 8. 최종 요약

가장 중요한 건 이 세 가지다.

1. **API 표면은 TLB/PIA에서 뽑는다.**
2. **파일 포맷은 ECMA-376 + MS-XLSX + MS-OI29500 축으로 읽는다.**
3. **행동 의미론은 Excel 오라클로 검증한다.**

이 세 축이 흔들리지 않으면, 이후 구현 우선순위나 crate 구조가 바뀌어도 프로젝트 방향은 잘 유지된다.
