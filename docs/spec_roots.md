# OOTD Spec Roots

`ootd`의 심화 구현 전에 먼저 고정해야 하는 공식 문서 루트를 정리한다.
목표는 "전체 Excel Object Model과 XLSX 동작을 어디서 가져올지"를 명확히 하는 것이다.

## Summary

- XLSX/패키징 정본은 `ECMA-376` + OPC + Microsoft Open Specifications(`MS-XLSX`, `MS-OI29500`, `MS-OSHARED`) 조합이다.
- Excel Object Model의 machine-readable 계약은 Microsoft Learn 설명 문서가 아니라, 설치된 Excel의 COM type library와 PIA가 우선이다.
- Microsoft Learn의 Excel VBA / Interop 문서는 typed contract 추출 뒤에 semantics와 behavioral oracle을 보강하는 용도로 쓴다.

## File Format Roots

### P0

- `ECMA-376`
  - URL: <https://ecma-international.org/publications-and-standards/standards/ecma-376/>
  - 역할: OOXML의 표준 루트. Part 1/2/3/4를 통해 SpreadsheetML, OPC, markup compatibility, transitional feature의 기준을 제공한다.
  - 용도: `excel-xlsx`, `office-opc`의 기본 구조와 writer/read path의 정본.

- `OPC Fundamentals`
  - URL: <https://learn.microsoft.com/en-us/previous-versions/windows/desktop/opc/open-packaging-conventions-overview>
  - 역할: 패키지, 파트, relationship, logical/physical model을 빠르게 재확인하는 운영용 루트.
  - 용도: lossless package preserve 정책과 relationship handling 검증.

- `MS-XLSX`
  - URL: <https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/f780b2d6-8252-4074-9fe3-5d7bc4830968>
  - 역할: Excel이 SpreadsheetML 위에 얹는 확장 구조와 동작을 설명한다.
  - 용도: `excel-xlsx`에서 Excel-specific extension, worksheet semantics, 확장 part/attr 처리.

### P1

- `MS-OI29500`
  - URL: <https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/cc8df712-c9b7-4450-87ba-208fd1079a2a>
  - 역할: Office가 ISO/IEC 29500을 구현할 때 택한 선택과 implementation note를 제공한다.
  - 용도: 표준만으로 모호한 영역에서 Excel 호환 동작을 좁히는 데 사용.

- `MS-OSHARED`
  - URL: <https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-oshared/d93502fa-5b8f-4f47-a3fe-5574046f4b8d>
  - 역할: 여러 Office 포맷에 공통으로 쓰이는 데이터형과 구조를 설명한다.
  - 용도: shared structure, 공통 object/data type, auxiliary parsing reference.

## Object Model Roots

### P0

- `COM type library + PIA acquisition path`
  - URL: <https://learn.microsoft.com/en-us/dotnet/framework/interop/importing-a-type-library-as-an-assembly>
  - URL: <https://learn.microsoft.com/en-us/dotnet/framework/interop/how-to-generate-primary-interop-assemblies-using-tlbimp-exe>
  - URL: <https://learn.microsoft.com/en-us/visualstudio/vsto/office-primary-interop-assemblies?view=visualstudio>
  - 역할: 실제 machine-readable OM contract를 가져오는 경로를 설명한다.
  - 용도: Windows + Excel 설치 환경에서 type library를 추출하고, `office-idl`의 canonical source로 정규화하는 단계의 기준.
  - 메모: Learn 문서가 OM 전체를 구조적으로 제공하는 것은 아니다. 최종 계약은 type library와 PIA projection에 더 가깝다.

### P1

- `Microsoft.Office.Interop.Excel namespace`
  - URL: <https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel?view=excel-pia>
  - 역할: managed projection 기준으로 인터페이스/클래스/enum surface를 훑는 루트.
  - 용도: `office-idl` 정규화 결과와 surface coverage를 대조하고, 누락 멤버를 빠르게 탐색하는 용도.

- `Excel VBA object model`
  - URL: <https://learn.microsoft.com/en-us/office/vba/api/overview/excel/object-model>
  - 역할: Excel OM의 설명형 루트.
  - 용도: `Workbook`, `Worksheet`, `Range`, `Application` 등의 의미론과 예제를 보강하는 behavioral reference.

- `Office object library reference`
  - URL: <https://learn.microsoft.com/en-us/office/vba/api/overview/library-reference/reference-object-library-reference-for-office>
  - 역할: Office 공통 object model 설명 루트.
  - 용도: Excel 외 Office shared object와 enum, shared semantics 확인.

### P2

- `개별 API 페이지 패턴`
  - VBA: `https://learn.microsoft.com/en-us/office/vba/api/excel.<object>`
  - VBA member: `https://learn.microsoft.com/en-us/office/vba/api/excel.<object>.<member>`
  - Interop example: `https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.application?view=excel-pia`
  - 용도: 특정 object/member semantics를 테스트 oracle이나 facade 설계에 바로 연결할 때 사용.

## Recommended Intake Order

1. `ECMA-376` Part 1/2/3/4와 `MS-XLSX`를 우선 pinning한다.
2. `MS-OI29500`와 `MS-OSHARED`를 보조 루트로 붙여 Excel-specific ambiguity를 줄인다.
3. Windows/Excel 설치 환경에서 COM type library와 PIA를 추출한다.
4. 추출 결과를 `office-idl` schema로 정규화해 `specs/pinned` 아래 canonical dataset으로 저장한다.
5. Learn의 VBA / Interop 문서로 behavior, examples, naming, version notes를 보강한다.

## Immediate Next Step

- `specs/pinned/` 아래에 OM source manifest를 만들고, Excel type library/PIA 추출 절차와 target Office version을 고정한다.
- 이후 `office-idl`에는 문서 설명이 아니라 추출된 type/member dataset을 넣는다.
- `excel-runtime`과 future facade crate는 그 canonical dataset을 기준으로 surface를 늘린다.
