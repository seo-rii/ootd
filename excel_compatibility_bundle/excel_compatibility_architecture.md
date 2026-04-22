# Excel Compatibility Core 설계 초안 (Rust + WASM)

## 1. 목표

이 프로젝트의 목표는 “엑셀 파일을 읽고 쓰는 라이브러리”가 아니라, **Excel Object Model 호환 코어**를 만드는 것이다.

핵심 목표:

1. `Application / Workbooks / Workbook / Worksheets / Worksheet / Range / Names / Tables` 중심의 **Object Model 호환 계층** 제공
2. `.xlsx/.xlsm/.xltx/.xltm`에 대한 **안정적인 load/save 및 round-trip**
3. 실제 Excel과 대조 가능한 **runtime + calculation engine**
4. Rust native + WebAssembly 빌드를 통해 **TS/JS에서도 사용 가능**
5. loader / writer / renderer / calc / wasm 바인딩을 **분리된 extension crate**로 설계

비목표(v0):

- VBA/XLM 실행기 구현
- `.xls` 완전 지원
- 차트/도형의 100% 시각 렌더링
- Excel UI 자체 재구현

---

## 2. 제품 정의

### 2.1 최종 제품 이미지

이 프로젝트는 다음 4개 층으로 구성된다.

1. **Object Model 계약층**
   - Excel VBA Object Model과 비슷한 public API 표면
   - 코드 생성 기반으로 일관성 유지

2. **Runtime 층**
   - `Application`, active workbook/sheet, selection, events, undo, calculation mode 관리
   - Object Model 호출을 실제 model/calc/io에 연결

3. **Persistence / Model 층**
   - Workbook/Worksheet/Table/Style/Name/Formula 등의 실제 저장 구조
   - 파일 포맷과 계산 엔진 사이의 공통 중간 표현

4. **Extension 층**
   - `xlsx loader/writer`
   - `renderer`
   - `wasm/js binding`
   - 나중에 `xlsb`, `xls` 추가

---

## 3. 스펙 우선순위

### 3.1 스펙 계층

스펙은 아래 우선순위로 사용한다.

#### A. Object Model API 계약
1. **Excel Type Library / PIA**
2. **Excel VBA reference**

용도:
- 객체/컬렉션 구조
- 속성/메서드/이벤트 이름
- enum 이름과 값
- optional parameter, 반환 타입, 컬렉션 멤버

#### B. 파일 포맷 / 패키징
1. **ECMA-376 / ISO 29500**
2. **MS-XLSX**
3. **MS-OI29500**
4. **MS-OSHARED**
5. 필요 시 **MS-ODRAWXML**, **MS-OFFCRYPTO**

용도:
- OPC/ZIP/relationship/content-types
- SpreadsheetML 기본 구조
- Office 구현 차이와 확장
- 공통 Office data structure
- drawing/ext 파트 해석

#### C. 행동 의미론
1. **실제 Microsoft Excel 실행 결과**
2. Excel recalculation / date system / function compatibility 문서

용도:
- 계산 순서
- dirty propagation
- dynamic array / spill
- 1900/1904 date system
- 정밀도/에러 동작
- 버전 차이

### 3.2 규칙

- 문서 스펙과 실제 Excel 동작이 다르면 **Excel을 정답 오라클**로 간주한다.
- 스펙이 모호하면:
  1. 실제 Excel에서 최소 재현 workbook 생성
  2. 저장 결과(xlsx/xml/pdf/json)를 수집
  3. 그 결과를 compatibility test로 승격

---

## 4. 패키지 구조

```text
crates/
  office-idl                # canonical OM schema (build-time only)
  office-codegen            # type library/PIA -> IDL -> Rust/TS codegen

  office-common             # Variant, Error, Color, units, ids, shared enums
  office-opc                # ZIP/OPC, content types, relationships
  office-ooxml-shared       # shared OOXML helpers, MC/ext helpers
  office-runtime            # session, handles, dispatch, events, undo

  office-om-common          # generated Office OM surface
  excel-om                  # generated Excel OM surface

  excel-model               # workbook graph / persisted structures
  excel-formula             # lexer/parser/AST/ref resolver
  excel-calc                # dependency graph / scheduler / evaluator
  excel-runtime             # concrete OM implementations for Excel

  excel-xlsx                # .xlsx/.xlsm/.xltx/.xltm codecs
  excel-render              # sheet renderer (HTML/SVG first)
  office-wasm               # wasm exports
  office-ts                 # npm wrapper + generated d.ts

  excel-oracle-win          # Windows-only Excel automation harness
  office-fixtures           # corpora metadata, golden JSON/PDF, snapshots
  office-fuzz               # fuzz targets

  excel-xlsb               # later
  excel-xls                # much later
```

### 4.1 crate 책임 분리 원칙

- `excel-model`은 **파일에 저장 가능한 상태만** 가진다.
- `excel-runtime`은 **런타임 상태**를 가진다.
  - active workbook
  - selection
  - calculation mode
  - event queue
  - UI-like state
- `excel-xlsx`는 **포맷 codec**일 뿐, calc/runtime을 모른다.
- `excel-calc`는 **XML/ZIP을 몰라야** 한다.
- `office-om-common` / `excel-om`은 **generated surface**이고 직접 비즈니스 로직을 많이 넣지 않는다.

---

## 5. Object Model 설계

## 5.1 철학

Object Model은 내부 저장 구조가 아니라 **façade**다.

예:
- `Range`는 셀 저장소 그 자체가 아니라,
  `(sheet_id, areas[], context)` 위의 view/handle
- `Selection`, `ActiveCell`, `UsedRange`, `CurrentRegion`은 runtime 해석 결과
- `Rows(5)`, `Columns("B")`, `Range("A1:C3")`는 **lazy handle**이어야 함

### 5.2 public object 집합 (v0.1)

우선 구현 표면:

- `Application`
- `Workbooks`
- `Workbook`
- `Worksheets`
- `Worksheet`
- `Range`
- `Rows`
- `Columns`
- `Names`
- `Name`
- `Table` / `ListObject`
- `Style` / `Styles`
- `Window` (최소 subset)
- `CellFormat` 관련 최소 subset

### 5.3 범위 객체 원칙

`Range`는 다음을 지원해야 한다.

- 단일 셀
- 단일 직사각형
- multi-area range
- row/column whole selection
- 3D reference를 위한 별도 내부 참조 표현

내부 표현 추천:

```rust
pub struct RangeRef {
    pub workbook_id: WorkbookId,
    pub sheet_scope: SheetScope,        // SingleSheet | MultiSheet3D
    pub areas: SmallVec<[Rect; 1]>,     // multi-area
}
```

### 5.4 Object Model 노출 형태

Rust public API는 2층으로 나눈다.

1. **Strong API**
   - 타입 안전
   - Rust 친화적

2. **Dispatch API**
   - VBA/COM에 가까운 late-bound 호출 흉내
   - wasm/js와의 동적 브리지

예:

```rust
pub trait ExcelObject {
    fn type_name(&self) -> &'static str;
    fn get(&self, member: &str, args: &[Value]) -> Result<Value, OmError>;
    fn set(&mut self, member: &str, value: Value) -> Result<(), OmError>;
    fn invoke(&mut self, member: &str, args: &[Value]) -> Result<Value, OmError>;
}
```

---

## 6. office-idl

## 6.1 목적

`office-idl`은 runtime crate가 아니라 **빌드 시점 canonical schema**다.

용도:
- type library / PIA / 문서에서 가져온 OM 구조를 하나의 정규화된 형식으로 저장
- Rust/TS dispatch metadata 생성
- 지원 상태(manual/auto/stub) 관리

### 6.2 IDL 핵심 엔티티

- library
- enum
- interface
- class
- member
- parameter
- property
- method
- event
- collection metadata
- alias / deprecated / version note
- support matrix

### 6.3 지원 상태

모든 멤버에 지원 상태를 넣는다.

- `generated_only`
- `stub`
- `partial`
- `implemented`
- `oracle_verified`
- `unsupported`

### 6.4 예시

```json
{
  "library": "Excel",
  "version": "16.0",
  "interfaces": [
    {
      "name": "Range",
      "kind": "dispatch",
      "members": [
        {
          "name": "Value2",
          "member_kind": "property",
          "access": "readwrite",
          "return_type": "Variant",
          "params": [],
          "support": "implemented"
        },
        {
          "name": "Calculate",
          "member_kind": "method",
          "params": [],
          "return_type": "Void",
          "support": "partial"
        }
      ]
    }
  ]
}
```

---

## 7. 코드 생성 파이프라인

## 7.1 파이프라인

```text
Excel Type Library / PIA
        ↓
extractor (Windows build tool)
        ↓
canonical office-idl JSON
        ↓
codegen
   ├─ excel-om Rust traits/handles
   ├─ dispatch metadata tables
   ├─ TS declaration files
   └─ docs/stub coverage report
```

### 7.2 extractor 책임

- interface/class/member 열거
- enum 값 수집
- optional/default parameter 정규화
- 별칭/숨김 멤버 분리
- PIA와 type library 사이 이름 충돌 정리
- canonical JSON 생성

### 7.3 codegen 산출물

- `excel_om_generated.rs`
- `excel_dispatch_table.rs`
- `excel_types.d.ts`
- coverage report (`implemented/partial/stub/unsupported`)

---

## 8. 런타임 설계

## 8.1 핵심 구조

```rust
pub struct OfficeSession {
    pub options: SessionOptions,
    pub workbooks: HandleMap<WorkbookHandle>,
    pub events: EventQueue,
    pub profiles: CompatibilityProfiles,
    pub calc_service: CalcService,
}

pub struct ExcelAppRuntime {
    pub active_workbook: Option<WorkbookId>,
    pub active_sheet: Option<SheetId>,
    pub selection: Option<RangeRef>,
    pub calculation_mode: CalculationMode,
    pub screen_updating: bool,
    pub display_alerts: bool,
}
```

### 8.2 핸들 vs 소유 데이터

- public OM 객체는 대부분 `Handle<T>`
- 실제 데이터는 model/runtime store에 존재
- wasm에서는 숫자 핸들/opaque handle id로 전달 가능

### 8.3 이벤트

초기 이벤트 subset:

- workbook open
- before save
- after save
- sheet change
- calculate

이벤트는 실제 Excel처럼 모두 동기 실행하려고 하지 말고,
초기에는 **명시적 dispatch queue**로 단순화한다.

---

## 9. 모델 설계

## 9.1 WorkbookGraph

```rust
pub struct WorkbookModel {
    pub metadata: WorkbookMetadata,
    pub sheets: SlotMap<SheetId, WorksheetModel>,
    pub names: Vec<DefinedName>,
    pub styles: StyleRepository,
    pub tables: Vec<TableModel>,
    pub connections: Vec<ConnectionModel>,
    pub calc_props: CalcProperties,
    pub unknown_parts: Vec<OpaquePart>,
}
```

### 9.2 WorksheetModel

```rust
pub struct WorksheetModel {
    pub sheet_info: SheetInfo,
    pub dimensions: Option<Rect>,
    pub cells: CellStore,
    pub merges: Vec<Rect>,
    pub row_meta: RowMetaStore,
    pub col_meta: ColMetaStore,
    pub validations: Vec<DataValidationRule>,
    pub conditional_formats: Vec<ConditionalFormatBlock>,
    pub drawings: Vec<DrawingAnchor>,
    pub comments: Vec<CommentRef>,
    pub tables: Vec<TableId>,
    pub unknown_nodes: Vec<OpaqueXmlNode>,
}
```

### 9.3 CellStore

희소 구조 권장:

- row chunking
- cell map inside chunk
- formula/value/style 분리 저장 가능
- dense row/column에 대한 최적화는 나중

예:

```rust
pub enum CellContent {
    Blank,
    Number(f64),
    Bool(bool),
    SharedString(SharedStringId),
    InlineString(String),
    Error(CellError),
    Formula(FormulaCell),
}
```

---

## 10. formula / calc

## 10.1 crate 분리

- `excel-formula`
  - lexer
  - parser
  - AST
  - A1/R1C1 reference normalization
  - structured reference parser
- `excel-calc`
  - dependency graph
  - dirty propagation
  - scheduler
  - evaluator
  - cache
  - iterative calculation

## 10.2 FormulaCell 구조

```rust
pub struct FormulaCell {
    pub original_text: String,          // round-trip용
    pub normalized_ast: FormulaAst,     // 계산용
    pub cached_value: Option<ScalarValue>,
    pub formula_kind: FormulaKind,      // Normal | Array | Shared | DataTable | DynamicArray
}
```

## 10.3 계산 엔진 설계 원칙

반드시 구현해야 하는 것:

- dependency graph 재구성
- dirty marking
- partial recalc
- full rebuild
- circular reference + iteration 설정
- calc chain를 힌트로만 사용
- cache invalidation 정책

### 10.4 compatibility profile

프로필별 차이를 registry로 둔다.

```rust
pub struct CompatibilityProfile {
    pub name: String,
    pub dynamic_arrays: bool,
    pub implicit_intersection_at: bool,
    pub function_registry: FunctionRegistry,
    pub date_system_default: DateSystem,
    pub quirks: QuirkFlags,
}
```

### 10.5 함수 registry 메타데이터

각 함수는 최소 다음 메타데이터를 가진다.

- 이름
- arity
- volatility
- return behavior (scalar/array/range)
- version gate
- thread-safe 여부
- error propagation rules
- argument coercion rules

---

## 11. 포맷 codec

## 11.1 xlsx v0 범위

우선 지원:

- `.xlsx`
- `.xlsm` (macro-preserving round-trip)
- `.xltx`
- `.xltm`
- Strict `.xlsx`는 read 우선, write는 뒤로

후순위:

- `.xlsb`
- `.xls`

## 11.2 macro policy

v0에서는 VBA/XLM을 **실행하지 않는다**.

대신 다음을 보존한다.

- `vbaProject.bin`
- 관련 relationship
- signature
- macro-related opaque parts
- extLst/unknown XML

## 11.3 unknown preservation

알 수 없는 요소는 삭제하지 않는다.

보존 대상:

- unknown part
- unknown relationship
- extLst
- AlternateContent
- namespace extension element/attr
- vendor-specific payload

writer는 손대지 않은 unknown payload를 최대한 그대로 다시 쓴다.

---

## 12. lossless mode / normalized mode

두 모드를 명시적으로 분리한다.

### 12.1 lossless mode
- 원문 XML의 모르는 노드/속성/part를 보존
- writer는 최대한 원문 구조를 유지
- round-trip 안정성 우선

### 12.2 normalized mode
- 내부 모델 재구성 후 canonical write
- diff 친화적
- 일부 원문 구조는 바뀔 수 있음

기본값은 **lossless mode**.

---

## 13. WASM / TypeScript

## 13.1 원칙

FFI 호출은 coarse-grained API로 제한한다.

나쁜 예:
- `get_cell_value(sheet, row, col)`를 수십만 번 호출

좋은 예:
- `load(bytes)`
- `save(profile)`
- `get_range_values(sheet, rect)`
- `set_range_values(sheet, rect, matrix)`
- `calculate(scope)`
- `render_sheet(viewport)`

### 13.2 JS export 예시

```ts
const book = await excel.load(bytes);
const values = book.getRangeValues("Sheet1", "A1:C1000");
book.setRangeValues("Sheet1", "D1:D1000", output);
book.calculate({ scope: "workbook" });
const out = book.save({ format: "xlsx" });
```

### 13.3 바인딩 정책

- wasm boundary에서는 UTF-8 string / bytes / JSON-like struct 중심
- 아주 세밀한 OM 조작은 native Rust API에 남겨두고, JS는 batch 중심으로 감싼다
- TS declaration은 `office-idl`에서 같이 생성

---

## 14. renderer

v0 목표는 “Excel UI 재현”이 아니라 **검증 가능한 시각 출력**이다.

순서:

1. HTML/SVG sheet renderer
2. snapshot comparison
3. 나중에 PDF renderer

초기 범위:
- grid
- row/col sizing
- number format
- basic font/alignment/fill/border
- merged cell
- freeze pane visual
- conditional formatting 일부
- image anchor read-preserve

차트/도형은 초기에는:
- read-preserve
- bbox/anchor 기반 최소 표현
- 나중에 full render

---

## 15. 구현 순서

## M0 — spec freeze / codegen
산출물:
- `specs/sources.toml`
- `office-idl` schema
- extractor prototype
- codegen skeleton

## M1 — OPC + xlsx round-trip
산출물:
- `office-opc`
- `excel-xlsx`
- `.xlsx/.xlsm/.xltx/.xltm` open/save
- unknown part preservation

## M2 — core object model
산출물:
- `Application/Workbook/Worksheet/Range/Names`
- range addressing
- used range / rows / columns
- Value / Value2 기초

## M3 — calc core
산출물:
- parser
- AST
- dependency graph
- basic evaluator
- dirty propagation
- full rebuild / partial recalc

## M4 — 실무 기능
산출물:
- styles / numFmt
- data validation
- conditional formatting
- tables / structured refs
- autofilter
- print settings
- freeze panes

## M5 — renderer / oracle hardening
산출물:
- HTML/SVG renderer
- PDF comparison workflow
- oracle regression suite

## M6 — binary formats
산출물:
- `.xlsb` read
- 이후 `.xls`

---

## 16. 테스트 전략

## 16.1 테스트 층위

1. unit tests
2. schema/format validation
3. round-trip tests
4. differential tests
5. Excel oracle tests
6. render snapshot tests
7. fuzz tests
8. corpus replay regression tests

### 16.2 unit test 대상

- A1 parser
- R1C1 parser
- formula lexer
- operator precedence
- structured reference parser
- date serial conversion
- error values
- number coercion
- style inheritance
- merge rules
- name resolution

### 16.3 schema/format validation

writer 산출물은 외부 validator로 검증한다.

- Open XML SDK `OpenXmlValidator`
- part-level + package-level validation
- strict/transitional profile 구분

### 16.4 round-trip 검증

두 종류로 본다.

1. **Semantic round-trip**
   - Excel로 열었을 때 의미가 유지되는가

2. **Lossless round-trip**
   - unknown/ext payload가 유지되는가
   - XML canonical diff
   - ZIP part inventory diff
   - relationship diff

### 16.5 differential test

동일 fixture를 여러 구현으로 읽어 비교:

- 우리 라이브러리
- Open XML SDK
- Apache POI
- LibreOffice(필요 시)

주의:
- 이들은 “정답”이 아니라 **비교 대상**
- 정답은 Excel oracle

### 16.6 render snapshot

- sheet -> SVG/PNG/PDF
- golden image와 pixel diff
- tolerance는 글꼴/OS별로 조절

### 16.7 fuzz

분리된 fuzz target:

- zip/opc parser
- xml reader
- formula lexer/parser
- xlsx workbook loader
- style parser
- relationship resolver

---

## 17. fixture corpus 전략

## 17.1 소스 분류

### A. 공식 Microsoft 샘플
- Office Scripts sample workbook
- Data Validation Examples.xlsx
- Power BI Financial Sample workbook
- MOS Excel / Excel Expert course materials

### B. 오픈소스 테스트 자산
- Open XML SDK test/sample/data
- Apache POI `test-data`
- LibreOffice `sc/qa/unit/data`

### C. 직접 생성하는 synthetic corpus
가장 중요함.

조합 매트릭스를 기반으로 자동 생성:
- styles
- merge
- validation
- CF
- names
- tables
- formulas
- dynamic arrays
- hidden sheets
- print settings
- freeze panes
- drawings/images
- strict vs transitional
- macro-preserving package

### D. 실사용 corpus
- 내부/공개 스프레드시트
- 권리/민감정보 정리 후 회귀용으로 보관

## 17.2 fixture manifest

각 fixture에는 메타데이터를 붙인다.

```toml
[[fixture]]
id = "dv_examples_ms"
path = "corpus/ms/Data Validation Examples.xlsx"
source = "microsoft"
tags = ["xlsx", "validation", "list", "date", "custom"]
oracle = true
roundtrip = true
render = false
notes = "공식 데이터 유효성 검사 예제"
```

---

## 18. Excel oracle 테스트

## 18.1 별도 crate

`excel-oracle-win`은 Windows + Excel 설치 환경에서만 동작하는 integration harness다.

### 18.2 목적

- 실제 Excel을 정답 오라클로 사용
- workbook open/calculate/export 결과를 golden으로 저장
- 우리 writer 결과를 다시 Excel에 열어 검증

### 18.3 기본 프로토콜

1. Excel.Application 생성
2. `Visible = false`
3. `DisplayAlerts = false`
4. `AutomationSecurity = ForceDisable`
5. `Workbooks.Open(...)`
6. `Application.CalculateFullRebuild()`
7. workbook/sheet/range 메타데이터 수집
8. `Range.Value2` 기반 값 추출
9. workbook/worksheet를 PDF로 export
10. JSON + PDF + log 저장
11. 종료 및 cleanup

### 18.4 수집 대상

- workbook calc/version/meta
- sheet name/order/visibility
- used range
- named ranges
- table definitions
- selected probe cells/ranges
- formulas
- Value2
- number format
- row/column metrics 일부
- export PDF

### 18.5 JSON 오라클 예시

```json
{
  "workbook": {
    "name": "sample.xlsx",
    "sheets": [
      {
        "name": "Sheet1",
        "used_range": "A1:F20",
        "probes": [
          {
            "addr": "A1",
            "formula": "=SUM(B1:B10)",
            "value2": 42.0,
            "number_format": "0.00"
          }
        ]
      }
    ]
  }
}
```

### 18.6 probe 전략

모든 셀을 다 덤프하지 않는다. 3계층으로 나눈다.

1. **핵심 probe**
   - 공식 표적 셀 목록
2. **자동 probe**
   - used range 내부의 formula/value/style boundary 셀 샘플링
3. **full dump mode**
   - 작은 fixture에만 적용

### 18.7 보안

- 매크로 강제 비활성화
- 외부 링크 갱신 금지
- 네트워크 경로 금지
- sandbox path에서만 실행
- 실행 후 orphaned Excel process cleanup

---

## 19. CI 전략

## 19.1 OS matrix

- Linux: unit, parser, xlsx, fuzz smoke
- macOS: core portability
- Windows: oracle-disabled 기본 + oracle-enabled self-hosted

## 19.2 job 분리

- `lint`
- `unit-core`
- `xlsx-roundtrip`
- `validator`
- `render-snapshots`
- `oracle-win`
- `fuzz-smoke`
- `corpus-replay`

## 19.3 릴리스 게이트

릴리스 기준 예:

- critical fixture pass >= 99%
- lossless round-trip critical corpus pass 100%
- validator critical corpus pass 100%
- oracle diff 허용 오차 이내
- no new crash from fuzz seeds

---

## 20. sources.toml 운영 규칙

- URL만 저장하지 말고 **revision/date/profile**까지 pin
- 스펙이 업데이트되면 diff를 남김
- codegen 대상 버전과 runtime 지원 버전을 분리

예:
- `om_contract.excel.type_library = 16.0`
- `ooxml.ecma376.part2 = 5th_edition_2021`
- `office_impl.ms_xlsx.revision = 2025-05-20`

---

## 21. 구현 중 결정해야 할 세부 정책

1. `Value` vs `Value2` 우선순위
   - runtime 내부 계산은 `Value2` 기준 권장
2. shared formula 표현 방식
3. style dedup 시점
4. unknown XML node raw preservation 방식
5. calc cache write-back 정책
6. strict write 지원 시점
7. 3D reference와 OM Range의 관계
8. structured reference canonicalization 규칙
9. thread-safe calc subset 정의
10. wasm memory management

---

## 22. v0.1 완료 기준

다음이 되면 v0.1로 본다.

- `.xlsx/.xlsm/.xltx/.xltm` open/save 가능
- macro payload preservation 가능
- `Application/Workbook/Worksheet/Range/Names` 기본 동작
- `Value2`, formula, basic style, names, table, validation, CF subset 지원
- dependency graph + partial recalc + full rebuild
- Excel oracle integration 동작
- 공식 corpus + synthetic corpus 기반 회귀 체계 구축
- wasm/js batch API 제공

---

## 23. 최종 권장

첫 구현은 “파일 포맷”보다 아래 3개를 먼저 끝내는 것이 맞다.

1. `specs/sources.toml`
2. `office-idl + codegen`
3. `excel-oracle-win`

이 3개가 먼저 있어야 이후 구현이 흔들리지 않는다.
