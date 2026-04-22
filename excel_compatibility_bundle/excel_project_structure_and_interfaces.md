# Excel Compatibility Core: 프로젝트 구조 및 핵심 인터페이스

이 문서는 `excel_compatibility_architecture.md`의 후속 문서다.
목표는 **Cargo workspace 구조**, **crate 간 경계**, **핵심 Rust 인터페이스**, **WASM/TS 경계**, **테스트/오라클 인터페이스**를 구체적으로 고정하는 것이다.

---

## 1. 설계 원칙

이 프로젝트는 세 층을 명확히 분리해야 한다.

1. **Object Model 표면**
   - `Application`, `Workbook`, `Worksheet`, `Range` 같은 Excel 호환 API
   - 사용자는 이 층을 통해 문서를 조작한다.

2. **실제 상태와 동작**
   - workbook graph, formula AST, calculation graph, event queue, selection, active window
   - Object Model 호출은 여기서 실제 의미를 갖는다.

3. **포맷/외부 세계**
   - `.xlsx` loader/writer
   - renderer
   - wasm/js 바인딩
   - Excel oracle, validator, fuzz, fixtures

핵심 규칙은 아래와 같다.

- Object Model은 **내부 저장 구조가 아니라 façade**다.
- 계산 엔진은 **파일 포맷을 몰라야** 한다.
- writer는 **런타임 상태를 몰라야** 한다.
- unknown OOXML part/ext/relationship는 **lossless 보존** 가능해야 한다.
- WASM 경계는 **배치 API 중심**이어야 한다.
- generated OM surface와 hand-written runtime implementation은 **분리**되어야 한다.

---

## 2. Cargo workspace 제안 구조

```text
.
├─ Cargo.toml
├─ rust-toolchain.toml
├─ docs/
│  ├─ architecture/
│  ├─ interfaces/
│  ├─ specs/
│  └─ test-protocols/
├─ specs/
│  ├─ sources.toml
│  ├─ office-idl.schema.json
│  ├─ pinned/
│  └─ generated/
├─ tools/
│  ├─ tlb-dump/
│  ├─ pia-dump/
│  ├─ xlsx-diff/
│  └─ oracle-runner/
├─ fixtures/
│  ├─ corpus/
│  ├─ synthetic/
│  ├─ golden/
│  └─ manifests/
├─ crates/
│  ├─ office-idl/
│  ├─ office-codegen/
│  ├─ office-common/
│  ├─ office-opc/
│  ├─ office-ooxml-shared/
│  ├─ office-runtime/
│  ├─ office-om-common/
│  ├─ excel-om/
│  ├─ excel-model/
│  ├─ excel-formula/
│  ├─ excel-calc/
│  ├─ excel-runtime/
│  ├─ excel-xlsx/
│  ├─ excel-render/
│  ├─ office-wasm/
│  ├─ office-ts/
│  ├─ excel-oracle-win/
│  ├─ office-fixtures/
│  ├─ office-fuzz/
│  ├─ excel-xlsb/
│  └─ excel-xls/
└─ .github/
   └─ workflows/
```

### 2.1 crate별 책임

#### build-time / spec 계층

- `office-idl`
  - canonical Object Model schema 정의
  - TLB/PIA에서 뽑은 정보를 정규화하는 중간 표현
  - runtime dependency 없음

- `office-codegen`
  - `office-idl`을 입력으로 받아 Rust/TS/generated metadata 생성
  - `excel-om`, `office-om-common`, `office-ts` 타입 생성에 사용

#### 공통 기반 계층

- `office-common`
  - id/newtype, error, color, address primitive, string interning, shared value type
  - Excel/Word/PowerPoint 공통으로 살아남을 것만 둔다

- `office-opc`
  - ZIP/OPC, content types, relationship, part uri, stream IO
  - OOXML 패키징만 담당

- `office-ooxml-shared`
  - Markup Compatibility, extLst, shared OOXML helpers
  - OOXML 공통 유틸리티

- `office-runtime`
  - object handle, dispatch call, event sink, undo stack, session registry
  - COM/VBA와 비슷한 late-bound 호출을 흉내 낼 공통 런타임

#### generated OM 계층

- `office-om-common`
  - generated common enums/interfaces
  - 예: `MsoTriState`, 공통 `ColorFormat` 계열 일부

- `excel-om`
  - generated Excel object model surface
  - `Application`, `Workbook`, `Worksheet`, `Range`, `Names`, `ListObject` 등

#### Excel core 계층

- `excel-model`
  - workbook graph, sheet data, styles, tables, names, validations, conditional formats
  - 파일에 저장 가능한 상태의 canonical model

- `excel-formula`
  - lexer, parser, AST, reference normalization, structured reference parser

- `excel-calc`
  - dependency graph, dirty propagation, evaluator, recalc scheduler, function registry

- `excel-runtime`
  - Excel object model 구현체
  - runtime state + model + calc + codec를 묶는 orchestrator

#### 외부 extension 계층

- `excel-xlsx`
  - `.xlsx/.xlsm/.xltx/.xltm` loader/writer
  - strict/read profile, macro-preserving round-trip, opaque preservation

- `excel-render`
  - sheet -> render tree -> HTML/SVG/PDF(후순위)

- `office-wasm`
  - coarse-grained WASM exports
  - byte array load/save, batch range ops, render snapshot

- `office-ts`
  - npm wrapper, friendly TS types, async/sync wrapper, browser/node adapters

#### 품질 보증 계층

- `excel-oracle-win`
  - Windows + Excel 데스크톱 오라클
  - open/calc/export/json dump

- `office-fixtures`
  - corpus manifest, fixture metadata, snapshot schema

- `office-fuzz`
  - ZIP/XML/formula/model mutation fuzz target

#### 후순위 포맷

- `excel-xlsb`
  - BIFF12 binary format

- `excel-xls`
  - BIFF8 + CFB legacy format

---

## 3. crate 의존성 규칙

### 3.1 허용 방향

```text
office-idl ───────┐
office-codegen ───┼──> office-om-common
                  ├──> excel-om
                  └──> office-ts (generated d.ts metadata)

office-common <────────────── office-runtime
office-common <────────────── excel-model
office-common <────────────── excel-formula
office-common <────────────── excel-calc
office-common <────────────── excel-xlsx
office-common <────────────── excel-render

office-opc <───────────────── excel-xlsx
office-ooxml-shared <──────── excel-xlsx

excel-model <──────────────── excel-runtime
excel-formula <────────────── excel-calc
excel-calc <───────────────── excel-runtime
excel-om <─────────────────── excel-runtime
office-runtime <───────────── excel-runtime
excel-xlsx <───────────────── excel-runtime
excel-render <─────────────── office-wasm / tools

excel-runtime <────────────── office-wasm
excel-runtime <────────────── excel-oracle-win
office-fixtures <──────────── tests / tools
```

### 3.2 금지 방향

아래 의존은 금지한다.

- `excel-calc -> excel-xlsx`
- `excel-calc -> excel-runtime`
- `excel-model -> excel-runtime`
- `excel-xlsx -> excel-runtime`
- `excel-xlsx -> excel-calc`
- `excel-om -> excel-runtime`
- `office-om-common -> office-runtime`
- `office-common -> excel-*`

즉, **generated surface는 구현을 몰라야 하고**, **포맷 codec은 calc/runtime을 몰라야 한다.**

---

## 4. 이름 규칙

코드를 길게 유지하려면 이름 규칙을 먼저 고정하는 편이 좋다.

- `*Id`
  - persisted entity id
  - 예: `WorkbookId`, `SheetId`, `StyleId`

- `*Handle`
  - runtime object handle
  - 예: `WorkbookHandle`, `RangeHandle`, `ObjectHandle`

- `*Model`
  - persistence graph 구조체
  - 예: `WorkbookModel`, `WorksheetModel`, `TableModel`

- `*Contract`
  - codegen이 만든 object model trait
  - 예: `WorkbookContract`, `WorksheetContract`

- `*Service`
  - 내부 orchestration/service trait
  - 예: `CalcService`, `WorkbookCodecService`

- `*Spec`
  - 입력 옵션/요청 파라미터
  - 예: `OpenWorkbookSpec`, `SaveWorkbookSpec`, `RenderSpec`

- `*Snapshot`
  - 테스트/디버그/오라클 비교용 직렬화 구조

- `*Profile`
  - Excel 버전/호환 모드

---

## 5. 핵심 타입 체계

이 프로젝트에서 제일 먼저 분리해야 하는 타입은 **셀 값**과 **Object Model 값**이다.

### 5.1 `CellValue`와 `OmValue`는 분리한다

#### `CellValue`
엑셀 파일/모델에 저장되는 값이다.

예:
- blank
- bool
- number(f64)
- shared string / inline string / rich text
- error
- formula cached value

#### `OmValue`
Object Model dispatch에서 오가는 값이다.

예:
- `Missing`
- `Empty`
- `Null`
- `Bool`
- `Number`
- `Text`
- `Error`
- `ObjectHandle`
- `Array`

이 둘을 섞으면 문제가 생긴다.

예를 들어:
- `Range.Value2`는 2차원 배열을 반환할 수 있다.
- `Worksheet.Cells(1, 1)`는 객체를 반환한다.
- `Missing`은 optional parameter 의미론에 필요하지만 파일 값이 아니다.

### 5.2 추천 기본 타입

```rust
pub struct WorkbookId(pub u64);
pub struct SheetId(pub u64);
pub struct NameId(pub u64);
pub struct TableId(pub u64);
pub struct StyleId(pub u64);
pub struct ObjectHandle(pub u64);

pub struct Rect {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

pub enum SheetScope {
    Single(SheetId),
    Multi3D { start: SheetId, end: SheetId },
}

pub struct RangeRef {
    pub workbook_id: WorkbookId,
    pub scope: SheetScope,
    pub areas: Vec<Rect>,
}
```

### 5.3 주소 계층

주소 표현도 최소 3개를 분리하는 편이 좋다.

- `A1Address`
- `R1C1Address`
- `RangeRef` (normalized internal form)

loader/writer는 텍스트를 다루고,
formula parser는 문법 노드를 다루며,
runtime/model/calc는 normalized `RangeRef`를 다루도록 나누는 게 안전하다.

---

## 6. 핵심 인터페이스 그룹

여기서는 실제 구현을 시작할 때 바로 기준선으로 삼을 trait 집합을 제안한다.

### 6.1 Object Model dispatch 공통 인터페이스

모든 OM 객체는 최소한 아래 4가지를 제공해야 한다.

1. type 이름
2. get(property read)
3. set(property write)
4. invoke(method call)

예시:

```rust
pub trait OmObject {
    fn type_name(&self) -> &'static str;
    fn get(&self, member: &str, args: &[OmValue]) -> OmResult<OmValue>;
    fn set(&mut self, member: &str, value: OmValue, args: &[OmValue]) -> OmResult<()>;
    fn invoke(&mut self, member: &str, args: &[OmValue]) -> OmResult<OmValue>;
}
```

이 인터페이스는 public ergonomic API가 아니라 **dispatch bridge**다.

실사용자는 가능하면 generated strong API를 통하게 한다.

### 6.2 Session / Object Registry

VBA/COM 같은 세계에서는 객체 정체성과 active object 상태가 중요하다.
따라서 `Application -> Workbooks -> Workbook -> Worksheet -> Range`가 모두 단순 참조가 아니라 **세션 안의 object handle** 위에 올라타야 한다.

```rust
pub trait OfficeSession {
    fn root_application(&self) -> ObjectHandle;
    fn resolve(&self, handle: ObjectHandle) -> OmResult<&dyn OmObject>;
    fn resolve_mut(&mut self, handle: ObjectHandle) -> OmResult<&mut dyn OmObject>;
    fn dispatch_get(&self, handle: ObjectHandle, member: &str, args: &[OmValue]) -> OmResult<OmValue>;
    fn dispatch_set(&mut self, handle: ObjectHandle, member: &str, value: OmValue, args: &[OmValue]) -> OmResult<()>;
    fn dispatch_invoke(&mut self, handle: ObjectHandle, member: &str, args: &[OmValue]) -> OmResult<OmValue>;
}
```

이 session은 아래를 소유한다.

- workbook registry
- active workbook / active sheet / selection
- event sinks
- undo manager
- calc scheduler handle
- extension service locator

### 6.3 generated contract 인터페이스

`excel-om`은 codegen으로 강한 typed trait을 생성한다.
핵심은 hand-written implementation과 분리하는 것이다.

예시:

```rust
pub trait WorkbookContract {
    fn name(&self) -> OmResult<String>;
    fn worksheets(&self) -> OmResult<ObjectHandle>;
    fn names(&self) -> OmResult<ObjectHandle>;
    fn close(&mut self, save_changes: Option<bool>) -> OmResult<()>;
    fn save(&mut self) -> OmResult<()>;
    fn save_as(&mut self, path: String, file_format: Option<i32>) -> OmResult<()>;
}
```

런타임 구현체는 아래처럼 가진다.

```rust
pub struct WorkbookObject {
    pub handle: ObjectHandle,
    pub workbook_id: WorkbookId,
}

impl WorkbookContract for WorkbookObject {
    /* hand-written implementation */
}

impl OmObject for WorkbookObject {
    /* generated dispatch metadata를 사용해 bridge */
}
```

### 6.4 Excel host orchestration 인터페이스

사용자 친화적인 상위 API는 `excel-runtime`에 둔다.
이건 COM 흉내가 아니라 Rust host API다.

```rust
pub trait ExcelHost {
    fn create_workbook(&mut self) -> OmResult<WorkbookHandle>;
    fn open_workbook(&mut self, spec: OpenWorkbookSpec) -> OmResult<WorkbookHandle>;
    fn save_workbook(&mut self, workbook: WorkbookHandle, spec: SaveWorkbookSpec) -> OmResult<Vec<u8>>;
    fn close_workbook(&mut self, workbook: WorkbookHandle, save: bool) -> OmResult<()>;

    fn calculate(&mut self, scope: CalcScope) -> OmResult<CalcReport>;
    fn render_sheet(&self, spec: RenderSpec) -> OmResult<RenderArtifact>;

    fn get_range_values(&self, spec: GetRangeValuesSpec) -> OmResult<OmArray>;
    fn set_range_values(&mut self, spec: SetRangeValuesSpec) -> OmResult<()>;
}
```

이 인터페이스는 native app, CLI, tests, WASM wrapper가 모두 재사용할 수 있다.

### 6.5 Model query / mutation 인터페이스

`excel-model`은 trait보다 struct 중심이어도 괜찮다. 다만 mutation 경계는 따로 두는 편이 좋다.

```rust
pub trait WorkbookQuery {
    fn workbook(&self, workbook_id: WorkbookId) -> OmResult<&WorkbookModel>;
    fn worksheet(&self, sheet_id: SheetId) -> OmResult<&WorksheetModel>;
    fn cell(&self, sheet_id: SheetId, row: u32, col: u32) -> OmResult<CellSnapshot>;
}

pub trait WorkbookMutation {
    fn set_value(&mut self, sheet_id: SheetId, row: u32, col: u32, value: CellValue) -> OmResult<()>;
    fn set_formula(&mut self, sheet_id: SheetId, row: u32, col: u32, formula: FormulaSource) -> OmResult<()>;
    fn apply_style(&mut self, target: &RangeRef, style_id: StyleId) -> OmResult<()>;
    fn insert_rows(&mut self, sheet_id: SheetId, before: u32, count: u32) -> OmResult<()>;
    fn insert_columns(&mut self, sheet_id: SheetId, before: u32, count: u32) -> OmResult<()>;
}
```

이 mutation 인터페이스는 반드시 다음과 연결되어야 한다.

- calc dirty marking
- dependency graph invalidation
- undo command recording
- table/name/reference retargeting

즉, 실제 public mutation entrypoint는 보통 `excel-runtime`에 두고,
`excel-model` 쪽 mutation은 낮은 레벨 primitive로 두는 것이 좋다.

### 6.6 Formula parser / resolver 인터페이스

`excel-formula`는 다음 3단계를 분리해야 한다.

1. parse text -> AST
2. normalize AST -> bound reference
3. serialize AST -> formula text

```rust
pub trait FormulaParser {
    fn parse_a1(&self, input: &str) -> OmResult<FormulaAst>;
    fn parse_r1c1(&self, input: &str) -> OmResult<FormulaAst>;
    fn print_a1(&self, ast: &FormulaAst) -> OmResult<String>;
    fn print_r1c1(&self, ast: &FormulaAst) -> OmResult<String>;
}

pub trait ReferenceBinder {
    fn bind(&self, workbook_id: WorkbookId, anchor: CellRef, ast: &FormulaAst) -> OmResult<BoundFormula>;
}
```

### 6.7 Calculation engine 인터페이스

`excel-calc`는 최소한 아래를 담당해야 한다.

- dependency graph build/rebuild
- dirty propagation
- partial recalc / full recalc
- evaluator
- function registry/profile
- spill result materialization

```rust
pub trait CalcEngine {
    fn mark_dirty(&mut self, workbook_id: WorkbookId, target: DirtyTarget) -> OmResult<()>;
    fn rebuild_dependencies(&mut self, workbook_id: WorkbookId) -> OmResult<()>;
    fn calculate(&mut self, scope: CalcScope) -> OmResult<CalcReport>;
    fn evaluate_cell(&mut self, workbook_id: WorkbookId, cell: CellRef) -> OmResult<EvaluatedCell>;
}

pub trait FunctionRegistry {
    fn resolve(&self, name: &str, profile: ExcelProfile) -> Option<&FunctionDescriptor>;
}
```

### 6.8 Loader / writer 인터페이스

포맷 codec은 `excel-runtime`과 느슨하게 연결되어야 한다.

```rust
pub trait WorkbookCodec {
    fn sniff(&self, bytes: &[u8]) -> bool;
    fn load(&self, input: WorkbookInput, options: LoadOptions) -> OmResult<LoadedWorkbook>;
    fn save(&self, workbook: &WorkbookModel, options: SaveOptions) -> OmResult<Vec<u8>>;
}
```

하지만 실제로는 workbook 전체만으로 writer를 돌리기 어려울 수 있다.
특히 unknown part 보존, macro-preserving round-trip 때문에 **lossless package view**가 필요하다.

그래서 v0.1에서는 다음 보조 인터페이스를 두는 것이 좋다.

```rust
pub trait OpaquePartStore {
    fn part(&self, uri: &str) -> Option<&OpaquePart>;
    fn parts(&self) -> Vec<&OpaquePart>;
    fn relationship_set(&self, source_uri: &str) -> Option<&OpaqueRelationshipSet>;
}
```

실제 save는 다음 형태가 더 안전하다.

```rust
pub trait WorkbookPackageCodec {
    fn load_package(&self, bytes: &[u8], options: LoadOptions) -> OmResult<LoadedPackageWorkbook>;
    fn save_package(&self, package: &LoadedPackageWorkbook, options: SaveOptions) -> OmResult<Vec<u8>>;
}
```

즉, **canonical model**과 **lossless package**를 둘 다 다뤄야 한다.

### 6.9 Renderer 인터페이스

renderer는 model/calc 결과를 읽고 출력만 담당한다.

```rust
pub trait SheetRenderer {
    fn render(&self, workbook: &WorkbookModel, spec: RenderSpec) -> OmResult<RenderTree>;
}

pub trait RenderBackend {
    fn to_svg(&self, tree: &RenderTree) -> OmResult<String>;
    fn to_html(&self, tree: &RenderTree) -> OmResult<String>;
    fn to_pdf(&self, tree: &RenderTree) -> OmResult<Vec<u8>>;
}
```

초기에는 `RenderTree -> SVG/HTML`까지만 안정화하고,
PDF는 후순위로 두는 것이 좋다.

### 6.10 WASM / TS 경계 인터페이스

WASM은 object-per-call 방식보다 batch API를 기본으로 두는 것이 낫다.

```rust
pub trait WasmFacade {
    fn load_xlsx(&mut self, bytes: &[u8]) -> OmResult<WorkbookHandle>;
    fn save_xlsx(&mut self, workbook: WorkbookHandle) -> OmResult<Vec<u8>>;
    fn get_range_values(&self, workbook: WorkbookHandle, range: RangeRef) -> OmResult<OmArray>;
    fn set_range_values(&mut self, workbook: WorkbookHandle, range: RangeRef, values: OmArray) -> OmResult<()>;
    fn calculate_workbook(&mut self, workbook: WorkbookHandle) -> OmResult<CalcReport>;
    fn render_sheet_svg(&self, spec: RenderSpec) -> OmResult<String>;
}
```

TS wrapper는 여기에 주소 문자열, `Uint8Array`, `readonly unknown[][]` 같은 ergonomic wrapper를 올리면 된다.

### 6.11 Event / undo 인터페이스

런타임은 최소한 아래 추상화를 가져야 한다.

```rust
pub trait EventSink {
    fn on_event(&mut self, event: OfficeEvent);
}

pub trait UndoManager {
    fn push(&mut self, command: UndoCommand);
    fn undo(&mut self) -> OmResult<()>;
    fn redo(&mut self) -> OmResult<()>;
}
```

이건 Excel 이벤트를 전부 복제하는 수준까지 갈 필요는 없다.
v0.1에서는 내부 consistency와 testing을 위해서만 먼저 도입하면 된다.

### 6.12 Oracle / fixture 인터페이스

오라클 테스트는 별도 crate로 떼어내고, 비교 산출물은 JSON schema로 고정한다.

```rust
pub trait ExcelOracleRunner {
    fn open_and_probe(&self, spec: OracleProbeSpec) -> OmResult<OracleWorkbookSnapshot>;
    fn export_pdf(&self, spec: OracleExportSpec) -> OmResult<Vec<u8>>;
}
```

fixture crate는 데이터 자체보다 **manifest**를 정규화하는 역할이 중요하다.

```rust
pub trait FixtureRegistry {
    fn list(&self) -> Vec<FixtureDescriptor>;
    fn open(&self, id: &str) -> OmResult<FixtureArtifactSet>;
}
```

---

## 7. codegen 구조

### 7.1 파이프라인

```text
Excel Type Library / PIA
        ↓ extract
office-idl JSON
        ↓ normalize
member descriptors / enums / class maps
        ↓ generate
excel-om / office-om-common / office-ts
```

### 7.2 generated crate가 내보낼 것

`excel-om`은 최소한 다음을 생성해야 한다.

- enum 정의
- interface trait (`WorkbookContract` 등)
- dispatch descriptor table
- optional parameter metadata
- collection default member metadata
- doc string / source mapping metadata

### 7.3 구현체 연결 방식

generated trait은 hand-written struct가 구현한다.

```text
excel-om                excel-runtime
---------               -----------------
WorkbookContract  <---  WorkbookObject
WorksheetContract <---  WorksheetObject
RangeContract     <---  RangeObject
```

dispatch bridge는 가능하면 codegen이 생성하고,
핵심 의미론만 hand-written으로 남기는 구조가 유지 보수에 유리하다.

---

## 8. 폴더 수준 제안

### 8.1 `excel-runtime`

```text
excel-runtime/
├─ src/
│  ├─ host.rs
│  ├─ session.rs
│  ├─ registry.rs
│  ├─ application.rs
│  ├─ workbook.rs
│  ├─ worksheet.rs
│  ├─ range.rs
│  ├─ names.rs
│  ├─ tables.rs
│  ├─ mutation/
│  ├─ dispatch/
│  ├─ events/
│  └─ undo/
└─ tests/
```

### 8.2 `excel-model`

```text
excel-model/
├─ src/
│  ├─ workbook.rs
│  ├─ worksheet.rs
│  ├─ cell.rs
│  ├─ style.rs
│  ├─ table.rs
│  ├─ names.rs
│  ├─ validation.rs
│  ├─ conditional_format.rs
│  ├─ merge.rs
│  ├─ view.rs
│  └─ opaque/
└─ tests/
```

### 8.3 `excel-xlsx`

```text
excel-xlsx/
├─ src/
│  ├─ load/
│  │  ├─ package.rs
│  │  ├─ workbook.rs
│  │  ├─ worksheet.rs
│  │  ├─ shared_strings.rs
│  │  ├─ styles.rs
│  │  ├─ tables.rs
│  │  └─ drawings.rs
│  ├─ save/
│  ├─ strict/
│  ├─ macro_preserve/
│  └─ lossless/
└─ tests/
```

### 8.4 `excel-calc`

```text
excel-calc/
├─ src/
│  ├─ graph/
│  ├─ evaluator/
│  ├─ functions/
│  ├─ spills/
│  ├─ profile/
│  └─ recalc/
└─ tests/
```

---

## 9. end-to-end 호출 흐름

### 9.1 파일 열기

```text
office-wasm / CLI / test
    ↓
excel-runtime::ExcelHost::open_workbook
    ↓ choose codec
excel-xlsx::load_package
    ↓
excel-model::WorkbookModel 생성
    ↓
excel-runtime session에 WorkbookObject 등록
    ↓
WorkbookHandle 반환
```

### 9.2 `Range.Value2 = ...`

```text
RangeObject.set("Value2", ...)
    ↓
excel-runtime mutation service
    ↓
excel-model cell mutation
    ↓
excel-calc mark_dirty
    ↓
undo command push
    ↓
필요 시 event queue enqueue
```

### 9.3 저장

```text
ExcelHost::save_workbook
    ↓
workbook/session lookup
    ↓
excel-xlsx writer
    ↓ merge canonical model + opaque parts
    ↓
Vec<u8>
```

---

## 10. 초기 public API 권장안

사용자 관점의 첫 public API는 아래 정도면 충분하다.

```rust
let mut host = ExcelEngine::new();
let wb = host.open_workbook(OpenWorkbookSpec::from_bytes(bytes))?;

let values = host.get_range_values(GetRangeValuesSpec::a1(wb, "Sheet1", "A1:C10"))?;
host.set_range_values(SetRangeValuesSpec::a1(wb, "Sheet1", "E1:F2", values2))?;

host.calculate(CalcScope::Workbook(wb))?;
let out = host.save_workbook(wb, SaveWorkbookSpec::xlsx())?;
```

즉, 첫 버전의 product identity는 `COM clone`이 아니라
**Rust/TS에서 쓰기 좋은 Excel compatibility engine**이어야 한다.

그 위에 object model dispatch와 generated OM surface가 얹히는 구조가 좋다.

---

## 11. 마일스톤별 인터페이스 고정 순서

### M0
먼저 고정할 것:
- `WorkbookId`, `SheetId`, `RangeRef`, `CellValue`, `OmValue`
- `OmObject`, `OfficeSession`
- `ExcelHost`
- `WorkbookCodec`
- `CalcEngine`

### M1
그다음 고정할 것:
- `WorkbookModel`, `WorksheetModel` 최소 필드
- `OpaquePartStore`
- `LoadOptions`, `SaveOptions`, `ExcelProfile`

### M2
그다음 고정할 것:
- generated `WorkbookContract`, `WorksheetContract`, `RangeContract`
- event/undo interfaces
- render interfaces

이 순서로 가면 가장 먼저 **공통 골조**가 고정되고,
이후 구현 crate를 병렬로 진행하기 쉬워진다.

---

## 12. 추천 시작점

바로 시작한다면 다음 순서가 가장 안전하다.

1. `office-common`에 id/value/error/address primitive 정의
2. `office-runtime`에 `OmObject`, `OfficeSession`, `ObjectHandle` 정의
3. `excel-model`에 `WorkbookModel`, `WorksheetModel`, `CellRecord` 최소 구조 정의
4. `excel-xlsx`에 `load/save` 최소 round-trip 구현
5. `excel-runtime`에 `ExcelHost` + `WorkbookObject` + `WorksheetObject` + `RangeObject` 최소 구현
6. `excel-calc`에 parser/evaluator basic subset 추가
7. `office-wasm`에 batch API만 먼저 export

이 문서의 기준은 **crate 경계를 먼저 고정하고 구현은 뒤에 얹는다**는 것이다.
초반에 구조를 흔들지 않으려면, runtime/model/calc/codec의 의존 방향을 이 문서대로 엄격히 지키는 편이 좋다.
