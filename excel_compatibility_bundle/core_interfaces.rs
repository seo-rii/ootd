#![allow(dead_code)]

use std::collections::BTreeMap;

pub type OmResult<T> = Result<T, OmError>;

#[derive(Debug, Clone)]
pub struct OmError {
    pub code: OmErrorCode,
    pub message: String,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OmErrorCode {
    InvalidArgument,
    NotFound,
    TypeMismatch,
    Unsupported,
    InvalidState,
    Io,
    Parse,
    Calculation,
    External,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct WorkbookId(pub u64);
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SheetId(pub u64);
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct StyleId(pub u64);
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ObjectHandle(pub u64);
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct WorkbookHandle(pub ObjectHandle);
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct WorksheetHandle(pub ObjectHandle);
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct RangeHandle(pub ObjectHandle);

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ExcelProfile {
    Excel2016,
    Excel2021,
    Excel365,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileFormat {
    Xlsx,
    Xlsm,
    Xltx,
    Xltm,
    StrictXlsx,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellError {
    Null,
    Div0,
    Value,
    Ref,
    Name,
    Num,
    NA,
    GettingData,
    Spill,
    Calc,
    Field,
    Blocked,
    Unknown,
}

#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    Blank,
    Bool(bool),
    Number(f64),
    Text(String),
    Error(CellError),
}

#[derive(Debug, Clone, PartialEq)]
pub enum OmValue {
    Missing,
    Empty,
    Null,
    Bool(bool),
    Number(f64),
    Text(String),
    Error(CellError),
    Object(ObjectHandle),
    Array(OmArray),
}

#[derive(Debug, Clone, PartialEq)]
pub struct OmArray {
    pub rows: usize,
    pub cols: usize,
    pub values: Vec<OmValue>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellRef {
    pub sheet_id: SheetId,
    pub row: u32,
    pub col: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Rect {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SheetScope {
    Single(SheetId),
    Multi3D { start: SheetId, end: SheetId },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RangeRef {
    pub workbook_id: WorkbookId,
    pub scope: SheetScope,
    pub areas: Vec<Rect>,
}

#[derive(Debug, Clone)]
pub struct FormulaSource {
    pub text: String,
    pub is_r1c1: bool,
}

#[derive(Debug, Clone)]
pub struct FormulaAst;

#[derive(Debug, Clone)]
pub struct BoundFormula;

#[derive(Debug, Clone)]
pub struct EvaluatedCell {
    pub cell: CellRef,
    pub value: CellValue,
}

#[derive(Debug, Clone)]
pub struct CalcReport {
    pub recalculated_cells: usize,
    pub spilled_ranges: usize,
    pub warnings: Vec<String>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CalcScope {
    Global,
    Workbook(WorkbookId),
    Sheet(SheetId),
    Cell(CellRef),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DirtyTarget {
    Workbook(WorkbookId),
    Sheet(SheetId),
    Range(RangeRef),
    Cell(CellRef),
}

#[derive(Debug, Clone)]
pub struct FunctionDescriptor {
    pub canonical_name: String,
    pub introduced_in: ExcelProfile,
    pub volatile: bool,
    pub may_spill: bool,
}

#[derive(Debug, Clone)]
pub struct CellSnapshot {
    pub cell: CellRef,
    pub formula: Option<FormulaSource>,
    pub value: CellValue,
    pub style_id: Option<StyleId>,
}

#[derive(Debug, Clone)]
pub struct WorkbookModel {
    pub id: WorkbookId,
    pub display_name: String,
}

#[derive(Debug, Clone)]
pub struct WorksheetModel {
    pub id: SheetId,
    pub workbook_id: WorkbookId,
    pub name: String,
}

#[derive(Debug, Clone)]
pub struct OpaquePart {
    pub uri: String,
    pub content_type: String,
    pub bytes: Vec<u8>,
}

#[derive(Debug, Clone)]
pub struct OpaqueRelationshipSet {
    pub source_uri: String,
    pub relationships: Vec<OpaqueRelationship>,
}

#[derive(Debug, Clone)]
pub struct OpaqueRelationship {
    pub id: String,
    pub kind: String,
    pub target: String,
    pub external: bool,
}

#[derive(Debug, Clone)]
pub struct LoadedWorkbook {
    pub model: WorkbookModel,
    pub opaque_parts: Vec<OpaquePart>,
    pub opaque_relationships: Vec<OpaqueRelationshipSet>,
}

#[derive(Debug, Clone)]
pub struct OpenWorkbookSpec {
    pub bytes: Vec<u8>,
    pub format_hint: Option<FileFormat>,
    pub profile: ExcelProfile,
    pub read_only: bool,
}

#[derive(Debug, Clone)]
pub struct SaveWorkbookSpec {
    pub format: FileFormat,
    pub profile: ExcelProfile,
    pub lossless: bool,
}

#[derive(Debug, Clone)]
pub struct LoadOptions {
    pub profile: ExcelProfile,
    pub preserve_unknown_parts: bool,
    pub read_calc_chain: bool,
}

#[derive(Debug, Clone)]
pub struct SaveOptions {
    pub profile: ExcelProfile,
    pub lossless: bool,
}

#[derive(Debug, Clone)]
pub struct GetRangeValuesSpec {
    pub workbook: WorkbookHandle,
    pub range: RangeRef,
}

#[derive(Debug, Clone)]
pub struct SetRangeValuesSpec {
    pub workbook: WorkbookHandle,
    pub range: RangeRef,
    pub values: OmArray,
}

#[derive(Debug, Clone)]
pub struct RenderSpec {
    pub workbook_id: WorkbookId,
    pub sheet_id: SheetId,
    pub viewport: Rect,
    pub include_gridlines: bool,
    pub include_headers: bool,
}

#[derive(Debug, Clone)]
pub struct RenderTree;

#[derive(Debug, Clone)]
pub enum RenderArtifact {
    Svg(String),
    Html(String),
    Pdf(Vec<u8>),
}

#[derive(Debug, Clone)]
pub struct OfficeEvent {
    pub kind: String,
    pub payload: BTreeMap<String, OmValue>,
}

#[derive(Debug, Clone)]
pub struct UndoCommand {
    pub name: String,
}

#[derive(Debug, Clone)]
pub struct OracleProbeSpec {
    pub path: String,
    pub profile: ExcelProfile,
}

#[derive(Debug, Clone)]
pub struct OracleExportSpec {
    pub path: String,
    pub pdf: bool,
}

#[derive(Debug, Clone)]
pub struct OracleWorkbookSnapshot {
    pub workbook_name: String,
    pub worksheets: Vec<String>,
}

pub trait OmObject {
    fn type_name(&self) -> &'static str;
    fn get(&self, member: &str, args: &[OmValue]) -> OmResult<OmValue>;
    fn set(&mut self, member: &str, value: OmValue, args: &[OmValue]) -> OmResult<()>;
    fn invoke(&mut self, member: &str, args: &[OmValue]) -> OmResult<OmValue>;
}

pub trait OfficeSession {
    fn root_application(&self) -> ObjectHandle;
    fn resolve(&self, handle: ObjectHandle) -> OmResult<&dyn OmObject>;
    fn resolve_mut(&mut self, handle: ObjectHandle) -> OmResult<&mut dyn OmObject>;
    fn dispatch_get(
        &self,
        handle: ObjectHandle,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue>;
    fn dispatch_set(
        &mut self,
        handle: ObjectHandle,
        member: &str,
        value: OmValue,
        args: &[OmValue],
    ) -> OmResult<()>;
    fn dispatch_invoke(
        &mut self,
        handle: ObjectHandle,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue>;
}

pub trait WorkbookContract {
    fn name(&self) -> OmResult<String>;
    fn worksheets(&self) -> OmResult<ObjectHandle>;
    fn save(&mut self) -> OmResult<()>;
    fn save_as(&mut self, path: String, file_format: Option<FileFormat>) -> OmResult<()>;
    fn close(&mut self, save_changes: Option<bool>) -> OmResult<()>;
}

pub trait WorksheetContract {
    fn name(&self) -> OmResult<String>;
    fn range(&self, a1: String) -> OmResult<ObjectHandle>;
    fn cells(&self, row: u32, col: u32) -> OmResult<ObjectHandle>;
}

pub trait RangeContract {
    fn value2(&self) -> OmResult<OmValue>;
    fn set_value2(&mut self, value: OmValue) -> OmResult<()>;
    fn formula(&self) -> OmResult<OmValue>;
    fn set_formula(&mut self, formula: String) -> OmResult<()>;
    fn rows(&self) -> OmResult<ObjectHandle>;
    fn columns(&self) -> OmResult<ObjectHandle>;
    fn address(&self) -> OmResult<String>;
}

pub trait ExcelHost {
    fn create_workbook(&mut self) -> OmResult<WorkbookHandle>;
    fn open_workbook(&mut self, spec: OpenWorkbookSpec) -> OmResult<WorkbookHandle>;
    fn save_workbook(
        &mut self,
        workbook: WorkbookHandle,
        spec: SaveWorkbookSpec,
    ) -> OmResult<Vec<u8>>;
    fn close_workbook(&mut self, workbook: WorkbookHandle, save: bool) -> OmResult<()>;

    fn calculate(&mut self, scope: CalcScope) -> OmResult<CalcReport>;
    fn render_sheet(&self, spec: RenderSpec) -> OmResult<RenderArtifact>;

    fn get_range_values(&self, spec: GetRangeValuesSpec) -> OmResult<OmArray>;
    fn set_range_values(&mut self, spec: SetRangeValuesSpec) -> OmResult<()>;
}

pub trait WorkbookQuery {
    fn workbook(&self, workbook_id: WorkbookId) -> OmResult<&WorkbookModel>;
    fn worksheet(&self, sheet_id: SheetId) -> OmResult<&WorksheetModel>;
    fn cell(&self, sheet_id: SheetId, row: u32, col: u32) -> OmResult<CellSnapshot>;
}

pub trait WorkbookMutation {
    fn set_value(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
        value: CellValue,
    ) -> OmResult<()>;
    fn set_formula(
        &mut self,
        sheet_id: SheetId,
        row: u32,
        col: u32,
        formula: FormulaSource,
    ) -> OmResult<()>;
    fn apply_style(&mut self, target: &RangeRef, style_id: StyleId) -> OmResult<()>;
    fn insert_rows(&mut self, sheet_id: SheetId, before: u32, count: u32) -> OmResult<()>;
    fn insert_columns(&mut self, sheet_id: SheetId, before: u32, count: u32) -> OmResult<()>;
}

pub trait FormulaParser {
    fn parse_a1(&self, input: &str) -> OmResult<FormulaAst>;
    fn parse_r1c1(&self, input: &str) -> OmResult<FormulaAst>;
    fn print_a1(&self, ast: &FormulaAst) -> OmResult<String>;
    fn print_r1c1(&self, ast: &FormulaAst) -> OmResult<String>;
}

pub trait ReferenceBinder {
    fn bind(
        &self,
        workbook_id: WorkbookId,
        anchor: CellRef,
        ast: &FormulaAst,
    ) -> OmResult<BoundFormula>;
}

pub trait CalcEngine {
    fn mark_dirty(&mut self, workbook_id: WorkbookId, target: DirtyTarget) -> OmResult<()>;
    fn rebuild_dependencies(&mut self, workbook_id: WorkbookId) -> OmResult<()>;
    fn calculate(&mut self, scope: CalcScope) -> OmResult<CalcReport>;
    fn evaluate_cell(&mut self, workbook_id: WorkbookId, cell: CellRef) -> OmResult<EvaluatedCell>;
}

pub trait FunctionRegistry {
    fn resolve(&self, name: &str, profile: ExcelProfile) -> Option<&FunctionDescriptor>;
}

pub trait WorkbookCodec {
    fn sniff(&self, bytes: &[u8]) -> bool;
    fn load(&self, input: &[u8], options: LoadOptions) -> OmResult<LoadedWorkbook>;
    fn save(&self, workbook: &WorkbookModel, options: SaveOptions) -> OmResult<Vec<u8>>;
}

pub trait OpaquePartStore {
    fn part(&self, uri: &str) -> Option<&OpaquePart>;
    fn parts(&self) -> Vec<&OpaquePart>;
    fn relationship_set(&self, source_uri: &str) -> Option<&OpaqueRelationshipSet>;
}

pub trait SheetRenderer {
    fn render(&self, workbook: &WorkbookModel, spec: RenderSpec) -> OmResult<RenderTree>;
}

pub trait RenderBackend {
    fn to_svg(&self, tree: &RenderTree) -> OmResult<String>;
    fn to_html(&self, tree: &RenderTree) -> OmResult<String>;
    fn to_pdf(&self, tree: &RenderTree) -> OmResult<Vec<u8>>;
}

pub trait WasmFacade {
    fn load_xlsx(&mut self, bytes: &[u8]) -> OmResult<WorkbookHandle>;
    fn save_xlsx(&mut self, workbook: WorkbookHandle) -> OmResult<Vec<u8>>;
    fn get_range_values(&self, workbook: WorkbookHandle, range: RangeRef) -> OmResult<OmArray>;
    fn set_range_values(
        &mut self,
        workbook: WorkbookHandle,
        range: RangeRef,
        values: OmArray,
    ) -> OmResult<()>;
    fn calculate_workbook(&mut self, workbook: WorkbookHandle) -> OmResult<CalcReport>;
    fn render_sheet_svg(&self, spec: RenderSpec) -> OmResult<String>;
}

pub trait EventSink {
    fn on_event(&mut self, event: OfficeEvent);
}

pub trait UndoManager {
    fn push(&mut self, command: UndoCommand);
    fn undo(&mut self) -> OmResult<()>;
    fn redo(&mut self) -> OmResult<()>;
}

pub trait ExcelOracleRunner {
    fn open_and_probe(&self, spec: OracleProbeSpec) -> OmResult<OracleWorkbookSnapshot>;
    fn export_pdf(&self, spec: OracleExportSpec) -> OmResult<Vec<u8>>;
}

pub trait FixtureRegistry {
    fn list(&self) -> Vec<String>;
    fn open(&self, id: &str) -> OmResult<Vec<u8>>;
}
