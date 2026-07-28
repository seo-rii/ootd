use std::error::Error;
use std::fmt::{Display, Formatter};

use serde::{Deserialize, Serialize};

mod name;
mod reference;

pub use name::{
    BuiltinName, DefinedName, DefinedNameId, DefinedNameKind, DefinedNameMetadata, NameKey,
    NameScope, NameValidationMode, canonicalize_excel_name,
};
pub use reference::{ExternalReference, RangeArea, RangeSet, ReferenceTarget};

pub type OmResult<T> = Result<T, OmError>;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum OmErrorCode {
    InvalidArgument,
    NotFound,
    TypeMismatch,
    Unsupported,
    InvalidState,
    Io,
    Parse,
    ResourceLimit,
    EncryptedWorkbookUnsupported,
    SignedPackageMutationUnsupported,
    ActiveContentConversionUnsupported,
    Calculation,
    External,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OmError {
    pub code: OmErrorCode,
    pub message: String,
}

impl OmError {
    pub fn new(code: OmErrorCode, message: impl Into<String>) -> Self {
        Self {
            code,
            message: message.into(),
        }
    }

    pub fn invalid_argument(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::InvalidArgument, message)
    }

    pub fn invalid_state(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::InvalidState, message)
    }

    pub fn io(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::Io, message)
    }

    pub fn parse(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::Parse, message)
    }

    pub fn resource_limit(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::ResourceLimit, message)
    }

    pub fn encrypted_workbook_unsupported(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::EncryptedWorkbookUnsupported, message)
    }

    pub fn signed_package_mutation_unsupported(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::SignedPackageMutationUnsupported, message)
    }

    pub fn active_content_conversion_unsupported(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::ActiveContentConversionUnsupported, message)
    }

    pub fn unsupported(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::Unsupported, message)
    }
}

impl Display for OmError {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        write!(f, "{:?}: {}", self.code, self.message)
    }
}

impl Error for OmError {}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct WorkbookId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct SheetId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct StyleId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct ChartId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct ChartObjectId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct DrawingId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct DrawingObjectId(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct Emu(pub i64);

impl Emu {
    pub fn to_points(self) -> Points {
        Points(self.0 as f64 / 12_700.0)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Serialize, Deserialize)]
pub struct Points(pub f64);

impl Points {
    pub fn to_emu(self) -> Emu {
        Emu((self.0 * 12_700.0).round() as i64)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct ObjectHandle(pub u64);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct WorkbookHandle(pub ObjectHandle);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct WorksheetHandle(pub ObjectHandle);

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
pub struct RangeHandle(pub ObjectHandle);

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum ExcelProfile {
    Excel2016,
    Excel2021,
    #[default]
    Excel365,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum FileFormat {
    Xlsx,
    Xlsm,
    Xltx,
    Xltm,
    StrictXlsx,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
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
    Busy,
    Connect,
    Python,
    Timeout,
    Unknown,
}

#[derive(Debug, Clone, Default, PartialEq, Serialize, Deserialize)]
pub enum CellValue {
    #[default]
    Blank,
    Bool(bool),
    Number(f64),
    Text(String),
    Error(CellError),
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
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

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct OmArray {
    pub rows: usize,
    pub cols: usize,
    pub values: Vec<OmValue>,
}

impl OmArray {
    pub fn new(rows: usize, cols: usize, values: Vec<OmValue>) -> OmResult<Self> {
        if rows.saturating_mul(cols) != values.len() {
            return Err(OmError::invalid_argument(
                "array dimensions do not match the number of values",
            ));
        }

        Ok(Self { rows, cols, values })
    }

    pub fn scalar(value: OmValue) -> Self {
        Self {
            rows: 1,
            cols: 1,
            values: vec![value],
        }
    }

    pub fn get(&self, row: usize, col: usize) -> Option<&OmValue> {
        if row >= self.rows || col >= self.cols {
            return None;
        }

        self.values.get(row * self.cols + col)
    }
}

impl From<CellValue> for OmValue {
    fn from(value: CellValue) -> Self {
        match value {
            CellValue::Blank => OmValue::Empty,
            CellValue::Bool(value) => OmValue::Bool(value),
            CellValue::Number(value) => OmValue::Number(value),
            CellValue::Text(value) => OmValue::Text(value),
            CellValue::Error(value) => OmValue::Error(value),
        }
    }
}

impl TryFrom<OmValue> for CellValue {
    type Error = OmError;

    fn try_from(value: OmValue) -> OmResult<Self> {
        match value {
            OmValue::Missing | OmValue::Empty | OmValue::Null => Ok(CellValue::Blank),
            OmValue::Bool(value) => Ok(CellValue::Bool(value)),
            OmValue::Number(value) => Ok(CellValue::Number(value)),
            OmValue::Text(value) => Ok(CellValue::Text(value)),
            OmValue::Error(value) => Ok(CellValue::Error(value)),
            OmValue::Object(_) | OmValue::Array(_) => Err(OmError::type_mismatch(
                "cannot coerce object or array values into a worksheet cell",
            )),
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub struct CellRef {
    pub sheet_id: SheetId,
    pub row: u32,
    pub col: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct Rect {
    pub row_first: u32,
    pub row_last: u32,
    pub col_first: u32,
    pub col_last: u32,
}

impl Rect {
    pub fn single_cell(row: u32, col: u32) -> Self {
        Self {
            row_first: row,
            row_last: row,
            col_first: col,
            col_last: col,
        }
    }

    pub fn width(&self) -> u32 {
        self.col_last.saturating_sub(self.col_first) + 1
    }

    pub fn height(&self) -> u32 {
        self.row_last.saturating_sub(self.row_first) + 1
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum SheetScope {
    Single(SheetId),
    Multi3D { start: SheetId, end: SheetId },
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct RangeRef {
    pub workbook_id: WorkbookId,
    pub scope: SheetScope,
    pub areas: Vec<Rect>,
}

impl RangeRef {
    pub fn single_cell(workbook_id: WorkbookId, sheet_id: SheetId, row: u32, col: u32) -> Self {
        Self {
            workbook_id,
            scope: SheetScope::Single(sheet_id),
            areas: vec![Rect::single_cell(row, col)],
        }
    }

    pub fn single_rect(workbook_id: WorkbookId, sheet_id: SheetId, rect: Rect) -> Self {
        Self {
            workbook_id,
            scope: SheetScope::Single(sheet_id),
            areas: vec![rect],
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct FormulaSource {
    pub text: String,
    pub is_r1c1: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct WorkbookModel {
    pub id: WorkbookId,
    pub display_name: String,
    pub format: FileFormat,
    #[serde(default)]
    pub date1904: bool,
    #[serde(default)]
    pub is_addin: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum SheetVisibility {
    #[default]
    Visible,
    Hidden,
    VeryHidden,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum SheetKind {
    #[default]
    Worksheet,
    ChartSheet,
    MacroSheet,
    DialogSheet,
}

impl SheetKind {
    pub fn is_worksheet(&self) -> bool {
        *self == Self::Worksheet
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize, Default)]
pub enum ObjectPlacement {
    MoveAndSize,
    MoveOnly,
    FreeFloating,
    #[default]
    Unknown,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub enum DrawingAnchor {
    TwoCell(TwoCellAnchor),
    OneCell(OneCellAnchor),
    Absolute(AbsoluteAnchor),
    UnsupportedRaw,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct TwoCellAnchor {
    pub from: CellMarker,
    pub to: CellMarker,
    pub edit_as: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OneCellAnchor {
    pub from: CellMarker,
    pub extents: SizeEmu,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct AbsoluteAnchor {
    pub position: PointEmu,
    pub extents: SizeEmu,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct CellMarker {
    pub col_zero_based: u32,
    pub col_offset: Emu,
    pub row_zero_based: u32,
    pub row_offset: Emu,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct PointEmu {
    pub x: Emu,
    pub y: Emu,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub struct SizeEmu {
    pub cx: Emu,
    pub cy: Emu,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct WorksheetModel {
    pub id: SheetId,
    pub workbook_id: WorkbookId,
    pub name: String,
    #[serde(default, skip_serializing_if = "SheetKind::is_worksheet")]
    pub kind: SheetKind,
    #[serde(default)]
    pub visibility: SheetVisibility,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub relationship_id: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub part_uri: Option<String>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OpaquePart {
    pub uri: String,
    pub content_type: String,
    pub bytes: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct OpenWorkbookSpec {
    pub bytes: Vec<u8>,
    pub format_hint: Option<FileFormat>,
    pub profile: ExcelProfile,
    pub read_only: bool,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SaveWorkbookSpec {
    pub format: FileFormat,
    pub profile: ExcelProfile,
    pub lossless: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct GetRangeValuesSpec {
    pub workbook: WorkbookHandle,
    pub range: RangeRef,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
pub struct SetRangeValuesSpec {
    pub workbook: WorkbookHandle,
    pub range: RangeRef,
    pub values: OmArray,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct LoadOptions {
    pub profile: ExcelProfile,
    pub preserve_unknown_parts: bool,
    pub read_calc_chain: bool,
}

impl Default for LoadOptions {
    fn default() -> Self {
        Self {
            profile: ExcelProfile::default(),
            preserve_unknown_parts: true,
            read_calc_chain: true,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
pub struct SaveOptions {
    pub profile: ExcelProfile,
    pub lossless: bool,
}

impl Default for SaveOptions {
    fn default() -> Self {
        Self {
            profile: ExcelProfile::default(),
            lossless: true,
        }
    }
}

impl OmError {
    pub fn type_mismatch(message: impl Into<String>) -> Self {
        Self::new(OmErrorCode::TypeMismatch, message)
    }
}

#[cfg(test)]
mod tests {
    use super::{
        CellValue, Emu, ExcelProfile, LoadOptions, ObjectHandle, OmArray, OmErrorCode, OmValue,
        Points, RangeRef, Rect, SaveOptions, SheetId, SheetScope, WorkbookId,
    };

    #[test]
    fn om_array_rejects_invalid_dimensions() {
        let error =
            OmArray::new(2, 2, vec![OmValue::Empty]).expect_err("invalid dimensions should fail");
        assert_eq!(error.code, OmErrorCode::InvalidArgument);
    }

    #[test]
    fn rect_single_cell_has_unit_extent() {
        let rect = Rect::single_cell(3, 7);
        assert_eq!(rect.width(), 1);
        assert_eq!(rect.height(), 1);
        assert_eq!(rect.row_first, 3);
        assert_eq!(rect.col_first, 7);
    }

    #[test]
    fn range_ref_single_cell_uses_single_sheet_scope() {
        let range = RangeRef::single_cell(WorkbookId(1), SheetId(2), 4, 5);
        assert_eq!(range.areas, vec![Rect::single_cell(4, 5)]);
        assert_eq!(range.scope, SheetScope::Single(SheetId(2)));
    }

    #[test]
    fn cell_value_round_trips_through_om_value() {
        let value = CellValue::Text("hello".to_string());
        let om_value = OmValue::from(value.clone());
        let restored = CellValue::try_from(om_value).expect("convert back to cell");

        assert_eq!(restored, value);
    }

    #[test]
    fn om_value_array_cannot_become_cell_value() {
        let error = CellValue::try_from(OmValue::Array(OmArray::scalar(OmValue::Number(1.0))))
            .expect_err("array coercion should fail");
        assert_eq!(error.code, OmErrorCode::TypeMismatch);
    }

    #[test]
    fn om_value_object_cannot_become_cell_value() {
        let error = CellValue::try_from(OmValue::Object(ObjectHandle(7)))
            .expect_err("object coercion should fail");
        assert_eq!(error.code, OmErrorCode::TypeMismatch);
    }

    #[test]
    fn missing_empty_and_null_om_values_coerce_to_blank_cells() {
        assert_eq!(
            CellValue::try_from(OmValue::Missing).expect("missing"),
            CellValue::Blank
        );
        assert_eq!(
            CellValue::try_from(OmValue::Empty).expect("empty"),
            CellValue::Blank
        );
        assert_eq!(
            CellValue::try_from(OmValue::Null).expect("null"),
            CellValue::Blank
        );
    }

    #[test]
    fn om_array_get_returns_none_for_out_of_bounds_indices() {
        let array = OmArray::new(
            2,
            2,
            vec![
                OmValue::Number(1.0),
                OmValue::Number(2.0),
                OmValue::Number(3.0),
                OmValue::Number(4.0),
            ],
        )
        .expect("array");

        assert_eq!(array.get(0, 0), Some(&OmValue::Number(1.0)));
        assert_eq!(array.get(2, 0), None);
        assert_eq!(array.get(0, 2), None);
    }

    #[test]
    fn default_load_and_save_options_preserve_lossless_profile_defaults() {
        let load = LoadOptions::default();
        let save = SaveOptions::default();

        assert_eq!(load.profile, ExcelProfile::Excel365);
        assert!(load.preserve_unknown_parts);
        assert!(load.read_calc_chain);
        assert_eq!(save.profile, ExcelProfile::Excel365);
        assert!(save.lossless);
    }

    #[test]
    fn emu_and_points_convert_with_excel_geometry_scale() {
        assert_eq!(Emu(12_700).to_points(), Points(1.0));
        assert_eq!(Points(1.5).to_emu(), Emu(19_050));
    }
}
