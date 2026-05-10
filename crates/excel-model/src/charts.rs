use office_common::{
    ChartId, ChartObjectId, DrawingAnchor, DrawingId, DrawingObjectId, FormulaSource,
    ObjectPlacement, ReferenceTarget, SheetId, WorkbookId,
};

#[derive(Debug, Clone, PartialEq)]
pub struct ChartModel {
    pub id: ChartId,
    pub workbook_id: WorkbookId,
    pub chart_type: ChartType,
    pub series: Vec<SeriesModel>,
    pub title: Option<ChartText>,
    pub legend: Option<LegendModel>,
    pub axes: Vec<AxisModel>,
    pub raw_part_uri: Option<String>,
    pub dirty: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ChartType {
    Unknown,
    Bar,
    Line,
    Scatter,
    Pie,
    Unsupported(String),
}

#[derive(Debug, Clone, PartialEq)]
pub struct SeriesModel {
    pub name: Option<ChartSourceExpr>,
    pub x_values: Option<ChartSourceExpr>,
    pub values: Option<ChartSourceExpr>,
    pub order: Option<u32>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartSourceExpr {
    pub raw: FormulaSource,
    pub resolved: Option<ReferenceTarget>,
    pub cache: Option<ChartCacheSnapshot>,
    pub dirty: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartCacheSnapshot {
    pub kind: ChartCacheKind,
    pub point_count: Option<u32>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ChartCacheKind {
    Number,
    String,
    MultiLevelString,
    Literal,
    Unknown,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartText {
    pub text: String,
}

#[derive(Debug, Clone, PartialEq)]
pub struct LegendModel {
    pub visible: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct AxisModel {
    pub raw_id: Option<String>,
    pub kind: ChartAxisKind,
    pub title: Option<ChartText>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartAxisKind {
    Category,
    Value,
    Date,
    Series,
}

#[derive(Debug, Clone, PartialEq)]
pub struct DrawingModel {
    pub id: DrawingId,
    pub workbook_id: WorkbookId,
    pub host_sheet_id: SheetId,
    pub objects: Vec<DrawingObjectModel>,
    pub raw_part_uri: Option<String>,
    pub dirty: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum DrawingObjectModel {
    ChartFrame(ChartObjectModel),
    UnsupportedRaw {
        id: DrawingObjectId,
        raw_part_uri: Option<String>,
    },
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartObjectModel {
    pub id: ChartObjectId,
    pub workbook_id: WorkbookId,
    pub host_sheet_id: SheetId,
    pub chart_id: ChartId,
    pub name: String,
    pub anchor: Option<DrawingAnchor>,
    pub placement: ObjectPlacement,
    pub z_order: Option<u32>,
    pub raw_binding: Option<String>,
    pub dirty: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartSheetBinding {
    pub sheet_id: SheetId,
    pub chart_id: ChartId,
    pub drawing_id: Option<DrawingId>,
    pub raw_part_uri: Option<String>,
}
