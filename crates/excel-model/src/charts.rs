use std::collections::BTreeMap;

use office_common::{
    CellError, CellValue, ChartId, ChartObjectId, DrawingAnchor, DrawingId, DrawingObjectId,
    ExternalReference, FormulaSource, NameScope, ObjectPlacement, OmArray, OmValue, RangeArea,
    RangeSet, Rect, ReferenceTarget, SheetId, SheetScope, WorkbookId, WorksheetModel,
};

use crate::names::DefinedNameTable;

const EXCEL_MAX_ROW_INDEX: u32 = 1_048_576;
const EXCEL_MAX_COLUMN_INDEX: u32 = 16_384;

#[derive(Debug, Clone, PartialEq)]
pub struct ChartModel {
    pub id: ChartId,
    pub workbook_id: WorkbookId,
    pub chart_type: ChartType,
    pub style: Option<u16>,
    pub series: Vec<SeriesModel>,
    pub title: Option<ChartText>,
    pub legend: Option<LegendModel>,
    pub axes: Vec<AxisModel>,
    pub vary_by_categories: Option<bool>,
    pub gap_width: Option<u16>,
    pub gap_depth: Option<u16>,
    pub overlap: Option<i16>,
    pub bar_shape: Option<ChartBarShape>,
    pub has_series_lines: Option<bool>,
    pub has_drop_lines: Option<bool>,
    pub has_hi_lo_lines: Option<bool>,
    pub has_up_down_bars: Option<bool>,
    pub first_slice_angle: Option<u16>,
    pub explosion: Option<u16>,
    pub bubble_scale: Option<u16>,
    pub show_negative_bubbles: Option<bool>,
    pub has_3d_shading: Option<bool>,
    pub doughnut_hole_size: Option<u16>,
    pub second_plot_size: Option<u16>,
    pub size_represents: Option<ChartSizeRepresents>,
    pub split_type: Option<ChartSplitType>,
    pub split_value: Option<f64>,
    pub data_labels: Option<ChartDataLabelsModel>,
    pub data_table: Option<ChartDataTableModel>,
    pub data_table_dirty: bool,
    pub show_data_labels_over_maximum: Option<bool>,
    pub display_blanks_as: Option<ChartDisplayBlanksAs>,
    pub plot_visible_only: Option<bool>,
    pub view_3d: Option<ChartView3DModel>,
    pub view_3d_dirty: bool,
    pub rounded_corners: Option<bool>,
    pub protection: Option<ChartProtectionModel>,
    pub protection_dirty: bool,
    pub raw_part_uri: Option<String>,
    pub dirty: bool,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ChartView3DModel {
    pub elevation: Option<i16>,
    pub height_percent: Option<u16>,
    pub rotation: Option<u16>,
    pub depth_percent: Option<u16>,
    pub right_angle_axes: Option<bool>,
    pub perspective: Option<u16>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct ChartProtectionModel {
    pub drawing_objects: bool,
    pub contents: bool,
    pub data: bool,
    pub formatting: bool,
    pub selection: bool,
    pub user_interface_only: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ChartType {
    Unknown,
    Area,
    Area3D,
    AreaStacked,
    Area3DStacked,
    AreaStacked100,
    Area3DStacked100,
    Bar,
    Bar3DClustered,
    BarStacked,
    Bar3DStacked,
    BarStacked100,
    Bar3DStacked100,
    Column,
    Column3D,
    ColumnStacked,
    Column3DClustered,
    Column3DStacked,
    ColumnStacked100,
    Column3DStacked100,
    CylinderColumn,
    CylinderColumnClustered,
    CylinderColumnStacked,
    CylinderColumnStacked100,
    CylinderBarClustered,
    CylinderBarStacked,
    CylinderBarStacked100,
    ConeColumn,
    ConeColumnClustered,
    ConeColumnStacked,
    ConeColumnStacked100,
    ConeBarClustered,
    ConeBarStacked,
    ConeBarStacked100,
    PyramidColumn,
    PyramidColumnClustered,
    PyramidColumnStacked,
    PyramidColumnStacked100,
    PyramidBarClustered,
    PyramidBarStacked,
    PyramidBarStacked100,
    Line,
    Line3D,
    LineMarkers,
    LineMarkersStacked,
    LineMarkersStacked100,
    LineStacked,
    LineStacked100,
    Scatter,
    ScatterLines,
    ScatterLinesNoMarkers,
    ScatterSmooth,
    ScatterSmoothNoMarkers,
    Bubble,
    Bubble3DEffect,
    Doughnut,
    DoughnutExploded,
    Pie,
    Pie3D,
    PieExploded,
    Pie3DExploded,
    PieOfPie,
    BarOfPie,
    Radar,
    RadarMarkers,
    RadarFilled,
    StockHLC,
    StockOHLC,
    Surface,
    SurfaceWireframe,
    SurfaceTopView,
    SurfaceTopViewWireframe,
    Unsupported(String),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartDisplayBlanksAs {
    Gap,
    Span,
    Zero,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartBarShape {
    Box,
    PyramidToPoint,
    PyramidToMax,
    Cylinder,
    ConeToPoint,
    ConeToMax,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartMarkerStyle {
    Automatic,
    Circle,
    Dash,
    Diamond,
    Dot,
    None,
    Picture,
    Plus,
    Square,
    Star,
    Triangle,
    X,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSizeRepresents {
    Area,
    Width,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartSplitType {
    Custom,
    PercentValue,
    Position,
    Value,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartDataLabelPosition {
    Above,
    Below,
    BestFit,
    Center,
    InsideBase,
    InsideEnd,
    Left,
    OutsideEnd,
    Right,
}

#[derive(Debug, Clone, PartialEq)]
pub struct SeriesModel {
    pub name: Option<ChartSourceExpr>,
    pub x_values: Option<ChartSourceExpr>,
    pub values: Option<ChartSourceExpr>,
    pub bubble_size: Option<ChartSourceExpr>,
    pub bar_shape: Option<ChartBarShape>,
    pub smooth: Option<bool>,
    pub marker_style: Option<ChartMarkerStyle>,
    pub marker_size: Option<u8>,
    pub invert_if_negative: Option<bool>,
    pub points: BTreeMap<u32, ChartPointModel>,
    pub data_labels: Option<ChartDataLabelsModel>,
    pub point_data_labels: BTreeMap<u32, ChartDataLabelsModel>,
    pub order: Option<u32>,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartPointModel {
    pub explosion: Option<u16>,
    pub dirty: bool,
    pub loaded: bool,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartDataLabelsModel {
    pub label_type: Option<i32>,
    pub show_legend_key: Option<bool>,
    pub has_leader_lines: Option<bool>,
    pub show_series_name: Option<bool>,
    pub show_category_name: Option<bool>,
    pub show_value: Option<bool>,
    pub show_percentage: Option<bool>,
    pub show_bubble_size: Option<bool>,
    pub number_format: Option<String>,
    pub number_format_linked: Option<bool>,
    pub position: Option<ChartDataLabelPosition>,
    pub separator: Option<String>,
    pub dirty: bool,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ChartDataTableModel {
    pub has_border_horizontal: Option<bool>,
    pub has_border_vertical: Option<bool>,
    pub has_border_outline: Option<bool>,
    pub show_legend_key: Option<bool>,
    pub dirty: bool,
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
    pub position: Option<ChartLegendPosition>,
    pub include_in_layout: Option<bool>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartLegendPosition {
    Bottom,
    Corner,
    Custom,
    Left,
    Right,
    Top,
}

#[derive(Debug, Clone, PartialEq)]
pub struct AxisModel {
    pub raw_id: Option<String>,
    pub kind: ChartAxisKind,
    pub title: Option<ChartText>,
    pub has_major_gridlines: Option<bool>,
    pub has_minor_gridlines: Option<bool>,
    pub major_tick_mark: Option<ChartTickMark>,
    pub minor_tick_mark: Option<ChartTickMark>,
    pub tick_label_position: Option<ChartTickLabelPosition>,
    pub tick_label_number_format: Option<String>,
    pub tick_label_number_format_linked: Option<bool>,
    pub tick_label_spacing: Option<u32>,
    pub tick_mark_spacing: Option<u32>,
    pub axis_between_categories: Option<bool>,
    pub category_type_auto: Option<bool>,
    pub base_unit: Option<ChartAxisTimeUnit>,
    pub major_unit_scale: Option<ChartAxisTimeUnit>,
    pub minor_unit_scale: Option<ChartAxisTimeUnit>,
    pub display_unit: Option<ChartAxisDisplayUnit>,
    pub has_display_unit_label: Option<bool>,
    pub display_unit_label: Option<ChartText>,
    pub reverse_plot_order: Option<bool>,
    pub scale_type: Option<ChartAxisScaleType>,
    pub log_base: Option<f64>,
    pub crosses: Option<ChartAxisCrosses>,
    pub crosses_at: Option<f64>,
    pub minimum_scale: Option<f64>,
    pub maximum_scale: Option<f64>,
    pub major_unit: Option<f64>,
    pub minor_unit: Option<f64>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartAxisKind {
    Category,
    Value,
    Date,
    Series,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartAxisScaleType {
    Linear,
    Logarithmic,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartAxisCrosses {
    Automatic,
    Custom,
    Maximum,
    Minimum,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartAxisTimeUnit {
    Days,
    Months,
    Years,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub enum ChartAxisDisplayUnit {
    BuiltIn(ChartBuiltInDisplayUnit),
    Custom(f64),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartBuiltInDisplayUnit {
    Hundreds,
    Thousands,
    TenThousands,
    HundredThousands,
    Millions,
    TenMillions,
    HundredMillions,
    ThousandMillions,
    MillionMillions,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartTickMark {
    Cross,
    Inside,
    None,
    Outside,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ChartTickLabelPosition {
    High,
    Low,
    NextToAxis,
    None,
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
    pub anchor_attrs: BTreeMap<String, String>,
    pub position_attrs: BTreeMap<String, String>,
    pub extents_attrs: BTreeMap<String, String>,
    pub marker_attrs: ChartMarkerXmlAttrs,
    pub graphic_frame_attrs: BTreeMap<String, String>,
    pub graphic_frame_transform_xml: Option<String>,
    pub graphic_data_attrs: BTreeMap<String, String>,
    pub graphic_data_child_xmls: Vec<String>,
    pub chart_reference_attrs: BTreeMap<String, String>,
    pub non_visual_frame_attrs: BTreeMap<String, String>,
    pub graphic_attrs: BTreeMap<String, String>,
    pub non_visual_id: Option<u32>,
    pub non_visual_attrs: BTreeMap<String, String>,
    pub non_visual_child_xml: Option<String>,
    pub non_visual_frame_properties_xml: Option<String>,
    pub client_data_attrs: BTreeMap<String, String>,
    pub client_data_xml: Option<String>,
    pub anchor_extension_xmls: Vec<String>,
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

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartMarkerXmlAttrs {
    pub from: ChartCellMarkerXmlAttrs,
    pub to: ChartCellMarkerXmlAttrs,
}

#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct ChartCellMarkerXmlAttrs {
    pub attrs: BTreeMap<String, String>,
    pub col_attrs: BTreeMap<String, String>,
    pub col_offset_attrs: BTreeMap<String, String>,
    pub row_attrs: BTreeMap<String, String>,
    pub row_offset_attrs: BTreeMap<String, String>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ChartSheetBinding {
    pub sheet_id: SheetId,
    pub chart_id: ChartId,
    pub drawing_id: Option<DrawingId>,
    pub raw_part_uri: Option<String>,
}

pub fn resolve_chart_source_reference(
    reference: &str,
    workbook_id: WorkbookId,
    workbook_display_name: Option<&str>,
    worksheets: &[WorksheetModel],
) -> Option<ReferenceTarget> {
    #[derive(Debug, Clone, Copy, PartialEq, Eq)]
    enum ChartA1Endpoint {
        Cell(u32, u32),
        Row(u32),
        Column(u32),
    }

    let parse_column_label_a1 = |input: &str| -> Option<u32> {
        let normalized = input.trim().replace('$', "").to_ascii_uppercase();
        if normalized.is_empty() || !normalized.chars().all(|ch| ch.is_ascii_alphabetic()) {
            return None;
        }

        let mut col = 0u32;
        for ch in normalized.bytes() {
            col = col.checked_mul(26)?.checked_add((ch - b'A' + 1) as u32)?;
        }
        if col == 0 || col > EXCEL_MAX_COLUMN_INDEX {
            return None;
        }
        Some(col)
    };

    let parse_cell_a1 = |input: &str| -> Option<(u32, u32)> {
        let trimmed = input.trim();
        let mut letters = String::new();
        let mut digits = String::new();
        for ch in trimmed.chars() {
            if ch == '$' {
                continue;
            }
            if ch.is_ascii_alphabetic() && digits.is_empty() {
                letters.push(ch.to_ascii_uppercase());
            } else if ch.is_ascii_digit() {
                digits.push(ch);
            } else {
                return None;
            }
        }
        if letters.is_empty() || digits.is_empty() {
            return None;
        }

        let col = parse_column_label_a1(&letters)?;
        let row = digits.parse::<u32>().ok()?;
        if row == 0 || row > EXCEL_MAX_ROW_INDEX {
            return None;
        }
        Some((row, col))
    };

    let parse_a1_endpoint = |input: &str| -> Option<ChartA1Endpoint> {
        let normalized = input.trim().replace('$', "");
        if normalized.is_empty() {
            return None;
        }
        if normalized.chars().all(|ch| ch.is_ascii_digit()) {
            let row = normalized.parse::<u32>().ok()?;
            if row == 0 || row > EXCEL_MAX_ROW_INDEX {
                return None;
            }
            return Some(ChartA1Endpoint::Row(row));
        }
        if normalized.chars().all(|ch| ch.is_ascii_alphabetic()) {
            return Some(ChartA1Endpoint::Column(parse_column_label_a1(input)?));
        }

        let (row, col) = parse_cell_a1(input)?;
        Some(ChartA1Endpoint::Cell(row, col))
    };

    let parse_rect_a1 = |input: &str| -> Option<Rect> {
        let input = input.trim();
        let mut parts = input.split(':');
        let first = parts.next()?;
        let second = parts.next();
        if parts.next().is_some() {
            return None;
        }
        let first = parse_a1_endpoint(first)?;
        let Some(second) = second else {
            return match first {
                ChartA1Endpoint::Cell(row, col) => Some(Rect::single_cell(row, col)),
                ChartA1Endpoint::Row(_) | ChartA1Endpoint::Column(_) => None,
            };
        };
        let second = parse_a1_endpoint(second)?;
        match (first, second) {
            (
                ChartA1Endpoint::Cell(first_row, first_col),
                ChartA1Endpoint::Cell(second_row, second_col),
            ) => Some(Rect {
                row_first: first_row.min(second_row),
                row_last: first_row.max(second_row),
                col_first: first_col.min(second_col),
                col_last: first_col.max(second_col),
            }),
            (ChartA1Endpoint::Row(first_row), ChartA1Endpoint::Row(second_row)) => Some(Rect {
                row_first: first_row.min(second_row),
                row_last: first_row.max(second_row),
                col_first: 1,
                col_last: EXCEL_MAX_COLUMN_INDEX,
            }),
            (ChartA1Endpoint::Column(first_col), ChartA1Endpoint::Column(second_col)) => {
                Some(Rect {
                    row_first: 1,
                    row_last: EXCEL_MAX_ROW_INDEX,
                    col_first: first_col.min(second_col),
                    col_last: first_col.max(second_col),
                })
            }
            _ => None,
        }
    };

    let reference = reference.trim();
    let reference = reference.strip_prefix('=').unwrap_or(reference).trim();
    if reference.is_empty() {
        return None;
    }

    let mut parts = Vec::new();
    let mut start = 0usize;
    let mut in_quote = false;
    let mut chars = reference.char_indices().peekable();
    while let Some((index, ch)) = chars.next() {
        match ch {
            '\'' => {
                if in_quote && chars.peek().is_some_and(|(_, next)| *next == '\'') {
                    chars.next();
                } else {
                    in_quote = !in_quote;
                }
            }
            ',' if !in_quote => {
                let part = reference[start..index].trim();
                if part.is_empty() {
                    return None;
                }
                parts.push(part);
                start = index + ch.len_utf8();
            }
            _ => {}
        }
    }
    if in_quote {
        return None;
    }

    let part = reference[start..].trim();
    if part.is_empty() {
        return None;
    }
    parts.push(part);

    let mut areas = Vec::new();
    for part in parts {
        let part = part.trim();
        let mut in_quote = false;
        let mut separator = None;
        let mut chars = part.char_indices().peekable();
        while let Some((index, ch)) = chars.next() {
            match ch {
                '\'' => {
                    if in_quote && chars.peek().is_some_and(|(_, next)| *next == '\'') {
                        chars.next();
                    } else {
                        in_quote = !in_quote;
                    }
                }
                '!' if !in_quote => separator = Some(index),
                _ => {}
            }
        }
        if in_quote {
            return None;
        }

        let separator = separator?;
        let sheet = part[..separator].trim();
        let area_reference = part[separator + 1..].trim();
        if sheet.is_empty() || area_reference.is_empty() {
            return None;
        }

        let sheet_name = parse_chart_source_sheet_name(sheet, workbook_display_name)?;
        let sheet_id = worksheets
            .iter()
            .find(|worksheet| worksheet.name.eq_ignore_ascii_case(&sheet_name))
            .map(|worksheet| worksheet.id)?;
        let rect = parse_rect_a1(area_reference)?;
        areas.push(RangeArea::new(SheetScope::Single(sheet_id), rect).ok()?);
    }
    RangeSet::new(workbook_id, areas)
        .ok()
        .map(ReferenceTarget::Range)
}

pub fn resolve_chart_source_reference_with_names(
    reference: &str,
    workbook_id: WorkbookId,
    workbook_display_name: Option<&str>,
    worksheets: &[WorksheetModel],
    defined_names: &DefinedNameTable,
    current_sheet: Option<SheetId>,
) -> Option<ReferenceTarget> {
    resolve_chart_source_reference(reference, workbook_id, workbook_display_name, worksheets)
        .or_else(|| resolve_chart_source_literal(reference))
        .or_else(|| {
            resolve_chart_source_defined_name(
                reference,
                workbook_id,
                workbook_display_name,
                worksheets,
                defined_names,
                current_sheet,
            )
        })
}

fn parse_chart_source_sheet_name(
    sheet: &str,
    workbook_display_name: Option<&str>,
) -> Option<String> {
    let mut sheet_name = if sheet.starts_with('\'') {
        let mut output = String::new();
        let mut chars = sheet.char_indices().peekable();
        let (_, first) = chars.next()?;
        if first != '\'' {
            return None;
        }

        let mut parsed = None;
        while let Some((index, ch)) = chars.next() {
            if ch == '\'' {
                if chars.peek().is_some_and(|(_, next)| *next == '\'') {
                    chars.next();
                    output.push('\'');
                } else if sheet[index + ch.len_utf8()..].trim().is_empty() {
                    parsed = Some(output);
                    break;
                } else {
                    return None;
                }
            } else {
                output.push(ch);
            }
        }
        parsed?
    } else {
        if sheet.contains('\'') {
            return None;
        }
        sheet.to_string()
    };
    if let Some(qualified) = sheet_name.strip_prefix('[') {
        let close_index = qualified.find(']')?;
        let source_workbook_name = &qualified[..close_index];
        let unqualified_sheet_name = &qualified[close_index + 1..];
        if source_workbook_name.is_empty()
            || unqualified_sheet_name.is_empty()
            || !workbook_display_name
                .is_some_and(|name| name.eq_ignore_ascii_case(source_workbook_name))
        {
            return None;
        }
        sheet_name = unqualified_sheet_name.to_string();
    }
    if sheet_name.is_empty()
        || sheet_name.contains('[')
        || sheet_name.contains(']')
        || sheet_name.contains(':')
    {
        return None;
    }
    Some(sheet_name)
}

fn resolve_chart_source_literal(reference: &str) -> Option<ReferenceTarget> {
    let parse_scalar_value = |text: &str| -> Option<CellValue> {
        let text = text.trim();
        if text.is_empty() {
            return None;
        }
        if let Some(inner) = text
            .strip_prefix('"')
            .and_then(|text| text.strip_suffix('"'))
        {
            let mut value = String::new();
            let mut chars = inner.chars().peekable();
            while let Some(ch) = chars.next() {
                if ch == '"' {
                    if chars.peek() == Some(&'"') {
                        chars.next();
                        value.push('"');
                    } else {
                        return None;
                    }
                } else {
                    value.push(ch);
                }
            }
            return Some(CellValue::Text(value));
        }
        if text.eq_ignore_ascii_case("TRUE") {
            return Some(CellValue::Bool(true));
        }
        if text.eq_ignore_ascii_case("FALSE") {
            return Some(CellValue::Bool(false));
        }
        let error = match text.to_ascii_uppercase().as_str() {
            "#NULL!" => Some(CellError::Null),
            "#DIV/0!" => Some(CellError::Div0),
            "#VALUE!" => Some(CellError::Value),
            "#REF!" => Some(CellError::Ref),
            "#NAME?" => Some(CellError::Name),
            "#NUM!" => Some(CellError::Num),
            "#N/A" => Some(CellError::NA),
            "#GETTING_DATA" => Some(CellError::GettingData),
            "#SPILL!" => Some(CellError::Spill),
            "#CALC!" => Some(CellError::Calc),
            "#FIELD!" => Some(CellError::Field),
            "#BLOCKED!" => Some(CellError::Blocked),
            _ => None,
        };
        if let Some(error) = error {
            return Some(CellValue::Error(error));
        }
        let number = text.parse::<f64>().ok()?;
        number.is_finite().then_some(CellValue::Number(number))
    };
    let split_array_axis = |text: &str, separator: char| -> Option<Vec<String>> {
        let mut parts = Vec::new();
        let mut start = 0usize;
        let mut in_string = false;
        let mut chars = text.char_indices().peekable();
        while let Some((index, ch)) = chars.next() {
            if ch == '"' {
                if in_string && chars.peek().is_some_and(|(_, next)| *next == '"') {
                    chars.next();
                } else {
                    in_string = !in_string;
                }
            } else if ch == separator && !in_string {
                let part = text[start..index].trim();
                if part.is_empty() {
                    return None;
                }
                parts.push(part.to_string());
                start = index + ch.len_utf8();
            }
        }
        if in_string {
            return None;
        }
        let part = text[start..].trim();
        if part.is_empty() {
            return None;
        }
        parts.push(part.to_string());
        Some(parts)
    };
    let parse_array_value = |text: &str| -> Option<OmArray> {
        let inner = text
            .trim()
            .strip_prefix('{')
            .and_then(|text| text.strip_suffix('}'))?
            .trim();
        if inner.is_empty() {
            return None;
        }
        let rows = split_array_axis(inner, ';')?;
        let mut values = Vec::new();
        let mut cols = None::<usize>;
        for row in &rows {
            let row_values = split_array_axis(row.as_str(), ',')?;
            if row_values.is_empty() {
                return None;
            }
            if let Some(cols) = cols {
                if row_values.len() != cols {
                    return None;
                }
            } else {
                cols = Some(row_values.len());
            }
            for value in row_values {
                values.push(OmValue::from(parse_scalar_value(value.as_str())?));
            }
        }
        OmArray::new(rows.len(), cols?, values).ok()
    };

    let reference = reference.trim();
    let reference = reference.strip_prefix('=').unwrap_or(reference).trim();
    parse_scalar_value(reference)
        .map(ReferenceTarget::Value)
        .or_else(|| parse_array_value(reference).map(ReferenceTarget::Array))
}

fn resolve_chart_source_defined_name(
    reference: &str,
    workbook_id: WorkbookId,
    workbook_display_name: Option<&str>,
    worksheets: &[WorksheetModel],
    defined_names: &DefinedNameTable,
    current_sheet: Option<SheetId>,
) -> Option<ReferenceTarget> {
    let reference = reference.trim();
    let mut reference = reference.strip_prefix('=').unwrap_or(reference).trim();
    if reference.is_empty() || reference.contains(',') || reference.contains(':') {
        return None;
    }

    let mut in_quote = false;
    let mut separator = None;
    let mut chars = reference.char_indices().peekable();
    while let Some((index, ch)) = chars.next() {
        match ch {
            '\'' => {
                if in_quote && chars.peek().is_some_and(|(_, next)| *next == '\'') {
                    chars.next();
                } else {
                    in_quote = !in_quote;
                }
            }
            '!' if !in_quote => separator = Some(index),
            _ => {}
        }
    }
    if in_quote {
        return None;
    }

    let defined_name = if let Some(separator) = separator {
        let sheet = reference[..separator].trim();
        let name = reference[separator + 1..].trim();
        if sheet.is_empty() || name.is_empty() {
            return None;
        }
        let sheet_name = parse_chart_source_sheet_name(sheet, workbook_display_name)?;
        let sheet_id = worksheets
            .iter()
            .find(|worksheet| worksheet.name.eq_ignore_ascii_case(&sheet_name))
            .map(|worksheet| worksheet.id)?;
        defined_names.lookup_in_scope(NameScope::Worksheet(sheet_id), name)?
    } else {
        if let Some(qualified) = reference.strip_prefix('[') {
            let close_index = qualified.find(']')?;
            let source_workbook_name = &qualified[..close_index];
            let unqualified_name = qualified[close_index + 1..].trim();
            if source_workbook_name.is_empty()
                || unqualified_name.is_empty()
                || !workbook_display_name
                    .is_some_and(|name| name.eq_ignore_ascii_case(source_workbook_name))
            {
                return None;
            }
            reference = unqualified_name;
        }
        defined_names.lookup(current_sheet, reference)?
    };

    if let Some(target) = resolve_chart_source_reference(
        &defined_name.refers_to.text,
        workbook_id,
        workbook_display_name,
        worksheets,
    ) {
        return Some(target);
    }

    let refers_to = defined_name
        .refers_to
        .text
        .trim()
        .strip_prefix('=')
        .unwrap_or(defined_name.refers_to.text.trim())
        .trim();
    if let Some(target) = resolve_chart_source_literal(refers_to) {
        return Some(target);
    }
    if refers_to.starts_with('[') {
        return Some(ReferenceTarget::External(ExternalReference {
            text: defined_name.refers_to.text.clone(),
        }));
    }
    if refers_to.is_empty() || refers_to.contains('!') || refers_to.contains(',') {
        return Some(ReferenceTarget::Formula(defined_name.refers_to.clone()));
    }

    let NameScope::Worksheet(sheet_id) = defined_name.scope else {
        return Some(ReferenceTarget::Formula(defined_name.refers_to.clone()));
    };
    let worksheet_name = worksheets
        .iter()
        .find(|worksheet| worksheet.id == sheet_id)
        .map(|worksheet| worksheet.name.as_str())?;
    let qualified_reference = format!("'{}'!{}", worksheet_name.replace('\'', "''"), refers_to);
    if let Some(target) = resolve_chart_source_reference(
        &qualified_reference,
        workbook_id,
        workbook_display_name,
        worksheets,
    ) {
        return Some(target);
    }

    Some(ReferenceTarget::Formula(defined_name.refers_to.clone()))
}

#[cfg(test)]
mod tests {
    use super::*;
    use office_common::{
        CellError, CellValue, NameValidationMode, OmValue, SheetKind, SheetVisibility,
    };

    fn worksheet(id: u32, workbook_id: WorkbookId, name: &str) -> WorksheetModel {
        WorksheetModel {
            id: SheetId(id.into()),
            workbook_id,
            name: name.to_string(),
            kind: SheetKind::Worksheet,
            visibility: SheetVisibility::Visible,
            relationship_id: None,
            part_uri: None,
        }
    }

    #[test]
    fn resolve_chart_source_reference_preserves_multi_area_order() {
        let workbook_id = WorkbookId(3);
        let worksheets = vec![worksheet(7, workbook_id, "Data 2026")];

        let target = resolve_chart_source_reference(
            "'Data 2026'!$C$1,'Data 2026'!$A$1:$A$3",
            workbook_id,
            None,
            &worksheets,
        )
        .expect("chart source target");

        let ReferenceTarget::Range(range) = target else {
            panic!("expected range target");
        };
        assert_eq!(range.workbook_id(), workbook_id);
        assert_eq!(range.areas().len(), 2);
        assert_eq!(range.areas()[0].scope, SheetScope::Single(SheetId(7)));
        assert_eq!(range.areas()[0].rect, Rect::single_cell(1, 3));
        assert_eq!(range.areas()[1].scope, SheetScope::Single(SheetId(7)));
        assert_eq!(
            range.areas()[1].rect,
            Rect {
                row_first: 1,
                row_last: 3,
                col_first: 1,
                col_last: 1,
            }
        );
    }

    #[test]
    fn resolve_chart_source_reference_accepts_sheet_case_and_axis_ranges() {
        let workbook_id = WorkbookId(4);
        let worksheets = vec![worksheet(2, workbook_id, "Data")];

        let target = resolve_chart_source_reference("data!$B:$D", workbook_id, None, &worksheets)
            .expect("column source target");
        let ReferenceTarget::Range(range) = target else {
            panic!("expected range target");
        };
        assert_eq!(
            range.areas()[0].rect,
            Rect {
                row_first: 1,
                row_last: EXCEL_MAX_ROW_INDEX,
                col_first: 2,
                col_last: 4,
            }
        );

        let target = resolve_chart_source_reference("Data!2:4", workbook_id, None, &worksheets)
            .expect("row source target");
        let ReferenceTarget::Range(range) = target else {
            panic!("expected range target");
        };
        assert_eq!(
            range.areas()[0].rect,
            Rect {
                row_first: 2,
                row_last: 4,
                col_first: 1,
                col_last: EXCEL_MAX_COLUMN_INDEX,
            }
        );
    }

    #[test]
    fn resolve_chart_source_reference_rejects_external_and_ambiguous_sources() {
        let workbook_id = WorkbookId(5);
        let worksheets = vec![worksheet(9, workbook_id, "Data")];

        assert!(
            resolve_chart_source_reference("[Book.xlsx]Data!$A$1", workbook_id, None, &worksheets)
                .is_none()
        );
        assert!(
            resolve_chart_source_reference("Data!$A:$1", workbook_id, None, &worksheets).is_none()
        );
        assert!(
            resolve_chart_source_reference("Missing!$A$1", workbook_id, None, &worksheets)
                .is_none()
        );
        assert!(
            resolve_chart_source_reference("'Data' trailing!$A$1", workbook_id, None, &worksheets)
                .is_none()
        );
    }

    #[test]
    fn resolve_chart_source_reference_accepts_current_workbook_qualifier_only() {
        let workbook_id = WorkbookId(6);
        let worksheets = vec![worksheet(11, workbook_id, "Data")];

        let target = resolve_chart_source_reference(
            "[Workbook]Data!$A$1",
            workbook_id,
            Some("Workbook"),
            &worksheets,
        )
        .expect("same workbook source target");
        let ReferenceTarget::Range(range) = target else {
            panic!("expected range target");
        };
        assert_eq!(range.areas()[0].scope, SheetScope::Single(SheetId(11)));
        assert_eq!(range.areas()[0].rect, Rect::single_cell(1, 1));
        assert!(
            resolve_chart_source_reference(
                "[Other.xlsx]Data!$A$1",
                workbook_id,
                Some("Workbook"),
                &worksheets,
            )
            .is_none()
        );
        assert!(
            resolve_chart_source_reference("[Workbook]Data!$A$1", workbook_id, None, &worksheets)
                .is_none()
        );
    }

    #[test]
    fn resolve_chart_source_reference_with_names_resolves_workbook_names() {
        let workbook_id = WorkbookId(7);
        let worksheets = vec![worksheet(2, workbook_id, "Data")];
        let mut defined_names = DefinedNameTable::default();
        defined_names
            .add(
                NameScope::Workbook,
                "SeriesValues",
                FormulaSource {
                    text: "Data!$B$1:$B$3".to_string(),
                    is_r1c1: false,
                },
                NameValidationMode::StrictExcel,
            )
            .expect("add workbook name");

        let target = resolve_chart_source_reference_with_names(
            "[Workbook]SeriesValues",
            workbook_id,
            Some("Workbook"),
            &worksheets,
            &defined_names,
            None,
        )
        .expect("defined name source target");

        let ReferenceTarget::Range(range) = target else {
            panic!("expected range target");
        };
        assert_eq!(range.areas()[0].scope, SheetScope::Single(SheetId(2)));
        assert_eq!(
            range.areas()[0].rect,
            Rect {
                row_first: 1,
                row_last: 3,
                col_first: 2,
                col_last: 2,
            }
        );
        assert!(
            resolve_chart_source_reference_with_names(
                "[Other]SeriesValues",
                workbook_id,
                Some("Workbook"),
                &worksheets,
                &defined_names,
                None,
            )
            .is_none()
        );
    }

    #[test]
    fn resolve_chart_source_reference_with_names_uses_sheet_scope() {
        let workbook_id = WorkbookId(8);
        let worksheets = vec![worksheet(4, workbook_id, "Data 2026")];
        let mut defined_names = DefinedNameTable::default();
        defined_names
            .add(
                NameScope::Workbook,
                "SeriesValues",
                FormulaSource {
                    text: "'Data 2026'!$A$1".to_string(),
                    is_r1c1: false,
                },
                NameValidationMode::StrictExcel,
            )
            .expect("add workbook name");
        defined_names
            .add(
                NameScope::Worksheet(SheetId(4)),
                "SeriesValues",
                FormulaSource {
                    text: "$C$1:$C$3".to_string(),
                    is_r1c1: false,
                },
                NameValidationMode::StrictExcel,
            )
            .expect("add sheet name");

        let target = resolve_chart_source_reference_with_names(
            "SeriesValues",
            workbook_id,
            None,
            &worksheets,
            &defined_names,
            Some(SheetId(4)),
        )
        .expect("current sheet name source target");
        let ReferenceTarget::Range(range) = target else {
            panic!("expected range target");
        };
        assert_eq!(
            range.areas()[0].rect,
            Rect {
                row_first: 1,
                row_last: 3,
                col_first: 3,
                col_last: 3,
            }
        );

        let target = resolve_chart_source_reference_with_names(
            "'Data 2026'!SeriesValues",
            workbook_id,
            None,
            &worksheets,
            &defined_names,
            None,
        )
        .expect("qualified sheet name source target");
        let ReferenceTarget::Range(range) = target else {
            panic!("expected range target");
        };
        assert_eq!(range.areas()[0].scope, SheetScope::Single(SheetId(4)));
        assert_eq!(range.areas()[0].rect.col_first, 3);
    }

    #[test]
    fn resolve_chart_source_reference_with_names_resolves_constant_targets() {
        let workbook_id = WorkbookId(9);
        let worksheets = vec![worksheet(2, workbook_id, "Data")];
        let mut defined_names = DefinedNameTable::default();
        for (name, refers_to) in [
            ("SeriesNumber", "42"),
            ("SeriesText", "\"Revenue \"\"FY26\"\"\""),
            ("SeriesBool", "TRUE"),
            ("SeriesError", "#N/A"),
        ] {
            defined_names
                .add(
                    NameScope::Workbook,
                    name,
                    FormulaSource {
                        text: refers_to.to_string(),
                        is_r1c1: false,
                    },
                    NameValidationMode::StrictExcel,
                )
                .expect("add constant name");
        }

        assert_eq!(
            resolve_chart_source_reference_with_names(
                "SeriesNumber",
                workbook_id,
                None,
                &worksheets,
                &defined_names,
                None,
            ),
            Some(ReferenceTarget::Value(CellValue::Number(42.0)))
        );
        assert_eq!(
            resolve_chart_source_reference_with_names(
                "SeriesText",
                workbook_id,
                None,
                &worksheets,
                &defined_names,
                None,
            ),
            Some(ReferenceTarget::Value(CellValue::Text(
                "Revenue \"FY26\"".to_string()
            )))
        );
        assert_eq!(
            resolve_chart_source_reference_with_names(
                "SeriesBool",
                workbook_id,
                None,
                &worksheets,
                &defined_names,
                None,
            ),
            Some(ReferenceTarget::Value(CellValue::Bool(true)))
        );
        assert_eq!(
            resolve_chart_source_reference_with_names(
                "SeriesError",
                workbook_id,
                None,
                &worksheets,
                &defined_names,
                None,
            ),
            Some(ReferenceTarget::Value(CellValue::Error(CellError::NA)))
        );
    }

    #[test]
    fn resolve_chart_source_reference_with_names_resolves_direct_literal_targets() {
        let workbook_id = WorkbookId(10);
        let worksheets = vec![worksheet(2, workbook_id, "Data")];
        let defined_names = DefinedNameTable::default();

        assert_eq!(
            resolve_chart_source_reference_with_names(
                "=\"Inline \"\"Name\"\"\"",
                workbook_id,
                None,
                &worksheets,
                &defined_names,
                None,
            ),
            Some(ReferenceTarget::Value(CellValue::Text(
                "Inline \"Name\"".to_string()
            )))
        );
        assert_eq!(
            resolve_chart_source_reference_with_names(
                "=FALSE",
                workbook_id,
                None,
                &worksheets,
                &defined_names,
                None,
            ),
            Some(ReferenceTarget::Value(CellValue::Bool(false)))
        );

        let target = resolve_chart_source_reference_with_names(
            "={1,2;\"Q\"\"3\"\"\",#N/A}",
            workbook_id,
            None,
            &worksheets,
            &defined_names,
            None,
        )
        .expect("array target");
        let ReferenceTarget::Array(array) = target else {
            panic!("expected array target");
        };
        assert_eq!(array.rows, 2);
        assert_eq!(array.cols, 2);
        assert_eq!(array.values[0], OmValue::Number(1.0));
        assert_eq!(array.values[1], OmValue::Number(2.0));
        assert_eq!(array.values[2], OmValue::Text("Q\"3\"".to_string()));
        assert_eq!(array.values[3], OmValue::Error(CellError::NA));
    }

    #[test]
    fn resolve_chart_source_reference_with_names_resolves_array_formula_external_targets() {
        let workbook_id = WorkbookId(11);
        let worksheets = vec![worksheet(3, workbook_id, "Data")];
        let mut defined_names = DefinedNameTable::default();
        defined_names
            .add(
                NameScope::Workbook,
                "SeriesArray",
                FormulaSource {
                    text: "{\"Q1\",\"Q2\";1,2}".to_string(),
                    is_r1c1: false,
                },
                NameValidationMode::StrictExcel,
            )
            .expect("add array name");
        defined_names
            .add(
                NameScope::Workbook,
                "SeriesFormula",
                FormulaSource {
                    text: "SUM(Data!$B$1:$B$3)".to_string(),
                    is_r1c1: false,
                },
                NameValidationMode::StrictExcel,
            )
            .expect("add formula name");
        defined_names
            .add(
                NameScope::Workbook,
                "ExternalSeries",
                FormulaSource {
                    text: "[Other.xlsx]Data!$A$1".to_string(),
                    is_r1c1: false,
                },
                NameValidationMode::StrictExcel,
            )
            .expect("add external name");

        let target = resolve_chart_source_reference_with_names(
            "SeriesArray",
            workbook_id,
            None,
            &worksheets,
            &defined_names,
            None,
        )
        .expect("array target");
        let ReferenceTarget::Array(array) = target else {
            panic!("expected array target");
        };
        assert_eq!(array.rows, 2);
        assert_eq!(array.cols, 2);
        assert_eq!(array.values[0], OmValue::Text("Q1".to_string()));
        assert_eq!(array.values[1], OmValue::Text("Q2".to_string()));
        assert_eq!(array.values[2], OmValue::Number(1.0));
        assert_eq!(array.values[3], OmValue::Number(2.0));

        assert_eq!(
            resolve_chart_source_reference_with_names(
                "SeriesFormula",
                workbook_id,
                None,
                &worksheets,
                &defined_names,
                None,
            ),
            Some(ReferenceTarget::Formula(FormulaSource {
                text: "SUM(Data!$B$1:$B$3)".to_string(),
                is_r1c1: false,
            }))
        );
        let Some(ReferenceTarget::External(external)) = resolve_chart_source_reference_with_names(
            "ExternalSeries",
            workbook_id,
            None,
            &worksheets,
            &defined_names,
            None,
        ) else {
            panic!("expected external target");
        };
        assert_eq!(external.text, "[Other.xlsx]Data!$A$1");
    }
}
