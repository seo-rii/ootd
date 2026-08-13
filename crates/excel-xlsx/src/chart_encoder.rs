use std::collections::{BTreeMap, BTreeSet};
use std::io::{Cursor, Write};

use excel_model::{
    AxisModel, ChartAxisCrosses, ChartAxisDisplayUnit, ChartAxisGroup, ChartAxisKind,
    ChartAxisScaleType, ChartAxisTimeUnit, ChartBarShape, ChartBuiltInDisplayUnit,
    ChartDataLabelPosition, ChartDataLabelsModel, ChartDataTableModel, ChartDisplayBlanksAs,
    ChartGroupModel, ChartLayoutMode, ChartLayoutTarget, ChartLegendPosition, ChartManualLayout,
    ChartMarkerStyle, ChartModel, ChartProtectionModel, ChartSizeRepresents, ChartSourceExpr,
    ChartSplitType, ChartTickLabelPosition, ChartTickMark, ChartType, ChartView3DModel,
    SeriesModel,
};
use office_common::{
    CellError, CellValue, OmError, OmErrorCode, OmResult, OmValue, ReferenceTarget,
};
use quick_xml::escape::partial_escape;
use quick_xml::events::{BytesEnd, BytesRef, BytesStart, BytesText, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, Reader, Writer};

const CHART_XML_NAMESPACE: &[u8] = b"http://schemas.openxmlformats.org/drawingml/2006/chart";

/// Encodes a chart model by patching loaded XML when possible and serializing a new part otherwise.
pub fn encode_chart_model_xml(source_xml: Option<&[u8]>, chart: &ChartModel) -> OmResult<Vec<u8>> {
    let candidate_xml = if let Some(source_xml) = source_xml {
        match patch_loaded_chart_model_xml(source_xml, chart)? {
            Some(candidate_xml) => candidate_xml,
            None => serialize_chart_model_xml(chart)?,
        }
    } else {
        serialize_chart_model_xml(chart)?
    };

    let mut reader = Reader::from_reader(Cursor::new(candidate_xml.as_slice()));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(chart_xml_error(error)),
        }
        buffer.clear();
    }

    let converged_xml = rewrite_loaded_chart_axis_additions(candidate_xml.as_slice(), chart)?;
    if converged_xml != candidate_xml {
        let difference = candidate_xml
            .iter()
            .zip(converged_xml.iter())
            .position(|(candidate, converged)| candidate != converged)
            .unwrap_or_else(|| candidate_xml.len().min(converged_xml.len()));
        let context_start = difference.saturating_sub(80);
        let candidate_context_end = (difference + 160).min(candidate_xml.len());
        let converged_context_end = (difference + 160).min(converged_xml.len());
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            format!(
                "chart XML encoder did not converge at byte {difference}: candidate={:?}, converged={:?}",
                String::from_utf8_lossy(&candidate_xml[context_start..candidate_context_end]),
                String::from_utf8_lossy(&converged_xml[context_start..converged_context_end]),
            ),
        ));
    }

    Ok(candidate_xml)
}

fn xml_local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn format_find_number(number: f64) -> String {
    if number.is_finite()
        && number.fract() == 0.0
        && number >= i64::MIN as f64
        && number <= i64::MAX as f64
    {
        (number as i64).to_string()
    } else {
        number.to_string()
    }
}

fn find_cell_value_text(value: &CellValue) -> String {
    match value {
        CellValue::Blank => String::new(),
        CellValue::Bool(true) => "TRUE".to_string(),
        CellValue::Bool(false) => "FALSE".to_string(),
        CellValue::Number(number) => format_find_number(*number),
        CellValue::Text(text) => text.clone(),
        CellValue::Error(error) => formula_cell_error_text(error).to_string(),
        CellValue::IsoDateTime(value) => value.as_str().to_string(),
        CellValue::RichText(value) => value.as_str().to_string(),
    }
}

fn chart_xml_error(error: impl std::fmt::Display) -> OmError {
    OmError::new(OmErrorCode::Parse, error.to_string())
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ChartSourceXmlSlot {
    Name,
    XValues,
    Values,
    BubbleSize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ChartTextXmlTarget {
    ChartTitle,
    AxisTitle(usize),
}

fn chart_type_from_group_name(local_name: &[u8]) -> Option<ChartType> {
    match local_name {
        b"areaChart" => Some(ChartType::Area),
        b"area3DChart" => Some(ChartType::Area3D),
        b"barChart" => Some(ChartType::Bar),
        b"bar3DChart" => Some(ChartType::Bar3DClustered),
        b"lineChart" => Some(ChartType::Line),
        b"line3DChart" => Some(ChartType::Line3D),
        b"scatterChart" => Some(ChartType::Scatter),
        b"bubbleChart" => Some(ChartType::Bubble),
        b"doughnutChart" => Some(ChartType::Doughnut),
        b"ofPieChart" => Some(ChartType::PieOfPie),
        b"pieChart" => Some(ChartType::Pie),
        b"pie3DChart" => Some(ChartType::Pie3D),
        b"radarChart" => Some(ChartType::Radar),
        b"stockChart" => Some(ChartType::StockHLC),
        b"surfaceChart" => Some(ChartType::SurfaceTopView),
        b"surface3DChart" => Some(ChartType::Surface),
        _ => None,
    }
}

fn chart_group_xml_name(chart_type: &ChartType) -> Option<&'static str> {
    match chart_type {
        ChartType::Area | ChartType::AreaStacked | ChartType::AreaStacked100 => Some("areaChart"),
        ChartType::Area3D | ChartType::Area3DStacked | ChartType::Area3DStacked100 => {
            Some("area3DChart")
        }
        ChartType::Bar
        | ChartType::BarStacked
        | ChartType::BarStacked100
        | ChartType::Column
        | ChartType::ColumnStacked
        | ChartType::ColumnStacked100 => Some("barChart"),
        ChartType::Bar3DClustered
        | ChartType::Bar3DStacked
        | ChartType::Bar3DStacked100
        | ChartType::Column3D
        | ChartType::Column3DClustered
        | ChartType::Column3DStacked
        | ChartType::Column3DStacked100
        | ChartType::CylinderColumn
        | ChartType::CylinderColumnClustered
        | ChartType::CylinderColumnStacked
        | ChartType::CylinderColumnStacked100
        | ChartType::CylinderBarClustered
        | ChartType::CylinderBarStacked
        | ChartType::CylinderBarStacked100
        | ChartType::ConeColumn
        | ChartType::ConeColumnClustered
        | ChartType::ConeColumnStacked
        | ChartType::ConeColumnStacked100
        | ChartType::ConeBarClustered
        | ChartType::ConeBarStacked
        | ChartType::ConeBarStacked100
        | ChartType::PyramidColumn
        | ChartType::PyramidColumnClustered
        | ChartType::PyramidColumnStacked
        | ChartType::PyramidColumnStacked100
        | ChartType::PyramidBarClustered
        | ChartType::PyramidBarStacked
        | ChartType::PyramidBarStacked100 => Some("bar3DChart"),
        ChartType::Line
        | ChartType::LineMarkers
        | ChartType::LineMarkersStacked
        | ChartType::LineMarkersStacked100
        | ChartType::LineStacked
        | ChartType::LineStacked100 => Some("lineChart"),
        ChartType::Line3D => Some("line3DChart"),
        ChartType::Scatter
        | ChartType::ScatterLines
        | ChartType::ScatterLinesNoMarkers
        | ChartType::ScatterSmooth
        | ChartType::ScatterSmoothNoMarkers => Some("scatterChart"),
        ChartType::Bubble | ChartType::Bubble3DEffect => Some("bubbleChart"),
        ChartType::Doughnut | ChartType::DoughnutExploded => Some("doughnutChart"),
        ChartType::Pie | ChartType::PieExploded => Some("pieChart"),
        ChartType::Pie3D | ChartType::Pie3DExploded => Some("pie3DChart"),
        ChartType::PieOfPie | ChartType::BarOfPie => Some("ofPieChart"),
        ChartType::Radar | ChartType::RadarMarkers | ChartType::RadarFilled => Some("radarChart"),
        ChartType::StockHLC
        | ChartType::StockOHLC
        | ChartType::StockVHLC
        | ChartType::StockVOHLC => Some("stockChart"),
        ChartType::Surface | ChartType::SurfaceWireframe => Some("surface3DChart"),
        ChartType::SurfaceTopView | ChartType::SurfaceTopViewWireframe => Some("surfaceChart"),
        ChartType::Unknown | ChartType::Unsupported(_) => None,
    }
}

fn chart_filtered_series_wrapper_name(chart_type: &ChartType) -> Option<&'static str> {
    match chart_type {
        ChartType::Area
        | ChartType::Area3D
        | ChartType::AreaStacked
        | ChartType::Area3DStacked
        | ChartType::AreaStacked100
        | ChartType::Area3DStacked100 => Some("filteredAreaSeries"),
        ChartType::Bar
        | ChartType::Bar3DClustered
        | ChartType::BarStacked
        | ChartType::Bar3DStacked
        | ChartType::BarStacked100
        | ChartType::Bar3DStacked100
        | ChartType::Column
        | ChartType::Column3D
        | ChartType::Column3DClustered
        | ChartType::ColumnStacked
        | ChartType::Column3DStacked
        | ChartType::ColumnStacked100
        | ChartType::Column3DStacked100
        | ChartType::CylinderColumn
        | ChartType::CylinderColumnClustered
        | ChartType::CylinderColumnStacked
        | ChartType::CylinderColumnStacked100
        | ChartType::CylinderBarClustered
        | ChartType::CylinderBarStacked
        | ChartType::CylinderBarStacked100
        | ChartType::ConeColumn
        | ChartType::ConeColumnClustered
        | ChartType::ConeColumnStacked
        | ChartType::ConeColumnStacked100
        | ChartType::ConeBarClustered
        | ChartType::ConeBarStacked
        | ChartType::ConeBarStacked100
        | ChartType::PyramidColumn
        | ChartType::PyramidColumnClustered
        | ChartType::PyramidColumnStacked
        | ChartType::PyramidColumnStacked100
        | ChartType::PyramidBarClustered
        | ChartType::PyramidBarStacked
        | ChartType::PyramidBarStacked100 => Some("filteredBarSeries"),
        ChartType::Line
        | ChartType::Line3D
        | ChartType::LineMarkers
        | ChartType::LineMarkersStacked
        | ChartType::LineMarkersStacked100
        | ChartType::LineStacked
        | ChartType::LineStacked100 => Some("filteredLineSeries"),
        ChartType::Scatter
        | ChartType::ScatterLines
        | ChartType::ScatterLinesNoMarkers
        | ChartType::ScatterSmooth
        | ChartType::ScatterSmoothNoMarkers => Some("filteredScatterSeries"),
        ChartType::Bubble | ChartType::Bubble3DEffect => Some("filteredBubbleSeries"),
        ChartType::Doughnut
        | ChartType::DoughnutExploded
        | ChartType::Pie
        | ChartType::Pie3D
        | ChartType::PieExploded
        | ChartType::Pie3DExploded
        | ChartType::PieOfPie
        | ChartType::BarOfPie => Some("filteredPieSeries"),
        ChartType::Radar | ChartType::RadarMarkers | ChartType::RadarFilled => {
            Some("filteredRadarSeries")
        }
        ChartType::Surface
        | ChartType::SurfaceWireframe
        | ChartType::SurfaceTopView
        | ChartType::SurfaceTopViewWireframe => Some("filteredSurfaceSeries"),
        ChartType::StockHLC
        | ChartType::StockOHLC
        | ChartType::StockVHLC
        | ChartType::StockVOHLC
        | ChartType::Unknown
        | ChartType::Unsupported(_) => None,
    }
}

fn chart_type_is_volume_stock(chart_type: &ChartType) -> bool {
    matches!(chart_type, ChartType::StockVHLC | ChartType::StockVOHLC)
}

fn volume_stock_series_count(chart_type: &ChartType) -> Option<usize> {
    match chart_type {
        ChartType::StockVHLC => Some(4),
        ChartType::StockVOHLC => Some(5),
        _ => None,
    }
}

fn chart_type_for_axis_group(chart: &ChartModel, axis_group: ChartAxisGroup) -> ChartType {
    match (&chart.chart_type, axis_group) {
        (ChartType::StockVHLC | ChartType::StockVOHLC, ChartAxisGroup::Primary) => {
            ChartType::Column
        }
        (ChartType::StockVHLC, ChartAxisGroup::Secondary) => ChartType::StockHLC,
        (ChartType::StockVOHLC, ChartAxisGroup::Secondary) => ChartType::StockOHLC,
        _ => chart.chart_type.clone(),
    }
}

fn chart_type_uses_xy_values(chart_type: &ChartType) -> bool {
    matches!(
        chart_type,
        ChartType::Scatter
            | ChartType::ScatterLines
            | ChartType::ScatterLinesNoMarkers
            | ChartType::ScatterSmooth
            | ChartType::ScatterSmoothNoMarkers
            | ChartType::Bubble
            | ChartType::Bubble3DEffect
    )
}

fn chart_type_uses_bubble_size(chart_type: &ChartType) -> bool {
    matches!(chart_type, ChartType::Bubble | ChartType::Bubble3DEffect)
}

fn chart_type_bar_direction_xml_value(chart_type: &ChartType) -> Option<&'static str> {
    match chart_type {
        ChartType::Bar
        | ChartType::BarStacked
        | ChartType::BarStacked100
        | ChartType::Bar3DClustered
        | ChartType::Bar3DStacked
        | ChartType::Bar3DStacked100
        | ChartType::CylinderBarClustered
        | ChartType::CylinderBarStacked
        | ChartType::CylinderBarStacked100
        | ChartType::ConeBarClustered
        | ChartType::ConeBarStacked
        | ChartType::ConeBarStacked100
        | ChartType::PyramidBarClustered
        | ChartType::PyramidBarStacked
        | ChartType::PyramidBarStacked100 => Some("bar"),
        ChartType::Column
        | ChartType::ColumnStacked
        | ChartType::ColumnStacked100
        | ChartType::Column3D
        | ChartType::Column3DClustered
        | ChartType::Column3DStacked
        | ChartType::Column3DStacked100
        | ChartType::CylinderColumn
        | ChartType::CylinderColumnClustered
        | ChartType::CylinderColumnStacked
        | ChartType::CylinderColumnStacked100
        | ChartType::ConeColumn
        | ChartType::ConeColumnClustered
        | ChartType::ConeColumnStacked
        | ChartType::ConeColumnStacked100
        | ChartType::PyramidColumn
        | ChartType::PyramidColumnClustered
        | ChartType::PyramidColumnStacked
        | ChartType::PyramidColumnStacked100 => Some("col"),
        _ => None,
    }
}

fn chart_type_grouping_xml_value(chart_type: &ChartType) -> Option<&'static str> {
    match chart_type {
        ChartType::Area | ChartType::Area3D => Some("standard"),
        ChartType::AreaStacked | ChartType::Area3DStacked => Some("stacked"),
        ChartType::AreaStacked100 | ChartType::Area3DStacked100 => Some("percentStacked"),
        ChartType::Bar
        | ChartType::Column
        | ChartType::Bar3DClustered
        | ChartType::Column3DClustered
        | ChartType::CylinderColumnClustered
        | ChartType::CylinderBarClustered
        | ChartType::ConeColumnClustered
        | ChartType::ConeBarClustered
        | ChartType::PyramidColumnClustered
        | ChartType::PyramidBarClustered => Some("clustered"),
        ChartType::Column3D
        | ChartType::CylinderColumn
        | ChartType::ConeColumn
        | ChartType::PyramidColumn => Some("standard"),
        ChartType::BarStacked
        | ChartType::ColumnStacked
        | ChartType::Bar3DStacked
        | ChartType::Column3DStacked
        | ChartType::CylinderColumnStacked
        | ChartType::CylinderBarStacked
        | ChartType::ConeColumnStacked
        | ChartType::ConeBarStacked
        | ChartType::PyramidColumnStacked
        | ChartType::PyramidBarStacked => Some("stacked"),
        ChartType::BarStacked100
        | ChartType::ColumnStacked100
        | ChartType::Bar3DStacked100
        | ChartType::Column3DStacked100
        | ChartType::CylinderColumnStacked100
        | ChartType::CylinderBarStacked100
        | ChartType::ConeColumnStacked100
        | ChartType::ConeBarStacked100
        | ChartType::PyramidColumnStacked100
        | ChartType::PyramidBarStacked100 => Some("percentStacked"),
        ChartType::Line | ChartType::LineMarkers => Some("standard"),
        ChartType::LineStacked | ChartType::LineMarkersStacked => Some("stacked"),
        ChartType::LineStacked100 | ChartType::LineMarkersStacked100 => Some("percentStacked"),
        _ => None,
    }
}

fn chart_bar_shape_xml_value(shape: ChartBarShape) -> &'static str {
    match shape {
        ChartBarShape::Box => "box",
        ChartBarShape::PyramidToPoint => "pyramid",
        ChartBarShape::PyramidToMax => "pyramidToMax",
        ChartBarShape::Cylinder => "cylinder",
        ChartBarShape::ConeToPoint => "cone",
        ChartBarShape::ConeToMax => "coneToMax",
    }
}

fn chart_marker_style_xml_value(style: ChartMarkerStyle) -> &'static str {
    match style {
        ChartMarkerStyle::Automatic => "auto",
        ChartMarkerStyle::Circle => "circle",
        ChartMarkerStyle::Dash => "dash",
        ChartMarkerStyle::Diamond => "diamond",
        ChartMarkerStyle::Dot => "dot",
        ChartMarkerStyle::None => "none",
        ChartMarkerStyle::Picture => "picture",
        ChartMarkerStyle::Plus => "plus",
        ChartMarkerStyle::Square => "square",
        ChartMarkerStyle::Star => "star",
        ChartMarkerStyle::Triangle => "triangle",
        ChartMarkerStyle::X => "x",
    }
}

fn chart_bar_shape_from_chart_type(chart_type: &ChartType) -> Option<ChartBarShape> {
    match chart_type {
        ChartType::Bar3DClustered
        | ChartType::Bar3DStacked
        | ChartType::Bar3DStacked100
        | ChartType::Column3D
        | ChartType::Column3DClustered
        | ChartType::Column3DStacked
        | ChartType::Column3DStacked100 => Some(ChartBarShape::Box),
        ChartType::CylinderColumn
        | ChartType::CylinderColumnClustered
        | ChartType::CylinderColumnStacked
        | ChartType::CylinderColumnStacked100
        | ChartType::CylinderBarClustered
        | ChartType::CylinderBarStacked
        | ChartType::CylinderBarStacked100 => Some(ChartBarShape::Cylinder),
        ChartType::ConeColumn
        | ChartType::ConeColumnClustered
        | ChartType::ConeColumnStacked
        | ChartType::ConeColumnStacked100
        | ChartType::ConeBarClustered
        | ChartType::ConeBarStacked
        | ChartType::ConeBarStacked100 => Some(ChartBarShape::ConeToPoint),
        ChartType::PyramidColumn
        | ChartType::PyramidColumnClustered
        | ChartType::PyramidColumnStacked
        | ChartType::PyramidColumnStacked100
        | ChartType::PyramidBarClustered
        | ChartType::PyramidBarStacked
        | ChartType::PyramidBarStacked100 => Some(ChartBarShape::PyramidToPoint),
        _ => None,
    }
}

fn chart_effective_bar_shape(chart: &ChartModel) -> Option<ChartBarShape> {
    if chart_type_supports_bar_shape(&chart.chart_type) {
        chart
            .bar_shape
            .or_else(|| chart_bar_shape_from_chart_type(&chart.chart_type))
    } else {
        None
    }
}

fn chart_type_line_marker_xml_value(chart_type: &ChartType) -> Option<&'static str> {
    match chart_type {
        ChartType::Line | ChartType::LineStacked | ChartType::LineStacked100 => Some("0"),
        ChartType::LineMarkers
        | ChartType::LineMarkersStacked
        | ChartType::LineMarkersStacked100 => Some("1"),
        _ => None,
    }
}

fn chart_type_scatter_style_xml_value(chart_type: &ChartType) -> Option<&'static str> {
    match chart_type {
        ChartType::Scatter => Some("marker"),
        ChartType::ScatterLines => Some("lineMarker"),
        ChartType::ScatterLinesNoMarkers => Some("line"),
        ChartType::ScatterSmooth => Some("smoothMarker"),
        ChartType::ScatterSmoothNoMarkers => Some("smooth"),
        _ => None,
    }
}

fn chart_type_radar_style_xml_value(chart_type: &ChartType) -> Option<&'static str> {
    match chart_type {
        ChartType::Radar => Some("standard"),
        ChartType::RadarMarkers => Some("marker"),
        ChartType::RadarFilled => Some("filled"),
        _ => None,
    }
}

fn chart_type_of_pie_xml_value(chart_type: &ChartType) -> Option<&'static str> {
    match chart_type {
        ChartType::PieOfPie => Some("pie"),
        ChartType::BarOfPie => Some("bar"),
        _ => None,
    }
}

fn chart_type_surface_wireframe_xml_value(chart_type: &ChartType) -> Option<&'static str> {
    match chart_type {
        ChartType::Surface | ChartType::SurfaceTopView => Some("0"),
        ChartType::SurfaceWireframe | ChartType::SurfaceTopViewWireframe => Some("1"),
        _ => None,
    }
}

fn chart_type_has_axes(chart_type: &ChartType) -> bool {
    !matches!(
        chart_type,
        ChartType::Doughnut
            | ChartType::DoughnutExploded
            | ChartType::Pie
            | ChartType::Pie3D
            | ChartType::PieExploded
            | ChartType::Pie3DExploded
            | ChartType::PieOfPie
            | ChartType::BarOfPie
    )
}

fn chart_type_supports_bar_shape(chart_type: &ChartType) -> bool {
    matches!(chart_group_xml_name(chart_type), Some("bar3DChart"))
}

fn chart_type_supports_series_smooth(chart_type: &ChartType) -> bool {
    matches!(
        chart_group_xml_name(chart_type),
        Some("lineChart" | "scatterChart")
    )
}

fn chart_type_supports_series_marker(chart_type: &ChartType) -> bool {
    matches!(
        chart_group_xml_name(chart_type),
        Some("lineChart" | "scatterChart" | "radarChart")
    )
}

fn chart_type_supports_gap_depth(chart_type: &ChartType) -> bool {
    matches!(
        chart_group_xml_name(chart_type),
        Some("area3DChart" | "bar3DChart" | "line3DChart")
    )
}

fn chart_type_supports_explosion(chart_type: &ChartType) -> bool {
    matches!(
        chart_type,
        ChartType::Doughnut
            | ChartType::DoughnutExploded
            | ChartType::Pie
            | ChartType::Pie3D
            | ChartType::PieExploded
            | ChartType::Pie3DExploded
    )
}

fn chart_explosion_xml_value(chart: &ChartModel) -> Option<String> {
    if !chart_type_supports_explosion(&chart.chart_type) {
        return None;
    }
    chart
        .explosion
        .or_else(|| {
            matches!(
                chart.chart_type,
                ChartType::PieExploded | ChartType::Pie3DExploded | ChartType::DoughnutExploded
            )
            .then_some(25)
        })
        .map(|value| value.to_string())
}

fn chart_axis_kind_from_xml_name(local_name: &[u8]) -> Option<ChartAxisKind> {
    match local_name {
        b"catAx" => Some(ChartAxisKind::Category),
        b"valAx" => Some(ChartAxisKind::Value),
        b"dateAx" => Some(ChartAxisKind::Date),
        b"serAx" => Some(ChartAxisKind::Series),
        _ => None,
    }
}

fn chart_axis_xml_name(kind: ChartAxisKind) -> &'static str {
    match kind {
        ChartAxisKind::Category => "catAx",
        ChartAxisKind::Value => "valAx",
        ChartAxisKind::Date => "dateAx",
        ChartAxisKind::Series => "serAx",
    }
}

fn chart_axis_orientation_xml_value(reverse_plot_order: bool) -> &'static str {
    if reverse_plot_order {
        "maxMin"
    } else {
        "minMax"
    }
}

fn chart_axis_log_base_xml_value(axis: &AxisModel) -> Option<f64> {
    if axis.scale_type == Some(ChartAxisScaleType::Logarithmic) || axis.log_base.is_some() {
        Some(axis.log_base.unwrap_or(10.0))
    } else {
        None
    }
}

fn chart_axis_has_scaling_xml(axis: &AxisModel) -> bool {
    chart_axis_log_base_xml_value(axis).is_some()
        || axis.minimum_scale.is_some()
        || axis.maximum_scale.is_some()
        || axis.reverse_plot_order.is_some()
}

fn chart_axis_time_unit_xml_value(value: ChartAxisTimeUnit) -> &'static str {
    match value {
        ChartAxisTimeUnit::Days => "days",
        ChartAxisTimeUnit::Months => "months",
        ChartAxisTimeUnit::Years => "years",
    }
}

fn chart_built_in_display_unit_xml_value(value: ChartBuiltInDisplayUnit) -> &'static str {
    match value {
        ChartBuiltInDisplayUnit::Hundreds => "hundreds",
        ChartBuiltInDisplayUnit::Thousands => "thousands",
        ChartBuiltInDisplayUnit::TenThousands => "tenThousands",
        ChartBuiltInDisplayUnit::HundredThousands => "hundredThousands",
        ChartBuiltInDisplayUnit::Millions => "millions",
        ChartBuiltInDisplayUnit::TenMillions => "tenMillions",
        ChartBuiltInDisplayUnit::HundredMillions => "hundredMillions",
        ChartBuiltInDisplayUnit::ThousandMillions => "billions",
        ChartBuiltInDisplayUnit::MillionMillions => "trillions",
    }
}

fn chart_axis_display_unit_label_text(axis: &AxisModel) -> String {
    if let Some(label) = axis.display_unit_label.as_ref() {
        return label.text.clone();
    }
    match axis.display_unit {
        Some(ChartAxisDisplayUnit::BuiltIn(ChartBuiltInDisplayUnit::Hundreds)) => {
            "Hundreds".to_string()
        }
        Some(ChartAxisDisplayUnit::BuiltIn(ChartBuiltInDisplayUnit::Thousands)) => {
            "Thousands".to_string()
        }
        Some(ChartAxisDisplayUnit::BuiltIn(ChartBuiltInDisplayUnit::TenThousands)) => {
            "Ten Thousands".to_string()
        }
        Some(ChartAxisDisplayUnit::BuiltIn(ChartBuiltInDisplayUnit::HundredThousands)) => {
            "Hundred Thousands".to_string()
        }
        Some(ChartAxisDisplayUnit::BuiltIn(ChartBuiltInDisplayUnit::Millions)) => {
            "Millions".to_string()
        }
        Some(ChartAxisDisplayUnit::BuiltIn(ChartBuiltInDisplayUnit::TenMillions)) => {
            "Ten Millions".to_string()
        }
        Some(ChartAxisDisplayUnit::BuiltIn(ChartBuiltInDisplayUnit::HundredMillions)) => {
            "Hundred Millions".to_string()
        }
        Some(ChartAxisDisplayUnit::BuiltIn(ChartBuiltInDisplayUnit::ThousandMillions)) => {
            "Billions".to_string()
        }
        Some(ChartAxisDisplayUnit::BuiltIn(ChartBuiltInDisplayUnit::MillionMillions)) => {
            "Trillions".to_string()
        }
        Some(ChartAxisDisplayUnit::Custom(value)) => chart_number_xml_value(value),
        None => String::new(),
    }
}

fn chart_axis_crosses_xml_value(value: ChartAxisCrosses) -> Option<&'static str> {
    match value {
        ChartAxisCrosses::Automatic => Some("autoZero"),
        ChartAxisCrosses::Custom => None,
        ChartAxisCrosses::Maximum => Some("max"),
        ChartAxisCrosses::Minimum => Some("min"),
    }
}

fn chart_axis_between_categories_xml_value(value: bool) -> &'static str {
    if value { "between" } else { "midCat" }
}

fn chart_legend_position_xml_value(position: ChartLegendPosition) -> &'static str {
    match position {
        ChartLegendPosition::Bottom => "b",
        ChartLegendPosition::Corner => "tr",
        ChartLegendPosition::Custom => "cust",
        ChartLegendPosition::Left => "l",
        ChartLegendPosition::Right => "r",
        ChartLegendPosition::Top => "t",
    }
}

fn chart_display_blanks_as_xml_value(value: ChartDisplayBlanksAs) -> &'static str {
    match value {
        ChartDisplayBlanksAs::Gap => "gap",
        ChartDisplayBlanksAs::Span => "span",
        ChartDisplayBlanksAs::Zero => "zero",
    }
}

fn chart_view_3d_xml_string(view_3d: &ChartView3DModel) -> Option<String> {
    let mut children = String::new();
    if let Some(value) = view_3d.elevation {
        children.push_str(&format!(r#"<c:rotX val="{value}"/>"#));
    }
    if let Some(value) = view_3d.height_percent {
        children.push_str(&format!(r#"<c:hPercent val="{value}"/>"#));
    }
    if let Some(value) = view_3d.rotation {
        children.push_str(&format!(r#"<c:rotY val="{value}"/>"#));
    }
    if let Some(value) = view_3d.depth_percent {
        children.push_str(&format!(r#"<c:depthPercent val="{value}"/>"#));
    }
    if let Some(value) = view_3d.right_angle_axes {
        children.push_str(&format!(
            r#"<c:rAngAx val="{}"/>"#,
            if value { "1" } else { "0" }
        ));
    }
    if let Some(value) = view_3d.perspective {
        children.push_str(&format!(r#"<c:perspective val="{value}"/>"#));
    }
    (!children.is_empty()).then(|| format!("<c:view3D>{children}</c:view3D>"))
}

fn write_chart_view_3d_element(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    view_3d: &ChartView3DModel,
) -> OmResult<bool> {
    let Some(xml) = chart_view_3d_xml_string(view_3d) else {
        return Ok(false);
    };
    writer
        .get_mut()
        .write_all(xml.as_bytes())
        .map_err(chart_xml_error)?;
    Ok(true)
}

fn chart_protection_xml(protection: ChartProtectionModel) -> String {
    let mut children = String::new();
    for (name, protected) in [
        ("chartObject", protection.contents),
        ("data", protection.data),
        ("formatting", protection.formatting),
        ("selection", protection.selection),
        ("userInterface", protection.user_interface_only),
    ] {
        if protected {
            children.push_str(format!(r#"<c:{name} val="1"/>"#).as_str());
        }
    }
    if children.is_empty() {
        String::new()
    } else {
        format!("<c:protection>{children}</c:protection>")
    }
}

fn write_chart_protection_element(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    protection: ChartProtectionModel,
) -> OmResult<bool> {
    let children = [
        ("c:chartObject", protection.contents),
        ("c:data", protection.data),
        ("c:formatting", protection.formatting),
        ("c:selection", protection.selection),
        ("c:userInterface", protection.user_interface_only),
    ];
    if !children.iter().any(|(_, protected)| *protected) {
        return Ok(false);
    }
    writer
        .write_event(Event::Start(BytesStart::new("c:protection")))
        .map_err(chart_xml_error)?;
    for (name, protected) in children {
        if protected {
            let mut child = BytesStart::new(name);
            child.push_attribute(("val", "1"));
            writer
                .write_event(Event::Empty(child))
                .map_err(chart_xml_error)?;
        }
    }
    writer
        .write_event(Event::End(BytesEnd::new("c:protection")))
        .map_err(chart_xml_error)?;
    Ok(true)
}

fn chart_tick_mark_xml_value(value: ChartTickMark) -> &'static str {
    match value {
        ChartTickMark::Cross => "cross",
        ChartTickMark::Inside => "in",
        ChartTickMark::None => "none",
        ChartTickMark::Outside => "out",
    }
}

fn chart_tick_label_position_xml_value(value: ChartTickLabelPosition) -> &'static str {
    match value {
        ChartTickLabelPosition::High => "high",
        ChartTickLabelPosition::Low => "low",
        ChartTickLabelPosition::NextToAxis => "nextTo",
        ChartTickLabelPosition::None => "none",
    }
}

fn chart_data_label_position_xml_value(value: ChartDataLabelPosition) -> &'static str {
    match value {
        ChartDataLabelPosition::Above => "t",
        ChartDataLabelPosition::Below => "b",
        ChartDataLabelPosition::BestFit => "bestFit",
        ChartDataLabelPosition::Center => "ctr",
        ChartDataLabelPosition::InsideBase => "inBase",
        ChartDataLabelPosition::InsideEnd => "inEnd",
        ChartDataLabelPosition::Left => "l",
        ChartDataLabelPosition::OutsideEnd => "outEnd",
        ChartDataLabelPosition::Right => "r",
    }
}

fn chart_size_represents_xml_value(value: ChartSizeRepresents) -> &'static str {
    match value {
        ChartSizeRepresents::Area => "area",
        ChartSizeRepresents::Width => "w",
    }
}

fn chart_split_type_xml_value(value: ChartSplitType) -> &'static str {
    match value {
        ChartSplitType::Custom => "cust",
        ChartSplitType::PercentValue => "percent",
        ChartSplitType::Position => "pos",
        ChartSplitType::Value => "val",
    }
}

fn om_value_text(value: &OmValue) -> Option<String> {
    match value {
        OmValue::Missing | OmValue::Empty | OmValue::Null => Some(String::new()),
        OmValue::Bool(true) => Some("TRUE".to_string()),
        OmValue::Bool(false) => Some("FALSE".to_string()),
        OmValue::Number(number) => Some(format_find_number(*number)),
        OmValue::Text(text) => Some(text.clone()),
        OmValue::Error(error) => Some(formula_cell_error_text(error).to_string()),
        OmValue::Object(_) | OmValue::Array(_) => None,
    }
}

fn chart_source_container_xml_string(
    chart_type: &ChartType,
    slot: ChartSourceXmlSlot,
    source: &ChartSourceExpr,
) -> OmResult<String> {
    let container_name = match (chart_type, slot) {
        (_, ChartSourceXmlSlot::Name) => "tx",
        (chart_type, ChartSourceXmlSlot::XValues) if chart_type_uses_xy_values(chart_type) => {
            "xVal"
        }
        (chart_type, ChartSourceXmlSlot::Values) if chart_type_uses_xy_values(chart_type) => "yVal",
        (_, ChartSourceXmlSlot::XValues) => "cat",
        (_, ChartSourceXmlSlot::Values) => "val",
        (_, ChartSourceXmlSlot::BubbleSize) => "bubbleSize",
    };

    if let Some(values) = chart_source_literal_values(source)? {
        if slot == ChartSourceXmlSlot::Name {
            let value = values.first().map(String::as_str).unwrap_or_default();
            return Ok(format!("<c:tx><c:v>{}</c:v></c:tx>", partial_escape(value)));
        }

        let literal_name =
            if slot == ChartSourceXmlSlot::XValues && !chart_type_uses_xy_values(chart_type) {
                "strLit"
            } else {
                "numLit"
            };
        let mut xml = format!(
            r#"<c:{container_name}><c:{literal_name}><c:ptCount val="{}"/>"#,
            values.len()
        );
        for (index, value) in values.iter().enumerate() {
            xml.push_str(&format!(
                r#"<c:pt idx="{index}"><c:v>{}</c:v></c:pt>"#,
                partial_escape(value)
            ));
        }
        xml.push_str(&format!("</c:{literal_name}></c:{container_name}>"));
        return Ok(xml);
    }

    let reference_name = match (chart_type, slot) {
        (_, ChartSourceXmlSlot::Name) => "strRef",
        (chart_type, ChartSourceXmlSlot::XValues) if chart_type_uses_xy_values(chart_type) => {
            "numRef"
        }
        (chart_type, ChartSourceXmlSlot::Values) if chart_type_uses_xy_values(chart_type) => {
            "numRef"
        }
        (_, ChartSourceXmlSlot::XValues) => "strRef",
        (_, ChartSourceXmlSlot::Values) => "numRef",
        (_, ChartSourceXmlSlot::BubbleSize) => "numRef",
    };
    let formula = partial_escape(source.raw.text.trim_start_matches('=')).to_string();
    let full_reference = source
        .full_reference
        .as_ref()
        .map(|reference| {
            let reference =
                partial_escape(reference.raw.text.trim_start_matches('=')).to_string();
            format!(
                r#"<c:extLst><c:ext uri="{{02D57815-91ED-43cb-92C2-25804820EDAC}}"><c15:fullRef xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"><c15:sqref>{reference}</c15:sqref></c15:fullRef></c:ext></c:extLst>"#
            )
        })
        .unwrap_or_default();
    Ok(format!(
        r#"<c:{container_name}><c:{reference_name}><c:f>{formula}</c:f>{full_reference}</c:{reference_name}></c:{container_name}>"#
    ))
}

fn chart_source_literal_values(source: &ChartSourceExpr) -> OmResult<Option<Vec<String>>> {
    match source.resolved.as_ref() {
        Some(ReferenceTarget::Array(array)) => array
            .values
            .iter()
            .map(|value| {
                om_value_text(value).ok_or_else(|| {
                    OmError::type_mismatch(
                        "chart literal arrays cannot contain object or nested array values",
                    )
                })
            })
            .collect::<OmResult<Vec<_>>>()
            .map(Some),
        Some(ReferenceTarget::Value(value)) => Ok(Some(vec![find_cell_value_text(value)])),
        _ => Ok(None),
    }
}

fn chart_number_xml_value(value: f64) -> String {
    if value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
}

fn push_chart_data_label_properties_xml(xml: &mut String, data_labels: &ChartDataLabelsModel) {
    if let Some(format_code) = data_labels.number_format.as_ref() {
        let format_code = partial_escape(format_code).to_string();
        let source_linked = if data_labels.number_format_linked.unwrap_or(true) {
            "1"
        } else {
            "0"
        };
        xml.push_str(&format!(
            r#"<c:numFmt formatCode="{format_code}" sourceLinked="{source_linked}"/>"#
        ));
    }
    if let Some(position) = data_labels.position {
        let position = chart_data_label_position_xml_value(position);
        xml.push_str(&format!(r#"<c:dLblPos val="{position}"/>"#));
    }
    for (element_name, value) in [
        ("showLegendKey", data_labels.show_legend_key),
        ("showLeaderLines", data_labels.has_leader_lines),
        ("showSerName", data_labels.show_series_name),
        ("showCatName", data_labels.show_category_name),
        ("showVal", data_labels.show_value),
        ("showPercent", data_labels.show_percentage),
        ("showBubbleSize", data_labels.show_bubble_size),
    ] {
        if let Some(value) = value {
            xml.push_str(&format!(
                r#"<c:{element_name} val="{}"/>"#,
                if value { "1" } else { "0" }
            ));
        }
    }
    if let Some(separator) = data_labels.separator.as_ref() {
        let separator = partial_escape(separator).to_string();
        xml.push_str(&format!(r#"<c:separator>{separator}</c:separator>"#));
    }
}

fn chart_data_labels_xml_string(data_labels: &ChartDataLabelsModel) -> String {
    let mut xml = String::from("<c:dLbls>");
    push_chart_data_label_properties_xml(&mut xml, data_labels);
    xml.push_str("</c:dLbls>");
    xml
}

fn chart_data_table_xml_string(data_table: &ChartDataTableModel) -> String {
    let mut xml = String::from("<c:dTable>");
    for (element_name, value) in [
        ("showHorzBorder", data_table.has_border_horizontal),
        ("showVertBorder", data_table.has_border_vertical),
        ("showOutline", data_table.has_border_outline),
        ("showKeys", data_table.show_legend_key),
    ] {
        if let Some(value) = value {
            xml.push_str(&format!(
                r#"<c:{element_name} val="{}"/>"#,
                if value { "1" } else { "0" }
            ));
        }
    }
    xml.push_str("</c:dTable>");
    xml
}

fn chart_manual_layout_xml_string(layout: &ChartManualLayout) -> String {
    let target = match layout.target {
        ChartLayoutTarget::Inner => "inner",
        ChartLayoutTarget::Outer => "outer",
    };
    let x_mode = match layout.x_mode {
        ChartLayoutMode::Edge => "edge",
        ChartLayoutMode::Factor => "factor",
    };
    let y_mode = match layout.y_mode {
        ChartLayoutMode::Edge => "edge",
        ChartLayoutMode::Factor => "factor",
    };
    let width_mode = match layout.width_mode {
        ChartLayoutMode::Edge => "edge",
        ChartLayoutMode::Factor => "factor",
    };
    let height_mode = match layout.height_mode {
        ChartLayoutMode::Edge => "edge",
        ChartLayoutMode::Factor => "factor",
    };
    let x = layout
        .x
        .map(|value| format!(r#"<c:x val="{}"/>"#, chart_number_xml_value(value)))
        .unwrap_or_default();
    let y = layout
        .y
        .map(|value| format!(r#"<c:y val="{}"/>"#, chart_number_xml_value(value)))
        .unwrap_or_default();
    let width = layout
        .width
        .map(|value| format!(r#"<c:w val="{}"/>"#, chart_number_xml_value(value)))
        .unwrap_or_default();
    let height = layout
        .height
        .map(|value| format!(r#"<c:h val="{}"/>"#, chart_number_xml_value(value)))
        .unwrap_or_default();
    format!(
        r#"<c:manualLayout><c:layoutTarget val="{target}"/><c:xMode val="{x_mode}"/><c:yMode val="{y_mode}"/><c:wMode val="{width_mode}"/><c:hMode val="{height_mode}"/>{x}{y}{width}{height}</c:manualLayout>"#
    )
}

fn write_chart_data_table_element<W: Write>(
    writer: &mut Writer<W>,
    data_table: &ChartDataTableModel,
) -> OmResult<()> {
    writer
        .write_event(Event::Start(BytesStart::new("c:dTable")))
        .map_err(chart_xml_error)?;
    for (element_name, value) in [
        ("c:showHorzBorder", data_table.has_border_horizontal),
        ("c:showVertBorder", data_table.has_border_vertical),
        ("c:showOutline", data_table.has_border_outline),
        ("c:showKeys", data_table.show_legend_key),
    ] {
        if let Some(value) = value {
            let mut element = BytesStart::new(element_name);
            element.push_attribute(("val", if value { "1" } else { "0" }));
            writer
                .write_event(Event::Empty(element))
                .map_err(chart_xml_error)?;
        }
    }
    writer
        .write_event(Event::End(BytesEnd::new("c:dTable")))
        .map_err(chart_xml_error)?;
    Ok(())
}

fn chart_series_data_labels_xml_string(series: &SeriesModel) -> String {
    if series.data_labels.is_none() && series.point_data_labels.is_empty() {
        return String::new();
    }
    let mut xml = String::from("<c:dLbls>");
    for (point_index, data_labels) in &series.point_data_labels {
        xml.push_str(&format!(r#"<c:dLbl><c:idx val="{point_index}"/>"#));
        push_chart_data_label_properties_xml(&mut xml, data_labels);
        xml.push_str("</c:dLbl>");
    }
    if let Some(data_labels) = series.data_labels.as_ref() {
        push_chart_data_label_properties_xml(&mut xml, data_labels);
    }
    xml.push_str("</c:dLbls>");
    xml
}

fn chart_extension_without_full_reference(extension_xml: &[u8]) -> OmResult<Option<Vec<u8>>> {
    let mut reader = Reader::from_reader(Cursor::new(extension_xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut skip_depth = 0usize;
    let mut retained_payload = false;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_)) if skip_depth > 0 => skip_depth += 1,
            Ok(Event::End(_)) if skip_depth > 0 => skip_depth -= 1,
            Ok(_) if skip_depth > 0 => {}
            Ok(Event::Start(element)) => {
                let element_name = element.name();
                let local_name = xml_local_name(element_name.as_ref());
                if depth > 0 && local_name == b"fullRef" {
                    skip_depth = 1;
                } else {
                    if depth == 1 {
                        retained_payload = true;
                    }
                    writer
                        .write_event(Event::Start(element.into_owned()))
                        .map_err(chart_xml_error)?;
                    depth += 1;
                }
            }
            Ok(Event::Empty(element)) => {
                let element_name = element.name();
                let local_name = xml_local_name(element_name.as_ref());
                if !(depth > 0 && local_name == b"fullRef") {
                    if depth == 1 {
                        retained_payload = true;
                    }
                    writer
                        .write_event(Event::Empty(element.into_owned()))
                        .map_err(chart_xml_error)?;
                }
            }
            Ok(Event::End(element)) => {
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(chart_xml_error)?;
                depth = depth.saturating_sub(1);
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer
                .write_event(event.into_owned())
                .map_err(chart_xml_error)?,
            Err(error) => return Err(chart_xml_error(error)),
        }
        buffer.clear();
    }

    Ok(retained_payload.then(|| writer.into_inner().into_inner()))
}

#[derive(Debug)]
struct LoadedChartXmlElementSpan {
    start: usize,
    end: usize,
    end_tag_start: usize,
    parent_start: Option<usize>,
    child_element_count: usize,
}

#[derive(Debug)]
struct LoadedChartXmlGroupSpan {
    local_name: Vec<u8>,
    end_tag_start: usize,
    last_direct_series_end: Option<usize>,
    first_after_series_start: Option<usize>,
    direct_ext_lst_start: Option<usize>,
}

#[derive(Debug)]
struct LoadedChartXmlSeriesSpan {
    start: usize,
    end: usize,
    raw_index: Option<u32>,
    is_filtered: bool,
    wrapper_start: Option<usize>,
    extension_start: Option<usize>,
    group_index: usize,
}

#[derive(Debug)]
struct LoadedChartXmlFrame {
    local_name: Vec<u8>,
    start: usize,
    parent_start: Option<usize>,
    child_element_count: usize,
    group_index: Option<usize>,
    series_index: Option<usize>,
}

fn is_chart_filtered_series_wrapper_name(local_name: &[u8]) -> bool {
    matches!(
        local_name,
        b"filteredAreaSeries"
            | b"filteredBarSeries"
            | b"filteredLineSeries"
            | b"filteredScatterSeries"
            | b"filteredBubbleSeries"
            | b"filteredPieSeries"
            | b"filteredRadarSeries"
            | b"filteredSurfaceSeries"
    )
}

fn chart_group_child_precedes_series(local_name: &[u8]) -> bool {
    matches!(
        local_name,
        b"barDir"
            | b"grouping"
            | b"varyColors"
            | b"scatterStyle"
            | b"radarStyle"
            | b"ofPieType"
            | b"wireframe"
    )
}

fn rewrite_chart_series_outer_name(series_xml: &[u8], qualified_name: &str) -> OmResult<Vec<u8>> {
    let mut reader = Reader::from_reader(Cursor::new(series_xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) => {
                depth += 1;
                if depth == 1 && xml_local_name(element.name().as_ref()) == b"ser" {
                    let mut replacement = BytesStart::new(qualified_name);
                    let mut has_chart_namespace = false;
                    for attr in element.attributes() {
                        let attr = attr.map_err(chart_xml_error)?;
                        has_chart_namespace |= attr.key.as_ref() == b"xmlns:c";
                        replacement.push_attribute(attr.to_owned());
                    }
                    if qualified_name == "c:ser" && !has_chart_namespace {
                        replacement.push_attribute((
                            "xmlns:c",
                            "http://schemas.openxmlformats.org/drawingml/2006/chart",
                        ));
                    }
                    writer
                        .write_event(Event::Start(replacement))
                        .map_err(chart_xml_error)?;
                } else {
                    writer
                        .write_event(Event::Start(element.into_owned()))
                        .map_err(chart_xml_error)?;
                }
            }
            Ok(Event::End(element)) => {
                if depth == 1 && xml_local_name(element.name().as_ref()) == b"ser" {
                    writer
                        .write_event(Event::End(BytesEnd::new(qualified_name)))
                        .map_err(chart_xml_error)?;
                } else {
                    writer
                        .write_event(Event::End(element.into_owned()))
                        .map_err(chart_xml_error)?;
                }
                depth = depth.saturating_sub(1);
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer
                .write_event(event.into_owned())
                .map_err(chart_xml_error)?,
            Err(error) => return Err(chart_xml_error(error)),
        }
        buffer.clear();
    }
    Ok(writer.into_inner().into_inner())
}

fn serialize_loaded_chart_group_shell(group: &ChartGroupModel) -> OmResult<Vec<u8>> {
    if chart_type_from_group_name(group.raw_name.as_bytes()).is_none() {
        return Err(OmError::unsupported(
            "new loaded chart groups require a recognized chart-group type",
        ));
    }
    if group.axis_ids.is_empty() {
        return Err(OmError::unsupported(
            "new loaded chart groups require existing target axis identifiers",
        ));
    }
    let qualified_name = format!("c:{}", group.raw_name);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut start = BytesStart::new(qualified_name.clone());
    start.push_attribute((
        "xmlns:c",
        "http://schemas.openxmlformats.org/drawingml/2006/chart",
    ));
    writer
        .write_event(Event::Start(start))
        .map_err(chart_xml_error)?;
    let write_value =
        |writer: &mut Writer<Cursor<Vec<u8>>>, local_name: &str, value: &str| -> OmResult<()> {
            let mut element = BytesStart::new(format!("c:{local_name}"));
            element.push_attribute(("val", value));
            writer
                .write_event(Event::Empty(element))
                .map_err(chart_xml_error)
        };
    for (local_name, value) in [
        ("barDir", group.bar_direction.as_deref()),
        ("grouping", group.chart_grouping.as_deref()),
        ("scatterStyle", group.scatter_style.as_deref()),
        ("radarStyle", group.radar_style.as_deref()),
        ("ofPieType", group.of_pie_type.as_deref()),
    ] {
        if let Some(value) = value {
            write_value(&mut writer, local_name, value)?;
        }
    }
    if let Some(wireframe) = group.surface_wireframe {
        write_value(&mut writer, "wireframe", if wireframe { "1" } else { "0" })?;
    }
    if let Some(vary) = group.vary_by_categories {
        write_value(&mut writer, "varyColors", if vary { "1" } else { "0" })?;
    }
    if let Some(has_markers) = group.line_has_markers {
        write_value(&mut writer, "marker", if has_markers { "1" } else { "0" })?;
    }
    if let Some(data_labels) = group.data_labels.as_ref() {
        writer
            .get_mut()
            .write_all(chart_data_labels_xml_string(data_labels).as_bytes())
            .map_err(chart_xml_error)?;
    }
    for local_name in [
        b"gapWidth".as_slice(),
        b"gapDepth".as_slice(),
        b"overlap".as_slice(),
        b"firstSliceAng".as_slice(),
        b"bubbleScale".as_slice(),
        b"showNegBubbles".as_slice(),
        b"bubble3D".as_slice(),
        b"holeSize".as_slice(),
        b"secondPieSize".as_slice(),
        b"sizeRepresents".as_slice(),
        b"splitType".as_slice(),
        b"splitPos".as_slice(),
    ] {
        if let Some(value) = chart_group_direct_property_value(group, local_name) {
            write_value(&mut writer, &String::from_utf8_lossy(local_name), &value)?;
        }
    }
    if let Some(shape) = group.bar_shape {
        write_value(&mut writer, "shape", chart_bar_shape_xml_value(shape))?;
    }
    for (local_name, enabled) in [
        ("serLines", group.has_series_lines),
        ("dropLines", group.has_drop_lines),
        ("hiLowLines", group.has_hi_lo_lines),
        ("upDownBars", group.has_up_down_bars),
    ] {
        if enabled == Some(true) {
            writer
                .write_event(Event::Empty(BytesStart::new(format!("c:{local_name}"))))
                .map_err(chart_xml_error)?;
        }
    }
    for axis_id in &group.axis_ids {
        write_value(&mut writer, "axId", axis_id)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new(qualified_name)))
        .map_err(chart_xml_error)?;
    Ok(writer.into_inner().into_inner())
}

struct LoadedChartGroupXmlSpan {
    raw_name: Vec<u8>,
    start: usize,
    end: usize,
}

struct LoadedChartGroupXmlTopology {
    groups: Vec<LoadedChartGroupXmlSpan>,
    first_axis_start: Option<usize>,
    plot_area_end: Option<usize>,
}

#[derive(Clone, Copy)]
enum LoadedChartGroupSequenceToken {
    Original(usize),
    Added(usize),
}

struct LoadedChartGroupSequence {
    tokens: Vec<LoadedChartGroupSequenceToken>,
    model_to_xml: Vec<usize>,
    surviving_original_groups: Vec<bool>,
}

fn loaded_chart_group_sequence(
    loaded_group_names: &[Vec<u8>],
    chart: &ChartModel,
) -> OmResult<LoadedChartGroupSequence> {
    let added_group_count = chart
        .groups
        .iter()
        .filter(|group| group.loaded_index.is_none())
        .count();
    let original_group_count = loaded_group_names
        .len()
        .checked_sub(added_group_count)
        .ok_or_else(|| {
            OmError::new(
                OmErrorCode::InvalidState,
                "loaded chart has fewer XML groups than runtime additions",
            )
        })?;
    let mut surviving_original_groups = vec![false; original_group_count];
    let mut additions_before = vec![Vec::<usize>::new(); original_group_count + 1];
    for (model_index, group) in chart.groups.iter().enumerate() {
        if let Some(loaded_index) = group.loaded_index {
            if loaded_index >= original_group_count || surviving_original_groups[loaded_index] {
                return Err(OmError::unsupported(
                    "loaded chart group identity changed before topology patching",
                ));
            }
            surviving_original_groups[loaded_index] = true;
        } else {
            let insertion_index = chart.groups[model_index + 1..]
                .iter()
                .find_map(|next| next.loaded_index)
                .unwrap_or(original_group_count);
            if insertion_index > original_group_count {
                return Err(OmError::new(
                    OmErrorCode::InvalidState,
                    "runtime chart group insertion index exceeds loaded topology",
                ));
            }
            additions_before[insertion_index].push(model_index);
        }
    }
    let mut tokens = Vec::<LoadedChartGroupSequenceToken>::with_capacity(loaded_group_names.len());
    for original_index in 0..original_group_count {
        tokens.extend(
            additions_before[original_index]
                .iter()
                .copied()
                .map(LoadedChartGroupSequenceToken::Added),
        );
        tokens.push(LoadedChartGroupSequenceToken::Original(original_index));
    }
    tokens.extend(
        additions_before[original_group_count]
            .iter()
            .copied()
            .map(LoadedChartGroupSequenceToken::Added),
    );
    if tokens.len() != loaded_group_names.len() {
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            "loaded chart group sequence does not match runtime topology",
        ));
    }
    let mut model_to_xml = vec![usize::MAX; chart.groups.len()];
    for (xml_index, token) in tokens.iter().copied().enumerate() {
        match token {
            LoadedChartGroupSequenceToken::Original(loaded_index) => {
                if let Some((model_index, model_group)) = chart
                    .groups
                    .iter()
                    .enumerate()
                    .find(|(_, group)| group.loaded_index == Some(loaded_index))
                {
                    if loaded_group_names[xml_index].as_slice() != model_group.raw_name.as_bytes() {
                        return Err(OmError::unsupported(
                            "loaded chart group identity changed before topology patching",
                        ));
                    }
                    model_to_xml[model_index] = xml_index;
                }
            }
            LoadedChartGroupSequenceToken::Added(model_index) => {
                if loaded_group_names[xml_index].as_slice()
                    != chart.groups[model_index].raw_name.as_bytes()
                {
                    return Err(OmError::unsupported(
                        "runtime-added chart group identity changed before topology patching",
                    ));
                }
                model_to_xml[model_index] = xml_index;
            }
        }
    }
    if model_to_xml
        .iter()
        .any(|xml_index| *xml_index == usize::MAX)
    {
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            "chart model group could not be mapped to loaded XML",
        ));
    }
    Ok(LoadedChartGroupSequence {
        tokens,
        model_to_xml,
        surviving_original_groups,
    })
}

fn scan_loaded_chart_group_xml_topology(
    existing_chart_xml: &[u8],
) -> OmResult<LoadedChartGroupXmlTopology> {
    let mut reader = Reader::from_reader(Cursor::new(existing_chart_xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut element_stack = Vec::<Vec<u8>>::new();
    let mut groups = Vec::<LoadedChartGroupXmlSpan>::new();
    let mut first_axis_start = None::<usize>;
    let mut plot_area_end = None::<usize>;
    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| OmError::new(OmErrorCode::InvalidState, "chart XML position overflow"))?;
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) => {
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                if element_stack.last().map(Vec::as_slice) == Some(b"plotArea".as_slice()) {
                    if chart_type_from_group_name(local_name.as_slice()).is_some() {
                        groups.push(LoadedChartGroupXmlSpan {
                            raw_name: local_name.clone(),
                            start: event_start,
                            end: 0,
                        });
                    } else if chart_axis_kind_from_xml_name(local_name.as_slice()).is_some()
                        && first_axis_start.is_none()
                    {
                        first_axis_start = Some(event_start);
                    }
                }
                element_stack.push(local_name);
            }
            Ok(Event::Empty(element)) => {
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                if element_stack.last().map(Vec::as_slice) == Some(b"plotArea".as_slice()) {
                    if chart_type_from_group_name(local_name.as_slice()).is_some() {
                        let event_end =
                            usize::try_from(reader.buffer_position()).map_err(|_| {
                                OmError::new(
                                    OmErrorCode::InvalidState,
                                    "chart XML position overflow",
                                )
                            })?;
                        groups.push(LoadedChartGroupXmlSpan {
                            raw_name: local_name,
                            start: event_start,
                            end: event_end,
                        });
                    } else if chart_axis_kind_from_xml_name(local_name.as_slice()).is_some()
                        && first_axis_start.is_none()
                    {
                        first_axis_start = Some(event_start);
                    }
                }
            }
            Ok(Event::End(element)) => {
                let element_name = element.name();
                let local_name = xml_local_name(element_name.as_ref());
                if element_stack.len() >= 2
                    && element_stack[element_stack.len() - 2].as_slice() == b"plotArea"
                    && chart_type_from_group_name(local_name).is_some()
                {
                    let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                    })?;
                    let group = groups.last_mut().ok_or_else(|| {
                        OmError::new(
                            OmErrorCode::Parse,
                            "loaded chart group end has no matching start",
                        )
                    })?;
                    if group.end != 0 || group.raw_name.as_slice() != local_name {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            "loaded chart group elements are not properly nested",
                        ));
                    }
                    group.end = event_end;
                }
                if local_name == b"plotArea" {
                    plot_area_end = Some(event_start);
                }
                element_stack.pop();
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(chart_xml_error(error)),
        }
        buffer.clear();
    }
    if groups.iter().any(|group| group.end == 0) {
        return Err(OmError::new(
            OmErrorCode::Parse,
            "loaded chart group has no closing XML boundary",
        ));
    }
    Ok(LoadedChartGroupXmlTopology {
        groups,
        first_axis_start,
        plot_area_end,
    })
}

fn apply_loaded_chart_xml_edits(
    existing_chart_xml: &[u8],
    mut edits: Vec<(usize, usize, Vec<u8>)>,
    label: &str,
) -> OmResult<Vec<u8>> {
    if edits.is_empty() {
        return Ok(existing_chart_xml.to_vec());
    }
    edits.sort_by_key(|(start, end, _)| (*start, *end));
    let additional_capacity = edits
        .iter()
        .map(|(_, _, replacement)| replacement.len())
        .sum::<usize>();
    let mut rewritten = Vec::with_capacity(existing_chart_xml.len() + additional_capacity);
    let mut cursor = 0usize;
    for (start, end, replacement) in edits {
        if start < cursor || end < start || end > existing_chart_xml.len() {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("loaded chart {label} XML edits overlap"),
            ));
        }
        rewritten.extend_from_slice(&existing_chart_xml[cursor..start]);
        rewritten.extend_from_slice(&replacement);
        cursor = end;
    }
    rewritten.extend_from_slice(&existing_chart_xml[cursor..]);
    Ok(rewritten)
}

fn rewrite_loaded_chart_group_additions(
    existing_chart_xml: &[u8],
    chart: &ChartModel,
) -> OmResult<Vec<u8>> {
    if chart
        .groups
        .iter()
        .all(|group| group.loaded_index.is_some())
    {
        return Ok(existing_chart_xml.to_vec());
    }
    let topology = scan_loaded_chart_group_xml_topology(existing_chart_xml)?;
    let insertion_tail = topology
        .first_axis_start
        .or(topology.plot_area_end)
        .ok_or_else(|| {
            OmError::new(
                OmErrorCode::Parse,
                "loaded chart plot area has no group insertion position",
            )
        })?;
    let mut loaded_seen = vec![false; topology.groups.len()];
    for group in &chart.groups {
        let Some(loaded_index) = group.loaded_index else {
            continue;
        };
        if loaded_index >= topology.groups.len()
            || loaded_seen[loaded_index]
            || topology.groups[loaded_index].raw_name.as_slice() != group.raw_name.as_bytes()
        {
            return Err(OmError::unsupported(
                "loaded chart group identity changed before group insertion",
            ));
        }
        loaded_seen[loaded_index] = true;
    }
    let mut additions = BTreeMap::<usize, Vec<u8>>::new();
    for (group_index, group) in chart.groups.iter().enumerate() {
        if group.loaded_index.is_some() {
            continue;
        }
        let insertion_position = chart.groups[group_index + 1..]
            .iter()
            .find_map(|next| next.loaded_index)
            .and_then(|loaded_index| topology.groups.get(loaded_index).map(|group| group.start))
            .unwrap_or(insertion_tail);
        additions
            .entry(insertion_position)
            .or_default()
            .extend_from_slice(&serialize_loaded_chart_group_shell(group)?);
    }
    apply_loaded_chart_xml_edits(
        existing_chart_xml,
        additions
            .into_iter()
            .map(|(position, addition)| (position, position, addition))
            .collect(),
        "group addition",
    )
}

fn rewrite_loaded_chart_group_removals(
    existing_chart_xml: &[u8],
    chart: &ChartModel,
) -> OmResult<Vec<u8>> {
    let topology = scan_loaded_chart_group_xml_topology(existing_chart_xml)?;
    let loaded_group_names = topology
        .groups
        .iter()
        .map(|group| group.raw_name.clone())
        .collect::<Vec<_>>();
    let sequence = loaded_chart_group_sequence(&loaded_group_names, chart)?;
    if sequence
        .surviving_original_groups
        .iter()
        .all(|survives| *survives)
    {
        return Ok(existing_chart_xml.to_vec());
    }
    let mut edits = Vec::<(usize, usize, Vec<u8>)>::new();
    for (xml_index, token) in sequence.tokens.into_iter().enumerate() {
        let xml_group = &topology.groups[xml_index];
        match token {
            LoadedChartGroupSequenceToken::Original(loaded_index) => {
                if !sequence.surviving_original_groups[loaded_index] {
                    edits.push((xml_group.start, xml_group.end, Vec::new()));
                }
            }
            LoadedChartGroupSequenceToken::Added(_) => {}
        }
    }
    apply_loaded_chart_xml_edits(existing_chart_xml, edits, "group removal")
}

fn serialize_loaded_chart_axis_shell(chart: &ChartModel, axis: &AxisModel) -> OmResult<Vec<u8>> {
    let axis_id = axis.raw_id.as_deref().ok_or_else(|| {
        OmError::new(
            OmErrorCode::InvalidState,
            "new loaded chart axes require a stable axis identity",
        )
    })?;
    let qualified_name = format!("c:{}", chart_axis_xml_name(axis.kind));
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut start = BytesStart::new(qualified_name.clone());
    start.push_attribute((
        "xmlns:c",
        "http://schemas.openxmlformats.org/drawingml/2006/chart",
    ));
    writer
        .write_event(Event::Start(start))
        .map_err(chart_xml_error)?;
    let mut axis_id_element = BytesStart::new("c:axId");
    axis_id_element.push_attribute(("val", axis_id));
    writer
        .write_event(Event::Empty(axis_id_element))
        .map_err(chart_xml_error)?;
    if let Some(deleted) = axis.deleted {
        let mut delete = BytesStart::new("c:delete");
        delete.push_attribute(("val", if deleted { "1" } else { "0" }));
        writer
            .write_event(Event::Empty(delete))
            .map_err(chart_xml_error)?;
    }
    if let Some(cross_axis_id) = chart_axis_cross_target_id(&chart.axes, axis) {
        let mut cross_axis = BytesStart::new("c:crossAx");
        cross_axis.push_attribute(("val", cross_axis_id.as_str()));
        writer
            .write_event(Event::Empty(cross_axis))
            .map_err(chart_xml_error)?;
    }
    writer
        .write_event(Event::End(BytesEnd::new(qualified_name)))
        .map_err(chart_xml_error)?;
    Ok(writer.into_inner().into_inner())
}

fn element_val_attribute(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> OmResult<Option<String>> {
    for attr in element.attributes() {
        let attr = attr.map_err(chart_xml_error)?;
        if attr.key.as_ref() == b"val" {
            return Ok(Some(
                attr.decode_and_unescape_value(decoder)
                    .map_err(chart_xml_error)?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn rewrite_loaded_chart_axis_additions(
    existing_chart_xml: &[u8],
    chart: &ChartModel,
) -> OmResult<Vec<u8>> {
    struct AxisCrossSpan {
        raw_id: Option<String>,
        start: usize,
        start_tag_end: usize,
        end: usize,
    }
    struct AxisChildSpan {
        start: usize,
        start_tag_end: usize,
        end: usize,
    }
    struct AxisDeleteSpan {
        span: AxisChildSpan,
        value: bool,
    }
    struct AxisSpan {
        kind: ChartAxisKind,
        raw_id: Option<String>,
        qualified_name: String,
        start: usize,
        start_tag_end: usize,
        end: usize,
        end_tag_start: usize,
        axis_id: Option<AxisChildSpan>,
        scaling: Option<AxisChildSpan>,
        delete: Option<AxisDeleteSpan>,
        cross_axis: Option<AxisCrossSpan>,
    }
    let mut reader = NsReader::from_reader(Cursor::new(existing_chart_xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut element_stack = Vec::<Vec<u8>>::new();
    let mut element_chart_namespace_stack = Vec::<bool>::new();
    let mut loaded_axes = Vec::<AxisSpan>::new();
    let mut active_axis_depth = None::<usize>;
    let mut saw_axis = false;
    let mut first_after_axes_start = None::<usize>;
    let mut plot_area_end = None::<usize>;
    let parse_axis_delete =
        |element: &BytesStart<'_>, decoder: quick_xml::encoding::Decoder| -> OmResult<bool> {
            Ok(match element_val_attribute(element, decoder)?.as_deref() {
                None | Some("1") => true,
                Some("0") => false,
                Some(value) if value.eq_ignore_ascii_case("true") => true,
                Some(value) if value.eq_ignore_ascii_case("false") => false,
                Some(value) => {
                    return Err(OmError::new(
                        OmErrorCode::Parse,
                        format!("invalid chart axis delete value: {value}"),
                    ));
                }
            })
        };
    let is_chart_namespace = |namespace: ResolveResult<'_>| match namespace {
        ResolveResult::Bound(namespace) => namespace.as_ref() == CHART_XML_NAMESPACE,
        ResolveResult::Unbound | ResolveResult::Unknown(_) => false,
    };
    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| OmError::new(OmErrorCode::InvalidState, "chart XML position overflow"))?;
        match reader.read_resolved_event_into(&mut buffer) {
            Ok((namespace, Event::Start(element))) => {
                let is_chart_namespace = is_chart_namespace(namespace);
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                if is_chart_namespace
                    && element_stack.last().map(Vec::as_slice) == Some(b"plotArea".as_slice())
                    && element_chart_namespace_stack.last() == Some(&true)
                {
                    if let Some(axis_kind) = chart_axis_kind_from_xml_name(local_name.as_slice()) {
                        loaded_axes.push(AxisSpan {
                            kind: axis_kind,
                            raw_id: None,
                            qualified_name: String::from_utf8_lossy(element.name().as_ref())
                                .into_owned(),
                            start: event_start,
                            start_tag_end: usize::try_from(reader.buffer_position()).map_err(
                                |_| {
                                    OmError::new(
                                        OmErrorCode::InvalidState,
                                        "chart XML position overflow",
                                    )
                                },
                            )?,
                            end: 0,
                            end_tag_start: 0,
                            axis_id: None,
                            scaling: None,
                            delete: None,
                            cross_axis: None,
                        });
                        active_axis_depth = Some(element_stack.len() + 1);
                        saw_axis = true;
                    } else if saw_axis && first_after_axes_start.is_none() {
                        first_after_axes_start = Some(event_start);
                    }
                } else if local_name.as_slice() == b"axId"
                    && is_chart_namespace
                    && active_axis_depth == Some(element_stack.len())
                    && let Some(axis) = loaded_axes.last_mut()
                {
                    if axis.axis_id.is_some() {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            "loaded chart axis has duplicate axId elements",
                        ));
                    }
                    axis.raw_id = element_val_attribute(&element, reader.decoder())?;
                    let start_tag_end =
                        usize::try_from(reader.buffer_position()).map_err(|_| {
                            OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                        })?;
                    axis.axis_id = Some(AxisChildSpan {
                        start: event_start,
                        start_tag_end,
                        end: 0,
                    });
                } else if local_name.as_slice() == b"scaling"
                    && is_chart_namespace
                    && active_axis_depth == Some(element_stack.len())
                    && let Some(axis) = loaded_axes.last_mut()
                {
                    let start_tag_end =
                        usize::try_from(reader.buffer_position()).map_err(|_| {
                            OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                        })?;
                    axis.scaling = Some(AxisChildSpan {
                        start: event_start,
                        start_tag_end,
                        end: 0,
                    });
                } else if local_name.as_slice() == b"delete"
                    && is_chart_namespace
                    && active_axis_depth == Some(element_stack.len())
                    && let Some(axis) = loaded_axes.last_mut()
                {
                    if axis.delete.is_some() {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            "loaded chart axis has duplicate delete elements",
                        ));
                    }
                    let start_tag_end =
                        usize::try_from(reader.buffer_position()).map_err(|_| {
                            OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                        })?;
                    axis.delete = Some(AxisDeleteSpan {
                        span: AxisChildSpan {
                            start: event_start,
                            start_tag_end,
                            end: 0,
                        },
                        value: parse_axis_delete(&element, reader.decoder())?,
                    });
                } else if local_name.as_slice() == b"crossAx"
                    && is_chart_namespace
                    && active_axis_depth == Some(element_stack.len())
                    && let Some(axis) = loaded_axes.last_mut()
                {
                    if axis.cross_axis.is_some() {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            "loaded chart axis has duplicate crossAx elements",
                        ));
                    }
                    let start_tag_end =
                        usize::try_from(reader.buffer_position()).map_err(|_| {
                            OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                        })?;
                    axis.cross_axis = Some(AxisCrossSpan {
                        raw_id: element_val_attribute(&element, reader.decoder())?,
                        start: event_start,
                        start_tag_end,
                        end: 0,
                    });
                }
                element_stack.push(local_name);
                element_chart_namespace_stack.push(is_chart_namespace);
            }
            Ok((namespace, Event::Empty(element))) => {
                let is_chart_namespace = is_chart_namespace(namespace);
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                if is_chart_namespace
                    && element_stack.last().map(Vec::as_slice) == Some(b"plotArea".as_slice())
                    && element_chart_namespace_stack.last() == Some(&true)
                {
                    if let Some(axis_kind) = chart_axis_kind_from_xml_name(local_name.as_slice()) {
                        let event_end =
                            usize::try_from(reader.buffer_position()).map_err(|_| {
                                OmError::new(
                                    OmErrorCode::InvalidState,
                                    "chart XML position overflow",
                                )
                            })?;
                        loaded_axes.push(AxisSpan {
                            kind: axis_kind,
                            raw_id: None,
                            qualified_name: String::from_utf8_lossy(element.name().as_ref())
                                .into_owned(),
                            start: event_start,
                            start_tag_end: event_end,
                            end: event_end,
                            end_tag_start: event_start,
                            axis_id: None,
                            scaling: None,
                            delete: None,
                            cross_axis: None,
                        });
                        saw_axis = true;
                    } else if saw_axis && first_after_axes_start.is_none() {
                        first_after_axes_start = Some(event_start);
                    }
                } else if local_name.as_slice() == b"axId"
                    && is_chart_namespace
                    && active_axis_depth == Some(element_stack.len())
                    && let Some(axis) = loaded_axes.last_mut()
                {
                    if axis.axis_id.is_some() {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            "loaded chart axis has duplicate axId elements",
                        ));
                    }
                    axis.raw_id = element_val_attribute(&element, reader.decoder())?;
                    let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                    })?;
                    axis.axis_id = Some(AxisChildSpan {
                        start: event_start,
                        start_tag_end: event_end,
                        end: event_end,
                    });
                } else if local_name.as_slice() == b"scaling"
                    && is_chart_namespace
                    && active_axis_depth == Some(element_stack.len())
                    && let Some(axis) = loaded_axes.last_mut()
                {
                    let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                    })?;
                    axis.scaling = Some(AxisChildSpan {
                        start: event_start,
                        start_tag_end: event_end,
                        end: event_end,
                    });
                } else if local_name.as_slice() == b"delete"
                    && is_chart_namespace
                    && active_axis_depth == Some(element_stack.len())
                    && let Some(axis) = loaded_axes.last_mut()
                {
                    if axis.delete.is_some() {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            "loaded chart axis has duplicate delete elements",
                        ));
                    }
                    let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                    })?;
                    axis.delete = Some(AxisDeleteSpan {
                        span: AxisChildSpan {
                            start: event_start,
                            start_tag_end: event_end,
                            end: event_end,
                        },
                        value: parse_axis_delete(&element, reader.decoder())?,
                    });
                } else if local_name.as_slice() == b"crossAx"
                    && is_chart_namespace
                    && active_axis_depth == Some(element_stack.len())
                    && let Some(axis) = loaded_axes.last_mut()
                {
                    if axis.cross_axis.is_some() {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            "loaded chart axis has duplicate crossAx elements",
                        ));
                    }
                    let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                    })?;
                    axis.cross_axis = Some(AxisCrossSpan {
                        raw_id: element_val_attribute(&element, reader.decoder())?,
                        start: event_start,
                        start_tag_end: event_end,
                        end: event_end,
                    });
                }
            }
            Ok((namespace, Event::End(element))) => {
                let is_chart_namespace = is_chart_namespace(namespace);
                let element_name = element.name();
                let local_name = xml_local_name(element_name.as_ref());
                if is_chart_namespace
                    && active_axis_depth
                        .is_some_and(|axis_depth| element_stack.len() == axis_depth + 1)
                    && let Some(axis) = loaded_axes.last_mut()
                {
                    let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                    })?;
                    match local_name {
                        b"axId" => {
                            if let Some(span) = axis.axis_id.as_mut() {
                                span.end = event_end;
                            }
                        }
                        b"scaling" => {
                            if let Some(span) = axis.scaling.as_mut() {
                                span.end = event_end;
                            }
                        }
                        b"delete" => {
                            if let Some(delete) = axis.delete.as_mut() {
                                delete.span.end = event_end;
                            }
                        }
                        _ => {}
                    }
                }
                if is_chart_namespace
                    && active_axis_depth
                        .is_some_and(|axis_depth| element_stack.len() == axis_depth + 1)
                    && local_name == b"crossAx"
                    && let Some(axis) = loaded_axes.last_mut()
                    && let Some(cross_axis) = axis.cross_axis.as_mut()
                {
                    cross_axis.end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                    })?;
                }
                if is_chart_namespace
                    && active_axis_depth == Some(element_stack.len())
                    && element_stack.len() >= 2
                    && element_stack[element_stack.len() - 2].as_slice() == b"plotArea"
                    && element_chart_namespace_stack[element_stack.len() - 2]
                    && let Some(axis_kind) = chart_axis_kind_from_xml_name(local_name)
                {
                    let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                        OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                    })?;
                    let axis = loaded_axes.last_mut().ok_or_else(|| {
                        OmError::new(
                            OmErrorCode::Parse,
                            "loaded chart axis end has no matching start",
                        )
                    })?;
                    if axis.end != 0 || axis.kind != axis_kind {
                        return Err(OmError::new(
                            OmErrorCode::Parse,
                            "loaded chart axis elements are not properly nested",
                        ));
                    }
                    axis.end = event_end;
                    axis.end_tag_start = event_start;
                    active_axis_depth = None;
                }
                if is_chart_namespace && local_name == b"plotArea" {
                    plot_area_end = Some(event_start);
                }
                element_stack.pop();
                element_chart_namespace_stack.pop();
            }
            Ok((_, Event::Eof)) => break,
            Ok((_, _)) => {}
            Err(error) => return Err(chart_xml_error(error)),
        }
        buffer.clear();
    }
    if loaded_axes.iter().any(|axis| {
        axis.end == 0
            || axis.end_tag_start == 0
            || axis
                .cross_axis
                .as_ref()
                .is_some_and(|cross_axis| cross_axis.end == 0)
            || axis.axis_id.as_ref().is_some_and(|span| span.end == 0)
            || axis.scaling.as_ref().is_some_and(|span| span.end == 0)
            || axis
                .delete
                .as_ref()
                .is_some_and(|delete| delete.span.end == 0)
    }) {
        return Err(OmError::new(
            OmErrorCode::Parse,
            "loaded chart axis or crossAx has no closing XML boundary",
        ));
    }
    if loaded_axes
        .iter()
        .any(|axis| axis.axis_id.is_none() || axis.raw_id.is_none())
    {
        return Err(OmError::new(
            OmErrorCode::Parse,
            "loaded chart axis is missing a stable axId",
        ));
    }
    let mut loaded_axis_ids = BTreeSet::new();
    for axis_id in loaded_axes.iter().filter_map(|axis| axis.raw_id.as_ref()) {
        if !loaded_axis_ids.insert(axis_id) {
            return Err(OmError::new(
                OmErrorCode::Parse,
                "loaded chart axis identity is duplicated in XML",
            ));
        }
    }
    let mut loaded_seen = vec![false; loaded_axes.len()];
    let mut loaded_model_indices = vec![None; loaded_axes.len()];
    let mut additions = Vec::<&AxisModel>::new();
    for (model_index, axis) in chart.axes.iter().enumerate() {
        let loaded_index = axis.raw_id.as_ref().and_then(|axis_id| {
            loaded_axes
                .iter()
                .position(|loaded| loaded.raw_id.as_ref() == Some(axis_id))
        });
        if let Some(loaded_index) = loaded_index {
            if loaded_seen[loaded_index] {
                return Err(OmError::unsupported(
                    "loaded chart axis identity is duplicated in the model",
                ));
            }
            loaded_seen[loaded_index] = true;
            loaded_model_indices[loaded_index] = Some(model_index);
        } else {
            additions.push(axis);
        }
    }
    let axis_topology_changed = !additions.is_empty() || loaded_seen.iter().any(|seen| !*seen);
    let mut edits = Vec::<(usize, usize, Vec<u8>)>::new();
    for (loaded_index, loaded_axis) in loaded_axes.iter().enumerate() {
        if loaded_seen[loaded_index] {
            continue;
        }
        if loaded_axis.raw_id.as_ref().is_some_and(|axis_id| {
            chart
                .groups
                .iter()
                .any(|group| group.axis_ids.contains(axis_id))
        }) {
            return Err(OmError::unsupported(
                "removing a chart axis that is still referenced by a group is not supported",
            ));
        }
        edits.push((loaded_axis.start, loaded_axis.end, Vec::new()));
    }
    let rewrite_cross_axis_start =
        |span: &AxisCrossSpan, cross_axis_id: &str| -> OmResult<Vec<u8>> {
            let mut reader = Reader::from_reader(Cursor::new(
                &existing_chart_xml[span.start..span.start_tag_end],
            ));
            reader.config_mut().trim_text(false);
            let mut buffer = Vec::new();
            let event = reader
                .read_event_into(&mut buffer)
                .map_err(chart_xml_error)?;
            let (element, is_empty) = match event {
                Event::Start(element) => (element, false),
                Event::Empty(element) => (element, true),
                _ => {
                    return Err(OmError::new(
                        OmErrorCode::Parse,
                        "loaded chart crossAx span does not start with an element",
                    ));
                }
            };
            let qualified_name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
            let mut rewritten = BytesStart::new(qualified_name);
            let mut wrote_value = false;
            for attr in element.attributes() {
                let attr = attr.map_err(chart_xml_error)?;
                let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
                let value = attr
                    .decode_and_unescape_value(reader.decoder())
                    .map_err(chart_xml_error)?
                    .into_owned();
                if attr.key.as_ref() == b"val" {
                    rewritten.push_attribute((key.as_str(), cross_axis_id));
                    wrote_value = true;
                } else {
                    rewritten.push_attribute((key.as_str(), value.as_str()));
                }
            }
            if !wrote_value {
                rewritten.push_attribute(("val", cross_axis_id));
            }
            let mut writer = Writer::new(Cursor::new(Vec::new()));
            writer
                .write_event(if is_empty {
                    Event::Empty(rewritten)
                } else {
                    Event::Start(rewritten)
                })
                .map_err(chart_xml_error)?;
            Ok(writer.into_inner().into_inner())
        };
    let rewrite_delete_start = |span: &AxisChildSpan, deleted: bool| -> OmResult<Vec<u8>> {
        let mut reader = Reader::from_reader(Cursor::new(
            &existing_chart_xml[span.start..span.start_tag_end],
        ));
        reader.config_mut().trim_text(false);
        let mut buffer = Vec::new();
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(chart_xml_error)?;
        let (element, is_empty) = match event {
            Event::Start(element) => (element, false),
            Event::Empty(element) => (element, true),
            _ => {
                return Err(OmError::new(
                    OmErrorCode::Parse,
                    "loaded chart delete span does not start with an element",
                ));
            }
        };
        let qualified_name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
        let mut rewritten = BytesStart::new(qualified_name);
        let replacement = if deleted { "1" } else { "0" };
        let mut wrote_value = false;
        for attr in element.attributes() {
            let attr = attr.map_err(chart_xml_error)?;
            let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
            let value = attr
                .decode_and_unescape_value(reader.decoder())
                .map_err(chart_xml_error)?
                .into_owned();
            if attr.key.as_ref() == b"val" {
                rewritten.push_attribute((key.as_str(), replacement));
                wrote_value = true;
            } else {
                rewritten.push_attribute((key.as_str(), value.as_str()));
            }
        }
        if !wrote_value {
            rewritten.push_attribute(("val", replacement));
        }
        let mut writer = Writer::new(Cursor::new(Vec::new()));
        writer
            .write_event(if is_empty {
                Event::Empty(rewritten)
            } else {
                Event::Start(rewritten)
            })
            .map_err(chart_xml_error)?;
        Ok(writer.into_inner().into_inner())
    };
    for (loaded_index, loaded_axis) in loaded_axes.iter().enumerate() {
        let Some(model_index) = loaded_model_indices[loaded_index] else {
            continue;
        };
        let Some(expected) = chart.axes[model_index].deleted else {
            continue;
        };
        match loaded_axis.delete.as_ref() {
            Some(delete) if delete.value != expected => edits.push((
                delete.span.start,
                delete.span.start_tag_end,
                rewrite_delete_start(&delete.span, expected)?,
            )),
            Some(_) => {}
            None => {
                let insertion_position = loaded_axis
                    .scaling
                    .as_ref()
                    .map(|span| span.end)
                    .or_else(|| loaded_axis.axis_id.as_ref().map(|span| span.end))
                    .unwrap_or(loaded_axis.start_tag_end);
                let qualified_name = loaded_axis
                    .qualified_name
                    .rsplit_once(':')
                    .map(|(prefix, _)| format!("{prefix}:delete"))
                    .unwrap_or_else(|| "delete".to_string());
                let mut delete = BytesStart::new(qualified_name);
                delete.push_attribute(("val", if expected { "1" } else { "0" }));
                let mut writer = Writer::new(Cursor::new(Vec::new()));
                writer
                    .write_event(Event::Empty(delete))
                    .map_err(chart_xml_error)?;
                edits.push((
                    insertion_position,
                    insertion_position,
                    writer.into_inner().into_inner(),
                ));
            }
        }
    }
    if axis_topology_changed {
        for (loaded_index, loaded_axis) in loaded_axes.iter().enumerate() {
            let Some(model_index) = loaded_model_indices[loaded_index] else {
                continue;
            };
            let expected_cross_axis_id =
                chart_axis_cross_target_id(&chart.axes, &chart.axes[model_index]);
            match (
                loaded_axis.cross_axis.as_ref(),
                expected_cross_axis_id.as_deref(),
            ) {
                (Some(cross_axis), Some(expected))
                    if cross_axis.raw_id.as_deref() != Some(expected) =>
                {
                    edits.push((
                        cross_axis.start,
                        cross_axis.start_tag_end,
                        rewrite_cross_axis_start(cross_axis, expected)?,
                    ));
                }
                (Some(cross_axis), None) => {
                    edits.push((cross_axis.start, cross_axis.end, Vec::new()));
                }
                (None, Some(expected)) => {
                    let mut cross_axis = BytesStart::new("c:crossAx");
                    cross_axis.push_attribute(("val", expected));
                    let mut writer = Writer::new(Cursor::new(Vec::new()));
                    writer
                        .write_event(Event::Empty(cross_axis))
                        .map_err(chart_xml_error)?;
                    edits.push((
                        loaded_axis.end_tag_start,
                        loaded_axis.end_tag_start,
                        writer.into_inner().into_inner(),
                    ));
                }
                _ => {}
            }
        }
    }
    if additions.is_empty() {
        return apply_loaded_chart_xml_edits(existing_chart_xml, edits, "axis update");
    }
    let insertion_position = first_after_axes_start.or(plot_area_end).ok_or_else(|| {
        OmError::new(
            OmErrorCode::Parse,
            "loaded chart plot area has no axis insertion position",
        )
    })?;
    let mut addition = Vec::new();
    for axis in additions {
        addition.extend_from_slice(&serialize_loaded_chart_axis_shell(chart, axis)?);
    }
    edits.push((insertion_position, insertion_position, addition));
    apply_loaded_chart_xml_edits(existing_chart_xml, edits, "axis topology")
}

fn rewrite_loaded_chart_series_topology(
    existing_chart_xml: &[u8],
    chart: &ChartModel,
) -> OmResult<Vec<u8>> {
    let mut reader = Reader::from_reader(Cursor::new(existing_chart_xml));
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut frames = Vec::<LoadedChartXmlFrame>::new();
    let mut element_spans = BTreeMap::<usize, LoadedChartXmlElementSpan>::new();
    let mut groups = Vec::<LoadedChartXmlGroupSpan>::new();
    let mut series_spans = Vec::<LoadedChartXmlSeriesSpan>::new();

    let parse_val = |element: &BytesStart<'_>,
                     decoder: quick_xml::encoding::Decoder|
     -> OmResult<Option<u32>> {
        for attr in element.attributes() {
            let attr = attr.map_err(chart_xml_error)?;
            if attr.key.as_ref() == b"val" {
                return Ok(attr
                    .decode_and_unescape_value(decoder)
                    .map_err(chart_xml_error)?
                    .parse::<u32>()
                    .ok());
            }
        }
        Ok(None)
    };

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| OmError::new(OmErrorCode::InvalidState, "chart XML position overflow"))?;
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) => {
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                let parent_local_name = frames.last().map(|frame| frame.local_name.clone());
                let parent_start = frames.last().map(|frame| frame.start);
                if let Some(parent) = frames.last_mut() {
                    parent.child_element_count += 1;
                }

                let mut group_index = frames.last().and_then(|frame| frame.group_index);
                if parent_local_name.as_deref() == Some(b"plotArea".as_slice())
                    && chart_type_from_group_name(local_name.as_slice()).is_some()
                {
                    group_index = Some(groups.len());
                    groups.push(LoadedChartXmlGroupSpan {
                        local_name: local_name.clone(),
                        end_tag_start: 0,
                        last_direct_series_end: None,
                        first_after_series_start: None,
                        direct_ext_lst_start: None,
                    });
                } else if let Some(group_index) = group_index {
                    let is_direct_group_child = frames.last().is_some_and(|frame| {
                        frame.group_index == Some(group_index)
                            && chart_type_from_group_name(frame.local_name.as_slice()).is_some()
                    });
                    if is_direct_group_child {
                        if local_name.as_slice() == b"extLst" {
                            groups[group_index].direct_ext_lst_start = Some(event_start);
                        }
                        if local_name.as_slice() != b"ser"
                            && !chart_group_child_precedes_series(local_name.as_slice())
                            && groups[group_index].first_after_series_start.is_none()
                        {
                            groups[group_index].first_after_series_start = Some(event_start);
                        }
                    }
                }

                let mut series_index = frames.last().and_then(|frame| frame.series_index);
                if local_name.as_slice() == b"ser" && series_index.is_none() {
                    let wrapper_frame = frames.iter().rev().find(|frame| {
                        is_chart_filtered_series_wrapper_name(frame.local_name.as_slice())
                    });
                    let is_filtered = wrapper_frame.is_some_and(|wrapper| {
                        let mut ancestors = frames.iter().rev();
                        ancestors
                            .next()
                            .is_some_and(|frame| frame.start == wrapper.start)
                            && ancestors
                                .next()
                                .is_some_and(|frame| frame.local_name.as_slice() == b"ext")
                            && ancestors
                                .next()
                                .is_some_and(|frame| frame.local_name.as_slice() == b"extLst")
                    });
                    let group_index = group_index.ok_or_else(|| {
                        OmError::unsupported(
                            "Series.IsFiltered requires a series inside a chart group",
                        )
                    })?;
                    let wrapper_start = is_filtered
                        .then(|| wrapper_frame.map(|frame| frame.start))
                        .flatten();
                    let extension_start = wrapper_start.and_then(|wrapper_start| {
                        frames
                            .iter()
                            .rev()
                            .skip_while(|frame| frame.start != wrapper_start)
                            .nth(1)
                            .filter(|frame| frame.local_name.as_slice() == b"ext")
                            .map(|frame| frame.start)
                    });
                    series_index = Some(series_spans.len());
                    series_spans.push(LoadedChartXmlSeriesSpan {
                        start: event_start,
                        end: 0,
                        raw_index: None,
                        is_filtered,
                        wrapper_start,
                        extension_start,
                        group_index,
                    });
                } else if let Some(series_index) = series_index
                    && frames.last().is_some_and(|frame| {
                        frame.series_index == Some(series_index)
                            && frame.local_name.as_slice() == b"ser"
                    })
                {
                    if local_name.as_slice() == b"idx" {
                        series_spans[series_index].raw_index =
                            parse_val(&element, reader.decoder())?;
                    }
                }

                frames.push(LoadedChartXmlFrame {
                    local_name,
                    start: event_start,
                    parent_start,
                    child_element_count: 0,
                    group_index,
                    series_index,
                });
            }
            Ok(Event::Empty(element)) => {
                if let Some(parent) = frames.last_mut() {
                    parent.child_element_count += 1;
                }
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                if let Some(series_index) = frames.last().and_then(|frame| frame.series_index)
                    && frames
                        .last()
                        .is_some_and(|frame| frame.local_name.as_slice() == b"ser")
                {
                    if local_name.as_slice() == b"idx" {
                        series_spans[series_index].raw_index =
                            parse_val(&element, reader.decoder())?;
                    }
                }
                if let Some(group_index) = frames.last().and_then(|frame| frame.group_index)
                    && frames.last().is_some_and(|frame| {
                        chart_type_from_group_name(frame.local_name.as_slice()).is_some()
                    })
                    && local_name.as_slice() != b"ser"
                    && !chart_group_child_precedes_series(local_name.as_slice())
                    && groups[group_index].first_after_series_start.is_none()
                {
                    groups[group_index].first_after_series_start = Some(event_start);
                }
            }
            Ok(Event::End(element)) => {
                let event_end = usize::try_from(reader.buffer_position()).map_err(|_| {
                    OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                })?;
                let frame = frames.pop().ok_or_else(|| {
                    OmError::new(OmErrorCode::Parse, "chart XML has an unmatched end element")
                })?;
                if xml_local_name(element.name().as_ref()) != frame.local_name.as_slice() {
                    return Err(OmError::new(
                        OmErrorCode::Parse,
                        "chart XML has mismatched element nesting",
                    ));
                }
                if frame.local_name.as_slice() == b"ser"
                    && let Some(series_index) = frame.series_index
                    && series_spans[series_index].start == frame.start
                {
                    series_spans[series_index].end = event_end;
                    if !series_spans[series_index].is_filtered {
                        groups[series_spans[series_index].group_index].last_direct_series_end =
                            Some(event_end);
                    }
                }
                if chart_type_from_group_name(frame.local_name.as_slice()).is_some()
                    && let Some(group_index) = frame.group_index
                {
                    groups[group_index].end_tag_start = event_start;
                }
                element_spans.insert(
                    frame.start,
                    LoadedChartXmlElementSpan {
                        start: frame.start,
                        end: event_end,
                        end_tag_start: event_start,
                        parent_start: frame.parent_start,
                        child_element_count: frame.child_element_count,
                    },
                );
            }
            Ok(Event::Eof) => break,
            Ok(_) => {}
            Err(error) => return Err(chart_xml_error(error)),
        }
        buffer.clear();
    }

    if groups.is_empty() {
        return Err(OmError::unsupported(
            "loaded series topology patching requires at least one chart group",
        ));
    }
    let group_sequence = if !chart.groups.is_empty() {
        let loaded_group_names = groups
            .iter()
            .map(|group| group.local_name.clone())
            .collect::<Vec<_>>();
        let sequence = loaded_chart_group_sequence(&loaded_group_names, chart)?;
        if !chart_group_overlay_is_stable(chart) {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                "loaded chart series-to-group partition is inconsistent",
            ));
        }
        Some(sequence)
    } else {
        None
    };

    let mut loaded_series_by_raw_index = BTreeMap::<u32, usize>::new();
    let mut duplicate_raw_indices = BTreeSet::<u32>::new();
    for (loaded_index, loaded) in series_spans.iter().enumerate() {
        let Some(raw_index) = loaded.raw_index else {
            continue;
        };
        if loaded_series_by_raw_index
            .insert(raw_index, loaded_index)
            .is_some()
        {
            duplicate_raw_indices.insert(raw_index);
        }
    }
    if !duplicate_raw_indices.is_empty() {
        return Err(OmError::unsupported(
            "loaded chart contains duplicate series c:idx identities",
        ));
    }
    let mut model_series_by_raw_index = BTreeMap::<u32, usize>::new();
    let mut model_to_loaded = vec![None; chart.series.len()];
    let mut model_group_indices = vec![None; chart.series.len()];
    for (model_index, series) in chart.series.iter().enumerate() {
        let raw_index = series.raw_index.ok_or_else(|| {
            OmError::unsupported(
                "loaded chart topology changes require stable series c:idx identities",
            )
        })?;
        if model_series_by_raw_index
            .insert(raw_index, model_index)
            .is_some()
        {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!("chart model contains duplicate series c:idx {raw_index}"),
            ));
        }
        model_to_loaded[model_index] = loaded_series_by_raw_index.get(&raw_index).copied();
        model_group_indices[model_index] = if chart.groups.is_empty() {
            loaded_series_by_raw_index
                .get(&raw_index)
                .map(|loaded_index| series_spans[*loaded_index].group_index)
        } else {
            let model_group_index = chart_group_index_for_series_raw_index(chart, raw_index)?;
            Some(
                group_sequence
                    .as_ref()
                    .and_then(|sequence| sequence.model_to_xml.get(model_group_index).copied())
                    .ok_or_else(|| {
                        OmError::new(
                            OmErrorCode::InvalidState,
                            "chart series target group has no loaded XML mapping",
                        )
                    })?,
            )
        };
    }

    #[derive(Debug)]
    struct ByteEdit {
        start: usize,
        end: usize,
        replacement: Vec<u8>,
    }

    #[derive(Debug, Default)]
    struct GroupFilterEdits {
        direct_additions: Vec<Vec<u8>>,
        filtered_additions: Vec<Vec<u8>>,
        removed_extension_starts: BTreeSet<usize>,
    }

    let mut edits = Vec::<ByteEdit>::new();
    let mut group_edits = (0..groups.len())
        .map(|_| GroupFilterEdits::default())
        .collect::<Vec<_>>();
    let mut removal_starts = BTreeSet::<usize>::new();
    let mut relocated_series_xml = vec![None::<Vec<u8>>; chart.series.len()];

    for loaded in &series_spans {
        let model_index = loaded
            .raw_index
            .and_then(|raw_index| model_series_by_raw_index.get(&raw_index).copied());
        let should_relocate = model_index.is_some_and(|model_index| {
            model_group_indices[model_index] != Some(loaded.group_index)
                || chart.series[model_index].is_filtered != loaded.is_filtered
        });
        if model_index.is_some() && !should_relocate {
            continue;
        }
        if let Some(model_index) = model_index {
            relocated_series_xml[model_index] =
                Some(existing_chart_xml[loaded.start..loaded.end].to_vec());
        }
        if !loaded.is_filtered {
            removal_starts.insert(loaded.start);
            edits.push(ByteEdit {
                start: loaded.start,
                end: loaded.end,
                replacement: Vec::new(),
            });
        } else {
            let removal_start = loaded
                .extension_start
                .filter(|extension_start| {
                    element_spans
                        .get(extension_start)
                        .is_some_and(|span| span.child_element_count == 1)
                })
                .or_else(|| {
                    loaded.wrapper_start.filter(|wrapper_start| {
                        element_spans
                            .get(wrapper_start)
                            .is_some_and(|span| span.child_element_count == 1)
                    })
                })
                .unwrap_or(loaded.start);
            let removal = element_spans.get(&removal_start).ok_or_else(|| {
                OmError::new(
                    OmErrorCode::Parse,
                    "filtered-series wrapper span is missing",
                )
            })?;
            if removal_starts.insert(removal.start) {
                edits.push(ByteEdit {
                    start: removal.start,
                    end: removal.end,
                    replacement: Vec::new(),
                });
            }
            if loaded.extension_start == Some(removal_start) {
                group_edits[loaded.group_index]
                    .removed_extension_starts
                    .insert(removal_start);
            }
        }
    }

    for (model_index, series) in chart.series.iter().enumerate() {
        let needs_addition =
            model_to_loaded[model_index].is_none() || relocated_series_xml[model_index].is_some();
        if !needs_addition {
            continue;
        }
        let raw_index = series.raw_index.ok_or_else(|| {
            OmError::new(
                OmErrorCode::InvalidState,
                "chart series has no c:idx identity",
            )
        })?;
        let group_index = model_group_indices[model_index].ok_or_else(|| {
            OmError::unsupported("chart series could not be assigned to a loaded chart group")
        })?;
        let series_xml = relocated_series_xml[model_index].take().unwrap_or_else(|| {
            format!(
                r#"<c:ser xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:idx val="{raw_index}"/><c:order val="{}"/></c:ser>"#,
                series.order.unwrap_or(model_index as u32),
            )
            .into_bytes()
        });
        if series.is_filtered {
            group_edits[group_index]
                .filtered_additions
                .push(rewrite_chart_series_outer_name(&series_xml, "c15:ser")?);
        } else {
            group_edits[group_index]
                .direct_additions
                .push(rewrite_chart_series_outer_name(&series_xml, "c:ser")?);
        }
    }

    for (group_index, group) in groups.iter().enumerate() {
        let group_edit = &group_edits[group_index];
        if !group_edit.direct_additions.is_empty() {
            let insertion_position = group
                .last_direct_series_end
                .or(group.first_after_series_start)
                .unwrap_or(group.end_tag_start);
            edits.push(ByteEdit {
                start: insertion_position,
                end: insertion_position,
                replacement: group_edit.direct_additions.concat(),
            });
        }

        if !group_edit.filtered_additions.is_empty() {
            let group_chart_type = chart_type_from_group_name(group.local_name.as_slice())
                .ok_or_else(|| {
                    OmError::unsupported(
                        "Series.IsFiltered requires a recognized loaded chart group",
                    )
                })?;
            let wrapper_name =
                chart_filtered_series_wrapper_name(&group_chart_type).ok_or_else(|| {
                    OmError::unsupported(format!(
                        "Series.IsFiltered is unavailable for loaded {} groups",
                        String::from_utf8_lossy(group.local_name.as_slice())
                    ))
                })?;
            let mut extension_xml = Vec::new();
            for series_xml in &group_edit.filtered_additions {
                extension_xml.extend_from_slice(
                    format!(
                        r#"<c:ext xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" uri="{{02D57815-91ED-43cb-92C2-25804820EDAC}}"><c15:{wrapper_name} xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">"#
                    )
                    .as_bytes(),
                );
                extension_xml.extend_from_slice(series_xml);
                extension_xml
                    .extend_from_slice(format!("</c15:{wrapper_name}></c:ext>").as_bytes());
            }
            if let Some(ext_lst_start) = group.direct_ext_lst_start {
                let ext_lst = element_spans.get(&ext_lst_start).ok_or_else(|| {
                    OmError::new(
                        OmErrorCode::Parse,
                        "chart-group extension list span is missing",
                    )
                })?;
                edits.push(ByteEdit {
                    start: ext_lst.end_tag_start,
                    end: ext_lst.end_tag_start,
                    replacement: extension_xml,
                });
            } else {
                let mut ext_lst_xml = br#"<c:extLst xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">"#.to_vec();
                ext_lst_xml.extend_from_slice(&extension_xml);
                ext_lst_xml.extend_from_slice(b"</c:extLst>");
                edits.push(ByteEdit {
                    start: group.end_tag_start,
                    end: group.end_tag_start,
                    replacement: ext_lst_xml,
                });
            }
        } else if let Some(ext_lst_start) = group.direct_ext_lst_start {
            let ext_lst = element_spans.get(&ext_lst_start).ok_or_else(|| {
                OmError::new(
                    OmErrorCode::Parse,
                    "chart-group extension list span is missing",
                )
            })?;
            if ext_lst.child_element_count > 0
                && ext_lst.child_element_count == group_edit.removed_extension_starts.len()
                && ext_lst.parent_start.is_some()
                && removal_starts.insert(ext_lst.start)
            {
                edits.retain(|edit| {
                    !group_edit.removed_extension_starts.contains(&edit.start)
                        && !(edit.start > ext_lst.start && edit.end < ext_lst.end)
                });
                edits.push(ByteEdit {
                    start: ext_lst.start,
                    end: ext_lst.end,
                    replacement: Vec::new(),
                });
            }
        }
    }

    edits.sort_by_key(|edit| (edit.start, if edit.start == edit.end { 0 } else { 1 }));
    let mut rewritten = Vec::with_capacity(existing_chart_xml.len());
    let mut cursor = 0usize;
    for edit in edits {
        if edit.start < cursor {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                "filtered-series structural edits overlap",
            ));
        }
        rewritten.extend_from_slice(&existing_chart_xml[cursor..edit.start]);
        rewritten.extend_from_slice(&edit.replacement);
        cursor = edit.end;
    }
    rewritten.extend_from_slice(&existing_chart_xml[cursor..]);
    Ok(rewritten)
}

fn chart_group_direct_property_name(local_name: &[u8]) -> bool {
    matches!(
        local_name,
        b"barDir"
            | b"grouping"
            | b"shape"
            | b"marker"
            | b"scatterStyle"
            | b"radarStyle"
            | b"ofPieType"
            | b"wireframe"
            | b"varyColors"
            | b"gapWidth"
            | b"gapDepth"
            | b"overlap"
            | b"firstSliceAng"
            | b"bubbleScale"
            | b"showNegBubbles"
            | b"bubble3D"
            | b"holeSize"
            | b"secondPieSize"
            | b"sizeRepresents"
            | b"splitType"
            | b"splitPos"
            | b"dLbls"
            | b"serLines"
            | b"dropLines"
            | b"hiLowLines"
            | b"upDownBars"
    )
}

fn chart_group_line_flag(group: &ChartGroupModel, local_name: &[u8]) -> Option<bool> {
    match local_name {
        b"serLines" => group.has_series_lines,
        b"dropLines" => group.has_drop_lines,
        b"hiLowLines" => group.has_hi_lo_lines,
        b"upDownBars" => group.has_up_down_bars,
        _ => None,
    }
}

fn chart_group_direct_property_value(group: &ChartGroupModel, local_name: &[u8]) -> Option<String> {
    match local_name {
        b"varyColors" => group
            .vary_by_categories
            .map(|value| if value { "1" } else { "0" }.to_string()),
        b"gapWidth" => group.gap_width.map(|value| value.to_string()),
        b"gapDepth" => group.gap_depth.map(|value| value.to_string()),
        b"overlap" => group.overlap.map(|value| value.to_string()),
        b"firstSliceAng" => group.first_slice_angle.map(|value| value.to_string()),
        b"bubbleScale" => group.bubble_scale.map(|value| value.to_string()),
        b"showNegBubbles" => group
            .show_negative_bubbles
            .map(|value| if value { "1" } else { "0" }.to_string()),
        b"bubble3D" => group
            .has_3d_shading
            .map(|value| if value { "1" } else { "0" }.to_string()),
        b"holeSize" => group.doughnut_hole_size.map(|value| value.to_string()),
        b"secondPieSize" => group.second_plot_size.map(|value| value.to_string()),
        b"sizeRepresents" => group.size_represents.map(|value| match value {
            ChartSizeRepresents::Area => "area".to_string(),
            ChartSizeRepresents::Width => "w".to_string(),
        }),
        b"splitType" => group.split_type.map(|value| match value {
            ChartSplitType::Custom => "cust".to_string(),
            ChartSplitType::PercentValue => "percent".to_string(),
            ChartSplitType::Position => "pos".to_string(),
            ChartSplitType::Value => "val".to_string(),
        }),
        b"splitPos" => group.split_value.map(|value| value.to_string()),
        _ => None,
    }
}

fn rewrite_chart_group_val_element(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    replacement: &str,
) -> OmResult<BytesStart<'static>> {
    let qualified_name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
    let mut rewritten = BytesStart::new(qualified_name);
    let mut wrote_value = false;
    for attr in element.attributes() {
        let attr = attr.map_err(chart_xml_error)?;
        let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
        let value = attr
            .decode_and_unescape_value(decoder)
            .map_err(chart_xml_error)?
            .into_owned();
        if attr.key.as_ref() == b"val" {
            rewritten.push_attribute((key.as_str(), replacement));
            wrote_value = true;
        } else {
            rewritten.push_attribute((key.as_str(), value.as_str()));
        }
    }
    if !wrote_value {
        rewritten.push_attribute(("val", replacement));
    }
    Ok(rewritten)
}

fn write_missing_chart_group_properties(
    writer: &mut Writer<Cursor<Vec<u8>>>,
    group: &ChartGroupModel,
    seen: &BTreeSet<Vec<u8>>,
    prefix: bool,
) -> OmResult<()> {
    let names: &[&[u8]] = if prefix {
        &[b"varyColors"]
    } else {
        &[
            b"gapWidth",
            b"gapDepth",
            b"overlap",
            b"firstSliceAng",
            b"bubbleScale",
            b"showNegBubbles",
            b"bubble3D",
            b"holeSize",
            b"secondPieSize",
            b"sizeRepresents",
            b"splitType",
            b"splitPos",
        ]
    };
    for local_name in names {
        if seen.contains(*local_name) {
            continue;
        }
        let Some(value) = chart_group_direct_property_value(group, local_name) else {
            continue;
        };
        let qualified_name = format!("c:{}", String::from_utf8_lossy(local_name));
        let mut element = BytesStart::new(qualified_name);
        element.push_attribute(("val", value.as_str()));
        writer
            .write_event(Event::Empty(element))
            .map_err(chart_xml_error)?;
    }
    if !prefix {
        if !seen.contains(b"dLbls".as_slice())
            && let Some(data_labels) = group.data_labels.as_ref().filter(|labels| labels.dirty)
        {
            writer
                .get_mut()
                .write_all(chart_data_labels_xml_string(data_labels).as_bytes())
                .map_err(chart_xml_error)?;
        }
        for (local_name, qualified_name) in [
            (b"serLines".as_slice(), "c:serLines"),
            (b"dropLines".as_slice(), "c:dropLines"),
            (b"hiLowLines".as_slice(), "c:hiLowLines"),
            (b"upDownBars".as_slice(), "c:upDownBars"),
        ] {
            if !seen.contains(local_name) && chart_group_line_flag(group, local_name) == Some(true)
            {
                writer
                    .write_event(Event::Empty(BytesStart::new(qualified_name)))
                    .map_err(chart_xml_error)?;
            }
        }
    }
    Ok(())
}

fn patch_loaded_chart_group_properties(
    existing_chart_xml: &[u8],
    chart: &ChartModel,
) -> OmResult<Vec<u8>> {
    if chart.groups.is_empty() || !chart.groups.iter().any(|group| group.dirty) {
        return Ok(existing_chart_xml.to_vec());
    }

    let mut reader = Reader::from_reader(Cursor::new(existing_chart_xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::with_capacity(existing_chart_xml.len())));
    let mut buffer = Vec::new();
    let mut element_stack = Vec::<Vec<u8>>::new();
    let mut current_group_index = None::<usize>;
    let mut current_group_depth = None::<usize>;
    let mut seen = vec![BTreeSet::<Vec<u8>>::new(); chart.groups.len()];
    let mut prefix_inserted = vec![false; chart.groups.len()];
    let mut tail_inserted = vec![false; chart.groups.len()];
    let mut axis_id_positions = vec![0usize; chart.groups.len()];
    let mut next_group_index = 0usize;
    let mut skip_depth = 0usize;

    loop {
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(_element)) if skip_depth > 0 => {
                skip_depth += 1;
            }
            Ok(Event::Empty(_)) if skip_depth > 0 => {}
            Ok(Event::End(_)) if skip_depth > 0 => {
                skip_depth -= 1;
            }
            Ok(Event::Text(_) | Event::CData(_) | Event::Comment(_) | Event::GeneralRef(_))
                if skip_depth > 0 => {}
            Ok(Event::Start(element)) => {
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                let parent_name = element_stack.last().map(Vec::as_slice);
                if parent_name == Some(b"plotArea".as_slice())
                    && chart_type_from_group_name(local_name.as_slice()).is_some()
                {
                    let group = chart.groups.get(next_group_index).ok_or_else(|| {
                        OmError::unsupported(
                            "loaded chart group count changed before lossless property patch",
                        )
                    })?;
                    if group.raw_name.as_bytes() != local_name.as_slice() {
                        return Err(OmError::unsupported(
                            "loaded chart group order or type changed before lossless property patch",
                        ));
                    }
                    current_group_index = Some(next_group_index);
                    current_group_depth = Some(element_stack.len() + 1);
                    next_group_index += 1;
                }
                if let Some(group_index) = current_group_index
                    && parent_name
                        .is_some_and(|parent| chart_type_from_group_name(parent).is_some())
                    && chart_group_direct_property_name(local_name.as_slice())
                {
                    seen[group_index].insert(local_name.clone());
                    if local_name.as_slice() == b"dLbls"
                        && let Some(data_labels) = chart.groups[group_index]
                            .data_labels
                            .as_ref()
                            .filter(|labels| labels.dirty)
                    {
                        writer
                            .get_mut()
                            .write_all(chart_data_labels_xml_string(data_labels).as_bytes())
                            .map_err(chart_xml_error)?;
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                    if let Some(enabled) =
                        chart_group_line_flag(&chart.groups[group_index], local_name.as_slice())
                        && !enabled
                    {
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                }
                if let Some(group_index) = current_group_index
                    && parent_name
                        .is_some_and(|parent| chart_type_from_group_name(parent).is_some())
                    && local_name.as_slice() == b"ser"
                    && !prefix_inserted[group_index]
                {
                    write_missing_chart_group_properties(
                        &mut writer,
                        &chart.groups[group_index],
                        &seen[group_index],
                        true,
                    )?;
                    prefix_inserted[group_index] = true;
                }
                if let Some(group_index) = current_group_index
                    && parent_name
                        .is_some_and(|parent| chart_type_from_group_name(parent).is_some())
                    && matches!(local_name.as_slice(), b"axId" | b"extLst")
                    && !tail_inserted[group_index]
                {
                    write_missing_chart_group_properties(
                        &mut writer,
                        &chart.groups[group_index],
                        &seen[group_index],
                        false,
                    )?;
                    tail_inserted[group_index] = true;
                }
                let mut output = element.into_owned();
                if let Some(group_index) = current_group_index
                    && parent_name
                        .is_some_and(|parent| chart_type_from_group_name(parent).is_some())
                    && local_name.as_slice() == b"axId"
                    && chart.groups[group_index].dirty
                {
                    let axis_id_position = axis_id_positions[group_index];
                    let axis_id = chart.groups[group_index]
                        .axis_ids
                        .get(axis_id_position)
                        .ok_or_else(|| {
                            OmError::unsupported(
                                "loaded chart group axis reference count changed before lossless property patch",
                            )
                        })?;
                    axis_id_positions[group_index] += 1;
                    output = rewrite_chart_group_val_element(
                        &output,
                        reader.decoder(),
                        axis_id.as_str(),
                    )?;
                } else if let Some(group_index) = current_group_index
                    && parent_name
                        .is_some_and(|parent| chart_type_from_group_name(parent).is_some())
                    && let Some(value) = chart_group_direct_property_value(
                        &chart.groups[group_index],
                        local_name.as_slice(),
                    )
                {
                    output =
                        rewrite_chart_group_val_element(&output, reader.decoder(), value.as_str())?;
                }
                writer
                    .write_event(Event::Start(output))
                    .map_err(chart_xml_error)?;
                element_stack.push(local_name);
            }
            Ok(Event::Empty(element)) => {
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                let parent_name = element_stack.last().map(Vec::as_slice);
                if let Some(group_index) = current_group_index
                    && parent_name
                        .is_some_and(|parent| chart_type_from_group_name(parent).is_some())
                    && matches!(local_name.as_slice(), b"axId" | b"extLst")
                    && !tail_inserted[group_index]
                {
                    write_missing_chart_group_properties(
                        &mut writer,
                        &chart.groups[group_index],
                        &seen[group_index],
                        false,
                    )?;
                    tail_inserted[group_index] = true;
                }
                if let Some(group_index) = current_group_index
                    && parent_name
                        .is_some_and(|parent| chart_type_from_group_name(parent).is_some())
                    && local_name.as_slice() == b"axId"
                    && chart.groups[group_index].dirty
                {
                    let axis_id_position = axis_id_positions[group_index];
                    let axis_id = chart.groups[group_index]
                        .axis_ids
                        .get(axis_id_position)
                        .ok_or_else(|| {
                            OmError::unsupported(
                                "loaded chart group axis reference count changed before lossless property patch",
                            )
                        })?;
                    axis_id_positions[group_index] += 1;
                    writer
                        .write_event(Event::Empty(rewrite_chart_group_val_element(
                            &element,
                            reader.decoder(),
                            axis_id.as_str(),
                        )?))
                        .map_err(chart_xml_error)?;
                    buffer.clear();
                    continue;
                }
                if let Some(group_index) = current_group_index
                    && parent_name
                        .is_some_and(|parent| chart_type_from_group_name(parent).is_some())
                    && chart_group_direct_property_name(local_name.as_slice())
                {
                    seen[group_index].insert(local_name.clone());
                    if local_name.as_slice() == b"dLbls"
                        && let Some(data_labels) = chart.groups[group_index]
                            .data_labels
                            .as_ref()
                            .filter(|labels| labels.dirty)
                    {
                        writer
                            .get_mut()
                            .write_all(chart_data_labels_xml_string(data_labels).as_bytes())
                            .map_err(chart_xml_error)?;
                        buffer.clear();
                        continue;
                    }
                    if let Some(enabled) =
                        chart_group_line_flag(&chart.groups[group_index], local_name.as_slice())
                    {
                        if enabled {
                            writer
                                .write_event(Event::Empty(element.into_owned()))
                                .map_err(chart_xml_error)?;
                        }
                        buffer.clear();
                        continue;
                    }
                    if let Some(value) = chart_group_direct_property_value(
                        &chart.groups[group_index],
                        local_name.as_slice(),
                    ) {
                        writer
                            .write_event(Event::Empty(rewrite_chart_group_val_element(
                                &element,
                                reader.decoder(),
                                value.as_str(),
                            )?))
                            .map_err(chart_xml_error)?;
                    } else {
                        writer
                            .write_event(Event::Empty(element.into_owned()))
                            .map_err(chart_xml_error)?;
                    }
                    buffer.clear();
                    continue;
                }
                writer
                    .write_event(Event::Empty(element.into_owned()))
                    .map_err(chart_xml_error)?;
            }
            Ok(Event::End(element)) => {
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                if let Some(group_index) = current_group_index
                    && current_group_depth == Some(element_stack.len())
                    && chart_type_from_group_name(local_name.as_slice()).is_some()
                {
                    if chart.groups[group_index].dirty
                        && axis_id_positions[group_index]
                            != chart.groups[group_index].axis_ids.len()
                    {
                        return Err(OmError::unsupported(
                            "loaded chart group axis reference count changed before lossless property patch",
                        ));
                    }
                    if !prefix_inserted[group_index] {
                        write_missing_chart_group_properties(
                            &mut writer,
                            &chart.groups[group_index],
                            &seen[group_index],
                            true,
                        )?;
                    }
                    if !tail_inserted[group_index] {
                        write_missing_chart_group_properties(
                            &mut writer,
                            &chart.groups[group_index],
                            &seen[group_index],
                            false,
                        )?;
                    }
                    current_group_index = None;
                    current_group_depth = None;
                }
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(chart_xml_error)?;
                element_stack.pop();
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer
                .write_event(event.into_owned())
                .map_err(chart_xml_error)?,
            Err(error) => return Err(chart_xml_error(error)),
        }
        buffer.clear();
    }
    if next_group_index != chart.groups.len() {
        return Err(OmError::unsupported(
            "loaded chart group count changed before lossless property patch",
        ));
    }
    Ok(writer.into_inner().into_inner())
}

fn copy_chart_xml_subtree(
    reader: &mut Reader<Cursor<&[u8]>>,
    writer: &mut Writer<Cursor<Vec<u8>>>,
    buffer: &mut Vec<u8>,
    start: BytesStart<'static>,
) -> OmResult<()> {
    writer
        .write_event(Event::Start(start))
        .map_err(chart_xml_error)?;
    let mut depth = 1usize;
    while depth > 0 {
        buffer.clear();
        match reader.read_event_into(buffer) {
            Ok(Event::Start(element)) => {
                depth += 1;
                writer
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(chart_xml_error)?;
            }
            Ok(Event::End(element)) => {
                depth -= 1;
                writer
                    .write_event(Event::End(element.into_owned()))
                    .map_err(chart_xml_error)?;
            }
            Ok(Event::Eof) => {
                return Err(OmError::new(
                    OmErrorCode::Parse,
                    "unexpected EOF while preserving chart group property subtree",
                ));
            }
            Ok(event) => writer
                .write_event(event.into_owned())
                .map_err(chart_xml_error)?,
            Err(error) => return Err(chart_xml_error(error)),
        }
    }
    Ok(())
}

fn patch_loaded_chart_model_xml(
    existing_chart_xml: &[u8],
    chart: &ChartModel,
) -> OmResult<Option<Vec<u8>>> {
    validate_volume_stock_chart(chart)?;
    let rewritten_group_topology_xml;
    let existing_chart_xml = if chart
        .groups
        .iter()
        .any(|group| group.loaded_index.is_none())
    {
        rewritten_group_topology_xml =
            rewrite_loaded_chart_group_additions(existing_chart_xml, chart)?;
        rewritten_group_topology_xml.as_slice()
    } else {
        existing_chart_xml
    };
    let rewritten_chart_xml;
    let existing_chart_xml =
        if chart.series_topology_dirty || chart.series.iter().any(|series| series.filter_dirty) {
            let rewritten = rewrite_loaded_chart_series_topology(existing_chart_xml, chart)?;
            rewritten_chart_xml = rewritten;
            rewritten_chart_xml.as_slice()
        } else {
            existing_chart_xml
        };
    let rewritten_group_removal_xml;
    let existing_chart_xml = if chart.series_topology_dirty && !chart.groups.is_empty() {
        rewritten_group_removal_xml =
            rewrite_loaded_chart_group_removals(existing_chart_xml, chart)?;
        rewritten_group_removal_xml.as_slice()
    } else {
        existing_chart_xml
    };
    let rewritten_axis_topology_xml;
    let existing_chart_xml = if !chart.axes.is_empty() {
        rewritten_axis_topology_xml =
            rewrite_loaded_chart_axis_additions(existing_chart_xml, chart)?;
        rewritten_axis_topology_xml.as_slice()
    } else {
        existing_chart_xml
    };
    let rewritten_group_xml;
    let existing_chart_xml = if chart.groups.iter().any(|group| group.dirty) {
        rewritten_group_xml = patch_loaded_chart_group_properties(existing_chart_xml, chart)?;
        rewritten_group_xml.as_slice()
    } else {
        existing_chart_xml
    };

    let expected_chart_style = chart.style.map(|style| style.to_string());
    let expected_chart_style = expected_chart_style.as_deref();
    let expected_chart_protection = chart.protection_dirty.then_some(chart.protection);
    let expected_chart_protection_needs_xml = expected_chart_protection
        .flatten()
        .is_some_and(|protection| protection != ChartProtectionModel::default());
    let expected_legend_position =
        chart
            .legend
            .as_ref()
            .filter(|legend| legend.visible)
            .map(|legend| {
                chart_legend_position_xml_value(
                    legend.position.unwrap_or(ChartLegendPosition::Right),
                )
            });
    let expected_legend_include_in_layout = chart
        .legend
        .as_ref()
        .filter(|legend| legend.visible)
        .and_then(|legend| legend.include_in_layout)
        .map(|include_in_layout| if include_in_layout { "0" } else { "1" });
    let expected_display_blanks_as = chart
        .display_blanks_as
        .map(chart_display_blanks_as_xml_value);
    let expected_plot_visible_only = chart
        .plot_visible_only
        .map(|value| if value { "1" } else { "0" });
    let expected_show_data_labels_over_maximum = chart
        .show_data_labels_over_maximum
        .map(|value| if value { "1" } else { "0" });
    let expected_dirty_view_3d = chart.view_3d_dirty.then_some(chart.view_3d.as_ref());
    let expected_rounded_corners = chart
        .rounded_corners
        .map(|value| if value { "1" } else { "0" });
    let expected_vary_colors = chart
        .vary_by_categories
        .map(|value| if value { "1" } else { "0" });
    let expected_bar_direction = chart_type_bar_direction_xml_value(&chart.chart_type);
    let expected_chart_grouping = chart_type_grouping_xml_value(&chart.chart_type);
    let expected_bar_shape = chart_effective_bar_shape(chart).map(chart_bar_shape_xml_value);
    let expected_series_invert_if_negative_values = chart
        .series
        .iter()
        .map(|series| {
            series
                .invert_if_negative
                .map(|value| if value { "1" } else { "0" })
        })
        .collect::<Vec<_>>();
    let expected_line_marker = chart_type_line_marker_xml_value(&chart.chart_type);
    let expected_scatter_style = chart_type_scatter_style_xml_value(&chart.chart_type);
    let expected_radar_style = chart_type_radar_style_xml_value(&chart.chart_type);
    let expected_of_pie_type = chart_type_of_pie_xml_value(&chart.chart_type);
    let expected_surface_wireframe = chart_type_surface_wireframe_xml_value(&chart.chart_type);
    let expected_gap_width = chart.gap_width.map(|value| value.to_string());
    let expected_gap_depth = chart
        .gap_depth
        .filter(|_| chart_type_supports_gap_depth(&chart.chart_type))
        .map(|value| value.to_string());
    let expected_overlap = chart.overlap.map(|value| value.to_string());
    let expected_first_slice_angle = chart.first_slice_angle.map(|value| value.to_string());
    let expected_bubble_scale = chart.bubble_scale.map(|value| value.to_string());
    let expected_show_negative_bubbles = chart
        .show_negative_bubbles
        .map(|value| if value { "1" } else { "0" });
    let expected_has_3d_shading = chart
        .has_3d_shading
        .map(|value| if value { "1" } else { "0" });
    let expected_doughnut_hole_size = chart.doughnut_hole_size.map(|value| value.to_string());
    let expected_second_plot_size = chart.second_plot_size.map(|value| value.to_string());
    let expected_size_represents = chart.size_represents.map(chart_size_represents_xml_value);
    let expected_split_type = chart.split_type.map(chart_split_type_xml_value);
    let expected_split_value = chart.split_value.map(chart_number_xml_value);
    let expected_dirty_data_labels = chart
        .data_labels
        .as_ref()
        .filter(|data_labels| data_labels.dirty);
    let expected_data_table = chart.data_table_dirty.then_some(chart.data_table.as_ref());
    let expected_plot_area_layout = chart
        .plot_area_layout_dirty
        .then_some(chart.plot_area_layout.as_ref());
    let expected_dirty_series_data_label_sets = chart
        .series
        .iter()
        .map(|series| {
            series
                .data_labels
                .as_ref()
                .is_some_and(|data_labels| data_labels.dirty)
                || series
                    .point_data_labels
                    .values()
                    .any(|data_labels| data_labels.dirty)
        })
        .collect::<Vec<_>>();
    let expected_gap_width = expected_gap_width.as_deref();
    let expected_gap_depth = expected_gap_depth.as_deref();
    let expected_overlap = expected_overlap.as_deref();
    let expected_first_slice_angle = expected_first_slice_angle.as_deref();
    let expected_bubble_scale = expected_bubble_scale.as_deref();
    let expected_doughnut_hole_size = expected_doughnut_hole_size.as_deref();
    let expected_second_plot_size = expected_second_plot_size.as_deref();
    let expected_split_value = expected_split_value.as_deref();
    let expected_has_hi_lo_lines = chart.has_hi_lo_lines.or_else(|| {
        matches!(
            chart.chart_type,
            ChartType::StockHLC
                | ChartType::StockOHLC
                | ChartType::StockVHLC
                | ChartType::StockVOHLC
        )
        .then_some(true)
    });
    let expected_has_up_down_bars = chart.has_up_down_bars.or_else(|| {
        matches!(
            chart.chart_type,
            ChartType::StockOHLC | ChartType::StockVOHLC
        )
        .then_some(true)
    });
    let expected_chart_group_line_flags: [(&[u8], &str, Option<bool>); 4] = [
        (b"serLines", "c:serLines", chart.has_series_lines),
        (b"dropLines", "c:dropLines", chart.has_drop_lines),
        (b"hiLowLines", "c:hiLowLines", expected_has_hi_lo_lines),
        (b"upDownBars", "c:upDownBars", expected_has_up_down_bars),
    ];
    let expected_chart_group_numeric_settings: [(&[u8], &str, Option<&str>); 9] = [
        (
            b"firstSliceAng",
            "c:firstSliceAng",
            expected_first_slice_angle,
        ),
        (b"bubbleScale", "c:bubbleScale", expected_bubble_scale),
        (
            b"showNegBubbles",
            "c:showNegBubbles",
            expected_show_negative_bubbles,
        ),
        (b"bubble3D", "c:bubble3D", expected_has_3d_shading),
        (b"holeSize", "c:holeSize", expected_doughnut_hole_size),
        (
            b"secondPieSize",
            "c:secondPieSize",
            expected_second_plot_size,
        ),
        (
            b"sizeRepresents",
            "c:sizeRepresents",
            expected_size_represents,
        ),
        (b"splitType", "c:splitType", expected_split_type),
        (b"splitPos", "c:splitPos", expected_split_value),
    ];

    let mut reader = Reader::from_reader(Cursor::new(existing_chart_xml));
    reader.config_mut().trim_text(false);
    let mut writer = Writer::new(Cursor::new(Vec::new()));
    let mut buffer = Vec::new();
    let mut element_stack = Vec::<Vec<u8>>::new();
    let mut next_loaded_series_index = 0usize;
    let mut current_series_index = None::<usize>;
    let mut current_point_index = None::<u32>;
    let mut source_stack = Vec::<ChartSourceXmlSlot>::new();
    let mut current_formula = None::<(ChartSourceXmlSlot, bool)>;
    let mut current_full_reference = None::<(ChartSourceXmlSlot, bool)>;
    let mut dirty_source_extension = None::<(Writer<Cursor<Vec<u8>>>, usize)>;
    let mut source_slots_seen = vec![[false; 4]; chart.series.len()];
    let mut series_order_seen = vec![false; chart.series.len()];
    let mut series_order_written = vec![false; chart.series.len()];
    let mut series_explosion_seen = vec![false; chart.series.len()];
    let mut series_explosion_written = vec![false; chart.series.len()];
    let mut series_explosion_inserted = vec![false; chart.series.len()];
    let mut series_explosion_removed = vec![false; chart.series.len()];
    let mut series_point_explosions_inserted = vec![BTreeSet::<u32>::new(); chart.series.len()];
    let mut series_bar_shape_seen = vec![false; chart.series.len()];
    let mut series_bar_shape_written = vec![false; chart.series.len()];
    let mut series_bar_shape_inserted = vec![false; chart.series.len()];
    let mut series_bar_shape_removed = vec![false; chart.series.len()];
    let mut series_smooth_seen = vec![false; chart.series.len()];
    let mut series_smooth_written = vec![false; chart.series.len()];
    let mut series_smooth_inserted = vec![false; chart.series.len()];
    let mut series_smooth_removed = vec![false; chart.series.len()];
    let mut series_marker_seen = vec![false; chart.series.len()];
    let mut series_marker_inserted = vec![false; chart.series.len()];
    let mut series_marker_removed = vec![false; chart.series.len()];
    let mut series_marker_style_seen = vec![false; chart.series.len()];
    let mut series_marker_style_written = vec![false; chart.series.len()];
    let mut series_marker_style_inserted = vec![false; chart.series.len()];
    let mut series_marker_size_seen = vec![false; chart.series.len()];
    let mut series_marker_size_written = vec![false; chart.series.len()];
    let mut series_marker_size_inserted = vec![false; chart.series.len()];
    let mut series_invert_if_negative_seen = vec![false; chart.series.len()];
    let mut series_invert_if_negative_written = vec![false; chart.series.len()];
    let mut series_invert_if_negative_inserted = vec![false; chart.series.len()];
    let mut series_emitted = vec![false; chart.series.len()];
    let mut patched_sources = 0usize;
    let mut chart_type = None::<ChartType>;
    let mut chart_type_rewritten = false;
    let mut chart_style_seen = false;
    let mut chart_style_written = false;
    let mut chart_style_inserted = false;
    let mut chart_protection_seen = false;
    let mut chart_protection_written = false;
    let mut chart_protection_inserted = false;
    let mut chart_protection_removed = false;
    let mut chart_title_seen = false;
    let mut chart_title_removed = false;
    let mut chart_title_inserted = false;
    let mut chart_title_text_written = false;
    let mut data_table_seen = false;
    let mut data_table_written = false;
    let mut data_table_inserted = false;
    let mut data_table_removed = false;
    let mut plot_area_layout_container_seen = false;
    let mut plot_area_manual_layout_seen = false;
    let mut plot_area_manual_layout_written = false;
    let mut plot_area_manual_layout_inserted = false;
    let mut plot_area_manual_layout_removed = false;
    let mut current_text_target = None::<(ChartTextXmlTarget, bool)>;
    let mut title_stack = Vec::<(usize, ChartTextXmlTarget)>::new();
    let mut legend_seen = false;
    let mut legend_removed = false;
    let mut legend_inserted = false;
    let mut legend_position_written = false;
    let mut legend_overlay_seen = false;
    let mut legend_overlay_written = false;
    let mut legend_overlay_inserted = false;
    let mut display_blanks_as_seen = false;
    let mut display_blanks_as_written = false;
    let mut display_blanks_as_inserted = false;
    let mut plot_visible_only_seen = false;
    let mut plot_visible_only_written = false;
    let mut plot_visible_only_inserted = false;
    let mut show_data_labels_over_maximum_seen = false;
    let mut show_data_labels_over_maximum_written = false;
    let mut show_data_labels_over_maximum_inserted = false;
    let mut view_3d_seen = false;
    let mut view_3d_written = false;
    let mut view_3d_inserted = false;
    let mut rounded_corners_seen = false;
    let mut rounded_corners_written = false;
    let mut rounded_corners_inserted = false;
    let mut vary_colors_seen = false;
    let mut vary_colors_written = false;
    let mut vary_colors_inserted = false;
    let mut bar_direction_seen = false;
    let mut bar_direction_written = false;
    let mut bar_direction_inserted = false;
    let mut chart_grouping_seen = false;
    let mut chart_grouping_written = false;
    let mut chart_grouping_inserted = false;
    let mut bar_shape_seen = false;
    let mut bar_shape_written = false;
    let mut bar_shape_inserted = false;
    let mut line_marker_seen = false;
    let mut line_marker_written = false;
    let mut line_marker_inserted = false;
    let mut scatter_style_seen = false;
    let mut scatter_style_written = false;
    let mut scatter_style_inserted = false;
    let mut radar_style_seen = false;
    let mut radar_style_written = false;
    let mut radar_style_inserted = false;
    let mut of_pie_type_seen = false;
    let mut of_pie_type_written = false;
    let mut of_pie_type_inserted = false;
    let mut surface_wireframe_seen = false;
    let mut surface_wireframe_written = false;
    let mut surface_wireframe_inserted = false;
    let mut gap_width_seen = false;
    let mut gap_width_written = false;
    let mut gap_width_inserted = false;
    let mut gap_depth_seen = false;
    let mut gap_depth_written = false;
    let mut gap_depth_inserted = false;
    let mut overlap_seen = false;
    let mut overlap_written = false;
    let mut overlap_inserted = false;
    let mut data_labels_seen = false;
    let mut data_labels_written = false;
    let mut data_labels_inserted = false;
    let mut series_data_labels_seen = vec![false; chart.series.len()];
    let mut series_data_labels_written = vec![false; chart.series.len()];
    let mut series_data_labels_inserted = vec![false; chart.series.len()];
    let mut chart_group_line_flag_seen = [false; 4];
    let mut chart_group_line_flag_written = [false; 4];
    let mut chart_group_line_flag_inserted = [false; 4];
    let mut chart_group_line_flag_removed = [false; 4];
    let mut chart_group_numeric_setting_seen = [false; 9];
    let mut chart_group_numeric_setting_written = [false; 9];
    let mut chart_group_numeric_setting_inserted = [false; 9];
    let mut current_axis_index = None::<usize>;
    let mut current_axis_depth = None::<usize>;
    let mut axis_kinds = Vec::<ChartAxisKind>::new();
    let mut axis_title_texts = Vec::<Option<String>>::new();
    let mut axis_title_text_written = Vec::<bool>::new();
    let mut axis_major_gridlines_seen = Vec::<bool>::new();
    let mut axis_major_gridlines_written = Vec::<bool>::new();
    let mut axis_major_gridlines_inserted = Vec::<bool>::new();
    let mut axis_major_gridlines_removed = Vec::<bool>::new();
    let mut axis_minor_gridlines_seen = Vec::<bool>::new();
    let mut axis_minor_gridlines_written = Vec::<bool>::new();
    let mut axis_minor_gridlines_inserted = Vec::<bool>::new();
    let mut axis_minor_gridlines_removed = Vec::<bool>::new();
    let mut axis_major_tick_mark_seen = Vec::<bool>::new();
    let mut axis_major_tick_mark_written = Vec::<bool>::new();
    let mut axis_major_tick_mark_inserted = Vec::<bool>::new();
    let mut axis_major_tick_mark_removed = Vec::<bool>::new();
    let mut axis_minor_tick_mark_seen = Vec::<bool>::new();
    let mut axis_minor_tick_mark_written = Vec::<bool>::new();
    let mut axis_minor_tick_mark_inserted = Vec::<bool>::new();
    let mut axis_minor_tick_mark_removed = Vec::<bool>::new();
    let mut axis_tick_label_position_seen = Vec::<bool>::new();
    let mut axis_tick_label_position_written = Vec::<bool>::new();
    let mut axis_tick_label_position_inserted = Vec::<bool>::new();
    let mut axis_tick_label_position_removed = Vec::<bool>::new();
    let mut axis_tick_label_number_format_seen = Vec::<bool>::new();
    let mut axis_tick_label_number_format_written = Vec::<bool>::new();
    let mut axis_tick_label_number_format_inserted = Vec::<bool>::new();
    let mut axis_tick_label_number_format_removed = Vec::<bool>::new();
    let mut axis_tick_label_spacing_seen = Vec::<bool>::new();
    let mut axis_tick_label_spacing_written = Vec::<bool>::new();
    let mut axis_tick_label_spacing_inserted = Vec::<bool>::new();
    let mut axis_tick_label_spacing_removed = Vec::<bool>::new();
    let mut axis_tick_mark_spacing_seen = Vec::<bool>::new();
    let mut axis_tick_mark_spacing_written = Vec::<bool>::new();
    let mut axis_tick_mark_spacing_inserted = Vec::<bool>::new();
    let mut axis_tick_mark_spacing_removed = Vec::<bool>::new();
    let mut axis_scaling_seen = Vec::<bool>::new();
    let mut axis_log_base_seen = Vec::<bool>::new();
    let mut axis_log_base_written = Vec::<bool>::new();
    let mut axis_log_base_inserted = Vec::<bool>::new();
    let mut axis_log_base_removed = Vec::<bool>::new();
    let mut axis_orientation_seen = Vec::<bool>::new();
    let mut axis_orientation_written = Vec::<bool>::new();
    let mut axis_orientation_inserted = Vec::<bool>::new();
    let mut axis_orientation_removed = Vec::<bool>::new();
    let mut axis_minimum_scale_seen = Vec::<bool>::new();
    let mut axis_minimum_scale_written = Vec::<bool>::new();
    let mut axis_minimum_scale_inserted = Vec::<bool>::new();
    let mut axis_minimum_scale_removed = Vec::<bool>::new();
    let mut axis_maximum_scale_seen = Vec::<bool>::new();
    let mut axis_maximum_scale_written = Vec::<bool>::new();
    let mut axis_maximum_scale_inserted = Vec::<bool>::new();
    let mut axis_maximum_scale_removed = Vec::<bool>::new();
    let mut axis_major_unit_seen = Vec::<bool>::new();
    let mut axis_major_unit_written = Vec::<bool>::new();
    let mut axis_major_unit_inserted = Vec::<bool>::new();
    let mut axis_major_unit_removed = Vec::<bool>::new();
    let mut axis_minor_unit_seen = Vec::<bool>::new();
    let mut axis_minor_unit_written = Vec::<bool>::new();
    let mut axis_minor_unit_inserted = Vec::<bool>::new();
    let mut axis_minor_unit_removed = Vec::<bool>::new();
    let mut axis_display_units_seen = Vec::<bool>::new();
    let mut axis_display_units_written = Vec::<bool>::new();
    let mut axis_display_units_inserted = Vec::<bool>::new();
    let mut axis_display_units_removed = Vec::<bool>::new();
    let mut axis_crosses_seen = Vec::<bool>::new();
    let mut axis_crosses_written = Vec::<bool>::new();
    let mut axis_crosses_inserted = Vec::<bool>::new();
    let mut axis_crosses_removed = Vec::<bool>::new();
    let mut axis_crosses_at_seen = Vec::<bool>::new();
    let mut axis_crosses_at_written = Vec::<bool>::new();
    let mut axis_crosses_at_inserted = Vec::<bool>::new();
    let mut axis_crosses_at_removed = Vec::<bool>::new();
    let mut axis_cross_between_seen = Vec::<bool>::new();
    let mut axis_cross_between_written = Vec::<bool>::new();
    let mut axis_cross_between_inserted = Vec::<bool>::new();
    let mut axis_cross_between_removed = Vec::<bool>::new();
    let mut axis_category_type_auto_seen = Vec::<bool>::new();
    let mut axis_category_type_auto_written = Vec::<bool>::new();
    let mut axis_category_type_auto_inserted = Vec::<bool>::new();
    let mut axis_category_type_auto_removed = Vec::<bool>::new();
    let mut axis_base_unit_seen = Vec::<bool>::new();
    let mut axis_base_unit_written = Vec::<bool>::new();
    let mut axis_base_unit_inserted = Vec::<bool>::new();
    let mut axis_base_unit_removed = Vec::<bool>::new();
    let mut axis_major_time_unit_seen = Vec::<bool>::new();
    let mut axis_major_time_unit_written = Vec::<bool>::new();
    let mut axis_major_time_unit_inserted = Vec::<bool>::new();
    let mut axis_major_time_unit_removed = Vec::<bool>::new();
    let mut axis_minor_time_unit_seen = Vec::<bool>::new();
    let mut axis_minor_time_unit_written = Vec::<bool>::new();
    let mut axis_minor_time_unit_inserted = Vec::<bool>::new();
    let mut axis_minor_time_unit_removed = Vec::<bool>::new();
    let mut current_chart_group_depth = None::<usize>;
    let mut chart_group_axis_refs_seen = Vec::<String>::new();
    let mut current_series_marker_index = None::<usize>;
    let mut skip_depth = 0usize;

    let slot_index = |slot: ChartSourceXmlSlot| -> usize {
        match slot {
            ChartSourceXmlSlot::Name => 0,
            ChartSourceXmlSlot::XValues => 1,
            ChartSourceXmlSlot::Values => 2,
            ChartSourceXmlSlot::BubbleSize => 3,
        }
    };
    let source_container_slot = |local_name: &[u8]| -> Option<ChartSourceXmlSlot> {
        match local_name {
            b"tx" => Some(ChartSourceXmlSlot::Name),
            b"cat" | b"xVal" => Some(ChartSourceXmlSlot::XValues),
            b"val" | b"yVal" => Some(ChartSourceXmlSlot::Values),
            b"bubbleSize" => Some(ChartSourceXmlSlot::BubbleSize),
            _ => None,
        }
    };
    let decode_general_ref_text = |reference: &BytesRef<'_>| -> OmResult<String> {
        let reference = reference.decode().map_err(chart_xml_error)?;
        let text_value = if let Some(number) = reference.strip_prefix("#x") {
            let codepoint = u32::from_str_radix(number, 16).map_err(chart_xml_error)?;
            char::from_u32(codepoint)
                .ok_or_else(|| {
                    OmError::new(
                        OmErrorCode::Parse,
                        format!("invalid XML character reference: &{reference};"),
                    )
                })?
                .to_string()
        } else if let Some(number) = reference.strip_prefix("#X") {
            let codepoint = u32::from_str_radix(number, 16).map_err(chart_xml_error)?;
            char::from_u32(codepoint)
                .ok_or_else(|| {
                    OmError::new(
                        OmErrorCode::Parse,
                        format!("invalid XML character reference: &{reference};"),
                    )
                })?
                .to_string()
        } else if let Some(number) = reference.strip_prefix('#') {
            let codepoint = number.parse::<u32>().map_err(chart_xml_error)?;
            char::from_u32(codepoint)
                .ok_or_else(|| {
                    OmError::new(
                        OmErrorCode::Parse,
                        format!("invalid XML character reference: &{reference};"),
                    )
                })?
                .to_string()
        } else {
            match reference.as_ref() {
                "amp" => "&".to_string(),
                "lt" => "<".to_string(),
                "gt" => ">".to_string(),
                "quot" => "\"".to_string(),
                "apos" => "'".to_string(),
                _ => format!("&{reference};"),
            }
        };
        Ok(text_value)
    };
    struct LoadedSeriesSignature {
        raw_index: Option<u32>,
        sources: [Option<String>; 4],
        group_index: usize,
    }
    let mut loaded_series_signatures = Vec::<LoadedSeriesSignature>::new();
    let mut loaded_chart_group_names = Vec::<Vec<u8>>::new();
    let mut loaded_axis_signatures = Vec::<(ChartAxisKind, Option<String>)>::new();
    let mut loaded_axis_start_positions = Vec::<usize>::new();
    {
        let mut signature_reader = NsReader::from_reader(Cursor::new(existing_chart_xml));
        signature_reader.config_mut().trim_text(false);
        let mut signature_buffer = Vec::new();
        let mut active_signature = None::<LoadedSeriesSignature>;
        let mut active_signature_depth = 0usize;
        let mut signature_source_stack = Vec::<ChartSourceXmlSlot>::new();
        let mut signature_formula = None::<(ChartSourceXmlSlot, String, usize)>;
        let mut signature_element_stack = Vec::<Vec<u8>>::new();
        let mut signature_chart_namespace_stack = Vec::<bool>::new();
        let mut active_group_index = None::<usize>;
        let mut active_axis_index = None::<usize>;
        let mut active_axis_depth = None::<usize>;
        let is_chart_namespace = |namespace: ResolveResult<'_>| match namespace {
            ResolveResult::Bound(namespace) => namespace.as_ref() == CHART_XML_NAMESPACE,
            ResolveResult::Unbound | ResolveResult::Unknown(_) => false,
        };

        loop {
            let event_start =
                usize::try_from(signature_reader.buffer_position()).map_err(|_| {
                    OmError::new(OmErrorCode::InvalidState, "chart XML position overflow")
                })?;
            match signature_reader.read_resolved_event_into(&mut signature_buffer) {
                Ok((namespace, Event::Start(element))) => {
                    let is_chart_namespace = is_chart_namespace(namespace);
                    let local_name = xml_local_name(element.name().as_ref()).to_vec();
                    let parent_name = signature_element_stack.last().map(Vec::as_slice);
                    if is_chart_namespace
                        && signature_chart_namespace_stack.last() == Some(&true)
                        && parent_name == Some(b"plotArea".as_slice())
                        && chart_type_from_group_name(local_name.as_slice()).is_some()
                    {
                        active_group_index = Some(loaded_chart_group_names.len());
                        loaded_chart_group_names.push(local_name.clone());
                    }
                    if is_chart_namespace
                        && signature_chart_namespace_stack.last() == Some(&true)
                        && parent_name == Some(b"plotArea".as_slice())
                        && let Some(axis_kind) =
                            chart_axis_kind_from_xml_name(local_name.as_slice())
                    {
                        active_axis_index = Some(loaded_axis_signatures.len());
                        active_axis_depth = Some(signature_element_stack.len() + 1);
                        loaded_axis_signatures.push((axis_kind, None));
                        loaded_axis_start_positions.push(event_start);
                    }
                    if local_name.as_slice() == b"axId"
                        && is_chart_namespace
                        && active_axis_depth == Some(signature_element_stack.len())
                        && let Some(axis_index) = active_axis_index
                    {
                        for attr in element.attributes() {
                            let attr = attr.map_err(chart_xml_error)?;
                            if attr.key.as_ref() == b"val" {
                                loaded_axis_signatures[axis_index].1 = Some(
                                    attr.decode_and_unescape_value(signature_reader.decoder())
                                        .map_err(chart_xml_error)?
                                        .into_owned(),
                                );
                                break;
                            }
                        }
                    }
                    if local_name.as_slice() == b"ser" && active_signature.is_none() {
                        active_signature = Some(LoadedSeriesSignature {
                            raw_index: None,
                            sources: [None, None, None, None],
                            group_index: active_group_index.unwrap_or(0),
                        });
                        active_signature_depth = 1;
                    } else if active_signature.is_some() {
                        if active_signature_depth == 1
                            && local_name.as_slice() == b"idx"
                            && let Some(signature) = active_signature.as_mut()
                        {
                            for attr in element.attributes() {
                                let attr = attr.map_err(chart_xml_error)?;
                                if attr.key.as_ref() == b"val" {
                                    signature.raw_index = attr
                                        .decode_and_unescape_value(signature_reader.decoder())
                                        .map_err(chart_xml_error)?
                                        .parse::<u32>()
                                        .ok();
                                    break;
                                }
                            }
                        }
                        active_signature_depth += 1;
                        if let Some((_, _, depth)) = signature_formula.as_mut() {
                            *depth += 1;
                        } else if let Some(slot) = source_container_slot(local_name.as_slice()) {
                            signature_source_stack.push(slot);
                        } else if local_name.as_slice() == b"f"
                            && let Some(slot) = signature_source_stack.last().copied()
                        {
                            signature_formula = Some((slot, String::new(), 1));
                        }
                    }
                    signature_element_stack.push(local_name);
                    signature_chart_namespace_stack.push(is_chart_namespace);
                }
                Ok((namespace, Event::Empty(element))) => {
                    let is_chart_namespace = is_chart_namespace(namespace);
                    let local_name = xml_local_name(element.name().as_ref()).to_vec();
                    if is_chart_namespace
                        && signature_chart_namespace_stack.last() == Some(&true)
                        && signature_element_stack.last().map(Vec::as_slice)
                            == Some(b"plotArea".as_slice())
                        && chart_type_from_group_name(local_name.as_slice()).is_some()
                    {
                        loaded_chart_group_names.push(local_name.clone());
                    }
                    if local_name.as_slice() == b"axId"
                        && is_chart_namespace
                        && active_axis_depth == Some(signature_element_stack.len())
                        && let Some(axis_index) = active_axis_index
                    {
                        for attr in element.attributes() {
                            let attr = attr.map_err(chart_xml_error)?;
                            if attr.key.as_ref() == b"val" {
                                loaded_axis_signatures[axis_index].1 = Some(
                                    attr.decode_and_unescape_value(signature_reader.decoder())
                                        .map_err(chart_xml_error)?
                                        .into_owned(),
                                );
                                break;
                            }
                        }
                    }
                    if active_signature_depth == 1
                        && local_name.as_slice() == b"idx"
                        && let Some(signature) = active_signature.as_mut()
                    {
                        for attr in element.attributes() {
                            let attr = attr.map_err(chart_xml_error)?;
                            if attr.key.as_ref() == b"val" {
                                signature.raw_index = attr
                                    .decode_and_unescape_value(signature_reader.decoder())
                                    .map_err(chart_xml_error)?
                                    .parse::<u32>()
                                    .ok();
                                break;
                            }
                        }
                    }
                }
                Ok((_, Event::Text(text))) => {
                    if let Some((_, formula, _)) = signature_formula.as_mut() {
                        formula.push_str(&text.xml_content().map_err(chart_xml_error)?);
                    }
                }
                Ok((_, Event::CData(data))) => {
                    if let Some((_, formula, _)) = signature_formula.as_mut() {
                        formula.push_str(&data.xml_content().map_err(chart_xml_error)?);
                    }
                }
                Ok((_, Event::GeneralRef(reference))) => {
                    if let Some((_, formula, _)) = signature_formula.as_mut() {
                        formula.push_str(&decode_general_ref_text(&reference)?);
                    }
                }
                Ok((namespace, Event::End(element))) => {
                    let is_chart_namespace = is_chart_namespace(namespace);
                    let local_name = xml_local_name(element.name().as_ref()).to_vec();
                    let closes_series =
                        local_name.as_slice() == b"ser" && active_signature_depth == 1;
                    if let Some((slot, formula, depth)) = signature_formula.as_mut() {
                        if local_name.as_slice() == b"f" && *depth == 1 {
                            if !formula.is_empty()
                                && let Some(signature) = active_signature.as_mut()
                            {
                                signature.sources[slot_index(*slot)]
                                    .get_or_insert_with(|| formula.clone());
                            }
                            signature_formula = None;
                        } else {
                            *depth = depth.saturating_sub(1);
                        }
                    } else if matches!(
                        local_name.as_slice(),
                        b"tx" | b"cat" | b"val" | b"xVal" | b"yVal" | b"bubbleSize"
                    ) {
                        signature_source_stack.pop();
                    }
                    if closes_series {
                        if let Some(signature) = active_signature.take() {
                            loaded_series_signatures.push(signature);
                            signature_source_stack.clear();
                        }
                        active_signature_depth = 0;
                    } else if active_signature_depth > 0 {
                        active_signature_depth -= 1;
                    }
                    if is_chart_namespace
                        && active_axis_depth == Some(signature_element_stack.len())
                        && signature_element_stack.len() >= 2
                        && signature_element_stack[signature_element_stack.len() - 2].as_slice()
                            == b"plotArea"
                        && signature_chart_namespace_stack
                            [signature_chart_namespace_stack.len() - 2]
                        && chart_axis_kind_from_xml_name(local_name.as_slice()).is_some()
                    {
                        active_axis_index = None;
                        active_axis_depth = None;
                    }
                    if active_group_index.is_some()
                        && chart_type_from_group_name(local_name.as_slice()).is_some()
                        && signature_element_stack
                            .get(signature_element_stack.len().saturating_sub(2))
                            .is_some_and(|name| name.as_slice() == b"plotArea")
                    {
                        active_group_index = None;
                    }
                    signature_element_stack.pop();
                    signature_chart_namespace_stack.pop();
                }
                Ok((_, Event::Eof)) => break,
                Ok((_, _)) => {}
                Err(error) => return Err(chart_xml_error(error)),
            }
            signature_buffer.clear();
        }
    }
    let model_series_matches_signature = |series: &SeriesModel,
                                          signature: &LoadedSeriesSignature|
     -> bool {
        if series
            .raw_index
            .zip(signature.raw_index)
            .is_some_and(|(model, loaded)| model == loaded)
        {
            return true;
        }
        let sources = [
            series.name.as_ref(),
            series.x_values.as_ref(),
            series.values.as_ref(),
            series.bubble_size.as_ref(),
        ];
        sources
            .iter()
            .zip(signature.sources.iter())
            .all(|(source, formula)| match (source, formula) {
                (Some(source), Some(formula)) => source.raw.text.trim_start_matches('=') == formula,
                (None, None) => true,
                _ => false,
            })
    };
    let mut used_model_series = vec![false; chart.series.len()];
    let loaded_series_model_indices = loaded_series_signatures
        .iter()
        .enumerate()
        .map(|(loaded_index, signature)| {
            if let Some(model_index) =
                chart
                    .series
                    .iter()
                    .enumerate()
                    .find_map(|(model_index, series)| {
                        (!used_model_series[model_index]
                            && model_series_matches_signature(series, signature))
                        .then_some(model_index)
                    })
            {
                used_model_series[model_index] = true;
                Some(model_index)
            } else if loaded_series_signatures.len() == chart.series.len()
                && loaded_index < chart.series.len()
                && !used_model_series[loaded_index]
            {
                used_model_series[loaded_index] = true;
                Some(loaded_index)
            } else {
                None
            }
        })
        .collect::<Vec<_>>();

    let loaded_chart_group_count = loaded_chart_group_names.len();
    let preserve_loaded_group_types = loaded_chart_group_count > 1;
    if !preserve_loaded_group_types
        && chart
            .series
            .iter()
            .any(|series| series.axis_group == ChartAxisGroup::Secondary)
    {
        return Ok(None);
    }
    let mut model_series_group_indices = vec![None; chart.series.len()];
    for (loaded_index, model_index) in loaded_series_model_indices.iter().enumerate() {
        if let Some(model_index) = model_index {
            model_series_group_indices[*model_index] = loaded_series_signatures
                .get(loaded_index)
                .map(|signature| signature.group_index);
        }
    }
    if preserve_loaded_group_types {
        let group_shape_matches = if chart_type_is_volume_stock(&chart.chart_type) {
            loaded_chart_group_names.as_slice()
                == [b"barChart".as_slice(), b"stockChart".as_slice()]
        } else {
            chart_group_xml_name(&chart.chart_type).is_some_and(|target_name| {
                loaded_chart_group_names
                    .first()
                    .is_some_and(|loaded_name| loaded_name.as_slice() == target_name.as_bytes())
            })
        };
        if !group_shape_matches {
            return Err(OmError::unsupported(
                "loaded multi-group chart type reshaping is not supported losslessly",
            ));
        }
        let loaded_raw_indices = loaded_series_signatures
            .iter()
            .filter_map(|signature| signature.raw_index)
            .collect::<BTreeSet<_>>();
        let model_raw_indices = chart
            .series
            .iter()
            .filter_map(|series| series.raw_index)
            .collect::<BTreeSet<_>>();
        let stable_series_topology = loaded_raw_indices.len() == loaded_series_signatures.len()
            && model_raw_indices.len() == chart.series.len()
            && loaded_raw_indices == model_raw_indices
            && loaded_series_model_indices.iter().all(Option::is_some)
            && model_series_group_indices.iter().all(Option::is_some);
        if !stable_series_topology {
            return Err(OmError::unsupported(
                "loaded multi-group chart series topology changes are not supported losslessly",
            ));
        }
        let axis_topology_matches = loaded_axis_signatures.len() == chart.axes.len()
            && loaded_axis_signatures.iter().zip(chart.axes.iter()).all(
                |((loaded_kind, loaded_id), axis)| {
                    *loaded_kind == axis.kind && loaded_id.as_ref() == axis.raw_id.as_ref()
                },
            );
        if !axis_topology_matches {
            return Err(OmError::unsupported(
                "loaded multi-group chart axis topology changes are not supported losslessly",
            ));
        }
        let axis_groups_match = chart
            .series
            .iter()
            .enumerate()
            .all(|(series_index, series)| {
                model_series_group_indices[series_index].is_some_and(|group_index| {
                    if chart.groups.is_empty() {
                        series.axis_group
                            == if group_index == 0 {
                                ChartAxisGroup::Primary
                            } else {
                                ChartAxisGroup::Secondary
                            }
                    } else {
                        chart
                            .groups
                            .get(group_index)
                            .is_some_and(|group| series.axis_group == group.axis_group)
                    }
                })
            });
        if !axis_groups_match {
            return Err(OmError::unsupported(
                "moving loaded series between chart groups is not supported losslessly",
            ));
        }
    }
    let model_series_chart_types = model_series_group_indices
        .iter()
        .map(|group_index| {
            if preserve_loaded_group_types {
                group_index
                    .and_then(|group_index| loaded_chart_group_names.get(group_index))
                    .and_then(|group_name| chart_type_from_group_name(group_name.as_slice()))
                    .unwrap_or_else(|| chart.chart_type.clone())
            } else {
                chart.chart_type.clone()
            }
        })
        .collect::<Vec<_>>();
    let expected_series_explosions = model_series_chart_types
        .iter()
        .enumerate()
        .map(|(series_index, chart_type)| {
            if !chart_type_supports_explosion(chart_type) {
                return None;
            }
            model_series_group_indices[series_index]
                .and_then(|group_index| chart.groups.get(group_index))
                .and_then(|group| group.explosion)
                .or(chart.explosion)
                .or_else(|| {
                    matches!(
                        chart_type,
                        ChartType::PieExploded
                            | ChartType::Pie3DExploded
                            | ChartType::DoughnutExploded
                    )
                    .then_some(25)
                })
                .map(|value| value.to_string())
        })
        .collect::<Vec<_>>();

    let expected_dirty_sources = chart
        .series
        .iter()
        .enumerate()
        .map(|(series_index, series)| {
            usize::from(series.name.as_ref().is_some_and(|source| source.dirty))
                + usize::from(series.x_values.as_ref().is_some_and(|source| source.dirty))
                + usize::from(series.values.as_ref().is_some_and(|source| source.dirty))
                + usize::from(
                    chart_type_uses_bubble_size(&model_series_chart_types[series_index])
                        && series
                            .bubble_size
                            .as_ref()
                            .is_some_and(|source| source.dirty),
                )
        })
        .sum::<usize>();

    let expected_dirty_point_explosions = chart
        .series
        .iter()
        .enumerate()
        .map(|(series_index, series)| {
            if chart_type_supports_explosion(&model_series_chart_types[series_index]) {
                series
                    .points
                    .iter()
                    .filter_map(|(point_index, point)| {
                        point
                            .dirty
                            .then_some(point.explosion)
                            .flatten()
                            .map(|explosion| (*point_index, explosion))
                    })
                    .collect::<Vec<_>>()
            } else {
                Vec::new()
            }
        })
        .collect::<Vec<_>>();
    let expected_series_bar_shapes = chart
        .series
        .iter()
        .enumerate()
        .map(|(series_index, series)| {
            chart_type_supports_bar_shape(&model_series_chart_types[series_index])
                .then_some(series.bar_shape)
                .flatten()
                .map(chart_bar_shape_xml_value)
        })
        .collect::<Vec<_>>();
    let expected_series_smooth_values = chart
        .series
        .iter()
        .enumerate()
        .map(|(series_index, series)| {
            chart_type_supports_series_smooth(&model_series_chart_types[series_index])
                .then_some(series.smooth)
                .flatten()
                .map(|value| if value { "1" } else { "0" })
        })
        .collect::<Vec<_>>();
    let expected_series_marker_styles = chart
        .series
        .iter()
        .enumerate()
        .map(|(series_index, series)| {
            chart_type_supports_series_marker(&model_series_chart_types[series_index])
                .then_some(series.marker_style)
                .flatten()
                .map(chart_marker_style_xml_value)
        })
        .collect::<Vec<_>>();
    let expected_series_marker_sizes = chart
        .series
        .iter()
        .enumerate()
        .map(|(series_index, series)| {
            chart_type_supports_series_marker(&model_series_chart_types[series_index])
                .then_some(series.marker_size)
                .flatten()
                .map(|value| value.to_string())
        })
        .collect::<Vec<_>>();

    let source_for_slot =
        |series_index: usize, slot: ChartSourceXmlSlot| -> Option<&ChartSourceExpr> {
            chart
                .series
                .get(series_index)
                .and_then(|series| match slot {
                    ChartSourceXmlSlot::Name => series.name.as_ref(),
                    ChartSourceXmlSlot::XValues => series.x_values.as_ref(),
                    ChartSourceXmlSlot::Values => series.values.as_ref(),
                    ChartSourceXmlSlot::BubbleSize
                        if chart_type_uses_bubble_size(&model_series_chart_types[series_index]) =>
                    {
                        series.bubble_size.as_ref()
                    }
                    ChartSourceXmlSlot::BubbleSize => None,
                })
        };
    let expected_point_explosion = |series_index: usize, point_index: u32| -> Option<u16> {
        expected_dirty_point_explosions
            .get(series_index)
            .and_then(|points| {
                points
                    .iter()
                    .find_map(|(index, explosion)| (*index == point_index).then_some(*explosion))
            })
    };
    let target_chart_group_name = chart_group_xml_name(&chart.chart_type);
    let qualified_replacement_name = |qualified_name: &[u8], replacement_local: &str| -> String {
        if let Some(prefix_len) = qualified_name.iter().position(|byte| *byte == b':') {
            format!(
                "{}:{replacement_local}",
                String::from_utf8_lossy(&qualified_name[..prefix_len])
            )
        } else {
            replacement_local.to_string()
        }
    };
    let source_slots_with_bubble_size = [
        ChartSourceXmlSlot::Name,
        ChartSourceXmlSlot::XValues,
        ChartSourceXmlSlot::Values,
        ChartSourceXmlSlot::BubbleSize,
    ];
    let source_slots_in_order: &[ChartSourceXmlSlot] = &source_slots_with_bubble_size;
    let source_container_target_local_name = |series_index: usize,
                                              slot: ChartSourceXmlSlot|
     -> &'static str {
        match (&model_series_chart_types[series_index], slot) {
            (_, ChartSourceXmlSlot::Name) => "tx",
            (chart_type, ChartSourceXmlSlot::XValues) if chart_type_uses_xy_values(chart_type) => {
                "xVal"
            }
            (chart_type, ChartSourceXmlSlot::Values) if chart_type_uses_xy_values(chart_type) => {
                "yVal"
            }
            (_, ChartSourceXmlSlot::XValues) => "cat",
            (_, ChartSourceXmlSlot::Values) => "val",
            (_, ChartSourceXmlSlot::BubbleSize) => "bubbleSize",
        }
    };
    let write_chart_source_container = |writer: &mut Writer<Cursor<Vec<u8>>>,
                                        series_index: usize,
                                        slot: ChartSourceXmlSlot,
                                        source: &ChartSourceExpr|
     -> OmResult<()> {
        let xml = chart_source_container_xml_string(
            &model_series_chart_types[series_index],
            slot,
            source,
        )?;
        writer
            .get_mut()
            .write_all(xml.as_bytes())
            .map_err(chart_xml_error)?;
        Ok(())
    };
    let rewrite_val_attribute_element = |element: &BytesStart<'_>,
                                         decoder: quick_xml::encoding::Decoder,
                                         replacement: &str|
     -> OmResult<BytesStart<'static>> {
        let qualified_name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
        let mut rewritten = BytesStart::new(qualified_name);
        let mut wrote_value = false;
        for attr in element.attributes() {
            let attr = attr.map_err(chart_xml_error)?;
            let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
            let value = attr
                .decode_and_unescape_value(decoder)
                .map_err(chart_xml_error)?
                .into_owned();
            if attr.key.as_ref() == b"val" {
                rewritten.push_attribute((key.as_str(), replacement));
                wrote_value = true;
            } else {
                rewritten.push_attribute((key.as_str(), value.as_str()));
            }
        }
        if !wrote_value {
            rewritten.push_attribute(("val", replacement));
        }
        Ok(rewritten)
    };
    let rewrite_element_name = |element: &BytesStart<'_>,
                                decoder: quick_xml::encoding::Decoder,
                                replacement_local: &str|
     -> OmResult<BytesStart<'static>> {
        let qualified_name = qualified_replacement_name(element.name().as_ref(), replacement_local);
        let mut rewritten = BytesStart::new(qualified_name);
        for attr in element.attributes() {
            let attr = attr.map_err(chart_xml_error)?;
            let key = String::from_utf8_lossy(attr.key.as_ref()).into_owned();
            let value = attr
                .decode_and_unescape_value(decoder)
                .map_err(chart_xml_error)?
                .into_owned();
            rewritten.push_attribute((key.as_str(), value.as_str()));
        }
        Ok(rewritten)
    };
    let write_chart_text_element =
        |writer: &mut Writer<Cursor<Vec<u8>>>, root_name: &str, text: &str| -> OmResult<()> {
            writer
                .write_event(Event::Start(BytesStart::new(root_name)))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::Start(BytesStart::new("c:tx")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::Start(BytesStart::new("c:rich")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::Start(BytesStart::new("a:p")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::Start(BytesStart::new("a:r")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::Start(BytesStart::new("a:t")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::Text(BytesText::from_escaped(partial_escape(text))))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:t")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:r")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::End(BytesEnd::new("a:p")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:rich")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:tx")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::End(BytesEnd::new(root_name)))
                .map_err(chart_xml_error)?;
            Ok(())
        };
    let chart_axis_id = |axis_index: usize, axis: &AxisModel| -> String {
        axis.raw_id
            .clone()
            .unwrap_or_else(|| ((axis_index + 1) * 10).to_string())
    };
    let chart_axis_ref_matches_model = |axis_id: &str| -> bool {
        chart
            .axes
            .iter()
            .enumerate()
            .any(|(axis_index, axis)| chart_axis_id(axis_index, axis) == axis_id)
    };
    let write_chart_axis_ref_element =
        |writer: &mut Writer<Cursor<Vec<u8>>>, axis_id: &str| -> OmResult<()> {
            let escaped_axis_id = partial_escape(axis_id).to_string();
            let mut axis_id_element = BytesStart::new("c:axId");
            axis_id_element.push_attribute(("val", escaped_axis_id.as_str()));
            writer
                .write_event(Event::Empty(axis_id_element))
                .map_err(chart_xml_error)?;
            Ok(())
        };
    let write_chart_val_element =
        |writer: &mut Writer<Cursor<Vec<u8>>>, element_name: &str, value: f64| -> OmResult<()> {
            let value = chart_number_xml_value(value);
            let mut element = BytesStart::new(element_name);
            element.push_attribute(("val", value.as_str()));
            writer
                .write_event(Event::Empty(element))
                .map_err(chart_xml_error)?;
            Ok(())
        };
    let write_chart_u32_val_element =
        |writer: &mut Writer<Cursor<Vec<u8>>>, element_name: &str, value: u32| -> OmResult<()> {
            let value = value.to_string();
            let mut element = BytesStart::new(element_name);
            element.push_attribute(("val", value.as_str()));
            writer
                .write_event(Event::Empty(element))
                .map_err(chart_xml_error)?;
            Ok(())
        };
    let write_chart_point_explosion_element =
        |writer: &mut Writer<Cursor<Vec<u8>>>, point_index: u32, explosion: u16| -> OmResult<()> {
            writer
                .write_event(Event::Start(BytesStart::new("c:dPt")))
                .map_err(chart_xml_error)?;
            write_chart_u32_val_element(writer, "c:idx", point_index)?;
            write_chart_u32_val_element(writer, "c:explosion", u32::from(explosion))?;
            writer
                .write_event(Event::End(BytesEnd::new("c:dPt")))
                .map_err(chart_xml_error)?;
            Ok(())
        };
    let write_chart_string_val_element =
        |writer: &mut Writer<Cursor<Vec<u8>>>, element_name: &str, value: &str| -> OmResult<()> {
            let mut element = BytesStart::new(element_name);
            element.push_attribute(("val", value));
            writer
                .write_event(Event::Empty(element))
                .map_err(chart_xml_error)?;
            Ok(())
        };
    let write_chart_series_marker_children = |writer: &mut Writer<Cursor<Vec<u8>>>,
                                              marker_style: Option<&str>,
                                              marker_size: Option<&str>|
     -> OmResult<()> {
        if let Some(marker_style) = marker_style {
            write_chart_string_val_element(writer, "c:symbol", marker_style)?;
        }
        if let Some(marker_size) = marker_size {
            write_chart_string_val_element(writer, "c:size", marker_size)?;
        }
        Ok(())
    };
    let write_chart_series_marker_element = |writer: &mut Writer<Cursor<Vec<u8>>>,
                                             marker_style: Option<&str>,
                                             marker_size: Option<&str>|
     -> OmResult<()> {
        writer
            .write_event(Event::Start(BytesStart::new("c:marker")))
            .map_err(chart_xml_error)?;
        write_chart_series_marker_children(writer, marker_style, marker_size)?;
        writer
            .write_event(Event::End(BytesEnd::new("c:marker")))
            .map_err(chart_xml_error)?;
        Ok(())
    };
    let write_chart_axis_scaling_element =
        |writer: &mut Writer<Cursor<Vec<u8>>>, axis: &AxisModel| -> OmResult<()> {
            if !chart_axis_has_scaling_xml(axis) {
                return Ok(());
            }
            writer
                .write_event(Event::Start(BytesStart::new("c:scaling")))
                .map_err(chart_xml_error)?;
            if let Some(value) = chart_axis_log_base_xml_value(axis) {
                write_chart_val_element(writer, "c:logBase", value)?;
            }
            if let Some(value) = axis.reverse_plot_order {
                write_chart_string_val_element(
                    writer,
                    "c:orientation",
                    chart_axis_orientation_xml_value(value),
                )?;
            }
            if let Some(value) = axis.minimum_scale {
                write_chart_val_element(writer, "c:min", value)?;
            }
            if let Some(value) = axis.maximum_scale {
                write_chart_val_element(writer, "c:max", value)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("c:scaling")))
                .map_err(chart_xml_error)?;
            Ok(())
        };
    let write_chart_axis_crossing_elements =
        |writer: &mut Writer<Cursor<Vec<u8>>>, axis: &AxisModel| -> OmResult<()> {
            if let Some(value) = axis.crosses_at {
                write_chart_val_element(writer, "c:crossesAt", value)?;
            } else if let Some(value) = axis.crosses.and_then(chart_axis_crosses_xml_value) {
                write_chart_string_val_element(writer, "c:crosses", value)?;
            }
            if let Some(value) = axis.axis_between_categories {
                write_chart_string_val_element(
                    writer,
                    "c:crossBetween",
                    chart_axis_between_categories_xml_value(value),
                )?;
            }
            Ok(())
        };
    let write_chart_data_labels_properties = |writer: &mut Writer<Cursor<Vec<u8>>>,
                                              data_labels: &ChartDataLabelsModel|
     -> OmResult<()> {
        if let Some(format_code) = data_labels.number_format.as_ref() {
            let mut element = BytesStart::new("c:numFmt");
            element.push_attribute(("formatCode", format_code.as_str()));
            element.push_attribute((
                "sourceLinked",
                if data_labels.number_format_linked.unwrap_or(true) {
                    "1"
                } else {
                    "0"
                },
            ));
            writer
                .write_event(Event::Empty(element))
                .map_err(chart_xml_error)?;
        }
        if let Some(position) = data_labels.position {
            let mut element = BytesStart::new("c:dLblPos");
            element.push_attribute(("val", chart_data_label_position_xml_value(position)));
            writer
                .write_event(Event::Empty(element))
                .map_err(chart_xml_error)?;
        }
        for (element_name, value) in [
            ("c:showLegendKey", data_labels.show_legend_key),
            ("c:showLeaderLines", data_labels.has_leader_lines),
            ("c:showSerName", data_labels.show_series_name),
            ("c:showCatName", data_labels.show_category_name),
            ("c:showVal", data_labels.show_value),
            ("c:showPercent", data_labels.show_percentage),
            ("c:showBubbleSize", data_labels.show_bubble_size),
        ] {
            if let Some(value) = value {
                let mut element = BytesStart::new(element_name);
                element.push_attribute(("val", if value { "1" } else { "0" }));
                writer
                    .write_event(Event::Empty(element))
                    .map_err(chart_xml_error)?;
            }
        }
        if let Some(separator) = data_labels.separator.as_ref() {
            writer
                .write_event(Event::Start(BytesStart::new("c:separator")))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                    separator,
                ))))
                .map_err(chart_xml_error)?;
            writer
                .write_event(Event::End(BytesEnd::new("c:separator")))
                .map_err(chart_xml_error)?;
        }
        Ok(())
    };
    let write_chart_data_labels_element =
        |writer: &mut Writer<Cursor<Vec<u8>>>,
         data_labels: Option<&ChartDataLabelsModel>,
         point_data_labels: Option<&BTreeMap<u32, ChartDataLabelsModel>>|
         -> OmResult<()> {
            writer
                .write_event(Event::Start(BytesStart::new("c:dLbls")))
                .map_err(chart_xml_error)?;
            if let Some(point_data_labels) = point_data_labels {
                for (point_index, data_labels) in point_data_labels {
                    writer
                        .write_event(Event::Start(BytesStart::new("c:dLbl")))
                        .map_err(chart_xml_error)?;
                    let mut index_element = BytesStart::new("c:idx");
                    let point_index = point_index.to_string();
                    index_element.push_attribute(("val", point_index.as_str()));
                    writer
                        .write_event(Event::Empty(index_element))
                        .map_err(chart_xml_error)?;
                    write_chart_data_labels_properties(writer, data_labels)?;
                    writer
                        .write_event(Event::End(BytesEnd::new("c:dLbl")))
                        .map_err(chart_xml_error)?;
                }
            }
            if let Some(data_labels) = data_labels {
                write_chart_data_labels_properties(writer, data_labels)?;
            }
            writer
                .write_event(Event::End(BytesEnd::new("c:dLbls")))
                .map_err(chart_xml_error)?;
            Ok(())
        };
    let write_chart_axis_display_units_element =
        |writer: &mut Writer<Cursor<Vec<u8>>>, axis: &AxisModel| -> OmResult<()> {
            let Some(display_unit) = axis.display_unit else {
                return Ok(());
            };
            writer
                .write_event(Event::Start(BytesStart::new("c:dispUnits")))
                .map_err(chart_xml_error)?;
            match display_unit {
                ChartAxisDisplayUnit::BuiltIn(value) => {
                    write_chart_string_val_element(
                        writer,
                        "c:builtInUnit",
                        chart_built_in_display_unit_xml_value(value),
                    )?;
                }
                ChartAxisDisplayUnit::Custom(value) => {
                    write_chart_val_element(writer, "c:custUnit", value)?;
                }
            }
            if axis.has_display_unit_label == Some(true) {
                if axis.display_unit_label.is_some() {
                    write_chart_text_element(
                        writer,
                        "c:dispUnitsLbl",
                        &chart_axis_display_unit_label_text(axis),
                    )?;
                } else {
                    writer
                        .write_event(Event::Empty(BytesStart::new("c:dispUnitsLbl")))
                        .map_err(chart_xml_error)?;
                }
            }
            writer
                .write_event(Event::End(BytesEnd::new("c:dispUnits")))
                .map_err(chart_xml_error)?;
            Ok(())
        };
    let write_chart_num_format_element = |writer: &mut Writer<Cursor<Vec<u8>>>,
                                          format_code: &str,
                                          source_linked: bool|
     -> OmResult<()> {
        let mut element = BytesStart::new("c:numFmt");
        element.push_attribute(("formatCode", format_code));
        element.push_attribute(("sourceLinked", if source_linked { "1" } else { "0" }));
        writer
            .write_event(Event::Empty(element))
            .map_err(chart_xml_error)?;
        Ok(())
    };
    let write_chart_axis_element = |writer: &mut Writer<Cursor<Vec<u8>>>,
                                    axis_index: usize,
                                    axis: &AxisModel|
     -> OmResult<()> {
        let axis_name = format!("c:{}", chart_axis_xml_name(axis.kind));
        writer
            .write_event(Event::Start(BytesStart::new(axis_name.as_str())))
            .map_err(chart_xml_error)?;
        let axis_id = chart_axis_id(axis_index, axis);
        write_chart_axis_ref_element(writer, &axis_id)?;
        write_chart_axis_scaling_element(writer, axis)?;
        if let Some(deleted) = axis.deleted {
            write_chart_string_val_element(writer, "c:delete", if deleted { "1" } else { "0" })?;
        }
        if axis.has_major_gridlines == Some(true) {
            writer
                .write_event(Event::Empty(BytesStart::new("c:majorGridlines")))
                .map_err(chart_xml_error)?;
        }
        if axis.has_minor_gridlines == Some(true) {
            writer
                .write_event(Event::Empty(BytesStart::new("c:minorGridlines")))
                .map_err(chart_xml_error)?;
        }
        if let Some(title) = axis.title.as_ref() {
            write_chart_text_element(writer, "c:title", &title.text)?;
        }
        if let Some(format_code) = axis.tick_label_number_format.as_deref() {
            write_chart_num_format_element(
                writer,
                format_code,
                axis.tick_label_number_format_linked.unwrap_or(true),
            )?;
        }
        if let Some(value) = axis.major_tick_mark {
            write_chart_string_val_element(
                writer,
                "c:majorTickMark",
                chart_tick_mark_xml_value(value),
            )?;
        }
        if let Some(value) = axis.minor_tick_mark {
            write_chart_string_val_element(
                writer,
                "c:minorTickMark",
                chart_tick_mark_xml_value(value),
            )?;
        }
        if let Some(value) = axis.tick_label_position {
            write_chart_string_val_element(
                writer,
                "c:tickLblPos",
                chart_tick_label_position_xml_value(value),
            )?;
        }
        if let Some(value) = axis.tick_label_spacing {
            write_chart_u32_val_element(writer, "c:tickLblSkip", value)?;
        }
        if let Some(value) = axis.tick_mark_spacing {
            write_chart_u32_val_element(writer, "c:tickMarkSkip", value)?;
        }
        if let Some(value) = axis.major_unit {
            write_chart_val_element(writer, "c:majorUnit", value)?;
        }
        if let Some(value) = axis.minor_unit {
            write_chart_val_element(writer, "c:minorUnit", value)?;
        }
        write_chart_axis_display_units_element(writer, axis)?;
        if let Some(axis_id) = chart_axis_cross_target_id(&chart.axes, axis) {
            write_chart_string_val_element(writer, "c:crossAx", axis_id.as_str())?;
        }
        write_chart_axis_crossing_elements(writer, axis)?;
        if let Some(value) = axis.category_type_auto {
            write_chart_string_val_element(writer, "c:auto", if value { "1" } else { "0" })?;
        }
        if let Some(value) = axis.base_unit {
            write_chart_string_val_element(
                writer,
                "c:baseTimeUnit",
                chart_axis_time_unit_xml_value(value),
            )?;
        }
        if let Some(value) = axis.major_unit_scale {
            write_chart_string_val_element(
                writer,
                "c:majorTimeUnit",
                chart_axis_time_unit_xml_value(value),
            )?;
        }
        if let Some(value) = axis.minor_unit_scale {
            write_chart_string_val_element(
                writer,
                "c:minorTimeUnit",
                chart_axis_time_unit_xml_value(value),
            )?;
        }
        writer
            .write_event(Event::End(BytesEnd::new(axis_name.as_str())))
            .map_err(chart_xml_error)?;
        Ok(())
    };

    loop {
        let event_start = usize::try_from(reader.buffer_position())
            .map_err(|_| OmError::new(OmErrorCode::InvalidState, "chart XML position overflow"))?;
        match reader.read_event_into(&mut buffer) {
            Ok(Event::Start(element)) if dirty_source_extension.is_some() => {
                let (capture, depth) = dirty_source_extension
                    .as_mut()
                    .expect("dirty source extension capture");
                capture
                    .write_event(Event::Start(element.into_owned()))
                    .map_err(chart_xml_error)?;
                *depth += 1;
            }
            Ok(Event::End(element)) if dirty_source_extension.is_some() => {
                let capture_complete = {
                    let (capture, depth) = dirty_source_extension
                        .as_mut()
                        .expect("dirty source extension capture");
                    capture
                        .write_event(Event::End(element.into_owned()))
                        .map_err(chart_xml_error)?;
                    *depth = depth.saturating_sub(1);
                    *depth == 0
                };
                if capture_complete {
                    let (capture, _) = dirty_source_extension
                        .take()
                        .expect("completed dirty source extension capture");
                    if let Some(extension_xml) = chart_extension_without_full_reference(
                        capture.into_inner().into_inner().as_slice(),
                    )? {
                        writer
                            .get_mut()
                            .write_all(extension_xml.as_slice())
                            .map_err(chart_xml_error)?;
                    }
                }
            }
            Ok(event) if dirty_source_extension.is_some() => {
                dirty_source_extension
                    .as_mut()
                    .expect("dirty source extension capture")
                    .0
                    .write_event(event.into_owned())
                    .map_err(chart_xml_error)?;
            }
            Ok(Event::Start(_)) if skip_depth > 0 => {
                skip_depth += 1;
            }
            Ok(Event::End(_)) if skip_depth > 0 => {
                skip_depth -= 1;
            }
            Ok(_) if skip_depth > 0 => {}
            Ok(Event::Start(element)) => {
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                let parent_name = element_stack.last().map(Vec::as_slice);
                let grandparent_name = element_stack
                    .len()
                    .checked_sub(2)
                    .and_then(|index| element_stack.get(index))
                    .map(Vec::as_slice);
                let depth = element_stack.len() + 1;
                if local_name.as_slice() == b"ext"
                    && let Some(series_index) = current_series_index
                    && let Some(slot) = source_stack.last().copied()
                    && source_for_slot(series_index, slot)
                        .is_some_and(|source| source.dirty && source.full_reference.is_none())
                {
                    let mut capture = Writer::new(Cursor::new(Vec::new()));
                    capture
                        .write_event(Event::Start(element.into_owned()))
                        .map_err(chart_xml_error)?;
                    dirty_source_extension = Some((capture, 1));
                    buffer.clear();
                    continue;
                }
                if parent_name == Some(b"plotArea".as_slice())
                    && local_name.as_slice() != b"layout"
                    && !plot_area_layout_container_seen
                    && let Some(Some(layout)) = expected_plot_area_layout
                {
                    writer
                        .get_mut()
                        .write_all(
                            format!(
                                "<c:layout>{}</c:layout>",
                                chart_manual_layout_xml_string(layout)
                            )
                            .as_bytes(),
                        )
                        .map_err(chart_xml_error)?;
                    plot_area_layout_container_seen = true;
                    plot_area_manual_layout_inserted = true;
                    plot_area_manual_layout_written = true;
                }
                if parent_name == Some(b"layout".as_slice())
                    && grandparent_name == Some(b"plotArea".as_slice())
                    && local_name.as_slice() != b"manualLayout"
                    && !plot_area_manual_layout_seen
                    && !plot_area_manual_layout_inserted
                    && let Some(Some(layout)) = expected_plot_area_layout
                {
                    writer
                        .get_mut()
                        .write_all(chart_manual_layout_xml_string(layout).as_bytes())
                        .map_err(chart_xml_error)?;
                    plot_area_manual_layout_inserted = true;
                    plot_area_manual_layout_written = true;
                }
                if let Some(next_chart_type) = chart_type_from_group_name(local_name.as_slice()) {
                    if chart_type.is_none() {
                        chart_type = Some(next_chart_type);
                    }
                    current_chart_group_depth = Some(depth);
                    chart_group_axis_refs_seen.clear();
                }
                if preserve_loaded_group_types
                    && parent_name
                        .is_some_and(|parent| chart_type_from_group_name(parent).is_some())
                    && chart_group_direct_property_name(local_name.as_slice())
                {
                    let element = element.into_owned();
                    buffer.clear();
                    copy_chart_xml_subtree(&mut reader, &mut writer, &mut buffer, element)?;
                    buffer.clear();
                    continue;
                }
                if local_name.as_slice() == b"ser" {
                    let Some(series_index) = loaded_series_model_indices
                        .get(next_loaded_series_index)
                        .copied()
                        .flatten()
                    else {
                        next_loaded_series_index += 1;
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    };
                    current_series_index = Some(series_index);
                    next_loaded_series_index += 1;
                    if let Some(emitted) = series_emitted.get_mut(series_index) {
                        *emitted = true;
                    }
                    source_stack.clear();
                } else if let Some(series_index) = current_series_index {
                    if local_name.as_slice() == b"dPt" && parent_name == Some(b"ser".as_slice()) {
                        current_point_index = None;
                    } else if local_name.as_slice() == b"idx"
                        && parent_name == Some(b"dPt".as_slice())
                    {
                        current_point_index = element_val_attribute(&element, reader.decoder())?
                            .and_then(|value| value.parse::<u32>().ok());
                    } else if local_name.as_slice() == b"dLbls"
                        && parent_name == Some(b"ser".as_slice())
                    {
                        if let Some(expected_points) =
                            expected_dirty_point_explosions.get(series_index)
                        {
                            for (point_index, explosion) in expected_points {
                                let already_inserted = series_point_explosions_inserted
                                    .get(series_index)
                                    .is_some_and(|inserted| inserted.contains(point_index));
                                if !already_inserted {
                                    write_chart_point_explosion_element(
                                        &mut writer,
                                        *point_index,
                                        *explosion,
                                    )?;
                                    if let Some(inserted) =
                                        series_point_explosions_inserted.get_mut(series_index)
                                    {
                                        inserted.insert(*point_index);
                                    }
                                }
                            }
                        }
                        if let Some(seen) = series_data_labels_seen.get_mut(series_index) {
                            *seen = true;
                        }
                        if expected_dirty_series_data_label_sets
                            .get(series_index)
                            .copied()
                            .unwrap_or(false)
                        {
                            if !series_data_labels_inserted
                                .get(series_index)
                                .copied()
                                .unwrap_or(false)
                            {
                                let series = &chart.series[series_index];
                                write_chart_data_labels_element(
                                    &mut writer,
                                    series.data_labels.as_ref(),
                                    Some(&series.point_data_labels),
                                )?;
                                if let Some(written) =
                                    series_data_labels_written.get_mut(series_index)
                                {
                                    *written = true;
                                }
                            }
                            skip_depth = 1;
                            buffer.clear();
                            continue;
                        }
                    } else if let Some(slot) = source_container_slot(local_name.as_slice()) {
                        let Some(source) = source_for_slot(series_index, slot) else {
                            skip_depth = 1;
                            buffer.clear();
                            continue;
                        };
                        if let Some(expected_points) =
                            expected_dirty_point_explosions.get(series_index)
                        {
                            for (point_index, explosion) in expected_points {
                                let already_inserted = series_point_explosions_inserted
                                    .get(series_index)
                                    .is_some_and(|inserted| inserted.contains(point_index));
                                if !already_inserted {
                                    write_chart_point_explosion_element(
                                        &mut writer,
                                        *point_index,
                                        *explosion,
                                    )?;
                                    if let Some(inserted) =
                                        series_point_explosions_inserted.get_mut(series_index)
                                    {
                                        inserted.insert(*point_index);
                                    }
                                }
                            }
                        }
                        if !series_data_labels_seen
                            .get(series_index)
                            .copied()
                            .unwrap_or(false)
                            && !series_data_labels_inserted
                                .get(series_index)
                                .copied()
                                .unwrap_or(false)
                            && expected_dirty_series_data_label_sets
                                .get(series_index)
                                .copied()
                                .unwrap_or(false)
                        {
                            let series = &chart.series[series_index];
                            write_chart_data_labels_element(
                                &mut writer,
                                series.data_labels.as_ref(),
                                Some(&series.point_data_labels),
                            )?;
                            if let Some(inserted) =
                                series_data_labels_inserted.get_mut(series_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = series_data_labels_written.get_mut(series_index)
                            {
                                *written = true;
                            }
                        }
                        for prior_slot in
                            source_slots_in_order.iter().copied().take(slot_index(slot))
                        {
                            if !source_slots_seen
                                .get(series_index)
                                .map(|seen| seen[slot_index(prior_slot)])
                                .unwrap_or(false)
                                && let Some(source) = source_for_slot(series_index, prior_slot)
                            {
                                write_chart_source_container(
                                    &mut writer,
                                    series_index,
                                    prior_slot,
                                    source,
                                )?;
                                if let Some(seen) = source_slots_seen.get_mut(series_index) {
                                    seen[slot_index(prior_slot)] = true;
                                }
                                if source.dirty {
                                    patched_sources += 1;
                                }
                            }
                        }
                        if source.dirty && chart_source_literal_values(source)?.is_some() {
                            write_chart_source_container(&mut writer, series_index, slot, source)?;
                            if let Some(seen) = source_slots_seen.get_mut(series_index) {
                                seen[slot_index(slot)] = true;
                            }
                            patched_sources += 1;
                            skip_depth = 1;
                            buffer.clear();
                            continue;
                        }
                        if let Some(seen) = source_slots_seen.get_mut(series_index) {
                            seen[slot_index(slot)] = true;
                        }
                        source_stack.push(slot);
                    } else if local_name.as_slice() == b"f"
                        && let Some(slot) = source_stack.last().copied()
                    {
                        current_formula = Some((slot, false));
                        if let Some(seen) = source_slots_seen.get_mut(series_index) {
                            seen[slot_index(slot)] = true;
                        }
                    } else if local_name.as_slice() == b"sqref"
                        && parent_name == Some(b"fullRef".as_slice())
                        && let Some(slot) = source_stack.last().copied()
                    {
                        current_full_reference = Some((slot, false));
                    }
                }
                if loaded_axis_start_positions
                    .binary_search(&event_start)
                    .is_ok()
                    && let Some(axis_kind) = chart_axis_kind_from_xml_name(local_name.as_slice())
                {
                    let axis_index = axis_kinds.len();
                    if chart
                        .axes
                        .get(axis_index)
                        .is_none_or(|axis| axis.kind != axis_kind)
                    {
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                    current_axis_index = Some(axis_index);
                    current_axis_depth = Some(depth);
                    axis_kinds.push(axis_kind);
                    axis_title_texts.push(None);
                    axis_title_text_written.push(false);
                    axis_major_gridlines_seen.push(false);
                    axis_major_gridlines_written.push(false);
                    axis_major_gridlines_inserted.push(false);
                    axis_major_gridlines_removed.push(false);
                    axis_minor_gridlines_seen.push(false);
                    axis_minor_gridlines_written.push(false);
                    axis_minor_gridlines_inserted.push(false);
                    axis_minor_gridlines_removed.push(false);
                    axis_major_tick_mark_seen.push(false);
                    axis_major_tick_mark_written.push(false);
                    axis_major_tick_mark_inserted.push(false);
                    axis_major_tick_mark_removed.push(false);
                    axis_minor_tick_mark_seen.push(false);
                    axis_minor_tick_mark_written.push(false);
                    axis_minor_tick_mark_inserted.push(false);
                    axis_minor_tick_mark_removed.push(false);
                    axis_tick_label_position_seen.push(false);
                    axis_tick_label_position_written.push(false);
                    axis_tick_label_position_inserted.push(false);
                    axis_tick_label_position_removed.push(false);
                    axis_tick_label_number_format_seen.push(false);
                    axis_tick_label_number_format_written.push(false);
                    axis_tick_label_number_format_inserted.push(false);
                    axis_tick_label_number_format_removed.push(false);
                    axis_tick_label_spacing_seen.push(false);
                    axis_tick_label_spacing_written.push(false);
                    axis_tick_label_spacing_inserted.push(false);
                    axis_tick_label_spacing_removed.push(false);
                    axis_tick_mark_spacing_seen.push(false);
                    axis_tick_mark_spacing_written.push(false);
                    axis_tick_mark_spacing_inserted.push(false);
                    axis_tick_mark_spacing_removed.push(false);
                    axis_scaling_seen.push(false);
                    axis_log_base_seen.push(false);
                    axis_log_base_written.push(false);
                    axis_log_base_inserted.push(false);
                    axis_log_base_removed.push(false);
                    axis_orientation_seen.push(false);
                    axis_orientation_written.push(false);
                    axis_orientation_inserted.push(false);
                    axis_orientation_removed.push(false);
                    axis_minimum_scale_seen.push(false);
                    axis_minimum_scale_written.push(false);
                    axis_minimum_scale_inserted.push(false);
                    axis_minimum_scale_removed.push(false);
                    axis_maximum_scale_seen.push(false);
                    axis_maximum_scale_written.push(false);
                    axis_maximum_scale_inserted.push(false);
                    axis_maximum_scale_removed.push(false);
                    axis_major_unit_seen.push(false);
                    axis_major_unit_written.push(false);
                    axis_major_unit_inserted.push(false);
                    axis_major_unit_removed.push(false);
                    axis_minor_unit_seen.push(false);
                    axis_minor_unit_written.push(false);
                    axis_minor_unit_inserted.push(false);
                    axis_minor_unit_removed.push(false);
                    axis_display_units_seen.push(false);
                    axis_display_units_written.push(false);
                    axis_display_units_inserted.push(false);
                    axis_display_units_removed.push(false);
                    axis_crosses_seen.push(false);
                    axis_crosses_written.push(false);
                    axis_crosses_inserted.push(false);
                    axis_crosses_removed.push(false);
                    axis_crosses_at_seen.push(false);
                    axis_crosses_at_written.push(false);
                    axis_crosses_at_inserted.push(false);
                    axis_crosses_at_removed.push(false);
                    axis_cross_between_seen.push(false);
                    axis_cross_between_written.push(false);
                    axis_cross_between_inserted.push(false);
                    axis_cross_between_removed.push(false);
                    axis_category_type_auto_seen.push(false);
                    axis_category_type_auto_written.push(false);
                    axis_category_type_auto_inserted.push(false);
                    axis_category_type_auto_removed.push(false);
                    axis_base_unit_seen.push(false);
                    axis_base_unit_written.push(false);
                    axis_base_unit_inserted.push(false);
                    axis_base_unit_removed.push(false);
                    axis_major_time_unit_seen.push(false);
                    axis_major_time_unit_written.push(false);
                    axis_major_time_unit_inserted.push(false);
                    axis_major_time_unit_removed.push(false);
                    axis_minor_time_unit_seen.push(false);
                    axis_minor_time_unit_written.push(false);
                    axis_minor_time_unit_inserted.push(false);
                    axis_minor_time_unit_removed.push(false);
                }
                if local_name.as_slice() == b"scaling"
                    && let Some(axis_index) = current_axis_index
                    && let Some(seen) = axis_scaling_seen.get_mut(axis_index)
                {
                    *seen = true;
                }
                if local_name.as_slice() == b"layout" && parent_name == Some(b"plotArea".as_slice())
                {
                    plot_area_layout_container_seen = true;
                }
                if local_name.as_slice() == b"manualLayout"
                    && parent_name == Some(b"layout".as_slice())
                    && grandparent_name == Some(b"plotArea".as_slice())
                {
                    plot_area_manual_layout_seen = true;
                    match expected_plot_area_layout {
                        Some(Some(layout)) => {
                            writer
                                .get_mut()
                                .write_all(chart_manual_layout_xml_string(layout).as_bytes())
                                .map_err(chart_xml_error)?;
                            plot_area_manual_layout_written = true;
                            skip_depth = 1;
                            buffer.clear();
                            continue;
                        }
                        Some(None) => {
                            plot_area_manual_layout_removed = true;
                            skip_depth = 1;
                            buffer.clear();
                            continue;
                        }
                        None => {}
                    }
                }
                if local_name.as_slice() == b"title" && parent_name == Some(b"chart".as_slice()) {
                    chart_title_seen = true;
                    if chart.title.is_none() {
                        chart_title_removed = true;
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    } else {
                        title_stack.push((depth, ChartTextXmlTarget::ChartTitle));
                    }
                } else if local_name.as_slice() == b"title"
                    && parent_name.is_some_and(|parent_name| {
                        chart_axis_kind_from_xml_name(parent_name).is_some()
                    })
                    && let Some(axis_index) = current_axis_index
                {
                    if chart
                        .axes
                        .get(axis_index)
                        .is_some_and(|axis| axis.title.is_none())
                    {
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                    title_stack.push((depth, ChartTextXmlTarget::AxisTitle(axis_index)));
                    if let Some(axis_title_text) = axis_title_texts.get_mut(axis_index) {
                        axis_title_text.get_or_insert_with(String::new);
                    }
                } else if local_name.as_slice() == b"dTable"
                    && parent_name == Some(b"plotArea".as_slice())
                {
                    data_table_seen = true;
                    match expected_data_table {
                        Some(Some(data_table)) => {
                            write_chart_data_table_element(&mut writer, data_table)?;
                            data_table_written = true;
                            skip_depth = 1;
                            buffer.clear();
                            continue;
                        }
                        Some(None) => {
                            data_table_removed = true;
                            skip_depth = 1;
                            buffer.clear();
                            continue;
                        }
                        None => {}
                    }
                } else if local_name.as_slice() == b"t"
                    && let Some((_, target)) = title_stack.last().copied()
                {
                    current_text_target = Some((target, false));
                } else if local_name.as_slice() == b"legend"
                    && parent_name == Some(b"chart".as_slice())
                {
                    legend_seen = true;
                    if expected_legend_position.is_none() {
                        legend_removed = true;
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"overlay"
                    && parent_name == Some(b"legend".as_slice())
                {
                    legend_overlay_seen = true;
                } else if local_name.as_slice() == b"view3D"
                    && parent_name == Some(b"chart".as_slice())
                {
                    view_3d_seen = true;
                    if let Some(Some(view_3d)) = expected_dirty_view_3d {
                        if write_chart_view_3d_element(&mut writer, view_3d)? {
                            view_3d_written = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"plotVisOnly"
                    && parent_name == Some(b"chart".as_slice())
                {
                    plot_visible_only_seen = true;
                } else if local_name.as_slice() == b"showDLblsOverMax"
                    && parent_name == Some(b"chart".as_slice())
                {
                    show_data_labels_over_maximum_seen = true;
                } else if local_name.as_slice() == b"dispBlanksAs"
                    && parent_name == Some(b"chart".as_slice())
                {
                    display_blanks_as_seen = true;
                } else if local_name.as_slice() == b"roundedCorners"
                    && parent_name == Some(b"chartSpace".as_slice())
                {
                    rounded_corners_seen = true;
                } else if local_name.as_slice() == b"style"
                    && parent_name == Some(b"chartSpace".as_slice())
                {
                    chart_style_seen = true;
                } else if local_name.as_slice() == b"protection"
                    && parent_name == Some(b"chartSpace".as_slice())
                {
                    chart_protection_seen = true;
                    if let Some(protection) = expected_chart_protection {
                        if let Some(protection) = protection {
                            if write_chart_protection_element(&mut writer, protection)? {
                                chart_protection_written = true;
                            } else {
                                chart_protection_removed = true;
                            }
                        } else {
                            chart_protection_removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"varyColors"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    vary_colors_seen = true;
                    if expected_vary_colors.is_some() {
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"barDir"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    bar_direction_seen = true;
                } else if local_name.as_slice() == b"grouping"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    chart_grouping_seen = true;
                } else if local_name.as_slice() == b"shape"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    bar_shape_seen = true;
                    if expected_bar_shape.is_none() {
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"marker"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    line_marker_seen = true;
                } else if local_name.as_slice() == b"scatterStyle"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    scatter_style_seen = true;
                } else if local_name.as_slice() == b"radarStyle"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    radar_style_seen = true;
                } else if local_name.as_slice() == b"ofPieType"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    of_pie_type_seen = true;
                } else if local_name.as_slice() == b"wireframe"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    surface_wireframe_seen = true;
                } else if local_name.as_slice() == b"gapWidth"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    gap_width_seen = true;
                } else if local_name.as_slice() == b"gapDepth"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    gap_depth_seen = true;
                } else if local_name.as_slice() == b"overlap"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    overlap_seen = true;
                } else if local_name.as_slice() == b"dLbls"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    data_labels_seen = true;
                    if let Some(data_labels) = expected_dirty_data_labels {
                        write_chart_data_labels_element(&mut writer, Some(data_labels), None)?;
                        data_labels_written = true;
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if current_chart_group_depth == Some(element_stack.len())
                    && let Some((setting_index, (_, _, expected))) =
                        expected_chart_group_numeric_settings
                            .iter()
                            .enumerate()
                            .find(|(_, (name, _, _))| local_name.as_slice() == *name)
                {
                    chart_group_numeric_setting_seen[setting_index] = true;
                    if expected.is_none() {
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if current_chart_group_depth == Some(element_stack.len())
                    && let Some((flag_index, (_, _, expected))) = expected_chart_group_line_flags
                        .iter()
                        .enumerate()
                        .find(|(_, (name, _, _))| local_name.as_slice() == *name)
                {
                    chart_group_line_flag_seen[flag_index] = true;
                    if *expected == Some(false) {
                        chart_group_line_flag_removed[flag_index] = true;
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                    if *expected == Some(true) {
                        chart_group_line_flag_written[flag_index] = true;
                    }
                }
                if matches!(local_name.as_slice(), b"majorGridlines" | b"minorGridlines")
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) =
                        if local_name.as_slice() == b"majorGridlines" {
                            (
                                &mut axis_major_gridlines_seen,
                                &mut axis_major_gridlines_written,
                                &mut axis_major_gridlines_removed,
                                axis.has_major_gridlines,
                            )
                        } else {
                            (
                                &mut axis_minor_gridlines_seen,
                                &mut axis_minor_gridlines_written,
                                &mut axis_minor_gridlines_removed,
                                axis.has_minor_gridlines,
                            )
                        };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    match expected {
                        Some(false) => {
                            if let Some(removed) = removed.get_mut(axis_index) {
                                *removed = true;
                            }
                            skip_depth = 1;
                            buffer.clear();
                            continue;
                        }
                        Some(true) => {
                            if let Some(written) = written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                        None => {}
                    }
                }

                if local_name.as_slice() == b"plotArea"
                    && !chart_title_seen
                    && let Some(title) = chart.title.as_ref()
                {
                    write_chart_text_element(&mut writer, "c:title", &title.text)?;
                    chart_title_inserted = true;
                    chart_title_text_written = true;
                }
                if local_name.as_slice() == b"plotArea"
                    && !view_3d_seen
                    && let Some(Some(view_3d)) = expected_dirty_view_3d
                    && write_chart_view_3d_element(&mut writer, view_3d)?
                {
                    view_3d_inserted = true;
                    view_3d_written = true;
                }
                if local_name.as_slice() == b"chart"
                    && parent_name == Some(b"chartSpace".as_slice())
                    && !rounded_corners_seen
                    && let Some(value) = expected_rounded_corners
                {
                    let mut rounded_corners = BytesStart::new("c:roundedCorners");
                    rounded_corners.push_attribute(("val", value));
                    writer
                        .write_event(Event::Empty(rounded_corners))
                        .map_err(chart_xml_error)?;
                    rounded_corners_inserted = true;
                    rounded_corners_written = true;
                }
                if local_name.as_slice() == b"chart"
                    && parent_name == Some(b"chartSpace".as_slice())
                    && !chart_style_seen
                    && let Some(value) = expected_chart_style
                {
                    let mut style = BytesStart::new("c:style");
                    style.push_attribute(("val", value));
                    writer
                        .write_event(Event::Empty(style))
                        .map_err(chart_xml_error)?;
                    chart_style_inserted = true;
                    chart_style_written = true;
                }
                if local_name.as_slice() == b"chart"
                    && parent_name == Some(b"chartSpace".as_slice())
                    && !chart_protection_seen
                    && expected_chart_protection_needs_xml
                    && let Some(Some(protection)) = expected_chart_protection
                    && write_chart_protection_element(&mut writer, protection)?
                {
                    chart_protection_inserted = true;
                    chart_protection_written = true;
                }

                let mut wrote_start_element = false;
                if local_name.as_slice() == b"order"
                    && let Some(series_index) = current_series_index
                    && let Some(series) = chart.series.get(series_index)
                {
                    if let Some(seen) = series_order_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    let order = series.order.unwrap_or(series_index as u32).to_string();
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            order.as_str(),
                        )?))
                        .map_err(chart_xml_error)?;
                    if let Some(written) = series_order_written.get_mut(series_index) {
                        *written = true;
                    }
                    wrote_start_element = true;
                }
                if !wrote_start_element
                    && local_name.as_slice() == b"explosion"
                    && parent_name == Some(b"dPt".as_slice())
                    && current_series_index.is_some_and(|series_index| {
                        !chart_type_supports_explosion(&model_series_chart_types[series_index])
                    })
                {
                    skip_depth = 1;
                    buffer.clear();
                    continue;
                }
                if !wrote_start_element
                    && local_name.as_slice() == b"explosion"
                    && parent_name == Some(b"dPt".as_slice())
                    && let Some(series_index) = current_series_index
                    && let Some(point_index) = current_point_index
                    && let Some(explosion) = expected_point_explosion(series_index, point_index)
                {
                    let explosion = explosion.to_string();
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            explosion.as_str(),
                        )?))
                        .map_err(chart_xml_error)?;
                    if let Some(inserted) = series_point_explosions_inserted.get_mut(series_index) {
                        inserted.insert(point_index);
                    }
                    wrote_start_element = true;
                }
                if !wrote_start_element
                    && local_name.as_slice() == b"explosion"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_explosion_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_explosions
                        .get(series_index)
                        .and_then(Option::as_deref)
                    {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_explosion_written.get_mut(series_index) {
                            *written = true;
                        }
                        wrote_start_element = true;
                    } else {
                        if let Some(removed) = series_explosion_removed.get_mut(series_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                }
                if !wrote_start_element
                    && local_name.as_slice() == b"shape"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_bar_shape_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_bar_shapes
                        .get(series_index)
                        .copied()
                        .flatten()
                    {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_bar_shape_written.get_mut(series_index) {
                            *written = true;
                        }
                        wrote_start_element = true;
                    } else {
                        if let Some(removed) = series_bar_shape_removed.get_mut(series_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                }
                if !wrote_start_element
                    && local_name.as_slice() == b"smooth"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_smooth_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_smooth_values
                        .get(series_index)
                        .copied()
                        .flatten()
                    {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_smooth_written.get_mut(series_index) {
                            *written = true;
                        }
                        wrote_start_element = true;
                    } else {
                        if let Some(removed) = series_smooth_removed.get_mut(series_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                }
                if !wrote_start_element
                    && local_name.as_slice() == b"invertIfNegative"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_invert_if_negative_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_invert_if_negative_values
                        .get(series_index)
                        .copied()
                        .flatten()
                    {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) =
                            series_invert_if_negative_written.get_mut(series_index)
                        {
                            *written = true;
                        }
                        wrote_start_element = true;
                    }
                }
                if !wrote_start_element
                    && local_name.as_slice() == b"marker"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_marker_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if !chart_type_supports_series_marker(&model_series_chart_types[series_index]) {
                        if let Some(removed) = series_marker_removed.get_mut(series_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                    current_series_marker_index = Some(series_index);
                    writer
                        .write_event(Event::Start(element.to_owned()))
                        .map_err(chart_xml_error)?;
                    wrote_start_element = true;
                }
                if !wrote_start_element
                    && local_name.as_slice() == b"symbol"
                    && parent_name == Some(b"marker".as_slice())
                    && let Some(series_index) = current_series_marker_index
                {
                    if let Some(seen) = series_marker_style_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_marker_styles
                        .get(series_index)
                        .copied()
                        .flatten()
                    {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_marker_style_written.get_mut(series_index) {
                            *written = true;
                        }
                        wrote_start_element = true;
                    }
                }
                if !wrote_start_element
                    && local_name.as_slice() == b"size"
                    && parent_name == Some(b"marker".as_slice())
                    && let Some(series_index) = current_series_marker_index
                {
                    if let Some(seen) = series_marker_size_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_marker_sizes
                        .get(series_index)
                        .and_then(Option::as_deref)
                    {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_marker_size_written.get_mut(series_index) {
                            *written = true;
                        }
                        wrote_start_element = true;
                    }
                }
                if !wrote_start_element
                    && let Some(slot) = source_container_slot(local_name.as_slice())
                    && let Some(series_index) = current_series_index
                {
                    let target_local_name = source_container_target_local_name(series_index, slot);
                    if target_local_name.as_bytes() != local_name.as_slice() {
                        writer
                            .write_event(Event::Start(rewrite_element_name(
                                &element,
                                reader.decoder(),
                                target_local_name,
                            )?))
                            .map_err(chart_xml_error)?;
                        wrote_start_element = true;
                    }
                }
                if !wrote_start_element
                    && let Some(source_chart_type) =
                        chart_type_from_group_name(local_name.as_slice())
                    && !preserve_loaded_group_types
                    && source_chart_type != chart.chart_type
                    && let Some(target_local_name) = target_chart_group_name
                {
                    writer
                        .write_event(Event::Start(rewrite_element_name(
                            &element,
                            reader.decoder(),
                            target_local_name,
                        )?))
                        .map_err(chart_xml_error)?;
                    chart_type_rewritten = true;
                    wrote_start_element = true;
                }
                if !wrote_start_element
                    && local_name.as_slice() == b"legendPos"
                    && expected_legend_position.is_some()
                    && element_stack
                        .iter()
                        .any(|name| name.as_slice() == b"legend")
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            expected_legend_position.expect("checked is_some"),
                        )?))
                        .map_err(chart_xml_error)?;
                    legend_position_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"overlay"
                    && parent_name == Some(b"legend".as_slice())
                    && let Some(value) = expected_legend_include_in_layout
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    legend_overlay_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"plotVisOnly"
                    && let Some(value) = expected_plot_visible_only
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    plot_visible_only_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"showDLblsOverMax"
                    && let Some(value) = expected_show_data_labels_over_maximum
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    show_data_labels_over_maximum_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"dispBlanksAs"
                    && let Some(value) = expected_display_blanks_as
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    display_blanks_as_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"roundedCorners"
                    && let Some(value) = expected_rounded_corners
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    rounded_corners_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"style"
                    && parent_name == Some(b"chartSpace".as_slice())
                    && let Some(value) = expected_chart_style
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    chart_style_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"barDir"
                    && let Some(value) = expected_bar_direction
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    bar_direction_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"grouping"
                    && let Some(value) = expected_chart_grouping
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    chart_grouping_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"shape"
                    && let Some(value) = expected_bar_shape
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    bar_shape_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"marker"
                    && let Some(value) = expected_line_marker
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    line_marker_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"scatterStyle"
                    && let Some(value) = expected_scatter_style
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    scatter_style_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"radarStyle"
                    && let Some(value) = expected_radar_style
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    radar_style_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"ofPieType"
                    && let Some(value) = expected_of_pie_type
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    of_pie_type_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"wireframe"
                    && let Some(value) = expected_surface_wireframe
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    surface_wireframe_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"gapWidth"
                    && let Some(value) = expected_gap_width
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    gap_width_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"gapDepth"
                    && let Some(value) = expected_gap_depth
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    gap_depth_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"overlap"
                    && let Some(value) = expected_overlap
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    overlap_written = true;
                } else if !wrote_start_element
                    && local_name.as_slice() == b"logBase"
                    && parent_name == Some(b"scaling".as_slice())
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_log_base_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = chart_axis_log_base_xml_value(axis) {
                        let value = chart_number_xml_value(value);
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value.as_str(),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = axis_log_base_written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) = axis_log_base_removed.get_mut(axis_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if !wrote_start_element
                    && local_name.as_slice() == b"orientation"
                    && parent_name == Some(b"scaling".as_slice())
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_orientation_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = axis.reverse_plot_order {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                chart_axis_orientation_xml_value(value),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = axis_orientation_written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else {
                        writer
                            .write_event(Event::Start(element.into_owned()))
                            .map_err(chart_xml_error)?;
                    }
                } else if !wrote_start_element
                    && matches!(local_name.as_slice(), b"min" | b"max")
                    && parent_name == Some(b"scaling".as_slice())
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) = if local_name.as_slice() == b"min" {
                        (
                            &mut axis_minimum_scale_seen,
                            &mut axis_minimum_scale_written,
                            &mut axis_minimum_scale_removed,
                            axis.minimum_scale,
                        )
                    } else {
                        (
                            &mut axis_maximum_scale_seen,
                            &mut axis_maximum_scale_written,
                            &mut axis_maximum_scale_removed,
                            axis.maximum_scale,
                        )
                    };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected {
                        let value = chart_number_xml_value(value);
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value.as_str(),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) = removed.get_mut(axis_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if !wrote_start_element
                    && matches!(local_name.as_slice(), b"majorUnit" | b"minorUnit")
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) =
                        if local_name.as_slice() == b"majorUnit" {
                            (
                                &mut axis_major_unit_seen,
                                &mut axis_major_unit_written,
                                &mut axis_major_unit_removed,
                                axis.major_unit,
                            )
                        } else {
                            (
                                &mut axis_minor_unit_seen,
                                &mut axis_minor_unit_written,
                                &mut axis_minor_unit_removed,
                                axis.minor_unit,
                            )
                        };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected {
                        let value = chart_number_xml_value(value);
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value.as_str(),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) = removed.get_mut(axis_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if !wrote_start_element
                    && local_name.as_slice() == b"dispUnits"
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_display_units_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if axis.display_unit.is_some() {
                        write_chart_axis_display_units_element(&mut writer, axis)?;
                        if let Some(written) = axis_display_units_written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else if let Some(removed) = axis_display_units_removed.get_mut(axis_index) {
                        *removed = true;
                    }
                    skip_depth = 1;
                    buffer.clear();
                    continue;
                } else if !wrote_start_element
                    && matches!(local_name.as_slice(), b"crosses" | b"crossesAt")
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if local_name.as_slice() == b"crosses" {
                        if let Some(seen) = axis_crosses_seen.get_mut(axis_index) {
                            *seen = true;
                        }
                        if axis.crosses_at.is_none()
                            && let Some(value) = axis.crosses.and_then(chart_axis_crosses_xml_value)
                        {
                            writer
                                .write_event(Event::Start(rewrite_val_attribute_element(
                                    &element,
                                    reader.decoder(),
                                    value,
                                )?))
                                .map_err(chart_xml_error)?;
                            if let Some(written) = axis_crosses_written.get_mut(axis_index) {
                                *written = true;
                            }
                        } else {
                            if let Some(removed) = axis_crosses_removed.get_mut(axis_index) {
                                *removed = true;
                            }
                            skip_depth = 1;
                            buffer.clear();
                            continue;
                        }
                    } else {
                        if let Some(seen) = axis_crosses_at_seen.get_mut(axis_index) {
                            *seen = true;
                        }
                        if let Some(value) = axis.crosses_at {
                            let value = chart_number_xml_value(value);
                            writer
                                .write_event(Event::Start(rewrite_val_attribute_element(
                                    &element,
                                    reader.decoder(),
                                    value.as_str(),
                                )?))
                                .map_err(chart_xml_error)?;
                            if let Some(written) = axis_crosses_at_written.get_mut(axis_index) {
                                *written = true;
                            }
                        } else {
                            if let Some(removed) = axis_crosses_at_removed.get_mut(axis_index) {
                                *removed = true;
                            }
                            skip_depth = 1;
                            buffer.clear();
                            continue;
                        }
                    }
                } else if !wrote_start_element
                    && local_name.as_slice() == b"crossBetween"
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_cross_between_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = axis.axis_between_categories {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                chart_axis_between_categories_xml_value(value),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = axis_cross_between_written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) = axis_cross_between_removed.get_mut(axis_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if !wrote_start_element
                    && local_name.as_slice() == b"auto"
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_category_type_auto_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = axis.category_type_auto {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                if value { "1" } else { "0" },
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = axis_category_type_auto_written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) = axis_category_type_auto_removed.get_mut(axis_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if !wrote_start_element
                    && matches!(
                        local_name.as_slice(),
                        b"baseTimeUnit" | b"majorTimeUnit" | b"minorTimeUnit"
                    )
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) = match local_name.as_slice() {
                        b"baseTimeUnit" => (
                            &mut axis_base_unit_seen,
                            &mut axis_base_unit_written,
                            &mut axis_base_unit_removed,
                            axis.base_unit.map(chart_axis_time_unit_xml_value),
                        ),
                        b"majorTimeUnit" => (
                            &mut axis_major_time_unit_seen,
                            &mut axis_major_time_unit_written,
                            &mut axis_major_time_unit_removed,
                            axis.major_unit_scale.map(chart_axis_time_unit_xml_value),
                        ),
                        b"minorTimeUnit" => (
                            &mut axis_minor_time_unit_seen,
                            &mut axis_minor_time_unit_written,
                            &mut axis_minor_time_unit_removed,
                            axis.minor_unit_scale.map(chart_axis_time_unit_xml_value),
                        ),
                        _ => unreachable!("matched axis time unit"),
                    };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) = removed.get_mut(axis_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if !wrote_start_element
                    && matches!(
                        local_name.as_slice(),
                        b"majorTickMark" | b"minorTickMark" | b"tickLblPos"
                    )
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) = match local_name.as_slice() {
                        b"majorTickMark" => (
                            &mut axis_major_tick_mark_seen,
                            &mut axis_major_tick_mark_written,
                            &mut axis_major_tick_mark_removed,
                            axis.major_tick_mark.map(chart_tick_mark_xml_value),
                        ),
                        b"minorTickMark" => (
                            &mut axis_minor_tick_mark_seen,
                            &mut axis_minor_tick_mark_written,
                            &mut axis_minor_tick_mark_removed,
                            axis.minor_tick_mark.map(chart_tick_mark_xml_value),
                        ),
                        b"tickLblPos" => (
                            &mut axis_tick_label_position_seen,
                            &mut axis_tick_label_position_written,
                            &mut axis_tick_label_position_removed,
                            axis.tick_label_position
                                .map(chart_tick_label_position_xml_value),
                        ),
                        _ => unreachable!("matched axis tick property"),
                    };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected {
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) = removed.get_mut(axis_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if !wrote_start_element
                    && local_name.as_slice() == b"numFmt"
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_tick_label_number_format_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(format_code) = axis.tick_label_number_format.as_deref() {
                        write_chart_num_format_element(
                            &mut writer,
                            format_code,
                            axis.tick_label_number_format_linked.unwrap_or(true),
                        )?;
                        if let Some(written) =
                            axis_tick_label_number_format_written.get_mut(axis_index)
                        {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) =
                            axis_tick_label_number_format_removed.get_mut(axis_index)
                        {
                            *removed = true;
                        }
                    }
                    skip_depth = 1;
                    buffer.clear();
                    continue;
                } else if !wrote_start_element
                    && matches!(local_name.as_slice(), b"tickLblSkip" | b"tickMarkSkip")
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) =
                        if local_name.as_slice() == b"tickLblSkip" {
                            (
                                &mut axis_tick_label_spacing_seen,
                                &mut axis_tick_label_spacing_written,
                                &mut axis_tick_label_spacing_removed,
                                axis.tick_label_spacing,
                            )
                        } else {
                            (
                                &mut axis_tick_mark_spacing_seen,
                                &mut axis_tick_mark_spacing_written,
                                &mut axis_tick_mark_spacing_removed,
                                axis.tick_mark_spacing,
                            )
                        };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected {
                        let value = value.to_string();
                        writer
                            .write_event(Event::Start(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value.as_str(),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) = removed.get_mut(axis_index) {
                            *removed = true;
                        }
                        skip_depth = 1;
                        buffer.clear();
                        continue;
                    }
                } else if !wrote_start_element
                    && let Some((setting_index, (_, _, Some(value)))) =
                        expected_chart_group_numeric_settings
                            .iter()
                            .enumerate()
                            .find(|(_, (name, _, _))| local_name.as_slice() == *name)
                {
                    writer
                        .write_event(Event::Start(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    chart_group_numeric_setting_written[setting_index] = true;
                } else if !wrote_start_element {
                    writer
                        .write_event(Event::Start(element.into_owned()))
                        .map_err(chart_xml_error)?;
                }
                if !preserve_loaded_group_types
                    && chart_type_from_group_name(local_name.as_slice()).is_some()
                    && let Some(value) = expected_vary_colors
                {
                    let mut vary_colors = BytesStart::new("c:varyColors");
                    vary_colors.push_attribute(("val", value));
                    writer
                        .write_event(Event::Empty(vary_colors))
                        .map_err(chart_xml_error)?;
                    vary_colors_inserted = true;
                    vary_colors_written = true;
                }
                element_stack.push(local_name);
            }
            Ok(Event::Empty(element)) => {
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                let parent_name = element_stack.last().map(Vec::as_slice);
                let grandparent_name = element_stack
                    .len()
                    .checked_sub(2)
                    .and_then(|index| element_stack.get(index))
                    .map(Vec::as_slice);
                if preserve_loaded_group_types
                    && parent_name
                        .is_some_and(|parent| chart_type_from_group_name(parent).is_some())
                    && chart_group_direct_property_name(local_name.as_slice())
                {
                    writer
                        .write_event(Event::Empty(element.into_owned()))
                        .map_err(chart_xml_error)?;
                    buffer.clear();
                    continue;
                }
                if parent_name == Some(b"plotArea".as_slice())
                    && local_name.as_slice() != b"layout"
                    && !plot_area_layout_container_seen
                    && let Some(Some(layout)) = expected_plot_area_layout
                {
                    writer
                        .get_mut()
                        .write_all(
                            format!(
                                "<c:layout>{}</c:layout>",
                                chart_manual_layout_xml_string(layout)
                            )
                            .as_bytes(),
                        )
                        .map_err(chart_xml_error)?;
                    plot_area_layout_container_seen = true;
                    plot_area_manual_layout_inserted = true;
                    plot_area_manual_layout_written = true;
                }
                if parent_name == Some(b"layout".as_slice())
                    && grandparent_name == Some(b"plotArea".as_slice())
                    && local_name.as_slice() != b"manualLayout"
                    && !plot_area_manual_layout_seen
                    && !plot_area_manual_layout_inserted
                    && let Some(Some(layout)) = expected_plot_area_layout
                {
                    writer
                        .get_mut()
                        .write_all(chart_manual_layout_xml_string(layout).as_bytes())
                        .map_err(chart_xml_error)?;
                    plot_area_manual_layout_inserted = true;
                    plot_area_manual_layout_written = true;
                }
                if local_name.as_slice() == b"layout" && parent_name == Some(b"plotArea".as_slice())
                {
                    plot_area_layout_container_seen = true;
                    if let Some(Some(layout)) = expected_plot_area_layout {
                        writer
                            .get_mut()
                            .write_all(
                                format!(
                                    "<c:layout>{}</c:layout>",
                                    chart_manual_layout_xml_string(layout)
                                )
                                .as_bytes(),
                            )
                            .map_err(chart_xml_error)?;
                        plot_area_manual_layout_inserted = true;
                        plot_area_manual_layout_written = true;
                        buffer.clear();
                        continue;
                    }
                }
                if local_name.as_slice() == b"manualLayout"
                    && parent_name == Some(b"layout".as_slice())
                    && grandparent_name == Some(b"plotArea".as_slice())
                {
                    plot_area_manual_layout_seen = true;
                    match expected_plot_area_layout {
                        Some(Some(layout)) => {
                            writer
                                .get_mut()
                                .write_all(chart_manual_layout_xml_string(layout).as_bytes())
                                .map_err(chart_xml_error)?;
                            plot_area_manual_layout_written = true;
                            buffer.clear();
                            continue;
                        }
                        Some(None) => {
                            plot_area_manual_layout_removed = true;
                            buffer.clear();
                            continue;
                        }
                        None => {}
                    }
                }
                if local_name.as_slice() == b"idx" && parent_name == Some(b"dPt".as_slice()) {
                    current_point_index = element_val_attribute(&element, reader.decoder())?
                        .and_then(|value| value.parse::<u32>().ok());
                }
                if local_name.as_slice() == b"scaling"
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_scaling_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if chart_axis_has_scaling_xml(axis) {
                        write_chart_axis_scaling_element(&mut writer, axis)?;
                        if chart_axis_log_base_xml_value(axis).is_some() {
                            if let Some(inserted) = axis_log_base_inserted.get_mut(axis_index) {
                                *inserted = true;
                            }
                            if let Some(written) = axis_log_base_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                        if axis.reverse_plot_order.is_some() {
                            if let Some(inserted) = axis_orientation_inserted.get_mut(axis_index) {
                                *inserted = true;
                            }
                            if let Some(written) = axis_orientation_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                        if axis.minimum_scale.is_some() {
                            if let Some(inserted) = axis_minimum_scale_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = axis_minimum_scale_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                        if axis.maximum_scale.is_some() {
                            if let Some(inserted) = axis_maximum_scale_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = axis_maximum_scale_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                        buffer.clear();
                        continue;
                    }
                }
                if local_name.as_slice() == b"logBase"
                    && parent_name == Some(b"scaling".as_slice())
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_log_base_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = chart_axis_log_base_xml_value(axis) {
                        let value = chart_number_xml_value(value);
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value.as_str(),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = axis_log_base_written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else if let Some(removed) = axis_log_base_removed.get_mut(axis_index) {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if local_name.as_slice() == b"orientation"
                    && parent_name == Some(b"scaling".as_slice())
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_orientation_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = axis.reverse_plot_order {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                chart_axis_orientation_xml_value(value),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = axis_orientation_written.get_mut(axis_index) {
                            *written = true;
                        }
                        buffer.clear();
                        continue;
                    }
                }
                if matches!(local_name.as_slice(), b"min" | b"max")
                    && parent_name == Some(b"scaling".as_slice())
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) = if local_name.as_slice() == b"min" {
                        (
                            &mut axis_minimum_scale_seen,
                            &mut axis_minimum_scale_written,
                            &mut axis_minimum_scale_removed,
                            axis.minimum_scale,
                        )
                    } else {
                        (
                            &mut axis_maximum_scale_seen,
                            &mut axis_maximum_scale_written,
                            &mut axis_maximum_scale_removed,
                            axis.maximum_scale,
                        )
                    };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected {
                        let value = chart_number_xml_value(value);
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value.as_str(),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else if let Some(removed) = removed.get_mut(axis_index) {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if matches!(local_name.as_slice(), b"majorUnit" | b"minorUnit")
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) =
                        if local_name.as_slice() == b"majorUnit" {
                            (
                                &mut axis_major_unit_seen,
                                &mut axis_major_unit_written,
                                &mut axis_major_unit_removed,
                                axis.major_unit,
                            )
                        } else {
                            (
                                &mut axis_minor_unit_seen,
                                &mut axis_minor_unit_written,
                                &mut axis_minor_unit_removed,
                                axis.minor_unit,
                            )
                        };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected {
                        let value = chart_number_xml_value(value);
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value.as_str(),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else if let Some(removed) = removed.get_mut(axis_index) {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if local_name.as_slice() == b"dispUnits"
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_display_units_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if axis.display_unit.is_some() {
                        write_chart_axis_display_units_element(&mut writer, axis)?;
                        if let Some(written) = axis_display_units_written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else if let Some(removed) = axis_display_units_removed.get_mut(axis_index) {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if matches!(local_name.as_slice(), b"crosses" | b"crossesAt")
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if local_name.as_slice() == b"crosses" {
                        if let Some(seen) = axis_crosses_seen.get_mut(axis_index) {
                            *seen = true;
                        }
                        if axis.crosses_at.is_none()
                            && let Some(value) = axis.crosses.and_then(chart_axis_crosses_xml_value)
                        {
                            writer
                                .write_event(Event::Empty(rewrite_val_attribute_element(
                                    &element,
                                    reader.decoder(),
                                    value,
                                )?))
                                .map_err(chart_xml_error)?;
                            if let Some(written) = axis_crosses_written.get_mut(axis_index) {
                                *written = true;
                            }
                        } else if let Some(removed) = axis_crosses_removed.get_mut(axis_index) {
                            *removed = true;
                        }
                    } else {
                        if let Some(seen) = axis_crosses_at_seen.get_mut(axis_index) {
                            *seen = true;
                        }
                        if let Some(value) = axis.crosses_at {
                            let value = chart_number_xml_value(value);
                            writer
                                .write_event(Event::Empty(rewrite_val_attribute_element(
                                    &element,
                                    reader.decoder(),
                                    value.as_str(),
                                )?))
                                .map_err(chart_xml_error)?;
                            if let Some(written) = axis_crosses_at_written.get_mut(axis_index) {
                                *written = true;
                            }
                        } else if let Some(removed) = axis_crosses_at_removed.get_mut(axis_index) {
                            *removed = true;
                        }
                    }
                    buffer.clear();
                    continue;
                }
                if local_name.as_slice() == b"crossBetween"
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_cross_between_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = axis.axis_between_categories {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                chart_axis_between_categories_xml_value(value),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = axis_cross_between_written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else if let Some(removed) = axis_cross_between_removed.get_mut(axis_index) {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if local_name.as_slice() == b"auto"
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_category_type_auto_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = axis.category_type_auto {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                if value { "1" } else { "0" },
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = axis_category_type_auto_written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else if let Some(removed) =
                        axis_category_type_auto_removed.get_mut(axis_index)
                    {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if matches!(
                    local_name.as_slice(),
                    b"baseTimeUnit" | b"majorTimeUnit" | b"minorTimeUnit"
                ) && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) = match local_name.as_slice() {
                        b"baseTimeUnit" => (
                            &mut axis_base_unit_seen,
                            &mut axis_base_unit_written,
                            &mut axis_base_unit_removed,
                            axis.base_unit.map(chart_axis_time_unit_xml_value),
                        ),
                        b"majorTimeUnit" => (
                            &mut axis_major_time_unit_seen,
                            &mut axis_major_time_unit_written,
                            &mut axis_major_time_unit_removed,
                            axis.major_unit_scale.map(chart_axis_time_unit_xml_value),
                        ),
                        b"minorTimeUnit" => (
                            &mut axis_minor_time_unit_seen,
                            &mut axis_minor_time_unit_written,
                            &mut axis_minor_time_unit_removed,
                            axis.minor_unit_scale.map(chart_axis_time_unit_xml_value),
                        ),
                        _ => unreachable!("matched axis time unit"),
                    };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else if let Some(removed) = removed.get_mut(axis_index) {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if matches!(
                    local_name.as_slice(),
                    b"majorTickMark" | b"minorTickMark" | b"tickLblPos"
                ) && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) = match local_name.as_slice() {
                        b"majorTickMark" => (
                            &mut axis_major_tick_mark_seen,
                            &mut axis_major_tick_mark_written,
                            &mut axis_major_tick_mark_removed,
                            axis.major_tick_mark.map(chart_tick_mark_xml_value),
                        ),
                        b"minorTickMark" => (
                            &mut axis_minor_tick_mark_seen,
                            &mut axis_minor_tick_mark_written,
                            &mut axis_minor_tick_mark_removed,
                            axis.minor_tick_mark.map(chart_tick_mark_xml_value),
                        ),
                        b"tickLblPos" => (
                            &mut axis_tick_label_position_seen,
                            &mut axis_tick_label_position_written,
                            &mut axis_tick_label_position_removed,
                            axis.tick_label_position
                                .map(chart_tick_label_position_xml_value),
                        ),
                        _ => unreachable!("matched axis tick property"),
                    };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else if let Some(removed) = removed.get_mut(axis_index) {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if local_name.as_slice() == b"numFmt"
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if let Some(seen) = axis_tick_label_number_format_seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(format_code) = axis.tick_label_number_format.as_deref() {
                        write_chart_num_format_element(
                            &mut writer,
                            format_code,
                            axis.tick_label_number_format_linked.unwrap_or(true),
                        )?;
                        if let Some(written) =
                            axis_tick_label_number_format_written.get_mut(axis_index)
                        {
                            *written = true;
                        }
                    } else if let Some(removed) =
                        axis_tick_label_number_format_removed.get_mut(axis_index)
                    {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if matches!(local_name.as_slice(), b"tickLblSkip" | b"tickMarkSkip")
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) =
                        if local_name.as_slice() == b"tickLblSkip" {
                            (
                                &mut axis_tick_label_spacing_seen,
                                &mut axis_tick_label_spacing_written,
                                &mut axis_tick_label_spacing_removed,
                                axis.tick_label_spacing,
                            )
                        } else {
                            (
                                &mut axis_tick_mark_spacing_seen,
                                &mut axis_tick_mark_spacing_written,
                                &mut axis_tick_mark_spacing_removed,
                                axis.tick_mark_spacing,
                            )
                        };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected {
                        let value = value.to_string();
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value.as_str(),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = written.get_mut(axis_index) {
                            *written = true;
                        }
                    } else if let Some(removed) = removed.get_mut(axis_index) {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if let Some(next_chart_type) = chart_type_from_group_name(local_name.as_slice()) {
                    if chart_type.is_none() {
                        chart_type = Some(next_chart_type);
                    }
                }
                if local_name.as_slice() == b"axId"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    if let Some(axis_id) = element_val_attribute(&element, reader.decoder())? {
                        if !chart_axis_ref_matches_model(&axis_id) {
                            buffer.clear();
                            continue;
                        }
                        chart_group_axis_refs_seen.push(axis_id);
                    }
                }
                if let Some(series_index) = current_series_index
                    && let Some(slot) = source_container_slot(local_name.as_slice())
                    && source_for_slot(series_index, slot).is_none()
                {
                    buffer.clear();
                    continue;
                }
                if local_name.as_slice() == b"order"
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_order_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(series) = chart.series.get(series_index) {
                        let order = series.order.unwrap_or(series_index as u32).to_string();
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                order.as_str(),
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_order_written.get_mut(series_index) {
                            *written = true;
                        }
                        buffer.clear();
                        continue;
                    }
                }
                if local_name.as_slice() == b"explosion"
                    && parent_name == Some(b"dPt".as_slice())
                    && current_series_index.is_some_and(|series_index| {
                        !chart_type_supports_explosion(&model_series_chart_types[series_index])
                    })
                {
                    buffer.clear();
                    continue;
                }
                if local_name.as_slice() == b"explosion"
                    && parent_name == Some(b"dPt".as_slice())
                    && let Some(series_index) = current_series_index
                    && let Some(point_index) = current_point_index
                    && let Some(explosion) = expected_point_explosion(series_index, point_index)
                {
                    let explosion = explosion.to_string();
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            explosion.as_str(),
                        )?))
                        .map_err(chart_xml_error)?;
                    if let Some(inserted) = series_point_explosions_inserted.get_mut(series_index) {
                        inserted.insert(point_index);
                    }
                    buffer.clear();
                    continue;
                }
                if local_name.as_slice() == b"explosion"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_explosion_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_explosions
                        .get(series_index)
                        .and_then(Option::as_deref)
                    {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_explosion_written.get_mut(series_index) {
                            *written = true;
                        }
                    } else if let Some(removed) = series_explosion_removed.get_mut(series_index) {
                        *removed = true;
                    }
                    buffer.clear();
                    continue;
                }
                if let Some(source_chart_type) = chart_type_from_group_name(local_name.as_slice())
                    && !preserve_loaded_group_types
                    && source_chart_type != chart.chart_type
                    && let Some(target_local_name) = target_chart_group_name
                {
                    writer
                        .write_event(Event::Empty(rewrite_element_name(
                            &element,
                            reader.decoder(),
                            target_local_name,
                        )?))
                        .map_err(chart_xml_error)?;
                    chart_type_rewritten = true;
                    buffer.clear();
                    continue;
                }
                if local_name.as_slice() == b"title" && parent_name == Some(b"chart".as_slice()) {
                    chart_title_seen = true;
                    if chart.title.is_none() {
                        chart_title_removed = true;
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"title"
                    && parent_name.is_some_and(|parent_name| {
                        chart_axis_kind_from_xml_name(parent_name).is_some()
                    })
                    && let Some(axis_index) = current_axis_index
                    && chart
                        .axes
                        .get(axis_index)
                        .is_some_and(|axis| axis.title.is_none())
                {
                    buffer.clear();
                    continue;
                } else if local_name.as_slice() == b"legend"
                    && parent_name == Some(b"chart".as_slice())
                {
                    legend_seen = true;
                    if expected_legend_position.is_none() {
                        legend_removed = true;
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"overlay"
                    && parent_name == Some(b"legend".as_slice())
                {
                    legend_overlay_seen = true;
                } else if local_name.as_slice() == b"view3D"
                    && parent_name == Some(b"chart".as_slice())
                {
                    view_3d_seen = true;
                    if let Some(Some(view_3d)) = expected_dirty_view_3d {
                        if write_chart_view_3d_element(&mut writer, view_3d)? {
                            view_3d_written = true;
                        }
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"plotVisOnly"
                    && parent_name == Some(b"chart".as_slice())
                {
                    plot_visible_only_seen = true;
                } else if local_name.as_slice() == b"showDLblsOverMax"
                    && parent_name == Some(b"chart".as_slice())
                {
                    show_data_labels_over_maximum_seen = true;
                } else if local_name.as_slice() == b"dispBlanksAs"
                    && parent_name == Some(b"chart".as_slice())
                {
                    display_blanks_as_seen = true;
                } else if local_name.as_slice() == b"roundedCorners"
                    && parent_name == Some(b"chartSpace".as_slice())
                {
                    rounded_corners_seen = true;
                } else if local_name.as_slice() == b"style"
                    && parent_name == Some(b"chartSpace".as_slice())
                {
                    chart_style_seen = true;
                } else if local_name.as_slice() == b"protection"
                    && parent_name == Some(b"chartSpace".as_slice())
                {
                    chart_protection_seen = true;
                    if let Some(protection) = expected_chart_protection {
                        if let Some(protection) = protection {
                            if write_chart_protection_element(&mut writer, protection)? {
                                chart_protection_written = true;
                            } else {
                                chart_protection_removed = true;
                            }
                        } else {
                            chart_protection_removed = true;
                        }
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"varyColors"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    vary_colors_seen = true;
                    if expected_vary_colors.is_some() {
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"barDir"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    bar_direction_seen = true;
                } else if local_name.as_slice() == b"grouping"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    chart_grouping_seen = true;
                } else if local_name.as_slice() == b"shape"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    bar_shape_seen = true;
                    if expected_bar_shape.is_none() {
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"marker"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    line_marker_seen = true;
                } else if local_name.as_slice() == b"scatterStyle"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    scatter_style_seen = true;
                } else if local_name.as_slice() == b"radarStyle"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    radar_style_seen = true;
                } else if local_name.as_slice() == b"ofPieType"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    of_pie_type_seen = true;
                } else if local_name.as_slice() == b"wireframe"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    surface_wireframe_seen = true;
                } else if local_name.as_slice() == b"gapWidth"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    gap_width_seen = true;
                } else if local_name.as_slice() == b"gapDepth"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    gap_depth_seen = true;
                } else if local_name.as_slice() == b"overlap"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    overlap_seen = true;
                } else if local_name.as_slice() == b"dTable"
                    && parent_name == Some(b"plotArea".as_slice())
                {
                    data_table_seen = true;
                    match expected_data_table {
                        Some(Some(data_table)) => {
                            write_chart_data_table_element(&mut writer, data_table)?;
                            data_table_written = true;
                            buffer.clear();
                            continue;
                        }
                        Some(None) => {
                            data_table_removed = true;
                            buffer.clear();
                            continue;
                        }
                        None => {}
                    }
                } else if local_name.as_slice() == b"dLbls"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_data_labels_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if expected_dirty_series_data_label_sets
                        .get(series_index)
                        .copied()
                        .unwrap_or(false)
                    {
                        if !series_data_labels_inserted
                            .get(series_index)
                            .copied()
                            .unwrap_or(false)
                        {
                            let series = &chart.series[series_index];
                            write_chart_data_labels_element(
                                &mut writer,
                                series.data_labels.as_ref(),
                                Some(&series.point_data_labels),
                            )?;
                            if let Some(written) = series_data_labels_written.get_mut(series_index)
                            {
                                *written = true;
                            }
                        }
                        buffer.clear();
                        continue;
                    }
                } else if local_name.as_slice() == b"dLbls"
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    data_labels_seen = true;
                    if let Some(data_labels) = expected_dirty_data_labels {
                        write_chart_data_labels_element(&mut writer, Some(data_labels), None)?;
                        data_labels_written = true;
                        buffer.clear();
                        continue;
                    }
                } else if current_chart_group_depth == Some(element_stack.len())
                    && let Some((setting_index, (_, _, expected))) =
                        expected_chart_group_numeric_settings
                            .iter()
                            .enumerate()
                            .find(|(_, (name, _, _))| local_name.as_slice() == *name)
                {
                    chart_group_numeric_setting_seen[setting_index] = true;
                    if expected.is_none() {
                        buffer.clear();
                        continue;
                    }
                } else if current_chart_group_depth == Some(element_stack.len())
                    && let Some((flag_index, (_, _, expected))) = expected_chart_group_line_flags
                        .iter()
                        .enumerate()
                        .find(|(_, (name, _, _))| local_name.as_slice() == *name)
                {
                    chart_group_line_flag_seen[flag_index] = true;
                    if *expected == Some(false) {
                        chart_group_line_flag_removed[flag_index] = true;
                        buffer.clear();
                        continue;
                    }
                    if *expected == Some(true) {
                        chart_group_line_flag_written[flag_index] = true;
                    }
                }
                if matches!(local_name.as_slice(), b"majorGridlines" | b"minorGridlines")
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    let (seen, written, removed, expected) =
                        if local_name.as_slice() == b"majorGridlines" {
                            (
                                &mut axis_major_gridlines_seen,
                                &mut axis_major_gridlines_written,
                                &mut axis_major_gridlines_removed,
                                axis.has_major_gridlines,
                            )
                        } else {
                            (
                                &mut axis_minor_gridlines_seen,
                                &mut axis_minor_gridlines_written,
                                &mut axis_minor_gridlines_removed,
                                axis.has_minor_gridlines,
                            )
                        };
                    if let Some(seen) = seen.get_mut(axis_index) {
                        *seen = true;
                    }
                    match expected {
                        Some(false) => {
                            if let Some(removed) = removed.get_mut(axis_index) {
                                *removed = true;
                            }
                            buffer.clear();
                            continue;
                        }
                        Some(true) => {
                            if let Some(written) = written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                        None => {}
                    }
                }
                if local_name.as_slice() == b"plotArea"
                    && !view_3d_seen
                    && let Some(Some(view_3d)) = expected_dirty_view_3d
                    && write_chart_view_3d_element(&mut writer, view_3d)?
                {
                    view_3d_inserted = true;
                    view_3d_written = true;
                }
                if local_name.as_slice() == b"legendPos"
                    && expected_legend_position.is_some()
                    && element_stack
                        .iter()
                        .any(|name| name.as_slice() == b"legend")
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            expected_legend_position.expect("checked is_some"),
                        )?))
                        .map_err(chart_xml_error)?;
                    legend_position_written = true;
                } else if local_name.as_slice() == b"overlay"
                    && parent_name == Some(b"legend".as_slice())
                    && let Some(value) = expected_legend_include_in_layout
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    legend_overlay_written = true;
                } else if local_name.as_slice() == b"plotVisOnly"
                    && let Some(value) = expected_plot_visible_only
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    plot_visible_only_written = true;
                } else if local_name.as_slice() == b"showDLblsOverMax"
                    && let Some(value) = expected_show_data_labels_over_maximum
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    show_data_labels_over_maximum_written = true;
                } else if local_name.as_slice() == b"dispBlanksAs"
                    && let Some(value) = expected_display_blanks_as
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    display_blanks_as_written = true;
                } else if local_name.as_slice() == b"roundedCorners"
                    && let Some(value) = expected_rounded_corners
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    rounded_corners_written = true;
                } else if local_name.as_slice() == b"style"
                    && parent_name == Some(b"chartSpace".as_slice())
                    && let Some(value) = expected_chart_style
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    chart_style_written = true;
                } else if local_name.as_slice() == b"barDir"
                    && let Some(value) = expected_bar_direction
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    bar_direction_written = true;
                } else if local_name.as_slice() == b"grouping"
                    && let Some(value) = expected_chart_grouping
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    chart_grouping_written = true;
                } else if local_name.as_slice() == b"shape"
                    && let Some(value) = expected_bar_shape
                    && parent_name != Some(b"ser".as_slice())
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    bar_shape_written = true;
                } else if local_name.as_slice() == b"shape"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_bar_shape_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_bar_shapes
                        .get(series_index)
                        .copied()
                        .flatten()
                    {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_bar_shape_written.get_mut(series_index) {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) = series_bar_shape_removed.get_mut(series_index) {
                            *removed = true;
                        }
                    }
                } else if local_name.as_slice() == b"smooth"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_smooth_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_smooth_values
                        .get(series_index)
                        .copied()
                        .flatten()
                    {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_smooth_written.get_mut(series_index) {
                            *written = true;
                        }
                    } else {
                        if let Some(removed) = series_smooth_removed.get_mut(series_index) {
                            *removed = true;
                        }
                    }
                } else if local_name.as_slice() == b"invertIfNegative"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_invert_if_negative_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_invert_if_negative_values
                        .get(series_index)
                        .copied()
                        .flatten()
                    {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) =
                            series_invert_if_negative_written.get_mut(series_index)
                        {
                            *written = true;
                        }
                    } else {
                        writer
                            .write_event(Event::Empty(element.into_owned()))
                            .map_err(chart_xml_error)?;
                    }
                } else if local_name.as_slice() == b"symbol"
                    && parent_name == Some(b"marker".as_slice())
                    && let Some(series_index) = current_series_marker_index
                {
                    if let Some(seen) = series_marker_style_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_marker_styles
                        .get(series_index)
                        .copied()
                        .flatten()
                    {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_marker_style_written.get_mut(series_index) {
                            *written = true;
                        }
                    } else {
                        writer
                            .write_event(Event::Empty(element.into_owned()))
                            .map_err(chart_xml_error)?;
                    }
                } else if local_name.as_slice() == b"size"
                    && parent_name == Some(b"marker".as_slice())
                    && let Some(series_index) = current_series_marker_index
                {
                    if let Some(seen) = series_marker_size_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if let Some(value) = expected_series_marker_sizes
                        .get(series_index)
                        .and_then(Option::as_deref)
                    {
                        writer
                            .write_event(Event::Empty(rewrite_val_attribute_element(
                                &element,
                                reader.decoder(),
                                value,
                            )?))
                            .map_err(chart_xml_error)?;
                        if let Some(written) = series_marker_size_written.get_mut(series_index) {
                            *written = true;
                        }
                    } else {
                        writer
                            .write_event(Event::Empty(element.into_owned()))
                            .map_err(chart_xml_error)?;
                    }
                } else if local_name.as_slice() == b"marker"
                    && parent_name == Some(b"ser".as_slice())
                    && let Some(series_index) = current_series_index
                {
                    if let Some(seen) = series_marker_seen.get_mut(series_index) {
                        *seen = true;
                    }
                    if !chart_type_supports_series_marker(&model_series_chart_types[series_index]) {
                        if let Some(removed) = series_marker_removed.get_mut(series_index) {
                            *removed = true;
                        }
                    } else {
                        let marker_style = expected_series_marker_styles
                            .get(series_index)
                            .copied()
                            .flatten();
                        let marker_size = expected_series_marker_sizes
                            .get(series_index)
                            .and_then(Option::as_deref);
                        if marker_style.is_some() || marker_size.is_some() {
                            write_chart_series_marker_element(
                                &mut writer,
                                marker_style,
                                marker_size,
                            )?;
                            if marker_style.is_some() {
                                if let Some(inserted) =
                                    series_marker_style_inserted.get_mut(series_index)
                                {
                                    *inserted = true;
                                }
                                if let Some(written) =
                                    series_marker_style_written.get_mut(series_index)
                                {
                                    *written = true;
                                }
                            }
                            if marker_size.is_some() {
                                if let Some(inserted) =
                                    series_marker_size_inserted.get_mut(series_index)
                                {
                                    *inserted = true;
                                }
                                if let Some(written) =
                                    series_marker_size_written.get_mut(series_index)
                                {
                                    *written = true;
                                }
                            }
                        } else {
                            writer
                                .write_event(Event::Empty(element.into_owned()))
                                .map_err(chart_xml_error)?;
                        }
                    }
                } else if local_name.as_slice() == b"marker"
                    && let Some(value) = expected_line_marker
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    line_marker_written = true;
                } else if local_name.as_slice() == b"scatterStyle"
                    && let Some(value) = expected_scatter_style
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    scatter_style_written = true;
                } else if local_name.as_slice() == b"radarStyle"
                    && let Some(value) = expected_radar_style
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    radar_style_written = true;
                } else if local_name.as_slice() == b"ofPieType"
                    && let Some(value) = expected_of_pie_type
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    of_pie_type_written = true;
                } else if local_name.as_slice() == b"wireframe"
                    && let Some(value) = expected_surface_wireframe
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    surface_wireframe_written = true;
                } else if local_name.as_slice() == b"gapWidth"
                    && let Some(value) = expected_gap_width
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    gap_width_written = true;
                } else if local_name.as_slice() == b"gapDepth"
                    && let Some(value) = expected_gap_depth
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    gap_depth_written = true;
                } else if local_name.as_slice() == b"overlap"
                    && let Some(value) = expected_overlap
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    overlap_written = true;
                } else if let Some((setting_index, (_, _, Some(value)))) =
                    expected_chart_group_numeric_settings
                        .iter()
                        .enumerate()
                        .find(|(_, (name, _, _))| local_name.as_slice() == *name)
                {
                    writer
                        .write_event(Event::Empty(rewrite_val_attribute_element(
                            &element,
                            reader.decoder(),
                            value,
                        )?))
                        .map_err(chart_xml_error)?;
                    chart_group_numeric_setting_written[setting_index] = true;
                } else {
                    writer
                        .write_event(Event::Empty(element.into_owned()))
                        .map_err(chart_xml_error)?;
                }
            }
            Ok(Event::Text(text)) => {
                if let Some((target, written)) = current_text_target.as_mut() {
                    match target {
                        ChartTextXmlTarget::ChartTitle => {
                            if let Some(title) = chart.title.as_ref() {
                                if !chart_title_text_written {
                                    writer
                                        .write_event(Event::Text(BytesText::from_escaped(
                                            partial_escape(&title.text),
                                        )))
                                        .map_err(chart_xml_error)?;
                                    chart_title_text_written = true;
                                } else {
                                    writer
                                        .write_event(Event::Text(BytesText::new("")))
                                        .map_err(chart_xml_error)?;
                                }
                                *written = true;
                            } else {
                                writer
                                    .write_event(Event::Text(text.into_owned()))
                                    .map_err(chart_xml_error)?;
                            }
                        }
                        ChartTextXmlTarget::AxisTitle(axis_index) => {
                            if let Some(title) = chart
                                .axes
                                .get(*axis_index)
                                .and_then(|axis| axis.title.as_ref())
                            {
                                if !axis_title_text_written
                                    .get(*axis_index)
                                    .copied()
                                    .unwrap_or(false)
                                {
                                    writer
                                        .write_event(Event::Text(BytesText::from_escaped(
                                            partial_escape(&title.text),
                                        )))
                                        .map_err(chart_xml_error)?;
                                    if let Some(written) =
                                        axis_title_text_written.get_mut(*axis_index)
                                    {
                                        *written = true;
                                    }
                                } else {
                                    writer
                                        .write_event(Event::Text(BytesText::new("")))
                                        .map_err(chart_xml_error)?;
                                }
                                *written = true;
                            } else {
                                if let Some(axis_title_text) = axis_title_texts
                                    .get_mut(*axis_index)
                                    .and_then(Option::as_mut)
                                {
                                    axis_title_text
                                        .push_str(&text.xml_content().map_err(chart_xml_error)?);
                                }
                                writer
                                    .write_event(Event::Text(text.into_owned()))
                                    .map_err(chart_xml_error)?;
                            }
                        }
                    }
                } else if let Some((slot, written)) = current_formula.as_mut()
                    && let Some(series_index) = current_series_index
                    && let Some(source) = source_for_slot(series_index, *slot)
                    && source.dirty
                {
                    if !*written {
                        let replacement = source.raw.text.trim_start_matches('=').to_string();
                        writer
                            .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                &replacement,
                            ))))
                            .map_err(chart_xml_error)?;
                        *written = true;
                        patched_sources += 1;
                    }
                } else if let Some((slot, written)) = current_full_reference.as_mut()
                    && let Some(series_index) = current_series_index
                    && let Some(reference) = source_for_slot(series_index, *slot)
                        .and_then(|source| source.full_reference.as_ref())
                    && source_for_slot(series_index, *slot).is_some_and(|source| source.dirty)
                {
                    if !*written {
                        writer
                            .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                reference.raw.text.trim_start_matches('='),
                            ))))
                            .map_err(chart_xml_error)?;
                        *written = true;
                    }
                } else {
                    writer
                        .write_event(Event::Text(text.into_owned()))
                        .map_err(chart_xml_error)?;
                }
            }
            Ok(Event::CData(data)) => {
                if let Some((target, written)) = current_text_target.as_mut() {
                    match target {
                        ChartTextXmlTarget::ChartTitle => {
                            if let Some(title) = chart.title.as_ref() {
                                if !chart_title_text_written {
                                    writer
                                        .write_event(Event::Text(BytesText::from_escaped(
                                            partial_escape(&title.text),
                                        )))
                                        .map_err(chart_xml_error)?;
                                    chart_title_text_written = true;
                                } else {
                                    writer
                                        .write_event(Event::Text(BytesText::new("")))
                                        .map_err(chart_xml_error)?;
                                }
                                *written = true;
                            } else {
                                writer
                                    .write_event(Event::CData(data.into_owned()))
                                    .map_err(chart_xml_error)?;
                            }
                        }
                        ChartTextXmlTarget::AxisTitle(axis_index) => {
                            if let Some(title) = chart
                                .axes
                                .get(*axis_index)
                                .and_then(|axis| axis.title.as_ref())
                            {
                                if !axis_title_text_written
                                    .get(*axis_index)
                                    .copied()
                                    .unwrap_or(false)
                                {
                                    writer
                                        .write_event(Event::Text(BytesText::from_escaped(
                                            partial_escape(&title.text),
                                        )))
                                        .map_err(chart_xml_error)?;
                                    if let Some(written) =
                                        axis_title_text_written.get_mut(*axis_index)
                                    {
                                        *written = true;
                                    }
                                } else {
                                    writer
                                        .write_event(Event::Text(BytesText::new("")))
                                        .map_err(chart_xml_error)?;
                                }
                                *written = true;
                            } else {
                                if let Some(axis_title_text) = axis_title_texts
                                    .get_mut(*axis_index)
                                    .and_then(Option::as_mut)
                                {
                                    axis_title_text
                                        .push_str(&data.xml_content().map_err(chart_xml_error)?);
                                }
                                writer
                                    .write_event(Event::CData(data.into_owned()))
                                    .map_err(chart_xml_error)?;
                            }
                        }
                    }
                } else if let Some((slot, written)) = current_formula.as_mut()
                    && let Some(series_index) = current_series_index
                    && let Some(source) = source_for_slot(series_index, *slot)
                    && source.dirty
                {
                    if !*written {
                        let replacement = source.raw.text.trim_start_matches('=').to_string();
                        writer
                            .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                &replacement,
                            ))))
                            .map_err(chart_xml_error)?;
                        *written = true;
                        patched_sources += 1;
                    }
                } else if let Some((slot, written)) = current_full_reference.as_mut()
                    && let Some(series_index) = current_series_index
                    && let Some(reference) = source_for_slot(series_index, *slot)
                        .and_then(|source| source.full_reference.as_ref())
                    && source_for_slot(series_index, *slot).is_some_and(|source| source.dirty)
                {
                    if !*written {
                        writer
                            .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                reference.raw.text.trim_start_matches('='),
                            ))))
                            .map_err(chart_xml_error)?;
                        *written = true;
                    }
                } else {
                    writer
                        .write_event(Event::CData(data.into_owned()))
                        .map_err(chart_xml_error)?;
                }
            }
            Ok(Event::GeneralRef(reference)) => {
                if let Some((target, written)) = current_text_target.as_mut() {
                    match target {
                        ChartTextXmlTarget::ChartTitle => {
                            if let Some(title) = chart.title.as_ref() {
                                if !chart_title_text_written {
                                    writer
                                        .write_event(Event::Text(BytesText::from_escaped(
                                            partial_escape(&title.text),
                                        )))
                                        .map_err(chart_xml_error)?;
                                    chart_title_text_written = true;
                                }
                                *written = true;
                            } else {
                                writer
                                    .write_event(Event::GeneralRef(reference.into_owned()))
                                    .map_err(chart_xml_error)?;
                            }
                        }
                        ChartTextXmlTarget::AxisTitle(axis_index) => {
                            if let Some(title) = chart
                                .axes
                                .get(*axis_index)
                                .and_then(|axis| axis.title.as_ref())
                            {
                                if !axis_title_text_written
                                    .get(*axis_index)
                                    .copied()
                                    .unwrap_or(false)
                                {
                                    writer
                                        .write_event(Event::Text(BytesText::from_escaped(
                                            partial_escape(&title.text),
                                        )))
                                        .map_err(chart_xml_error)?;
                                    if let Some(written) =
                                        axis_title_text_written.get_mut(*axis_index)
                                    {
                                        *written = true;
                                    }
                                }
                                *written = true;
                            } else {
                                if let Some(axis_title_text) = axis_title_texts
                                    .get_mut(*axis_index)
                                    .and_then(Option::as_mut)
                                {
                                    axis_title_text.push_str(&decode_general_ref_text(&reference)?);
                                }
                                writer
                                    .write_event(Event::GeneralRef(reference.into_owned()))
                                    .map_err(chart_xml_error)?;
                            }
                        }
                    }
                } else if let Some((slot, written)) = current_formula.as_mut()
                    && let Some(series_index) = current_series_index
                    && let Some(source) = source_for_slot(series_index, *slot)
                    && source.dirty
                {
                    if !*written {
                        let replacement = source.raw.text.trim_start_matches('=').to_string();
                        writer
                            .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                &replacement,
                            ))))
                            .map_err(chart_xml_error)?;
                        *written = true;
                        patched_sources += 1;
                    }
                } else if let Some((slot, written)) = current_full_reference.as_mut()
                    && let Some(series_index) = current_series_index
                    && let Some(reference) = source_for_slot(series_index, *slot)
                        .and_then(|source| source.full_reference.as_ref())
                    && source_for_slot(series_index, *slot).is_some_and(|source| source.dirty)
                {
                    if !*written {
                        writer
                            .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                                reference.raw.text.trim_start_matches('='),
                            ))))
                            .map_err(chart_xml_error)?;
                        *written = true;
                    }
                } else {
                    writer
                        .write_event(Event::GeneralRef(reference.into_owned()))
                        .map_err(chart_xml_error)?;
                }
            }
            Ok(Event::End(element)) => {
                let local_name = xml_local_name(element.name().as_ref()).to_vec();
                let parent_name = element_stack
                    .len()
                    .checked_sub(2)
                    .and_then(|index| element_stack.get(index))
                    .map(Vec::as_slice);
                if local_name.as_slice() == b"f"
                    && let Some((slot, written)) = current_formula.take()
                    && !written
                    && let Some(series_index) = current_series_index
                    && let Some(source) = source_for_slot(series_index, slot)
                    && source.dirty
                {
                    let replacement = source.raw.text.trim_start_matches('=').to_string();
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                            &replacement,
                        ))))
                        .map_err(chart_xml_error)?;
                    patched_sources += 1;
                }
                if local_name.as_slice() == b"sqref"
                    && let Some((slot, written)) = current_full_reference.take()
                    && !written
                    && let Some(series_index) = current_series_index
                    && let Some(reference) = source_for_slot(series_index, slot)
                        .and_then(|source| source.full_reference.as_ref())
                    && source_for_slot(series_index, slot).is_some_and(|source| source.dirty)
                {
                    writer
                        .write_event(Event::Text(BytesText::from_escaped(partial_escape(
                            reference.raw.text.trim_start_matches('='),
                        ))))
                        .map_err(chart_xml_error)?;
                }
                if local_name.as_slice() == b"marker"
                    && let Some(series_index) = current_series_marker_index
                {
                    let marker_style = expected_series_marker_styles
                        .get(series_index)
                        .copied()
                        .flatten();
                    let marker_size = expected_series_marker_sizes
                        .get(series_index)
                        .and_then(Option::as_deref);
                    if let Some(marker_style) = marker_style
                        && !series_marker_style_seen
                            .get(series_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        write_chart_string_val_element(&mut writer, "c:symbol", marker_style)?;
                        if let Some(inserted) = series_marker_style_inserted.get_mut(series_index) {
                            *inserted = true;
                        }
                        if let Some(written) = series_marker_style_written.get_mut(series_index) {
                            *written = true;
                        }
                    }
                    if let Some(marker_size) = marker_size
                        && !series_marker_size_seen
                            .get(series_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        write_chart_string_val_element(&mut writer, "c:size", marker_size)?;
                        if let Some(inserted) = series_marker_size_inserted.get_mut(series_index) {
                            *inserted = true;
                        }
                        if let Some(written) = series_marker_size_written.get_mut(series_index) {
                            *written = true;
                        }
                    }
                }
                if local_name.as_slice() == b"ser"
                    && let Some(series_index) = current_series_index
                {
                    if !series_explosion_seen
                        .get(series_index)
                        .copied()
                        .unwrap_or(false)
                        && let Some(value) = expected_series_explosions
                            .get(series_index)
                            .and_then(Option::as_deref)
                    {
                        let mut explosion = BytesStart::new("c:explosion");
                        explosion.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(explosion))
                            .map_err(chart_xml_error)?;
                        if let Some(inserted) = series_explosion_inserted.get_mut(series_index) {
                            *inserted = true;
                        }
                        if let Some(written) = series_explosion_written.get_mut(series_index) {
                            *written = true;
                        }
                    }
                    if let Some(expected_points) = expected_dirty_point_explosions.get(series_index)
                    {
                        for (point_index, explosion) in expected_points {
                            let already_inserted = series_point_explosions_inserted
                                .get(series_index)
                                .is_some_and(|inserted| inserted.contains(point_index));
                            if !already_inserted {
                                write_chart_point_explosion_element(
                                    &mut writer,
                                    *point_index,
                                    *explosion,
                                )?;
                                if let Some(inserted) =
                                    series_point_explosions_inserted.get_mut(series_index)
                                {
                                    inserted.insert(*point_index);
                                }
                            }
                        }
                    }
                    if !series_data_labels_seen
                        .get(series_index)
                        .copied()
                        .unwrap_or(false)
                        && !series_data_labels_inserted
                            .get(series_index)
                            .copied()
                            .unwrap_or(false)
                        && expected_dirty_series_data_label_sets
                            .get(series_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        let series = &chart.series[series_index];
                        write_chart_data_labels_element(
                            &mut writer,
                            series.data_labels.as_ref(),
                            Some(&series.point_data_labels),
                        )?;
                        if let Some(inserted) = series_data_labels_inserted.get_mut(series_index) {
                            *inserted = true;
                        }
                        if let Some(written) = series_data_labels_written.get_mut(series_index) {
                            *written = true;
                        }
                    }
                    for slot in source_slots_in_order.iter().copied() {
                        if !source_slots_seen
                            .get(series_index)
                            .map(|seen| seen[slot_index(slot)])
                            .unwrap_or(false)
                            && let Some(source) = source_for_slot(series_index, slot)
                        {
                            write_chart_source_container(&mut writer, series_index, slot, source)?;
                            if let Some(seen) = source_slots_seen.get_mut(series_index) {
                                seen[slot_index(slot)] = true;
                            }
                            if source.dirty {
                                patched_sources += 1;
                            }
                        }
                    }
                    if !series_marker_seen
                        .get(series_index)
                        .copied()
                        .unwrap_or(false)
                    {
                        let marker_style = expected_series_marker_styles
                            .get(series_index)
                            .copied()
                            .flatten();
                        let marker_size = expected_series_marker_sizes
                            .get(series_index)
                            .and_then(Option::as_deref);
                        if marker_style.is_some() || marker_size.is_some() {
                            write_chart_series_marker_element(
                                &mut writer,
                                marker_style,
                                marker_size,
                            )?;
                            if let Some(inserted) = series_marker_inserted.get_mut(series_index) {
                                *inserted = true;
                            }
                            if marker_style.is_some() {
                                if let Some(inserted) =
                                    series_marker_style_inserted.get_mut(series_index)
                                {
                                    *inserted = true;
                                }
                                if let Some(written) =
                                    series_marker_style_written.get_mut(series_index)
                                {
                                    *written = true;
                                }
                            }
                            if marker_size.is_some() {
                                if let Some(inserted) =
                                    series_marker_size_inserted.get_mut(series_index)
                                {
                                    *inserted = true;
                                }
                                if let Some(written) =
                                    series_marker_size_written.get_mut(series_index)
                                {
                                    *written = true;
                                }
                            }
                        }
                    }
                    if !series_invert_if_negative_seen
                        .get(series_index)
                        .copied()
                        .unwrap_or(false)
                        && let Some(value) = expected_series_invert_if_negative_values
                            .get(series_index)
                            .copied()
                            .flatten()
                    {
                        let mut invert_if_negative = BytesStart::new("c:invertIfNegative");
                        invert_if_negative.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(invert_if_negative))
                            .map_err(chart_xml_error)?;
                        if let Some(inserted) =
                            series_invert_if_negative_inserted.get_mut(series_index)
                        {
                            *inserted = true;
                        }
                        if let Some(written) =
                            series_invert_if_negative_written.get_mut(series_index)
                        {
                            *written = true;
                        }
                    }
                }
                if local_name.as_slice() == b"t"
                    && let Some((target, written)) = current_text_target.take()
                {
                    match target {
                        ChartTextXmlTarget::ChartTitle => {
                            if !written
                                && !chart_title_text_written
                                && let Some(title) = chart.title.as_ref()
                            {
                                writer
                                    .write_event(Event::Text(BytesText::from_escaped(
                                        partial_escape(&title.text),
                                    )))
                                    .map_err(chart_xml_error)?;
                                chart_title_text_written = true;
                            }
                        }
                        ChartTextXmlTarget::AxisTitle(axis_index) => {
                            if !written
                                && !axis_title_text_written
                                    .get(axis_index)
                                    .copied()
                                    .unwrap_or(false)
                                && let Some(title) = chart
                                    .axes
                                    .get(axis_index)
                                    .and_then(|axis| axis.title.as_ref())
                            {
                                writer
                                    .write_event(Event::Text(BytesText::from_escaped(
                                        partial_escape(&title.text),
                                    )))
                                    .map_err(chart_xml_error)?;
                                if let Some(written) = axis_title_text_written.get_mut(axis_index) {
                                    *written = true;
                                }
                            }
                        }
                    }
                }

                if local_name.as_slice() == b"scaling"
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if chart_axis_log_base_xml_value(axis).is_some()
                        && !axis_log_base_seen.get(axis_index).copied().unwrap_or(false)
                    {
                        if let Some(value) = chart_axis_log_base_xml_value(axis) {
                            write_chart_val_element(&mut writer, "c:logBase", value)?;
                            if let Some(inserted) = axis_log_base_inserted.get_mut(axis_index) {
                                *inserted = true;
                            }
                            if let Some(written) = axis_log_base_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                    }
                    if axis.reverse_plot_order.is_some()
                        && !axis_orientation_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.reverse_plot_order {
                            write_chart_string_val_element(
                                &mut writer,
                                "c:orientation",
                                chart_axis_orientation_xml_value(value),
                            )?;
                            if let Some(inserted) = axis_orientation_inserted.get_mut(axis_index) {
                                *inserted = true;
                            }
                            if let Some(written) = axis_orientation_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                    }
                    if axis.minimum_scale.is_some()
                        && !axis_minimum_scale_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.minimum_scale {
                            write_chart_val_element(&mut writer, "c:min", value)?;
                            if let Some(inserted) = axis_minimum_scale_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = axis_minimum_scale_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                    }
                    if axis.maximum_scale.is_some()
                        && !axis_maximum_scale_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.maximum_scale {
                            write_chart_val_element(&mut writer, "c:max", value)?;
                            if let Some(inserted) = axis_maximum_scale_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = axis_maximum_scale_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                    }
                }

                if local_name.as_slice() == b"chart"
                    && !legend_seen
                    && let Some(position) = expected_legend_position
                {
                    writer
                        .write_event(Event::Start(BytesStart::new("c:legend")))
                        .map_err(chart_xml_error)?;
                    let mut legend_pos = BytesStart::new("c:legendPos");
                    legend_pos.push_attribute(("val", position));
                    writer
                        .write_event(Event::Empty(legend_pos))
                        .map_err(chart_xml_error)?;
                    if let Some(value) = expected_legend_include_in_layout {
                        let mut overlay = BytesStart::new("c:overlay");
                        overlay.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(overlay))
                            .map_err(chart_xml_error)?;
                        legend_overlay_inserted = true;
                        legend_overlay_written = true;
                    }
                    writer
                        .write_event(Event::End(BytesEnd::new("c:legend")))
                        .map_err(chart_xml_error)?;
                    legend_inserted = true;
                    legend_position_written = true;
                }
                if local_name.as_slice() == b"chart" {
                    if !plot_visible_only_seen && let Some(value) = expected_plot_visible_only {
                        let mut plot_visible_only = BytesStart::new("c:plotVisOnly");
                        plot_visible_only.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(plot_visible_only))
                            .map_err(chart_xml_error)?;
                        plot_visible_only_inserted = true;
                        plot_visible_only_written = true;
                    }
                    if !display_blanks_as_seen && let Some(value) = expected_display_blanks_as {
                        let mut display_blanks_as = BytesStart::new("c:dispBlanksAs");
                        display_blanks_as.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(display_blanks_as))
                            .map_err(chart_xml_error)?;
                        display_blanks_as_inserted = true;
                        display_blanks_as_written = true;
                    }
                    if !show_data_labels_over_maximum_seen
                        && let Some(value) = expected_show_data_labels_over_maximum
                    {
                        let mut show_data_labels = BytesStart::new("c:showDLblsOverMax");
                        show_data_labels.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(show_data_labels))
                            .map_err(chart_xml_error)?;
                        show_data_labels_over_maximum_inserted = true;
                        show_data_labels_over_maximum_written = true;
                    }
                }
                if current_axis_depth == Some(element_stack.len())
                    && let Some(axis_index) = current_axis_index
                    && let Some(axis) = chart.axes.get(axis_index)
                {
                    if chart_axis_has_scaling_xml(axis)
                        && !axis_scaling_seen.get(axis_index).copied().unwrap_or(false)
                    {
                        write_chart_axis_scaling_element(&mut writer, axis)?;
                        if chart_axis_log_base_xml_value(axis).is_some() {
                            if let Some(inserted) = axis_log_base_inserted.get_mut(axis_index) {
                                *inserted = true;
                            }
                            if let Some(written) = axis_log_base_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                        if axis.reverse_plot_order.is_some() {
                            if let Some(inserted) = axis_orientation_inserted.get_mut(axis_index) {
                                *inserted = true;
                            }
                            if let Some(written) = axis_orientation_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                        if axis.minimum_scale.is_some() {
                            if let Some(inserted) = axis_minimum_scale_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = axis_minimum_scale_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                        if axis.maximum_scale.is_some() {
                            if let Some(inserted) = axis_maximum_scale_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = axis_maximum_scale_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                    }
                    if axis.has_major_gridlines == Some(true)
                        && !axis_major_gridlines_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        writer
                            .write_event(Event::Empty(BytesStart::new("c:majorGridlines")))
                            .map_err(chart_xml_error)?;
                        if let Some(inserted) = axis_major_gridlines_inserted.get_mut(axis_index) {
                            *inserted = true;
                        }
                        if let Some(written) = axis_major_gridlines_written.get_mut(axis_index) {
                            *written = true;
                        }
                    }
                    if axis.has_minor_gridlines == Some(true)
                        && !axis_minor_gridlines_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        writer
                            .write_event(Event::Empty(BytesStart::new("c:minorGridlines")))
                            .map_err(chart_xml_error)?;
                        if let Some(inserted) = axis_minor_gridlines_inserted.get_mut(axis_index) {
                            *inserted = true;
                        }
                        if let Some(written) = axis_minor_gridlines_written.get_mut(axis_index) {
                            *written = true;
                        }
                    }
                    if axis_title_texts
                        .get(axis_index)
                        .is_some_and(|title_text| title_text.is_none())
                        && let Some(title) = axis.title.as_ref()
                    {
                        write_chart_text_element(&mut writer, "c:title", &title.text)?;
                        if let Some(written) = axis_title_text_written.get_mut(axis_index) {
                            *written = true;
                        }
                    }
                    if axis.tick_label_number_format.is_some()
                        && !axis_tick_label_number_format_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(format_code) = axis.tick_label_number_format.as_deref() {
                            write_chart_num_format_element(
                                &mut writer,
                                format_code,
                                axis.tick_label_number_format_linked.unwrap_or(true),
                            )?;
                            if let Some(inserted) =
                                axis_tick_label_number_format_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) =
                                axis_tick_label_number_format_written.get_mut(axis_index)
                            {
                                *written = true;
                            }
                        }
                    }
                    if axis.major_tick_mark.is_some()
                        && !axis_major_tick_mark_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.major_tick_mark {
                            write_chart_string_val_element(
                                &mut writer,
                                "c:majorTickMark",
                                chart_tick_mark_xml_value(value),
                            )?;
                            if let Some(inserted) =
                                axis_major_tick_mark_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = axis_major_tick_mark_written.get_mut(axis_index)
                            {
                                *written = true;
                            }
                        }
                    }
                    if axis.minor_tick_mark.is_some()
                        && !axis_minor_tick_mark_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.minor_tick_mark {
                            write_chart_string_val_element(
                                &mut writer,
                                "c:minorTickMark",
                                chart_tick_mark_xml_value(value),
                            )?;
                            if let Some(inserted) =
                                axis_minor_tick_mark_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = axis_minor_tick_mark_written.get_mut(axis_index)
                            {
                                *written = true;
                            }
                        }
                    }
                    if axis.tick_label_position.is_some()
                        && !axis_tick_label_position_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.tick_label_position {
                            write_chart_string_val_element(
                                &mut writer,
                                "c:tickLblPos",
                                chart_tick_label_position_xml_value(value),
                            )?;
                            if let Some(inserted) =
                                axis_tick_label_position_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) =
                                axis_tick_label_position_written.get_mut(axis_index)
                            {
                                *written = true;
                            }
                        }
                    }
                    if axis.tick_label_spacing.is_some()
                        && !axis_tick_label_spacing_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.tick_label_spacing {
                            write_chart_u32_val_element(&mut writer, "c:tickLblSkip", value)?;
                            if let Some(inserted) =
                                axis_tick_label_spacing_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) =
                                axis_tick_label_spacing_written.get_mut(axis_index)
                            {
                                *written = true;
                            }
                        }
                    }
                    if axis.tick_mark_spacing.is_some()
                        && !axis_tick_mark_spacing_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.tick_mark_spacing {
                            write_chart_u32_val_element(&mut writer, "c:tickMarkSkip", value)?;
                            if let Some(inserted) =
                                axis_tick_mark_spacing_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) =
                                axis_tick_mark_spacing_written.get_mut(axis_index)
                            {
                                *written = true;
                            }
                        }
                    }
                    if axis.major_unit.is_some()
                        && !axis_major_unit_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.major_unit {
                            write_chart_val_element(&mut writer, "c:majorUnit", value)?;
                            if let Some(inserted) = axis_major_unit_inserted.get_mut(axis_index) {
                                *inserted = true;
                            }
                            if let Some(written) = axis_major_unit_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                    }
                    if axis.minor_unit.is_some()
                        && !axis_minor_unit_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.minor_unit {
                            write_chart_val_element(&mut writer, "c:minorUnit", value)?;
                            if let Some(inserted) = axis_minor_unit_inserted.get_mut(axis_index) {
                                *inserted = true;
                            }
                            if let Some(written) = axis_minor_unit_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                    }
                    if axis.display_unit.is_some()
                        && !axis_display_units_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        write_chart_axis_display_units_element(&mut writer, axis)?;
                        if let Some(inserted) = axis_display_units_inserted.get_mut(axis_index) {
                            *inserted = true;
                        }
                        if let Some(written) = axis_display_units_written.get_mut(axis_index) {
                            *written = true;
                        }
                    }
                    if axis.crosses_at.is_some() {
                        if !axis_crosses_at_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                            && let Some(value) = axis.crosses_at
                        {
                            write_chart_val_element(&mut writer, "c:crossesAt", value)?;
                            if let Some(inserted) = axis_crosses_at_inserted.get_mut(axis_index) {
                                *inserted = true;
                            }
                            if let Some(written) = axis_crosses_at_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                    } else if let Some(value) = axis.crosses.and_then(chart_axis_crosses_xml_value)
                        && !axis_crosses_seen.get(axis_index).copied().unwrap_or(false)
                    {
                        write_chart_string_val_element(&mut writer, "c:crosses", value)?;
                        if let Some(inserted) = axis_crosses_inserted.get_mut(axis_index) {
                            *inserted = true;
                        }
                        if let Some(written) = axis_crosses_written.get_mut(axis_index) {
                            *written = true;
                        }
                    }
                    if axis.axis_between_categories.is_some()
                        && !axis_cross_between_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.axis_between_categories {
                            write_chart_string_val_element(
                                &mut writer,
                                "c:crossBetween",
                                chart_axis_between_categories_xml_value(value),
                            )?;
                            if let Some(inserted) = axis_cross_between_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = axis_cross_between_written.get_mut(axis_index) {
                                *written = true;
                            }
                        }
                    }
                    if axis.category_type_auto.is_some()
                        && !axis_category_type_auto_seen
                            .get(axis_index)
                            .copied()
                            .unwrap_or(false)
                    {
                        if let Some(value) = axis.category_type_auto {
                            write_chart_string_val_element(
                                &mut writer,
                                "c:auto",
                                if value { "1" } else { "0" },
                            )?;
                            if let Some(inserted) =
                                axis_category_type_auto_inserted.get_mut(axis_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) =
                                axis_category_type_auto_written.get_mut(axis_index)
                            {
                                *written = true;
                            }
                        }
                    }
                }
                if local_name.as_slice() == b"plotArea" {
                    while let Some(axis) = chart.axes.get(axis_kinds.len()) {
                        let axis_index = axis_kinds.len();
                        write_chart_axis_element(&mut writer, axis_index, axis)?;
                        axis_kinds.push(axis.kind);
                        axis_title_texts.push(axis.title.as_ref().map(|title| title.text.clone()));
                        axis_title_text_written.push(axis.title.is_some());
                        axis_major_gridlines_seen.push(axis.has_major_gridlines == Some(true));
                        axis_major_gridlines_written.push(axis.has_major_gridlines == Some(true));
                        axis_major_gridlines_inserted.push(axis.has_major_gridlines == Some(true));
                        axis_major_gridlines_removed.push(false);
                        axis_minor_gridlines_seen.push(axis.has_minor_gridlines == Some(true));
                        axis_minor_gridlines_written.push(axis.has_minor_gridlines == Some(true));
                        axis_minor_gridlines_inserted.push(axis.has_minor_gridlines == Some(true));
                        axis_minor_gridlines_removed.push(false);
                        axis_major_tick_mark_seen.push(axis.major_tick_mark.is_some());
                        axis_major_tick_mark_written.push(axis.major_tick_mark.is_some());
                        axis_major_tick_mark_inserted.push(axis.major_tick_mark.is_some());
                        axis_major_tick_mark_removed.push(false);
                        axis_minor_tick_mark_seen.push(axis.minor_tick_mark.is_some());
                        axis_minor_tick_mark_written.push(axis.minor_tick_mark.is_some());
                        axis_minor_tick_mark_inserted.push(axis.minor_tick_mark.is_some());
                        axis_minor_tick_mark_removed.push(false);
                        axis_tick_label_position_seen.push(axis.tick_label_position.is_some());
                        axis_tick_label_position_written.push(axis.tick_label_position.is_some());
                        axis_tick_label_position_inserted.push(axis.tick_label_position.is_some());
                        axis_tick_label_position_removed.push(false);
                        axis_tick_label_number_format_seen
                            .push(axis.tick_label_number_format.is_some());
                        axis_tick_label_number_format_written
                            .push(axis.tick_label_number_format.is_some());
                        axis_tick_label_number_format_inserted
                            .push(axis.tick_label_number_format.is_some());
                        axis_tick_label_number_format_removed.push(false);
                        axis_tick_label_spacing_seen.push(axis.tick_label_spacing.is_some());
                        axis_tick_label_spacing_written.push(axis.tick_label_spacing.is_some());
                        axis_tick_label_spacing_inserted.push(axis.tick_label_spacing.is_some());
                        axis_tick_label_spacing_removed.push(false);
                        axis_tick_mark_spacing_seen.push(axis.tick_mark_spacing.is_some());
                        axis_tick_mark_spacing_written.push(axis.tick_mark_spacing.is_some());
                        axis_tick_mark_spacing_inserted.push(axis.tick_mark_spacing.is_some());
                        axis_tick_mark_spacing_removed.push(false);
                        axis_cross_between_seen.push(axis.axis_between_categories.is_some());
                        axis_cross_between_written.push(axis.axis_between_categories.is_some());
                        axis_cross_between_inserted.push(axis.axis_between_categories.is_some());
                        axis_cross_between_removed.push(false);
                        axis_category_type_auto_seen.push(axis.category_type_auto.is_some());
                        axis_category_type_auto_written.push(axis.category_type_auto.is_some());
                        axis_category_type_auto_inserted.push(axis.category_type_auto.is_some());
                        axis_category_type_auto_removed.push(false);
                        axis_base_unit_seen.push(axis.base_unit.is_some());
                        axis_base_unit_written.push(axis.base_unit.is_some());
                        axis_base_unit_inserted.push(axis.base_unit.is_some());
                        axis_base_unit_removed.push(false);
                        axis_major_time_unit_seen.push(axis.major_unit_scale.is_some());
                        axis_major_time_unit_written.push(axis.major_unit_scale.is_some());
                        axis_major_time_unit_inserted.push(axis.major_unit_scale.is_some());
                        axis_major_time_unit_removed.push(false);
                        axis_minor_time_unit_seen.push(axis.minor_unit_scale.is_some());
                        axis_minor_time_unit_written.push(axis.minor_unit_scale.is_some());
                        axis_minor_time_unit_inserted.push(axis.minor_unit_scale.is_some());
                        axis_minor_time_unit_removed.push(false);
                        let has_log_base = chart_axis_log_base_xml_value(axis).is_some();
                        let has_scaling = chart_axis_has_scaling_xml(axis);
                        axis_scaling_seen.push(has_scaling);
                        axis_log_base_seen.push(has_log_base);
                        axis_log_base_written.push(has_log_base);
                        axis_log_base_inserted.push(has_log_base);
                        axis_log_base_removed.push(false);
                        axis_orientation_seen.push(axis.reverse_plot_order.is_some());
                        axis_orientation_written.push(axis.reverse_plot_order.is_some());
                        axis_orientation_inserted.push(axis.reverse_plot_order.is_some());
                        axis_orientation_removed.push(false);
                        axis_minimum_scale_seen.push(axis.minimum_scale.is_some());
                        axis_minimum_scale_written.push(axis.minimum_scale.is_some());
                        axis_minimum_scale_inserted.push(axis.minimum_scale.is_some());
                        axis_minimum_scale_removed.push(false);
                        axis_maximum_scale_seen.push(axis.maximum_scale.is_some());
                        axis_maximum_scale_written.push(axis.maximum_scale.is_some());
                        axis_maximum_scale_inserted.push(axis.maximum_scale.is_some());
                        axis_maximum_scale_removed.push(false);
                        axis_major_unit_seen.push(axis.major_unit.is_some());
                        axis_major_unit_written.push(axis.major_unit.is_some());
                        axis_major_unit_inserted.push(axis.major_unit.is_some());
                        axis_major_unit_removed.push(false);
                        axis_minor_unit_seen.push(axis.minor_unit.is_some());
                        axis_minor_unit_written.push(axis.minor_unit.is_some());
                        axis_minor_unit_inserted.push(axis.minor_unit.is_some());
                        axis_minor_unit_removed.push(false);
                        axis_display_units_seen.push(axis.display_unit.is_some());
                        axis_display_units_written.push(axis.display_unit.is_some());
                        axis_display_units_inserted.push(axis.display_unit.is_some());
                        axis_display_units_removed.push(false);
                        let has_crosses = axis.crosses_at.is_none()
                            && axis
                                .crosses
                                .is_some_and(|value| chart_axis_crosses_xml_value(value).is_some());
                        axis_crosses_seen.push(has_crosses);
                        axis_crosses_written.push(has_crosses);
                        axis_crosses_inserted.push(has_crosses);
                        axis_crosses_removed.push(false);
                        axis_crosses_at_seen.push(axis.crosses_at.is_some());
                        axis_crosses_at_written.push(axis.crosses_at.is_some());
                        axis_crosses_at_inserted.push(axis.crosses_at.is_some());
                        axis_crosses_at_removed.push(false);
                    }
                }
                if chart_type_from_group_name(local_name.as_slice()).is_some()
                    && current_chart_group_depth == Some(element_stack.len())
                    && !preserve_loaded_group_types
                {
                    for (series_index, series) in chart.series.iter().enumerate() {
                        if series_emitted.get(series_index).copied().unwrap_or(false) {
                            continue;
                        }
                        let mut source_seen = [false; 4];
                        writer
                            .write_event(Event::Start(BytesStart::new("c:ser")))
                            .map_err(chart_xml_error)?;
                        let series_index_text =
                            series.raw_index.unwrap_or(series_index as u32).to_string();
                        let mut idx_element = BytesStart::new("c:idx");
                        idx_element.push_attribute(("val", series_index_text.as_str()));
                        writer
                            .write_event(Event::Empty(idx_element))
                            .map_err(chart_xml_error)?;
                        let order_text = series.order.unwrap_or(series_index as u32).to_string();
                        let mut order_element = BytesStart::new("c:order");
                        order_element.push_attribute(("val", order_text.as_str()));
                        writer
                            .write_event(Event::Empty(order_element))
                            .map_err(chart_xml_error)?;
                        if let Some(value) = expected_series_explosions
                            .get(series_index)
                            .and_then(Option::as_deref)
                        {
                            let mut explosion = BytesStart::new("c:explosion");
                            explosion.push_attribute(("val", value));
                            writer
                                .write_event(Event::Empty(explosion))
                                .map_err(chart_xml_error)?;
                            if let Some(seen) = series_explosion_seen.get_mut(series_index) {
                                *seen = true;
                            }
                            if let Some(inserted) = series_explosion_inserted.get_mut(series_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = series_explosion_written.get_mut(series_index) {
                                *written = true;
                            }
                        }
                        if let Some(expected_points) =
                            expected_dirty_point_explosions.get(series_index)
                        {
                            for (point_index, explosion) in expected_points {
                                write_chart_point_explosion_element(
                                    &mut writer,
                                    *point_index,
                                    *explosion,
                                )?;
                                if let Some(inserted) =
                                    series_point_explosions_inserted.get_mut(series_index)
                                {
                                    inserted.insert(*point_index);
                                }
                            }
                        }
                        if expected_dirty_series_data_label_sets
                            .get(series_index)
                            .copied()
                            .unwrap_or(false)
                        {
                            write_chart_data_labels_element(
                                &mut writer,
                                series.data_labels.as_ref(),
                                Some(&series.point_data_labels),
                            )?;
                            if let Some(seen) = series_data_labels_seen.get_mut(series_index) {
                                *seen = true;
                            }
                            if let Some(inserted) =
                                series_data_labels_inserted.get_mut(series_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = series_data_labels_written.get_mut(series_index)
                            {
                                *written = true;
                            }
                        }
                        let marker_style = expected_series_marker_styles
                            .get(series_index)
                            .copied()
                            .flatten();
                        let marker_size = expected_series_marker_sizes
                            .get(series_index)
                            .and_then(Option::as_deref);
                        if marker_style.is_some() || marker_size.is_some() {
                            write_chart_series_marker_element(
                                &mut writer,
                                marker_style,
                                marker_size,
                            )?;
                            if let Some(seen) = series_marker_seen.get_mut(series_index) {
                                *seen = true;
                            }
                            if let Some(inserted) = series_marker_inserted.get_mut(series_index) {
                                *inserted = true;
                            }
                            if marker_style.is_some() {
                                if let Some(inserted) =
                                    series_marker_style_inserted.get_mut(series_index)
                                {
                                    *inserted = true;
                                }
                                if let Some(written) =
                                    series_marker_style_written.get_mut(series_index)
                                {
                                    *written = true;
                                }
                            }
                            if marker_size.is_some() {
                                if let Some(inserted) =
                                    series_marker_size_inserted.get_mut(series_index)
                                {
                                    *inserted = true;
                                }
                                if let Some(written) =
                                    series_marker_size_written.get_mut(series_index)
                                {
                                    *written = true;
                                }
                            }
                        }
                        if let Some(value) = expected_series_invert_if_negative_values
                            .get(series_index)
                            .copied()
                            .flatten()
                        {
                            let mut invert_if_negative = BytesStart::new("c:invertIfNegative");
                            invert_if_negative.push_attribute(("val", value));
                            writer
                                .write_event(Event::Empty(invert_if_negative))
                                .map_err(chart_xml_error)?;
                            if let Some(seen) = series_invert_if_negative_seen.get_mut(series_index)
                            {
                                *seen = true;
                            }
                            if let Some(inserted) =
                                series_invert_if_negative_inserted.get_mut(series_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) =
                                series_invert_if_negative_written.get_mut(series_index)
                            {
                                *written = true;
                            }
                        }
                        for slot in source_slots_in_order.iter().copied() {
                            if let Some(source) = source_for_slot(series_index, slot) {
                                write_chart_source_container(
                                    &mut writer,
                                    series_index,
                                    slot,
                                    source,
                                )?;
                                source_seen[slot_index(slot)] = true;
                                if source.dirty {
                                    patched_sources += 1;
                                }
                            }
                        }
                        if let Some(value) = expected_series_bar_shapes
                            .get(series_index)
                            .copied()
                            .flatten()
                        {
                            let mut shape = BytesStart::new("c:shape");
                            shape.push_attribute(("val", value));
                            writer
                                .write_event(Event::Empty(shape))
                                .map_err(chart_xml_error)?;
                            if let Some(seen) = series_bar_shape_seen.get_mut(series_index) {
                                *seen = true;
                            }
                            if let Some(inserted) = series_bar_shape_inserted.get_mut(series_index)
                            {
                                *inserted = true;
                            }
                            if let Some(written) = series_bar_shape_written.get_mut(series_index) {
                                *written = true;
                            }
                        }
                        if let Some(value) = expected_series_smooth_values
                            .get(series_index)
                            .copied()
                            .flatten()
                        {
                            let mut smooth = BytesStart::new("c:smooth");
                            smooth.push_attribute(("val", value));
                            writer
                                .write_event(Event::Empty(smooth))
                                .map_err(chart_xml_error)?;
                            if let Some(seen) = series_smooth_seen.get_mut(series_index) {
                                *seen = true;
                            }
                            if let Some(inserted) = series_smooth_inserted.get_mut(series_index) {
                                *inserted = true;
                            }
                            if let Some(written) = series_smooth_written.get_mut(series_index) {
                                *written = true;
                            }
                        }
                        writer
                            .write_event(Event::End(BytesEnd::new("c:ser")))
                            .map_err(chart_xml_error)?;
                        if let Some(seen) = source_slots_seen.get_mut(series_index) {
                            *seen = source_seen;
                        }
                        if let Some(seen) = series_order_seen.get_mut(series_index) {
                            *seen = true;
                        }
                        if let Some(written) = series_order_written.get_mut(series_index) {
                            *written = true;
                        }
                        if let Some(emitted) = series_emitted.get_mut(series_index) {
                            *emitted = true;
                        }
                    }
                    if !bar_direction_seen && let Some(value) = expected_bar_direction {
                        let mut bar_direction = BytesStart::new("c:barDir");
                        bar_direction.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(bar_direction))
                            .map_err(chart_xml_error)?;
                        bar_direction_inserted = true;
                        bar_direction_written = true;
                    }
                    if !chart_grouping_seen && let Some(value) = expected_chart_grouping {
                        let mut chart_grouping = BytesStart::new("c:grouping");
                        chart_grouping.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(chart_grouping))
                            .map_err(chart_xml_error)?;
                        chart_grouping_inserted = true;
                        chart_grouping_written = true;
                    }
                    if !bar_shape_seen && let Some(value) = expected_bar_shape {
                        let mut bar_shape = BytesStart::new("c:shape");
                        bar_shape.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(bar_shape))
                            .map_err(chart_xml_error)?;
                        bar_shape_inserted = true;
                        bar_shape_written = true;
                    }
                    if !line_marker_seen && let Some(value) = expected_line_marker {
                        let mut line_marker = BytesStart::new("c:marker");
                        line_marker.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(line_marker))
                            .map_err(chart_xml_error)?;
                        line_marker_inserted = true;
                        line_marker_written = true;
                    }
                    if !scatter_style_seen && let Some(value) = expected_scatter_style {
                        let mut scatter_style = BytesStart::new("c:scatterStyle");
                        scatter_style.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(scatter_style))
                            .map_err(chart_xml_error)?;
                        scatter_style_inserted = true;
                        scatter_style_written = true;
                    }
                    if !radar_style_seen && let Some(value) = expected_radar_style {
                        let mut radar_style = BytesStart::new("c:radarStyle");
                        radar_style.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(radar_style))
                            .map_err(chart_xml_error)?;
                        radar_style_inserted = true;
                        radar_style_written = true;
                    }
                    if !of_pie_type_seen && let Some(value) = expected_of_pie_type {
                        let mut of_pie_type = BytesStart::new("c:ofPieType");
                        of_pie_type.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(of_pie_type))
                            .map_err(chart_xml_error)?;
                        of_pie_type_inserted = true;
                        of_pie_type_written = true;
                    }
                    if !surface_wireframe_seen && let Some(value) = expected_surface_wireframe {
                        let mut surface_wireframe = BytesStart::new("c:wireframe");
                        surface_wireframe.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(surface_wireframe))
                            .map_err(chart_xml_error)?;
                        surface_wireframe_inserted = true;
                        surface_wireframe_written = true;
                    }
                    if !gap_width_seen && let Some(value) = expected_gap_width {
                        let mut gap_width = BytesStart::new("c:gapWidth");
                        gap_width.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(gap_width))
                            .map_err(chart_xml_error)?;
                        gap_width_inserted = true;
                        gap_width_written = true;
                    }
                    if !gap_depth_seen && let Some(value) = expected_gap_depth {
                        let mut gap_depth = BytesStart::new("c:gapDepth");
                        gap_depth.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(gap_depth))
                            .map_err(chart_xml_error)?;
                        gap_depth_inserted = true;
                        gap_depth_written = true;
                    }
                    if !overlap_seen && let Some(value) = expected_overlap {
                        let mut overlap = BytesStart::new("c:overlap");
                        overlap.push_attribute(("val", value));
                        writer
                            .write_event(Event::Empty(overlap))
                            .map_err(chart_xml_error)?;
                        overlap_inserted = true;
                        overlap_written = true;
                    }
                    for (setting_index, (_, qualified_name, expected)) in
                        expected_chart_group_numeric_settings.iter().enumerate()
                    {
                        if !chart_group_numeric_setting_seen[setting_index]
                            && let Some(value) = expected
                        {
                            let mut setting = BytesStart::new(*qualified_name);
                            setting.push_attribute(("val", *value));
                            writer
                                .write_event(Event::Empty(setting))
                                .map_err(chart_xml_error)?;
                            chart_group_numeric_setting_inserted[setting_index] = true;
                            chart_group_numeric_setting_written[setting_index] = true;
                        }
                    }
                    for (flag_index, (_, qualified_name, expected)) in
                        expected_chart_group_line_flags.iter().enumerate()
                    {
                        if !chart_group_line_flag_seen[flag_index] && *expected == Some(true) {
                            writer
                                .write_event(Event::Empty(BytesStart::new(*qualified_name)))
                                .map_err(chart_xml_error)?;
                            chart_group_line_flag_inserted[flag_index] = true;
                            chart_group_line_flag_written[flag_index] = true;
                        }
                    }
                    if !data_labels_seen && let Some(data_labels) = expected_dirty_data_labels {
                        write_chart_data_labels_element(&mut writer, Some(data_labels), None)?;
                        data_labels_inserted = true;
                        data_labels_written = true;
                    }
                    if !chart_group_axis_refs_seen.is_empty() {
                        for (axis_index, axis) in chart.axes.iter().enumerate() {
                            let axis_id = chart_axis_id(axis_index, axis);
                            if !chart_group_axis_refs_seen
                                .iter()
                                .any(|seen_axis_id| seen_axis_id == &axis_id)
                            {
                                write_chart_axis_ref_element(&mut writer, &axis_id)?;
                                chart_group_axis_refs_seen.push(axis_id);
                            }
                        }
                    }
                    current_chart_group_depth = None;
                }
                if chart_type_from_group_name(local_name.as_slice()).is_some()
                    && current_chart_group_depth == Some(element_stack.len())
                {
                    current_chart_group_depth = None;
                    chart_group_axis_refs_seen.clear();
                }

                if local_name.as_slice() == b"legend"
                    && expected_legend_position.is_some()
                    && !legend_overlay_seen
                    && let Some(value) = expected_legend_include_in_layout
                {
                    let mut overlay = BytesStart::new("c:overlay");
                    overlay.push_attribute(("val", value));
                    writer
                        .write_event(Event::Empty(overlay))
                        .map_err(chart_xml_error)?;
                    legend_overlay_inserted = true;
                    legend_overlay_written = true;
                }

                if local_name.as_slice() == b"layout"
                    && parent_name == Some(b"plotArea".as_slice())
                    && !plot_area_manual_layout_seen
                    && !plot_area_manual_layout_inserted
                    && let Some(Some(layout)) = expected_plot_area_layout
                {
                    writer
                        .get_mut()
                        .write_all(chart_manual_layout_xml_string(layout).as_bytes())
                        .map_err(chart_xml_error)?;
                    plot_area_manual_layout_inserted = true;
                    plot_area_manual_layout_written = true;
                }

                if local_name.as_slice() == b"plotArea"
                    && !plot_area_layout_container_seen
                    && let Some(Some(layout)) = expected_plot_area_layout
                {
                    writer
                        .get_mut()
                        .write_all(
                            format!(
                                "<c:layout>{}</c:layout>",
                                chart_manual_layout_xml_string(layout)
                            )
                            .as_bytes(),
                        )
                        .map_err(chart_xml_error)?;
                    plot_area_layout_container_seen = true;
                    plot_area_manual_layout_inserted = true;
                    plot_area_manual_layout_written = true;
                }

                if local_name.as_slice() == b"plotArea"
                    && !data_table_seen
                    && let Some(Some(data_table)) = expected_data_table
                {
                    write_chart_data_table_element(&mut writer, data_table)?;
                    data_table_inserted = true;
                    data_table_written = true;
                }

                if local_name.as_slice() == b"dPt"
                    && let Some(series_index) = current_series_index
                    && let Some(point_index) = current_point_index
                    && let Some(explosion) = expected_point_explosion(series_index, point_index)
                {
                    let already_inserted = series_point_explosions_inserted
                        .get(series_index)
                        .is_some_and(|inserted| inserted.contains(&point_index));
                    if !already_inserted {
                        write_chart_u32_val_element(
                            &mut writer,
                            "c:explosion",
                            u32::from(explosion),
                        )?;
                        if let Some(inserted) =
                            series_point_explosions_inserted.get_mut(series_index)
                        {
                            inserted.insert(point_index);
                        }
                    }
                }
                if local_name.as_slice() == b"ser"
                    && let Some(series_index) = current_series_index
                    && !series_bar_shape_seen
                        .get(series_index)
                        .copied()
                        .unwrap_or(false)
                    && let Some(value) = expected_series_bar_shapes
                        .get(series_index)
                        .copied()
                        .flatten()
                {
                    let mut shape = BytesStart::new("c:shape");
                    shape.push_attribute(("val", value));
                    writer
                        .write_event(Event::Empty(shape))
                        .map_err(chart_xml_error)?;
                    if let Some(inserted) = series_bar_shape_inserted.get_mut(series_index) {
                        *inserted = true;
                    }
                    if let Some(written) = series_bar_shape_written.get_mut(series_index) {
                        *written = true;
                    }
                }
                if local_name.as_slice() == b"ser"
                    && let Some(series_index) = current_series_index
                    && !series_smooth_seen
                        .get(series_index)
                        .copied()
                        .unwrap_or(false)
                    && let Some(value) = expected_series_smooth_values
                        .get(series_index)
                        .copied()
                        .flatten()
                {
                    let mut smooth = BytesStart::new("c:smooth");
                    smooth.push_attribute(("val", value));
                    writer
                        .write_event(Event::Empty(smooth))
                        .map_err(chart_xml_error)?;
                    if let Some(inserted) = series_smooth_inserted.get_mut(series_index) {
                        *inserted = true;
                    }
                    if let Some(written) = series_smooth_written.get_mut(series_index) {
                        *written = true;
                    }
                }

                if chart_type_from_group_name(local_name.as_slice()).is_some()
                    && !preserve_loaded_group_types
                    && chart_type.as_ref() != Some(&chart.chart_type)
                    && let Some(target_local_name) = target_chart_group_name
                {
                    writer
                        .write_event(Event::End(BytesEnd::new(qualified_replacement_name(
                            element.name().as_ref(),
                            target_local_name,
                        ))))
                        .map_err(chart_xml_error)?;
                } else if let Some(slot) = source_container_slot(local_name.as_slice())
                    && current_series_index.is_some()
                {
                    let target_local_name = source_container_target_local_name(
                        current_series_index.expect("series source container"),
                        slot,
                    );
                    if target_local_name.as_bytes() != local_name.as_slice() {
                        writer
                            .write_event(Event::End(BytesEnd::new(qualified_replacement_name(
                                element.name().as_ref(),
                                target_local_name,
                            ))))
                            .map_err(chart_xml_error)?;
                    } else {
                        writer
                            .write_event(Event::End(element.to_owned()))
                            .map_err(chart_xml_error)?;
                    }
                } else {
                    writer
                        .write_event(Event::End(element.to_owned()))
                        .map_err(chart_xml_error)?;
                }

                if current_series_index.is_some() {
                    match local_name.as_slice() {
                        b"tx" | b"cat" | b"val" | b"xVal" | b"yVal" | b"bubbleSize" => {
                            source_stack.pop();
                        }
                        b"ser" => {
                            current_series_index = None;
                            current_point_index = None;
                            current_series_marker_index = None;
                            source_stack.clear();
                        }
                        b"marker" => {
                            current_series_marker_index = None;
                        }
                        b"dPt" => {
                            current_point_index = None;
                        }
                        _ => {}
                    }
                }
                if local_name.as_slice() == b"title"
                    && let Some((title_depth, _)) = title_stack.last()
                    && *title_depth == element_stack.len()
                {
                    title_stack.pop();
                }
                if current_axis_depth == Some(element_stack.len()) {
                    current_axis_index = None;
                    current_axis_depth = None;
                }
                element_stack.pop();
            }
            Ok(Event::Eof) => break,
            Ok(event) => writer
                .write_event(event.into_owned())
                .map_err(chart_xml_error)?,
            Err(error) => return Err(chart_xml_error(error)),
        }
        buffer.clear();
    }

    let chart_type_matches = preserve_loaded_group_types
        || match chart_type.as_ref() {
            Some(source_chart_type) if source_chart_type == &chart.chart_type => true,
            Some(_) => chart_type_rewritten,
            None => match &chart.chart_type {
                ChartType::Unsupported(raw_name) => {
                    !chart.series_topology_dirty
                        && chart.groups.len() == 1
                        && chart.groups[0].loaded_index == Some(0)
                        && chart.groups[0].raw_name == *raw_name
                }
                _ => false,
            },
        };
    let series_sources_match = source_slots_seen.len() == chart.series.len()
        && source_slots_seen
            .iter()
            .zip(chart.series.iter())
            .enumerate()
            .all(|(series_index, (seen, series))| {
                seen[0] == series.name.is_some()
                    && seen[1] == series.x_values.is_some()
                    && seen[2] == series.values.is_some()
                    && seen[3]
                        == (series.bubble_size.is_some()
                            && chart_type_uses_bubble_size(&model_series_chart_types[series_index]))
            });
    let series_orders_match = series_order_seen.len() == chart.series.len()
        && series_order_seen
            .iter()
            .zip(series_order_written.iter())
            .zip(chart.series.iter())
            .enumerate()
            .all(|(series_index, ((seen, written), series))| {
                let expected_order = series.order.unwrap_or(series_index as u32);
                if *seen {
                    *written
                } else {
                    expected_order == series_index as u32
                }
            });
    let series_explosions_match = series_explosion_seen.len() == chart.series.len()
        && series_explosion_seen
            .iter()
            .zip(series_explosion_written.iter())
            .zip(series_explosion_inserted.iter())
            .zip(series_explosion_removed.iter())
            .zip(expected_series_explosions.iter())
            .all(
                |((((seen, written), inserted), removed), expected)| match expected {
                    Some(_) if *seen => *written,
                    Some(_) => *inserted,
                    None if *seen => *removed,
                    None => true,
                },
            );
    let series_point_explosions_match = expected_dirty_point_explosions.iter().enumerate().all(
        |(series_index, expected_points)| {
            expected_points.iter().all(|(point_index, _)| {
                series_point_explosions_inserted
                    .get(series_index)
                    .is_some_and(|inserted| inserted.contains(point_index))
            })
        },
    );
    let series_bar_shapes_match =
        expected_series_bar_shapes
            .iter()
            .enumerate()
            .all(|(series_index, expected)| {
                match (
                    expected,
                    series_bar_shape_seen
                        .get(series_index)
                        .copied()
                        .unwrap_or(false),
                ) {
                    (Some(_), true) => series_bar_shape_written
                        .get(series_index)
                        .copied()
                        .unwrap_or(false),
                    (Some(_), false) => series_bar_shape_inserted
                        .get(series_index)
                        .copied()
                        .unwrap_or(false),
                    (None, true) => series_bar_shape_removed
                        .get(series_index)
                        .copied()
                        .unwrap_or(false),
                    (None, false) => true,
                }
            });
    let series_smooth_values_match =
        expected_series_smooth_values
            .iter()
            .enumerate()
            .all(|(series_index, expected)| {
                match (
                    expected,
                    series_smooth_seen
                        .get(series_index)
                        .copied()
                        .unwrap_or(false),
                ) {
                    (Some(_), true) => series_smooth_written
                        .get(series_index)
                        .copied()
                        .unwrap_or(false),
                    (Some(_), false) => series_smooth_inserted
                        .get(series_index)
                        .copied()
                        .unwrap_or(false),
                    (None, true) => series_smooth_removed
                        .get(series_index)
                        .copied()
                        .unwrap_or(false),
                    (None, false) => true,
                }
            });
    let series_markers_match = chart.series.iter().enumerate().all(|(series_index, _)| {
        let marker_seen = series_marker_seen
            .get(series_index)
            .copied()
            .unwrap_or(false);
        if !chart_type_supports_series_marker(&model_series_chart_types[series_index]) {
            return !marker_seen
                || series_marker_removed
                    .get(series_index)
                    .copied()
                    .unwrap_or(false);
        }
        let marker_inserted = series_marker_inserted
            .get(series_index)
            .copied()
            .unwrap_or(false);
        let marker_available = marker_seen || marker_inserted;
        let style_matches = if expected_series_marker_styles
            .get(series_index)
            .copied()
            .flatten()
            .is_some()
        {
            marker_available
                && (series_marker_style_written
                    .get(series_index)
                    .copied()
                    .unwrap_or(false)
                    || series_marker_style_inserted
                        .get(series_index)
                        .copied()
                        .unwrap_or(false))
        } else {
            true
        };
        let size_matches = if expected_series_marker_sizes
            .get(series_index)
            .and_then(Option::as_ref)
            .is_some()
        {
            marker_available
                && (series_marker_size_written
                    .get(series_index)
                    .copied()
                    .unwrap_or(false)
                    || series_marker_size_inserted
                        .get(series_index)
                        .copied()
                        .unwrap_or(false))
        } else {
            true
        };
        style_matches && size_matches
    });
    let series_invert_if_negative_values_match = expected_series_invert_if_negative_values
        .iter()
        .enumerate()
        .all(|(series_index, expected)| {
            if expected.is_some() {
                series_invert_if_negative_written
                    .get(series_index)
                    .copied()
                    .unwrap_or(false)
                    || series_invert_if_negative_inserted
                        .get(series_index)
                        .copied()
                        .unwrap_or(false)
            } else {
                true
            }
        });
    let title_matches = match (chart.title.as_ref(), chart_title_seen) {
        (Some(_), true) => chart_title_text_written,
        (Some(_), false) => chart_title_inserted,
        (None, true) => chart_title_removed,
        (None, false) => true,
    };
    let legend_matches = match (expected_legend_position, legend_seen) {
        (Some(_), true) => legend_position_written,
        (Some(_), false) => legend_inserted,
        (None, true) => legend_removed,
        (None, false) => true,
    };
    let legend_overlay_matches = match (
        expected_legend_include_in_layout,
        legend_seen,
        legend_overlay_seen,
    ) {
        (Some(_), true, true) => legend_overlay_written,
        (Some(_), true, false) => legend_overlay_inserted,
        (Some(_), false, _) => legend_overlay_inserted,
        (None, _, _) => true,
    };
    let display_blanks_as_matches = match (expected_display_blanks_as, display_blanks_as_seen) {
        (Some(_), true) => display_blanks_as_written,
        (Some(_), false) => display_blanks_as_inserted,
        (None, _) => true,
    };
    let plot_visible_only_matches = match (expected_plot_visible_only, plot_visible_only_seen) {
        (Some(_), true) => plot_visible_only_written,
        (Some(_), false) => plot_visible_only_inserted,
        (None, _) => true,
    };
    let show_data_labels_over_maximum_matches = match (
        expected_show_data_labels_over_maximum,
        show_data_labels_over_maximum_seen,
    ) {
        (Some(_), true) => show_data_labels_over_maximum_written,
        (Some(_), false) => show_data_labels_over_maximum_inserted,
        (None, _) => true,
    };
    let view_3d_matches = match expected_dirty_view_3d {
        Some(Some(_)) if view_3d_seen => view_3d_written,
        Some(Some(_)) => view_3d_inserted,
        Some(None) | None => true,
    };
    let rounded_corners_matches = match (expected_rounded_corners, rounded_corners_seen) {
        (Some(_), true) => rounded_corners_written,
        (Some(_), false) => rounded_corners_inserted,
        (None, _) => true,
    };
    let chart_style_matches = match (expected_chart_style, chart_style_seen) {
        (Some(_), true) => chart_style_written,
        (Some(_), false) => chart_style_inserted,
        (None, _) => true,
    };
    let chart_protection_matches = match expected_chart_protection {
        None => true,
        Some(_) if chart_protection_seen => chart_protection_written || chart_protection_removed,
        Some(_) => !expected_chart_protection_needs_xml || chart_protection_inserted,
    };
    let data_table_matches = match expected_data_table {
        Some(Some(_)) if data_table_seen => data_table_written,
        Some(Some(_)) => data_table_inserted,
        Some(None) if data_table_seen => data_table_removed,
        Some(None) => true,
        None => true,
    };
    let plot_area_layout_matches = match expected_plot_area_layout {
        Some(Some(_)) if plot_area_manual_layout_seen => plot_area_manual_layout_written,
        Some(Some(_)) => plot_area_manual_layout_inserted,
        Some(None) if plot_area_manual_layout_seen => plot_area_manual_layout_removed,
        Some(None) => true,
        None => true,
    };
    let vary_colors_matches = match (expected_vary_colors, vary_colors_seen) {
        (Some(_), true) => vary_colors_written,
        (Some(_), false) => vary_colors_inserted,
        (None, _) => true,
    };
    let bar_direction_matches = match (expected_bar_direction, bar_direction_seen) {
        (Some(_), true) => bar_direction_written,
        (Some(_), false) => bar_direction_inserted,
        (None, _) => true,
    };
    let chart_grouping_matches = match (expected_chart_grouping, chart_grouping_seen) {
        (Some(_), true) => chart_grouping_written,
        (Some(_), false) => chart_grouping_inserted,
        (None, _) => true,
    };
    let bar_shape_matches = match (expected_bar_shape, bar_shape_seen) {
        (Some(_), true) => bar_shape_written,
        (Some(_), false) => bar_shape_inserted,
        (None, _) => true,
    };
    let line_marker_matches = match (expected_line_marker, line_marker_seen) {
        (Some(_), true) => line_marker_written,
        (Some(_), false) => line_marker_inserted,
        (None, _) => true,
    };
    let scatter_style_matches = match (expected_scatter_style, scatter_style_seen) {
        (Some(_), true) => scatter_style_written,
        (Some(_), false) => scatter_style_inserted,
        (None, _) => true,
    };
    let radar_style_matches = match (expected_radar_style, radar_style_seen) {
        (Some(_), true) => radar_style_written,
        (Some(_), false) => radar_style_inserted,
        (None, _) => true,
    };
    let of_pie_type_matches = match (expected_of_pie_type, of_pie_type_seen) {
        (Some(_), true) => of_pie_type_written,
        (Some(_), false) => of_pie_type_inserted,
        (None, _) => true,
    };
    let surface_wireframe_matches = match (expected_surface_wireframe, surface_wireframe_seen) {
        (Some(_), true) => surface_wireframe_written,
        (Some(_), false) => surface_wireframe_inserted,
        (None, _) => true,
    };
    let gap_width_matches = match (expected_gap_width, gap_width_seen) {
        (Some(_), true) => gap_width_written,
        (Some(_), false) => gap_width_inserted,
        (None, _) => true,
    };
    let gap_depth_matches = match (expected_gap_depth, gap_depth_seen) {
        (Some(_), true) => gap_depth_written,
        (Some(_), false) => gap_depth_inserted,
        (None, _) => true,
    };
    let overlap_matches = match (expected_overlap, overlap_seen) {
        (Some(_), true) => overlap_written,
        (Some(_), false) => overlap_inserted,
        (None, _) => true,
    };
    let data_labels_match = match (expected_dirty_data_labels, data_labels_seen) {
        (Some(_), true) => data_labels_written,
        (Some(_), false) => data_labels_inserted,
        (None, _) => true,
    };
    let series_data_labels_match = expected_dirty_series_data_label_sets
        .iter()
        .enumerate()
        .all(|(series_index, expected)| {
            match (
                *expected,
                series_data_labels_seen
                    .get(series_index)
                    .copied()
                    .unwrap_or(false),
            ) {
                (true, true) => series_data_labels_written
                    .get(series_index)
                    .copied()
                    .unwrap_or(false),
                (true, false) => series_data_labels_inserted
                    .get(series_index)
                    .copied()
                    .unwrap_or(false),
                (false, _) => true,
            }
        });
    let chart_group_numeric_settings_match = expected_chart_group_numeric_settings
        .iter()
        .enumerate()
        .all(|(setting_index, (_, _, expected))| {
            match (expected, chart_group_numeric_setting_seen[setting_index]) {
                (Some(_), true) => chart_group_numeric_setting_written[setting_index],
                (Some(_), false) => chart_group_numeric_setting_inserted[setting_index],
                (None, _) => true,
            }
        });
    let chart_group_line_flags_match =
        expected_chart_group_line_flags
            .iter()
            .enumerate()
            .all(|(flag_index, (_, _, expected))| {
                match (*expected, chart_group_line_flag_seen[flag_index]) {
                    (Some(true), true) => chart_group_line_flag_written[flag_index],
                    (Some(true), false) => chart_group_line_flag_inserted[flag_index],
                    (Some(false), true) => chart_group_line_flag_removed[flag_index],
                    (Some(false), false) => true,
                    (None, _) => true,
                }
            });
    let axes_match = axis_kinds.len() == chart.axes.len()
        && axis_kinds
            .iter()
            .zip(axis_title_texts.iter())
            .zip(axis_title_text_written.iter())
            .zip(chart.axes.iter())
            .all(|(((kind, title_text), title_written), axis)| {
                *kind == axis.kind
                    && match axis.title.as_ref() {
                        Some(title) => *title_written || title_text.as_ref() == Some(&title.text),
                        None => title_text.is_none(),
                    }
            });
    let axis_numeric_field_matches = |axis_index: usize,
                                      expected: Option<f64>,
                                      seen: &[bool],
                                      written: &[bool],
                                      inserted: &[bool],
                                      removed: &[bool]|
     -> bool {
        match expected {
            Some(_) => {
                written.get(axis_index).copied().unwrap_or(false)
                    || inserted.get(axis_index).copied().unwrap_or(false)
            }
            None => {
                !seen.get(axis_index).copied().unwrap_or(false)
                    || removed.get(axis_index).copied().unwrap_or(false)
            }
        }
    };
    let axis_optional_field_matches = |axis_index: usize,
                                       expected: bool,
                                       seen: &[bool],
                                       written: &[bool],
                                       inserted: &[bool],
                                       removed: &[bool]|
     -> bool {
        if expected {
            written.get(axis_index).copied().unwrap_or(false)
                || inserted.get(axis_index).copied().unwrap_or(false)
        } else {
            !seen.get(axis_index).copied().unwrap_or(false)
                || removed.get(axis_index).copied().unwrap_or(false)
        }
    };
    let axis_scale_units_match = chart.axes.iter().enumerate().all(|(axis_index, axis)| {
        axis_numeric_field_matches(
            axis_index,
            chart_axis_log_base_xml_value(axis),
            &axis_log_base_seen,
            &axis_log_base_written,
            &axis_log_base_inserted,
            &axis_log_base_removed,
        ) && axis_numeric_field_matches(
            axis_index,
            axis.minimum_scale,
            &axis_minimum_scale_seen,
            &axis_minimum_scale_written,
            &axis_minimum_scale_inserted,
            &axis_minimum_scale_removed,
        ) && axis_numeric_field_matches(
            axis_index,
            axis.maximum_scale,
            &axis_maximum_scale_seen,
            &axis_maximum_scale_written,
            &axis_maximum_scale_inserted,
            &axis_maximum_scale_removed,
        ) && axis_numeric_field_matches(
            axis_index,
            axis.major_unit,
            &axis_major_unit_seen,
            &axis_major_unit_written,
            &axis_major_unit_inserted,
            &axis_major_unit_removed,
        ) && axis_numeric_field_matches(
            axis_index,
            axis.minor_unit,
            &axis_minor_unit_seen,
            &axis_minor_unit_written,
            &axis_minor_unit_inserted,
            &axis_minor_unit_removed,
        ) && axis_optional_field_matches(
            axis_index,
            axis.display_unit.is_some(),
            &axis_display_units_seen,
            &axis_display_units_written,
            &axis_display_units_inserted,
            &axis_display_units_removed,
        ) && axis_optional_field_matches(
            axis_index,
            axis.base_unit.is_some(),
            &axis_base_unit_seen,
            &axis_base_unit_written,
            &axis_base_unit_inserted,
            &axis_base_unit_removed,
        ) && axis_optional_field_matches(
            axis_index,
            axis.major_unit_scale.is_some(),
            &axis_major_time_unit_seen,
            &axis_major_time_unit_written,
            &axis_major_time_unit_inserted,
            &axis_major_time_unit_removed,
        ) && axis_optional_field_matches(
            axis_index,
            axis.minor_unit_scale.is_some(),
            &axis_minor_time_unit_seen,
            &axis_minor_time_unit_written,
            &axis_minor_time_unit_inserted,
            &axis_minor_time_unit_removed,
        )
    });
    let axis_orientation_match = chart.axes.iter().enumerate().all(|(axis_index, axis)| {
        if axis.reverse_plot_order.is_some() {
            axis_orientation_written
                .get(axis_index)
                .copied()
                .unwrap_or(false)
                || axis_orientation_inserted
                    .get(axis_index)
                    .copied()
                    .unwrap_or(false)
        } else {
            true
        }
    });
    let axis_crossing_match = chart.axes.iter().enumerate().all(|(axis_index, axis)| {
        axis_numeric_field_matches(
            axis_index,
            axis.crosses_at,
            &axis_crosses_at_seen,
            &axis_crosses_at_written,
            &axis_crosses_at_inserted,
            &axis_crosses_at_removed,
        ) && axis_optional_field_matches(
            axis_index,
            axis.crosses_at.is_none()
                && axis
                    .crosses
                    .is_some_and(|value| chart_axis_crosses_xml_value(value).is_some()),
            &axis_crosses_seen,
            &axis_crosses_written,
            &axis_crosses_inserted,
            &axis_crosses_removed,
        ) && axis_optional_field_matches(
            axis_index,
            axis.axis_between_categories.is_some(),
            &axis_cross_between_seen,
            &axis_cross_between_written,
            &axis_cross_between_inserted,
            &axis_cross_between_removed,
        )
    });
    let axis_category_type_match = chart.axes.iter().enumerate().all(|(axis_index, axis)| {
        axis_optional_field_matches(
            axis_index,
            axis.category_type_auto.is_some(),
            &axis_category_type_auto_seen,
            &axis_category_type_auto_written,
            &axis_category_type_auto_inserted,
            &axis_category_type_auto_removed,
        )
    });
    let axis_tick_settings_match = chart.axes.iter().enumerate().all(|(axis_index, axis)| {
        axis_optional_field_matches(
            axis_index,
            axis.major_tick_mark.is_some(),
            &axis_major_tick_mark_seen,
            &axis_major_tick_mark_written,
            &axis_major_tick_mark_inserted,
            &axis_major_tick_mark_removed,
        ) && axis_optional_field_matches(
            axis_index,
            axis.minor_tick_mark.is_some(),
            &axis_minor_tick_mark_seen,
            &axis_minor_tick_mark_written,
            &axis_minor_tick_mark_inserted,
            &axis_minor_tick_mark_removed,
        ) && axis_optional_field_matches(
            axis_index,
            axis.tick_label_position.is_some(),
            &axis_tick_label_position_seen,
            &axis_tick_label_position_written,
            &axis_tick_label_position_inserted,
            &axis_tick_label_position_removed,
        ) && axis_optional_field_matches(
            axis_index,
            axis.tick_label_number_format.is_some(),
            &axis_tick_label_number_format_seen,
            &axis_tick_label_number_format_written,
            &axis_tick_label_number_format_inserted,
            &axis_tick_label_number_format_removed,
        ) && axis_numeric_field_matches(
            axis_index,
            axis.tick_label_spacing.map(f64::from),
            &axis_tick_label_spacing_seen,
            &axis_tick_label_spacing_written,
            &axis_tick_label_spacing_inserted,
            &axis_tick_label_spacing_removed,
        ) && axis_numeric_field_matches(
            axis_index,
            axis.tick_mark_spacing.map(f64::from),
            &axis_tick_mark_spacing_seen,
            &axis_tick_mark_spacing_written,
            &axis_tick_mark_spacing_inserted,
            &axis_tick_mark_spacing_removed,
        )
    });
    let axis_gridlines_match = chart.axes.iter().enumerate().all(|(axis_index, axis)| {
        let major_matches = match axis.has_major_gridlines {
            Some(true) => {
                axis_major_gridlines_written
                    .get(axis_index)
                    .copied()
                    .unwrap_or(false)
                    || axis_major_gridlines_inserted
                        .get(axis_index)
                        .copied()
                        .unwrap_or(false)
            }
            Some(false) => {
                !axis_major_gridlines_seen
                    .get(axis_index)
                    .copied()
                    .unwrap_or(false)
                    || axis_major_gridlines_removed
                        .get(axis_index)
                        .copied()
                        .unwrap_or(false)
            }
            None => true,
        };
        let minor_matches = match axis.has_minor_gridlines {
            Some(true) => {
                axis_minor_gridlines_written
                    .get(axis_index)
                    .copied()
                    .unwrap_or(false)
                    || axis_minor_gridlines_inserted
                        .get(axis_index)
                        .copied()
                        .unwrap_or(false)
            }
            Some(false) => {
                !axis_minor_gridlines_seen
                    .get(axis_index)
                    .copied()
                    .unwrap_or(false)
                    || axis_minor_gridlines_removed
                        .get(axis_index)
                        .copied()
                        .unwrap_or(false)
            }
            None => true,
        };
        major_matches && minor_matches
    });

    if chart_type_matches
        && series_sources_match
        && series_orders_match
        && series_explosions_match
        && series_point_explosions_match
        && series_bar_shapes_match
        && series_smooth_values_match
        && series_markers_match
        && series_invert_if_negative_values_match
        && patched_sources == expected_dirty_sources
        && title_matches
        && legend_matches
        && legend_overlay_matches
        && display_blanks_as_matches
        && plot_visible_only_matches
        && show_data_labels_over_maximum_matches
        && view_3d_matches
        && rounded_corners_matches
        && chart_style_matches
        && chart_protection_matches
        && data_table_matches
        && plot_area_layout_matches
        && (preserve_loaded_group_types || vary_colors_matches)
        && (preserve_loaded_group_types || bar_direction_matches)
        && (preserve_loaded_group_types || chart_grouping_matches)
        && (preserve_loaded_group_types || bar_shape_matches)
        && (preserve_loaded_group_types || line_marker_matches)
        && (preserve_loaded_group_types || scatter_style_matches)
        && (preserve_loaded_group_types || radar_style_matches)
        && (preserve_loaded_group_types || of_pie_type_matches)
        && (preserve_loaded_group_types || surface_wireframe_matches)
        && (preserve_loaded_group_types || gap_width_matches)
        && (preserve_loaded_group_types || gap_depth_matches)
        && (preserve_loaded_group_types || overlap_matches)
        && (preserve_loaded_group_types || data_labels_match)
        && series_data_labels_match
        && (preserve_loaded_group_types || chart_group_numeric_settings_match)
        && (preserve_loaded_group_types || chart_group_line_flags_match)
        && axes_match
        && axis_scale_units_match
        && axis_orientation_match
        && axis_crossing_match
        && axis_category_type_match
        && axis_tick_settings_match
        && axis_gridlines_match
    {
        let patched_xml = writer.into_inner().into_inner();
        let patched_xml = if chart.axes.is_empty() {
            patched_xml
        } else {
            rewrite_loaded_chart_axis_additions(patched_xml.as_slice(), chart)?
        };
        Ok(Some(patched_xml))
    } else if preserve_loaded_group_types {
        Err(OmError::unsupported(
            "loaded multi-group chart edit could not be applied losslessly",
        ))
    } else {
        Ok(None)
    }
}

fn validate_volume_stock_chart(chart: &ChartModel) -> OmResult<()> {
    if chart_type_is_volume_stock(&chart.chart_type) {
        let expected_series_count = volume_stock_series_count(&chart.chart_type)
            .expect("volume-stock chart types have a fixed series count");
        if chart.series.len() != expected_series_count {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                format!(
                    "volume stock charts require exactly {expected_series_count} ordered series"
                ),
            ));
        }
        let mut ordered_series = chart.series.iter().enumerate().collect::<Vec<_>>();
        ordered_series.sort_by_key(|(series_index, series)| {
            (series.order.unwrap_or(*series_index as u32), *series_index)
        });
        if ordered_series
            .first()
            .is_none_or(|(_, series)| series.axis_group != ChartAxisGroup::Primary)
            || ordered_series
                .iter()
                .skip(1)
                .any(|(_, series)| series.axis_group != ChartAxisGroup::Secondary)
        {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                "volume stock charts require the volume series on the primary axis and price series on the secondary axis",
            ));
        }
        let has_primary_category = chart.axes.iter().any(|axis| {
            axis.axis_group == ChartAxisGroup::Primary
                && matches!(axis.kind, ChartAxisKind::Category | ChartAxisKind::Date)
        });
        let has_primary_value = chart.axes.iter().any(|axis| {
            axis.axis_group == ChartAxisGroup::Primary && axis.kind == ChartAxisKind::Value
        });
        let has_secondary_value = chart.axes.iter().any(|axis| {
            axis.axis_group == ChartAxisGroup::Secondary && axis.kind == ChartAxisKind::Value
        });
        if !(has_primary_category && has_primary_value && has_secondary_value) {
            return Err(OmError::new(
                OmErrorCode::InvalidState,
                "volume stock charts require shared category, primary volume, and secondary price axes",
            ));
        }
    }
    Ok(())
}

fn serialize_chart_model_xml(chart: &ChartModel) -> OmResult<Vec<u8>> {
    validate_volume_stock_chart(chart)?;
    let series_xml_for_axis_group = |axis_group: ChartAxisGroup| -> OmResult<(String, String)> {
        let group_chart_type = chart_type_for_axis_group(chart, axis_group);
        let mut series_xml = String::new();
        let mut filtered_series_extensions_xml = String::new();
        let mut ordered_series = chart.series.iter().enumerate().collect::<Vec<_>>();
        ordered_series.sort_by_key(|(series_index, series)| {
            (series.order.unwrap_or(*series_index as u32), *series_index)
        });
        for (series_index, series) in ordered_series {
            if series.axis_group != axis_group {
                continue;
            }
            let series_start = series_xml.len();
            let order = series.order.unwrap_or(series_index as u32);
            let raw_index = series.raw_index.unwrap_or(series_index as u32);
            let series_tag = if series.is_filtered {
                "c15:ser"
            } else {
                "c:ser"
            };
            series_xml.push_str(&format!(
                r#"<{series_tag}><c:idx val="{}"/><c:order val="{}"/>"#,
                raw_index, order,
            ));
            if let Some(explosion) = chart_explosion_xml_value(chart) {
                series_xml.push_str(&format!(r#"<c:explosion val="{explosion}"/>"#));
            }
            if chart_type_supports_explosion(&group_chart_type) {
                for (point_index, point) in &series.points {
                    if let Some(explosion) = point.explosion {
                        series_xml.push_str(&format!(
                            r#"<c:dPt><c:idx val="{point_index}"/><c:explosion val="{explosion}"/></c:dPt>"#
                        ));
                    }
                }
            }
            series_xml.push_str(&chart_series_data_labels_xml_string(series));
            if chart_type_supports_series_marker(&group_chart_type)
                && (series.marker_style.is_some() || series.marker_size.is_some())
            {
                series_xml.push_str("<c:marker>");
                if let Some(marker_style) = series.marker_style {
                    let marker_style = chart_marker_style_xml_value(marker_style);
                    series_xml.push_str(&format!(r#"<c:symbol val="{marker_style}"/>"#));
                }
                if let Some(marker_size) = series.marker_size {
                    series_xml.push_str(&format!(r#"<c:size val="{marker_size}"/>"#));
                }
                series_xml.push_str("</c:marker>");
            }
            if let Some(name) = series.name.as_ref() {
                series_xml.push_str(&chart_source_container_xml_string(
                    &group_chart_type,
                    ChartSourceXmlSlot::Name,
                    name,
                )?);
            }
            if let Some(x_values) = series.x_values.as_ref() {
                series_xml.push_str(&chart_source_container_xml_string(
                    &group_chart_type,
                    ChartSourceXmlSlot::XValues,
                    x_values,
                )?);
            }
            if let Some(values) = series.values.as_ref() {
                series_xml.push_str(&chart_source_container_xml_string(
                    &group_chart_type,
                    ChartSourceXmlSlot::Values,
                    values,
                )?);
            }
            if chart_type_uses_bubble_size(&group_chart_type)
                && let Some(bubble_size) = series.bubble_size.as_ref()
            {
                series_xml.push_str(&chart_source_container_xml_string(
                    &group_chart_type,
                    ChartSourceXmlSlot::BubbleSize,
                    bubble_size,
                )?);
            }
            if let Some(invert_if_negative) = series.invert_if_negative {
                series_xml.push_str(&format!(
                    r#"<c:invertIfNegative val="{}"/>"#,
                    if invert_if_negative { "1" } else { "0" }
                ));
            }
            if chart_type_supports_bar_shape(&group_chart_type)
                && let Some(bar_shape) = series.bar_shape
            {
                let bar_shape = chart_bar_shape_xml_value(bar_shape);
                series_xml.push_str(&format!(r#"<c:shape val="{bar_shape}"/>"#));
            }
            if chart_type_supports_series_smooth(&group_chart_type)
                && let Some(smooth) = series.smooth
            {
                series_xml.push_str(&format!(
                    r#"<c:smooth val="{}"/>"#,
                    if smooth { "1" } else { "0" }
                ));
            }
            series_xml.push_str(&format!("</{series_tag}>"));
            if series.is_filtered {
                let series_fragment = series_xml.split_off(series_start);
                let wrapper_name = chart_filtered_series_wrapper_name(&group_chart_type)
                    .ok_or_else(|| {
                        OmError::unsupported(
                            "saving filtered series requires a supported chart type",
                        )
                    })?;
                filtered_series_extensions_xml.push_str(&format!(
                    r#"<c:ext uri="{{02D57815-91ED-43cb-92C2-25804820EDAC}}"><c15:{wrapper_name} xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">{series_fragment}</c15:{wrapper_name}></c:ext>"#
                ));
            }
        }
        Ok((series_xml, filtered_series_extensions_xml))
    };
    let (primary_series_xml, primary_filtered_series_xml) =
        series_xml_for_axis_group(ChartAxisGroup::Primary)?;
    let (secondary_series_xml, secondary_filtered_series_xml) =
        series_xml_for_axis_group(ChartAxisGroup::Secondary)?;
    let chart_group_name = chart_group_xml_name(&chart.chart_type).ok_or_else(|| {
        OmError::unsupported("saving dirty charts requires a supported chart type")
    })?;
    let vary_colors_xml = chart
        .vary_by_categories
        .map(|value| format!(r#"<c:varyColors val="{}"/>"#, if value { "1" } else { "0" }))
        .unwrap_or_default();
    let bar_direction_xml = chart_type_bar_direction_xml_value(&chart.chart_type)
        .map(|value| format!(r#"<c:barDir val="{value}"/>"#))
        .unwrap_or_default();
    let chart_grouping_xml = chart_type_grouping_xml_value(&chart.chart_type)
        .map(|value| format!(r#"<c:grouping val="{value}"/>"#))
        .unwrap_or_default();
    let bar_shape_xml = chart_effective_bar_shape(chart)
        .map(chart_bar_shape_xml_value)
        .map(|value| format!(r#"<c:shape val="{value}"/>"#))
        .unwrap_or_default();
    let line_marker_xml = chart_type_line_marker_xml_value(&chart.chart_type)
        .map(|value| format!(r#"<c:marker val="{value}"/>"#))
        .unwrap_or_default();
    let scatter_style_xml = chart_type_scatter_style_xml_value(&chart.chart_type)
        .map(|value| format!(r#"<c:scatterStyle val="{value}"/>"#))
        .unwrap_or_default();
    let radar_style_xml = chart_type_radar_style_xml_value(&chart.chart_type)
        .map(|value| format!(r#"<c:radarStyle val="{value}"/>"#))
        .unwrap_or_default();
    let of_pie_type_xml = chart_type_of_pie_xml_value(&chart.chart_type)
        .map(|value| format!(r#"<c:ofPieType val="{value}"/>"#))
        .unwrap_or_default();
    let surface_wireframe_xml = chart_type_surface_wireframe_xml_value(&chart.chart_type)
        .map(|value| format!(r#"<c:wireframe val="{value}"/>"#))
        .unwrap_or_default();
    let gap_width_xml = chart
        .gap_width
        .map(|value| format!(r#"<c:gapWidth val="{value}"/>"#))
        .unwrap_or_default();
    let gap_depth_xml = chart
        .gap_depth
        .filter(|_| chart_type_supports_gap_depth(&chart.chart_type))
        .map(|value| format!(r#"<c:gapDepth val="{value}"/>"#))
        .unwrap_or_default();
    let overlap_xml = chart
        .overlap
        .map(|value| format!(r#"<c:overlap val="{value}"/>"#))
        .unwrap_or_default();
    let series_lines_xml = if chart.has_series_lines == Some(true) {
        r#"<c:serLines/>"#
    } else {
        ""
    };
    let drop_lines_xml = if chart.has_drop_lines == Some(true) {
        r#"<c:dropLines/>"#
    } else {
        ""
    };
    let has_hi_lo_lines = chart.has_hi_lo_lines.or_else(|| {
        matches!(
            chart.chart_type,
            ChartType::StockHLC
                | ChartType::StockOHLC
                | ChartType::StockVHLC
                | ChartType::StockVOHLC
        )
        .then_some(true)
    });
    let hi_lo_lines_xml = if has_hi_lo_lines == Some(true) {
        r#"<c:hiLowLines/>"#
    } else {
        ""
    };
    let has_up_down_bars = chart.has_up_down_bars.or_else(|| {
        matches!(
            chart.chart_type,
            ChartType::StockOHLC | ChartType::StockVOHLC
        )
        .then_some(true)
    });
    let up_down_bars_xml = if has_up_down_bars == Some(true) {
        r#"<c:upDownBars/>"#
    } else {
        ""
    };
    let first_slice_angle_xml = chart
        .first_slice_angle
        .map(|value| format!(r#"<c:firstSliceAng val="{value}"/>"#))
        .unwrap_or_default();
    let bubble_scale_xml = chart
        .bubble_scale
        .map(|value| format!(r#"<c:bubbleScale val="{value}"/>"#))
        .unwrap_or_default();
    let show_negative_bubbles_xml = chart
        .show_negative_bubbles
        .map(|value| {
            format!(
                r#"<c:showNegBubbles val="{}"/>"#,
                if value { "1" } else { "0" }
            )
        })
        .unwrap_or_default();
    let has_3d_shading_xml = chart
        .has_3d_shading
        .or_else(|| matches!(chart.chart_type, ChartType::Bubble3DEffect).then_some(true))
        .map(|value| format!(r#"<c:bubble3D val="{}"/>"#, if value { "1" } else { "0" }))
        .unwrap_or_default();
    let doughnut_hole_size_xml = chart
        .doughnut_hole_size
        .map(|value| format!(r#"<c:holeSize val="{value}"/>"#))
        .unwrap_or_default();
    let second_plot_size_xml = chart
        .second_plot_size
        .map(|value| format!(r#"<c:secondPieSize val="{value}"/>"#))
        .unwrap_or_default();
    let size_represents_xml = chart
        .size_represents
        .map(|value| {
            format!(
                r#"<c:sizeRepresents val="{}"/>"#,
                chart_size_represents_xml_value(value)
            )
        })
        .unwrap_or_default();
    let split_type_xml = chart
        .split_type
        .map(|value| {
            format!(
                r#"<c:splitType val="{}"/>"#,
                chart_split_type_xml_value(value)
            )
        })
        .unwrap_or_default();
    let split_value_xml = chart
        .split_value
        .map(|value| format!(r#"<c:splitPos val="{}"/>"#, chart_number_xml_value(value)))
        .unwrap_or_default();
    let data_labels_xml = chart
        .data_labels
        .as_ref()
        .map(chart_data_labels_xml_string)
        .unwrap_or_default();
    let title_xml = chart
        .title
        .as_ref()
        .map(|title| {
            let title_text = partial_escape(&title.text).to_string();
            format!(
                r#"<c:title><c:tx><c:rich><a:p><a:r><a:t>{title_text}</a:t></a:r></a:p></c:rich></c:tx></c:title>"#
            )
        })
        .unwrap_or_default();
    let legend_xml = chart
        .legend
        .as_ref()
        .filter(|legend| legend.visible)
        .map(|legend| {
            let legend_position = chart_legend_position_xml_value(
                legend.position.unwrap_or(ChartLegendPosition::Right),
            );
            let legend_overlay_xml = legend
                .include_in_layout
                .map(|include_in_layout| {
                    format!(
                        r#"<c:overlay val="{}"/>"#,
                        if include_in_layout { "0" } else { "1" }
                    )
                })
                .unwrap_or_default();
            format!(
                r#"<c:legend><c:legendPos val="{legend_position}"/>{legend_overlay_xml}</c:legend>"#
            )
        })
        .unwrap_or_default();
    let plot_visible_only_xml = chart
        .plot_visible_only
        .map(|value| {
            format!(
                r#"<c:plotVisOnly val="{}"/>"#,
                if value { "1" } else { "0" }
            )
        })
        .unwrap_or_default();
    let display_blanks_as_xml = chart
        .display_blanks_as
        .map(|value| {
            format!(
                r#"<c:dispBlanksAs val="{}"/>"#,
                chart_display_blanks_as_xml_value(value)
            )
        })
        .unwrap_or_default();
    let show_data_labels_over_maximum_xml = chart
        .show_data_labels_over_maximum
        .map(|value| {
            format!(
                r#"<c:showDLblsOverMax val="{}"/>"#,
                if value { "1" } else { "0" }
            )
        })
        .unwrap_or_default();
    let view_3d_xml = chart
        .view_3d
        .as_ref()
        .and_then(chart_view_3d_xml_string)
        .unwrap_or_default();
    let rounded_corners_xml = chart
        .rounded_corners
        .map(|value| {
            format!(
                r#"<c:roundedCorners val="{}"/>"#,
                if value { "1" } else { "0" }
            )
        })
        .unwrap_or_default();
    let chart_style_xml = chart
        .style
        .map(|value| format!(r#"<c:style val="{value}"/>"#))
        .unwrap_or_default();
    let chart_protection_xml = chart
        .protection
        .map(chart_protection_xml)
        .unwrap_or_default();
    let data_table_xml = chart
        .data_table
        .as_ref()
        .map(chart_data_table_xml_string)
        .unwrap_or_default();
    let chart_has_axes = chart_type_has_axes(&chart.chart_type);
    let has_secondary_series = chart_type_is_volume_stock(&chart.chart_type)
        || !secondary_series_xml.is_empty()
        || !secondary_filtered_series_xml.is_empty();
    let has_secondary_category_axis = chart.axes.iter().any(|axis| {
        axis.axis_group == ChartAxisGroup::Secondary
            && matches!(axis.kind, ChartAxisKind::Category | ChartAxisKind::Date)
    });
    let has_secondary_series_axis = chart.axes.iter().any(|axis| {
        axis.axis_group == ChartAxisGroup::Secondary && axis.kind == ChartAxisKind::Series
    });
    let mut all_axis_refs = String::new();
    let mut primary_axis_refs = String::new();
    let mut secondary_axis_refs = String::new();
    let mut axes_xml = String::new();
    if chart_has_axes {
        for (axis_index, axis) in chart.axes.iter().enumerate() {
            let axis_id = axis
                .raw_id
                .clone()
                .unwrap_or_else(|| ((axis_index + 1) * 10).to_string());
            let escaped_axis_id = partial_escape(&axis_id).to_string();
            let axis_ref = format!(r#"<c:axId val="{escaped_axis_id}"/>"#);
            all_axis_refs.push_str(&axis_ref);
            match axis.axis_group {
                ChartAxisGroup::Primary => {
                    primary_axis_refs.push_str(&axis_ref);
                    if has_secondary_series
                        && ((matches!(axis.kind, ChartAxisKind::Category | ChartAxisKind::Date)
                            && !has_secondary_category_axis)
                            || (axis.kind == ChartAxisKind::Series && !has_secondary_series_axis))
                    {
                        secondary_axis_refs.push_str(&axis_ref);
                    }
                }
                ChartAxisGroup::Secondary => secondary_axis_refs.push_str(&axis_ref),
            }
            let axis_tag = chart_axis_xml_name(axis.kind);
            let mut scaling_xml = String::new();
            if chart_axis_has_scaling_xml(axis) {
                scaling_xml.push_str("<c:scaling>");
                if let Some(value) = chart_axis_log_base_xml_value(axis) {
                    scaling_xml.push_str(&format!(
                        r#"<c:logBase val="{}"/>"#,
                        chart_number_xml_value(value)
                    ));
                }
                if let Some(value) = axis.reverse_plot_order {
                    scaling_xml.push_str(&format!(
                        r#"<c:orientation val="{}"/>"#,
                        chart_axis_orientation_xml_value(value)
                    ));
                }
                if let Some(value) = axis.minimum_scale {
                    scaling_xml.push_str(&format!(
                        r#"<c:min val="{}"/>"#,
                        chart_number_xml_value(value)
                    ));
                }
                if let Some(value) = axis.maximum_scale {
                    scaling_xml.push_str(&format!(
                        r#"<c:max val="{}"/>"#,
                        chart_number_xml_value(value)
                    ));
                }
                scaling_xml.push_str("</c:scaling>");
            }
            let axis_title_xml = axis
                .title
                .as_ref()
                .map(|title| {
                    let title_text = partial_escape(&title.text).to_string();
                    format!(
                        r#"<c:title><c:tx><c:rich><a:p><a:r><a:t>{title_text}</a:t></a:r></a:p></c:rich></c:tx></c:title>"#
                    )
                })
                .unwrap_or_default();
            let axis_deleted_xml = axis
                .deleted
                .map(|value| format!(r#"<c:delete val="{}"/>"#, if value { "1" } else { "0" }))
                .unwrap_or_default();
            let major_gridlines_xml = if axis.has_major_gridlines == Some(true) {
                r#"<c:majorGridlines/>"#
            } else {
                ""
            };
            let minor_gridlines_xml = if axis.has_minor_gridlines == Some(true) {
                r#"<c:minorGridlines/>"#
            } else {
                ""
            };
            let major_tick_mark_xml = axis
                .major_tick_mark
                .map(|value| {
                    format!(
                        r#"<c:majorTickMark val="{}"/>"#,
                        chart_tick_mark_xml_value(value)
                    )
                })
                .unwrap_or_default();
            let minor_tick_mark_xml = axis
                .minor_tick_mark
                .map(|value| {
                    format!(
                        r#"<c:minorTickMark val="{}"/>"#,
                        chart_tick_mark_xml_value(value)
                    )
                })
                .unwrap_or_default();
            let tick_label_position_xml = axis
                .tick_label_position
                .map(|value| {
                    format!(
                        r#"<c:tickLblPos val="{}"/>"#,
                        chart_tick_label_position_xml_value(value)
                    )
                })
                .unwrap_or_default();
            let tick_label_number_format_xml = axis
                .tick_label_number_format
                .as_ref()
                .map(|format_code| {
                    let format_code = partial_escape(format_code).to_string();
                    let source_linked = if axis.tick_label_number_format_linked.unwrap_or(true) {
                        "1"
                    } else {
                        "0"
                    };
                    format!(
                        r#"<c:numFmt formatCode="{format_code}" sourceLinked="{source_linked}"/>"#
                    )
                })
                .unwrap_or_default();
            let tick_label_spacing_xml = axis
                .tick_label_spacing
                .map(|value| format!(r#"<c:tickLblSkip val="{value}"/>"#))
                .unwrap_or_default();
            let tick_mark_spacing_xml = axis
                .tick_mark_spacing
                .map(|value| format!(r#"<c:tickMarkSkip val="{value}"/>"#))
                .unwrap_or_default();
            let major_unit_xml = axis
                .major_unit
                .map(|value| format!(r#"<c:majorUnit val="{}"/>"#, chart_number_xml_value(value)))
                .unwrap_or_default();
            let minor_unit_xml = axis
                .minor_unit
                .map(|value| format!(r#"<c:minorUnit val="{}"/>"#, chart_number_xml_value(value)))
                .unwrap_or_default();
            let display_units_xml = axis
                .display_unit
                .map(|display_unit| {
                    let unit_xml = match display_unit {
                        ChartAxisDisplayUnit::BuiltIn(value) => format!(
                            r#"<c:builtInUnit val="{}"/>"#,
                            chart_built_in_display_unit_xml_value(value)
                        ),
                        ChartAxisDisplayUnit::Custom(value) => {
                            format!(r#"<c:custUnit val="{}"/>"#, chart_number_xml_value(value))
                        }
                    };
                    let label_xml = if axis.has_display_unit_label == Some(true) {
                        if axis.display_unit_label.is_some() {
                            let label_text =
                                partial_escape(&chart_axis_display_unit_label_text(axis))
                                    .to_string();
                            format!(
                                r#"<c:dispUnitsLbl><c:tx><c:rich><a:p><a:r><a:t>{label_text}</a:t></a:r></a:p></c:rich></c:tx></c:dispUnitsLbl>"#
                            )
                        } else {
                            "<c:dispUnitsLbl/>".to_string()
                        }
                    } else {
                        String::new()
                    };
                    format!(r#"<c:dispUnits>{unit_xml}{label_xml}</c:dispUnits>"#)
                })
                .unwrap_or_default();
            let crossing_xml = if let Some(value) = axis.crosses_at {
                format!(r#"<c:crossesAt val="{}"/>"#, chart_number_xml_value(value))
            } else {
                axis.crosses
                    .and_then(chart_axis_crosses_xml_value)
                    .map(|value| format!(r#"<c:crosses val="{value}"/>"#))
                    .unwrap_or_default()
            };
            let cross_axis_xml = chart_axis_cross_target_id(&chart.axes, axis)
                .map(|axis_id| {
                    format!(r#"<c:crossAx val="{}"/>"#, partial_escape(axis_id.as_str()))
                })
                .unwrap_or_default();
            let cross_between_xml = axis
                .axis_between_categories
                .map(|value| {
                    format!(
                        r#"<c:crossBetween val="{}"/>"#,
                        chart_axis_between_categories_xml_value(value)
                    )
                })
                .unwrap_or_default();
            let category_type_auto_xml = axis
                .category_type_auto
                .map(|value| format!(r#"<c:auto val="{}"/>"#, if value { "1" } else { "0" }))
                .unwrap_or_default();
            let base_unit_xml = axis
                .base_unit
                .map(|value| {
                    format!(
                        r#"<c:baseTimeUnit val="{}"/>"#,
                        chart_axis_time_unit_xml_value(value)
                    )
                })
                .unwrap_or_default();
            let major_time_unit_xml = axis
                .major_unit_scale
                .map(|value| {
                    format!(
                        r#"<c:majorTimeUnit val="{}"/>"#,
                        chart_axis_time_unit_xml_value(value)
                    )
                })
                .unwrap_or_default();
            let minor_time_unit_xml = axis
                .minor_unit_scale
                .map(|value| {
                    format!(
                        r#"<c:minorTimeUnit val="{}"/>"#,
                        chart_axis_time_unit_xml_value(value)
                    )
                })
                .unwrap_or_default();
            axes_xml.push_str(&format!(
                r#"<c:{axis_tag}><c:axId val="{escaped_axis_id}"/>{scaling_xml}{axis_deleted_xml}{major_gridlines_xml}{minor_gridlines_xml}{axis_title_xml}{tick_label_number_format_xml}{major_tick_mark_xml}{minor_tick_mark_xml}{tick_label_position_xml}{tick_label_spacing_xml}{tick_mark_spacing_xml}{major_unit_xml}{minor_unit_xml}{display_units_xml}{cross_axis_xml}{crossing_xml}{cross_between_xml}{category_type_auto_xml}{base_unit_xml}{major_time_unit_xml}{minor_time_unit_xml}</c:{axis_tag}>"#
            ));
        }
    }
    if primary_axis_refs.is_empty() {
        primary_axis_refs.push_str(&all_axis_refs);
    }
    if secondary_axis_refs.is_empty() {
        secondary_axis_refs.push_str(&all_axis_refs);
    }
    let chart_group_xml = |series_xml: &str, filtered_series_xml: &str, axis_refs: &str| {
        let filtered_series_extensions_xml = if filtered_series_xml.is_empty() {
            String::new()
        } else {
            format!("<c:extLst>{filtered_series_xml}</c:extLst>")
        };
        format!(
            r#"<c:{chart_group_name}>{bar_direction_xml}{chart_grouping_xml}{bar_shape_xml}{line_marker_xml}{scatter_style_xml}{radar_style_xml}{of_pie_type_xml}{surface_wireframe_xml}{vary_colors_xml}{series_xml}{gap_width_xml}{gap_depth_xml}{overlap_xml}{first_slice_angle_xml}{bubble_scale_xml}{show_negative_bubbles_xml}{has_3d_shading_xml}{doughnut_hole_size_xml}{second_plot_size_xml}{size_represents_xml}{split_type_xml}{split_value_xml}{data_labels_xml}{series_lines_xml}{drop_lines_xml}{hi_lo_lines_xml}{up_down_bars_xml}{axis_refs}{filtered_series_extensions_xml}</c:{chart_group_name}>"#
        )
    };
    let filtered_series_extensions = |filtered_series_xml: &str| {
        if filtered_series_xml.is_empty() {
            String::new()
        } else {
            format!("<c:extLst>{filtered_series_xml}</c:extLst>")
        }
    };
    let chart_groups_xml = if chart_type_is_volume_stock(&chart.chart_type) {
        let volume_extensions_xml = filtered_series_extensions(&primary_filtered_series_xml);
        let stock_extensions_xml = filtered_series_extensions(&secondary_filtered_series_xml);
        format!(
            r#"<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>{vary_colors_xml}{primary_series_xml}{gap_width_xml}{overlap_xml}{primary_axis_refs}{volume_extensions_xml}</c:barChart><c:stockChart>{secondary_series_xml}{data_labels_xml}{drop_lines_xml}{hi_lo_lines_xml}{up_down_bars_xml}{secondary_axis_refs}{stock_extensions_xml}</c:stockChart>"#
        )
    } else if has_secondary_series {
        let mut chart_groups_xml = String::new();
        if !primary_series_xml.is_empty() || !primary_filtered_series_xml.is_empty() {
            chart_groups_xml.push_str(&chart_group_xml(
                &primary_series_xml,
                &primary_filtered_series_xml,
                &primary_axis_refs,
            ));
        }
        chart_groups_xml.push_str(&chart_group_xml(
            &secondary_series_xml,
            &secondary_filtered_series_xml,
            &secondary_axis_refs,
        ));
        chart_groups_xml
    } else {
        chart_group_xml(
            &primary_series_xml,
            &primary_filtered_series_xml,
            &all_axis_refs,
        )
    };
    let plot_area_layout_xml = chart
        .plot_area_layout
        .as_ref()
        .map(|layout| {
            format!(
                "<c:layout>{}</c:layout>",
                chart_manual_layout_xml_string(layout)
            )
        })
        .unwrap_or_default();
    Ok(format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  {rounded_corners_xml}{chart_style_xml}{chart_protection_xml}<c:chart>{title_xml}{view_3d_xml}<c:plotArea>{plot_area_layout_xml}{chart_groups_xml}{axes_xml}{data_table_xml}</c:plotArea>{legend_xml}{plot_visible_only_xml}{display_blanks_as_xml}{show_data_labels_over_maximum_xml}</c:chart>
</c:chartSpace>"#
    )
    .into_bytes())
}

fn chart_axis_cross_target_id(axes: &[AxisModel], axis: &AxisModel) -> Option<String> {
    if let Some(axis_id) = axis.cross_axis_raw_id.as_ref()
        && axes
            .iter()
            .any(|candidate| candidate.raw_id.as_ref() == Some(axis_id))
    {
        return Some(axis_id.clone());
    }
    match axis.kind {
        ChartAxisKind::Value => axes
            .iter()
            .find(|candidate| {
                candidate.axis_group == axis.axis_group
                    && matches!(
                        candidate.kind,
                        ChartAxisKind::Category | ChartAxisKind::Date
                    )
            })
            .or_else(|| {
                axes.iter().find(|candidate| {
                    candidate.axis_group == ChartAxisGroup::Primary
                        && matches!(
                            candidate.kind,
                            ChartAxisKind::Category | ChartAxisKind::Date
                        )
                })
            }),
        ChartAxisKind::Category | ChartAxisKind::Date | ChartAxisKind::Series => axes
            .iter()
            .find(|candidate| {
                candidate.axis_group == axis.axis_group && candidate.kind == ChartAxisKind::Value
            })
            .or_else(|| {
                axes.iter().find(|candidate| {
                    candidate.axis_group == ChartAxisGroup::Primary
                        && candidate.kind == ChartAxisKind::Value
                })
            })
            .or_else(|| {
                axes.iter()
                    .find(|candidate| candidate.kind == ChartAxisKind::Value)
            }),
    }
    .and_then(|target| target.raw_id.clone())
}

fn chart_group_overlay_is_stable(chart: &ChartModel) -> bool {
    if chart.groups.is_empty() {
        return false;
    }
    let mut series_by_raw_index = BTreeMap::<u32, ChartAxisGroup>::new();
    for series in &chart.series {
        let Some(raw_index) = series.raw_index else {
            return false;
        };
        if series_by_raw_index
            .insert(raw_index, series.axis_group)
            .is_some()
        {
            return false;
        }
    }
    let mut grouped_raw_indices = BTreeSet::new();
    for group in &chart.groups {
        for raw_index in &group.series_raw_indices {
            if !grouped_raw_indices.insert(*raw_index)
                || series_by_raw_index.get(raw_index) != Some(&group.axis_group)
            {
                return false;
            }
        }
    }
    grouped_raw_indices.len() == series_by_raw_index.len()
}

fn chart_group_index_for_series_raw_index(chart: &ChartModel, raw_index: u32) -> OmResult<usize> {
    let mut matches = chart
        .groups
        .iter()
        .enumerate()
        .filter_map(|(group_index, group)| {
            group
                .series_raw_indices
                .contains(&raw_index)
                .then_some(group_index)
        });
    let group_index = matches.next().ok_or_else(|| {
        OmError::new(
            OmErrorCode::InvalidState,
            format!("chart series c:idx {raw_index} has no owning chart group"),
        )
    })?;
    if matches.next().is_some() {
        return Err(OmError::new(
            OmErrorCode::InvalidState,
            format!("chart series c:idx {raw_index} belongs to multiple chart groups"),
        ));
    }
    Ok(group_index)
}

fn formula_cell_error_text(error: &CellError) -> &str {
    error.as_lexical_str()
}
