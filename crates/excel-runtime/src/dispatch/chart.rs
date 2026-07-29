use super::super::{
    BorderParent, ChartFormatParent, ChartGroupLineKind, ChartGroupShortcutKind,
    ChartObjectsParent, ExcelRuntime, MSO_FALSE, MSO_FLIP_HORIZONTAL, MSO_FLIP_VERTICAL,
    MSO_SCALE_FROM_BOTTOM_RIGHT, MSO_SCALE_FROM_MIDDLE, MSO_SCALE_FROM_TOP_LEFT, MSO_SHAPE_CHART,
    MSO_SHAPE_MIXED, MSO_TRUE, RuntimeObjectKind, ShapeRangeSource, XL_AUTOMATIC_SCALE, XL_BOX,
    XL_CATEGORY, XL_CATEGORY_SCALE, XL_CREATOR_CODE, XL_DISPLAY_UNIT_CUSTOM, XL_DISPLAY_UNIT_NONE,
    XL_FREE_FLOATING, XL_LEGEND_POSITION_BOTTOM, XL_LEGEND_POSITION_CORNER,
    XL_LEGEND_POSITION_CUSTOM, XL_LEGEND_POSITION_LEFT, XL_LEGEND_POSITION_RIGHT,
    XL_LEGEND_POSITION_TOP, XL_MARKER_STYLE_AUTOMATIC, XL_ORIENTATION_HORIZONTAL,
    XL_PLOT_BY_COLUMNS, XL_PLOT_BY_ROWS, XL_PRIMARY, XL_READING_ORDER_CONTEXT, XL_SECONDARY,
    XL_SERIES_AXIS, XL_SIZE_IS_AREA, XL_SIZE_IS_WIDTH, XL_SPLIT_BY_CUSTOM_SPLIT,
    XL_SPLIT_BY_PERCENT_VALUE, XL_SPLIT_BY_POSITION, XL_SPLIT_BY_VALUE,
    XL_TICK_LABEL_ORIENTATION_AUTOMATIC, XL_TIME_SCALE, XL_VALUE, attach_series_to_chart_group,
    chart_axis_crosses_to_excel_value, chart_axis_display_unit_label_text,
    chart_axis_group_from_excel_value, chart_axis_scale_type_to_excel_value,
    chart_axis_time_unit_to_excel_value, chart_bar_shape_to_excel_value,
    chart_built_in_display_unit_to_excel_value, chart_data_label_position_to_excel_value,
    chart_data_labels_count_for_chart_series, chart_data_labels_type_to_excel_value,
    chart_data_labels_visible, chart_display_blanks_as_to_excel_value, chart_effective_bar_shape,
    chart_group_axis_group, chart_group_chart_type, chart_group_indices,
    chart_group_overlay_is_stable, chart_marker_style_to_excel_value, chart_object_placement_value,
    chart_object_z_order_operation, chart_plot_by_from_optional_arg,
    chart_point_effective_data_labels, chart_series_effective_data_labels,
    chart_series_point_count, chart_source_expr_for_range, chart_source_expr_for_range_areas,
    chart_source_expr_text, chart_source_value_text_for_index,
    chart_tick_label_position_to_excel_value, chart_tick_mark_to_excel_value,
    chart_type_default_series_smooth, chart_type_for_series, chart_type_is_volume_stock,
    chart_type_supports_high_low_lines, chart_type_supports_radar_axis_labels,
    chart_type_supports_series_marker, chart_type_supports_series_smooth,
    chart_type_supports_up_down_bars, chart_type_to_excel_value, coerce_optional_bool_arg,
    coerce_positive_index, coerce_u32_arg, default_chart_axes_for_type,
    ensure_chart_supports_3d_view, ensure_chart_supports_bar_shape,
    ensure_chart_supports_gap_depth, ensure_chart_supports_right_angle_axes,
    graphic_frame_lock_aspect_ratio, graphic_frame_rotation_units,
    graphic_frame_transform_bool_attr, is_chart_group_shortcut_member, next_chart_series_raw_index,
    normalize_volume_stock_chart, om_value_is_omitted, resolve_series_insert_group,
    rotation_units_from_degrees, series_collection_indices, series_formula_text, series_plot_order,
    set_graphic_frame_rotation_xml, set_graphic_frame_transform_bool_attr_xml,
    update_series_plot_order, validate_copy_picture_args,
};
use excel_model::{
    ChartAxisCrosses, ChartAxisDisplayUnit, ChartAxisGroup, ChartAxisKind, ChartAxisScaleType,
    ChartAxisTimeUnit, ChartDataLabelPosition, ChartDisplayBlanksAs, ChartLayoutMode,
    ChartLegendPosition, ChartModel, ChartObjectModel, ChartSizeRepresents, ChartSourceExpr,
    ChartSplitType, ChartTickLabelPosition, ChartTickMark, DrawingModel, DrawingObjectModel,
    SeriesModel,
};
use office_common::{
    AbsoluteAnchor, ChartId, ChartObjectId, DrawingAnchor, DrawingId, ObjectPlacement, OmError,
    OmErrorCode, OmResult, OmValue, PointEmu, RangeSet, Rect, SheetId, SheetKind, SheetScope,
    SizeEmu, WorkbookHandle,
};
use std::collections::BTreeMap;

impl ExcelRuntime {
    pub(crate) fn dispatch_get_chart_objects(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        parent: ChartObjectsParent,
        chart_object_ids: Option<&[ChartObjectId]>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if matches!(
            member,
            "Count"
                | "Item"
                | "Application"
                | "Parent"
                | "Creator"
                | "ShapeRange"
                | "Left"
                | "Top"
                | "Width"
                | "Height"
                | "Visible"
                | "Placement"
                | "PrintObject"
                | "Locked"
                | "ProtectChartObject"
                | "RoundedCorners"
        ) {
            self.focus_member_supported("ChartObjects", member, false)?;
        }

        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Count does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(
                    self.chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?
                        .len() as f64,
                ))
            }
            "Item" => self.dispatch_invoke_chart_objects(
                workbook,
                sheet_id,
                parent,
                chart_object_ids,
                member,
                args,
            ),
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Application does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Parent does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(match parent {
                    ChartObjectsParent::Worksheet(parent_sheet_id) => {
                        self.register_worksheet_handle(workbook, parent_sheet_id).0
                    }
                    ChartObjectsParent::Chart(parent_chart_id) => {
                        self.register_chart_handle(workbook, parent_chart_id)
                    }
                }))
            }
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Creator does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "ShapeRange" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.ShapeRange does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.register_shape_range_handle(
                    workbook,
                    ShapeRangeSource::ChartObjects {
                        sheet_id,
                        parent,
                        chart_object_ids: chart_object_ids.map(|ids| ids.to_vec()),
                    },
                )))
            }
            "Left" | "Top" | "Width" | "Height" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "ChartObjects.{member} does not accept arguments"
                    )));
                }
                let mut bounds: Option<(f64, f64, f64, f64)> = None;
                for (chart_object_id, _) in
                    self.chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?
                {
                    let chart_object = self.chart_object_model(workbook, chart_object_id)?;
                    let left = Self::chart_object_geometry_value(chart_object, "Left")?;
                    let top = Self::chart_object_geometry_value(chart_object, "Top")?;
                    let right = left + Self::chart_object_geometry_value(chart_object, "Width")?;
                    let bottom = top + Self::chart_object_geometry_value(chart_object, "Height")?;
                    bounds = Some(match bounds {
                        Some((current_left, current_top, current_right, current_bottom)) => (
                            current_left.min(left),
                            current_top.min(top),
                            current_right.max(right),
                            current_bottom.max(bottom),
                        ),
                        None => (left, top, right, bottom),
                    });
                }
                let value = match (member, bounds) {
                    (_, None) => 0.0,
                    ("Left", Some((left, _, _, _))) => left,
                    ("Top", Some((_, top, _, _))) => top,
                    ("Width", Some((left, _, right, _))) => right - left,
                    ("Height", Some((_, top, _, bottom))) => bottom - top,
                    _ => unreachable!("ChartObjects geometry member was matched"),
                };
                Ok(OmValue::Number(value))
            }
            "Visible" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Visible does not accept arguments",
                    ));
                }
                let mut visible = true;
                for (chart_object_id, _) in
                    self.chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?
                {
                    if self
                        .chart_object_model(workbook, chart_object_id)?
                        .non_visual_attrs
                        .get("hidden")
                        .is_some_and(|value| {
                            value == "1"
                                || value.eq_ignore_ascii_case("true")
                                || value.eq_ignore_ascii_case("on")
                        })
                    {
                        visible = false;
                        break;
                    }
                }
                Ok(OmValue::Bool(visible))
            }
            "Placement" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Placement does not accept arguments",
                    ));
                }
                let placement = self
                    .chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?
                    .into_iter()
                    .next()
                    .map(|(chart_object_id, _)| {
                        chart_object_placement_value(
                            &self
                                .chart_object_model(workbook, chart_object_id)?
                                .placement,
                        )
                    })
                    .transpose()?
                    .unwrap_or(XL_FREE_FLOATING);
                Ok(OmValue::Number(f64::from(placement)))
            }
            "PrintObject" | "Locked" | "ProtectChartObject" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "ChartObjects.{member} does not accept arguments"
                    )));
                }
                let attr_name = match member {
                    "PrintObject" => "fPrintsWithSheet",
                    "Locked" | "ProtectChartObject" => "fLocksWithSheet",
                    _ => unreachable!("ChartObjects clientData boolean member was matched"),
                };
                let mut enabled = true;
                for (chart_object_id, _) in
                    self.chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?
                {
                    if self
                        .chart_object_model(workbook, chart_object_id)?
                        .client_data_attrs
                        .get(attr_name)
                        .is_some_and(|value| {
                            value == "0"
                                || value.eq_ignore_ascii_case("false")
                                || value.eq_ignore_ascii_case("off")
                        })
                    {
                        enabled = false;
                        break;
                    }
                }
                Ok(OmValue::Bool(enabled))
            }
            "RoundedCorners" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.RoundedCorners does not accept arguments",
                    ));
                }
                let mut saw_chart_object = false;
                let mut rounded_corners = true;
                for (chart_object_id, _) in
                    self.chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?
                {
                    saw_chart_object = true;
                    let chart_id = self.chart_object_model(workbook, chart_object_id)?.chart_id;
                    if !self
                        .chart_model(workbook, chart_id)?
                        .rounded_corners
                        .unwrap_or(false)
                    {
                        rounded_corners = false;
                        break;
                    }
                }
                Ok(OmValue::Bool(saw_chart_object && rounded_corners))
            }
            _ => Err(OmError::unsupported(format!(
                "ChartObjects.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_chart_objects(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        parent: ChartObjectsParent,
        chart_object_ids: Option<&[ChartObjectId]>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if matches!(
            member,
            "Item"
                | "Add"
                | "Delete"
                | "Copy"
                | "Cut"
                | "Duplicate"
                | "CopyPicture"
                | "Select"
                | "BringToFront"
                | "SendToBack"
        ) {
            self.focus_member_supported("ChartObjects", member, false)?;
        }

        match member {
            "Item" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Item expects a single chart object index or name",
                    ));
                };
                let entries =
                    self.chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?;
                let resolve_chart_object = |selector: &OmValue| -> OmResult<ChartObjectId> {
                    match selector {
                        OmValue::Number(_) => {
                            let index =
                                coerce_u32_arg(selector, "ChartObjects.Item index")? as usize;
                            if index == 0 {
                                return Err(OmError::invalid_argument(
                                    "ChartObjects.Item index is out of bounds",
                                ));
                            }
                            entries
                                .get(index - 1)
                                .map(|(chart_object_id, _)| *chart_object_id)
                                .ok_or_else(|| {
                                    OmError::invalid_argument(
                                        "ChartObjects.Item index is out of bounds",
                                    )
                                })
                        }
                        OmValue::Text(name) => entries
                            .iter()
                            .find(|(_, chart_object_name)| {
                                chart_object_name.eq_ignore_ascii_case(name)
                            })
                            .map(|(chart_object_id, _)| *chart_object_id)
                            .ok_or_else(|| {
                                OmError::new(
                                    OmErrorCode::NotFound,
                                    format!("chart object '{name}' was not found"),
                                )
                            }),
                        _ => Err(OmError::type_mismatch(
                            "ChartObjects.Item expects a numeric index, chart object name, or array of indexes and names",
                        )),
                    }
                };
                if let OmValue::Array(array) = index {
                    if array.values.is_empty() {
                        return Err(OmError::invalid_argument(
                            "ChartObjects.Item array must not be empty",
                        ));
                    }
                    let mut selected_chart_object_ids = Vec::with_capacity(array.values.len());
                    for selector in &array.values {
                        let chart_object_id = resolve_chart_object(selector)?;
                        if selected_chart_object_ids.contains(&chart_object_id) {
                            return Err(OmError::invalid_argument(
                                "ChartObjects.Item array must not contain duplicate chart objects",
                            ));
                        }
                        selected_chart_object_ids.push(chart_object_id);
                    }
                    return Ok(OmValue::Object(self.register_shape_range_handle(
                        workbook,
                        ShapeRangeSource::ChartObjects {
                            sheet_id,
                            parent,
                            chart_object_ids: Some(selected_chart_object_ids),
                        },
                    )));
                }
                let chart_object_id = resolve_chart_object(index)?;
                Ok(OmValue::Object(
                    self.register_chart_object_handle_with_parent_origin(
                        workbook,
                        chart_object_id,
                        Some(parent),
                    ),
                ))
            }
            "Delete" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Delete does not accept arguments",
                    ));
                }
                let chart_object_ids = self
                    .chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?
                    .into_iter()
                    .map(|(chart_object_id, _)| chart_object_id)
                    .collect::<Vec<_>>();
                if self.runtime_workbook(workbook)?.read_only {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        "cannot modify a read-only workbook",
                    ));
                }
                for chart_object_id in chart_object_ids {
                    self.delete_chart_object(workbook, chart_object_id)?;
                }
                Ok(OmValue::Empty)
            }
            "Copy" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Copy does not accept arguments",
                    ));
                }
                self.chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?;
                self.set_headless_copy_mode();
                Ok(OmValue::Empty)
            }
            "Cut" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Cut does not accept arguments",
                    ));
                }
                self.chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?;
                if self.runtime_workbook(workbook)?.read_only {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        "cannot modify a read-only workbook",
                    ));
                }
                self.set_headless_cut_mode();
                Ok(OmValue::Empty)
            }
            "Duplicate" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Duplicate does not accept arguments",
                    ));
                }
                let chart_object_ids = self
                    .chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?
                    .into_iter()
                    .map(|(chart_object_id, _)| chart_object_id)
                    .collect::<Vec<_>>();
                let mut duplicated_chart_object_ids = Vec::with_capacity(chart_object_ids.len());
                for chart_object_id in chart_object_ids {
                    duplicated_chart_object_ids
                        .push(self.duplicate_chart_object(workbook, chart_object_id)?);
                }
                match duplicated_chart_object_ids.as_slice() {
                    [] => Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    )),
                    [chart_object_id] => Ok(OmValue::Object(
                        self.register_chart_object_handle_with_parent_origin(
                            workbook,
                            *chart_object_id,
                            Some(parent),
                        ),
                    )),
                    _ => Ok(OmValue::Object(self.register_shape_range_handle(
                        workbook,
                        ShapeRangeSource::ChartObjects {
                            sheet_id,
                            parent,
                            chart_object_ids: Some(duplicated_chart_object_ids),
                        },
                    ))),
                }
            }
            "CopyPicture" => {
                validate_copy_picture_args(args, 2, "ChartObjects.CopyPicture")?;
                self.chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?;
                self.set_headless_copy_mode();
                Ok(OmValue::Empty)
            }
            "BringToFront" | "SendToBack" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "ChartObjects.{member} does not accept arguments"
                    )));
                }
                let chart_object_ids = self
                    .chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?
                    .into_iter()
                    .map(|(chart_object_id, _)| chart_object_id)
                    .collect::<Vec<_>>();
                self.move_chart_objects_z_order(workbook, &chart_object_ids, member)?;
                Ok(OmValue::Empty)
            }
            "Select" => {
                if args.len() > 1 {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Select accepts at most a Replace argument",
                    ));
                }
                if let Some(value) = args.first()
                    && !om_value_is_omitted(value)
                {
                    coerce_optional_bool_arg(value, true, "ChartObjects.Select Replace")?;
                }
                let Some((chart_object_id, _)) = self
                    .chart_object_entries_for_selection(workbook, sheet_id, chart_object_ids)?
                    .into_iter()
                    .next()
                else {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                };
                let chart_object = self.chart_object_model(workbook, chart_object_id)?.clone();
                self.chart_model(workbook, chart_object.chart_id)?;
                self.ensure_worksheet_visible(workbook, sheet_id, "ChartObjects.Select")?;
                self.set_selection(workbook, sheet_id, Rect::single_cell(1, 1));
                self.active_chart = Some((workbook, chart_object.chart_id, Some(parent)));
                Ok(OmValue::Empty)
            }
            "Add" => {
                let [left, top, width, height] = args else {
                    return Err(OmError::invalid_argument(
                        "ChartObjects.Add expects Left, Top, Width, and Height arguments",
                    ));
                };
                let mut geometry = Vec::new();
                for (value, label, require_non_negative) in [
                    (left, "ChartObjects.Add Left", false),
                    (top, "ChartObjects.Add Top", false),
                    (width, "ChartObjects.Add Width", true),
                    (height, "ChartObjects.Add Height", true),
                ] {
                    let OmValue::Number(number) = value else {
                        return Err(OmError::type_mismatch(format!(
                            "{label} expects a numeric points value"
                        )));
                    };
                    if !number.is_finite()
                        || *number < i64::MIN as f64 / 12_700.0
                        || *number > i64::MAX as f64 / 12_700.0
                        || require_non_negative && *number < 0.0
                    {
                        return Err(OmError::invalid_argument(format!(
                            "{label} expects a finite points value{}",
                            if require_non_negative {
                                " greater than or equal to zero"
                            } else {
                                ""
                            }
                        )));
                    }
                    geometry.push(office_common::Points(*number).to_emu());
                }
                let default_chart_type = self.default_chart_type.clone();
                let default_chart_axes = default_chart_axes_for_type(&default_chart_type);

                let chart_object_id = {
                    let runtime = self.runtime_workbook_mut(workbook)?;
                    if runtime.read_only {
                        return Err(OmError::new(
                            OmErrorCode::InvalidState,
                            "cannot modify a read-only workbook",
                        ));
                    }
                    let Some(worksheet) = runtime
                        .loaded
                        .state
                        .worksheets()
                        .iter()
                        .find(|worksheet| worksheet.id == sheet_id)
                    else {
                        return Err(OmError::new(OmErrorCode::NotFound, "unknown worksheet"));
                    };
                    if !matches!(worksheet.kind, SheetKind::Worksheet | SheetKind::ChartSheet) {
                        return Err(OmError::unsupported(
                            "ChartObjects.Add is only available on worksheets and chart sheets",
                        ));
                    }

                    let workbook_id = runtime.loaded.state.model().id;
                    let chart_id = ChartId(
                        runtime
                            .loaded
                            .state
                            .charts
                            .keys()
                            .map(|chart_id| chart_id.0)
                            .max()
                            .unwrap_or_default()
                            + 1,
                    );
                    let chart_object_id = ChartObjectId(
                        runtime
                            .loaded
                            .state
                            .drawings
                            .values()
                            .flat_map(|drawing| drawing.objects.iter())
                            .filter_map(|object| match object {
                                DrawingObjectModel::ChartFrame(chart_object) => {
                                    Some(chart_object.id.0)
                                }
                                DrawingObjectModel::UnsupportedRaw { id, .. } => Some(id.0),
                            })
                            .max()
                            .unwrap_or_default()
                            + 1,
                    );
                    let existing_chart_object_names = runtime
                        .loaded
                        .state
                        .drawings
                        .values()
                        .filter(|drawing| drawing.host_sheet_id == sheet_id)
                        .flat_map(|drawing| drawing.objects.iter())
                        .filter_map(|object| match object {
                            DrawingObjectModel::ChartFrame(chart_object) => {
                                Some(chart_object.name.as_str())
                            }
                            DrawingObjectModel::UnsupportedRaw { .. } => None,
                        })
                        .collect::<Vec<_>>();
                    let mut chart_object_number = existing_chart_object_names.len() + 1;
                    let chart_object_name = loop {
                        let candidate = format!("Chart {chart_object_number}");
                        if !existing_chart_object_names
                            .iter()
                            .any(|name| name.eq_ignore_ascii_case(&candidate))
                        {
                            break candidate;
                        }
                        chart_object_number += 1;
                    };
                    let drawing_id = if let Some(existing_drawing_id) = runtime
                        .loaded
                        .state
                        .drawings
                        .iter()
                        .find(|(_, drawing)| drawing.host_sheet_id == sheet_id)
                        .map(|(drawing_id, _)| *drawing_id)
                    {
                        existing_drawing_id
                    } else {
                        let drawing_id = DrawingId(
                            runtime
                                .loaded
                                .state
                                .drawings
                                .keys()
                                .map(|drawing_id| drawing_id.0)
                                .max()
                                .unwrap_or_default()
                                + 1,
                        );
                        runtime.loaded.state.drawings.insert(
                            drawing_id,
                            DrawingModel {
                                id: drawing_id,
                                workbook_id,
                                host_sheet_id: sheet_id,
                                objects: Vec::new(),
                                raw_part_uri: None,
                                dirty: true,
                            },
                        );
                        drawing_id
                    };
                    runtime.loaded.state.charts.insert(
                        chart_id,
                        ChartModel {
                            id: chart_id,
                            workbook_id,
                            chart_type: default_chart_type,
                            style: None,
                            series: Vec::new(),
                            title: None,
                            legend: None,
                            axes: default_chart_axes,
                            groups: Vec::new(),
                            vary_by_categories: None,
                            gap_width: None,
                            gap_depth: None,
                            overlap: None,
                            bar_shape: None,
                            has_series_lines: None,
                            has_drop_lines: None,
                            has_hi_lo_lines: None,
                            has_up_down_bars: None,
                            first_slice_angle: None,
                            explosion: None,
                            bubble_scale: None,
                            show_negative_bubbles: None,
                            has_3d_shading: None,
                            doughnut_hole_size: None,
                            second_plot_size: None,
                            size_represents: None,
                            split_type: None,
                            split_value: None,
                            data_labels: None,
                            data_table: None,
                            data_table_dirty: false,
                            plot_area_layout: None,
                            plot_area_layout_dirty: false,
                            show_data_labels_over_maximum: None,
                            display_blanks_as: None,
                            plot_visible_only: None,
                            view_3d: None,
                            view_3d_dirty: false,
                            rounded_corners: None,
                            protection: None,
                            protection_dirty: false,
                            raw_part_uri: None,
                            series_topology_dirty: false,
                            content_dirty: false,
                            dirty: true,
                        },
                    );
                    let drawing = runtime
                        .loaded
                        .state
                        .drawings
                        .get_mut(&drawing_id)
                        .expect("drawing was inserted above");
                    let z_order = drawing
                        .objects
                        .iter()
                        .enumerate()
                        .filter_map(|(fallback_index, object)| match object {
                            DrawingObjectModel::ChartFrame(chart_object) => chart_object
                                .z_order
                                .or_else(|| u32::try_from(fallback_index).ok()),
                            DrawingObjectModel::UnsupportedRaw { .. } => {
                                u32::try_from(fallback_index).ok()
                            }
                        })
                        .max()
                        .map_or(Some(0), |existing_z_order| existing_z_order.checked_add(1));
                    drawing
                        .objects
                        .push(DrawingObjectModel::ChartFrame(ChartObjectModel {
                            id: chart_object_id,
                            anchor_attrs: BTreeMap::new(),
                            position_attrs: BTreeMap::new(),
                            extents_attrs: BTreeMap::new(),
                            marker_attrs: Default::default(),
                            graphic_frame_attrs: BTreeMap::new(),
                            graphic_frame_transform_xml: None,
                            graphic_data_attrs: BTreeMap::new(),
                            graphic_data_child_xmls: Vec::new(),
                            chart_reference_attrs: BTreeMap::new(),
                            non_visual_frame_attrs: BTreeMap::new(),
                            graphic_attrs: BTreeMap::new(),
                            non_visual_id: u32::try_from(chart_object_id.0).ok(),
                            non_visual_attrs: BTreeMap::new(),
                            non_visual_child_xml: None,
                            non_visual_frame_properties_xml: None,
                            client_data_attrs: BTreeMap::new(),
                            client_data_xml: None,
                            anchor_extension_xmls: Vec::new(),
                            workbook_id,
                            host_sheet_id: sheet_id,
                            chart_id,
                            name: chart_object_name,
                            anchor: Some(DrawingAnchor::Absolute(AbsoluteAnchor {
                                position: PointEmu {
                                    x: geometry[0],
                                    y: geometry[1],
                                },
                                extents: SizeEmu {
                                    cx: geometry[2],
                                    cy: geometry[3],
                                },
                            })),
                            placement: ObjectPlacement::FreeFloating,
                            z_order,
                            raw_binding: None,
                            dirty: true,
                        }));
                    drawing.dirty = true;
                    runtime.prompt_dirty = true;
                    self.find_state = None;
                    self.cut_copy_mode = None;
                    self.clipboard = None;
                    chart_object_id
                };

                Ok(OmValue::Object(
                    self.register_chart_object_handle_with_parent_origin(
                        workbook,
                        chart_object_id,
                        Some(parent),
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "ChartObjects.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_shape_range(
        &mut self,
        workbook: WorkbookHandle,
        source: ShapeRangeSource,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("ShapeRange", member) {
            self.focus_member_supported("ShapeRange", member, false)?;
        }

        match member {
            "Item" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Item expects a single shape index or name",
                    ));
                };
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                let chart_object_id = match index {
                    OmValue::Number(_) => {
                        let index = coerce_u32_arg(index, "ShapeRange.Item index")? as usize;
                        if index == 0 {
                            return Err(OmError::invalid_argument(
                                "ShapeRange.Item index is out of bounds",
                            ));
                        }
                        entries
                            .get(index - 1)
                            .map(|(chart_object_id, _)| *chart_object_id)
                            .ok_or_else(|| {
                                OmError::invalid_argument("ShapeRange.Item index is out of bounds")
                            })?
                    }
                    OmValue::Text(name) => entries
                        .iter()
                        .find(|(_, candidate)| candidate.eq_ignore_ascii_case(name))
                        .map(|(chart_object_id, _)| *chart_object_id)
                        .ok_or_else(|| {
                            OmError::new(
                                OmErrorCode::NotFound,
                                format!("shape '{name}' was not found"),
                            )
                        })?,
                    _ => {
                        return Err(OmError::type_mismatch(
                            "ShapeRange.Item expects a numeric index or shape name",
                        ));
                    }
                };
                Ok(OmValue::Object(self.register_shape_range_handle(
                    workbook,
                    ShapeRangeSource::ChartObject {
                        chart_object_id,
                        parent: source.parent(),
                    },
                )))
            }
            "Copy" | "Cut" | "CopyPicture" => {
                let delegated = self.shape_range_delegate_handle(workbook, &source, member)?;
                self.dispatch_invoke(delegated, member, args)
            }
            "Select" => {
                if args.len() > 1 {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Select accepts at most a Replace argument",
                    ));
                }
                if let Some(value) = args.first()
                    && !om_value_is_omitted(value)
                {
                    coerce_optional_bool_arg(value, true, "ShapeRange.Select Replace")?;
                }
                let Some((chart_object_id, _)) = self
                    .shape_range_chart_object_entries(workbook, &source)?
                    .into_iter()
                    .next()
                else {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                };
                let chart_object = self.chart_object_model(workbook, chart_object_id)?.clone();
                self.chart_model(workbook, chart_object.chart_id)?;
                self.ensure_worksheet_visible(
                    workbook,
                    chart_object.host_sheet_id,
                    "ShapeRange.Select",
                )?;
                self.set_selection(
                    workbook,
                    chart_object.host_sheet_id,
                    Rect::single_cell(1, 1),
                );
                self.active_chart = Some((workbook, chart_object.chart_id, Some(source.parent())));
                Ok(OmValue::Empty)
            }
            "Delete" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Delete does not accept arguments",
                    ));
                }
                let chart_object_ids = self
                    .shape_range_chart_object_entries(workbook, &source)?
                    .into_iter()
                    .map(|(chart_object_id, _)| chart_object_id)
                    .collect::<Vec<_>>();
                for chart_object_id in chart_object_ids {
                    self.delete_chart_object(workbook, chart_object_id)?;
                }
                Ok(OmValue::Empty)
            }
            "Duplicate" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Duplicate does not accept arguments",
                    ));
                }
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                let [(chart_object_id, _)] = entries.as_slice() else {
                    let delegated =
                        self.shape_range_delegate_handle(workbook, &source, "Duplicate")?;
                    return self.dispatch_invoke(delegated, member, args);
                };
                let duplicated_id = self.duplicate_chart_object(workbook, *chart_object_id)?;
                Ok(OmValue::Object(self.register_shape_range_handle(
                    workbook,
                    ShapeRangeSource::ChartObject {
                        chart_object_id: duplicated_id,
                        parent: source.parent(),
                    },
                )))
            }
            "IncrementLeft" | "IncrementTop" => {
                let [increment] = args else {
                    return Err(OmError::invalid_argument(format!(
                        "ShapeRange.{member} expects a single increment argument"
                    )));
                };
                let OmValue::Number(increment) = increment else {
                    return Err(OmError::type_mismatch(format!(
                        "ShapeRange.{member} expects a numeric points value"
                    )));
                };
                if !increment.is_finite()
                    || *increment < i64::MIN as f64 / 12_700.0
                    || *increment > i64::MAX as f64 / 12_700.0
                {
                    return Err(OmError::invalid_argument(format!(
                        "ShapeRange.{member} expects a finite points value"
                    )));
                }

                let geometry_member = match member {
                    "IncrementLeft" => "Left",
                    "IncrementTop" => "Top",
                    _ => unreachable!("ShapeRange increment member was matched"),
                };
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                for (chart_object_id, _) in entries {
                    let delegated = self.register_chart_object_handle(workbook, chart_object_id);
                    let current = match self.dispatch_get(delegated, geometry_member, &[])? {
                        OmValue::Number(current) => current,
                        _ => {
                            return Err(OmError::new(
                                OmErrorCode::InvalidState,
                                "ChartObject geometry did not return a numeric value",
                            ));
                        }
                    };
                    self.dispatch_set(
                        delegated,
                        geometry_member,
                        OmValue::Number(current + *increment),
                        &[],
                    )?;
                }
                Ok(OmValue::Empty)
            }
            "IncrementRotation" => {
                let [increment] = args else {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.IncrementRotation expects a single increment argument",
                    ));
                };
                let OmValue::Number(increment) = increment else {
                    return Err(OmError::type_mismatch(
                        "ShapeRange.IncrementRotation expects a numeric degrees value",
                    ));
                };
                let increment_units =
                    rotation_units_from_degrees(*increment, "ShapeRange.IncrementRotation")?;
                let chart_object_ids = self
                    .shape_range_chart_object_entries(workbook, &source)?
                    .into_iter()
                    .map(|(chart_object_id, _)| chart_object_id)
                    .collect::<Vec<_>>();
                if chart_object_ids.is_empty() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                }
                let runtime = self.runtime_workbook_mut(workbook)?;
                if runtime.read_only {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        "cannot modify a read-only workbook",
                    ));
                }
                let mut workbook_dirty = false;
                for drawing in runtime.loaded.state.drawings.values_mut() {
                    for object in &mut drawing.objects {
                        let DrawingObjectModel::ChartFrame(chart_object) = object else {
                            continue;
                        };
                        if !chart_object_ids.contains(&chart_object.id) {
                            continue;
                        }
                        let current_units = graphic_frame_rotation_units(
                            chart_object.graphic_frame_transform_xml.as_deref(),
                        )?;
                        let updated_units =
                            current_units.checked_add(increment_units).ok_or_else(|| {
                                OmError::invalid_argument(
                                    "ShapeRange.IncrementRotation result is out of range",
                                )
                            })?;
                        let updated_transform = set_graphic_frame_rotation_xml(
                            chart_object.graphic_frame_transform_xml.as_deref(),
                            updated_units,
                        )?;
                        if chart_object.graphic_frame_transform_xml != updated_transform {
                            chart_object.graphic_frame_transform_xml = updated_transform;
                            chart_object.dirty = true;
                            drawing.dirty = true;
                            workbook_dirty = true;
                        }
                    }
                }
                if workbook_dirty {
                    runtime.prompt_dirty = true;
                    self.find_state = None;
                    self.cut_copy_mode = None;
                    self.clipboard = None;
                }
                Ok(OmValue::Empty)
            }
            "ScaleWidth" | "ScaleHeight" => {
                if args.len() < 2 || args.len() > 3 {
                    return Err(OmError::invalid_argument(format!(
                        "ShapeRange.{member} expects Factor, RelativeToOriginalSize, and optional Scale arguments"
                    )));
                }
                let OmValue::Number(factor) = &args[0] else {
                    return Err(OmError::type_mismatch(format!(
                        "ShapeRange.{member} Factor must be numeric"
                    )));
                };
                if !factor.is_finite() || *factor < 0.0 || *factor > i64::MAX as f64 / 12_700.0 {
                    return Err(OmError::invalid_argument(format!(
                        "ShapeRange.{member} Factor must be a finite non-negative value"
                    )));
                }
                match &args[1] {
                    OmValue::Missing | OmValue::Empty | OmValue::Null | OmValue::Bool(_) => {}
                    OmValue::Number(value) => {
                        if !value.is_finite()
                            || value.fract() != 0.0
                            || !matches!(*value as i32, MSO_TRUE | MSO_FALSE)
                        {
                            return Err(OmError::invalid_argument(format!(
                                "ShapeRange.{member} RelativeToOriginalSize must be msoTrue or msoFalse"
                            )));
                        }
                    }
                    _ => {
                        return Err(OmError::type_mismatch(format!(
                            "ShapeRange.{member} RelativeToOriginalSize must be boolean or MsoTriState"
                        )));
                    }
                }
                let scale_from = match args.get(2) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => {
                        MSO_SCALE_FROM_TOP_LEFT
                    }
                    Some(OmValue::Number(value)) => {
                        if !value.is_finite()
                            || value.fract() != 0.0
                            || *value < i32::MIN as f64
                            || *value > i32::MAX as f64
                        {
                            return Err(OmError::invalid_argument(format!(
                                "ShapeRange.{member} Scale must be an integral MsoScaleFrom value"
                            )));
                        }
                        match *value as i32 {
                            MSO_SCALE_FROM_TOP_LEFT
                            | MSO_SCALE_FROM_MIDDLE
                            | MSO_SCALE_FROM_BOTTOM_RIGHT => *value as i32,
                            _ => {
                                return Err(OmError::invalid_argument(format!(
                                    "ShapeRange.{member} Scale must be msoScaleFromTopLeft, msoScaleFromMiddle, or msoScaleFromBottomRight"
                                )));
                            }
                        }
                    }
                    Some(_) => {
                        return Err(OmError::type_mismatch(format!(
                            "ShapeRange.{member} Scale must be numeric when provided"
                        )));
                    }
                };

                let (position_member, size_member) = match member {
                    "ScaleWidth" => ("Left", "Width"),
                    "ScaleHeight" => ("Top", "Height"),
                    _ => unreachable!("ShapeRange scale member was matched"),
                };
                let shape_range = self.register_shape_range_handle(workbook, source);
                let current_position = match self.dispatch_get(shape_range, position_member, &[])? {
                    OmValue::Number(value) => value,
                    _ => {
                        return Err(OmError::new(
                            OmErrorCode::InvalidState,
                            "ShapeRange position did not return a numeric value",
                        ));
                    }
                };
                let current_size = match self.dispatch_get(shape_range, size_member, &[])? {
                    OmValue::Number(value) => value,
                    _ => {
                        return Err(OmError::new(
                            OmErrorCode::InvalidState,
                            "ShapeRange size did not return a numeric value",
                        ));
                    }
                };
                let new_size = current_size * *factor;
                if !new_size.is_finite() || new_size > i64::MAX as f64 / 12_700.0 {
                    return Err(OmError::invalid_argument(format!(
                        "ShapeRange.{member} scaled size is out of range"
                    )));
                }
                let new_position = match scale_from {
                    MSO_SCALE_FROM_TOP_LEFT => current_position,
                    MSO_SCALE_FROM_MIDDLE => current_position + (current_size - new_size) / 2.0,
                    MSO_SCALE_FROM_BOTTOM_RIGHT => current_position + current_size - new_size,
                    _ => unreachable!("validated MsoScaleFrom value"),
                };
                if new_position != current_position {
                    self.dispatch_set(
                        shape_range,
                        position_member,
                        OmValue::Number(new_position),
                        &[],
                    )?;
                }
                if new_size != current_size {
                    self.dispatch_set(shape_range, size_member, OmValue::Number(new_size), &[])?;
                }
                Ok(OmValue::Empty)
            }
            "Flip" => {
                let [command] = args else {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Flip expects a single MsoFlipCmd argument",
                    ));
                };
                let OmValue::Number(command) = command else {
                    return Err(OmError::type_mismatch(
                        "ShapeRange.Flip expects a numeric MsoFlipCmd argument",
                    ));
                };
                if !command.is_finite()
                    || command.fract() != 0.0
                    || *command < i32::MIN as f64
                    || *command > i32::MAX as f64
                {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Flip expects an integral MsoFlipCmd argument",
                    ));
                }
                let (attr_name, attr_name_text) = match *command as i32 {
                    MSO_FLIP_HORIZONTAL => (b"flipH".as_slice(), "flipH"),
                    MSO_FLIP_VERTICAL => (b"flipV".as_slice(), "flipV"),
                    _ => {
                        return Err(OmError::invalid_argument(
                            "ShapeRange.Flip expects msoFlipHorizontal or msoFlipVertical",
                        ));
                    }
                };
                let chart_object_ids = self
                    .shape_range_chart_object_entries(workbook, &source)?
                    .into_iter()
                    .map(|(chart_object_id, _)| chart_object_id)
                    .collect::<Vec<_>>();
                if chart_object_ids.is_empty() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                }
                let runtime = self.runtime_workbook_mut(workbook)?;
                if runtime.read_only {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        "cannot modify a read-only workbook",
                    ));
                }
                let mut workbook_dirty = false;
                for drawing in runtime.loaded.state.drawings.values_mut() {
                    for object in &mut drawing.objects {
                        let DrawingObjectModel::ChartFrame(chart_object) = object else {
                            continue;
                        };
                        if !chart_object_ids.contains(&chart_object.id) {
                            continue;
                        }
                        let current = graphic_frame_transform_bool_attr(
                            chart_object.graphic_frame_transform_xml.as_deref(),
                            attr_name,
                        )?;
                        let updated_transform = set_graphic_frame_transform_bool_attr_xml(
                            chart_object.graphic_frame_transform_xml.as_deref(),
                            attr_name,
                            attr_name_text,
                            !current,
                        )?;
                        if chart_object.graphic_frame_transform_xml != updated_transform {
                            chart_object.graphic_frame_transform_xml = updated_transform;
                            chart_object.dirty = true;
                            drawing.dirty = true;
                            workbook_dirty = true;
                        }
                    }
                }
                if workbook_dirty {
                    runtime.prompt_dirty = true;
                    self.find_state = None;
                    self.cut_copy_mode = None;
                    self.clipboard = None;
                }
                Ok(OmValue::Empty)
            }
            "ZOrder" => {
                let [command] = args else {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.ZOrder expects a single MsoZOrderCmd argument",
                    ));
                };
                let operation = chart_object_z_order_operation(command, "ShapeRange")?;
                let chart_object_ids = self
                    .shape_range_chart_object_entries(workbook, &source)?
                    .into_iter()
                    .map(|(chart_object_id, _)| chart_object_id)
                    .collect::<Vec<_>>();
                self.move_chart_objects_z_order(workbook, &chart_object_ids, operation)?;
                Ok(OmValue::Empty)
            }
            _ => Err(OmError::unsupported(format!(
                "ShapeRange.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_get_chart_object(
        &mut self,
        workbook: WorkbookHandle,
        chart_object_id: ChartObjectId,
        parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if matches!(
            member,
            "Name"
                | "Chart"
                | "Index"
                | "ZOrder"
                | "Placement"
                | "Left"
                | "Top"
                | "Width"
                | "Height"
                | "Visible"
                | "OnAction"
                | "PrintObject"
                | "Locked"
                | "ProtectChartObject"
                | "RoundedCorners"
                | "ShapeRange"
                | "TopLeftCell"
                | "BottomRightCell"
                | "Creator"
                | "Application"
                | "Parent"
        ) {
            self.focus_member_supported("ChartObject", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "ChartObject.{member} does not accept arguments"
            )));
        }

        match member {
            "Name" => Ok(OmValue::Text(
                self.chart_object_model(workbook, chart_object_id)?
                    .name
                    .clone(),
            )),
            "Chart" => {
                let chart_object = self.chart_object_model(workbook, chart_object_id)?;
                let chart_id = chart_object.chart_id;
                let host_sheet_id = chart_object.host_sheet_id;
                let parent = parent.unwrap_or(ChartObjectsParent::Worksheet(host_sheet_id));
                Ok(OmValue::Object(
                    self.register_chart_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        Some(parent),
                    ),
                ))
            }
            "ShapeRange" => {
                let host_sheet_id = self
                    .chart_object_model(workbook, chart_object_id)?
                    .host_sheet_id;
                let parent = parent.unwrap_or(ChartObjectsParent::Worksheet(host_sheet_id));
                Ok(OmValue::Object(self.register_shape_range_handle(
                    workbook,
                    ShapeRangeSource::ChartObject {
                        chart_object_id,
                        parent,
                    },
                )))
            }
            "Index" | "ZOrder" => {
                let sheet_id = self
                    .chart_object_model(workbook, chart_object_id)?
                    .host_sheet_id;
                let index = self
                    .chart_object_entries_for_sheet(workbook, sheet_id)?
                    .iter()
                    .position(|(candidate_id, _)| *candidate_id == chart_object_id)
                    .map(|index| index + 1)
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart object not found"))?;
                Ok(OmValue::Number(index as f64))
            }
            "Placement" => Ok(OmValue::Number(f64::from(chart_object_placement_value(
                &self
                    .chart_object_model(workbook, chart_object_id)?
                    .placement,
            )?))),
            "Visible" => {
                let visible = !self
                    .chart_object_model(workbook, chart_object_id)?
                    .non_visual_attrs
                    .get("hidden")
                    .is_some_and(|value| {
                        value == "1"
                            || value.eq_ignore_ascii_case("true")
                            || value.eq_ignore_ascii_case("on")
                    });
                Ok(OmValue::Bool(visible))
            }
            "OnAction" => Ok(OmValue::Text(
                self.chart_object_model(workbook, chart_object_id)?
                    .graphic_frame_attrs
                    .get("macro")
                    .cloned()
                    .unwrap_or_default(),
            )),
            "PrintObject" => {
                let print_object = self
                    .chart_object_model(workbook, chart_object_id)?
                    .client_data_attrs
                    .get("fPrintsWithSheet")
                    .map_or(true, |value| {
                        !(value == "0"
                            || value.eq_ignore_ascii_case("false")
                            || value.eq_ignore_ascii_case("off"))
                    });
                Ok(OmValue::Bool(print_object))
            }
            "Locked" | "ProtectChartObject" => {
                let locked = self
                    .chart_object_model(workbook, chart_object_id)?
                    .client_data_attrs
                    .get("fLocksWithSheet")
                    .map_or(true, |value| {
                        !(value == "0"
                            || value.eq_ignore_ascii_case("false")
                            || value.eq_ignore_ascii_case("off"))
                    });
                Ok(OmValue::Bool(locked))
            }
            "RoundedCorners" => {
                let chart_id = self.chart_object_model(workbook, chart_object_id)?.chart_id;
                Ok(OmValue::Bool(
                    self.chart_model(workbook, chart_id)?
                        .rounded_corners
                        .unwrap_or(false),
                ))
            }
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Left" | "Top" | "Width" | "Height" => {
                Ok(OmValue::Number(Self::chart_object_geometry_value(
                    self.chart_object_model(workbook, chart_object_id)?,
                    member,
                )?))
            }
            "TopLeftCell" | "BottomRightCell" => {
                let chart_object = self.chart_object_model(workbook, chart_object_id)?.clone();
                let Some(anchor) = chart_object.anchor.as_ref() else {
                    return Err(OmError::unsupported(format!(
                        "ChartObject.{member} is unavailable for unsupported drawing anchors"
                    )));
                };
                let marker = match (member, anchor) {
                    ("TopLeftCell", DrawingAnchor::OneCell(anchor)) => &anchor.from,
                    ("TopLeftCell", DrawingAnchor::TwoCell(anchor)) => &anchor.from,
                    ("BottomRightCell", DrawingAnchor::TwoCell(anchor)) => &anchor.to,
                    (_, DrawingAnchor::Absolute(_) | DrawingAnchor::UnsupportedRaw)
                    | ("BottomRightCell", DrawingAnchor::OneCell(_)) => {
                        return Err(OmError::unsupported(format!(
                            "ChartObject.{member} is unavailable for this drawing anchor"
                        )));
                    }
                    _ => {
                        return Err(OmError::unsupported(format!(
                            "ChartObject.{member} is not supported"
                        )));
                    }
                };
                let row = marker
                    .row_zero_based
                    .checked_add(1)
                    .ok_or_else(|| OmError::invalid_argument("chart marker row is out of range"))?;
                let col = marker.col_zero_based.checked_add(1).ok_or_else(|| {
                    OmError::invalid_argument("chart marker column is out of range")
                })?;
                let workbook_id = self.workbook_model(workbook)?.id;
                let range = RangeSet::single_rect(
                    workbook_id,
                    chart_object.host_sheet_id,
                    Rect::single_cell(row, col),
                )?;
                Ok(OmValue::Object(
                    self.register_range_set_handle(workbook, range).0,
                ))
            }
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => {
                let sheet_id = self
                    .chart_object_model(workbook, chart_object_id)?
                    .host_sheet_id;
                let parent = parent.unwrap_or(ChartObjectsParent::Worksheet(sheet_id));
                Ok(OmValue::Object(match parent {
                    ChartObjectsParent::Worksheet(parent_sheet_id) => {
                        self.register_worksheet_handle(workbook, parent_sheet_id).0
                    }
                    ChartObjectsParent::Chart(parent_chart_id) => {
                        self.register_chart_handle(workbook, parent_chart_id)
                    }
                }))
            }
            _ => Err(OmError::unsupported(format!(
                "ChartObject.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_shape_range(
        &mut self,
        workbook: WorkbookHandle,
        source: ShapeRangeSource,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("ShapeRange", member) {
            self.focus_member_supported("ShapeRange", member, false)?;
        }

        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Count does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(
                    self.shape_range_chart_object_entries(workbook, &source)?
                        .len() as f64,
                ))
            }
            "Type" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Type does not accept arguments",
                    ));
                }
                if self
                    .shape_range_chart_object_entries(workbook, &source)?
                    .is_empty()
                {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                }
                Ok(OmValue::Number(f64::from(MSO_SHAPE_CHART)))
            }
            "ID" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.ID does not accept arguments",
                    ));
                }
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                let [(chart_object_id, _)] = entries.as_slice() else {
                    return Ok(OmValue::Number(f64::from(MSO_SHAPE_MIXED)));
                };
                let chart_object = self.chart_object_model(workbook, *chart_object_id)?;
                let id = chart_object
                    .non_visual_id
                    .map(f64::from)
                    .unwrap_or(chart_object.id.0 as f64);
                Ok(OmValue::Number(id))
            }
            "AlternativeText" | "Title" => {
                let (attr_name, property_name) = match member {
                    "AlternativeText" => ("descr", "AlternativeText"),
                    "Title" => ("title", "Title"),
                    _ => unreachable!("ShapeRange non-visual text member was matched"),
                };
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "ShapeRange.{property_name} does not accept arguments"
                    )));
                }
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                if entries.is_empty() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                }
                let mut shared_alternative_text: Option<Option<String>> = None;
                for (chart_object_id, _) in entries {
                    let alternative_text = self
                        .chart_object_model(workbook, chart_object_id)?
                        .non_visual_attrs
                        .get(attr_name)
                        .cloned();
                    match &shared_alternative_text {
                        None => shared_alternative_text = Some(alternative_text),
                        Some(shared) if *shared == alternative_text => {}
                        Some(_) => return Ok(OmValue::Text(String::new())),
                    }
                }
                Ok(OmValue::Text(
                    shared_alternative_text.flatten().unwrap_or_default(),
                ))
            }
            "HasChart" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.HasChart does not accept arguments",
                    ));
                }
                let has_chart = !self
                    .shape_range_chart_object_entries(workbook, &source)?
                    .is_empty();
                Ok(OmValue::Number(f64::from(if has_chart {
                    MSO_TRUE
                } else {
                    MSO_FALSE
                })))
            }
            "HasSmartArt" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.HasSmartArt does not accept arguments",
                    ));
                }
                self.shape_range_chart_object_entries(workbook, &source)?;
                Ok(OmValue::Number(f64::from(MSO_FALSE)))
            }
            "AutoShapeType" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.AutoShapeType does not accept arguments",
                    ));
                }
                if self
                    .shape_range_chart_object_entries(workbook, &source)?
                    .is_empty()
                {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                }
                Ok(OmValue::Number(f64::from(MSO_SHAPE_MIXED)))
            }
            "LockAspectRatio" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.LockAspectRatio does not accept arguments",
                    ));
                }
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                if entries.is_empty() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                }
                let mut shared_value: Option<i32> = None;
                for (chart_object_id, _) in entries {
                    let chart_object = self.chart_object_model(workbook, chart_object_id)?;
                    let value = if graphic_frame_lock_aspect_ratio(
                        chart_object.non_visual_frame_properties_xml.as_deref(),
                    )? {
                        MSO_TRUE
                    } else {
                        MSO_FALSE
                    };
                    match shared_value {
                        None => shared_value = Some(value),
                        Some(shared) if shared == value => {}
                        Some(_) => return Ok(OmValue::Number(f64::from(MSO_SHAPE_MIXED))),
                    }
                }
                Ok(OmValue::Number(f64::from(
                    shared_value.unwrap_or(MSO_FALSE),
                )))
            }
            "Rotation" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Rotation does not accept arguments",
                    ));
                }
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                if entries.is_empty() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                }
                let mut shared_units: Option<i64> = None;
                for (chart_object_id, _) in entries {
                    let chart_object = self.chart_object_model(workbook, chart_object_id)?;
                    let units = graphic_frame_rotation_units(
                        chart_object.graphic_frame_transform_xml.as_deref(),
                    )?;
                    match shared_units {
                        None => shared_units = Some(units),
                        Some(shared) if shared == units => {}
                        Some(_) => return Ok(OmValue::Number(f64::from(MSO_SHAPE_MIXED))),
                    }
                }
                Ok(OmValue::Number(
                    shared_units.unwrap_or_default() as f64 / 60_000.0,
                ))
            }
            "HorizontalFlip" | "VerticalFlip" => {
                let (attr_name, property_name) = match member {
                    "HorizontalFlip" => (b"flipH".as_slice(), "HorizontalFlip"),
                    "VerticalFlip" => (b"flipV".as_slice(), "VerticalFlip"),
                    _ => unreachable!("ShapeRange flip member was matched"),
                };
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "ShapeRange.{property_name} does not accept arguments"
                    )));
                }
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                if entries.is_empty() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                }
                let mut shared_value: Option<i32> = None;
                for (chart_object_id, _) in entries {
                    let chart_object = self.chart_object_model(workbook, chart_object_id)?;
                    let value = if graphic_frame_transform_bool_attr(
                        chart_object.graphic_frame_transform_xml.as_deref(),
                        attr_name,
                    )? {
                        MSO_TRUE
                    } else {
                        MSO_FALSE
                    };
                    match shared_value {
                        None => shared_value = Some(value),
                        Some(shared) if shared == value => {}
                        Some(_) => return Ok(OmValue::Number(f64::from(MSO_SHAPE_MIXED))),
                    }
                }
                Ok(OmValue::Number(f64::from(
                    shared_value.unwrap_or(MSO_FALSE),
                )))
            }
            "OnAction" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.OnAction does not accept arguments",
                    ));
                }
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                if entries.is_empty() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                }
                let mut shared_on_action: Option<String> = None;
                for (chart_object_id, _) in entries {
                    let on_action = self
                        .chart_object_model(workbook, chart_object_id)?
                        .graphic_frame_attrs
                        .get("macro")
                        .cloned()
                        .unwrap_or_default();
                    match &shared_on_action {
                        None => shared_on_action = Some(on_action),
                        Some(shared) if *shared == on_action => {}
                        Some(_) => return Ok(OmValue::Text(String::new())),
                    }
                }
                Ok(OmValue::Text(shared_on_action.unwrap_or_default()))
            }
            "TopLeftCell" | "BottomRightCell" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "ShapeRange.{member} does not accept arguments"
                    )));
                }
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                if entries.is_empty() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart object not found",
                    ));
                }
                let mut host_sheet_id = None;
                let mut marker_row_zero_based = None;
                let mut marker_col_zero_based = None;
                for (chart_object_id, _) in entries {
                    let chart_object = self.chart_object_model(workbook, chart_object_id)?;
                    match host_sheet_id {
                        None => host_sheet_id = Some(chart_object.host_sheet_id),
                        Some(sheet_id) if sheet_id == chart_object.host_sheet_id => {}
                        Some(_) => {
                            return Err(OmError::unsupported(format!(
                                "ShapeRange.{member} is unavailable for mixed-sheet ranges"
                            )));
                        }
                    }
                    let Some(anchor) = chart_object.anchor.as_ref() else {
                        return Err(OmError::unsupported(format!(
                            "ShapeRange.{member} is unavailable for unsupported drawing anchors"
                        )));
                    };
                    let marker = match (member, anchor) {
                        ("TopLeftCell", DrawingAnchor::OneCell(anchor)) => anchor.from,
                        ("TopLeftCell", DrawingAnchor::TwoCell(anchor)) => anchor.from,
                        ("BottomRightCell", DrawingAnchor::TwoCell(anchor)) => anchor.to,
                        (_, DrawingAnchor::Absolute(_) | DrawingAnchor::UnsupportedRaw)
                        | ("BottomRightCell", DrawingAnchor::OneCell(_)) => {
                            return Err(OmError::unsupported(format!(
                                "ShapeRange.{member} is unavailable for this drawing anchor"
                            )));
                        }
                        _ => {
                            return Err(OmError::unsupported(format!(
                                "ShapeRange.{member} is not supported"
                            )));
                        }
                    };
                    match member {
                        "TopLeftCell" => {
                            marker_row_zero_based = Some(
                                marker_row_zero_based.map_or(marker.row_zero_based, |row: u32| {
                                    row.min(marker.row_zero_based)
                                }),
                            );
                            marker_col_zero_based = Some(
                                marker_col_zero_based.map_or(marker.col_zero_based, |col: u32| {
                                    col.min(marker.col_zero_based)
                                }),
                            );
                        }
                        "BottomRightCell" => {
                            marker_row_zero_based = Some(
                                marker_row_zero_based.map_or(marker.row_zero_based, |row: u32| {
                                    row.max(marker.row_zero_based)
                                }),
                            );
                            marker_col_zero_based = Some(
                                marker_col_zero_based.map_or(marker.col_zero_based, |col: u32| {
                                    col.max(marker.col_zero_based)
                                }),
                            );
                        }
                        _ => unreachable!("ShapeRange anchor cell member was matched"),
                    }
                }
                let row = marker_row_zero_based
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart object not found"))?
                    .checked_add(1)
                    .ok_or_else(|| OmError::invalid_argument("chart marker row is out of range"))?;
                let col = marker_col_zero_based
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart object not found"))?
                    .checked_add(1)
                    .ok_or_else(|| {
                        OmError::invalid_argument("chart marker column is out of range")
                    })?;
                let workbook_id = self.workbook_model(workbook)?.id;
                let range = RangeSet::single_rect(
                    workbook_id,
                    host_sheet_id.ok_or_else(|| {
                        OmError::new(OmErrorCode::NotFound, "chart object not found")
                    })?,
                    Rect::single_cell(row, col),
                )?;
                Ok(OmValue::Object(
                    self.register_range_set_handle(workbook, range).0,
                ))
            }
            "Item" => self.dispatch_invoke_shape_range(workbook, source, member, args),
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Application does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Creator does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.Parent does not accept arguments",
                    ));
                }
                let parent_handle = match source {
                    ShapeRangeSource::ChartObjects { parent, .. } => match parent {
                        ChartObjectsParent::Worksheet(parent_sheet_id) => {
                            self.register_worksheet_handle(workbook, parent_sheet_id).0
                        }
                        ChartObjectsParent::Chart(parent_chart_id) => {
                            self.register_chart_handle(workbook, parent_chart_id)
                        }
                    },
                    ShapeRangeSource::ChartObject { parent, .. } => match parent {
                        ChartObjectsParent::Worksheet(parent_sheet_id) => {
                            self.register_worksheet_handle(workbook, parent_sheet_id).0
                        }
                        ChartObjectsParent::Chart(parent_chart_id) => {
                            self.register_chart_handle(workbook, parent_chart_id)
                        }
                    },
                };
                Ok(OmValue::Object(parent_handle))
            }
            "ZOrderPosition" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ShapeRange.ZOrderPosition does not accept arguments",
                    ));
                }
                let entries = self.shape_range_chart_object_entries(workbook, &source)?;
                let [(chart_object_id, _)] = entries.as_slice() else {
                    return Ok(OmValue::Number(f64::from(MSO_SHAPE_MIXED)));
                };
                let sheet_id = self
                    .chart_object_model(workbook, *chart_object_id)?
                    .host_sheet_id;
                let position = self
                    .chart_object_entries_for_sheet(workbook, sheet_id)?
                    .iter()
                    .position(|(candidate_id, _)| candidate_id == chart_object_id)
                    .map(|index| index + 1)
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart object not found"))?;
                Ok(OmValue::Number(position as f64))
            }
            _ => {
                let delegated = self.shape_range_delegate_handle(workbook, &source, member)?;
                self.dispatch_get(delegated, member, args)
            }
        }
    }

    pub(crate) fn dispatch_get_chart(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("Chart", member) {
            self.focus_member_supported("Chart", member, false)?;
        }

        match member {
            "Name" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.Name does not accept arguments",
                    ));
                }
                let state = &self.runtime_workbook(workbook)?.loaded.state;
                if let Some(worksheet) = state.worksheets().iter().find(|worksheet| {
                    state
                        .chart_sheets
                        .get(&worksheet.id)
                        .is_some_and(|binding| binding.chart_id == chart_id)
                }) {
                    return Ok(OmValue::Text(worksheet.name.clone()));
                }
                if let Some((_, _, name)) =
                    self.embedded_chart_object_for_chart(workbook, chart_id)?
                {
                    return Ok(OmValue::Text(name));
                }
                Err(OmError::new(OmErrorCode::NotFound, "chart not found"))
            }
            "ChartType" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.ChartType does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(f64::from(chart_type_to_excel_value(
                    &self.chart_model(workbook, chart_id)?.chart_type,
                )?)))
            }
            "ChartStyle" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.ChartStyle does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(f64::from(
                    self.chart_model(workbook, chart_id)?.style.unwrap_or(0),
                )))
            }
            "BarShape" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.BarShape does not accept arguments",
                    ));
                }
                let chart = self.chart_model(workbook, chart_id)?;
                ensure_chart_supports_bar_shape(&chart.chart_type)?;
                Ok(OmValue::Number(f64::from(
                    chart_effective_bar_shape(chart)
                        .map(chart_bar_shape_to_excel_value)
                        .unwrap_or(XL_BOX),
                )))
            }
            "Elevation" | "HeightPercent" | "Rotation" | "DepthPercent" | "Perspective" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "Chart.{member} does not accept arguments"
                    )));
                }
                let chart = self.chart_model(workbook, chart_id)?;
                ensure_chart_supports_3d_view(&chart.chart_type, member)?;
                let view_3d = chart.view_3d.unwrap_or_default();
                let value = match member {
                    "Elevation" => i32::from(view_3d.elevation.unwrap_or(15)),
                    "HeightPercent" => i32::from(view_3d.height_percent.unwrap_or(100)),
                    "Rotation" => i32::from(view_3d.rotation.unwrap_or(20)),
                    "DepthPercent" => i32::from(view_3d.depth_percent.unwrap_or(100)),
                    "Perspective" => i32::from(view_3d.perspective.unwrap_or(30)),
                    _ => unreachable!("checked chart 3D view numeric getter"),
                };
                Ok(OmValue::Number(f64::from(value)))
            }
            "RightAngleAxes" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.RightAngleAxes does not accept arguments",
                    ));
                }
                let chart = self.chart_model(workbook, chart_id)?;
                ensure_chart_supports_right_angle_axes(&chart.chart_type)?;
                Ok(OmValue::Bool(
                    chart
                        .view_3d
                        .and_then(|view_3d| view_3d.right_angle_axes)
                        .unwrap_or(true),
                ))
            }
            "GapDepth" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.GapDepth does not accept arguments",
                    ));
                }
                let chart = self.chart_model(workbook, chart_id)?;
                ensure_chart_supports_gap_depth(&chart.chart_type)?;
                Ok(OmValue::Number(f64::from(chart.gap_depth.unwrap_or(150))))
            }
            "DisplayBlanksAs" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.DisplayBlanksAs does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(f64::from(
                    chart_display_blanks_as_to_excel_value(
                        self.chart_model(workbook, chart_id)?
                            .display_blanks_as
                            .unwrap_or(ChartDisplayBlanksAs::Gap),
                    ),
                )))
            }
            "PlotVisibleOnly" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.PlotVisibleOnly does not accept arguments",
                    ));
                }
                Ok(OmValue::Bool(
                    self.chart_model(workbook, chart_id)?
                        .plot_visible_only
                        .unwrap_or(true),
                ))
            }
            "ShowDataLabelsOverMaximum" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.ShowDataLabelsOverMaximum does not accept arguments",
                    ));
                }
                Ok(OmValue::Bool(
                    self.chart_model(workbook, chart_id)?
                        .show_data_labels_over_maximum
                        .unwrap_or(false),
                ))
            }
            "ProtectContents"
            | "ProtectDrawingObjects"
            | "ProtectData"
            | "ProtectFormatting"
            | "ProtectSelection" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "Chart.{member} does not accept arguments"
                    )));
                }
                let protection = self
                    .chart_model(workbook, chart_id)?
                    .protection
                    .unwrap_or_default();
                Ok(OmValue::Bool(match member {
                    "ProtectContents" => protection.contents,
                    "ProtectDrawingObjects" => protection.drawing_objects,
                    "ProtectData" => protection.data,
                    "ProtectFormatting" => protection.formatting,
                    "ProtectSelection" => protection.selection,
                    _ => unreachable!("chart protection getter was matched"),
                }))
            }
            "ProtectionMode" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.ProtectionMode does not accept arguments",
                    ));
                }
                Ok(OmValue::Bool(
                    self.chart_model(workbook, chart_id)?
                        .protection
                        .unwrap_or_default()
                        .user_interface_only,
                ))
            }
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.Creator does not accept arguments",
                    ));
                }
                self.chart_model(workbook, chart_id)?;
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "Index" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.Index does not accept arguments",
                    ));
                }
                self.chart_model(workbook, chart_id)?;
                let (chart_sheet_index, embedded_chart_object) = {
                    let state = &self.runtime_workbook(workbook)?.loaded.state;
                    let chart_sheet_index = state
                        .worksheets()
                        .iter()
                        .filter(|worksheet| worksheet.kind == SheetKind::ChartSheet)
                        .position(|worksheet| {
                            state
                                .chart_sheets
                                .get(&worksheet.id)
                                .is_some_and(|binding| binding.chart_id == chart_id)
                        })
                        .map(|index| index + 1);
                    let embedded_chart_object = self
                        .embedded_chart_object_for_chart(workbook, chart_id)?
                        .map(|(host_sheet_id, chart_object_id, _)| {
                            (host_sheet_id, chart_object_id)
                        });
                    (chart_sheet_index, embedded_chart_object)
                };
                if let Some(index) = chart_sheet_index {
                    return Ok(OmValue::Number(index as f64));
                }
                if let Some((host_sheet_id, chart_object_id)) = embedded_chart_object {
                    let index = self
                        .chart_object_entries_for_sheet(workbook, host_sheet_id)?
                        .iter()
                        .position(|(candidate_id, _)| *candidate_id == chart_object_id)
                        .map(|index| index + 1)
                        .ok_or_else(|| {
                            OmError::new(OmErrorCode::NotFound, "chart object not found")
                        })?;
                    return Ok(OmValue::Number(index as f64));
                }
                Err(OmError::new(OmErrorCode::NotFound, "chart not found"))
            }
            "Next" | "Previous" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "Chart.{member} does not accept arguments"
                    )));
                }
                self.chart_model(workbook, chart_id)?;
                let adjacent_sheet_id = {
                    let state = &self.runtime_workbook(workbook)?.loaded.state;
                    let base_sheet_id = state
                        .chart_sheets
                        .iter()
                        .find_map(|(sheet_id, binding)| {
                            (binding.chart_id == chart_id).then_some(*sheet_id)
                        })
                        .or(self
                            .embedded_chart_object_for_chart(workbook, chart_id)?
                            .map(|(host_sheet_id, _, _)| host_sheet_id))
                        .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart not found"))?;
                    let index = state
                        .worksheets()
                        .iter()
                        .position(|worksheet| worksheet.id == base_sheet_id)
                        .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "unknown worksheet"))?;
                    if member == "Next" {
                        state
                            .worksheets()
                            .get(index + 1)
                            .map(|worksheet| worksheet.id)
                    } else if index == 0 {
                        None
                    } else {
                        state
                            .worksheets()
                            .get(index - 1)
                            .map(|worksheet| worksheet.id)
                    }
                };
                Ok(adjacent_sheet_id
                    .map(|sheet_id| self.register_sheet_object_handle(workbook, sheet_id))
                    .transpose()?
                    .map(OmValue::Object)
                    .unwrap_or(OmValue::Empty))
            }
            "Visible" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.Visible does not accept arguments",
                    ));
                }
                if let Some(sheet_id) = self.chart_sheet_id_for_chart(workbook, chart_id)? {
                    let sheet_handle = self.register_worksheet_handle(workbook, sheet_id).0;
                    return self.dispatch_get(sheet_handle, "Visible", &[]);
                }
                if let Some((_, chart_object_id, _)) =
                    self.embedded_chart_object_for_chart(workbook, chart_id)?
                {
                    let chart_object_handle =
                        self.register_chart_object_handle(workbook, chart_object_id);
                    return self.dispatch_get(chart_object_handle, "Visible", &[]);
                }
                Err(OmError::new(OmErrorCode::NotFound, "chart not found"))
            }
            "ChartArea" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.ChartArea does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.register_object(
                    RuntimeObjectKind::ChartArea {
                        workbook,
                        chart_id,
                        chart_object_parent,
                    },
                )))
            }
            "PlotArea" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.PlotArea does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.register_object(
                    RuntimeObjectKind::PlotArea {
                        workbook,
                        chart_id,
                        chart_object_parent,
                    },
                )))
            }
            "HasTitle" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.HasTitle does not accept arguments",
                    ));
                }
                Ok(OmValue::Bool(
                    self.chart_model(workbook, chart_id)?.title.is_some(),
                ))
            }
            "ChartTitle" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.ChartTitle does not accept arguments",
                    ));
                }
                if self.chart_model(workbook, chart_id)?.title.is_none() {
                    return Err(OmError::new(OmErrorCode::NotFound, "chart title not found"));
                }
                Ok(OmValue::Object(
                    self.register_chart_title_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        chart_object_parent,
                    ),
                ))
            }
            "HasDataTable" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.HasDataTable does not accept arguments",
                    ));
                }
                Ok(OmValue::Bool(
                    self.chart_model(workbook, chart_id)?.data_table.is_some(),
                ))
            }
            "DataTable" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.DataTable does not accept arguments",
                    ));
                }
                if self.chart_model(workbook, chart_id)?.data_table.is_none() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart data table not found",
                    ));
                }
                Ok(OmValue::Object(
                    self.register_data_table_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        chart_object_parent,
                    ),
                ))
            }
            "HasLegend" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.HasLegend does not accept arguments",
                    ));
                }
                Ok(OmValue::Bool(
                    self.chart_model(workbook, chart_id)?
                        .legend
                        .as_ref()
                        .is_some_and(|legend| legend.visible),
                ))
            }
            "HasAxis" => {
                if args.is_empty() || args.len() > 2 {
                    return Err(OmError::invalid_argument(
                        "Chart.HasAxis expects axis type and optional axis group",
                    ));
                }
                let axis_type = coerce_u32_arg(&args[0], "Chart.HasAxis axis type")? as i32;
                if !matches!(axis_type, XL_CATEGORY | XL_VALUE | XL_SERIES_AXIS) {
                    return Err(OmError::invalid_argument(
                        "Chart.HasAxis supports category, value, and series axes",
                    ));
                }
                let axis_group = args
                    .get(1)
                    .map(|value| coerce_u32_arg(value, "Chart.HasAxis axis group"))
                    .transpose()?
                    .unwrap_or(XL_PRIMARY);
                let axis_group =
                    chart_axis_group_from_excel_value(axis_group, "Chart.HasAxis axis group")?;
                Ok(OmValue::Bool(
                    self.chart_axis_index_for_type(workbook, chart_id, axis_type, axis_group)?
                        .is_some(),
                ))
            }
            "Legend" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.Legend does not accept arguments",
                    ));
                }
                if !self
                    .chart_model(workbook, chart_id)?
                    .legend
                    .as_ref()
                    .is_some_and(|legend| legend.visible)
                {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart legend not found",
                    ));
                }
                Ok(OmValue::Object(
                    self.register_legend_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        chart_object_parent,
                    ),
                ))
            }
            "ChartGroups" => {
                let handle = self.register_chart_groups_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    None,
                    chart_object_parent,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            member if is_chart_group_shortcut_member(member) => self
                .dispatch_get_chart_group_shortcut(
                    workbook,
                    chart_id,
                    chart_object_parent,
                    member,
                    args,
                ),
            "Axes" => {
                let handle = self.register_axes_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    chart_object_parent,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "SeriesCollection" => {
                let handle = self
                    .register_series_collection_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        chart_object_parent,
                    );
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "FullSeriesCollection" => {
                let handle = self
                    .register_full_series_collection_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        chart_object_parent,
                    );
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "ChartObjects" => {
                let Some(sheet_id) = self.chart_sheet_id_for_chart(workbook, chart_id)? else {
                    return Err(OmError::unsupported(
                        "Chart.ChartObjects is only available for chart sheets",
                    ));
                };
                let handle = self.register_chart_objects_handle_with_parent(
                    workbook,
                    sheet_id,
                    ChartObjectsParent::Chart(chart_id),
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.Application does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Chart.Parent does not accept arguments",
                    ));
                }
                if let Some((host_sheet_id, chart_object_id, _)) =
                    self.embedded_chart_object_for_chart(workbook, chart_id)?
                {
                    let parent =
                        chart_object_parent.unwrap_or(ChartObjectsParent::Worksheet(host_sheet_id));
                    return Ok(OmValue::Object(
                        self.register_chart_object_handle_with_parent_origin(
                            workbook,
                            chart_object_id,
                            Some(parent),
                        ),
                    ));
                }
                Ok(OmValue::Object(workbook.0))
            }
            _ => Err(OmError::unsupported(format!(
                "Chart.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_chart_child_container(
        &mut self,
        surface: &str,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "{surface}.{member} does not accept arguments"
            )));
        }
        self.chart_model(workbook, chart_id)?;

        match member {
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: match surface {
                        "ChartArea" => ChartFormatParent::ChartArea {
                            chart_id,
                            chart_object_parent,
                        },
                        "PlotArea" => ChartFormatParent::PlotArea {
                            chart_id,
                            chart_object_parent,
                        },
                        _ => unreachable!("chart child surface was provided by runtime dispatch"),
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: match surface {
                        "ChartArea" => BorderParent::ChartArea {
                            chart_id,
                            chart_object_parent,
                        },
                        "PlotArea" => BorderParent::PlotArea {
                            chart_id,
                            chart_object_parent,
                        },
                        _ => unreachable!("chart child surface was provided by runtime dispatch"),
                    },
                },
            ))),
            "Name" => Ok(OmValue::Text(match surface {
                "ChartArea" => "Chart Area".to_string(),
                "PlotArea" => "Plot Area".to_string(),
                _ => unreachable!("chart child surface was provided by runtime dispatch"),
            })),
            "Left" | "Top" | "Width" | "Height" if surface == "ChartArea" => {
                let state = &self.runtime_workbook(workbook)?.loaded.state;
                let chart_object = state
                    .drawings
                    .values()
                    .flat_map(|drawing| drawing.objects.iter())
                    .find_map(|object| match object {
                        DrawingObjectModel::ChartFrame(chart_object)
                            if chart_object.chart_id == chart_id =>
                        {
                            Some(chart_object)
                        }
                        DrawingObjectModel::UnsupportedRaw { .. } => None,
                        _ => None,
                    })
                    .ok_or_else(|| {
                        OmError::unsupported(
                            "ChartArea geometry is only available for embedded charts",
                        )
                    })?;
                Ok(OmValue::Number(Self::chart_object_geometry_value(
                    chart_object,
                    member,
                )?))
            }
            "Left" | "Top" | "Width" | "Height" | "InsideLeft" | "InsideTop" | "InsideWidth"
            | "InsideHeight"
                if surface == "PlotArea" =>
            {
                let Some(layout) = self.chart_model(workbook, chart_id)?.plot_area_layout else {
                    return Ok(OmValue::Number(0.0));
                };
                let state = &self.runtime_workbook(workbook)?.loaded.state;
                let Some(chart_object) = state
                    .drawings
                    .values()
                    .flat_map(|drawing| drawing.objects.iter())
                    .find_map(|object| match object {
                        DrawingObjectModel::ChartFrame(chart_object)
                            if chart_object.chart_id == chart_id =>
                        {
                            Some(chart_object)
                        }
                        DrawingObjectModel::UnsupportedRaw { .. } => None,
                        _ => None,
                    })
                else {
                    return Ok(OmValue::Number(0.0));
                };
                let chart_width = Self::chart_object_geometry_value(chart_object, "Width")?;
                let chart_height = Self::chart_object_geometry_value(chart_object, "Height")?;
                let left_fraction = layout.x.unwrap_or(0.0);
                let top_fraction = layout.y.unwrap_or(0.0);
                let value = match member {
                    "Left" | "InsideLeft" => left_fraction * chart_width,
                    "Top" | "InsideTop" => top_fraction * chart_height,
                    "Width" | "InsideWidth" => layout
                        .width
                        .map(|width| {
                            if layout.width_mode == ChartLayoutMode::Edge {
                                (width - left_fraction).max(0.0) * chart_width
                            } else {
                                width * chart_width
                            }
                        })
                        .unwrap_or(0.0),
                    "Height" | "InsideHeight" => layout
                        .height
                        .map(|height| {
                            if layout.height_mode == ChartLayoutMode::Edge {
                                (height - top_fraction).max(0.0) * chart_height
                            } else {
                                height * chart_height
                            }
                        })
                        .unwrap_or(0.0),
                    _ => unreachable!("PlotArea geometry member was matched"),
                };
                Ok(OmValue::Number(value))
            }
            "RoundedCorners" if surface == "ChartArea" => Ok(OmValue::Bool(
                self.chart_model(workbook, chart_id)?
                    .rounded_corners
                    .unwrap_or(false),
            )),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_chart_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "{surface}.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn chart_format_parent_object_kind(
        &self,
        workbook: WorkbookHandle,
        parent: ChartFormatParent,
    ) -> OmResult<RuntimeObjectKind> {
        match parent {
            ChartFormatParent::ChartArea {
                chart_id,
                chart_object_parent,
            } => {
                self.chart_model(workbook, chart_id)?;
                Ok(RuntimeObjectKind::ChartArea {
                    workbook,
                    chart_id,
                    chart_object_parent,
                })
            }
            ChartFormatParent::PlotArea {
                chart_id,
                chart_object_parent,
            } => {
                self.chart_model(workbook, chart_id)?;
                Ok(RuntimeObjectKind::PlotArea {
                    workbook,
                    chart_id,
                    chart_object_parent,
                })
            }
            ChartFormatParent::ChartTitle {
                chart_id,
                chart_object_parent,
            } => {
                if self.chart_model(workbook, chart_id)?.title.is_none() {
                    return Err(OmError::new(OmErrorCode::NotFound, "chart title not found"));
                }
                Ok(RuntimeObjectKind::ChartTitle {
                    workbook,
                    chart_id,
                    chart_object_parent,
                })
            }
            ChartFormatParent::Legend {
                chart_id,
                chart_object_parent,
            } => {
                if !self
                    .chart_model(workbook, chart_id)?
                    .legend
                    .as_ref()
                    .is_some_and(|legend| legend.visible)
                {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart legend not found",
                    ));
                }
                Ok(RuntimeObjectKind::Legend {
                    workbook,
                    chart_id,
                    chart_object_parent,
                })
            }
            ChartFormatParent::LegendEntry {
                chart_id,
                entry_index,
                chart_object_parent,
            } => {
                self.validate_legend_entry_index(workbook, chart_id, entry_index)?;
                Ok(RuntimeObjectKind::LegendEntry {
                    workbook,
                    chart_id,
                    entry_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::LegendKey {
                chart_id,
                entry_index,
                chart_object_parent,
            } => {
                self.validate_legend_entry_index(workbook, chart_id, entry_index)?;
                Ok(RuntimeObjectKind::LegendKey {
                    workbook,
                    chart_id,
                    entry_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::DataTable {
                chart_id,
                chart_object_parent,
            } => {
                if self.chart_model(workbook, chart_id)?.data_table.is_none() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart data table not found",
                    ));
                }
                Ok(RuntimeObjectKind::DataTable {
                    workbook,
                    chart_id,
                    chart_object_parent,
                })
            }
            ChartFormatParent::Axis {
                chart_id,
                axis_index,
                chart_object_parent,
            } => {
                self.axis_model(workbook, chart_id, axis_index)?;
                Ok(RuntimeObjectKind::Axis {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::AxisTitle {
                chart_id,
                axis_index,
                chart_object_parent,
            } => {
                if self
                    .axis_model(workbook, chart_id, axis_index)?
                    .title
                    .is_none()
                {
                    return Err(OmError::new(OmErrorCode::NotFound, "axis title not found"));
                }
                Ok(RuntimeObjectKind::AxisTitle {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::DisplayUnitLabel {
                chart_id,
                axis_index,
                chart_object_parent,
            } => {
                let axis = self.axis_model(workbook, chart_id, axis_index)?;
                if axis.kind != ChartAxisKind::Value {
                    return Err(OmError::unsupported(
                        "DisplayUnitLabel applies only to value axes",
                    ));
                }
                if axis.display_unit.is_none() || axis.has_display_unit_label != Some(true) {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "display unit label not found",
                    ));
                }
                Ok(RuntimeObjectKind::DisplayUnitLabel {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::TickLabels {
                chart_id,
                axis_index,
                chart_object_parent,
            } => {
                self.axis_model(workbook, chart_id, axis_index)?;
                Ok(RuntimeObjectKind::TickLabels {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::Gridlines {
                chart_id,
                axis_index,
                major,
                chart_object_parent,
            } => {
                let axis = self.axis_model(workbook, chart_id, axis_index)?;
                let has_gridlines = if major {
                    axis.has_major_gridlines
                } else {
                    axis.has_minor_gridlines
                };
                if has_gridlines != Some(true) {
                    return Err(OmError::new(OmErrorCode::NotFound, "gridlines not found"));
                }
                Ok(RuntimeObjectKind::Gridlines {
                    workbook,
                    chart_id,
                    axis_index,
                    major,
                    chart_object_parent,
                })
            }
            ChartFormatParent::Series {
                chart_id,
                series_index,
                chart_object_parent,
            } => {
                self.series_model(workbook, chart_id, series_index)?;
                Ok(RuntimeObjectKind::Series {
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::DataLabels {
                chart_id,
                series_index,
                chart_object_parent,
            } => {
                self.series_model(workbook, chart_id, series_index)?;
                Ok(RuntimeObjectKind::DataLabels {
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::LeaderLines {
                chart_id,
                series_index,
                chart_object_parent,
            } => {
                if !self.leader_lines_visible(workbook, chart_id, series_index)? {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "leader lines not found",
                    ));
                }
                Ok(RuntimeObjectKind::LeaderLines {
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::DataLabel {
                chart_id,
                series_index,
                point_index,
                chart_object_parent,
            } => {
                self.validate_data_label_index(workbook, chart_id, series_index, point_index)?;
                Ok(RuntimeObjectKind::DataLabel {
                    workbook,
                    chart_id,
                    series_index,
                    point_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::Point {
                chart_id,
                series_index,
                point_index,
                chart_object_parent,
            } => {
                self.validate_point_index(workbook, chart_id, series_index, point_index)?;
                Ok(RuntimeObjectKind::Point {
                    workbook,
                    chart_id,
                    series_index,
                    point_index,
                    chart_object_parent,
                })
            }
            ChartFormatParent::ChartGroupLines {
                chart_id,
                group_index,
                kind,
                chart_object_parent,
            } => {
                self.chart_group_model(workbook, chart_id, group_index)?;
                Ok(RuntimeObjectKind::ChartGroupLines {
                    workbook,
                    chart_id,
                    group_index,
                    kind,
                    chart_object_parent,
                })
            }
        }
    }

    pub(crate) fn border_parent_object_kind(
        &self,
        workbook: WorkbookHandle,
        parent: BorderParent,
    ) -> OmResult<RuntimeObjectKind> {
        match parent {
            BorderParent::ChartArea {
                chart_id,
                chart_object_parent,
            } => {
                self.chart_model(workbook, chart_id)?;
                Ok(RuntimeObjectKind::ChartArea {
                    workbook,
                    chart_id,
                    chart_object_parent,
                })
            }
            BorderParent::PlotArea {
                chart_id,
                chart_object_parent,
            } => {
                self.chart_model(workbook, chart_id)?;
                Ok(RuntimeObjectKind::PlotArea {
                    workbook,
                    chart_id,
                    chart_object_parent,
                })
            }
            BorderParent::ChartTitle {
                chart_id,
                chart_object_parent,
            } => {
                if self.chart_model(workbook, chart_id)?.title.is_none() {
                    return Err(OmError::new(OmErrorCode::NotFound, "chart title not found"));
                }
                Ok(RuntimeObjectKind::ChartTitle {
                    workbook,
                    chart_id,
                    chart_object_parent,
                })
            }
            BorderParent::Legend {
                chart_id,
                chart_object_parent,
            } => {
                if !self
                    .chart_model(workbook, chart_id)?
                    .legend
                    .as_ref()
                    .is_some_and(|legend| legend.visible)
                {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart legend not found",
                    ));
                }
                Ok(RuntimeObjectKind::Legend {
                    workbook,
                    chart_id,
                    chart_object_parent,
                })
            }
            BorderParent::LegendKey {
                chart_id,
                entry_index,
                chart_object_parent,
            } => {
                self.validate_legend_entry_index(workbook, chart_id, entry_index)?;
                Ok(RuntimeObjectKind::LegendKey {
                    workbook,
                    chart_id,
                    entry_index,
                    chart_object_parent,
                })
            }
            BorderParent::Axis {
                chart_id,
                axis_index,
                chart_object_parent,
            } => {
                self.axis_model(workbook, chart_id, axis_index)?;
                Ok(RuntimeObjectKind::Axis {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                })
            }
            BorderParent::AxisTitle {
                chart_id,
                axis_index,
                chart_object_parent,
            } => {
                if self
                    .axis_model(workbook, chart_id, axis_index)?
                    .title
                    .is_none()
                {
                    return Err(OmError::new(OmErrorCode::NotFound, "axis title not found"));
                }
                Ok(RuntimeObjectKind::AxisTitle {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                })
            }
            BorderParent::DisplayUnitLabel {
                chart_id,
                axis_index,
                chart_object_parent,
            } => {
                let axis = self.axis_model(workbook, chart_id, axis_index)?;
                if axis.kind != ChartAxisKind::Value {
                    return Err(OmError::unsupported(
                        "DisplayUnitLabel applies only to value axes",
                    ));
                }
                if axis.display_unit.is_none() || axis.has_display_unit_label != Some(true) {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "display unit label not found",
                    ));
                }
                Ok(RuntimeObjectKind::DisplayUnitLabel {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                })
            }
            BorderParent::DataTable {
                chart_id,
                chart_object_parent,
            } => {
                if self.chart_model(workbook, chart_id)?.data_table.is_none() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart data table not found",
                    ));
                }
                Ok(RuntimeObjectKind::DataTable {
                    workbook,
                    chart_id,
                    chart_object_parent,
                })
            }
            BorderParent::Gridlines {
                chart_id,
                axis_index,
                major,
                chart_object_parent,
            } => {
                let axis = self.axis_model(workbook, chart_id, axis_index)?;
                let has_gridlines = if major {
                    axis.has_major_gridlines
                } else {
                    axis.has_minor_gridlines
                };
                if has_gridlines != Some(true) {
                    return Err(OmError::new(OmErrorCode::NotFound, "gridlines not found"));
                }
                Ok(RuntimeObjectKind::Gridlines {
                    workbook,
                    chart_id,
                    axis_index,
                    major,
                    chart_object_parent,
                })
            }
            BorderParent::DataLabels {
                chart_id,
                series_index,
                chart_object_parent,
            } => {
                self.series_model(workbook, chart_id, series_index)?;
                Ok(RuntimeObjectKind::DataLabels {
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                })
            }
            BorderParent::LeaderLines {
                chart_id,
                series_index,
                chart_object_parent,
            } => {
                if !self.leader_lines_visible(workbook, chart_id, series_index)? {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "leader lines not found",
                    ));
                }
                Ok(RuntimeObjectKind::LeaderLines {
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                })
            }
            BorderParent::DataLabel {
                chart_id,
                series_index,
                point_index,
                chart_object_parent,
            } => {
                self.validate_data_label_index(workbook, chart_id, series_index, point_index)?;
                Ok(RuntimeObjectKind::DataLabel {
                    workbook,
                    chart_id,
                    series_index,
                    point_index,
                    chart_object_parent,
                })
            }
            BorderParent::Point {
                chart_id,
                series_index,
                point_index,
                chart_object_parent,
            } => {
                self.validate_point_index(workbook, chart_id, series_index, point_index)?;
                Ok(RuntimeObjectKind::Point {
                    workbook,
                    chart_id,
                    series_index,
                    point_index,
                    chart_object_parent,
                })
            }
            BorderParent::ChartGroupLines {
                chart_id,
                group_index,
                kind,
                chart_object_parent,
            } => {
                let chart = self.chart_group_model(workbook, chart_id, group_index)?;
                let has_lines = match kind {
                    ChartGroupLineKind::SeriesLines => chart.has_series_lines,
                    ChartGroupLineKind::DropLines => chart.has_drop_lines,
                    ChartGroupLineKind::HiLoLines => chart.has_hi_lo_lines,
                    ChartGroupLineKind::UpBars | ChartGroupLineKind::DownBars => {
                        chart.has_up_down_bars
                    }
                };
                if has_lines != Some(true) {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        format!("{} not found", kind.display_name()),
                    ));
                }
                Ok(RuntimeObjectKind::ChartGroupLines {
                    workbook,
                    chart_id,
                    group_index,
                    kind,
                    chart_object_parent,
                })
            }
        }
    }

    pub(crate) fn dispatch_get_legend(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if !args.is_empty() && member != "LegendEntries" {
            return Err(OmError::invalid_argument(format!(
                "Legend.{member} does not accept arguments"
            )));
        }
        let legend = self
            .chart_model(workbook, chart_id)?
            .legend
            .as_ref()
            .filter(|legend| legend.visible)
            .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart legend not found"))?;

        match member {
            "Name" => Ok(OmValue::Text("Legend".to_string())),
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::Legend {
                        chart_id,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::Legend {
                        chart_id,
                        chart_object_parent,
                    },
                },
            ))),
            "IncludeInLayout" => Ok(OmValue::Bool(legend.include_in_layout.unwrap_or(true))),
            "Left" | "Top" | "Width" | "Height" => Ok(OmValue::Number(0.0)),
            "Position" => {
                let position = legend.position.ok_or_else(|| {
                    OmError::unsupported("Legend.Position is unavailable for unknown position")
                })?;
                Ok(OmValue::Number(f64::from(match position {
                    ChartLegendPosition::Bottom => XL_LEGEND_POSITION_BOTTOM,
                    ChartLegendPosition::Corner => XL_LEGEND_POSITION_CORNER,
                    ChartLegendPosition::Custom => XL_LEGEND_POSITION_CUSTOM,
                    ChartLegendPosition::Left => XL_LEGEND_POSITION_LEFT,
                    ChartLegendPosition::Right => XL_LEGEND_POSITION_RIGHT,
                    ChartLegendPosition::Top => XL_LEGEND_POSITION_TOP,
                })))
            }
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_chart_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    chart_object_parent,
                ),
            )),
            "LegendEntries" => {
                let handle = self.register_legend_entries_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    chart_object_parent,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            _ => Err(OmError::unsupported(format!(
                "Legend.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_legend_entries(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("LegendEntries", member) {
            self.focus_member_supported("LegendEntries", member, false)?;
        }

        if !args.is_empty() && member != "Item" {
            return Err(OmError::invalid_argument(format!(
                "LegendEntries.{member} does not accept arguments"
            )));
        }

        match member {
            "Item" => self.dispatch_invoke_legend_entries(
                workbook,
                chart_id,
                chart_object_parent,
                member,
                args,
            ),
            "Count" => {
                let chart = self.chart_model(workbook, chart_id)?;
                if !chart.legend.as_ref().is_some_and(|legend| legend.visible) {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "chart legend not found",
                    ));
                }
                Ok(OmValue::Number(chart.series.len() as f64))
            }
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_legend_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "LegendEntries.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_legend_entries(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("LegendEntries", member) {
            self.focus_member_supported("LegendEntries", member, false)?;
        }

        match member {
            "Item" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "LegendEntries.Item expects a single 1-based index",
                    ));
                };
                let index = coerce_u32_arg(index, "LegendEntries.Item index")? as usize;
                if index == 0 {
                    return Err(OmError::invalid_argument(
                        "LegendEntries.Item index is out of bounds",
                    ));
                }
                self.validate_legend_entry_index(workbook, chart_id, index - 1)?;
                Ok(OmValue::Object(
                    self.register_legend_entry_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        index - 1,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "LegendEntries.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_get_legend_entry(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        entry_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("LegendEntry", member) {
            self.focus_member_supported("LegendEntry", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "LegendEntry.{member} does not accept arguments"
            )));
        }
        self.validate_legend_entry_index(workbook, chart_id, entry_index)?;

        match member {
            "Index" => Ok(OmValue::Number((entry_index + 1) as f64)),
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::LegendEntry {
                        chart_id,
                        entry_index,
                        chart_object_parent,
                    },
                },
            ))),
            "LegendKey" => Ok(OmValue::Object(
                self.register_legend_key_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    entry_index,
                    chart_object_parent,
                ),
            )),
            "Left" | "Top" | "Width" | "Height" => Ok(OmValue::Number(0.0)),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_legend_entries_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "LegendEntry.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_legend_key(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        entry_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("LegendKey", member) {
            self.focus_member_supported("LegendKey", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "LegendKey.{member} does not accept arguments"
            )));
        }
        self.validate_legend_entry_index(workbook, chart_id, entry_index)?;

        match member {
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::LegendKey {
                        chart_id,
                        entry_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::LegendKey {
                        chart_id,
                        entry_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Left" | "Top" | "Width" | "Height" => Ok(OmValue::Number(0.0)),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_legend_entry_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    entry_index,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "LegendKey.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_data_table(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "DataTable.{member} does not accept arguments"
            )));
        }
        let data_table = self
            .chart_model(workbook, chart_id)?
            .data_table
            .as_ref()
            .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart data table not found"))?;
        let has_border_horizontal = data_table.has_border_horizontal;
        let has_border_vertical = data_table.has_border_vertical;
        let has_border_outline = data_table.has_border_outline;
        let show_legend_key = data_table.show_legend_key;

        match member {
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::DataTable {
                        chart_id,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::DataTable {
                        chart_id,
                        chart_object_parent,
                    },
                },
            ))),
            "HasBorderHorizontal" => Ok(OmValue::Bool(has_border_horizontal.unwrap_or(true))),
            "HasBorderVertical" => Ok(OmValue::Bool(has_border_vertical.unwrap_or(true))),
            "HasBorderOutline" => Ok(OmValue::Bool(has_border_outline.unwrap_or(true))),
            "ShowLegendKey" => Ok(OmValue::Bool(show_legend_key.unwrap_or(false))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_chart_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "DataTable.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_chart_group_shortcut(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        let shortcut = ChartGroupShortcutKind::from_member(member).ok_or_else(|| {
            OmError::invalid_argument(format!("{member} is not a chart-group collection"))
        })?;
        let chart = self.chart_model(workbook, chart_id)?;
        if chart_group_indices(chart, Some(shortcut)).is_empty() {
            return Err(OmError::new(
                OmErrorCode::NotFound,
                format!("{member} is not available for this chart type"),
            ));
        }

        let handle = self.register_chart_groups_handle_with_chart_object_parent_origin(
            workbook,
            chart_id,
            Some(shortcut),
            chart_object_parent,
        );
        if args.is_empty() {
            Ok(OmValue::Object(handle))
        } else {
            self.dispatch_invoke(handle, "Item", args)
        }
    }

    pub(crate) fn dispatch_get_chart_groups(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        shortcut: Option<ChartGroupShortcutKind>,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("ChartGroups", member) {
            self.focus_member_supported("ChartGroups", member, false)?;
        }

        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartGroups.Count does not accept arguments",
                    ));
                }
                let chart = self.chart_model(workbook, chart_id)?;
                Ok(OmValue::Number(
                    chart_group_indices(chart, shortcut).len() as f64
                ))
            }
            "Item" => self.dispatch_invoke_chart_groups(
                workbook,
                chart_id,
                shortcut,
                chart_object_parent,
                member,
                args,
            ),
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartGroups.Creator does not accept arguments",
                    ));
                }
                self.chart_model(workbook, chart_id)?;
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartGroups.Application does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "ChartGroups.Parent does not accept arguments",
                    ));
                }
                self.chart_model(workbook, chart_id)?;
                Ok(OmValue::Object(
                    self.register_chart_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "ChartGroups.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_chart_groups(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        shortcut: Option<ChartGroupShortcutKind>,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("ChartGroups", member) {
            self.focus_member_supported("ChartGroups", member, false)?;
        }

        match member {
            "Item" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "ChartGroups.Item expects a single 1-based index",
                    ));
                };
                let index = coerce_u32_arg(index, "ChartGroups.Item index")? as usize;
                if index == 0 {
                    return Err(OmError::invalid_argument(
                        "ChartGroups.Item index is out of bounds",
                    ));
                }
                let chart = self.chart_model(workbook, chart_id)?;
                let group_indices = chart_group_indices(chart, shortcut);
                if index > group_indices.len() {
                    return Err(OmError::invalid_argument(
                        "ChartGroups.Item index is out of bounds",
                    ));
                }
                let group_index = group_indices[index - 1];
                Ok(OmValue::Object(
                    self.register_chart_group_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        group_index,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "ChartGroups.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_get_chart_group(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        group_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("ChartGroup", member) {
            self.focus_member_supported("ChartGroup", member, false)?;
        }

        match member {
            "SeriesCollection" => {
                let (axis_group, group_index_filter) = {
                    let chart = self.chart_group_model(workbook, chart_id, group_index)?;
                    (
                        chart_group_axis_group(chart, group_index)?,
                        chart_group_overlay_is_stable(chart).then_some(group_index),
                    )
                };
                let handle = self.register_object(RuntimeObjectKind::SeriesCollection {
                    workbook,
                    chart_id,
                    axis_group_filter: Some(axis_group),
                    group_index_filter,
                    full: false,
                    chart_object_parent,
                });
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "CategoryCollection" | "FullCategoryCollection" => {
                self.chart_group_model(workbook, chart_id, group_index)?;
                let full = member == "FullCategoryCollection";
                let handle = self
                    .register_category_collection_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        group_index,
                        full,
                        chart_object_parent,
                    );
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            _ => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "ChartGroup.{member} does not accept arguments"
                    )));
                }
                let chart = self.chart_group_model(workbook, chart_id, group_index)?;
                let group = chart_group_overlay_is_stable(chart)
                    .then(|| chart.groups.get(group_index))
                    .flatten();
                match member {
                    "ChartType" => {
                        let group_chart_type = chart_group_chart_type(chart, group_index)?;
                        Ok(OmValue::Number(f64::from(chart_type_to_excel_value(
                            &group_chart_type,
                        )?)))
                    }
                    "Index" => Ok(OmValue::Number((group_index + 1) as f64)),
                    "AxisGroup" => Ok(OmValue::Number(f64::from(
                        match chart_group_axis_group(chart, group_index)? {
                            ChartAxisGroup::Primary => XL_PRIMARY,
                            ChartAxisGroup::Secondary => XL_SECONDARY,
                        },
                    ))),
                    "SeriesLines" | "DropLines" | "HiLoLines" | "UpBars" | "DownBars" => {
                        let kind = ChartGroupLineKind::from_chart_group_member(member)
                            .expect("matched ChartGroup line object member");
                        let group_chart_type = chart_group_chart_type(chart, group_index)?;
                        let restrict_to_primitive_group =
                            chart_type_is_volume_stock(&chart.chart_type);
                        let has_lines = match kind {
                            ChartGroupLineKind::SeriesLines => group
                                .and_then(|group| group.has_series_lines)
                                .or(chart.has_series_lines),
                            ChartGroupLineKind::DropLines
                                if !restrict_to_primitive_group
                                    || chart_type_supports_high_low_lines(&group_chart_type) =>
                            {
                                group
                                    .and_then(|group| group.has_drop_lines)
                                    .or(chart.has_drop_lines)
                            }
                            ChartGroupLineKind::HiLoLines
                                if !restrict_to_primitive_group
                                    || chart_type_supports_high_low_lines(&group_chart_type) =>
                            {
                                group
                                    .and_then(|group| group.has_hi_lo_lines)
                                    .or(chart.has_hi_lo_lines)
                            }
                            ChartGroupLineKind::UpBars | ChartGroupLineKind::DownBars => {
                                if !restrict_to_primitive_group
                                    || chart_type_supports_up_down_bars(&group_chart_type)
                                {
                                    group
                                        .and_then(|group| group.has_up_down_bars)
                                        .or(chart.has_up_down_bars)
                                } else {
                                    Some(false)
                                }
                            }
                            _ => Some(false),
                        };
                        if has_lines != Some(true) {
                            return Err(OmError::new(
                                OmErrorCode::NotFound,
                                format!("{} not found", kind.display_name()),
                            ));
                        }
                        Ok(OmValue::Object(self.register_object(
                            RuntimeObjectKind::ChartGroupLines {
                                workbook,
                                chart_id,
                                group_index,
                                kind,
                                chart_object_parent,
                            },
                        )))
                    }
                    "RadarAxisLabels" => {
                        let axis_index =
                            self.chart_group_radar_axis_index(workbook, chart_id, group_index)?;
                        Ok(OmValue::Object(self.register_object(
                            RuntimeObjectKind::TickLabels {
                                workbook,
                                chart_id,
                                axis_index,
                                chart_object_parent,
                            },
                        )))
                    }
                    "HasRadarAxisLabels" => {
                        if !chart_type_supports_radar_axis_labels(&chart.chart_type) {
                            return Err(OmError::unsupported(
                                "ChartGroup.HasRadarAxisLabels is only supported for radar chart groups",
                            ));
                        }
                        let position = chart
                            .axes
                            .iter()
                            .find(|axis| {
                                matches!(axis.kind, ChartAxisKind::Category | ChartAxisKind::Date)
                            })
                            .and_then(|axis| axis.tick_label_position)
                            .unwrap_or(ChartTickLabelPosition::NextToAxis);
                        Ok(OmValue::Bool(position != ChartTickLabelPosition::None))
                    }
                    "VaryByCategories" => Ok(OmValue::Bool(
                        group
                            .and_then(|group| group.vary_by_categories)
                            .or(chart.vary_by_categories)
                            .unwrap_or(false),
                    )),
                    "GapWidth" => Ok(OmValue::Number(f64::from(
                        group
                            .and_then(|group| group.gap_width)
                            .or(chart.gap_width)
                            .unwrap_or(150),
                    ))),
                    "Overlap" => Ok(OmValue::Number(f64::from(
                        group
                            .and_then(|group| group.overlap)
                            .or(chart.overlap)
                            .unwrap_or(0),
                    ))),
                    "HasSeriesLines" => Ok(OmValue::Bool(
                        group
                            .and_then(|group| group.has_series_lines)
                            .or(chart.has_series_lines)
                            .unwrap_or(false),
                    )),
                    "HasDropLines" => Ok(OmValue::Bool(
                        group
                            .and_then(|group| group.has_drop_lines)
                            .or(chart.has_drop_lines)
                            .unwrap_or(false),
                    )),
                    "HasHiLoLines" => Ok(OmValue::Bool(
                        (!chart_type_is_volume_stock(&chart.chart_type)
                            || chart_type_supports_high_low_lines(&chart_group_chart_type(
                                chart,
                                group_index,
                            )?))
                            && group
                                .and_then(|group| group.has_hi_lo_lines)
                                .or(chart.has_hi_lo_lines)
                                .unwrap_or(false),
                    )),
                    "HasUpDownBars" => Ok(OmValue::Bool(
                        (!chart_type_is_volume_stock(&chart.chart_type)
                            || chart_type_supports_up_down_bars(&chart_group_chart_type(
                                chart,
                                group_index,
                            )?))
                            && group
                                .and_then(|group| group.has_up_down_bars)
                                .or(chart.has_up_down_bars)
                                .unwrap_or(false),
                    )),
                    "FirstSliceAngle" => Ok(OmValue::Number(f64::from(
                        group
                            .and_then(|group| group.first_slice_angle)
                            .or(chart.first_slice_angle)
                            .unwrap_or(0),
                    ))),
                    "Explosion" => Ok(OmValue::Number(f64::from(
                        group
                            .and_then(|group| group.explosion)
                            .or(chart.explosion)
                            .unwrap_or(0),
                    ))),
                    "BubbleScale" => Ok(OmValue::Number(f64::from(
                        group
                            .and_then(|group| group.bubble_scale)
                            .or(chart.bubble_scale)
                            .unwrap_or(100),
                    ))),
                    "ShowNegativeBubbles" => Ok(OmValue::Bool(
                        group
                            .and_then(|group| group.show_negative_bubbles)
                            .or(chart.show_negative_bubbles)
                            .unwrap_or(false),
                    )),
                    "Has3DShading" => Ok(OmValue::Bool(
                        group
                            .and_then(|group| group.has_3d_shading)
                            .or(chart.has_3d_shading)
                            .unwrap_or(false),
                    )),
                    "DoughnutHoleSize" => Ok(OmValue::Number(f64::from(
                        group
                            .and_then(|group| group.doughnut_hole_size)
                            .or(chart.doughnut_hole_size)
                            .unwrap_or(75),
                    ))),
                    "SecondPlotSize" => Ok(OmValue::Number(f64::from(
                        group
                            .and_then(|group| group.second_plot_size)
                            .or(chart.second_plot_size)
                            .unwrap_or(75),
                    ))),
                    "SizeRepresents" => Ok(OmValue::Number(f64::from(
                        match group
                            .and_then(|group| group.size_represents)
                            .or(chart.size_represents)
                            .unwrap_or(ChartSizeRepresents::Area)
                        {
                            ChartSizeRepresents::Area => XL_SIZE_IS_AREA,
                            ChartSizeRepresents::Width => XL_SIZE_IS_WIDTH,
                        },
                    ))),
                    "SplitType" => Ok(OmValue::Number(f64::from(
                        match group
                            .and_then(|group| group.split_type)
                            .or(chart.split_type)
                            .unwrap_or(ChartSplitType::Position)
                        {
                            ChartSplitType::Custom => XL_SPLIT_BY_CUSTOM_SPLIT,
                            ChartSplitType::PercentValue => XL_SPLIT_BY_PERCENT_VALUE,
                            ChartSplitType::Position => XL_SPLIT_BY_POSITION,
                            ChartSplitType::Value => XL_SPLIT_BY_VALUE,
                        },
                    ))),
                    "SplitValue" => Ok(OmValue::Number(
                        group
                            .and_then(|group| group.split_value)
                            .or(chart.split_value)
                            .unwrap_or(0.0),
                    )),
                    "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
                    "Application" => Ok(OmValue::Object(self.root_application())),
                    "Parent" => Ok(OmValue::Object(
                        self.register_chart_handle_with_chart_object_parent_origin(
                            workbook,
                            chart_id,
                            chart_object_parent,
                        ),
                    )),
                    _ => Err(OmError::unsupported(format!(
                        "ChartGroup.{member} is not implemented"
                    ))),
                }
            }
        }
    }

    pub(crate) fn dispatch_invoke_chart_group(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        group_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("ChartGroup", member) {
            self.focus_member_supported("ChartGroup", member, false)?;
        }

        match member {
            "SeriesCollection" | "CategoryCollection" | "FullCategoryCollection" => self
                .dispatch_get_chart_group(
                    workbook,
                    chart_id,
                    group_index,
                    chart_object_parent,
                    member,
                    args,
                ),
            _ => Err(OmError::unsupported(format!(
                "ChartGroup.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_get_chart_group_lines(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        group_index: usize,
        kind: ChartGroupLineKind,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        let surface = kind.surface_name();
        if self.focus_member_declared(surface, member) {
            self.focus_member_supported(surface, member, false)?;
        }
        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "{surface}.{member} does not accept arguments"
            )));
        }
        self.chart_group_model(workbook, chart_id, group_index)?;
        match member {
            "Name" => Ok(OmValue::Text(kind.display_name().to_string())),
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::ChartGroupLines {
                        chart_id,
                        group_index,
                        kind,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::ChartGroupLines {
                        chart_id,
                        group_index,
                        kind,
                        chart_object_parent,
                    },
                },
            ))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_chart_group_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    group_index,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "{surface}.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_chart_group_lines(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        group_index: usize,
        kind: ChartGroupLineKind,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        let surface = kind.surface_name();
        if self.focus_member_declared(surface, member) {
            self.focus_member_supported(surface, member, false)?;
        }
        match member {
            "Copy" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "{surface}.Copy does not accept arguments"
                    )));
                }
                let chart = self.chart_model(workbook, chart_id)?;
                chart_group_axis_group(chart, group_index)?;
                let group = chart_group_overlay_is_stable(chart)
                    .then(|| chart.groups.get(group_index))
                    .flatten();
                let exists = match kind {
                    ChartGroupLineKind::SeriesLines => group
                        .and_then(|group| group.has_series_lines)
                        .or(chart.has_series_lines),
                    ChartGroupLineKind::DropLines => group
                        .and_then(|group| group.has_drop_lines)
                        .or(chart.has_drop_lines),
                    ChartGroupLineKind::HiLoLines => group
                        .and_then(|group| group.has_hi_lo_lines)
                        .or(chart.has_hi_lo_lines),
                    ChartGroupLineKind::UpBars | ChartGroupLineKind::DownBars => group
                        .and_then(|group| group.has_up_down_bars)
                        .or(chart.has_up_down_bars),
                };
                if exists != Some(true) {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        format!("{surface} not found"),
                    ));
                }
                self.set_headless_copy_mode();
                Ok(OmValue::Empty)
            }
            "Select" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "{surface}.Select does not accept arguments"
                    )));
                }
                self.chart_group_model(workbook, chart_id, group_index)?;
                let chart = self.register_chart_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    chart_object_parent,
                );
                self.dispatch_invoke(chart, "Select", &[])
            }
            "Delete" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "{surface}.Delete does not accept arguments"
                    )));
                }
                {
                    let runtime = self.runtime_workbook_mut(workbook)?;
                    if runtime.read_only {
                        return Err(OmError::new(
                            OmErrorCode::InvalidState,
                            "cannot modify a read-only workbook",
                        ));
                    }
                    let chart = runtime
                        .loaded
                        .state
                        .charts
                        .get_mut(&chart_id)
                        .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart not found"))?;
                    chart_group_axis_group(chart, group_index)?;
                    let current = chart.groups.get(group_index).and_then(|group| match kind {
                        ChartGroupLineKind::SeriesLines => group.has_series_lines,
                        ChartGroupLineKind::DropLines => group.has_drop_lines,
                        ChartGroupLineKind::HiLoLines => group.has_hi_lo_lines,
                        ChartGroupLineKind::UpBars | ChartGroupLineKind::DownBars => {
                            group.has_up_down_bars
                        }
                    });
                    let fallback = match kind {
                        ChartGroupLineKind::SeriesLines => chart.has_series_lines,
                        ChartGroupLineKind::DropLines => chart.has_drop_lines,
                        ChartGroupLineKind::HiLoLines => chart.has_hi_lo_lines,
                        ChartGroupLineKind::UpBars | ChartGroupLineKind::DownBars => {
                            chart.has_up_down_bars
                        }
                    };
                    if current.or(fallback) != Some(true) {
                        return Err(OmError::new(
                            OmErrorCode::NotFound,
                            format!("{surface} not found"),
                        ));
                    }
                    match kind {
                        ChartGroupLineKind::SeriesLines => chart.has_series_lines = Some(false),
                        ChartGroupLineKind::DropLines => chart.has_drop_lines = Some(false),
                        ChartGroupLineKind::HiLoLines => chart.has_hi_lo_lines = Some(false),
                        ChartGroupLineKind::UpBars | ChartGroupLineKind::DownBars => {
                            chart.has_up_down_bars = Some(false)
                        }
                    }
                    if let Some(group) = chart.groups.get_mut(group_index) {
                        match kind {
                            ChartGroupLineKind::SeriesLines => group.has_series_lines = Some(false),
                            ChartGroupLineKind::DropLines => group.has_drop_lines = Some(false),
                            ChartGroupLineKind::HiLoLines => group.has_hi_lo_lines = Some(false),
                            ChartGroupLineKind::UpBars | ChartGroupLineKind::DownBars => {
                                group.has_up_down_bars = Some(false)
                            }
                        }
                        group.dirty = true;
                    }
                    chart.content_dirty = true;
                    chart.dirty = true;
                    runtime.prompt_dirty = true;
                }
                self.stale_chart_group_line_handles_for_group(
                    workbook,
                    chart_id,
                    group_index,
                    kind,
                );
                self.find_state = None;
                self.cut_copy_mode = None;
                self.clipboard = None;
                Ok(OmValue::Empty)
            }
            "ClearFormats" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "{surface}.ClearFormats does not accept arguments"
                    )));
                }
                let runtime = self.runtime_workbook_mut(workbook)?;
                if runtime.read_only {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        "cannot modify a read-only workbook",
                    ));
                }
                let chart = runtime
                    .loaded
                    .state
                    .charts
                    .get_mut(&chart_id)
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart not found"))?;
                chart_group_axis_group(chart, group_index)?;
                let group = chart_group_overlay_is_stable(chart)
                    .then(|| chart.groups.get(group_index))
                    .flatten();
                let exists = match kind {
                    ChartGroupLineKind::SeriesLines => group
                        .and_then(|group| group.has_series_lines)
                        .or(chart.has_series_lines),
                    ChartGroupLineKind::DropLines => group
                        .and_then(|group| group.has_drop_lines)
                        .or(chart.has_drop_lines),
                    ChartGroupLineKind::HiLoLines => group
                        .and_then(|group| group.has_hi_lo_lines)
                        .or(chart.has_hi_lo_lines),
                    ChartGroupLineKind::UpBars | ChartGroupLineKind::DownBars => group
                        .and_then(|group| group.has_up_down_bars)
                        .or(chart.has_up_down_bars),
                };
                if exists != Some(true) {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        format!("{surface} not found"),
                    ));
                }
                chart.content_dirty = true;
                chart.dirty = true;
                runtime.prompt_dirty = true;
                self.find_state = None;
                self.cut_copy_mode = None;
                self.clipboard = None;
                Ok(OmValue::Empty)
            }
            _ => Err(OmError::unsupported(format!(
                "{surface}.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_get_category_collection(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        group_index: usize,
        full: bool,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("CategoryCollection", member) {
            self.focus_member_supported("CategoryCollection", member, false)?;
        }

        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "CategoryCollection.Count does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(self.chart_category_count_for_group(
                    workbook,
                    chart_id,
                    group_index,
                    full,
                )? as f64))
            }
            "Item" => self.dispatch_invoke_category_collection(
                workbook,
                chart_id,
                group_index,
                full,
                chart_object_parent,
                member,
                args,
            ),
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "CategoryCollection.Creator does not accept arguments",
                    ));
                }
                self.chart_group_model(workbook, chart_id, group_index)?;
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "CategoryCollection.Application does not accept arguments",
                    ));
                }
                self.chart_group_model(workbook, chart_id, group_index)?;
                Ok(OmValue::Object(self.root_application()))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "CategoryCollection.Parent does not accept arguments",
                    ));
                }
                self.chart_group_model(workbook, chart_id, group_index)?;
                Ok(OmValue::Object(
                    self.register_chart_group_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        group_index,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "CategoryCollection.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_category_collection(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        group_index: usize,
        full: bool,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("CategoryCollection", member) {
            self.focus_member_supported("CategoryCollection", member, false)?;
        }

        match member {
            "Item" => {
                let [selector] = args else {
                    return Err(OmError::invalid_argument(
                        "CategoryCollection.Item expects a single 1-based index or category name",
                    ));
                };
                let category_count =
                    self.chart_category_count_for_group(workbook, chart_id, group_index, full)?;
                let category_index = match selector {
                    OmValue::Number(index) => {
                        let index = coerce_positive_index(*index, "CategoryCollection.Item index")?
                            as usize;
                        if index > category_count {
                            return Err(OmError::invalid_argument(
                                "CategoryCollection.Item index is out of bounds",
                            ));
                        }
                        index - 1
                    }
                    OmValue::Text(name) => {
                        let mut matched_index = None;
                        for index in 0..category_count {
                            if self
                                .chart_category_name_for_index(
                                    workbook,
                                    chart_id,
                                    group_index,
                                    index,
                                    full,
                                )?
                                .eq_ignore_ascii_case(name)
                            {
                                matched_index = Some(index);
                                break;
                            }
                        }
                        matched_index.ok_or_else(|| {
                            OmError::new(
                                OmErrorCode::NotFound,
                                "CategoryCollection.Item category not found",
                            )
                        })?
                    }
                    _ => {
                        return Err(OmError::type_mismatch(
                            "CategoryCollection.Item expects a numeric index or category name",
                        ));
                    }
                };
                Ok(OmValue::Object(
                    self.register_chart_category_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        group_index,
                        category_index,
                        full,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "CategoryCollection.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_get_chart_category(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        group_index: usize,
        category_index: usize,
        full: bool,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("ChartCategory", member) {
            self.focus_member_supported("ChartCategory", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "ChartCategory.{member} does not accept arguments"
            )));
        }
        let category_count =
            self.chart_category_count_for_group(workbook, chart_id, group_index, full)?;
        if category_index >= category_count {
            return Err(OmError::new(
                OmErrorCode::NotFound,
                "chart category not found",
            ));
        }
        match member {
            "Name" => Ok(OmValue::Text(self.chart_category_name_for_index(
                workbook,
                chart_id,
                group_index,
                category_index,
                full,
            )?)),
            "Index" => Ok(OmValue::Number((category_index + 1) as f64)),
            "IsFiltered" => Ok(OmValue::Bool(self.chart_category_is_filtered_for_index(
                workbook,
                chart_id,
                group_index,
                category_index,
                full,
            )?)),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_category_collection_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    group_index,
                    full,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "ChartCategory.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_axes(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("Axes", member) {
            self.focus_member_supported("Axes", member, false)?;
        }

        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Axes.Count does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(
                    self.chart_model(workbook, chart_id)?
                        .axes
                        .iter()
                        .filter(|axis| axis.deleted != Some(true))
                        .count() as f64,
                ))
            }
            "Item" => {
                self.dispatch_invoke_axes(workbook, chart_id, chart_object_parent, member, args)
            }
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Axes.Creator does not accept arguments",
                    ));
                }
                self.chart_model(workbook, chart_id)?;
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Axes.Application does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Axes.Parent does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(
                    self.register_chart_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Axes.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_axes(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("Axes", member) {
            self.focus_member_supported("Axes", member, false)?;
        }

        match member {
            "Item" => {
                if args.is_empty() || args.len() > 2 {
                    return Err(OmError::invalid_argument(
                        "Axes.Item expects axis type and optional axis group",
                    ));
                }
                let axis_type = coerce_u32_arg(&args[0], "Axes.Item axis type")? as i32;
                if !matches!(axis_type, XL_CATEGORY | XL_VALUE | XL_SERIES_AXIS) {
                    return Err(OmError::invalid_argument(
                        "Axes.Item supports category, value, and series axes",
                    ));
                }
                let axis_group = args
                    .get(1)
                    .map(|value| coerce_u32_arg(value, "Axes.Item axis group"))
                    .transpose()?
                    .unwrap_or(XL_PRIMARY);
                let axis_group =
                    chart_axis_group_from_excel_value(axis_group, "Axes.Item axis group")?;
                let axis_index = self
                    .chart_axis_index_for_type(workbook, chart_id, axis_type, axis_group)?
                    .ok_or_else(|| {
                        OmError::invalid_argument("Axes.Item axis type is not available")
                    })?;
                Ok(OmValue::Object(
                    self.register_axis_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        axis_index,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Axes.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_adjustments(
        &mut self,
        workbook: WorkbookHandle,
        parent: ChartFormatParent,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        match member {
            "Item" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "Adjustments.Item expects a single index",
                    ));
                };
                self.chart_format_parent_object_kind(workbook, parent)?;
                let _ = coerce_u32_arg(index, "Adjustments.Item index")?;
                Err(OmError::invalid_argument(
                    "Adjustments.Item index is out of bounds",
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Adjustments.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_get_axis(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        axis_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("Axis", member) {
            self.focus_member_supported("Axis", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Axis.{member} does not accept arguments"
            )));
        }
        let axis = self.axis_model(workbook, chart_id, axis_index)?;

        match member {
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::Axis {
                        chart_id,
                        axis_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::Axis {
                        chart_id,
                        axis_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Type" => Ok(OmValue::Number(f64::from(match axis.kind {
                ChartAxisKind::Category | ChartAxisKind::Date => XL_CATEGORY,
                ChartAxisKind::Value => XL_VALUE,
                ChartAxisKind::Series => XL_SERIES_AXIS,
            }))),
            "AxisGroup" => Ok(OmValue::Number(f64::from(match axis.axis_group {
                ChartAxisGroup::Primary => XL_PRIMARY,
                ChartAxisGroup::Secondary => XL_SECONDARY,
            }))),
            "AxisBetweenCategories" => {
                if !matches!(axis.kind, ChartAxisKind::Category | ChartAxisKind::Date) {
                    return Err(OmError::unsupported(
                        "Axis.AxisBetweenCategories applies only to category axes",
                    ));
                }
                let chart = self.chart_model(workbook, chart_id)?;
                Ok(OmValue::Bool(
                    chart
                        .axes
                        .iter()
                        .find(|value_axis| {
                            value_axis.deleted != Some(true)
                                && value_axis.axis_group == axis.axis_group
                                && value_axis.kind == ChartAxisKind::Value
                        })
                        .and_then(|value_axis| value_axis.axis_between_categories)
                        .unwrap_or(true),
                ))
            }
            "CategoryType" => {
                if !matches!(axis.kind, ChartAxisKind::Category | ChartAxisKind::Date) {
                    return Err(OmError::unsupported(
                        "Axis.CategoryType applies only to category axes",
                    ));
                }
                Ok(OmValue::Number(f64::from(
                    if axis.category_type_auto == Some(true) {
                        XL_AUTOMATIC_SCALE
                    } else if axis.kind == ChartAxisKind::Date {
                        XL_TIME_SCALE
                    } else {
                        XL_CATEGORY_SCALE
                    },
                )))
            }
            "DisplayUnit" => {
                if axis.kind == ChartAxisKind::Series {
                    return Err(OmError::unsupported(
                        "Axis.DisplayUnit applies only to value or category axes",
                    ));
                }
                Ok(OmValue::Number(f64::from(match axis.display_unit {
                    Some(ChartAxisDisplayUnit::BuiltIn(value)) => {
                        chart_built_in_display_unit_to_excel_value(value)
                    }
                    Some(ChartAxisDisplayUnit::Custom(_)) => XL_DISPLAY_UNIT_CUSTOM,
                    None => XL_DISPLAY_UNIT_NONE,
                })))
            }
            "DisplayUnitCustom" => {
                if axis.kind != ChartAxisKind::Value {
                    return Err(OmError::unsupported(
                        "Axis.DisplayUnitCustom applies only to value axes",
                    ));
                }
                match axis.display_unit {
                    Some(ChartAxisDisplayUnit::Custom(value)) => Ok(OmValue::Number(value)),
                    Some(ChartAxisDisplayUnit::BuiltIn(_)) | None => Err(OmError::new(
                        OmErrorCode::NotFound,
                        "axis display unit is not custom",
                    )),
                }
            }
            "HasDisplayUnitLabel" => {
                if axis.kind != ChartAxisKind::Value {
                    return Err(OmError::unsupported(
                        "Axis.HasDisplayUnitLabel applies only to value axes",
                    ));
                }
                Ok(OmValue::Bool(
                    axis.display_unit.is_some() && axis.has_display_unit_label.unwrap_or(false),
                ))
            }
            "DisplayUnitLabel" => {
                if axis.kind != ChartAxisKind::Value {
                    return Err(OmError::unsupported(
                        "Axis.DisplayUnitLabel applies only to value axes",
                    ));
                }
                if axis.display_unit.is_none() || axis.has_display_unit_label != Some(true) {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "display unit label not found",
                    ));
                }
                Ok(OmValue::Object(self.register_object(
                    RuntimeObjectKind::DisplayUnitLabel {
                        workbook,
                        chart_id,
                        axis_index,
                        chart_object_parent,
                    },
                )))
            }
            "BaseUnit" => {
                if axis.kind != ChartAxisKind::Date {
                    return Err(OmError::unsupported(
                        "Axis.BaseUnit applies only to date category axes",
                    ));
                }
                Ok(OmValue::Number(f64::from(
                    chart_axis_time_unit_to_excel_value(
                        axis.base_unit.unwrap_or(ChartAxisTimeUnit::Days),
                    ),
                )))
            }
            "BaseUnitIsAuto" => {
                if axis.kind != ChartAxisKind::Date {
                    return Err(OmError::unsupported(
                        "Axis.BaseUnitIsAuto applies only to date category axes",
                    ));
                }
                Ok(OmValue::Bool(axis.base_unit.is_none()))
            }
            "HasTitle" => Ok(OmValue::Bool(axis.title.is_some())),
            "HasMajorGridlines" => Ok(OmValue::Bool(axis.has_major_gridlines == Some(true))),
            "HasMinorGridlines" => Ok(OmValue::Bool(axis.has_minor_gridlines == Some(true))),
            "ReversePlotOrder" => Ok(OmValue::Bool(axis.reverse_plot_order.unwrap_or(false))),
            "ScaleType" => Ok(OmValue::Number(f64::from(
                chart_axis_scale_type_to_excel_value(if axis.log_base.is_some() {
                    ChartAxisScaleType::Logarithmic
                } else {
                    axis.scale_type.unwrap_or(ChartAxisScaleType::Linear)
                }),
            ))),
            "LogBase" => axis
                .log_base
                .map(OmValue::Number)
                .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "axis scale is linear")),
            "Crosses" => Ok(OmValue::Number(f64::from(
                chart_axis_crosses_to_excel_value(if axis.crosses_at.is_some() {
                    ChartAxisCrosses::Custom
                } else {
                    axis.crosses.unwrap_or(ChartAxisCrosses::Automatic)
                }),
            ))),
            "CrossesAt" => axis
                .crosses_at
                .map(OmValue::Number)
                .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "axis crossing is automatic")),
            "MajorUnitScale" => {
                if axis.kind != ChartAxisKind::Date {
                    return Err(OmError::unsupported(
                        "Axis.MajorUnitScale applies only to date category axes",
                    ));
                }
                Ok(OmValue::Number(f64::from(
                    chart_axis_time_unit_to_excel_value(
                        axis.major_unit_scale.unwrap_or(ChartAxisTimeUnit::Days),
                    ),
                )))
            }
            "MinorUnitScale" => {
                if axis.kind != ChartAxisKind::Date {
                    return Err(OmError::unsupported(
                        "Axis.MinorUnitScale applies only to date category axes",
                    ));
                }
                Ok(OmValue::Number(f64::from(
                    chart_axis_time_unit_to_excel_value(
                        axis.minor_unit_scale.unwrap_or(ChartAxisTimeUnit::Days),
                    ),
                )))
            }
            "MajorTickMark" => Ok(OmValue::Number(f64::from(chart_tick_mark_to_excel_value(
                axis.major_tick_mark.unwrap_or(ChartTickMark::Outside),
            )))),
            "MinorTickMark" => Ok(OmValue::Number(f64::from(chart_tick_mark_to_excel_value(
                axis.minor_tick_mark.unwrap_or(ChartTickMark::None),
            )))),
            "TickLabelPosition" => Ok(OmValue::Number(f64::from(
                chart_tick_label_position_to_excel_value(
                    axis.tick_label_position
                        .unwrap_or(ChartTickLabelPosition::NextToAxis),
                ),
            ))),
            "TickLabels" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::TickLabels {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                },
            ))),
            "TickLabelSpacing" => Ok(OmValue::Number(f64::from(
                axis.tick_label_spacing.unwrap_or(1),
            ))),
            "TickLabelSpacingIsAuto" => Ok(OmValue::Bool(axis.tick_label_spacing.is_none())),
            "TickMarkSpacing" => Ok(OmValue::Number(f64::from(
                axis.tick_mark_spacing.unwrap_or(1),
            ))),
            "MinimumScale" => axis
                .minimum_scale
                .map(OmValue::Number)
                .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "axis minimum scale is auto")),
            "MaximumScale" => axis
                .maximum_scale
                .map(OmValue::Number)
                .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "axis maximum scale is auto")),
            "MajorUnit" => axis
                .major_unit
                .map(OmValue::Number)
                .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "axis major unit is auto")),
            "MinorUnit" => axis
                .minor_unit
                .map(OmValue::Number)
                .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "axis minor unit is auto")),
            "MinimumScaleIsAuto" => Ok(OmValue::Bool(axis.minimum_scale.is_none())),
            "MaximumScaleIsAuto" => Ok(OmValue::Bool(axis.maximum_scale.is_none())),
            "MajorUnitIsAuto" => Ok(OmValue::Bool(axis.major_unit.is_none())),
            "MinorUnitIsAuto" => Ok(OmValue::Bool(axis.minor_unit.is_none())),
            "MajorGridlines" | "MinorGridlines" => {
                let major = member == "MajorGridlines";
                let has_gridlines = if major {
                    axis.has_major_gridlines
                } else {
                    axis.has_minor_gridlines
                };
                if has_gridlines != Some(true) {
                    return Err(OmError::new(OmErrorCode::NotFound, "gridlines not found"));
                }
                Ok(OmValue::Object(self.register_object(
                    RuntimeObjectKind::Gridlines {
                        workbook,
                        chart_id,
                        axis_index,
                        major,
                        chart_object_parent,
                    },
                )))
            }
            "AxisTitle" => {
                if axis.title.is_none() {
                    return Err(OmError::new(OmErrorCode::NotFound, "axis title not found"));
                }
                Ok(OmValue::Object(self.register_object(
                    RuntimeObjectKind::AxisTitle {
                        workbook,
                        chart_id,
                        axis_index,
                        chart_object_parent,
                    },
                )))
            }
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_chart_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "Axis.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_tick_labels(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        axis_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("TickLabels", member) {
            self.focus_member_supported("TickLabels", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "TickLabels.{member} does not accept arguments"
            )));
        }
        let axis = self.axis_model(workbook, chart_id, axis_index)?;

        match member {
            "Name" => Ok(OmValue::Text("Tick Labels".to_string())),
            "AutoScaleFont" => Ok(OmValue::Bool(true)),
            "Depth" => Ok(OmValue::Number(1.0)),
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::TickLabels {
                        chart_id,
                        axis_index,
                        chart_object_parent,
                    },
                },
            ))),
            "NumberFormat" | "NumberFormatLocal" => Ok(OmValue::Text(
                axis.tick_label_number_format
                    .clone()
                    .unwrap_or_else(|| "General".to_string()),
            )),
            "MultiLevel" => Ok(OmValue::Bool(true)),
            "NumberFormatLinked" => Ok(OmValue::Bool(
                axis.tick_label_number_format_linked.unwrap_or(true),
            )),
            "Offset" => Ok(OmValue::Number(100.0)),
            "Orientation" => Ok(OmValue::Number(f64::from(
                XL_TICK_LABEL_ORIENTATION_AUTOMATIC,
            ))),
            "ReadingOrder" => Ok(OmValue::Number(f64::from(XL_READING_ORDER_CONTEXT))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Axis {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                },
            ))),
            _ => Err(OmError::unsupported(format!(
                "TickLabels.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_axis_title(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        axis_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("AxisTitle", member) {
            self.focus_member_supported("AxisTitle", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "AxisTitle.{member} does not accept arguments"
            )));
        }
        let title = self
            .axis_model(workbook, chart_id, axis_index)?
            .title
            .as_ref()
            .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "axis title not found"))?;

        match member {
            "Name" => Ok(OmValue::Text("Axis Title".to_string())),
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::AxisTitle {
                        chart_id,
                        axis_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::AxisTitle {
                        chart_id,
                        axis_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Text" | "Caption" => Ok(OmValue::Text(title.text.clone())),
            "Left" | "Top" | "Width" | "Height" => Ok(OmValue::Number(0.0)),
            "Orientation" => Ok(OmValue::Number(f64::from(XL_ORIENTATION_HORIZONTAL))),
            "ReadingOrder" => Ok(OmValue::Number(f64::from(XL_READING_ORDER_CONTEXT))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Axis {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                },
            ))),
            _ => Err(OmError::unsupported(format!(
                "AxisTitle.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_display_unit_label(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        axis_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("DisplayUnitLabel", member) {
            self.focus_member_supported("DisplayUnitLabel", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "DisplayUnitLabel.{member} does not accept arguments"
            )));
        }
        let axis = self.axis_model(workbook, chart_id, axis_index)?;
        if axis.kind != ChartAxisKind::Value {
            return Err(OmError::unsupported(
                "DisplayUnitLabel applies only to value axes",
            ));
        }
        if axis.display_unit.is_none() || axis.has_display_unit_label != Some(true) {
            return Err(OmError::new(
                OmErrorCode::NotFound,
                "display unit label not found",
            ));
        }

        match member {
            "Name" => Ok(OmValue::Text("Display Unit Label".to_string())),
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::DisplayUnitLabel {
                        chart_id,
                        axis_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::DisplayUnitLabel {
                        chart_id,
                        axis_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Text" | "Caption" => Ok(OmValue::Text(chart_axis_display_unit_label_text(axis))),
            "Left" | "Top" | "Width" | "Height" => Ok(OmValue::Number(0.0)),
            "Orientation" => Ok(OmValue::Number(f64::from(XL_ORIENTATION_HORIZONTAL))),
            "ReadingOrder" => Ok(OmValue::Number(f64::from(XL_READING_ORDER_CONTEXT))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Axis {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                },
            ))),
            _ => Err(OmError::unsupported(format!(
                "DisplayUnitLabel.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_gridlines(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        axis_index: usize,
        major: bool,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("Gridlines", member) {
            self.focus_member_supported("Gridlines", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Gridlines.{member} does not accept arguments"
            )));
        }
        let axis = self.axis_model(workbook, chart_id, axis_index)?;
        let has_gridlines = if major {
            axis.has_major_gridlines
        } else {
            axis.has_minor_gridlines
        };
        if has_gridlines != Some(true) {
            return Err(OmError::new(OmErrorCode::NotFound, "gridlines not found"));
        }

        match member {
            "Name" => Ok(OmValue::Text(if major {
                "Major Gridlines".to_string()
            } else {
                "Minor Gridlines".to_string()
            })),
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::Gridlines {
                        chart_id,
                        axis_index,
                        major,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::Gridlines {
                        chart_id,
                        axis_index,
                        major,
                        chart_object_parent,
                    },
                },
            ))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Axis {
                    workbook,
                    chart_id,
                    axis_index,
                    chart_object_parent,
                },
            ))),
            _ => Err(OmError::unsupported(format!(
                "Gridlines.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_chart_title(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "ChartTitle.{member} does not accept arguments"
            )));
        }
        let title = self
            .chart_model(workbook, chart_id)?
            .title
            .as_ref()
            .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart title not found"))?;

        match member {
            "Name" => Ok(OmValue::Text("Chart Title".to_string())),
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::ChartTitle {
                        chart_id,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::ChartTitle {
                        chart_id,
                        chart_object_parent,
                    },
                },
            ))),
            "Text" | "Caption" => Ok(OmValue::Text(title.text.clone())),
            "Left" | "Top" | "Width" | "Height" => Ok(OmValue::Number(0.0)),
            "Orientation" => Ok(OmValue::Number(f64::from(XL_ORIENTATION_HORIZONTAL))),
            "ReadingOrder" => Ok(OmValue::Number(f64::from(XL_READING_ORDER_CONTEXT))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_chart_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "ChartTitle.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_series_collection(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        axis_group_filter: Option<ChartAxisGroup>,
        group_index_filter: Option<usize>,
        full: bool,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("SeriesCollection", member) {
            self.focus_member_supported("SeriesCollection", member, false)?;
        }

        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "SeriesCollection.Count does not accept arguments",
                    ));
                }
                let chart = self.chart_model(workbook, chart_id)?;
                let count =
                    series_collection_indices(chart, axis_group_filter, group_index_filter, full)
                        .len();
                Ok(OmValue::Number(count as f64))
            }
            "Item" => self.dispatch_invoke_series_collection(
                workbook,
                chart_id,
                axis_group_filter,
                group_index_filter,
                full,
                chart_object_parent,
                member,
                args,
            ),
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "SeriesCollection.Creator does not accept arguments",
                    ));
                }
                self.chart_model(workbook, chart_id)?;
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "SeriesCollection.Application does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "SeriesCollection.Parent does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(
                    self.register_chart_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "SeriesCollection.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_series_collection(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        axis_group_filter: Option<ChartAxisGroup>,
        group_index_filter: Option<usize>,
        full: bool,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("SeriesCollection", member) {
            self.focus_member_supported("SeriesCollection", member, false)?;
        }

        match member {
            "Item" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "SeriesCollection.Item expects a single 1-based index or series name",
                    ));
                };
                let index = match index {
                    OmValue::Number(_) => {
                        let index = coerce_u32_arg(index, "SeriesCollection.Item index")? as usize;
                        let chart = self.chart_model(workbook, chart_id)?;
                        let matching_indices = series_collection_indices(
                            chart,
                            axis_group_filter,
                            group_index_filter,
                            full,
                        );
                        if index == 0 || index > matching_indices.len() {
                            return Err(OmError::invalid_argument(
                                "SeriesCollection.Item index is out of bounds",
                            ));
                        }
                        matching_indices[index - 1]
                    }
                    OmValue::Text(name) => {
                        let lookup = name.trim();
                        let state = &self.runtime_workbook(workbook)?.loaded.state;
                        let chart = state.charts.get(&chart_id).ok_or_else(|| {
                            OmError::new(OmErrorCode::NotFound, "chart not found")
                        })?;
                        let mut matched_index = None;
                        for series_index in series_collection_indices(
                            chart,
                            axis_group_filter,
                            group_index_filter,
                            full,
                        ) {
                            let series = &chart.series[series_index];
                            let Some(source) = series.name.as_ref() else {
                                continue;
                            };
                            let raw = source.raw.text.trim();
                            let raw_without_equals = raw.trim_start_matches('=');
                            let display_name = chart_source_value_text_for_index(state, source, 0);
                            let formula_name = chart_source_expr_text(Some(source));
                            if raw.eq_ignore_ascii_case(lookup)
                                || raw_without_equals.eq_ignore_ascii_case(lookup)
                                || display_name
                                    .as_deref()
                                    .is_some_and(|value| value.eq_ignore_ascii_case(lookup))
                                || formula_name
                                    .as_deref()
                                    .is_some_and(|value| value.eq_ignore_ascii_case(lookup))
                            {
                                matched_index = Some(series_index);
                                break;
                            }
                        }
                        matched_index.ok_or_else(|| {
                            OmError::new(
                                OmErrorCode::NotFound,
                                "SeriesCollection.Item series name not found",
                            )
                        })?
                    }
                    _ => {
                        return Err(OmError::type_mismatch(
                            "SeriesCollection.Item expects a numeric index or series name",
                        ));
                    }
                };
                Ok(OmValue::Object(
                    self.register_series_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        index,
                        chart_object_parent,
                    ),
                ))
            }
            "NewSeries" => {
                if full {
                    return Err(OmError::unsupported(
                        "FullSeriesCollection.NewSeries is not supported",
                    ));
                }
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "SeriesCollection.NewSeries does not accept arguments",
                    ));
                }
                let axis_group = axis_group_filter.unwrap_or(ChartAxisGroup::Primary);
                let target_group_index = {
                    let chart = self.chart_model(workbook, chart_id)?;
                    resolve_series_insert_group(chart, axis_group, group_index_filter)?
                };
                if target_group_index.is_none() && axis_group == ChartAxisGroup::Secondary {
                    self.set_chart_axis_presence(
                        workbook,
                        chart_id,
                        XL_VALUE,
                        ChartAxisGroup::Secondary,
                        true,
                    )?;
                }
                let series_index = {
                    let runtime = self.runtime_workbook_mut(workbook)?;
                    if runtime.read_only {
                        return Err(OmError::new(
                            OmErrorCode::InvalidState,
                            "cannot modify a read-only workbook",
                        ));
                    }
                    let chart = runtime
                        .loaded
                        .state
                        .charts
                        .get_mut(&chart_id)
                        .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart not found"))?;
                    let series_index = chart.series.len();
                    let raw_index = next_chart_series_raw_index(chart)?;
                    chart.series.push(SeriesModel {
                        name: None,
                        x_values: None,
                        values: None,
                        bubble_size: None,
                        bar_shape: None,
                        smooth: None,
                        marker_style: None,
                        marker_size: None,
                        invert_if_negative: None,
                        points: BTreeMap::new(),
                        data_labels: None,
                        point_data_labels: BTreeMap::new(),
                        raw_index: Some(raw_index),
                        order: u32::try_from(series_index).ok(),
                        axis_group,
                        is_filtered: false,
                        filter_dirty: false,
                    });
                    if let Some(group_index) = target_group_index {
                        attach_series_to_chart_group(chart, group_index, raw_index);
                    }
                    if let Some(plot_order) = series_index
                        .checked_add(1)
                        .and_then(|index| u32::try_from(index).ok())
                    {
                        update_series_plot_order(&mut chart.series, series_index, plot_order);
                    }
                    normalize_volume_stock_chart(chart);
                    chart.content_dirty = true;
                    chart.dirty = true;
                    runtime.prompt_dirty = true;
                    series_index
                };
                self.find_state = None;
                self.cut_copy_mode = None;
                self.clipboard = None;
                Ok(OmValue::Object(
                    self.register_series_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        series_index,
                        chart_object_parent,
                    ),
                ))
            }
            "Add" => {
                if full {
                    return Err(OmError::unsupported(
                        "FullSeriesCollection.Add is not supported",
                    ));
                }
                if args.is_empty() || args.len() > 5 {
                    return Err(OmError::invalid_argument(
                        "SeriesCollection.Add expects Source and optional Rowcol, SeriesLabels, CategoryLabels, and Replace arguments",
                    ));
                }
                let plot_by =
                    chart_plot_by_from_optional_arg(args.get(1), "SeriesCollection.Add Rowcol")?;
                let series_labels = args
                    .get(2)
                    .map(|value| {
                        coerce_optional_bool_arg(value, false, "SeriesCollection.Add SeriesLabels")
                    })
                    .transpose()?
                    .unwrap_or(false);
                let category_labels = args
                    .get(3)
                    .map(|value| {
                        coerce_optional_bool_arg(
                            value,
                            false,
                            "SeriesCollection.Add CategoryLabels",
                        )
                    })
                    .transpose()?
                    .unwrap_or(false);
                let replace = args
                    .get(4)
                    .map(|value| {
                        coerce_optional_bool_arg(value, false, "SeriesCollection.Add Replace")
                    })
                    .transpose()?
                    .unwrap_or(false);
                let source_range =
                    self.chart_source_range_from_arg(workbook, &args[0], "SeriesCollection.Add")?;
                let chart = self.chart_model(workbook, chart_id)?;
                let existing_series_len = chart.series.len();
                let first_raw_index = next_chart_series_raw_index(chart)?;
                let target_axis_group = axis_group_filter.unwrap_or(ChartAxisGroup::Primary);
                let target_group_index =
                    resolve_series_insert_group(chart, target_axis_group, group_index_filter)?;
                let mut new_series = self.chart_series_models_for_range_source(
                    workbook,
                    &source_range,
                    plot_by,
                    existing_series_len,
                    first_raw_index,
                    series_labels,
                    category_labels,
                )?;
                if target_group_index.is_none() && target_axis_group == ChartAxisGroup::Secondary {
                    self.set_chart_axis_presence(
                        workbook,
                        chart_id,
                        XL_VALUE,
                        ChartAxisGroup::Secondary,
                        true,
                    )?;
                }
                if axis_group_filter.is_some() || target_group_index.is_some() {
                    for series in &mut new_series {
                        series.axis_group = target_axis_group;
                    }
                }
                let replacement_category_sources = if replace {
                    if !category_labels {
                        return Err(OmError::invalid_argument(
                            "SeriesCollection.Add Replace requires CategoryLabels",
                        ));
                    }
                    let category_sources = new_series
                        .iter()
                        .filter_map(|series| series.x_values.clone())
                        .collect::<Vec<_>>();
                    if category_sources.is_empty() {
                        return Err(OmError::invalid_argument(
                            "SeriesCollection.Add Replace requires category label source data",
                        ));
                    }
                    Some(category_sources)
                } else {
                    None
                };
                let first_new_series_index = {
                    let runtime = self.runtime_workbook_mut(workbook)?;
                    if runtime.read_only {
                        return Err(OmError::new(
                            OmErrorCode::InvalidState,
                            "cannot modify a read-only workbook",
                        ));
                    }
                    let chart = runtime
                        .loaded
                        .state
                        .charts
                        .get_mut(&chart_id)
                        .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "chart not found"))?;
                    if let Some(category_sources) = replacement_category_sources
                        && !chart.series.is_empty()
                    {
                        let first_category_source = &category_sources[0];
                        if category_sources
                            .iter()
                            .all(|source| source.raw.text == first_category_source.raw.text)
                        {
                            for series in &mut chart.series {
                                series.x_values = Some(first_category_source.clone());
                            }
                        } else if category_sources.len() == chart.series.len() {
                            for (series, category_source) in
                                chart.series.iter_mut().zip(category_sources.iter())
                            {
                                series.x_values = Some(category_source.clone());
                            }
                        } else {
                            return Err(OmError::unsupported(
                                "SeriesCollection.Add Replace with multiple category label ranges requires one range or one range per existing series",
                            ));
                        }
                    }
                    let first_new_series_index = chart.series.len();
                    let new_raw_indices = new_series
                        .iter()
                        .filter_map(|series| series.raw_index)
                        .collect::<Vec<_>>();
                    chart.series.append(&mut new_series);
                    if let Some(group_index) = target_group_index {
                        for raw_index in new_raw_indices {
                            attach_series_to_chart_group(chart, group_index, raw_index);
                        }
                    }
                    normalize_volume_stock_chart(chart);
                    chart.content_dirty = true;
                    chart.dirty = true;
                    runtime.prompt_dirty = true;
                    first_new_series_index
                };
                self.find_state = None;
                self.cut_copy_mode = None;
                self.clipboard = None;
                Ok(OmValue::Object(
                    self.register_series_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        first_new_series_index,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "SeriesCollection.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn chart_source_range_from_arg(
        &self,
        workbook: WorkbookHandle,
        value: &OmValue,
        owner: &str,
    ) -> OmResult<RangeSet> {
        match value {
            OmValue::Object(handle) => match self.runtime_object(*handle)? {
                RuntimeObjectKind::Range {
                    workbook: range_workbook,
                    range,
                    ..
                } => {
                    if range_workbook != workbook {
                        return Err(OmError::unsupported(format!(
                            "{owner} cross-workbook ranges are not supported"
                        )));
                    }
                    for area in range.areas() {
                        if matches!(area.scope, SheetScope::Multi3D { .. }) {
                            return Err(OmError::unsupported(format!(
                                "{owner} 3D ranges are not supported"
                            )));
                        }
                    }
                    Ok(range)
                }
                _ => Err(OmError::type_mismatch(format!(
                    "{owner} Source expects a Range object"
                ))),
            },
            _ => Err(OmError::type_mismatch(format!(
                "{owner} Source expects a Range object"
            ))),
        }
    }

    pub(crate) fn chart_series_models_for_range_source(
        &self,
        workbook: WorkbookHandle,
        source_range: &RangeSet,
        plot_by: Option<i32>,
        first_order: usize,
        first_raw_index: u32,
        series_labels: bool,
        category_labels: bool,
    ) -> OmResult<Vec<SeriesModel>> {
        let workbook_id = self.workbook_model(workbook)?.id;
        let mut new_series = Vec::new();
        let plot_by = if plot_by.is_none() && (series_labels || category_labels) {
            Some(XL_PLOT_BY_COLUMNS)
        } else {
            plot_by
        };
        let mut make_series = |name: Option<ChartSourceExpr>,
                               x_values: Option<ChartSourceExpr>,
                               values: ChartSourceExpr| {
            let order = u32::try_from(first_order + new_series.len()).ok();
            let raw_index = u32::try_from(new_series.len())
                .ok()
                .and_then(|offset| first_raw_index.checked_add(offset));
            new_series.push(SeriesModel {
                name,
                x_values,
                values: Some(values),
                bubble_size: None,
                bar_shape: None,
                smooth: None,
                marker_style: None,
                marker_size: None,
                invert_if_negative: None,
                points: BTreeMap::new(),
                data_labels: None,
                point_data_labels: BTreeMap::new(),
                raw_index,
                order,
                axis_group: ChartAxisGroup::Primary,
                is_filtered: false,
                filter_dirty: false,
            });
        };
        match plot_by {
            Some(XL_PLOT_BY_ROWS) => {
                for area in source_range.areas() {
                    let SheetScope::Single(source_sheet_id) = area.scope else {
                        unreachable!("3D ranges were rejected above");
                    };
                    let worksheet_name = self
                        .worksheet_model(workbook, source_sheet_id)?
                        .name
                        .clone();
                    let data_row_first = if category_labels {
                        area.rect.row_first.checked_add(1).ok_or_else(|| {
                            OmError::invalid_argument(
                                "SeriesCollection.Add category labels leave no data rows",
                            )
                        })?
                    } else {
                        area.rect.row_first
                    };
                    let data_col_first = if series_labels {
                        area.rect.col_first.checked_add(1).ok_or_else(|| {
                            OmError::invalid_argument(
                                "SeriesCollection.Add series labels leave no data columns",
                            )
                        })?
                    } else {
                        area.rect.col_first
                    };
                    if data_row_first > area.rect.row_last {
                        return Err(OmError::invalid_argument(
                            "SeriesCollection.Add category labels leave no data rows",
                        ));
                    }
                    if data_col_first > area.rect.col_last {
                        return Err(OmError::invalid_argument(
                            "SeriesCollection.Add series labels leave no data columns",
                        ));
                    }
                    let x_values = if category_labels {
                        Some(chart_source_expr_for_range(
                            workbook_id,
                            source_sheet_id,
                            Rect {
                                row_first: area.rect.row_first,
                                row_last: area.rect.row_first,
                                col_first: data_col_first,
                                col_last: area.rect.col_last,
                            },
                            &worksheet_name,
                        )?)
                    } else {
                        None
                    };
                    for row in data_row_first..=area.rect.row_last {
                        let name = if series_labels {
                            Some(chart_source_expr_for_range(
                                workbook_id,
                                source_sheet_id,
                                Rect {
                                    row_first: row,
                                    row_last: row,
                                    col_first: area.rect.col_first,
                                    col_last: area.rect.col_first,
                                },
                                &worksheet_name,
                            )?)
                        } else {
                            None
                        };
                        let values_rect = Rect {
                            row_first: row,
                            row_last: row,
                            col_first: data_col_first,
                            col_last: area.rect.col_last,
                        };
                        make_series(
                            name,
                            x_values.clone(),
                            chart_source_expr_for_range(
                                workbook_id,
                                source_sheet_id,
                                values_rect,
                                &worksheet_name,
                            )?,
                        );
                    }
                }
            }
            Some(XL_PLOT_BY_COLUMNS) => {
                for area in source_range.areas() {
                    let SheetScope::Single(source_sheet_id) = area.scope else {
                        unreachable!("3D ranges were rejected above");
                    };
                    let worksheet_name = self
                        .worksheet_model(workbook, source_sheet_id)?
                        .name
                        .clone();
                    let data_row_first = if series_labels {
                        area.rect.row_first.checked_add(1).ok_or_else(|| {
                            OmError::invalid_argument(
                                "SeriesCollection.Add series labels leave no data rows",
                            )
                        })?
                    } else {
                        area.rect.row_first
                    };
                    let data_col_first = if category_labels {
                        area.rect.col_first.checked_add(1).ok_or_else(|| {
                            OmError::invalid_argument(
                                "SeriesCollection.Add category labels leave no data columns",
                            )
                        })?
                    } else {
                        area.rect.col_first
                    };
                    if data_row_first > area.rect.row_last {
                        return Err(OmError::invalid_argument(
                            "SeriesCollection.Add series labels leave no data rows",
                        ));
                    }
                    if data_col_first > area.rect.col_last {
                        return Err(OmError::invalid_argument(
                            "SeriesCollection.Add category labels leave no data columns",
                        ));
                    }
                    let x_values = if category_labels {
                        Some(chart_source_expr_for_range(
                            workbook_id,
                            source_sheet_id,
                            Rect {
                                row_first: data_row_first,
                                row_last: area.rect.row_last,
                                col_first: area.rect.col_first,
                                col_last: area.rect.col_first,
                            },
                            &worksheet_name,
                        )?)
                    } else {
                        None
                    };
                    for col in data_col_first..=area.rect.col_last {
                        let name = if series_labels {
                            Some(chart_source_expr_for_range(
                                workbook_id,
                                source_sheet_id,
                                Rect {
                                    row_first: area.rect.row_first,
                                    row_last: area.rect.row_first,
                                    col_first: col,
                                    col_last: col,
                                },
                                &worksheet_name,
                            )?)
                        } else {
                            None
                        };
                        let values_rect = Rect {
                            row_first: data_row_first,
                            row_last: area.rect.row_last,
                            col_first: col,
                            col_last: col,
                        };
                        make_series(
                            name,
                            x_values.clone(),
                            chart_source_expr_for_range(
                                workbook_id,
                                source_sheet_id,
                                values_rect,
                                &worksheet_name,
                            )?,
                        );
                    }
                }
            }
            Some(_) => unreachable!("unsupported PlotBy was rejected"),
            None => {
                let mut areas = Vec::with_capacity(source_range.len());
                for area in source_range.areas() {
                    let SheetScope::Single(source_sheet_id) = area.scope else {
                        unreachable!("3D ranges were rejected above");
                    };
                    let worksheet_name = self
                        .worksheet_model(workbook, source_sheet_id)?
                        .name
                        .clone();
                    areas.push((source_sheet_id, area.rect, worksheet_name));
                }
                make_series(
                    None,
                    None,
                    chart_source_expr_for_range_areas(workbook_id, areas)?,
                );
            }
        }
        Ok(new_series)
    }

    pub(crate) fn dispatch_get_series(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        series_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("Series", member) {
            self.focus_member_supported("Series", member, false)?;
        }

        if !args.is_empty() && member != "DataLabels" && member != "Points" {
            return Err(OmError::invalid_argument(format!(
                "Series.{member} does not accept arguments"
            )));
        }

        match member {
            "Name" => Ok(chart_source_expr_text(
                self.series_model(workbook, chart_id, series_index)?
                    .name
                    .as_ref(),
            )
            .map(OmValue::Text)
            .unwrap_or(OmValue::Empty)),
            "Format" => {
                self.series_model(workbook, chart_id, series_index)?;
                Ok(OmValue::Object(self.register_object(
                    RuntimeObjectKind::ChartFormat {
                        workbook,
                        parent: ChartFormatParent::Series {
                            chart_id,
                            series_index,
                            chart_object_parent,
                        },
                    },
                )))
            }
            "ChartType" => {
                let chart = self.chart_model(workbook, chart_id)?;
                let Some(series) = chart.series.get(series_index) else {
                    return Err(OmError::new(OmErrorCode::NotFound, "series not found"));
                };
                let series_chart_type = chart_type_for_series(chart, series);
                Ok(OmValue::Number(f64::from(chart_type_to_excel_value(
                    &series_chart_type,
                )?)))
            }
            "Values" => Ok(chart_source_expr_text(
                self.series_model(workbook, chart_id, series_index)?
                    .values
                    .as_ref(),
            )
            .map(OmValue::Text)
            .unwrap_or(OmValue::Empty)),
            "XValues" => Ok(chart_source_expr_text(
                self.series_model(workbook, chart_id, series_index)?
                    .x_values
                    .as_ref(),
            )
            .map(OmValue::Text)
            .unwrap_or(OmValue::Empty)),
            "BubbleSizes" => Ok(chart_source_expr_text(
                self.series_model(workbook, chart_id, series_index)?
                    .bubble_size
                    .as_ref(),
            )
            .map(OmValue::Text)
            .unwrap_or(OmValue::Empty)),
            "BarShape" => {
                let chart = self.chart_model(workbook, chart_id)?;
                ensure_chart_supports_bar_shape(&chart.chart_type)?;
                let series = chart
                    .series
                    .get(series_index)
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "series not found"))?;
                Ok(OmValue::Number(f64::from(
                    series
                        .bar_shape
                        .or_else(|| chart_effective_bar_shape(chart))
                        .map(chart_bar_shape_to_excel_value)
                        .unwrap_or(XL_BOX),
                )))
            }
            "Smooth" => {
                let chart = self.chart_model(workbook, chart_id)?;
                if !chart_type_supports_series_smooth(&chart.chart_type) {
                    return Err(OmError::unsupported(
                        "Series.Smooth is only supported for line and scatter chart types",
                    ));
                }
                let series = chart
                    .series
                    .get(series_index)
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "series not found"))?;
                Ok(OmValue::Bool(series.smooth.unwrap_or_else(|| {
                    chart_type_default_series_smooth(&chart.chart_type)
                })))
            }
            "MarkerStyle" => {
                let chart = self.chart_model(workbook, chart_id)?;
                if !chart_type_supports_series_marker(&chart.chart_type) {
                    return Err(OmError::unsupported(
                        "Series.MarkerStyle is only supported for line, scatter, and radar chart types",
                    ));
                }
                let series = chart
                    .series
                    .get(series_index)
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "series not found"))?;
                Ok(OmValue::Number(f64::from(
                    series
                        .marker_style
                        .map(chart_marker_style_to_excel_value)
                        .unwrap_or(XL_MARKER_STYLE_AUTOMATIC),
                )))
            }
            "MarkerSize" => {
                let chart = self.chart_model(workbook, chart_id)?;
                if !chart_type_supports_series_marker(&chart.chart_type) {
                    return Err(OmError::unsupported(
                        "Series.MarkerSize is only supported for line, scatter, and radar chart types",
                    ));
                }
                let series = chart
                    .series
                    .get(series_index)
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "series not found"))?;
                Ok(OmValue::Number(f64::from(series.marker_size.unwrap_or(5))))
            }
            "InvertIfNegative" => {
                let series = self.series_model(workbook, chart_id, series_index)?;
                Ok(OmValue::Bool(series.invert_if_negative.unwrap_or(false)))
            }
            "IsFiltered" => Ok(OmValue::Bool(
                self.series_model(workbook, chart_id, series_index)?
                    .is_filtered,
            )),
            "Formula" => {
                let series = self.series_model(workbook, chart_id, series_index)?;
                Ok(OmValue::Text(series_formula_text(series, series_index)))
            }
            "AxisGroup" => {
                let series = self.series_model(workbook, chart_id, series_index)?;
                Ok(OmValue::Number(f64::from(match series.axis_group {
                    ChartAxisGroup::Primary => XL_PRIMARY,
                    ChartAxisGroup::Secondary => XL_SECONDARY,
                })))
            }
            "HasDataLabels" => Ok(OmValue::Bool(chart_data_labels_visible(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                ),
            ))),
            "HasLeaderLines" => Ok(OmValue::Bool(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )
                .and_then(|labels| labels.has_leader_lines)
                .unwrap_or(false),
            )),
            "LeaderLines" => {
                if !self.leader_lines_visible(workbook, chart_id, series_index)? {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        "leader lines not found",
                    ));
                }
                Ok(OmValue::Object(
                    self.register_leader_lines_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        series_index,
                        chart_object_parent,
                    ),
                ))
            }
            "DataLabels" => {
                self.series_model(workbook, chart_id, series_index)?;
                let handle = self.register_data_labels_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "Points" => {
                self.series_model(workbook, chart_id, series_index)?;
                let handle = self.register_points_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "PlotOrder" => Ok(OmValue::Number(f64::from(series_plot_order(
                self.series_model(workbook, chart_id, series_index)?,
                series_index,
            )))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => {
                let is_filtered = self
                    .series_model(workbook, chart_id, series_index)?
                    .is_filtered;
                let parent = if is_filtered {
                    self.register_full_series_collection_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        chart_object_parent,
                    )
                } else {
                    self.register_series_collection_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        chart_object_parent,
                    )
                };
                Ok(OmValue::Object(parent))
            }
            _ => Err(OmError::unsupported(format!(
                "Series.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_leader_lines(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        series_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("LeaderLines", member) {
            self.focus_member_supported("LeaderLines", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "LeaderLines.{member} does not accept arguments"
            )));
        }
        self.leader_lines_visible(workbook, chart_id, series_index)?;
        match member {
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::LeaderLines {
                        chart_id,
                        series_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::LeaderLines {
                        chart_id,
                        series_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_series_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "LeaderLines.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_data_labels(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        series_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("DataLabels", member) {
            self.focus_member_supported("DataLabels", member, false)?;
        }

        if !args.is_empty() && member != "Item" {
            return Err(OmError::invalid_argument(format!(
                "DataLabels.{member} does not accept arguments"
            )));
        }

        match member {
            "Item" => self.dispatch_invoke_data_labels(
                workbook,
                chart_id,
                series_index,
                chart_object_parent,
                member,
                args,
            ),
            "Name" => Ok(OmValue::Text("Data Labels".to_string())),
            "Format" => {
                self.series_model(workbook, chart_id, series_index)?;
                Ok(OmValue::Object(self.register_object(
                    RuntimeObjectKind::ChartFormat {
                        workbook,
                        parent: ChartFormatParent::DataLabels {
                            chart_id,
                            series_index,
                            chart_object_parent,
                        },
                    },
                )))
            }
            "Border" => {
                self.series_model(workbook, chart_id, series_index)?;
                Ok(OmValue::Object(self.register_object(
                    RuntimeObjectKind::Border {
                        workbook,
                        parent: BorderParent::DataLabels {
                            chart_id,
                            series_index,
                            chart_object_parent,
                        },
                    },
                )))
            }
            "Count" => {
                let chart = self.chart_model(workbook, chart_id)?;
                if chart.series.get(series_index).is_none() {
                    return Err(OmError::new(OmErrorCode::NotFound, "series not found"));
                }
                Ok(OmValue::Number(
                    chart_data_labels_count_for_chart_series(chart, series_index) as f64,
                ))
            }
            "Type" => Ok(OmValue::Number(f64::from(
                chart_data_labels_type_to_excel_value(chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )),
            ))),
            "ShowLegendKey" => Ok(OmValue::Bool(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )
                .and_then(|labels| labels.show_legend_key)
                .unwrap_or(false),
            )),
            "HasLeaderLines" => Ok(OmValue::Bool(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )
                .and_then(|labels| labels.has_leader_lines)
                .unwrap_or(false),
            )),
            "ShowSeriesName" => Ok(OmValue::Bool(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )
                .and_then(|labels| labels.show_series_name)
                .unwrap_or(false),
            )),
            "ShowCategoryName" => Ok(OmValue::Bool(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )
                .and_then(|labels| labels.show_category_name)
                .unwrap_or(false),
            )),
            "ShowValue" => Ok(OmValue::Bool(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )
                .and_then(|labels| labels.show_value)
                .unwrap_or(false),
            )),
            "ShowPercentage" => Ok(OmValue::Bool(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )
                .and_then(|labels| labels.show_percentage)
                .unwrap_or(false),
            )),
            "ShowBubbleSize" => Ok(OmValue::Bool(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )
                .and_then(|labels| labels.show_bubble_size)
                .unwrap_or(false),
            )),
            "NumberFormat" | "NumberFormatLocal" => Ok(OmValue::Text(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )
                .and_then(|labels| labels.number_format.clone())
                .unwrap_or_else(|| "General".to_string()),
            )),
            "NumberFormatLinked" => Ok(OmValue::Bool(
                chart_series_effective_data_labels(
                    self.chart_model(workbook, chart_id)?,
                    series_index,
                )
                .and_then(|labels| labels.number_format_linked)
                .unwrap_or(true),
            )),
            "Position" => Ok(OmValue::Number(f64::from(
                chart_data_label_position_to_excel_value(
                    chart_series_effective_data_labels(
                        self.chart_model(workbook, chart_id)?,
                        series_index,
                    )
                    .and_then(|labels| labels.position)
                    .unwrap_or(ChartDataLabelPosition::BestFit),
                ),
            ))),
            "Separator" => Ok(chart_series_effective_data_labels(
                self.chart_model(workbook, chart_id)?,
                series_index,
            )
            .and_then(|labels| labels.separator.clone())
            .map(OmValue::Text)
            .unwrap_or(OmValue::Empty)),
            "Orientation" => Ok(OmValue::Number(f64::from(XL_ORIENTATION_HORIZONTAL))),
            "ReadingOrder" => Ok(OmValue::Number(f64::from(XL_READING_ORDER_CONTEXT))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_series_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "DataLabels.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_data_labels(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        series_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("DataLabels", member) {
            self.focus_member_supported("DataLabels", member, false)?;
        }

        match member {
            "Item" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "DataLabels.Item expects a single 1-based index",
                    ));
                };
                let index = coerce_u32_arg(index, "DataLabels.Item index")? as usize;
                let chart = self.chart_model(workbook, chart_id)?;
                if chart.series.get(series_index).is_none() {
                    return Err(OmError::new(OmErrorCode::NotFound, "series not found"));
                }
                let label_count = chart_data_labels_count_for_chart_series(chart, series_index);
                if index == 0 || index > label_count {
                    return Err(OmError::invalid_argument(
                        "DataLabels.Item index is out of bounds",
                    ));
                }
                Ok(OmValue::Object(
                    self.register_data_label_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        series_index,
                        index - 1,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "DataLabels.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_get_data_label(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        series_index: usize,
        point_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("DataLabel", member) {
            self.focus_member_supported("DataLabel", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "DataLabel.{member} does not accept arguments"
            )));
        }
        let data_labels = self
            .validate_data_label_index(workbook, chart_id, series_index, point_index)?
            .cloned();

        match member {
            "Name" => Ok(OmValue::Text(format!("Data Label {}", point_index + 1))),
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::DataLabel {
                        chart_id,
                        series_index,
                        point_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::DataLabel {
                        chart_id,
                        series_index,
                        point_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Index" => Ok(OmValue::Number((point_index + 1) as f64)),
            "Type" => Ok(OmValue::Number(f64::from(
                chart_data_labels_type_to_excel_value(data_labels.as_ref()),
            ))),
            "ShowLegendKey" => Ok(OmValue::Bool(
                data_labels
                    .as_ref()
                    .and_then(|labels| labels.show_legend_key)
                    .unwrap_or(false),
            )),
            "HasLeaderLines" => Ok(OmValue::Bool(
                data_labels
                    .as_ref()
                    .and_then(|labels| labels.has_leader_lines)
                    .unwrap_or(false),
            )),
            "ShowSeriesName" => Ok(OmValue::Bool(
                data_labels
                    .as_ref()
                    .and_then(|labels| labels.show_series_name)
                    .unwrap_or(false),
            )),
            "ShowCategoryName" => Ok(OmValue::Bool(
                data_labels
                    .as_ref()
                    .and_then(|labels| labels.show_category_name)
                    .unwrap_or(false),
            )),
            "ShowValue" => Ok(OmValue::Bool(
                data_labels
                    .as_ref()
                    .and_then(|labels| labels.show_value)
                    .unwrap_or(false),
            )),
            "ShowPercentage" => Ok(OmValue::Bool(
                data_labels
                    .as_ref()
                    .and_then(|labels| labels.show_percentage)
                    .unwrap_or(false),
            )),
            "ShowBubbleSize" => Ok(OmValue::Bool(
                data_labels
                    .as_ref()
                    .and_then(|labels| labels.show_bubble_size)
                    .unwrap_or(false),
            )),
            "NumberFormat" | "NumberFormatLocal" => Ok(OmValue::Text(
                data_labels
                    .as_ref()
                    .and_then(|labels| labels.number_format.clone())
                    .unwrap_or_else(|| "General".to_string()),
            )),
            "NumberFormatLinked" => Ok(OmValue::Bool(
                data_labels
                    .as_ref()
                    .and_then(|labels| labels.number_format_linked)
                    .unwrap_or(true),
            )),
            "Position" => Ok(OmValue::Number(f64::from(
                chart_data_label_position_to_excel_value(
                    data_labels
                        .as_ref()
                        .and_then(|labels| labels.position)
                        .unwrap_or(ChartDataLabelPosition::BestFit),
                ),
            ))),
            "Separator" => Ok(data_labels
                .as_ref()
                .and_then(|labels| labels.separator.clone())
                .map(OmValue::Text)
                .unwrap_or(OmValue::Empty)),
            "Orientation" => Ok(OmValue::Number(f64::from(XL_ORIENTATION_HORIZONTAL))),
            "ReadingOrder" => Ok(OmValue::Number(f64::from(XL_READING_ORDER_CONTEXT))),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_data_labels_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "DataLabel.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_points(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        series_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("Points", member) {
            self.focus_member_supported("Points", member, false)?;
        }

        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Points.Count does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(chart_series_point_count(self.series_model(
                    workbook,
                    chart_id,
                    series_index,
                )?) as f64))
            }
            "Item" => self.dispatch_invoke_points(
                workbook,
                chart_id,
                series_index,
                chart_object_parent,
                member,
                args,
            ),
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Points.Creator does not accept arguments",
                    ));
                }
                self.series_model(workbook, chart_id, series_index)?;
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Points.Application does not accept arguments",
                    ));
                }
                self.series_model(workbook, chart_id, series_index)?;
                Ok(OmValue::Object(self.root_application()))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Points.Parent does not accept arguments",
                    ));
                }
                self.series_model(workbook, chart_id, series_index)?;
                Ok(OmValue::Object(
                    self.register_series_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        series_index,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Points.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_points(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        series_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("Points", member) {
            self.focus_member_supported("Points", member, false)?;
        }

        match member {
            "Item" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "Points.Item expects a single 1-based index",
                    ));
                };
                let index = coerce_u32_arg(index, "Points.Item index")? as usize;
                let point_count = chart_series_point_count(self.series_model(
                    workbook,
                    chart_id,
                    series_index,
                )?);
                if index == 0 || index > point_count {
                    return Err(OmError::invalid_argument(
                        "Points.Item index is out of bounds",
                    ));
                }
                Ok(OmValue::Object(
                    self.register_point_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        series_index,
                        index - 1,
                        chart_object_parent,
                    ),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Points.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_get_point(
        &mut self,
        workbook: WorkbookHandle,
        chart_id: ChartId,
        series_index: usize,
        point_index: usize,
        chart_object_parent: Option<ChartObjectsParent>,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if self.focus_member_declared("Point", member) {
            self.focus_member_supported("Point", member, false)?;
        }

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Point.{member} does not accept arguments"
            )));
        }
        self.validate_point_index(workbook, chart_id, series_index, point_index)?;

        match member {
            "Name" => Ok(OmValue::Text(format!("Point {}", point_index + 1))),
            "Format" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::ChartFormat {
                    workbook,
                    parent: ChartFormatParent::Point {
                        chart_id,
                        series_index,
                        point_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Border" => Ok(OmValue::Object(self.register_object(
                RuntimeObjectKind::Border {
                    workbook,
                    parent: BorderParent::Point {
                        chart_id,
                        series_index,
                        point_index,
                        chart_object_parent,
                    },
                },
            ))),
            "Index" => Ok(OmValue::Number((point_index + 1) as f64)),
            "Explosion" => {
                let point_index = u32::try_from(point_index).map_err(|_| {
                    OmError::invalid_argument("Point.Explosion index is out of bounds")
                })?;
                let series = self.series_model(workbook, chart_id, series_index)?;
                Ok(OmValue::Number(f64::from(
                    series
                        .points
                        .get(&point_index)
                        .and_then(|point| point.explosion)
                        .unwrap_or(0),
                )))
            }
            "HasDataLabel" => {
                self.validate_point_index(workbook, chart_id, series_index, point_index)?;
                Ok(OmValue::Bool(chart_data_labels_visible(
                    chart_point_effective_data_labels(
                        self.chart_model(workbook, chart_id)?,
                        series_index,
                        point_index,
                    ),
                )))
            }
            "DataLabel" => {
                self.validate_data_label_index(workbook, chart_id, series_index, point_index)?;
                Ok(OmValue::Object(
                    self.register_data_label_handle_with_chart_object_parent_origin(
                        workbook,
                        chart_id,
                        series_index,
                        point_index,
                        chart_object_parent,
                    ),
                ))
            }
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => Ok(OmValue::Object(
                self.register_points_handle_with_chart_object_parent_origin(
                    workbook,
                    chart_id,
                    series_index,
                    chart_object_parent,
                ),
            )),
            _ => Err(OmError::unsupported(format!(
                "Point.{member} is not implemented"
            ))),
        }
    }
}
