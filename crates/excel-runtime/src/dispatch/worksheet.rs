use super::super::{
    EXCEL_MAX_COLUMN_INDEX, EXCEL_MAX_ROW_INDEX, ExcelRuntime, RangeProjection, RuntimeNamesScope,
    RuntimeObjectKind, RuntimeSheetCollectionKind, WORKBOOK_PART_NAME, WorkbookCalculationState,
    XL_CREATOR_CODE, XL_SHEET_TYPE_CHART, XL_SHEET_TYPE_DIALOG_SHEET,
    XL_SHEET_TYPE_EXCEL4_MACRO_SHEET, XL_SHEET_TYPE_WORKSHEET, coerce_evaluate_expression_arg,
    coerce_optional_bool_arg, coerce_positive_index, coerce_u32_arg, om_value_is_omitted,
    parse_cells_args, parse_rect_a1, reorder_workbook_sheet_entries,
    sheet_visibility_to_excel_value, validate_check_spelling_args,
    validate_export_as_fixed_format_args, validate_optional_integer_arg,
    validate_optional_text_arg, validate_print_out_args, validate_print_preview_args,
};
use office_common::{
    OmError, OmErrorCode, OmResult, OmValue, Rect, SheetId, SheetKind, WorkbookHandle,
};
use std::collections::BTreeSet;

impl ExcelRuntime {
    pub(crate) fn dispatch_get_sheet_collection(
        &mut self,
        workbook: WorkbookHandle,
        collection_kind: RuntimeSheetCollectionKind,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        let collection_name = collection_kind.member_name();
        self.focus_member_supported(collection_kind.focus_surface_name(), member, false)?;
        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "{collection_name}.Count does not accept arguments"
                    )));
                }
                Ok(OmValue::Number(
                    self.runtime_workbook(workbook)?
                        .loaded
                        .state
                        .worksheets
                        .iter()
                        .filter(|worksheet| collection_kind.includes(worksheet.kind))
                        .count() as f64,
                ))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "{collection_name}.Parent does not accept arguments"
                    )));
                }
                Ok(OmValue::Object(workbook.0))
            }
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "{collection_name}.Application does not accept arguments"
                    )));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "{collection_name}.Creator does not accept arguments"
                    )));
                }
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "Visible" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "{collection_name}.Visible does not accept arguments"
                    )));
                }
                let visibilities = self
                    .runtime_workbook(workbook)?
                    .loaded
                    .state
                    .worksheets
                    .iter()
                    .filter(|worksheet| collection_kind.includes(worksheet.kind))
                    .map(|worksheet| worksheet.visibility)
                    .collect::<Vec<_>>();
                let Some(first_visibility) = visibilities.first().copied() else {
                    return Ok(OmValue::Empty);
                };
                if visibilities
                    .iter()
                    .all(|visibility| *visibility == first_visibility)
                {
                    Ok(OmValue::Number(f64::from(sheet_visibility_to_excel_value(
                        first_visibility,
                    ))))
                } else {
                    Ok(OmValue::Null)
                }
            }
            "Item" => self.resolve_sheet_collection_item(workbook, collection_kind, args),
            _ => Err(OmError::unsupported(format!(
                "{collection_name}.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_worksheet(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Worksheet", member, false)?;

        match member {
            "Name" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Name does not accept arguments",
                    ));
                }
                let worksheet = self.worksheet_model(workbook, sheet_id)?;
                Ok(OmValue::Text(worksheet.name.clone()))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Parent does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(workbook.0))
            }
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Application does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Names" => {
                let handle =
                    self.register_names_handle(workbook, RuntimeNamesScope::Worksheet(sheet_id));
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "ChartObjects" => {
                let sheet_kind = self.worksheet_model(workbook, sheet_id)?.kind;
                if !matches!(
                    sheet_kind,
                    SheetKind::Worksheet | SheetKind::ChartSheet | SheetKind::DialogSheet
                ) {
                    return Err(OmError::unsupported(
                        "Worksheet.ChartObjects is only available on worksheets, chart sheets, and dialog sheets",
                    ));
                }
                let handle = self.register_chart_objects_handle(workbook, sheet_id);
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "Index" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Index does not accept arguments",
                    ));
                }
                let index = self
                    .runtime_workbook(workbook)?
                    .loaded
                    .state
                    .worksheets
                    .iter()
                    .position(|worksheet| worksheet.id == sheet_id)
                    .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "unknown worksheet"))?;
                Ok(OmValue::Number((index + 1) as f64))
            }
            "Next" | "Previous" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "Worksheet.{member} does not accept arguments"
                    )));
                }
                let adjacent_sheet_id = {
                    let worksheets = &self.runtime_workbook(workbook)?.loaded.state.worksheets;
                    let index = worksheets
                        .iter()
                        .position(|worksheet| worksheet.id == sheet_id)
                        .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "unknown worksheet"))?;
                    if member == "Next" {
                        worksheets.get(index + 1).map(|worksheet| worksheet.id)
                    } else if index == 0 {
                        None
                    } else {
                        worksheets.get(index - 1).map(|worksheet| worksheet.id)
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
                        "Worksheet.Visible does not accept arguments",
                    ));
                }
                let worksheet = self.worksheet_model(workbook, sheet_id)?;
                Ok(OmValue::Number(f64::from(sheet_visibility_to_excel_value(
                    worksheet.visibility,
                ))))
            }
            "Type" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Type does not accept arguments",
                    ));
                }
                let sheet_type = match self.worksheet_model(workbook, sheet_id)?.kind {
                    SheetKind::Worksheet => XL_SHEET_TYPE_WORKSHEET,
                    SheetKind::ChartSheet => XL_SHEET_TYPE_CHART,
                    SheetKind::MacroSheet => XL_SHEET_TYPE_EXCEL4_MACRO_SHEET,
                    SheetKind::DialogSheet => XL_SHEET_TYPE_DIALOG_SHEET,
                };
                Ok(OmValue::Number(f64::from(sheet_type)))
            }
            "UsedRange" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.UsedRange does not accept arguments",
                    ));
                }
                self.ensure_grid_worksheet(workbook, sheet_id, "Worksheet.UsedRange")?;
                let rect = self.used_range_rect(workbook, sheet_id)?;
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, rect).0,
                ))
            }
            "Cells" => {
                self.ensure_grid_worksheet(workbook, sheet_id, "Worksheet.Cells")?;
                let rect = match args {
                    [] | [OmValue::Missing] | [OmValue::Empty] | [OmValue::Null] => Rect {
                        row_first: 1,
                        row_last: EXCEL_MAX_ROW_INDEX,
                        col_first: 1,
                        col_last: EXCEL_MAX_COLUMN_INDEX,
                    },
                    _ => {
                        let (row, col) = parse_cells_args(args)?;
                        Rect::single_cell(row, col)
                    }
                };
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, rect).0,
                ))
            }
            "Rows" => {
                self.ensure_grid_worksheet(workbook, sheet_id, "Worksheet.Rows")?;
                let rect = match args {
                    [] | [OmValue::Missing] | [OmValue::Empty] | [OmValue::Null] => Rect {
                        row_first: 1,
                        row_last: EXCEL_MAX_ROW_INDEX,
                        col_first: 1,
                        col_last: EXCEL_MAX_COLUMN_INDEX,
                    },
                    [OmValue::Text(reference)] => {
                        let reference = reference.trim().replace('$', "");
                        let parts: Vec<_> = reference.split(':').collect();
                        let parse_row = |part: &str| -> OmResult<u32> {
                            if part.is_empty() || !part.chars().all(|ch| ch.is_ascii_digit()) {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Rows text selector must be a row number or range like \"2:3\"",
                                ));
                            }
                            let index = part.parse::<u32>().map_err(|_| {
                                OmError::invalid_argument(
                                    "Worksheet.Rows text selector is not a valid row number",
                                )
                            })?;
                            if index == 0 || index > EXCEL_MAX_ROW_INDEX {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Rows text selector is out of bounds",
                                ));
                            }
                            Ok(index)
                        };
                        let (row_first, row_last) = match parts.as_slice() {
                            [single] => {
                                let index = parse_row(single)?;
                                (index, index)
                            }
                            [start, end] => {
                                let start = parse_row(start)?;
                                let end = parse_row(end)?;
                                (start.min(end), start.max(end))
                            }
                            _ => {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Rows text selector must be a row number or range like \"2:3\"",
                                ));
                            }
                        };
                        Rect {
                            row_first,
                            row_last,
                            col_first: 1,
                            col_last: EXCEL_MAX_COLUMN_INDEX,
                        }
                    }
                    [index] => {
                        let index = coerce_u32_arg(index, "Worksheet.Rows index")?;
                        if index == 0 || index > EXCEL_MAX_ROW_INDEX {
                            return Err(OmError::invalid_argument(
                                "Worksheet.Rows index is out of bounds",
                            ));
                        }
                        Rect {
                            row_first: index,
                            row_last: index,
                            col_first: 1,
                            col_last: EXCEL_MAX_COLUMN_INDEX,
                        }
                    }
                    _ => {
                        return Err(OmError::invalid_argument(
                            "Worksheet.Rows expects an optional row index or text range",
                        ));
                    }
                };
                Ok(OmValue::Object(
                    self.register_projected_range_handle(
                        workbook,
                        sheet_id,
                        rect,
                        RangeProjection::Rows,
                    )
                    .0,
                ))
            }
            "Columns" => {
                self.ensure_grid_worksheet(workbook, sheet_id, "Worksheet.Columns")?;
                let rect = match args {
                    [] | [OmValue::Missing] | [OmValue::Empty] | [OmValue::Null] => Rect {
                        row_first: 1,
                        row_last: EXCEL_MAX_ROW_INDEX,
                        col_first: 1,
                        col_last: EXCEL_MAX_COLUMN_INDEX,
                    },
                    [OmValue::Number(number)] => {
                        let index = coerce_positive_index(*number, "Worksheet.Columns index")?;
                        if index > EXCEL_MAX_COLUMN_INDEX {
                            return Err(OmError::invalid_argument(
                                "Worksheet.Columns index is out of bounds",
                            ));
                        }
                        Rect {
                            row_first: 1,
                            row_last: EXCEL_MAX_ROW_INDEX,
                            col_first: index,
                            col_last: index,
                        }
                    }
                    [OmValue::Text(reference)] => {
                        let reference = reference.trim().replace('$', "").to_ascii_uppercase();
                        let parts: Vec<_> = reference.split(':').collect();
                        let parse_column = |part: &str| -> OmResult<u32> {
                            if part.is_empty() || !part.chars().all(|ch| ch.is_ascii_alphabetic()) {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Columns text selector must be a column label or range like \"B:C\"",
                                ));
                            }

                            let mut index = 0u32;
                            for ch in part.bytes() {
                                index = index
                                    .checked_mul(26)
                                    .and_then(|value| value.checked_add((ch - b'A' + 1) as u32))
                                    .ok_or_else(|| {
                                        OmError::invalid_argument(
                                            "Worksheet.Columns text selector overflows column bounds",
                                        )
                                    })?;
                            }
                            if index > EXCEL_MAX_COLUMN_INDEX {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Columns text selector is out of bounds",
                                ));
                            }
                            Ok(index)
                        };
                        let (col_first, col_last) = match parts.as_slice() {
                            [single] => {
                                let index = parse_column(single)?;
                                (index, index)
                            }
                            [start, end] => {
                                let start = parse_column(start)?;
                                let end = parse_column(end)?;
                                (start.min(end), start.max(end))
                            }
                            _ => {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Columns text selector must be a column label or range like \"B:C\"",
                                ));
                            }
                        };
                        Rect {
                            row_first: 1,
                            row_last: EXCEL_MAX_ROW_INDEX,
                            col_first,
                            col_last,
                        }
                    }
                    [_] => {
                        return Err(OmError::type_mismatch(
                            "Worksheet.Columns expects a numeric index or column label string",
                        ));
                    }
                    _ => {
                        return Err(OmError::invalid_argument(
                            "Worksheet.Columns expects an optional column index or text range",
                        ));
                    }
                };
                Ok(OmValue::Object(
                    self.register_projected_range_handle(
                        workbook,
                        sheet_id,
                        rect,
                        RangeProjection::Columns,
                    )
                    .0,
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Worksheet.{member} is not implemented as a property"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_sheet_collection(
        &mut self,
        workbook: WorkbookHandle,
        collection_kind: RuntimeSheetCollectionKind,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        let collection_name = collection_kind.member_name();
        self.focus_member_supported(collection_kind.focus_surface_name(), member, false)?;
        match member {
            "Add" if collection_kind != RuntimeSheetCollectionKind::Charts => {
                let added_sheet = self.add_worksheet(workbook, args)?;
                let RuntimeObjectKind::Worksheet { sheet_id, .. } =
                    self.runtime_object(added_sheet.0)?
                else {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        format!("{collection_name}.Add did not create a sheet"),
                    ));
                };
                Ok(OmValue::Object(
                    self.register_sheet_object_handle(workbook, sheet_id)?,
                ))
            }
            "Add" => {
                if args.len() > 3 {
                    return Err(OmError::invalid_argument(
                        "Charts.Add accepts at most Before, After, and Count arguments",
                    ));
                }
                let add_args = [
                    args.first().cloned().unwrap_or(OmValue::Missing),
                    args.get(1).cloned().unwrap_or(OmValue::Missing),
                    args.get(2).cloned().unwrap_or(OmValue::Missing),
                    OmValue::Number(f64::from(XL_SHEET_TYPE_CHART)),
                ];
                let added_sheet = self.add_worksheet(workbook, &add_args)?;
                let RuntimeObjectKind::Worksheet { sheet_id, .. } =
                    self.runtime_object(added_sheet.0)?
                else {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        "Charts.Add did not create a chart sheet",
                    ));
                };
                let chart_id = self
                    .runtime_workbook(workbook)?
                    .loaded
                    .state
                    .chart_sheets
                    .get(&sheet_id)
                    .map(|binding| binding.chart_id)
                    .ok_or_else(|| {
                        OmError::new(
                            OmErrorCode::InvalidState,
                            "Charts.Add created a sheet without a chart binding",
                        )
                    })?;
                Ok(OmValue::Object(
                    self.register_chart_handle(workbook, chart_id),
                ))
            }
            "Delete" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "{collection_name}.Delete does not accept arguments"
                    )));
                }
                let sheet_ids = self.sheet_collection_ids_in_order(workbook, collection_kind)?;
                if sheet_ids.is_empty() {
                    return Ok(OmValue::Bool(true));
                }
                Self::ensure_pivot_sheet_lifecycle_supported(
                    self.runtime_workbook(workbook)?,
                    &format!("{collection_name}.Delete"),
                )?;
                self.ensure_sheet_block_can_be_deleted_from_workbook(
                    workbook,
                    sheet_ids.as_slice(),
                )?;
                if self.display_alerts {
                    return Ok(OmValue::Bool(false));
                }
                for sheet_id in sheet_ids {
                    self.delete_worksheet(workbook, sheet_id, false)?;
                }
                Ok(OmValue::Bool(true))
            }
            "Copy" => {
                if args.len() > 2 {
                    return Err(OmError::invalid_argument(format!(
                        "{collection_name}.Copy accepts at most Before and After arguments"
                    )));
                }
                let sheet_ids = self.sheet_collection_ids_in_order(workbook, collection_kind)?;
                if sheet_ids.is_empty() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        format!("{collection_name} collection is empty"),
                    ));
                }
                let operation = format!("{collection_name}.Copy");
                let placement_target = self.worksheet_placement_target(
                    workbook,
                    args.first(),
                    args.get(1),
                    operation.as_str(),
                )?;
                Self::ensure_pivot_sheet_lifecycle_supported(
                    self.runtime_workbook(workbook)?,
                    operation.as_str(),
                )?;
                if let Some((target_workbook, _)) = placement_target
                    && target_workbook != workbook
                {
                    Self::ensure_pivot_sheet_lifecycle_supported(
                        self.runtime_workbook(target_workbook)?,
                        operation.as_str(),
                    )?;
                }
                if let Some((target_workbook, base_insertion_index)) = placement_target {
                    let copied_sheet_ids = self.copy_sheet_block_to_workbook(
                        workbook,
                        sheet_ids.as_slice(),
                        target_workbook,
                        base_insertion_index,
                        operation.as_str(),
                    )?;
                    if let Some(active_sheet_id) = copied_sheet_ids.first().copied() {
                        self.set_selection(
                            target_workbook,
                            active_sheet_id,
                            Rect::single_cell(1, 1),
                        );
                    }
                    return Ok(OmValue::Empty);
                }

                self.create_workbook_from_sheet_block(
                    workbook,
                    sheet_ids.as_slice(),
                    operation.as_str(),
                )?;
                Ok(OmValue::Empty)
            }
            "Move" => {
                if args.len() > 2 {
                    return Err(OmError::invalid_argument(format!(
                        "{collection_name}.Move accepts at most Before and After arguments"
                    )));
                }
                let sheet_ids = self.sheet_collection_ids_in_order(workbook, collection_kind)?;
                if sheet_ids.is_empty() {
                    return Err(OmError::new(
                        OmErrorCode::NotFound,
                        format!("{collection_name} collection is empty"),
                    ));
                }
                let placement_target = self.worksheet_placement_target(
                    workbook,
                    args.first(),
                    args.get(1),
                    &format!("{collection_name}.Move"),
                )?;
                let (source_read_only, source_sheet_count) = {
                    let runtime = self.runtime_workbook(workbook)?;
                    (runtime.read_only, runtime.loaded.state.worksheets.len())
                };
                if source_read_only {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        "cannot modify a read-only workbook",
                    ));
                }

                if placement_target
                    .as_ref()
                    .is_none_or(|(target_workbook, _)| *target_workbook != workbook)
                {
                    let operation = format!("{collection_name}.Move");
                    Self::ensure_pivot_sheet_lifecycle_supported(
                        self.runtime_workbook(workbook)?,
                        operation.as_str(),
                    )?;
                    if let Some((target_workbook, _)) = placement_target {
                        Self::ensure_pivot_sheet_lifecycle_supported(
                            self.runtime_workbook(target_workbook)?,
                            operation.as_str(),
                        )?;
                    }
                }

                if let Some((target_workbook, base_insertion_index)) = placement_target {
                    if target_workbook == workbook {
                        let moving_sheet_ids = sheet_ids.iter().copied().collect::<BTreeSet<_>>();
                        let active_sheet_id = sheet_ids[0];
                        {
                            let runtime = self.runtime_workbook_mut(workbook)?;
                            let target_index =
                                base_insertion_index.min(runtime.loaded.state.worksheets.len());
                            let removed_before = runtime
                                .loaded
                                .state
                                .worksheets
                                .iter()
                                .take(target_index)
                                .filter(|worksheet| moving_sheet_ids.contains(&worksheet.id))
                                .count();
                            let mut moving_sheets = Vec::with_capacity(sheet_ids.len());
                            let mut remaining_sheets = Vec::with_capacity(
                                runtime
                                    .loaded
                                    .state
                                    .worksheets
                                    .len()
                                    .saturating_sub(sheet_ids.len()),
                            );
                            for worksheet in runtime.loaded.state.worksheets.drain(..) {
                                if moving_sheet_ids.contains(&worksheet.id) {
                                    moving_sheets.push(worksheet);
                                } else {
                                    remaining_sheets.push(worksheet);
                                }
                            }
                            if moving_sheets.is_empty() {
                                return Err(OmError::new(
                                    OmErrorCode::InvalidState,
                                    format!("{collection_name}.Move found no sheets to move"),
                                ));
                            }
                            let adjusted_index = target_index
                                .saturating_sub(removed_before)
                                .min(remaining_sheets.len());
                            for (offset, worksheet) in moving_sheets.into_iter().enumerate() {
                                remaining_sheets.insert(adjusted_index + offset, worksheet);
                            }
                            runtime.loaded.state.worksheets = remaining_sheets;
                            let workbook_xml = runtime
                                .loaded
                                .package
                                .part(WORKBOOK_PART_NAME)
                                .ok_or_else(|| {
                                    OmError::new(
                                        OmErrorCode::Parse,
                                        format!("workbook package is missing {WORKBOOK_PART_NAME}"),
                                    )
                                })?
                                .bytes
                                .clone();
                            runtime.loaded.package.replace_part_bytes(
                                WORKBOOK_PART_NAME,
                                reorder_workbook_sheet_entries(
                                    workbook_xml.as_slice(),
                                    &runtime.loaded.state.worksheets,
                                )?,
                            )?;
                            runtime.prompt_dirty = true;
                        }
                        self.find_state = None;
                        self.cut_copy_mode = None;
                        self.clipboard = None;
                        self.set_selection(workbook, active_sheet_id, Rect::single_cell(1, 1));
                        return Ok(OmValue::Empty);
                    }

                    let copied_sheet_ids = self.copy_sheet_block_to_workbook(
                        workbook,
                        sheet_ids.as_slice(),
                        target_workbook,
                        base_insertion_index,
                        &format!("{collection_name}.Move"),
                    )?;
                    if source_sheet_count > sheet_ids.len() {
                        self.ensure_sheet_block_can_be_deleted_from_workbook(
                            workbook,
                            sheet_ids.as_slice(),
                        )?;
                        for sheet_id in sheet_ids {
                            self.delete_worksheet(workbook, sheet_id, false)?;
                        }
                    } else {
                        self.close_workbook(workbook)?;
                    }
                    if let Some(active_sheet_id) = copied_sheet_ids.first().copied() {
                        self.set_selection(
                            target_workbook,
                            active_sheet_id,
                            Rect::single_cell(1, 1),
                        );
                    }
                    return Ok(OmValue::Empty);
                }

                let (moved_workbook, moved_sheet_ids) = self.create_workbook_from_sheet_block(
                    workbook,
                    sheet_ids.as_slice(),
                    &format!("{collection_name}.Move"),
                )?;
                if source_sheet_count > sheet_ids.len() {
                    self.ensure_sheet_block_can_be_deleted_from_workbook(
                        workbook,
                        sheet_ids.as_slice(),
                    )?;
                    for sheet_id in sheet_ids {
                        self.delete_worksheet(workbook, sheet_id, false)?;
                    }
                } else {
                    self.close_workbook(workbook)?;
                }
                if let Some(active_sheet_id) = moved_sheet_ids.first().copied() {
                    self.set_selection(moved_workbook, active_sheet_id, Rect::single_cell(1, 1));
                }
                Ok(OmValue::Empty)
            }
            "PrintPreview" | "PrintOut" => {
                match member {
                    "PrintPreview" => validate_print_preview_args(args, collection_name)?,
                    "PrintOut" => validate_print_out_args(args, collection_name)?,
                    _ => unreachable!("sheet collection print method branch"),
                }
                if collection_kind == RuntimeSheetCollectionKind::Charts {
                    let chart_ids = {
                        let runtime = self.runtime_workbook(workbook)?;
                        runtime
                            .loaded
                            .state
                            .worksheets
                            .iter()
                            .filter(|worksheet| worksheet.kind == SheetKind::ChartSheet)
                            .map(|worksheet| {
                                runtime
                                    .loaded
                                    .state
                                    .chart_sheets
                                    .get(&worksheet.id)
                                    .map(|binding| binding.chart_id)
                                    .ok_or_else(|| {
                                        OmError::new(
                                            OmErrorCode::InvalidState,
                                            "chart sheet is missing a chart binding",
                                        )
                                    })
                            })
                            .collect::<OmResult<Vec<_>>>()?
                    };
                    for chart_id in chart_ids {
                        self.chart_model(workbook, chart_id)?;
                    }
                } else {
                    self.runtime_workbook(workbook)?;
                }
                Ok(OmValue::Empty)
            }
            "Select" => {
                if args.len() > 1 {
                    return Err(OmError::invalid_argument(format!(
                        "{collection_name}.Select accepts at most a Replace argument"
                    )));
                }
                if let Some(value) = args.first()
                    && !om_value_is_omitted(value)
                {
                    coerce_optional_bool_arg(
                        value,
                        true,
                        &format!("{collection_name}.Select Replace"),
                    )?;
                }
                let sheet_id = self
                    .sheet_collection_ids_in_order(workbook, collection_kind)?
                    .first()
                    .copied()
                    .ok_or_else(|| {
                        OmError::new(
                            OmErrorCode::NotFound,
                            format!("{collection_name} collection is empty"),
                        )
                    })?;
                self.ensure_worksheet_visible(
                    workbook,
                    sheet_id,
                    &format!("{collection_name}.Select"),
                )?;
                self.set_selection(workbook, sheet_id, Rect::single_cell(1, 1));
                Ok(OmValue::Empty)
            }
            "Item" => self.resolve_sheet_collection_item(workbook, collection_kind, args),
            _ => Err(OmError::unsupported(format!(
                "{collection_name}.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_worksheet(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Worksheet", member, false)?;
        match member {
            "Range" => {
                self.ensure_grid_worksheet(workbook, sheet_id, "Worksheet.Range")?;
                if let [OmValue::Text(reference)] = args {
                    let range = self.resolve_worksheet_range_text(workbook, sheet_id, reference)?;
                    self.remember_range_selection(workbook, &range, "Worksheet.Range")?;
                    return Ok(OmValue::Object(
                        self.register_range_set_handle(workbook, range).0,
                    ));
                }
                if let [OmValue::Object(handle)] = args {
                    match self.runtime_object(*handle)? {
                        RuntimeObjectKind::Range {
                            workbook: range_workbook,
                            range: object_range,
                            ..
                        } => {
                            let (range_sheet_id, _) =
                                Self::range_set_single_sheet_rects(&object_range)?;
                            if range_workbook != workbook || range_sheet_id != sheet_id {
                                return Err(OmError::invalid_argument(
                                    "Worksheet.Range object argument must belong to the same worksheet",
                                ));
                            }
                            self.remember_range_selection(
                                workbook,
                                &object_range,
                                "Worksheet.Range",
                            )?;
                            return Ok(OmValue::Object(
                                self.register_range_set_handle(workbook, object_range).0,
                            ));
                        }
                        _ => {
                            return Err(OmError::type_mismatch(
                                "Worksheet.Range expects A1 references or Range objects",
                            ));
                        }
                    }
                }
                let rect = match args {
                    [start, end] => {
                        let start = match start {
                            OmValue::Text(a1) => parse_rect_a1(a1)?,
                            OmValue::Object(handle) => match self.runtime_object(*handle)? {
                                RuntimeObjectKind::Range {
                                    workbook: range_workbook,
                                    range: object_range,
                                    ..
                                } => {
                                    if object_range.areas().len() != 1 {
                                        return Err(OmError::unsupported(
                                            "Worksheet.Range endpoint arguments require single-area Range objects",
                                        ));
                                    }
                                    let (range_sheet_id, rect) =
                                        Self::range_set_single_area(&object_range)?;
                                    if range_workbook != workbook || range_sheet_id != sheet_id {
                                        return Err(OmError::invalid_argument(
                                            "Worksheet.Range object arguments must belong to the same worksheet",
                                        ));
                                    }
                                    rect
                                }
                                _ => {
                                    return Err(OmError::type_mismatch(
                                        "Worksheet.Range expects A1 references or Range objects",
                                    ));
                                }
                            },
                            _ => {
                                return Err(OmError::type_mismatch(
                                    "Worksheet.Range expects A1 references or Range objects",
                                ));
                            }
                        };
                        let end = match end {
                            OmValue::Text(a1) => parse_rect_a1(a1)?,
                            OmValue::Object(handle) => match self.runtime_object(*handle)? {
                                RuntimeObjectKind::Range {
                                    workbook: range_workbook,
                                    range: object_range,
                                    ..
                                } => {
                                    if object_range.areas().len() != 1 {
                                        return Err(OmError::unsupported(
                                            "Worksheet.Range endpoint arguments require single-area Range objects",
                                        ));
                                    }
                                    let (range_sheet_id, rect) =
                                        Self::range_set_single_area(&object_range)?;
                                    if range_workbook != workbook || range_sheet_id != sheet_id {
                                        return Err(OmError::invalid_argument(
                                            "Worksheet.Range object arguments must belong to the same worksheet",
                                        ));
                                    }
                                    rect
                                }
                                _ => {
                                    return Err(OmError::type_mismatch(
                                        "Worksheet.Range expects A1 references or Range objects",
                                    ));
                                }
                            },
                            _ => {
                                return Err(OmError::type_mismatch(
                                    "Worksheet.Range expects A1 references or Range objects",
                                ));
                            }
                        };
                        Rect {
                            row_first: start.row_first.min(end.row_first),
                            row_last: start.row_last.max(end.row_last),
                            col_first: start.col_first.min(end.col_first),
                            col_last: start.col_last.max(end.col_last),
                        }
                    }
                    _ => {
                        return Err(OmError::invalid_argument(
                            "Worksheet.Range expects one A1 reference or Range object, or two A1/range endpoints",
                        ));
                    }
                };
                self.remember_selection(workbook, sheet_id, rect);
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, rect).0,
                ))
            }
            "Activate" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Activate does not accept arguments",
                    ));
                }
                self.ensure_worksheet_visible(workbook, sheet_id, "Worksheet.Activate")?;
                let rect = self
                    .selection
                    .filter(|selection| {
                        selection.workbook == workbook && selection.sheet_id == sheet_id
                    })
                    .map(|selection| selection.rect)
                    .unwrap_or(Rect::single_cell(1, 1));
                self.set_selection(workbook, sheet_id, rect);
                Ok(OmValue::Empty)
            }
            "Select" => {
                match args {
                    []
                    | [OmValue::Missing | OmValue::Empty | OmValue::Null]
                    | [OmValue::Bool(_)] => {}
                    [_] => {
                        return Err(OmError::type_mismatch(
                            "Worksheet.Select Replace expects a boolean when provided",
                        ));
                    }
                    _ => {
                        return Err(OmError::invalid_argument(
                            "Worksheet.Select accepts at most one Replace argument",
                        ));
                    }
                }
                self.ensure_worksheet_visible(workbook, sheet_id, "Worksheet.Select")?;
                let rect = self
                    .selection
                    .filter(|selection| {
                        selection.workbook == workbook && selection.sheet_id == sheet_id
                    })
                    .map(|selection| selection.rect)
                    .unwrap_or(Rect::single_cell(1, 1));
                self.set_selection(workbook, sheet_id, rect);
                Ok(OmValue::Empty)
            }
            "PrintPreview" => {
                validate_print_preview_args(args, "Worksheet")?;
                self.worksheet_model(workbook, sheet_id)?;
                Ok(OmValue::Empty)
            }
            "PrintOut" => {
                validate_print_out_args(args, "Worksheet")?;
                self.worksheet_model(workbook, sheet_id)?;
                Ok(OmValue::Empty)
            }
            "CheckSpelling" => {
                validate_check_spelling_args(args, "Worksheet")?;
                self.worksheet_model(workbook, sheet_id)?;
                Ok(OmValue::Empty)
            }
            "ExportAsFixedFormat" => {
                validate_export_as_fixed_format_args(args, "Worksheet")?;
                self.worksheet_model(workbook, sheet_id)?;
                Ok(OmValue::Empty)
            }
            "Paste" => {
                self.ensure_grid_worksheet(workbook, sheet_id, "Worksheet.Paste")?;
                if args.len() > 2 {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Paste accepts at most Destination and Link arguments",
                    ));
                }
                if let Some(value) = args.get(1) {
                    let link = coerce_optional_bool_arg(value, false, "Worksheet.Paste Link")?;
                    if link {
                        return Err(OmError::unsupported(
                            "Worksheet.Paste Link is not supported",
                        ));
                    }
                }
                let destination = match args.first() {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => None,
                    Some(OmValue::Object(handle)) => {
                        let RuntimeObjectKind::Range {
                            workbook: destination_workbook,
                            range: destination_range,
                            ..
                        } = self.runtime_object(*handle)?
                        else {
                            return Err(OmError::type_mismatch(
                                "Worksheet.Paste Destination expects a Range object",
                            ));
                        };
                        if destination_range.areas().len() != 1 {
                            return Err(OmError::unsupported(
                                "Worksheet.Paste Destination requires a single-area range for cell materialization",
                            ));
                        }
                        let (destination_sheet_id, _) =
                            Self::range_set_single_area(&destination_range)?;
                        if destination_workbook != workbook || destination_sheet_id != sheet_id {
                            return Err(OmError::invalid_argument(
                                "Worksheet.Paste Destination must belong to the same worksheet",
                            ));
                        }
                        Some(*handle)
                    }
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Worksheet.Paste Destination expects a Range object",
                        ));
                    }
                };
                let destination = if let Some(destination) = destination {
                    destination
                } else {
                    if let Some(selection_range) = self.selection_range.as_ref()
                        && selection_range.workbook_id() == self.workbook_model(workbook)?.id
                        && let Ok((selection_sheet_id, _)) =
                            Self::range_set_single_sheet_rects(selection_range)
                        && selection_sheet_id == sheet_id
                        && selection_range.len() != 1
                    {
                        return Err(OmError::unsupported(
                            "Worksheet.Paste without a Destination requires a single-area selection",
                        ));
                    }
                    let rect = self
                        .selection
                        .filter(|selection| {
                            selection.workbook == workbook && selection.sheet_id == sheet_id
                        })
                        .map(|selection| selection.rect)
                        .unwrap_or(Rect::single_cell(1, 1));
                    self.register_range_handle(workbook, sheet_id, rect).0
                };
                self.dispatch_invoke(destination, "PasteSpecial", &[])
            }
            "PasteSpecial" => {
                self.ensure_grid_worksheet(workbook, sheet_id, "Worksheet.PasteSpecial")?;
                if args.len() > 7 {
                    return Err(OmError::invalid_argument(
                        "Worksheet.PasteSpecial accepts at most Format, Link, DisplayAsIcon, IconFileName, IconIndex, IconLabel, and NoHTMLFormatting arguments",
                    ));
                }
                validate_optional_text_arg(args, 0, "Worksheet.PasteSpecial Format")?;
                if matches!(args.first(), Some(OmValue::Text(format)) if !format.is_empty()) {
                    return Err(OmError::unsupported(
                        "Worksheet.PasteSpecial Format is not supported",
                    ));
                }
                if let Some(value) = args.get(1) {
                    let link =
                        coerce_optional_bool_arg(value, false, "Worksheet.PasteSpecial Link")?;
                    if link {
                        return Err(OmError::unsupported(
                            "Worksheet.PasteSpecial Link is not supported",
                        ));
                    }
                }
                if let Some(value) = args.get(2) {
                    let display_as_icon = coerce_optional_bool_arg(
                        value,
                        false,
                        "Worksheet.PasteSpecial DisplayAsIcon",
                    )?;
                    if display_as_icon {
                        return Err(OmError::unsupported(
                            "Worksheet.PasteSpecial DisplayAsIcon is not supported",
                        ));
                    }
                }
                validate_optional_text_arg(args, 3, "Worksheet.PasteSpecial IconFileName")?;
                if matches!(args.get(3), Some(OmValue::Text(icon_file_name)) if !icon_file_name.is_empty())
                {
                    return Err(OmError::unsupported(
                        "Worksheet.PasteSpecial IconFileName is not supported",
                    ));
                }
                validate_optional_integer_arg(args, 4, "Worksheet.PasteSpecial IconIndex")?;
                if args.get(4).is_some_and(|value| !om_value_is_omitted(value)) {
                    return Err(OmError::unsupported(
                        "Worksheet.PasteSpecial IconIndex is not supported",
                    ));
                }
                validate_optional_text_arg(args, 5, "Worksheet.PasteSpecial IconLabel")?;
                if matches!(args.get(5), Some(OmValue::Text(icon_label)) if !icon_label.is_empty())
                {
                    return Err(OmError::unsupported(
                        "Worksheet.PasteSpecial IconLabel is not supported",
                    ));
                }
                if let Some(value) = args.get(6) {
                    coerce_optional_bool_arg(
                        value,
                        false,
                        "Worksheet.PasteSpecial NoHTMLFormatting",
                    )?;
                }
                self.dispatch_invoke_worksheet(workbook, sheet_id, "Paste", &[])
            }
            "Calculate" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Calculate does not accept arguments",
                    ));
                }
                self.ensure_grid_worksheet(workbook, sheet_id, "Worksheet.Calculate")?;
                self.calculate_sheet_formulas(workbook, sheet_id, None)?;
                self.record_calculation_snapshot(
                    workbook,
                    WorkbookCalculationState::PartiallyCalculated,
                )?;
                Ok(OmValue::Empty)
            }
            "Evaluate" => {
                let expression = coerce_evaluate_expression_arg(args, "Worksheet.Evaluate")?;
                self.evaluate_formula_expression(
                    workbook,
                    sheet_id,
                    expression,
                    "Worksheet.Evaluate",
                )
            }
            "Cells" => {
                self.ensure_grid_worksheet(workbook, sheet_id, "Worksheet.Cells")?;
                let (row, col) = parse_cells_args(args)?;
                let rect = Rect::single_cell(row, col);
                self.remember_selection(workbook, sheet_id, rect);
                Ok(OmValue::Object(
                    self.register_range_handle(workbook, sheet_id, rect).0,
                ))
            }
            "Copy" => {
                if args.len() > 2 {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Copy accepts at most Before and After arguments",
                    ));
                }
                let placement_target = self.worksheet_placement_target(
                    workbook,
                    args.first(),
                    args.get(1),
                    "Worksheet.Copy",
                )?;
                Self::ensure_pivot_sheet_lifecycle_supported(
                    self.runtime_workbook(workbook)?,
                    "Worksheet.Copy",
                )?;
                if let Some((target_workbook, _)) = placement_target
                    && target_workbook != workbook
                {
                    Self::ensure_pivot_sheet_lifecycle_supported(
                        self.runtime_workbook(target_workbook)?,
                        "Worksheet.Copy",
                    )?;
                }
                if let Some((target_workbook, insertion_index)) = placement_target {
                    self.copy_basic_worksheet_to_workbook(
                        workbook,
                        sheet_id,
                        target_workbook,
                        insertion_index,
                    )?;
                } else {
                    let rect = self
                        .selection
                        .filter(|selection| {
                            selection.workbook == workbook && selection.sheet_id == sheet_id
                        })
                        .map(|selection| selection.rect)
                        .unwrap_or(Rect::single_cell(1, 1));
                    self.spawn_single_sheet_workbook_from_source(workbook, sheet_id, rect)?;
                }
                Ok(OmValue::Empty)
            }
            "Move" => {
                self.move_worksheet(workbook, sheet_id, args)?;
                Ok(OmValue::Empty)
            }
            "Delete" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Worksheet.Delete does not accept arguments",
                    ));
                }
                Self::ensure_pivot_sheet_lifecycle_supported(
                    self.runtime_workbook(workbook)?,
                    "Worksheet.Delete",
                )?;
                Ok(OmValue::Bool(
                    self.delete_worksheet(workbook, sheet_id, true)?,
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Worksheet.{member} is not implemented as a method"
            ))),
        }
    }
}
