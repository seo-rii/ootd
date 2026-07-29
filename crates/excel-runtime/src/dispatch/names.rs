use super::super::{
    ExcelRuntime, RuntimeNamesScope, XL_CREATOR_CODE, coerce_optional_bool_arg, coerce_u32_arg,
    convert_formula_a1_to_r1c1, convert_formula_r1c1_to_a1, om_value_is_omitted,
    runtime_name_scope,
};
use excel_model::{ChartSourceExpr, DrawingObjectModel, resolve_chart_source_reference_with_names};
use office_common::{
    DefinedNameId, DefinedNameMetadata, FormulaSource, NameScope, NameValidationMode, OmError,
    OmErrorCode, OmResult, OmValue, SheetId, WorkbookHandle,
};
use std::collections::BTreeMap;

impl ExcelRuntime {
    pub(crate) fn dispatch_get_names(
        &mut self,
        workbook: WorkbookHandle,
        scope: RuntimeNamesScope,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Names", member, false)?;

        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Names.Count does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(
                    self.names_in_scope(workbook, scope)?.len() as f64
                ))
            }
            "Item" => self.dispatch_invoke_names(workbook, scope, member, args),
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Names.Application does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Names.Creator does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Names.Parent does not accept arguments",
                    ));
                }
                match scope {
                    RuntimeNamesScope::Workbook => Ok(OmValue::Object(workbook.0)),
                    RuntimeNamesScope::Worksheet(sheet_id) => Ok(OmValue::Object(
                        self.register_worksheet_handle(workbook, sheet_id).0,
                    )),
                }
            }
            _ => Err(OmError::unsupported(format!(
                "Names.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_names(
        &mut self,
        workbook: WorkbookHandle,
        scope: RuntimeNamesScope,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Names", member, false)?;

        match member {
            "Item" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "Names.Item expects a single name or 1-based index",
                    ));
                };
                let name_id = match index {
                    OmValue::Text(name) => {
                        let name_scope = runtime_name_scope(scope);
                        self.runtime_workbook(workbook)?
                            .loaded
                            .state
                            .lookup_name_in_scope(name_scope, name)
                            .map(|defined_name| defined_name.id)
                            .ok_or_else(|| {
                                OmError::new(
                                    OmErrorCode::NotFound,
                                    format!("defined name '{name}' was not found"),
                                )
                            })?
                    }
                    OmValue::Number(_) => {
                        let index = coerce_u32_arg(index, "Names.Item index")? as usize;
                        if index == 0 {
                            return Err(OmError::invalid_argument(
                                "Names.Item index is out of bounds",
                            ));
                        }
                        *self
                            .names_in_scope(workbook, scope)?
                            .get(index - 1)
                            .ok_or_else(|| {
                                OmError::invalid_argument("Names.Item index is out of bounds")
                            })?
                    }
                    _ => {
                        return Err(OmError::type_mismatch(
                            "Names.Item expects a name string or numeric index",
                        ));
                    }
                };
                Ok(OmValue::Object(
                    self.register_name_handle(workbook, name_id),
                ))
            }
            "Add" => {
                if args.len() < 2 || args.len() > 11 {
                    return Err(OmError::invalid_argument(
                        "Names.Add expects Name, RefersTo, and optional Visible, MacroType, ShortcutKey, Category, NameLocal, RefersToLocal, CategoryLocal, RefersToR1C1, and RefersToR1C1Local arguments",
                    ));
                }
                let name_local = match args.get(6) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => None,
                    Some(OmValue::Text(name)) => Some(name.clone()),
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Names.Add NameLocal expects a string when provided",
                        ));
                    }
                };
                let name = match (&args[0], name_local) {
                    (OmValue::Text(name), None) => name.clone(),
                    (OmValue::Missing | OmValue::Empty | OmValue::Null, Some(name)) => name,
                    (OmValue::Text(_), Some(_)) => {
                        return Err(OmError::invalid_argument(
                            "Names.Add accepts only one of Name and NameLocal",
                        ));
                    }
                    _ => {
                        return Err(OmError::type_mismatch(
                            "Names.Add Name expects a string unless NameLocal is provided",
                        ));
                    }
                };
                let refers_to_local = match args.get(7) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => None,
                    Some(OmValue::Text(refers_to)) => {
                        Some(refers_to.trim_start_matches('=').to_string())
                    }
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Names.Add RefersToLocal expects a string when provided",
                        ));
                    }
                };
                let refers_to_a1 = match (&args[1], refers_to_local) {
                    (OmValue::Text(refers_to), None) => {
                        Some(refers_to.trim_start_matches('=').to_string())
                    }
                    (OmValue::Missing | OmValue::Empty | OmValue::Null, Some(refers_to)) => {
                        Some(refers_to)
                    }
                    (OmValue::Text(_), Some(_)) => {
                        return Err(OmError::invalid_argument(
                            "Names.Add accepts only one of RefersTo and RefersToLocal",
                        ));
                    }
                    (OmValue::Missing | OmValue::Empty | OmValue::Null, None) => None,
                    _ => {
                        return Err(OmError::type_mismatch(
                            "Names.Add RefersTo expects a string unless RefersToLocal or RefersToR1C1 is provided",
                        ));
                    }
                };
                let refers_to_r1c1_local = match args.get(10) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => None,
                    Some(OmValue::Text(refers_to)) => {
                        Some(refers_to.trim_start_matches('=').to_string())
                    }
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Names.Add RefersToR1C1Local expects a string when provided",
                        ));
                    }
                };
                let refers_to_r1c1_direct = match args.get(9) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => None,
                    Some(OmValue::Text(refers_to)) => {
                        Some(refers_to.trim_start_matches('=').to_string())
                    }
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Names.Add RefersToR1C1 expects a string when provided",
                        ));
                    }
                };
                let refers_to_r1c1 = match (refers_to_r1c1_direct, refers_to_r1c1_local) {
                    (Some(_), Some(_)) => {
                        return Err(OmError::invalid_argument(
                            "Names.Add accepts only one of RefersToR1C1 and RefersToR1C1Local",
                        ));
                    }
                    (Some(refers_to), None) | (None, Some(refers_to)) => Some(refers_to),
                    (None, None) => None,
                };
                let (refers_to, is_r1c1) = match (refers_to_a1, refers_to_r1c1) {
                    (Some(refers_to), None) => (refers_to, false),
                    (None, Some(refers_to)) => (refers_to, true),
                    (Some(_), Some(_)) => {
                        return Err(OmError::invalid_argument(
                            "Names.Add accepts only one A1 or R1C1 reference argument",
                        ));
                    }
                    (None, None) => {
                        return Err(OmError::type_mismatch(
                            "Names.Add expects RefersTo, RefersToLocal, RefersToR1C1, or RefersToR1C1Local",
                        ));
                    }
                };
                let visible = args
                    .get(2)
                    .map(|value| coerce_optional_bool_arg(value, true, "Names.Add Visible"))
                    .transpose()?
                    .unwrap_or(true);
                let macro_type = match args.get(3) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => 3,
                    Some(value) => coerce_u32_arg(value, "Names.Add MacroType")?,
                };
                if !matches!(macro_type, 1 | 2 | 3) {
                    return Err(OmError::invalid_argument(
                        "Names.Add MacroType supports 1, 2, 3, or omitted",
                    ));
                }
                let shortcut_key = match args.get(4) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => None,
                    Some(OmValue::Text(shortcut_key)) => {
                        let mut chars = shortcut_key.chars();
                        match (chars.next(), chars.next()) {
                            (Some(ch), None) if ch.is_ascii_alphabetic() => {
                                Some(shortcut_key.clone())
                            }
                            _ => {
                                return Err(OmError::invalid_argument(
                                    "Names.Add ShortcutKey expects a single letter",
                                ));
                            }
                        }
                    }
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Names.Add ShortcutKey expects a string when provided",
                        ));
                    }
                };
                if shortcut_key.is_some() && macro_type != 2 {
                    return Err(OmError::invalid_argument(
                        "Names.Add ShortcutKey applies only when MacroType is 2",
                    ));
                }
                let category_arg = match (
                    args.get(5).filter(|value| !om_value_is_omitted(value)),
                    args.get(8).filter(|value| !om_value_is_omitted(value)),
                ) {
                    (None, None) => None,
                    (Some(_), Some(_)) => {
                        return Err(OmError::invalid_argument(
                            "Names.Add accepts only one of Category and CategoryLocal",
                        ));
                    }
                    (Some(value), None) => Some((value, "Names.Add Category")),
                    (None, Some(value)) => Some((value, "Names.Add CategoryLocal")),
                };
                let function_group_id = if let Some((value, label)) = category_arg {
                    if !matches!(macro_type, 1 | 2) {
                        return Err(OmError::invalid_argument(
                            "Names.Add Category applies only when MacroType is 1 or 2",
                        ));
                    }
                    match value {
                        OmValue::Number(_) => Some(coerce_u32_arg(value, label)?),
                        OmValue::Text(category) => {
                            let normalized = category.trim().to_ascii_lowercase();
                            let category_id = match normalized.as_str() {
                                "financial" => 1,
                                "date and time" | "date & time" => 2,
                                "math and trig" | "math & trig" => 3,
                                "statistical" => 4,
                                "lookup and reference" | "lookup & reference" => 5,
                                "database" => 6,
                                "text" => 7,
                                "logical" => 8,
                                "information" => 9,
                                "commands" | "command" => 10,
                                "customizing" | "customization" => 11,
                                "macro control" => 12,
                                "dde / external" | "dde/external" | "dde external" => 13,
                                "user defined" | "user-defined" => 14,
                                "engineering" => 15,
                                "cube" => 16,
                                "" => {
                                    return Err(OmError::invalid_argument(format!(
                                        "{label} expects a category name or number"
                                    )));
                                }
                                _ => {
                                    return Err(OmError::unsupported(format!(
                                        "{label} custom categories are not implemented"
                                    )));
                                }
                            };
                            Some(category_id)
                        }
                        _ => {
                            return Err(OmError::type_mismatch(format!(
                                "{label} expects a category name or number"
                            )));
                        }
                    }
                } else {
                    None
                };
                let name_id = {
                    let runtime = self.runtime_workbook_mut(workbook)?;
                    if runtime.read_only {
                        return Err(OmError::new(
                            OmErrorCode::InvalidState,
                            "cannot modify a read-only workbook",
                        ));
                    }
                    let mut metadata = DefinedNameMetadata::default();
                    metadata.hidden = !visible;
                    metadata.function = macro_type == 1;
                    metadata.vb_procedure = macro_type == 2;
                    metadata.function_group_id = function_group_id;
                    metadata.shortcut_key = shortcut_key;
                    let name_id = runtime.loaded.state.defined_names.add_with_metadata(
                        runtime_name_scope(scope),
                        name,
                        FormulaSource {
                            text: refers_to,
                            is_r1c1,
                        },
                        metadata,
                        NameValidationMode::StrictExcel,
                    )?;
                    let workbook_id = runtime.loaded.state.model().id;
                    let workbook_display_name = runtime.loaded.state.model().display_name.clone();
                    let worksheets = runtime.loaded.state.worksheets().to_vec();
                    let defined_names = runtime.loaded.state.defined_names.clone();
                    let mut chart_source_current_sheets = BTreeMap::new();
                    for (chart_sheet_id, binding) in &runtime.loaded.state.chart_sheets {
                        chart_source_current_sheets.insert(binding.chart_id, *chart_sheet_id);
                    }
                    for drawing in runtime.loaded.state.drawings.values() {
                        for object in &drawing.objects {
                            let DrawingObjectModel::ChartFrame(chart_object) = object else {
                                continue;
                            };
                            chart_source_current_sheets
                                .insert(chart_object.chart_id, drawing.host_sheet_id);
                        }
                    }
                    let refresh_chart_source =
                        |source: &mut Option<ChartSourceExpr>, current_sheet: Option<SheetId>| {
                            let Some(source) = source.as_mut() else {
                                return;
                            };
                            source.resolved = resolve_chart_source_reference_with_names(
                                source.raw.text.as_str(),
                                workbook_id,
                                Some(&workbook_display_name),
                                &worksheets,
                                &defined_names,
                                current_sheet,
                            );
                        };
                    for chart in runtime.loaded.state.charts.values_mut() {
                        let chart_current_sheet =
                            chart_source_current_sheets.get(&chart.id).copied();
                        for series in &mut chart.series {
                            refresh_chart_source(&mut series.name, chart_current_sheet);
                            refresh_chart_source(&mut series.x_values, chart_current_sheet);
                            refresh_chart_source(&mut series.values, chart_current_sheet);
                            refresh_chart_source(&mut series.bubble_size, chart_current_sheet);
                        }
                    }
                    runtime.prompt_dirty = true;
                    self.find_state = None;
                    self.cut_copy_mode = None;
                    self.clipboard = None;
                    name_id
                };
                Ok(OmValue::Object(
                    self.register_name_handle(workbook, name_id),
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Names.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn dispatch_get_name(
        &mut self,
        workbook: WorkbookHandle,
        name_id: DefinedNameId,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Name", member, false)?;

        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Name.{member} does not accept arguments"
            )));
        }

        match member {
            "Name" => Ok(OmValue::Text(
                self.defined_name(workbook, name_id)?.display_name.clone(),
            )),
            "Visible" => Ok(OmValue::Bool(
                !self.defined_name(workbook, name_id)?.metadata.hidden,
            )),
            "RefersTo" | "RefersToLocal" | "RefersToR1C1" | "RefersToR1C1Local" => {
                let refers_to = self.defined_name(workbook, name_id)?.refers_to.clone();
                let reference = refers_to.text.trim_start_matches('=');
                let reference = if matches!(member, "RefersToR1C1" | "RefersToR1C1Local") {
                    if refers_to.is_r1c1 {
                        reference.to_string()
                    } else {
                        convert_formula_a1_to_r1c1(reference, 1, 1)
                    }
                } else if refers_to.is_r1c1 {
                    convert_formula_r1c1_to_a1(reference, 1, 1)
                } else {
                    reference.to_string()
                };
                Ok(OmValue::Text(format!("={reference}")))
            }
            "RefersToRange" => {
                let refers_to = self.defined_name(workbook, name_id)?.refers_to.clone();
                let mut reference = refers_to.text.trim_start_matches('=').to_string();
                if refers_to.is_r1c1 {
                    reference = convert_formula_r1c1_to_a1(reference.as_str(), 1, 1);
                }
                let (target_workbook, range) =
                    self.resolve_application_range_text(reference.as_str())?;
                if target_workbook != workbook {
                    return Err(OmError::unsupported(
                        "Name.RefersToRange cross-workbook references are not supported",
                    ));
                }
                self.ensure_range_set_targets_grid_worksheets(
                    target_workbook,
                    &range,
                    "Name.RefersToRange",
                )?;
                Ok(OmValue::Object(
                    self.register_range_set_handle(workbook, range).0,
                ))
            }
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Parent" => {
                let scope = self.defined_name(workbook, name_id)?.scope;
                match scope {
                    NameScope::Workbook => Ok(OmValue::Object(workbook.0)),
                    NameScope::Worksheet(sheet_id) => Ok(OmValue::Object(
                        self.register_worksheet_handle(workbook, sheet_id).0,
                    )),
                }
            }
            _ => Err(OmError::unsupported(format!(
                "Name.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_name(
        &mut self,
        workbook: WorkbookHandle,
        name_id: DefinedNameId,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Name", member, false)?;

        match member {
            "Delete" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Name.Delete does not accept arguments",
                    ));
                }
                let runtime = self.runtime_workbook_mut(workbook)?;
                if runtime.read_only {
                    return Err(OmError::new(
                        OmErrorCode::InvalidState,
                        "cannot modify a read-only workbook",
                    ));
                }
                runtime.loaded.state.defined_names.remove_by_id(name_id)?;
                let workbook_id = runtime.loaded.state.model().id;
                let workbook_display_name = runtime.loaded.state.model().display_name.clone();
                let worksheets = runtime.loaded.state.worksheets().to_vec();
                let defined_names = runtime.loaded.state.defined_names.clone();
                let mut chart_source_current_sheets = BTreeMap::new();
                for (chart_sheet_id, binding) in &runtime.loaded.state.chart_sheets {
                    chart_source_current_sheets.insert(binding.chart_id, *chart_sheet_id);
                }
                for drawing in runtime.loaded.state.drawings.values() {
                    for object in &drawing.objects {
                        let DrawingObjectModel::ChartFrame(chart_object) = object else {
                            continue;
                        };
                        chart_source_current_sheets
                            .insert(chart_object.chart_id, drawing.host_sheet_id);
                    }
                }
                let refresh_chart_source =
                    |source: &mut Option<ChartSourceExpr>, current_sheet: Option<SheetId>| {
                        let Some(source) = source.as_mut() else {
                            return;
                        };
                        source.resolved = resolve_chart_source_reference_with_names(
                            source.raw.text.as_str(),
                            workbook_id,
                            Some(&workbook_display_name),
                            &worksheets,
                            &defined_names,
                            current_sheet,
                        );
                    };
                for chart in runtime.loaded.state.charts.values_mut() {
                    let chart_current_sheet = chart_source_current_sheets.get(&chart.id).copied();
                    for series in &mut chart.series {
                        refresh_chart_source(&mut series.name, chart_current_sheet);
                        refresh_chart_source(&mut series.x_values, chart_current_sheet);
                        refresh_chart_source(&mut series.values, chart_current_sheet);
                        refresh_chart_source(&mut series.bubble_size, chart_current_sheet);
                    }
                }
                runtime.prompt_dirty = true;
                self.find_state = None;
                self.cut_copy_mode = None;
                self.clipboard = None;
                Ok(OmValue::Empty)
            }
            _ => Err(OmError::unsupported(format!(
                "Name.{member} is not implemented as a method"
            ))),
        }
    }
}
