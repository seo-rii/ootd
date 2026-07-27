use super::super::{
    APPLICATION_NAME, APPLICATION_VERSION, ExcelRuntime, RuntimeNamesScope, RuntimeObjectKind,
    RuntimeSelection, RuntimeSheetCollectionKind, XL_4_DIGIT_YEARS, XL_24_HOUR_CLOCK,
    XL_ALTERNATE_ARRAY_SEPARATOR, XL_COLUMN_SEPARATOR, XL_COUNTRY_CODE, XL_COUNTRY_SETTING,
    XL_CURRENCY_BEFORE, XL_CURRENCY_CODE, XL_CURRENCY_DIGITS, XL_CURRENCY_LEADING_ZEROS,
    XL_CURRENCY_MINUS_SIGN, XL_CURRENCY_NEGATIVE, XL_CURRENCY_SPACE_BEFORE,
    XL_CURRENCY_TRAILING_ZEROS, XL_DATE_ORDER, XL_DATE_SEPARATOR, XL_DAY_CODE, XL_DAY_LEADING_ZERO,
    XL_DECIMAL_SEPARATOR, XL_GENERAL_FORMAT_NAME, XL_HOUR_CODE, XL_LEFT_BRACE, XL_LEFT_BRACKET,
    XL_LIST_SEPARATOR, XL_LOWER_CASE_COLUMN_LETTER, XL_LOWER_CASE_ROW_LETTER, XL_MDY, XL_METRIC,
    XL_MINUTE_CODE, XL_MONTH_CODE, XL_MONTH_LEADING_ZERO, XL_MONTH_NAME_CHARS,
    XL_NON_ENGLISH_FUNCTIONS, XL_NONCURRENCY_DIGITS, XL_RIGHT_BRACE, XL_RIGHT_BRACKET,
    XL_ROW_SEPARATOR, XL_SECOND_CODE, XL_THOUSANDS_SEPARATOR, XL_TIME_LEADING_ZERO,
    XL_TIME_SEPARATOR, XL_UPPER_CASE_COLUMN_LETTER, XL_UPPER_CASE_ROW_LETTER,
    XL_WEEKDAY_NAME_CHARS, XL_YEAR_CODE, coerce_evaluate_expression_arg,
};
use office_common::{
    ObjectHandle, OmError, OmResult, OmValue, RangeArea, RangeSet, Rect, SheetScope, WorkbookHandle,
};

impl ExcelRuntime {
    pub(crate) fn dispatch_get_application(
        &mut self,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Application", member, false)?;
        if !args.is_empty()
            && !matches!(
                member,
                "Cells"
                    | "Rows"
                    | "Columns"
                    | "Workbooks"
                    | "Worksheets"
                    | "Sheets"
                    | "Charts"
                    | "Names"
                    | "International"
            )
        {
            return Err(OmError::invalid_argument(format!(
                "Application.{member} does not accept index arguments"
            )));
        }

        match member {
            "Workbooks" => {
                let handle = self.register_object(RuntimeObjectKind::WorkbooksCollection);
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "Worksheets" | "Sheets" | "Charts" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                let handle = self.register_object(RuntimeObjectKind::WorksheetsCollection {
                    workbook: active_workbook,
                    kind: match member {
                        "Worksheets" => RuntimeSheetCollectionKind::Worksheets,
                        "Sheets" => RuntimeSheetCollectionKind::Sheets,
                        "Charts" => RuntimeSheetCollectionKind::Charts,
                        _ => unreachable!("sheet collection member"),
                    },
                });
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "Names" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                let handle =
                    self.register_names_handle(active_workbook, RuntimeNamesScope::Workbook);
                if args.is_empty() {
                    Ok(OmValue::Object(handle))
                } else {
                    self.dispatch_invoke(handle, "Item", args)
                }
            }
            "WorksheetFunction" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.WorksheetFunction does not accept index arguments",
                    ));
                }
                Ok(OmValue::Object(
                    self.register_object(RuntimeObjectKind::WorksheetFunction),
                ))
            }
            "ActiveWorkbook" => Ok(self
                .active_workbook
                .map(|workbook| OmValue::Object(workbook.0))
                .unwrap_or(OmValue::Empty)),
            "ActiveSheet" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                let sheet_id = self.active_sheet_id(active_workbook)?;
                Ok(OmValue::Object(
                    self.register_sheet_object_handle(active_workbook, sheet_id)?,
                ))
            }
            "ActiveChart" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                if let Some((chart_workbook, chart_id, chart_object_parent)) = self.active_chart {
                    if chart_workbook == active_workbook
                        && self.chart_model(chart_workbook, chart_id).is_ok()
                    {
                        return Ok(OmValue::Object(
                            self.register_chart_handle_with_chart_object_parent_origin(
                                chart_workbook,
                                chart_id,
                                chart_object_parent,
                            ),
                        ));
                    }
                    self.active_chart = None;
                }
                let sheet_id = self.active_sheet_id(active_workbook)?;
                let chart_id = self
                    .runtime_workbook(active_workbook)?
                    .loaded
                    .state
                    .chart_sheets
                    .get(&sheet_id)
                    .map(|binding| binding.chart_id);
                Ok(chart_id
                    .map(|chart_id| {
                        OmValue::Object(self.register_chart_handle(active_workbook, chart_id))
                    })
                    .unwrap_or(OmValue::Empty))
            }
            "ActiveCell" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                let selection = self
                    .selection
                    .filter(|selection| selection.workbook == active_workbook)
                    .unwrap_or(self.default_selection(active_workbook)?);
                self.ensure_grid_worksheet(
                    active_workbook,
                    selection.sheet_id,
                    "Application.ActiveCell",
                )?;
                Ok(OmValue::Object(
                    self.register_range_handle(
                        active_workbook,
                        selection.sheet_id,
                        Rect::single_cell(selection.rect.row_first, selection.rect.col_first),
                    )
                    .0,
                ))
            }
            "Selection" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                if let Some((chart_workbook, chart_id, chart_object_parent)) = self.active_chart {
                    if chart_workbook == active_workbook
                        && self.chart_model(chart_workbook, chart_id).is_ok()
                    {
                        return Ok(OmValue::Object(
                            self.register_chart_handle_with_chart_object_parent_origin(
                                chart_workbook,
                                chart_id,
                                chart_object_parent,
                            ),
                        ));
                    }
                    self.active_chart = None;
                }
                let selection = self
                    .selection
                    .filter(|selection| selection.workbook == active_workbook)
                    .unwrap_or(self.default_selection(active_workbook)?);
                let active_sheet_chart_id = {
                    self.runtime_workbook(active_workbook)?
                        .loaded
                        .state
                        .chart_sheets
                        .get(&selection.sheet_id)
                        .map(|binding| binding.chart_id)
                };
                if let Some(chart_id) = active_sheet_chart_id {
                    return Ok(OmValue::Object(
                        self.register_chart_handle(active_workbook, chart_id),
                    ));
                }
                if let Some(range) = self.selection_range.clone() {
                    let active_workbook_id = self.workbook_model(active_workbook)?.id;
                    if range.workbook_id() == active_workbook_id {
                        let (range_sheet_id, range_rect) = Self::range_set_first_area(&range)?;
                        if range_sheet_id == selection.sheet_id && range_rect == selection.rect {
                            return Ok(OmValue::Object(
                                self.register_range_set_handle(active_workbook, range).0,
                            ));
                        }
                    }
                }
                Ok(OmValue::Object(
                    self.register_range_handle(active_workbook, selection.sheet_id, selection.rect)
                        .0,
                ))
            }
            "Name" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.Name does not accept index arguments",
                    ));
                }
                Ok(OmValue::Text(APPLICATION_NAME.to_string()))
            }
            "Version" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.Version does not accept index arguments",
                    ));
                }
                Ok(OmValue::Text(APPLICATION_VERSION.to_string()))
            }
            "UserName" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.UserName does not accept index arguments",
                    ));
                }
                Ok(OmValue::Text(self.user_name.clone()))
            }
            "DefaultFilePath" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.DefaultFilePath does not accept index arguments",
                    ));
                }
                Ok(OmValue::Text(self.default_file_path.clone()))
            }
            "Caption" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.Caption does not accept index arguments",
                    ));
                }
                Ok(OmValue::Text(self.caption.clone()))
            }
            "DisplayAlerts" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.DisplayAlerts does not accept index arguments",
                    ));
                }
                Ok(OmValue::Bool(self.display_alerts))
            }
            "Calculation" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.Calculation does not accept index arguments",
                    ));
                }
                Ok(OmValue::Number(f64::from(self.calculation)))
            }
            "ScreenUpdating" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.ScreenUpdating does not accept index arguments",
                    ));
                }
                Ok(OmValue::Bool(self.screen_updating))
            }
            "EnableEvents" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.EnableEvents does not accept index arguments",
                    ));
                }
                Ok(OmValue::Bool(self.enable_events))
            }
            "StatusBar" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.StatusBar does not accept index arguments",
                    ));
                }
                Ok(self
                    .status_bar
                    .as_ref()
                    .map(|value| OmValue::Text(value.clone()))
                    .unwrap_or(OmValue::Bool(false)))
            }
            "DisplayStatusBar" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.DisplayStatusBar does not accept index arguments",
                    ));
                }
                Ok(OmValue::Bool(self.display_status_bar))
            }
            "DisplayFormulaBar" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.DisplayFormulaBar does not accept index arguments",
                    ));
                }
                Ok(OmValue::Bool(self.display_formula_bar))
            }
            "DisplayScrollBars" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.DisplayScrollBars does not accept index arguments",
                    ));
                }
                Ok(OmValue::Bool(self.display_scroll_bars))
            }
            "DisplayFullScreen" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.DisplayFullScreen does not accept index arguments",
                    ));
                }
                Ok(OmValue::Bool(self.display_full_screen))
            }
            "UseSystemSeparators" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.UseSystemSeparators does not accept index arguments",
                    ));
                }
                Ok(OmValue::Bool(self.use_system_separators))
            }
            "DecimalSeparator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.DecimalSeparator does not accept index arguments",
                    ));
                }
                Ok(OmValue::Text(self.decimal_separator.clone()))
            }
            "ThousandsSeparator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.ThousandsSeparator does not accept index arguments",
                    ));
                }
                Ok(OmValue::Text(self.thousands_separator.clone()))
            }
            "International" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "Application.International expects one XlApplicationInternational index",
                    ));
                };
                let OmValue::Number(index) = index else {
                    return Err(OmError::type_mismatch(
                        "Application.International index expects an XlApplicationInternational numeric value",
                    ));
                };
                if !index.is_finite()
                    || index.fract() != 0.0
                    || *index < i32::MIN as f64
                    || *index > i32::MAX as f64
                {
                    return Err(OmError::invalid_argument(
                        "Application.International index expects an integral XlApplicationInternational value",
                    ));
                }
                let list_separator = || {
                    if self.decimal_separator == "," {
                        ";"
                    } else {
                        ","
                    }
                };
                match *index as i32 {
                    XL_COUNTRY_CODE | XL_COUNTRY_SETTING => Ok(OmValue::Number(1.0)),
                    XL_DECIMAL_SEPARATOR => Ok(OmValue::Text(self.decimal_separator.clone())),
                    XL_THOUSANDS_SEPARATOR => Ok(OmValue::Text(self.thousands_separator.clone())),
                    XL_LIST_SEPARATOR => Ok(OmValue::Text(list_separator().to_string())),
                    XL_UPPER_CASE_ROW_LETTER => Ok(OmValue::Text("R".to_string())),
                    XL_UPPER_CASE_COLUMN_LETTER => Ok(OmValue::Text("C".to_string())),
                    XL_LOWER_CASE_ROW_LETTER => Ok(OmValue::Text("r".to_string())),
                    XL_LOWER_CASE_COLUMN_LETTER => Ok(OmValue::Text("c".to_string())),
                    XL_LEFT_BRACKET => Ok(OmValue::Text("[".to_string())),
                    XL_RIGHT_BRACKET => Ok(OmValue::Text("]".to_string())),
                    XL_LEFT_BRACE => Ok(OmValue::Text("{".to_string())),
                    XL_RIGHT_BRACE => Ok(OmValue::Text("}".to_string())),
                    XL_COLUMN_SEPARATOR => Ok(OmValue::Text(list_separator().to_string())),
                    XL_ROW_SEPARATOR => Ok(OmValue::Text(";".to_string())),
                    XL_ALTERNATE_ARRAY_SEPARATOR => Ok(OmValue::Text("\\".to_string())),
                    XL_DATE_SEPARATOR => Ok(OmValue::Text("/".to_string())),
                    XL_TIME_SEPARATOR => Ok(OmValue::Text(":".to_string())),
                    XL_YEAR_CODE => Ok(OmValue::Text("y".to_string())),
                    XL_MONTH_CODE => Ok(OmValue::Text("m".to_string())),
                    XL_DAY_CODE => Ok(OmValue::Text("d".to_string())),
                    XL_HOUR_CODE => Ok(OmValue::Text("h".to_string())),
                    XL_MINUTE_CODE => Ok(OmValue::Text("m".to_string())),
                    XL_SECOND_CODE => Ok(OmValue::Text("s".to_string())),
                    XL_CURRENCY_CODE => Ok(OmValue::Text("$".to_string())),
                    XL_GENERAL_FORMAT_NAME => Ok(OmValue::Text("General".to_string())),
                    XL_CURRENCY_DIGITS | XL_NONCURRENCY_DIGITS => Ok(OmValue::Number(2.0)),
                    XL_CURRENCY_NEGATIVE => Ok(OmValue::Number(0.0)),
                    XL_MONTH_NAME_CHARS | XL_WEEKDAY_NAME_CHARS => Ok(OmValue::Number(3.0)),
                    XL_DATE_ORDER => Ok(OmValue::Number(0.0)),
                    XL_24_HOUR_CLOCK => Ok(OmValue::Bool(false)),
                    XL_NON_ENGLISH_FUNCTIONS => Ok(OmValue::Bool(false)),
                    XL_METRIC => Ok(OmValue::Bool(false)),
                    XL_CURRENCY_SPACE_BEFORE => Ok(OmValue::Bool(false)),
                    XL_CURRENCY_BEFORE => Ok(OmValue::Bool(true)),
                    XL_CURRENCY_MINUS_SIGN => Ok(OmValue::Bool(false)),
                    XL_CURRENCY_TRAILING_ZEROS => Ok(OmValue::Bool(true)),
                    XL_CURRENCY_LEADING_ZEROS => Ok(OmValue::Bool(true)),
                    XL_MONTH_LEADING_ZERO => Ok(OmValue::Bool(false)),
                    XL_DAY_LEADING_ZERO => Ok(OmValue::Bool(false)),
                    XL_4_DIGIT_YEARS => Ok(OmValue::Bool(true)),
                    XL_MDY => Ok(OmValue::Bool(true)),
                    XL_TIME_LEADING_ZERO => Ok(OmValue::Bool(false)),
                    other => Err(OmError::unsupported(format!(
                        "Application.International index {other} is not implemented"
                    ))),
                }
            }
            "ShowWindowsInTaskbar" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.ShowWindowsInTaskbar does not accept index arguments",
                    ));
                }
                Ok(OmValue::Bool(self.show_windows_in_taskbar))
            }
            "Interactive" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.Interactive does not accept index arguments",
                    ));
                }
                Ok(OmValue::Bool(self.interactive))
            }
            "CutCopyMode" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.CutCopyMode does not accept index arguments",
                    ));
                }
                Ok(self
                    .cut_copy_mode
                    .map(|value| OmValue::Number(f64::from(value)))
                    .unwrap_or(OmValue::Bool(false)))
            }
            "Cells" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                let sheet_id = self.active_sheet_id(active_workbook)?;
                if args.is_empty() {
                    self.dispatch_get_worksheet(active_workbook, sheet_id, "Cells", &[])
                } else {
                    self.dispatch_invoke_worksheet(active_workbook, sheet_id, "Cells", args)
                }
            }
            "Rows" | "Columns" => {
                let Some(active_workbook) = self.active_workbook else {
                    return Ok(OmValue::Empty);
                };
                let sheet_id = self.active_sheet_id(active_workbook)?;
                self.dispatch_get_worksheet(active_workbook, sheet_id, member, args)
            }
            _ => Err(OmError::unsupported(format!(
                "Application.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_application(
        &mut self,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Application", member, false)?;

        match member {
            "Quit" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.Quit does not accept arguments",
                    ));
                }
                let workbooks = self
                    .workbooks
                    .keys()
                    .copied()
                    .map(|handle| WorkbookHandle(ObjectHandle(handle)))
                    .collect::<Vec<_>>();
                for workbook in workbooks {
                    self.close_workbook(workbook)?;
                }
                Ok(OmValue::Empty)
            }
            "Calculate" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.Calculate does not accept arguments",
                    ));
                }
                self.calculate_all_open_workbooks()?;
                Ok(OmValue::Empty)
            }
            "CalculateFull" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.CalculateFull does not accept arguments",
                    ));
                }
                self.calculate_all_open_workbooks()?;
                Ok(OmValue::Empty)
            }
            "CalculateFullRebuild" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Application.CalculateFullRebuild does not accept arguments",
                    ));
                }
                let workbooks = self
                    .workbooks
                    .keys()
                    .copied()
                    .map(|handle| WorkbookHandle(ObjectHandle(handle)))
                    .collect::<Vec<_>>();
                for workbook in workbooks {
                    if self.invalidate_workbook_calc_chain(workbook)? {
                        self.runtime_workbook_mut(workbook)?.prompt_dirty = true;
                    }
                    self.calculate_workbook_formulas(workbook)?;
                }
                Ok(OmValue::Empty)
            }
            "Evaluate" => {
                let expression = coerce_evaluate_expression_arg(args, "Application.Evaluate")?;
                let Some(active_workbook) = self.active_workbook else {
                    return Err(OmError::invalid_state("application has no active workbook"));
                };
                let sheet_id = self.active_sheet_id(active_workbook)?;
                self.evaluate_formula_expression(
                    active_workbook,
                    sheet_id,
                    expression,
                    "Application.Evaluate",
                )
            }
            "Goto" => {
                if args.len() > 2 {
                    return Err(OmError::invalid_argument(
                        "Application.Goto accepts at most reference and scroll arguments",
                    ));
                }

                match args.get(1) {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => {}
                    Some(OmValue::Bool(_)) => {}
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Application.Goto scroll expects a boolean when provided",
                        ));
                    }
                }

                let (workbook, range) = match args.first() {
                    None | Some(OmValue::Missing | OmValue::Empty | OmValue::Null) => {
                        self.last_goto_range.clone().ok_or_else(|| {
                            OmError::invalid_state(
                                "Application.Goto without a reference requires a prior Application.Goto selection",
                            )
                        })?
                    }
                    Some(OmValue::Text(reference)) => {
                        self.resolve_application_range_text(reference)?
                    }
                    Some(OmValue::Object(handle)) => match self.runtime_object(*handle)? {
                        RuntimeObjectKind::Range {
                            workbook, range, ..
                        } => (workbook, range),
                        _ => {
                            return Err(OmError::type_mismatch(
                                "Application.Goto expects a Range object or A1-style text reference",
                            ));
                        }
                    },
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Application.Goto expects a Range object or A1-style text reference",
                        ));
                    }
                };

                self.ensure_range_set_targets_grid_worksheets(
                    workbook,
                    &range,
                    "Application.Goto",
                )?;
                let (sheet_id, rect) = Self::range_set_first_area(&range)?;
                self.set_range_selection(workbook, range.clone(), "Application.Goto")?;
                self.last_goto_selection = Some(RuntimeSelection {
                    workbook,
                    sheet_id,
                    rect,
                });
                self.last_goto_range = Some((workbook, range));
                Ok(OmValue::Empty)
            }
            "Range" => match args {
                [OmValue::Text(reference)] => {
                    let (workbook, range) = self.resolve_application_range_text(reference)?;
                    self.ensure_range_set_targets_grid_worksheets(
                        workbook,
                        &range,
                        "Application.Range",
                    )?;
                    self.remember_range_selection(workbook, &range, "Application.Range")?;
                    Ok(OmValue::Object(
                        self.register_range_set_handle(workbook, range).0,
                    ))
                }
                _ => {
                    let Some(active_workbook) = self.active_workbook else {
                        return Err(OmError::invalid_state("application has no active workbook"));
                    };
                    let sheet_id = self.active_sheet_id(active_workbook)?;
                    self.dispatch_invoke_worksheet(active_workbook, sheet_id, "Range", args)
                }
            },
            "Intersect" => {
                if !(2..=30).contains(&args.len()) {
                    return Err(OmError::invalid_argument(
                        "Application.Intersect expects 2 to 30 range arguments",
                    ));
                }

                let parse_range = |value: &OmValue,
                                   label: &str,
                                   runtime: &ExcelRuntime|
                 -> OmResult<(WorkbookHandle, RangeSet)> {
                    match value {
                        OmValue::Object(handle) => match runtime.runtime_object(*handle)? {
                            RuntimeObjectKind::Range {
                                workbook, range, ..
                            } => Ok((workbook, range)),
                            _ => Err(OmError::type_mismatch(format!(
                                "Application.Intersect {label} expects range objects"
                            ))),
                        },
                        _ => Err(OmError::type_mismatch(format!(
                            "Application.Intersect {label} expects range objects"
                        ))),
                    }
                };

                let (workbook, first_range) = parse_range(&args[0], "Arg1", self)?;
                let (sheet_id, mut current_rects) =
                    Self::range_set_single_sheet_rects(&first_range)?;
                for (index, arg) in args.iter().enumerate().skip(1) {
                    let (next_workbook, next_range) =
                        parse_range(arg, &format!("Arg{}", index + 1), self)?;
                    let (next_sheet_id, next_rects) =
                        Self::range_set_single_sheet_rects(&next_range)?;
                    if next_workbook != workbook || next_sheet_id != sheet_id {
                        return Err(OmError::invalid_argument(
                            "Application.Intersect expects ranges from the same worksheet",
                        ));
                    }
                    let mut intersections = Vec::new();
                    for left in &current_rects {
                        for right in &next_rects {
                            let row_first = left.row_first.max(right.row_first);
                            let row_last = left.row_last.min(right.row_last);
                            let col_first = left.col_first.max(right.col_first);
                            let col_last = left.col_last.min(right.col_last);
                            if row_first <= row_last && col_first <= col_last {
                                intersections.push(Rect {
                                    row_first,
                                    row_last,
                                    col_first,
                                    col_last,
                                });
                            }
                        }
                    }
                    current_rects = intersections;
                    if current_rects.is_empty() {
                        return Ok(OmValue::Empty);
                    }
                }

                let workbook_id = self.workbook_model(workbook)?.id;
                let areas = current_rects
                    .into_iter()
                    .map(|rect| RangeArea::new(SheetScope::Single(sheet_id), rect))
                    .collect::<OmResult<Vec<_>>>()?;

                Ok(OmValue::Object(
                    self.register_range_set_handle(workbook, RangeSet::new(workbook_id, areas)?)
                        .0,
                ))
            }
            "Union" => {
                if !(2..=30).contains(&args.len()) {
                    return Err(OmError::invalid_argument(
                        "Application.Union expects 2 to 30 range arguments",
                    ));
                }

                let parse_range = |value: &OmValue,
                                   label: &str,
                                   runtime: &ExcelRuntime|
                 -> OmResult<(WorkbookHandle, RangeSet)> {
                    match value {
                        OmValue::Object(handle) => match runtime.runtime_object(*handle)? {
                            RuntimeObjectKind::Range {
                                workbook, range, ..
                            } => Ok((workbook, range)),
                            _ => Err(OmError::type_mismatch(format!(
                                "Application.Union {label} expects range objects"
                            ))),
                        },
                        _ => Err(OmError::type_mismatch(format!(
                            "Application.Union {label} expects range objects"
                        ))),
                    }
                };

                let (workbook, first_range) = parse_range(&args[0], "Arg1", self)?;
                let (sheet_id, first_rects) = Self::range_set_single_sheet_rects(&first_range)?;
                let workbook_id = self.workbook_model(workbook)?.id;
                let mut areas = first_rects
                    .into_iter()
                    .map(|rect| RangeArea::new(SheetScope::Single(sheet_id), rect))
                    .collect::<OmResult<Vec<_>>>()?;
                for (index, arg) in args.iter().enumerate().skip(1) {
                    let (next_workbook, next_range) =
                        parse_range(arg, &format!("Arg{}", index + 1), self)?;
                    let (next_sheet_id, next_rects) =
                        Self::range_set_single_sheet_rects(&next_range)?;
                    if next_workbook != workbook || next_sheet_id != sheet_id {
                        return Err(OmError::invalid_argument(
                            "Application.Union expects ranges from the same worksheet",
                        ));
                    }
                    for rect in next_rects {
                        areas.push(RangeArea::new(SheetScope::Single(sheet_id), rect)?);
                    }
                }

                Ok(OmValue::Object(
                    self.register_range_set_handle(workbook, RangeSet::new(workbook_id, areas)?)
                        .0,
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Application.{member} is not implemented as a method"
            ))),
        }
    }
}
