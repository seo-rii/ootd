use super::super::{
    EXCEL_MAX_COLUMN_INDEX, EXCEL_MAX_ROW_INDEX, ExcelRuntime, RangeAddressReferenceStyle,
    RangeProjection, RuntimeObjectKind, XL_CREATOR_CODE, coerce_optional_bool_arg,
    coerce_optional_reference_style_arg, coerce_u32_arg, convert_formula_a1_to_r1c1,
    convert_formula_r1c1_to_a1, format_external_address_qualifier, format_rect_address_with_flags,
    format_rect_r1c1_address_with_flags, formula_sheet_address_qualifier, om_value_is_omitted,
    render_range_text_value,
};
use office_common::{
    CellValue, FormulaSource, GetRangeValuesSpec, OmArray, OmError, OmErrorCode, OmResult, OmValue,
    RangeArea, RangeSet, Rect, SetRangeValuesSpec, SheetId, SheetScope, WorkbookHandle,
};

impl ExcelRuntime {
    pub(crate) fn dispatch_get_range(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        rect: Rect,
        projection: RangeProjection,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Range", member, false)?;
        if !args.is_empty() && !matches!(member, "Address" | "Rows" | "Columns" | "Cells") {
            return Err(OmError::invalid_argument(format!(
                "Range.{member} does not accept arguments"
            )));
        }

        match member {
            "Value" | "Value2" | "Formula" | "FormulaR1C1" | "Formula2" | "Formula2R1C1"
            | "FormulaLocal" | "FormulaR1C1Local" | "Formula2Local" | "Formula2R1C1Local" => {
                let array = if matches!(
                    member,
                    "Formula"
                        | "FormulaR1C1"
                        | "Formula2"
                        | "Formula2R1C1"
                        | "FormulaLocal"
                        | "FormulaR1C1Local"
                        | "Formula2Local"
                        | "Formula2R1C1Local"
                ) {
                    let wants_r1c1 = matches!(
                        member,
                        "FormulaR1C1" | "Formula2R1C1" | "FormulaR1C1Local" | "Formula2R1C1Local"
                    );
                    let worksheet = self
                        .runtime_workbook(workbook)?
                        .loaded
                        .state
                        .worksheet_data_for_sheet(sheet_id)?;
                    let mut values = Vec::with_capacity((rect.height() * rect.width()) as usize);
                    for row in rect.row_first..=rect.row_last {
                        for col in rect.col_first..=rect.col_last {
                            let value = match worksheet.cells.get(&(row, col)) {
                                Some(cell) => match &cell.formula {
                                    Some(formula) => {
                                        let text = if wants_r1c1 {
                                            if formula.is_r1c1 {
                                                formula.text.clone()
                                            } else {
                                                convert_formula_a1_to_r1c1(&formula.text, row, col)
                                            }
                                        } else if formula.is_r1c1 {
                                            convert_formula_r1c1_to_a1(&formula.text, row, col)
                                        } else {
                                            formula.text.clone()
                                        };
                                        OmValue::Text(format!("={text}"))
                                    }
                                    None => OmValue::from(cell.value.clone()),
                                },
                                None => OmValue::Empty,
                            };
                            values.push(value);
                        }
                    }
                    OmArray::new(rect.height() as usize, rect.width() as usize, values)?
                } else {
                    self.get_range_values(GetRangeValuesSpec {
                        workbook,
                        range: self.range_ref(workbook, sheet_id, rect)?,
                    })?
                };
                if array.rows == 1 && array.cols == 1 {
                    Ok(array.values.into_iter().next().unwrap_or(OmValue::Empty))
                } else {
                    Ok(OmValue::Array(array))
                }
            }
            "Text" => {
                let array = self.get_range_values(GetRangeValuesSpec {
                    workbook,
                    range: self.range_ref(workbook, sheet_id, rect)?,
                })?;
                let Some(first) = array.values.first() else {
                    return Ok(OmValue::Text(String::new()));
                };
                let first_text = render_range_text_value(first);
                if array.values.len() == 1 {
                    Ok(OmValue::Text(first_text))
                } else if array
                    .values
                    .iter()
                    .all(|value| render_range_text_value(value) == first_text)
                {
                    Ok(OmValue::Text(first_text))
                } else {
                    Ok(OmValue::Null)
                }
            }
            "HasFormula" => {
                let worksheet_data = self
                    .runtime_workbook(workbook)?
                    .loaded
                    .state
                    .worksheet_data_for_sheet(sheet_id)?;
                let mut has_formula = false;
                let mut has_non_formula = false;

                for row in rect.row_first..=rect.row_last {
                    for col in rect.col_first..=rect.col_last {
                        if worksheet_data
                            .cells
                            .get(&(row, col))
                            .and_then(|cell| cell.formula.as_ref())
                            .is_some()
                        {
                            has_formula = true;
                        } else {
                            has_non_formula = true;
                        }

                        if has_formula && has_non_formula {
                            return Ok(OmValue::Null);
                        }
                    }
                }

                Ok(OmValue::Bool(has_formula))
            }
            "Address" => {
                if args.len() > 5 {
                    return Err(OmError::invalid_argument(
                        "Range.Address accepts optional RowAbsolute, ColumnAbsolute, ReferenceStyle, External, and RelativeTo arguments",
                    ));
                }
                let row_absolute = args
                    .first()
                    .map(|value| {
                        coerce_optional_bool_arg(value, true, "Range.Address row absolute")
                    })
                    .transpose()?
                    .unwrap_or(true);
                let column_absolute = args
                    .get(1)
                    .map(|value| {
                        coerce_optional_bool_arg(value, true, "Range.Address column absolute")
                    })
                    .transpose()?
                    .unwrap_or(true);
                let reference_style = args
                    .get(2)
                    .map(coerce_optional_reference_style_arg)
                    .transpose()?
                    .unwrap_or(RangeAddressReferenceStyle::A1);
                let external = args
                    .get(3)
                    .map(|value| coerce_optional_bool_arg(value, false, "Range.Address external"))
                    .transpose()?
                    .unwrap_or(false);
                let external_prefix = if external {
                    let workbook_name = self.workbook_model(workbook)?.display_name.clone();
                    let worksheet_name = self
                        .worksheets(workbook)?
                        .iter()
                        .find(|worksheet| worksheet.id == sheet_id)
                        .map(|worksheet| worksheet.name.clone())
                        .ok_or_else(|| OmError::new(OmErrorCode::NotFound, "unknown worksheet"))?;
                    Some(format_external_address_qualifier(
                        &workbook_name,
                        &worksheet_name,
                    ))
                } else {
                    None
                };
                let relative_to = match args.get(4) {
                    None => None,
                    Some(value) if om_value_is_omitted(value) => None,
                    Some(OmValue::Object(handle)) => match self.runtime_object(*handle)? {
                        RuntimeObjectKind::Range { range, .. } => {
                            let (_, relative_rect) = Self::range_set_first_area(&range)?;
                            Some((relative_rect.row_first, relative_rect.col_first))
                        }
                        _ => {
                            return Err(OmError::type_mismatch(
                                "Range.Address RelativeTo expects a Range object when provided",
                            ));
                        }
                    },
                    Some(_) => {
                        return Err(OmError::type_mismatch(
                            "Range.Address RelativeTo expects a Range object when provided",
                        ));
                    }
                };
                let mut address = match reference_style {
                    RangeAddressReferenceStyle::A1 => {
                        format_rect_address_with_flags(rect, row_absolute, column_absolute)
                    }
                    RangeAddressReferenceStyle::R1C1 => {
                        let (base_row, base_col) =
                            relative_to.unwrap_or((rect.row_first, rect.col_first));
                        format_rect_r1c1_address_with_flags(
                            rect,
                            row_absolute,
                            column_absolute,
                            base_row,
                            base_col,
                        )
                    }
                };
                if let Some(prefix) = external_prefix {
                    address.insert_str(0, prefix.as_str());
                }
                Ok(OmValue::Text(address))
            }
            "Parent" => Ok(OmValue::Object(
                self.register_worksheet_handle(workbook, sheet_id).0,
            )),
            "Application" => Ok(OmValue::Object(self.root_application())),
            "Worksheet" => Ok(OmValue::Object(
                self.register_worksheet_handle(workbook, sheet_id).0,
            )),
            "Row" => Ok(OmValue::Number(rect.row_first as f64)),
            "Column" => Ok(OmValue::Number(rect.col_first as f64)),
            "Rows" => {
                let handle = self.register_projected_range_handle(
                    workbook,
                    sheet_id,
                    rect,
                    RangeProjection::Rows,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle.0))
                } else {
                    self.dispatch_invoke(handle.0, "Item", args)
                }
            }
            "Columns" => {
                let handle = self.register_projected_range_handle(
                    workbook,
                    sheet_id,
                    rect,
                    RangeProjection::Columns,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle.0))
                } else {
                    self.dispatch_invoke(handle.0, "Item", args)
                }
            }
            "Cells" => {
                let handle = self.register_projected_range_handle(
                    workbook,
                    sheet_id,
                    rect,
                    RangeProjection::Cells,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle.0))
                } else {
                    self.dispatch_invoke(handle.0, "Item", args)
                }
            }
            "CurrentRegion" => Ok(OmValue::Object(
                self.register_range_handle(
                    workbook,
                    sheet_id,
                    self.current_region_rect(workbook, sheet_id, rect)?,
                )
                .0,
            )),
            "EntireRow" => Ok(OmValue::Object(
                self.register_range_handle(
                    workbook,
                    sheet_id,
                    Rect {
                        row_first: rect.row_first,
                        row_last: rect.row_last,
                        col_first: 1,
                        col_last: EXCEL_MAX_COLUMN_INDEX,
                    },
                )
                .0,
            )),
            "EntireColumn" => Ok(OmValue::Object(
                self.register_range_handle(
                    workbook,
                    sheet_id,
                    Rect {
                        row_first: 1,
                        row_last: EXCEL_MAX_ROW_INDEX,
                        col_first: rect.col_first,
                        col_last: rect.col_last,
                    },
                )
                .0,
            )),
            "Count" | "CountLarge" => Ok(OmValue::Number(match projection {
                RangeProjection::Cells => u64::from(rect.width()) * u64::from(rect.height()),
                RangeProjection::Rows => u64::from(rect.height()),
                RangeProjection::Columns => u64::from(rect.width()),
            } as f64)),
            _ => Err(OmError::unsupported(format!(
                "Range.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_get_range_set(
        &mut self,
        workbook: WorkbookHandle,
        range: RangeSet,
        projection: RangeProjection,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if range.areas().len() == 1 {
            let (sheet_id, rect) = Self::range_set_single_area(&range)?;
            return self.dispatch_get_range(workbook, sheet_id, rect, projection, member, args);
        }

        self.focus_member_supported("Range", member, false)?;
        if !args.is_empty() && !matches!(member, "Address" | "Rows" | "Columns" | "Cells") {
            return Err(OmError::invalid_argument(format!(
                "Range.{member} does not accept arguments"
            )));
        }

        match member {
            "Areas" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Range.Areas does not accept index arguments",
                    ));
                }
                Ok(OmValue::Object(self.register_areas_handle(workbook, range)))
            }
            "Address" => self.range_set_address(workbook, &range, args),
            "Text" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Range.Text does not accept arguments",
                    ));
                }
                let state = &self.runtime_workbook(workbook)?.loaded.state;
                let mut first_text = None;
                for area in range.areas() {
                    let SheetScope::Single(sheet_id) = area.scope else {
                        return Err(OmError::unsupported(
                            "3D range handles are not supported by this operation yet",
                        ));
                    };
                    let worksheet_data = state.worksheet_data_for_sheet(sheet_id)?;
                    for row in area.rect.row_first..=area.rect.row_last {
                        for col in area.rect.col_first..=area.rect.col_last {
                            let value = worksheet_data
                                .cells
                                .get(&(row, col))
                                .map(|cell| OmValue::from(cell.value.clone()))
                                .unwrap_or(OmValue::Empty);
                            let text = render_range_text_value(&value);
                            match &first_text {
                                Some(first_text) if first_text != &text => {
                                    return Ok(OmValue::Null);
                                }
                                Some(_) => {}
                                None => first_text = Some(text),
                            }
                        }
                    }
                }

                Ok(OmValue::Text(first_text.unwrap_or_default()))
            }
            "HasFormula" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Range.HasFormula does not accept arguments",
                    ));
                }
                let state = &self.runtime_workbook(workbook)?.loaded.state;
                let mut has_formula = false;
                let mut has_non_formula = false;
                for area in range.areas() {
                    let SheetScope::Single(sheet_id) = area.scope else {
                        return Err(OmError::unsupported(
                            "3D range handles are not supported by this operation yet",
                        ));
                    };
                    let worksheet_data = state.worksheet_data_for_sheet(sheet_id)?;
                    for row in area.rect.row_first..=area.rect.row_last {
                        for col in area.rect.col_first..=area.rect.col_last {
                            if worksheet_data
                                .cells
                                .get(&(row, col))
                                .and_then(|cell| cell.formula.as_ref())
                                .is_some()
                            {
                                has_formula = true;
                            } else {
                                has_non_formula = true;
                            }

                            if has_formula && has_non_formula {
                                return Ok(OmValue::Null);
                            }
                        }
                    }
                }

                Ok(OmValue::Bool(has_formula))
            }
            "CurrentRegion" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Range.CurrentRegion does not accept arguments",
                    ));
                }
                let mut areas = Vec::with_capacity(range.len());
                for area in range.areas() {
                    let SheetScope::Single(sheet_id) = area.scope else {
                        return Err(OmError::unsupported(
                            "3D range handles are not supported by this operation yet",
                        ));
                    };
                    areas.push(RangeArea::new(
                        area.scope,
                        self.current_region_rect(workbook, sheet_id, area.rect)?,
                    )?);
                }
                Ok(OmValue::Object(
                    self.register_range_set_handle(
                        workbook,
                        RangeSet::new(range.workbook_id(), areas)?,
                    )
                    .0,
                ))
            }
            "EntireRow" | "EntireColumn" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(format!(
                        "Range.{member} does not accept arguments"
                    )));
                }
                let mut areas = Vec::with_capacity(range.len());
                for area in range.areas() {
                    if matches!(area.scope, SheetScope::Multi3D { .. }) {
                        return Err(OmError::unsupported(
                            "3D range handles are not supported by this operation yet",
                        ));
                    }
                    let rect = if member == "EntireRow" {
                        Rect {
                            row_first: area.rect.row_first,
                            row_last: area.rect.row_last,
                            col_first: 1,
                            col_last: EXCEL_MAX_COLUMN_INDEX,
                        }
                    } else {
                        Rect {
                            row_first: 1,
                            row_last: EXCEL_MAX_ROW_INDEX,
                            col_first: area.rect.col_first,
                            col_last: area.rect.col_last,
                        }
                    };
                    areas.push(RangeArea::new(area.scope, rect)?);
                }
                Ok(OmValue::Object(
                    self.register_range_set_handle(
                        workbook,
                        RangeSet::new(range.workbook_id(), areas)?,
                    )
                    .0,
                ))
            }
            "Count" | "CountLarge" => Ok(OmValue::Number(
                range
                    .areas()
                    .iter()
                    .map(|area| match projection {
                        RangeProjection::Cells => {
                            u64::from(area.rect.width()) * u64::from(area.rect.height())
                        }
                        RangeProjection::Rows => u64::from(area.rect.height()),
                        RangeProjection::Columns => u64::from(area.rect.width()),
                    })
                    .sum::<u64>() as f64,
            )),
            "Rows" => {
                let (sheet_id, rect) = Self::range_set_first_area(&range)?;
                let handle = self.register_projected_range_handle(
                    workbook,
                    sheet_id,
                    rect,
                    RangeProjection::Rows,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle.0))
                } else {
                    self.dispatch_invoke(handle.0, "Item", args)
                }
            }
            "Columns" => {
                let (sheet_id, rect) = Self::range_set_first_area(&range)?;
                let handle = self.register_projected_range_handle(
                    workbook,
                    sheet_id,
                    rect,
                    RangeProjection::Columns,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle.0))
                } else {
                    self.dispatch_invoke(handle.0, "Item", args)
                }
            }
            "Cells" => {
                let handle = self.register_projected_range_set_handle(
                    workbook,
                    range,
                    RangeProjection::Cells,
                );
                if args.is_empty() {
                    Ok(OmValue::Object(handle.0))
                } else {
                    self.dispatch_invoke(handle.0, "Item", args)
                }
            }
            _ => {
                let (sheet_id, rect) = Self::range_set_first_area(&range)?;
                self.dispatch_get_range(workbook, sheet_id, rect, projection, member, args)
            }
        }
    }

    pub(crate) fn dispatch_get_areas(
        &mut self,
        workbook: WorkbookHandle,
        range: RangeSet,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Areas", member, false)?;

        match member {
            "Count" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Areas.Count does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(range.len() as f64))
            }
            "Item" => self.dispatch_invoke_areas(workbook, range, member, args),
            "Application" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Areas.Application does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(self.root_application()))
            }
            "Creator" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Areas.Creator does not accept arguments",
                    ));
                }
                Ok(OmValue::Number(f64::from(XL_CREATOR_CODE)))
            }
            "Parent" => {
                if !args.is_empty() {
                    return Err(OmError::invalid_argument(
                        "Areas.Parent does not accept arguments",
                    ));
                }
                Ok(OmValue::Object(
                    self.register_range_set_handle(workbook, range).0,
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Areas.{member} is not implemented"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_areas(
        &mut self,
        workbook: WorkbookHandle,
        range: RangeSet,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("Areas", member, false)?;

        match member {
            "Item" => {
                let [index] = args else {
                    return Err(OmError::invalid_argument(
                        "Areas.Item expects a single 1-based index",
                    ));
                };
                let index = coerce_u32_arg(index, "Areas.Item index")? as usize;
                if index == 0 {
                    return Err(OmError::invalid_argument(
                        "Areas.Item index is out of bounds",
                    ));
                }
                let Some(area) = range.areas().get(index - 1).copied() else {
                    return Err(OmError::invalid_argument(
                        "Areas.Item index is out of bounds",
                    ));
                };
                let area_range = RangeSet::new(range.workbook_id(), vec![area])?;
                Ok(OmValue::Object(
                    self.register_range_set_handle(workbook, area_range).0,
                ))
            }
            _ => Err(OmError::unsupported(format!(
                "Areas.{member} is not implemented as a method"
            ))),
        }
    }

    pub(crate) fn range_set_address(
        &mut self,
        workbook: WorkbookHandle,
        range: &RangeSet,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        if args.len() > 5 {
            return Err(OmError::invalid_argument(
                "Range.Address accepts optional RowAbsolute, ColumnAbsolute, ReferenceStyle, External, and RelativeTo arguments",
            ));
        }
        let row_absolute = args
            .first()
            .map(|value| coerce_optional_bool_arg(value, true, "Range.Address row absolute"))
            .transpose()?
            .unwrap_or(true);
        let column_absolute = args
            .get(1)
            .map(|value| coerce_optional_bool_arg(value, true, "Range.Address column absolute"))
            .transpose()?
            .unwrap_or(true);
        let reference_style = args
            .get(2)
            .map(coerce_optional_reference_style_arg)
            .transpose()?
            .unwrap_or(RangeAddressReferenceStyle::A1);
        let external = args
            .get(3)
            .map(|value| coerce_optional_bool_arg(value, false, "Range.Address external"))
            .transpose()?
            .unwrap_or(false);
        let relative_to = match args.get(4) {
            None => None,
            Some(value) if om_value_is_omitted(value) => None,
            Some(OmValue::Object(handle)) => match self.runtime_object(*handle)? {
                RuntimeObjectKind::Range { range, .. } => {
                    let (_, relative_rect) = Self::range_set_first_area(&range)?;
                    Some((relative_rect.row_first, relative_rect.col_first))
                }
                _ => {
                    return Err(OmError::type_mismatch(
                        "Range.Address RelativeTo expects a Range object when provided",
                    ));
                }
            },
            Some(_) => {
                return Err(OmError::type_mismatch(
                    "Range.Address RelativeTo expects a Range object when provided",
                ));
            }
        };

        let workbook_name = if external {
            Some(self.workbook_model(workbook)?.display_name.clone())
        } else {
            None
        };
        let mut common_sheet_id = None;
        let mut spans_multiple_sheets = false;
        for area in range.areas() {
            let SheetScope::Single(sheet_id) = area.scope else {
                return Err(OmError::unsupported(
                    "3D range handles are not supported by Range.Address yet",
                ));
            };
            match common_sheet_id {
                Some(existing_sheet_id) if existing_sheet_id != sheet_id => {
                    spans_multiple_sheets = true;
                    break;
                }
                Some(_) => {}
                None => common_sheet_id = Some(sheet_id),
            }
        }
        let mut parts = Vec::with_capacity(range.areas().len());
        for area in range.areas() {
            let SheetScope::Single(sheet_id) = area.scope else {
                return Err(OmError::unsupported(
                    "3D range handles are not supported by Range.Address yet",
                ));
            };
            let mut address = match reference_style {
                RangeAddressReferenceStyle::A1 => {
                    format_rect_address_with_flags(area.rect, row_absolute, column_absolute)
                }
                RangeAddressReferenceStyle::R1C1 => {
                    let (base_row, base_col) =
                        relative_to.unwrap_or((area.rect.row_first, area.rect.col_first));
                    format_rect_r1c1_address_with_flags(
                        area.rect,
                        row_absolute,
                        column_absolute,
                        base_row,
                        base_col,
                    )
                }
            };
            if let Some(workbook_name) = workbook_name.as_ref() {
                let worksheet_name = self.worksheet_model(workbook, sheet_id)?.name.clone();
                address.insert_str(
                    0,
                    format_external_address_qualifier(workbook_name, &worksheet_name).as_str(),
                );
            } else if spans_multiple_sheets {
                let worksheet_name = self.worksheet_model(workbook, sheet_id)?.name.clone();
                address.insert_str(0, formula_sheet_address_qualifier(&worksheet_name).as_str());
            }
            parts.push(address);
        }

        Ok(OmValue::Text(parts.join(",")))
    }

    pub(crate) fn dispatch_set_range(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        rect: Rect,
        _projection: RangeProjection,
        member: &str,
        value: OmValue,
        args: &[OmValue],
    ) -> OmResult<()> {
        self.focus_member_supported("Range", member, true)?;
        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "Range.{member} does not accept index arguments"
            )));
        }

        match member {
            "Value" | "Value2" | "Formula" | "FormulaR1C1" | "Formula2" | "Formula2R1C1"
            | "FormulaLocal" | "FormulaR1C1Local" | "Formula2Local" | "Formula2R1C1Local" => {
                let values = match value {
                    OmValue::Array(array) => array,
                    scalar => OmArray::new(
                        rect.height() as usize,
                        rect.width() as usize,
                        vec![scalar; rect.height() as usize * rect.width() as usize],
                    )?,
                };
                if matches!(member, "Formula2" | "Formula2Local") {
                    self.set_range_dynamic_array_formulas(SetRangeValuesSpec {
                        workbook,
                        range: self.range_ref(workbook, sheet_id, rect)?,
                        values,
                    })?;
                } else if matches!(member, "Formula" | "FormulaLocal") {
                    self.set_range_formulas(SetRangeValuesSpec {
                        workbook,
                        range: self.range_ref(workbook, sheet_id, rect)?,
                        values,
                    })?;
                } else if matches!(
                    member,
                    "FormulaR1C1" | "Formula2R1C1" | "FormulaR1C1Local" | "Formula2R1C1Local"
                ) {
                    if values.rows != rect.height() as usize || values.cols != rect.width() as usize
                    {
                        return Err(OmError::invalid_argument(format!(
                            "range dimensions {}x{} do not match formula matrix {}x{}",
                            rect.height(),
                            rect.width(),
                            values.rows,
                            values.cols,
                        )));
                    }

                    let mut updates = Vec::with_capacity(values.values.len());
                    for row_offset in 0..values.rows {
                        for col_offset in 0..values.cols {
                            let row = rect.row_first + row_offset as u32;
                            let col = rect.col_first + col_offset as u32;
                            let value =
                                values.values[row_offset * values.cols + col_offset].clone();
                            let (cell_value, formula) = match value {
                                OmValue::Text(text) => {
                                    if let Some(formula_text) = text.strip_prefix('=') {
                                        (
                                            CellValue::Blank,
                                            Some(FormulaSource {
                                                text: formula_text.to_string(),
                                                is_r1c1: true,
                                            }),
                                        )
                                    } else {
                                        (CellValue::Text(text), None)
                                    }
                                }
                                other => (CellValue::try_from(other)?, None),
                            };
                            updates.push(((row, col), cell_value, formula));
                        }
                    }

                    let runtime = self.runtime_workbook_mut(workbook)?;
                    if runtime.read_only {
                        return Err(OmError::new(
                            OmErrorCode::InvalidState,
                            "cannot modify a read-only workbook",
                        ));
                    }
                    let worksheet = runtime
                        .loaded
                        .state
                        .worksheet_data_for_sheet_mut(sheet_id)?;
                    for (key, cell_value, formula) in updates {
                        let is_dynamic_formula =
                            matches!(member, "Formula2R1C1" | "Formula2R1C1Local")
                                && formula.is_some();
                        let unchanged = worksheet.cells.get(&key).is_some_and(|existing| {
                            existing.value == cell_value
                                && existing.formula == formula
                                && worksheet.dynamic_array_formulas.contains(&key)
                                    == is_dynamic_formula
                        });
                        if unchanged {
                            continue;
                        }
                        worksheet.prepare_cell_for_edit(key);
                        if let Some(existing) = worksheet.cells.get_mut(&key) {
                            existing.value = cell_value;
                            existing.formula = formula;
                            if matches!(existing.value, CellValue::Blank)
                                && existing.style_id.is_none()
                                && existing.formula.is_none()
                            {
                                worksheet.cells.remove(&key);
                            }
                            worksheet.dirty = true;
                            worksheet.dirty_cells.insert(key);
                        } else if !matches!(cell_value, CellValue::Blank) || formula.is_some() {
                            worksheet.cells.insert(
                                key,
                                excel_model::CellData {
                                    value: cell_value,
                                    formula,
                                    style_id: None,
                                },
                            );
                            worksheet.dirty = true;
                            worksheet.dirty_cells.insert(key);
                        }
                        if is_dynamic_formula {
                            worksheet.dynamic_array_formulas.insert(key);
                        }
                    }
                } else {
                    self.set_range_values(SetRangeValuesSpec {
                        workbook,
                        range: self.range_ref(workbook, sheet_id, rect)?,
                        values,
                    })?;
                }
                Ok(())
            }
            _ => Err(OmError::unsupported(format!(
                "Range.{member} is not writable"
            ))),
        }
    }
}
