use super::super::{
    ExcelRuntime, RuntimeObjectKind, XL_CREATOR_CODE, format_formula_string_literal,
    format_rect_address_with_flags, formula_cell_error_text, formula_sheet_address_qualifier,
    worksheet_function_formula_name,
};
use office_common::{OmArray, OmError, OmResult, OmValue, RangeSet, SheetScope, WorkbookHandle};

impl ExcelRuntime {
    pub(crate) fn dispatch_get_worksheet_function(
        &mut self,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        self.focus_member_supported("WorksheetFunction", member, false)?;
        if !args.is_empty() {
            return Err(OmError::invalid_argument(format!(
                "WorksheetFunction.{member} does not accept arguments"
            )));
        }

        match member {
            "Application" | "Parent" => Ok(OmValue::Object(self.root_application())),
            "Creator" => Ok(OmValue::Number(f64::from(XL_CREATOR_CODE))),
            _ => Err(OmError::unsupported(format!(
                "WorksheetFunction.{member} is not implemented as a property"
            ))),
        }
    }

    pub(crate) fn dispatch_invoke_worksheet_function(
        &mut self,
        member: &str,
        args: &[OmValue],
    ) -> OmResult<OmValue> {
        let Some(active_workbook) = self.active_workbook else {
            return Err(OmError::invalid_state(
                "Application.WorksheetFunction requires an active workbook",
            ));
        };
        let active_sheet = self.active_sheet_id(active_workbook)?;
        let function_name = worksheet_function_formula_name(member)?;
        let mut formula_args = Vec::with_capacity(args.len());
        for arg in args {
            formula_args.push(self.worksheet_function_formula_arg(active_workbook, arg)?);
        }
        let expression = format!("={function_name}({})", formula_args.join(","));
        self.evaluate_formula_expression(
            active_workbook,
            active_sheet,
            &expression,
            &format!("WorksheetFunction.{member}"),
        )
    }

    fn worksheet_function_formula_arg(
        &self,
        active_workbook: WorkbookHandle,
        value: &OmValue,
    ) -> OmResult<String> {
        match value {
            OmValue::Missing | OmValue::Empty | OmValue::Null => Ok(String::new()),
            OmValue::Bool(value) => Ok(if *value { "TRUE" } else { "FALSE" }.to_string()),
            OmValue::Number(value) => {
                if !value.is_finite() {
                    return Err(OmError::invalid_argument(
                        "WorksheetFunction numeric arguments must be finite",
                    ));
                }
                Ok(value.to_string())
            }
            OmValue::Text(value) => Ok(format_formula_string_literal(value)),
            OmValue::Error(value) => Ok(formula_cell_error_text(value).to_string()),
            OmValue::Object(handle) => match self.runtime_object(*handle)? {
                RuntimeObjectKind::Range {
                    workbook, range, ..
                } => {
                    if workbook != active_workbook {
                        return Err(OmError::unsupported(
                            "WorksheetFunction range arguments must belong to the active workbook",
                        ));
                    }
                    self.worksheet_function_range_reference_text(workbook, &range)
                }
                _ => Err(OmError::type_mismatch(
                    "WorksheetFunction object arguments must be Range objects",
                )),
            },
            OmValue::Array(array) => self.worksheet_function_array_literal(active_workbook, array),
        }
    }

    fn worksheet_function_array_literal(
        &self,
        active_workbook: WorkbookHandle,
        array: &OmArray,
    ) -> OmResult<String> {
        let mut rows = Vec::with_capacity(array.rows);
        for row in 0..array.rows {
            let mut cols = Vec::with_capacity(array.cols);
            for col in 0..array.cols {
                let value = array.get(row, col).ok_or_else(|| {
                    OmError::invalid_argument("WorksheetFunction array dimensions are invalid")
                })?;
                cols.push(self.worksheet_function_formula_arg(active_workbook, value)?);
            }
            rows.push(cols.join(","));
        }
        Ok(format!("{{{}}}", rows.join(";")))
    }

    fn worksheet_function_range_reference_text(
        &self,
        workbook: WorkbookHandle,
        range: &RangeSet,
    ) -> OmResult<String> {
        let mut parts = Vec::with_capacity(range.areas().len());
        for area in range.areas() {
            let SheetScope::Single(sheet_id) = area.scope else {
                return Err(OmError::unsupported(
                    "WorksheetFunction range arguments do not support 3D references yet",
                ));
            };
            let worksheet_name = self.worksheet_model(workbook, sheet_id)?.name.clone();
            parts.push(format!(
                "{}{}",
                formula_sheet_address_qualifier(&worksheet_name),
                format_rect_address_with_flags(area.rect, true, true)
            ));
        }
        Ok(parts.join(","))
    }
}
