use super::calc::FormulaEvalError;
use super::{
    EXCEL_MAX_COLUMN_INDEX, EXCEL_MAX_ROW_INDEX, ExcelRuntime, FormulaArrayResult,
    FormulaEvaluator, RuntimeCalculationSnapshot, WorkbookCalculationState,
};
use excel_model::WorkbookState;
use office_common::{CellError, CellValue, ObjectHandle, OmResult, Rect, SheetId, WorkbookHandle};
use sha2::{Digest, Sha256};

/// A one-based formula-cell address within a runtime workbook.
#[derive(Clone, Copy, Debug, PartialEq, Eq, PartialOrd, Ord)]
pub struct CalculationCell {
    pub sheet_id: SheetId,
    pub row: u32,
    pub column: u32,
}

/// A formula cell that completed evaluation with an Excel error value.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CalculationCellError {
    pub cell: CalculationCell,
    pub error: CellError,
}

/// Outcomes from one workbook calculation pass.
///
/// `volatile` is an annotation and can overlap `evaluated` or `errors`. Unsupported and external
/// formulas retain their existing cached values; circular formulas remain mapped to `#CALC!`.
/// All three unresolved categories make the report incomplete.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct CalculationReport {
    pub evaluated: Vec<CalculationCell>,
    pub unsupported: Vec<CalculationCell>,
    pub external: Vec<CalculationCell>,
    pub circular: Vec<CalculationCell>,
    pub volatile: Vec<CalculationCell>,
    pub errors: Vec<CalculationCellError>,
}

impl CalculationReport {
    /// Returns whether every formula was evaluated without unresolved dependencies or semantics.
    pub fn is_complete(&self) -> bool {
        self.unsupported.is_empty() && self.external.is_empty() && self.circular.is_empty()
    }
}

pub(super) fn calculation_input_digest(state: &WorkbookState) -> [u8; 32] {
    let mut digest = Sha256::new();
    let update_bytes = |digest: &mut Sha256, value: &[u8]| {
        digest.update(u64::try_from(value.len()).unwrap_or(u64::MAX).to_le_bytes());
        digest.update(value);
    };

    digest.update([u8::from(state.model().date1904)]);
    digest.update(
        u64::try_from(state.worksheets.len())
            .unwrap_or(u64::MAX)
            .to_le_bytes(),
    );
    for worksheet in &state.worksheets {
        digest.update(worksheet.id.0.to_le_bytes());
        update_bytes(&mut digest, worksheet.name.as_bytes());
        digest.update([match worksheet.kind {
            office_common::SheetKind::Worksheet => 0,
            office_common::SheetKind::ChartSheet => 1,
            office_common::SheetKind::MacroSheet => 2,
            office_common::SheetKind::DialogSheet => 3,
        }]);
        digest.update([match worksheet.visibility {
            office_common::SheetVisibility::Visible => 0,
            office_common::SheetVisibility::Hidden => 1,
            office_common::SheetVisibility::VeryHidden => 2,
        }]);
    }
    digest.update(
        u64::try_from(state.worksheet_data().len())
            .unwrap_or(u64::MAX)
            .to_le_bytes(),
    );
    for (sheet_id, worksheet) in state.worksheet_data() {
        digest.update(sheet_id.0.to_le_bytes());
        digest.update(
            u64::try_from(worksheet.cells.len())
                .unwrap_or(u64::MAX)
                .to_le_bytes(),
        );
        for (&(row, column), cell) in &worksheet.cells {
            digest.update(row.to_le_bytes());
            digest.update(column.to_le_bytes());
            match &cell.value {
                CellValue::Blank => digest.update([0]),
                CellValue::Number(value) => {
                    digest.update([1]);
                    digest.update(value.to_bits().to_le_bytes());
                }
                CellValue::Text(value) => {
                    digest.update([2]);
                    update_bytes(&mut digest, value.as_bytes());
                }
                CellValue::Bool(value) => digest.update([3, u8::from(*value)]),
                CellValue::Error(value) => digest.update([4, *value as u8]),
            }
            if let Some(formula) = &cell.formula {
                digest.update([1, u8::from(formula.is_r1c1)]);
                update_bytes(&mut digest, formula.text.as_bytes());
            } else {
                digest.update([0]);
            }
        }
        digest.update(
            u64::try_from(worksheet.dynamic_array_formulas.len())
                .unwrap_or(u64::MAX)
                .to_le_bytes(),
        );
        for &(row, column) in &worksheet.dynamic_array_formulas {
            digest.update(row.to_le_bytes());
            digest.update(column.to_le_bytes());
        }
        digest.update(
            u64::try_from(worksheet.spill_ranges.len())
                .unwrap_or(u64::MAX)
                .to_le_bytes(),
        );
        for (&(row, column), spill) in &worksheet.spill_ranges {
            digest.update(row.to_le_bytes());
            digest.update(column.to_le_bytes());
            digest.update(spill.row_first.to_le_bytes());
            digest.update(spill.row_last.to_le_bytes());
            digest.update(spill.col_first.to_le_bytes());
            digest.update(spill.col_last.to_le_bytes());
        }
        digest.update(
            u64::try_from(worksheet.spill_owners.len())
                .unwrap_or(u64::MAX)
                .to_le_bytes(),
        );
        for (&(row, column), &(owner_row, owner_column)) in &worksheet.spill_owners {
            digest.update(row.to_le_bytes());
            digest.update(column.to_le_bytes());
            digest.update(owner_row.to_le_bytes());
            digest.update(owner_column.to_le_bytes());
        }
    }
    digest.update(
        u64::try_from(state.defined_names.len())
            .unwrap_or(u64::MAX)
            .to_le_bytes(),
    );
    for defined_name in state.defined_names.iter() {
        digest.update(defined_name.id.0.to_le_bytes());
        match defined_name.scope {
            office_common::NameScope::Workbook => digest.update([0]),
            office_common::NameScope::Worksheet(sheet_id) => {
                digest.update([1]);
                digest.update(sheet_id.0.to_le_bytes());
            }
        }
        update_bytes(&mut digest, defined_name.display_name.as_bytes());
        digest.update([u8::from(defined_name.refers_to.is_r1c1)]);
        update_bytes(&mut digest, defined_name.refers_to.text.as_bytes());
    }
    digest.finalize().into()
}

impl ExcelRuntime {
    pub(super) fn record_calculation_snapshot(
        &mut self,
        workbook: WorkbookHandle,
        state: WorkbookCalculationState,
    ) -> OmResult<()> {
        let input_digest = calculation_input_digest(&self.runtime_workbook(workbook)?.loaded.state);
        self.runtime_workbook_mut(workbook)?.calculation_snapshot =
            Some(RuntimeCalculationSnapshot {
                input_digest,
                state,
            });
        Ok(())
    }

    pub(super) fn calculate_all_open_workbooks(&mut self) -> OmResult<()> {
        let workbooks = self
            .workbooks
            .keys()
            .copied()
            .map(|handle| WorkbookHandle(ObjectHandle(handle)))
            .collect::<Vec<_>>();
        for workbook in workbooks {
            self.calculate_workbook_formulas(workbook)?;
        }
        Ok(())
    }

    /// Calculates one workbook and returns address-level outcomes for every formula in scope.
    pub fn calculate_workbook_with_report(
        &mut self,
        workbook: WorkbookHandle,
    ) -> OmResult<CalculationReport> {
        self.calculate_workbook_formulas(workbook)
    }

    pub(super) fn calculate_workbook_formulas(
        &mut self,
        workbook: WorkbookHandle,
    ) -> OmResult<CalculationReport> {
        let sheet_ids = self
            .runtime_workbook(workbook)?
            .loaded
            .state
            .worksheets
            .iter()
            .map(|worksheet| worksheet.id)
            .collect::<Vec<_>>();
        let mut report = CalculationReport::default();
        for sheet_id in sheet_ids {
            let sheet_report = self.calculate_sheet_formulas(workbook, sheet_id, None)?;
            report.evaluated.extend(sheet_report.evaluated);
            report.unsupported.extend(sheet_report.unsupported);
            report.external.extend(sheet_report.external);
            report.circular.extend(sheet_report.circular);
            report.volatile.extend(sheet_report.volatile);
            report.errors.extend(sheet_report.errors);
        }
        self.record_calculation_snapshot(
            workbook,
            if report.is_complete() {
                WorkbookCalculationState::FullyCalculated
            } else {
                WorkbookCalculationState::PartiallyCalculated
            },
        )?;
        Ok(report)
    }

    pub(super) fn calculate_sheet_formulas(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        scope: Option<Rect>,
    ) -> OmResult<CalculationReport> {
        let snapshot = self.runtime_workbook(workbook)?.loaded.state.clone();
        let worksheet = snapshot.worksheet_data_for_sheet(sheet_id)?;
        let formula_cells = worksheet
            .cells
            .iter()
            .filter_map(|(&(row, col), cell)| {
                let formula = cell.formula.as_ref()?;
                scope
                    .is_none_or(|rect| {
                        (rect.row_first..=rect.row_last).contains(&row)
                            && (rect.col_first..=rect.col_last).contains(&col)
                    })
                    .then(|| {
                        (
                            row,
                            col,
                            formula.text.clone(),
                            worksheet.dynamic_array_formulas.contains(&(row, col)),
                        )
                    })
            })
            .collect::<Vec<_>>();
        if formula_cells.is_empty() {
            return Ok(CalculationReport::default());
        }

        let formula_has_external_workbook_reference = |formula: &str| {
            let bytes = formula.as_bytes();
            let mut index = 0usize;
            let mut in_string = false;
            let mut saw_open_bracket = false;
            let mut saw_closed_bracket = false;
            while index < bytes.len() {
                let byte = bytes[index];
                if byte == b'"' {
                    if in_string && bytes.get(index + 1) == Some(&b'"') {
                        index += 2;
                        continue;
                    }
                    in_string = !in_string;
                    index += 1;
                    continue;
                }
                if in_string {
                    index += 1;
                    continue;
                }
                match byte {
                    b'[' => {
                        saw_open_bracket = true;
                        saw_closed_bracket = false;
                    }
                    b']' if saw_open_bracket => saw_closed_bracket = true,
                    b'!' if saw_closed_bracket => return true,
                    b'+' | b'-' | b'*' | b'/' | b'^' | b'&' | b'=' | b'<' | b'>' | b',' | b';'
                    | b'(' | b')' => {
                        saw_open_bracket = false;
                        saw_closed_bracket = false;
                    }
                    _ => {}
                }
                index += 1;
            }
            false
        };
        let formula_is_volatile = |formula: &str| {
            let bytes = formula.as_bytes();
            let mut index = 0usize;
            let mut in_string = false;
            while index < bytes.len() {
                if bytes[index] == b'"' {
                    if in_string && bytes.get(index + 1) == Some(&b'"') {
                        index += 2;
                        continue;
                    }
                    in_string = !in_string;
                    index += 1;
                    continue;
                }
                if in_string {
                    index += 1;
                    continue;
                }
                if !bytes[index].is_ascii_alphabetic() && bytes[index] != b'_' {
                    index += 1;
                    continue;
                }
                let start = index;
                index += 1;
                while index < bytes.len()
                    && (bytes[index].is_ascii_alphanumeric() || matches!(bytes[index], b'_' | b'.'))
                {
                    index += 1;
                }
                let mut next = index;
                while bytes.get(next).is_some_and(u8::is_ascii_whitespace) {
                    next += 1;
                }
                if bytes.get(next) != Some(&b'(') {
                    continue;
                }
                let function_name = formula[start..index].rsplit('.').next().unwrap_or_default();
                if [
                    "CELL",
                    "INDIRECT",
                    "INFO",
                    "NOW",
                    "OFFSET",
                    "RAND",
                    "RANDARRAY",
                    "RANDBETWEEN",
                    "TODAY",
                ]
                .iter()
                .any(|candidate| function_name.eq_ignore_ascii_case(candidate))
                {
                    return true;
                }
            }
            false
        };

        let mut report = CalculationReport::default();
        let mut formula_cache_changed = false;
        let mut dynamic_updates = Vec::<((u32, u32), FormulaArrayResult)>::new();
        let mut scalar_formula_cells = Vec::new();
        for (row, col, formula, is_dynamic) in formula_cells {
            let calculation_cell = CalculationCell {
                sheet_id,
                row,
                column: col,
            };
            if formula_is_volatile(&formula) {
                report.volatile.push(calculation_cell);
            }
            if formula_has_external_workbook_reference(&formula) {
                report.external.push(calculation_cell);
                continue;
            }
            if is_dynamic {
                let mut evaluator = FormulaEvaluator::new(&snapshot);
                match evaluator.evaluate_dynamic_array_formula_cell_result(sheet_id, row, col) {
                    Ok(result) => {
                        if let Some(CellValue::Error(error)) = result.values.first() {
                            report.errors.push(CalculationCellError {
                                cell: calculation_cell,
                                error: *error,
                            });
                        } else {
                            report.evaluated.push(calculation_cell);
                        }
                        dynamic_updates.push(((row, col), result));
                    }
                    Err(FormulaEvalError::Unsupported) => {
                        report.unsupported.push(calculation_cell);
                    }
                    Err(FormulaEvalError::Circular) => {
                        report.circular.push(calculation_cell);
                        dynamic_updates.push((
                            (row, col),
                            FormulaArrayResult {
                                rows: 1,
                                cols: 1,
                                values: vec![CellValue::Error(CellError::Calc)],
                            },
                        ));
                    }
                    Err(error) => {
                        if let Some(CellValue::Error(cell_error)) = error.into_cell_value() {
                            report.errors.push(CalculationCellError {
                                cell: calculation_cell,
                                error: cell_error,
                            });
                            dynamic_updates.push((
                                (row, col),
                                FormulaArrayResult {
                                    rows: 1,
                                    cols: 1,
                                    values: vec![CellValue::Error(cell_error)],
                                },
                            ));
                        }
                    }
                }
            } else {
                scalar_formula_cells.push((row, col));
            }
        }

        if !dynamic_updates.is_empty() {
            let worksheet = self
                .runtime_workbook_mut(workbook)?
                .loaded
                .state
                .worksheet_data_for_sheet_mut(sheet_id)?;
            for (anchor @ (row, col), result) in dynamic_updates {
                worksheet.clear_owned_spill(anchor)?;
                worksheet.dirty = true;
                worksheet.dirty_cells.insert(anchor);
                formula_cache_changed = true;
                let Some(row_last) = row
                    .checked_add(u32::try_from(result.rows).unwrap_or(u32::MAX))
                    .and_then(|value| value.checked_sub(1))
                else {
                    if let Some(cell) = worksheet.cells.get_mut(&anchor) {
                        cell.value = CellValue::Error(CellError::Spill);
                    }
                    continue;
                };
                let Some(col_last) = col
                    .checked_add(u32::try_from(result.cols).unwrap_or(u32::MAX))
                    .and_then(|value| value.checked_sub(1))
                else {
                    if let Some(cell) = worksheet.cells.get_mut(&anchor) {
                        cell.value = CellValue::Error(CellError::Spill);
                    }
                    continue;
                };
                if row_last > EXCEL_MAX_ROW_INDEX || col_last > EXCEL_MAX_COLUMN_INDEX {
                    if let Some(cell) = worksheet.cells.get_mut(&anchor) {
                        cell.value = CellValue::Error(CellError::Spill);
                    }
                    continue;
                }
                let spill_rect = Rect {
                    row_first: row,
                    row_last,
                    col_first: col,
                    col_last,
                };
                let obstructed = (row..=row_last).any(|target_row| {
                    (col..=col_last).any(|target_col| {
                        let key = (target_row, target_col);
                        key != anchor
                            && worksheet.cells.get(&key).is_some_and(|cell| {
                                cell.formula.is_some() || !matches!(cell.value, CellValue::Blank)
                            })
                    })
                });
                if obstructed {
                    if let Some(cell) = worksheet.cells.get_mut(&anchor) {
                        cell.value = CellValue::Error(CellError::Spill);
                    }
                    worksheet.dirty = true;
                    worksheet.dirty_cells.insert(anchor);
                    continue;
                }

                for (index, value) in result.values.into_iter().enumerate() {
                    let row_offset = index / result.cols;
                    let col_offset = index % result.cols;
                    let key = (row + row_offset as u32, col + col_offset as u32);
                    if key == anchor {
                        if let Some(cell) = worksheet.cells.get_mut(&key) {
                            cell.value = value;
                        }
                    } else if let Some(cell) = worksheet.cells.get_mut(&key) {
                        cell.value = value;
                        cell.formula = None;
                        worksheet.spill_owners.insert(key, anchor);
                    } else {
                        worksheet.cells.insert(
                            key,
                            excel_model::CellData {
                                value,
                                formula: None,
                                style_id: None,
                            },
                        );
                        worksheet.spill_owners.insert(key, anchor);
                    }
                    worksheet.dirty_cells.insert(key);
                }
                worksheet.spill_ranges.insert(anchor, spill_rect);
                worksheet.dirty = true;
            }
        }

        if !scalar_formula_cells.is_empty() {
            let scalar_snapshot = self.runtime_workbook(workbook)?.loaded.state.clone();
            let mut scalar_updates = Vec::with_capacity(scalar_formula_cells.len());
            for (row, col) in scalar_formula_cells {
                let calculation_cell = CalculationCell {
                    sheet_id,
                    row,
                    column: col,
                };
                let mut evaluator = FormulaEvaluator::new(&scalar_snapshot);
                match evaluator.evaluate_formula_cell_result(sheet_id, row, col) {
                    Ok(CellValue::Error(error)) => {
                        report.errors.push(CalculationCellError {
                            cell: calculation_cell,
                            error,
                        });
                        scalar_updates.push(((row, col), CellValue::Error(error)));
                    }
                    Ok(value) => {
                        report.evaluated.push(calculation_cell);
                        scalar_updates.push(((row, col), value));
                    }
                    Err(FormulaEvalError::Unsupported) => {
                        report.unsupported.push(calculation_cell);
                    }
                    Err(FormulaEvalError::Circular) => {
                        report.circular.push(calculation_cell);
                        scalar_updates.push(((row, col), CellValue::Error(CellError::Calc)));
                    }
                    Err(error) => {
                        if let Some(CellValue::Error(cell_error)) = error.into_cell_value() {
                            report.errors.push(CalculationCellError {
                                cell: calculation_cell,
                                error: cell_error,
                            });
                            scalar_updates.push(((row, col), CellValue::Error(cell_error)));
                        }
                    }
                }
            }
            let worksheet = self
                .runtime_workbook_mut(workbook)?
                .loaded
                .state
                .worksheet_data_for_sheet_mut(sheet_id)?;
            for ((row, col), value) in scalar_updates {
                let coordinates = (row, col);
                if let Some(cell) = worksheet.cells.get_mut(&coordinates) {
                    if cell.value == value {
                        continue;
                    }
                    cell.value = value;
                    worksheet.dirty_cells.insert(coordinates);
                    worksheet.dirty = true;
                    formula_cache_changed = true;
                }
            }
        }
        if formula_cache_changed {
            self.runtime_workbook_mut(workbook)?.formula_cache_dirty = true;
        }
        report.evaluated.sort_unstable();
        report.unsupported.sort_unstable();
        report.external.sort_unstable();
        report.circular.sort_unstable();
        report.volatile.sort_unstable();
        report.errors.sort_by_key(|outcome| outcome.cell);
        Ok(report)
    }
}
