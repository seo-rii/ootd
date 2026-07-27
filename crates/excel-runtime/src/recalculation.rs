use super::{
    EXCEL_MAX_COLUMN_INDEX, EXCEL_MAX_ROW_INDEX, ExcelRuntime, FormulaArrayResult, FormulaEvaluator,
};
use office_common::{CellError, CellValue, ObjectHandle, OmResult, Rect, SheetId, WorkbookHandle};

impl ExcelRuntime {
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

    pub(super) fn calculate_workbook_formulas(&mut self, workbook: WorkbookHandle) -> OmResult<()> {
        let sheet_ids = self
            .runtime_workbook(workbook)?
            .loaded
            .state
            .worksheets
            .iter()
            .map(|worksheet| worksheet.id)
            .collect::<Vec<_>>();
        for sheet_id in sheet_ids {
            self.calculate_sheet_formulas(workbook, sheet_id, None)?;
        }
        Ok(())
    }

    pub(super) fn calculate_sheet_formulas(
        &mut self,
        workbook: WorkbookHandle,
        sheet_id: SheetId,
        scope: Option<Rect>,
    ) -> OmResult<()> {
        let snapshot = self.runtime_workbook(workbook)?.loaded.state.clone();
        let worksheet = snapshot.worksheet_data_for_sheet(sheet_id)?;
        let formula_cells = worksheet
            .cells
            .iter()
            .filter_map(|(&(row, col), cell)| {
                if cell.formula.is_some()
                    && scope.is_none_or(|rect| {
                        (rect.row_first..=rect.row_last).contains(&row)
                            && (rect.col_first..=rect.col_last).contains(&col)
                    })
                {
                    Some((row, col))
                } else {
                    None
                }
            })
            .collect::<Vec<_>>();
        if formula_cells.is_empty() {
            return Ok(());
        }

        let mut dynamic_updates = Vec::<((u32, u32), FormulaArrayResult)>::new();
        let mut scalar_formula_cells = Vec::new();
        for (row, col) in formula_cells {
            if worksheet.dynamic_array_formulas.contains(&(row, col)) {
                let mut evaluator = FormulaEvaluator::new(&snapshot);
                if let Some(value) =
                    evaluator.evaluate_dynamic_array_formula_cell(sheet_id, row, col)
                {
                    dynamic_updates.push(((row, col), value));
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
                worksheet.clear_owned_spill(anchor);
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
                let mut evaluator = FormulaEvaluator::new(&scalar_snapshot);
                if let Some(value) = evaluator.evaluate_formula_cell(sheet_id, row, col) {
                    scalar_updates.push(((row, col), value));
                }
            }
            let worksheet = self
                .runtime_workbook_mut(workbook)?
                .loaded
                .state
                .worksheet_data_for_sheet_mut(sheet_id)?;
            for ((row, col), value) in scalar_updates {
                if let Some(cell) = worksheet.cells.get_mut(&(row, col)) {
                    cell.value = value;
                }
            }
        }
        Ok(())
    }
}
