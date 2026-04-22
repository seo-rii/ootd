use std::collections::{BTreeMap, BTreeSet};

use office_common::{
    CellValue, FormulaSource, OmArray, OmError, OmErrorCode, OmResult, OmValue, OpaquePart,
    RangeRef, Rect, SheetId, SheetScope, StyleId, WorkbookId, WorkbookModel, WorksheetModel,
};

#[derive(Debug, Clone, PartialEq)]
pub struct CellData {
    pub value: CellValue,
    pub formula: Option<FormulaSource>,
    pub style_id: Option<StyleId>,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct WorksheetData {
    pub cells: BTreeMap<(u32, u32), CellData>,
    pub source_xml: Vec<u8>,
    pub dirty: bool,
    pub dirty_cells: BTreeSet<(u32, u32)>,
}

#[derive(Debug, Clone, PartialEq)]
pub struct WorkbookState {
    pub model: WorkbookModel,
    pub worksheets: Vec<WorksheetModel>,
    pub worksheet_data: BTreeMap<SheetId, WorksheetData>,
    pub opaque_parts: Vec<OpaquePart>,
}

impl WorkbookState {
    pub fn assign_workbook_id(&mut self, workbook_id: WorkbookId) {
        self.model.id = workbook_id;
        for worksheet in &mut self.worksheets {
            worksheet.workbook_id = workbook_id;
        }
    }

    pub fn set_worksheet_source_xml(&mut self, sheet_id: SheetId, source_xml: Vec<u8>) {
        self.worksheet_data(sheet_id).source_xml = source_xml;
    }

    pub fn insert_cell(&mut self, sheet_id: SheetId, row: u32, col: u32, cell: CellData) {
        let worksheet = self.worksheet_data(sheet_id);
        worksheet.cells.insert((row, col), cell);
        worksheet.dirty = true;
        worksheet.dirty_cells.insert((row, col));
    }

    pub fn cell(&self, sheet_id: SheetId, row: u32, col: u32) -> Option<&CellData> {
        self.worksheet_data
            .get(&sheet_id)
            .and_then(|worksheet| worksheet.cells.get(&(row, col)))
    }

    pub fn worksheet_data_for_sheet(&self, sheet_id: SheetId) -> OmResult<&WorksheetData> {
        self.worksheet_data.get(&sheet_id).ok_or_else(|| {
            OmError::new(
                OmErrorCode::NotFound,
                format!("unknown worksheet data for sheet {}", sheet_id.0),
            )
        })
    }

    pub fn worksheet_data_for_sheet_mut(
        &mut self,
        sheet_id: SheetId,
    ) -> OmResult<&mut WorksheetData> {
        self.worksheet_data.get_mut(&sheet_id).ok_or_else(|| {
            OmError::new(
                OmErrorCode::NotFound,
                format!("unknown worksheet data for sheet {}", sheet_id.0),
            )
        })
    }

    pub fn get_range_values(&self, range: &RangeRef) -> OmResult<OmArray> {
        let (sheet_id, rect) = self.single_sheet_rect(range)?;
        let worksheet = self.worksheet_data_for_sheet(sheet_id)?;
        let mut values = Vec::with_capacity((rect.height() * rect.width()) as usize);

        for row in rect.row_first..=rect.row_last {
            for col in rect.col_first..=rect.col_last {
                let value = worksheet
                    .cells
                    .get(&(row, col))
                    .map(|cell| OmValue::from(cell.value.clone()))
                    .unwrap_or(OmValue::Empty);
                values.push(value);
            }
        }

        OmArray::new(rect.height() as usize, rect.width() as usize, values)
    }

    pub fn get_range_formulas(&self, range: &RangeRef) -> OmResult<OmArray> {
        let (sheet_id, rect) = self.single_sheet_rect(range)?;
        let worksheet = self.worksheet_data_for_sheet(sheet_id)?;
        let mut values = Vec::with_capacity((rect.height() * rect.width()) as usize);

        for row in rect.row_first..=rect.row_last {
            for col in rect.col_first..=rect.col_last {
                let value = match worksheet.cells.get(&(row, col)) {
                    Some(cell) => match &cell.formula {
                        Some(formula) => OmValue::Text(format!("={}", formula.text)),
                        None => OmValue::from(cell.value.clone()),
                    },
                    None => OmValue::Empty,
                };
                values.push(value);
            }
        }

        OmArray::new(rect.height() as usize, rect.width() as usize, values)
    }

    pub fn set_range_values(&mut self, range: &RangeRef, values: &OmArray) -> OmResult<()> {
        let (sheet_id, rect) = self.single_sheet_rect(range)?;
        if values.rows != rect.height() as usize || values.cols != rect.width() as usize {
            return Err(OmError::invalid_argument(format!(
                "range dimensions {}x{} do not match value matrix {}x{}",
                rect.height(),
                rect.width(),
                values.rows,
                values.cols,
            )));
        }

        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        for row_offset in 0..values.rows {
            for col_offset in 0..values.cols {
                let row = rect.row_first + row_offset as u32;
                let col = rect.col_first + col_offset as u32;
                let key = (row, col);
                let value = CellValue::try_from(
                    values.values[row_offset * values.cols + col_offset].clone(),
                )?;

                if let Some(existing) = worksheet.cells.get_mut(&key) {
                    if existing.value == value && existing.formula.is_none() {
                        continue;
                    }

                    existing.value = value;
                    existing.formula = None;
                    if matches!(existing.value, CellValue::Blank) && existing.style_id.is_none() {
                        worksheet.cells.remove(&key);
                    }
                    worksheet.dirty = true;
                    worksheet.dirty_cells.insert(key);
                    continue;
                }

                if !matches!(value, CellValue::Blank) {
                    worksheet.cells.insert(
                        key,
                        CellData {
                            value,
                            formula: None,
                            style_id: None,
                        },
                    );
                    worksheet.dirty = true;
                    worksheet.dirty_cells.insert(key);
                }
            }
        }

        Ok(())
    }

    pub fn set_range_formulas(&mut self, range: &RangeRef, values: &OmArray) -> OmResult<()> {
        let (sheet_id, rect) = self.single_sheet_rect(range)?;
        if values.rows != rect.height() as usize || values.cols != rect.width() as usize {
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
                let value = values.values[row_offset * values.cols + col_offset].clone();
                let (cell_value, formula) = match value {
                    OmValue::Text(text) => {
                        if let Some(formula_text) = text.strip_prefix('=') {
                            (
                                CellValue::Blank,
                                Some(FormulaSource {
                                    text: formula_text.to_string(),
                                    is_r1c1: false,
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

        let worksheet = self.worksheet_data_for_sheet_mut(sheet_id)?;
        for (key, cell_value, formula) in updates {
            if let Some(existing) = worksheet.cells.get_mut(&key) {
                if existing.value == cell_value && existing.formula == formula {
                    continue;
                }

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
                continue;
            }

            if !matches!(cell_value, CellValue::Blank) || formula.is_some() {
                worksheet.cells.insert(
                    key,
                    CellData {
                        value: cell_value,
                        formula,
                        style_id: None,
                    },
                );
                worksheet.dirty = true;
                worksheet.dirty_cells.insert(key);
            }
        }

        Ok(())
    }

    fn worksheet_data(&mut self, sheet_id: SheetId) -> &mut WorksheetData {
        self.worksheet_data.entry(sheet_id).or_default()
    }

    fn single_sheet_rect(&self, range: &RangeRef) -> OmResult<(SheetId, Rect)> {
        if range.workbook_id != self.model.id {
            return Err(OmError::invalid_argument(format!(
                "range workbook {} does not match loaded workbook {}",
                range.workbook_id.0, self.model.id.0,
            )));
        }

        if range.areas.len() != 1 {
            return Err(OmError::unsupported(
                "multi-area ranges are not supported in the core worksheet model yet",
            ));
        }

        let rect = range.areas[0];
        if rect.row_first == 0 || rect.row_last == 0 || rect.col_first == 0 || rect.col_last == 0 {
            return Err(OmError::invalid_argument(
                "worksheet coordinates are 1-based and must be greater than zero",
            ));
        }
        if rect.row_first > rect.row_last || rect.col_first > rect.col_last {
            return Err(OmError::invalid_argument(
                "worksheet range bounds must be ordered",
            ));
        }

        let sheet_id = match range.scope {
            SheetScope::Single(sheet_id) => sheet_id,
            SheetScope::Multi3D { .. } => {
                return Err(OmError::unsupported(
                    "3D ranges are not supported in the core worksheet model yet",
                ));
            }
        };

        if !self
            .worksheets
            .iter()
            .any(|worksheet| worksheet.id == sheet_id)
        {
            return Err(OmError::new(
                OmErrorCode::NotFound,
                format!("unknown worksheet {}", sheet_id.0),
            ));
        }

        Ok((sheet_id, rect))
    }
}

#[cfg(test)]
mod tests {
    use super::{CellData, WorkbookState, WorksheetData};
    use std::collections::{BTreeMap, BTreeSet};

    use office_common::{
        CellValue, FileFormat, FormulaSource, ObjectHandle, OmArray, OmErrorCode, OmValue,
        RangeRef, Rect, SheetId, SheetScope, StyleId, WorkbookId, WorkbookModel, WorksheetModel,
    };

    fn sample_state() -> WorkbookState {
        let workbook_id = WorkbookId(7);
        let sheet_id = SheetId(3);
        WorkbookState {
            model: WorkbookModel {
                id: workbook_id,
                display_name: "Workbook".to_string(),
                format: FileFormat::Xlsx,
            },
            worksheets: vec![WorksheetModel {
                id: sheet_id,
                workbook_id,
                name: "Sheet1".to_string(),
                relationship_id: Some("rId1".to_string()),
                part_uri: Some("xl/worksheets/sheet1.xml".to_string()),
            }],
            worksheet_data: BTreeMap::from([(
                sheet_id,
                WorksheetData {
                    cells: BTreeMap::from([
                        (
                            (1, 1),
                            CellData {
                                value: CellValue::Number(42.0),
                                formula: None,
                                style_id: None,
                            },
                        ),
                        (
                            (2, 2),
                            CellData {
                                value: CellValue::Text("hello".to_string()),
                                formula: None,
                                style_id: Some(StyleId(9)),
                            },
                        ),
                    ]),
                    source_xml: b"<worksheet/>".to_vec(),
                    dirty: false,
                    dirty_cells: BTreeSet::new(),
                },
            )]),
            opaque_parts: Vec::new(),
        }
    }

    #[test]
    fn get_range_values_reads_sparse_cells_as_empty() {
        let state = sample_state();
        let values = state
            .get_range_values(&RangeRef::single_rect(
                WorkbookId(7),
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 2,
                    col_first: 1,
                    col_last: 2,
                },
            ))
            .expect("range values");

        assert_eq!(values.rows, 2);
        assert_eq!(values.cols, 2);
        assert_eq!(values.get(0, 0), Some(&OmValue::Number(42.0)));
        assert_eq!(values.get(0, 1), Some(&OmValue::Empty));
        assert_eq!(values.get(1, 0), Some(&OmValue::Empty));
        assert_eq!(values.get(1, 1), Some(&OmValue::Text("hello".to_string())));
    }

    #[test]
    fn set_range_values_updates_cells_in_row_major_order() {
        let mut state = sample_state();
        state
            .set_range_values(
                &RangeRef::single_rect(
                    WorkbookId(7),
                    SheetId(3),
                    Rect {
                        row_first: 1,
                        row_last: 2,
                        col_first: 1,
                        col_last: 2,
                    },
                ),
                &OmArray::new(
                    2,
                    2,
                    vec![
                        OmValue::Number(1.0),
                        OmValue::Text("two".to_string()),
                        OmValue::Bool(true),
                        OmValue::Empty,
                    ],
                )
                .expect("array"),
            )
            .expect("set range values");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(worksheet.dirty);
        assert_eq!(
            worksheet.dirty_cells,
            BTreeSet::from([(1, 1), (1, 2), (2, 1), (2, 2)])
        );
        assert_eq!(
            worksheet.cells.get(&(1, 1)).expect("A1").value,
            CellValue::Number(1.0)
        );
        assert_eq!(
            worksheet.cells.get(&(1, 2)).expect("B1").value,
            CellValue::Text("two".to_string())
        );
        assert_eq!(
            worksheet.cells.get(&(2, 1)).expect("A2").value,
            CellValue::Bool(true)
        );
        assert_eq!(
            worksheet.cells.get(&(2, 2)).expect("B2").style_id,
            Some(StyleId(9))
        );
        assert_eq!(
            worksheet.cells.get(&(2, 2)).expect("B2").value,
            CellValue::Blank
        );
    }

    #[test]
    fn set_range_values_rejects_mismatched_dimensions() {
        let mut state = sample_state();
        let result = state.set_range_values(
            &RangeRef::single_rect(
                WorkbookId(7),
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 2,
                    col_first: 1,
                    col_last: 2,
                },
            ),
            &OmArray::scalar(OmValue::Number(1.0)),
        );

        assert!(result.is_err());
    }

    #[test]
    fn set_range_values_rejects_multi_area_ranges() {
        let mut state = sample_state();
        let result = state.set_range_values(
            &RangeRef {
                workbook_id: WorkbookId(7),
                scope: office_common::SheetScope::Single(SheetId(3)),
                areas: vec![Rect::single_cell(1, 1), Rect::single_cell(2, 2)],
            },
            &OmArray::new(1, 2, vec![OmValue::Number(1.0), OmValue::Number(2.0)]).expect("array"),
        );

        assert!(result.is_err());
    }

    #[test]
    fn blank_write_removes_unstyled_cells() {
        let mut state = sample_state();
        state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 1, 1),
                &OmArray::scalar(OmValue::Empty),
            )
            .expect("clear cell");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(!worksheet.cells.contains_key(&(1, 1)));
        assert!(worksheet.dirty_cells.contains(&(1, 1)));
    }

    #[test]
    fn get_range_values_rejects_foreign_workbook_id() {
        let state = sample_state();
        let result =
            state.get_range_values(&RangeRef::single_cell(WorkbookId(999), SheetId(3), 1, 1));

        assert!(result.is_err());
    }

    #[test]
    fn get_range_formulas_returns_formula_text_and_constant_values() {
        let mut state = sample_state();
        let worksheet = state
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data");
        worksheet.cells.get_mut(&(1, 1)).expect("A1").formula = Some(FormulaSource {
            text: "SUM(B1:B2)".to_string(),
            is_r1c1: false,
        });

        let formulas = state
            .get_range_formulas(&RangeRef::single_rect(
                WorkbookId(7),
                SheetId(3),
                Rect {
                    row_first: 1,
                    row_last: 2,
                    col_first: 1,
                    col_last: 2,
                },
            ))
            .expect("range formulas");

        assert_eq!(
            formulas,
            OmArray::new(
                2,
                2,
                vec![
                    OmValue::Text("=SUM(B1:B2)".to_string()),
                    OmValue::Empty,
                    OmValue::Empty,
                    OmValue::Text("hello".to_string()),
                ],
            )
            .expect("array"),
        );
    }

    #[test]
    fn set_range_values_rejects_foreign_workbook_id() {
        let mut state = sample_state();
        let result = state.set_range_values(
            &RangeRef::single_cell(WorkbookId(999), SheetId(3), 1, 1),
            &OmArray::scalar(OmValue::Number(1.0)),
        );

        assert!(result.is_err());
    }

    #[test]
    fn get_range_values_rejects_zero_based_coordinates() {
        let state = sample_state();
        let result = state.get_range_values(&RangeRef::single_rect(
            WorkbookId(7),
            SheetId(3),
            Rect {
                row_first: 0,
                row_last: 1,
                col_first: 1,
                col_last: 1,
            },
        ));

        assert!(result.is_err());
    }

    #[test]
    fn set_range_values_rejects_unordered_rect_bounds() {
        let mut state = sample_state();
        let result = state.set_range_values(
            &RangeRef::single_rect(
                WorkbookId(7),
                SheetId(3),
                Rect {
                    row_first: 3,
                    row_last: 2,
                    col_first: 1,
                    col_last: 1,
                },
            ),
            &OmArray::scalar(OmValue::Number(1.0)),
        );

        assert!(result.is_err());
    }

    #[test]
    fn set_range_values_clears_formula_even_when_value_is_unchanged() {
        let mut state = sample_state();
        let worksheet = state
            .worksheet_data_for_sheet_mut(SheetId(3))
            .expect("worksheet data");
        worksheet.cells.get_mut(&(1, 1)).expect("A1").formula = Some(FormulaSource {
            text: "SUM(B1:B2)".to_string(),
            is_r1c1: false,
        });

        state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 1, 1),
                &OmArray::scalar(OmValue::Number(42.0)),
            )
            .expect("overwrite formula");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        let cell = worksheet.cells.get(&(1, 1)).expect("A1");
        assert_eq!(cell.value, CellValue::Number(42.0));
        assert!(cell.formula.is_none());
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(1, 1)));
    }

    #[test]
    fn set_range_values_keeps_clean_when_existing_cell_state_is_unchanged() {
        let mut state = sample_state();
        state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 2, 2),
                &OmArray::scalar(OmValue::Text("hello".to_string())),
            )
            .expect("rewrite same value");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(!worksheet.dirty);
        assert!(worksheet.dirty_cells.is_empty());
    }

    #[test]
    fn insert_cell_marks_worksheet_dirty_and_tracks_coordinate() {
        let mut state = sample_state();
        state.insert_cell(
            SheetId(3),
            3,
            1,
            CellData {
                value: CellValue::Number(9.0),
                formula: None,
                style_id: None,
            },
        );

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(3, 1)));
        assert_eq!(
            worksheet.cells.get(&(3, 1)).expect("A3").value,
            CellValue::Number(9.0)
        );
    }

    #[test]
    fn assign_workbook_id_updates_model_and_worksheet_ownership() {
        let mut state = sample_state();

        state.assign_workbook_id(WorkbookId(99));

        assert_eq!(state.model.id, WorkbookId(99));
        assert_eq!(state.worksheets[0].workbook_id, WorkbookId(99));
    }

    #[test]
    fn worksheet_data_lookup_returns_not_found_for_unknown_sheet() {
        let state = sample_state();

        let error = state
            .worksheet_data_for_sheet(SheetId(999))
            .expect_err("unknown worksheet data should fail");

        assert_eq!(error.code, OmErrorCode::NotFound);
        assert!(error.message.contains("999"));
    }

    #[test]
    fn get_range_values_rejects_unknown_sheet_id_with_not_found() {
        let state = sample_state();
        let error = state
            .get_range_values(&RangeRef::single_cell(WorkbookId(7), SheetId(999), 1, 1))
            .expect_err("unknown sheet should fail");

        assert_eq!(error.code, OmErrorCode::NotFound);
        assert!(error.message.contains("999"));
    }

    #[test]
    fn set_range_values_rejects_3d_ranges_with_unsupported_code() {
        let mut state = sample_state();
        let error = state
            .set_range_values(
                &RangeRef {
                    workbook_id: WorkbookId(7),
                    scope: SheetScope::Multi3D {
                        start: SheetId(3),
                        end: SheetId(4),
                    },
                    areas: vec![Rect::single_cell(1, 1)],
                },
                &OmArray::scalar(OmValue::Number(1.0)),
            )
            .expect_err("3D ranges should fail");

        assert_eq!(error.code, OmErrorCode::Unsupported);
    }

    #[test]
    fn set_range_values_rejects_object_values_with_type_mismatch_and_keeps_clean_state() {
        let mut state = sample_state();
        let error = state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 1, 1),
                &OmArray::scalar(OmValue::Object(ObjectHandle(5))),
            )
            .expect_err("object values should fail");

        assert_eq!(error.code, OmErrorCode::TypeMismatch);
        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(!worksheet.dirty);
        assert!(worksheet.dirty_cells.is_empty());
        assert_eq!(
            worksheet.cells.get(&(1, 1)).expect("A1").value,
            CellValue::Number(42.0)
        );
    }

    #[test]
    fn set_range_formulas_updates_formula_cells_and_plain_text_constants() {
        let mut state = sample_state();
        state
            .set_range_formulas(
                &RangeRef::single_rect(
                    WorkbookId(7),
                    SheetId(3),
                    Rect {
                        row_first: 1,
                        row_last: 1,
                        col_first: 1,
                        col_last: 2,
                    },
                ),
                &OmArray::new(
                    1,
                    2,
                    vec![
                        OmValue::Text("=SUM(B1:B2)".to_string()),
                        OmValue::Text("literal".to_string()),
                    ],
                )
                .expect("array"),
            )
            .expect("set formulas");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        assert!(worksheet.dirty);
        assert_eq!(worksheet.dirty_cells, BTreeSet::from([(1, 1), (1, 2)]));
        let a1 = worksheet.cells.get(&(1, 1)).expect("A1");
        assert_eq!(a1.value, CellValue::Blank);
        assert_eq!(
            a1.formula,
            Some(FormulaSource {
                text: "SUM(B1:B2)".to_string(),
                is_r1c1: false,
            })
        );
        let b1 = worksheet.cells.get(&(1, 2)).expect("B1");
        assert_eq!(b1.value, CellValue::Text("literal".to_string()));
        assert!(b1.formula.is_none());
    }

    #[test]
    fn blank_write_preserves_styled_cell_shell() {
        let mut state = sample_state();
        state
            .set_range_values(
                &RangeRef::single_cell(WorkbookId(7), SheetId(3), 2, 2),
                &OmArray::scalar(OmValue::Empty),
            )
            .expect("clear styled cell");

        let worksheet = state
            .worksheet_data_for_sheet(SheetId(3))
            .expect("worksheet data");
        let cell = worksheet.cells.get(&(2, 2)).expect("B2");
        assert_eq!(cell.value, CellValue::Blank);
        assert_eq!(cell.style_id, Some(StyleId(9)));
        assert!(worksheet.dirty);
        assert!(worksheet.dirty_cells.contains(&(2, 2)));
    }
}
